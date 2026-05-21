const path = require('path');
const fs = require('fs').promises;
const Logger = require('../utils/logger');
const { SCRIPTS_PATH } = require('../utils/files');

const logger = new Logger({ minLevel: process.env.COMPRAS_LOGGER_MIN_LEVEL || 'debug' });

const SNE_CERTIDOES = path.join(SCRIPTS_PATH, 'sne', 'CERTIDOES');

const VALIDITY_WARN_DAYS = parseInt(process.env.COMPRAS_SNE_VALIDITY_WARN_DAYS, 10) || 5;
const COMPANY_NAME_MAX_LENGTH = parseInt(process.env.COMPRAS_SNE_COMPANY_NAME_MAX_LENGTH, 10) || 25;

const FALLBACK_TRUNCATION_MARKER = '[...]';
const _rawTruncationMarker = (process.env.COMPRAS_SNE_TRUNCATION_MARKER || '').trim();
const TRUNCATION_MARKER = _rawTruncationMarker.length > 0 && _rawTruncationMarker.length <= 10
  ? _rawTruncationMarker
  : FALLBACK_TRUNCATION_MARKER;

// Maps filename keywords (priority order) to canonical cert type names
const TYPE_KEYWORDS = [
  { keywords: ['relatório de credenciamento'], type: 'Credenciamento' },
  { keywords: ['registrada no sicaf', 'sicaf'], type: 'SICAF' },
  { keywords: ['receita federal', 'pgfn'], type: 'Receita Federal' },
  { keywords: ['fgts', 'cef'], type: 'FGTS' },
  { keywords: ['trabalhista', 'tst'], type: 'Trabalhista' },
  { keywords: ['receita estadual', 'estadual', 'distrital', 'dívida ativa do estado', 'divida ativa do estado'], type: 'Estadual' },
];

// Verification rules per cert type — add/edit rules here to extend behavior.
// Patterns run against normalized text (newlines collapsed to single space).
const CERT_CHECKS = {
  sicaf: {
    impedimento: {
      // "Impedimento de Licitar:Nada Consta" (no space after colon in SICAF)
      // Capture at most 20 chars — stops before next field's colon (e.g. "Ocorrências Impeditivas indiretas:")
      pattern: /Impedimento de Licitar[:\s]+([^:]{1,20})/i,
      expected: /nada\s+consta/i,
      failSuffix: 'IMPEDIDO',
      // When impediment is detected, confirm scope includes UFSC or federal government.
      // Captures up to 300 chars after the section heading (appears near end of the document).
      scopePattern: /Impedimento de Licitar no [Âa]mbito[:\s]+(.{1,300})/i,
      scopeRequired: /[Óo]rg[aã]os do Governo Federal|Universidade Federal de Santa Catarina/i,
    },
    dates: [
      // "Receita Federal e PGFN22/04/2026Automática" (date immediately after label)
      { label: 'Receita Federal', pattern: /Receita Federal e PGFN(\d{2}\/\d{2}\/\d{4})/i },
      // "FGTS04/04/2026Automática" — scoped to III block (before "IV -")
      { label: 'FGTS', pattern: /\bFGTS(\d{2}\/\d{2}\/\d{4})/i },
      // "TrabalhistaValidade:02/05/2026" or "Trabalhista Validade: 02/05/2026"
      { label: 'Trabalhista', pattern: /Trabalhista[^0-9]{0,20}(\d{2}\/\d{2}\/\d{4})/i },
      // "Receita Estadual/DistritalValidade:19/05/2026"
      { label: 'Estadual', pattern: /Receita Estadual(?:\/Distrital)?[^0-9]{0,20}(\d{2}\/\d{2}\/\d{4})/i },
    ],
  },
  fgts: {
    // "Validade:25/03/2026 a 23/04/2026" — get expiration (end of range);
    // fallback for single-date format. firstMatch: true stops after first hit.
    dates: [
      { label: 'FGTS', pattern: /[Vv]alidade:\d{2}\/\d{2}\/\d{4}\s+a\s+(\d{2}\/\d{2}\/\d{4})/i, firstMatch: true },
      { label: 'FGTS', pattern: /[Vv]alidade:(\d{2}\/\d{2}\/\d{4})/i, firstMatch: true },
    ],
  },
  'receita federal': {
    dates: [
      { label: 'Receita Federal', pattern: /[Vv][áa]lida?\s+at[eé]\s*:?\s*(\d{2}[\/.\-]\d{2}[\/.\-]\d{4})/i },
      { label: 'Receita Federal', pattern: /[Vv]alidade[^0-9]*(\d{2}[\/.\-]\d{2}[\/.\-]\d{4})/i },
    ],
  },
  trabalhista: {
    dates: [
      { label: 'Trabalhista', pattern: /[Vv][áa]lida?\s+at[eé]\s*:?\s*(\d{2}[\/.\-]\d{2}[\/.\-]\d{4})/i },
      { label: 'Trabalhista', pattern: /[Vv]alidade[^0-9]*(\d{2}[\/.\-]\d{2}[\/.\-]\d{4})/i },
    ],
  },
  estadual: {
    // Both patterns require a colon before the date, preventing false matches when PDF text
    // extraction outputs labels and values out of order (labels first, then values in a separate column).
    dates: [
      { label: 'Estadual', pattern: /[Vv]álida?\s+até\s*:?\s*(\d{2}[\/.\-]\d{2}[\/.\-]\d{4})/i },
      // Handles "Validade: DD/MM/YYYY" and "Validade (conforme Lei nº xxxx): DD/MM/YYYY"
      { label: 'Estadual', pattern: /[Vv]alidade[^:]*:\s*(\d{2}[\/.\-]\d{2}[\/.\-]\d{4})/i },
    ],
  },
};

function detectType(filename) {
  const lower = filename.toLowerCase();
  for (const entry of TYPE_KEYWORDS) {
    if (entry.keywords.some(kw => lower.includes(kw))) {
      return entry.type;
    }
  }
  return null;
}

// Collapse newlines into spaces so multi-line fields become single-line
function normalizeText(raw) {
  return raw.replace(/\r?\n\s*/g, ' ').replace(/\s{2,}/g, ' ');
}

function extractCnpj(text) {
  // Matches formatted CNPJ: XX.XXX.XXX/XXXX-XX or digits-only variants
  const match = text.match(/\d{2}[\.\s]?\d{3}[\.\s]?\d{3}[\.\s\/]?\d{4}[\.\s\-]?\d{2}/);
  if (!match) return null;
  return match[0].replace(/[.\-\/\s]/g, '');
}

// Extracts company name from RAW (non-normalized) text so newlines act as natural field delimiters.
// \s* between label tokens matches optional newlines in multi-line formats (e.g. FGTS PDFs).
// [^\n\r:]{2,} captures everything up to the next newline or colon.
function extractCompanyName(rawText) {
  const patterns = [
    /Raz[aã]o\s*Social\s*:[ \t]*(?:\r?\n[ \t]*)?([^\n\r:]{2,}(?:\r?\n[ \t]+[^\n\r:]+)?)/i,
    /Nome\s*Empresarial\s*:[ \t]*(?:\r?\n[ \t]*)?([^\n\r:]{2,}(?:\r?\n[ \t]+[^\n\r:]+)?)/i,
    /Nome\/Raz[aã]o\s*Social\s*:[ \t]*(?:\r?\n[ \t]*)?([^\n\r:]{2,}(?:\r?\n[ \t]+[^\n\r:]+)?)/i,
    // "Nome (razão social):" — used by SC Estadual and some other certidões
    /Nome\s*\(raz[aã]o\s*social\)\s*:[ \t]*(?:\r?\n[ \t]*)?([^\n\r:]{2,}(?:\r?\n[ \t]+[^\n\r:]+)?)/i,
    /Contribuinte\s*:[ \t]*(?:\r?\n[ \t]*)?([^\n\r:]{2,}(?:\r?\n[ \t]+[^\n\r:]+)?)/i,
    // Two-column layout: "NOME: CNPJ\nCOMPANY" — CNPJ label on same line as NOME, name on next line
    /^Nome\s*:[ \t]*CNPJ[^\n\r]*\r?\n[ \t]*([A-Za-zÀ-ÿ][^\n\r:]{2,})/im,
    // Two-column layout: "NOME:\nCNPJ\nCOMPANY" — CNPJ label on the line right after NOME
    /^Nome\s*:[ \t]*\r?\n[ \t]*CNPJ[^\n\r]*\r?\n[ \t]*([A-Za-zÀ-ÿ][^\n\r:]{2,})/im,
    // "Nome: AIVY VARIEDADES LTDA" — Receita Federal uses bare "Nome:" label, value on same line
    /^Nome\s*:[ \t]*([^\n\r:]{2,})/im,
  ];
  // Words that are field labels, not company names
  const LABEL_WORDS = /^(cnpj|cpf|ie|im|insc\.?\s*estadual|insc\.?\s*municipal)$/i;
  for (const p of patterns) {
    const m = rawText.match(p);
    if (m) {
      const name = m[1].replace(/\s+/g, ' ').trim();
      if (name.length > 1 && !LABEL_WORDS.test(name)) return name;
    }
  }
  return null;
}

function parseDate(str) {
  const [d, m, y] = str.split(/[\/.\-]/).map(Number);
  return new Date(y, m - 1, d);
}

function classifyDate(dateStr) {
  const exp = parseDate(dateStr);
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const diffDays = Math.ceil((exp - today) / 86400000);
  const datePart = dateStr.replace(/\//g, '-');
  if (diffDays < 0) return { status: 'VENCIDA', label: 'VENCIDA' };
  if (diffDays <= VALIDITY_WARN_DAYS) return { status: 'A_VENCER', label: `A VENCER EM ${datePart}` };
  return { status: 'VALIDA', label: `VÁLIDA ATÉ ${datePart}` };
}

function nearestDate(dates) {
  return dates.reduce((dateA, dateB) => parseDate(dateA) <= parseDate(dateB) ? dateA : dateB);
}

const VALIDITY_SORT_ORDER = { VALIDA: 3, A_VENCER: 2, VENCIDA: 1, SEM_VALIDADE: 0 };

// Remove chars forbidden in Windows filenames
function sanitizeFilename(str) {
  return str.replace(/[<>:"/\\|?*\n\r\t]/g, '').trim();
}

function analyzeCertidaoFromRaw(filename, rawText) {
  const text = normalizeText(rawText);
  const type = detectType(text);
  if (!type) {
    return { filename, error: 'Tipo de certidão não identificado', newName: null, type: null };
  }
  const cnpj = extractCnpj(text);
  const rawCompany = extractCompanyName(rawText);
  const company = rawCompany && rawCompany.length > COMPANY_NAME_MAX_LENGTH
    ? rawCompany.slice(0, COMPANY_NAME_MAX_LENGTH).trimEnd() + TRUNCATION_MARKER
    : rawCompany;
  const typeKey = type.toLowerCase();
  const checks = CERT_CHECKS[typeKey] || {};

  let impedido = false;
  const suffixes = [];

  if (checks.impedimento) {
    const m = text.match(checks.impedimento.pattern);
    if (m && !checks.impedimento.expected.test(m[1].trim())) {
      let scopeMatches = true;
      if (checks.impedimento.scopePattern) {
        const sm = text.match(checks.impedimento.scopePattern);
        scopeMatches = sm ? checks.impedimento.scopeRequired.test(sm[1]) : false;
      }
      if (scopeMatches) {
        suffixes.push(checks.impedimento.failSuffix);
        impedido = true;
      }
    }
  }

  const foundDates = [];
  const componentDates = {};
  if (type !== 'Credenciamento' && checks.dates) {
    for (const rule of checks.dates) {
      const m = text.match(rule.pattern);
      if (m) {
        foundDates.push(m[1]);
        if (rule.label && !(rule.label in componentDates)) componentDates[rule.label] = m[1];
        if (rule.firstMatch) break;
      }
    }
  }

  let validity = null;
  const componentValidity = {};
  if (type !== 'Credenciamento' && !impedido) {
    for (const [label, dateStr] of Object.entries(componentDates)) {
      componentValidity[label] = classifyDate(dateStr).status;
    }
    if (foundDates.length > 0) {
      const nearest = nearestDate(foundDates);
      validity = classifyDate(nearest);
      suffixes.push(validity.label);
    }
  }

  let emissionDate = null;
  if (type === 'Credenciamento') {
    const em = text.match(/[Ee]mitido\s+em[:\s]+(\d{2}\/\d{2}\/\d{4})/i);
    if (em) emissionDate = em[1];
  }

  const warnings = [];
  if (!cnpj) warnings.push('CNPJ não encontrado');
  if (!company) warnings.push('Razão social não encontrada');
  if (type !== 'Credenciamento' && foundDates.length === 0 && !impedido) warnings.push('Data de validade não encontrada');

  const nameParts = [];
  if (cnpj) nameParts.push(sanitizeFilename(cnpj));
  if (company) nameParts.push(sanitizeFilename(company));
  nameParts.push(type);
  if (suffixes.length > 0) nameParts.push(suffixes.join(' - '));

  const newName = nameParts.length > 1 ? `${nameParts.join(' - ')}.pdf` : null;

  return {
    filename,
    newName,
    type,
    cnpj,
    company,
    dates: foundDates,
    componentValidity: Object.keys(componentValidity).length > 0 ? componentValidity : undefined,
    validity: validity ? validity.status : (type === 'Credenciamento' ? 'SEM_VALIDADE' : null),
    validityLabel: validity ? validity.label : null,
    emissionDate: emissionDate || undefined,
    impedido,
    warnings,
    error: null,
  };
}

async function analyzeCertidao(filename, pdfParse) {
  const filePath = path.join(SNE_CERTIDOES, filename);
  const buffer = await fs.readFile(filePath);

  let rawText = '';
  try {
    const data = await pdfParse(buffer);
    rawText = data.text || '';
  } catch (e) {
    logger.warn(`Erro ao parsear PDF "${filename}": ${e.message}`, 'SNE');
    return { filename, error: 'Erro ao ler PDF', newName: null, type: null };
  }

  return analyzeCertidaoFromRaw(filename, rawText);
}

function sortByCnpj(certA, certB) {
  if (certA.cnpj && certB.cnpj) {
    return certA.cnpj !== certB.cnpj
      ? certA.cnpj.localeCompare(certB.cnpj)
      : (certA.type || '').localeCompare(certB.type || '');
  }
  if (certA.cnpj) return -1;
  if (certB.cnpj) return 1;
  return certA.filename.localeCompare(certB.filename);
}

async function listCertidoes() {
  let created = false;
  try {
    await fs.access(SNE_CERTIDOES);
  } catch {
    await fs.mkdir(SNE_CERTIDOES, { recursive: true });
    created = true;
  }

  const entries = await fs.readdir(SNE_CERTIDOES);
  const files = entries
    .filter(filename => filename.toLowerCase().endsWith('.pdf'))
    .map(filename => ({ name: filename, type: detectType(filename) }));

  files.sort((fileA, fileB) => {
    const cnpjA = (fileA.name.match(/\d{14}/) || [''])[0];
    const cnpjB = (fileB.name.match(/\d{14}/) || [''])[0];
    if (cnpjA && cnpjB) return cnpjA !== cnpjB ? cnpjA.localeCompare(cnpjB) : fileA.name.localeCompare(fileB.name);
    if (cnpjA) return -1;
    if (cnpjB) return 1;
    return fileA.name.localeCompare(fileB.name);
  });

  return { files, folderPath: SNE_CERTIDOES, folderCreated: created };
}

async function analyzeCertidoes(onProgress) {
  // Use lib path to avoid test-file loading at module init (pdf-parse v1 quirk)
  const pdfParse = require('pdf-parse/lib/pdf-parse.js');

  const { files, folderPath } = await listCertidoes();
  const total = files.length;
  const results = [];

  for (let i = 0; i < files.length; i++) {
    if (onProgress) onProgress(i + 1, total, files[i].name);
    const result = await analyzeCertidao(files[i].name, pdfParse);
    results.push(result);
  }

  results.sort(sortByCnpj);
  logger.info(`Análise de ${results.length} certidão(ões) concluída`, 'SNE');
  return { results, folderPath };
}

async function fileExists(filePath) {
  try {
    await fs.access(filePath);
    return true;
  } catch {
    return false;
  }
}

async function renameCertidoes() {
  const pdfParse = require('pdf-parse/lib/pdf-parse.js');

  const { files, folderPath } = await listCertidoes();

  // Phase 1: analyze all files
  const analyses = [];
  for (const file of files) {
    analyses.push(await analyzeCertidao(file.name, pdfParse));
  }

  // Phase 1.5: deduplicate by (cnpj, type) — keep best validity, delete others
  const cnpjTypeGroups = new Map();
  for (const analysis of analyses) {
    if (!analysis.cnpj || !analysis.type || analysis.error) continue;
    const key = `${analysis.cnpj}::${analysis.type}`;
    const list = cnpjTypeGroups.get(key) || [];
    list.push(analysis);
    cnpjTypeGroups.set(key, list);
  }

  const supersededFiles = new Set();
  for (const group of cnpjTypeGroups.values()) {
    if (group.length < 2) continue;

    const withMtime = await Promise.all(group.map(async analysis => {
      try {
        const stat = await fs.stat(path.join(SNE_CERTIDOES, analysis.filename));
        return { analysis, mtime: stat.mtime.getTime() };
      } catch {
        return { analysis, mtime: 0 };
      }
    }));

    withMtime.sort((itemA, itemB) => {
      const scoreDiff = (VALIDITY_SORT_ORDER[itemB.analysis.validity] ?? 0) - (VALIDITY_SORT_ORDER[itemA.analysis.validity] ?? 0);
      if (scoreDiff !== 0) return scoreDiff;
      const latestDateA = itemA.analysis.dates.length ? itemA.analysis.dates.reduce((dateA, dateB) => parseDate(dateA) >= parseDate(dateB) ? dateA : dateB) : null;
      const latestDateB = itemB.analysis.dates.length ? itemB.analysis.dates.reduce((dateA, dateB) => parseDate(dateA) >= parseDate(dateB) ? dateA : dateB) : null;
      if (latestDateA && latestDateB) {
        const dateDiff = parseDate(latestDateB) - parseDate(latestDateA);
        if (dateDiff !== 0) return dateDiff;
      }
      return itemB.mtime - itemA.mtime;
    });

    const [, ...losers] = withMtime;
    for (const { analysis } of losers) supersededFiles.add(analysis.filename);
  }

  // Phase 2: group by target name (files with no newName go straight to results)
  const groups = new Map(); // targetName → analysis[]
  const results = [];

  for (const analysis of analyses) {
    if (supersededFiles.has(analysis.filename)) {
      try {
        await fs.unlink(path.join(SNE_CERTIDOES, analysis.filename));
        analysis.renamed = false;
        analysis.deleted = true;
        logger.info(`"${analysis.filename}" eliminado (${analysis.type} mais antiga/vencida para CNPJ ${analysis.cnpj})`, 'SNE');
      } catch (e) {
        analysis.renamed = false;
        analysis.renameError = `Erro ao eliminar: ${e.message}`;
        logger.warn(`Erro ao eliminar "${analysis.filename}": ${e.message}`, 'SNE');
      }
      results.push(analysis);
      continue;
    }

    if (!analysis.newName) {
      analysis.renamed = false;
      analysis.renameSkipped = true;
      results.push(analysis);
      continue;
    }
    const list = groups.get(analysis.newName) || [];
    list.push(analysis);
    groups.set(analysis.newName, list);
  }

  // Phase 3: resolve each group
  for (const [targetName, group] of groups) {
    if (group.length === 1) {
      const analysis = group[0];
      if (analysis.filename === targetName) {
        analysis.renamed = false;
        analysis.noChange = true;
      } else {
        try {
          await fs.rename(path.join(SNE_CERTIDOES, analysis.filename), path.join(SNE_CERTIDOES, targetName));
          analysis.renamed = true;
        } catch (e) {
          analysis.renamed = false;
          analysis.renameError = e.message;
          logger.warn(`Erro ao renomear "${analysis.filename}": ${e.message}`, 'SNE');
        }
      }
      results.push(analysis);
      continue;
    }

    // Conflict: multiple files resolve to the same target name
    const withStats = await Promise.all(group.map(async analysis => {
      try {
        const stat = await fs.stat(path.join(SNE_CERTIDOES, analysis.filename));
        return { analysis, mtime: stat.mtime.getTime() };
      } catch {
        return { analysis, mtime: null };
      }
    }));

    const allHaveDates = withStats.every(entry => entry.mtime !== null);
    const maxMtime = allHaveDates ? Math.max(...withStats.map(entry => entry.mtime)) : null;
    const winners = allHaveDates ? withStats.filter(entry => entry.mtime === maxMtime) : [];
    const canDetermine = allHaveDates && winners.length === 1;

    if (canDetermine) {
      const winner = winners[0].analysis;

      for (const { analysis } of withStats) {
        if (analysis === winner) continue;
        try {
          await fs.unlink(path.join(SNE_CERTIDOES, analysis.filename));
          analysis.renamed = false;
          analysis.deleted = true;
          analysis.deletedKeptAs = targetName;
          logger.info(`"${analysis.filename}" eliminado (mais antigo; mantido: "${targetName}")`, 'SNE');
        } catch (e) {
          analysis.renamed = false;
          analysis.renameError = `Erro ao eliminar: ${e.message}`;
        }
        results.push(analysis);
      }

      if (winner.filename === targetName) {
        winner.renamed = false;
        winner.noChange = true;
      } else {
        try {
          await fs.rename(path.join(SNE_CERTIDOES, winner.filename), path.join(SNE_CERTIDOES, targetName));
          winner.renamed = true;
        } catch (e) {
          winner.renamed = false;
          winner.renameError = e.message;
          logger.warn(`Erro ao renomear "${winner.filename}": ${e.message}`, 'SNE');
        }
      }
      results.push(winner);
    } else {
      // Can't determine which is newer — rename all with DUPLICADO N suffix
      const baseName = targetName.replace(/\.pdf$/i, '');
      let suffixCounter = 1;

      for (const { analysis } of withStats) {
        let dupName;
        do {
          dupName = `${baseName} - DUPLICADO ${suffixCounter}.pdf`;
          suffixCounter++;
        } while (await fileExists(path.join(SNE_CERTIDOES, dupName)));

        try {
          await fs.rename(path.join(SNE_CERTIDOES, analysis.filename), path.join(SNE_CERTIDOES, dupName));
          analysis.newName = dupName;
          analysis.renamed = true;
          analysis.conflictDuplicated = true;
        } catch (e) {
          analysis.renamed = false;
          analysis.renameError = e.message;
        }
        results.push(analysis);
      }
      logger.warn(`Conflito sem resolução por data: ${group.map(analysis => `"${analysis.filename}"`).join(', ')} → renomeados como DUPLICADO`, 'SNE');
    }
  }

  results.sort(sortByCnpj);
  const renamed = results.filter(r => r.renamed).length;
  logger.info(`Renomeação concluída: ${renamed}/${results.length} arquivo(s) renomeado(s)`, 'SNE');
  return { results, folderPath };
}

// =============================================
// SNE — Empenhos
// =============================================

const SNE_EMPENHOS = path.join(SCRIPTS_PATH, 'sne', 'SNEs');
const SNE_AFS = path.join(SCRIPTS_PATH, 'sne', 'AFs');

function extractSneNumber(text) {
  // SNE number is always YYYYXXXXX (9 digits, year-first).
  // "NOTA DE EMPENHO" is ASCII-safe and always precedes the number on the same line.
  let m = text.match(/NOTA\s+DE\s+EMPENHO\s+(20\d{7})/i);
  if (m) return m[1];
  // Fallback: number appears before the keyword in some PDF column extraction orders
  m = text.match(/(20\d{7})[^\d]{0,100}EMPENHO/i);
  return m ? m[1] : null;
}

function extractAfInfo(text) {
  const m = text.match(/\bAF:\s*(\d+)\s*\/\s*(\d+)/i);
  if (!m) return null;
  return { number: m[1], year: m[2] };
}

function extractCredorFromText(rawText) {
  // Line after "credor:" header contains: COMPANY TYPE_CODE CNPJ  (normal column order)
  //                                    or: CNPJ TYPE_CODE COMPANY  (reversed — PDF reads right-to-left)
  const m = rawText.match(/credor:[^\n]*\n([^\n]+)/i);
  if (!m) return { company: null, cnpj: null };

  const line = m[1];
  const cnpjMatch = line.match(/(\d{2}\.\d{3}\.\d{3}\/\d{4}-\d{2})/);
  if (!cnpjMatch) return { company: null, cnpj: null };

  const cnpj = cnpjMatch[1].replace(/[.\-\/]/g, '');
  const cnpjIdx = line.indexOf(cnpjMatch[1]);

  // Normal order: company and type code come before the CNPJ
  const beforeCnpj = line.slice(0, cnpjIdx);
  const companyBefore = beforeCnpj.replace(/\s+\d{1,2}\s*$/, '').trim();
  if (companyBefore.length > 1) return { company: companyBefore, cnpj };

  // Reversed order: CNPJ is first, then a 2-digit type code, then the company name
  const afterCnpj = line.slice(cnpjIdx + cnpjMatch[1].length);
  const companyAfter = afterCnpj.replace(/^\d{2}/, '').trim();
  if (companyAfter.length > 1) return { company: companyAfter, cnpj };

  return { company: null, cnpj };
}

async function listEmpenhos() {
  let created = false;
  try {
    await fs.access(SNE_EMPENHOS);
  } catch {
    await fs.mkdir(SNE_EMPENHOS, { recursive: true });
    created = true;
  }

  const entries = await fs.readdir(SNE_EMPENHOS);
  const files = entries.filter(f => f.toLowerCase().endsWith('.pdf'));
  files.sort((a, b) => a.localeCompare(b));
  return { files, folderPath: SNE_EMPENHOS, folderCreated: created };
}

async function analyzeEmpenhos(onProgress) {
  const pdfParse = require('pdf-parse/lib/pdf-parse.js');
  const { files, folderPath } = await listEmpenhos();
  const total = files.length;
  const results = [];

  for (let i = 0; i < files.length; i++) {
    const filename = files[i];
    if (onProgress) onProgress(i + 1, total, filename);
    const filePath = path.join(SNE_EMPENHOS, filename);
    try {
      const buffer = await fs.readFile(filePath);
      let rawText = '';
      try {
        const pdfData = await pdfParse(buffer);
        rawText = pdfData.text || '';
      } catch (e) {
        logger.warn(`Erro ao parsear PDF "${filename}": ${e.message}`, 'SNE');
        results.push({ filename, error: 'Erro ao ler PDF' });
        continue;
      }

      const text = normalizeText(rawText);
      const sneNumber = extractSneNumber(text);
      const af = extractAfInfo(text);
      const { company, cnpj } = extractCredorFromText(rawText);
      const tombamento = /tombamento/i.test(text);

      results.push({ filename, sneNumber, af, company, cnpj, tombamento, error: null });
    } catch (e) {
      logger.error(`Erro ao processar "${filename}": ${e.message}`, 'SNE');
      results.push({ filename, error: e.message });
    }
  }

  results.sort((a, b) => {
    const afA = a.af ? `${a.af.year}${a.af.number.padStart(8, '0')}` : 'z';
    const afB = b.af ? `${b.af.year}${b.af.number.padStart(8, '0')}` : 'z';
    if (afA !== afB) return afA.localeCompare(afB);
    return (a.sneNumber || '').localeCompare(b.sneNumber || '');
  });

  logger.info(`Análise de ${results.length} empenho(s) concluída`, 'SNE');
  return { results, folderPath };
}

async function criarAFs(filenames = null, moveSnes = false) {
  const { results } = await analyzeEmpenhos();
  const toProcess = filenames ? results.filter(r => filenames.includes(r.filename)) : results;
  await fs.mkdir(SNE_AFS, { recursive: true });

  // Map CNPJ → certidão filenames from CERTIDOES folder
  const certidoesEntries = await fs.readdir(SNE_CERTIDOES).catch(() => []);
  const certidoesByCnpj = new Map();
  for (const f of certidoesEntries) {
    if (!f.toLowerCase().endsWith('.pdf')) continue;
    // After renaming, filename starts with 14-digit CNPJ
    const cnpjMatch = f.match(/^(\d{14})/);
    if (!cnpjMatch) continue;
    const cnpj = cnpjMatch[1];
    const list = certidoesByCnpj.get(cnpj) || [];
    list.push(f);
    certidoesByCnpj.set(cnpj, list);
  }

  const created = { afs: new Set(), snes: [], errors: [] };

  for (const r of toProcess) {
    if (r.error || !r.af || !r.sneNumber) continue;

    const afFolderName = sanitizeFilename(`AF ${r.af.number}-${r.af.year}`);
    const sneFolderName = sanitizeFilename(`SNE ${r.sneNumber}`);
    const afPath = path.join(SNE_AFS, afFolderName);
    const snePath = path.join(afPath, sneFolderName);

    try {
      await fs.mkdir(snePath, { recursive: true });
      created.afs.add(afFolderName);
      created.snes.push(`${afFolderName}/${sneFolderName}`);

      const srcSne = path.join(SNE_EMPENHOS, r.filename);
      const dstSne = path.join(snePath, r.filename);
      try {
        if (moveSnes) {
          await fs.rename(srcSne, dstSne).catch(async () => {
            await fs.copyFile(srcSne, dstSne);
            await fs.unlink(srcSne);
          });
        } else {
          await fs.copyFile(srcSne, dstSne);
        }
      } catch (e) {
        created.errors.push(`Erro ao ${moveSnes ? 'mover' : 'copiar'} SNE "${r.filename}": ${e.message}`);
      }

      const certFiles = r.cnpj ? (certidoesByCnpj.get(r.cnpj) || []) : [];
      for (const certFile of certFiles) {
        const src = path.join(SNE_CERTIDOES, certFile);
        const dst = path.join(snePath, certFile);
        try {
          await fs.copyFile(src, dst);
        } catch (e) {
          created.errors.push(`Erro ao copiar "${certFile}": ${e.message}`);
        }
      }
    } catch (e) {
      created.errors.push(`Erro ao criar pasta "${snePath}": ${e.message}`);
    }
  }

  logger.info(`AFs criadas: ${created.afs.size} AF(s), ${created.snes.length} SNE(s)`, 'SNE');
  return { afsPath: SNE_AFS, afs: [...created.afs], snes: created.snes, errors: created.errors };
}

async function analyzeSneFolder(snePath, pdfParse) {
  const entries = await fs.readdir(snePath, { withFileTypes: true });
  const allFiles = entries.filter(e => e.isFile()).map(e => e.name).sort((a, b) => a.localeCompare(b));
  const pdfFiles = allFiles.filter(f => f.toLowerCase().endsWith('.pdf'));

  let empenho = null;
  const certidoes = [];

  for (const filename of pdfFiles) {
    const filePath = path.join(snePath, filename);
    let rawText = '';
    try {
      const buffer = await fs.readFile(filePath);
      const pdfData = await pdfParse(buffer);
      rawText = pdfData.text || '';
    } catch (e) {
      logger.warn(`Erro ao parsear PDF "${filename}": ${e.message}`, 'SNE');
      certidoes.push({ filename, error: 'Erro ao ler PDF', type: null });
      continue;
    }

    const text = normalizeText(rawText);
    const sneNumber = extractSneNumber(text);

    if (sneNumber && !empenho) {
      const af = extractAfInfo(text);
      const { company, cnpj } = extractCredorFromText(rawText);
      const tombamento = /tombamento/i.test(text);
      empenho = { filename, sneNumber, af, company, cnpj, tombamento, error: null };
    } else {
      certidoes.push(analyzeCertidaoFromRaw(filename, rawText));
    }
  }

  return { empenho, certidoes, files: allFiles };
}

async function analyzeAFs(onProgress) {
  const pdfParse = require('pdf-parse/lib/pdf-parse.js');

  try {
    await fs.access(SNE_AFS);
  } catch {
    await fs.mkdir(SNE_AFS, { recursive: true });
    return { afs: [], folderPath: SNE_AFS };
  }

  const afEntries = await fs.readdir(SNE_AFS, { withFileTypes: true });
  const afFolders = afEntries
    .filter(e => e.isDirectory())
    .map(e => e.name)
    .sort((a, b) => a.localeCompare(b, 'pt-BR', { numeric: true }));

  const total = afFolders.length;
  const afs = [];
  for (let i = 0; i < afFolders.length; i++) {
    const afName = afFolders[i];
    if (onProgress) onProgress(i + 1, total, afName);
    const afPath = path.join(SNE_AFS, afName);
    const sneEntries = await fs.readdir(afPath, { withFileTypes: true });
    const sneFolders = sneEntries
      .filter(e => e.isDirectory())
      .map(e => e.name)
      .sort((a, b) => a.localeCompare(b, 'pt-BR', { numeric: true }));

    const snes = [];
    for (const sneName of sneFolders) {
      const snePath = path.join(afPath, sneName);
      const { empenho, certidoes, files } = await analyzeSneFolder(snePath, pdfParse);
      snes.push({ name: sneName, path: snePath, files, empenho, certidoes });
    }

    afs.push({ name: afName, path: afPath, snes });
  }

  logger.info(`AFs analisadas: ${afs.length} AF(s)`, 'SNE');
  return { afs, folderPath: SNE_AFS };
}

async function sincronizarCertidoesSne(afName, sneName, cnpj) {
  const snePath = path.join(SNE_AFS, afName, sneName);

  const sneEntries = await fs.readdir(snePath);
  const stale = sneEntries.filter(f => f.toLowerCase().endsWith('.pdf') && f.startsWith(cnpj));
  for (const f of stale) {
    await fs.unlink(path.join(snePath, f));
  }

  const certdoesEntries = await fs.readdir(SNE_CERTIDOES).catch(() => []);
  const fresh = certdoesEntries.filter(f => f.toLowerCase().endsWith('.pdf') && f.startsWith(cnpj));
  for (const f of fresh) {
    await fs.copyFile(path.join(SNE_CERTIDOES, f), path.join(snePath, f));
  }

  logger.info(`Certidões sincronizadas em "${afName}/${sneName}": ${stale.length} removida(s), ${fresh.length} copiada(s)`, 'SNE');
  return { removed: stale.length, copied: fresh.length };
}

async function listAFs() {
  try {
    await fs.access(SNE_AFS);
  } catch {
    await fs.mkdir(SNE_AFS, { recursive: true });
    return { afs: [], folderPath: SNE_AFS };
  }

  const afEntries = await fs.readdir(SNE_AFS, { withFileTypes: true });
  const afFolders = afEntries
    .filter(e => e.isDirectory())
    .map(e => e.name)
    .sort((a, b) => a.localeCompare(b, 'pt-BR', { numeric: true }));

  const afs = [];
  for (const afName of afFolders) {
    const afPath = path.join(SNE_AFS, afName);
    const sneEntries = await fs.readdir(afPath, { withFileTypes: true });
    const sneFolders = sneEntries
      .filter(e => e.isDirectory())
      .map(e => e.name)
      .sort((a, b) => a.localeCompare(b, 'pt-BR', { numeric: true }));

    const snes = [];
    for (const sneName of sneFolders) {
      const snePath = path.join(afPath, sneName);
      const fileEntries = await fs.readdir(snePath, { withFileTypes: true });
      const files = fileEntries
        .filter(e => e.isFile())
        .map(e => e.name)
        .sort((a, b) => a.localeCompare(b));
      snes.push({ name: sneName, path: snePath, files });
    }

    afs.push({ name: afName, path: afPath, snes });
  }

  logger.info(`AFs listadas: ${afs.length} AF(s)`, 'SNE');
  return { afs, folderPath: SNE_AFS };
}

module.exports = {
  SNE_CERTIDOES,
  SNE_EMPENHOS,
  SNE_AFS,
  TYPE_KEYWORDS,
  CERT_CHECKS,
  listCertidoes,
  analyzeCertidoes,
  renameCertidoes,
  listEmpenhos,
  analyzeEmpenhos,
  criarAFs,
  listAFs,
  analyzeAFs,
  sincronizarCertidoesSne,
};
