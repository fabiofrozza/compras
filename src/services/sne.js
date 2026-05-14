const path = require('path');
const fs = require('fs').promises;
const Logger = require('../utils/logger');
const { SCRIPTS_PATH } = require('../utils/files');

const logger = new Logger({ minLevel: process.env.COMPRAS_LOGGER_MIN_LEVEL || 'debug' });

const SNE_CERTIDOES = path.join(SCRIPTS_PATH, 'sne', 'CERTIDOES');

const VALIDITY_WARN_DAYS = 5;

// Maps filename keywords (priority order) to canonical cert type names
const TYPE_KEYWORDS = [
  { keywords: ['relatório de credenciamento'], type: 'Credenciamento' },
  { keywords: ['registrada no sicaf', 'sicaf'], type: 'SICAF' },
  { keywords: ['receita federal', 'pgfn'], type: 'Receita Federal' },
  { keywords: ['fgts', 'cef'], type: 'FGTS' },
  { keywords: ['trabalhista', 'tst'], type: 'Trabalhista' },
  { keywords: ['receita estadual', 'estadual', 'distrital'], type: 'Estadual' },
  { keywords: ['receita municipal', 'municipal', 'iss'], type: 'Municipal' },
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
    dates: [
      { label: 'Estadual', pattern: /[Vv][áa]lidade.*?(\d{2}[\/.\-]\d{2}[\/.\-]\d{4})/i },
      { label: 'Estadual', pattern: /[Vv]alidade[^0-9]*(\d{2}[\/.\-]\d{2}[\/.\-]\d{4})/i },
    ],
  },
  municipal: {
    dates: [
      { label: 'Municipal', pattern: /[Vv][áa]lidade.*?\s*(\d{2}[\/.\-]\d{2}[\/.\-]\d{4})/i },
      { label: 'Municipal', pattern: /[Vv]alidade[^0-9]*(\d{2}[\/.\-]\d{2}[\/.\-]\d{4})/i },
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
  ];
  for (const p of patterns) {
    const m = rawText.match(p);
    if (m) {
      const name = m[1].replace(/\s+/g, ' ').trim();
      if (name.length > 1) return name;
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
  return dates.reduce((a, b) => parseDate(a) <= parseDate(b) ? a : b);
}

const VALIDITY_ORDER = { VALIDA: 3, A_VENCER: 2, VENCIDA: 1, SEM_VALIDADE: 0 };

// Remove chars forbidden in Windows filenames
function sanitizeFilename(str) {
  return str.replace(/[<>:"/\\|?*\n\r\t]/g, '').trim();
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

  const text = normalizeText(rawText);
  const type = detectType(text);
  if (!type) {
    return { filename, error: 'Tipo de certidão não identificado', newName: null, type: null };
  }
  const cnpj = extractCnpj(text);
  const rawCompany = extractCompanyName(rawText); // raw text preserves newlines as delimiters
  const company = rawCompany && rawCompany.length > 25
    ? rawCompany.slice(0, 25).trimEnd() + '[...]'
    : rawCompany;
  const typeKey = type.toLowerCase();
  const checks = CERT_CHECKS[typeKey] || {};

  let impedido = false;
  const suffixes = [];

  if (checks.impedimento) {
    const m = text.match(checks.impedimento.pattern);
    if (m && !checks.impedimento.expected.test(m[1].trim())) {
      suffixes.push(checks.impedimento.failSuffix);
      impedido = true;
    }
  }

  const foundDates = [];
  const componentDates = {}; // label → first matching dateStr (for multi-component certs like SICAF)
  if (type !== 'Credenciamento' && checks.dates) {
    for (const rule of checks.dates) {
      const m = text.match(rule.pattern);
      if (m) {
        foundDates.push(m[1]);
        if (rule.label && !(rule.label in componentDates)) componentDates[rule.label] = m[1];
        if (rule.firstMatch) break; // stop after first hit (e.g. FGTS range vs single)
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

  logger.debug(`"${filename}": type=${type}, cnpj=${cnpj ?? 'null'}, company="${company ?? 'null'}", dates=[${foundDates.join(', ')}]`, 'SNE');

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
    impedido,
    warnings,
    error: null,
  };
}

function sortByCnpj(a, b) {
  if (a.cnpj && b.cnpj) {
    return a.cnpj !== b.cnpj
      ? a.cnpj.localeCompare(b.cnpj)
      : (a.type || '').localeCompare(b.type || '');
  }
  if (a.cnpj) return -1;
  if (b.cnpj) return 1;
  return a.filename.localeCompare(b.filename);
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
    .filter(f => f.toLowerCase().endsWith('.pdf'))
    .map(f => ({ name: f, type: detectType(f) }));

  files.sort((a, b) => {
    const ca = (a.name.match(/\d{14}/) || [''])[0];
    const cb = (b.name.match(/\d{14}/) || [''])[0];
    if (ca && cb) return ca !== cb ? ca.localeCompare(cb) : a.name.localeCompare(b.name);
    if (ca) return -1;
    if (cb) return 1;
    return a.name.localeCompare(b.name);
  });

  return { files, folderPath: SNE_CERTIDOES, folderCreated: created };
}

async function analyzeCertidoes() {
  // Use lib path to avoid test-file loading at module init (pdf-parse v1 quirk)
  const pdfParse = require('pdf-parse/lib/pdf-parse.js');

  const { files, folderPath } = await listCertidoes();
  const results = [];

  for (const file of files) {
    const result = await analyzeCertidao(file.name, pdfParse);
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
  for (const a of analyses) {
    if (!a.cnpj || !a.type || a.error) continue;
    const key = `${a.cnpj}::${a.type}`;
    const list = cnpjTypeGroups.get(key) || [];
    list.push(a);
    cnpjTypeGroups.set(key, list);
  }

  const supersededFiles = new Set();
  for (const group of cnpjTypeGroups.values()) {
    if (group.length < 2) continue;

    const withMtime = await Promise.all(group.map(async a => {
      try {
        const stat = await fs.stat(path.join(SNE_CERTIDOES, a.filename));
        return { a, mtime: stat.mtime.getTime() };
      } catch {
        return { a, mtime: 0 };
      }
    }));

    withMtime.sort((x, y) => {
      const scoreDiff = (VALIDITY_ORDER[y.a.validity] ?? 0) - (VALIDITY_ORDER[x.a.validity] ?? 0);
      if (scoreDiff !== 0) return scoreDiff;
      const xDate = x.a.dates.length ? x.a.dates.reduce((a, b) => parseDate(a) >= parseDate(b) ? a : b) : null;
      const yDate = y.a.dates.length ? y.a.dates.reduce((a, b) => parseDate(a) >= parseDate(b) ? a : b) : null;
      if (xDate && yDate) {
        const dateDiff = parseDate(yDate) - parseDate(xDate);
        if (dateDiff !== 0) return dateDiff;
      }
      return y.mtime - x.mtime;
    });

    const [, ...losers] = withMtime;
    for (const { a } of losers) supersededFiles.add(a.filename);
  }

  // Phase 2: group by target name (files with no newName go straight to results)
  const groups = new Map(); // targetName → analysis[]
  const results = [];

  for (const a of analyses) {
    if (supersededFiles.has(a.filename)) {
      try {
        await fs.unlink(path.join(SNE_CERTIDOES, a.filename));
        a.renamed = false;
        a.deleted = true;
        logger.info(`"${a.filename}" eliminado (${a.type} mais antiga/vencida para CNPJ ${a.cnpj})`, 'SNE');
      } catch (e) {
        a.renamed = false;
        a.renameError = `Erro ao eliminar: ${e.message}`;
        logger.warn(`Erro ao eliminar "${a.filename}": ${e.message}`, 'SNE');
      }
      results.push(a);
      continue;
    }

    if (!a.newName) {
      a.renamed = false;
      a.renameSkipped = true;
      logger.debug(`"${a.filename}": sem nome gerado (type=${a.type ?? 'null'}, cnpj=${a.cnpj ?? 'null'}, company="${a.company ?? 'null'}") — pulado`, 'SNE');
      results.push(a);
      continue;
    }
    const list = groups.get(a.newName) || [];
    list.push(a);
    groups.set(a.newName, list);
  }

  // Phase 3: resolve each group
  for (const [targetName, group] of groups) {
    if (group.length === 1) {
      const a = group[0];
      if (a.filename === targetName) {
        a.renamed = false;
        a.noChange = true;
      } else {
        try {
          await fs.rename(path.join(SNE_CERTIDOES, a.filename), path.join(SNE_CERTIDOES, targetName));
          a.renamed = true;
          logger.debug(`"${a.filename}" → "${targetName}"`, 'SNE');
        } catch (e) {
          a.renamed = false;
          a.renameError = e.message;
          logger.warn(`Erro ao renomear "${a.filename}": ${e.message}`, 'SNE');
        }
      }
      results.push(a);
      continue;
    }

    // Conflict: multiple files resolve to the same target name
    const withStats = await Promise.all(group.map(async a => {
      try {
        const stat = await fs.stat(path.join(SNE_CERTIDOES, a.filename));
        return { a, mtime: stat.mtime.getTime() };
      } catch {
        return { a, mtime: null };
      }
    }));

    const allHaveDates = withStats.every(e => e.mtime !== null);
    const maxMtime = allHaveDates ? Math.max(...withStats.map(e => e.mtime)) : null;
    const winners = allHaveDates ? withStats.filter(e => e.mtime === maxMtime) : [];
    const canDetermine = allHaveDates && winners.length === 1;

    if (canDetermine) {
      const winner = winners[0].a;

      for (const { a } of withStats) {
        if (a === winner) continue;
        try {
          await fs.unlink(path.join(SNE_CERTIDOES, a.filename));
          a.renamed = false;
          a.deleted = true;
          a.deletedKeptAs = targetName;
          logger.info(`"${a.filename}" eliminado (mais antigo; mantido: "${targetName}")`, 'SNE');
        } catch (e) {
          a.renamed = false;
          a.renameError = `Erro ao eliminar: ${e.message}`;
        }
        results.push(a);
      }

      if (winner.filename === targetName) {
        winner.renamed = false;
        winner.noChange = true;
      } else {
        try {
          await fs.rename(path.join(SNE_CERTIDOES, winner.filename), path.join(SNE_CERTIDOES, targetName));
          winner.renamed = true;
          logger.debug(`"${winner.filename}" → "${targetName}"`, 'SNE');
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

      for (const { a } of withStats) {
        let dupName;
        do {
          dupName = `${baseName} - DUPLICADO ${suffixCounter}.pdf`;
          suffixCounter++;
        } while (await fileExists(path.join(SNE_CERTIDOES, dupName)));

        try {
          await fs.rename(path.join(SNE_CERTIDOES, a.filename), path.join(SNE_CERTIDOES, dupName));
          a.newName = dupName;
          a.renamed = true;
          a.conflictDuplicated = true;
        } catch (e) {
          a.renamed = false;
          a.renameError = e.message;
        }
        results.push(a);
      }
      logger.warn(`Conflito sem resolução por data: ${group.map(a => `"${a.filename}"`).join(', ')} → renomeados como DUPLICADO`, 'SNE');
    }
  }

  results.sort(sortByCnpj);
  const renamed = results.filter(r => r.renamed).length;
  logger.info(`Renomeação concluída: ${renamed}/${results.length} arquivo(s) renomeado(s)`, 'SNE');
  return { results, folderPath };
}

module.exports = {
  SNE_CERTIDOES,
  TYPE_KEYWORDS,
  CERT_CHECKS,
  listCertidoes,
  analyzeCertidoes,
  renameCertidoes,
};
