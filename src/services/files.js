const path = require('path');
const fs = require('fs').promises;
const { spawn } = require('child_process');
const Logger = require('../utils/logger');

const logger = new Logger({ minLevel: process.env.COMPRAS_LOGGER_MIN_LEVEL || 'debug' });

const SCRIPTS_PATH = path.resolve(path.join(__dirname, '..', '..', 'scripts'));

const ALLOWED_DELETE_FOLDERS = [
  'atas_finalizadas',
  'sicaf',
  'arquivos_gerados',
  'tr',
  'listas',
  'mapas',
  'arquivos_importar',
  'relatorios',
  'resumo_pedidos',
  'log',
  'temp',
  'para_importar',
  'dados'
];

function isPathSafe(targetPath, baseDir) {
  const realPath = path.resolve(path.normalize(targetPath));
  const safeBase = path.resolve(path.normalize(baseDir));
  return realPath.toLowerCase().startsWith(safeBase.toLowerCase());
}

function isProtectedFile(filename) {
  return filename.endsWith('.R') || filename.endsWith('.js') || filename.toLowerCase() === 'dados_atas.xlsx';
}

function matchFilePattern(filename, patterns) {
  if (!patterns || patterns.length === 0) return true;

  return patterns.some(pattern => {
    let p = pattern.trim();
    if (!p || p === '*') return true;

    if (!p.includes('*') && !p.includes('?') && !p.startsWith('.')) {
      return filename.toLowerCase().endsWith('.' + p.toLowerCase());
    }

    if (p.startsWith('.')) {
      return filename.toLowerCase().endsWith(p.toLowerCase());
    }

    if (!p.includes('.') && !p.startsWith('*')) {
      p = '*.' + p;
    }

    const regexString = '^' + p.replace(/[.+^${}()|[\]\\]/g, '\\$&')
      .replace(/\*/g, '.*')
      .replace(/\?/g, '.') + '$';
    return new RegExp(regexString, 'i').test(filename);
  });
}

function openInInterface(targetPath, isFile = false) {
  if (process.platform === 'win32') {
    const isUrl = /^https?:\/\//i.test(targetPath);
    const winPath = isUrl ? targetPath : path.normalize(targetPath);
    if (isFile) spawn('cmd', ['/c', 'start', '', winPath]);
    else spawn('explorer', [winPath]);
  } else if (process.platform === 'darwin') {
    spawn('open', [targetPath]);
  } else {
    spawn('xdg-open', [targetPath]);
  }
}

async function resolveScriptFolder(res, scriptName, innerFolder, context) {
  let mappedFolder = innerFolder;
  if (innerFolder === 'RAIZ') {
    mappedFolder = '.';
  }
  const folderPath = path.join(SCRIPTS_PATH, scriptName, mappedFolder);

  if (!isPathSafe(folderPath, SCRIPTS_PATH)) {
    res.status(403).json({ error: 'Acesso negado: fora do diretório permitido' });
    return null;
  }

  let created = false;
  try {
    await fs.access(folderPath);
  } catch {
    await fs.mkdir(folderPath, { recursive: true });
    logger.debug(`Pasta criada: ${folderPath}`, context);
    created = true;
  }

  return { folderPath, created };
}

function parseFilters(extensions, nameContains) {
  return {
    patterns: extensions ? extensions.split(',') : null,
    filterNameContains: nameContains ? nameContains.toLowerCase().split('_') : []
  };
}

function filePassesFilter(file, patterns, filterNameContains) {
  if (filterNameContains.length > 0) {
    if (!filterNameContains.every(term => file.toLowerCase().includes(term.trim()))) return false;
  }
  return matchFilePattern(file, patterns);
}

async function listFilesWithStats(folderPath, filterFn) {
  const entries = await fs.readdir(folderPath);
  const filtered = filterFn ? entries.filter(filterFn) : entries;
  return Promise.all(filtered.map(async (name) => {
    const stats = await fs.stat(path.join(folderPath, name));
    return { name, isDirectory: stats.isDirectory(), size: stats.size, modifiedDate: stats.mtime };
  }));
}

async function deleteFilesInFolder(folderPath, filterFn) {
  const files = await fs.readdir(folderPath);
  let deletedCount = 0;
  for (const file of files) {
    if (isProtectedFile(file)) continue;
    if (filterFn && !filterFn(file)) continue;
    const filePath = path.join(folderPath, file);
    const stats = await fs.stat(filePath);
    if (stats.isFile()) {
      await fs.unlink(filePath);
      deletedCount++;
    }
  }
  return deletedCount;
}

async function validateAndDeleteFile(filePath, baseDir) {
  if (!isPathSafe(filePath, baseDir)) {
    return { error: 'Acesso negado', statusCode: 403 };
  }
  if (isProtectedFile(path.basename(filePath))) {
    return { error: 'Este tipo de arquivo não pode ser excluído.', statusCode: 403 };
  }
  try { await fs.access(filePath); } catch {
    return { error: 'Arquivo não encontrado', statusCode: 404 };
  }
  const stats = await fs.stat(filePath);
  if (!stats.isFile()) {
    return { error: 'O caminho não é um arquivo', statusCode: 400 };
  }
  await fs.unlink(filePath);
  return { success: true };
}

// Suppliers paths
const FORNECEDORES_DADOS = path.join(SCRIPTS_PATH, 'fornecedores', 'DADOS');
const FORNECEDORES_IMPORTAR = path.join(SCRIPTS_PATH, 'fornecedores', 'PARA_IMPORTAR');

async function listPregoes(res) {
  const result = await resolveScriptFolder(res, 'fornecedores', 'DADOS', 'FornecedoresPregoes');
  if (!result) return null;
  const { folderPath: dadosPath, created } = result;

  const entries = await fs.readdir(dadosPath, { withFileTypes: true });
  const folders = entries.filter(e => e.isDirectory()).map(e => e.name);

  const pregoes = await Promise.all(folders.map(async (nome) => {
    const pastaPath = path.join(dadosPath, nome);
    const arquivos = await fs.readdir(pastaPath);
    const xlsxFiles = arquivos.filter(f => /^[^~].*\.xlsx?$/i.test(f) && !/_CONFERENCIA\.xlsx$/i.test(f));

    const statusName = `PE_${nome}_STATUS.json`;
    let resultado = null;

    try {
      const statusContent = await fs.readFile(path.join(dadosPath, nome, statusName), 'utf-8');
      const statusData = JSON.parse(statusContent);
      resultado = statusData.resultado || null;
    } catch { /* sem arquivo de status */ }

    let status = 'pendente';
    if (resultado === 'sucesso') status = 'sucesso';
    else if (resultado === 'parcial') status = 'parcial';
    else if (resultado === 'sem_saida') status = 'sem_saida';

    return {
      nome,
      qtdArquivos: xlsxFiles.length,
      resultado,
      status
    };
  }));

  pregoes.sort((a, b) => a.nome.localeCompare(b.nome));
  return { pregoes, folderPath: dadosPath, folderCreated: created };
}

async function listPregaoFiles(pregao) {
  const pastaPath = path.join(FORNECEDORES_DADOS, pregao);

  if (!isPathSafe(pastaPath, SCRIPTS_PATH)) {
    return { error: 'Acesso negado', statusCode: 403 };
  }

  try {
    await fs.access(pastaPath);
  } catch {
    return { error: 'Pasta do pregão não encontrada', statusCode: 404 };
  }

  const statusName = `PE_${pregao}_STATUS.json`;
  let errosPorArquivo = [];
  let resultado = null;
  try {
    const statusContent = await fs.readFile(path.join(FORNECEDORES_DADOS, pregao, statusName), 'utf-8');
    const statusData = JSON.parse(statusContent);
    resultado = statusData.resultado || null;
    if (Array.isArray(statusData.erros)) {
      errosPorArquivo = statusData.erros;
    }
  } catch { /* sem arquivo de status ou erro de parse */ }

  const baseFiles = await listFilesWithStats(pastaPath, f => /^[^~].*\.xlsx?$/i.test(f));
  const fileDetails = baseFiles.map(file => {
    const erro = errosPorArquivo.find(e => file.name.toUpperCase() === e.arquivo?.toUpperCase());
    return {
      ...file,
      fullPath: path.join(pastaPath, file.name),
      hasError: !!erro,
      errorType: erro?.tipo || null,
    };
  });

  return {
    arquivos: fileDetails,
    folderPath: pastaPath,
    pregao,
    resultado
  };
}

async function listImportFiles() {
  const result = await resolveScriptFolder(null, 'fornecedores', 'PARA_IMPORTAR', 'FornecedoresImportar');
  if (!result) throw new Error('Erro ao resolver pasta de importação');
  const { folderPath: importarPath, created } = result;

  const baseFiles = await listFilesWithStats(importarPath, f => f.endsWith('.csv'));

  const fileDetails = await Promise.all(baseFiles.map(async (file) => {
    const match = file.name.match(/^PE_(.+)\.csv$/i);
    const pregao = match ? match[1] : '';

    const statusName = `PE_${pregao}_STATUS.json`;
    let hasError = null;
    try {
      const statusContent = await fs.readFile(path.join(FORNECEDORES_DADOS, pregao, statusName), 'utf-8');
      const statusData = JSON.parse(statusContent);
      hasError = statusData.resultado === 'parcial';
    } catch { /* fallback silencioso */ }

    const confName = path.join(FORNECEDORES_DADOS, pregao, `PE_${pregao}_CONFERENCIA.xlsx`);
    let hasConferencia = false;
    try {
      await fs.access(confName);
      hasConferencia = true;
    } catch { /* sem conferência */ }

    return {
      ...file,
      fullPath: path.join(importarPath, file.name),
      pregao,
      hasError,
      hasConferencia,
      conferenciaPath: hasConferencia ? confName : null
    };
  }));

  fileDetails.sort((a, b) => b.name.localeCompare(a.name));
  return { arquivos: fileDetails, folderPath: importarPath, folderCreated: created };
}

async function deletePregao(pregao) {
  if (pregao.includes('/') || pregao.includes('\\') || pregao === '..' || pregao === '.') {
    return { error: 'Nome de pregão inválido', statusCode: 400 };
  }

  const pastaPath = path.join(FORNECEDORES_DADOS, pregao);

  if (!isPathSafe(pastaPath, SCRIPTS_PATH)) {
    return { error: 'Acesso negado', statusCode: 403 };
  }

  try {
    await fs.access(pastaPath);
  } catch {
    return { error: 'Pasta do pregão não encontrada', statusCode: 404 };
  }

  const deletedCount = await deleteFilesInFolder(pastaPath);

  const remaining = await fs.readdir(pastaPath);
  if (remaining.length === 0) {
    await fs.rmdir(pastaPath);
  }

  logger.info(`Pregão ${pregao}: ${deletedCount} arquivo(s) excluído(s)`, 'Fornecedores');
  return { success: true, message: `Pregão ${pregao} excluído com sucesso`, deletedCount };
}

async function deleteSupplierFile(filePath) {
  const relativePath = path.relative(SCRIPTS_PATH, path.resolve(filePath));
  const parts = relativePath.split(path.sep);
  const innerFolder = parts.length >= 2 ? parts[1] : '';
  if (!ALLOWED_DELETE_FOLDERS.includes(innerFolder.toLowerCase())) {
    return { error: 'A exclusão nesta pasta não é permitida por segurança.', statusCode: 403 };
  }

  const result = await validateAndDeleteFile(filePath, SCRIPTS_PATH);
  if (result.success) {
    logger.info(`Arquivo excluído: ${filePath}`, 'Fornecedores');
  }
  return result;
}

async function moveSupplierFiles(pregao, os) {
  const csvName = `PE_${pregao}.csv`;
  const confName = `PE_${pregao}_CONFERENCIA.xlsx`;
  const docsPath = path.join(os.homedir(), 'Documents');

  try {
    await fs.access(docsPath);
  } catch {
    await fs.mkdir(docsPath, { recursive: true });
  }

  let movedFiles = [];

  const csvSrc = path.join(FORNECEDORES_IMPORTAR, csvName);
  try {
    await fs.access(csvSrc);
    const csvDest = path.join(docsPath, csvName);
    await fs.copyFile(csvSrc, csvDest);
    await fs.unlink(csvSrc);
    movedFiles.push(csvName);
  } catch { /* CSV não encontrado */ }

  const confSrc = path.join(FORNECEDORES_IMPORTAR, confName);
  try {
    await fs.access(confSrc);
    const confDest = path.join(docsPath, confName);
    await fs.copyFile(confSrc, confDest);
    await fs.unlink(confSrc);
    movedFiles.push(confName);
  } catch { /* Conferência não encontrada */ }

  if (movedFiles.length === 0) {
    return { error: 'Nenhum arquivo encontrado para mover', statusCode: 404 };
  }

  logger.info(`Arquivos movidos para Documentos: ${movedFiles.join(', ')}`, 'Fornecedores');
  return {
    success: true,
    message: `${movedFiles.length} arquivo(s) movido(s) para Documentos`,
    movedFiles,
    destination: docsPath
  };
}

module.exports = {
  SCRIPTS_PATH,
  ALLOWED_DELETE_FOLDERS,
  FORNECEDORES_DADOS,
  FORNECEDORES_IMPORTAR,
  isPathSafe,
  isProtectedFile,
  matchFilePattern,
  openInInterface,
  resolveScriptFolder,
  parseFilters,
  filePassesFilter,
  listFilesWithStats,
  deleteFilesInFolder,
  validateAndDeleteFile,
  listPregoes,
  listPregaoFiles,
  listImportFiles,
  deletePregao,
  deleteSupplierFile,
  moveSupplierFiles,
};
