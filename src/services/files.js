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

module.exports = {
  SCRIPTS_PATH,
  ALLOWED_DELETE_FOLDERS,
  isPathSafe,
  isProtectedFile,
  matchFilePattern,
  openInInterface,
  resolveScriptFolder,
  parseFilters,
  filePassesFilter,
};
