require('dotenv').config({ path: require('path').join(__dirname, '..', '.env'), quiet: true });
const express = require('express');
const WebSocket = require('ws');
const path = require('path');
const fs = require('fs').promises;
const os = require('os');
const dns = require('dns');
const { spawn, execSync } = require('child_process');
const Logger = require('./utils/logger');
const { executarMailmerge } = require('./services/mailmerge');

let logConsentEnabled = true; // null/true = salvar logs (permissivo por padrão); false = opt-out

const logger = new Logger({ minLevel: process.env.COMPRAS_LOGGER_MIN_LEVEL || 'debug' });

const { registerConsoleRoutes, handleConsoleMessage } = require('./services/consoleService');
const app = express();
const PORT = process.env.COMPRAS_PORT || 3000;
const SCRIPTS_PATH = path.resolve(path.join(__dirname, '..', 'scripts'));
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

// Middlewares
app.use(express.json({ limit: '5mb' }));
app.use(express.static(path.join(__dirname, '..', 'public')));

// Utilitário para evitar Path Traversal
function isPathSafe(targetPath, baseDir) {
  const realPath = path.resolve(path.normalize(targetPath));
  const safeBase = path.resolve(path.normalize(baseDir));
  return realPath.toLowerCase().startsWith(safeBase.toLowerCase());
}

// Dupla proteção: Nunca apagar scripts de código (.R, .js) ou o arquivo de dados de atas
function isProtectedFile(filename) {
  return filename.endsWith('.R') || filename.endsWith('.js') || filename.toLowerCase() === 'dados_atas.xlsx';
}

function matchFilePattern(filename, patterns) {
  if (!patterns || patterns.length === 0) return true;

  return patterns.some(pattern => {
    let p = pattern.trim();
    if (!p || p === '*') return true;

    // Se não tem curingas e não começa com ponto, assume que é extensão
    if (!p.includes('*') && !p.includes('?') && !p.startsWith('.')) {
      return filename.toLowerCase().endsWith('.' + p.toLowerCase());
    }

    if (p.startsWith('.')) {
      return filename.toLowerCase().endsWith(p.toLowerCase());
    }

    // Se tem curingas mas não tem ponto, trata como padrão de extensão (ex: "xls*" → "*.xls*")
    if (!p.includes('.') && !p.startsWith('*')) {
      p = '*.' + p;
    }

    const regexString = '^' + p.replace(/[.+^${}()|[\]\\]/g, '\\$&')
      .replace(/\*/g, '.*')
      .replace(/\?/g, '.') + '$';
    return new RegExp(regexString, 'i').test(filename);
  });
}

function abrirNaInterface(targetPath, isFile = false) {
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

// Retorna null com o erro já respondido via res quando o caminho é inválido.
// Caso contrário, retorna { folderPath, created } onde created indica se a pasta foi criada agora.
async function resolverPastaScript(res, scriptName, innerFolder, context) {
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

function parsearFiltros(extensions, nameContains) {
  return {
    patterns: extensions ? extensions.split(',') : null,
    filterNameContains: nameContains ? nameContains.toLowerCase().split('_') : []
  };
}

function arquivoPassaNoFiltro(file, patterns, filterNameContains) {
  if (filterNameContains.length > 0) {
    if (!filterNameContains.every(term => file.toLowerCase().includes(term.trim()))) return false;
  }
  return matchFilePattern(file, patterns);
}


// API - Open folder
app.post('/api/open-folder', async (req, res) => {
  try {
    const { folderPath } = req.body;

    if (!folderPath) {
      logger.warn('Solicitação para abrir pasta sem caminho', 'API');
      return res.status(400).json({ error: 'Caminho da pasta não fornecido' });
    }

    try {
      await fs.access(folderPath);
    } catch {
      logger.warn(`Pasta não encontrada: ${folderPath}`, 'API');
      return res.status(404).json({ error: `Pasta não encontrada: ${folderPath}` });
    }

    if (!isPathSafe(folderPath, SCRIPTS_PATH)) {
      return res.status(403).json({ error: 'Acesso negado: fora do diretório permitido' });
    }

    try {
      abrirNaInterface(folderPath);
      logger.debug(`Pasta aberta: ${folderPath}`, 'API');
    } catch (spawnError) {
      logger.error(`Erro ao abrir pasta: ${spawnError.message}`, 'API', spawnError);
      return res.status(500).json({ error: `Erro ao abrir pasta: ${spawnError.message}` });
    }

    res.json({ success: true, message: 'Pasta aberta' });
  } catch (error) {
    logger.error(`Erro geral ao abrir pasta: ${error.message}`, 'API', error);
    res.status(500).json({ error: `Erro: ${error.message}` });
  }
});

// API - Open file
app.post('/api/open-file', async (req, res) => {
  try {
    const { filePath } = req.body;

    if (!filePath) {
      logger.warn('Solicitação para abrir arquivo sem caminho', 'API');
      return res.status(400).json({ error: 'Caminho do arquivo não fornecido' });
    }

    try {
      const stats = await fs.stat(filePath);
      if (!stats.isFile()) {
        return res.status(400).json({ error: `O caminho fornecido não é um arquivo: ${filePath}` });
      }
    } catch {
      logger.warn(`Arquivo não encontrado: ${filePath}`, 'API');
      return res.status(404).json({ error: `Arquivo não encontrado: ${filePath}` });
    }

    if (!isPathSafe(filePath, SCRIPTS_PATH)) {
      return res.status(403).json({ error: 'Acesso negado: fora do diretório permitido' });
    }

    try {
      abrirNaInterface(filePath, true);
      logger.debug(`Arquivo aberto: ${filePath}`, 'API');
    } catch (spawnError) {
      logger.error(`Erro ao abrir arquivo: ${spawnError.message}`, 'API', spawnError);
      return res.status(500).json({ error: `Erro ao abrir arquivo: ${spawnError.message}` });
    }

    res.json({ success: true, message: 'Arquivo aberto' });
  } catch (error) {
    logger.error(`Erro geral ao abrir arquivo: ${error.message}`, 'API', error);
    res.status(500).json({ error: `Erro: ${error.message}` });
  }
});

// API - Check Atas Data Status
app.get('/api/check-atas-data', async (req, res) => {
  try {
    const filePath = path.join(__dirname, '..', 'scripts', 'atas', 'dados_atas.xlsx');
    const configPath = path.join(__dirname, '..', 'scripts', '_common', 'config.json');

    try {
      const stats = await fs.stat(filePath);

      let atasConfig = {};
      try {
        const configRaw = await fs.readFile(configPath, 'utf-8');
        const configJson = JSON.parse(configRaw);
        if (configJson.ATAS) {
          atasConfig = configJson.ATAS;
        }
      } catch (e) {
        // Ignora se não conseguir ler o config.json
      }

      res.json({
        exists: true,
        modified: stats.mtime,
        atasConfig
      });
    } catch (e) {
      res.json({ exists: false });
    }
  } catch (error) {
    logger.error(`Erro ao verificar dados de atas: ${error.message}`, 'API', error);
    res.status(500).json({ error: error.message });
  }
});

// API - GET list files
app.get('/api/list-files/:scriptName/:innerFolder', async (req, res) => {
  try {
    const { scriptName, innerFolder } = req.params;
    const { extensions, nameContains, sort } = req.query;

    const result = await resolverPastaScript(res, scriptName, innerFolder, 'ListFiles');
    if (!result) return;
    const { folderPath, created } = result;

    const { patterns, filterNameContains } = parsearFiltros(extensions, nameContains);
    const files = await fs.readdir(folderPath);
    const filteredFiles = files.filter(file => arquivoPassaNoFiltro(file, patterns, filterNameContains));

    const fileDetails = await Promise.all(
      filteredFiles.map(async (file) => {
        const filePath = path.join(folderPath, file);
        const stats = await fs.stat(filePath);
        return {
          name: file,
          isDirectory: stats.isDirectory(),
          size: stats.size,
          modifiedDate: stats.mtime,
        };
      })
    );

    if (sort === 'desc') {
      fileDetails.sort((a, b) => b.name.localeCompare(a.name));
    } else if (sort === 'asc') {
      fileDetails.sort((a, b) => a.name.localeCompare(b.name));
    }

    logger.debug(`Listados ${filteredFiles.length} itens de ${scriptName}/${innerFolder}`, 'ListFiles');
    const canDelete = ALLOWED_DELETE_FOLDERS.includes(innerFolder.toLowerCase());
    res.json({ files: fileDetails, folderPath, canDelete, folderCreated: created });
  } catch (error) {
    logger.error(`Erro ao listar arquivos: ${error.message}`, 'ListFiles', error);
    res.status(500).json({ error: error.message });
  }
});

// API - DELETE clear folder
app.delete('/api/clear-folder/:scriptName/:innerFolder', async (req, res) => {
  try {
    const { scriptName, innerFolder } = req.params;
    const { extensions, nameContains } = req.query;

    // Medida de segurança: Só permitir limpar pastas específicas de relatórios/saída
    if (!ALLOWED_DELETE_FOLDERS.includes(innerFolder.toLowerCase())) {
      return res.status(403).json({ error: 'A exclusão nesta pasta não é permitida por segurança.' });
    }

    const result = await resolverPastaScript(res, scriptName, innerFolder, 'ClearFolder');
    if (!result) return;
    const { folderPath } = result;

    const { patterns, filterNameContains } = parsearFiltros(extensions, nameContains);
    const files = await fs.readdir(folderPath);
    let deletedCount = 0;

    for (const file of files) {
      // Dupla proteção: Nunca apagar scripts de código ou arquivos protegidos
      if (isProtectedFile(file)) continue;
      if (!arquivoPassaNoFiltro(file, patterns, filterNameContains)) continue;

      const filePath = path.join(folderPath, file);
      const stats = await fs.stat(filePath);

      if (stats.isFile()) {
        await fs.unlink(filePath);
        deletedCount++;
      }
    }

    logger.info(`Excluído(s) ${deletedCount} arquivo(s) de ${scriptName}/${innerFolder}`, 'ClearFolder');
    res.json({
      message: `${deletedCount} arquivo(s) excluído(s) com sucesso!`,
      deletedCount
    });
  } catch (error) {
    logger.error(`Erro ao limpar pasta: ${error.message}`, 'ClearFolder', error);
    res.status(500).json({ error: error.message });
  }
});

// API - DELETE single file from allowed folder
app.delete('/api/delete-file/:scriptName/:innerFolder', async (req, res) => {
  try {
    const { scriptName, innerFolder } = req.params;
    const { fileName } = req.body;

    if (!fileName) {
      return res.status(400).json({ error: 'Nome do arquivo não fornecido' });
    }

    if (!ALLOWED_DELETE_FOLDERS.includes(innerFolder.toLowerCase())) {
      return res.status(403).json({ error: 'A exclusão nesta pasta não é permitida por segurança.' });
    }

    const result = await resolverPastaScript(res, scriptName, innerFolder, 'DeleteFile');
    if (!result) return;
    const { folderPath } = result;

    // Prevent path traversal via fileName
    if (fileName.includes('/') || fileName.includes('\\') || fileName === '..' || fileName === '.') {
      return res.status(400).json({ error: 'Nome de arquivo inválido' });
    }

    // Never delete scripts or protected files
    if (isProtectedFile(fileName)) {
      return res.status(403).json({ error: 'Este tipo de arquivo não pode ser excluído.' });
    }

    const filePath = path.join(folderPath, fileName);

    if (!isPathSafe(filePath, folderPath)) {
      return res.status(403).json({ error: 'Acesso negado' });
    }

    try { await fs.access(filePath); } catch {
      return res.status(404).json({ error: 'Arquivo não encontrado' });
    }

    const stats = await fs.stat(filePath);
    if (!stats.isFile()) {
      return res.status(400).json({ error: 'O caminho não é um arquivo' });
    }

    await fs.unlink(filePath);
    logger.info(`Arquivo excluído: ${fileName} de ${scriptName}/${innerFolder}`, 'DeleteFile');
    res.json({ success: true, message: 'Arquivo excluído com sucesso' });
  } catch (error) {
    logger.error(`Erro ao excluir arquivo: ${error.message}`, 'DeleteFile', error);
    res.status(500).json({ error: error.message });
  }
});

// API - Validate Link (proxy server-side para evitar CORS)
app.post('/api/validate-link', async (req, res) => {
  const { url } = req.body;

  if (!url) {
    return res.json({
      isValid: false,
      status: 'info',
      msg: 'Informe o link da aba LISTA FINAL e aguarde.',
    });
  }

  try {
    new URL(url);
  } catch {
    return res.json({
      isValid: false,
      status: 'error',
      msg: 'Link inválido.',
    });
  }

  try {
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), 10000);

    const response = await fetch(url, {
      signal: controller.signal,
      headers: { 'User-Agent': 'Mozilla/5.0' }
    });
    clearTimeout(timeout);

    if (!response.ok) {
      const httpMessages = {
        401: 'Sem permissão para acessar a planilha. Verifique se o link é público.',
        403: 'Acesso negado à planilha. Verifique as permissões de compartilhamento.',
        404: 'Planilha não encontrada. Verifique se o link está correto.',
        429: 'Muitas requisições ao servidor. Aguarde um momento e tente novamente.',
      };
      const msg = httpMessages[response.status]
        || (response.status >= 500 ? 'Erro no servidor do Google. Tente novamente mais tarde.' : `Erro ao acessar o link (HTTP ${response.status}).`);
      return res.json({
        isValid: false,
        status: 'error',
        msg,
      });
    }

    const htmlContent = await response.text();

    if (htmlContent.includes('LISTA FINAL')) {
      const inputValueRegex = /<input[^>]*value="([^"]+)"[^>]*>/gi;
      const inputValues = [];
      let inputMatch;
      while ((inputMatch = inputValueRegex.exec(htmlContent)) !== null) {
        if (inputMatch[1] && inputMatch[1].trim()) {
          inputValues.push(inputMatch[1].trim());
        }
      }
      const grupoMateriais = inputValues.length > 0 ? inputValues.join(', ') : 'Grupo não identificado';

      const regexSPA = /23080\.\d{6}\/\d{4}-\d{2}/g;
      const processosSPA = [...new Set(htmlContent.match(regexSPA) || [])].sort();
      // Busca "VALIDAÇÃO MANUAL" apenas dentro de células da tabela (<td>),
      // ignorando metadados, nomes de abas e filter views no HTML.
      const tdContentRegex = /<td[^>]*>([\s\S]*?)<\/td>/gi;
      let temValidacaoManual = false;
      let tdMatch;
      while ((tdMatch = tdContentRegex.exec(htmlContent)) !== null) {
        if (tdMatch[1].includes('VALIDAÇÃO MANUAL')) {
          temValidacaoManual = true;
          break;
        }
      }

      return res.json({
        isValid: true,
        status: 'success',
        msg: grupoMateriais,
        processosSPA,
        temValidacaoManual,
      });
    } else {
      return res.json({
        isValid: true,
        status: 'warning',
        msg: 'Este não parece ser um link de planilha de inserção de demandas.',
        processosSPA: [],
      });
    }

  } catch (error) {
    logger.error(`Erro ao validar link: ${error.message}`, 'ValidateLink');
    const isTimeout = error.name === 'AbortError';
    return res.json({
      isValid: false,
      status: 'error',
      msg: isTimeout
        ? 'A requisição excedeu o tempo limite. Verifique sua conexão ou tente novamente.'
        : 'Não foi possível acessar o link informado. Verifique sua conexão.',
      error: error.message,
    });
  }
});

// =============================================
// API - Fornecedores
// =============================================

const FORNECEDORES_DADOS = path.join(SCRIPTS_PATH, 'fornecedores', 'DADOS');
const FORNECEDORES_IMPORTAR = path.join(SCRIPTS_PATH, 'fornecedores', 'PARA_IMPORTAR');

// Lista pregões (pastas em DADOS) com status de processamento
app.get('/api/fornecedores/pregoes', async (_req, res) => {
  try {
    const result = await resolverPastaScript(res, 'fornecedores', 'DADOS', 'FornecedoresPregoes');
    if (!result) return;
    const { folderPath: dadosPath, created } = result;

    const entries = await fs.readdir(dadosPath, { withFileTypes: true });
    const folders = entries.filter(e => e.isDirectory()).map(e => e.name);

    // Para cada pregão, verificar se já foi processado e se há erros
    const pregoes = await Promise.all(folders.map(async (nome) => {
      const pastaPath = path.join(dadosPath, nome);
      const arquivos = await fs.readdir(pastaPath);
      const xlsxFiles = arquivos.filter(f => /^[^~].*\.xlsx?$/i.test(f) && !/_CONFERENCIA\.xlsx$/i.test(f));

      // Ler status.json da pasta do pregão
      const statusName = `PE_${nome}_STATUS.json`;
      let resultado = null;

      try {
        const statusContent = await fs.readFile(path.join(dadosPath, nome, statusName), 'utf-8');
        const statusData = JSON.parse(statusContent);
        resultado = statusData.resultado || null;
      } catch { /* sem arquivo de status */ }

      // Status para o card: 'pendente', 'sucesso', 'parcial', 'sem_saida'
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
    res.json({ pregoes, folderPath: dadosPath, folderCreated: created });
  } catch (error) {
    logger.error(`Erro ao listar pregões: ${error.message}`, 'Fornecedores', error);
    res.status(500).json({ error: error.message });
  }
});

// Lista arquivos de um pregão específico
app.get('/api/fornecedores/pregao/:pregao/arquivos', async (req, res) => {
  try {
    const { pregao } = req.params;
    const pastaPath = path.join(FORNECEDORES_DADOS, pregao);

    if (!isPathSafe(pastaPath, SCRIPTS_PATH)) {
      return res.status(403).json({ error: 'Acesso negado' });
    }

    try { await fs.access(pastaPath); } catch {
      return res.status(404).json({ error: 'Pasta do pregão não encontrada' });
    }

    const arquivos = await fs.readdir(pastaPath);
    const xlsxFiles = arquivos.filter(f => /^[^~].*\.xlsx?$/i.test(f));

    // Verificar erros por arquivo lendo o _STATUS.json da pasta do pregão
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

    const fileDetails = await Promise.all(xlsxFiles.map(async (file) => {
      const filePath = path.join(pastaPath, file);
      const stats = await fs.stat(filePath);
      const upperFileName = file.toUpperCase();
      const erro = errosPorArquivo.find(e => upperFileName === e.arquivo?.toUpperCase());

      return {
        name: file,
        size: stats.size,
        modifiedDate: stats.mtime,
        hasError: !!erro,
        errorType: erro?.tipo || null,
        fullPath: filePath
      };
    }));

    res.json({
      arquivos: fileDetails,
      folderPath: pastaPath,
      pregao,
      resultado
    });
  } catch (error) {
    logger.error(`Erro ao listar arquivos do pregão: ${error.message}`, 'Fornecedores', error);
    res.status(500).json({ error: error.message });
  }
});

// Lista arquivos para importar (PARA_IMPORTAR)
app.get('/api/fornecedores/importar', async (_req, res) => {
  try {
    const result = await resolverPastaScript(res, 'fornecedores', 'PARA_IMPORTAR', 'FornecedoresImportar');
    if (!result) return;
    const { folderPath: importarPath, created } = result;

    const arquivos = await fs.readdir(importarPath);
    const csvFiles = arquivos.filter(f => f.endsWith('.csv'));

    const fileDetails = await Promise.all(csvFiles.map(async (file) => {
      const filePath = path.join(importarPath, file);
      const stats = await fs.stat(filePath);

      // Extrair número do pregão do nome: PE_XXXXX.csv
      const match = file.match(/^PE_(.+)\.csv$/i);
      const pregao = match ? match[1] : '';

      // Verificar se há erro usando o STATUS.json da pasta do pregão
      const statusName = `PE_${pregao}_STATUS.json`;
      let hasError = null;
      try {
        const statusContent = await fs.readFile(path.join(FORNECEDORES_DADOS, pregao, statusName), 'utf-8');
        const statusData = JSON.parse(statusContent);
        hasError = statusData.resultado === 'parcial';
      } catch { /* fallback silencioso: pasta do pregão não existe */ }

      // Verificar se há arquivo de conferência
      const confName = path.join(FORNECEDORES_DADOS, pregao, `PE_${pregao}_CONFERENCIA.xlsx`);
      let hasConferencia = false;
      try {
        await fs.access(confName);
        hasConferencia = true;
      } catch { /* sem conferência */ }

      return {
        name: file,
        pregao,
        size: stats.size,
        modifiedDate: stats.mtime,
        hasError,
        hasConferencia,
        fullPath: filePath,
        conferenciaPath: hasConferencia ? confName : null
      };
    }));

    fileDetails.sort((a, b) => b.name.localeCompare(a.name));
    res.json({ arquivos: fileDetails, folderPath: importarPath, folderCreated: created });
  } catch (error) {
    logger.error(`Erro ao listar arquivos importar: ${error.message}`, 'Fornecedores', error);
    res.status(500).json({ error: error.message });
  }
});

// Deletar pasta do pregão (exclui apenas arquivos permitidos, depois remove a pasta se vazia)
app.delete('/api/fornecedores/pregao/:pregao', async (req, res) => {
  try {
    const { pregao } = req.params;

    // Mesma checagem da API genérica: só permitir exclusão em pastas autorizadas
    if (!ALLOWED_DELETE_FOLDERS.includes('dados')) {
      return res.status(403).json({ error: 'A exclusão nesta pasta não é permitida por segurança.' });
    }

    // Prevent path traversal via pregão name
    if (pregao.includes('/') || pregao.includes('\\') || pregao === '..' || pregao === '.') {
      return res.status(400).json({ error: 'Nome de pregão inválido' });
    }

    const pastaPath = path.join(FORNECEDORES_DADOS, pregao);

    if (!isPathSafe(pastaPath, SCRIPTS_PATH)) {
      return res.status(403).json({ error: 'Acesso negado' });
    }

    try { await fs.access(pastaPath); } catch {
      return res.status(404).json({ error: 'Pasta do pregão não encontrada' });
    }

    // Dupla proteção: Nunca apagar scripts de código ou arquivos protegidos
    const files = await fs.readdir(pastaPath);
    let deletedCount = 0;
    for (const file of files) {
      if (isProtectedFile(file)) continue;
      const filePath = path.join(pastaPath, file);
      const stats = await fs.stat(filePath);
      if (stats.isFile()) {
        await fs.unlink(filePath);
        deletedCount++;
      }
    }

    // Remover a pasta somente se ficou vazia
    const remaining = await fs.readdir(pastaPath);
    if (remaining.length === 0) {
      await fs.rmdir(pastaPath);
    }

    logger.info(`Pregão ${pregao}: ${deletedCount} arquivo(s) excluído(s)`, 'Fornecedores');
    res.json({ success: true, message: `Pregão ${pregao} excluído com sucesso` });
  } catch (error) {
    logger.error(`Erro ao excluir pregão: ${error.message}`, 'Fornecedores', error);
    res.status(500).json({ error: error.message });
  }
});

// Deletar arquivo individual
app.delete('/api/fornecedores/arquivo', async (req, res) => {
  try {
    const { filePath } = req.body;

    if (!filePath) {
      return res.status(400).json({ error: 'Caminho do arquivo não fornecido' });
    }

    if (!isPathSafe(filePath, SCRIPTS_PATH)) {
      return res.status(403).json({ error: 'Acesso negado' });
    }

    // Mesma checagem da API genérica: verificar se a pasta-pai está na lista permitida
    const relativePath = path.relative(SCRIPTS_PATH, path.resolve(filePath));
    const parts = relativePath.split(path.sep);
    // Estrutura esperada: fornecedores/DADOS|PARA_IMPORTAR/..., checar a subpasta (parts[1])
    const innerFolder = parts.length >= 2 ? parts[1] : '';
    if (!ALLOWED_DELETE_FOLDERS.includes(innerFolder.toLowerCase())) {
      return res.status(403).json({ error: 'A exclusão nesta pasta não é permitida por segurança.' });
    }

    // Nunca apagar scripts de código ou arquivos protegidos
    const fileName = path.basename(filePath);
    if (isProtectedFile(fileName)) {
      return res.status(403).json({ error: 'Este tipo de arquivo não pode ser excluído.' });
    }

    try { await fs.access(filePath); } catch {
      return res.status(404).json({ error: 'Arquivo não encontrado' });
    }

    const stats = await fs.stat(filePath);
    if (!stats.isFile()) {
      return res.status(400).json({ error: 'O caminho não é um arquivo' });
    }

    await fs.unlink(filePath);
    logger.info(`Arquivo excluído: ${filePath}`, 'Fornecedores');
    res.json({ success: true, message: 'Arquivo excluído com sucesso' });
  } catch (error) {
    logger.error(`Erro ao excluir arquivo: ${error.message}`, 'Fornecedores', error);
    res.status(500).json({ error: error.message });
  }
});

// Mover arquivos para pasta Documentos do usuário
app.post('/api/fornecedores/mover', async (req, res) => {
  try {
    const { pregao } = req.body;

    if (!pregao) {
      return res.status(400).json({ error: 'Número do pregão não fornecido' });
    }

    const csvName = `PE_${pregao}.csv`;
    const confName = `PE_${pregao}_CONFERENCIA.xlsx`;
    const docsPath = path.join(os.homedir(), 'Documents');

    try { await fs.access(docsPath); } catch {
      await fs.mkdir(docsPath, { recursive: true });
    }

    let movedFiles = [];

    // Mover CSV
    const csvSrc = path.join(FORNECEDORES_IMPORTAR, csvName);
    try {
      await fs.access(csvSrc);
      const csvDest = path.join(docsPath, csvName);
      await fs.copyFile(csvSrc, csvDest);
      await fs.unlink(csvSrc);
      movedFiles.push(csvName);
    } catch { /* CSV não encontrado */ }

    // Mover CONFERENCIA
    const confSrc = path.join(FORNECEDORES_IMPORTAR, confName);
    try {
      await fs.access(confSrc);
      const confDest = path.join(docsPath, confName);
      await fs.copyFile(confSrc, confDest);
      await fs.unlink(confSrc);
      movedFiles.push(confName);
    } catch { /* Conferência não encontrada */ }

    if (movedFiles.length === 0) {
      return res.status(404).json({ error: 'Nenhum arquivo encontrado para mover' });
    }

    logger.info(`Arquivos movidos para Documentos: ${movedFiles.join(', ')}`, 'Fornecedores');
    res.json({
      success: true,
      message: `${movedFiles.length} arquivo(s) movido(s) para Documentos`,
      movedFiles,
      destination: docsPath
    });
  } catch (error) {
    logger.error(`Erro ao mover arquivos: ${error.message}`, 'Fornecedores', error);
    res.status(500).json({ error: error.message });
  }
});

// API - POST consentimento de gravação de logs (decisão persiste no localStorage do cliente)
app.post('/api/log-consent', (req, res) => {
  const { consent } = req.body;
  if (typeof consent !== 'boolean') {
    return res.status(400).json({ error: 'consent deve ser boolean' });
  }

  logConsentEnabled = consent;
  if (consent) {
    logger.enableFileLogging(path.join(__dirname, '..', process.env.COMPRAS_LOG_DIR || 'logs'));
  } else {
    logger.disableFileLogging();
  }

  res.json({ ok: true });
});

// Rota principal
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, '..', 'public', 'index.html'));
});


function getComputerName() {
  return os.hostname() || 'Computador';
}

// API - GET user info
registerConsoleRoutes(app, logger);

app.get('/api/user-info', (req, res) => {
  try {
    res.json({
      computerName: getComputerName(),
      userName: process.env.USERNAME || os.userInfo().username || '',
      timestamp: new Date().toISOString(),
    });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// API - GET home background images (lidas do .env)
app.get('/api/home-backgrounds', (req, res) => {
  try {
    const raw = process.env.COMPRAS_HOME_BACKGROUND_IMAGES || '';
    const images = raw.split(',').map(u => u.trim()).filter(Boolean);
    res.json({ images });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

app.get('/api/app-config', (_req, res) => {
  res.json({
    companyName: process.env.COMPRAS_COMPANY_NAME || '',
    departmentName: process.env.COMPRAS_DEPARTMENT_NAME || '',
    companySite: process.env.COMPRAS_COMPANY_SITE || '',
    departmentSite: process.env.COMPRAS_DEPARTMENT_SITE || '',
    manualSite: process.env.COMPRAS_MANUAL_SITE || '',
    backgroundRefreshTime: parseInt(process.env.COMPRAS_HOME_BACKGROUND_REFRESH_TIME, 10) || 0,
  });
});

// API - Observatório (planilha de controle)
const { registerObservatorioRoute } = require('./services/observatorioService');
registerObservatorioRoute(app, logger);

const server = app.listen(PORT, async () => {

  const localIp = Object.values(os.networkInterfaces())
    .flat()
    .find(iface => iface.family === 'IPv4' && !iface.internal);

  logger.section('Servidor iniciado');
  logger.info(`==============================================`);
  logger.info(``);
  logger.info(`Se a página não abrir automaticamente, acessar no navegador:`);
  logger.info(`http://localhost:${PORT}`);
  if (localIp) {
    logger.info(``);
    logger.info(`Acesso pela rede local (outros computadores/celulares):`);
    logger.info(`http://${localIp.address}:${PORT}`);
  }
  logger.info(``);
  logger.info(`==============================================`);
  logger.debug(`Pasta: ${__dirname}`, 'Server');

  abrirNaInterface(`http://localhost:${PORT}`);

  function parseClientLabel(ip, userAgent) {
    const isLocal = ['::1', '127.0.0.1', '::ffff:127.0.0.1'].includes(ip);

    let osName = 'Desconhecido';
    if (/Windows/.test(userAgent)) osName = 'Windows';
    else if (/Android/.test(userAgent)) osName = 'Android';
    else if (/iPhone|iPad/.test(userAgent)) osName = 'iOS';
    else if (/Mac OS/.test(userAgent)) osName = 'macOS';
    else if (/Linux/.test(userAgent)) osName = 'Linux';

    let browser = '';
    if (/Edg\//.test(userAgent)) browser = 'Edge';
    else if (/Chrome\//.test(userAgent)) browser = 'Chrome';
    else if (/Firefox\//.test(userAgent)) browser = 'Firefox';
    else if (/Safari\//.test(userAgent) && !/Chrome/.test(userAgent)) browser = 'Safari';

    const device = browser ? `${osName}/${browser}` : osName;
    return isLocal ? `local (${device})` : `${ip.replace(/^::ffff:/, '')} (${device})`;
  }

  const wss = new WebSocket.Server({ server });
  let clientCount = 0;

  logger.section('WebSocket Server iniciado');

  const interval = setInterval(() => {
    wss.clients.forEach((ws) => {
      if (ws.isAlive === false) return ws.terminate();
      ws.isAlive = false;
      ws.ping();
    });
  }, 30000);

  wss.on('connection', (ws, req) => {
    ws.isAlive = true;
    ws.on('pong', () => { ws.isAlive = true; });

    clientCount++;

    const ip = req.socket.remoteAddress;
    const userAgent = req.headers['user-agent'] || 'Desconhecido';
    const clientLabel = parseClientLabel(ip, userAgent);
    ws.clientLabel = clientLabel;

    logger.info(`Cliente conectado: ${clientLabel} (Total: ${clientCount})`, 'WebSocket');
    logger.debug(`IP: ${ip} | Navegador: ${userAgent}`, 'WebSocket');

    // Tentar resolver hostname via DNS reverso
    const cleanIp = ip.replace(/^::ffff:/, '');
    dns.reverse(cleanIp, (err, hostnames) => {
      if (!err && hostnames && hostnames.length > 0) {
        ws.clientLabel = `${hostnames[0]} (${clientLabel})`;
        logger.debug(`Hostname resolvido: ${ws.clientLabel}`, 'WebSocket');
      }
    });

    ws.on('message', (message) => {
      try {
        const data = JSON.parse(message);
        if (data.action === 'execute-r-script' || data.action === 'execute-npm-update') {
          handleConsoleMessage(ws, data, logger, isPathSafe, logConsentEnabled);
        }
      } catch (err) {
        logger.error(`Erro ao processar mensagem: ${err.message}`, 'WebSocket', err);
      }
    });

    ws.on('close', () => {
      clientCount--;
      logger.info(`Cliente desconectado: ${ws.clientLabel} (Total: ${clientCount})`, 'WebSocket');

      if (clientCount === 0) {
        logger.info('Nenhum cliente conectado. Aguardando reconexão durante 4 segundos...', 'Server');

        // Timeout maior que o intervalo de reconexão do frontend (3s) para tolerar reloads
        setTimeout(() => {
          if (clientCount === 0) {
            server.close(() => {
              clearInterval(interval)
              logger.info('Nenhum cliente reconectado. Servidor encerrado.', 'Server');
              process.exit(0);
            });
          }
        }, 4000);
      }
    });

    ws.on('error', (error) => {
      logger.error(`Erro WebSocket: ${error.message}`, 'WebSocket', error);
    });
  });
});

