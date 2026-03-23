require('dotenv').config({ path: require('path').join(__dirname, '..', '.env'), quiet: true });
const express = require('express');
const WebSocket = require('ws');
const path = require('path');
const fs = require('fs').promises;
const os = require('os');
const { spawn, execSync } = require('child_process');
const Logger = require('./utils/logger');
const { executarMailmerge } = require('./services/mailmerge');

const logger = new Logger({
  minLevel: 'debug',
  logDir: path.join(__dirname, '..', 'logs')
});

const app = express();
const PORT = process.env.PORT || 3000;
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
  'para_importar'
];
let cachedRScriptPath = null;

// Middlewares
app.use(express.json({ limit: '5mb' }));
app.use(express.static(path.join(__dirname, '..', 'public')));

// Utilitário para evitar Path Traversal
function isPathSafe(targetPath, baseDir) {
  const realPath = path.resolve(path.normalize(targetPath));
  const safeBase = path.resolve(path.normalize(baseDir));
  return realPath.toLowerCase().startsWith(safeBase.toLowerCase());
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

// Retorna null com o erro já respondido via res quando o caminho é inválido ou inexistente.
async function resolverPastaScript(res, scriptName, innerFolder, context) {
  const folderPath = path.join(SCRIPTS_PATH, scriptName, innerFolder);

  if (!isPathSafe(folderPath, SCRIPTS_PATH)) {
    res.status(403).json({ error: 'Acesso negado: fora do diretório permitido' });
    return null;
  }

  try {
    await fs.access(folderPath);
  } catch {
    logger.debug(`Pasta não encontrada: ${folderPath}`, context);
    res.status(404).json({ error: 'Pasta não encontrada' });
    return null;
  }

  return folderPath;
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

// Compara versões no formato "x.y.z" — retorna negativo, zero ou positivo
function compareVersions(a, b) {
  const pa = a.split('.').map(Number);
  const pb = b.split('.').map(Number);
  for (let i = 0; i < Math.max(pa.length, pb.length); i++) {
    const diff = (pa[i] || 0) - (pb[i] || 0);
    if (diff !== 0) return diff;
  }
  return 0;
}

async function getRScriptPath() {
  if (cachedRScriptPath) return cachedRScriptPath;

  if (process.platform === 'win32') {
    // Registro do Windows primeiro (HKLM e HKCU), versões ordenadas da mais nova para a mais antiga
    const hives = ['HKLM\\SOFTWARE\\R-core\\R', 'HKCU\\SOFTWARE\\R-core\\R'];
    for (const hive of hives) {
      try {
        const output = execSync(`reg query "${hive}"`, { encoding: 'utf8', stdio: ['pipe', 'pipe', 'ignore'] });
        const versoes = output.split('\n')
          .map(l => l.trim())
          .filter(l => l.toLowerCase().startsWith(hive.toLowerCase() + '\\'))
          .map(l => l.slice(hive.length + 1).trim())
          .filter(v => v && !v.includes('\\'));

        versoes.sort((a, b) => compareVersions(b, a));

        for (const versao of versoes) {
          try {
            const instOutput = execSync(`reg query "${hive}\\${versao}" /v InstallPath`, { encoding: 'utf8', stdio: ['pipe', 'pipe', 'ignore'] });
            const match = instOutput.match(/InstallPath\s+REG_SZ\s+(.+)/);
            if (match) {
              const rscriptPath = path.join(match[1].trim(), 'bin', 'Rscript.exe');
              try {
                await fs.access(rscriptPath);
                cachedRScriptPath = rscriptPath;
                return cachedRScriptPath;
              } catch { }
            }
          } catch { }
        }
      } catch { }
    }

    // Pastas comuns de instalação
    const localAppData = process.env['LOCALAPPDATA'] || '';
    const programFiles = process.env['ProgramFiles'] || 'C:\\Program Files';
    const programFilesX86 = process.env['ProgramFiles(x86)'] || '';
    const rDirs = [
      localAppData && path.join(localAppData, 'Programs', 'R'),
      path.join(programFiles, 'R'),
      programFilesX86 && path.join(programFilesX86, 'R'),
      'C:\\R',
    ].filter(Boolean);

    for (const rDir of rDirs) {
      try {
        await fs.access(rDir);
        const entries = await fs.readdir(rDir);
        const versoes = entries.filter(f => f.startsWith('R-'));
        versoes.sort((a, b) => compareVersions(b.slice(2), a.slice(2)));

        for (const versao of versoes) {
          const rscriptPath = path.join(rDir, versao, 'bin', 'Rscript.exe');
          try {
            await fs.access(rscriptPath);
            cachedRScriptPath = rscriptPath;
            return cachedRScriptPath;
          } catch { }
        }
      } catch { }
    }

    // Entradas do PATH com padrão de versão do R (ex: R-4.x.x\bin), ordenadas pela mais recente
    const systemPath = process.env['PATH'] || '';
    const pathCandidates = [];
    for (const dir of systemPath.split(';')) {
      const vMatch = dir.match(/[\\\/]R-([\d.]+)[\\\/]bin$/i);
      if (vMatch) pathCandidates.push({ dir, version: vMatch[1] });
    }
    pathCandidates.sort((a, b) => compareVersions(b.version, a.version));
    for (const { dir } of pathCandidates) {
      const rscriptPath = path.join(dir, 'Rscript.exe');
      try {
        await fs.access(rscriptPath);
        cachedRScriptPath = rscriptPath;
        return cachedRScriptPath;
      } catch { }
    }

    // Fallback: where Rscript (pega o que estiver no PATH, verifica existência)
    try {
      const lines = execSync('where Rscript', { encoding: 'utf8', stdio: ['pipe', 'pipe', 'ignore'] }).split('\n').map(l => l.trim()).filter(Boolean);
      for (const found of lines) {
        try {
          await fs.access(found);
          cachedRScriptPath = found;
          return cachedRScriptPath;
        } catch { }
      }
    } catch { }
  }

  return null;
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

    try {
      const stats = await fs.stat(filePath);
      res.json({
        exists: true,
        modified: stats.mtime
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

    const folderPath = await resolverPastaScript(res, scriptName, innerFolder, 'ListFiles');
    if (!folderPath) return;

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
    res.json({ files: fileDetails, folderPath, canDelete });
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

    const folderPath = await resolverPastaScript(res, scriptName, innerFolder, 'ClearFolder');
    if (!folderPath) return;

    const { patterns, filterNameContains } = parsearFiltros(extensions, nameContains);
    const files = await fs.readdir(folderPath);
    let deletedCount = 0;

    for (const file of files) {
      // Dupla proteção: Nunca apagar scripts de código ou o próprio excel
      if (file.endsWith('.R') || file.endsWith('.js') || file.toLowerCase() === 'dados_atas.xlsx') continue;
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

    const folderPath = await resolverPastaScript(res, scriptName, innerFolder, 'DeleteFile');
    if (!folderPath) return;

    // Prevent path traversal via fileName
    if (fileName.includes('/') || fileName.includes('\\') || fileName === '..' || fileName === '.') {
      return res.status(400).json({ error: 'Nome de arquivo inválido' });
    }

    // Never delete scripts or protected files
    if (fileName.endsWith('.R') || fileName.endsWith('.js') || fileName.toLowerCase() === 'dados_atas.xlsx') {
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

      return res.json({
        isValid: true,
        status: 'success',
        msg: grupoMateriais,
        processosSPA,
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
    try { await fs.access(FORNECEDORES_DADOS); } catch {
      await fs.mkdir(FORNECEDORES_DADOS, { recursive: true });
    }

    const entries = await fs.readdir(FORNECEDORES_DADOS, { withFileTypes: true });
    const folders = entries.filter(e => e.isDirectory()).map(e => e.name);

    // Para cada pregão, verificar se já foi processado e se há erros
    const pregoes = await Promise.all(folders.map(async (nome) => {
      const pastaPath = path.join(FORNECEDORES_DADOS, nome);
      const arquivos = await fs.readdir(pastaPath);
      const xlsxFiles = arquivos.filter(f => /^[^~].*\.xlsx?$/i.test(f));

      // Ler status.json da pasta do pregão
      const statusName = `PE_${nome}_STATUS.json`;
      let resultado = null;

      try {
        const statusContent = await fs.readFile(path.join(FORNECEDORES_DADOS, nome, statusName), 'utf-8');
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

    pregoes.sort((a, b) => b.nome.localeCompare(a.nome));
    res.json({ pregoes, folderPath: FORNECEDORES_DADOS });
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
    try { await fs.access(FORNECEDORES_IMPORTAR); } catch {
      await fs.mkdir(FORNECEDORES_IMPORTAR, { recursive: true });
    }

    const arquivos = await fs.readdir(FORNECEDORES_IMPORTAR);
    const csvFiles = arquivos.filter(f => f.endsWith('.csv'));

    const fileDetails = await Promise.all(csvFiles.map(async (file) => {
      const filePath = path.join(FORNECEDORES_IMPORTAR, file);
      const stats = await fs.stat(filePath);

      // Extrair número do pregão do nome: PE_XXXXX.csv
      const match = file.match(/^PE_(.+)\.csv$/i);
      const pregao = match ? match[1] : '';

      // Verificar se há erro usando o STATUS.json da pasta do pregão
      const statusName = `PE_${pregao}_STATUS.json`;
      let hasError = false;
      try {
        const statusContent = await fs.readFile(path.join(FORNECEDORES_DADOS, pregao, statusName), 'utf-8');
        const statusData = JSON.parse(statusContent);
        hasError = statusData.resultado === 'parcial';
      } catch { /* fallback silencioso */ }

      // Verificar se há arquivo de conferência
      const confName = `PE_${pregao}_CONFERENCIA.xlsx`;
      let hasConferencia = false;
      try {
        await fs.access(path.join(FORNECEDORES_IMPORTAR, confName));
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
        conferenciaPath: hasConferencia ? path.join(FORNECEDORES_IMPORTAR, confName) : null
      };
    }));

    fileDetails.sort((a, b) => b.name.localeCompare(a.name));
    res.json({ arquivos: fileDetails, folderPath: FORNECEDORES_IMPORTAR });
  } catch (error) {
    logger.error(`Erro ao listar arquivos importar: ${error.message}`, 'Fornecedores', error);
    res.status(500).json({ error: error.message });
  }
});

// Deletar pasta do pregão
app.delete('/api/fornecedores/pregao/:pregao', async (req, res) => {
  try {
    const { pregao } = req.params;
    const pastaPath = path.join(FORNECEDORES_DADOS, pregao);

    if (!isPathSafe(pastaPath, SCRIPTS_PATH)) {
      return res.status(403).json({ error: 'Acesso negado' });
    }

    try { await fs.access(pastaPath); } catch {
      return res.status(404).json({ error: 'Pasta do pregão não encontrada' });
    }

    await fs.rm(pastaPath, { recursive: true, force: true });
    logger.info(`Pasta do pregão ${pregao} excluída`, 'Fornecedores');
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

    try { await fs.access(filePath); } catch {
      return res.status(404).json({ error: 'Arquivo não encontrado' });
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

// Rota principal
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, '..', 'public', 'index.html'));
});

app.get('/api/check-r', async (_req, res) => {
  cachedRScriptPath = null;
  const rscriptCmd = await getRScriptPath();

  if (!rscriptCmd) {
    return res.json({ installed: false, message: 'R não encontrado. Por favor, instale o R em https://cran.r-project.org/bin/windows/base/' });
  }

  const rCheck = spawn(rscriptCmd, ['--version']);

  let output = '';
  let responded = false;
  rCheck.stdout.on('data', (data) => { output += data.toString(); });
  rCheck.stderr.on('data', (data) => { output += data.toString(); });

  rCheck.on('close', (code) => {
    if (responded) return;
    responded = true;
    if (code === 0 || output.includes('R scripting')) {
      res.json({ installed: true, version: output.trim(), path: rscriptCmd });
    } else {
      res.json({ installed: false, message: 'R não encontrado. Por favor, instale o R em https://cran.r-project.org/bin/windows/base/' });
    }
  });

  rCheck.on('error', () => {
    if (responded) return;
    responded = true;
    res.json({ installed: false, message: 'R não encontrado. Por favor, instale o R em https://cran.r-project.org/bin/windows/base/' });
  });
});

app.get('/api/check-r-latest', async (_req, res) => {
  try {
    const response = await fetch('https://cran.r-project.org/bin/windows/base/');
    const html = await response.text();

    // The page heading is "R-X.X.X for Windows"
    const versionMatch = html.match(/R-(\d+\.\d+\.\d+)\s+for\s+Windows/);
    // The download link is "R-X.X.X-win.exe"
    const exeMatch = html.match(/href=["']?(R-[\d.]+-win\.exe)["']?/i);

    if (versionMatch) {
      const latestVersion = versionMatch[1];
      const exeFile = exeMatch ? exeMatch[1] : `R-${latestVersion}-win.exe`;
      const downloadUrl = `https://cran.r-project.org/bin/windows/base/${exeFile}`;
      res.json({ latest: latestVersion, downloadUrl });
    } else {
      res.json({ error: 'Não foi possível identificar a versão mais recente.' });
    }
  } catch (err) {
    logger.error(`Erro ao consultar versão do R no CRAN: ${err.message}`, 'RLatest');
    res.json({ error: 'Não foi possível consultar o CRAN.' });
  }
});

function getComputerName() {
  return os.hostname() || 'Computador';
}

// API - GET user info
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
    const raw = process.env.HOME_BACKGROUND_IMAGES || '';
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
  });
});

const server = app.listen(PORT, async () => {

  logger.section('Servidor iniciado');
  logger.info(`==============================================`);
  logger.info(``);
  logger.info(`Se a página não abrir automaticamente, acessar no navegador:`);
  logger.info(`http://localhost:${PORT}`);
  logger.info(``);
  logger.info(`==============================================`);
  logger.debug(`Pasta: ${__dirname}`, 'Server');

  abrirNaInterface(`http://localhost:${PORT}`);

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

    logger.info(`Cliente conectado (Total: ${clientCount})`, 'WebSocket');
    logger.debug(`IP: ${ip} | Navegador: ${userAgent}`, 'WebSocket');

    ws.on('message', (message) => {
      try {
        const data = JSON.parse(message);
        if (data.action === 'execute-r-script') {
          if (data.scriptName === 'atas_mailmerge') {
            logger.section(`Executando mailmerge de atas (JavaScript): ${data.scriptName}`);
            executeMailmergeJS(ws, data.params);
          } else {
            logger.section(`Executando script R: ${data.scriptName}`);
            executeRScript(ws, data.scriptName, data.params);
          }
        }
      } catch (err) {
        logger.error(`Erro ao processar mensagem: ${err.message}`, 'WebSocket', err);
      }
    });

    ws.on('close', () => {
      clientCount--;
      logger.info(`Cliente desconectado (Total: ${clientCount})`, 'WebSocket');
      logger.debug(`IP: ${ip}`, 'WebSocket');

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

async function executeMailmergeJS(ws, params) {
  try {
    ws.send(JSON.stringify({
      type: 'start',
      message: 'Iniciando mailmerge de atas...'
    }));

    const sendLog = (message, level = 'info') => {
      ws.send(JSON.stringify({
        type: 'output',
        message: message,
        level: level
      }));
    };

    const resultado = await executarMailmerge(params, sendLog);

    const mailmergeType = (resultado.status === 'success' || resultado.status === 'warning') ? resultado.status : 'error';
    const notificationMessage = buildNotificationMessage('atas_mailmerge', mailmergeType, null, resultado.message);

    if (mailmergeType === 'error') {
      logger.warn(notificationMessage, 'Mailmerge');
    } else {
      logger.success(notificationMessage, 'Mailmerge');
    }

    ws.send(JSON.stringify({
      type: mailmergeType,
      message: resultado.message,
      notificationMessage,
      scriptName: 'atas_mailmerge',
      detalhes: resultado
    }));

  } catch (err) {
    const errorMsg = `Erro ao executar mailmerge: ${err.message}`;
    const notificationMessage = buildNotificationMessage('atas_mailmerge', 'error', null, errorMsg);
    logger.error(notificationMessage, 'Mailmerge', err);
    ws.send(JSON.stringify({
      type: 'error',
      message: errorMsg,
      notificationMessage,
      scriptName: 'atas_mailmerge'
    }));
  }
}

async function executeRScript(ws, scriptFolder, params) {
  const scriptName = scriptFolder + '.R';
  const scriptPath = path.join(__dirname, '..', 'scripts', scriptFolder, scriptName);

  const scriptsDir = path.resolve(path.join(__dirname, '..', 'scripts'));
  const workingDir = path.resolve(path.join(scriptsDir, scriptFolder));

  if (!isPathSafe(workingDir, scriptsDir)) {
    logger.error(`Tentativa de Path Traversal bloqueada: ${scriptFolder}`, 'Security');
    ws.send(JSON.stringify({ type: 'error', message: 'Acesso negado ao diretório.' }));
    return;
  }

  try {
    await fs.access(scriptPath);
  } catch (error) {
    logger.warn(`Script R não encontrado: ${scriptPath}`, 'RScript');
    ws.send(JSON.stringify({
      type: 'error',
      scriptName: scriptFolder,
      message: 'Script R não encontrado: ' + scriptPath
    }));
    return;
  }

  const rscriptCmd = await getRScriptPath();

  if (!rscriptCmd) {
    const notFoundMsg = 'R não encontrado. Instale o R pela aba Instalação antes de executar scripts.';
    logger.error(notFoundMsg, 'RDetect');
    ws.send(JSON.stringify({
      type: 'error',
      scriptName: scriptFolder,
      message: notFoundMsg,
      notificationMessage: `${TAB_NAMES[scriptFolder] || scriptFolder}: ${notFoundMsg}`
    }));
    return;
  }

  const args = [scriptPath, ...Object.values(params), 'json-output'];

  const spawnOptions = {
    cwd: workingDir,
    windowsHide: true,
  };

  ws.send(JSON.stringify({
    type: 'start',
    message: 'Iniciando execução do script R...'
  }));

  try {
    const rProcess = spawn(rscriptCmd, args, spawnOptions);
    let lineBuffer = '';

    let currentLogState = 'info';
    let finalMessage = null;

    rProcess.stdout.on('data', (data) => {
      lineBuffer += data.toString();
      const lines = lineBuffer.split('\n');
      lineBuffer = lines.pop() || '';

      lines.forEach(line => {
        if (line.trim()) {
          // Logs em JSON emitidos pelo utils.R
          if (line.trim().startsWith('{"json_log":true')) {
            try {
              const parsed = JSON.parse(line.trim());

              if (parsed.type === 'progress') {
                ws.send(JSON.stringify(parsed));
                return;
              }

              if (parsed.type === 'config_data') {
                if (parsed.data && parsed.data.final_message) {
                  finalMessage = parsed.data.final_message;
                }
                return;
              }

              if (parsed.message === '') return;
              ws.send(JSON.stringify({ type: 'output', message: parsed.message, level: parsed.level }));
              return;
            } catch (e) { /* Falhou em ler JSON, sigo parseando como texto puro */ }
          }

          const lowerLine = line.toLowerCase();

          if (lowerLine.includes('alerta')) currentLogState = 'warning';
          else if (lowerLine.includes('erro ') || lowerLine.includes('error')) currentLogState = 'error';
          else if (lowerLine.includes('sucesso') || lowerLine.includes('success')) currentLogState = 'success';
          else if (line.includes('╭') || lowerLine.includes('início script') || line.includes('===')) currentLogState = 'section';

          let lineLevel = currentLogState;

          // Linha sem caracteres de caixa: reinicia estado e faz detecção sem contexto
          if (!line.includes('│') && !line.includes('█') && !line.includes('╭') && !line.includes('╰') && !line.includes('├') && !line.includes('▄') && !line.includes('▀') && !line.includes('▒') && !line.includes('░')) {
            currentLogState = 'info';
            lineLevel = detectLogLevel(line);
          } else if (currentLogState === 'info') {
            // Mesmo dentro de caixa, aplica detecção sem contexto quando estado é 'info'
            lineLevel = detectLogLevel(line);
          }

          ws.send(JSON.stringify({ type: 'output', message: line, level: lineLevel }));
        }
      });
    });

    rProcess.stderr.on('data', (data) => {
      const lines = data.toString().split('\n');
      lines.forEach(line => {
        if (line.trim()) {
          ws.send(JSON.stringify({ type: 'output', message: line, level: line.includes('Error') ? 'error' : 'warning' }));
        }
      });
    });

    rProcess.on('close', (code) => {
      if (lineBuffer.trim()) {
        ws.send(JSON.stringify({ type: 'output', message: lineBuffer.trim(), level: 'info' }));
      }

      let logType = 'success';
      let logMsg = '';

      if (code === 1) {
        logType = 'error';
        logMsg = `código de erro: ${code}`;
      } else if (code === 2) {
        logType = 'warning';
        logMsg = 'código: 2';
      } else if (code !== 0) {
        logType = 'error';
        logMsg = `código inesperado: ${code}`;
      }

      const notificationMessage = buildNotificationMessage(scriptFolder, logType, finalMessage, logMsg);

      if (logType === 'success') {
        logger.success(notificationMessage, 'RScript');
      } else if (logType === 'warning') {
        logger.warn(notificationMessage, 'RScript');
      } else {
        logger.error(notificationMessage, 'RScript');
      }

      ws.send(JSON.stringify({
        type: logType,
        message: logMsg,
        notificationMessage,
        exitCode: code,
        scriptName: scriptFolder
      }));
    });

    rProcess.on('error', (err) => {
      logger.error(`Erro no processo R: ${err.message}`, 'RScript', err);
      ws.send(JSON.stringify({ type: 'error', scriptName: scriptFolder, message: 'Erro fatal ao executar R: ' + err.message }));
    });

  } catch (err) {
    logger.error(`Erro ao iniciar Rscript: ${err.message}`, 'RScript', err);
    ws.send(JSON.stringify({
      type: 'error',
      scriptName: scriptFolder,
      message: `Erro ao iniciar Rscript. O R está instalado? Detalhe: ${err.message}`
    }));
  }
}

const TAB_NAMES = {
  atas: 'Atas', atas_mailmerge: 'Atas (Mailmerge)', catmat: 'Catmat',
  fornecedores: 'Fornecedores', importacao: 'Importação',
  mapas: 'Mapas', powerbi: 'Power BI', instalacao: 'Instalação'
};

const TYPE_LABELS = {
  success: 'concluído com sucesso',
  warning: 'concluído com alertas',
  error: 'falhou'
};

function buildNotificationMessage(scriptName, type, finalMessage, message) {
  const sourceName = TAB_NAMES[scriptName] || scriptName;
  const label = TYPE_LABELS[type] || 'finalizado';
  const detail = finalMessage || message || '';
  return `Script "${sourceName}" ${label}${detail ? ': ' + detail : ''}`;
}

function detectLogLevel(line) {
  const lowerLine = line.toLowerCase();

  if (lowerLine.includes('error') || lowerLine.includes('✗') || lowerLine.includes('erro ')) {
    return 'error';
  } else if (lowerLine.includes('warning') || lowerLine.includes('⚠') || lowerLine.includes('alerta')) {
    return 'warning';
  } else if (lowerLine.includes('✓') || lowerLine.includes('success') || lowerLine.includes('sucesso')) {
    return 'success';
  } else if (line.includes('╭') || line.includes('╰') || line.includes('├') || line.includes('│') || line.includes('===') || lowerLine.includes('início script')) {
    return 'section';
  } else if (line.startsWith('>') || lowerLine.includes('executando') || lowerLine.includes('■■■')) {
    return 'command';
  } else {
    return 'info';
  }
}
