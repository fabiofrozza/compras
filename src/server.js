// server.js - Servidor Node.js

require('dotenv').config({ path: require('path').join(__dirname, '..', '.env') });
const express = require('express');
const WebSocket = require('ws');
const path = require('path');
const fs = require('fs').promises;
const os = require('os');
const { spawn } = require('child_process');
const Logger = require('./utils/logger');
const { executarMailmerge } = require('./services/mailmerge');

// Inicializar logger (use 'debug' para mais detalhes durante desenvolvimento)
const logger = new Logger({ minLevel: 'debug' });

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
  'temp'
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

// Helper para correspondência de padrões de arquivo (wildcards e extensões)
function matchFilePattern(filename, patterns) {
  if (!patterns || patterns.length === 0) return true;

  return patterns.some(pattern => {
    let p = pattern.trim();
    if (!p || p === '*') return true;

    // Se não tem curingas e não começa com ponto, assume que é extensão
    if (!p.includes('*') && !p.includes('?') && !p.startsWith('.')) {
      return filename.toLowerCase().endsWith('.' + p.toLowerCase());
    }

    // Se começa com ponto, é extensão exata
    if (p.startsWith('.')) {
      return filename.toLowerCase().endsWith(p.toLowerCase());
    }

    // Conversão simples de glob para regex
    const regexString = '^' + p.replace(/[.+^${}()|[\]\\]/g, '\\$&')
      .replace(/\*/g, '.*')
      .replace(/\?/g, '.') + '$';
    return new RegExp(regexString, 'i').test(filename);
  });
}

// Abre uma pasta, arquivo ou URL no aplicativo padrão do sistema operacional
function abrirNaInterface(targetPath, isFile = false) {
  if (process.platform === 'win32') {
    if (isFile) spawn('cmd', ['/c', 'start', '', targetPath]);
    else spawn('explorer', [targetPath]);
  } else if (process.platform === 'darwin') {
    spawn('open', [targetPath]);
  } else {
    spawn('xdg-open', [targetPath]);
  }
}

// Resolve e valida o caminho de uma subpasta de scripts.
// Retorna o folderPath ou null (com o erro já enviado via res).
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

// Parseia os filtros de extensão e nome a partir dos query params
function parsearFiltros(extensions, nameContains) {
  return {
    patterns: extensions ? extensions.split(',') : null,
    filterNameContains: nameContains ? nameContains.toLowerCase().split('_') : []
  };
}

// Verifica se um arquivo passa nos filtros de padrão e nome
function arquivoPassaNoFiltro(file, patterns, filterNameContains) {
  if (filterNameContains.length > 0) {
    if (!filterNameContains.every(term => file.toLowerCase().includes(term.trim()))) return false;
  }
  return matchFilePattern(file, patterns);
}

async function getRScriptPath() {
  if (cachedRScriptPath) return cachedRScriptPath;

  let rscriptCmd = 'Rscript'; // Padrão

  if (process.platform === 'win32') {
    try {
      const { execSync } = require('child_process');
      const whereResult = execSync('where Rscript', { encoding: 'utf8', stdio: ['pipe', 'pipe', 'ignore'] });
      if (whereResult) {
        cachedRScriptPath = whereResult.split('\n')[0].trim();
        return cachedRScriptPath;
      }
    } catch (e) { /* Não está no PATH */ }

    // Busca em caminhos comuns
    const possibleRPaths = [
      'C:\\Program Files\\R\\R-4.4.2\\bin\\Rscript.exe',
      'C:\\Program Files\\R\\R-4.4.1\\bin\\Rscript.exe',
      'C:\\Program Files\\R\\R-4.3.3\\bin\\Rscript.exe',
      // ... adicione as outras versões aqui ...
    ];

    for (const rPath of possibleRPaths) {
      try {
        await fs.access(rPath);
        cachedRScriptPath = rPath;
        return cachedRScriptPath;
      } catch (e) { /* Ignora e tenta o próximo */ }
    }

    // Busca dinâmica no Program Files
    try {
      const programFiles = process.env['ProgramFiles'] || 'C:\\Program Files';
      const rDir = path.join(programFiles, 'R');
      await fs.access(rDir);

      const rVersions = (await fs.readdir(rDir)).filter(f => f.startsWith('R-'));
      if (rVersions.length > 0) {
        rVersions.sort().reverse(); // Pega a mais recente
        const rscriptPath = path.join(rDir, rVersions[0], 'bin', 'Rscript.exe');
        await fs.access(rscriptPath);
        cachedRScriptPath = rscriptPath;
        return cachedRScriptPath;
      }
    } catch (e) { /* Falhou busca dinâmica */ }
  }

  cachedRScriptPath = rscriptCmd; // Fallback para o comando genérico
  return cachedRScriptPath;
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

  // Valida formato da URL
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
      return res.json({
        isValid: false,
        status: 'error',
        msg: `Erro ao acessar o link (HTTP ${response.status}).`,
      });
    }

    const htmlContent = await response.text();

    if (htmlContent.includes('LISTA FINAL')) {
      // Extrair grupo de materiais dos campos input (value)
      const inputValueRegex = /<input[^>]*value="([^"]+)"[^>]*>/gi;
      const inputValues = [];
      let inputMatch;
      while ((inputMatch = inputValueRegex.exec(htmlContent)) !== null) {
        if (inputMatch[1] && inputMatch[1].trim()) {
          inputValues.push(inputMatch[1].trim());
        }
      }
      const grupoMateriais = inputValues.length > 0 ? inputValues.join(', ') : 'Grupo não identificado';

      // Extrair processos SPA
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
    return res.json({
      isValid: false,
      status: 'error',
      msg: 'Erro ao acessar o link informado.',
      error: error.message,
    });
  }
});

// Rota principal
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, '..', 'public', 'index.html'));
});

// Endpoint para verificar instalação do R
app.get('/api/check-r', (req, res) => {
  const rCheck = spawn('Rscript', ['--version']);

  let output = '';
  rCheck.stdout.on('data', (data) => {
    output += data.toString();
  });

  rCheck.stderr.on('data', (data) => {
    output += data.toString();
  });

  rCheck.on('close', (code) => {
    if (code === 0 || output.includes('R scripting')) {
      res.json({
        installed: true,
        version: output.trim(),
        message: 'R está instalado e pronto para uso'
      });
    } else {
      res.json({
        installed: false,
        message: 'R não encontrado. Por favor, instale o R em seu sistema.'
      });
    }
  });
});

// Obter nome do computador
function getComputerName() {
  return os.hostname() || 'Computador';
}

// API - GET user info
app.get('/api/user-info', (req, res) => {
  try {
    res.json({
      computerName: getComputerName(),
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

// Criar servidor HTTP
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

  // Configurar WebSocket e rastrear conexões
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

  // Capturar informações do cliente
  const ip = req.socket.remoteAddress;
  const userAgent = req.headers['user-agent'] || 'Desconhecido';

  logger.info(`Cliente conectado (Total: ${clientCount})`, 'WebSocket');
  logger.debug(`IP: ${ip} | Navegador: ${userAgent}`, 'WebSocket');

  ws.on('message', (message) => {
    try {
      const data = JSON.parse(message);
      if (data.action === 'execute-r-script') {
        // Verificar se é mailmerge de atas (usar versão JavaScript)
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

    // Se não houver mais clientes, fechar o servidor após 2 segundos
    if (clientCount === 0) {
      logger.info('Nenhum cliente conectado. Encerrando servidor...', 'Server');

      // ENCERRAMENTO DESATIVADO TEMPORARIAMENTE

      //setTimeout(() => {
      //    server.close(() => {
      //        clearInterval(interval)
      //        logger.info('Servidor encerrado.', 'Server');
      //        process.exit(0);
      //    });
      //}, 500); // meio segundo de espera
    }
  });

  ws.on('error', (error) => {
    logger.error(`Erro WebSocket: ${error.message}`, 'WebSocket', error);
  });
  });
});

// Função para executar mailmerge de atas em JavaScript
async function executeMailmergeJS(ws, params) {
  try {
    // Enviar início da execução
    ws.send(JSON.stringify({
      type: 'start',
      message: 'Iniciando mailmerge de atas...'
    }));

    // Criar função logger que envia mensagens via WebSocket
    const sendLog = (message, level = 'info') => {
      ws.send(JSON.stringify({
        type: 'output',
        message: message,
        level: level
      }));
    };

    // Executar mailmerge
    const resultado = await executarMailmerge(params, sendLog);

    // Enviar resultado final
    if (resultado.status === 'success' || resultado.status === 'warning') {
      logger.success(`Mailmerge concluído: ${resultado.message}`, 'Mailmerge');
      ws.send(JSON.stringify({
        type: 'success',
        message: resultado.message,
        scriptName: 'atas_mailmerge',
        detalhes: resultado
      }));
    } else {
      logger.warn(`Erro no mailmerge: ${resultado.message}`, 'Mailmerge');
      ws.send(JSON.stringify({
        type: 'error',
        message: resultado.message,
        scriptName: 'atas_mailmerge',
        detalhes: resultado
      }));
    }

  } catch (err) {
    logger.error(`Erro ao executar mailmerge: ${err.message}`, 'Mailmerge', err);
    ws.send(JSON.stringify({
      type: 'error',
      message: `Erro ao executar mailmerge: ${err.message}`,
      scriptName: 'atas_mailmerge'
    }));
  }
}

async function executeRScript(ws, scriptFolder, params) {
  const scriptName = scriptFolder + '.R';
  const scriptPath = path.join(__dirname, '..', 'scripts', scriptFolder, scriptName);

  const scriptsDir = path.resolve(path.join(__dirname, '..', 'scripts'));
  const workingDir = path.resolve(path.join(scriptsDir, scriptFolder));
  //const workingDir = path.join(__dirname, '..', 'scripts', scriptFolder);

  // Validar se o caminho não está tentando escapar da pasta raiz
  if (!isPathSafe(workingDir, scriptsDir)) {
    logger.error(`Tentativa de Path Traversal bloqueada: ${scriptFolder}`, 'Security');
    ws.send(JSON.stringify({ type: 'error', message: 'Acesso negado ao diretório.' }));
    return;
  }

  // 1. Verifica se o script existe (de forma limpa)
  try {
    await fs.access(scriptPath);
  } catch (error) {
    logger.warn(`Script R não encontrado: ${scriptPath}`, 'RScript');
    ws.send(JSON.stringify({
      type: 'error',
      scriptName: scriptFolder,
      message: 'Script R não encontrado: ' + scriptPath
    }));
    return; // Sai da função cedo
  }

  // 2. Pega o caminho do R (com cache e await)
  const rscriptCmd = await getRScriptPath();

  if (rscriptCmd === 'Rscript' && process.platform === 'win32') {
    // Se retornou o padrão no Windows e não encontrou, talvez não esteja instalado
    logger.warn('Caminho absoluto do R não encontrado no Windows, tentando via PATH...', 'RDetect');
  }

  // 3. Preparar argumentos e ambiente
  const args = [scriptPath, ...Object.values(params), 'json-output'];

  const spawnOptions = {
    cwd: workingDir,
    windowsHide: true,
    //shell: process.platform === 'win32' && !rscriptCmd.includes(' ')
  };

  ws.send(JSON.stringify({
    type: 'start',
    message: 'Iniciando execução do script R...'
  }));

  // 4. Executar
  try {
    const rProcess = spawn(rscriptCmd, args, spawnOptions);
    let lineBuffer = '';

    let currentLogState = 'info';

    rProcess.stdout.on('data', (data) => {
      lineBuffer += data.toString();
      const lines = lineBuffer.split('\n');
      lineBuffer = lines.pop() || '';

      lines.forEach(line => {
        if (line.trim()) {
          // Detectar logs em JSON emitidos pelo utils.R
          if (line.trim().startsWith('{"json_log":true')) {
            try {
              const parsed = JSON.parse(line.trim());

              // Se for um bloco de progresso
              if (parsed.type === 'progress') {
                ws.send(JSON.stringify(parsed));
                return;
              }

              if (parsed.message === '') return;
              ws.send(JSON.stringify({ type: 'output', message: parsed.message, level: parsed.level }));
              return;
            } catch (e) { /* Falhou em ler JSON, sigo parseando como texto puro */ }
          }

          const lowerLine = line.toLowerCase();

          // Change state if specific keywords are found inside a block
          if (lowerLine.includes('alerta')) currentLogState = 'warning';
          else if (lowerLine.includes('erro ') || lowerLine.includes('error')) currentLogState = 'error';
          else if (lowerLine.includes('sucesso') || lowerLine.includes('success')) currentLogState = 'success';
          else if (line.includes('╭') || lowerLine.includes('início script') || line.includes('===')) currentLogState = 'section';

          let lineLevel = currentLogState;

          // If the line is not drawn with box-drawing characters, reset to info and use stateless detection
          if (!line.includes('│') && !line.includes('█') && !line.includes('╭') && !line.includes('╰') && !line.includes('├') && !line.includes('▄') && !line.includes('▀') && !line.includes('▒') && !line.includes('░')) {
            currentLogState = 'info';
            lineLevel = detectLogLevel(line);
          } else if (currentLogState === 'info') {
            // Even if box-drawn, try stateless detection for the line if we are in 'info'
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
      let logMsg = 'Script R executado com sucesso!';

      if (code === 1) {
        logType = 'error';
        logMsg = `Script R finalizou com código de erro: ${code}`;
      } else if (code === 2) {
        logType = 'warning';
        logMsg = 'Script R finalizou com alertas (código: 2)';
      } else if (code !== 0) {
        logType = 'error';
        logMsg = `Script R finalizou com código inesperado: ${code}`;
      }

      if (logType === 'success') {
        logger.success(logMsg, 'RScript');
      } else if (logType === 'warning') {
        logger.warn(logMsg, 'RScript');
      } else {
        logger.error(logMsg, 'RScript');
      }

      ws.send(JSON.stringify({ type: logType, message: logMsg, exitCode: code, scriptName: scriptFolder }));
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

// Detectar nível de log baseado no conteúdo da linha
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
