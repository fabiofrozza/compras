require('dotenv').config({ path: require('path').join(__dirname, '..', '.env'), quiet: true });
const express = require('express');
const WebSocket = require('ws');
const path = require('path');
const fs = require('fs').promises;
const os = require('os');
const dns = require('dns');
const Logger = require('./utils/logger');
const { executarMailmerge } = require('./services/mailmerge');
const { registerConsoleRoutes, handleConsoleMessage } = require('./services/consoleService');
const { validateLink } = require('./services/spreadsheet');
const {
  SCRIPTS_PATH,
  ALLOWED_DELETE_FOLDERS,
  isPathSafe,
  isProtectedFile,
  openInInterface,
  resolveScriptFolder,
  parseFilters,
  filePassesFilter,
  listPregoes,
  listPregaoFiles,
  listImportFiles,
  deletePregao,
  deleteSupplierFile,
  moveSupplierFiles,
} = require('./services/files');

let logConsentEnabled = true; // null/true = salvar logs (permissivo por padrão); false = opt-out

const logger = new Logger({ minLevel: process.env.COMPRAS_LOGGER_MIN_LEVEL || 'debug' });

const app = express();
const PORT = process.env.COMPRAS_PORT || 3000;

// Middlewares
app.use(express.json({ limit: '5mb' }));
app.use(express.static(path.join(__dirname, '..', 'public')));


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
      openInInterface(folderPath);
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
      openInInterface(filePath, true);
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

    const result = await resolveScriptFolder(res, scriptName, innerFolder, 'ListFiles');
    if (!result) return;
    const { folderPath, created } = result;

    const { patterns, filterNameContains } = parseFilters(extensions, nameContains);
    const files = await fs.readdir(folderPath);
    const filteredFiles = files.filter(file => filePassesFilter(file, patterns, filterNameContains));

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

    const result = await resolveScriptFolder(res, scriptName, innerFolder, 'ClearFolder');
    if (!result) return;
    const { folderPath } = result;

    const { patterns, filterNameContains } = parseFilters(extensions, nameContains);
    const files = await fs.readdir(folderPath);
    let deletedCount = 0;

    for (const file of files) {
      // Dupla proteção: Nunca apagar scripts de código ou arquivos protegidos
      if (isProtectedFile(file)) continue;
      if (!filePassesFilter(file, patterns, filterNameContains)) continue;

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

    const result = await resolveScriptFolder(res, scriptName, innerFolder, 'DeleteFile');
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
  try {
    const { url } = req.body;
    const result = await validateLink(url);
    res.json(result);
  } catch (error) {
    logger.error(`Erro ao validar link: ${error.message}`, 'ValidateLink', error);
    res.status(500).json({ error: error.message });
  }
});

// =============================================
// API - Fornecedores
// =============================================

app.get('/api/fornecedores/pregoes', async (_req, res) => {
  try {
    const result = await listPregoes(res);
    if (!result) return;
    res.json(result);
  } catch (error) {
    logger.error(`Erro ao listar pregões: ${error.message}`, 'Fornecedores', error);
    res.status(500).json({ error: error.message });
  }
});

app.get('/api/fornecedores/pregao/:pregao/arquivos', async (req, res) => {
  try {
    const { pregao } = req.params;
    const result = await listPregaoFiles(pregao);

    if (result.error) {
      return res.status(result.statusCode).json({ error: result.error });
    }

    res.json(result);
  } catch (error) {
    logger.error(`Erro ao listar arquivos do pregão: ${error.message}`, 'Fornecedores', error);
    res.status(500).json({ error: error.message });
  }
});

app.get('/api/fornecedores/importar', async (_req, res) => {
  try {
    const result = await listImportFiles();
    res.json(result);
  } catch (error) {
    logger.error(`Erro ao listar arquivos importar: ${error.message}`, 'Fornecedores', error);
    res.status(500).json({ error: error.message });
  }
});

app.delete('/api/fornecedores/pregao/:pregao', async (req, res) => {
  try {
    const { pregao } = req.params;

    if (!ALLOWED_DELETE_FOLDERS.includes('dados')) {
      return res.status(403).json({ error: 'A exclusão nesta pasta não é permitida por segurança.' });
    }

    const result = await deletePregao(pregao);

    if (result.error) {
      return res.status(result.statusCode).json({ error: result.error });
    }

    res.json(result);
  } catch (error) {
    logger.error(`Erro ao excluir pregão: ${error.message}`, 'Fornecedores', error);
    res.status(500).json({ error: error.message });
  }
});

app.delete('/api/fornecedores/arquivo', async (req, res) => {
  try {
    const { filePath } = req.body;

    if (!filePath) {
      return res.status(400).json({ error: 'Caminho do arquivo não fornecido' });
    }

    const result = await deleteSupplierFile(filePath);

    if (result.error) {
      return res.status(result.statusCode).json({ error: result.error });
    }

    res.json(result);
  } catch (error) {
    logger.error(`Erro ao excluir arquivo: ${error.message}`, 'Fornecedores', error);
    res.status(500).json({ error: error.message });
  }
});

app.post('/api/fornecedores/mover', async (req, res) => {
  try {
    const { pregao } = req.body;

    if (!pregao) {
      return res.status(400).json({ error: 'Número do pregão não fornecido' });
    }

    const result = await moveSupplierFiles(pregao, os);

    if (result.error) {
      return res.status(result.statusCode).json({ error: result.error });
    }

    res.json(result);
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

  openInInterface(`http://localhost:${PORT}`);

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

