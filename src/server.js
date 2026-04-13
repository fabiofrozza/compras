require('dotenv').config({ path: require('path').join(__dirname, '..', '.env'), quiet: true });
const express = require('express');
const WebSocket = require('ws');
const path = require('path');
const os = require('os');
const dns = require('dns');
const Logger = require('./utils/logger');
const { registerConsoleRoutes, handleConsoleMessage, checkRAvailable } = require('./services/consoleService');
const { registerFileRoutes } = require('./services/fileRoutes');
const { isPathSafe, openInInterface } = require('./services/files');

let logConsentEnabled = true; // null/true = salvar logs (permissivo por padrão); false = opt-out

const logger = new Logger({ minLevel: process.env.COMPRAS_LOGGER_MIN_LEVEL || 'debug' });

const app = express();
const PORT = process.env.COMPRAS_PORT || 3000;

// Middlewares
app.use(express.json({ limit: '5mb' }));
app.use(express.static(path.join(__dirname, '..', 'public')));


registerFileRoutes(app, logger);

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

    checkRAvailable().then(available => {
      if (ws.readyState === WebSocket.OPEN) {
        ws.send(JSON.stringify({ type: 'r-status', available }));
      }
    });

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

