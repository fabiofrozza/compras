const fs = require('fs').promises;
const path = require('path');
const os = require('os');
const { spawn, execSync } = require('child_process');
const { executarMailmerge } = require('./mailmerge');
let logger;
let cachedRScriptPath = null;
let __dirname_server = path.join(__dirname, '..');
const isPathSafe = (targetPath, baseDir) => {
  const realPath = path.resolve(path.normalize(targetPath));
  const safeBase = path.resolve(path.normalize(baseDir));
  return realPath.toLowerCase().startsWith(safeBase.toLowerCase());
};

// Compara versões no formato "x.y.z" - retorna negativo, zero ou positivo
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
function registerConsoleRoutes(app, l) {
  logger = l;
  app.get('/api/check-node', (_req, res) => {
  
    res.json({
  
      installed: true,
  
      version: process.version.replace(/^v/, ''),
  
      path: process.execPath
  
    });
  
  });
  
  
  
  app.get('/api/check-node-latest', async (_req, res) => {
  
    try {
  
      const response = await fetch('https://nodejs.org/dist/latest/');
  
      const html = await response.text();
  
      const match = html.match(/node-v(\d+\.\d+\.\d+)/);
  
      if (match) {
  
        res.json({ latest: match[1] });
  
      } else {
  
        res.json({ error: 'Não foi possível identificar a versão mais recente.' });
      }
    } catch (err) {
      logger.error(`Erro ao consultar versão do Node.js: ${err.message}`, 'NodeLatest');
      res.json({ error: 'Não foi possível consultar o site do Node.js.' });
    }
  });

  app.get('/api/npm-outdated', async (_req, res) => {
    const npmCli = path.join(__dirname_server, '..', 'bin', 'node_modules', 'npm', 'bin', 'npm-cli.js');
    const projectDir = path.join(__dirname_server, '..');

    try {
      const proc = spawn(process.execPath, [npmCli, 'outdated', '--json'], {
        cwd: projectDir,
        windowsHide: true
      });

      let stdoutData = '';
  
      proc.stdout.on('data', (data) => { stdoutData += data.toString(); });
  
      proc.stderr.on('data', () => { });
  
  
  
      await new Promise(resolve => { proc.on('close', resolve); });
  
  
  
      const outdated = stdoutData.trim() ? JSON.parse(stdoutData) : {};
  
      if (outdated.error) {
  
        return res.json({ error: outdated.error.summary || 'Erro ao verificar pacotes npm.' });
  
      }
  
      const packages = Object.entries(outdated)
  
        .filter(([, info]) => info && info.current !== undefined)
  
        .map(([name, info]) => ({
  
          name,
  
          current: info.current,
  
          wanted: info.wanted,
  
          latest: info.latest
  
        }));
  
  
  
      res.json({ packages });
  
    } catch (err) {
  
      logger.error(`Erro ao verificar pacotes npm: ${err.message}`, 'NpmOutdated');
  
      res.json({ error: 'Não foi possível verificar os pacotes npm.' });
  
    }
  
  });
  
  
  
  app.get('/api/check-npm', async (_req, res) => {
  
    const npmCli = path.join(__dirname_server, '..', 'bin', 'node_modules', 'npm', 'bin', 'npm-cli.js');
  
  
  
    try {
  
      const versionProc = spawn(process.execPath, [npmCli, '--version'], { windowsHide: true });
  
      let versionOut = '';
  
      versionProc.stdout.on('data', (data) => { versionOut += data.toString(); });
  
      await new Promise(resolve => { versionProc.on('close', resolve); });
  
      const currentVersion = versionOut.trim();
  
  
  
      const response = await fetch('https://registry.npmjs.org/npm/latest');
  
      const data = await response.json();
  
      const latestVersion = data.version;
  
  
  
      res.json({ current: currentVersion, latest: latestVersion });
  
    } catch (err) {
  
      logger.error(`Erro ao verificar versão do npm: ${err.message}`, 'NpmVersion');
  
      res.json({ error: 'Não foi possível verificar a versão do npm.' });
  
    }
  
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
  
}
async function executeNpmUpdate(ws) {
  const npmCli = path.join(__dirname_server, '..', 'bin', 'node_modules', 'npm', 'bin', 'npm-cli.js');
  const nodeExe = process.execPath;
  const projectDir = path.join(__dirname_server, '..');
  const binDir = path.join(__dirname_server, '..', 'bin');
  ws.send(JSON.stringify({
    type: 'start',
    message: 'Verificando npm e pacotes...'
  }));
  const sendLog = (message, level = 'info') => {
    ws.send(JSON.stringify({ type: 'output', message, level }));
  };
  const spawnNpm = (args, cwd = projectDir) => spawn(nodeExe, [npmCli, ...args], {
    cwd,
    windowsHide: true
  });
  const streamOutput = (proc) => {
    let lineBuffer = '';
    proc.stdout.on('data', (data) => {
      lineBuffer += data.toString();
      const lines = lineBuffer.split('\n');
      lineBuffer = lines.pop() || '';
      lines.forEach(line => { if (line.trim()) sendLog(line.trim()); });
    });
    proc.stderr.on('data', (data) => {
      const lines = data.toString().split('\n');
      lines.forEach(line => {
        if (!line.trim()) return;
        if (line.toLowerCase().includes('warn')) {
          sendLog(line.trim(), 'warning');
        } else if (line.toLowerCase().includes('err')) {
          sendLog(line.trim(), 'error');
        } else {
          sendLog(line.trim());
        }
      });
    });
  };
  try {
    const results = [];
    // === Etapa 1: Verificar e atualizar o próprio npm ===
    sendLog('Verificando versão do npm...', 'section');
    const versionProc = spawnNpm(['--version']);
    let currentNpmVersion = '';
    versionProc.stdout.on('data', (data) => { currentNpmVersion += data.toString(); });
    await new Promise(resolve => { versionProc.on('close', resolve); });
    currentNpmVersion = currentNpmVersion.trim();
    let latestNpmVersion = '';
    try {
      const response = await fetch('https://registry.npmjs.org/npm/latest');
      const data = await response.json();
      latestNpmVersion = data.version;
    } catch {
      sendLog('Não foi possível consultar a versão mais recente do npm.', 'warning');
    }
    if (currentNpmVersion && latestNpmVersion) {
      sendLog(`  npm: ${currentNpmVersion} → ${latestNpmVersion}`);
      const npmParts = currentNpmVersion.split('.').map(Number);
      const latestParts = latestNpmVersion.split('.').map(Number);
      let npmOutdated = false;
      for (let i = 0; i < Math.max(npmParts.length, latestParts.length); i++) {
        if ((latestParts[i] || 0) > (npmParts[i] || 0)) { npmOutdated = true; break; }
        if ((latestParts[i] || 0) < (npmParts[i] || 0)) break;
      }
      if (npmOutdated) {
        sendLog('');
        sendLog('Atualizando npm...', 'section');
        const npmUpdateProc = spawnNpm(['install', 'npm@latest', '--prefix', binDir], binDir);
        streamOutput(npmUpdateProc);
        const npmUpdateCode = await new Promise(resolve => { npmUpdateProc.on('close', resolve); });
        if (npmUpdateCode === 0) {
          sendLog(`✓ npm atualizado: ${currentNpmVersion} → ${latestNpmVersion}`, 'success');
          results.push(`npm atualizado (${currentNpmVersion} → ${latestNpmVersion})`);
        } else {
          sendLog('Erro ao atualizar o npm.', 'error');
          results.push('falha ao atualizar npm');
        }
      } else {
        sendLog('✓ npm já está atualizado.', 'success');
      }
    }
    sendLog('');
    // === Etapa 2: Verificar e atualizar pacotes do projeto ===
    sendLog('Verificando pacotes do projeto...', 'section');
    const outdatedProcess = spawnNpm(['outdated', '--json']);
    let stdoutData = '';
    outdatedProcess.stdout.on('data', (data) => { stdoutData += data.toString(); });
    outdatedProcess.stderr.on('data', () => { });
    await new Promise(resolve => { outdatedProcess.on('close', resolve); });
    const outdated = stdoutData.trim() ? JSON.parse(stdoutData) : {};
    const packages = Object.keys(outdated);
    if (packages.length === 0) {
      sendLog('✓ Todos os pacotes já estão atualizados.', 'success');
    } else {
      for (const pkg of packages) {
        const info = outdated[pkg];
        sendLog(`  ${pkg}: ${info.current} → ${info.wanted}${info.latest !== info.wanted ? ` (última: ${info.latest})` : ''}`);
      }
      sendLog('');
      sendLog('Atualizando pacotes...', 'section');
      const updateProcess = spawnNpm(['update']);
      streamOutput(updateProcess);
      const updateCode = await new Promise(resolve => { updateProcess.on('close', resolve); });
      if (updateCode === 0) {
        sendLog('');
        sendLog(`✓ ${packages.length} pacote(s) atualizado(s) com sucesso.`, 'success');
        results.push(`${packages.length} pacote(s) atualizado(s)`);
      } else {
        sendLog('Erro ao atualizar pacotes npm.', 'error');
        results.push('falha ao atualizar pacotes');
      }
    }
    // === Resultado final ===
    const hasError = results.some(r => r.startsWith('falha'));
    if (results.length === 0) {
      const msg = 'npm e pacotes Node.js já estão atualizados.';
      const notificationMessage = buildNotificationMessage('npm_update', 'success', null, msg);
      logger.success(notificationMessage, 'NpmUpdate');
      ws.send(JSON.stringify({ type: 'success', message: msg, notificationMessage, scriptName: 'npm_update' }));
    } else if (hasError) {
      const msg = `Atualização concluída com erros: ${results.join('; ')}.`;
      const notificationMessage = buildNotificationMessage('npm_update', 'error', null, msg);
      logger.error(notificationMessage, 'NpmUpdate');
      ws.send(JSON.stringify({ type: 'error', message: msg, notificationMessage, scriptName: 'npm_update' }));
    } else {
      const msg = `Atualização concluída: ${results.join('; ')}.`;
      const notificationMessage = buildNotificationMessage('npm_update', 'success', null, msg);
      logger.success(notificationMessage, 'NpmUpdate');
      ws.send(JSON.stringify({ type: 'success', message: msg, notificationMessage, scriptName: 'npm_update' }));
    }
  } catch (err) {
    const msg = `Erro ao executar npm: ${err.message}`;
    const notificationMessage = buildNotificationMessage('npm_update', 'error', null, msg);
    logger.error(notificationMessage, 'NpmUpdate', err);
    ws.send(JSON.stringify({ type: 'error', message: msg, notificationMessage, scriptName: 'npm_update' }));
  }
}
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
async function executeRScript(ws, scriptFolder, params, isPathSafe, logConsentEnabled) {
  const scriptName = scriptFolder + '.R';
  const scriptPath = path.join(__dirname_server, '..', 'scripts', scriptFolder, scriptName);
  const scriptsDir = path.resolve(path.join(__dirname_server, '..', 'scripts'));
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
  const consentFlag = logConsentEnabled ? 'log-consent=true' : 'log-consent=false';
  const args = [scriptPath, ...Object.values(params), 'json-output', consentFlag];
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
                ws.send(JSON.stringify({ type: 'config_data', scriptName: scriptFolder, data: parsed.data }));
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
  mapas: 'Mapas', powerbi: 'Power BI', instalacao: 'Instalação',
  npm_update: 'Pacotes Node.js'
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
function handleConsoleMessage(ws, data, l, isPathSafeFn, logConsentEnabled) {
  logger = l;
  if (data.action === 'execute-r-script') {
    if (data.scriptName === 'atas_mailmerge') {
      logger.section(`Executando mailmerge de atas (JavaScript): ${data.scriptName} [${ws.clientLabel}]`);
      executeMailmergeJS(ws, data.params);
    } else {
      logger.section(`Executando script R: ${data.scriptName} [${ws.clientLabel}]`);
      executeRScript(ws, data.scriptName, data.params, isPathSafeFn, logConsentEnabled);
    }
  } else if (data.action === 'execute-npm-update') {
    logger.section(`Executando atualização de pacotes npm`);
    executeNpmUpdate(ws);
  }
}
async function checkRAvailable() {
  const rpath = await getRScriptPath();
  return rpath !== null;
}

module.exports = { registerConsoleRoutes, handleConsoleMessage, checkRAvailable };
