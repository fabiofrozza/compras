// Referências para elementos DOM — inicializadas sob demanda via ensureConsoleDOM()
let consoleContainer, consoleHeader, consoleSummary, summaryTitle, summaryDescription;
let closeButton, minimizeButton, consoleOutput, scriptRunningOverlay, overlayLoader;

let isScriptRunning = false;
let hasEverRun = false;
let _consoleDOMReady = false;

function ensureConsoleDOM() {
  if (_consoleDOMReady) return;
  _consoleDOMReady = true;

  // --- Overlay de bloqueio durante execução de script ---
  let overlayHTML = `
    <div id="script-running-overlay" role="status" aria-live="polite" aria-label="Script em execução">
  `;
  overlayHTML += getMatrixHtml();
  overlayHTML += `
        <div class="overlay-badge">
            <div class="overlay-loader" id="overlay-loader"></div>
            <span class="d-none">Script em execução. Aguarde...</span>
        </div>
    </div>
    `;
  overlayHTML += getMatrixCSS();

  // --- Console flutuante ---
  const consoleHTML = `
    <div id="console-container">
        <div class="console-header" id="console-header">
            <div class="status-indicators">
                <div class="status-circle" id="status-1"></div>
                <div class="status-circle" id="status-2"></div>
                <div class="status-circle" id="status-3"></div>
            </div>
            <div class="header-buttons">
                <button class="btn-minimize" type="button" aria-label="Minimizar" data-bs-toggle="tooltip"
                    title="Minimizar">
                    <i class="material-symbols-outlined">minimize</i>
                </button>
                <button class="btn-close" type="button" aria-label="Fechar" data-bs-toggle="tooltip" title="Fechar">
                    <i class="material-symbols-outlined">close</i>
                </button>
            </div>
        </div>
        <div id="console-output"></div>
        <div id="console-summary">
            <div class="summary-icon"></div>
            <div class="summary-text">
                <h4></h4>
                <p style="height: 16px;"></p>
            </div>
        </div>
    </div>`;

  // --- Drawer flutuante de logs ---
  const logsDrawerHTML = `
    <div id="logs-drawer" class="d-none shadow-lg collapsed">
        <div class="logs-drawer-header" onclick="toggleLogsDrawer()">
            <span id="logs-drawer-title">
                <i class="material-symbols-outlined me-2" data-bs-toggle="tooltip" title="Exibir logs das últimas execuções">description</i>
                <span>Logs</span>
            </span>
            <div class="logs-drawer-actions">
                <button id="btn-reopen-console" class="logs-drawer-action-btn" data-bs-toggle="tooltip"
                    title="Reabrir console" style="display:none" onclick="event.stopPropagation(); reopenConsole();">
                    <i class="material-symbols-outlined">terminal</i>
                </button>
                <i class="material-symbols-outlined" id="logs-minimize-icon" data-bs-toggle="tooltip"
                    title="Minimizar">minimize</i>
            </div>
        </div>
        <div class="logs-drawer-content">
            <div id="logs-file-list" class="files-list" data-folder="log" data-script-name="_common"
                data-selectable="false" data-extensions="log" data-name-contains="" data-sort="desc">
            </div>
        </div>
    </div>`;

  // Inserir no body (antes dos scripts)
  document.body.insertAdjacentHTML('beforeend', overlayHTML + consoleHTML + logsDrawerHTML);

  consoleContainer = document.getElementById('console-container');
  consoleHeader = document.getElementById('console-header');
  consoleSummary = document.getElementById('console-summary');
  summaryTitle = consoleSummary.querySelector('h4');
  summaryDescription = consoleSummary.querySelector('p');
  closeButton = consoleContainer.querySelector('.btn-close');
  minimizeButton = consoleContainer.querySelector('.btn-minimize');
  consoleOutput = document.getElementById('console-output');
  scriptRunningOverlay = document.getElementById('script-running-overlay');
  overlayLoader = document.getElementById('overlay-loader');

  closeButton.addEventListener('click', () => {
    if (!consoleContainer.classList.contains('running')) {
      consoleContainer.classList.remove('show');
    }
  });

  minimizeButton.addEventListener('click', () => {
    consoleContainer.classList.toggle('minimized');
    if (consoleContainer.classList.contains('minimized')) {
      minimizeButton.innerHTML = '<i class="material-symbols-outlined">open_in_full</i>';
      minimizeButton.setAttribute('title', 'Restaurar');
    } else {
      minimizeButton.innerHTML = '<i class="material-symbols-outlined">minimize</i>';
      minimizeButton.setAttribute('title', 'Minimizar');
    }
  });

  // --- Lógica de arrastar e limites ---
  const consoleResizeObserver = new ResizeObserver(() => {
    if (!isDragging) {
      enforceConsoleConstraints();
    }
  });
  consoleResizeObserver.observe(consoleContainer);
  window.addEventListener('resize', enforceConsoleConstraints);
}

async function showScriptRunningOverlay() {
  isScriptRunning = true;
  ensureConsoleDOM();
  if (scriptRunningOverlay) {
    scriptRunningOverlay.classList.add('active');
    // Lazy-load loaders.js sob demanda para getLoader()
    if (typeof getLoader !== 'function' && typeof loadScript === 'function') {
      try { await loadScript('js/loaders.js'); } catch (e) { /* fallback */ }
    }
    overlayLoader.innerHTML = typeof getLoader === 'function' ? getLoader() : '';
  }
}

function hideScriptRunningOverlay() {
  isScriptRunning = false;
  if (scriptRunningOverlay) {
    scriptRunningOverlay.classList.remove('active');
  }
}

function reopenConsole() {
  if (!hasEverRun) return; // nada a exibir ainda
  consoleContainer.classList.add('show');
}

/** @param {string} scriptName */
function prepareConsoleForExecution(scriptName) {
  ensureConsoleDOM();

  const btnRun = document.getElementById('btn-run-' + scriptName);
  if (btnRun) {
    btnRun.disabled = true;
  }

  if (!hasEverRun) {
    hasEverRun = true;
    const btnReopen = document.getElementById('btn-reopen-console');
    if (btnReopen) btnReopen.style.removeProperty('display');
  }

  showScriptRunningOverlay();
  consoleContainer.classList.remove('minimized');

  const statusClasses = ['finished-success', 'finished-warning', 'finished-error'];
  statusClasses.forEach(cls => consoleContainer.classList.remove(cls));

  const pbContainer = document.getElementById('console-progress-container');
  if (pbContainer) pbContainer.classList.add('d-none');

  consoleOutput.innerHTML = '';
  consoleContainer.classList.add('show', 'running');
  consoleSummary.className = 'summary-running';
  summaryTitle.textContent = 'Executando...';

  // Se for o mailmerge (que é JS) não exibe 'script R'
  if (scriptName === 'atas_mailmerge') {
    summaryDescription.textContent = 'Aguarde o processamento da geração das atas.';
  } else {
    summaryDescription.textContent = 'Aguarde o término do processamento do script R.';
  }

  consoleHeader.addEventListener('mousedown', onDragStart);
}

/** @param {{status: 'success'|'warning'|'error', message: string, log: string, scriptName: string}} result */
function handleScriptResult(result) {
  consoleContainer.classList.remove('running');
  hideScriptRunningOverlay();

  const statusClasses = ['finished-success', 'finished-warning', 'finished-error'];
  statusClasses.forEach(cls => consoleContainer.classList.remove(cls));

  const pbContainer = document.getElementById('console-progress-container');
  if (pbContainer) pbContainer.classList.add('d-none');

  const btnRun = document.getElementById('btn-run-' + result.scriptName);
  if (btnRun) btnRun.disabled = false;

  switch (result.status) {
    case 'success':
      consoleSummary.className = 'summary-success';
      summaryTitle.textContent = 'Execução concluída com sucesso';
      consoleContainer.classList.add('finished-success');
      break;
    case 'warning':
      consoleSummary.className = 'summary-warning';
      summaryTitle.textContent = 'Execução concluída com alertas';
      consoleContainer.classList.add('finished-warning');
      break;
    case 'error':
      consoleSummary.className = 'summary-error';
      summaryTitle.textContent = 'Falha na execução';
      consoleContainer.classList.add('finished-error');
      break;
  }

  summaryDescription.textContent = result.message || 'Verifique o log para mais detalhes.';

  if (consoleOutput.innerHTML.trim() === '') {
    consoleOutput.innerHTML = result.log || 'Nenhum log detalhado foi retornado.';
  }
}

/** 
 * Lida com as atualizações da barra de progresso do R via WebSocket 
 * @param {{action: 'start'|'update'|'close', value: number, max: number, label: string}} data 
 */
let _progressBarGeneration = 0; // Contador para evitar que um close atrasado feche uma barra nova

function updateProgressBar(data) {
  let container = document.getElementById('console-progress-container');
  let bar = document.getElementById('console-progress-bar');

  // Create elements if they don't exist
  if (!container) {
    const textContainer = document.querySelector('#console-summary .summary-text');
    if (textContainer) {
      textContainer.classList.add('flex-grow-1');

      container = document.createElement('div');
      container.id = 'console-progress-container';
      container.className = 'progress d-none  ';
      container.style.height = '16px';

      bar = document.createElement('div');
      bar.id = 'console-progress-bar';
      bar.className = 'progress-bar progress-bar-striped progress-bar-animated text-bg-info';
      bar.setAttribute('role', 'progressbar');
      bar.style.width = '0%';

      // Popup flutuante de percentual
      const popup = document.createElement('span');
      popup.className = 'progress-percent-popup';
      popup.textContent = '0%';
      bar.appendChild(popup);

      container.appendChild(bar);
      textContainer.appendChild(container);
    }
  }

  if (!container || !bar) return;

  // Garante que o popup existe (caso a barra tenha sido criada antes desta versão)
  let popup = bar.querySelector('.progress-percent-popup');
  if (!popup) {
    popup = document.createElement('span');
    popup.className = 'progress-percent-popup';
    popup.textContent = '0%';
    bar.appendChild(popup);
  }

  if (data.action === 'start') {
    _progressBarGeneration++; // Nova geração: invalida qualquer close pendente
    container.classList.remove('d-none');
    summaryDescription.classList.add('d-none');
    bar.style.width = '0%';
    bar.setAttribute('aria-valuemax', data.max);
    popup.textContent = '0%';
    if (data.label) {
      bar.innerText = data.label;
      bar.appendChild(popup); // re-append pois innerText remove filhos
    }
  } else if (data.action === 'update') {
    const max = parseFloat(bar.getAttribute('aria-valuemax') || '100');
    const percentage = (data.value / max) * 100;
    bar.style.width = `${percentage}%`;
    popup.textContent = `${Math.round(percentage)}%`;
    if (data.label) {
      bar.innerText = data.label;
      bar.appendChild(popup); // re-append pois innerText remove filhos
    }
  } else if (data.action === 'close') {
    const generationAtClose = _progressBarGeneration;
    setTimeout(() => {
      // Só esconde se nenhuma nova barra foi iniciada desde o close
      if (_progressBarGeneration === generationAtClose) {
        container.classList.add('d-none');
        summaryDescription.classList.remove('d-none');
      }
    }, 300);
  }
}

function runRScript(scriptName, customParams = null) {
  if (isScriptRunning) {
    showToast('Aguarde o término do script em execução antes de iniciar outro.', 'warning', 5000, 'execução');
    return;
  }

  if (!ws || ws.readyState !== WebSocket.OPEN) {
    showToast('Servidor não está conectado. Execute novamente o arquivo start.cmd para iniciar o servidor.', 'error', 10000, 'inicialização');
    return;
  }

  let dados = {};

  if (customParams) {
    dados = customParams;
  } else {
    const aba = document.querySelector(`#${scriptName}`);
    if (!aba) {
      console.error(`Aba ${scriptName} não encontrada`);
      return;
    }

    const campos = aba.querySelectorAll('input, select, textarea');
    campos.forEach(campo => {
      if (campo.id) {
        dados[campo.id] = (campo.type === 'checkbox' || campo.type === 'radio') ? campo.checked : campo.value;
      }
    });

    if (scriptName === 'catmat' && selectedFiles['catmat-lista-itens-tr']) {
      dados['arquivo_selecionado'] = selectedFiles['catmat-lista-itens-tr'];
    }
  }

  prepareConsoleForExecution(scriptName);

  const message = {
    action: 'execute-r-script',
    scriptName: scriptName,
    params: dados
  };

  try {
    ws.send(JSON.stringify(message));
  } catch (error) {
    console.error('Erro ao enviar mensagem:', error);
    // Reverter estado da UI em caso de falha no envio
    handleScriptResult({
      status: 'error',
      message: `Falha ao enviar comando para o servidor: ${error.message}`,
      log: '',
      scriptName: scriptName
    });
  }
}

// --- Lógica de arraste e limites do console ---

let isDragging = false;
let offsetX, offsetY;

const enforceConsoleConstraints = () => {
  if (!consoleContainer) return;
  const rect = consoleContainer.getBoundingClientRect();
  if (rect.width === 0 && rect.height === 0) return;

  const navbar = document.getElementById('navbar');

  const minLeft = 0;
  const minTop = navbar ? navbar.offsetHeight : 0;

  let currentLeft = rect.left;
  let currentTop = rect.top;
  let currentWidth = rect.width;
  let currentHeight = rect.height;

  const isBottomRight = getComputedStyle(consoleContainer).bottom !== 'auto' || getComputedStyle(consoleContainer).right !== 'auto';
  let needsPositionUpdate = false;
  let needsSizeUpdate = false;

  let newLeft = currentLeft;
  let newTop = currentTop;

  if (currentLeft < minLeft) {
    newLeft = minLeft;
    needsPositionUpdate = true;
  }

  if (currentTop < minTop) {
    newTop = minTop;
    needsPositionUpdate = true;
  }

  let newWidth = currentWidth;
  let newHeight = currentHeight;

  if (newLeft + newWidth > window.innerWidth) {
    newWidth = window.innerWidth - newLeft;
    needsSizeUpdate = true;
  }

  if (newTop + newHeight > window.innerHeight) {
    newHeight = window.innerHeight - newTop;
    needsSizeUpdate = true;
  }

  if (isBottomRight && (needsPositionUpdate || needsSizeUpdate)) {
    consoleContainer.style.left = `${newLeft}px`;
    consoleContainer.style.top = `${newTop}px`;
    consoleContainer.style.bottom = 'auto';
    consoleContainer.style.right = 'auto';
    if (needsSizeUpdate) {
      consoleContainer.style.width = `${newWidth}px`;
      consoleContainer.style.height = `${newHeight}px`;
    }
  } else if (!isBottomRight) {
    if (needsPositionUpdate) {
      consoleContainer.style.left = `${newLeft}px`;
      consoleContainer.style.top = `${newTop}px`;
    }
    if (needsSizeUpdate) {
      consoleContainer.style.width = `${newWidth}px`;
      consoleContainer.style.height = `${newHeight}px`;
    }
  }

  if (!isBottomRight || needsPositionUpdate || needsSizeUpdate) {
    consoleContainer.style.maxWidth = `${window.innerWidth - newLeft}px`;
    consoleContainer.style.maxHeight = `${window.innerHeight - newTop}px`;
  } else {
    const cs = getComputedStyle(consoleContainer);
    consoleContainer.style.maxWidth = `${window.innerWidth - minLeft - (parseInt(cs.right) || 0)}px`;
    consoleContainer.style.maxHeight = `${window.innerHeight - minTop - (parseInt(cs.bottom) || 0)}px`;
  }
};

const onDragStart = (e) => {
  if (e.target.closest('button')) { // ignorar cliques em botões do header
    return;
  }

  isDragging = true;

  // Na primeira vez, o console usa 'bottom'/'right'; convertemos para 'top'/'left' para o arraste funcionar
  const rect = consoleContainer.getBoundingClientRect();
  if (getComputedStyle(consoleContainer).bottom !== 'auto' || getComputedStyle(consoleContainer).right !== 'auto') {
    consoleContainer.style.top = `${rect.top}px`;
    consoleContainer.style.left = `${rect.left}px`;
    consoleContainer.style.bottom = 'auto';
    consoleContainer.style.right = 'auto';
    enforceConsoleConstraints();
  }

  offsetX = e.clientX - consoleContainer.offsetLeft;
  offsetY = e.clientY - consoleContainer.offsetTop;

  document.body.classList.add('dragging-console');
  document.addEventListener('mousemove', onDragMove);
  document.addEventListener('mouseup', onDragEnd, { once: true });
};

const onDragMove = (e) => {
  if (!isDragging) return;

  const navbar = document.getElementById('navbar');

  const minLeft = 0;
  const minTop = navbar ? navbar.offsetHeight : 0;

  const maxLeft = window.innerWidth - consoleContainer.offsetWidth;
  const maxTop = window.innerHeight - consoleContainer.offsetHeight;

  let newLeft = e.clientX - offsetX;
  let newTop = e.clientY - offsetY;

  newLeft = Math.max(minLeft, Math.min(newLeft, maxLeft));
  newTop = Math.max(minTop, Math.min(newTop, maxTop));

  consoleContainer.style.left = `${newLeft}px`;
  consoleContainer.style.top = `${newTop}px`;

  consoleContainer.style.maxWidth = `${window.innerWidth - newLeft}px`;
  consoleContainer.style.maxHeight = `${window.innerHeight - newTop}px`;
};

const onDragEnd = () => {
  isDragging = false;
  document.body.classList.remove('dragging-console');
  document.removeEventListener('mousemove', onDragMove);
  enforceConsoleConstraints();
};

// Event delegation para .btn-run — permite que botões carregados dinamicamente funcionem
document.addEventListener('click', (event) => {
  const btn = event.target.closest('.btn-run');

  if (btn) {
    // Botões com onclick gerenciam sua própria execução (ex: importação, mailmerge);
    // isScriptRunning é verificado internamente pelo próprio runRScript.
    if (btn.hasAttribute('onclick')) return;

    // Bloqueia clique se há script em execução — apenas para botões sem onclick
    if (isScriptRunning) {
      event.preventDefault();
      showToast('Aguarde o término do script em execução antes de iniciar outro.', 'warning', 5000, 'execução');
      return;
    }

    if (!btn.disabled) {
      const scriptName = btn.getAttribute('data-script-name');
      if (scriptName) {
        event.preventDefault();
        runRScript(scriptName);
      }
    }
  }
});

function getMatrixHtml() {
  return `
  <!-- From Uiverse.io by https://uiverse.io/whoisyourdeadie/foolish-rabbit-13 --> 
<div class="matrix-container">
  <div class="matrix-pattern">
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
  </div>
  <div class="matrix-pattern">
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
  </div>
  <div class="matrix-pattern">
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
  </div>
  <div class="matrix-pattern">
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
  </div>
  <div class="matrix-pattern">
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
    <div class="matrix-column"></div>
  </div>
</div>
`;
}

function getMatrixCSS() {
  return `
    <style>
/* From Uiverse.io by whoisyourdeadie */ 
.matrix-container {
  position: relative;
  width: 100%;
  height: 100%;
  display: flex;
}

.matrix-pattern {
  position: relative;
  width: 1000px;
  height: 100%;
  flex-shrink: 0;
}

.matrix-column {
  position: absolute;
  top: -100%;
  width: 20px;
  height: 100%;
  font-size: 16px;
  line-height: 18px;
  font-weight: bold;
  animation: fall linear infinite;
  white-space: nowrap;
}

.matrix-column::before {
  content: "アイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨラリルレロワヲンABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
  position: absolute;
  top: 0;
  left: 0;
  background: linear-gradient(
    to bottom,
    #ffffff 0%,
    #ffffff 5%,
    #00ff41 10%,
    #00ff41 20%,
    #00dd33 30%,
    #00bb22 40%,
    #009911 50%,
    #007700 60%,
    #005500 70%,
    #003300 80%,
    rgba(0, 255, 65, 0.5) 90%,
    transparent 100%
  );
  -webkit-background-clip: text;
  background-clip: text;
  -webkit-text-fill-color: transparent;
  writing-mode: vertical-lr;
  letter-spacing: 1px;
  text-rendering: optimizeLegibility;
  -webkit-font-smoothing: antialiased;
  -moz-osx-font-smoothing: grayscale;
}

.matrix-column:nth-child(1) {
  left: 0px;
  animation-delay: -2.5s;
  animation-duration: 3s;
}
.matrix-column:nth-child(2) {
  left: 25px;
  animation-delay: -3.2s;
  animation-duration: 4s;
}
.matrix-column:nth-child(3) {
  left: 50px;
  animation-delay: -1.8s;
  animation-duration: 2.5s;
}
.matrix-column:nth-child(4) {
  left: 75px;
  animation-delay: -2.9s;
  animation-duration: 3.5s;
}
.matrix-column:nth-child(5) {
  left: 100px;
  animation-delay: -1.5s;
  animation-duration: 3s;
}
.matrix-column:nth-child(6) {
  left: 125px;
  animation-delay: -3.8s;
  animation-duration: 4.5s;
}
.matrix-column:nth-child(7) {
  left: 150px;
  animation-delay: -2.1s;
  animation-duration: 2.8s;
}
.matrix-column:nth-child(8) {
  left: 175px;
  animation-delay: -2.7s;
  animation-duration: 3.2s;
}
.matrix-column:nth-child(9) {
  left: 200px;
  animation-delay: -3.4s;
  animation-duration: 3.8s;
}
.matrix-column:nth-child(10) {
  left: 225px;
  animation-delay: -1.9s;
  animation-duration: 2.7s;
}
.matrix-column:nth-child(11) {
  left: 250px;
  animation-delay: -3.6s;
  animation-duration: 4.2s;
}
.matrix-column:nth-child(12) {
  left: 275px;
  animation-delay: -2.3s;
  animation-duration: 3.1s;
}
.matrix-column:nth-child(13) {
  left: 300px;
  animation-delay: -3.1s;
  animation-duration: 3.6s;
}
.matrix-column:nth-child(14) {
  left: 325px;
  animation-delay: -2.6s;
  animation-duration: 2.9s;
}
.matrix-column:nth-child(15) {
  left: 350px;
  animation-delay: -3.7s;
  animation-duration: 4.1s;
}
.matrix-column:nth-child(16) {
  left: 375px;
  animation-delay: -2.8s;
  animation-duration: 3.3s;
}
.matrix-column:nth-child(17) {
  left: 400px;
  animation-delay: -3.3s;
  animation-duration: 3.7s;
}
.matrix-column:nth-child(18) {
  left: 425px;
  animation-delay: -2.2s;
  animation-duration: 2.6s;
}
.matrix-column:nth-child(19) {
  left: 450px;
  animation-delay: -3.9s;
  animation-duration: 4.3s;
}
.matrix-column:nth-child(20) {
  left: 475px;
  animation-delay: -2.4s;
  animation-duration: 3.4s;
}
.matrix-column:nth-child(21) {
  left: 500px;
  animation-delay: -1.7s;
  animation-duration: 2.4s;
}
.matrix-column:nth-child(22) {
  left: 525px;
  animation-delay: -3.5s;
  animation-duration: 3.9s;
}
.matrix-column:nth-child(23) {
  left: 550px;
  animation-delay: -2s;
  animation-duration: 3s;
}
.matrix-column:nth-child(24) {
  left: 575px;
  animation-delay: -4s;
  animation-duration: 4.4s;
}
.matrix-column:nth-child(25) {
  left: 600px;
  animation-delay: -1.6s;
  animation-duration: 2.3s;
}
.matrix-column:nth-child(26) {
  left: 625px;
  animation-delay: -3s;
  animation-duration: 3.5s;
}
.matrix-column:nth-child(27) {
  left: 650px;
  animation-delay: -3.8s;
  animation-duration: 4s;
}
.matrix-column:nth-child(28) {
  left: 675px;
  animation-delay: -2.5s;
  animation-duration: 2.8s;
}
.matrix-column:nth-child(29) {
  left: 700px;
  animation-delay: -3.2s;
  animation-duration: 3.6s;
}
.matrix-column:nth-child(30) {
  left: 725px;
  animation-delay: -2.7s;
  animation-duration: 3.2s;
}
.matrix-column:nth-child(31) {
  left: 750px;
  animation-delay: -1.8s;
  animation-duration: 2.7s;
}
.matrix-column:nth-child(32) {
  left: 775px;
  animation-delay: -3.6s;
  animation-duration: 4.1s;
}
.matrix-column:nth-child(33) {
  left: 800px;
  animation-delay: -2.1s;
  animation-duration: 3.1s;
}
.matrix-column:nth-child(34) {
  left: 825px;
  animation-delay: -3.4s;
  animation-duration: 3.7s;
}
.matrix-column:nth-child(35) {
  left: 850px;
  animation-delay: -2.8s;
  animation-duration: 2.9s;
}
.matrix-column:nth-child(36) {
  left: 875px;
  animation-delay: -3.7s;
  animation-duration: 4.2s;
}
.matrix-column:nth-child(37) {
  left: 900px;
  animation-delay: -2.3s;
  animation-duration: 3.3s;
}
.matrix-column:nth-child(38) {
  left: 925px;
  animation-delay: -1.9s;
  animation-duration: 2.5s;
}
.matrix-column:nth-child(39) {
  left: 950px;
  animation-delay: -3.5s;
  animation-duration: 3.8s;
}
.matrix-column:nth-child(40) {
  left: 975px;
  animation-delay: -2.6s;
  animation-duration: 3.4s;
}

.matrix-column:nth-child(odd)::before {
  content: "アイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨラリルレロワヲン123456789";
}

.matrix-column:nth-child(even)::before {
  content: "ガギグゲゴザジズゼゾダヂヅデドバビブベボパピプペポヴァィゥェォャュョッABCDEFGHIJKLMNOPQRSTUVWXYZ";
}

.matrix-column:nth-child(3n)::before {
  content: "アカサタナハマヤラワイキシチニヒミリウクスツヌフムユルエケセテネヘメレオコソトノホモヨロヲン0987654321";
}

.matrix-column:nth-child(4n)::before {
  content: "ンヲロヨモホノトソコオレメヘネテセケエルユムフヌツスクウリミヒニチシキイワラヤマハナタサカア";
}

.matrix-column:nth-child(5n)::before {
  content: "ガザダバパギジヂビピグズヅブプゲゼデベペゴゾドボポヴァィゥェォャュョッ!@#$%^&*()_+-=[]{}|;:,.<>?";
}

@keyframes fall {
  0% {
    transform: translateY(-10%);
    opacity: 1;
  }
  100% {
    transform: translateY(200%);
    opacity: 0;
  }
}

@media (max-width: 768px) {
  .matrix-column {
    font-size: 14px;
    line-height: 16px;
    width: 18px;
  }
}

@media (max-width: 480px) {
  .matrix-column {
    font-size: 12px;
    line-height: 14px;
    width: 15px;
  }
}
</style>
`;

}