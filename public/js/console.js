// Referências para elementos DOM - inicializadas sob demanda via ensureConsoleDOM()
let consoleContainer, consoleHeader, consoleSummary, summaryTitle, summaryDescription;
let closeButton, consoleOutput, scriptRunningOverlay, overlayLoader;

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
  //overlayHTML += getMatrixHtml();
  overlayHTML += `
        <div class="overlay-badge">
            <div class="overlay-loader" id="overlay-loader"></div>
            <span class="d-none">Script em execução. Aguarde...</span>
        </div>
    </div>
    `;
  //overlayHTML += getMatrixCSS();

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
                <button class="btn-close" type="button" aria-label="Fechar">
                    <i class="material-symbols-outlined">close</i>
                </button>
            </div>
        </div>
        <div class="console-output-wrapper"><div id="console-output"></div></div>
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
    <div id="logs-drawer" class="logs-drawer d-none shadow-lg collapsed">
        <div class="logs-drawer-header" onclick="toggleLogsDrawer()">
            <span class="logs-drawer-title" data-bs-toggle="tooltip" data-bs-title="Logs das últimas execuções">
                <i class="material-symbols-outlined me-2">bug_report</i>
                <span>Logs</span>
            </span>
            <div class="logs-drawer-actions">
                <button id="btn-reopen-console" class="logs-drawer-action-btn" data-bs-toggle="tooltip"
                    data-bs-title="Reabrir console" style="display:none" onclick="event.stopPropagation(); reopenConsole();">
                    <i class="material-symbols-outlined">terminal</i>
                </button>
                <button class="logs-drawer-action-btn btn-minimize-logs-drawer" onclick="event.stopPropagation(); toggleLogsDrawer();">
                    <i class="material-symbols-outlined">minimize</i>
                </button>
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

  // Registra tooltips dos elementos recém-inseridos dinamicamente
  if (typeof initializeTooltips === 'function') initializeTooltips();

  consoleContainer = document.getElementById('console-container');
  consoleHeader = document.getElementById('console-header');
  consoleSummary = document.getElementById('console-summary');
  summaryTitle = consoleSummary.querySelector('h4');
  summaryDescription = consoleSummary.querySelector('p');
  closeButton = consoleContainer.querySelector('.btn-close');
  consoleOutput = document.getElementById('console-output');
  scriptRunningOverlay = document.getElementById('script-running-overlay');
  overlayLoader = document.getElementById('overlay-loader');

  closeButton.addEventListener('click', () => {
    if (!consoleContainer.classList.contains('running')) {
      consoleContainer.classList.remove('show');
    }
  });

  consoleContainer.addEventListener('mousedown', () => { mousedownOnConsole = true; });
  document.addEventListener('mouseup', (e) => {
    if (mousedownOnConsole) {
      if (!e.composedPath().includes(consoleContainer)) justDragged = true;
    }
    mousedownOnConsole = false;
  });

  document.addEventListener('click', (e) => {
    if (justDragged) {
      justDragged = false;
      return;
    }
    const path = e.composedPath();
    if (consoleContainer.classList.contains('show') && !consoleContainer.classList.contains('running')) {
      if (!path.includes(consoleContainer)) consoleContainer.classList.remove('show');
    }
    const logsDrawer = document.getElementById('logs-drawer');
    if (logsDrawer && !logsDrawer.classList.contains('d-none') && !logsDrawer.classList.contains('collapsed')) {
      if (!path.includes(logsDrawer)) logsDrawer.classList.add('collapsed');
    }
  });

  // Converte a ancoragem inicial (bottom/right da folha de estilos) em top/left
  // assim que o console fica visível, antes do usuário arrastar ou redimensionar.
  // Após a conversão, o observer é descartado: manter o ResizeObserver ativo durante
  // o resize nativo do CSS atrasa o repaint (cada frame força leitura síncrona de layout).
  const consoleResizeObserver = new ResizeObserver(() => {
    if (isDragging) return;
    const rect = consoleContainer.getBoundingClientRect();
    if (rect.width === 0 && rect.height === 0) return;
    if (consoleContainer.style.bottom === 'auto' && consoleContainer.style.right === 'auto') {
      consoleResizeObserver.disconnect();
      return;
    }
    consoleContainer.style.top = Math.round(rect.top) + 'px';
    consoleContainer.style.left = Math.round(rect.left) + 'px';
    consoleContainer.style.bottom = 'auto';
    consoleContainer.style.right = 'auto';
    enforceConsoleConstraints();
    consoleResizeObserver.disconnect();
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
    overlayLoader.innerHTML = (typeof getLoader === 'function' && !document.body.classList.contains('no-animations')) ? getLoader() : '';
    // Scripts inseridos via innerHTML não são executados pelo browser;
    // é preciso recriá-los como elementos DOM reais.
    overlayLoader.querySelectorAll('script').forEach(oldScript => {
      const newScript = document.createElement('script');
      if (oldScript.src) {
        newScript.src = oldScript.src;
      } else {
        newScript.textContent = oldScript.textContent;
      }
      oldScript.replaceWith(newScript);
    });
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
  const statusClasses = ['finished-success', 'finished-warning', 'finished-error'];
  statusClasses.forEach(cls => consoleContainer.classList.remove(cls));

  const pbContainer = document.getElementById('console-progress-container');
  if (pbContainer) pbContainer.classList.add('d-none');

  consoleOutput.innerHTML = '';
  consoleContainer.classList.add('show', 'running');
  consoleSummary.className = 'summary-running';
  summaryTitle.textContent = 'Executando...';

  if (scriptName === 'atas_mailmerge') {
    summaryDescription.textContent = 'Aguarde o processamento da geração das atas.';
  } else if (scriptName === 'npm_update') {
    summaryDescription.textContent = 'Aguarde a atualização dos pacotes Node.js.';
  } else {
    summaryDescription.textContent = 'Aguarde o término do processamento do script R.';
  }

  consoleHeader.addEventListener('mousedown', onDragStart);
  consoleHeader.addEventListener('touchstart', onDragStart, { passive: false });
}

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
    const processedRadioNames = new Set();
    campos.forEach(campo => {
      if (campo.type === 'radio') {
        if (!processedRadioNames.has(campo.name)) {
          processedRadioNames.add(campo.name);
          const checked = aba.querySelector(`input[type="radio"][name="${campo.name}"]:checked`);
          if (checked) dados[campo.name] = checked.value;
        }
        return;
      }
      if (campo.id) {
        dados[campo.id] = campo.type === 'checkbox' ? campo.checked : campo.value;
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
let justDragged = false;
let mousedownOnConsole = false;
let offsetX, offsetY;

const enforceConsoleConstraints = () => {
  if (!consoleContainer) return;

  // Converte bottom/right para top/left se necessário (facilita os cálculos)
  if (consoleContainer.style.bottom !== 'auto' || consoleContainer.style.right !== 'auto') {
    const rect = consoleContainer.getBoundingClientRect();
    consoleContainer.style.top = Math.round(rect.top) + 'px';
    consoleContainer.style.left = Math.round(rect.left) + 'px';
    consoleContainer.style.bottom = 'auto';
    consoleContainer.style.right = 'auto';
  }

  const rect = consoleContainer.getBoundingClientRect();
  if (rect.width === 0 && rect.height === 0) return;

  const navbar = document.getElementById('navbar');

  const minLeft = 0;
  const minTop = navbar ? navbar.offsetHeight : 0;

  // Usar Math.round evita que frações de subpixel em rect.top/left façam a condicional abaixo
  // ser acionada recursivamente em loops do ResizeObserver (o que congela o redimensionamento nativo)
  let currentLeft = Math.round(rect.left);
  let currentTop = Math.round(rect.top);
  let currentWidth = Math.round(rect.width);
  let currentHeight = Math.round(rect.height);

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
    const overflowX = (newLeft + newWidth) - window.innerWidth;
    const availableShiftX = newLeft - minLeft;
    
    if (availableShiftX > 0) {
      const shiftX = Math.min(overflowX, availableShiftX);
      newLeft -= shiftX;
      needsPositionUpdate = true;
    }

    if (newLeft + newWidth > window.innerWidth) {
      newWidth = window.innerWidth - newLeft;
      needsSizeUpdate = true;
    }
  }

  if (newTop + newHeight > window.innerHeight) {
    const overflowY = (newTop + newHeight) - window.innerHeight;
    const availableShiftY = newTop - minTop;

    if (availableShiftY > 0) {
      const shiftY = Math.min(overflowY, availableShiftY);
      newTop -= shiftY;
      needsPositionUpdate = true;
    }

    if (newTop + newHeight > window.innerHeight) {
      newHeight = window.innerHeight - newTop;
      needsSizeUpdate = true;
    }
  }

  if (needsPositionUpdate) {
    if (consoleContainer.style.left !== `${newLeft}px`) consoleContainer.style.left = `${newLeft}px`;
    if (consoleContainer.style.top !== `${newTop}px`) consoleContainer.style.top = `${newTop}px`;
  }
  if (needsSizeUpdate) {
    if (consoleContainer.style.width !== `${newWidth}px`) consoleContainer.style.width = `${newWidth}px`;
    if (consoleContainer.style.height !== `${newHeight}px`) consoleContainer.style.height = `${newHeight}px`;
  }
};

const onDragStart = (e) => {
  if (e.target.closest('button')) { // ignorar cliques em botões do header
    return;
  }

  const point = e.touches ? e.touches[0] : e;

  isDragging = true;

  // Na primeira vez, o console usa 'bottom'/'right'; convertemos para 'top'/'left' para o arraste funcionar
  if (consoleContainer.style.bottom !== 'auto' || consoleContainer.style.right !== 'auto') {
    const rect = consoleContainer.getBoundingClientRect();
    consoleContainer.style.top = `${Math.round(rect.top)}px`;
    consoleContainer.style.left = `${Math.round(rect.left)}px`;
    consoleContainer.style.bottom = 'auto';
    consoleContainer.style.right = 'auto';
    enforceConsoleConstraints();
  }

  offsetX = point.clientX - consoleContainer.offsetLeft;
  offsetY = point.clientY - consoleContainer.offsetTop;

  document.body.classList.add('dragging-console');
  document.addEventListener('mousemove', onDragMove);
  document.addEventListener('mouseup', onDragEnd, { once: true });
  document.addEventListener('touchmove', onDragMove, { passive: false });
  document.addEventListener('touchend', onDragEnd, { once: true });
};

const onDragMove = (e) => {
  if (!isDragging) return;
  if (e.cancelable) e.preventDefault(); // impede scroll da página durante arraste touch

  const point = e.touches ? e.touches[0] : e;
  const navbar = document.getElementById('navbar');

  const minLeft = 0;
  const minTop = navbar ? navbar.offsetHeight : 0;

  const maxLeft = window.innerWidth - consoleContainer.offsetWidth;
  const maxTop = window.innerHeight - consoleContainer.offsetHeight;

  let newLeft = Math.round(Math.max(minLeft, Math.min(point.clientX - offsetX, maxLeft)));
  let newTop = Math.round(Math.max(minTop, Math.min(point.clientY - offsetY, maxTop)));

  consoleContainer.style.left = `${newLeft}px`;
  consoleContainer.style.top = `${newTop}px`;
};

const onDragEnd = () => {
  isDragging = false;
  justDragged = true;
  document.body.classList.remove('dragging-console');
  document.removeEventListener('mousemove', onDragMove);
  enforceConsoleConstraints();
};

// Event delegation para .btn-run - permite que botões carregados dinamicamente funcionem
document.addEventListener('click', (event) => {
  const btn = event.target.closest('.btn-run');

  if (btn) {
    // Botões com onclick gerenciam sua própria execução (ex: importação, mailmerge);
    // isScriptRunning é verificado internamente pelo próprio runRScript.
    if (btn.hasAttribute('onclick')) return;

    // Bloqueia clique se há script em execução - apenas para botões sem onclick
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

