// --- LÓGICA INTEGRADA DO CONSOLE ---

// Referências para elementos DOM — inicializadas sob demanda via ensureConsoleDOM()
let consoleContainer, consoleHeader, consoleSummary, summaryTitle, summaryDescription;
let closeButton, minimizeButton, consoleOutput, scriptRunningOverlay, overlayLoader;

// Flag global que indica se há um script em execução
let isScriptRunning = false;

// Flag que indica se ao menos um script já foi executado na sessão
let hasEverRun = false;

// Flag que indica se o DOM do console já foi injetado
let _consoleDOMReady = false;

/**
 * Injeta o HTML do console, overlay e logs-drawer no body (uma única vez).
 * Também configura todos os event listeners de arrastar, fechar, minimizar e resize.
 */
function ensureConsoleDOM() {
    if (_consoleDOMReady) return;
    _consoleDOMReady = true;

    // --- Overlay de bloqueio durante execução de script ---
    const overlayHTML = `
    <div id="script-running-overlay" role="status" aria-live="polite" aria-label="Script em execução">
        <div class="overlay-badge">
            <div class="overlay-loader" id="overlay-loader"></div>
            <span class="d-none">Script em execução. Aguarde...</span>
        </div>
    </div>`;

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
                    <i class="fas fa-window-minimize"></i>
                </button>
                <button class="btn-close" type="button" aria-label="Fechar" data-bs-toggle="tooltip" title="Fechar">
                    <i class="fas fa-window-close"></i>
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
                <i class="fas fa-file-alt me-2" data-bs-toggle="tooltip" title="Exibir logs das últimas execuções"></i>
                <span>Logs</span>
            </span>
            <div class="logs-drawer-actions">
                <button id="btn-reopen-console" class="logs-drawer-action-btn" data-bs-toggle="tooltip"
                    title="Reabrir console" style="display:none" onclick="event.stopPropagation(); reopenConsole();">
                    <i class="fas fa-terminal"></i>
                </button>
                <i class="fas fa-window-minimize" id="logs-minimize-icon" data-bs-toggle="tooltip"
                    title="Minimizar"></i>
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

    // --- Cachear referências DOM ---
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

    // --- Configurar event listeners ---

    // Fechar console
    closeButton.addEventListener('click', () => {
        if (!consoleContainer.classList.contains('running')) {
            consoleContainer.classList.remove('show');
        }
    });

    // Minimizar/restaurar console
    minimizeButton.addEventListener('click', () => {
        consoleContainer.classList.toggle('minimized');
        if (consoleContainer.classList.contains('minimized')) {
            minimizeButton.innerHTML = '<i class="fas fa-window-restore"></i>';
            minimizeButton.setAttribute('title', 'Restaurar');
        } else {
            minimizeButton.innerHTML = '<i class="fas fa-window-minimize"></i>';
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
    const sidebarEl = document.getElementById('sidebar');
    if (sidebarEl) consoleResizeObserver.observe(sidebarEl);
    window.addEventListener('resize', enforceConsoleConstraints);
}

/** Ativa o overlay de bloqueio */
async function showScriptRunningOverlay() {
    isScriptRunning = true;
    ensureConsoleDOM();
    if (scriptRunningOverlay) {
        scriptRunningOverlay.classList.add('active');
        // Lazy-load other.js sob demanda para getLoader()
        if (typeof getLoader !== 'function' && typeof loadScript === 'function') {
            try { await loadScript('js/other.js'); } catch (e) { /* fallback */ }
        }
        overlayLoader.innerHTML = typeof getLoader === 'function' ? getLoader() : '';
    }
}

/** Desativa o overlay de bloqueio */
function hideScriptRunningOverlay() {
    isScriptRunning = false;
    if (scriptRunningOverlay) {
        scriptRunningOverlay.classList.remove('active');
    }
}

/**
 * Reabre o console flutuante — só funciona se ao menos um script já tiver sido executado.
 * Chamada pelo botão no logs-drawer.
 */
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

    // Registra que ao menos um script já foi executado e libera o botão de reabrir console
    if (!hasEverRun) {
        hasEverRun = true;
        const btnReopen = document.getElementById('btn-reopen-console');
        if (btnReopen) btnReopen.style.removeProperty('display');
    }

    // Ativa o overlay de bloqueio para impedir execuções simultâneas
    showScriptRunningOverlay();

    // Remove o estado minimizado
    consoleContainer.classList.remove('minimized');

    // Remove as classes de status anterior
    const statusClasses = ['finished-success', 'finished-warning', 'finished-error'];
    statusClasses.forEach(cls => consoleContainer.classList.remove(cls));

    // Esconde possíveis barras de progresso anteriores
    const pbContainer = document.getElementById('console-progress-container');
    if (pbContainer) pbContainer.classList.add('d-none');

    // Limpa e configura o estado de "executando"
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
    // Remove a classe de executando
    consoleContainer.classList.remove('running');

    // Desativa o overlay de bloqueio
    hideScriptRunningOverlay();

    // Remove as classes de status correspondente
    const statusClasses = ['finished-success', 'finished-warning', 'finished-error'];
    statusClasses.forEach(cls => consoleContainer.classList.remove(cls));

    // Esconde a barra de progresso
    const pbContainer = document.getElementById('console-progress-container');
    if (pbContainer) pbContainer.classList.add('d-none');

    const btnRun = document.getElementById('btn-run-' + result.scriptName);
    if (btnRun) {
        btnRun.disabled = false;
    }

    // Atualiza o resumo com base no status
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

    // Adiciona uma mensagem final ao log se não houver uma
    if (consoleOutput.innerHTML.trim() === '') {
        consoleOutput.innerHTML = result.log || 'Nenhum log detalhado foi retornado.';
    }
}

/** 
 * Lida com as atualizações da barra de progresso do R via WebSocket 
 * @param {{action: 'start'|'update'|'close', value: number, max: number, label: string}} data 
 */
function updateProgressBar(data) {
    let container = document.getElementById('console-progress-container');
    let bar = document.getElementById('console-progress-bar');

    // Create elements if they don't exist
    if (!container) {
        const textContainer = document.querySelector('#console-summary .summary-text');
        if (textContainer) {
            textContainer.classList.add('flex-grow-1'); // ensure takes full width

            container = document.createElement('div');
            container.id = 'console-progress-container';
            container.className = 'progress d-none  ';
            container.style.height = '16px';

            bar = document.createElement('div');
            bar.id = 'console-progress-bar';
            bar.className = 'progress-bar progress-bar-striped progress-bar-animated text-bg-info';
            bar.setAttribute('role', 'progressbar');
            bar.style.width = '0%';

            container.appendChild(bar);
            textContainer.appendChild(container);
        }
    }

    if (!container || !bar) return;

    if (data.action === 'start') {
        container.classList.remove('d-none');
        summaryDescription.classList.add('d-none');
        bar.style.width = '0%';
        bar.setAttribute('aria-valuemax', data.max);
        if (data.label) {
            bar.innerText = data.label;
        }
    } else if (data.action === 'update') {
        const max = parseFloat(bar.getAttribute('aria-valuemax') || '100');
        const percentage = (data.value / max) * 100;
        bar.style.width = `${percentage}%`;
        if (data.label) {
            bar.innerText = data.label;
        }
    } else if (data.action === 'close') {
        setTimeout(() => {
            container.classList.add('d-none');
            summaryDescription.classList.remove('d-none');
        }, 300);
    }
}

function runRScript(scriptName, customParams = null) {
    // 0. Bloquear se já há um script em execução
    if (isScriptRunning) {
        showToast('Aguarde o término do script em execução antes de iniciar outro.', 'warning', 5000, 'execução');
        return;
    }

    // 1. Verificar se está conectado
    if (!ws || ws.readyState !== WebSocket.OPEN) {
        showToast('Servidor não está conectado. Execute novamente o arquivo start.cmd para iniciar o servidor.', 'error', 10000, 'inicialização');
        return;
    }

    let dados = {};

    if (customParams) {
        dados = customParams;
    } else {
        // 2. Coletar dados do formulário
        const aba = document.querySelector(`#${scriptName}`);
        if (!aba) {
            console.error(`Aba ${scriptName} não encontrada`);
            return;
        }

        // Lógica específica para cada script
        if (scriptName === 'atas') {
            // Para Atas, usar a função específica de processamento
            dados = processarDadosAtas();
        } else {
            // Para outros scripts, manter comportamento original
            const campos = aba.querySelectorAll('input, select, textarea');
            campos.forEach(campo => {
                if (campo.id) {
                    if (campo.type === 'checkbox' || campo.type === 'radio') {
                        dados[campo.id] = campo.checked;
                    } else {
                        dados[campo.id] = campo.value;
                    }
                }
            });
        }

        // Adicionar arquivos selecionados aos parâmetros
        if (scriptName === 'catmat' && selectedFiles['catmat-lista-itens-tr']) {
            dados['arquivo_selecionado'] = selectedFiles['catmat-lista-itens-tr'];
        }
    }

    // 4. Preparar a UI do console
    prepareConsoleForExecution(scriptName);

    // 5. Enviar a mensagem via WebSocket
    const message = {
        action: 'execute-r-script',
        scriptName: scriptName,
        params: dados
    };

    try {
        ws.send(JSON.stringify(message));
    } catch (error) {
        console.error('Erro ao enviar mensagem:', error);
        // Reverter o estado da UI em caso de falha no envio
        handleScriptResult({
            status: 'error',
            message: `Falha ao enviar comando para o servidor: ${error.message}`,
            log: '',
            scriptName: scriptName
        });
    }
}

// Lógica para tornar o console arrastável e respeitar limites

let isDragging = false;
let offsetX, offsetY;

// Função para garantir que o console respeite os limites
const enforceConsoleConstraints = () => {
    if (!consoleContainer) return;
    const rect = consoleContainer.getBoundingClientRect();
    if (rect.width === 0 && rect.height === 0) return;

    const sidebar = document.getElementById('sidebar');
    const navbar = document.getElementById('navbar');

    const minLeft = sidebar ? sidebar.offsetWidth : 0;
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

// Função para iniciar o arraste
const onDragStart = (e) => {
    // Ignorar se o clique foi no botão de fechar ou em outro botão no header
    if (e.target.closest('button')) {
        return;
    }

    isDragging = true;

    // A primeira vez que arrastamos, o console é posicionado com 'bottom' e 'right'.
    // Convertemos para 'top' e 'left' para que o arraste funcione corretamente.
    const rect = consoleContainer.getBoundingClientRect();
    if (getComputedStyle(consoleContainer).bottom !== 'auto' || getComputedStyle(consoleContainer).right !== 'auto') {
        consoleContainer.style.top = `${rect.top}px`;
        consoleContainer.style.left = `${rect.left}px`;
        consoleContainer.style.bottom = 'auto';
        consoleContainer.style.right = 'auto';
        enforceConsoleConstraints();
    }

    // Calcula o deslocamento do mouse em relação ao canto superior esquerdo do console
    offsetX = e.clientX - consoleContainer.offsetLeft;
    offsetY = e.clientY - consoleContainer.offsetTop;

    document.body.classList.add('dragging-console');

    document.addEventListener('mousemove', onDragMove);
    document.addEventListener('mouseup', onDragEnd, { once: true });
};

// Função para mover o console
const onDragMove = (e) => {
    if (!isDragging) return;

    const sidebar = document.getElementById('sidebar');
    const navbar = document.getElementById('navbar');

    const minLeft = sidebar ? sidebar.offsetWidth : 0;
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

// Função para finalizar o arraste
const onDragEnd = () => {
    isDragging = false;
    document.body.classList.remove('dragging-console');
    document.removeEventListener('mousemove', onDragMove);
    enforceConsoleConstraints();
};

// Configurar Event Delegation para botões de execução (.btn-run)
// Isso permite que botões carregados dinamicamente (Lazy Loading) funcionem sem reatribuir listeners
document.addEventListener('click', (event) => {
    const btn = event.target.closest('.btn-run');

    if (btn) {
        // Botões com onclick gerenciam sua própria execução (ex: importação, mailmerge).
        // O onclick já foi disparado antes desta delegação (fase target vs bubbling),
        // então a verificação de isScriptRunning é feita pelo próprio runRScript internamente.
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