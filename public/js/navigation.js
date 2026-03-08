// --- Constante centralizada de abas ---
const TAB_LIST = [
    { id: 'home', label: 'Página inicial', icon: 'fa-solid fa-home', color: '#3b82f6', hidden: true },
    { id: 'atas', label: 'Atas', icon: 'fa-solid fa-file-alt', color: '#e11d48' },
    { id: 'catmat', label: 'Catmat', icon: 'fa-solid fa-list', color: '#fb923c' },
    { id: 'fornecedores', label: 'Fornecedores', icon: 'fa-solid fa-building', color: '#10b981' },
    { id: 'importacao', label: 'Importação', icon: 'fa-solid fa-gavel', color: '#facc15' },
    { id: 'mapas', label: 'Mapas', icon: 'fa-solid fa-basket-shopping', color: '#84cc16' },
    { id: 'powerbi', label: 'Power BI', icon: 'fa-solid fa-chart-line', color: '#8b5cf6', separator: 'after' },
    { id: 'instalacao', label: 'Instalação', icon: 'fa-solid fa-wrench', color: '#0ea5e9' },
];

// --- Lazy Loading de Scripts sob demanda ---
const _loadedScripts = new Set();
function loadScript(src) {
    if (_loadedScripts.has(src)) return Promise.resolve();
    return new Promise((resolve, reject) => {
        const script = document.createElement('script');
        script.src = src;
        script.onload = () => { _loadedScripts.add(src); resolve(); };
        script.onerror = () => reject(new Error(`Falha ao carregar script: ${src}`));
        document.head.appendChild(script);
    });
}

const TAB_SCRIPTS = {
    atas: 'js/atas.js',
    importacao: 'js/importacao.js',
    powerbi: 'js/powerbi.js'
};

const mainContent = document.getElementById('mainContent');
const navbar = document.getElementById('navbar');

let globalComputerName = '';

// --- Navbar App Buttons ---

function buildNavbarApps() {
    const center = document.querySelector('.navbar-center');
    if (!center) return;

    center.innerHTML = '';

    const container = document.createElement('div');
    container.className = 'navbar-apps';

    TAB_LIST.filter(tab => tab.id !== 'instalacao').forEach(tab => {
        const btn = document.createElement('button');
        btn.className = 'navbar-app-btn';
        btn.dataset.tabId = tab.id;
        btn.setAttribute('aria-label', tab.label);
        btn.setAttribute('data-bs-toggle', 'tooltip');
        btn.setAttribute('data-bs-placement', 'bottom');
        btn.setAttribute('data-bs-title', tab.label);
        btn.style.setProperty('--color', tab.color);
        btn.innerHTML = `<i class="${tab.icon}"></i>`;
        container.appendChild(btn);
    });

    center.appendChild(container);

    container.addEventListener('click', (e) => {
        const btn = e.target.closest('.navbar-app-btn');
        if (!btn) return;

        container.querySelectorAll('.navbar-app-btn').forEach(b => {
            const t = bootstrap.Tooltip.getInstance(b);
            if (t) t.hide();
        });

        const tabId = btn.dataset.tabId;
        const tabButton = document.getElementById(`${tabId}-tab`);
        if (tabButton) {
            const tab = new bootstrap.Tab(tabButton);
            tab.show();
        }
    });
}

// --- App Launcher ---

function buildAppLauncher() {
    const grid = document.getElementById('app-launcher-grid');
    if (!grid) return;

    grid.innerHTML = TAB_LIST.filter(tab => !tab.hidden).map(tab => {
        const button = `<button class="app-launcher-item" data-tab-id="${tab.id}" aria-label="${tab.label}">
            <div class="app-launcher-icon">
                <i class="${tab.icon}"></i>
            </div>
            <span class="app-launcher-label">${tab.label}</span>
        </button>`;
        return tab.separator === 'after' ? button + '<hr class="app-launcher-separator">' : button;
    }).join('');

    grid.addEventListener('click', (e) => {
        const item = e.target.closest('.app-launcher-item');
        if (!item) return;

        const tabId = item.dataset.tabId;
        const tabButton = document.getElementById(`${tabId}-tab`);
        if (tabButton) {
            const tab = new bootstrap.Tab(tabButton);
            tab.show();
        }
        closeAllPanels();
    });
}

// --- Navegação por abas ---

function getTabLabel(tabId) {
    const tab = TAB_LIST.find(t => t.id === tabId);
    return tab ? tab.label : tabId;
}

document.addEventListener('shown.bs.tab', async (event) => {
    if (event.target) {
        const tabButton = event.target;

        // Sub-tabs (ex: importação) não devem alterar título nem processar lógica de aba principal
        const isMainTab = tabButton.closest('#hidden-tablist');

        const targetSelector = tabButton.getAttribute('data-bs-target');
        if (targetSelector) {
            const targetPane = document.querySelector(targetSelector);
            if (targetPane && targetPane.hasAttribute('data-load-url') && !targetPane.hasAttribute('data-loaded')) {
                try {
                    const response = await fetch(targetPane.getAttribute('data-load-url'));
                    if (!response.ok) throw new Error('Erro ao carregar conteúdo da aba');

                    const html = await response.text();
                    targetPane.innerHTML = html;

                    const tabIdForScript = tabButton.getAttribute('id')?.replace('-tab', '') || tabButton.getAttribute('aria-controls');
                    if (tabIdForScript && TAB_SCRIPTS[tabIdForScript]) {
                        await loadScript(TAB_SCRIPTS[tabIdForScript]);
                    }

                    targetPane.setAttribute('data-loaded', 'true');
                } catch (error) {
                    console.error('Erro no lazy loading:', error);
                    let msgError = `<div class="alert alert-danger">Erro ao carregar aba: ${error.message}</div>`;
                    if (typeof catLoader === 'function') {
                        msgError += catLoader();
                    } else {
                        try {
                            await loadScript('js/loaders.js');
                            if (typeof catLoader === 'function') msgError += catLoader();
                        } catch (e) { /* fallback: sem animação */ }
                    }
                    targetPane.innerHTML = msgError;
                }
            }
        }

        if (!isMainTab) return;

        let tabId = tabButton.getAttribute('id')?.replace('-tab', '') || tabButton.getAttribute('aria-controls');

        if (tabId) {
            localStorage.setItem('lastActiveTab', tabId);

            const navSubtitle = document.getElementById('nav-subtitle');
            const isHome = tabId === 'home';

            if (navSubtitle) {
                navSubtitle.innerHTML = isHome ? '' : `<span>${getTabLabel(tabId)}</span>`;
                navSubtitle.style.display = isHome ? 'none' : '';
            }

            // Destacar item ativo no App Launcher e Navbar
            document.querySelectorAll('.app-launcher-item').forEach(item => {
                item.classList.toggle('active', item.dataset.tabId === tabId);
            });
            document.querySelectorAll('.navbar-app-btn').forEach(btn => {
                btn.classList.toggle('active', btn.dataset.tabId === tabId);
            });

            await refreshScriptFileLists(tabId);
            setupLiveValidation(tabId);
            validateTabFields(tabId);
            createRequiredFieldsTooltip();

            if (tabId === 'atas' && typeof inicializarAtas === 'function') {
                inicializarAtas();
            }

            if (tabId === 'importacao' && typeof inicializarImportacao === 'function') {
                inicializarImportacao();
            }

            if (globalComputerName && (tabId === 'atas' || tabId === 'catmat' || tabId === 'importacao' || tabId === 'mapas' || tabId === 'powerbi')) {
                if (typeof ensureConsoleDOM === 'function') ensureConsoleDOM();

                const logsList = document.getElementById('logs-file-list');
                const logsDrawer = document.getElementById('logs-drawer');

                if (logsList && logsDrawer) {
                    logsList.dataset.nameContains = `${tabId}_${globalComputerName}`.toLowerCase();
                    logsDrawer.classList.remove('d-none');
                    loadFiles('logs-file-list', '_common', 'log', false);
                }
            } else {
                const logsDrawer = document.getElementById('logs-drawer');
                if (logsDrawer) logsDrawer.classList.add('d-none');
            }
        }
    }
});

async function fetchComputerName() {
    try {
        const response = await fetch('/api/user-info');
        const data = await response.json();
        globalComputerName = data.computerName;
        const computerNameEl = document.getElementById('logs-drawer-computer');
        if (computerNameEl) {
            computerNameEl.textContent = globalComputerName;
        }
    } catch (e) {
        console.error("Erro ao obter nome do computador:", e);
    }
}

fetchComputerName();

document.addEventListener('DOMContentLoaded', () => {
    buildNavbarApps();
    buildAppLauncher();
});
