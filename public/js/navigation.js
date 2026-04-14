// --- Constante centralizada de abas ---
const TAB_LIST = [
    { id: 'home', label: 'Página inicial', icon: 'home', hiddenAppButton: true, separator: 'after', hiddenHomeCard: true },
    { id: 'atas', label: 'Atas', icon: 'description', description: 'Gere as Atas de Registro de Preços', color: 'text-primary', logsDrawer: true },
    { id: 'catmat', label: 'Catmat', icon: 'percent_discount', description: 'Verifique as margens de preferência dos itens do TR', color: 'text-info', logsDrawer: true },
    { id: 'fornecedores', label: 'Fornecedores', icon: 'storefront', description: 'Atualize os dados dos fornecedores', color: 'text-success', logsDrawer: true },
    { id: 'importacao', label: 'Importação', icon: 'upload_file', description: 'Gere os arquivos para importação dos pedidos e relatórios gerenciais', color: 'text-warning', logsDrawer: true },
    { id: 'mapas', label: 'Mapas', icon: 'shopping_basket', description: 'Transforme Mapas de licitação em listas prévias', logsDrawer: true },
    { id: 'powerbi', label: 'Power BI', icon: 'add_chart', description: 'Gere os dados para o Observatório', color: 'text-danger', separator: 'after', logsDrawer: true },
    { id: 'instalacao', label: 'Instalação', icon: 'build', hiddenHomeCard: true, logsDrawer: true },
];

// --- Lazy Loading de Scripts e Estilos sob demanda ---
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

const _loadedStyles = new Set();
function loadStylesheet(href) {
    if (_loadedStyles.has(href)) return Promise.resolve();
    return new Promise((resolve, reject) => {
        const link = document.createElement('link');
        link.rel = 'stylesheet';
        link.href = href;
        link.onload = () => { _loadedStyles.add(href); resolve(); };
        link.onerror = () => reject(new Error(`Falha ao carregar CSS: ${href}`));
        document.head.appendChild(link);
    });
}

// Recursos compartilhados por todas as abas de script (não usados na home)
function loadScriptTabAssets() {
    return Promise.all([
        loadScript('js/file_system.js'),
        loadStylesheet('css/forms.css'),
        loadStylesheet('css/files.css'),
        loadStylesheet('css/console.css'),
    ]);
}

const TAB_SCRIPTS = {
    atas: 'js/atas.js',
    fornecedores: 'js/fornecedores.js',
    importacao: 'js/importacao.js',
    powerbi: 'js/powerbi.js',
    instalacao: 'js/instalacao.js'
};

const navbar = document.getElementById('navbar');

let globalComputerName = '';

// --- Navbar App Buttons ---

function buildHiddenTablist() {
    const tablist = document.getElementById('hidden-tablist');
    if (!tablist) return;

    tablist.innerHTML = TAB_LIST.map(tab =>
        `<a id="${tab.id}-tab" data-bs-toggle="pill" data-bs-target="#${tab.id}" role="tab" aria-controls="${tab.id}"></a>`
    ).join('');
}

function buildTabPanes() {
    const tabContent = document.getElementById('tabContent');
    if (!tabContent) return;

    // Verificar aba preferida salva para definir qual pane inicia ativa
    let defaultTab = localStorage.getItem('lastActiveTab') || 'home';
    try {
        const saved = JSON.parse(localStorage.getItem(STORAGE_KEY) || '{}');
        if (saved?.preferences?.preferredTab) {
            defaultTab = saved.preferences.preferredTab;
        }
    } catch (e) { /* fallback para lastActiveTab */ }

    tabContent.innerHTML = TAB_LIST.map(tab => {
        const isDefault = tab.id === defaultTab;
        const activeClass = isDefault ? ' show active' : ' fade';
        return `
            <div class="tab-pane${activeClass}" id="${tab.id}" role="tabpanel" aria-labelledby="${tab.id}-tab" data-load-url="tabs/${tab.id}.html">
                <div class="custom-spinner-container">
                    <div class="custom-spinner text-primary">
                        <div class="spinner-border"></div>
                        <span role="status">Carregando aba...</span>
                    </div>
                </div>
            </div>`;
    }).join('');
}

function buildNavbarApps() {
    const center = document.querySelector('.navbar-center');
    if (!center) return;

    center.innerHTML = '';

    const container = document.createElement('div');
    container.className = 'navbar-apps';

    TAB_LIST.filter(tab => !tab.hiddenNavbar).forEach(tab => {
        const btn = document.createElement('button');
        btn.className = 'navbar-app-btn';
        btn.dataset.tabId = tab.id;
        btn.setAttribute('aria-label', tab.label);
        btn.style.setProperty('--color', tab.color);
        btn.innerHTML = `<i class="material-symbols-outlined">${tab.icon}</i><span class="navbar-app-label">${tab.label}</span>`;
        container.appendChild(btn);

        if (tab.separator === 'after') {
            const separator = document.createElement('div');
            separator.className = 'header-separator';
            separator.setAttribute('role', 'separator');
            container.appendChild(separator);
        }
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

// --- Home Cards ---

function buildHomeCards() {
    const grid = document.getElementById('home-cards-grid');
    if (!grid) return;

    const preferredTab = appState?.preferences?.preferredTab || '';

    grid.innerHTML = TAB_LIST
        .filter(tab => !tab.hiddenHomeCard)
        .map(tab => {
            const colorClass = tab.color ? ` ${tab.color}` : '';
            const isPreferred = preferredTab && tab.id === preferredTab;
            const starActiveClass = isPreferred ? ' active' : '';
            const starTitle = isPreferred ? 'Aba preferida' : 'Definir como aba preferida';
            return `<a class="card-link card-glass-container" data-tab-id="${tab.id}" role="button"
                aria-label="${tab.label}">
                <button class="card-glass-star${starActiveClass}" data-tab-id="${tab.id}"
                    data-bs-toggle="tooltip" data-bs-title="${starTitle}" aria-label="${starTitle}">
                    <i class="material-symbols-outlined">star</i>
                </button>
                <div class="card-glass-title">${tab.label}</div>
                <hr>
                <div class="card-glass-text">${tab.description || ''}</div>
                <div class="card-glass-icon">
                    <i class="material-symbols-outlined${colorClass}">${tab.icon}</i>
                </div>
            </a>`;
        }).join('');

    grid.addEventListener('click', (e) => {
        const starBtn = e.target.closest('.card-glass-star');
        if (starBtn) {
            e.preventDefault();
            e.stopPropagation();

            if (starBtn.classList.contains('active')) {
                // Estrela preferida → abre painel de configurações
                if (typeof openPanel === 'function') openPanel('settings-panel');
            } else {
                // Estrela cinza → salva como preferida
                const tabId = starBtn.dataset.tabId;
                appState.preferences.preferredTab = tabId;
                saveAppState();
                applyUserPreferences(appState.preferences);
                if (typeof showToast === 'function') {
                    showToast('Aba preferida atualizada', 'success', 2000, 'configurações');
                }
            }
            return;
        }

        const card = e.target.closest('.card-link');
        if (!card) return;
        e.preventDefault();

        const tabId = card.dataset.tabId;
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

    grid.innerHTML = TAB_LIST.filter(tab => !tab.hiddenAppButton).map(tab => {
        const button = `<button class="app-launcher-item" data-tab-id="${tab.id}" aria-label="${tab.label}">
            <div class="app-launcher-icon">
                <i class="material-symbols-outlined">${tab.icon}</i>
            </div>
            <span class="app-launcher-label">${tab.label}</span>
        </button>`;
        return tab.separator === 'after' ? button + '<hr>' : button;
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

        // Carregar recursos compartilhados das abas de script antes de injetar o HTML,
        // para evitar FOUC e garantir disponibilidade de file_system.js
        if (isMainTab) {
            const earlyTabId = tabButton.getAttribute('id')?.replace('-tab', '') || tabButton.getAttribute('aria-controls');
            if (earlyTabId && earlyTabId !== 'home') {
                try {
                    await loadScriptTabAssets();
                } catch (error) {
                    console.error('Erro ao carregar recursos da aba:', error);
                }
            }
        }

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

                    if (tabIdForScript === 'home') buildHomeCards();

                    initScriptFormHeaders(targetPane);
                    targetPane.setAttribute('data-loaded', 'true');
                } catch (error) {
                    console.error('Erro no lazy loading:', error);
                    let msgError = `<div class="alert alert-danger">Erro ao carregar aba: ${error.message}</div>`;
                    if (typeof getLoader === 'function') {
                        msgError += getLoader(4);
                    } else {
                        try {
                            await loadScript('js/loaders.js');
                            if (typeof getLoader === 'function') msgError += getLoader(4);
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

            if (tabId !== 'home') {
                await refreshScriptFileLists(tabId);
            }
            setupLiveValidation(tabId);
            validateTabFields(tabId);
            atualizarIndicadoresSubTabs();
            createRequiredFieldsTooltip();

            if (tabId === 'atas' && typeof inicializarAtas === 'function') {
                inicializarAtas();
            }

            if (tabId === 'importacao' && typeof inicializarImportacao === 'function') {
                inicializarImportacao();
            }

            if (tabId === 'fornecedores' && typeof inicializarFornecedores === 'function') {
                inicializarFornecedores();
            }

            // Guard: após awaits anteriores, verificar se esta aba ainda é a ativa
            if (!tabButton.classList.contains('active')) return;

            const tabConfig = TAB_LIST.find(t => t.id === tabId);
            if (globalComputerName && tabConfig?.logsDrawer) {
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

// Executados imediatamente (o DOM já está parseado com defer) para que
// outros scripts com DOMContentLoaded encontrem os tab-panes existentes
buildTabPanes();
buildHiddenTablist();

document.addEventListener('DOMContentLoaded', () => {
    buildNavbarApps();
    buildAppLauncher();
});
