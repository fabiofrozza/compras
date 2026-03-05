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
const sidebar = document.getElementById('sidebar');

const sidebarShadow = document.querySelector('.sidebar-sombra');
const roundCorner = document.querySelector('.round-corner');

const toggleSidebarBtnClose = document.getElementById('toggleSidebarBtnClose');
const toggleSidebarIcon = document.getElementById('toggleSidebarIcon');

const companyLogo = document.getElementById('companyLogo');
const companyLogoExpanded = document.getElementById('companyLogoExpanded');
const navbarLogoContainer = document.getElementById('navbar-top-container');
const logoDepartmentContainer = document.getElementById('logo-department-container');

// Estado: 
// - false = colapsado permanente (hover expande temporariamente)
// - true = expandido permanente (hover não tem efeito)
let isPermanentlyExpanded = false;
let globalComputerName = ''; // Armazena o nome do computador

function setToggleIcon(isExpanded) {
    if (!toggleSidebarIcon) return;
    if (isExpanded) {
        toggleSidebarIcon.classList.remove('fa-bars');
        toggleSidebarIcon.classList.add('fa-close');
    } else {
        toggleSidebarIcon.classList.remove('fa-close');
        toggleSidebarIcon.classList.add('fa-bars');
    }
}

function initializeNavigation() {
    toggleSidebarBtnClose.addEventListener('click', () => {
        if (isPermanentlyExpanded) {
            collapsePermanently();
        } else {
            expandPermanently();
        }
    });
}

function toggleClasses(elements, classAdd, classRemove) {
    const elementsArray = Array.isArray(elements) ? elements : [elements];
    elementsArray.forEach(el => {
        if (classRemove) el.classList.remove(classRemove);
        if (classAdd) el.classList.add(classAdd);
    });
}

function expandPermanently() {
    isPermanentlyExpanded = true;

    toggleClasses([sidebar, sidebarShadow, roundCorner], 'expanded', 'collapsed');
    sidebar.classList.remove('hover-expand');
    toggleClasses(navbar, 'reduced', 'full');
    toggleClasses(mainContent, 'reduced', 'full');
    toggleClasses(navbarLogoContainer, 'expanded', 'collapsed');
    toggleClasses(logoDepartmentContainer, 'd-flex', 'd-none');
    toggleClasses(toggleSidebarBtnClose, null, 'd-none');
    toggleClasses(companyLogo, 'd-none');
    toggleClasses(companyLogoExpanded, null, 'd-none');
    setToggleIcon(true);
    toggleSidebarBtnClose.setAttribute('data-bs-title', 'Fechar barra lateral');
}

function collapsePermanently() {
    isPermanentlyExpanded = false;

    toggleClasses([sidebar, sidebarShadow, roundCorner], 'collapsed', 'expanded');
    sidebar.classList.remove('hover-expand');
    toggleClasses(navbar, 'full', 'reduced');
    toggleClasses(mainContent, 'full', 'reduced');
    toggleClasses(navbarLogoContainer, 'collapsed', 'expanded');
    toggleClasses(logoDepartmentContainer, 'd-none', 'd-flex');
    toggleClasses(companyLogo, null, 'd-none');
    toggleClasses(companyLogoExpanded, 'd-none');
    toggleClasses(toggleSidebarBtnClose, 'd-none');
    setToggleIcon(false);
    toggleSidebarBtnClose.setAttribute('data-bs-title', 'Fixar barra lateral');
}

function expandSidebar() { expandPermanently(); }
function collapseSidebar() { collapsePermanently(); }

sidebar.addEventListener('mouseenter', () => {
    if (!isPermanentlyExpanded) {
        toggleClasses([sidebar, sidebarShadow, roundCorner], 'expanded', 'collapsed');
        sidebar.classList.add('hover-expand');
        toggleClasses(companyLogo, 'd-none');
        toggleClasses(toggleSidebarBtnClose, null, 'd-none');
        setToggleIcon(false); // no hover, ainda mostra hamburguer
    }
});

sidebar.addEventListener('mouseleave', () => {
    if (!isPermanentlyExpanded && sidebar.classList.contains('hover-expand')) {
        toggleClasses([sidebar, sidebarShadow, roundCorner], 'collapsed', 'expanded');
        sidebar.classList.remove('hover-expand');
        toggleClasses(companyLogo, null, 'd-none');
        toggleClasses(toggleSidebarBtnClose, 'd-none');
    }
});

sidebar.addEventListener('click', (event) => {
    const isInteractiveElement = event.target.closest('button, a, .nav-item, .nav-link');
    if (!isInteractiveElement) {
        if (isPermanentlyExpanded) {
            collapsePermanently();
        } else {
            expandPermanently();
        }
    }
});

document.addEventListener('shown.bs.tab', async (event) => {
    if (event.target) {
        // --- Lazy Loading: Carregar conteúdo HTML se necessário ---
        const tabButton = event.target;
        const targetSelector = tabButton.getAttribute('data-bs-target');
        if (targetSelector) {
            const targetPane = document.querySelector(targetSelector);
            if (targetPane && targetPane.hasAttribute('data-load-url') && !targetPane.hasAttribute('data-loaded')) {
                try {
                    const response = await fetch(targetPane.getAttribute('data-load-url'));
                    if (!response.ok) throw new Error('Erro ao carregar conteúdo da aba');

                    const html = await response.text();
                    targetPane.innerHTML = html;

                    // Lazy-load do script da aba ANTES de marcar data-loaded,
                    // pois o MutationObserver do storage.js reage a data-loaded
                    // e chama loadConfig() que precisa de funções do script da aba
                    // (ex: popularAnosSelector definida em atas.js)
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

        let tabId = tabButton.getAttribute('id')?.replace('-tab', '') || tabButton.getAttribute('aria-controls');

        if (tabId) {
            // Salvar no localStorage apenas se for uma aba da sidebar
            if (tabButton.closest('#sidebar')) {
                localStorage.setItem('lastActiveTab', tabId);
            }

            const navSubtitle = document.getElementById('nav-subtitle');
            const sidebarTabButton = document.getElementById(`${tabId}-tab`);

            if (navSubtitle && sidebarTabButton) {
                const nameText = sidebarTabButton.querySelector('.link-text')?.textContent || '';

                if (tabId === 'home') {
                    navSubtitle.innerHTML = `<span>Compras</span>`;
                } else {
                    navSubtitle.innerHTML = `<span>${nameText}</span>`;
                }
            }

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
                    // Formato esperado: "nomeScript_nomeComputador"
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
initializeNavigation();
collapsePermanently();