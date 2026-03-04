// SIDEBAR TOGGLE CONTROLS

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

// Mapeamento de abas para scripts sob demanda
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

// Troca o ícone do botão entre hamburguer e X
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

// --- Funções auxiliares ---
// Helper: toggle classes em múltiplos elementos
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

    // Botão vira X e fica visível; logo empresa some do slot fixo e aparece no container expandido
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

    // Volta ao estado inicial: logo empresa no slot fixo, expanded logo escondido, botão hamburguer escondido
    toggleClasses(companyLogo, null, 'd-none');
    toggleClasses(companyLogoExpanded, 'd-none');
    toggleClasses(toggleSidebarBtnClose, 'd-none');
    setToggleIcon(false);
    toggleSidebarBtnClose.setAttribute('data-bs-title', 'Fixar barra lateral');
}

// Expandir sidebar
function expandSidebar() {
    expandPermanently();
}

// Colapsar sidebar
function collapseSidebar() {
    collapsePermanently();
}

// --- Hover: expande temporariamente apenas se NÃO estiver permanentemente expandido ---
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
    // Se é um hover temporário e não está fixo, recolhe
    if (!isPermanentlyExpanded && sidebar.classList.contains('hover-expand')) {
        toggleClasses([sidebar, sidebarShadow, roundCorner], 'collapsed', 'expanded');
        sidebar.classList.remove('hover-expand');
        toggleClasses(companyLogo, null, 'd-none');
        toggleClasses(toggleSidebarBtnClose, 'd-none');
    }
});

sidebar.addEventListener('click', (event) => {
    // Verificar se o clique foi em um elemento interativo (botão, link, etc)
    const isInteractiveElement = event.target.closest('button, a, .nav-item, .nav-link');

    // Se o clique foi no espaço vazio (não em elemento interativo), executar toggle
    if (!isInteractiveElement) {
        if (isPermanentlyExpanded) {
            collapsePermanently();
        } else {
            expandPermanently();
        }
    }
});

// Rastrear mudanças nas abas e 
// carregar arquivos da aba atual e atualizar o ícone e nome da aba na navbar
document.addEventListener('shown.bs.tab', async (event) => {
    if (event.target) {
        // --- Lazy Loading: Carregar conteúdo HTML se necessário ---
        const tabButton = event.target;
        const targetSelector = tabButton.getAttribute('data-bs-target');
        if (targetSelector) {
            const targetPane = document.querySelector(targetSelector);
            // Verifica se tem URL configurada e se ainda NÃO foi carregado
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
                    // Lazy-load other.js para exibir catLoader (animação de erro)
                    if (typeof catLoader === 'function') {
                        msgError += catLoader();
                    } else {
                        try {
                            await loadScript('js/other.js');
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

            // Atualizar o ícone e nome da aba na navbar obtendo as informações da sidebar
            const navSubtitle = document.getElementById('nav-subtitle');
            const sidebarTabButton = document.getElementById(`${tabId}-tab`);

            if (navSubtitle && sidebarTabButton) {
                //const iconClass = sidebarTabButton.querySelector('i')?.className || '';
                const nameText = sidebarTabButton.querySelector('.link-text')?.textContent || '';

                if (tabId === 'home') {
                    navSubtitle.innerHTML = `<span>Compras</span>`;
                } else {
                    navSubtitle.innerHTML = `<span>${nameText}</span>`;
                }
            }

            // Carregar/atualizar arquivos da aba que foi ativada
            await refreshScriptFileLists(tabId);

            // Configurar validação em tempo real
            setupLiveValidation(tabId);

            // Validar campos da aba ativada
            validateTabFields(tabId);

            // Criar tooltips para campos obrigatórios
            createRequiredFieldsTooltip();

            if (tabId === 'atas' && typeof inicializarAtas === 'function') {
                inicializarAtas();
            }

            if (tabId === 'importacao' && typeof inicializarImportacao === 'function') {
                inicializarImportacao();
            }


            // Atualiza os logs após a troca de aba
            if (globalComputerName && (tabId === 'atas' || tabId === 'catmat' || tabId === 'importacao' || tabId === 'mapas' || tabId === 'powerbi')) {
                // Garante que o DOM do console/logs-drawer exista
                if (typeof ensureConsoleDOM === 'function') ensureConsoleDOM();

                const logsList = document.getElementById('logs-file-list');
                const logsDrawer = document.getElementById('logs-drawer');

                if (logsList && logsDrawer) {
                    // Formato esperado de log: "nomeScript_nomeComputador" ou variações com '_'
                    let logsNameFilter = `${tabId}_${globalComputerName}`.toLowerCase();
                    logsList.dataset.nameContains = logsNameFilter;

                    // Mostra o drawer se estiver oculto
                    logsDrawer.classList.remove('d-none');

                    // Carrega os arquivos filtrados de log
                    loadFiles('logs-file-list', '_common', 'log', false);
                }
            } else {
                // Esconde drawer em home (se existir)
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

// --- Estado inicial: colapsado, sem expansão permanente ---
collapsePermanently(); // já define collapsed = true, isPermanentlyExpanded = false