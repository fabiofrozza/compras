const darkModeBtn = document.getElementById('darkModeBtn');

function applyDarkMode(darkModeEnabled = null) {

    const isDark = darkModeEnabled ?? !document.body.classList.contains('dark-mode');

    document.body.classList.toggle('dark-mode', isDark);
    darkModeBtn.innerHTML = isDark ? '<i class="fa-solid fa-sun"></i>' : '<i class="fa-solid fa-moon"></i>';

    if (isDark) {
        mainContent.setAttribute('data-bs-theme', 'dark');
    } else {
        mainContent.removeAttribute('data-bs-theme');
    }

    mainContent.classList.toggle('bg-dark', isDark);
    mainContent.classList.toggle('bg-white', !isDark);

}

// ====== BACKGROUND ALEATÓRIO DA HOME ======

async function setHomeBackground() {
    const homePane = document.getElementById('home');
    if (!homePane) return;

    // Buscar URLs das imagens do servidor (.env)
    let homeBackgroundImages = [];
    try {
        const response = await fetch('/api/home-backgrounds');
        const data = await response.json();
        homeBackgroundImages = data.images || [];
    } catch (err) {
        console.error('Erro ao buscar imagens de fundo:', err);
        return;
    }

    if (homeBackgroundImages.length === 0) return;

    const randomIndex = Math.floor(Math.random() * homeBackgroundImages.length);
    const imageUrl = homeBackgroundImages[randomIndex];

    // Pré-carregar a imagem antes de aplicar
    const img = new Image();
    img.onload = () => {
        // Usamos CSS custom property para o ::before consumir via var(--home-bg-url)
        homePane.style.setProperty('--home-bg-url', `url('${imageUrl}')`);
        homePane.classList.add('home-bg-loaded');

        // Atualizar o link de crédito da foto (pode carrregar após lazy load)
        const setCredit = () => {
            const credit = document.getElementById('home-bg-credit');
            if (credit) credit.href = imageUrl;
        };

        setCredit();
        // Observar inserção do elemento caso ainda não exista (lazy load)
        if (!document.getElementById('home-bg-credit')) {
            const observer = new MutationObserver(() => {
                if (document.getElementById('home-bg-credit')) {
                    setCredit();
                    observer.disconnect();
                }
            });
            observer.observe(homePane, { childList: true, subtree: true });
        }
    };
    img.src = imageUrl;
}

// Inicializar background da Home
document.addEventListener('DOMContentLoaded', setHomeBackground);

// Event Delegation para links de navegação (.card-link)
// Permite que funcione mesmo em conteúdo carregado dinamicamente (Lazy Loading)
document.addEventListener('click', (e) => {
    const link = e.target.closest('.card-link');
    if (link) {
        e.preventDefault();
        const target = link.getAttribute('data-bs-target');
        const tabTrigger = document.querySelector(`[data-bs-target="${target}"]`);
        if (tabTrigger) {
            const tab = new bootstrap.Tab(tabTrigger);
            tab.show();
        }
    }
});

function toggleLogsDrawer() {
    const logsDrawer = document.getElementById('logs-drawer');
    if (logsDrawer) {
        logsDrawer.classList.toggle('collapsed');
    }
}

// ====== PAINÉIS DO HEADER (Settings / User) ======

const overlay = document.getElementById('header-panel-overlay');

function openPanel(panelId, triggerBtn) {
    // Fechar qualquer painel aberto antes de abrir o novo
    document.querySelectorAll('.header-panel.open').forEach(p => {
        if (p.id !== panelId) closeAllPanels();
    });

    const panel = document.getElementById(panelId);
    if (!panel) return;

    const isOpen = panel.classList.contains('open');

    closeAllPanels();

    if (!isOpen) {
        panel.classList.add('open');
        panel.setAttribute('aria-hidden', 'false');
        overlay.classList.add('active');
        if (triggerBtn) triggerBtn.classList.add('active');

        // Callbacks específicos por painel
        if (panelId === 'settings-panel') {
            inicializarPreferenciasPanel();
        } else if (panelId === 'user-panel') {
            popularUserPanel();
        } else if (panelId === 'notifications-panel') {
            renderNotificationList();
        }
    }
}

function closeAllPanels() {
    document.querySelectorAll('.header-panel').forEach(p => {
        p.classList.remove('open');
        p.setAttribute('aria-hidden', 'true');
    });
    overlay.classList.remove('active');
    document.querySelectorAll('.header-icon-btn').forEach(b => b.classList.remove('active'));
}

// Preencher painel de usuário
async function popularUserPanel() {
    const nameEl = document.getElementById('userComputerName');
    const syncEl = document.getElementById('lastSyncTime');

    if (nameEl && nameEl.textContent === 'Carregando...') {
        try {
            const response = await fetch('/api/user-info');
            const data = await response.json();
            nameEl.textContent = data.computerName || 'Desconhecido';
        } catch {
            nameEl.textContent = 'Local (Navegador)';
        }
    }
    if (syncEl) {
        syncEl.textContent = new Date().toLocaleString('pt-BR');
    }
}

// Inicializar painel de preferências
function inicializarPreferenciasPanel() {
    if (typeof populateFavoriteTabSelectList === 'function') {
        populateFavoriteTabSelectList();
    }

    loadAppState();
    const prefs = appState.preferences;

    const darkModeCheckbox = document.getElementById('darkMode');
    if (darkModeCheckbox) darkModeCheckbox.checked = prefs.darkMode;

    const sidebarCheckbox = document.getElementById('sidebarExpanded');
    if (sidebarCheckbox) sidebarCheckbox.checked = prefs.sidebarExpanded;

    const tabContainer = document.getElementById('preferredTab');
    if (tabContainer) {
        const val = prefs.preferredTab || '';
        const radio = tabContainer.querySelector(`input[name="preferredTabRadio"][value="${val}"]`);
        if (radio) radio.checked = true;
    }

    // Configurar auto-save nos campos do painel
    if (typeof setupAutoSave === 'function') {
        setupAutoSave(document.getElementById('settings-panel'));
    }
}

// Wiring dos botões
document.addEventListener('DOMContentLoaded', () => {
    const settingsBtn = document.getElementById('settingsBtn');
    const userBtn = document.getElementById('userBtn');
    const notificationsBtn = document.getElementById('notificationsBtn');
    const settingsPanelClose = document.getElementById('settingsPanelClose');
    const userPanelClose = document.getElementById('userPanelClose');
    const notificationsPanelClose = document.getElementById('notificationsPanelClose');

    if (settingsBtn) settingsBtn.addEventListener('click', () => openPanel('settings-panel', settingsBtn));
    if (userBtn) userBtn.addEventListener('click', () => openPanel('user-panel', userBtn));
    if (notificationsBtn) notificationsBtn.addEventListener('click', () => openPanel('notifications-panel', notificationsBtn));
    if (settingsPanelClose) settingsPanelClose.addEventListener('click', closeAllPanels);
    if (userPanelClose) userPanelClose.addEventListener('click', closeAllPanels);
    if (notificationsPanelClose) notificationsPanelClose.addEventListener('click', closeAllPanels);

    // Fechar painel ao clicar no overlay
    if (overlay) overlay.addEventListener('click', closeAllPanels);

    // Fechar com Escape
    document.addEventListener('keydown', (e) => {
        if (e.key === 'Escape') closeAllPanels();
    });
});

/**
 * Limpa todos os campos de um formulário e dispara eventos necessários
 * @param {string} formId O ID do formulário a ser limpo
 */
function limparFormulario(formId) {
    const form = document.getElementById(formId);
    if (!form) return;

    // Busca todos os campos de input, select e textarea dentro do formulário
    const campos = form.querySelectorAll('input, select, textarea');

    campos.forEach(campo => {
        // Ignorar botões e campos escondidos
        if (['button', 'submit', 'reset', 'hidden'].includes(campo.type)) return;

        if (campo.type === 'checkbox' || campo.type === 'radio') {
            campo.checked = false;
        } else {
            campo.value = '';
        }

        // Disparar eventos para que componentes que observam alterações (como numerarAtas) reajam
        campo.dispatchEvent(new Event('input', { bubbles: true }));
        campo.dispatchEvent(new Event('change', { bubbles: true }));
    });

}
