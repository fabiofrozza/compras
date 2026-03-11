const darkModeBtn = document.getElementById('darkModeBtn');

function applyDarkMode(darkModeEnabled = null) {

    const isDark = darkModeEnabled ?? !document.body.classList.contains('dark-mode');

    document.body.classList.toggle('dark-mode', isDark);
    darkModeBtn.innerHTML = isDark ? '<i class="material-symbols-outlined">light_mode</i>' : '<i class="material-symbols-outlined">dark_mode</i>';

    const mc = document.getElementById('mainContent');
    if (isDark) {
        mc.setAttribute('data-bs-theme', 'dark');
    } else {
        mc.removeAttribute('data-bs-theme');
    }

    mc.classList.toggle('bg-dark', isDark);
    mc.classList.toggle('bg-white', !isDark);

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

// ====== FILTRO SVG ALEATÓRIO (Liquid Glass) ======

function randomizeGlassFilter() {
    const totalFilters = 6;
    const randomIndex = Math.floor(Math.random() * totalFilters) + 1;
    document.documentElement.style.setProperty('--filter-name', `url(#container-glass-${randomIndex})`);
}

document.addEventListener('DOMContentLoaded', () => {
    setHomeBackground();
    randomizeGlassFilter();
});

// Event delegation para .card-link — funciona mesmo em conteúdo carregado dinamicamente
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

function inicializarPreferenciasPanel() {
    if (typeof populateFavoriteTabSelectList === 'function') {
        populateFavoriteTabSelectList();
    }

    loadAppState();
    const prefs = appState.preferences;

    const darkModeCheckbox = document.getElementById('darkMode');
    if (darkModeCheckbox) darkModeCheckbox.checked = prefs.darkMode;

    const tabContainer = document.getElementById('preferredTab');
    if (tabContainer) {
        const val = prefs.preferredTab || '';
        const radio = tabContainer.querySelector(`input[name="preferredTabRadio"][value="${val}"]`);
        if (radio) radio.checked = true;
    }

    if (typeof setupAutoSave === 'function') {
        setupAutoSave(document.getElementById('settings-panel'));
    }
}

document.addEventListener('DOMContentLoaded', () => {
    const companyLogo = document.getElementById('companyLogo');
    const settingsBtn = document.getElementById('settingsBtn');
    const userBtn = document.getElementById('userBtn');
    const notificationsBtn = document.getElementById('notificationsBtn');
    const appLauncherBtn = document.getElementById('appLauncherBtn');
    const settingsPanelClose = document.getElementById('settingsPanelClose');
    const userPanelClose = document.getElementById('userPanelClose');
    const notificationsPanelClose = document.getElementById('notificationsPanelClose');
    const appLauncherPanelClose = document.getElementById('appLauncherPanelClose');

    if (companyLogo) {
        companyLogo.style.cursor = 'pointer';
        companyLogo.addEventListener('click', () => {
            const homeTab = document.getElementById('home-tab');
            if (homeTab) new bootstrap.Tab(homeTab).show();
        });
    }
    if (settingsBtn) settingsBtn.addEventListener('click', () => openPanel('settings-panel', settingsBtn));
    if (userBtn) userBtn.addEventListener('click', () => openPanel('user-panel', userBtn));
    if (notificationsBtn) notificationsBtn.addEventListener('click', () => openPanel('notifications-panel', notificationsBtn));
    if (appLauncherBtn) appLauncherBtn.addEventListener('click', () => openPanel('app-launcher-panel', appLauncherBtn));
    if (settingsPanelClose) settingsPanelClose.addEventListener('click', closeAllPanels);
    if (userPanelClose) userPanelClose.addEventListener('click', closeAllPanels);
    if (notificationsPanelClose) notificationsPanelClose.addEventListener('click', closeAllPanels);
    if (appLauncherPanelClose) appLauncherPanelClose.addEventListener('click', closeAllPanels);

    if (overlay) overlay.addEventListener('click', closeAllPanels);
    document.addEventListener('keydown', (e) => {
        if (e.key === 'Escape') closeAllPanels();
    });
});

function limparFormulario(formId) {
    const form = document.getElementById(formId);
    if (!form) return;

    const campos = form.querySelectorAll('input, select, textarea');

    campos.forEach(campo => {
        if (['button', 'submit', 'reset', 'hidden'].includes(campo.type)) return;

        if (campo.type === 'checkbox' || campo.type === 'radio') {
            campo.checked = false;
        } else {
            campo.value = '';
        }

        campo.dispatchEvent(new Event('input', { bubbles: true }));
        campo.dispatchEvent(new Event('change', { bubbles: true }));
    });

}
