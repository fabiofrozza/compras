const darkModeBtn = document.getElementById('darkModeBtn');

// ====== TELA INTEIRA ======

function toggleFullscreen() {
    if (!document.fullscreenElement) {
        document.documentElement.requestFullscreen();
    } else {
        document.exitFullscreen();
    }
}

document.addEventListener('fullscreenchange', () => {
    const btn = document.getElementById('fullscreenBtn');
    if (!btn) return;
    const icon = btn.querySelector('i');
    if (document.fullscreenElement) {
        icon.textContent = 'minimize';
        btn.setAttribute('data-bs-title', 'Sair da tela inteira');
    } else {
        icon.textContent = 'open_in_full';
        btn.setAttribute('data-bs-title', 'Tela inteira');
    }
    const tooltip = bootstrap.Tooltip.getInstance(btn);
    if (tooltip) tooltip.dispose();
    new bootstrap.Tooltip(btn);
});

function applyDarkMode(darkModeEnabled = null) {

    const isDark = darkModeEnabled ?? !document.body.classList.contains('dark');

    document.body.classList.toggle('dark', isDark);
    darkModeBtn.innerHTML = isDark ? '<i class="material-symbols-outlined">light_mode</i>' : '<i class="material-symbols-outlined">dark_mode</i>';

    if (isDark) {
        document.documentElement.setAttribute('data-bs-theme', 'dark');
        document.documentElement.setAttribute('data-theme', 'dark');
    } else {
        document.documentElement.removeAttribute('data-bs-theme');
        document.documentElement.removeAttribute('data-theme')
    }

}

// ====== BACKGROUND ALEATÓRIO DA HOME ======

// bgSource: 'local' | 'bing' | 'random'
async function setHomeBackground(bgSource) {
    const homePane = document.getElementById('home');
    if (!homePane) return;

    if (bgSource === undefined) {
        loadAppState();
        bgSource = appState.preferences.bgSource || 'random';
    }

    const [localImages, bingImages] = await Promise.all([
        bgSource !== 'bing'
            ? fetch('/api/home-backgrounds').then(r => r.json()).then(d => d.images || []).catch(() => [])
            : Promise.resolve([]),
        bgSource !== 'local'
            ? fetch('https://peapix.com/bing/feed?country=br').then(r => r.json()).then(d => Array.isArray(d) ? d : []).catch(() => [])
            : Promise.resolve([])
    ]);

    const hasLocal = localImages.length > 0;
    const hasBing = bingImages.length > 0;

    let imageUrl, creditHref, creditTooltip;

    const pickLocal = () => {
        const url = localImages[Math.floor(Math.random() * localImages.length)];
        return { imageUrl: url, creditHref: url, creditTooltip: 'Ver foto original' };
    };

    const pickBing = () => {
        const item = bingImages[Math.floor(Math.random() * bingImages.length)];
        return { imageUrl: item.fullUrl, creditHref: item.pageUrl, creditTooltip: `${item.title} - ${item.copyright}` };
    };

    if (bgSource === 'local' || (bgSource === 'random' && !hasBing)) {
        if (!hasLocal) return;
        ({ imageUrl, creditHref, creditTooltip } = pickLocal());
    } else if (bgSource === 'bing' || (bgSource === 'random' && !hasLocal)) {
        if (!hasBing) return;
        ({ imageUrl, creditHref, creditTooltip } = pickBing());
    } else {
        // random com ambas as fontes disponíveis
        ({ imageUrl, creditHref, creditTooltip } = Math.random() < 0.5 ? pickBing() : pickLocal());
    }

    // Pré-carregar a imagem antes de aplicar
    const img = new Image();
    img.onload = () => {
        // Usamos CSS custom property para o ::before consumir via var(--home-bg-url)
        homePane.style.setProperty('--home-bg-url', `url('${imageUrl}')`);
        homePane.classList.add('home-bg-loaded');

        // Atualizar o link e tooltip de crédito da foto (pode carregar após lazy load)
        const applyCredit = () => {
            const credit = document.getElementById('home-bg-credit');
            if (!credit) return;
            credit.href = creditHref;
            credit.setAttribute('data-bs-title', creditTooltip);
            const bsTooltip = bootstrap.Tooltip.getInstance(credit);
            if (bsTooltip) {
                bsTooltip.dispose();
                new bootstrap.Tooltip(credit);
            }
        };

        applyCredit();
        if (!document.getElementById('home-bg-credit')) {
            const observer = new MutationObserver(() => {
                if (document.getElementById('home-bg-credit')) {
                    applyCredit();
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
    document.documentElement.style.setProperty('--compras-filter', `url(#container-glass-${randomIndex})`);
}

document.addEventListener('DOMContentLoaded', () => {
    setHomeBackground();
    randomizeGlassFilter();
});

// Event delegation para .card-link - funciona mesmo em conteúdo carregado dinamicamente
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

function openPanel(panelId, triggerBtn) {
    const panel = document.getElementById(panelId);
    if (!panel) return;

    const isOpen = panel.classList.contains('open');

    closeAllPanels();

    if (!isOpen) {
        panel.classList.add('open');
        panel.setAttribute('aria-hidden', 'false');
        if (triggerBtn) triggerBtn.classList.add('active');

        if (panelId === 'settings-panel') {
            inicializarPreferenciasPanel();
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
    document.querySelectorAll('.header-icon-btn').forEach(b => b.classList.remove('active'));
}

async function inicializarPreferenciasPanel() {
    if (typeof populateFavoriteTabSelectList === 'function') {
        populateFavoriteTabSelectList();
    }

    loadAppState();
    const prefs = appState.preferences;

    const darkModeCheckbox = document.getElementById('darkMode');
    if (darkModeCheckbox) darkModeCheckbox.checked = prefs.darkMode;

    const animatedBgCheckbox = document.getElementById('animatedBg');
    if (animatedBgCheckbox) animatedBgCheckbox.checked = prefs.animatedBg !== false;

    const bgSourceSelect = document.getElementById('bgSource');
    if (bgSourceSelect) bgSourceSelect.value = prefs.bgSource || 'random';

    const preferredTabSelect = document.getElementById('preferredTab');
    if (preferredTabSelect) preferredTabSelect.value = prefs.preferredTab || '';

    const syncEl = document.getElementById('lastSyncTime');
    if (syncEl) syncEl.textContent = new Date().toLocaleString('pt-BR');

    const userNameEl = document.getElementById('userName');
    const computerNameEl = document.getElementById('userComputerName');
    if (userNameEl && userNameEl.textContent === 'Carregando...') {
        try {
            const res = await fetch('/api/user-info');
            const data = await res.json();
            userNameEl.textContent = data.userName || 'Desconhecido';
            if (computerNameEl) computerNameEl.textContent = data.computerName || '';
        } catch {
            userNameEl.textContent = 'Local (Navegador)';
        }
    }

    if (typeof setupAutoSave === 'function') {
        setupAutoSave(document.getElementById('settings-panel'));
    }
}

document.addEventListener('DOMContentLoaded', () => {
    const companyLogo = document.getElementById('companyLogo');
    const fullscreenBtn = document.getElementById('fullscreenBtn');
    const settingsBtn = document.getElementById('settingsBtn');
    const notificationsBtn = document.getElementById('notificationsBtn');
    const helpBtn = document.getElementById('helpBtn');
    const appLauncherBtn = document.getElementById('appLauncherBtn');
    const settingsPanelClose = document.getElementById('settingsPanelClose');
    const notificationsPanelClose = document.getElementById('notificationsPanelClose');
    const helpPanelClose = document.getElementById('helpPanelClose');
    const appLauncherPanelClose = document.getElementById('appLauncherPanelClose');

    if (companyLogo) {
        companyLogo.style.cursor = 'pointer';
        companyLogo.addEventListener('click', () => {
            const homeTab = document.getElementById('home-tab');
            if (homeTab) new bootstrap.Tab(homeTab).show();
        });
    }
    if (fullscreenBtn) fullscreenBtn.addEventListener('click', toggleFullscreen);
    if (settingsBtn) settingsBtn.addEventListener('click', () => openPanel('settings-panel', settingsBtn));
    if (notificationsBtn) notificationsBtn.addEventListener('click', () => openPanel('notifications-panel', notificationsBtn));
    if (helpBtn) helpBtn.addEventListener('click', () => openPanel('help-panel', helpBtn));
    if (appLauncherBtn) appLauncherBtn.addEventListener('click', () => openPanel('app-launcher-panel', appLauncherBtn));
    if (settingsPanelClose) settingsPanelClose.addEventListener('click', closeAllPanels);
    if (notificationsPanelClose) notificationsPanelClose.addEventListener('click', closeAllPanels);
    if (helpPanelClose) helpPanelClose.addEventListener('click', closeAllPanels);
    if (appLauncherPanelClose) appLauncherPanelClose.addEventListener('click', closeAllPanels);

    document.addEventListener('click', (e) => {
        if (!document.querySelector('.header-panel.open')) return;
        if (e.target.closest('.header-panel') || e.target.closest('.header-icon-btn')) return;
        closeAllPanels();
    });
    document.addEventListener('keydown', (e) => {
        if (e.key === 'Escape') closeAllPanels();
    });

    buildHelpPanel();
});

async function buildHelpPanel() {
    const body = document.getElementById('help-panel-body');
    if (!body) return;

    let cfg = {};
    try {
        const res = await fetch('/api/app-config');
        cfg = await res.json();
    } catch (err) {
        console.error('Erro ao buscar configurações:', err);
    }

    const links = [
        { icon: 'business', label: cfg.companyName || 'Site da Empresa', url: cfg.companySite || '#' },
        { icon: 'store', label: cfg.departmentName || 'Site do Departamento', url: cfg.departmentSite || '#' },
        { icon: 'article', label: 'Manual do Departamento', url: cfg.manualSite || '#' },
        { separator: true },
        { icon: 'code', label: 'Repositório', url: 'https://github.com/fabiofrozza/compras' },
    ];

    const ul = document.createElement('ul');
    ul.className = 'help-link-list';

    for (const item of links) {
        const li = document.createElement('li');
        if (item.separator) {
            li.innerHTML = '<hr>';
        } else {
            const a = document.createElement('a');
            a.href = item.url;
            a.target = '_blank';
            a.rel = 'noopener noreferrer';
            a.innerHTML = `<i class="material-symbols-outlined">${item.icon}</i>${item.label}`;
            li.appendChild(a);
        }
        ul.appendChild(li);
    }

    body.appendChild(ul);
}

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
