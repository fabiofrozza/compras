const darkModeBtn = document.getElementById('darkModeBtn');

// ====== CABEÇALHOS DE SCRIPT-FORM ======

function initScriptFormHeaders(container) {
    const headers = container.querySelectorAll('.script-form-header:not([data-initialized])');
    headers.forEach(h4 => {
        const icon = h4.dataset.comprasIcon;
        const title = h4.dataset.comprasTitle;
        const helpText = h4.dataset.comprasHelp;
        const titleId = h4.dataset.comprasTitleId;

        // Localiza o div.collapse antes de inserir qualquer elemento
        const collapseDiv = h4.parentElement.querySelector(':scope > .collapse');
        const collapseId = collapseDiv?.id;

        // Título: ícone + texto
        const titleSpan = document.createElement('span');
        const iconEl = document.createElement('i');
        iconEl.className = 'material-symbols-outlined';
        iconEl.textContent = icon;
        titleSpan.appendChild(iconEl);
        titleSpan.append(' ');
        if (titleId) {
            const textSpan = document.createElement('span');
            textSpan.id = titleId;
            textSpan.textContent = title;
            titleSpan.appendChild(textSpan);
        } else {
            titleSpan.append(title);
        }

        // Container de botões
        const buttonsDiv = document.createElement('div');
        buttonsDiv.className = 'script-form-header-buttons';

        // Botão de ajuda (se data-compras-help presente)
        if (helpText) {
            const helpId = collapseId
                ? collapseId.replace('-collapse', '-help')
                : `help-${crypto.randomUUID().slice(0, 8)}`;
            const anchorName = `--${helpId}`;

            const helpBtn = document.createElement('button');
            helpBtn.type = 'button';
            helpBtn.className = 'form-header-help';
            helpBtn.setAttribute('popovertarget', helpId);
            helpBtn.style.setProperty('anchor-name', anchorName);
            helpBtn.innerHTML = '<i class="material-symbols-outlined">help</i>';
            buttonsDiv.appendChild(helpBtn);

            const popoverDiv = document.createElement('div');
            popoverDiv.setAttribute('popover', '');
            popoverDiv.id = helpId;
            popoverDiv.className = 'form-help-popover';
            popoverDiv.style.setProperty('position-anchor', anchorName);
            popoverDiv.innerHTML = helpText;
            h4.after(popoverDiv);
        }

        // Botão de colapsar (sempre presente quando há div.collapse)
        if (collapseId) {
            const collapseBtn = document.createElement('button');
            collapseBtn.type = 'button';
            collapseBtn.className = 'form-collapse-btn';
            collapseBtn.setAttribute('data-bs-toggle', 'collapse');
            collapseBtn.setAttribute('data-bs-target', `#${collapseId}`);
            collapseBtn.setAttribute('aria-expanded', 'true');
            collapseBtn.setAttribute('aria-controls', collapseId);
            collapseBtn.innerHTML = '<i class="material-symbols-outlined"></i>';
            buttonsDiv.appendChild(collapseBtn);
        }

        h4.innerHTML = '';
        h4.appendChild(titleSpan);
        h4.appendChild(buttonsDiv);
        h4.setAttribute('data-initialized', 'true');
        ['comprasIcon', 'comprasTitle', 'comprasHelp', 'comprasTitleId'].forEach(attr => delete h4.dataset[attr]);
    });
}

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
        icon.textContent = 'fullscreen_exit';
        btn.setAttribute('data-bs-title', 'Sair da tela inteira');
    } else {
        icon.textContent = 'fullscreen';
        btn.setAttribute('data-bs-title', 'Tela inteira');
    }
    const tooltip = bootstrap.Tooltip.getInstance(btn);
    if (tooltip) tooltip.dispose();
    new bootstrap.Tooltip(btn, {
        container: 'body',
        trigger: 'hover',
        customClass: 'custom-tooltip'
    });
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
            if (bsTooltip) bsTooltip.dispose();
            new bootstrap.Tooltip(credit, {
                container: 'body',
                trigger: 'hover',
                customClass: 'custom-tooltip'
            });
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

document.addEventListener('DOMContentLoaded', async () => {
    setHomeBackground();
    randomizeGlassFilter();

    try {
        const cfg = await fetch('/api/app-config').then(r => r.json());
        const refreshMs = (cfg.backgroundRefreshTime || 0) * 1000;
        if (refreshMs > 0) setInterval(setHomeBackground, refreshMs);
    } catch (_) { }
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
        { separator: true },
        { icon: 'build', label: 'Instalação', tabId: 'instalacao' },
    ];

    const ul = document.createElement('ul');
    ul.className = 'help-link-list';

    for (const item of links) {
        const li = document.createElement('li');
        if (item.separator) {
            li.innerHTML = '<hr>';
        } else if (item.tabId) {
            const btn = document.createElement('button');
            btn.dataset.tabId = item.tabId;
            btn.innerHTML = `<i class="material-symbols-outlined">${item.icon}</i>${item.label}`;
            btn.addEventListener('click', () => {
                const tabButton = document.getElementById(`${item.tabId}-tab`);
                if (tabButton) new bootstrap.Tab(tabButton).show();
                closeAllPanels();
            });
            li.appendChild(btn);
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

// Bootstrap não remove o elemento .tooltip do body quando o trigger é
// removido/recriado dinamicamente sem chamar dispose(), o que pode deixar
// tooltips "fantasmas" na tela. Esta varredura periódica remove tooltips
// cujo trigger (referenciado via aria-describedby) não existe mais no DOM.
function cleanupOrphanTooltips() {
    // Busca tooltips em qualquer lugar do documento, não apenas no body direto
    document.querySelectorAll('.tooltip').forEach(tooltipEl => {
        const id = tooltipEl.id;
        // Usa ~= para encontrar o id em uma lista (aria-describedby pode ter vários ids)
        const trigger = id ? document.querySelector(`[aria-describedby~="${id}"]`) : null;

        // Se o trigger não existe, não está no DOM ou está oculto (ex: aba trocada), removemos a tooltip
        if (!trigger || !document.body.contains(trigger) || trigger.getClientRects().length === 0) {
            // Tenta dispose via Bootstrap se o trigger ainda existir, senão remove direto o elemento
            const instance = trigger ? bootstrap.Tooltip.getInstance(trigger) : null;
            if (instance) {
                instance.dispose();
            } else {
                tooltipEl.remove();
            }
        }
    });
}

setInterval(cleanupOrphanTooltips, 1000);

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
