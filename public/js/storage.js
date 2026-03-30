// ====== SISTEMA DE PREFERÊNCIAS E CONFIGURAÇÕES DO USUÁRIO ======

const STORAGE_KEY = 'compras_web_state';
const autoSaveBoundFields = new WeakSet();

let appState = {
    preferences: {
        darkMode: false,
        animatedBg: true,
        enableAnimations: true,
        preferredTab: '',
        bgSource: 'random'
    }
};

function loadAppState() {
    try {
        const saved = localStorage.getItem(STORAGE_KEY);
        if (saved) {
            const parsed = JSON.parse(saved);
            Object.keys(parsed).forEach(key => {
                if (typeof parsed[key] === 'object') {
                    appState[key] = { ...(appState[key] || {}), ...parsed[key] };
                } else {
                    appState[key] = parsed[key];
                }
            });
        }
    } catch (error) {
        console.error('Erro ao ler LocalStorage:', error);
    }
}

function saveAppState() {
    try {
        localStorage.setItem(STORAGE_KEY, JSON.stringify(appState));
    } catch (error) {
        console.error('Erro ao salvar no LocalStorage:', error);
    }
}

// ==== PREFERÊNCIAS DE UI ====

function populateFavoriteTabSelectList() {
    const selectEl = document.getElementById('preferredTab');
    if (!selectEl || typeof TAB_LIST === 'undefined') return;

    const options = [
        { value: '', label: 'Abrir última aba visualizada' },
        ...TAB_LIST.map(tab => ({ value: tab.id, label: tab.label }))
    ];

    selectEl.innerHTML = options.map(opt =>
        `<option value="${opt.value}">${opt.label}</option>`
    ).join('');
}

function showPreferencesSaveIndicator() {
    if (typeof showToast === 'function') {
        showToast('Preferências salvas', 'success', 2000, 'configurações');
    }
}

async function getUserInfo() {
    try {
        const response = await fetch('/api/user-info');
        const data = await response.json();
        return data;
    } catch (error) {
        console.error('Erro ao obter informações do usuário:', error);
        return { computerName: 'Local (Navegador)' };
    }
}

function initPreferencesPage() {
    try {
        loadAppState();
        applyUserPreferences(appState.preferences);
    } catch (error) {
        console.error('Erro ao inicializar página de Preferências:', error);
    }
}

function applyUserPreferences(preferences) {
    if (preferences.darkMode) {
        if (typeof applyDarkMode === 'function') applyDarkMode(true);
        const el = document.getElementById('darkMode');
        if (el) el.checked = true;
    } else {
        if (typeof applyDarkMode === 'function') applyDarkMode(false);
        const el = document.getElementById('darkMode');
        if (el) el.checked = false;
    }

    const animatedBg = preferences.animatedBg !== false;
    document.body.classList.toggle('no-animated-bg', !animatedBg);
    const animatedBgEl = document.getElementById('animatedBg');
    if (animatedBgEl) animatedBgEl.checked = animatedBg;

    const enableAnimations = preferences.enableAnimations !== false;
    document.body.classList.toggle('no-animations', !enableAnimations);
    const enableAnimationsEl = document.getElementById('enableAnimations');
    if (enableAnimationsEl) enableAnimationsEl.checked = enableAnimations;

    const bgSourceSelect = document.getElementById('bgSource');
    if (bgSourceSelect) bgSourceSelect.value = preferences.bgSource || 'random';

    const preferredTabSelect = document.getElementById('preferredTab');
    if (preferredTabSelect) preferredTabSelect.value = preferences.preferredTab || '';

    // Atualizar estrelas de aba preferida nos cards da home
    document.querySelectorAll('.card-glass-star').forEach(star => {
        const isPreferred = preferences.preferredTab && star.dataset.tabId === preferences.preferredTab;
        star.classList.toggle('active', isPreferred);

        const title = isPreferred ? 'Aba preferida' : 'Definir como aba preferida';
        star.setAttribute('data-bs-title', title);
        star.setAttribute('aria-label', title);

        // Atualizar tooltip do Bootstrap se existir
        const tooltip = bootstrap.Tooltip.getInstance(star);
        if (tooltip) {
            tooltip.setContent({ '.tooltip-inner': title });
        }
    });

    // Aplicar navegação (abrir a aba) se estivermos na inicialização
    if (!window.navigationInitialized) {
        window.navigationInitialized = true;
        if (preferences.preferredTab) {
            setTimeout(() => {
                const tabButton = document.getElementById(preferences.preferredTab + '-tab');
                const tabContent = document.getElementById(preferences.preferredTab);
                if (tabButton && tabContent) {
                    const tab = new bootstrap.Tab(tabButton);
                    tab.show();
                }
            }, 500);
        } else {
            const lastTab = localStorage.getItem('lastActiveTab');
            if (lastTab) {
                setTimeout(() => {
                    const tabButton = document.getElementById(lastTab + '-tab') || document.getElementById('' + lastTab + '-tab');
                    const tabContent = document.getElementById(lastTab);
                    if (tabButton && tabContent) {
                        const tab = new bootstrap.Tab(tabButton);
                        tab.show();
                    } else {
                        const homeTab = new bootstrap.Tab(document.getElementById('home-tab'));
                        homeTab.show();
                    }
                }, 500);
            } else {
                setTimeout(() => {
                    const homeTab = new bootstrap.Tab(document.getElementById('home-tab'));
                    homeTab.show();
                }, 500);
            }
        }
    }
}

// ==== CAMPOS DO FORMULÁRIO (CONFIG) ====

let saveTimeout = null;

function setupAutoSave(container) {
    const fields = container.querySelectorAll('[data-field]');
    fields.forEach(field => {
        if (autoSaveBoundFields.has(field)) return;
        autoSaveBoundFields.add(field);
        ['input', 'change'].forEach(eventType => {
            field.addEventListener(eventType, (e) => {
                const target = e.target;
                target.classList.remove('is-invalid');

                const fieldName = target.dataset.field;
                // Radio buttons in a group share the same data-field
                let value;
                if (target.type === 'radio') {
                    if (target.checked) value = target.value;
                    else return; // only save the checked one
                } else {
                    value = target.type === 'checkbox' ? target.checked : target.value;
                }

                // Preferências UI
                if (['darkMode', 'animatedBg', 'enableAnimations', 'preferredTab', 'bgSource'].includes(fieldName)) {
                    appState.preferences[fieldName] = value;
                    saveAppState();
                    applyUserPreferences(appState.preferences);

                    if (eventType === 'change') {
                        showPreferencesSaveIndicator();
                    }
                } else {
                    const tabPane = target.closest('.tab-pane');
                    const tabId = tabPane ? tabPane.id : 'global';

                    if (!appState[tabId]) appState[tabId] = {};
                    appState[tabId][fieldName] = value;

                    if (saveTimeout) clearTimeout(saveTimeout);
                    saveTimeout = setTimeout(() => {
                        saveAppState();
                        if (typeof showToast === 'function') {
                            showToast('Alterações salvas com sucesso!', 'success', 2000, 'configuração');
                        }
                    }, 500);
                }
            });
        });
    });
}

function loadConfig(container = document) {
    try {
        loadAppState();

        if (typeof popularAnosSelector === 'function') {
            popularAnosSelector();
        }

        const scope = container || document;

        scope.querySelectorAll('[data-field]').forEach(field => {
            const fieldName = field.dataset.field;

            // Preferências são aplicadas por applyUserPreferences - não sobrescrever aqui
            if (['darkMode', 'animatedBg', 'preferredTab', 'bgSource'].includes(fieldName)) return;

            const tabPane = field.closest('.tab-pane');
            const tabId = tabPane ? tabPane.id : 'global';

            if (appState[tabId] && appState[tabId].hasOwnProperty(fieldName)) {
                let valor = appState[tabId][fieldName] || '';

                if (field.type === 'checkbox') {
                    field.checked = valor === true || valor === 'true';
                } else if (field.type === 'radio') {
                    field.checked = (field.value === valor);
                } else {
                    field.value = valor;
                }

                field.dispatchEvent(new Event('input', { bubbles: true }));
            }
        });

        setupAutoSave(scope);

    } catch (error) {
        console.error('Erro ao carregar configuração:', error);
    }
}

// ==== STARTUP HOOKS ====

document.addEventListener('DOMContentLoaded', () => {
    // Inicializa a UI (modo escuro e abre aba)
    initPreferencesPage();

    // Observador para detectar lazy load de abas
    const observer = new MutationObserver((mutations) => {
        mutations.forEach((mutation) => {
            if (mutation.type === 'attributes' && mutation.attributeName === 'data-loaded') {
                const target = mutation.target;
                if (target.getAttribute('data-loaded') === 'true') {
                    // Usar timeout para permitir que o navegador renderize o conteúdo antes
                    setTimeout(() => loadConfig(target), 0);
                }
            }
        });
    });

    document.querySelectorAll('.tab-pane').forEach(pane => {
        observer.observe(pane, { attributes: true });
    });
});
