const TOOLTIP_DEFAULTS = {
    container: 'body',
    trigger: 'hover',
    html: true,
    customClass: 'custom-tooltip',
    delay: { show: 300, hide: 100 }
};

function showToast(message, type = 'success', duration, source = 'origem não informada') {
    const container = document.getElementById('toast-container');
    const defaultDurations = {
        'success': 3000,
        'error': 10000,
        'warning': 7000,
        'info': 5000
    };
    duration = duration || defaultDurations[type] || 5000;

    const toast = document.createElement('div');
    toast.className = `app-toast ${type} shadow`;

    const messageContainer = document.createElement('span');
    messageContainer.className = 'toast-message';
    messageContainer.textContent = message;

    const closeButton = document.createElement('button');
    closeButton.className = 'toast-close-btn';
    closeButton.innerHTML = '<i class="material-symbols-outlined">close</i>';

    const removeToast = () => {
        toast.classList.add('removing');
        setTimeout(() => {
            toast.remove();
        }, 300); // Aguardar a animação de saída
    };

    closeButton.addEventListener('click', removeToast);

    toast.style.setProperty('--compras-toast-duration', `${duration}ms`);
    toast.appendChild(messageContainer);
    toast.appendChild(closeButton);
    container.appendChild(toast);

    setTimeout(removeToast, duration);
}

function getLabelIconsContainer(labelEl) {
    let container = labelEl.querySelector(':scope > .label-icons');
    if (!container) {
        container = document.createElement('div');
        container.className = 'label-icons';
        labelEl.appendChild(container);
    }
    return container;
}

function transformLabelTooltips(container = document) {
    container.querySelectorAll('label[data-bs-toggle="tooltip"]').forEach(label => {
        const title = label.getAttribute('data-bs-title');
        if (!title) return;

        // Descarta instância prévia, caso initializeTooltips tenha rodado antes
        const existing = bootstrap.Tooltip.getInstance(label);
        if (existing) existing.dispose();

        label.removeAttribute('data-bs-toggle');
        label.removeAttribute('data-bs-title');

        const icon = document.createElement('i');
        icon.className = 'material-symbols-outlined form-label-help';
        icon.textContent = 'help';
        icon.setAttribute('data-bs-toggle', 'tooltip');
        icon.setAttribute('data-bs-title', title);
        getLabelIconsContainer(label).appendChild(icon);
    });
}

function createRequiredFieldsTooltip(container = document) {
    transformLabelTooltips(container);
    const requiredFields = container.querySelectorAll('label:has(+ [data-validate-rule]):not(:has(.indicator-required)), label:has(+ :required):not(:has(.indicator-required)), label:has(+ * :required):not(:has(.indicator-required)), h4:has(~ [data-validate-rule], ~ * [data-validate-rule], ~ :required, ~ * :required):not(:has(.indicator-required))');
    requiredFields.forEach(field => {
        const asteriskSpan = document.createElement('span');
        asteriskSpan.className = 'indicator-required';
        asteriskSpan.dataset.bsTitle = 'Campo obrigatório';
        asteriskSpan.dataset.bsToggle = 'tooltip';
        asteriskSpan.innerHTML = '<i class="material-symbols-outlined">emergency</i>';

        if (field.tagName === 'LABEL') {
            getLabelIconsContainer(field).appendChild(asteriskSpan);
        } else {
            const titleSpan = field.querySelector(':scope > span');
            (titleSpan || field).appendChild(asteriskSpan);
        }
    });

    initializeTooltips(container);
}

function initializeTooltips(container = document) {
    const tooltipTriggerList = container.querySelectorAll('[data-bs-toggle="tooltip"]');
    tooltipTriggerList.forEach(tooltipTriggerEl => {
        // Evita inicializar múltiplas vezes o mesmo elemento
        if (!bootstrap.Tooltip.getInstance(tooltipTriggerEl)) {
            try {
                // Previne erro do Bootstrap 5 ao inicializar se title não estiver propriamente atrelado
                const t = tooltipTriggerEl.getAttribute('title');
                const dt = tooltipTriggerEl.getAttribute('data-bs-title');
                const dtot = tooltipTriggerEl.getAttribute('data-bs-original-title');
                const isTitleEmpty = (str) => !str || str.trim() === '';

                if (isTitleEmpty(t) && isTitleEmpty(dt) && isTitleEmpty(dtot)) {
                    return; // Ignora inicialização silenciosamente (title é nulo ou inteiramente vazio)
                }
                new bootstrap.Tooltip(tooltipTriggerEl, { ...TOOLTIP_DEFAULTS });
            } catch (e) {
                console.warn('Tooltips: ignoring initialization for', tooltipTriggerEl, e);
            }
        }
    });
}

function removeTooltip(element) {
    // Ocultar e destruir tooltip antes de remover o elemento do DOM
    const tooltip = bootstrap.Tooltip.getInstance(element);
    if (tooltip) {
        tooltip.hide();
        tooltip.dispose();
    }
}
transformLabelTooltips();
initializeTooltips();

function testToasts() {
    showToast('Testando toast de sucesso', 'success', 15000)
    showToast('Testando toast de erro', 'error', 11000)
    showToast('Testando toast de aviso', 'warning', 12000)
    showToast('Testando toast de informação', 'info', 10000)
}

// ==========================================================================
// CENTRAL DE NOTIFICAÇÕES (Painel no Header)
// ==========================================================================

/** @type {{ id: number, message: string, type: string, source: string, timestamp: Date }[]} */
const notifications = [];
let notificationIdCounter = 0;
const notificationActionCallbacks = new Map();

const notificationIcons = {
    success: 'check_circle',
    error: 'cancel',
    warning: 'warning',
    info: 'info'
};

/**
 * Adiciona uma notificação à central de notificações no header.
 * @param {Object} options
 * @param {string} options.message - Mensagem da notificação
 * @param {'success'|'error'|'warning'|'info'} [options.type='info'] - Tipo da notificação
 * @param {string} [options.source='Sistema'] - Aba ou serviço de origem
 * @param {Array<{label: string, callback: Function}>} [options.actions] - Botões de ação opcionais
 */
async function addNotification({ message, type = 'info', source = 'Sistema', actions = null }) {
    const notification = {
        id: ++notificationIdCounter,
        message,
        type,
        source,
        timestamp: new Date(),
        actions: null
    };

    if (actions && actions.length > 0) {
        notification.actions = actions.map((a, idx) => ({ label: a.label, idx }));
        actions.forEach((a, idx) => {
            notificationActionCallbacks.set(`${notification.id}_${idx}`, a.callback);
        });
    }

    notifications.unshift(notification); // Mais recente no topo
    easterEggNotification();

    const delay = (ms) => new Promise(res => setTimeout(res, ms));
    await delay(1000);

    updateNotificationBadge();
    renderNotificationList();

    return notification.id;
}

let _easterEggsLoadersReady = false;
async function easterEggNotification() {
    if (document.body.classList.contains('no-animations')) return;

    const btn = document.getElementById('notificationsBtn');
    const headerActions = document.getElementById('header-actions');
    if (!btn) return;

    // Lazy-load do loaders.js se ainda não carregado
    if (!_easterEggsLoadersReady && typeof getMario !== 'function') {
        try {
            await loadScript('js/loaders.js');
            _easterEggsLoadersReady = true;
        } catch (e) {
            return;
        }
    }

    const useLink = Math.random() < 0.5;

    const container = document.createElement('div');
    container.className = 'easter-egg-container';

    if (useLink) {
        // Link idle por 2s, depois troca para espada levantada e empurra o header
        container.classList.add('link-character');
        container.innerHTML = getLinkIdle();
        btn.appendChild(container);
        requestAnimationFrame(() => requestAnimationFrame(() => container.classList.add('visible')));

        setTimeout(() => {
            container.innerHTML = getLinkSword();
            headerActions.classList.add('bell-bump');
            setTimeout(() => headerActions.classList.remove('bell-bump'), 1000);
        }, 1000);

        // Fade-out e remoção após idle (2s) + exibição da espada (1s)
        setTimeout(() => {
            container.classList.remove('visible');
            setTimeout(() => container.remove(), 300);
        }, 1700);
    } else {
        container.innerHTML = getMario();
        btn.appendChild(container);
        requestAnimationFrame(() => requestAnimationFrame(() => container.classList.add('visible')));

        // Sincronizar o "pulo" do sino com a cabeçada do Mario
        setTimeout(() => {
            headerActions.classList.add('bell-bump');
            setTimeout(() => headerActions.classList.remove('bell-bump'), 400);
        }, 500);

        // Fade-out e remoção após um ciclo da animação original (2s)
        setTimeout(() => {
            container.classList.remove('visible');
            setTimeout(() => container.remove(), 300);
        }, 1700);
    }
}

/** Atualiza o badge vermelho no botão do sino */
function updateNotificationBadge() {
    const btn = document.getElementById('notificationsBtn');
    if (!btn) return;

    // Remove badge existente
    const existingBadge = btn.querySelector('.notification-badge');
    if (existingBadge) existingBadge.remove();

    const count = notifications.length;
    if (count > 0) {
        const badge = document.createElement('span');
        badge.className = 'notification-badge';
        badge.textContent = count > 9 ? '9+' : count;
        btn.appendChild(badge);
    }
}

/** Renderiza a lista de notificações no painel */
function renderNotificationList() {
    const list = document.getElementById('notification-list');
    const footer = document.getElementById('notification-footer');
    if (!list) return;

    list.innerHTML = '';

    if (notifications.length === 0) {
        list.innerHTML = `
            <div class="notification-empty">
                <i class="material-symbols-outlined">notifications_off</i>
                <span>Nenhuma notificação</span>
            </div>`;
        if (footer) footer.style.display = 'none';
        return;
    }

    if (footer) footer.style.display = 'flex';

    notifications.forEach(notif => {
        const item = document.createElement('div');
        item.className = `notification-item ${notif.type}`;
        item.dataset.notifId = notif.id;

        const iconClass = notificationIcons[notif.type] || notificationIcons.info;
        const timeStr = notif.timestamp.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' });

        item.innerHTML = `
            <div class="notification-icon"><i class="material-symbols-outlined">${iconClass}</i></div>
            <div class="notification-body">
                <div class="notification-message">${notif.message}</div>
                <div class="notification-meta">
                    <span class="notification-time"><i class="material-symbols-outlined me-1">schedule</i>${timeStr}</span>
                    <span class="notification-source">${notif.source}</span>
                </div>
            </div>
            <button class="notification-dismiss" aria-label="Dispensar" data-notif-id="${notif.id}">
                <i class="material-symbols-outlined">close</i>
            </button>`;

        if (notif.actions && notif.actions.length > 0) {
            const actionsEl = document.createElement('div');
            actionsEl.className = 'notification-actions';
            notif.actions.forEach((a, i) => {
                const btn = document.createElement('button');
                const isLast = i === notif.actions.length - 1;
                btn.className = `btn btn-sm ${isLast ? 'btn-primary' : 'btn-secondary'}`;
                btn.dataset.notifId = notif.id;
                btn.dataset.actionIdx = a.idx;
                btn.textContent = a.label;
                actionsEl.appendChild(btn);
            });
            item.querySelector('.notification-body').appendChild(actionsEl);
        }

        list.appendChild(item);
    });
}

/** Remove callbacks de ação associados a uma notificação */
function _cleanupNotificationCallbacks(id) {
    for (const key of notificationActionCallbacks.keys()) {
        if (key.startsWith(`${id}_`)) notificationActionCallbacks.delete(key);
    }
}

/** Remove uma notificação individual pelo ID */
function dismissNotification(id) {
    const item = document.querySelector(`.notification-item[data-notif-id="${id}"]`);

    if (item) {
        item.classList.add('removing');
        setTimeout(() => {
            const index = notifications.findIndex(n => n.id === id);
            if (index !== -1) notifications.splice(index, 1);
            _cleanupNotificationCallbacks(id);
            updateNotificationBadge();
            renderNotificationList();
            if (notifications.length === 0 && typeof closeAllPanels === 'function') closeAllPanels();
        }, 200);
    } else {
        const index = notifications.findIndex(n => n.id === id);
        if (index !== -1) notifications.splice(index, 1);
        _cleanupNotificationCallbacks(id);
        updateNotificationBadge();
        renderNotificationList();
        if (notifications.length === 0 && typeof closeAllPanels === 'function') closeAllPanels();
    }
}

/** Remove todas as notificações */
function clearAllNotifications() {
    notifications.length = 0;
    notificationActionCallbacks.clear();
    updateNotificationBadge();
    renderNotificationList();
    if (typeof closeAllPanels === 'function') closeAllPanels();
}

// Event delegation para dismiss de notificações individuais
document.addEventListener('click', (e) => {
    const dismissBtn = e.target.closest('.notification-dismiss');
    if (dismissBtn) {
        const id = parseInt(dismissBtn.dataset.notifId, 10);
        if (!isNaN(id)) dismissNotification(id);
    }

    const actionBtn = e.target.closest('.btn');
    if (actionBtn) {
        const notifId = parseInt(actionBtn.dataset.notifId, 10);
        const actionIdx = parseInt(actionBtn.dataset.actionIdx, 10);
        const cb = notificationActionCallbacks.get(`${notifId}_${actionIdx}`);
        if (cb) cb();
    }
});

document.addEventListener('DOMContentLoaded', () => {
    const clearAllBtn = document.getElementById('notification-clear-all');
    if (clearAllBtn) {
        clearAllBtn.addEventListener('click', clearAllNotifications);
    }
});

renderNotificationList();

/**
 * Exibe um modal de confirmação genérico
 * @param {Object} options - Opções de configuração
 * @returns {Promise<boolean>} - Retorna true se confirmado
 */
function showConfirmationModal({
    title = 'Confirmar',
    message = 'Tem certeza?',
    detail = null,
    confirmText = 'Confirmar',
    confirmColor = 'btn-danger',
    icon = 'warning'
} = {}) {
    return new Promise((resolve) => {
        const modalId = 'genericConfirmModal';
        const existingModal = document.getElementById(modalId);
        if (existingModal) existingModal.remove();

        const modalHtml = `
            <div class="modal fade" id="${modalId}" tabindex="-1" aria-hidden="true">
                <div class="modal-dialog modal-dialog-centered">
                    <div class="modal-content">
                        <div class="modal-header">
                            <h5 class="modal-title ${confirmColor === 'btn-danger' ? 'text-danger' : ''}">
                                <i class="material-symbols-outlined me-2">${icon}</i>${title}
                            </h5>
                            <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
                        </div>
                        <div class="modal-body">
                            <p>${message}</p>
                            <div id="${modalId}-detail-container" class="modal-detail p-2 border rounded mb-3 text-break small" style="display: none;">
                                <span id="${modalId}-detail-text"></span>
                            </div>
                            ${confirmColor === 'btn-danger' ? '<p class="text-danger mb-0 fw-bold">Esta ação não pode ser desfeita!</p>' : ''}
                        </div>
                        <div class="modal-footer">
                            <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Cancelar</button>
                            <button type="button" class="btn ${confirmColor}" id="${modalId}-btn">${confirmText}</button>
                        </div>
                    </div>
                </div>
            </div>`;

        document.body.insertAdjacentHTML('beforeend', modalHtml);
        const modalEl = document.getElementById(modalId);

        if (detail) {
            document.getElementById(`${modalId}-detail-container`).style.display = 'block';
            document.getElementById(`${modalId}-detail-text`).innerHTML = detail;
        }

        const modal = new bootstrap.Modal(modalEl);
        let isConfirmed = false;

        document.getElementById(`${modalId}-btn`).addEventListener('click', () => {
            isConfirmed = true;
            modal.hide();
        });

        modalEl.addEventListener('hidden.bs.modal', () => {
            modalEl.remove();
            resolve(isConfirmed);
        });

        modal.show();
    });
}
