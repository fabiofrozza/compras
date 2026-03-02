// TOASTS, TOOLTIPS AND MODALS

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
    closeButton.innerHTML = '<i class="fas fa-x fa-2xs"></i>';

    const removeToast = () => {
        toast.classList.add('removing');
        setTimeout(() => {
            toast.remove();
        }, 300); // Aguardar a animação de saída
    };

    closeButton.addEventListener('click', removeToast);

    toast.appendChild(messageContainer);
    toast.appendChild(closeButton);
    container.appendChild(toast);

    setTimeout(removeToast, duration);
}

function createRequiredFieldsTooltip() {
    const requiredFields = document.querySelectorAll('label:has(+ [data-validate-rule]):not(:has(.indicator-required)), label:has(+ :required):not(:has(.indicator-required)), h4:has(+ [data-validate-rule]):not(:has(.indicator-required))');
    requiredFields.forEach(field => {
        const asteriskSpan = document.createElement('span');
        asteriskSpan.className = 'indicator-required';
        asteriskSpan.dataset.bsTitle = 'Campo obrigatório';
        asteriskSpan.dataset.bsToggle = 'tooltip';
        asteriskSpan.innerHTML = '<i class="fas fa-asterisk"></i>'

        field.appendChild(asteriskSpan);
    });

    // Re-inicializar tooltips para incluir os novos campos obrigatórios
    initializeTooltips();
}

function initializeTooltips() {
    // Initialize tooltips (Bootstrap 5)
    const tooltipTriggerList = document.querySelectorAll('[data-bs-toggle="tooltip"]');
    const tooltipList = [...tooltipTriggerList].map(tooltipTriggerEl => new bootstrap.Tooltip(tooltipTriggerEl, {
        container: 'body',
        trigger: 'hover',
        customClass: 'custom-tooltip'
    }));
}

function removeTooltip(element) {
    // Ocultar tooltip antes de remover o elemento do DOM
    const tooltip = bootstrap.Tooltip.getInstance(element);
    if (tooltip) tooltip.hide();
}

initializeTooltips();

function testToasts() {
    showToast('Testando toast de sucesso', 'success', 500000)
    showToast('Testando toast de erro', 'error', 500000)
    showToast('Testando toast de aviso', 'warning', 500000)
    showToast('Testando toast de informação', 'info', 500000)
}

//testToasts();

// ==========================================================================
// CENTRAL DE NOTIFICAÇÕES (Painel no Header)
// ==========================================================================

/** @type {{ id: number, message: string, type: string, source: string, timestamp: Date }[]} */
const notifications = [];
let notificationIdCounter = 0;

const notificationIcons = {
    success: 'fas fa-circle-check',
    error: 'fas fa-circle-xmark',
    warning: 'fas fa-circle-exclamation',
    info: 'fas fa-circle-info'
};

/**
 * Adiciona uma notificação à central de notificações no header.
 * @param {Object} options
 * @param {string} options.message - Mensagem da notificação
 * @param {'success'|'error'|'warning'|'info'} [options.type='info'] - Tipo da notificação
 * @param {string} [options.source='Sistema'] - Aba ou serviço de origem
 */
function addNotification({ message, type = 'info', source = 'Sistema' }) {
    const notification = {
        id: ++notificationIdCounter,
        message,
        type,
        source,
        timestamp: new Date()
    };

    notifications.unshift(notification); // Mais recente no topo
    updateNotificationBadge();
    shakeBell();
    renderNotificationList();
}

/** Dispara a animação de balanço no ícone do sino */
function shakeBell() {
    const btn = document.getElementById('notificationsBtn');
    if (!btn) return;
    btn.classList.remove('bell-shake');
    // forçar reflow para reiniciar a animação caso chegue outra notificação imediatamente
    void btn.offsetWidth;
    btn.classList.add('bell-shake');
    setTimeout(() => btn.classList.remove('bell-shake'), 700);
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
                <i class="fas fa-bell-slash"></i>
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
            <div class="notification-icon"><i class="${iconClass}"></i></div>
            <div class="notification-body">
                <div class="notification-message">${notif.message}</div>
                <div class="notification-meta">
                    <span class="notification-time"><i class="fas fa-clock me-1"></i>${timeStr}</span>
                    <span class="notification-source">${notif.source}</span>
                </div>
            </div>
            <button class="notification-dismiss" aria-label="Dispensar" data-notif-id="${notif.id}">
                <i class="fas fa-xmark"></i>
            </button>`;

        list.appendChild(item);
    });
}

/** Remove uma notificação individual pelo ID */
function dismissNotification(id) {
    const item = document.querySelector(`.notification-item[data-notif-id="${id}"]`);

    if (item) {
        item.classList.add('removing');
        setTimeout(() => {
            const index = notifications.findIndex(n => n.id === id);
            if (index !== -1) notifications.splice(index, 1);
            updateNotificationBadge();
            renderNotificationList();
        }, 200);
    } else {
        const index = notifications.findIndex(n => n.id === id);
        if (index !== -1) notifications.splice(index, 1);
        updateNotificationBadge();
        renderNotificationList();
    }
}

/** Remove todas as notificações */
function clearAllNotifications() {
    notifications.length = 0;
    updateNotificationBadge();
    renderNotificationList();
}

// Event delegation para dismiss de notificações individuais
document.addEventListener('click', (e) => {
    const dismissBtn = e.target.closest('.notification-dismiss');
    if (dismissBtn) {
        const id = parseInt(dismissBtn.dataset.notifId, 10);
        if (!isNaN(id)) dismissNotification(id);
    }
});

// Botão 'Limpar tudo'
document.addEventListener('DOMContentLoaded', () => {
    const clearAllBtn = document.getElementById('notification-clear-all');
    if (clearAllBtn) {
        clearAllBtn.addEventListener('click', clearAllNotifications);
    }
});

// Inicializar a lista vazia ao carregar
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
    icon = 'fas fa-exclamation-triangle'
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
                                <i class="${icon} me-2"></i>${title}
                            </h5>
                            <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
                        </div>
                        <div class="modal-body">
                            <p>${message}</p>
                            <div id="${modalId}-detail-container" class="p-2 bg-light border rounded mb-3 text-break small" style="display: none;">
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
