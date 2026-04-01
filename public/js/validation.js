function validateTabFields(abaName) {
    const fields = Array.from(document.querySelectorAll(`#${abaName} [data-field]`));
    fields.forEach(field => validateSingleField(field));
}

/**
 * Valida um campo individual e aplica a classe is-invalid se necessário
 * @param {HTMLElement} field - O campo a validar
 * @returns {boolean} true se válido, false caso contrário
 */
function validateSingleField(field) {
    const { value, dataset: { validateRule, regexPattern }, attributes } = field;
    const isRequired = field.hasAttribute('required');
    const minValue = attributes['min']?.value;
    const maxValue = attributes['max']?.value;

    let hasError = false;

    if (isRequired && value === '') {
        hasError = true;
    }
    else if (value !== '' && validateRule) {
        switch (validateRule) {
            case 'numeric':
                if (!/^\d+$/.test(value)) {
                    hasError = true;
                } else {
                    const numValue = Number(value);
                    if ((minValue && numValue < Number(minValue)) ||
                        (maxValue && numValue > Number(maxValue))) {
                        hasError = true;
                    }
                }
                break;

            case 'email':
                hasError = !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(value);
                break;

            case 'pattern':
                hasError = regexPattern && !new RegExp(regexPattern).test(value);
                break;

            case 'date':
                hasError = isNaN(Date.parse(value));
                break;

            case 'file-selectable':
                hasError = field.querySelectorAll('.selected').length === 0;
                break;
        }
    }

    field.classList.toggle('is-invalid', hasError);
    toggleValidationMsg(field, hasError);
    setButtonState(field);

    return !hasError;
}

function toggleValidationMsg(field, hasError) {
    const inputGroup = field.closest('.input-group');

    if (inputGroup) {
        const msgField = inputGroup.querySelector('[data-validation-msg]');
        if (!msgField) return;

        const anyError = hasError || Array.from(inputGroup.querySelectorAll('[data-field]'))
            .filter(f => f !== field)
            .some(f => f.classList.contains('is-invalid'));

        const feedbackId = msgField.id + '-feedback';
        let feedback = document.getElementById(feedbackId);
        if (!feedback) {
            feedback = document.createElement('div');
            feedback.id = feedbackId;
            feedback.className = 'validation-msg';
            feedback.textContent = msgField.dataset.validationMsg;
            inputGroup.insertAdjacentElement('afterend', feedback);
        }
        feedback.classList.toggle('show', anyError);
        return;
    }

    const msg = field.dataset.validationMsg;
    if (!msg) return;

    const feedbackId = field.id + '-feedback';
    let feedback = document.getElementById(feedbackId);
    if (!feedback) {
        feedback = document.createElement('div');
        feedback.id = feedbackId;
        feedback.className = 'validation-msg';
        feedback.textContent = msg;
        field.insertAdjacentElement('afterend', feedback);
    }
    feedback.classList.toggle('show', hasError);
}

function setButtonState(field) {
    const formElement = field.closest('.script-form');
    if (!formElement) return;

    const formId = formElement.id;

    if (formId === 'form-powerbi-path') {
        const hasInvalidFields = Array.from(document.querySelectorAll(`#${formId} [data-field]`)).some(f => f.classList.contains('is-invalid'));
        const btnMaintenance = document.getElementById('btn-run-powerbi-maintenance');
        if (btnMaintenance) btnMaintenance.disabled = hasInvalidFields;
        if (typeof atualizarBotaoPowerBIPanel === 'function') atualizarBotaoPowerBIPanel();
        return;
    }

    if (formId === 'form-importacao-link' || formId === 'importacaoTabsForm') {
        const hasInvalidFields = Array.from(document.querySelectorAll('#form-importacao-link [data-field], #form-importacao-config [data-field]'))
            .some(f => f.classList.contains('is-invalid'));
        ['btn-importacao-arquivos', 'btn-importacao-resumo', 'btn-importacao-relatorio'].forEach(btnId => {
            const btn = document.getElementById(btnId);
            if (btn) btn.disabled = hasInvalidFields;
        });
        return;
    }

    const button = document.getElementById('btn-' + formId);

    if (!button) return;

    const hasInvalidFields = Array.from(document.querySelectorAll(`#${formId} [data-field]`)).some(f => f.classList.contains('is-invalid'));

    let shouldDisable = hasInvalidFields;

    // Regras de negócio: botão de atas também exige dados do SICAF disponíveis
    if (button.id === 'btn-form-atas-modelos') {
        shouldDisable = hasInvalidFields || !atasData.dadosDisponiveis;
    }

    // Botão catmat: opção API exige internet
    if (button.id === 'btn-form-itens-tr') {
        const selected = document.querySelector('input[name="catmat-metodo"]:checked');
        shouldDisable = hasInvalidFields || (selected?.value === 'api' && !navigator.onLine);
    }

    button.disabled = shouldDisable;
}

function atualizarBotaoCatmat() {
    const btn = document.getElementById('btn-form-itens-tr');
    if (!btn) return;
    const hasInvalidFields = Array.from(document.querySelectorAll('#form-itens-tr [data-field]'))
        .some(f => f.classList.contains('is-invalid'));
    const selected = document.querySelector('input[name="catmat-metodo"]:checked');
    btn.disabled = hasInvalidFields || (selected?.value === 'api' && !navigator.onLine);
}

document.addEventListener('change', (e) => {
    if (e.target.name === 'catmat-metodo') atualizarBotaoCatmat();
});

function setupLiveValidation(aba) {
    const container = document.querySelector('#' + aba);
    const fields = container.querySelectorAll('[data-field]');
    fields.forEach(field => {
        ['input', 'change', 'blur', 'click'].forEach(eventType => {
            field.addEventListener(eventType, () => {
                validateSingleField(field);
                atualizarIndicadoresSubTabs(field);
            });
        });
    });
}

/**
 * Verifica cada tab-pane de sub-tabs e marca o nav-link correspondente
 * quando há campos requeridos vazios ou inválidos no painel.
 * @param {HTMLElement} [campo] - Campo que disparou a atualização (otimiza para atualizar apenas o grupo relevante)
 */
function atualizarIndicadoresSubTabs(campo) {
    const tabContents = campo
        ? [campo.closest('.tab-content')].filter(Boolean)
        : document.querySelectorAll('.sub-tabs + .sub-tabs-content .tab-content, .sub-tabs + .script-form .tab-content');

    tabContents.forEach(tabContent => {
        tabContent.querySelectorAll('.tab-pane').forEach(pane => {
            const navLink = document.querySelector(`[data-bs-target="#${pane.id}"]`);
            if (!navLink) return;

            const temProblema = Array.from(pane.querySelectorAll('[required]')).some(f =>
                f.classList.contains('is-invalid') || ('value' in f && !f.value)
            );

            navLink.classList.toggle('has-validation-issue', temProblema);
        });
    });
}