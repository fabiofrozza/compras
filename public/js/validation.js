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
    setButtonState(field);

    return !hasError;
}

function setButtonState(field) {
    const formElement = field.closest('.script-form');
    if (!formElement) return;

    const formId = formElement.id;

    if (formId === 'form-powerbi-path') {
        const hasInvalidFields = Array.from(document.querySelectorAll(`#${formId} [data-field]`)).some(f => f.classList.contains('is-invalid'));
        ['btn-run-powerbi-panel', 'btn-run-powerbi-maintenance'].forEach(btnId => {
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

    button.disabled = shouldDisable;
}

function setupLiveValidation(aba) {
    const container = document.querySelector('#' + aba);
    const fields = container.querySelectorAll('[data-field]');
    fields.forEach(field => {
        ['input', 'change', 'blur', 'click'].forEach(eventType => {
            field.addEventListener(eventType, () => {
                validateSingleField(field);
            });
        });
    });
}