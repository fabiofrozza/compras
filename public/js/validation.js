/**
 * Sistema centralizado de validação de formulários
 * Usa atributos data-* para definir regras de validação
 */

function validateTabFields(abaName) {
    const selector = `#${abaName} [data-field]`;

    let fields = [];

    fields = Array.from(document.querySelectorAll(selector));

    fields.forEach(field => {
        validateSingleField(field)
    });
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
    const button = document.getElementById('btn-' + formId);

    if (!button) return;

    const hasInvalidFields = Array.from(document.querySelectorAll(`#${formId} [data-field]`)).some(f => f.classList.contains('is-invalid'));
    
    // Por padrão, habilitar se não há campos inválidos
    let shouldDisable = hasInvalidFields;
    
    // Verificações adicionais de regras de negócio por tipo de botão
    if (button.id === 'btn-form-atas-modelos') {
        // Para o botão de gerar atas, também verificar se os dados do SICAF estão disponíveis
        shouldDisable = hasInvalidFields || !atasData.dadosDisponiveis;
    }
    
    button.disabled = shouldDisable;
}

/**
 * Configura listeners de validação em tempo real para um container
 * Valida o campo ao digitar e mostra erro imediatamente se inválido
 * @param {string|HTMLElement} aba - Seletor da aba para configurar validação
 */
function setupLiveValidation(aba) {
    let container = null;

    container = document.querySelector('#' + aba);

    const fields = container.querySelectorAll('[data-field]');
    fields.forEach(field => {
        // Valida em tempo real ao digitar/mudar valor
        ['input', 'change', 'blur', 'click'].forEach(eventType => {
            field.addEventListener(eventType, () => {
                validateSingleField(field);
            });
        });
    });
}