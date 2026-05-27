// ─── BUTTON REGISTRY ────────────────────────────────────────────────────────
//
// Cada entrada mapeia um botão às condições que devem ser satisfeitas para
// habilitá-lo. Quando uma condição falha, seu `label` vira razão no tooltip.
//
// Tipos de condição:
//   server   — AppState.serverConnected
//   r        — AppState.rAvailable
//   internet — navigator.onLine
//   form     — todos [data-field] de #formId sem is-invalid e obrigatórios preenchidos
//   file-selected  — selectedFiles[containerId] truthy
//   folder-not-empty — #containerId tem pelo menos 1 tr[id] ou .item-card
//   custom   — função check() → boolean (não bloqueia se o elemento não existe ainda)
//
// activeWhen (opcional) — função que retorna false para desativar a condição
// (condição inativa = passa automaticamente, sem bloquear o botão)
//

const rMessage = 'R não encontrado. Instale pela aba "Instalação"';
const internetMessage = 'Sem conexão com a internet';
const serverMessage = 'Sem conexão com o servidor. Execute novamente "start.cmd" para reiniciá-lo';
const powerbiMessage = 'Informe o caminho da pasta da base de dados';
const processoMessage = 'Nenhum processo encontrado na planilha. Certifique-se que o link informado é o da aba "LISTA FINAL"';
const configuracaoMessage = 'Preencha os campos de configuração';
const linkMessage = 'Informe um link válido da aba "LISTA FINAL" da planilha de inserção de demandas do Google Drive';
const BUTTON_REGISTRY = {

    // ── Atas ──────────────────────────────────────────────────────────────────

    'btn-panel-info-pregao': {
        conditions: [
            { type: 'form', formId: 'panel-info-pregao', label: 'Preencha as informações do pregão' },
            { type: 'folder-not-empty', containerId: 'atas-relatorios-sicaf', label: 'Coloque os relatórios de credenciamento do SICAF na pasta' },
            { type: 'r', label: rMessage },
            { type: 'server', label: serverMessage },
        ],
    },

    'btn-panel-atas-modelos': {
        conditions: [
            { type: 'custom', check: () => typeof atasData !== 'undefined' && atasData.dadosDisponiveis, label: 'Dados não disponíveis. Execute "Obter dados dos SICAF" primeiro' },
            { type: 'file-selected', containerId: 'atas-modelos', label: 'Selecione um modelo de ata' },
            { type: 'server', label: serverMessage },
        ],
    },

    // ── CATMAT ────────────────────────────────────────────────────────────────

    'btn-panel-itens-tr': {
        conditions: [
            { type: 'file-selected', containerId: 'catmat-lista-itens-tr', label: 'Selecione um arquivo em "Itens do TR"' },
            {
                type: 'custom',
                check: () => {
                    const container = document.getElementById('catmat-arquivos-auxiliares');
                    if (!container) return true;
                    return container.textContent.includes('Margens.xlsx');
                },
                label: '"Margens.xlsx" não encontrado na pasta de arquivos auxiliares',
            },
            {
                type: 'custom',
                check: () => {
                    const selected = document.querySelector('input[name="catmat-metodo"]:checked');
                    if (selected?.value !== 'lista') return true; // condição só para modo lista
                    const container = document.getElementById('catmat-arquivos-auxiliares');
                    if (!container) return true;
                    return container.textContent.includes('Lista CATMAT.xlsx');
                },
                label: '"Lista CATMAT.xlsx" não encontrado na pasta de arquivos auxiliares',
            },
            { type: 'r', label: rMessage },
            { type: 'internet', label: internetMessage, activeWhen: () => document.querySelector('input[name="catmat-metodo"]:checked')?.value === 'api' },
            { type: 'server', label: serverMessage },
        ],
    },

    // ── Fornecedores ──────────────────────────────────────────────────────────

    'btn-analisar-certidoes': {
        conditions: [
            { type: 'server', label: serverMessage },
        ],
    },

    'btn-obter-dados-fornecedores': {
        conditions: [
            {
                type: 'custom',
                check: () => typeof pregaoSelecionado === 'undefined' || !!pregaoSelecionado,
                label: 'Selecione um pregão',
            },
            {
                type: 'custom',
                check: () => typeof pregaoFolderFileCount === 'undefined' || pregaoFolderFileCount > 0,
                label: 'A pasta do pregão selecionado está vazia',
                activeWhen: () => typeof pregaoSelecionado !== 'undefined' && !!pregaoSelecionado,
            },
            { type: 'r', label: rMessage },
            { type: 'server', label: serverMessage },
        ],
    },

    // ── Importação ────────────────────────────────────────────────────────────

    'btn-importacao-principal': {
        conditions: [
            { type: 'custom', check: () => typeof importacaoLinkValido === 'undefined' || importacaoLinkValido, label: linkMessage },
            { type: 'form', formId: 'form-importacao-config', label: configuracaoMessage },
            {
                type: 'custom',
                check: () => {
                    const activeTab = document.querySelector('#importacaoTabs .nav-link.active')?.id;
                    if (activeTab === 'tab-relatorio-gerencial') return true;
                    const badge = document.getElementById('badge-processos');
                    return !badge || !badge.classList.contains('d-none');
                },
                label: processoMessage,
                activeWhen: () => typeof importacaoLinkValido !== 'undefined' && importacaoLinkValido
            },
            {
                type: 'custom',
                check: () => {
                    const activeTab = document.querySelector('#importacaoTabs .nav-link.active')?.id;
                    if (activeTab !== 'tab-resumo-pedidos') return true;
                    const container = document.getElementById('importacao-resumos-pdf');
                    return !!container && container.querySelectorAll('tr[data-filepath], .item-card').length > 0;
                },
                label: 'Nenhum arquivo .pdf em "Prints das telas dos pedidos"',
                activeWhen: () => document.querySelector('#importacaoTabs .nav-link.active')?.id === 'tab-resumo-pedidos'
            },
            { type: 'r', label: rMessage },
            { type: 'internet', label: internetMessage },
            { type: 'server', label: serverMessage },
        ],
    },

    // ── Mapas ─────────────────────────────────────────────────────────────────

    'btn-run-mapas': {
        conditions: [
            { type: 'folder-not-empty', containerId: 'mapas-mapas-a-processar', label: 'Coloque os mapas de licitação obtidos no Solar na pasta' },
            { type: 'r', label: rMessage },
            { type: 'server', label: serverMessage },
        ],
    },

    // ── Power BI ──────────────────────────────────────────────────────────────

    'btn-run-powerbi-panel': {
        conditions: [
            { type: 'form', formId: 'panel-powerbi-path', label: powerbiMessage },
            { type: 'r', label: rMessage },
            {
                type: 'internet',
                label: internetMessage,
                // Necessita internet para todos os modos exceto 'licitacao' e 'execucao'
                activeWhen: () => {
                    if (typeof POWERBI_OPCOES_INTERNET === 'undefined') return true;
                    const sel = document.querySelector('#panel-powerbi input[name="powerbi-tipo"]:checked');
                    return !sel || POWERBI_OPCOES_INTERNET.includes(sel.value);
                },
            },
            { type: 'server', label: serverMessage },
        ],
    },

    'btn-run-powerbi-maintenance': {
        conditions: [
            { type: 'form', formId: 'panel-powerbi-path', label: powerbiMessage },
            { type: 'r', label: rMessage },
            { type: 'server', label: serverMessage },
        ],
    },

    'btn-run-powerbi-observatorio': {
        conditions: [
            { type: 'form', formId: 'panel-powerbi-path', label: powerbiMessage },
            { type: 'internet', label: internetMessage },
            { type: 'server', label: serverMessage },
        ],
    },

    // ── SNE ───────────────────────────────────────────────────────────────────

    'btn-atualizar-empenhos': {
        conditions: [
            { type: 'server', label: serverMessage },
        ],
    },

    'btn-criar-afs': {
        conditions: [
            { type: 'server', label: serverMessage },
        ],
    },

    'btn-atualizar-afs': {
        conditions: [
            { type: 'server', label: serverMessage },
        ],
    },

    // ── Instalação ────────────────────────────────────────────────────────────

    'btn-run-instalacao': {
        conditions: [
            { type: 'r', label: rMessage },
            { type: 'internet', label: internetMessage },
            { type: 'server', label: serverMessage },
        ],
    },

    'btn-run-npm_update': {
        conditions: [
            { type: 'internet', label: internetMessage },
            { type: 'server', label: serverMessage },
        ],
    },
};

// ─── CONDITION CHECKER ───────────────────────────────────────────────────────

function checkCondition(cond) {
    // Condição desativada (ex: internet só para modo API) → passa automaticamente
    if (cond.activeWhen && !cond.activeWhen()) return true;

    switch (cond.type) {
        case 'server':
            return AppState.serverConnected;

        case 'r':
            return AppState.rAvailable;

        case 'internet':
            return navigator.onLine;

        case 'form': {
            const fields = document.querySelectorAll(`#${cond.formId} [data-field]`);
            if (fields.length === 0) return true;
            return !Array.from(fields).some(f =>
                f.classList.contains('is-invalid') || (f.hasAttribute('required') && !f.value)
            );
        }

        case 'file-selected':
            return !!selectedFiles[cond.containerId];

        case 'folder-not-empty': {
            const container = document.getElementById(cond.containerId);
            if (!container) return true; // aba não carregada ainda
            return container.querySelectorAll('tr[data-filepath], .item-card').length > 0;
        }

        case 'custom':
            try { return cond.check(); } catch { return true; }

        default:
            return true;
    }
}

// ─── BUTTON EVALUATION ENGINE ────────────────────────────────────────────────

function evaluateButton(buttonId) {
    const config = BUTTON_REGISTRY[buttonId];
    const button = document.getElementById(buttonId);
    if (!button || !config) return;

    // Não sobrescreve botões em execução (ex: Observatório durante SSE)
    if (button.dataset.executing === 'true') return;

    const reasons = config.conditions
        .filter(cond => !checkCondition(cond))
        .map(cond => cond.label);

    button.disabled = reasons.length > 0;
    updateButtonTooltip(button, reasons);
}

function evaluateAllButtons() {
    Object.keys(BUTTON_REGISTRY).forEach(evaluateButton);
    atualizarIndicadoresSubTabs();
}

// ─── FIELD VALIDATION ────────────────────────────────────────────────────────

function validateTabFields(abaName) {
    const fields = Array.from(document.querySelectorAll(`#${abaName} [data-field]`));
    fields.forEach(field => validateSingleField(field));
}

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
    evaluateAllButtons();

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

// ─── TOOLTIP HELPERS ─────────────────────────────────────────────────────────

// Envolve um botão num wrapper <span> para permitir tooltip em botões disabled
function wrapButtonForTooltip(button) {
    if (button.parentElement?.classList.contains('btn-tooltip-wrapper')) return button.parentElement;
    const wrapper = document.createElement('span');
    wrapper.className = 'btn-tooltip-wrapper';
    button.parentElement.insertBefore(wrapper, button);
    wrapper.appendChild(button);
    return wrapper;
}

function updateButtonTooltip(button, reasons) {
    const wrapper = button.parentElement?.classList.contains('btn-tooltip-wrapper')
        ? button.parentElement
        : wrapButtonForTooltip(button);

    // Clean up any legacy tooltip on the wrapper itself
    const wrapperTooltip = bootstrap.Tooltip.getInstance(wrapper);
    if (wrapperTooltip) wrapperTooltip.dispose();
    wrapper.removeAttribute('data-bs-toggle');
    wrapper.removeAttribute('data-bs-title');

    let warningIcon = wrapper.querySelector('.btn-unavailable-icon');

    if (reasons.length === 0) {
        if (warningIcon) {
            const iconTooltip = bootstrap.Tooltip.getInstance(warningIcon);
            if (iconTooltip) iconTooltip.dispose();
            warningIcon.remove();
        }
        return;
    }

    const header = '<div class="tooltip-error-title">Botão indisponível</div>';
    const body = '<ul class="tooltip-error-list">' + reasons.map(r => `<li>${r}</li>`).join('') + '</ul>';
    const title = header + body;

    if (!warningIcon) {
        warningIcon = document.createElement('i');
        warningIcon.className = 'material-symbols-outlined btn-unavailable-icon';
        warningIcon.textContent = 'cancel';
        button.insertAdjacentElement('afterend', warningIcon);
    }

    warningIcon.setAttribute('data-bs-html', 'true');
    warningIcon.setAttribute('data-bs-title', title);

    const existingIconTooltip = bootstrap.Tooltip.getInstance(warningIcon);
    if (existingIconTooltip) {
        existingIconTooltip.setContent({ '.tooltip-inner': title });
    } else {
        new bootstrap.Tooltip(warningIcon, {
            ...TOOLTIP_DEFAULTS,
            customClass: 'custom-tooltip tooltip-disabled-btn'
        });
    }
}

// ─── LIVE VALIDATION SETUP ───────────────────────────────────────────────────

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
    evaluateAllButtons();
}

// ─── GLOBAL EVENT LISTENERS ──────────────────────────────────────────────────

// Seleção de arquivo em file-list selecionável
document.addEventListener('file-selected', () => evaluateAllButtons());

// Carregamento/atualização de pasta
document.addEventListener('files-loaded', () => evaluateAllButtons());

// Mudança de modo (catmat, powerbi) afeta condições ativas
document.addEventListener('change', (e) => {
    if (e.target.name === 'catmat-metodo' || e.target.name === 'powerbi-tipo') {
        evaluateAllButtons();
    }
});

// ─── SUB-TABS INDICATOR ──────────────────────────────────────────────────────

/**
 * Verifica cada tab-pane de sub-tabs e marca o nav-link correspondente
 * quando há campos requeridos vazios ou inválidos no painel.
 * @param {HTMLElement} [campo] - Campo que disparou a atualização (otimiza para atualizar apenas o grupo relevante)
 */
function atualizarIndicadoresSubTabs(campo) {
    const tabContents = campo
        ? [campo.closest('.tab-content')].filter(Boolean)
        : document.querySelectorAll('.sub-tabs + .sub-tabs-content .tab-content, .sub-tabs + .panel .tab-content');

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
