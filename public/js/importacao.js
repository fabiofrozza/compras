let importacaoLinkValido = false;
let importacaoInicializado = false;
let importacaoLastValidatedLink = '';

function inicializarImportacao() {
    if (importacaoInicializado) return;

    if (document.getElementById('importacao-link')) {
        configurarValidacaoContinuaImportacao();
        validarLinkGoogle();
        importacaoInicializado = true;
    }
}

function configurarValidacaoContinuaImportacao() {
    const linkInput = document.getElementById('importacao-link');
    const configForm = document.getElementById('form-importacao-config');

    if (linkInput) {
        linkInput.addEventListener('change', validarLinkGoogle);
        linkInput.addEventListener('input', () => {
            importacaoLastValidatedLink = '';
            linkInput.classList.remove('is-valid', 'is-invalid');
            atualizarFeedbackPlanilha('info', 'Informe o link da aba LISTA FINAL e aguarde.', 'info');
            verificarLiberacaoBotoesImportacao();
        });
    }

    if (configForm) {
        configForm.addEventListener('input', verificarLiberacaoBotoesImportacao);
        configForm.addEventListener('change', verificarLiberacaoBotoesImportacao);
    }

    // Quando a lista de arquivos é atualizada, recalcular os botões
    document.addEventListener('files-loaded', (e) => {
        if (e.detail && e.detail.containerId === 'importacao-resumos-pdf') {
            verificarLiberacaoBotoesImportacao();
        }
    });
}

function revalidarLinkGoogle() {
    importacaoLastValidatedLink = '';
    validarLinkGoogle();
}

function abrirLinkPlanilha() {
    const url = document.getElementById('importacao-link')?.value.trim();
    if (url) window.open(url, '_blank');
}

async function validarLinkGoogle() {
    const linkInput = document.getElementById('importacao-link');
    if (!linkInput) return;

    const url = linkInput.value.trim();

    if (importacaoLastValidatedLink === url && url !== '') return;

    if (!url) {
        importacaoLastValidatedLink = '';
        atualizarFeedbackPlanilha('info', 'Informe o link da aba LISTA FINAL e aguarde.', 'info');
        popularComboProcessosSPA([]);
        return;
    }

    try {
        new URL(url);
    } catch {
        atualizarFeedbackPlanilha('error', 'Link inválido.', 'cancel');
        popularComboProcessosSPA([]);
        return;
    }

    atualizarFeedbackPlanilha('loading', 'Aguarde... Acessando o link informado...', 'hourglass_empty');

    try {
        const response = await fetch('/api/validate-link', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ url })
        });

        const result = await response.json();
        importacaoLastValidatedLink = url;

        switch (result.status) {
            case 'success':
                atualizarFeedbackPlanilha('success', result.msg, 'check_circle');
                break;
            case 'warning':
                atualizarFeedbackPlanilha('warning', result.msg, 'warning');
                break;
            case 'error':
                atualizarFeedbackPlanilha('error', result.msg, 'cancel');
                if (result.error) console.error('Erro ao acessar link:', result.error);
                break;
            default:
                atualizarFeedbackPlanilha('info', result.msg, 'info');
                break;
        }

        popularComboProcessosSPA(result.processosSPA || []);

    } catch (error) {
        console.error('Erro ao validar link:', error);
        atualizarFeedbackPlanilha('error', 'Erro ao acessar o link informado. Veja erro no console.', 'cancel');
        popularComboProcessosSPA([]);
    }
}

/**
 * Atualiza o feedback visual do link da planilha.
 * 
 * @param {'success'|'error'|'warning'|'info'|'loading'} tipo - Tipo do feedback
 * @param {string} mensagem - Texto da mensagem (sem ícone)
 * @param {string} iconeClass - Classe Font Awesome do ícone (ex: 'check_circle')
 */
function atualizarFeedbackPlanilha(tipo, mensagem, iconeClass) {
    const linkInput = document.getElementById('importacao-link');
    const statusContainer = document.getElementById('importacao-sheet-status');
    const statusIcon = document.getElementById('icon-status-link');
    const statusText = document.getElementById('texto-status-link');

    if (!linkInput || !statusContainer || !statusIcon || !statusText) return;

    const btnAbrir = document.getElementById('btn-abrir-link');
    if (btnAbrir) btnAbrir.disabled = !linkInput.value.trim();

    // Remove todas as classes de estado anteriores
    const classesEstado = ['alert-success', 'alert-danger', 'alert-warning', 'alert-light',
        'border-success', 'border-danger', 'border-warning', 'border'];
    statusContainer.classList.remove(...classesEstado);

    const classesIcone = ['text-success', 'text-danger', 'text-warning', 'text-muted', 'text-info'];
    statusIcon.classList.remove(...classesIcone);

    // Atualiza ícone
    statusIcon.innerHTML = `<i class="material-symbols-outlined">${iconeClass}</i>`;
    statusText.innerHTML = mensagem;

    switch (tipo) {
        case 'success':
            statusContainer.classList.add('alert-success', 'border-success');
            statusIcon.classList.add('text-success');
            linkInput.classList.remove('is-invalid');
            linkInput.classList.add('is-valid');
            importacaoLinkValido = true;
            break;

        case 'error':
            statusContainer.classList.add('alert-danger', 'border-danger');
            statusIcon.classList.add('text-danger');
            linkInput.classList.remove('is-valid');
            linkInput.classList.add('is-invalid');
            importacaoLinkValido = false;
            break;

        case 'warning':
            statusContainer.classList.add('alert-warning', 'border-warning');
            statusIcon.classList.add('text-warning');
            linkInput.classList.remove('is-invalid');
            linkInput.classList.add('is-valid');
            importacaoLinkValido = true;
            break;

        case 'loading':
            statusContainer.classList.add('alert-light', 'border');
            statusIcon.classList.add('text-info');
            linkInput.classList.remove('is-valid', 'is-invalid');
            importacaoLinkValido = false;
            break;

        case 'info':
        default:
            statusContainer.classList.add('alert-light', 'border');
            statusIcon.classList.add('text-muted');
            linkInput.classList.remove('is-valid', 'is-invalid');
            importacaoLinkValido = false;
            break;
    }

    verificarLiberacaoBotoesImportacao();
}

function popularComboProcessosSPA(processos) {
    const comboSPA = document.getElementById('importacao-combo-spa');
    if (!comboSPA) return;

    comboSPA.innerHTML = '<option value="todos">Relatório consolidado (todos os processos)</option>';
    processos.forEach(proc => {
        const option = document.createElement('option');
        option.value = proc;
        option.textContent = proc;
        comboSPA.appendChild(option);
    });
}

function verificarLiberacaoBotoesImportacao() {
    const formFields = document.querySelectorAll('#form-importacao-link [data-field], #form-importacao-config [data-field]');
    let formValido = true;

    formFields.forEach(field => {
        if (field.classList.contains('is-invalid') || (field.hasAttribute('required') && !field.value)) {
            formValido = false;
        }
    });

    const podeExecutar = importacaoLinkValido && formValido;

    const btnArquivos = document.getElementById('btn-importacao-arquivos');
    const btnResumo = document.getElementById('btn-importacao-resumo');
    const btnRelatorio = document.getElementById('btn-importacao-relatorio');

    if (btnArquivos) btnArquivos.disabled = !podeExecutar;
    if (btnRelatorio) btnRelatorio.disabled = !podeExecutar;

    // O botão de resumo exige condição extra: a lista de PDFs não pode estar vazia
    if (btnResumo) {
        const listaPdf = document.getElementById('importacao-resumos-pdf');
        const temArquivosPdf = listaPdf && listaPdf.querySelectorAll('table tr').length > 0;
        btnResumo.disabled = !podeExecutar || !temArquivosPdf;
    }
}

function executarAcaoImportacao(abaDestino) {
    // Mapeia aliases de abas: HTML usa 'gerar', form usa 'arquivos'
    const btnIdMap = {
        'gerar': 'btn-importacao-arquivos',
        'resumo': 'btn-importacao-resumo',
        'relatorio': 'btn-importacao-relatorio'
    };

    const botaoID = btnIdMap[abaDestino] || `btn-importacao-${abaDestino === 'importacao' ? 'arquivos' : abaDestino}`;
    const btn = document.getElementById(botaoID);

    if (btn && btn.disabled) return;

    // Parâmetros enviados em ordem específica (o script R os lê posicionalmente)
    const paramsOrdenados = {};
    paramsOrdenados['argumento_aba'] = abaDestino;

    document.querySelectorAll('#form-importacao-link [data-field], #form-importacao-config [data-field]').forEach(campo => {
        paramsOrdenados[campo.dataset.field] = campo.value;
    });

    if (abaDestino === 'relatorio') {
        const comboSPA = document.getElementById('importacao-combo-spa');
        if (comboSPA) paramsOrdenados['processo_selecionado'] = comboSPA.value;
    }

    runRScript('importacao', paramsOrdenados);
}