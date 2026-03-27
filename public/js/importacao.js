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

        const processos = result.processosSPA || [];

        switch (result.status) {
            case 'success':
                atualizarFeedbackPlanilha('success', result.msg, 'check_circle');
                atualizarWidgetsExtras(processos, result.temValidacaoManual || false);
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

        popularComboProcessosSPA(processos);

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
 * @param {string} iconeClass - Classe do ícone Material Symbols (ex: 'check_circle')
 */
function atualizarFeedbackPlanilha(tipo, mensagem, iconeClass) {
    const linkInput = document.getElementById('importacao-link');
    const statusText = document.getElementById('texto-status-link');
    const widgetStatus = document.getElementById('widget-status-link');

    if (!linkInput || !statusText || !widgetStatus) return;

    const btnAbrir = document.getElementById('btn-abrir-link');
    if (btnAbrir) btnAbrir.disabled = !linkInput.value.trim();

    const classesIcone = ['border-success', 'border-danger', 'border-warning', 'border-muted', 'border-info'];
    widgetStatus.classList.remove(...classesIcone);

    statusText.innerHTML = mensagem;

    // Atualiza o tooltip com o mesmo texto (sem HTML)
    statusText.setAttribute('data-bs-title', statusText.textContent.trim());
    const tooltipInstance = bootstrap.Tooltip.getInstance(statusText);
    if (tooltipInstance) tooltipInstance.setContent({ '.tooltip-inner': statusText.textContent.trim() });

    switch (tipo) {
        case 'success':
            widgetStatus.classList.add('border-success');
            linkInput.classList.remove('is-invalid');
            linkInput.classList.add('is-valid');
            importacaoLinkValido = true;
            break;

        case 'error':
            widgetStatus.classList.add('border-danger');
            linkInput.classList.remove('is-valid');
            linkInput.classList.add('is-invalid');
            importacaoLinkValido = false;
            break;

        case 'warning':
            widgetStatus.classList.add('border-warning');
            linkInput.classList.remove('is-invalid');
            linkInput.classList.add('is-valid');
            importacaoLinkValido = true;
            break;

        case 'loading':
            widgetStatus.classList.add('border-info');
            linkInput.classList.remove('is-valid', 'is-invalid');
            importacaoLinkValido = false;
            break;

        case 'info':
        default:
            widgetStatus.classList.add('border-muted');
            linkInput.classList.remove('is-valid', 'is-invalid');
            importacaoLinkValido = false;
            break;
    }

    // Esconde widgets extras quando não é sucesso
    if (tipo !== 'success') {
        atualizarWidgetsExtras([], false);
    }

    verificarLiberacaoBotoesImportacao();
}

/**
 * Atualiza os widgets de processos e validação manual.
 */
function atualizarWidgetsExtras(processos, temValidacaoManual) {
    const widgetProcessos = document.getElementById('widget-processos');
    const textoProcessos = document.getElementById('texto-processos');
    const widgetValidacao = document.getElementById('widget-validacao');
    const textoValidacao = document.getElementById('texto-validacao');

    if (widgetProcessos && textoProcessos) {
        if (processos.length > 0) {
            widgetProcessos.classList.remove('d-none');
            textoProcessos.textContent = processos.length;
        } else {
            widgetProcessos.classList.add('d-none');
        }
    }

    if (widgetValidacao && textoValidacao) {
        if (temValidacaoManual) {
            widgetValidacao.classList.remove('d-none');
            textoValidacao.innerHTML = '<span class="badge rounded-pill warning"><i class="material-symbols-outlined">warning</i> Há itens a validar</span>';
        } else {
            widgetValidacao.classList.add('d-none');
        }
    }
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

    atualizarIndicadoresSubTabs();

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

function exibirResultadoImportacao(data) {
    const col = document.getElementById('importacao-resultado-col');
    const linkCol = document.getElementById('importacao-link-col');
    const grid = document.getElementById('importacao-resultado-widgets');
    const msgDiv = document.getElementById('importacao-resultado-msg');
    if (!col || !grid || !msgDiv) return;

    grid.innerHTML = '';
    msgDiv.innerHTML = '';

    const conf = data.conferencia;
    if (conf && Object.keys(conf).length > 0) {
        grid.classList.remove('d-none');

        const widgets = [
            { icon: 'category', label: 'Grupo', value: conf.grupo + ' - ' + conf.nome_grupo, size: 'full' },
            { icon: 'calendar_today', label: 'Etapa', value: conf.etapa + ' / ' + conf.ano },
            { icon: 'inventory_2', label: 'Itens', value: conf.n_itens },
            { icon: 'assignment', label: 'Solicitações', value: conf.n_solicitacoes }
        ];

        widgets.forEach(w => {
            if (w.value == null) return;
            const div = document.createElement('div');
            div.className = 'dashboard-widget ' + (w.size === 'full' ? 'widget-full' : '');
            div.innerHTML = `
                <div class="widget-icon"><i class="material-symbols-outlined">${w.icon}</i></div>
                <div class="widget-content">
                    <span class="widget-label">${w.label}</span>
                    <span class="widget-value">${w.value}</span>
                </div>`;
            grid.appendChild(div);
        });
    }

    // msg_erro: array onde o primeiro item é o nome do log (ignorar), os demais são mensagens
    const erros = data.msg_erro;
    if (Array.isArray(erros) && erros.length > 1) {
        const mensagens = erros.slice(1);
        let msgErros = `
            <div class="alert alert-warning alert-sm py-1 px-2 mb-1">
                <i class="material-symbols-outlined">warning</i>
        `;
        msgErros += mensagens
            .map(m => ` ${m}<br>`)
            .join('');
        msgErros += `</div>`;

        msgDiv.classList.remove('d-none');
        msgDiv.innerHTML = msgErros;
    }

    col.classList.remove('d-none');
    linkCol.classList.remove('col');
    linkCol.classList.add('col-sm-6')
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