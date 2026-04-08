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
            if (linkInput.value.trim() === '') {
                atualizarFeedbackPlanilha('info', 'Insira o link da planilha para validação.');
            } else {
                linkInput.classList.remove('is-valid', 'is-invalid');
                atualizarFeedbackPlanilha('info', 'Informe o link da aba LISTA FINAL e aguarde.');
            }
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
        atualizarFeedbackPlanilha('info', 'Informe o link da aba LISTA FINAL e aguarde.');
        popularComboProcessosSPA([]);
        return;
    }

    try {
        new URL(url);
    } catch {
        atualizarFeedbackPlanilha('error', 'Link inválido.');
        popularComboProcessosSPA([]);
        return;
    }

    atualizarFeedbackPlanilha('loading', 'Aguarde... Acessando o link informado...');

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
                atualizarFeedbackPlanilha('success', result.msg);
                atualizarWidgetsExtras(processos, result.temValidacaoManual || false);
                break;
            case 'warning':
                atualizarFeedbackPlanilha('warning', result.msg);
                break;
            case 'error':
                atualizarFeedbackPlanilha('error', result.msg);
                if (result.error) console.error('Erro ao acessar link:', result.error);
                break;
            default:
                atualizarFeedbackPlanilha('info', result.msg);
                break;
        }

        popularComboProcessosSPA(processos);

    } catch (error) {
        console.error('Erro ao validar link:', error);
        atualizarFeedbackPlanilha('error', 'Erro ao acessar o link informado. Veja erro no console.');
        popularComboProcessosSPA([]);
    }
}

function atualizarFeedbackPlanilha(tipo, mensagem) {
    const linkInput = document.getElementById('importacao-link');
    const badge = document.getElementById('badge-status-link');

    if (!linkInput || !badge) return;

    const btnAbrir = document.getElementById('btn-abrir-link');
    if (btnAbrir) btnAbrir.disabled = !linkInput.value.trim();

    badge.classList.remove('text-bg-success', 'text-bg-danger', 'text-bg-warning', 'text-bg-secondary', 'text-bg-info');
    badge.textContent = mensagem;

    switch (tipo) {
        case 'success':
            badge.classList.add('text-bg-success');
            linkInput.classList.remove('is-invalid');
            linkInput.classList.add('is-valid');
            importacaoLinkValido = true;
            break;

        case 'error':
            badge.classList.add('text-bg-danger');
            linkInput.classList.remove('is-valid');
            linkInput.classList.add('is-invalid');
            importacaoLinkValido = false;
            break;

        case 'warning':
            badge.classList.add('text-bg-warning');
            linkInput.classList.remove('is-invalid');
            linkInput.classList.add('is-valid');
            importacaoLinkValido = true;
            break;

        case 'loading':
            badge.classList.add('text-bg-info');
            linkInput.classList.remove('is-valid', 'is-invalid');
            importacaoLinkValido = false;
            break;

        case 'info':
        default:
            badge.classList.add('text-bg-secondary');
            if (linkInput.value.trim() !== '') {
                linkInput.classList.remove('is-valid', 'is-invalid');
            }
            importacaoLinkValido = false;
            break;
    }

    if (tipo !== 'success') {
        atualizarWidgetsExtras([], false);
        ocultarResultadoImportacao();
    }

    verificarLiberacaoBotoesImportacao();
}

function atualizarWidgetsExtras(processos, temValidacaoManual) {
    const badgeProcessos = document.getElementById('badge-processos');
    const badgeValidacao = document.getElementById('badge-validacao');

    if (badgeProcessos) {
        if (processos.length > 0) {
            badgeProcessos.classList.remove('d-none');
            badgeProcessos.textContent = `${processos.length} processo${processos.length !== 1 ? 's' : ''}`;
        } else {
            badgeProcessos.classList.add('d-none');
        }
    }

    if (badgeValidacao) {
        badgeValidacao.classList.toggle('d-none', !temValidacaoManual);
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

    const reasons = [];
    if (!importacaoLinkValido) reasons.push('Informe um link válido do Comprasnet');
    if (!formValido) reasons.push('Preencha os campos obrigatórios');
    if (!navigator.onLine) reasons.push('Sem conexão com a internet');

    const btnArquivos = document.getElementById('btn-importacao-arquivos');
    const btnResumo = document.getElementById('btn-importacao-resumo');
    const btnRelatorio = document.getElementById('btn-importacao-relatorio');

    [btnArquivos, btnRelatorio].forEach(btn => {
        if (!btn) return;
        btn.disabled = reasons.length > 0;
        updateButtonTooltip(btn, reasons);
    });

    if (btnResumo) {
        const listaPdf = document.getElementById('importacao-resumos-pdf');
        const temArquivosPdf = listaPdf && listaPdf.querySelectorAll('table tr').length > 0;
        const resumoReasons = [...reasons];
        if (!temArquivosPdf) resumoReasons.push('Nenhum arquivo PDF disponível');
        btnResumo.disabled = resumoReasons.length > 0;
        updateButtonTooltip(btnResumo, resumoReasons);
    }
}

function ocultarResultadoImportacao() {
    const resultadoCol = document.getElementById('importacao-resultado-col');
    const linkCol = document.getElementById('importacao-link-col');
    if (!resultadoCol || resultadoCol.classList.contains('d-none')) return;

    resultadoCol.classList.add('d-none');
    resultadoCol.classList.remove('col-sm-6');
    linkCol.classList.remove('col-sm-6');
    linkCol.classList.add('col');
}

function exibirResultadoImportacao(data) {
    const resultadoCol = document.getElementById('importacao-resultado-col');
    const linkCol = document.getElementById('importacao-link-col');
    const grid = document.getElementById('importacao-resultado-widgets');
    const msgDiv = document.getElementById('importacao-resultado-msg');
    if (!resultadoCol || !grid || !msgDiv) return;

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

    resultadoCol.classList.remove('d-none');
    resultadoCol.classList.add('col-sm-6');
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