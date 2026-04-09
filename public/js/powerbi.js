const POWERBI_OPCOES_INTERNET = ['planejamento', 'paalteracoes', 'todos'];

function atualizarBotaoPowerBIPanel() {
    const btn = document.getElementById('btn-run-powerbi-panel');
    if (!btn) return;

    const reasons = [];
    const hasInvalidFields = Array.from(document.querySelectorAll('#form-powerbi-path [data-field]'))
        .some(f => f.classList.contains('is-invalid'));
    if (hasInvalidFields) reasons.push('Preencha os campos obrigatórios');

    const selected = document.querySelector('#form-powerbi-panel input[name="powerbi-tipo"]:checked');
    const necessitaInternet = !selected || POWERBI_OPCOES_INTERNET.includes(selected.value);
    if (necessitaInternet && !navigator.onLine) reasons.push('Sem conexão com a internet');

    btn.disabled = reasons.length > 0;
    updateButtonTooltip(btn, reasons);
}

document.addEventListener('change', (e) => {
    if (e.target.name === 'powerbi-tipo' && e.target.closest('#form-powerbi-panel')) {
        atualizarBotaoPowerBIPanel();
    }
});

function executarPowerBI(button) {
    const form = button ? button.closest('form') : document;
    const selected = form.querySelector('input[name="powerbi-tipo"]:checked');
    if (!selected) {
        showToast('Selecione uma opção antes de executar.', 'warning', 4000, 'Power BI');
        return;
    }

    const pastaInput = document.getElementById('powerbi-pasta');
    if (!pastaInput || !validateSingleField(pastaInput)) {
        showToast('Informe um caminho válido para a pasta da base de dados antes de executar.', 'warning', 4000, 'Power BI');
        if (pastaInput) pastaInput.focus();
        return;
    }
    const pasta = pastaInput.value.trim();

    const tipo = selected.value;

    runRScript('powerbi', { tipo, pasta });
}

let observatorioRegistros = [];
let observatorioFiltroProblemas = false;
let observatorioOrdenacao = { coluna: null, direcao: 'asc' };

function mostrarProgressoObservatorio() {
    const container = document.getElementById('powerbi-observatorio-progress-container');
    const bar = document.getElementById('powerbi-observatorio-progress-bar');
    if (!container || !bar) return;
    container.classList.remove('d-none');
    bar.style.width = '0%';
    bar.innerText = '';
    const popup = document.createElement('span');
    popup.className = 'progress-percent-popup';
    popup.textContent = '0%';
    bar.appendChild(popup);
}

function atualizarProgressoObservatorio(current, total, label) {
    const bar = document.getElementById('powerbi-observatorio-progress-bar');
    if (!bar) return;
    const percentage = total > 0 ? (current / total) * 100 : 0;
    bar.style.width = `${percentage}%`;
    if (label) bar.innerText = label;
    let popup = bar.querySelector('.progress-percent-popup');
    if (!popup) {
        popup = document.createElement('span');
        popup.className = 'progress-percent-popup';
        bar.appendChild(popup);
    } else if (label) {
        bar.appendChild(popup); // re-append pois innerText remove filhos
    }
    popup.textContent = `${Math.round(percentage)}%`;
}

function esconderProgressoObservatorio() {
    const container = document.getElementById('powerbi-observatorio-progress-container');
    if (container) container.classList.add('d-none');
}

function popularFiltrosObservatorio() {
    const anos = [...new Set(observatorioRegistros.map(r => r.ano))].sort();
    const situacoes = [...new Set(observatorioRegistros.map(r => r.situacao).filter(Boolean))].sort();
    const statusLicitacao = [...new Set(observatorioRegistros.map(r => r.obsLicitacaoStatus).filter(Boolean))];
    const statusExecucao = [...new Set(observatorioRegistros.map(r => r.obsExecucaoStatus).filter(Boolean))];

    const buildStatusOptions = values =>
        values.map(v => `<option value="${escapeHtml(v)}">${escapeHtml(OBSERVATORIO_STATUS_LABEL[v] || v)}</option>`).join('');

    document.getElementById('filtro-obs-ano').innerHTML =
        '<option value="">Todos</option>' + anos.map(a => `<option value="${a}">${a}</option>`).join('');
    document.getElementById('filtro-obs-situacao').innerHTML =
        '<option value="">Todas</option>' + situacoes.map(s => `<option value="${escapeHtml(s)}">${escapeHtml(s)}</option>`).join('');
    document.getElementById('filtro-obs-licitacao').innerHTML =
        '<option value="">Todos</option>' + buildStatusOptions(statusLicitacao);
    document.getElementById('filtro-obs-execucao').innerHTML =
        '<option value="">Todos</option>' + buildStatusOptions(statusExecucao);
}

function criarLinhaObservatorio(r) {
    const tr = document.createElement('tr');
    tr.innerHTML = `
        <td>${r.ano}</td>
        <td>${escapeHtml(r.processo)}</td>
        <td>${escapeHtml(r.situacao)}</td>
        <td>${escapeHtml(r.dataFinalizacao)}</td>
        <td>${renderObservatorioBadge(r.obsLicitacao, r.obsLicitacaoStatus)}</td>
        <td>${renderObservatorioBadge(r.obsExecucao, r.obsExecucaoStatus)}</td>
    `;
    return tr;
}

function aplicarFiltrosObservatorio() {
    const ano = document.getElementById('filtro-obs-ano').value;
    const processo = document.getElementById('filtro-obs-processo').value.trim().toLowerCase();
    const situacao = document.getElementById('filtro-obs-situacao').value;
    const data = document.getElementById('filtro-obs-data').value.trim().toLowerCase();
    const licitacao = document.getElementById('filtro-obs-licitacao').value;
    const execucao = document.getElementById('filtro-obs-execucao').value;

    let registrosFiltrados = observatorioRegistros.filter(r => {
        if (ano && String(r.ano) !== ano) return false;
        if (processo && !r.processo.toLowerCase().includes(processo)) return false;
        if (situacao && r.situacao !== situacao) return false;
        if (data && !String(r.dataFinalizacao ?? '').toLowerCase().includes(data)) return false;
        if (observatorioFiltroProblemas) {
            const PROBLEMAS = ['pendente', 'divergente'];
            if (!PROBLEMAS.includes(r.obsLicitacaoStatus) && !PROBLEMAS.includes(r.obsExecucaoStatus)) return false;
        } else {
            if (licitacao && r.obsLicitacaoStatus !== licitacao) return false;
            if (execucao && r.obsExecucaoStatus !== execucao) return false;
        }
        return true;
    });

    if (observatorioOrdenacao.coluna) {
        const { coluna, direcao } = observatorioOrdenacao;
        registrosFiltrados = [...registrosFiltrados].sort((a, b) => {
            const va = getObservatorioSortValue(a, coluna);
            const vb = getObservatorioSortValue(b, coluna);
            const cmp = va < vb ? -1 : va > vb ? 1 : 0;
            return direcao === 'asc' ? cmp : -cmp;
        });
    }

    const tbody = document.getElementById('powerbi-observatorio-tbody');
    tbody.innerHTML = '';
    const fragment = document.createDocumentFragment();
    for (const r of registrosFiltrados) fragment.appendChild(criarLinhaObservatorio(r));
    tbody.appendChild(fragment);

    const temFiltroOuOrdenacao = ano || processo || situacao || data || licitacao || execucao || observatorioFiltroProblemas || observatorioOrdenacao.coluna !== null;
    const btnTopo = document.getElementById('btn-limpar-filtros-obs-bar');
    if (btnTopo) btnTopo.hidden = !temFiltroOuOrdenacao;
    const btnBaixo = document.getElementById('btn-limpar-filtros-obs-bar-bottom');
    if (btnBaixo) btnBaixo.hidden = !temFiltroOuOrdenacao;

    const temFiltro = ano || processo || situacao || data || licitacao || execucao || observatorioFiltroProblemas;
    const countText = temFiltro
        ? `${registrosFiltrados.length} de ${observatorioRegistros.length} processo(s)`
        : `${observatorioRegistros.length} processo(s)`;
    document.getElementById('observatorio-info-count').textContent = countText;
    const countBottomEl = document.getElementById('observatorio-info-count-bottom');
    if (countBottomEl) countBottomEl.textContent = countText;

    atualizarDestaqueCabecalhos();
}

function toggleFiltroProblemas(button) {
    observatorioFiltroProblemas = !observatorioFiltroProblemas;
    button.classList.toggle('active', observatorioFiltroProblemas);

    const selLicitacao = document.getElementById('filtro-obs-licitacao');
    const selExecucao = document.getElementById('filtro-obs-execucao');
    selLicitacao.disabled = observatorioFiltroProblemas;
    selExecucao.disabled = observatorioFiltroProblemas;
    if (observatorioFiltroProblemas) {
        selLicitacao.value = '';
        selExecucao.value = '';
    }
    aplicarFiltrosObservatorio();
}

function limparFiltrosEOrdenacaoObservatorio() {
    observatorioOrdenacao = { coluna: null, direcao: 'asc' };
    atualizarIndicadoresOrdenacao();

    document.getElementById('filtro-obs-ano').value = '';
    document.getElementById('filtro-obs-processo').value = '';
    document.getElementById('filtro-obs-situacao').value = '';
    document.getElementById('filtro-obs-data').value = '';
    document.getElementById('filtro-obs-licitacao').value = '';
    document.getElementById('filtro-obs-licitacao').disabled = false;
    document.getElementById('filtro-obs-execucao').value = '';
    document.getElementById('filtro-obs-execucao').disabled = false;
    if (observatorioFiltroProblemas) {
        observatorioFiltroProblemas = false;
        document.getElementById('btn-filtro-problemas')?.classList.remove('active');
    }
    aplicarFiltrosObservatorio();
}

function atualizarDestaqueCabecalhos() {
    const filtrosAtivos = {
        ano: !!document.getElementById('filtro-obs-ano').value,
        processo: !!document.getElementById('filtro-obs-processo').value.trim(),
        situacao: !!document.getElementById('filtro-obs-situacao').value,
        data: !!document.getElementById('filtro-obs-data').value.trim(),
        licitacao: observatorioFiltroProblemas || !!document.getElementById('filtro-obs-licitacao').value,
        execucao: observatorioFiltroProblemas || !!document.getElementById('filtro-obs-execucao').value,
    };

    // Cabeçalhos (Linha 1): Destaca se tiver filtro OU ordenação
    document.querySelectorAll('#powerbi-observatorio-result thead tr:first-child th[data-col]').forEach(th => {
        const col = th.dataset.col;
        const temFiltroOuOrdenacao = filtrosAtivos[col] || (observatorioOrdenacao.coluna === col);
        th.classList.toggle('obs-header-highlight', temFiltroOuOrdenacao);
    });

    // Células de filtro (Linha 2): Destaca se tiver filtro
    document.querySelectorAll('#observatorio-filter-row th[data-col]').forEach(th => {
        const col = th.dataset.col;
        th.classList.toggle('obs-filter-highlight', filtrosAtivos[col]);
    });
}

function getObservatorioSortValue(r, coluna) {
    switch (coluna) {
        case 'ano': return r.ano;
        case 'processo': return r.processo.toLowerCase();
        case 'situacao': return (r.situacao ?? '').toLowerCase();
        case 'data': {
            const [d, m, y] = (r.dataFinalizacao || '').split('/');
            return y && m && d ? `${y}${m}${d}` : '';
        }
        case 'licitacao': return (r.obsLicitacao ?? '').toLowerCase();
        case 'execucao': return (r.obsExecucao ?? '').toLowerCase();
        default: return '';
    }
}

const OBSERVATORIO_COL_LABELS = {
    ano: 'Ano', processo: 'Processo', situacao: 'Situação',
    data: 'Data finalização', licitacao: 'Obs. Licitação', execucao: 'Obs. Execução',
};

function atualizarIndicadoresOrdenacao() {
    document.querySelectorAll('#powerbi-observatorio-result th[data-col]').forEach(th => {
        const icon = th.querySelector('.sort-icon');
        if (!icon) return;
        const ativa = observatorioOrdenacao.coluna === th.dataset.col;
        icon.textContent = ativa
            ? (observatorioOrdenacao.direcao === 'asc' ? 'arrow_upward' : 'arrow_downward')
            : 'unfold_more';
        th.classList.toggle('sorted', ativa);
    });
    const colLabel = OBSERVATORIO_COL_LABELS[observatorioOrdenacao.coluna] ?? '';
    const sorted = !!observatorioOrdenacao.coluna;

    const indicator = document.getElementById('observatorio-sort-indicator');
    const label = document.getElementById('observatorio-sort-label');
    if (indicator && label) {
        indicator.hidden = !sorted;
        label.textContent = colLabel;
    }

    const indicatorBottom = document.getElementById('observatorio-sort-indicator-bottom');
    const labelBottom = document.getElementById('observatorio-sort-label-bottom');
    if (indicatorBottom && labelBottom) {
        indicatorBottom.hidden = !sorted;
        labelBottom.textContent = colLabel;
    }
}

function ordenarObservatorio(coluna) {
    if (observatorioOrdenacao.coluna === coluna) {
        if (observatorioOrdenacao.direcao === 'asc') {
            observatorioOrdenacao.direcao = 'desc';
        } else {
            observatorioOrdenacao = { coluna: null, direcao: 'asc' };
        }
    } else {
        observatorioOrdenacao = { coluna, direcao: 'asc' };
    }
    atualizarIndicadoresOrdenacao();
    aplicarFiltrosObservatorio();
}



document.getElementById('observatorio-filter-row')?.addEventListener('input', aplicarFiltrosObservatorio);
document.getElementById('observatorio-filter-row')?.addEventListener('change', aplicarFiltrosObservatorio);

function executarObservatorio(button) {
    if (!navigator.onLine) {
        showToast('Sem conexão com a internet.', 'warning', 4000, 'Observatório');
        return;
    }
    const pastaInput = document.getElementById('powerbi-pasta');
    if (!pastaInput || !validateSingleField(pastaInput)) {
        showToast('Informe um caminho válido para a pasta da base de dados antes de executar.', 'warning', 4000, 'Observatório');
        if (pastaInput) pastaInput.focus();
        return;
    }
    const pasta = pastaInput.value.trim();

    const status = document.getElementById('powerbi-observatorio-status');
    const result = document.getElementById('powerbi-observatorio-result');
    const tbody = document.getElementById('powerbi-observatorio-tbody');
    const downloadBtn = document.getElementById('btn-download-powerbi-observatorio');

    button.disabled = true;
    status.textContent = 'Acessando a Planilha de Controle...';
    status.classList.remove('text-danger');
    result.hidden = true;
    tbody.innerHTML = '';
    downloadBtn.hidden = true;
    mostrarProgressoObservatorio();

    const es = new EventSource(`/api/observatorio/planilha-controle?pasta=${encodeURIComponent(pasta)}`);
    let finalizado = false;

    const finalizar = (msgErro) => {
        finalizado = true;
        es.close();
        esconderProgressoObservatorio();
        button.disabled = false;
        if (msgErro) {
            status.textContent = `Erro: ${msgErro}`;
            status.classList.add('text-danger');
            showToast(`Falha ao recuperar dados: ${msgErro}`, 'error', 5000, 'Observatório');
        }
    };

    es.addEventListener('progress', (e) => {
        const d = JSON.parse(e.data);
        atualizarProgressoObservatorio(d.current, d.total, d.label);
    });

    es.addEventListener('done', (e) => {
        const data = JSON.parse(e.data);
        observatorioRegistros = data.registros || [];

        observatorioFiltroProblemas = false;
        observatorioOrdenacao = { coluna: null, direcao: 'asc' };
        document.getElementById('btn-filtro-problemas')?.classList.remove('active');
        document.getElementById('filtro-obs-licitacao').disabled = false;
        document.getElementById('filtro-obs-execucao').disabled = false;

        popularFiltrosObservatorio();
        atualizarIndicadoresOrdenacao();
        aplicarFiltrosObservatorio();

        result.hidden = false;
        downloadBtn.hidden = observatorioRegistros.length === 0;
        document.getElementById('btn-filtro-problemas').hidden = observatorioRegistros.length === 0;

        const partes = [];
        if (data.abasComErro && data.abasComErro.length) {
            const detalhes = data.abasComErro.map(e => `${e.ano} (${e.erro})`).join('; ');
            partes.push(`Abas ignoradas: ${detalhes}.`);
        }
        if (data.arquivosComErro && data.arquivosComErro.length) {
            const detalhes = data.arquivosComErro.map(a => `${a.arquivo} (${a.erro})`).join('; ');
            partes.push(`Comparação não realizada para: ${detalhes}.`);
        }
        status.textContent = partes.join(' ');
        finalizar();
    });

    es.addEventListener('fail', (e) => {
        let msg = 'Erro desconhecido';
        try { msg = JSON.parse(e.data).message || msg; } catch { /* ignore */ }
        finalizar(msg);
    });

    es.onerror = () => {
        if (!finalizado) finalizar('Falha na conexão com o servidor');
    };
}

function escapeHtml(str) {
    return String(str ?? '').replace(/[&<>"']/g, c => ({
        '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
    }[c]));
}

const OBSERVATORIO_STATUS_INFO = {
    atendido: { variant: 'success', titulo: 'Arquivo já incluído na base' },
    pendente: { variant: 'warning', titulo: 'Incluir o arquivo correspondente na base' },
    divergente: { variant: 'danger', titulo: 'Divergência entre a Planilha de Controle e a base — verificar' },
    na: { variant: 'secondary', titulo: 'Não se aplica' },
    analise: { variant: '', titulo: 'Em análise' },
};

function renderObservatorioBadge(texto, status) {
    if (!texto) return '';
    const info = OBSERVATORIO_STATUS_INFO[status] || OBSERVATORIO_STATUS_INFO.analise;
    if (!info.variant) return escapeHtml(texto);
    return `<span class="badge text-bg-${info.variant} text-wrap" title="${escapeHtml(info.titulo)}">${escapeHtml(texto)}</span>`;
}

const OBSERVATORIO_STATUS_LABEL = {
    atendido: 'Atendido',
    pendente: 'Pendente',
    divergente: 'Divergente',
    na: 'Não se aplica',
    analise: 'Em análise',
};

function baixarObservatorioCsv() {
    if (!observatorioRegistros.length) return;

    const cabecalho = [
        'Ano', 'Processo (CAPL)', 'Situação', 'Data de finalização',
        'Observatório - Licitação', 'Status Licitação',
        'Observatório - Execução', 'Status Execução',
    ];
    const escapeCsv = v => {
        const s = String(v ?? '');
        return /[",;\n\r]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
    };
    const statusLabel = s => OBSERVATORIO_STATUS_LABEL[s] || '';
    const linhas = [cabecalho.map(escapeCsv).join(';')];
    for (const r of observatorioRegistros) {
        linhas.push([
            r.ano, r.processo, r.situacao, r.dataFinalizacao,
            r.obsLicitacao, statusLabel(r.obsLicitacaoStatus),
            r.obsExecucao, statusLabel(r.obsExecucaoStatus),
        ].map(escapeCsv).join(';'));
    }

    const blob = new Blob(['\uFEFF' + linhas.join('\r\n')], { type: 'text/csv;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `observatorio_planilha_controle_${new Date().toISOString().slice(0, 10)}.csv`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}
