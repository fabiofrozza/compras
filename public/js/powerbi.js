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

function executarObservatorio(button) {
    if (!navigator.onLine) {
        showToast('Sem conexão com a internet.', 'warning', 4000, 'Observatório');
        return;
    }
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

    const es = new EventSource('/api/observatorio/planilha-controle');
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

        const fragment = document.createDocumentFragment();
        for (const r of observatorioRegistros) {
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td>${r.ano}</td>
                <td>${escapeHtml(r.processo)}</td>
                <td>${escapeHtml(r.situacao)}</td>
                <td>${escapeHtml(r.dataFinalizacao)}</td>
                <td>${escapeHtml(r.obsLicitacao)}</td>
                <td>${escapeHtml(r.obsExecucao)}</td>
            `;
            fragment.appendChild(tr);
        }
        tbody.appendChild(fragment);
        result.hidden = false;
        downloadBtn.hidden = observatorioRegistros.length === 0;

        const partes = [`${data.total} processo(s) recuperado(s) das abas ${data.anos.join(', ')}.`];
        if (data.abasComErro && data.abasComErro.length) {
            const detalhes = data.abasComErro.map(e => `${e.ano} (${e.erro})`).join('; ');
            partes.push(`Abas ignoradas: ${detalhes}.`);
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

function baixarObservatorioCsv() {
    if (!observatorioRegistros.length) return;

    const cabecalho = ['Ano', 'Processo (CAPL)', 'Situação', 'Data de finalização', 'Observatório - Licitação', 'Observatório - Execução'];
    const escapeCsv = v => {
        const s = String(v ?? '');
        return /[",;\n\r]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
    };
    const linhas = [cabecalho.map(escapeCsv).join(';')];
    for (const r of observatorioRegistros) {
        linhas.push([r.ano, r.processo, r.situacao, r.dataFinalizacao, r.obsLicitacao, r.obsExecucao].map(escapeCsv).join(';'));
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
