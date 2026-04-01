const POWERBI_OPCOES_INTERNET = ['planejamento', 'paalteracoes', 'todos'];

function atualizarBotaoPowerBIPanel() {
    const btn = document.getElementById('btn-run-powerbi-panel');
    if (!btn) return;
    const hasInvalidFields = Array.from(document.querySelectorAll('#form-powerbi-path [data-field]'))
        .some(f => f.classList.contains('is-invalid'));
    const selected = document.querySelector('#form-powerbi-panel input[name="powerbi-tipo"]:checked');
    const necessitaInternet = !selected || POWERBI_OPCOES_INTERNET.includes(selected.value);
    btn.disabled = hasInvalidFields || (necessitaInternet && !navigator.onLine);
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
