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
