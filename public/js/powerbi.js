// --- LÓGICA DA ABA POWER BI ---

/**
 * Executa o script R do Power BI com a opção selecionada.
 * Usa a função global `runRScript` (definida em console.js)
 * passando parâmetros customizados.
 */
function executarPowerBI() {
    const selecionado = document.querySelector('input[name="powerbi-tipo"]:checked');
    if (!selecionado) {
        showToast('Selecione uma opção antes de executar.', 'warning', 4000, 'Power BI');
        return;
    }

    // Validar pasta da base de dados
    const pastaInput = document.getElementById('powerbi-pasta');
    const pasta = pastaInput ? pastaInput.value.trim() : '';
    if (!pasta) {
        showToast('Informe a pasta da base de dados antes de executar.', 'warning', 4000, 'Power BI');
        if (pastaInput) pastaInput.focus();
        return;
    }

    const tipo = selecionado.value;

    // Envia para o servidor via WebSocket (console.js → runRScript)
    // O R script receberá: tipo e pasta como argumentos
    runRScript('powerbi', { tipo, pasta });
}
