// --- LÓGICA DA ABA POWER BI ---

/** Descrições de cada opção do Power BI (reproduzidas do script AHK) */
const powerbiDescricoes = {
    planejamento:
        'Painel Visão Planejamento — Será acessada a planilha de controle para recuperar os dados dos processos, Unidades requerentes e situação do envio da documentação.',
    licitacao:
        'Painel Visão Licitação — Serão obtidos os dados dos Mapas de Licitação da pasta POWERBI/Mapa de licitações.',
    execucao:
        'Painel Visão Execução — Serão obtidos os dados dos relatórios de execução das AFs/Empenhos da pasta POWERBI/Execucao AF Empenho.',
    paalteracoes:
        'Painel Processos Administrativos e Alterações Contratuais — Será acessada a planilha de controle para recuperar os dados dos processos administrativos, trocas de marca, cancelamentos e reequilíbrios.',
    renomear:
        'Os arquivos das pastas Mapa de licitações e Execucao AF Empenho serão renomeados conforme o padrão TIPO - ANO - ETAPA - PROCESSO - PREGÃO.',
    todos:
        'Serão gerados dados para todos os Painéis do Observatório.'
};

/**
 * Atualiza a descrição exibida quando o usuário troca de opção.
 * Chamada pelo event delegation abaixo.
 */
function atualizarDescricaoPowerBI() {
    const selecionado = document.querySelector('input[name="powerbi-tipo"]:checked');
    const textoEl = document.getElementById('powerbi-descricao-texto');
    if (selecionado && textoEl) {
        textoEl.textContent = powerbiDescricoes[selecionado.value] || '';
    }
}

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
    const descricao = powerbiDescricoes[tipo] || '';

    // Confirmar antes de executar
    const confirmar = confirm(`${descricao}\n\nDeseja continuar?`);
    if (!confirmar) return;

    // Envia para o servidor via WebSocket (console.js → runRScript)
    // O R script receberá: tipo e pasta como argumentos
    runRScript('powerbi', { tipo, pasta });
}

// Event delegation para trocar a descrição quando o radio muda
document.addEventListener('change', (event) => {
    if (event.target.name === 'powerbi-tipo') {
        atualizarDescricaoPowerBI();
    }
});
