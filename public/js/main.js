let ws; // Variável global para WebSocket
let selectedFiles = {}; // Armazena arquivos selecionados por containerId

// Conectar ao WebSocket
function connectWebSocket() {
    ws = new WebSocket('ws://localhost:3000');

    ws.onopen = () => {
        showToast('Conectado ao servidor', 'success', 3000, 'inicialização');
    };

    ws.onmessage = (event) => {
        try {
            const data = JSON.parse(event.data);

            // Processar output e progresso apenas se o console estiver visível
            if (consoleContainer && consoleContainer.classList.contains('show')) {
                if (data.type === 'output') {
                    // Escolher a formatação baseada no nível do log
                    let colorStyle = '';
                    switch (data.level) {
                        case 'error': colorStyle = 'color: #ff6b6b; font-weight: bold;'; break; // Vermelho
                        case 'warning': colorStyle = 'color: #ffc107; font-weight: bold;'; break; // Amarelo
                        case 'success': colorStyle = 'color: #20c997; font-weight: bold;'; break; // Verde
                        case 'command': colorStyle = 'color: #0dcaf0;'; break; // Ciano
                        case 'section': colorStyle = 'color: #6ea8fe; font-weight: bold;'; break; // Azul
                        case 'info':
                        default: colorStyle = 'color: inherit;'; break;
                    }
                    consoleOutput.innerHTML += `<span style="${colorStyle}">${data.message}\n</span>`;
                    consoleOutput.scrollTop = consoleOutput.scrollHeight;
                } else if (data.type === 'progress') {
                    if (typeof updateProgressBar === 'function') {
                        updateProgressBar(data);
                    }
                }
            }

            // Finalização do script: SEMPRE processar, independente do console estar visível
            // Isso garante que isScriptRunning seja resetado e o overlay seja removido
            if (data.type === 'success' || data.type === 'warning' || data.type === 'error') {
                // Finalizar o console e resetar a flag isScriptRunning
                const result = {
                    status: data.type,
                    message: data.message,
                    log: consoleOutput ? consoleOutput.innerHTML : '',
                    scriptName: data.scriptName
                };
                handleScriptResult(result);

                // Registrar na central de notificações
                const tabNames = {
                    atas: 'Atas', atas_mailmerge: 'Atas (Mailmerge)', catmat: 'Catmat',
                    fornecedores: 'Fornecedores', importacao: 'Importação',
                    mapas: 'Mapas', powerbi: 'Power BI', instalacao: 'Instalação'
                };
                const sourceName = tabNames[data.scriptName] || data.scriptName;
                const typeLabels = { success: 'concluído com sucesso', warning: 'concluído com alertas', error: 'falhou' };
                addNotification({
                    message: `Script "${sourceName}" ${typeLabels[data.type] || 'finalizado'}${data.message ? ': ' + data.message : ''}`,
                    type: data.type,
                    source: sourceName
                });

                // Atualizar as listas de arquivos
                if (data.scriptName) {
                    setTimeout(async () => {
                        const tabName = data.scriptName === 'atas_mailmerge' ? 'atas' : data.scriptName;
                        await refreshScriptFileLists(tabName);

                        if (data.scriptName === 'atas_mailmerge' && typeof carregarAtasFinalizadas === 'function') {
                            carregarAtasFinalizadas();
                        }

                        if (data.scriptName === 'atas' && typeof verificarStatusDadosAtas === 'function') {
                            verificarStatusDadosAtas();
                        }

                        if (document.getElementById('logs-file-list')) {
                            loadFiles('logs-file-list', '_common', 'log', false);
                        }
                    }, 500);
                }
            }
        } catch (e) {
            // Ignorar erros de parse silenciosamente
        }
    };

    ws.onclose = () => {
        // Resetar flag de script em execução para evitar bloqueio permanente após desconexão
        if (isScriptRunning) {
            hideScriptRunningOverlay();
        }
        showToast('Desconectado do servidor. Tentando reconectar...', 'error', 2000, 'inicialização');
        addNotification({ message: 'Desconectado do servidor. Tentando reconectar...', type: 'warning', source: 'Sistema' });
        // Reconectar após 3 segundos
        setTimeout(connectWebSocket, 3000);
    };

    ws.onerror = (error) => {
        console.error('Erro WebSocket:', error);
    };
}


// Inicializar tudo quando a página carregar
window.addEventListener('load', async () => {
    connectWebSocket();
    await loadConfig();
    await initPreferencesPage();
});
