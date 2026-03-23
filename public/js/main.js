let ws;
let selectedFiles = {}; // Armazena arquivos selecionados por containerId
let wsDisconnectedAlertShown = false;
let wsDisconnectedNotifId = null;

function connectWebSocket() {
    ws = new WebSocket('ws://localhost:3000');

    ws.onopen = () => {
        if (wsDisconnectedAlertShown) {
            showToast('Reconectado ao servidor', 'success', 3000, 'inicialização');
            if (wsDisconnectedNotifId !== null) {
                dismissNotification(wsDisconnectedNotifId);
                wsDisconnectedNotifId = null;
            }
        }
        wsDisconnectedAlertShown = false;
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
                        case 'error': colorStyle = 'color: #ff6b6b; font-weight: bold;'; break;
                        case 'warning': colorStyle = 'color: #ffc107; font-weight: bold;'; break;
                        case 'success': colorStyle = 'color: #20c997; font-weight: bold;'; break;
                        case 'command': colorStyle = 'color: #0dcaf0;'; break;
                        case 'section': colorStyle = 'color: #6ea8fe; font-weight: bold;'; break;
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

                const notificationMessage = data.notificationMessage || data.message;
                addNotification({
                    message: notificationMessage,
                    type: data.type,
                    source: data.scriptName
                });

                handleScriptResult({
                    status: data.type,
                    message: notificationMessage,
                    log: consoleOutput ? consoleOutput.innerHTML : '',
                    scriptName: data.scriptName
                });

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

                        if (data.scriptName === 'fornecedores') {
                            if (typeof carregarPregoes === 'function') carregarPregoes();
                            if (typeof carregarImportar === 'function') carregarImportar();
                            if (typeof atualizarInfoPregao === 'function' && pregaoSelecionado) {
                                atualizarInfoPregao(pregaoSelecionado);
                                carregarConteudoPasta(pregaoSelecionado);
                            }
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
        if (isScriptRunning) hideScriptRunningOverlay();
        if (!wsDisconnectedAlertShown) {
            wsDisconnectedAlertShown = true;
            showToast('Desconectado. Execute novamente o arquivo start.cmd para reiniciar o servidor.', 'error', 10000, 'inicialização');
            addNotification({ message: 'Desconectado. Execute novamente o arquivo start.cmd para reiniciar o servidor.', type: 'error', source: 'Sistema' })
                .then(id => { wsDisconnectedNotifId = id; });
        }
        setTimeout(connectWebSocket, 3000);
    };

    ws.onerror = (error) => {
        console.error('Erro WebSocket:', error);
    };
}

window.addEventListener('load', async () => {
    connectWebSocket();
    await loadConfig();
    await initPreferencesPage();
});
