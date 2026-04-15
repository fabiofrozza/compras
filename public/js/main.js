// Sinaliza que JavaScript está habilitado (oculta avisos de noscript)
document.documentElement.classList.replace('no-js', 'js-enabled');

let ws;
let selectedFiles = {}; // Armazena arquivos selecionados por containerId
let wsDisconnectedAlertShown = false;
let wsDisconnectedNotifId = null;
let internetOfflineNotifId = null;
let rNotFoundNotifId = null;

// Estado global da aplicação — fonte de verdade para condições que não derivam do DOM
const AppState = {
    serverConnected: false,
    rAvailable: false,
};

function setState(patch) {
    Object.assign(AppState, patch);
    if (typeof evaluateAllButtons === 'function') evaluateAllButtons();
}

function handleInternetOffline() {
    if (internetOfflineNotifId === null) {
        showToast('Sem conexão com a internet', 'warning', 5000, 'Sistema');
        addNotification({ message: 'Sem conexão com a internet.', type: 'warning', source: 'Sistema' })
            .then(id => { internetOfflineNotifId = id; });
    }
    if (typeof evaluateAllButtons === 'function') evaluateAllButtons();
}

function handleInternetOnline() {
    if (internetOfflineNotifId !== null) {
        showToast('Conexão com a internet restaurada', 'success', 3000, 'Sistema');
        dismissNotification(internetOfflineNotifId);
        internetOfflineNotifId = null;
    }
    if (typeof evaluateAllButtons === 'function') evaluateAllButtons();
}

window.addEventListener('offline', handleInternetOffline);
window.addEventListener('online', handleInternetOnline);

function connectWebSocket() {
    ws = new WebSocket(`ws://${window.location.host}`);

    ws.onopen = () => {
        if (wsDisconnectedAlertShown) {
            showToast('Reconectado ao servidor', 'success', 3000, 'inicialização');
            if (wsDisconnectedNotifId !== null) {
                dismissNotification(wsDisconnectedNotifId);
                wsDisconnectedNotifId = null;
            }
        }
        wsDisconnectedAlertShown = false;
        setState({ serverConnected: true });
    };

    ws.onmessage = (event) => {
        try {
            const data = JSON.parse(event.data);

            if (data.type === 'r-status') {
                setState({ rAvailable: data.available });
                if (!data.available && rNotFoundNotifId === null) {
                    showToast('R não encontrado. Acesse a aba Instalação para baixar e instalar.', 'warning', 8000, 'Sistema');
                    addNotification({
                        message: 'R não está instalado neste computador. Sem ele, não é possível executar os scripts da aplicação.',
                        type: 'warning',
                        source: 'Sistema',
                        actions: [{
                            label: 'Ir para Instalação',
                            callback: () => {
                                const tabBtn = document.getElementById('instalacao-tab');
                                if (tabBtn) new bootstrap.Tab(tabBtn).show();
                            }
                        }]
                    }).then(id => { rNotFoundNotifId = id; });
                } else if (data.available && rNotFoundNotifId !== null) {
                    dismissNotification(rNotFoundNotifId);
                    rNotFoundNotifId = null;
                }
                return;
            }

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

            if (data.type === 'config_data') {
                if (data.scriptName === 'importacao' && typeof exibirResultadoImportacao === 'function') {
                    exibirResultadoImportacao(data.data);
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

                // Sugestão extra: ao falhar um script R, orientar o usuário a atualizar pacotes
                const nonRScripts = ['npm_update', 'atas_mailmerge'];
                if (data.type === 'error' && data.scriptName && !nonRScripts.includes(data.scriptName)) {
                    addNotification({
                        message: 'Se o erro persistir, tente atualizar os pacotes R na aba Instalação.',
                        type: 'info',
                        source: 'Sistema',
                        actions: [{
                            label: 'Ir para Instalação',
                            callback: () => {
                                const tabBtn = document.getElementById('instalacao-tab');
                                if (tabBtn) new bootstrap.Tab(tabBtn).show();
                            }
                        }]
                    });
                }

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

                        if (data.scriptName === 'npm_update' && typeof carregarNpmPackages === 'function') {
                            carregarNpmPackages();
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
        setState({ serverConnected: false });
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
    if (!navigator.onLine) handleInternetOffline();
    await loadConfig();
    await initPreferencesPage();
});
