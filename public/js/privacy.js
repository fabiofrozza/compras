// ====== CONSENTIMENTO DE PRIVACIDADE / GRAVAÇÃO DE LOGS ======

async function setLogConsent(value) {
    appState.preferences.logConsent = value;
    saveAppState();

    try {
        await fetch('/api/log-consent', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ consent: value })
        });
    } catch (e) {
        console.warn('Erro ao sincronizar consentimento com o servidor:', e);
    }

    _applyConsentUI(value);
}

function _applyConsentUI(consent) {
    const toggle = document.getElementById('logConsent');
    if (toggle) toggle.checked = consent === true;
}

document.addEventListener('DOMContentLoaded', () => {
    const consent = appState.preferences.logConsent ?? null;
    _applyConsentUI(consent);

    // Sincroniza o estado com o servidor (que não persiste entre reinicializações).
    // Somente sincroniza com o servidor se houver decisão explícita (true/false).
    // null = ainda sem decisão = NÃO habilita logs automaticamente (opt-in).
    if (consent === true || consent === false) {
        fetch('/api/log-consent', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ consent: consent })
        }).catch(() => { /* servidor pode estar iniciando */ });
    }

    // Se ainda não houve decisão, solicita via painel de notificações.
    // O usuário pode ignorar e será solicitado novamente na próxima sessão.
    if (consent === null) {
        let privacyNotifId = null;
        addNotification({
            message: 'Para possibilitar o tratamento de erros, podem ser salvos logs de execução, incluindo nome de usuário, hostname e detalhes das operações realizadas. Deseja permitir?',
            type: 'info',
            source: 'Privacidade',
            actions: [
                {
                    label: 'Recusar',
                    callback: () => {
                        setLogConsent(false);
                        if (privacyNotifId !== null) dismissNotification(privacyNotifId);
                    }
                },
                {
                    label: 'Aceitar',
                    callback: () => {
                        setLogConsent(true);
                        if (privacyNotifId !== null) dismissNotification(privacyNotifId);
                    }
                }
            ]
        }).then(id => { privacyNotifId = id; });
    }

    document.getElementById('logConsent')?.addEventListener('change', (e) => {
        setLogConsent(e.target.checked);
        if (typeof showPreferencesSaveIndicator === 'function') {
            showPreferencesSaveIndicator();
        }
    });
});
