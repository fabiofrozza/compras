let _rDownloadUrl = '';

function executarInstalacao(modo) {
    runRScript('instalacao', { modo });
}

function executarInstalacaoSelecionada() {
    const modo = document.querySelector('input[name="instalacao-modo"]:checked')?.value;
    if (modo) executarInstalacao(modo);
}

function executarNpmUpdate() {
    if (isScriptRunning) {
        showToast('Aguarde o término do script em execução antes de iniciar outro.', 'warning', 5000, 'execução');
        return;
    }

    if (!ws || ws.readyState !== WebSocket.OPEN) {
        showToast('Servidor não está conectado. Execute novamente o arquivo start.cmd para iniciar o servidor.', 'error', 10000, 'inicialização');
        return;
    }

    prepareConsoleForExecution('npm_update');

    try {
        ws.send(JSON.stringify({ action: 'execute-npm-update' }));
    } catch (error) {
        console.error('Erro ao enviar mensagem:', error);
        handleScriptResult({
            status: 'error',
            message: `Falha ao enviar comando para o servidor: ${error.message}`,
            log: '',
            scriptName: 'npm_update'
        });
    }
}

function downloadR() {
    if (!_rDownloadUrl) return;
    window.open(_rDownloadUrl, '_blank');
    addNotification({
        message: 'Download do instalador do R iniciado.'
            + '<ul class="mb-0 ps-3">'
            + '<li>Aguarde o término do download.</li>'
            + '<li>Execute-o, aceitando as opções padrão.</li>'
            + '<li>Reinicie esta aplicação após a instalação.</li>'
            + '<li>Versões anteriores podem ser desinstaladas pelo Painel de Controle, se desejar.</li>'
            + '</ul>',
        type: 'info',
        source: 'Instalação'
    });
}

function compareVersions(a, b) {
    const pa = a.split('.').map(Number);
    const pb = b.split('.').map(Number);
    for (let i = 0; i < Math.max(pa.length, pb.length); i++) {
        const na = pa[i] || 0;
        const nb = pb[i] || 0;
        if (na > nb) return 1;
        if (na < nb) return -1;
    }
    return 0;
}

async function carregarInfoR() {
    const loading = document.getElementById('r-info-loading');
    const content = document.getElementById('r-info-content');
    const versionEl = document.getElementById('r-info-version');
    const pathEl = document.getElementById('r-info-path');
    const latestEl = document.getElementById('r-info-latest');
    const statusEl = document.getElementById('r-info-status');
    const downloadBtn = document.getElementById('btn-download-r');

    if (!content) return;

    loading.classList.remove('d-none');
    content.classList.add('d-none');
    downloadBtn.classList.add('d-none');

    // Consulta versão local e versão mais recente em paralelo
    const [localRes, latestRes] = await Promise.allSettled([
        fetch('/api/check-r').then(r => r.json()),
        fetch('/api/check-r-latest').then(r => r.json())
    ]);

    loading.classList.add('d-none');
    content.classList.remove('d-none');

    // Versão instalada
    let installedVersion = '';
    if (localRes.status === 'fulfilled' && localRes.value.installed) {
        const versionMatch = localRes.value.version.match(/(\d+\.\d+\.\d+)/);
        installedVersion = versionMatch ? versionMatch[1] : '';
        versionEl.textContent = installedVersion || localRes.value.version;
        pathEl.textContent = localRes.value.path;
    } else {
        versionEl.textContent = 'R não encontrado';
        versionEl.classList.add('text-danger');
        pathEl.textContent = '-';
    }

    // Versão mais recente
    if (latestRes.status === 'fulfilled' && latestRes.value.latest) {
        const latest = latestRes.value.latest;
        latestEl.innerHTML = `<a href="https://cran.r-project.org/bin/windows/base/" target="_blank">${latest}</a>`;
        _rDownloadUrl = latestRes.value.downloadUrl;

        if (installedVersion) {
            const cmp = compareVersions(installedVersion, latest);
            if (cmp >= 0) {
                statusEl.textContent = 'Atualizado';
                statusEl.className = 'badge ms-2 text-bg-success';
            } else {
                statusEl.textContent = 'Nova versão disponível';
                statusEl.className = 'badge ms-2 text-bg-warning';
                downloadBtn.classList.remove('d-none');
            }
        } else {
            downloadBtn.classList.remove('d-none');
        }
    } else {
        latestEl.textContent = 'Não foi possível verificar';
        latestEl.classList.add('text-body-secondary');
        statusEl.textContent = '';
    }
}

async function carregarInfoNode() {
    const loading = document.getElementById('node-info-loading');
    const content = document.getElementById('node-info-content');
    if (!content) return;

    loading.classList.remove('d-none');
    content.classList.add('d-none');

    const [localRes, latestRes] = await Promise.allSettled([
        fetch('/api/check-node').then(r => r.json()),
        fetch('/api/check-node-latest').then(r => r.json())
    ]);

    loading.classList.add('d-none');
    content.classList.remove('d-none');

    const versionEl = document.getElementById('node-info-version');
    const pathEl = document.getElementById('node-info-path');
    const latestEl = document.getElementById('node-info-latest');
    const statusEl = document.getElementById('node-info-status');

    let installedVersion = '';
    if (localRes.status === 'fulfilled' && localRes.value.installed) {
        installedVersion = localRes.value.version;
        versionEl.textContent = installedVersion;
        pathEl.textContent = localRes.value.path;
    } else {
        versionEl.textContent = 'Node.js não encontrado';
        versionEl.classList.add('text-danger');
        pathEl.textContent = '-';
    }

    if (latestRes.status === 'fulfilled' && latestRes.value.latest) {
        const latest = latestRes.value.latest;
        latestEl.innerHTML = `<a href="https://nodejs.org/en/download/" target="_blank">${latest}</a>`;

        if (installedVersion) {
            const cmp = compareVersions(installedVersion, latest);
            if (cmp >= 0) {
                statusEl.textContent = 'Atualizado';
                statusEl.className = 'badge ms-2 text-bg-success';
            } else {
                statusEl.textContent = 'Nova versão disponível';
                statusEl.className = 'badge ms-2 text-bg-warning';
            }
        }
    } else {
        latestEl.textContent = 'Não foi possível verificar';
        latestEl.classList.add('text-body-secondary');
    }
}

let _npmOutdated = false;

async function carregarInfoNpm() {
    const loading = document.getElementById('npm-version-loading');
    const content = document.getElementById('npm-version-content');
    if (!content) return;

    loading.classList.remove('d-none');
    content.classList.add('d-none');

    try {
        const res = await fetch('/api/check-npm');
        const data = await res.json();

        loading.classList.add('d-none');
        content.classList.remove('d-none');

        const versionEl = document.getElementById('npm-info-version');
        const latestEl = document.getElementById('npm-info-latest');
        const statusEl = document.getElementById('npm-info-status');

        if (data.error) {
            versionEl.textContent = 'Não foi possível verificar';
            versionEl.classList.add('text-body-secondary');
            return;
        }

        versionEl.textContent = data.current || '-';
        latestEl.textContent = data.latest || '-';

        if (data.current && data.latest) {
            const cmp = compareVersions(data.current, data.latest);
            if (cmp >= 0) {
                statusEl.textContent = 'Atualizado';
                statusEl.className = 'badge ms-2 text-bg-success';
            } else {
                statusEl.textContent = 'Nova versão disponível';
                statusEl.className = 'badge ms-2 text-bg-warning';
                _npmOutdated = true;
                document.getElementById('btn-run-npm_update')?.classList.remove('d-none');
            }
        }
    } catch {
        loading.classList.add('d-none');
        content.classList.remove('d-none');
        document.getElementById('npm-info-version').textContent = 'Erro ao verificar';
    }
}

async function carregarNpmPackages() {
    const loading = document.getElementById('npm-packages-loading');
    const content = document.getElementById('npm-packages-content');
    const statusEl = document.getElementById('npm-packages-status');
    const listEl = document.getElementById('npm-packages-list');
    const updateBtn = document.getElementById('btn-run-npm_update');

    if (!content) return;

    loading.classList.remove('d-none');
    content.classList.add('d-none');

    try {
        const res = await fetch('/api/npm-outdated');
        const data = await res.json();

        loading.classList.add('d-none');
        content.classList.remove('d-none');

        if (data.error) {
            statusEl.innerHTML = `<span class="text-body-secondary">${data.error}</span>`;
            return;
        }

        if (data.packages.length === 0) {
            statusEl.innerHTML = '<span class="badge text-bg-success">Todos os pacotes estão atualizados</span>';
            listEl.innerHTML = '';
            return;
        }

        statusEl.innerHTML = `<span class="badge text-bg-warning">${data.packages.length} pacote(s) com atualização disponível</span>`;

        const rows = data.packages.map(pkg =>
            `<tr>
                <td>${pkg.name}</td>
                <td>${pkg.current}</td>
                <td>${pkg.wanted}</td>
                <td>${pkg.latest}</td>
            </tr>`
        ).join('');

        listEl.innerHTML =
            `<div class="files-table-container">
                <table class="files-table">
                    <thead><tr>
                        <th>Pacote</th>
                        <th>Atual</th>
                        <th>Compatível</th>
                        <th>Última</th>
                    </tr></thead>
                    <tbody>${rows}</tbody>
                </table>
            </div>`;

        updateBtn.classList.remove('d-none');
    } catch {
        loading.classList.add('d-none');
        content.classList.remove('d-none');
        statusEl.innerHTML = '<span class="text-body-secondary">Não foi possível verificar os pacotes.</span>';
    }
}

document.getElementById('r-info-loading').innerHTML = customSpinnerHTML('Verificando R...');
document.getElementById('node-info-loading').innerHTML = customSpinnerHTML('Verificando Node.js...');
document.getElementById('npm-version-loading').innerHTML = customSpinnerHTML('Verificando npm...');
document.getElementById('npm-packages-loading').innerHTML = customSpinnerHTML('Verificando pacotes...');

carregarInfoR();
carregarInfoNode();

fetch('/api/app-config')
    .then(r => r.json())
    .then(cfg => {
        if (cfg.isDev) {
            document.getElementById('npm-packages-panel')?.classList.remove('d-none');
            carregarInfoNpm();
            carregarNpmPackages();
        }
    })
    .catch(() => { /* sem bloco npm em caso de falha */ });
