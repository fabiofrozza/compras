let _rDownloadUrl = '';

function executarInstalacao(modo) {
    runRScript('instalacao', { modo });
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
        pathEl.textContent = '—';
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

carregarInfoR();
