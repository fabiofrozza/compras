// ===========================================
// fornecedores.js — Lógica da aba Fornecedores
// ===========================================

let pregaoSelecionado = null;
let pregoesDadosFolderPath = '';
let importarFolderPath = '';

// ====== INICIALIZAÇÃO ======

function inicializarFornecedores() {
    carregarPregoes();
    carregarImportar();

    const btnRefreshPregoes = document.getElementById('btn-refresh-pregoes');
    const btnObterDados = document.getElementById('btn-obter-dados-fornecedores');

    if (btnRefreshPregoes) {
        btnRefreshPregoes.addEventListener('click', () => {
            removeTooltip(btnRefreshPregoes);
            carregarPregoes();
        });
    }

    if (btnObterDados) {
        btnObterDados.addEventListener('click', (e) => {
            e.stopPropagation();
            if (!pregaoSelecionado) {
                showToast('Selecione um pregão primeiro.', 'warning');
                return;
            }
            executarFornecedores(pregaoSelecionado);
        });
    }

    const pregoesContainer = document.getElementById('fornecedores-pregoes-list');
    if (pregoesContainer) {
        pregoesContainer.addEventListener('click', (e) => {
            const btn = e.target.closest('.pregao-card-open');
            if (!btn) return;
            e.stopPropagation();
            removeTooltip(btn);
            const pregao = btn.closest('.pregao-card')?.dataset.pregao;
            if (pregao && pregoesDadosFolderPath) {
                openFolder(pregoesDadosFolderPath + '\\' + pregao);
                showToast(`Abrindo pasta ${pregao}. Verifique na barra de tarefas...`, 'info', 5000);
            }
        });
    }
}

// ====== PREGÕES ======

async function carregarPregoes() {
    const container = document.getElementById('fornecedores-pregoes-list');
    if (!container) return;

    container.innerHTML = `
        <div class="custom-spinner-container">
            <div class="custom-spinner text-primary">
                <div class="spinner-border"></div>
                <span role="status">Carregando pregões...</span>
            </div>
        </div>
    `;

    try {
        const response = await fetch('/api/fornecedores/pregoes');
        const data = await response.json();

        if (data.error) {
            container.innerHTML = `<div class="alert alert-danger"><i class="material-symbols-outlined">error</i> ${data.error}</div>`;
            return;
        }

        pregoesDadosFolderPath = data.folderPath;
        const displayPath = data.folderPath.replace(/\\/g, '/');

        if (!data.pregoes || data.pregoes.length === 0) {
            container.innerHTML = buildFolderPathHTML(displayPath) +
                `<div class="alert alert-warning"><i class="material-symbols-outlined">warning</i> Nenhum pregão encontrado na pasta DADOS</div>`;
            setupFolderPathButtons(container);
            initializeTooltips();
            return;
        }

        let html = buildFolderPathHTML(displayPath) + '<div class="fornecedores-pregoes-grid">';
        data.pregoes.forEach(pregao => {
                const statusClass = {
                sucesso: 'status-sucesso',
                parcial: 'status-parcial',
                sem_saida: 'status-sem_saida'
            }[pregao.status] || 'status-pendente';
            const statusIcon = {
                sucesso: 'check_circle',
                parcial: 'warning',
                sem_saida: 'cancel'
            }[pregao.status] || 'hourglass_empty';
            const statusTooltip = {
                sucesso: 'Arquivo gerado com sucesso',
                parcial: 'Arquivo gerado com erros',
                sem_saida: 'Script executado sem gerar arquivo',
            }[pregao.status] || 'Pendente de geração do arquivo';
            const isSelected = pregaoSelecionado === pregao.nome ? ' pregao-selected' : '';

            html += `
                <div class="pregao-card ${statusClass}${isSelected}" data-pregao="${pregao.nome}"
                     onclick="selecionarPregao('${pregao.nome}')">
                    <div class="pregao-card-icon">
                        <i class="material-symbols-outlined">folder</i>
                    </div>
                    <div class="pregao-card-name">${pregao.nome}</div>
                    <div class="pregao-card-status">
                        <i class="material-symbols-outlined"
                        data-bs-toggle="tooltip" data-bs-title="${statusTooltip}">
                            ${statusIcon}
                        </i>
                    </div>
                    <button class="pregao-card-open btn btn-sm" data-bs-toggle="tooltip"
                        data-bs-title="Abrir pasta ${pregao.nome}"
                        data-bs-placement="right">
                        <i class="material-symbols-outlined">folder_open</i>
                    </button>
                    <button class="pregao-card-delete btn btn-sm" data-bs-toggle="tooltip"
                        data-bs-title="Excluir pregão ${pregao.nome}"
                        data-bs-placement="right"
                        onclick="event.stopPropagation(); excluirPregao('${pregao.nome}')">
                        <i class="material-symbols-outlined">delete</i>
                    </button>
                </div>
            `;
        });
        html += '</div>';

        container.innerHTML = html;
        setupFolderPathButtons(container);
        initializeTooltips();
    } catch (error) {
        container.innerHTML = `<div class="alert alert-danger"><i class="material-symbols-outlined">error</i> Erro ao carregar pregões: ${error.message}</div>`;
    }
}

function selecionarPregao(nome) {
    pregaoSelecionado = nome;

    // Atualizar visual de seleção
    document.querySelectorAll('.pregao-card').forEach(card => {
        card.classList.toggle('pregao-selected', card.dataset.pregao === nome);
    });

    carregarConteudoPasta(nome);
    atualizarInfoPregao(nome);
}

async function excluirPregao(nome) {
    const confirmed = await showConfirmationModal({
        title: 'Excluir Pregão',
        message: `Tem certeza que deseja excluir a pasta do pregão <strong>${nome}</strong>?`,
        detail: '<i class="material-symbols-outlined me-1">warning</i> Esta ação não pode ser desfeita.',
        confirmText: 'Excluir',
        confirmColor: 'btn-danger'
    });

    if (!confirmed) return;

    try {
        const response = await fetch(`/api/fornecedores/pregao/${nome}`, { method: 'DELETE' });
        const data = await response.json();

        if (!response.ok) {
            showToast(`Erro ao excluir: ${data.error}`, 'error');
            return;
        }

        showToast(data.message, 'success');

        if (pregaoSelecionado === nome) {
            pregaoSelecionado = null;
            limparInfoPregao();
            limparConteudoPasta();
        }

        carregarPregoes();
    } catch (error) {
        showToast(`Erro ao excluir pregão: ${error.message}`, 'error');
    }
}

// ====== CONTEÚDO DA PASTA ======

async function carregarConteudoPasta(pregao) {
    const container = document.getElementById('fornecedores-conteudo-list');
    const titulo = document.getElementById('conteudo-pasta-titulo');
    if (!container) return;

    if (titulo) titulo.textContent = `Fornecedores - Pregão ${pregao}`;

    container.innerHTML = `
        <div class="custom-spinner-container">
            <div class="custom-spinner text-primary">
                <div class="spinner-border"></div>
                <span role="status">Carregando arquivos...</span>
            </div>
        </div>
    `;

    try {
        const response = await fetch(`/api/fornecedores/pregao/${pregao}/arquivos`);
        const data = await response.json();

        if (data.error) {
            container.innerHTML = `<div class="alert alert-danger"><i class="material-symbols-outlined">error</i> ${data.error}</div>`;
            return;
        }

        if (!data.arquivos || data.arquivos.length === 0) {
            container.innerHTML = `<div class="alert alert-warning"><i class="material-symbols-outlined">warning</i> Nenhum arquivo encontrado</div>`;
            return;
        }

        const displayPath = data.folderPath.replace(/\\/g, '/');
        let html = buildFolderPathHTML(displayPath);
        html += `
            <div class="files-table-container">
                <table class="files-table">
                    <thead>
                        <tr>
                            <th colspan="2">Dados dos Fornecedores</th>
                            <th>Erros</th>
                            <th></th>
                        </tr>
                    </thead>
                    <tbody>
        `;

        data.arquivos.forEach(file => {
            const notProcessed = !data.resultado || data.resultado === 'sem_saida';
            const iconColor = file.hasError ? 'text-danger' : '';
            const erroText = file.hasError ? (file.errorType || 'Sim') : (notProcessed ? '—' : 'Não');
            const erroClass = file.hasError ? 'text-danger fw-bold' : (notProcessed ? '' : 'text-success');
            const safeFilePath = file.fullPath.replace(/\\/g, '\\\\').replace(/"/g, '&quot;');

            html += `
                <tr ondblclick="openFile('${safeFilePath}')">
                    <td><span class="file-icon"><i class="material-symbols-outlined ${iconColor}">table_chart</i></span></td>
                    <td>${file.name}</td>
                    <td class="${erroClass}">${erroText}</td>
                    <td class="table-btn-column text-nowrap">
                        <button class="btn btn-sm text-primary file-row-btn" onclick="event.stopPropagation(); openFile('${safeFilePath}')"
                            data-bs-toggle="tooltip" data-bs-title="Abrir arquivo">
                            <i class="material-symbols-outlined">open_in_new</i>
                        </button>
                        <button class="btn btn-sm text-danger file-row-btn" onclick="event.stopPropagation(); excluirArquivo('${safeFilePath}', '${pregao}')"
                            data-bs-toggle="tooltip" data-bs-title="Excluir arquivo">
                            <i class="material-symbols-outlined">delete</i>
                        </button>
                    </td>
                </tr>
            `;
        });

        html += '</tbody></table></div>';
        container.innerHTML = html;
        setupFolderPathButtons(container);
        setupConteudoFooterButtons(container, pregao, data.arquivos.length > 0);
        initializeTooltips();
    } catch (error) {
        container.innerHTML = `<div class="alert alert-danger"><i class="material-symbols-outlined">error</i> Erro: ${error.message}</div>`;
    }
}

function setupConteudoFooterButtons(container, pregao, hasFiles) {
    const footer = document.createElement('div');
    footer.className = 'folder-buttons-container';
    footer.innerHTML = `
        <button data-bs-toggle="tooltip" data-bs-title="Excluir todos os arquivos" class="btn btn-outline-danger btn-sm btn-clear-conteudo" ${hasFiles ? '' : 'disabled'}>
            <i class="material-symbols-outlined">delete</i>
        </button>
        <button data-bs-toggle="tooltip" data-bs-title="Atualizar lista" class="btn btn-outline-primary btn-sm btn-refresh-conteudo">
            <i class="material-symbols-outlined">refresh</i>
        </button>
    `;
    container.appendChild(footer);

    const btnClear = footer.querySelector('.btn-clear-conteudo');
    btnClear.addEventListener('click', () => {
        removeTooltip(btnClear);
        excluirPregao(pregao);
    });

    const btnRefresh = footer.querySelector('.btn-refresh-conteudo');
    btnRefresh.addEventListener('click', () => {
        removeTooltip(btnRefresh);
        carregarConteudoPasta(pregao);
    });

    initializeTooltips();
}

function limparConteudoPasta() {
    const container = document.getElementById('fornecedores-conteudo-list');
    const titulo = document.getElementById('conteudo-pasta-titulo');
    if (container) {
        container.innerHTML = `<div class="alert alert-info"><i class="material-symbols-outlined">info</i> Selecione um pregão para ver os arquivos</div>`;
    }
    if (titulo) titulo.textContent = 'Conteúdo da pasta';
}

async function excluirArquivo(filePath, pregao) {
    const fileName = filePath.split(/[/\\]/).pop();
    const confirmed = await showConfirmationModal({
        title: 'Excluir Arquivo',
        message: `Tem certeza que deseja excluir o arquivo <strong>${fileName}</strong>?`,
        confirmText: 'Excluir',
        confirmColor: 'btn-danger'
    });

    if (!confirmed) return;

    try {
        const response = await fetch('/api/fornecedores/arquivo', {
            method: 'DELETE',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ filePath })
        });
        const data = await response.json();

        if (!response.ok) {
            showToast(`Erro ao excluir: ${data.error}`, 'error');
            return;
        }

        showToast(data.message, 'success');
        if (pregao) {
            carregarConteudoPasta(pregao);
            carregarPregoes();
        }
    } catch (error) {
        showToast(`Erro ao excluir arquivo: ${error.message}`, 'error');
    }
}

// ====== INFORMAÇÕES DO PREGÃO ======

async function atualizarInfoPregao(nome) {
    const elNumero = document.getElementById('info-pregao-numero');
    const elQtd = document.getElementById('info-pregao-qtd');
    const elGerado = document.getElementById('info-pregao-gerado');
    const elErros = document.getElementById('info-pregao-erros');

    if (elNumero) elNumero.textContent = nome;

    try {
        const response = await fetch('/api/fornecedores/pregoes');
        const data = await response.json();

        const pregao = data.pregoes?.find(p => p.nome === nome);
        if (!pregao) return;

        if (elQtd) elQtd.textContent = pregao.qtdArquivos;

        const btn = document.getElementById('btn-obter-dados-fornecedores');
        if (btn) btn.disabled = pregao.qtdArquivos === 0;

        if (elGerado) {
            const gerado = pregao.resultado === 'sucesso' || pregao.resultado === 'parcial';
            elGerado.textContent = gerado ? 'Sim' : 'Não';
            elGerado.className = pregao.resultado === 'sucesso' ? 'text-success fw-bold' :
                pregao.resultado === 'parcial' ? 'text-warning fw-bold' : '';
        }
        if (elErros) {
            if (!pregao.resultado) {
                elErros.textContent = '—';
                elErros.className = '';
            } else if (pregao.resultado === 'sucesso') {
                elErros.textContent = 'Não';
                elErros.className = 'text-success';
            } else {
                elErros.textContent = 'Sim';
                elErros.className = 'text-danger fw-bold';
            }
        }
    } catch (error) {
        console.error('Erro ao atualizar info do pregão:', error);
    }
}

function limparInfoPregao() {
    const fields = ['info-pregao-numero', 'info-pregao-qtd', 'info-pregao-gerado', 'info-pregao-erros'];
    fields.forEach(id => {
        const el = document.getElementById(id);
        if (el) {
            el.textContent = '—';
            el.className = '';
        }
    });

    const btn = document.getElementById('btn-obter-dados-fornecedores');
    if (btn) btn.disabled = true;
}

// ====== ARQUIVOS A IMPORTAR ======

function setupImportarFooterButtons(container, hasFiles) {
    const footer = document.createElement('div');
    footer.className = 'folder-buttons-container';
    footer.innerHTML = `
        <button data-bs-toggle="tooltip" data-bs-title="Excluir todos os arquivos" class="btn btn-outline-danger btn-sm btn-clear-importar" ${hasFiles ? '' : 'disabled'}>
            <i class="material-symbols-outlined">delete</i>
        </button>
        <button data-bs-toggle="tooltip" data-bs-title="Atualizar lista" class="btn btn-outline-primary btn-sm btn-refresh-importar">
            <i class="material-symbols-outlined">refresh</i>
        </button>
    `;
    container.appendChild(footer);

    const btnClear = footer.querySelector('.btn-clear-importar');
    btnClear.addEventListener('click', () => {
        removeTooltip(btnClear);
        limparPastaImportar();
    });

    const btnRefresh = footer.querySelector('.btn-refresh-importar');
    btnRefresh.addEventListener('click', () => {
        removeTooltip(btnRefresh);
        carregarImportar();
    });

    initializeTooltips();
}

async function carregarImportar() {
    const container = document.getElementById('fornecedores-importar-list');
    if (!container) return;

    container.innerHTML = `
        <div class="custom-spinner-container">
            <div class="custom-spinner text-primary">
                <div class="spinner-border"></div>
                <span role="status">Carregando arquivos...</span>
            </div>
        </div>
    `;

    try {
        const response = await fetch('/api/fornecedores/importar');
        const data = await response.json();

        if (data.error) {
            container.innerHTML = `<div class="alert alert-danger"><i class="material-symbols-outlined">error</i> ${data.error}</div>`;
            setupImportarFooterButtons(container, false);
            return;
        }

        importarFolderPath = data.folderPath;

        if (!data.arquivos || data.arquivos.length === 0) {
            let html = buildFolderPathHTML(data.folderPath.replace(/\\/g, '/'));
            html += `<div class="alert alert-warning"><i class="material-symbols-outlined">warning</i> Nenhum arquivo para importar</div>`;
            container.innerHTML = html;
            setupFolderPathButtons(container);
            setupImportarFooterButtons(container, false);
            initializeTooltips();
            return;
        }

        const displayPath = data.folderPath.replace(/\\/g, '/');
        let html = buildFolderPathHTML(displayPath);
        html += `
            <div class="files-table-container">
                <table class="files-table">
                    <thead>
                        <tr>
                            <th colspan="2">Importar</th>
                            <th>Erros</th>
                            <th></th>
                        </tr>
                    </thead>
                    <tbody>
        `;

        data.arquivos.forEach(file => {
            const iconColor = file.hasError ? 'text-danger' : '';
            const erroText = file.hasError ? 'Sim' : 'Não';
            const erroClass = file.hasError ? 'text-danger fw-bold' : 'text-success';
            const safeFilePath = file.fullPath.replace(/\\/g, '\\\\').replace(/"/g, '&quot;');
            const safeConfPath = file.conferenciaPath ? file.conferenciaPath.replace(/\\/g, '\\\\').replace(/"/g, '&quot;') : '';

            html += `
                <tr class="importar-row" data-pregao="${file.pregao}"
                    onclick="selecionarImportar('${file.pregao}')">
                    <td><span class="file-icon"><i class="material-symbols-outlined ${iconColor}">table_chart</i></span></td>
                    <td>${file.name}</td>
                    <td class="${erroClass}">${erroText}</td>
                    <td class="table-btn-column text-nowrap">
                        ${file.hasConferencia ? `
                        <button class="btn btn-sm text-primary file-row-btn" onclick="event.stopPropagation(); openFile('${safeConfPath}')"
                            data-bs-toggle="tooltip" data-bs-title="Abrir arquivo de conferência">
                            <i class="material-symbols-outlined">open_in_new</i>
                        </button>` : ''}
                        <button class="btn btn-sm text-success file-row-btn" onclick="event.stopPropagation(); moverArquivos('${file.pregao}')"
                            data-bs-toggle="tooltip" data-bs-title="Mover para Documentos">
                            <i class="material-symbols-outlined">folder</i>
                        </button>
                        <button class="btn btn-sm text-danger file-row-btn" onclick="event.stopPropagation(); excluirArquivoImportar('${file.pregao}')"
                            data-bs-toggle="tooltip" data-bs-title="Excluir arquivos do pregão">
                            <i class="material-symbols-outlined">delete</i>
                        </button>
                    </td>
                </tr>
            `;
        });

        html += '</tbody></table></div>';
        container.innerHTML = html;
        setupFolderPathButtons(container);
        setupImportarFooterButtons(container, true);
        initializeTooltips();
    } catch (error) {
        container.innerHTML = `<div class="alert alert-danger"><i class="material-symbols-outlined">error</i> Erro: ${error.message}</div>`;
    }
}

async function selecionarImportar(pregao) {
    const pregaoCard = document.querySelector(`.pregao-card[data-pregao="${pregao}"]`);

    if (pregaoCard) {
        selecionarPregao(pregao);
    } else {
        // Pasta do pregão não existe — limpar seleção e atualizar listas
        pregaoSelecionado = null;
        document.querySelectorAll('.pregao-card').forEach(card => {
            card.classList.remove('pregao-selected');
        });
        limparInfoPregao();
        limparConteudoPasta();
        await carregarImportar();
        showToast(`A pasta do pregão ${pregao} não existe mais.`, 'warning');
    }
}

async function excluirArquivoImportar(pregao) {
    const confirmed = await showConfirmationModal({
        title: 'Excluir Arquivos',
        message: `Tem certeza que deseja excluir todos os arquivos do pregão <strong>${pregao}</strong>?`,
        detail: '<i class="material-symbols-outlined me-1">info</i> Serão excluídos o CSV e o arquivo de conferência (se existirem).',
        confirmText: 'Excluir',
        confirmColor: 'btn-danger'
    });

    if (!confirmed) return;

    const suffixes = ['.csv', '_CONFERENCIA.xlsx'];
    let deleted = 0;

    for (const suffix of suffixes) {
        try {
            const fullPath = importarFolderPath + '\\PE_' + pregao + suffix;
            const response = await fetch('/api/fornecedores/arquivo', {
                method: 'DELETE',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ filePath: fullPath })
            });
            if (response.ok) deleted++;
        } catch { /* arquivo pode não existir */ }
    }

    showToast(`${deleted} arquivo(s) do pregão ${pregao} excluído(s).`, 'success');
    carregarImportar();
    carregarPregoes();
}

async function limparPastaImportar() {
    const confirmed = await showConfirmationModal({
        title: 'Confirmar Exclusão',
        message: 'Tem certeza que deseja excluir <strong>TODOS</strong> os arquivos da pasta?',
        detail: `<i class="material-symbols-outlined me-1">folder</i> ${importarFolderPath}`,
        confirmText: 'Excluir',
        confirmColor: 'btn-danger'
    });

    if (!confirmed) return;

    try {
        const response = await fetch('/api/clear-folder/fornecedores/PARA_IMPORTAR?extensions=.xlsx,.csv,.json', {
            method: 'DELETE'
        });
        const data = await response.json();

        if (!response.ok) {
            showToast(`Erro ao limpar pasta: ${data.error}`, 'error', 10000);
            return;
        }

        showToast(data.message, 'success');
        carregarImportar();
        carregarPregoes();
    } catch (error) {
        showToast(`Erro ao limpar pasta: ${error.message}`, 'error', 10000);
    }
}

async function moverArquivos(pregao) {
    const confirmed = await showConfirmationModal({
        title: 'Mover Arquivos',
        message: `Mover os arquivos do pregão <strong>${pregao}</strong> para a pasta Documentos?`,
        detail: '<i class="material-symbols-outlined me-1">folder</i> Os arquivos CSV e XLSX de conferência serão movidos.',
        confirmText: 'Mover',
        confirmColor: 'btn-success'
    });

    if (!confirmed) return;

    try {
        const response = await fetch('/api/fornecedores/mover', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ pregao })
        });
        const data = await response.json();

        if (!response.ok) {
            showToast(`Erro ao mover: ${data.error}`, 'error');
            return;
        }

        showToast(`${data.message}. Destino: ${data.destination}`, 'success', 8000);
        carregarImportar();
        carregarPregoes();
    } catch (error) {
        showToast(`Erro ao mover arquivos: ${error.message}`, 'error');
    }
}

// ====== EXECUÇÃO DO SCRIPT R ======

function executarFornecedores(pregao) {
    if (typeof runRScript !== 'function') {
        showToast('Erro: função de execução não disponível.', 'error');
        return;
    }

    runRScript('fornecedores', { pregao });
}

// Quando o script finalizar, atualizar as listas
document.addEventListener('DOMContentLoaded', () => {
    // Após o script de fornecedores finalizar, recarregar as listas
    const originalHandleScriptResult = window.handleScriptResult;
});
