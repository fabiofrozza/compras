let pregaoSelecionado = null;
let pregaoFolderFileCount = 0; // qtd de arquivos na pasta do pregão selecionado
let pregoesDadosFolderPath = '';
let importarFolderPath = '';

// ====== INICIALIZAÇÃO ======

function inicializarFornecedores() {
    carregarPregoes();
    carregarImportar();

    const btnObterDados = document.getElementById('btn-obter-dados-fornecedores');

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
            const btn = e.target.closest('.item-card-open');
            if (!btn) return;
            e.stopPropagation();
            removeTooltip(btn);
            const pregao = btn.closest('.item-card')?.dataset.pregao;
            if (pregao && pregoesDadosFolderPath) {
                openFolder(pregoesDadosFolderPath + '\\' + pregao, pregao);
            }
        });
    }
}

// ====== PREGÕES ======

async function carregarPregoes() {
    const container = document.getElementById('fornecedores-pregoes-list');
    if (!container) return;

    container.innerHTML = customSpinnerHTML('Carregando pregões...');

    try {
        const response = await fetch('/api/fornecedores/pregoes');
        const data = await response.json();

        if (data.error) {
            container.innerHTML = `<div class="alert alert-danger"><i class="material-symbols-outlined">error</i> ${data.error}</div>`;
            return;
        }

        pregoesDadosFolderPath = data.folderPath;
        const displayPath = data.folderPath.replace(/\\/g, '/');

        const refreshPregoes = `
            <button data-bs-toggle="tooltip" data-bs-title="Atualizar lista de pregões" class="folder-path-btn btn-refresh-pregoes">
                <i class="material-symbols-outlined">refresh</i>
            </button>`;

        if (!data.pregoes || data.pregoes.length === 0) {
            const emptyMsg = data.folderCreated
                ? '<i class="material-symbols-outlined">create_new_folder</i> Pasta não encontrada e criada'
                : '<i class="material-symbols-outlined">warning</i> Nenhum pregão encontrado na pasta DADOS';
            container.innerHTML = buildFolderPathHTML(displayPath, '', refreshPregoes) +
                `<div class="alert alert-warning">${emptyMsg}</div>`;
            setupFolderPathButtons(container);
            setupRefreshPregoesButton(container);
            initializeTooltips();
            return;
        }

        let html = buildFolderPathHTML(displayPath, '', refreshPregoes) + '<div class="items-grid">';
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
            const isSelected = pregaoSelecionado === pregao.nome ? ' item-selected selected' : '';

            html += `
                <div class="item-card ${statusClass}${isSelected}" data-pregao="${pregao.nome}"
                     onclick="selecionarPregao('${pregao.nome}')">
                    <div class="item-card-icon">
                        <i class="material-symbols-outlined">folder</i>
                    </div>
                    <div class="item-card-name">${pregao.nome}</div>
                    <div class="item-card-status">
                        <i class="material-symbols-outlined"
                        data-bs-toggle="tooltip" data-bs-title="${statusTooltip}">
                            ${statusIcon}
                        </i>
                    </div>
                    <button class="item-card-open btn btn-sm text-primary file-row-btn" data-bs-toggle="tooltip"
                        data-bs-title="Abrir pasta ${pregao.nome}"
                        data-bs-placement="right">
                        <i class="material-symbols-outlined">folder_open</i>
                    </button>
                    <button class="item-card-delete btn btn-sm text-danger file-row-btn" data-bs-toggle="tooltip"
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
        setupRefreshPregoesButton(container);
        initializeTooltips();
    } catch (error) {
        container.innerHTML = `<div class="alert alert-danger"><i class="material-symbols-outlined">error</i> Erro ao carregar pregões: ${error.message}</div>`;
    }
}

function setupRefreshPregoesButton(container) {
    const btn = container.querySelector('.btn-refresh-pregoes');
    if (btn) {
        btn.addEventListener('click', () => {
            removeTooltip(btn);
            carregarPregoes();
        });
    }
}

function selecionarPregao(nome) {
    pregaoSelecionado = nome;
    pregaoFolderFileCount = 0; // será atualizado por atualizarInfoPregao

    // Atualizar visual de seleção
    document.querySelectorAll('.item-card[data-pregao]').forEach(card => {
        const isSelected = card.dataset.pregao === nome;
        card.classList.toggle('item-selected', isSelected);
        card.classList.toggle('selected', isSelected);
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

    container.innerHTML = customSpinnerHTML('Carregando arquivos...');

    try {
        const response = await fetch(`/api/fornecedores/pregao/${pregao}/arquivos`);
        const data = await response.json();

        if (data.error) {
            container.innerHTML = `<div class="alert alert-danger"><i class="material-symbols-outlined">error</i> ${data.error}</div>`;
            return;
        }

        const displayPath = data.folderPath.replace(/\\/g, '/');
        let html = buildFolderPathHTML(displayPath);

        if (!data.arquivos || data.arquivos.length === 0) {
            html += `<div class="alert alert-warning"><i class="material-symbols-outlined">warning</i> Nenhum arquivo encontrado</div>`;
            container.innerHTML = html;
            setupFolderPathButtons(container);
            setupConteudoFooterButtons(container, pregao, false);
            initializeTooltips();
            return;
        }
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
            // Don´t show CONFERENCIA file
            if (file.name === 'PE_' + pregao + '_CONFERENCIA.xlsx') return;

            const notProcessed = !data.resultado || data.resultado === 'sem_saida';
            const iconColor = file.hasError ? 'text-danger' : '';
            const erroText = file.hasError ? (file.errorType || 'Sim') : (notProcessed ? '-' : 'Não');
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
    const folderPathContainer = container.querySelector('.folder-path-container');
    const copyIcon = folderPathContainer.querySelector('.copy-icon');

    const btnClear = document.createElement('button');
    btnClear.className = 'folder-path-btn text-danger btn-clear-conteudo';
    btnClear.setAttribute('data-bs-toggle', 'tooltip');
    btnClear.setAttribute('data-bs-title', 'Excluir todos os arquivos');
    btnClear.innerHTML = '<i class="material-symbols-outlined">delete</i>';
    if (!hasFiles) btnClear.disabled = true;
    copyIcon.parentNode.insertBefore(btnClear, copyIcon);

    const btnRefresh = document.createElement('button');
    btnRefresh.className = 'folder-path-btn btn-refresh-conteudo';
    btnRefresh.setAttribute('data-bs-toggle', 'tooltip');
    btnRefresh.setAttribute('data-bs-title', 'Atualizar lista');
    btnRefresh.innerHTML = '<i class="material-symbols-outlined">refresh</i>';
    const btnOpen = folderPathContainer.querySelector('.btn-open');
    btnOpen.parentNode.insertBefore(btnRefresh, btnOpen);

    btnClear.addEventListener('click', () => {
        removeTooltip(btnClear);
        excluirPregao(pregao);
    });

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

        pregaoFolderFileCount = pregao.qtdArquivos;
        if (typeof evaluateAllButtons === 'function') evaluateAllButtons();

        if (pregao.qtdArquivos === 0) {
            elQtd.closest('.dashboard-widget').classList.add('border-danger');
        } else {
            elQtd.closest('.dashboard-widget').classList.remove('border-danger');
        }


        if (elGerado) {
            const gerado = pregao.resultado === 'sucesso' || pregao.resultado === 'parcial';
            if (gerado) {
                const badgeClass = pregao.resultado === 'sucesso' ? 'success' : 'warning';
                const badgeIcon = pregao.resultado === 'sucesso' ? 'check' : 'warning';
                elGerado.innerHTML = `<span class="badge rounded-pill ${badgeClass}"><i class="material-symbols-outlined">${badgeIcon}</i> Sim</span>`;
            } else {
                elGerado.innerHTML = `<span class="badge rounded-pill blank">Não</span>`;
            }
            elGerado.className = 'widget-value';
        }
        if (elErros) {
            if (!pregao.resultado) {
                elErros.textContent = '-';
                elErros.className = 'widget-value';
            } else if (pregao.resultado === 'sucesso') {
                elErros.innerHTML = `<span class="badge rounded-pill success"><i class="material-symbols-outlined">check</i> Não</span>`;
                elErros.className = 'widget-value';
            } else {
                elErros.innerHTML = `<span class="badge rounded-pill danger"><i class="material-symbols-outlined">close</i> Sim</span>`;
                elErros.className = 'widget-value';
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
            el.textContent = '-';
            el.className = 'widget-value';
        }
    });

    pregaoFolderFileCount = 0;
    if (typeof evaluateAllButtons === 'function') evaluateAllButtons();
}

// ====== ARQUIVOS A IMPORTAR ======

function setupImportarFooterButtons(container, hasFiles) {
    const folderPathContainer = container.querySelector('.folder-path-container');
    const copyIcon = folderPathContainer.querySelector('.copy-icon');

    const btnClear = document.createElement('button');
    btnClear.className = 'folder-path-btn text-danger btn-clear-importar';
    btnClear.setAttribute('data-bs-toggle', 'tooltip');
    btnClear.setAttribute('data-bs-title', 'Excluir todos os arquivos');
    btnClear.innerHTML = '<i class="material-symbols-outlined">delete</i>';
    if (!hasFiles) btnClear.disabled = true;
    copyIcon.parentNode.insertBefore(btnClear, copyIcon);

    const btnRefresh = document.createElement('button');
    btnRefresh.className = 'folder-path-btn btn-refresh-importar';
    btnRefresh.setAttribute('data-bs-toggle', 'tooltip');
    btnRefresh.setAttribute('data-bs-title', 'Atualizar lista');
    btnRefresh.innerHTML = '<i class="material-symbols-outlined">refresh</i>';
    const btnOpen = folderPathContainer.querySelector('.btn-open');
    btnOpen.parentNode.insertBefore(btnRefresh, btnOpen);

    btnClear.addEventListener('click', () => {
        removeTooltip(btnClear);
        limparPastaImportar();
    });

    btnRefresh.addEventListener('click', () => {
        removeTooltip(btnRefresh);
        carregarImportar();
    });

    initializeTooltips();
}

async function carregarImportar() {
    const container = document.getElementById('fornecedores-importar-list');
    if (!container) return;

    container.innerHTML = customSpinnerHTML('Carregando arquivos...');

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
            const emptyMsg = data.folderCreated
                ? '<i class="material-symbols-outlined">create_new_folder</i> Pasta não encontrada e criada'
                : '<i class="material-symbols-outlined">warning</i> Nenhum arquivo para importar';
            let html = buildFolderPathHTML(data.folderPath.replace(/\\/g, '/'));
            html += `<div class="alert alert-warning">${emptyMsg}</div>`;
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
            const erroText = file.hasError === null ? '-' : (file.hasError ? 'Sim' : 'Não');
            const erroClass = file.hasError === null ? '' : (file.hasError ? 'text-danger fw-bold' : 'text-success');
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
    const pregaoCard = document.querySelector(`.item-card[data-pregao="${pregao}"]`);

    if (pregaoCard) {
        selecionarPregao(pregao);
    } else {
        // Pasta do pregão não existe - limpar seleção e atualizar listas
        pregaoSelecionado = null;
        pregaoFolderFileCount = 0;
        document.querySelectorAll('.item-card[data-pregao]').forEach(card => {
            card.classList.remove('item-selected', 'selected');
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
        const response = await fetch('/api/clear-folder/fornecedores/para_importar?extensions=.xlsx,.csv,.json', {
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
        message: `Mover o arquivo para importação dos dados dos fornecedores do pregão <strong>${pregao}</strong> para a pasta Documentos?`,
        detail: '<i class="material-symbols-outlined me-1">folder</i> Os arquivos auxiliares (conferência e status) continuarão na pasta do pregão e poderão ser excluídos posteriormente com segurança.',
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
