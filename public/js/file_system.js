const FILE_ICON_MAP = {
    doc:  ['file-word', 'description'],
    docx: ['file-word', 'description'],
    docm: ['file-word', 'description'],
    log:  ['file-word', 'fingerprint'],
    xls:  ['file-excel', 'table_chart'],
    xlsx: ['file-excel', 'table_chart'],
    xlsm: ['file-excel', 'table_chart'],
    csv:  ['file-excel', 'table_chart'],
    pdf:  ['file-pdf', 'picture_as_pdf'],
    txt:  ['file-text', 'article'],
    rtf:  ['file-text', 'description'],
    jpg:  ['file-image', 'image'],
    jpeg: ['file-image', 'image'],
    png:  ['file-image', 'image'],
    gif:  ['file-image', 'image'],
    bmp:  ['file-image', 'image'],
    zip:  ['file-archive', 'folder_zip'],
    rar:  ['file-archive', 'folder_zip'],
    '7z': ['file-archive', 'folder_zip'],
    r:    ['file-code', 'code'],
};

function fileIcon(icon, colorClass = '') {
    const cls = `material-symbols-outlined${colorClass ? ` ${colorClass}` : ''}`;
    return `<i class="${cls}">${icon}</i>`;
}

function getFileIcon(fileName, isDirectory) {
    if (isDirectory) return fileIcon('folder');
    if (fileName.startsWith('~')) return fileIcon('help');

    const extension = fileName.split('.').pop().toLowerCase();
    const entry = FILE_ICON_MAP[extension];

    return entry ? fileIcon(entry[1], entry[0]) : fileIcon('insert_drive_file');
}

async function openFolder(folderPath) {
    try {
        const response = await fetch('/api/open-folder', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ folderPath: folderPath })
        });

        const data = await response.json();

        if (!response.ok) {
            console.error('Erro ao abrir pasta:', data.error);
            showToast(`Erro ao abrir pasta:\n\n${data.error}\n\nCaminho: ${folderPath}`, 'error', 10000);
        }
    } catch (error) {
        showToast(`Erro ao abrir pasta:\n\n${error.message}\n\nCaminho: ${folderPath}\n\nVerifique o console para mais detalhes.`, 'error', 10000);
    }
}

async function openFile(filePath) {
    try {
        const response = await fetch('/api/open-file', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ filePath: filePath })
        });

        const data = await response.json();

        if (!response.ok) {
            console.error('Erro ao abrir arquivo:', data.error);
            showToast(`Erro ao abrir arquivo:\n\n${data.error}\n\nCaminho: ${filePath}`, 'error', 10000);
        } else {
            showToast('Abrindo arquivo...', 'info', 3000);
        }
    } catch (error) {
        console.error('Erro ao abrir arquivo:', error);
        showToast(`Erro ao abrir arquivo:\n\n${error.message}\n\nCaminho: ${filePath}\n\nVerifique o console para mais detalhes.`, 'error', 10000);
    }
}

async function clearFolderFiles(containerId, folderPath, scriptName, innerFolder) {
    // Escapar caminho para exibição segura em HTML
    const safePath = folderPath.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&#039;");

    const filesList = document.getElementById(containerId);
    const extensions = filesList ? filesList.dataset.extensions : null;
    const nameContains = filesList ? filesList.dataset.nameContains : null;

    let message = 'Tem certeza que deseja excluir <strong>TODOS</strong> os arquivos da pasta?';
    if (extensions || nameContains) {
        let filtros = [];
        if (extensions) filtros.push(`extensões: ${extensions}`);
        if (nameContains) filtros.push(`filtro: ${nameContains}`);
        message = `Tem certeza que deseja excluir os arquivos filtrados (<strong>${filtros.join(', ')}</strong>) da pasta?`;
    }

    const confirmed = await showConfirmationModal({
        title: 'Confirmar Exclusão',
        message: message,
        detail: `<i class="material-symbols-outlined me-1">folder</i> ${safePath}`,
        confirmText: 'Excluir',
        confirmColor: 'btn-danger'
    });

    if (!confirmed) return;

    try {
        filesList.innerHTML = `
            <div class="custom-spinner text-danger">
                <div class="spinner-border"></div>
                <span role="status">Excluindo arquivos...</span>
            </div>
        `;

        let url = `/api/clear-folder/${scriptName}/${innerFolder}`;
        let params = new URLSearchParams();
        if (extensions) {
            params.append('extensions', extensions);
        }
        if (nameContains) {
            params.append('nameContains', nameContains);
        }

        let qs = params.toString();
        if (qs) {
            url += `?${qs}`;
        }

        const response = await fetch(url, {
            method: 'DELETE'
        });

        const data = await response.json();

        if (!response.ok) {
            console.error('Erro ao limpar pasta:', data.error);
            showToast(`Erro ao limpar pasta:\n\n${data.error}`, 'error', 10000);
            await loadFiles(containerId, scriptName, innerFolder, false);
            return;
        }

        showToast('Arquivos excluídos com sucesso', 'success');
        await loadFiles(containerId, scriptName, innerFolder, false);
    } catch (error) {
        console.error('Erro ao limpar pasta:', error);
        showToast(`Erro ao limpar pasta:\n\n${error.message}`, 'error', 10000);
        await loadFiles(containerId, scriptName, innerFolder, false);
    }
}

async function refreshScriptFileLists(scriptName) {
    const aba = document.getElementById(scriptName);
    const fileLists = aba.querySelectorAll('.files-list');

    const promises = [];
    fileLists.forEach((fileList) => {
        const containerId = fileList.id;
        const folder = fileList.dataset.folder;
        const selectable = fileList.dataset.selectable === 'true';

        promises.push(loadFiles(containerId, scriptName, folder, selectable));
    });

    await Promise.all(promises);
}

function refreshFileList(containerId, scriptName) {
    const fileList = document.getElementById(containerId);

    if (!scriptName) {
        const mainTabPane = fileList.closest('.tab-pane[data-load-url]');
        if (mainTabPane) {
            scriptName = mainTabPane.id;
        } else {
            scriptName = fileList.closest('.tab-pane').id;
        }
    }

    const folder = fileList.dataset.folder;
    const selectable = fileList.dataset.selectable === 'true';

    return loadFiles(containerId, scriptName, folder, selectable);
}

async function loadFiles(containerId, scriptName, innerFolder, selectable) {
    const filesList = document.getElementById(containerId);
    if (!filesList) return;
    const extensions = filesList.dataset.extensions;
    const nameContains = filesList.dataset.nameContains;
    const sort = filesList.dataset.sort;

    try {
        filesList.innerHTML = `
            <div class="custom-spinner text-primary">
                <div class="spinner-border"></div>
                <span role="status">Atualizando lista de arquivos...</span>
            </div>
        `;

        let url = `/api/list-files/${scriptName}/${innerFolder}`;
        let params = new URLSearchParams();
        if (extensions) {
            params.append('extensions', extensions);
        }
        if (nameContains) {
            params.append('nameContains', nameContains);
        }
        if (sort) {
            params.append('sort', sort);
        }

        let qs = params.toString();
        if (qs) url += `?${qs}`;

        const response = await fetch(url);
        const data = await response.json();

        if (data.error) {
            filesList.innerHTML = `
                <div class="alert alert-danger" role="alert">
                    <i class="material-symbols-outlined">close</i> Erro ao carregar arquivos: ${data.error}
                </div>
            `;
            return;
        }

        let filesHTML = `
            <div class="folder-path-container">
                <i class="material-symbols-outlined">folder</i>
                <span class="folder-path">${data.folderPath}</span>
                <button data-bs-toggle="tooltip" data-bs-title="Clique para copiar o caminho da pasta" class="btn btn-outline-secondary btn-sm copy-icon"}>
                    <i class="material-symbols-outlined">content_copy</i>
                </button>
            </div>
        `;

        const hasFiles = data.files && data.files.length > 0;
        const canDelete = data.canDelete;

        let filesHTMLButtons = `
            <div class="folder-buttons-container">
                ${canDelete ? `
                <button data-bs-toggle="tooltip" data-bs-title="Excluir todos os arquivos da pasta" class="btn btn-outline-danger btn-sm btn-clear" data-container-id="${containerId}" data-folder-path="${data.folderPath}" data-script-name="${scriptName}" data-inner-folder="${innerFolder}" ${hasFiles ? '' : 'disabled'}>
                    <i class="material-symbols-outlined">delete</i>
                </button>
                ` : ''}
                <button data-bs-toggle="tooltip" data-bs-title="Abrir pasta" class="btn btn-outline-secondary btn-sm btn-open" data-folder-path="${data.folderPath}">
                    <i class="material-symbols-outlined">folder_open</i>
                </button>
                <button data-bs-toggle="tooltip" data-bs-title="Atualizar lista de arquivos" class="btn btn-outline-secondary btn-sm btn-refresh" data-container-id="${containerId}" data-script-name="${scriptName}">
                    <i class="material-symbols-outlined">refresh</i>
                </button>
            </div>
        `;

        if (!data.files || data.files.length === 0) {
            filesHTML += `
                <div class="alert alert-warning" role="alert">
                    <i class="material-symbols-outlined">warning</i> Nenhum arquivo encontrado na pasta
                </div>
            `;
            filesHTML += filesHTMLButtons;
            filesList.innerHTML = filesHTML;
            setupFileListButtons(filesList);
            return;
        }

        filesHTML += `
            <div class="files-table-container">
                <table class="files-table">
                    <tbody>
        `;

        data.files.forEach((file, index) => {
            const modDate = new Date(file.modifiedDate).toLocaleString('pt-BR');
            const icon = getFileIcon(file.name, file.isDirectory);
            const fileId = `${containerId}-file-${index}`;
            const safeFolderPath = data.folderPath.replace(/"/g, '&quot;');
            const safeFileName = file.name.replace(/"/g, '&quot;');

            // Usamos '/' para unir que o backend entende em qualquer SO
            const fileFullPath = `${safeFolderPath}/${safeFileName}`;

            const isSelected = selectedFiles[containerId] === file.name ? 'selected' : '';
            if (selectable) {
                filesHTML += `
                    <tr class="selectable ${isSelected}" id="${fileId}" data-filepath="${fileFullPath}" onclick='selectFile("${containerId}", "${file.name}", "${fileId}")' ondblclick='openFile(this.dataset.filepath)'>
                `;
            } else {
                filesHTML += `
                    <tr data-filepath="${fileFullPath}" ondblclick='openFile(this.dataset.filepath)'>
                `;
            }
            filesHTML += `
                    <td><span class="file-icon">${icon}</span></td><td>${file.name}</td>
                    <td>${modDate}</td>
                </tr>
            `;
        });

        filesHTML += '</tbody></table></div>';
        filesHTML += filesHTMLButtons;

        filesList.innerHTML = filesHTML;

        colorSelectedRow(containerId);
        setupFileListButtons(filesList);
        initializeTooltips();

    } catch (error) {
        console.error('Erro ao carregar arquivos:', error);
        filesList.innerHTML = `
            <div class="alert alert-danger" role="alert">
                <i class="material-symbols-outlined">error</i> Erro ao carregar arquivos: ${error.message}
            </div>
        `;
    }

    document.dispatchEvent(new CustomEvent('files-loaded', {
        detail: { containerId }
    }));
}

function setupFileListButtons(filesList) {
    const btnRefresh = filesList.querySelector('.btn-refresh');
    const btnOpen = filesList.querySelector('.btn-open');
    const btnClear = filesList.querySelector('.btn-clear');
    const folderPathSpan = filesList.querySelector('.folder-path');
    const copyIcon = filesList.querySelector('.copy-icon');

    if (copyIcon) {
        copyIcon.addEventListener('click', () => {
            const textToCopy = folderPathSpan.textContent;

            navigator.clipboard.writeText(textToCopy).then(() => {
                showToast('Caminho copiado para a área de transferência', 'success');
            }).catch((err) => {
                console.error('Erro ao copiar para área de transferência:', err);
                showToast('Erro ao copiar o caminho para a área de transferência', 'error');
            });
        });
    }

    if (btnRefresh) {
        btnRefresh.addEventListener('click', async (e) => {
            e.preventDefault();

            removeTooltip(btnRefresh);

            const containerId = btnRefresh.dataset.containerId;
            const scriptName = btnRefresh.dataset.scriptName;
            await refreshFileList(containerId, scriptName);

            filesList = document.getElementById(containerId);
            validateSingleField(filesList);
            colorSelectedRow(containerId);
        });
    }

    if (btnOpen) {
        btnOpen.addEventListener('click', (e) => {
            e.preventDefault();
            const folderPath = btnOpen.dataset.folderPath;
            openFolder(folderPath);
            showToast('Abrindo pasta. Verifique na barra de tarefas...', 'info', 5000);
        });
    }

    if (btnClear) {
        btnClear.addEventListener('click', (e) => {
            e.preventDefault();

            removeTooltip(btnClear);

            const containerId = btnClear.dataset.containerId;
            const folderPath = btnClear.dataset.folderPath;
            const scriptName = btnClear.dataset.scriptName;
            const innerFolder = btnClear.dataset.innerFolder;
            clearFolderFiles(containerId, folderPath, scriptName, innerFolder);
        });
    }

    // Se é a lista de relatórios SICAF, numerar as atas após carregamento
    if (filesList.id === 'atas-relatorios-sicaf' && typeof numerarAtas === 'function') {
        numerarAtas();
    }
}

function selectFile(containerId, fileName, fileId) {
    const container = document.getElementById(containerId);
    if (container) {
        const previousSelected = container.querySelector('tr.selected');
        if (previousSelected) {
            previousSelected.classList.remove('selected');
            previousSelected.style.borderLeft = '';
            previousSelected.style.backgroundColor = '';
        }
    }

    const fileRow = document.getElementById(fileId);
    if (fileRow) {
        fileRow.classList.add('selected');
        colorSelectedRow(containerId);
    }

    selectedFiles[containerId] = fileName;

    const event = new CustomEvent('file-selected', {
        detail: { containerId: containerId, fileName: fileName }
    });
    document.dispatchEvent(event);
}

function colorSelectedRow(containerId) {
    const container = document.getElementById(containerId);
    if (!container) return;

    const fileRow = container.querySelector('tr.selected');
    if (!fileRow) return;

    // Pegar a cor do ícone da primeira célula
    const iconElement = fileRow.querySelector('.file-icon i');
    let iconColor = '#28a745'; // cor padrão

    if (iconElement && iconElement.style.color) {
        iconColor = iconElement.style.color;
    }

    fileRow.style.borderLeft = `4px solid ${iconColor}`;

    // Tornar o fundo mais claro aplicando transparência (15%)
    if (iconColor.startsWith('#')) {
        fileRow.style.backgroundColor = `${iconColor}26`;
    } else if (iconColor.startsWith('rgb')) {
        fileRow.style.backgroundColor = iconColor.replace('rgb', 'rgba').replace(')', ', 0.15)');
    } else {
        fileRow.style.backgroundColor = 'rgba(0, 0, 0, 0.05)';
    }

}
