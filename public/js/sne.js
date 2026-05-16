let sneAnalysis = [];
let sneGrouped = new Map();
let sneSelectedCnpj = null;
let sneSortBy = 'nome';
let sneFolderPath = '';

const VALIDITY_ORDER = { 'VALIDA': 5, 'SEM_VALIDADE': 4, 'A_VENCER': 3, 'VENCIDA': 1 };
// Mínimo de certidões individuais para cobrir um SICAF ausente ou vencido
const MANDATORY_INDIVIDUAL = ['Receita Federal', 'FGTS', 'Trabalhista'];

// ====== INICIALIZAÇÃO ======

function inicializarSne() {
    carregarFornecedores();

    document.getElementById('btn-analisar-certidoes')?.addEventListener('click', analisarCertidoes);

    document.querySelectorAll('input[name="sne-sort"]').forEach(radio => {
        radio.addEventListener('change', (e) => {
            sneSortBy = e.target.value;
            renderFornecedores();
        });
    });
}

// ====== CARREGAMENTO ======

async function carregarFornecedores() {
    const container = document.getElementById('sne-fornecedores-list');
    if (!container) return;

    container.innerHTML = customSpinnerHTML('Analisando certidões...');

    try {
        const response = await fetch('/api/sne/certidoes/analisar');
        const data = await response.json();

        if (data.error) {
            container.innerHTML = alertHTML('danger', 'error', data.error);
            return;
        }

        sneAnalysis = data.results || [];
        sneFolderPath = data.folderPath || '';
        sneGrouped = groupByCnpj(sneAnalysis);
        renderFornecedores();

        if (sneSelectedCnpj && sneGrouped.has(sneSelectedCnpj)) {
            selecionarFornecedor(sneSelectedCnpj);
        } else if (sneSelectedCnpj) {
            sneSelectedCnpj = null;
            resetCertidoesPanel();
        }
    } catch (error) {
        container.innerHTML = alertHTML('danger', 'error', `Erro ao analisar certidões: ${error.message}`);
    }
}

// ====== AGRUPAMENTO ======

function groupByCnpj(results) {
    const grouped = new Map();

    for (const r of results) {
        const key = r.cnpj || '__sem_cnpj__';
        if (!grouped.has(key)) {
            grouped.set(key, { cnpj: r.cnpj, company: r.company, certidoes: [] });
        }
        const group = grouped.get(key);
        if (!group.company && r.company) group.company = r.company;
        group.certidoes.push(r);
    }

    return grouped;
}

// ====== STATUS DO FORNECEDOR ======

function betterValidity(a, b) {
    return (VALIDITY_ORDER[a] || 0) >= (VALIDITY_ORDER[b] || 0) ? a : b;
}

function computeSupplierStatus(certidoes) {
    if (certidoes.some(c => c.impedido)) return 'impedido';

    const typeBest = {};
    const hasParseErrors = certidoes.some(c => c.error);

    for (const c of certidoes) {
        if (c.error || !c.type) continue;
        typeBest[c.type] = typeBest[c.type] !== undefined
            ? betterValidity(typeBest[c.type], c.validity)
            : c.validity;
    }

    if (!typeBest['Credenciamento']) return 'erro';

    const sicafValidity = typeBest['SICAF'];
    let sicafCoveredByIndividual = false;

    if (!sicafValidity || sicafValidity === 'VENCIDA') {
        // SICAF ausente ou vencido — verifica cobertura por certidões individuais
        const sicafCert = certidoes.find(c => !c.error && c.type === 'SICAF');
        if (sicafCert?.componentValidity) {
            // Cada componente vencido no SICAF deve ter uma certidão avulsa válida ou ao menos presente.
            // null = arquivo existe mas não verificável → passa (hasUnverifiable produz 'alerta' adiante).
            // undefined = arquivo ausente → não cobre → 'erro'.
            sicafCoveredByIndividual = Object.entries(sicafCert.componentValidity).every(([label, cv]) => {
                if (cv !== 'VENCIDA') return true;
                const standalone = typeBest[label];
                if (standalone === null) return true;
                return standalone && standalone !== 'VENCIDA';
            });
        } else {
            sicafCoveredByIndividual = MANDATORY_INDIVIDUAL.every(t => {
                const v = typeBest[t];
                if (v === null) return true;
                return v && v !== 'VENCIDA';
            });
        }
        if (!sicafCoveredByIndividual) return 'erro';
    }

    for (const [type, best] of Object.entries(typeBest)) {
        if (type === 'Credenciamento') continue;
        if (type === 'SICAF' && sicafCoveredByIndividual) continue;
        if (best === 'VENCIDA') return 'erro';
    }

    const hasAVencer = Object.entries(typeBest).some(([type, v]) => {
        if (type === 'SICAF' && sicafCoveredByIndividual) return false;
        return v === 'A_VENCER';
    });
    const hasUnverifiable = certidoes.some(c =>
        !c.error && c.type && c.type !== 'Credenciamento' && !c.validity
    );

    if (hasParseErrors || hasAVencer || hasUnverifiable) return 'alerta';
    return 'ok';
}

// ====== RENDERIZAÇÃO DE FORNECEDORES ======

function renderFornecedores() {
    const container = document.getElementById('sne-fornecedores-list');
    if (!container) return;

    const displayPath = sneFolderPath.replace(/\\/g, '/');

    const refreshBtn = `
        <button class="folder-path-btn btn-refresh-fornecedores"
            data-bs-toggle="tooltip" data-bs-title="Atualizar lista">
            <i class="material-symbols-outlined">refresh</i>
        </button>`;

    if (sneGrouped.size === 0) {
        container.innerHTML = buildFolderPathHTML(displayPath, '', refreshBtn) +
            `<div class="alert alert-warning"><i class="material-symbols-outlined">warning</i> Nenhuma certidão encontrada na pasta CERTIDOES</div>`;
        setupFolderPathButtons(container);
        setupRefreshFornecedoresButton(container);
        initializeTooltips();
        return;
    }

    let entries = [...sneGrouped.entries()];

    if (sneSortBy === 'cnpj') {
        entries.sort(([, a], [, b]) => {
            if (a.cnpj && b.cnpj) return a.cnpj.localeCompare(b.cnpj);
            if (a.cnpj) return -1;
            if (b.cnpj) return 1;
            return (a.company || '').localeCompare(b.company || '');
        });
    } else {
        entries.sort(([, a], [, b]) => {
            const na = (a.company || '').toLowerCase();
            const nb = (b.company || '').toLowerCase();
            if (na && nb) return na.localeCompare(nb);
            if (na) return -1;
            if (nb) return 1;
            return (a.cnpj || '').localeCompare(b.cnpj || '');
        });
    }

    const STATUS_MAP = {
        ok:       { cssClass: 'status-sucesso',  icon: 'check_circle', tooltip: 'Todas as certidões válidas' },
        alerta:   { cssClass: 'status-parcial',  icon: 'warning',      tooltip: 'Certidões a vencer ou com data não verificada' },
        erro:     { cssClass: 'status-sem_saida', icon: 'cancel',      tooltip: 'Certidões vencidas ou ausentes' },
        impedido: { cssClass: 'status-sem_saida', icon: 'gavel',       tooltip: 'Fornecedor impedido de licitar' },
    };

    let html = buildFolderPathHTML(displayPath, '', refreshBtn) + '<div class="items-grid">';

    for (const [key, group] of entries) {
        const status = computeSupplierStatus(group.certidoes);
        const { cssClass, icon, tooltip } = STATUS_MAP[status];
        const isSelected = sneSelectedCnpj === key ? ' item-selected selected' : '';
        const displayName = group.company || '(Razão social não identificada)';
        const displayCnpj = group.cnpj ? formatCnpj(group.cnpj) : '(CNPJ não identificado)';
        const safeKey = key.replace(/['"\\]/g, '');;

        html += `
            <div class="item-card sne-supplier-card ${cssClass}${isSelected}" data-supplier="${safeKey}"
                 onclick="selecionarFornecedor('${safeKey}')">
                <div class="item-card-status">
                    <i class="material-symbols-outlined" data-bs-toggle="tooltip"
                        data-bs-title="${tooltip}">${icon}</i>
                </div>
                <div class="sne-supplier-name">${displayName}</div>
                <div class="sne-supplier-cnpj">${displayCnpj}</div>
                <button class="item-card-open btn btn-sm text-primary file-row-btn"
                    data-bs-toggle="tooltip" data-bs-title="Abrir arquivos do fornecedor" data-bs-placement="right"
                    onclick="event.stopPropagation(); openFornecedorFiles('${safeKey}')">
                    <i class="material-symbols-outlined">folder_open</i>
                </button>
                <button class="item-card-delete btn btn-sm text-danger file-row-btn"
                    data-bs-toggle="tooltip" data-bs-title="Excluir certidões deste fornecedor" data-bs-placement="right"
                    onclick="event.stopPropagation(); excluirFornecedor('${safeKey}')">
                    <i class="material-symbols-outlined">delete</i>
                </button>
            </div>`;
    }

    html += '</div>';
    container.innerHTML = html;
    setupFolderPathButtons(container);
    setupRefreshFornecedoresButton(container);
    initializeTooltips();
}

function setupRefreshFornecedoresButton(container) {
    container.querySelector('.btn-refresh-fornecedores')?.addEventListener('click', () => {
        carregarFornecedores();
    });
}

// ====== SELEÇÃO DE FORNECEDOR ======

function selecionarFornecedor(key) {
    sneSelectedCnpj = key;

    document.querySelectorAll('.item-card[data-supplier]').forEach(card => {
        const isSelected = card.dataset.supplier === key;
        card.classList.toggle('item-selected', isSelected);
        card.classList.toggle('selected', isSelected);
    });

    const group = sneGrouped.get(key);
    if (group) renderCertidoesDoFornecedor(group);
}

function renderCertidoesDoFornecedor(group) {
    const container = document.getElementById('sne-certidoes-list');
    const titulo = document.getElementById('sne-certidoes-titulo');
    if (!container) return;

    if (titulo) {
        titulo.textContent = group.company || (group.cnpj ? formatCnpj(group.cnpj) : 'Certidões');
    }

    const results = group.certidoes;

    if (!results || results.length === 0) {
        container.innerHTML = `<div class="alert alert-warning"><i class="material-symbols-outlined">warning</i> Nenhuma certidão encontrada.</div>`;
        return;
    }

    let html = `
        <div class="files-table-container">
            <table class="files-table">
                <thead>
                    <tr>
                        <th colspan="2">Arquivo</th>
                        <th>Tipo</th>
                        <th>Situação</th>
                        <th></th>
                    </tr>
                </thead>
                <tbody>`;

    results.forEach(r => {
        const { icon, cssClass } = validityDisplay(r);
        const typeLabel = r.type || '<span class="text-muted">Não identificado</span>';
        const typeClass = r.type ? '' : 'text-warning';
        const warnings = r.warnings?.length > 0
            ? `<br><small class="text-warning">${r.warnings.join('; ')}</small>`
            : '';
        const safeFilePath = (sneFolderPath + '\\' + r.filename).replace(/\\/g, '\\\\').replace(/"/g, '&quot;');
        const safeFilename = r.filename.replace(/'/g, "\\'");

        html += `
            <tr>
                <td><i class="material-symbols-outlined ${cssClass}" data-bs-toggle="tooltip"
                    data-bs-title="${validityTooltip(r)}">${icon}</i></td>
                <td class="text-break">${r.filename}${warnings}</td>
                <td class="${typeClass}">${typeLabel}</td>
                <td class="${cssClass}"><small>${getSituacaoLabel(r)}</small></td>
                <td class="table-btn-column text-nowrap">
                    <button class="btn btn-sm text-primary file-row-btn"
                        onclick="openFile('${safeFilePath}')"
                        data-bs-toggle="tooltip" data-bs-title="Abrir arquivo">
                        <i class="material-symbols-outlined">open_in_new</i>
                    </button>
                    <button class="btn btn-sm text-danger file-row-btn"
                        onclick="excluirArquivoCertidao('${safeFilename}')"
                        data-bs-toggle="tooltip" data-bs-title="Excluir arquivo">
                        <i class="material-symbols-outlined">delete</i>
                    </button>
                </td>
            </tr>`;
    });

    html += '</tbody></table></div>';
    container.innerHTML = html;
    initializeTooltips();
}

function resetCertidoesPanel() {
    const container = document.getElementById('sne-certidoes-list');
    const titulo = document.getElementById('sne-certidoes-titulo');
    if (container) {
        container.innerHTML = `<div class="alert alert-info" role="alert">
            <i class="material-symbols-outlined">info</i> Selecione um fornecedor para ver as certidões
        </div>`;
    }
    if (titulo) titulo.textContent = 'Certidões';
}

// ====== ABERTURA DE ARQUIVOS ======

function openFornecedorFiles(key) {
    const group = sneGrouped.get(key);
    if (!group) return;
    group.certidoes.forEach(c => openFile(sneFolderPath + '\\' + c.filename));
}

// ====== EXCLUSÃO ======

async function excluirFornecedor(key) {
    const group = sneGrouped.get(key);
    if (!group) return;

    const displayName = group.company || formatCnpj(group.cnpj) || key;
    const filenames = group.certidoes.map(c => c.filename);

    const confirmed = await showConfirmationModal({
        title: 'Excluir Certidões do Fornecedor',
        message: `Tem certeza que deseja excluir todas as certidões de <strong>${displayName}</strong>?`,
        detail: `<i class="material-symbols-outlined me-1">warning</i> ${filenames.length} arquivo(s) serão excluídos permanentemente.`,
        confirmText: 'Excluir',
        confirmColor: 'btn-danger',
    });

    if (!confirmed) return;

    await deleteCertidoes(filenames);
}

async function excluirArquivoCertidao(filename) {
    const confirmed = await showConfirmationModal({
        title: 'Excluir Certidão',
        message: `Tem certeza que deseja excluir <strong>${filename}</strong>?`,
        detail: '<i class="material-symbols-outlined me-1">warning</i> Esta ação não pode ser desfeita.',
        confirmText: 'Excluir',
        confirmColor: 'btn-danger',
    });

    if (!confirmed) return;

    await deleteCertidoes([filename]);
}

async function deleteCertidoes(filenames) {
    try {
        const response = await fetch('/api/sne/certidoes', {
            method: 'DELETE',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ filenames }),
        });
        const data = await response.json();

        if (!response.ok && response.status !== 207) {
            showToast(`Erro ao excluir: ${data.error}`, 'error');
            return;
        }

        showToast(data.message, 'success');
        carregarFornecedores();
    } catch (error) {
        showToast(`Erro ao excluir: ${error.message}`, 'error');
    }
}

// ====== ANALISAR E RENOMEAR ======

async function analisarCertidoes() {
    const confirmed = await showConfirmationModal({
        title: 'Analisar e Renomear Certidões',
        message: 'Os arquivos PDF serão analisados e renomeados com CNPJ, razão social e status de validade.',
        detail: '<i class="material-symbols-outlined me-1">warning</i> Esta ação não pode ser desfeita. Em caso de nomes duplicados, o arquivo mais antigo será <strong>excluído permanentemente</strong>.',
        confirmText: 'Analisar e Renomear',
        confirmColor: 'btn-primary',
    });

    if (!confirmed) return;

    prepareConsoleForExecution('sne_analisar');

    try {
        const response = await fetch('/api/sne/certidoes/renomear', { method: 'POST' });
        const data = await response.json();

        if (!response.ok) {
            handleScriptResult({ scriptName: 'sne_analisar', status: 'error', message: data.error || 'Erro ao processar certidões.', log: '' });
            return;
        }

        const renamed = data.results.filter(r => r.renamed).length;
        const deleted = data.results.filter(r => r.deleted).length;
        const errors  = data.results.filter(r => r.renameError).length;
        const skipped = data.results.filter(r => r.renameSkipped).length;
        const noChange = data.results.filter(r => r.noChange).length;

        const C = {
            section: 'color: #6ea8fe; font-weight: bold;',
            success: 'color: #20c997;',
            warning: 'color: #ffc107;',
            error:   'color: #ff6b6b; font-weight: bold;',
            muted:   'color: #adb5bd;',
            info:    'color: inherit;',
        };
        const line = (style, text) => `<span style="${style}">${text}\n</span>`;

        let log = '';
        log += line(C.section, '[SNE] Análise e renomeação de certidões');
        log += line(C.info,    `[SNE] ${data.results.length} arquivo(s) na pasta CERTIDOES`);
        log += line(C.info,    '');

        for (const r of data.results) {
            const typeLabel = r.type || 'Tipo desconhecido';
            const who = r.company || (r.cnpj ? formatCnpj(r.cnpj) : null);
            const whoStr = who ? ` — ${who}` : '';

            if (r.renamed && r.conflictDuplicated) {
                log += line(C.warning, `  ⚠ Duplicata renomeada: "${r.filename}"`);
                log += line(C.warning, `    → "${r.newName}"`);
            } else if (r.renamed) {
                log += line(C.success, `  ✓ ${typeLabel}${whoStr}`);
                log += line(C.muted,   `    "${r.filename}"`);
                log += line(C.muted,   `    → "${r.newName}"`);
            } else if (r.deleted) {
                const kept = r.deletedKeptAs ? ` (mantido: "${r.deletedKeptAs}")` : '';
                log += line(C.warning, `  ✗ Eliminado (duplicata)${kept}: "${r.filename}"`);
            } else if (r.renameError) {
                log += line(C.error, `  ✗ Erro: "${r.filename}": ${r.renameError}`);
            } else if (r.renameSkipped) {
                const reason = r.error ? ` (${r.error})` : '';
                log += line(C.muted, `  — Sem nome${reason}: "${r.filename}"`);
            } else if (r.noChange) {
                log += line(C.muted, `  — Sem alteração: "${r.filename}"`);
            }
        }

        log += line(C.info, '');

        const summaryParts = [];
        if (renamed > 0)  summaryParts.push(`${renamed} renomeado(s)`);
        if (deleted > 0)  summaryParts.push(`${deleted} eliminado(s)`);
        if (errors > 0)   summaryParts.push(`${errors} com erro`);
        if (skipped > 0)  summaryParts.push(`${skipped} sem nome`);
        if (noChange > 0) summaryParts.push(`${noChange} sem alteração`);
        log += line(C.section, `[SNE] ${summaryParts.join(' · ')}`);

        consoleOutput.innerHTML = log;

        const parts = [`${renamed} renomeado(s)`];
        if (deleted > 0) parts.push(`${deleted} eliminado(s) por conflito`);
        if (errors + skipped > 0) parts.push(`${errors + skipped} com falha ou sem nome identificado`);
        const status = (errors + skipped) === 0 ? 'success' : 'warning';
        handleScriptResult({ scriptName: 'sne_analisar', status, message: parts.join(', ') + '.', log: '' });

        addNotification({
            message: 'Análise de certidões concluída. ' + parts.join(', ') + '.',
            type: status === 'success' ? 'success' : 'warning',
            source: 'SNE',
        });

        carregarFornecedores();
    } catch (error) {
        handleScriptResult({ scriptName: 'sne_analisar', status: 'error', message: `Erro: ${error.message}`, log: '' });
    }
}

// ====== UTILITÁRIOS DE EXIBIÇÃO ======

function formatCnpj(cnpj) {
    if (!cnpj || cnpj.length !== 14) return cnpj || '';
    return cnpj.replace(/(\d{2})(\d{3})(\d{3})(\d{4})(\d{2})/, '$1.$2.$3/$4-$5');
}

function getSituacaoLabel(r) {
    if (r.error) return 'Erro';
    if (r.impedido) return 'Impedido de licitar';
    if (r.type === 'SICAF' && r.componentValidity) {
        const vencidas = Object.entries(r.componentValidity).filter(([, v]) => v === 'VENCIDA').map(([l]) => l);
        const aVencer = Object.entries(r.componentValidity).filter(([, v]) => v === 'A_VENCER').map(([l]) => l);
        const parts = [];
        if (vencidas.length > 0) parts.push(`Vencida: ${vencidas.join(', ')}`);
        if (aVencer.length > 0) parts.push(`A vencer: ${aVencer.join(', ')}`);
        if (parts.length > 0) return parts.join(' | ');
    }
    if (r.validity === 'VENCIDA') return 'Vencida';
    if (r.validity === 'A_VENCER') return r.validityLabel || 'A vencer';
    if (r.validity === 'VALIDA') return r.validityLabel || 'Válida';
    if (r.validity === 'SEM_VALIDADE') return 'Credenciamento';
    return '—';
}

function validityDisplay(r) {
    if (r.error) return { icon: 'error', cssClass: 'text-danger' };
    if (r.impedido) return { icon: 'shield_lock', cssClass: 'text-danger' };
    if (r.validity === 'VENCIDA') return { icon: 'cancel', cssClass: 'text-danger' };
    if (r.validity === 'A_VENCER') return { icon: 'warning', cssClass: 'text-warning' };
    if (r.validity === 'VALIDA') return { icon: 'check_circle', cssClass: 'text-success' };
    if (r.validity === 'SEM_VALIDADE') return { icon: 'description', cssClass: 'text-muted' };
    return { icon: 'hourglass_empty', cssClass: 'text-muted' };
}

function validityTooltip(r) {
    if (r.error) return r.error;
    if (r.impedido) return 'Impedimento de licitar detectado';
    if (r.validityLabel) return r.validityLabel;
    if (r.validity === 'SEM_VALIDADE') return 'Credenciamento — sem verificação de validade';
    return 'Validade não encontrada';
}

function alertHTML(type, icon, message) {
    return `<div class="alert alert-${type}" role="alert">
        <i class="material-symbols-outlined">${icon}</i> ${message}
    </div>`;
}
