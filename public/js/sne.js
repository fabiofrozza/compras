let sneAnalysis = [];
let sneGrouped = new Map();
let sneSelectedCnpj = null;
let sneSortBy = 'nome';
let sneFolderPath = '';

let sneEmpenhos = [];
let sneEmpenhosFolderPath = '';
let sneEmpenhosSortState = { column: null, direction: 'asc' };

const VALIDITY_COVERAGE_ORDER = { 'VALIDA': 5, 'SEM_VALIDADE': 4, 'A_VENCER': 3, 'VENCIDA': 1 };
// Mínimo de certidões individuais para cobrir um SICAF ausente ou vencido
const MANDATORY_INDIVIDUAL = ['Receita Federal', 'FGTS', 'Trabalhista'];

// ====== INICIALIZAÇÃO ======

function inicializarSne() {
    carregarFornecedores();

    document.getElementById('btn-analisar-certidoes')?.addEventListener('click', analisarCertidoes);
    document.getElementById('btn-atualizar-empenhos')?.addEventListener('click', carregarEmpenhos);
    document.getElementById('btn-criar-afs')?.addEventListener('click', executarCriarAFs);

    document.querySelectorAll('input[name="sne-sort"]').forEach(radio => {
        radio.addEventListener('change', (e) => {
            sneSortBy = e.target.value;
            renderFornecedores();
        });
    });

    document.getElementById('tab-sne-empenhos')?.addEventListener('shown.bs.tab', () => {
        if (sneEmpenhosFolderPath) renderEmpenhos(); else carregarEmpenhos();
    });

    document.getElementById('btn-atualizar-afs')?.addEventListener('click', carregarAFs);
    document.getElementById('tab-sne-afs')?.addEventListener('shown.bs.tab', () => {
        if (sneAfsFolderPath) renderAFs(); else carregarAFs();
    });
}

// ====== CARREGAMENTO ======

async function carregarFornecedores() {
    const container = document.getElementById('sne-fornecedores-list');
    if (!container) return;

    container.innerHTML = customSpinnerHTML('Lendo certidões...');

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
    return (VALIDITY_COVERAGE_ORDER[a] || 0) >= (VALIDITY_COVERAGE_ORDER[b] || 0) ? a : b;
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

    const deleteBtnCertidoes = `
        <button class="folder-path-btn text-danger btn-clear-certidoes"
            data-bs-toggle="tooltip" data-bs-title="Excluir todos os arquivos da pasta"
            ${sneGrouped.size === 0 ? 'disabled' : ''}>
            <i class="material-symbols-outlined">delete</i>
        </button>`;

    if (sneGrouped.size === 0) {
        container.innerHTML = buildFolderPathHTML(displayPath, deleteBtnCertidoes, refreshBtn) +
            `<div class="alert alert-warning"><i class="material-symbols-outlined">warning</i> Nenhuma certidão encontrada na pasta CERTIDOES</div>`;
        setupFolderPathButtons(container);
        setupRefreshFornecedoresButton(container);
        initializeTooltips();
        evaluateAllButtons();
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
        ok: { cssClass: 'status-sucesso', icon: 'check_circle', tooltip: 'Todas as certidões válidas' },
        alerta: { cssClass: 'status-parcial', icon: 'warning', tooltip: 'Certidões a vencer ou com data não verificada' },
        erro: { cssClass: 'status-sem_saida', icon: 'cancel', tooltip: 'Certidões vencidas ou ausentes' },
        impedido: { cssClass: 'status-sem_saida', icon: 'gavel', tooltip: 'Fornecedor impedido de licitar' },
    };

    let html = buildFolderPathHTML(displayPath, deleteBtnCertidoes, refreshBtn) + '<div class="items-grid">';

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
    evaluateAllButtons();
}

function setupRefreshFornecedoresButton(container) {
    container.querySelector('.btn-refresh-fornecedores')?.addEventListener('click', carregarFornecedores);
    container.querySelector('.btn-clear-certidoes')?.addEventListener('click', (e) => {
        removeTooltip(e.currentTarget);
        excluirTodosCertidoes();
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
            <tr ondblclick="openFile('${safeFilePath}')">
                <td><i class="material-symbols-outlined ${cssClass}" data-bs-toggle="tooltip"
                    data-bs-title="${validityTooltip(r)}">${icon}</i></td>
                <td class="text-break">${r.filename}${warnings}</td>
                <td class="${typeClass}">${typeLabel}</td>
                <td class="${cssClass}">${getSituacaoLabel(r)}</td>
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
        const errors = data.results.filter(r => r.renameError).length;
        const skipped = data.results.filter(r => r.renameSkipped).length;
        const noChange = data.results.filter(r => r.noChange).length;

        const C = {
            section: 'color: #6ea8fe; font-weight: bold;',
            success: 'color: #20c997;',
            warning: 'color: #ffc107;',
            error: 'color: #ff6b6b; font-weight: bold;',
            muted: 'color: #adb5bd;',
            info: 'color: inherit;',
        };
        const line = (style, text) => `<span style="${style}">${text}\n</span>`;

        let log = '';
        log += line(C.section, '[SNE] Análise e renomeação de certidões');
        log += line(C.info, `[SNE] ${data.results.length} arquivo(s) na pasta CERTIDOES`);
        log += line(C.info, '');

        for (const r of data.results) {
            const typeLabel = r.type || 'Tipo desconhecido';
            const who = r.company || (r.cnpj ? formatCnpj(r.cnpj) : null);
            const whoStr = who ? ` — ${who}` : '';

            if (r.renamed && r.conflictDuplicated) {
                log += line(C.warning, `  ⚠ Duplicata renomeada: "${r.filename}"`);
                log += line(C.warning, `    → "${r.newName}"`);
            } else if (r.renamed) {
                log += line(C.success, `  ✓ ${typeLabel}${whoStr}`);
                log += line(C.muted, `    "${r.filename}"`);
                log += line(C.muted, `    → "${r.newName}"`);
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
        if (renamed > 0) summaryParts.push(`${renamed} renomeado(s)`);
        if (deleted > 0) summaryParts.push(`${deleted} eliminado(s)`);
        if (errors > 0) summaryParts.push(`${errors} com erro`);
        if (skipped > 0) summaryParts.push(`${skipped} sem nome`);
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

        await carregarFornecedores();
        if (sneEmpenhosFolderPath) renderEmpenhos();
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
    if (r.validity === 'SEM_VALIDADE') return r.emissionDate ? `Emitido em ${r.emissionDate}` : 'Credenciamento';
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
    if (r.validity === 'SEM_VALIDADE') return r.emissionDate ? `Credenciamento emitido em ${r.emissionDate}` : 'Credenciamento — sem verificação de validade';
    return 'Validade não encontrada';
}

function alertHTML(type, icon, message) {
    return `<div class="alert alert-${type}" role="alert">
        <i class="material-symbols-outlined">${icon}</i> ${message}
    </div>`;
}

// ====== EMPENHOS ======

const EMPENHO_STATUS_MAP = {
    ok: { cssClass: 'text-success', icon: 'check_circle', label: 'Certidões OK' },
    alerta: { cssClass: 'text-warning', icon: 'warning', label: 'Certidões com alerta' },
    erro: { cssClass: 'text-danger', icon: 'cancel', label: 'Certidões irregulares' },
    impedido: { cssClass: 'text-danger', icon: 'gavel', label: 'Fornecedor impedido' },
};

function navegarParaCertidoesDoCnpj(cnpj) {
    if (!cnpj || !sneGrouped.has(cnpj)) return;
    const tab = document.getElementById('tab-sne-certidoes');
    if (tab) bootstrap.Tab.getOrCreateInstance(tab).show();
    selecionarFornecedor(cnpj);
}

function scrollParaAF(afName) {
    const group = [...document.querySelectorAll('.sne-af-group')].find(el => el.dataset.afName === afName);
    if (group) group.scrollIntoView({ behavior: 'smooth', block: 'center' });
}

function navegarParaAF(afName) {
    snePendingAfScroll = afName;
    const tabEl = document.getElementById('tab-sne-afs');
    if (!tabEl) return;
    if (tabEl.getAttribute('aria-selected') === 'true') {
        requestAnimationFrame(() => {
            scrollParaAF(snePendingAfScroll);
            snePendingAfScroll = null;
        });
    } else {
        bootstrap.Tab.getOrCreateInstance(tabEl).show();
    }
}

function getSupplierStatusByCnpj(cnpj) {
    if (!cnpj) return null;
    const group = sneGrouped.get(cnpj);
    if (!group) return null;
    return computeSupplierStatus(group.certidoes);
}

function buildTombCell(hasTombamento) {
    return hasTombamento
        ? `<i class="material-symbols-outlined text-warning" data-bs-toggle="tooltip" data-bs-title="Contém tombamento">inventory_2</i>`
        : `<span class="text-muted">—</span>`;
}

function buildStatusCell(cnpj) {
    const supplierStatus = getSupplierStatusByCnpj(cnpj);
    if (!supplierStatus) return '<span class="text-muted">—</span>';
    const { cssClass, icon, label } = EMPENHO_STATUS_MAP[supplierStatus];
    return `<button class="btn btn-sm file-row-btn sne-status-link ${cssClass}"
        data-bs-toggle="tooltip" data-bs-title="${label}<br><em>Clique para exibir</em>"
        onclick="event.stopPropagation(); navegarParaCertidoesDoCnpj('${cnpj}')">
        <i class="material-symbols-outlined">${icon}</i>
    </button>`;
}

function buildAfStatusCell(certidoes, afName, sneName, cnpj) {
    const safeAf = afName.replace(/'/g, "\\'");
    const safeSne = sneName.replace(/'/g, "\\'");
    const syncOnClick = `event.stopPropagation(); sincronizarCertidoesSne('${safeAf}', '${safeSne}')`;

    if (!certidoes || certidoes.length === 0) {
        if (!cnpj) return '<span class="text-muted">—</span>';
        return `<button class="btn btn-sm file-row-btn sne-status-link text-muted"
            data-bs-toggle="tooltip" data-bs-title="Sem certidões<br><em>Clique para atualizar certidões</em>"
            onclick="${syncOnClick}">
            <i class="material-symbols-outlined">refresh</i>
        </button>`;
    }

    const supplierStatus = computeSupplierStatus(certidoes);
    const { cssClass, icon, label } = EMPENHO_STATUS_MAP[supplierStatus];

    return `<button class="btn btn-sm file-row-btn sne-status-link ${cssClass}"
        data-bs-toggle="tooltip" data-bs-title="${label}<br><em>Clique para atualizar certidões</em>"
        onclick="${syncOnClick}">
        <i class="material-symbols-outlined">${icon}</i>
    </button>`;
}

async function sincronizarCertidoesSne(afName, sneName) {
    const afData = sneAfsAnalysis.find(af => af.name === afName);
    const sneData = afData?.snes.find(s => s.name === sneName);
    const cnpj = sneData?.empenho?.cnpj;

    if (!cnpj) {
        showToast('CNPJ não identificado no empenho da SNE.', 'warning');
        return;
    }

    try {
        const resp = await fetch('/api/sne/afs/sincronizar-certidoes', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ afName, sneName, cnpj }),
        });
        const data = await resp.json();

        if (!resp.ok) {
            showToast(`Erro ao sincronizar: ${data.error}`, 'error');
            return;
        }

        const msg = data.copied > 0
            ? `${data.copied} certidão(ões) atualizada(s) em ${sneName}.`
            : `Nenhuma certidão encontrada na pasta Certidões para este CNPJ.`;
        showToast(msg, data.copied > 0 ? 'success' : 'warning');
        carregarAFs();
    } catch (e) {
        showToast(`Erro: ${e.message}`, 'error');
    }
}

function getEmpenhoForSne(afName, sneName) {
    const afMatch = afName.match(/^AF (\d+)-(\d+)$/);
    if (!afMatch) return null;
    const sneNum = sneName.replace(/^SNE /, '');
    return sneEmpenhos.find(r => !r.error && r.af &&
        r.af.number === afMatch[1] && r.af.year === afMatch[2] &&
        r.sneNumber === sneNum) || null;
}

async function carregarEmpenhos() {
    const container = document.getElementById('sne-empenhos-container');
    if (!container) return;

    container.innerHTML = customSpinnerHTML('Lendo empenhos...');

    try {
        const [empenhoResp, afsResp] = await Promise.all([
            fetch('/api/sne/empenhos/analisar'),
            fetch('/api/sne/afs'),
        ]);
        const data = await empenhoResp.json();
        const afsData = await afsResp.json();

        if (data.error) {
            container.innerHTML = alertHTML('danger', 'error', data.error);
            return;
        }

        sneEmpenhos = data.results || [];
        sneEmpenhosFolderPath = data.folderPath || '';

        if (!afsData.error) {
            sneAfsData = afsData.afs || [];
        }

        renderEmpenhos();
    } catch (error) {
        container.innerHTML = alertHTML('danger', 'error', `Erro ao analisar empenhos: ${error.message}`);
    }
}

function getSneEmpenhoSortValue(r, column) {
    switch (column) {
        case 'arquivo':    return r.filename.toLowerCase();
        case 'sne':        return parseInt(r.sneNumber || '0', 10);
        case 'af':         return r.af ? parseInt(r.af.year) * 1e6 + parseInt(r.af.number) : -1;
        case 'cnpj':       return r.cnpj || '';
        case 'fornecedor': return (r.company || '').toLowerCase();
        case 'tombamento': return r.tombamento ? 1 : 0;
        case 'certidoes': {
            const order = { ok: 1, alerta: 2, erro: 3, impedido: 4 };
            return order[getSupplierStatusByCnpj(r.cnpj)] || 0;
        }
        case 'pasta-sne': {
            if (!r.af || !r.sneNumber) return -1;
            const afData = sneAfsData.find(af => af.name === `AF ${r.af.number}-${r.af.year}`);
            return afData?.snes?.some(s => s.name === `SNE ${r.sneNumber}`) ? 1 : 0;
        }
        default: return '';
    }
}

function sortSneEmpenhos(column) {
    const current = sneEmpenhosSortState;
    if (current.column === column) {
        sneEmpenhosSortState = current.direction === 'asc'
            ? { column, direction: 'desc' }
            : { column: null, direction: 'asc' };
    } else {
        sneEmpenhosSortState = { column, direction: 'asc' };
    }
    renderEmpenhos();
}

function renderEmpenhos() {
    const container = document.getElementById('sne-empenhos-container');
    if (!container) return;

    const displayPath = sneEmpenhosFolderPath.replace(/\\/g, '/');

    const refreshBtn = `
        <button class="folder-path-btn btn-refresh-empenhos"
            data-bs-toggle="tooltip" data-bs-title="Atualizar lista">
            <i class="material-symbols-outlined">refresh</i>
        </button>`;

    const deleteBtnEmpenhos = `
        <button class="folder-path-btn text-danger btn-clear-empenhos"
            data-bs-toggle="tooltip" data-bs-title="Excluir todos os arquivos da pasta"
            ${sneEmpenhos.length === 0 ? 'disabled' : ''}>
            <i class="material-symbols-outlined">delete</i>
        </button>`;

    if (sneEmpenhos.length === 0) {
        container.innerHTML = buildFolderPathHTML(displayPath, deleteBtnEmpenhos, refreshBtn) +
            alertHTML('warning', 'warning', 'Nenhum empenho encontrado na pasta SNEs');
        setupFolderPathButtons(container);
        setupRefreshEmpenhoButton(container);
        initializeTooltips();
        return;
    }

    const { column: sortCol, direction: sortDir } = sneEmpenhosSortState;
    const sortedEmpenhos = sortCol ? [...sneEmpenhos].sort((a, b) => {
        if (a.error && !b.error) return 1;
        if (!a.error && b.error) return -1;
        if (a.error && b.error) return 0;
        const va = getSneEmpenhoSortValue(a, sortCol);
        const vb = getSneEmpenhoSortValue(b, sortCol);
        const cmp = va < vb ? -1 : va > vb ? 1 : 0;
        return sortDir === 'asc' ? cmp : -cmp;
    }) : sneEmpenhos;

    const sortTh = (col, label, extraClass = '', extraAttrs = '') => {
        const active = sortCol === col;
        const icon = active ? (sortDir === 'asc' ? 'arrow_upward' : 'arrow_downward') : 'unfold_more';
        const cls = ['files-col-sortable', active ? 'sorted' : '', extraClass].filter(Boolean).join(' ');
        return `<th class="${cls}" ${extraAttrs} onclick="sortSneEmpenhos('${col}')">${label}<i class="material-symbols-outlined sort-icon">${icon}</i></th>`;
    };

    let html = buildFolderPathHTML(displayPath, deleteBtnEmpenhos, refreshBtn);
    html += `
        <div class="files-table-container">
            <table class="files-table sne-empenhos-table">
                <thead>
                    <tr>
                        <th class="sne-select-col text-center">
                            <input type="checkbox" id="sne-select-all" class="form-check-input">
                        </th>
                        ${sortTh('arquivo', 'Arquivo', '', 'colspan="2"')}
                        ${sortTh('sne', 'SNE')}
                        ${sortTh('af', 'AF')}
                        ${sortTh('cnpj', 'CNPJ')}
                        ${sortTh('fornecedor', 'Fornecedor')}
                        ${sortTh('tombamento', 'Tombamento', 'text-center')}
                        ${sortTh('certidoes', 'Certidões', 'text-center')}
                        ${sortTh('pasta-sne', 'Pasta SNE', 'text-center')}
                        <th></th>
                    </tr>
                </thead>
                <tbody>`;

    for (const r of sortedEmpenhos) {
        const safeFilename = r.filename.replace(/'/g, "\\'");
        const safeFilePath = (sneEmpenhosFolderPath + '\\' + r.filename).replace(/\\/g, '\\\\').replace(/"/g, '&quot;');

        if (r.error) {
            html += `
                <tr class="text-danger">
                    <td class="sne-select-col"></td>
                    <td><i class="material-symbols-outlined">error</i></td>
                    <td class="text-break" colspan="8">${r.filename}: ${r.error}</td>
                    <td class="table-btn-column text-nowrap">
                        <button class="btn btn-sm text-danger file-row-btn"
                            onclick="excluirEmpenho('${safeFilename}')"
                            data-bs-toggle="tooltip" data-bs-title="Excluir arquivo">
                            <i class="material-symbols-outlined">delete</i>
                        </button>
                    </td>
                </tr>`;
            continue;
        }

        const sneNum = r.sneNumber || '—';
        const afLabel = r.af ? `${r.af.number} / ${r.af.year}` : '—';
        const cnpjLabel = r.cnpj ? formatCnpj(r.cnpj) : '—';

        const tombCell = buildTombCell(r.tombamento);

        const statusCell = buildStatusCell(r.cnpj);

        let pastaSneCell = '<span class="text-muted">—</span>';
        if (r.af && r.sneNumber) {
            const afFolderName = `AF ${r.af.number}-${r.af.year}`;
            const sneFolderName = `SNE ${r.sneNumber}`;
            const afData = sneAfsData.find(af => af.name === afFolderName);
            const sneExists = afData?.snes?.some(s => s.name === sneFolderName);
            const safeAfFolderName = afFolderName.replace(/'/g, "\\'");
            pastaSneCell = sneExists
                ? `<button class="btn btn-sm file-row-btn sne-status-link text-success"
                        data-bs-toggle="tooltip" data-bs-title="Pasta ${sneFolderName} criada<br><em>Clique para exibir</em>"
                        onclick="event.stopPropagation(); navegarParaAF('${safeAfFolderName}')">
                        <i class="material-symbols-outlined">folder</i>
                    </button>`
                : `<i class="material-symbols-outlined text-muted" data-bs-toggle="tooltip" data-bs-title="Pasta SNE não criada">folder_off</i>`;
        }

        const safeFilenameAttr = r.filename.replace(/"/g, '&quot;');
        html += `
            <tr ondblclick="openFile('${safeFilePath}')">
                <td class="sne-select-col text-center" ondblclick="event.stopPropagation()">
                    <input type="checkbox" class="form-check-input sne-row-check" data-filename="${safeFilenameAttr}">
                </td>
                <td><i class="material-symbols-outlined text-muted">receipt_long</i></td>
                <td class="text-break">${r.filename}</td>
                <td class="text-nowrap">${sneNum}</td>
                <td class="text-nowrap">${afLabel}</td>
                <td class="text-nowrap">${cnpjLabel}</td>
                <td>${r.company || '<span class="text-muted">—</span>'}</td>
                <td class="text-center">${tombCell}</td>
                <td class="text-center">${statusCell}</td>
                <td class="text-center">${pastaSneCell}</td>
                <td class="table-btn-column text-nowrap">
                    <button class="btn btn-sm text-primary file-row-btn"
                        onclick="openFile('${safeFilePath}')"
                        data-bs-toggle="tooltip" data-bs-title="Abrir arquivo">
                        <i class="material-symbols-outlined">open_in_new</i>
                    </button>
                    <button class="btn btn-sm text-danger file-row-btn"
                        onclick="excluirEmpenho('${safeFilename}')"
                        data-bs-toggle="tooltip" data-bs-title="Excluir arquivo">
                        <i class="material-symbols-outlined">delete</i>
                    </button>
                </td>
            </tr>`;
    }

    html += `</tbody></table></div>`;
    container.innerHTML = html;

    const table = container.querySelector('.sne-empenhos-table');
    table?.querySelector('#sne-select-all')?.addEventListener('change', function () {
        table.querySelectorAll('.sne-row-check').forEach(cb => { cb.checked = this.checked; });
    });

    setupFolderPathButtons(container);
    setupRefreshEmpenhoButton(container);
    initializeTooltips();
}

function setupRefreshEmpenhoButton(container) {
    container.querySelector('.btn-refresh-empenhos')?.addEventListener('click', carregarEmpenhos);
    container.querySelector('.btn-clear-empenhos')?.addEventListener('click', (e) => {
        removeTooltip(e.currentTarget);
        excluirTodosEmpenhos();
    });
}

async function excluirEmpenho(filename) {
    const confirmed = await showConfirmationModal({
        title: 'Excluir Empenho',
        message: `Tem certeza que deseja excluir <strong>${filename}</strong>?`,
        detail: '<i class="material-symbols-outlined me-1">warning</i> Esta ação não pode ser desfeita.',
        confirmText: 'Excluir',
        confirmColor: 'btn-danger',
    });

    if (!confirmed) return;

    try {
        const response = await fetch('/api/sne/empenhos', {
            method: 'DELETE',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ filenames: [filename] }),
        });
        const data = await response.json();

        if (!response.ok && response.status !== 207) {
            showToast(`Erro ao excluir: ${data.error}`, 'error');
            return;
        }

        showToast(data.message, 'success');
        carregarEmpenhos();
    } catch (error) {
        showToast(`Erro ao excluir: ${error.message}`, 'error');
    }
}

// ====== AFs ======

let sneAfsData = [];
let sneAfsAnalysis = [];
let sneAfsFolderPath = '';
let sneAfsSortState = { column: null, direction: 'asc' };
let snePendingAfScroll = null;

async function carregarAFs() {
    const container = document.getElementById('sne-afs-container');
    if (!container) return;

    container.innerHTML = customSpinnerHTML('Analisando arquivos das AFs...');

    try {
        const resp = await fetch('/api/sne/afs/analisar');
        const data = await resp.json();

        if (data.error) {
            container.innerHTML = alertHTML('danger', 'error', data.error);
            return;
        }

        sneAfsAnalysis = data.afs || [];
        sneAfsFolderPath = data.folderPath || '';

        renderAFs();
    } catch (error) {
        container.innerHTML = alertHTML('danger', 'error', `Erro ao carregar AFs: ${error.message}`);
    }
}

function getSneAfSortValue(af, column) {
    switch (column) {
        case 'af': {
            const m = af.name.match(/^AF (\d+)-(\d+)$/);
            return m ? parseInt(m[2]) * 1e6 + parseInt(m[1]) : 0;
        }
        case 'tombamento': return af.snes.some(s => s.empenho?.tombamento) ? 1 : 0;
        case 'certidoes': {
            const order = { ok: 1, alerta: 2, erro: 3, impedido: 4 };
            let worst = 0;
            for (const sne of af.snes) {
                if (!sne.certidoes?.length) continue;
                const val = order[computeSupplierStatus(sne.certidoes)] || 0;
                if (val > worst) worst = val;
            }
            return worst;
        }
        default: return '';
    }
}

function sortAFs(column) {
    const current = sneAfsSortState;
    if (current.column === column) {
        sneAfsSortState = current.direction === 'asc'
            ? { column, direction: 'desc' }
            : { column: null, direction: 'asc' };
    } else {
        sneAfsSortState = { column, direction: 'asc' };
    }
    renderAFs();
}

function renderAFs() {
    const container = document.getElementById('sne-afs-container');
    if (!container) return;

    const displayPath = sneAfsFolderPath.replace(/\\/g, '/');
    const refreshBtn = `
        <button class="folder-path-btn btn-refresh-afs"
            data-bs-toggle="tooltip" data-bs-title="Atualizar lista">
            <i class="material-symbols-outlined">refresh</i>
        </button>`;

    const deleteBtnAfs = `
        <button class="folder-path-btn text-danger btn-clear-afs"
            data-bs-toggle="tooltip" data-bs-title="Excluir todos os arquivos da pasta"
            ${sneAfsAnalysis.length === 0 ? 'disabled' : ''}>
            <i class="material-symbols-outlined">delete</i>
        </button>`;

    if (sneAfsAnalysis.length === 0) {
        container.innerHTML = buildFolderPathHTML(displayPath, deleteBtnAfs, refreshBtn) +
            alertHTML('info', 'info', 'Nenhuma AF encontrada. Use <strong>Criar AFs</strong> na aba Empenhos para gerar a estrutura.');
        setupFolderPathButtons(container);
        setupRefreshAfsButton(container);
        initializeTooltips();
        return;
    }

    const { column: sortCol, direction: sortDir } = sneAfsSortState;
    const sortedAfs = sortCol ? [...sneAfsAnalysis].sort((a, b) => {
        const va = getSneAfSortValue(a, sortCol);
        const vb = getSneAfSortValue(b, sortCol);
        const cmp = va < vb ? -1 : va > vb ? 1 : 0;
        return sortDir === 'asc' ? cmp : -cmp;
    }) : sneAfsAnalysis;

    const sortTh = (col, label, extraClass = '', extraAttrs = '') => {
        const active = sortCol === col;
        const icon = active ? (sortDir === 'asc' ? 'arrow_upward' : 'arrow_downward') : 'unfold_more';
        const cls = ['files-col-sortable', active ? 'sorted' : '', extraClass].filter(Boolean).join(' ');
        return `<th class="${cls}" ${extraAttrs} onclick="sortAFs('${col}')">${label}<i class="material-symbols-outlined sort-icon">${icon}</i></th>`;
    };

    let html = buildFolderPathHTML(displayPath, deleteBtnAfs, refreshBtn);
    html += `
        <div class="files-table-container">
            <table class="files-table sne-afs-table">
                <thead>
                    <tr>
                        ${sortTh('af', 'AF', '', 'colspan="2"')}
                        <th></th>
                        <th colspan="2">SNE</th>
                        ${sortTh('tombamento', 'Tombamento', 'text-center')}
                        ${sortTh('certidoes', 'Certidões', 'text-center')}
                        <th>Arquivos</th>
                        <th></th>
                    </tr>
                </thead>`;

    for (const af of sortedAfs) {
        const safeAfName = af.name.replace(/'/g, "\\'");
        const safeAfPath = af.path.replace(/\\/g, '\\\\').replace(/"/g, '&quot;');
        const rowspan = Math.max(af.snes.length, 1);
        const rowspanAttr = rowspan > 1 ? ` rowspan="${rowspan}"` : '';

        const afIconCell = `<td class="sne-af-cell"${rowspanAttr}><i class="material-symbols-outlined text-primary">folder</i></td>`;
        const afNameCell = `<td class="sne-af-cell fw-semibold"${rowspanAttr}>${af.name}</td>`;
        const afActionsCell = `<td class="sne-af-cell table-btn-column text-nowrap"${rowspanAttr}>
                        <button class="btn btn-sm text-primary file-row-btn"
                            onclick="openFolder('${safeAfPath}')"
                            data-bs-toggle="tooltip" data-bs-title="Abrir pasta da AF">
                            <i class="material-symbols-outlined">folder_open</i>
                        </button>
                        <button class="btn btn-sm text-danger file-row-btn"
                            onclick="excluirAFFolder('${safeAfName}')"
                            data-bs-toggle="tooltip" data-bs-title="Excluir AF">
                            <i class="material-symbols-outlined">delete</i>
                        </button>
                    </td>`;

        html += `<tbody class="sne-af-group" data-af-name="${af.name}">`;

        if (af.snes.length === 0) {
            html += `<tr>${afIconCell}${afNameCell}${afActionsCell}
                    <td colspan="6" class="text-muted fst-italic">Sem SNEs</td>
                </tr>`;
        } else {
            for (let i = 0; i < af.snes.length; i++) {
                const sne = af.snes[i];
                const safeSneNum = sne.name.replace(/'/g, "\\'");
                const safeSnePath = sne.path.replace(/\\/g, '\\\\').replace(/"/g, '&quot;');

                const tombCell = buildTombCell(sne.empenho?.tombamento);

                const statusCell = buildAfStatusCell(sne.certidoes, af.name, sne.name, sne.empenho?.cnpj);

                const filesHtml = sne.files.length > 0
                    ? sne.files.map(f => {
                        const safeFilePath = (sne.path + '\\' + f).replace(/\\/g, '\\\\').replace(/'/g, "\\'");
                        return `<div class="sne-file-item" ondblclick="openFile('${safeFilePath}')">
                                <i class="material-symbols-outlined text-muted">description</i>
                                <span>${f}</span>
                                <button class="btn btn-sm text-primary file-row-btn"
                                    onclick="openFile('${safeFilePath}')"
                                    data-bs-toggle="tooltip" data-bs-title="Abrir arquivo">
                                    <i class="material-symbols-outlined">open_in_new</i>
                                </button>
                            </div>`;
                    }).join('')
                    : '<span class="text-muted fst-italic">Vazia</span>';

                html += `<tr>`;
                if (i === 0) html += afIconCell + afNameCell + afActionsCell;
                html += `
                    <td><i class="material-symbols-outlined text-muted">receipt_long</i></td>
                    <td class="text-nowrap">${sne.name}</td>
                    <td class="text-center">${tombCell}</td>
                    <td class="text-center">${statusCell}</td>
                    <td>${filesHtml}</td>
                    <td class="table-btn-column text-nowrap">
                        <button class="btn btn-sm text-primary file-row-btn"
                            onclick="openFolder('${safeSnePath}')"
                            data-bs-toggle="tooltip" data-bs-title="Abrir pasta da SNE">
                            <i class="material-symbols-outlined">folder_open</i>
                        </button>
                        <button class="btn btn-sm text-danger file-row-btn"
                            onclick="excluirSNEFolder('${safeAfName}', '${safeSneNum}')"
                            data-bs-toggle="tooltip" data-bs-title="Excluir SNE">
                            <i class="material-symbols-outlined">delete</i>
                        </button>
                    </td>
                </tr>`;
            }
        }

        html += `</tbody>`;
    }

    html += `</table></div>`;
    container.innerHTML = html;
    setupFolderPathButtons(container);
    setupRefreshAfsButton(container);
    initializeTooltips();

    if (snePendingAfScroll) {
        const afName = snePendingAfScroll;
        snePendingAfScroll = null;
        requestAnimationFrame(() => scrollParaAF(afName));
    }
}

function setupRefreshAfsButton(container) {
    container.querySelector('.btn-refresh-afs')?.addEventListener('click', carregarAFs);
    container.querySelector('.btn-clear-afs')?.addEventListener('click', (e) => {
        removeTooltip(e.currentTarget);
        excluirTodasAFs();
    });
}

async function excluirTodosCertidoes() {
    const confirmed = await showConfirmationModal({
        title: 'Excluir Todas as Certidões',
        message: 'Tem certeza que deseja excluir <strong>TODOS</strong> os arquivos da pasta?',
        detail: '<i class="material-symbols-outlined me-1">warning</i> Todos os arquivos serão excluídos permanentemente.',
        confirmText: 'Excluir',
        confirmColor: 'btn-danger',
    });

    if (!confirmed) return;

    try {
        const response = await fetch('/api/sne/certidoes', {
            method: 'DELETE',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ deleteAll: true }),
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

async function excluirTodosEmpenhos() {
    const confirmed = await showConfirmationModal({
        title: 'Excluir Todos os Empenhos',
        message: 'Tem certeza que deseja excluir <strong>TODOS</strong> os arquivos da pasta?',
        detail: '<i class="material-symbols-outlined me-1">warning</i> Todos os arquivos serão excluídos permanentemente.',
        confirmText: 'Excluir',
        confirmColor: 'btn-danger',
    });

    if (!confirmed) return;

    try {
        const response = await fetch('/api/sne/empenhos', {
            method: 'DELETE',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ deleteAll: true }),
        });
        const data = await response.json();

        if (!response.ok && response.status !== 207) {
            showToast(`Erro ao excluir: ${data.error}`, 'error');
            return;
        }

        showToast(data.message, 'success');
        carregarEmpenhos();
    } catch (error) {
        showToast(`Erro ao excluir: ${error.message}`, 'error');
    }
}

async function excluirTodasAFs() {
    if (sneAfsAnalysis.length === 0) return;

    const confirmed = await showConfirmationModal({
        title: 'Excluir Todas as AFs',
        message: 'Tem certeza que deseja excluir <strong>TODAS</strong> as AFs da pasta?',
        detail: `<i class="material-symbols-outlined me-1">warning</i> ${sneAfsAnalysis.length} AF(s) e todo o seu conteúdo serão excluídas permanentemente.`,
        confirmText: 'Excluir',
        confirmColor: 'btn-danger',
    });

    if (!confirmed) return;

    try {
        const response = await fetch('/api/sne/afs', {
            method: 'DELETE',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ deleteAll: true }),
        });
        const data = await response.json();

        if (!response.ok) {
            showToast(`Erro ao excluir: ${data.error}`, 'error');
            return;
        }

        showToast(data.message, 'success');
        carregarAFs();
    } catch (error) {
        showToast(`Erro ao excluir: ${error.message}`, 'error');
    }
}

async function excluirAFFolder(afName) {
    const confirmed = await showConfirmationModal({
        title: 'Excluir AF',
        message: `Tem certeza que deseja excluir a pasta <strong>${afName}</strong> e todo o seu conteúdo?`,
        detail: '<i class="material-symbols-outlined me-1">warning</i> Todas as SNEs e arquivos dentro desta AF serão excluídos permanentemente.',
        confirmText: 'Excluir',
        confirmColor: 'btn-danger',
    });

    if (!confirmed) return;

    try {
        const response = await fetch('/api/sne/afs', {
            method: 'DELETE',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ afName }),
        });
        const data = await response.json();

        if (!response.ok) {
            showToast(`Erro ao excluir: ${data.error}`, 'error');
            return;
        }

        showToast(data.message, 'success');
        carregarAFs();
    } catch (error) {
        showToast(`Erro ao excluir: ${error.message}`, 'error');
    }
}

async function excluirSNEFolder(afName, sneName) {
    const confirmed = await showConfirmationModal({
        title: 'Excluir SNE',
        message: `Tem certeza que deseja excluir a pasta <strong>${sneName}</strong>?`,
        detail: '<i class="material-symbols-outlined me-1">warning</i> Todos os arquivos dentro desta SNE serão excluídos permanentemente.',
        confirmText: 'Excluir',
        confirmColor: 'btn-danger',
    });

    if (!confirmed) return;

    try {
        const response = await fetch('/api/sne/afs', {
            method: 'DELETE',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ afName, sneName }),
        });
        const data = await response.json();

        if (!response.ok) {
            showToast(`Erro ao excluir: ${data.error}`, 'error');
            return;
        }

        showToast(data.message, 'success');
        carregarAFs();
    } catch (error) {
        showToast(`Erro ao excluir: ${error.message}`, 'error');
    }
}

async function executarCriarAFs() {
    if (sneEmpenhos.length === 0) {
        showToast('Nenhum empenho carregado. Atualize a lista primeiro.', 'warning');
        return;
    }

    const scope = document.querySelector('input[name="sne-afs-scope"]:checked')?.value;

    let filenames = null;
    let scopeDetail = '';

    if (scope === 'ok') {
        const filtered = sneEmpenhos.filter(r => !r.error && r.af && r.sneNumber && getSupplierStatusByCnpj(r.cnpj) === 'ok');
        if (filtered.length === 0) {
            showToast('Nenhuma SNE com certidões OK encontrada.', 'warning');
            return;
        }
        filenames = filtered.map(r => r.filename);
        scopeDetail = `<br><i class="material-symbols-outlined me-1">filter_list</i> Somente certidões OK: <strong>${filenames.length}</strong> de ${sneEmpenhos.length} empenho(s).`;
    } else if (scope === 'selecionadas') {
        const checked = document.querySelectorAll('.sne-row-check:checked');
        if (checked.length === 0) {
            showToast('Selecione pelo menos uma SNE.', 'warning');
            return;
        }
        filenames = Array.from(checked).map(cb => cb.dataset.filename);
        scopeDetail = `<br><i class="material-symbols-outlined me-1">checklist</i> Somente selecionadas: <strong>${filenames.length}</strong> de ${sneEmpenhos.length} empenho(s).`;
    }

    const moveSnes = document.querySelector('input[name="sne-afs-action"]:checked')?.value === 'mover';
    const actionDetail = moveSnes
        ? `<br><i class="material-symbols-outlined me-1">drive_file_move</i> Os empenhos serão <strong>movidos</strong> para as pastas das AFs.`
        : '';

    const confirmed = await showConfirmationModal({
        title: 'Criar estrutura de AFs',
        message: 'Serão criadas pastas para cada AF e SNE identificados, com as certidões de cada fornecedor copiadas para cada pasta.',
        detail: `<i class="material-symbols-outlined me-1">folder</i> As pastas serão criadas em <strong>scripts/sne/AFs/</strong>.${scopeDetail}${actionDetail}`,
        confirmText: 'Criar AFs',
        confirmColor: 'btn-primary',
    });

    if (!confirmed) return;

    try {
        const response = await fetch('/api/sne/empenhos/criar-afs', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ ...(filenames && { filenames }), moveSnes }),
        });
        const data = await response.json();

        if (!response.ok) {
            showToast(`Erro ao criar AFs: ${data.error}`, 'error');
            return;
        }

        const parts = [`${data.afs.length} AF(s)`, `${data.snes.length} SNE(s)`];
        if (data.errors?.length > 0) parts.push(`${data.errors.length} erro(s)`);
        const status = data.errors?.length > 0 ? 'warning' : 'success';
        showToast(`Estrutura criada: ${parts.join(', ')}.`, status);

        addNotification({
            message: `AFs criadas: ${parts.join(', ')}.`,
            type: status,
            source: 'SNE',
        });

        if (moveSnes) {
            await carregarEmpenhos();
        } else {
            renderEmpenhos();
        }
        carregarAFs();
    } catch (error) {
        showToast(`Erro: ${error.message}`, 'error');
    }
}
