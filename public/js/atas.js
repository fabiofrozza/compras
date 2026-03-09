const atasData = {
    arquivos: [],
    ataAtual: null,
};

// Flag para evitar re-inicialização e duplicação de listeners
let atasInicializado = false;

function popularAnosSelector() {
    const anoAtual = new Date().getFullYear() + 1; // +1 para incluir pregões do próximo ano
    const selectorAnoPregao = document.getElementById('atas-ano-pregao');
    const selectorAnoAta = document.getElementById('atas-ano-ata');

    const popularSelector = (selector) => {
        if (!selector) return;
        const valorAtual = selector.value;
        while (selector.options.length > 1) selector.remove(1);
        for (let i = 0; i < 10; i++) {
            const ano = anoAtual - i;
            const option = document.createElement('option');
            option.value = ano;
            option.textContent = ano;
            selector.appendChild(option);
        }
        if (valorAtual) selector.value = valorAtual;
    };

    popularSelector(selectorAnoPregao);
    popularSelector(selectorAnoAta);
}

function numerarAtas() {
    const container = document.getElementById('atas-relatorios-sicaf');
    const numeroPrimeiraAta = parseInt(document.getElementById('atas-primeira-ata').value);

    if (!container) return;

    const linhas = container.querySelectorAll('tr');

    linhas.forEach((linha, index) => {
        // Tenta localizar a célula de numeração pela classe para não duplicar ou sobrescrever
        let cellNumero = linha.querySelector('.coluna-numero-ata');

        // Se não encontrar, insere uma nova célula na posição 1 (entre ícone e nome)
        if (!cellNumero) {
            cellNumero = linha.insertCell(1);
            cellNumero.className = 'coluna-numero-ata';
            cellNumero.style.verticalAlign = "middle";
            cellNumero.style.paddingRight = "10px";
        }

        if (isNaN(numeroPrimeiraAta) || numeroPrimeiraAta < 1) {
            cellNumero.textContent = "Ata ----";
        } else {
            const numeroAta = numeroPrimeiraAta + index;
            cellNumero.textContent = "Ata " + String(numeroAta).padStart(4, '0');
        }
    });

    if (!isNaN(numeroPrimeiraAta) && numeroPrimeiraAta >= 1) {
        atasData.ataAtual = numeroPrimeiraAta;
    }
}

async function verificarStatusDadosAtas() {
    const statusContainer = document.getElementById('status-dados-atas');
    const statusText = document.getElementById('texto-status-dados');
    const statusIcon = document.getElementById('icon-status-dados');

    if (!statusContainer) return;

    try {
        const response = await fetch('/api/check-atas-data');
        const data = await response.json();

        if (data.exists) {
            atasData.dadosDisponiveis = true;
            const dataModificacao = new Date(data.modified).toLocaleString('pt-BR');

            statusContainer.classList.remove('alert-warning', 'alert-light', 'border-warning', 'border');
            statusContainer.classList.add('alert-success', 'border-success');
            statusIcon.classList.remove('text-warning', 'text-muted');
            statusIcon.classList.add('text-success');
            statusIcon.innerHTML = '<i class="material-symbols-outlined">check_circle</i>';
            statusText.innerHTML = `Dados obtidos em ${dataModificacao}`;
        } else {
            atasData.dadosDisponiveis = false;

            statusContainer.classList.remove('alert-success', 'border-success', 'alert-light', 'border');
            statusContainer.classList.add('alert-warning', 'border-warning');
            statusIcon.classList.remove('text-success', 'text-muted');
            statusIcon.classList.add('text-warning');
            statusIcon.innerHTML = '<i class="material-symbols-outlined">warning</i>';
            statusText.innerHTML = 'Arquivo de dados não encontrado. Execute "Obter dados dos SICAF" primeiro.';
        }
    } catch (error) {
        console.error('Erro ao verificar status dos dados:', error);
        atasData.dadosDisponiveis = false;
        statusText.textContent = 'Erro ao verificar status dos dados.';
    }

    atualizarBotaoGerarAtas();
}

function atualizarBotaoGerarAtas() {
    const statusContainer = document.getElementById('status-dados-atas');
    const btnGerar = document.getElementById('btn-form-atas-modelos');
    if (!btnGerar || !statusContainer) return;

    const modeloSelecionado = selectedFiles['atas-modelos'];

    if (atasData.dadosDisponiveis && modeloSelecionado) {
        btnGerar.disabled = false;
    } else {
        btnGerar.disabled = true;
    }
}

function setupAtasListeners() {
    const inputPrimeiraAta = document.getElementById('atas-primeira-ata');
    const inputProcessoSPA = document.getElementById('atas-processo-spa');

    if (inputPrimeiraAta) {
        inputPrimeiraAta.addEventListener('change', numerarAtas);
        inputPrimeiraAta.addEventListener('input', numerarAtas);
    }

    // Listener para máscara de processo SPA
    if (inputProcessoSPA) {
        inputProcessoSPA.addEventListener('input', (e) => {
            let value = e.target.value.replace(/\D/g, ''); // Remove tudo que não é dígito
            if (value.length > 0) {
                // Aplicar máscara: 23080.XXXXXX/XXXX-XX
                if (value.length <= 5) {
                    e.target.value = value;
                } else if (value.length <= 11) {
                    e.target.value = value.substring(0, 5) + '.' + value.substring(5);
                } else if (value.length <= 15) {
                    e.target.value = value.substring(0, 5) + '.' + value.substring(5, 11) + '/' + value.substring(11);
                } else {
                    e.target.value = value.substring(0, 5) + '.' + value.substring(5, 11) + '/' + value.substring(11, 15) + '-' + value.substring(15, 17);
                }
            }
        });
    }

    // Listener global para seleção de arquivos (disparado pelo file_system.js)
    document.addEventListener('file-selected', (e) => {
        if (e.detail && e.detail.containerId === 'atas-modelos') {
            atualizarBotaoGerarAtas();
        }
    });
}

function executarMailmergeAtas() {
    const btnGerar = document.getElementById('btn-form-atas-modelos');

    if (btnGerar && btnGerar.disabled) {
        return;
    }

    const modeloSelecionado = selectedFiles['atas-modelos'];

    if (!atasData.dadosDisponiveis) {
        showToast('É necessário obter os dados do SICAF antes de gerar as Atas.', 'error', 10000, 'atas');
        return;
    }

    if (!modeloSelecionado) {
        showToast('Selecione um modelo antes de gerar as Atas.', 'warning', 10000, 'atas');
        return;
    }

    const dados = {
        modelo_ata: modeloSelecionado
    };

    runRScript('atas_mailmerge', dados);
}

function inicializarAtas() {
    if (atasInicializado) return;
    popularAnosSelector();
    verificarStatusDadosAtas();
    setupAtasListeners();
    atasInicializado = true;
}
