// atas.js - Lógica específica para a aba Atas

// Objeto para armazenar dados da aba Atas
const atasData = {
    arquivos: [],
    ataAtual: null,
};

// Flag para evitar re-inicialização e duplicação de listeners
let atasInicializado = false;

/**
 * Converte data de "dd de mês de yyyy" para "YYYY-MM-DD"
 */
function converterDataBrasilParaISO(dataBrasil) {
    if (!dataBrasil) return '';

    // Padrão: "17 de fevereiro de 2026"
    const meses = {
        'janeiro': '01', 'fevereiro': '02', 'março': '03', 'abril': '04',
        'maio': '05', 'junho': '06', 'julho': '07', 'agosto': '08',
        'setembro': '09', 'outubro': '10', 'novembro': '11', 'dezembro': '12'
    };

    const partes = dataBrasil.trim().split(' de ');
    if (partes.length !== 3) return '';

    const dia = String(parseInt(partes[0])).padStart(2, '0');
    const mesNome = partes[1].toLowerCase();
    const ano = partes[2];
    const mes = meses[mesNome];

    if (!mes) return '';

    return `${ano}-${mes}-${dia}`;
}

/**
 * Formata data de YYYY-MM-DD para "dd de mês de yyyy"
 */
function formatarDataBrasil(dataISO) {
    if (!dataISO) return '';

    const data = new Date(dataISO + 'T00:00:00');
    const dia = String(data.getDate()).padStart(2, '0');
    const meses = [
        'janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho',
        'julho', 'agosto', 'setembro', 'outubro', 'novembro', 'dezembro'
    ];
    const mes = meses[data.getMonth()];
    const ano = data.getFullYear();

    return `${parseInt(dia)} de ${mes} de ${ano}`;
}

/**
 * Popula os selectors de ano com os últimos 10 anos
 * Preserva os valores existentes carregados da configuração
 */
function popularAnosSelector() {
    const anoAtual = new Date().getFullYear() + 1; // Considerar o próximo ano para pregões futuros
    const selectorAnoPregao = document.getElementById('atas-ano-pregao');
    const selectorAnoAta = document.getElementById('atas-ano-ata');

    if (selectorAnoPregao) {
        // Salvar o valor atual antes de limpar opções
        const valorAtualPregao = selectorAnoPregao.value;

        // Limpar opções existentes (menos a opção vazia)
        while (selectorAnoPregao.options.length > 1) {
            selectorAnoPregao.remove(1);
        }

        for (let i = 0; i < 10; i++) {
            const ano = anoAtual - i;
            const option = document.createElement('option');
            option.value = ano;
            option.textContent = ano;
            selectorAnoPregao.appendChild(option);
        }

        // Restaurar o valor anterior se estiver disponível nas novas opções
        if (valorAtualPregao) {
            selectorAnoPregao.value = valorAtualPregao;
        }
    }

    if (selectorAnoAta) {
        // Salvar o valor atual antes de limpar opções
        const valorAtualAta = selectorAnoAta.value;

        // Limpar opções existentes (menos a opção vazia)
        while (selectorAnoAta.options.length > 1) {
            selectorAnoAta.remove(1);
        }

        for (let i = 0; i < 10; i++) {
            const ano = anoAtual - i;
            const option = document.createElement('option');
            option.value = ano;
            option.textContent = ano;
            selectorAnoAta.appendChild(option);
        }

        // Restaurar o valor anterior se estiver disponível nas novas opções
        if (valorAtualAta) {
            selectorAnoAta.value = valorAtualAta;
        }
    }
}

/**
 * Carrega lista de PDFs da pasta SICAF e cria tabela dinâmica com numeração
 */
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

/**
 * Verifica se o arquivo dados_atas.xlsx existe e atualiza a interface
 */
async function verificarStatusDadosAtas() {
    const statusContainer = document.getElementById('status-dados-atas');
    const statusText = document.getElementById('texto-status-dados');
    const statusIcon = document.getElementById('icon-status-dados');

    if (!statusContainer) return;

    try {
        const response = await fetch('/api/check-atas-data');
        const data = await response.json();

        if (data.exists) {
            statusContainer.dataset.dadosDisponiveis = 'true';
            atasData.dadosDisponiveis = true;
            const dataModificacao = new Date(data.modified).toLocaleString('pt-BR');

            statusContainer.classList.remove('alert-warning', 'alert-light', 'border-warning', 'border');
            statusContainer.classList.add('alert-success', 'border-success');
            statusIcon.classList.remove('text-warning', 'text-muted');
            statusIcon.classList.add('text-success');
            statusIcon.innerHTML = '<i class="fas fa-check-circle"></i>';
            statusText.innerHTML = `Dados obtidos em ${dataModificacao}`;
        } else {
            statusContainer.dataset.dadosDisponiveis = 'false';
            atasData.dadosDisponiveis = false;

            statusContainer.classList.remove('alert-success', 'border-success', 'alert-light', 'border');
            statusContainer.classList.add('alert-warning', 'border-warning');
            statusIcon.classList.remove('text-success', 'text-muted');
            statusIcon.classList.add('text-warning');
            statusIcon.innerHTML = '<i class="fas fa-exclamation-triangle"></i>';
            statusText.innerHTML = 'Arquivo de dados não encontrado. Execute "Obter dados dos SICAF" primeiro.';
        }
    } catch (error) {
        console.error('Erro ao verificar status dos dados:', error);
        statusContainer.dataset.dadosDisponiveis = 'false';
        atasData.dadosDisponiveis = false;
        statusText.textContent = 'Erro ao verificar status dos dados.';
    }

    atualizarBotaoGerarAtas();
}

/**
 * Habilita ou desabilita o botão de gerar atas baseado nos requisitos
 */
function atualizarBotaoGerarAtas() {
    const statusContainer = document.getElementById('status-dados-atas');
    const btnGerar = document.getElementById('btn-form-atas-modelos');
    if (!btnGerar || !statusContainer) return;

    const modeloSelecionado = selectedFiles['atas-modelos'];
    const dadosOk = statusContainer.dataset.dadosDisponiveis === 'true';

    if (dadosOk && modeloSelecionado) {
        btnGerar.disabled = false;
    } else {
        btnGerar.disabled = true;
    }
}

/**
 * Setup de listeners para campos da aba Atas
 */
function setupAtasListeners() {
    const inputPrimeiraAta = document.getElementById('atas-primeira-ata');
    const inputDataPregao = document.getElementById('atas-data-pregao');
    const inputProcessoSPA = document.getElementById('atas-processo-spa');

    // Listener para quando n_ata muda - atualiza numeração da lista
    if (inputPrimeiraAta) {
        inputPrimeiraAta.addEventListener('change', numerarAtas);
        inputPrimeiraAta.addEventListener('input', numerarAtas);
    }

    // Listener para data - converte para formato Brasil ao sair do campo
    if (inputDataPregao) {
        inputDataPregao.addEventListener('change', () => {
            if (inputDataPregao.value) {
                // Armazenar a data em formato ISO para salvar
                // Mas podemos validar aqui se necessário
            }
        });
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

/**
 * Processa dados da ata para envio ao servidor
 */
function processarDadosAtas() {
    const form = document.getElementById('form-info-pregao');
    const dados = {};

    if (!form) {
        console.error('Formulário form-info-pregao não encontrado');
        return dados;
    }

    const campos = form.querySelectorAll('[data-field]');

    campos.forEach(campo => {
        const fieldName = campo.dataset.field;
        if (fieldName) {
            // Se for campo de data, manter no formato ISO para o servidor (não converter)
            // O servidor ou script R pode fazer a conversão se necessário
            dados[fieldName] = campo.value;
        }
    });

    return dados;
}

/**
 * Executa o mailmerge para gerar as Atas finalizadas
 */
function executarMailmergeAtas() {
    const btnGerar = document.getElementById('btn-form-atas-modelos');

    // Validação primeiro: se o botão está desabilitado, não fazer nada
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

    console.log('Iniciando mailmerge com modelo:', modeloSelecionado);

    // Preparar dados para enviar
    const dados = {
        modelo_ata: modeloSelecionado
    };

    runRScript('atas_mailmerge', dados);
}

/**
 * Inicializa a aba Atas
 */
function inicializarAtas() {
    // Se já foi inicializado, não faz nada para evitar duplicar listeners
    if (atasInicializado) return;

    // Popular selects de ano (necessário pois o HTML acabou de ser injetado)
    popularAnosSelector();

    // Verificar status dos dados
    verificarStatusDadosAtas();

    // Setup de listeners
    setupAtasListeners();

    atasInicializado = true;
}
