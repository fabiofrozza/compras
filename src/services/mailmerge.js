// mailmerge.js - Função de mailmerge para atas em JavaScript
const ExcelJS = require('exceljs');
const fs = require('fs').promises;
const path = require('path');
const JSZip = require('jszip');

/**
 * Executa o processo de mailmerge para atas
 * @param {Object} params - Parâmetros da execução
 * @param {string} params.modelo_ata - Nome do arquivo do modelo de ata
 * @param {Function} logger - Função para enviar logs
 * @returns {Object} Resultado da execução
 */
async function executarMailmerge(params, logger) {
    try {
        const modeloAta = params.modelo_ata;

        if (!modeloAta || modeloAta === 'undefined' || modeloAta === '') {
            logger('Erro: Modelo de Ata não foi informado. Encerrando...', 'error');
            return {
                status: 'error',
                message: 'Modelo de Ata não foi informado',
                arquivos_gerados: [],
                erros: []
            };
        }

        logger(`📋 Modelo selecionado: ${modeloAta}`, 'info');

        // Definir caminhos
        const scriptsDir = path.join(__dirname, '..', '..', 'scripts', 'atas');
        const caminhoTemplate = path.join(scriptsDir, 'ATAS_MODELOS', modeloAta);
        const caminhoData = path.join(scriptsDir, 'dados_atas.xlsx');

        // Validar arquivo de dados
        try {
            await fs.access(caminhoData);
        } catch (e) {
            logger(`Erro: Arquivo dados_atas.xlsx não encontrado em: ${caminhoData}`, 'error');
            return {
                status: 'error',
                message: 'Arquivo dados_atas.xlsx não encontrado',
                arquivos_gerados: [],
                erros: ['Arquivo de dados não encontrado']
            };
        }

        // Validar template
        try {
            await fs.access(caminhoTemplate);
        } catch (e) {
            logger(`Erro: Template não encontrado em: ${caminhoTemplate}`, 'error');
            return {
                status: 'error',
                message: 'Template de ata não encontrado',
                arquivos_gerados: [],
                erros: ['Template não encontrado']
            };
        }

        logger(`📂 Carregando dados de: ${caminhoData}`, 'info');

        // Ler dados do Excel
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(caminhoData);
        const worksheet = workbook.getWorksheet('Dados_para_Atas');

        if (!worksheet) {
            logger('Erro: Aba "Dados_para_Atas" não encontrada no arquivo Excel', 'error');
            return {
                status: 'error',
                message: 'Aba "Dados_para_Atas" não encontrada',
                arquivos_gerados: [],
                erros: ['Aba de dados não encontrada']
            };
        }

        const headers = worksheet.getRow(1).values.filter(v => v);
        const numRegistros = worksheet.rowCount - 1;

        logger(`📊 Total de registros: ${numRegistros}`, 'info');

        // Validar formato do template
        const extensaoTemplate = path.extname(caminhoTemplate).toLowerCase();

        // Rejeitar arquivos .doc - apenas .docx é permitido
        if (extensaoTemplate === '.doc') {
            logger('Erro: Apenas arquivos .docx são permitidos. Arquivo .doc não é suportado.', 'error');
            return {
                status: 'error',
                message: 'Apenas arquivos .docx são permitidos',
                arquivos_gerados: [],
                erros: ['Template .doc não é permitido. Use .docx']
            };
        }

        if (extensaoTemplate !== '.docx') {
            logger(`Erro: Formato de arquivo não suportado: ${extensaoTemplate}. Apenas .docx é permitido.`, 'error');
            return {
                status: 'error',
                message: 'Formato de arquivo não suportado',
                arquivos_gerados: [],
                erros: [`Apenas .docx é permitido. Recebido: ${extensaoTemplate}`]
            };
        }

        logger(`📝 Tipo de documento: ${extensaoTemplate}`, 'info');

        logger('⚙️ Iniciando processamento de atas...', 'info');

        const arquivosGerados = [];
        const erros = [];

        // Ler o template uma única vez
        let templateBuffer;
        try {
            templateBuffer = await fs.readFile(caminhoTemplate);
            logger(`✓ Template carregado com sucesso (${templateBuffer.length} bytes)`, 'info');
        } catch (err) {
            logger(`Erro ao ler template: ${err.message}`, 'error');
            return {
                status: 'error',
                message: 'Erro ao ler arquivo de template',
                arquivos_gerados: [],
                erros: [`Falha ao ler: ${err.message}`]
            };
        }

        // Processar cada linha de dados
        for (let i = 2; i <= worksheet.rowCount; i++) {
            let numeroAta = i - 1; // Valor padrão caso algo falhe
            try {
                // Ler dados da linha
                const row = worksheet.getRow(i);
                const dadosLinha = {};

                headers.forEach((header, idx) => {
                    const valor = row.getCell(idx + 1).value;
                    dadosLinha[header] = valor || '';
                });

                // Usar número da ata do arquivo de dados ao invés de sequencial
                numeroAta = parseInt(dadosLinha['ata']) || (i - 1);
                logger(`📄 Processando Ata ${numeroAta} de ${numRegistros}...`, 'info');

                // Substituir dados no documento
                const documentoProcessado = await processarDocumento(
                    templateBuffer,
                    dadosLinha,
                    numeroAta,
                    extensaoTemplate,
                    logger
                );

                // Validar documento processado
                if (!documentoProcessado || documentoProcessado.length === 0) {
                    throw new Error('Documento processado está vazio');
                }

                // Criar um nome de arquivo de saída
                let razaoSocial = (dadosLinha['razao_social'] || 'Fornecedor').toString().trim();
                razaoSocial = razaoSocial.substring(0, 50);
                razaoSocial = razaoSocial.replace(/[\\/:*?"<>|]/g, '');

                const nomeArquivoSaida = `Ata ${numeroAta.toString().padStart(4, '0')} - ${razaoSocial}${extensaoTemplate}`;
                const caminhoSaida = path.join(scriptsDir, 'ATAS_FINALIZADAS', nomeArquivoSaida);

                // Salvar arquivo
                await fs.writeFile(caminhoSaida, documentoProcessado);

                // Verificar se arquivo foi criado corretamente
                const stats = await fs.stat(caminhoSaida);
                logger(`   ✓ Ata gerada: ${nomeArquivoSaida} (${stats.size} bytes)`, 'success');

                arquivosGerados.push(nomeArquivoSaida);

            } catch (err) {
                logger(`⚠️ Erro ao processar Ata ${numeroAta}: ${err.message}`, 'warning');
                erros.push(`Ata ${numeroAta} - ${err.message}`);
            }
        }

        // Resultado final
        logger('', 'info');
        logger(`✅ ${arquivosGerados.length} ata(s) gerada(s) com sucesso!`, 'success');

        if (erros.length > 0) {
            logger(`⚠️ ${erros.length} erro(s) encontrado(s):`, 'warning');
            erros.forEach(erro => {
                logger(`   • ${erro}`, 'warning');
            });
        }

        return {
            status: erros.length === 0 ? 'success' : 'warning',
            message: `${arquivosGerados.length} ata(s) gerada(s)`,
            arquivos_gerados: arquivosGerados,
            erros: erros
        };

    } catch (err) {
        logger(`Erro geral: ${err.message}`, 'error');
        return {
            status: 'error',
            message: err.message,
            arquivos_gerados: [],
            erros: [err.message]
        };
    }
}

/**
 * Processa documento DOCX substituindo placeholders
 */
async function processarDocumento(templateBuffer, dados, numeroAta, extensao, logger) {
    try {
        if (extensao === '.docx') {
            return await processarDocumento_DOCX(templateBuffer, dados, numeroAta);
        } else {
            throw new Error(`Formato de arquivo não suportado: ${extensao}`);
        }
    } catch (err) {
        throw new Error(`Erro ao processar documento: ${err.message}`);
    }
}

/**
 * Processa arquivo DOCX (ZIP com XMLs)
 */
async function processarDocumento_DOCX(templateBuffer, dados, numeroAta) {
    try {
        const zip = new JSZip();
        const carregado = await zip.loadAsync(templateBuffer);

        // Encontrar arquivos XML que contêm o texto
        const arquivosProcessar = [];

        for (const nomeArquivo of Object.keys(carregado.files)) {
            if (nomeArquivo.includes('word/document') ||
                (nomeArquivo.includes('word/') && nomeArquivo.endsWith('.xml')) ||
                nomeArquivo === 'document.xml' ||
                nomeArquivo === 'content.xml') {
                arquivosProcessar.push(nomeArquivo);
            }
        }

        if (arquivosProcessar.length === 0) {
            for (const nomeArquivo of Object.keys(carregado.files)) {
                if (nomeArquivo.endsWith('.xml') && !nomeArquivo.includes('_rels')) {
                    arquivosProcessar.push(nomeArquivo);
                }
            }
        }

        if (arquivosProcessar.length === 0) {
            throw new Error('Nenhum arquivo XML de documento encontrado');
        }

        // Processar cada arquivo XML
        for (const nomeArquivo of arquivosProcessar) {
            const conteudoAtual = await carregado.file(nomeArquivo).async('text');
            const conteudoProcessado = substituirCampos(conteudoAtual, dados, numeroAta);
            carregado.file(nomeArquivo, conteudoProcessado);
        }

        // Gerar buffer do documento processado
        const docBuffer = await carregado.generateAsync({ type: 'nodebuffer' });
        return docBuffer;

    } catch (err) {
        throw new Error(`Erro ao processar DOCX: ${err.message}`);
    }
}

/**
 * Substitui placeholders pelos dados fornecidos em XML (para DOCX)
 */
function substituirCampos(conteudo, dados, numeroAta) {
    let resultado = conteudo;

    // Cria um objeto com todas as chaves em minúsculas (para facilitar a comparação)
    const dadosLower = {};
    Object.keys(dados).forEach(key => {
        dadosLower[key.toLowerCase()] = dados[key];
    });

    // Adiciona o NUMERO_ATA (também em minúsculas)
    dadosLower['numero_ata'] = String(numeroAta);

    // Lista de padrões de placeholder (incluindo «...»)
    const padroes = [
        /«([a-zA-Z0-9_]+)»/g,          // «campo»
        /<<\{([a-zA-Z0-9_]+)\}>>/g,    // <<{campo}>>
        /\[([a-zA-Z0-9_]+)\]/g,         // [campo]
        /\{([a-zA-Z0-9_]+)\}/g,         // {campo}
        /<<([a-zA-Z0-9_]+)>>/g          // <<campo>>
    ];

    // Aplica cada padrão
    padroes.forEach(padrao => {
        resultado = resultado.replace(padrao, (match, nomeCampo) => {
            const nomeLower = nomeCampo.toLowerCase();
            if (dadosLower.hasOwnProperty(nomeLower)) {
                return String(dadosLower[nomeLower]).trim();
            }
            return match; // mantém o placeholder se não encontrar
        });
    });

    return resultado;
}
module.exports = { executarMailmerge };
