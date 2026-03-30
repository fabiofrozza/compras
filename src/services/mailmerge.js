// mailmerge.js - Função de mailmerge para atas em JavaScript
const ExcelJS = require('exceljs');
const XLSX = require('xlsx');
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

        // Verificar se existem arquivos .xls de tabela na pasta SICAF
        const pastaSicaf = path.join(scriptsDir, 'SICAF');
        let tabelasDisponiveis = false;
        try {
            const arquivosSicaf = await fs.readdir(pastaSicaf);
            const xlsFiles = arquivosSicaf.filter(f => f.toLowerCase().endsWith('.xls') || f.toLowerCase().endsWith('.xlsx'));
            tabelasDisponiveis = xlsFiles.length > 0;
            if (tabelasDisponiveis) {
                logger(`📊 ${xlsFiles.length} arquivo(s) de tabela (.xls) encontrado(s) na pasta SICAF`, 'info');
            }
        } catch (e) {
            // Pasta SICAF não encontrada ou inacessível - segue sem tabelas
        }

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

                    if (header === 'data' && dadosLinha[header]) {
                        dadosLinha[header] = formatarDataBrasil(dadosLinha[header]);
                    }
                });

                // Usar número da ata do arquivo de dados ao invés de sequencial
                numeroAta = parseInt(dadosLinha['ata']) || (i - 1);
                logger(`📄 Processando Ata ${i - 1} de ${numRegistros}...`, 'info');

                // Buscar arquivo .xls de tabela correspondente na pasta SICAF
                let dadosTabela = null;
                if (tabelasDisponiveis) {
                    // Procurar pelo mesmo nome do PDF (número da ata)
                    const possiveisNomes = [
                        `${numeroAta}.xls`,
                        `${numeroAta}.xlsx`,
                    ];

                    for (const nomeXls of possiveisNomes) {
                        const caminhoXls = path.join(pastaSicaf, nomeXls);
                        try {
                            await fs.access(caminhoXls);
                            dadosTabela = lerTabelaXls(caminhoXls);

                            // Verificação cruzada de CNPJ
                            const cnpjDados = String(dadosLinha['cnpj'] || '').trim();
                            const cnpjTabela = String(dadosTabela.cnpj || '').trim();
                            if (cnpjDados && cnpjTabela && cnpjDados !== cnpjTabela) {
                                logger(`   ⚠️ CNPJ divergente no arquivo ${nomeXls}: esperado ${cnpjDados}, encontrado ${cnpjTabela}`, 'warning');
                            }

                            // Verificação cruzada de número da ata
                            const ataTabela = String(dadosTabela.ata || '').trim();
                            if (ataTabela && String(numeroAta) !== ataTabela) {
                                logger(`   ⚠️ Nº Ata divergente no arquivo ${nomeXls}: esperado ${numeroAta}, encontrado ${ataTabela}`, 'warning');
                            }

                            logger(`   📊 Tabela carregada: ${dadosTabela.itens.length} item(ns) de ${nomeXls}`, 'info');
                            break;
                        } catch (e) {
                            // Arquivo não encontrado, tentar próximo nome
                        }
                    }
                }

                // Substituir dados no documento
                const documentoProcessado = await processarDocumento(
                    templateBuffer,
                    dadosLinha,
                    numeroAta,
                    extensaoTemplate,
                    logger,
                    dadosTabela
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
async function processarDocumento(templateBuffer, dados, numeroAta, extensao, logger, dadosTabela) {
    try {
        if (extensao === '.docx') {
            return await processarDocumento_DOCX(templateBuffer, dados, numeroAta, dadosTabela);
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
async function processarDocumento_DOCX(templateBuffer, dados, numeroAta, dadosTabela) {
    try {
        const zip = new JSZip();
        const carregado = await zip.loadAsync(templateBuffer);

        // Processar todos os XMLs do ZIP exceto arquivos de relacionamento
        const arquivosProcessar = Object.keys(carregado.files).filter(
            nome => nome.endsWith('.xml') && !nome.includes('_rels')
        );

        if (arquivosProcessar.length === 0) {
            throw new Error('Nenhum arquivo XML de documento encontrado');
        }

        // Gerar XML da tabela se houver dados
        const tabelaXml = dadosTabela && dadosTabela.itens.length > 0
            ? gerarTabelaOoxml(dadosTabela)
            : null;

        // Processar cada arquivo XML
        for (const nomeArquivo of arquivosProcessar) {
            let conteudoAtual = await carregado.file(nomeArquivo).async('text');
            let conteudoNormalizado = normalizarPlaceholdersSplit(conteudoAtual);
            let conteudoProcessado = substituirCampos(conteudoNormalizado, dados, numeroAta);

            // Inserir tabela no lugar do placeholder TABELA_ITENS
            if (tabelaXml && conteudoProcessado.includes('TABELA_ITENS')) {
                conteudoProcessado = inserirTabelaNoXml(conteudoProcessado, tabelaXml);
            }

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
 * Normaliza placeholders divididos pelo Word em múltiplos runs XML.
 * O Word às vezes divide <<campo>> em runs separados, ex:
 *   &lt;&lt;</w:t></w:r><w:r><w:t>campo</w:t></w:r><w:r><w:t>&gt;&gt;
 * Esta função reune o placeholder num único token antes da substituição.
 */
function normalizarPlaceholdersSplit(xml) {
    let resultado = xml;

    // Normaliza &lt;&lt;...&gt;&gt; (<<campo>> com escape XML, possivelmente dividido por runs)
    resultado = resultado.replace(
        /&lt;&lt;((?:(?!&lt;&lt;|&gt;&gt;|<\/w:p>)[\s\S]){0,2000})&gt;&gt;/g,
        (match, inner) => {
            const campo = inner.replace(/<[^>]*>/g, '').replace(/\s/g, '');
            if (/^[a-zA-Z0-9_]+$/.test(campo)) {
                return `&lt;&lt;${campo}&gt;&gt;`;
            }
            return match;
        }
    );

    // Normaliza <<...>> (literal, possivelmente dividido por runs)
    resultado = resultado.replace(
        /<<((?:(?!<<|>>|<\/w:p>)[\s\S]){0,2000})>>/g,
        (match, inner) => {
            const campo = inner.replace(/<[^>]*>/g, '').replace(/\s/g, '');
            if (/^[a-zA-Z0-9_]+$/.test(campo)) {
                return `<<${campo}>>`;
            }
            return match;
        }
    );

    return resultado;
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
        /«([a-zA-Z0-9_]+)»/g,                                  // «campo»
        /&lt;&lt;\{([a-zA-Z0-9_]+)\}&gt;&gt;/g,                // <<{campo}>> (XML escaped)
        /<<\{([a-zA-Z0-9_]+)\}>>/g,                            // <<{campo}>>
        /\[([a-zA-Z0-9_]+)\]/g,                                // [campo]
        /\{([a-zA-Z0-9_]+)\}/g,                                // {campo}
        /&lt;&lt;([a-zA-Z0-9_]+)&gt;&gt;/g,                    // <<campo>> (XML escaped)
        /<<([a-zA-Z0-9_]+)>>/g                                  // <<campo>>
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

/**
 * Formata data de Date object (UTC) ou string "YYYY-MM-DD" para "dd de mês de yyyy"
 */
function formatarDataBrasil(valorData) {
    if (!valorData) return '';

    let dia, mesIndex, ano;

    if (valorData instanceof Date) {
        // ExcelJS reads dates as UTC midnight
        dia = valorData.getUTCDate();
        mesIndex = valorData.getUTCMonth();
        ano = valorData.getUTCFullYear();
    } else if (typeof valorData === 'string') {
        const partes = valorData.substring(0, 10).split('-');
        if (partes.length === 3) {
            ano = parseInt(partes[0]);
            mesIndex = parseInt(partes[1]) - 1;
            dia = parseInt(partes[2]);
        } else {
            return valorData;
        }
    } else {
        return valorData;
    }

    if (isNaN(dia) || isNaN(ano)) return valorData;

    const meses = [
        'janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho',
        'julho', 'agosto', 'setembro', 'outubro', 'novembro', 'dezembro'
    ];

    return `${dia} de ${meses[mesIndex]} de ${ano}`;
}

/**
 * Lê os dados de itens de um arquivo .xls de tabela do fornecedor
 * @param {string} caminhoXls - Caminho para o arquivo .xls
 * @returns {Object} { cnpj, ata, itens: [{item, descricao, unidade, qtde, valor, total}], valorTotal }
 */
function lerTabelaXls(caminhoXls) {
    const wb = XLSX.readFile(caminhoXls);
    const ws = wb.Sheets[wb.SheetNames[0]];
    const dados = XLSX.utils.sheet_to_json(ws, { header: 1 });

    // Extrair CNPJ e nº da ata do cabeçalho
    let cnpj = '';
    let ata = '';
    let valorTotal = 0;

    for (const row of dados) {
        if (!row || row.length === 0) continue;
        const primeira = String(row[0] || '').trim();
        if (primeira.startsWith('CNPJ:')) {
            cnpj = String(row[3] || '').trim();
        } else if (primeira === 'Nº ata:' || primeira.startsWith('N\u00ba ata')) {
            ata = String(row[3] || '').trim();
        } else if (primeira === 'Valor total:') {
            valorTotal = row[3] || 0;
        }
    }

    // Encontrar a linha de cabeçalho da tabela de itens
    let headerIdx = -1;
    for (let i = 0; i < dados.length; i++) {
        const row = dados[i];
        if (!row) continue;
        const primeira = String(row[0] || '').trim();
        if (primeira === 'Grupo/Item' || primeira === 'Item') {
            headerIdx = i;
            break;
        }
    }

    const itens = [];
    if (headerIdx >= 0) {
        // Ler linhas de dados (após o cabeçalho, até "Valor total" ou fim)
        for (let i = headerIdx + 1; i < dados.length; i++) {
            const row = dados[i];
            if (!row || row.length === 0) continue;
            const primeira = String(row[0] || '').trim();
            if (primeira === 'Valor total') {
                // Linha de total da tabela - captura o valor total se não veio do cabeçalho
                if (!valorTotal && row[8]) valorTotal = row[8];
                break;
            }
            if (primeira === '') continue;

            // Colunas: [0]Grupo/Item, [1]_, [2]Descrição, [3]_, [4]Unid.medida, [5]Qtde, [6]Valor, [7]_, [8]Total
            const item = String(row[0] || '');
            const descricao = String(row[2] || '');
            const unidade = String(row[4] || '');
            const qtde = row[5] || 0;
            const valor = row[6] || 0;
            const total = row[8] || 0;

            itens.push({ item, descricao, unidade, qtde, valor, total });
        }
    }

    return { cnpj, ata, itens, valorTotal };
}

/**
 * Formata número como moeda brasileira (1.234,5600)
 */
function formatarMoeda(valor) {
    if (typeof valor !== 'number') valor = parseFloat(valor) || 0;
    return valor.toLocaleString('pt-BR', {
        minimumFractionDigits: 4,
        maximumFractionDigits: 4
    });
}

/**
 * Escapa caracteres especiais para XML
 */
function escaparXml(texto) {
    return String(texto)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

/**
 * Gera uma célula OOXML (<w:tc>) com texto
 */
function gerarCelula(texto, negrito, spanCols, preenchimento) {
    const fill = preenchimento || 'FFFFFF';
    let tcPr = '<w:tcPr><w:tcW w:w="0" w:type="auto"/>';
    if (spanCols) {
        tcPr += `<w:gridSpan w:val="${spanCols}"/>`;
    }
    tcPr += '<w:tcBorders>'
        + '<w:top w:val="single" w:sz="4" w:space="0" w:color="010000"/>'
        + '<w:left w:val="single" w:sz="4" w:space="0" w:color="010000"/>'
        + '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="010000"/>'
        + '<w:right w:val="single" w:sz="4" w:space="0" w:color="010000"/>'
        + '</w:tcBorders>';
    tcPr += `<w:shd w:val="clear" w:color="000000" w:fill="${fill}"/>`;
    tcPr += '</w:tcPr>';

    const rPr = '<w:rPr>'
        + '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>'
        + (negrito ? '<w:b/><w:bCs/>' : '')
        + '<w:color w:val="010000"/>'
        + '<w:sz w:val="22"/><w:szCs w:val="22"/>'
        + '</w:rPr>';

    // Tratar quebras de linha na descrição
    const linhas = String(texto).split('\n');
    let runs = '';
    linhas.forEach((linha, idx) => {
        if (idx > 0) {
            runs += `<w:r>${rPr}<w:br/></w:r>`;
        }
        runs += `<w:r>${rPr}<w:t xml:space="preserve">${escaparXml(linha)}</w:t></w:r>`;
    });

    return `<w:tc>${tcPr}<w:p><w:pPr><w:rPr>${rPr.replace('<w:rPr>', '').replace('</w:rPr>', '')}</w:rPr></w:pPr>${runs}</w:p></w:tc>`;
}

/**
 * Gera o XML OOXML completo de uma tabela de itens
 */
function gerarTabelaOoxml(dadosTabela) {
    const { itens } = dadosTabela;

    let xml = '<w:tbl>';

    // Propriedades da tabela
    xml += '<w:tblPr>'
        + '<w:tblW w:w="0" w:type="auto"/>'
        + '<w:tblInd w:w="75" w:type="dxa"/>'
        + '<w:tblCellMar><w:left w:w="70" w:type="dxa"/><w:right w:w="70" w:type="dxa"/></w:tblCellMar>'
        + '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>'
        + '</w:tblPr>';

    // Grid (larguras das colunas)
    xml += '<w:tblGrid>'
        + '<w:gridCol w:w="587"/>'
        + '<w:gridCol w:w="6222"/>'
        + '<w:gridCol w:w="943"/>'
        + '<w:gridCol w:w="655"/>'
        + '<w:gridCol w:w="864"/>'
        + '<w:gridCol w:w="1260"/>'
        + '</w:tblGrid>';

    // Linha de cabeçalho (cinza)
    xml += '<w:tr><w:trPr><w:trHeight w:val="435"/></w:trPr>';
    xml += gerarCelula('Item', true, null, 'CCCCCC');
    xml += gerarCelula('Descri\u00e7\u00e3o', true, null, 'CCCCCC');
    xml += gerarCelula('Unid. medida', true, null, 'CCCCCC');
    xml += gerarCelula('Qtde.', true, null, 'CCCCCC');
    xml += gerarCelula('Valor', true, null, 'CCCCCC');
    xml += gerarCelula('Total', true, null, 'CCCCCC');
    xml += '</w:tr>';

    // Linhas de itens
    for (const item of itens) {
        const numItem = String(item.item).padStart(4, '0');
        xml += '<w:tr><w:trPr><w:trHeight w:val="416"/></w:trPr>';
        xml += gerarCelula(numItem, false);
        xml += gerarCelula(item.descricao, false);
        xml += gerarCelula(item.unidade, false);
        xml += gerarCelula(String(item.qtde), false);
        xml += gerarCelula(formatarMoeda(item.valor), false);
        xml += gerarCelula(formatarMoeda(item.total), false);
        xml += '</w:tr>';
    }

    // Linha de valor total - soma calculada dos itens (o .xls traz o total do pregão, não da ata)
    const totalCalculado = itens.reduce((soma, item) => soma + (typeof item.total === 'number' ? item.total : parseFloat(item.total) || 0), 0);
    const totalFormatado = formatarMoeda(totalCalculado);
    xml += '<w:tr><w:trPr><w:trHeight w:val="218"/></w:trPr>';
    xml += gerarCelula('Valor total', true, 5);
    xml += gerarCelula(totalFormatado, true);
    xml += '</w:tr>';

    xml += '</w:tbl>';
    return xml;
}

/**
 * Substitui o parágrafo que contém o placeholder TABELA_ITENS pelo XML da tabela
 */
function inserirTabelaNoXml(conteudoXml, tabelaXml) {
    // Encontrar o <w:p> que contém "TABELA_ITENS" (qualquer formato de placeholder)
    const regex = /<w:p\b[^>]*>(?:(?!<w:p\b).)*?TABELA_ITENS(?:(?!<w:p\b).)*?<\/w:p>/gs;
    return conteudoXml.replace(regex, tabelaXml);
}

module.exports = { executarMailmerge };
