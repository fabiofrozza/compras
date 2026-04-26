const path = require('path');

// Lê as abas "Licitação YYYY" da Planilha de Controle (Google Sheets) e
// retorna os processos com as colunas Observatório - Licitação/Execução
// preenchidas conforme regras de negócio definidas em powerbi.html.

function parseCsv(text) {
  const rows = [];
  let row = [];
  let field = '';
  let inQuotes = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (inQuotes) {
      if (ch === '"') {
        if (text[i + 1] === '"') { field += '"'; i++; }
        else inQuotes = false;
      } else field += ch;
    } else {
      if (ch === '"') inQuotes = true;
      else if (ch === ',') { row.push(field); field = ''; }
      else if (ch === '\n') { row.push(field); rows.push(row); row = []; field = ''; }
      else if (ch === '\r') { /* ignora */ }
      else field += ch;
    }
  }
  if (field.length > 0 || row.length > 0) { row.push(field); rows.push(row); }
  return rows;
}

async function buscarAbaCsv(spreadsheetId, sheetName) {
  const url = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/gviz/tq?tqx=out:csv&sheet=${encodeURIComponent(sheetName)}`;
  const response = await fetch(url, { redirect: 'follow' });
  if (!response.ok) {
    throw new Error(`HTTP ${response.status} ao acessar a aba "${sheetName}"`);
  }
  const text = await response.text();
  if (text.trim().startsWith('<')) {
    throw new Error(`A aba "${sheetName}" não pôde ser lida (planilha pode não estar pública).`);
  }
  return parseCsv(text);
}

function extrairIdPlanilha(url) {
  const match = (url || '').match(/\/d\/([a-zA-Z0-9-_]+)/);
  return match ? match[1] : null;
}

function parseDataFinalizacao(valor) {
  if (!valor) return null;
  const trimmed = String(valor).trim();
  if (!trimmed || trimmed === '#N/D' || /processo em andamento/i.test(trimmed)) return null;
  // Aceita YYYY-MM-DD e DD/MM/YYYY
  const isoMatch = trimmed.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (isoMatch) return new Date(Number(isoMatch[1]), Number(isoMatch[2]) - 1, Number(isoMatch[3]));
  const brMatch = trimmed.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (brMatch) return new Date(Number(brMatch[3]), Number(brMatch[2]) - 1, Number(brMatch[1]));
  const generic = new Date(trimmed);
  return isNaN(generic.getTime()) ? null : generic;
}

// Lê uma aba de um xlsx e devolve um array de objetos indexados pelo cabeçalho
// (linha 1). Retorna null se a aba não existir.
async function lerAbaXlsx(caminho, nomeAba) {
  // Realiza o require apenas no momento da execução para evitar carregamento desnecessário na inicialização
  const ExcelJS = require('exceljs');
  
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(caminho);
  const worksheet = workbook.getWorksheet(nomeAba);
  if (!worksheet) return null;

  const headers = [];
  const headerRow = worksheet.getRow(1);
  headerRow.eachCell({ includeEmpty: false }, (cell, colNumber) => {
    headers[colNumber] = String(cell.value ?? '').trim();
  });

  const linhas = [];
  for (let r = 2; r <= worksheet.rowCount; r++) {
    const row = worksheet.getRow(r);
    const obj = {};
    let vazio = true;
    for (let c = 1; c < headers.length; c++) {
      const nome = headers[c];
      if (!nome) continue;
      const valor = row.getCell(c).value;
      const texto = valor == null ? '' : String(typeof valor === 'object' && 'text' in valor ? valor.text : valor).trim();
      obj[nome] = texto;
      if (texto) vazio = false;
    }
    if (!vazio) linhas.push(obj);
  }
  return linhas;
}

// Constrói um índice processo → Set de valores de uma coluna específica.
// Usado para checar rapidamente se um processo aparece e quais valores carrega.
function indexarPorProcesso(linhas, colunaProcesso, colunaValor) {
  const mapa = new Map();
  if (!linhas) return mapa;
  for (const linha of linhas) {
    const proc = (linha[colunaProcesso] || '').trim();
    if (!proc) continue;
    const valor = (linha[colunaValor] || '').trim();
    if (!mapa.has(proc)) mapa.set(proc, new Set());
    mapa.get(proc).add(valor);
  }
  return mapa;
}

// Determina o status de cada dimensão (Licitação/Execução) cruzando o que a
// Planilha de Controle declara com o que o usuário já disponibilizou nos xlsx.
function classificarStatus(reg, mapaLicitacao, mapaExecucao) {
  // Licitação
  if (reg.obsLicitacao === 'Migrado para o ano seguinte' || reg.obsLicitacao === 'Inativo') {
    reg.obsLicitacaoStatus = 'na';
  } else if (reg.obsLicitacao === 'Incluir Pré DPL') {
    const resultados = mapaLicitacao.get(reg.processo);
    if (!resultados) {
      reg.obsLicitacaoStatus = 'pendente';
    } else if (Array.from(resultados).every(v => v === 'Não licitado')) {
      reg.obsLicitacaoStatus = 'atendido';
    } else {
      reg.obsLicitacaoStatus = 'divergente';
    }
  } else if (reg.obsLicitacao === 'Incluir Pós DPL') {
    const resultados = mapaLicitacao.get(reg.processo);
    if (!resultados) {
      reg.obsLicitacaoStatus = 'pendente';
    } else if (Array.from(resultados).some(v => v && v !== 'Não licitado')) {
      reg.obsLicitacaoStatus = 'atendido';
    } else {
      reg.obsLicitacaoStatus = 'divergente';
    }
  } else {
    reg.obsLicitacaoStatus = 'analise';
  }

  // Execução
  if (reg.obsExecucao === 'Migrado para o ano seguinte' || reg.obsExecucao === 'Inativo') {
    reg.obsExecucaoStatus = 'na';
  } else if (reg.obsExecucao === 'Incluir Execução') {
    const valores = mapaExecucao.get(reg.processo);
    if (!valores) {
      reg.obsExecucaoStatus = 'pendente';
    } else if (valores.has('Sim')) {
      reg.obsExecucaoStatus = 'atendido';
    } else {
      reg.obsExecucaoStatus = 'pendente';
    }
  } else {
    reg.obsExecucaoStatus = 'analise';
  }
}

function registerObservatorioRoute(app, logger) {
  app.get('/api/observatorio/planilha-controle', async (req, res) => {
    // Stream de progresso via Server-Sent Events: cada aba lida emite uma etapa,
    // mais uma etapa final de classificação dos processos.
    res.setHeader('Content-Type', 'text/event-stream');
    res.setHeader('Cache-Control', 'no-cache');
    res.setHeader('Connection', 'keep-alive');
    res.flushHeaders();

    const send = (event, payload) => {
      res.write(`event: ${event}\ndata: ${JSON.stringify(payload)}\n\n`);
    };

    try {
      const url = process.env.COMPRAS_URL_PLANILHA_CONTROLE;
      if (!url) {
        send('fail', { message: 'COMPRAS_URL_PLANILHA_CONTROLE não definido no .env' });
        return res.end();
      }
      const spreadsheetId = extrairIdPlanilha(url);
      if (!spreadsheetId) {
        send('fail', { message: 'URL da planilha de controle inválida' });
        return res.end();
      }

      const pastaBase = (req.query.pasta || '').toString().trim();

      const anoAtual = new Date().getFullYear();
      const anos = [];
      for (let ano = 2022; ano <= anoAtual; ano++) anos.push(ano);

      // Etapas: N abas + classificação + (se houver pasta) leitura dos 2 xlsx + comparação
      const totalEtapas = anos.length + 1 + (pastaBase ? 3 : 0);
      let etapa = 0;

      // Colunas exatas e por prefixo (a coluna de data tem título longo e variável no Sheets)
      const COLUNAS_EXATAS = {
        processo: 'Processo (CAPL)',
        situacao: 'Situação',
        obsLicitacao: 'Observatório - Licitação',
        obsExecucao: 'Observatório - Execução',
      };
      const COLUNAS_PREFIXO = {
        dataFinalizacao: 'Data de finalização',
      };

      const resultados = [];
      const abasComErro = [];

      // Posição padrão da coluna "Data de finalização" quando o título não vier
      // no cabeçalho (célula mesclada nas abas mais antigas faz o gviz devolver vazio).
      const DATA_FINALIZACAO_FALLBACK_IDX = 24;

      const normalizar = s => (s || '').replace(/\\s+/g, ' ').trim();
      const prefixoDataFinalizacao = normalizar(COLUNAS_PREFIXO.dataFinalizacao);

      // Primeira passada: baixa todas as abas e calcula índices de colunas exatas.
      const abasCarregadas = [];
      let idxDataFinalizacaoAprendido = -1;

      for (const ano of anos) {
        etapa++;
        send('progress', { current: etapa, total: totalEtapas, label: `Lendo aba "Licitação ${ano}"...` });
        const sheetName = `Licitação ${ano}`;
        let linhas;
        try {
          linhas = await buscarAbaCsv(spreadsheetId, sheetName);
        } catch (err) {
          abasComErro.push({ ano, erro: err.message });
          continue;
        }

        if (linhas.length < 3) continue;

        // Linha 1 (índice 0) é descartada; linha 2 (índice 1) contém os nomes das colunas
        // Normaliza espaços/quebras de linha porque títulos no Sheets podem ter Alt+Enter
        const cabecalho = linhas[1].map(normalizar);
        const idx = {};
        for (const [chave, nome] of Object.entries(COLUNAS_EXATAS)) {
          idx[chave] = cabecalho.indexOf(normalizar(nome));
        }
        if (idx.processo === -1) {
          abasComErro.push({ ano, erro: 'Coluna "Processo (CAPL)" não encontrada' });
          continue;
        }

        const idxDataFinalizacaoLocal = cabecalho.findIndex(h => h.startsWith(prefixoDataFinalizacao));
        if (idxDataFinalizacaoLocal !== -1 && idxDataFinalizacaoAprendido === -1) {
          idxDataFinalizacaoAprendido = idxDataFinalizacaoLocal;
        }

        abasCarregadas.push({ ano, linhas, idx });
      }

      etapa++;
      send('progress', { current: etapa, total: totalEtapas, label: 'Classificando processos...' });

      // Aplica o índice aprendido (ou o fallback fixo) para todas as abas.
      const idxDataFinalizacao = idxDataFinalizacaoAprendido !== -1
        ? idxDataFinalizacaoAprendido
        : DATA_FINALIZACAO_FALLBACK_IDX;

      for (const { ano, linhas, idx } of abasCarregadas) {
        idx.dataFinalizacao = idxDataFinalizacao;
        for (let i = 2; i < linhas.length; i++) {
          const linha = linhas[i];
          const processo = (linha[idx.processo] || '').trim();
          if (!processo) continue;
          resultados.push({
            ano,
            processo,
            situacao: (linha[idx.situacao] || '').trim(),
            dataFinalizacao: (linha[idx.dataFinalizacao] || '').trim(),
            obsLicitacao: '',
            obsExecucao: '',
          });
        }
      }

      // Identifica processos migrados para o ano seguinte
      const processosPorAno = new Map();
      for (const reg of resultados) {
        if (!processosPorAno.has(reg.ano)) processosPorAno.set(reg.ano, new Set());
        processosPorAno.get(reg.ano).add(reg.processo);
      }

      const hoje = new Date();
      const umAnoAtras = new Date(hoje.getFullYear() - 1, hoje.getMonth(), hoje.getDate());

      for (const reg of resultados) {
        const proxAno = processosPorAno.get(reg.ano + 1);
        const migrado = proxAno && proxAno.has(reg.processo);

        if (migrado) {
          reg.obsLicitacao = 'Migrado para o ano seguinte';
          reg.obsExecucao = 'Migrado para o ano seguinte';
          continue;
        }

        const situacao = reg.situacao;
        if (situacao === 'Inativo') {
          reg.obsLicitacao = 'Inativo';
          reg.obsExecucao = 'Inativo';
        } else if (situacao === 'PROAD') {
          const dataFim = parseDataFinalizacao(reg.dataFinalizacao);
          reg.obsLicitacao = dataFim ? 'Incluir Pós DPL' : 'Incluir Pré DPL';
          if (dataFim && dataFim < umAnoAtras) {
            reg.obsExecucao = 'Incluir Execução';
          }
        }
      }

      // Cruzamento com os arquivos "Dados Visão Licitação.xlsx" e "Dados Visão Execução.xlsx"
      // gerados pelo script R, se o usuário informou a pasta da base.
      let mapaLicitacao = new Map();
      let mapaExecucao = new Map();
      const arquivosComErro = [];

      if (pastaBase) {
        const caminhoLicitacao = path.join(pastaBase, 'Dados Visão Licitação.xlsx');
        const caminhoExecucao = path.join(pastaBase, 'Dados Visão Execução.xlsx');

        etapa++;
        send('progress', { current: etapa, total: totalEtapas, label: 'Lendo "Dados Visão Licitação.xlsx"...' });
        try {
          const linhas = await lerAbaXlsx(caminhoLicitacao, 'Mapa de licitações');
          if (!linhas) throw new Error('Aba "Mapa de licitações" não encontrada');
          mapaLicitacao = indexarPorProcesso(linhas, 'Processo', 'resultado');
        } catch (err) {
          arquivosComErro.push({ arquivo: 'Dados Visão Licitação.xlsx', erro: err.message });
        }

        etapa++;
        send('progress', { current: etapa, total: totalEtapas, label: 'Lendo "Dados Visão Execução.xlsx"...' });
        try {
          const linhas = await lerAbaXlsx(caminhoExecucao, 'Mapa de licitações');
          if (!linhas) throw new Error('Aba "Mapa de licitações" não encontrada');
          mapaExecucao = indexarPorProcesso(linhas, 'processo', 'executado');
        } catch (err) {
          arquivosComErro.push({ arquivo: 'Dados Visão Execução.xlsx', erro: err.message });
        }

        etapa++;
        send('progress', { current: etapa, total: totalEtapas, label: 'Comparando processos com a base...' });
      }

      for (const reg of resultados) {
        classificarStatus(reg, mapaLicitacao, mapaExecucao);
      }

      send('done', {
        anos,
        total: resultados.length,
        registros: resultados,
        abasComErro,
        arquivosComErro,
        comparou: Boolean(pastaBase),
      });
      res.end();
    } catch (error) {
      logger.error(`Erro ao processar planilha de controle: ${error.message}`, 'Observatório', error);
      send('fail', { message: error.message });
      res.end();
    }
  });
}

module.exports = { registerObservatorioRoute };
