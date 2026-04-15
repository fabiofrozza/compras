const Logger = require('../utils/logger');

const logger = new Logger({ minLevel: process.env.COMPRAS_LOGGER_MIN_LEVEL || 'debug' });

async function validateLink(url) {
  if (!url) {
    return {
      isValid: false,
      status: 'info',
      msg: 'Informe o link da aba LISTA FINAL e aguarde.',
    };
  }

  try {
    new URL(url);
  } catch {
    return {
      isValid: false,
      status: 'error',
      msg: 'Link inválido.',
    };
  }

  try {
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), 10000);

    const response = await fetch(url, {
      signal: controller.signal,
      headers: { 'User-Agent': 'Mozilla/5.0' }
    });
    clearTimeout(timeout);

    if (!response.ok) {
      const httpMessages = {
        401: 'Sem permissão para acessar a planilha. Verifique se o link é público.',
        403: 'Acesso negado à planilha. Verifique as permissões de compartilhamento.',
        404: 'Planilha não encontrada. Verifique se o link está correto.',
        429: 'Muitas requisições ao servidor. Aguarde um momento e tente novamente.',
      };
      const msg = httpMessages[response.status]
        || (response.status >= 500 ? 'Erro no servidor do Google. Tente novamente mais tarde.' : `Erro ao acessar o link (HTTP ${response.status}).`);
      return {
        isValid: false,
        status: 'error',
        msg,
      };
    }

    const htmlContent = await response.text();

    if (htmlContent.includes('LISTA FINAL')) {
      const inputValueRegex = /<input[^>]*value="([^"]+)"[^>]*>/gi;
      const inputValues = [];
      let inputMatch;
      while ((inputMatch = inputValueRegex.exec(htmlContent)) !== null) {
        if (inputMatch[1] && inputMatch[1].trim()) {
          inputValues.push(inputMatch[1].trim());
        }
      }
      const grupoMateriais = inputValues.length > 0 ? inputValues.join(', ') : 'Grupo não identificado';

      const regexSPA = /23080\.\d{6}\/\d{4}-\d{2}/g;
      const processosSPA = [...new Set(htmlContent.match(regexSPA) || [])].sort();

      const tdContentRegex = /<td[^>]*>([\s\S]*?)<\/td>/gi;
      let temValidacaoManual = false;
      let tdMatch;
      while ((tdMatch = tdContentRegex.exec(htmlContent)) !== null) {
        if (tdMatch[1].includes('VALIDAÇÃO MANUAL')) {
          temValidacaoManual = true;
          break;
        }
      }

      return {
        isValid: true,
        status: 'success',
        msg: grupoMateriais,
        processosSPA,
        temValidacaoManual,
      };
    } else {
      return {
        isValid: true,
        status: 'warning',
        msg: 'Este não parece ser um link de planilha de inserção de demandas.',
        processosSPA: [],
      };
    }

  } catch (error) {
    logger.error(`Erro ao validar link: ${error.message}`, 'ValidateLink');
    const isTimeout = error.name === 'AbortError';
    return {
      isValid: false,
      status: 'error',
      msg: isTimeout
        ? 'A requisição excedeu o tempo limite. Verifique sua conexão ou tente novamente.'
        : 'Não foi possível acessar o link informado. Verifique sua conexão.',
      error: error.message,
    };
  }
}

module.exports = {
  validateLink,
};
