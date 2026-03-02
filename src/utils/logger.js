// logger.js - Sistema de logging com níveis e filtros
// Níveis: 'debug', 'info', 'warn', 'error'

class Logger {
    constructor(options = {}) {
        this.minLevel = options.minLevel || 'info';
        this.showTimestamp = options.showTimestamp !== false;
        this.levels = {
            debug: 0,
            info: 1,
            warn: 2,
            error: 3
        };
    }

    /**
     * Verifica se a mensagem deve ser logada baseado no nível
     */
    shouldLog(level) {
        return this.levels[level] >= this.levels[this.minLevel];
    }

    /**
     * Formata a mensagem com timestamp e nível
     */
    formatMessage(level, message, context = '') {
        const timestamp = this.showTimestamp 
            ? new Date().toLocaleTimeString('pt-BR') 
            : '';
        
        const levelSymbols = {
            debug: '[debug]',
            info:  '[info ]',
            warn:  '[warn ]',
            error: '[error]'
        };

        const prefix = levelSymbols[level] || '○';
        const ctxStr = context ? ` [${context}]` : '';
        const timeStr = timestamp ? ` ${timestamp}` : '';

        return `${timeStr} ${prefix}${ctxStr} ${message}`;
    }

    /**
     * Log de debug (desenvolvimento)
     */
    debug(message, context = '') {
        if (this.shouldLog('debug')) {
            console.log(this.formatMessage('debug', message, context));
        }
    }

    /**
     * Log de informação (eventos normais)
     */
    info(message, context = '') {
        if (this.shouldLog('info')) {
            console.log(this.formatMessage('info', message, context));
        }
    }

    /**
     * Log de aviso (possíveis problemas)
     */
    warn(message, context = '') {
        if (this.shouldLog('warn')) {
            console.warn(this.formatMessage('warn', message, context));
        }
    }

    /**
     * Log de erro (problemas críticos)
     */
    error(message, context = '', errorObj = null) {
        if (this.shouldLog('error')) {
            console.error(this.formatMessage('error', message, context));
            if (errorObj instanceof Error) {
                console.error('  Stack:', errorObj.stack);
            } else if (errorObj) {
                console.error('  Details:', errorObj);
            }
        }
    }

    /**
     * Seção de início (comando importante)
     */
    section(title) {
        this.info('►►► ' + title);
    }

    /**
     * Seção de sucesso
     */
    success(message, context = '') {
        if (this.shouldLog('info')) {
            console.log(this.formatMessage('info', '✓ ' + message, context));
        }
    }

    /**
     * Definir nível mínimo de log
     */
    setLevel(level) {
        if (this.levels[level] !== undefined) {
            this.minLevel = level;
        }
    }
}

module.exports = Logger;
