const fs = require('fs');
const path = require('path');
const os = require('os');

const MAX_LOG_FILES = parseInt(process.env.COMPRAS_MAX_LOG_FILES) || 20;

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
        this.fileStream = null;
        this.logDir = null;

        if (options.logDir) {
            this._initFileLogging(options.logDir);
        }
    }

    _initFileLogging(logDir) {
        this.logDir = logDir;
        try {
            fs.mkdirSync(logDir, { recursive: true });

            const now = new Date();
            const date = now.getFullYear()
                + '-' + String(now.getMonth() + 1).padStart(2, '0')
                + '-' + String(now.getDate()).padStart(2, '0');
            const time = String(now.getHours()).padStart(2, '0')
                + '-' + String(now.getMinutes()).padStart(2, '0')
                + '-' + String(now.getSeconds()).padStart(2, '0');
            const user = (process.env.USERNAME || os.userInfo().username).toUpperCase();
            const node = os.hostname().toUpperCase();

            const logName = `Log_NODEJS_${date}_${time}_${user}-${node}.log`;
            const logPath = path.join(logDir, logName);
            this.fileStream = fs.createWriteStream(logPath, { flags: 'a', encoding: 'utf8' });

            this._pruneOldLogs(logDir);
        } catch (err) {
            console.error('Failed to initialize file logging:', err.message);
        }
    }

    _pruneOldLogs(logDir) {
        try {
            const files = fs.readdirSync(logDir)
                .filter(f => f.startsWith('Log_NODEJS_') && f.endsWith('.log'))
                .map(f => ({ name: f, time: fs.statSync(path.join(logDir, f)).mtimeMs }))
                .sort((a, b) => b.time - a.time);

            for (const file of files.slice(MAX_LOG_FILES)) {
                fs.unlinkSync(path.join(logDir, file.name));
            }
        } catch { /* best-effort cleanup */ }
    }

    _writeToFile(formatted, extra) {
        if (!this.fileStream) return;
        this.fileStream.write(formatted + '\n');
        if (extra) this.fileStream.write(extra + '\n');
    }

    shouldLog(level) {
        return this.levels[level] >= this.levels[this.minLevel];
    }

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

    debug(message, context = '') {
        const formatted = this.formatMessage('debug', message, context);
        this._writeToFile(formatted);
        if (this.shouldLog('debug')) {
            console.log(formatted);
        }
    }

    info(message, context = '') {
        const formatted = this.formatMessage('info', message, context);
        this._writeToFile(formatted);
        if (this.shouldLog('info')) {
            console.log(formatted);
        }
    }

    warn(message, context = '') {
        const formatted = this.formatMessage('warn', message, context);
        this._writeToFile(formatted);
        if (this.shouldLog('warn')) {
            console.warn(formatted);
        }
    }

    error(message, context = '', errorObj = null) {
        const formatted = this.formatMessage('error', message, context);
        let extra = null;
        if (errorObj instanceof Error) {
            extra = '  Stack: ' + errorObj.stack;
        } else if (errorObj) {
            extra = '  Details: ' + JSON.stringify(errorObj);
        }
        this._writeToFile(formatted, extra);
        if (this.shouldLog('error')) {
            console.error(formatted);
            if (errorObj instanceof Error) {
                console.error('  Stack:', errorObj.stack);
            } else if (errorObj) {
                console.error('  Details:', errorObj);
            }
        }
    }

    section(title) {
        this.info('►►► ' + title);
    }

    success(message, context = '') {
        const formatted = this.formatMessage('info', '✓ ' + message, context);
        this._writeToFile(formatted);
        if (this.shouldLog('info')) {
            console.log(formatted);
        }
    }

    setLevel(level) {
        if (this.levels[level] !== undefined) {
            this.minLevel = level;
        }
    }

    enableFileLogging(logDir) {
        const dir = logDir || this.logDir;
        if (dir && !this.fileStream) {
            this._initFileLogging(dir);
        }
    }

    disableFileLogging() {
        if (this.fileStream) {
            this.fileStream.end();
            this.fileStream = null;
        }
    }
}

module.exports = Logger;
