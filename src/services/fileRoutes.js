const fs = require('fs').promises;
const path = require('path');
const os = require('os');
const {
  SCRIPTS_PATH,
  ALLOWED_DELETE_FOLDERS,
  isPathSafe,
  openInInterface,
  resolveScriptFolder,
  parseFilters,
  filePassesFilter,
  listFilesWithStats,
  deleteFilesInFolder,
  validateAndDeleteFile,
  listPregoes,
  listPregaoFiles,
  listImportFiles,
  deletePregao,
  deleteSupplierFile,
  moveSupplierFiles,
} = require('./files');
const { validateLink } = require('./spreadsheet');

function registerFileRoutes(app, logger) {
  app.post('/api/open-folder', async (req, res) => {
    try {
      const { folderPath } = req.body;

      if (!folderPath) {
        logger.warn('Solicitação para abrir pasta sem caminho', 'API');
        return res.status(400).json({ error: 'Caminho da pasta não fornecido' });
      }

      try {
        await fs.access(folderPath);
      } catch {
        logger.warn(`Pasta não encontrada: ${folderPath}`, 'API');
        return res.status(404).json({ error: `Pasta não encontrada: ${folderPath}` });
      }

      if (!isPathSafe(folderPath, SCRIPTS_PATH)) {
        return res.status(403).json({ error: 'Acesso negado: fora do diretório permitido' });
      }

      try {
        openInInterface(folderPath);
        logger.debug(`Pasta aberta: ${folderPath}`, 'API');
      } catch (spawnError) {
        logger.error(`Erro ao abrir pasta: ${spawnError.message}`, 'API', spawnError);
        return res.status(500).json({ error: `Erro ao abrir pasta: ${spawnError.message}` });
      }

      res.json({ success: true, message: 'Pasta aberta' });
    } catch (error) {
      logger.error(`Erro geral ao abrir pasta: ${error.message}`, 'API', error);
      res.status(500).json({ error: `Erro: ${error.message}` });
    }
  });

  app.post('/api/open-file', async (req, res) => {
    try {
      const { filePath } = req.body;

      if (!filePath) {
        logger.warn('Solicitação para abrir arquivo sem caminho', 'API');
        return res.status(400).json({ error: 'Caminho do arquivo não fornecido' });
      }

      try {
        const stats = await fs.stat(filePath);
        if (!stats.isFile()) {
          return res.status(400).json({ error: `O caminho fornecido não é um arquivo: ${filePath}` });
        }
      } catch {
        logger.warn(`Arquivo não encontrado: ${filePath}`, 'API');
        return res.status(404).json({ error: `Arquivo não encontrado: ${filePath}` });
      }

      if (!isPathSafe(filePath, SCRIPTS_PATH)) {
        return res.status(403).json({ error: 'Acesso negado: fora do diretório permitido' });
      }

      try {
        openInInterface(filePath, true);
        logger.debug(`Arquivo aberto: ${filePath}`, 'API');
      } catch (spawnError) {
        logger.error(`Erro ao abrir arquivo: ${spawnError.message}`, 'API', spawnError);
        return res.status(500).json({ error: `Erro ao abrir arquivo: ${spawnError.message}` });
      }

      res.json({ success: true, message: 'Arquivo aberto' });
    } catch (error) {
      logger.error(`Erro geral ao abrir arquivo: ${error.message}`, 'API', error);
      res.status(500).json({ error: `Erro: ${error.message}` });
    }
  });

  app.get('/api/check-atas-data', async (req, res) => {
    try {
      const filePath = path.join(__dirname, '..', '..', 'scripts', 'atas', 'dados_atas.xlsx');
      const configPath = path.join(__dirname, '..', '..', 'scripts', '_common', 'config.json');

      try {
        const stats = await fs.stat(filePath);

        let atasConfig = {};
        try {
          const configRaw = await fs.readFile(configPath, 'utf-8');
          const configJson = JSON.parse(configRaw);
          if (configJson.ATAS) {
            atasConfig = configJson.ATAS;
          }
        } catch (e) {
          // Ignora se não conseguir ler o config.json
        }

        res.json({
          exists: true,
          modified: stats.mtime,
          atasConfig
        });
      } catch (e) {
        res.json({ exists: false });
      }
    } catch (error) {
      logger.error(`Erro ao verificar dados de atas: ${error.message}`, 'API', error);
      res.status(500).json({ error: error.message });
    }
  });

  app.get('/api/list-files/:scriptName/:innerFolder', async (req, res) => {
    try {
      const { scriptName, innerFolder } = req.params;
      const { extensions, nameContains, sort } = req.query;

      const result = await resolveScriptFolder(res, scriptName, innerFolder, 'ListFiles');
      if (!result) return;
      const { folderPath, created } = result;

      const { patterns, filterNameContains } = parseFilters(extensions, nameContains);
      const fileDetails = await listFilesWithStats(folderPath, file => filePassesFilter(file, patterns, filterNameContains));

      if (sort === 'desc') {
        fileDetails.sort((a, b) => b.name.localeCompare(a.name));
      } else if (sort === 'asc') {
        fileDetails.sort((a, b) => a.name.localeCompare(b.name));
      }

      logger.debug(`Listados ${fileDetails.length} itens de ${scriptName}/${innerFolder}`, 'ListFiles');
      const canDelete = ALLOWED_DELETE_FOLDERS.includes(innerFolder.toLowerCase());
      res.json({ files: fileDetails, folderPath, canDelete, folderCreated: created });
    } catch (error) {
      logger.error(`Erro ao listar arquivos: ${error.message}`, 'ListFiles', error);
      res.status(500).json({ error: error.message });
    }
  });

  app.delete('/api/clear-folder/:scriptName/:innerFolder', async (req, res) => {
    try {
      const { scriptName, innerFolder } = req.params;
      const { extensions, nameContains } = req.query;

      if (!ALLOWED_DELETE_FOLDERS.includes(innerFolder.toLowerCase())) {
        return res.status(403).json({ error: 'A exclusão nesta pasta não é permitida por segurança.' });
      }

      const result = await resolveScriptFolder(res, scriptName, innerFolder, 'ClearFolder');
      if (!result) return;
      const { folderPath } = result;

      const { patterns, filterNameContains } = parseFilters(extensions, nameContains);
      const deletedCount = await deleteFilesInFolder(folderPath, file => filePassesFilter(file, patterns, filterNameContains));

      logger.info(`Excluído(s) ${deletedCount} arquivo(s) de ${scriptName}/${innerFolder}`, 'ClearFolder');
      res.json({
        message: `${deletedCount} arquivo(s) excluído(s) com sucesso!`,
        deletedCount
      });
    } catch (error) {
      logger.error(`Erro ao limpar pasta: ${error.message}`, 'ClearFolder', error);
      res.status(500).json({ error: error.message });
    }
  });

  app.delete('/api/delete-file/:scriptName/:innerFolder', async (req, res) => {
    try {
      const { scriptName, innerFolder } = req.params;
      const { fileName } = req.body;

      if (!fileName) {
        return res.status(400).json({ error: 'Nome do arquivo não fornecido' });
      }

      if (!ALLOWED_DELETE_FOLDERS.includes(innerFolder.toLowerCase())) {
        return res.status(403).json({ error: 'A exclusão nesta pasta não é permitida por segurança.' });
      }

      const resolved = await resolveScriptFolder(res, scriptName, innerFolder, 'DeleteFile');
      if (!resolved) return;
      const { folderPath } = resolved;

      if (fileName.includes('/') || fileName.includes('\\') || fileName === '..' || fileName === '.') {
        return res.status(400).json({ error: 'Nome de arquivo inválido' });
      }

      const filePath = path.join(folderPath, fileName);
      const result = await validateAndDeleteFile(filePath, folderPath);

      if (result.error) {
        return res.status(result.statusCode).json({ error: result.error });
      }

      logger.info(`Arquivo excluído: ${fileName} de ${scriptName}/${innerFolder}`, 'DeleteFile');
      res.json({ success: true, message: 'Arquivo excluído com sucesso' });
    } catch (error) {
      logger.error(`Erro ao excluir arquivo: ${error.message}`, 'DeleteFile', error);
      res.status(500).json({ error: error.message });
    }
  });

  app.post('/api/validate-link', async (req, res) => {
    try {
      const { url } = req.body;
      const result = await validateLink(url);
      res.json(result);
    } catch (error) {
      logger.error(`Erro ao validar link: ${error.message}`, 'ValidateLink', error);
      res.status(500).json({ error: error.message });
    }
  });

  // =============================================
  // Fornecedores
  // =============================================

  app.get('/api/fornecedores/pregoes', async (_req, res) => {
    try {
      const result = await listPregoes(res);
      if (!result) return;
      res.json(result);
    } catch (error) {
      logger.error(`Erro ao listar pregões: ${error.message}`, 'Fornecedores', error);
      res.status(500).json({ error: error.message });
    }
  });

  app.get('/api/fornecedores/pregao/:pregao/arquivos', async (req, res) => {
    try {
      const { pregao } = req.params;
      const result = await listPregaoFiles(pregao);

      if (result.error) {
        return res.status(result.statusCode).json({ error: result.error });
      }

      res.json(result);
    } catch (error) {
      logger.error(`Erro ao listar arquivos do pregão: ${error.message}`, 'Fornecedores', error);
      res.status(500).json({ error: error.message });
    }
  });

  app.get('/api/fornecedores/importar', async (_req, res) => {
    try {
      const result = await listImportFiles();
      res.json(result);
    } catch (error) {
      logger.error(`Erro ao listar arquivos importar: ${error.message}`, 'Fornecedores', error);
      res.status(500).json({ error: error.message });
    }
  });

  app.delete('/api/fornecedores/pregao/:pregao', async (req, res) => {
    try {
      const { pregao } = req.params;

      if (!ALLOWED_DELETE_FOLDERS.includes('dados')) {
        return res.status(403).json({ error: 'A exclusão nesta pasta não é permitida por segurança.' });
      }

      const result = await deletePregao(pregao);

      if (result.error) {
        return res.status(result.statusCode).json({ error: result.error });
      }

      res.json(result);
    } catch (error) {
      logger.error(`Erro ao excluir pregão: ${error.message}`, 'Fornecedores', error);
      res.status(500).json({ error: error.message });
    }
  });

  app.delete('/api/fornecedores/arquivo', async (req, res) => {
    try {
      const { filePath } = req.body;

      if (!filePath) {
        return res.status(400).json({ error: 'Caminho do arquivo não fornecido' });
      }

      const result = await deleteSupplierFile(filePath);

      if (result.error) {
        return res.status(result.statusCode).json({ error: result.error });
      }

      res.json(result);
    } catch (error) {
      logger.error(`Erro ao excluir arquivo: ${error.message}`, 'Fornecedores', error);
      res.status(500).json({ error: error.message });
    }
  });

  app.post('/api/fornecedores/mover', async (req, res) => {
    try {
      const { pregao } = req.body;

      if (!pregao) {
        return res.status(400).json({ error: 'Número do pregão não fornecido' });
      }

      const result = await moveSupplierFiles(pregao, os);

      if (result.error) {
        return res.status(result.statusCode).json({ error: result.error });
      }

      res.json(result);
    } catch (error) {
      logger.error(`Erro ao mover arquivos: ${error.message}`, 'Fornecedores', error);
      res.status(500).json({ error: error.message });
    }
  });
}

module.exports = { registerFileRoutes };
