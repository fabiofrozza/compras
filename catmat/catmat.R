catmat_obter_tr <- function(lista_itens_TR) {
  
  pasta <- get_config("pasta")
  
  tryCatch({
    tr <- read_excel(file.path(pasta$tr, lista_itens_TR))
  },
  error = function(e) {
    log_erro(sprintf("Não foi possível ler a lista de itens do TR %s. Encerrando...", lista_itens_TR),
             e,
             finalizar = TRUE)
  })
  
  processo <- str_extract_all(as.character(tr), "23080\\.\\d{6}/\\d{4}-\\d{2}", simplify = TRUE)
  processo <- unique(processo[nchar(processo) > 0])
  processo <- if(length(processo) > 0) {
    str_sub(gsub("/", "-", processo[1]), 7, 17)
  } else {
    "Processo não encontrado"
  }
  
  # Encontrar a linha que contém os cabeçalhos
  linha_cabecalho <- which(apply(tr, 1, function(x) {
    any(grepl("Grupo/Item", x, ignore.case = TRUE)) & 
    any(grepl("Descrição", x, ignore.case = TRUE))
  }))
  
  if (length(linha_cabecalho) == 0 || is.na(linha_cabecalho)) {
    log_erro("Linha inicial das colunas 'Grupo/Item' e 'Descrição' não encontrada. Encerrando...",
             finalizar = TRUE)
  }
  
  # Extrair os cabeçalhos e definir como nomes das colunas
  cabecalhos <- unlist(tr[linha_cabecalho, ])
  colnames(tr) <- cabecalhos
  
  # Extrair os dados a partir da linha seguinte aos cabeçalhos
  tr_itens <- tr[-c(1:linha_cabecalho), ]
  
  catmat_tr <- data.frame(
    Item = tr_itens$`Grupo/Item`,
    CATMAT = trimws(str_split_i(tr_itens$Descrição, "-", 1))
  )
  
  catmat_tr <- na.omit(catmat_tr)
  
  if (nrow(catmat_tr) == 0) {
    log_erro("Nenhum código CATMAT encontrado no relatório. Encerrando...", 
             finalizar = TRUE)
  }

  return(list(catmat_tr = catmat_tr,
              processo  = processo))  
}

catmat_obter_margens <- function() {

  pasta <- get_config("pasta")
  
  tryCatch({
    planilha_margens <- read_excel(file.path(pasta$fontes, "Margens.xlsx"),
                                   col_types = c("text", "text", "text", "numeric", "text", "numeric"))
  },
  error = function(e) {
    log_erro("Não foi possível ler a planilha de margens de preferência (_fontes/Margens.xlsx). Encerrando...",
             e,
             finalizar = TRUE)
  })

  colunas_exigidas <- c("NCM", "Descrição", "Regra de origem", "Margem normal", "Regra de qualificação", "Margem adicional")
  colunas_planilha <- names(planilha_margens)
  
  if (!all(colunas_exigidas == colunas_planilha)) {
    log_erro("Há diferença no nome das colunas do arquivo Margens.xlsx. Elas devem ser exatamente iguais a:", colunas_exigidas,
             finalizar = TRUE)
  } 

  # Adiciona coluna com números das linhas do Excel (começando de 1)
  planilha_margens$linha_planilha <- 1:nrow(planilha_margens) + 1  # +1 porque o Excel começa na linha 1 (cabeçalho)
  
  margens <- planilha_margens %>%
    filter(!is.na(NCM)) %>%
    mutate(NCM = gsub(".", "", NCM, fixed = TRUE),
           NCM = gsub(",", "", NCM, fixed = TRUE),
           NCMvalido = str_sub(NCM, 1, 2) == str_sub(Descrição, 1, 2))
  
  if(any(!margens$NCMvalido)) {
    
    NCMinvalidos <- margens %>%
      filter(!NCMvalido) %>%
      mutate(Descrição = paste0(str_sub(Descrição, 1, 20), "...")) %>%
      select(linha_planilha, NCM, Descrição)
    
    log_erro("Os seguintes NCMs não correspondem à descrição. Verifique a planilha de margens de preferência nas seguintes linhas:",
             NCMinvalidos,
             finalizar = TRUE)
  }
    
  margens <- margens %>%
    select(-c(`Descrição`, NCMvalido, linha_planilha))

  return(margens)
} 

catmat_aplicar_margens <- function(resultados) {
  
  margens <- catmat_obter_margens()
  
  colunas_adicionais <- names(margens)[-1]
  
  resultados_final <- resultados %>%
    left_join(margens, by = "NCM") %>%
    relocate(all_of(colunas_adicionais), .after = NCM)
  
  return(resultados_final)
  
}

catmat_consultar_api <- function(codigo_item) {
  
  url     <- paste0(get_config("url")$api_catmat, codigo_item)
  headers <- c('Accept' = '*/*')
  
  resposta <- tryCatch({
    VERB("GET", url = url, add_headers(headers))
  },
  error = function(e) {
    log_erro("Não foi possível a consulta à API. Verifique sua conexão com a internet.",
             alerta = TRUE)
    return(e)
  })

  # Se não foi possível a conexão com a API...
  if ("error" %in% class(resposta)) {
    return(data.frame(
      "Código CATMAT" = codigo_item,
      "Status do item" = NA_character_,
      "NCM" = NA,
      "Aplica margem preferência" = NA,
      "Resultado consulta" = paste("Erro na consulta:", resposta$message),
      "descricaoItem" = NA,
      "codigoGrupo" = NA,
      "nomeGrupo" = NA,
      "codigoClasse" = NA,
      "nomeClasse" = NA,
      "codigoPdm" = NA,
      "nomePdm" = NA,
      "codigoItem" = NA,
      "itemSustentavel" = NA,
      "descricao_ncm" = NA,
      "dataHoraAtualizacao" = NA,
      stringsAsFactors = FALSE
    ))
  }
  
  # Se houve retorno da API...
  if (status_code(resposta) == 200) {
    conteudo <- content(resposta, "text", encoding = "UTF-8")
    dados <- fromJSON(conteudo)
    
    if (!is.null(dados$resultado) && length(dados$resultado) > 0) {
      return(data.frame(
        "Código CATMAT" = codigo_item,
        "Status do item" = ifelse(is.null(dados$resultado$statusItem), NA, ifelse(dados$resultado$statusItem == TRUE, "Localizado", "Não localizado")),
        "NCM" = ifelse(is.null(dados$resultado$codigo_ncm), NA, dados$resultado$codigo_ncm),
        "Aplica margem preferência" = ifelse(is.null(dados$resultado$aplica_margem_preferencia), NA, dados$resultado$aplica_margem_preferencia),
        "Resultado consulta" = "Sucesso",
        "descricaoItem" = ifelse(is.null(dados$resultado$descricaoItem), NA, dados$resultado$descricaoItem),
        "codigoGrupo" = ifelse(is.null(dados$resultado$codigoGrupo), NA, dados$resultado$codigoGrupo),
        "nomeGrupo" = ifelse(is.null(dados$resultado$nomeGrupo), NA, dados$resultado$nomeGrupo),
        "codigoClasse" = ifelse(is.null(dados$resultado$codigoClasse), NA, dados$resultado$codigoClasse),
        "nomeClasse" = ifelse(is.null(dados$resultado$nomeClasse), NA, dados$resultado$nomeClasse),
        "codigoPdm" = ifelse(is.null(dados$resultado$codigoPdm), NA, dados$resultado$codigoPdm),
        "nomePdm" = ifelse(is.null(dados$resultado$nomePdm), NA, dados$resultado$nomePdm),
        "codigoItem" = ifelse(is.null(dados$resultado$codigoItem), NA, dados$resultado$codigoItem),
        "itemSustentavel" = ifelse(is.null(dados$resultado$itemSustentavel), NA, dados$resultado$itemSustentavel),
        "descricao_ncm" = ifelse(is.null(dados$resultado$descricao_ncm), NA, dados$resultado$descricao_ncm),
        "dataHoraAtualizacao" = ifelse(is.null(dados$resultado$dataHoraAtualizacao), NA, dados$resultado$dataHoraAtualizacao),
        stringsAsFactors = FALSE
      ))
    } else {
      return(data.frame(
        "Código CATMAT" = codigo_item,
        "Status do item" = NA_character_,
        "NCM" = NA,
        "Aplica margem preferência" = NA,
        "Resultado consulta" = "Nenhum dado encontrado",
        "descricaoItem" = NA,
        "codigoGrupo" = NA,
        "nomeGrupo" = NA,
        "codigoClasse" = NA,
        "nomeClasse" = NA,
        "codigoPdm" = NA,
        "nomePdm" = NA,
        "codigoItem" = NA,
        "itemSustentavel" = NA,
        "descricao_ncm" = NA,
        "dataHoraAtualizacao" = NA,
        stringsAsFactors = FALSE
      ))
    }
  } 
  
  # Se não houve retorno da API...
  # Traduz conteúdo da resposta de hexadecimal para caracteres
  hex_content <- resposta[["content"]]
  texto_resposta <- rawToChar(as.raw(strtoi(hex_content, 16L)))

  return(data.frame(
    "Código CATMAT" = codigo_item,
    "Status do item" = NA_character_,
    "NCM" = NA,
    "Aplica margem preferência" = NA,
    "Resultado consulta" = paste("Erro na API - Status code:", status_code(resposta), " - ", texto_resposta),
    "descricaoItem" = NA,
    "codigoGrupo" = NA,
    "nomeGrupo" = NA,
    "codigoClasse" = NA,
    "nomeClasse" = NA,
    "codigoPdm" = NA,
    "nomePdm" = NA,
    "codigoItem" = NA,
    "itemSustentavel" = NA,
    "descricao_ncm" = NA,
    "dataHoraAtualizacao" = NA,
    stringsAsFactors = FALSE
  ))
    
}

catmat_resultados_api <- function(catmat_tr) {
  
  codigos <- catmat_tr$CATMAT
  
  log_info(sprintf("Consultando %s código(s) CATMAT...", length(codigos)), 
           cores = "verde")
  
  pb <- log_barra_progresso("Consultando CATMAT...", length(codigos))
  
  resultados <- map_df(codigos, ~{
    log_barra_progresso(paste("Consultando código:", .x), pb = pb)
    catmat_consultar_api(.x)
  })
  
  log_barra_progresso(pb = pb)
  
  resultados <- bind_cols(catmat_tr, select(resultados, -Código.CATMAT))
  
  return(resultados)
}

catmat_consultar_lista <- function(origem = "local") {
  
  # Opção não utilizada, pois a lista de CATMATs é muito extensa e é melhor mantê-la localmente
  if (origem == "google") {
    gs4_deauth()
    
    url_lista <- get_config("url")$lista_catmat
    
    lista_catmat <- 
        range_read(url_lista, 
                   sheet = "Materiais", 
                   col_names = TRUE)
    
    return(lista_catmat)
    
  } 
    
  pasta <- get_config("pasta")
  
  log_info("Aguarde a leitura da lista de CATMATs....", 
           cores = "vermelho")
  
  tryCatch({
    lista_catmat <- read_excel(file.path(pasta$fontes, "Lista CATMAT.xlsx"), 
                               skip = 2)
  },
  error = function(e) {
    log_erro("Não foi possível ler a lista de CATMATs (_fontes/Lista CATMAT.xlsx). Encerrando...",
             e,
             finalizar = TRUE)
  })
  
  colunas_exigidas <- c("Código do Grupo", "Nome do Grupo", "Código da Classe", "Nome da Classe", "Código do PDM", "Nome do PDM", "Código do Item", "Descrição do Item", "Código NCM")
  colunas_planilha <- names(lista_catmat)
  
  if (all(colunas_exigidas == colunas_planilha)) {
    
    return(lista_catmat)
    
  } else {
    
    log_erro("Há diferença no nome das colunas do arquivo Lista CATMAT.xlsx. Elas devem começar na linha 3 e ser exatamente iguais a:", colunas_exigidas,
             finalizar = TRUE)
    
  }

}

catmat_resultados_lista <- function(catmat_tr) {
  
  lista_catmat <- catmat_consultar_lista()
  
  resultados_lista <- catmat_tr %>%
    left_join(lista_catmat, by = c("CATMAT" = "Código do Item")) %>%
    mutate(
      "Status item" = 
        ifelse(is.na(`Código do Grupo`), "Não localizado", "Localizado")) %>%
    rename(`NCM` = "Código NCM") %>%
    select("Item", "CATMAT", "Status item", "NCM", "Código do Grupo", "Nome do Grupo", "Código da Classe", "Nome da Classe", "Código do PDM", "Nome do PDM", "Descrição do Item")
  
  return(resultados_lista)
}

catmat_salvar <- function(resultado_final, processo, script_a_executar) {
  
  pasta <- get_config("pasta")
  
  arquivo_saida <- sprintf("CATMAT %s - %s.xlsx", 
                           toupper(script_a_executar), 
                           processo)
  aba           <- sprintf("Resultados - %s", 
                           toupper(script_a_executar))
  
  tryCatch({
    wb <- createWorkbook()
      addWorksheet(wb, aba)
      writeDataTable(wb, aba, 
                     resultado_final, 
                     startCol = 1, startRow = 1, 
                     tableStyle = "TableStyleMedium2", 
                     withFilter = TRUE)
      setColWidths(wb, aba, 
                   cols = 1:8, widths = 18)
      addStyle(wb, aba, style = createStyle(numFmt = "PERCENTAGE"), 
               rows = 2:(nrow(resultado_final) + 1), 
               cols = 6, gridExpand = TRUE)
      addStyle(wb, aba, style = createStyle(numFmt = "PERCENTAGE"), 
               rows = 2:(nrow(resultado_final) + 1), 
               cols = 8, gridExpand = TRUE)
      conditionalFormatting(wb, aba,
                            cols = 3, rows = 2:(nrow(resultado_final) + 1),
                            type = "contains", rule = "Não localizado", 
                            style = createStyle(fontColour = "red"))
    saveWorkbook(wb, file.path(pasta$arquivos_gerados, arquivo_saida), overwrite = TRUE)
  }, 
  error = function(e) {
    log_erro(sprintf("Não foi possível salvar o arquivo %s", arquivo_saida), 
             e,
             finalizar = TRUE)
  }, 
  warning = function(w) {
    log_erro(sprintf("Não foi possível salvar o arquivo %s", arquivo_saida), 
             w,
             finalizar = TRUE)
  })
  
  log_info("Arquivo salvo com sucesso:", 
           arquivo_saida, 
           cores = "verde")
  
  log_info("Consulta concluída!",
           "-",
           "Resultados salvos em", 
           arquivo_saida, 
           cores = "vermelho")
  
}

catmat_main <- function() {
  
  cat("=== INICIANDO SCRIPT. AGUARDE... ===\n")
  
  source(file.path("..", "_common", "config.R"), chdir = TRUE)
  
  pasta                  = list(atual = getwd())
  pasta$tr               = file.path(pasta$atual, "TR")
  pasta$arquivos_gerados = file.path(pasta$atual, "ARQUIVOS GERADOS")
  pasta$fontes           = file.path(pasta$atual, "_fontes")
  pasta$criar            = c(pasta$tr, pasta$arquivos_gerados)
  
  pacotes = c("readxl", "httr", "openxlsx", "dplyr", "purrr", "stringr")
  
  config_inicializar(pacotes, pasta)

  log_secao("Obtendo lista de CATMATs do TR")
  
  lista_itens_TR <- commandArgs(trailingOnly = TRUE)[2]
  
  if (is.na(lista_itens_TR)) {
    log_erro("Nenhum arquivo com lista de itens do TR encontrado. Encerrando...", 
             finalizar = TRUE)
  }
  
  tr        <- catmat_obter_tr(lista_itens_TR)
  catmat_tr <- tr$catmat_tr
  processo  <- tr$processo
  
  script_a_executar <- 
    utils_verificar_script(scripts_permitidos = c("api", "lista"),
                           script_padrao = "api")
  
  switch(
    script_a_executar,
    "api" = {
      log_secao("Consulta à API")
      resultados <- catmat_resultados_api(catmat_tr)
    },
    "lista" = {
      log_secao("Consulta à Lista de CATMATs")
      resultados <- catmat_resultados_lista(catmat_tr)
    }
  )
  
  log_secao("Aplicando margens de preferência")
  
  resultado_final <- catmat_aplicar_margens(resultados)
  
  log_secao("Salvando dados")
  
  catmat_salvar(resultado_final, processo, script_a_executar)
  
  config_finalizar(sucesso = TRUE)
  
}

catmat_main()
