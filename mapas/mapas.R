mapas_analisar <- function(mapas_dados, mapas_arquivos) {

  processar_mapa <- function(dados, arquivo) {
    log_barra_progresso(label = sprintf("Analisando mapa %s", basename(arquivo)), pb = pb)

    resultados <- list(
      cancelado = c(
        "Sem vencedor - Cancelado",
        "Sem vencedor - Anulado",
        "Sem vencedor - Revogado",
        "Sem vencedor - Suspenso",
        "Sem vencedor - Aceito e Habilitado com intenção de recurso"
      ),
      fracassado = c(
        "Sem vencedor - Fracassado na Análise",
        "Sem vencedor - Fracassado no Julgamento",
        "Sem vencedor - Fracassado na Disputa",
        "Sem vencedor - Cancelado no julgamento"
      ),
      deserto = c(
        "Sem vencedor - Deserto",
        "Sem vencedor - Cancelado por inexistência de proposta"
      )
    )
    
    colunas_manter <- c("processo", "edital", "data_homologacao", "item", "cod_item", "unidade", 
                        "descricao", "especificacao", "detalhamento", "vencedor", 
                        "valor_referencia", "valor_homologado", "quantidade")
    
    tryCatch({
      dados <- dados %>% 
        select(-c(1:3, 7:8)) %>%
        setNames(colunas_manter) %>%
        mutate(
          across(c(item, quantidade), as.numeric),
          across(c(valor_referencia, valor_homologado), utils_corrigir_valor),
          data_homologacao = as.Date(data_homologacao, format = "%d/%m/%Y"),
          .groups = 'drop'
        ) %>%
        group_by(item) %>%
        mutate(quantidade = sum(quantidade)) %>%
        distinct(item, .keep_all = TRUE) %>%
        mutate(
          valor = ifelse(valor_homologado == 0, valor_referencia, valor_homologado),
          resultado = case_when(
            vencedor %in% resultados$deserto ~ "Deserto",
            vencedor %in% resultados$cancelado ~ "Cancelado",
            vencedor %in% resultados$fracassado ~ "Fracassado",
            TRUE ~ "Sucesso"
          )
        ) %>%
        select(cod_item, unidade, descricao, especificacao, detalhamento,
               resultado, item, valor, quantidade,
               processo, edital, data_homologacao, vencedor)
      
      return(dados)
    }, 
    error = function(e) {
      log_erro(sprintf("Não foi possível analisar os dados. Abra o arquivo %s e verifique se está de acordo com o Manual. Encerrando...", basename(arquivo)),
               e, 
               finalizar = TRUE)
    })
  }
  
  pb <- log_barra_progresso("Aguarde...", length(mapas_arquivos))
  
  mapas <- lapply(seq_along(mapas_dados), function(i) {
    processar_mapa(mapas_dados[[i]], mapas_arquivos[i])
  })
  
  log_barra_progresso(pb = pb)

  mapas_tratados <- bind_rows(Filter(Negate(is.null), mapas))

  return(mapas_tratados)
  
}

mapas_ler <- function(mapas_arquivos) {
  # Inicia a barra de progresso
  pb <- log_barra_progresso("Lendo mapas...", length(mapas_arquivos))
  
  # Função interna para ler um único arquivo com tratamento de erro
  ler_um_arquivo <- function(arquivo) {
    tryCatch({
      log_barra_progresso(sprintf("Processando: %s", basename(arquivo)), pb = pb)
      dados <- readWorkbook(arquivo)
      
      return(dados)
    }, 
    error = function(e) {
      log_erro(sprintf("Falha ao ler %s. Abra o arquivo e verifique se está conforme o Manual. Continuando com o próximo...", arquivo), 
               e)
    })
  }
  
  # Aplica a função em todos os arquivos usando lapply
  mapas_lidos <- lapply(mapas_arquivos, ler_um_arquivo)
  
  arquivos_validos <- mapas_arquivos[sapply(mapas_lidos, Negate(is.null))]
  
  # Remove entradas NULL (se houver falhas não fatais)
  mapas_lidos <- Filter(Negate(is.null), mapas_lidos)
  
  # Finaliza a barra de progresso
  log_barra_progresso(pb = pb)
  
  if (length(mapas_lidos) == 0) {
    log_erro("Nenhum arquivo lido. Encerrando...", 
             finalizar = TRUE)
  } else {
    return(list(dados    = mapas_lidos, 
                arquivos = arquivos_validos))
  }
}

mapas_obter <- function() {
  
  pasta <- get_config("pasta")
  
  tryCatch({
    #get .xls files
    mapas_arquivos <- list.files(path = file.path(pasta$mapas), 
                                 pattern = "^[^~]*.xls*", 
                                 full.names = TRUE, 
                                 recursive = TRUE)
    
    if (length(mapas_arquivos) == 0) stop("Não há Mapas na pasta.")
    
    return(mapas_arquivos)
  },
  error = function(e) {
    log_erro("Não foi possível obter a lista de arquivos. Verifique a pasta MAPAS. Encerrando...",
             e,
             finalizar = TRUE)
  })
  
}

mapas_salvar <- function(mapas, script_a_executar) {

  salvar_planilha <- function(dados, nome_arquivo) {
    
    pasta <- get_config("pasta")
    
    tryCatch({
      wb <- createWorkbook()
        addWorksheet(wb, "Mapa de Licitação")
        writeData(wb, "Mapa de Licitação", dados)
      saveWorkbook(wb, file.path(pasta$listas, nome_arquivo), overwrite = TRUE)
      
      log_info(sprintf("Arquivo salvo: %s", nome_arquivo),
               cores = "verde")
      
      return(TRUE)
    }, 
    error = function(e) {
      log_erro(sprintf("Não foi possível salvar o arquivo %s", nome_arquivo), 
               e)
      return(FALSE)
    }, 
    warning = function(w) {
      log_erro(sprintf("Não foi possível salvar o arquivo %s", nome_arquivo), 
               w)
      return(FALSE)
    })
  }
  
  options("openxlsx.dateFormat" = "dd/mm/yyyy")
  
  if (script_a_executar == "processo") {
    
    processos <- unique(mapas$processo)
    
    pb <- log_barra_progresso("Salvando...", length(processos))
    
    for (proc in processos) {
      
      #salva planilha com mapa da licitação organizado
      mapa    <- filter(mapas, processo == proc)
      spa     <- gsub("/", "-", str_sub(proc, 7, 17))
      pregao  <- gsub("/", "-", mapa$edital[1])
      arquivo <- sprintf("Mapa - SPA %s - Pregão %s.xlsx", 
                         spa,
                         pregao) 
      
      log_barra_progresso(sprintf("Salvando arquivo %s", arquivo), pb = pb)

      salvar_planilha(mapa, arquivo)
      
    }
    
    log_barra_progresso(pb = pb)
    
  } else {

    #salva planilha com mapa da licitação organizado
    grupo <- paste(sort(unique(str_sub(mapas$cod_item, 1, 6))), collapse = " ")
    
    if (nchar(grupo) > 20) {
      grupo <- paste(str_sub(grupo, 1, 20), "[...]")
    }
    
    arquivo <- sprintf("Mapa - Grupo %s.xlsx", grupo)

    salvar_planilha(mapas, arquivo)
  }
}

mapas_main <- function() {
  cat("=== INICIANDO SCRIPT. AGUARDE... ===\n")
  
  source(file.path("..", "_common", "config.R"), chdir = TRUE)
  
  pasta        = list()
  pasta$atual  = getwd()
  pasta$mapas  = file.path(pasta$atual, "MAPAS")
  pasta$listas = file.path(pasta$atual, "LISTAS")
  pasta$criar  = c(pasta$mapas, pasta$listas)
  
  pacotes     = c("openxlsx", "dplyr", "stringr")
  
  config_inicializar(pacotes, pasta)

  script_a_executar <- 
    utils_verificar_script(scripts_permitidos = c("processo", "grupo"),
                           script_padrao = "grupo")
  
  log_secao("OBTENDO LISTA DE ARQUIVOS DAS PASTA MAPAS")
  
  mapas_arquivos <- mapas_obter()

  log_secao("OBTENDO CONTEÚDO DOS MAPAS")
  
  mapas_lidos <- mapas_ler(mapas_arquivos)

  mapas_dados    <- mapas_lidos$dados
  mapas_arquivos <- mapas_lidos$arquivos
  
  log_secao("ANALISANDO INFORMAÇÕES")

  mapas_analisados <- mapas_analisar(mapas_dados, mapas_arquivos)
  
  log_secao(sprintf("SALVANDO ARQUIVO POR %s", toupper(script_a_executar)))
  
  mapas_salvar(mapas_analisados, script_a_executar)

  config_finalizar(sucesso = TRUE)
}

mapas_main()
