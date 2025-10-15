# ---- PRINCIPAL ----

power_bi_main <- function() {

  source(file.path("..", "_common", "config.R"), chdir = TRUE)

  pacotes = c("openxlsx2", "googlesheets4", "readxl", 
              "dplyr", "stringr", "purrr", "tidyr")
  
  config_inicializar(pacotes)
  
  log_secao("DEFINIÇÕES INICIAIS")
  
  script_a_executar <- 
    utils_verificar_script(
      scripts_permitidos = c("planejamento", "licitacao", "execucao", 
                             "paalteracoes", "renomear", "todos"),
      script_padrao = "execucao")
  
  if (script_a_executar != "renomear") utils <- power_bi_utils()
  
  switch(
    script_a_executar,
    "planejamento" = {
      power_bi_planejamento_main(utils, 
                                 script_a_executar)
    },
    "licitacao" = {
      mapas_dados <- power_bi_renomear(script_a_executar)
      
      power_bi_licitacao_main(utils, 
                              mapas_dados,
                              script_a_executar)
    },
    "execucao" = {
      resultado_renomear <- power_bi_renomear(script_a_executar)
      
      power_bi_execucao_main(utils, 
                             resultado_renomear,
                             script_a_executar)
    },
    "paalteracoes" = {
      power_bi_paalteracoes_main(utils,
                                 script_a_executar)      
    },
    "renomear" = {
      power_bi_renomear(script_a_executar)
    },
    "todos" = {
      power_bi_planejamento_main(utils, 
                                 script_a_executar)
      
      power_bi_paalteracoes_main(utils,
                                 script_a_executar)
      
      resultado_renomear <- power_bi_renomear(script_a_executar)
      
      if (!is.null(resultado_renomear)) {
        power_bi_execucao_main(utils, 
                               resultado_renomear,
                               script_a_executar)
        
        power_bi_licitacao_main(utils, 
                                resultado_renomear$mapas_dados,
                                script_a_executar)
      }
    }
  )  
  
  config_finalizar(sucesso = TRUE)

}

power_bi_utils <- function() {

  unidades <- utils_unidades_requerentes_obter("power bi")
  
  catalogo <- utils_catalogo()
  
  # Definir aqui o ano inicial do painel Visão Planejamento
  ano_inicial   <- 2021
  ano_corrente  <- format(Sys.Date(), "%Y")
  ano_licitacao <- seq(ano_inicial, ano_corrente)

  # Definir aqui o ano inicial do painel 
  # Visão Processos Administrativos e Alterações Contratuais
  # ANA CORINA, edite apenas o ano da linha a seguir :)
  ano_inicial_processos_administrativos <- 2020
  
  return(list(unidades      = unidades,
              catalogo      = catalogo,
              ano_licitacao = ano_licitacao,
              ano_inicial_processos_administrativos = 
                ano_inicial_processos_administrativos))
    
}

power_bi_salvar <- function(nome_arquivo, dados, abas, vazios) {
  
  pasta <- get_config("pasta")
  
  arquivo_a_salvar <- file.path(pasta$superior, 
                                sprintf("Dados Visão %s.xlsx", nome_arquivo))
  
  pb <- log_barra_progresso("Salvando arquivo para o Power BI. Aguarde...", 
                            length(dados) + 1)
  
  tryCatch({
    wb <- wb_workbook()
    
    for (i in seq_along(dados)) {
      
      wb <- wb_add_worksheet(wb, 
                             sheet = abas[i], 
                             tab_color = rainbow(length(dados))[i]) %>%
        wb_add_data(abas[i], 
                    dados[[i]], 
                    na.strings = vazios[i], 
                    col_names = TRUE)
      
      log_barra_progresso(sprintf("Salvando arquivo para o Power BI. Aba %s incluída", abas[i]), pb = pb)
      
    }
    
    gc(full = TRUE)
    rm(dados)
    
    log_barra_progresso("Finalizando...", pb = pb)
    wb_save(wb, file = arquivo_a_salvar)
    
    log_barra_progresso(pb = pb)
    
    log_info(sprintf("Arquivo %s salvo com sucesso", arquivo_a_salvar),
             cores = "verde")
  },
  error = function(e) { 
    log_erro("Erro ao salvar o arquivo para o Power BI. Encerrando...", 
             e,
             finalizar = TRUE)
  },
  warning = function(w) { 
    log_erro("Erro ao salvar o arquivo para o Power BI. Encerrando...", 
             w,
             finalizar = TRUE)
  })
  
}

# ---- RENOMEAR ARQUIVOS ----

power_bi_renomear <- function(script_a_executar) {
  
  obter_arquivos <- function(pasta) {
    
    pasta_superior <- get_config("pasta")$superior
    
    tryCatch({
      arquivos <- list.files(path = file.path(pasta_superior, pasta),
                             pattern = "^[^~]",
                             recursive = TRUE,
                             full.names = TRUE)
    },
    error = function(e) {
      log_erro(sprintf("Não foi possível obter a lista de arquivos da pasta %s. Encerrando...", pasta),
               e,
               finalizar = TRUE)
    })
    
    if (length(arquivos) == 0) {
      log_erro(sprintf("Nenhum arquivo encontrado na pasta %s. Encerrando...", pasta), 
               finalizar = TRUE)
    } else {
      return(arquivos)
    }
    
  }

  ler_arquivos <- function(arquivos, opcao) {
    
    pb <- log_barra_progresso("Aguarde...", length(arquivos))
    
    lista_arquivos <- lapply(arquivos, function(arquivo) {
      log_barra_progresso(sprintf("Lendo %s: %s", opcao, basename(arquivo)), pb = pb)
      
      tryCatch({
        if (opcao == "relatório de execução") {
          
          read_excel(arquivo, .name_repair = ~ make.names(., unique = TRUE)) %>%
            mutate(
              arquivo_relatorio = arquivo,
              pregao_relatorio  = Edital[1] 
            )
          
        } else if (opcao == "mapa") {
          
          read_excel(arquivo) %>% 
            mutate(
              mapas_arquivos        = arquivo,
              Ano                   = str_split_i(arquivo, "/", length(str_split_1(arquivo, "/")) - 2),
              Etapa                 = str_split_i(arquivo, "/", length(str_split_1(arquivo, "/")) - 1),
              pregao_mapa           = str_remove(Edital, "^0+"),
              arquivo_mapa_ajustado = paste0(dirname(arquivo),
                                             "/Mapa ", Ano, " - ", Etapa, 
                                             " - SPA ", str_sub(Processo, 7, 12), "-", str_sub(Processo, 14, 17),
                                             " - PE ", gsub("/", "-", pregao_mapa),
                                             ".xlsx")
            )
          
        }
      }, 
      error = function(e) {
        log_erro(sprintf("Erro ao processar o %s %s. Ignorando...", opcao, arquivo),
                 e,
                 alerta = TRUE)
        return(NULL)
      })
    })
    
    log_barra_progresso(pb = pb)
    
    arquivos_dados <- Filter(Negate(is.null), lista_arquivos)
    
    if (length(arquivos_dados) == 0) {
      log_erro("Nenhum dado válido para processar. Encerrando...", 
               finalizar = TRUE)
    } else {
      log_info(sprintf("Foram encontrados %d arquivos de %s válidos.", length(arquivos_dados), opcao),
               cores = "verde")
    }
    
    arquivos_dados <- bind_rows(arquivos_dados)
    
    return(arquivos_dados)
  }
  
  verificar_duplicidade <- function(dados, agrupar_por, verificar_por) {

    dados_duplicados <- dados %>%
      group_by({{agrupar_por}}) %>%
      summarise(qtd_arquivos = n_distinct({{verificar_por}}),
                arquivos = paste(unique({{verificar_por}}), collapse = " ←←← e →→→")) %>%
      filter(qtd_arquivos > 1) %>%
      select(-qtd_arquivos)
    
    if (nrow(dados_duplicados) > 0) {
      log_erro("Há dados duplicados em um ou mais arquivos:", 
               dados_duplicados, 
               alerta = TRUE)
      return(TRUE)
    } else {
      log_info("Nenhum arquivo duplicado.",
               cores = "verde")
      return(FALSE)
    }
    
  }
  
  conferir_pregoes <- function(mapas_dados, relatorios_dados) {
    
    controle_pregoes <- relatorios_dados %>%
      distinct(pregao_relatorio, arquivo_relatorio) %>%
      mutate(
        npregao = str_trim(str_sub(pregao_relatorio, 3)),
        processo = map_chr(
          npregao,
          ~ mapas_dados %>%
            filter(pregao_mapa == .x) %>%
            distinct(Processo) %>%
            pull(Processo) %>%
            first()),
        arquivo_mapa = map_chr(
          npregao,
          ~ mapas_dados %>%
            filter(pregao_mapa == .x) %>%
            distinct(arquivo_mapa_ajustado) %>%
            pull(arquivo_mapa_ajustado) %>%
            first()),
        arquivo_execucao_ajustado = 
          file.path(
            dirname(arquivo_relatorio),
            gsub("Mapa", "Execução", 
                 gsub("\\.xlsx$", ".xls", basename(arquivo_mapa))))      
      )

    pregoes_sem_mapa <- controle_pregoes %>% 
      filter(!complete.cases(.)) %>% 
      mutate(mensagem = paste(pregao_relatorio, "-", arquivo_relatorio)) %>% 
      select(pregao_relatorio, arquivo_relatorio, mensagem)
    
    if (nrow(pregoes_sem_mapa) > 0) {
      log_erro("Os seguintes pregões não possuem mapa correspondente e serão ignorados:", 
               pregoes_sem_mapa$mensagem, 
               alerta = TRUE)
      
      relatorios_dados <- relatorios_dados %>% 
        anti_join(pregoes_sem_mapa, by = "pregao_relatorio")
      
      controle_pregoes <- controle_pregoes %>% 
        filter(complete.cases(.))
    }    
    
    mapas_dados <- mapas_dados %>%
      mutate(executado = if_else(
        pregao_mapa %in% controle_pregoes$npregao, 
        "Sim", 
        "Não"
      ))
    
    relatorios_dados <- relatorios_dados %>%
      left_join(controle_pregoes %>% 
                  select(pregao_relatorio, processo),
                by = "pregao_relatorio")
    
    return(list(mapas_dados      = mapas_dados,
                relatorios_dados = relatorios_dados,
                controle_pregoes = controle_pregoes))
  }
  
  renomear_arquivos <- function(dados, coluna_antigo, coluna_novo, tipo_arquivo) {
    
    alteracoes <- dados %>%
      filter({{coluna_antigo}} != {{coluna_novo}}) %>%
      select(
        antigo = {{coluna_antigo}}, 
        novo   = {{coluna_novo}})
    
    if (nrow(alteracoes) > 0) {
      
      walk2(alteracoes$antigo, alteracoes$novo, ~{
        tryCatch({
          if (file.rename(.x, .y)) {
            cat(paste0(" →→→ ", tipo_arquivo, 
                       " RENOMEADO DE →→→ ", basename(.x), 
                       " PARA →→→ ", basename(.y), "\n"))
          }
        }, 
        warning = function(w) {
          log_erro(sprintf("Problemas ao renomear %s. Passando para o próximo...", tipo_arquivo),
                   w)
        })
      })
      
    } else {
      
      log_info(sprintf("Nenhum %s para ser renomeado.", tipo_arquivo), 
               cores = "vermelho")
      
    }

  }
  
  # ---- INICIANDO ----
  
  geral <- get_config("geral")
  geral$script_nome <- "POWER" # "PowerBI - Renomear"
  set_config(geral = geral)
  
  log_secao("OBTENDO LISTA DE MAPAS DE LICITAÇÕES")
  
  mapas <- obter_arquivos("Mapa de Licitações")
  
  log_secao("OBTENDO INFORMAÇÃO DOS MAPAS")
  
  mapas_dados <- ler_arquivos(mapas, "mapa")
  
  log_secao("VERIFICANDO MAPAS DUPLICADOS")

  mapas_duplicados      <- verificar_duplicidade(mapas_dados, Processo, mapas_arquivos)
  relatorios_duplicados <- FALSE

  if (script_a_executar != "licitacao") {
    
    log_secao("OBTENDO LISTA DE RELATÓRIOS DE EXECUÇÃO")
    
    relatorios <- obter_arquivos("Execução AF Empenho")
    
    log_secao("OBTENDO INFORMAÇÕES DOS RELATÓRIOS")
    
    relatorios_dados <- ler_arquivos(relatorios, "relatório de execução")
  
    log_secao("VERIFICANDO RELATORIOS DUPLICADOS")
    
    relatorios_duplicados <- verificar_duplicidade(relatorios_dados, pregao_relatorio, arquivo_relatorio)
    
  }
  
  if (mapas_duplicados || relatorios_duplicados) {
    log_erro("Arquivos duplicados encontrados. Encerrando...",
             "Corrija os arquivos e execute o script novamente.",
             finalizar = script_a_executar != "todos")
    return(NULL)
  }
  
  if (script_a_executar != "licitacao") {
    
    log_secao("CRUZANDO DADOS DOS MAPAS E DOS RELATÓRIOS")

    controle_pregoes <- conferir_pregoes(mapas_dados, relatorios_dados)
    mapas_dados      <- controle_pregoes$mapas_dados
    relatorios_dados <- controle_pregoes$relatorios_dados
    controle_pregoes <- controle_pregoes$controle_pregoes
    
  }
  
  log_secao("RENOMEANDO ARQUIVOS CONFORME PADRÃO")
  
  controle_mapas <- distinct(mapas_dados, mapas_arquivos, arquivo_mapa_ajustado)
  renomear_arquivos(controle_mapas, mapas_arquivos, arquivo_mapa_ajustado, "MAPA")
  
  if (script_a_executar != "licitacao") {
    renomear_arquivos(controle_pregoes, arquivo_relatorio, arquivo_execucao_ajustado, "RELATÓRIO")
  }
  
  switch(
    script_a_executar,
    "renomear"  = config_finalizar(),
    "licitacao" = return(mapas_dados),
                  return(list(mapas_dados = mapas_dados,
                              relatorios_dados = relatorios_dados))
  )  
}

# ---- PLANEJAMENTO ----

power_bi_planejamento_main <- function(utils, script_a_executar) {
  
  planilha_obter <- function(ano_licitacao) {
    
    url_planilha <- get_config("url")$controle
    
    log_info("Planilha: Controle de Processos DCOM",
             paste0("Link: ", url_planilha),
             paste0("Anos: ", paste(ano_licitacao, collapse = ", ")),
             cores = "verde")
    
    pb <- log_barra_progresso("Lendo planilha de Controle de Processos", length(ano_licitacao))
  
    nomes_colunas <- c("material", "processo", "data_inicio", 
                       "requerente", "pedido", 
                       "itens", "compartilhado", "itens_consolidados", 
                       "etapa", "situacao", "data_finalizacao")
      
    planilha_original <- lapply(ano_licitacao, function(ano) {
      
      log_barra_progresso(sprintf("Lendo planilha de Controle de Processos - Ano %s", ano), pb = pb)

      colunas       <- if (ano == 2021) "B2:W" else "B2:Y"
      tipos_colunas <- if (ano == 2021) "ccDccc__licc_________c" else "ccDcc_c__licc__________c"
        
      planilha_ano <- tryCatch({
          range_read(
            url_planilha, 
            sheet     = paste("Licitação", ano),
            col_names = TRUE,
            col_types = tipos_colunas,
            range     = colunas
          )
      },
      error = function(e) {
        log_erro(sprintf("Erro ao acessar aba %s da Planilha de Controle. Verifique se há permissão para leitura e se o link %s está correto. ", paste("Licitação", ano), url_planilha), 
                 e)
        return(NULL)
      })
      
      if (is.null(planilha_ano)) return(NULL)

      if (ncol(planilha_ano) != length(nomes_colunas)) {
        stop(sprintf("A aba Licitação %s não possui o número correto de colunas.", ano))
      } else {
        colnames(planilha_ano) <- nomes_colunas
      }
      
      planilha_ano <- planilha_ano[!is.na(planilha_ano$processo), ]
      
      if (nrow(planilha_ano) == 0) {
        stop(sprintf("A aba Licitação %s não possui dados.", ano))
      }
      
      return(planilha_ano)
    })
    
    log_barra_progresso(pb = pb)
    
    planilha_original <- Filter(Negate(is.null), planilha_original)
    
    return(if (length(planilha_original) == 0) NULL else as.list(planilha_original))
    
  }
  
  planilha_organizar <- function(planilha_original, ano_licitacao) {
    
    names(planilha_original) <- ano_licitacao
    planilha <- bind_rows(planilha_original, .id = "ano")
    
    planilha <- planilha %>%
      mutate(
        data_finalizacao = as.POSIXct(data_finalizacao, format = "%d/%m/%Y"),
        etapa = ifelse(
          etapa %in% c("Sem etapa", "Recondução"),
          etapa, 
          paste0(str_split_i(etapa, "/", 2), "/", str_split_i(etapa, "/", 1))
        )
      )
    
    return(planilha)
  }
  
  planilha_analisar <- function(planilha) {
    
    padronizar_colunas <- function(coluna, tamanho, preencher) {
      if(length(coluna) < tamanho) c(coluna, rep(preencher, tamanho - length(coluna))) else coluna
    }
    
    inconsistencias_colunas <- planilha %>%
      rowwise() %>%
      mutate(
        n_requerentes = length(strsplit(as.character(requerente), "\n")[[1]]),
        n_pedidos     = length(strsplit(as.character(pedido), "\n")[[1]]),
        n_itens       = length(strsplit(as.character(itens), "\n")[[1]]),
        consistente   = (n_requerentes == n_pedidos) && (n_pedidos == n_itens)
      ) %>%
      filter(!consistente) %>%
      ungroup() %>%
      select(processo, ano)
    
    if (nrow(inconsistencias_colunas) > 0) {
      msg_erro <- sprintf(" Processo %s em %s: diferença nas colunas de requerentes, pedidos e itens",
                          inconsistencias_colunas$processo, inconsistencias_colunas$ano)
      log_erro("Houve problemas na captura dos dados.", 
               unique(msg_erro),
               alerta = TRUE)
    }
    
    dados_observatorio <- planilha %>%
      mutate(
        requerente_lista = strsplit(as.character(requerente), "\n"),
        pedido_lista     = strsplit(as.character(pedido), "\n"),
        itens_lista      = strsplit(as.character(itens), "\n"),
        tamanho_maximo = pmax(
          lengths(requerente_lista),
          lengths(pedido_lista),
          lengths(itens_lista)
        )
      ) %>%
      select(-requerente, -pedido, -itens) %>%
      mutate(
        requerente = map2(requerente_lista, tamanho_maximo, padronizar_colunas, "ERRO"),
        pedido     = map2(pedido_lista, tamanho_maximo, padronizar_colunas, "ERRO"),
        itens      = map2(itens_lista, tamanho_maximo, padronizar_colunas, 0)
      ) %>%
      select(-requerente_lista, -pedido_lista, -itens_lista) %>%
      unnest(c(requerente, pedido, itens)) %>%
      mutate(
        itens            = suppressWarnings(as.numeric(itens)),
        data_inicio      = format(data_inicio, "%d/%m/%Y"),
        data_finalizacao = format(data_finalizacao, "%d/%m/%Y")
      ) %>%
      select(
        ano, material, processo, data_inicio, 
        requerente, pedido, 
        itens, compartilhado, itens_consolidados, 
        etapa, situacao, data_finalizacao
      )
    
    return(dados_observatorio)
  }
  
  planilha_resultados <- function(dados_observatorio) {
    
    resultado <- dados_observatorio %>% 
      group_by(processo, ano) %>%
      mutate(
        resultado = 
          case_when(
            situacao == "Inativo"                      ~ "Inativo",
            grepl("ok", tolower(pedido), fixed = TRUE) ~ "Enviado",
            grepl("CAPL", situacao, fixed = TRUE)      ~ "Aguardando envio ou em análise",
            .default                                   = "Não enviado"
          ),
        indice_otimizacao = 1 - (itens_consolidados / sum(itens, na.rm = TRUE))
      ) %>%
      ungroup()
    
    return(resultado)
  }
  
  # ---- INICIANDO ----
  
  geral <- get_config("geral")
  geral$script_nome <- "POWER" # "PowerBI - Planejamento"
  set_config(geral = geral)
  
  ano_licitacao <- utils$ano_licitacao
  unidades      <- utils$unidades
  
  log_secao("OBTENDO DADOS DA PLANILHA DE CONTROLE")
  
  planilha_original <- planilha_obter(ano_licitacao)
  
  if (is.null(planilha_original)) {
    log_erro("Não foi possível obter os dados da planilha de Controle de Processos DCOM. Encerrando...",
             finalizar = script_a_executar != "todos")
    return(NULL)
  }
  
  log_secao("ORGANIZANDO PLANILHA")
  
  planilha <- planilha_organizar(planilha_original, ano_licitacao)
  
  log_secao("ANALISANDO INFORMAÇÕES")
  
  dados_observatorio <- planilha_analisar(planilha)
  
  log_secao("IDENTIFICANDO SITUAÇÃO DOS PEDIDOS")
  
  resultado <- planilha_resultados(dados_observatorio)
  
  log_secao("SALVANDO ARQUIVO PARA O POWER BI")
  
  power_bi_salvar(nome_arquivo = "Planejamento",
                  dados        = list(resultado, 
                                      unidades),
                  abas         = c("Visão Planejamento", 
                                   "Unidades"),
                  vazios       = c("", 
                                   "Não informado"))
  
}

# ---- LICITAÇÃO ----

power_bi_licitacao_main <- function(utils, mapas_dados, script_a_executar) {
  
  verificar_resultados <- function(mapas_dados) {
  
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
    
    mapas_resultados <- mapas_dados %>%
      mutate(
        pregaoEdital = paste(Edital, Processo, sep = " - "),
        idUnico      = paste(Processo, Item, sep = "-"),
        idUnicoSetor = paste(Processo, Item, `Setor Entrega`, sep = "-"),
        grupo        = str_sub(`Cód. Item`, 1, 6),
        resultado    = case_when(
          Situação == "Aguardando certame" ~ "Não licitado",
          Edital == "000000/0" ~ "Não licitado",
          Vencedor %in% resultados$deserto ~ "Deserto",
          Vencedor %in% resultados$cancelado ~ "Cancelado",
          Vencedor %in% resultados$fracassado ~ "Fracassado",
          TRUE ~ "Sucesso") 
      ) %>%
      select(-any_of(c("Setor Solicitante", "Modalidade", "Situação", 
                       "mapas_arquivos", "pregao_mapa", "arquivo_mapa_ajustado")))
    
    return(mapas_resultados)
    
  }
  
  verificar_duplicados <- function(mapas_resultados) {
  # Verifica se há linhas duplicadas.
  # Já aconteceu de haver duas situações/valores para determinado item
  # Neste caso, se houver mais de um idUnicoSetor, escolhe apenas o com valor diferente de R$ 0,00
    
    duplicados <- mapas_resultados %>% 
      group_by(idUnicoSetor) %>% 
      filter(n() > 1) %>% 
      ungroup()
    
    if (nrow(duplicados) > 0) {
      log_erro(
        "Itens com informações duplicadas. Mantendo apenas os com valor homologado.", 
        distinct(duplicados, Processo, Item, `Setor Entrega`),
        alerta = TRUE)
      
      mapas_resultados <- mapas_resultados %>%
        group_by(idUnicoSetor) %>%
        filter(`Valor homologado` != "0,00" | n() == 1) %>%
        ungroup()
    }
    
    return(mapas_resultados)
    
  }

  # ---- INICIANDO ----
  
  geral <- get_config("geral")
  geral$script_nome <- "POWER" # "Power BI - Licitação"
  set_config(geral = geral)
  
  unidades <- utils$unidades
  catalogo <- utils$catalogo
  
  log_secao("VERIFICANDO RESULTADO DOS ITENS DOS MAPAS")
  
  mapas_resultados <- verificar_resultados(mapas_dados)
  
  log_secao("VERIFICANDO ITENS DUPLICADOS")
  
  mapas_dados_finais <- verificar_duplicados(mapas_resultados)
  
  log_secao("SALVANDO ARQUIVO PARA O POWER BI")
  
  power_bi_salvar(nome_arquivo = "Licitação", 
                  dados        = list(mapas_dados_finais, 
                                      unidades, 
                                      catalogo),
                  abas         = c("Mapa de licitações", 
                                   "Unidades", 
                                   "Catálogo"),
                  vazios       = c("", 
                                   "", 
                                   "Não informado"))

}

# ---- EXECUÇÃO ----

power_bi_execucao_main <- function(utils, resultado_renomear, script_a_executar) {
  
  dados_geral <- function(relatorios_dados, mapas_dados, catalogo) {
    
    dados_lista <- list()

    processos <- unique(relatorios_dados$processo)
    
    pb <- log_barra_progresso("Aguarde...", size = length(processos))
    
    for (processo_a_analisar in processos) {
      dados <- relatorios_dados %>%
        filter(processo == processo_a_analisar)
      npregao_raw <- dados$Edital[1]
      npregao     <- str_trim(str_sub(npregao_raw, 3))
      divisores   <- which(dados$Edital == npregao_raw)
      
      log_barra_progresso(sprintf("Analisando pregão %s", npregao), pb = pb)
      
      mapa <- mapas_dados %>% 
        filter(pregao_mapa == npregao) %>% 
        mutate(Item = str_remove(Item, "^0+"))

      for (i in seq_along(divisores)) {
        inicio <- divisores[i]
        fim    <- if (i == length(divisores)) nrow(dados) else divisores[i + 1] - 1
        
        dados_item  <- dados[inicio:fim, ]
        item_numero <- dados_item$Item[1]
        item_mapa   <- which(mapa$Item == item_numero)
        
        if (length(item_mapa) == 0) next
        
        dados_lista[[length(dados_lista) + 1]] <- list(
          id_unico        = paste(npregao, item_numero, sep = "-"),
          ano             = mapa$Ano[item_mapa][1],
          etapa           = mapa$Etapa[item_mapa][1],
          processo        = mapa$Processo[item_mapa][1],
          pregao          = npregao,
          pregao_processo = paste(npregao, mapa$Processo[item_mapa][1], sep = " - "),
          objeto          = mapa$Objeto[item_mapa][1],
          item            = item_numero,
          grupo           = catalogo %>% filter(Grupo == str_sub(mapa$`Cód. Item`[item_mapa][1], 1, 6)) %>% pull(grupoDescricao),
          un_medida       = mapa$`Unidade de medida`[item_mapa][1],
          descricao       = mapa$`Descrição resumida`[item_mapa][1],
          especif         = mapa$Especificação[item_mapa][1],
          detalh          = mapa$Detalhamento[item_mapa][1],
          qt_total        = utils_corrigir_valor(dados_item$`Qtde.Licitação`[1]),
          valor           = utils_corrigir_valor(mapa$`Valor homologado`[item_mapa][1]),
          fornecedor      = mapa$Vencedor[item_mapa][1]
        )
      }
    }
    
    log_barra_progresso(pb = pb)
    
    dados_geral <- bind_rows(dados_lista)

    return(dados_geral)
  }
  
  processar_relatorios <- function(relatorios_dados, tipo) {
    
    config <- switch(tipo,
                     "Saldos" = list(
                       coluna_id = 2,
                       cabecalho = NULL,
                       start_offset = 5,
                       colunas = list(
                         id_unico = "id_unico",
                         pregao = "pregao",
                         item = "item",
                         setor = 2,
                         qt_ata = 3,
                         qt_empenhada = 4
                       )
                     ),
                     "AFs" = list(
                       coluna_id = 2,
                       cabecalho = "Protocolo",
                       start_offset = 1,
                       colunas = list(
                         id_unico = "id_unico",
                         pregao = "pregao",
                         item = "item",
                         af = 2,
                         setor_af = 3,
                         data_af = 4,
                         qt_af = 6,
                         situacao_af = 7
                       )
                     ),
                     "Empenhos" = list(
                       coluna_id = 2,
                       cabecalho = "Solicitação Empenho",
                       start_offset = 1,
                       colunas = list(
                         id_unico = "id_unico",
                         pregao = "pregao",
                         item = "item",
                         sne = 2,
                         spa_sne = 3,
                         unidade_sne = 4,
                         data_sne = 5,
                         qt_sne = 7
                       )
                     ),
                     stop("Tipo inválido. Escolha 'saldos', 'af' ou 'empenhos'.")
    )
    
    resultados_lista <- list()
    
    processos <- unique(relatorios_dados$processo)
    
    pb <- log_barra_progresso("Aguarde...", length(processos))
    
    for (processo_a_analisar in processos) {
      dados_orig <- relatorios_dados %>%
        filter(processo == processo_a_analisar)
      
      npregao <- as.character(dados_orig[1, 1])
      if (is.na(npregao) || nchar(trimws(npregao)) == 0) {
        log_barra_progresso(sprintf("Pregão inválido no relatório do processo %s, pulando...", processo_a_analisar), pb = pb)
        next
      }
      
      pregao_base <- str_trim(str_sub(npregao, 3))
      divisores <- which(dados_orig[, 1] == npregao)
      
      if (length(divisores) == 0) next
      
      log_barra_progresso(sprintf("Processando pregão %s (%s)", npregao, tipo), pb = pb)
      
      for (i in seq_along(divisores)) {
        inicio <- divisores[i]
        fim <- ifelse(i == length(divisores), nrow(dados_orig), divisores[i + 1] - 1)
        if (inicio > fim) next
        
        dados_item <- dados_orig[inicio:fim, , drop = FALSE]
        item_val <- dados_item[1, 2]
        if (is.na(item_val) || nchar(trimws(item_val)) == 0) next
        
        id_unico_val <- paste(pregao_base, item_val, sep = "-")
        
        # Encontra linha de início dos dados
        start_row <- if (is.null(config$cabecalho)) {
          config$start_offset
        } else {
          cabecalho_row <- which(dados_item[, config$coluna_id] == config$cabecalho)
          if (length(cabecalho_row) == 0) next
          cabecalho_row[1] + config$start_offset
        }
        
        if (start_row > nrow(dados_item)) next
        
        # Processa cada linha de dados
        for (j in start_row:nrow(dados_item)) {
          celula <- dados_item[j, config$coluna_id]
          if (is.na(celula) || trimws(celula) == "") break
          
          linha <- list(
            id_unico = id_unico_val,
            pregao = pregao_base,
            item = as.character(item_val)
          )
          
          # Adiciona colunas específicas
          for (nome in names(config$colunas)) {
            if (nome %in% c("id_unico", "pregao", "item")) next
            col <- config$colunas[[nome]]
            linha[[nome]] <- as.character(dados_item[j, col])
          }
          
          resultados_lista[[length(resultados_lista) + 1]] <- linha
        }
      }
    }
    
    log_barra_progresso(pb = pb)
    
    return(bind_rows(resultados_lista))
  }
  
  # ---- INICIANDO ----
  
  geral <- get_config("geral")
  geral$script_nome <- "POWER" # "PowerBI - Execução"
  set_config(geral = geral)
  
  unidades <- utils$unidades
  catalogo <- utils$catalogo
  
  relatorios_dados <- resultado_renomear$relatorios_dados
  mapas_dados      <- resultado_renomear$mapas_dados
  
  log_secao("GERAL - ANALISANDO DADOS DOS ARQUIVOS")
  
  dados_geral <- dados_geral(relatorios_dados, mapas_dados, catalogo)
  
  log_secao("SALDOS - ANALISANDO DADOS DOS ARQUIVOS")

  dados_saldos <- processar_relatorios(relatorios_dados, tipo = "Saldos")
  
  log_secao("AFs - ANALISANDO DADOS DOS ARQUIVOS")
  
  dados_af <- processar_relatorios(relatorios_dados, tipo = "AFs")
  
  log_secao("EMPENHOS - ANALISANDO DADOS DOS ARQUIVOS")
  
  dados_empenhos <- processar_relatorios(relatorios_dados, tipo = "Empenhos")
  
  log_info("QUANTIDADE DE MAPAS DE LICITAÇÕES", length(unique(mapas_dados$Processo)),
           "-",
           "QUANTIDADE DE PREGOES FINALIZADOS", length(unique(relatorios_dados$pregao_relatorio)),
           cores = "verde")
  
  log_secao("SALVANDO ARQUIVO PARA O POWER BI")
  
  gc(full = TRUE)
  rm(relatorios_dados)
  
  mapas_dados <- mapas_dados %>% 
    distinct(pregao_mapa, Processo, Ano, Etapa, executado) %>%
    rename(processo = Processo,
           ano = Ano,
           etapa = Etapa)
  
  power_bi_salvar(nome_arquivo = "Execução", 
                  dados        = list(dados_geral, 
                                      dados_saldos, 
                                      dados_af, 
                                      dados_empenhos, 
                                      mapas_dados, 
                                      unidades),
                  abas         = c("Geral", 
                                   "Saldos", 
                                   "AFs", 
                                   "Empenhos", 
                                   "Mapa de licitações", 
                                   "Unidades"),
                  vazios       = c("Não disponível", 
                                   "Não disponível", 
                                   "Não disponível", 
                                   "Não disponível", 
                                   "-", 
                                   "Não informado"))
  
}

# ---- PROCESSOS ADMINISTRATIVOS E ALTERAÇÕES CONTRATUAIS ----

power_bi_paalteracoes_main <- function(utils, script_a_executar) {
  
  planilha_obter <- function(ano_inicial_processos_administrativos) {
    
    url_planilha <- get_config("url")$saa
    
    log_info("Planilha: Processos SAA/DCOM",
             paste0("Link: ", url_planilha),
             cores = "verde")
    
    abas <- c("Processos Administrativos", "Troca de Marca", "Cancelamento", "Reequilíbrio")
    
    pb <- log_barra_progresso("Lendo planilha de Processos SAA/DCOM", length(abas))
    
    planilha_original <- lapply(abas, function(aba) {
      
      log_barra_progresso(sprintf("Lendo planilha de Processos SAA/DCOM - Aba %s", aba), pb = pb)
      
      if (aba == "Processos Administrativos") {
        colunas <- "A:L"
        tipos_coluna <- "cccccccDcccc"
      } else {
        colunas <- "A2:F"
        tipos_coluna <- "cccccc"
      }
        
      planilha_aba <- tryCatch({
          range_read(
            url_planilha,
            sheet = aba,
            range = colunas,
            col_names = TRUE,
            col_types = tipos_coluna
          )
      },
      error = function(e) {
        log_erro(sprintf("Erro ao acessar aba %s de Processos SAA/DCOM. Verifique se há permissão para leitura e se o link %s está correto. ", aba, url_planilha), 
                 e)
        return(NULL)
      })
      
      if (is.null(planilha_aba)) return(NULL)
      
      if (aba == "Processos Administrativos") {
        planilha_aba <- planilha_aba %>%
          filter(!is.na(`Nº. PROCESSO`),
                 ABERTURA >= as.Date(paste0(ano_inicial_processos_administrativos, "-01-01")))
        
        # Criar mapeamento de fornecedores para IDs sequenciais
        cnpj_unicos <- unique(planilha_aba$CNPJ)
        fornecedor_map <- data.frame(
          CNPJ          = cnpj_unicos,
          FORNECEDOR_ID = sprintf("Fornecedor %04d", seq_along(cnpj_unicos))
        )
        
        # Aplicar as transformações para sigilo
        planilha_aba <- planilha_aba %>%
          left_join(fornecedor_map, by = "CNPJ") %>%
          mutate(
            FORNECEDOR = FORNECEDOR_ID,
            `Nº. PROCESSO` = str_replace(`Nº. PROCESSO`, "^(.+?)/", "23080.******/"),
            CNPJ = str_replace_all(CNPJ, "(?<=\\.)\\d{3}", "***")
          ) %>%
          select(-FORNECEDOR_ID)
        
      } else {
        
        planilha_aba <- planilha_aba %>%
        filter(!is.na(Processo))
        
      }
      
      if (nrow(planilha_aba) == 0) {
        stop(sprintf("Não foi possível recuperar os dados da aba %s. Encerrando...", aba))
      }
      
      return(planilha_aba)
    })
    
    log_barra_progresso(pb = pb)
    
    names(planilha_original) <- abas
    
    planilha_original <- Filter(Negate(is.null), planilha_original)
    
    return(if (length(planilha_original) != 4) NULL else as.list(planilha_original))
    
  }
  
  # ---- INICIANDO ----
  
  geral <- get_config("geral")
  geral$script_nome <- "POWER" # "PowerBI - PAs e alterações"
  set_config(geral = geral)
  
  ano_inicial_processos_administrativos <- utils$ano_inicial_processos_administrativos
  unidades                              <- utils$unidades
  
  log_secao("OBTENDO DADOS DA PLANILHA DE PROCESSOS SAA")
  
  planilha_original <- planilha_obter(ano_inicial_processos_administrativos)
  
  if (is.null(planilha_original)) {
    log_erro("Não foi possível obter os dados da planilha de Processos SAA/DCOM.",
             finalizar = script_a_executar != "todos")
    return(NULL)
  }
  
  log_secao("SALVANDO ARQUIVO PARA O POWER BI")
  
  power_bi_salvar(nome_arquivo = "SAA", 
                  dados        = list(planilha_original$`Processos Administrativos`,
                                      planilha_original$`Troca de Marca`,
                                      planilha_original$Cancelamento,
                                      planilha_original$Reequilíbrio,
                                      unidades),
                  abas         = c("Processos Administrativos", 
                                   "Troca de Marca", 
                                   "Cancelamento", 
                                   "Reequilíbrio", 
                                   "Unidades"),
                  vazios       = c("Não informado", 
                                   "Não informado", 
                                   "Não informado", 
                                   "Não informado", 
                                   "Não informado"))
  
}

# ---- INICIAR ----

power_bi_main()