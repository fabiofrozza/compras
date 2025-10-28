# OBTENÇÃO DOS DADOS DA LISTA FINAL====

main_dados_obter <- function(script_a_executar) {
# Função principal para baixar a Lista Final do Google Drive,
# filtrar as informações, gravar estatísticas, limpar dados
# e organizar os dataframes para a geração dos arquivos para importar
# ou para os resumos da importação ou relatórios gerenciais
  
  log_secao("BAIXANDO DADOS DO GOOGLE DRIVE", "MAIN")
  
  lista_final <- main_lista_final_obter()
  
  log_secao("ORGANIZANDO DADOS", "MAIN")
 
  dados_filtrados <- main_dados_filtrar(lista_final) 

  log_secao("AJUSTANDO PLANILHA", "MAIN")
  
  lista_final$planilha <- main_dados_limpar(dados_filtrados$planilha)
  
  log_secao("MONTANDO ARQUIVO", "MAIN")
 
  dados_montados <- main_dados_montar(lista_final$planilha, lista_final$unidades, lista_final$info, script_a_executar)
  
  if (is.null(dados_montados)) {
    log_erro("Não foi possível montar o arquivo. Encerrando...",
             finalizar = TRUE)
  } else {
    dados <- 
      append(dados_montados, 
             list(
               unidades = lista_final$unidades,
               original = lista_final$ajustada,
               filtros  = dados_filtrados$filtros, 
               stats    = dados_filtrados$stats))
    
    return(dados)
  }
}

main_dados_filtrar <- function(lista_final) {
  
  planilha_controle_obter <- function() {
  # Função para obter os dados da Planilha de Controle
  # Será verificada a aba Licitação do ano corrente
  # Serão obtidos as informações das Unidades e pedidos/SDs
    
    #link da planilha de Controle de Processos DCOM
    url_planilha <- get_config("url")$controle
    
    log_info("Planilha: Controle de Processos DCOM",
             paste0("Link: ", url_planilha),
             cores = "verde")
    
    # Definido o ano corrente para posteriormente ler a aba Licitação deste ano
    ano_corrente  <- format(Sys.Date(), "%Y")
    nomes_colunas <- c("processo", "requerente", "pedido")
    colunas       <- "C2:F"
    tipos_colunas <- "c_cc"
    
    tryCatch({
      
      # Lê a planilha
      planilha_controle <- 
        range_read(
          url_planilha, 
          sheet     = paste("Licitação", ano_corrente),
          col_names = TRUE,
          col_types = tipos_colunas,
          range     = colunas
        )
      
      # Se a planilha lida não tiver as colunas informadas acima, gera erro
      if (ncol(planilha_controle) != length(nomes_colunas)) {
        stop(sprintf("A aba Licitação %s não possui o número correto de colunas.", ano_corrente))
      } else {
        colnames(planilha_controle) <- nomes_colunas
      }
      
      # Mantém apenas as linhas em que o número do processo não esteja ausente
      planilha_controle <- planilha_controle %>% 
        filter(!is.na(processo))
      
      # Se não retornar nenhuma linha, gera erro
      if (nrow(planilha_controle) == 0) {
        stop(sprintf("A aba Licitação %s não possui dados.", ano_corrente))
      }
      
      return(planilha_controle)
    },
    error = function(e) {
      log_erro("Erro ao acessar Planilha de Controle. Verifique se há permissão para leitura e se o link está correto.", 
               e)
      return(NULL)
    })
    
  }
  
  planilha_controle_analisar <- function(planilha_controle_original, planilha, unidades) {
  # Função para obter dados da planilha de controle
    
    # Função auxiliar para igualar número de linhas das colunas Requerentes, Pedidos/SD e Itens
    padronizar_colunas <- function(coluna, tamanho, preencher) {
      if(length(coluna) < tamanho) c(coluna, rep(preencher, tamanho - length(coluna))) else coluna
    }
    
    # Obtém os responsáveis pela orçamentação conforme a coluna da Lista Final
    responsaveis_orcamentacao <- planilha %>%
      filter(!is.na(Processo)) %>%
      distinct(Processo, `Responsável pela Pesquisa`) %>%
      arrange(Processo, `Responsável pela Pesquisa`) %>%
      rename(processo = Processo, Sigla = `Responsável pela Pesquisa`) %>%
      mutate(tipo = "Pedido")
    
    # Organiza a planilha de controle, cruzando dados da Lista Final
    # Se a Unidade na planilha de controle está nos responsáveis pela orçamentação na Lista Final
    # registra como pedido, senão SD
    # Se tem OK no Pedido/SD na planilha de controle, informa como enviado, senão não
    planilha_controle <- planilha_controle_original %>%
      filter(processo %in% unique(planilha$Processo)) %>%
      mutate(
        requerente_lista = strsplit(as.character(requerente), "\n"),
        pedido_lista     = strsplit(as.character(pedido), "\n"),
        tamanho_maximo = pmax(
          lengths(requerente_lista),
          lengths(pedido_lista)
        )
      ) %>%
      select(-requerente, -pedido) %>%
      mutate(
        requerente = map2(requerente_lista, tamanho_maximo, padronizar_colunas, "ERRO"),
        pedido     = map2(pedido_lista, tamanho_maximo, padronizar_colunas, "ERRO"),
      ) %>%
      select(-requerente_lista, -pedido_lista) %>%
      unnest(c(requerente, pedido)) %>%
      left_join(
        unidades %>% select(Sigla, sigla_solar),
        by = c("requerente" = "sigla_solar")
      ) %>%
      left_join(
        responsaveis_orcamentacao,
        by = c("processo", "Sigla")
      ) %>%
      mutate(
        enviado = case_when(
          grepl("ok", tolower(pedido), fixed = TRUE) ~ TRUE,
          .default = FALSE
        ),
        tipo = ifelse(is.na(tipo), "SD", tipo)
      ) %>%
      mutate(
        pedido = str_extract(pedido, "\\b\\d+/\\d{2,4}\\b"),
        oficio = "",
        peculiaridades = "",
        etp = "",
        equipe_apoio = ""
      ) %>%
      select(
        processo, Sigla, requerente, pedido, enviado, tipo, oficio, peculiaridades, etp, equipe_apoio
      )
    
    return(planilha_controle)
  }
  
  filtro_planilha_controle <- function(filtros, planilha, unidades, script_a_executar, importacao) {
  # Se estiver sendo gerado o relatório após a geração dos processos
  # lê a planilha de controle para verificar pedidos enviados ou não
    
    if (script_a_executar == "relatorio" & importacao$pos_processos) {
      
      planilha_controle_original <- planilha_controle_obter()
      filtros$planilha_controle  <- planilha_controle_analisar(planilha_controle_original, planilha, unidades)
    }
    
    return(filtros)
  }
  
  filtro_processo_especifico <- function (filtros, planilha, script_a_executar, importacao) {
  # Se não for relatório consolidado, filtra a planilha (Lista Final) e planilha de controle
  # com o processo que está sendo analisado
    
    if (!importacao$consolidado) {
      planilha <- planilha %>% 
        filter(Processo == importacao$processo_para_relatorio)
      
      if (script_a_executar == "relatorio" & importacao$pos_processos) {
        filtros$planilha_controle <- filtros$planilha_controle %>% 
          filter(processo == importacao$processo_para_relatorio)
        }
    }
    
    return(list(planilha = planilha,
                filtros  = filtros))

  }

  filtro_estatisticas <- function(filtros, stats, planilha, importacao) {
    
    demandas_unidades          <- planilha[ , importacao$coluna_inicial:importacao$coluna_final_unidades]
    contagem_demandas_unidades <- colSums(!is.na(demandas_unidades))
    
    filtros$unidades_com_demanda <- sort(contagem_demandas_unidades[which(contagem_demandas_unidades != 0)], decreasing = TRUE)
    filtros$unidades_sem_demanda <- contagem_demandas_unidades[which(contagem_demandas_unidades == 0)]
    
    stats$unidades_com_demanda <- length(filtros$unidades_com_demanda)
    stats$unidades_sem_demanda <- length(filtros$unidades_sem_demanda)
    
    stats$itens_lista_final <- sum(!is.na(planilha$`Código do item`))
    
    return(list(filtros = filtros,
                stats   = stats))
    
  }
  
  filtro_orcamentos_nao_enviados <- function(filtros, stats, planilha, script_a_executar, importacao) {
  
    # Se estiver sendo gerado o relatório após a geração dos processos
    # verifica orçamentos não enviados
    if (script_a_executar == "relatorio" & importacao$pos_processos) {

      # Transforma Lista Final em formato longo (uma linha para cada item e cada Unidade)      
      subset_planilha_ajustes_exclusao <- planilha %>%
        pivot_longer(
          cols           = all_of(names(planilha)[importacao$coluna_inicial:importacao$coluna_final_unidades]),
          names_to       = "Unidade",
          values_to      = "Quantitativo",
          values_drop_na = TRUE
        )
    
      # Registra Unidades que tem pedido na planilha de controle e não enviaram documentação
      filtros$orcamentos_nao_enviados <- subset_planilha_ajustes_exclusao %>%
        right_join(
          filtros$planilha_controle %>%
            group_by(Sigla) %>%
            filter(tipo == "Pedido" & !enviado) %>%
            select(Sigla, pedido), 
          by = c("Responsável pela Pesquisa" = "Sigla")) %>%
        select(Processo, pedido, `Responsável pela Pesquisa`, `Código do item`, `Descrição Resumida`, linha_planilha, Unidade, Quantitativo) %>%
        mutate(Quantitativo = suppressWarnings(as.numeric(Quantitativo)),
               Quantitativo = ifelse(is.na(Quantitativo), 0, Quantitativo))
      
      # Registra estatísticas dos orçamentos não enviados
      stats$unidades_nao_orcamentos <- 
        length(unique(filtros$orcamentos_nao_enviados$`Responsável pela Pesquisa`))
      stats$qtde_itens_orcamentos_nao_enviados <- 
        length(unique(filtros$orcamentos_nao_enviados$linha_planilha))
      stats$qtde_unidades_afetadas_orcamentos <- 
        length(unique(filtros$orcamentos_nao_enviados$Unidade))
      
      # Exibe o resultado do filtro
      filtro_exibir(filtros$orcamentos_nao_enviados)
    }
    
    return(list(filtros = filtros, 
                stats   = stats))
    
  }
  
  filtro_sem_demanda <- function(filtros, stats, planilha) {
    
    # Registra os itens que não tiveram demanda
    filtros$sem_demanda <- planilha %>%
      filter(`Qtd. Total` == 0 & `Código do item` != "") %>%
      select(linha_planilha, `Código do item`, `Descrição Resumida`)
    
    # Registra, para estatística, os itens com e sem demanda
    stats$itens_sem_demanda <- nrow(filtros$sem_demanda)
    stats$itens_com_demanda <- stats$itens_lista_final - stats$itens_sem_demanda
    
    # Exclui do dataframe principal as linhas com a quantidade igual à zero (sem demanda)
    planilha <- planilha %>%
      filter(`Qtd. Total` > 0 & `Código do item` != "")
    
    # Exibe resultado do filtro
    filtro_exibir(filtros$sem_demanda)
    
    return(list(filtros  = filtros,
                stats    = stats,
                planilha = planilha))
  }
  
  filtro_abaixo_qtde <- function(filtros, stats, planilha, importacao) {
    
    # Registra os itens com valor total menor que o mínimo informado
    filtros$abaixo_qtde <- planilha %>%
      filter(`Qtd. Total` < importacao$qtde_minima) %>%
      select(linha_planilha, `Nº Item`, `Qtd. Total`, Processo, `Responsável pela Pesquisa`, `Código do item`, `Descrição Resumida`)
    
    # Registra, para estatística, a quantidade de itens abaixo do valor mínimo
    stats$abaixo_qtde <- nrow(filtros$abaixo_qtde)
    
    # Excluir do dataframe principal as linhas com a quantidade total menor que a mínima definida
    planilha <- planilha %>%
      filter(planilha$`Qtd. Total` >= importacao$qtde_minima)
    
    # Exibe resultado do filtro
    filtro_exibir(filtros$abaixo_qtde)
    
    return(list(filtros  = filtros,
                stats    = stats,
                planilha = planilha))
  }
  
  filtro_valor_minimo <- function(filtros, stats, planilha, importacao) {
  
    # Verifica se há algum problema com a coluna Valor total
    # Se não houver valor numérico, gera erro e continua
    valor_total_com_problemas <- planilha %>%
      filter(is.na(`Valor total`)) %>%
      pull(linha_planilha) %>%
      paste(collapse = ", ") %>%
      noquote()
    
    if (nchar(valor_total_com_problemas) > 0) {
      log_erro(sprintf("Há algum problema na coluna 'Valores Totais' na(s) linha(s) %s. Verifique a planilha.",
                       valor_total_com_problemas))
    }

    # Registra informações dos itens abaixo do valor mínimo
    filtros$abaixo_valor <- planilha %>%
      filter(`Valor total` > 0, `Valor total` < importacao$valor_minimo) %>%
      select(linha_planilha, `Nº Item`, `Qtd. Total`, Processo, `Responsável pela Pesquisa`, 
             `Código do item`, `Descrição Resumida`, `Valor unitário`, `Valor total`) %>%
      mutate(`Valor total` = paste0("R$ ", formatC(`Valor total`, format = "f", digits = 2, 
                                                   big.mark = ".", decimal.mark = ",")))
    
    # Registra estatística da quantidade de itens abaixo do valor mínimo
    stats$abaixo_valor <- nrow(filtros$abaixo_valor)
  
    # Mantém apenas linhas com o valor total igual a 0 ou maior ou igual ao mínimo definido
    planilha <- planilha %>%
      filter(`Valor total` == 0 | `Valor total` >= importacao$valor_minimo)
    
    # Exibe resultado do filtro
    filtro_exibir(filtros$abaixo_valor)
    
    return(list(filtros  = filtros,
                stats    = stats,
                planilha = planilha))
  
  }
  
  filtro_ajustes_exclusoes <- function(filtros, stats, planilha, importacao) {
  # Função para verificar ajustes (solicitação pela Unidade ou DCOM, 0 na planilha)
  # ou exclusões (desistência ou não envio de documentos, X na planilha)
    
    # Transforma os dados da Lista Final no formato longo
    # Uma linha para cada item para cada Unidade com demanda
    subset_planilha_ajustes_exclusao <- planilha %>%
      pivot_longer(
        cols           = all_of(names(planilha)[importacao$coluna_inicial:importacao$coluna_final_unidades]),
        names_to       = "Unidade",
        values_to      = "Quantitativo",
        values_drop_na = TRUE
      )
    
    # Verifica ajustes (0 na planilha)
    filtros$com_ajustes <- subset_planilha_ajustes_exclusao %>%
      filter(Quantitativo == "0") %>%
      group_by(`Processo`, `Nº Item`, linha_planilha, `Código do item`, `Descrição Resumida`, Unidade) %>%
      summarise(Quantidade_0 = n(), .groups = 'drop')
    
    stats$unidades_com_ajustes <- length(unique(filtros$com_ajustes$Unidade))
    stats$qtde_ajustes         <- sum(filtros$com_ajustes$Quantidade_0)
    
    # Verifica exclusões (X na planilha)
    filtros$com_demanda_excluida <- subset_planilha_ajustes_exclusao %>%
      filter(Quantitativo == "X" | Quantitativo == "x") %>%
      group_by(`Processo`, `Nº Item`, linha_planilha, `Código do item`, `Descrição Resumida`, Unidade) %>%
      summarise(Quantidade_X = n(), .groups = 'drop')
    
    stats$unidades_com_demanda_excluida <- length(unique(filtros$com_demanda_excluida$Unidade))
    stats$qtde_demanda_excluida         <- sum(filtros$com_demanda_excluida$Quantidade_X)
  
    return(list(filtros = filtros,
                stats   = stats))
  }
  
  filtro_excluidos <- function(filtros, stats, planilha) {
    
    # Registra os itens excluídos manualmente (marcados na coluna F da planilha)
    filtros$excluidos <- planilha %>%
      filter(`Item excl.` == "TRUE") %>%
      select(
        linha_planilha, `Nº Item`, `Qtd. Total`, Processo, `Responsável pela Pesquisa`,
        `Código do item`, `Descrição Resumida`, `Valor total`
      ) %>%
      mutate(`Valor total` = paste0("R$ ", formatC(`Valor total`, 
                                                   format = "f", digits = 2, 
                                                   big.mark = ".", decimal.mark = ",")))
    
    # Registra para estatística o número de itens excluídos manualmente
    stats$excluidos <- nrow(filtros$excluidos)
    
    # Elimina itens excluídos (coluna F da planilha)
    planilha <- planilha %>%
      filter(`Item excl.` == "FALSE")
    
    # Exibe resultado do filtro
    filtro_exibir(filtros$excluidos)
  
    return(list(filtros  = filtros,
                stats    = stats,
                planilha = planilha))
    
  }
  
  filtro_exibir <- function(filtro) {
  # Função para exibir resultado do filtro
    
    if (nrow(filtro) > 0) {
      print.data.frame(filtro, right = FALSE, row.names = FALSE)
    } else {
      cat("  ˣˣˣ\n")
    }
    
  }
  
  # Função para filtrar os dados da planilha (Lista Final)
  
  # Define variáveis
  planilha          <- lista_final$ajustada
  unidades          <- lista_final$unidades
  script_a_executar <- config_json("script_a_executar", opcao = "get")
  stats             <- list()
  filtros           <- list()
  
  importacao <- get_config("importacao")
  
  importacao$pos_processos    <- ifelse(sum(!is.na(planilha$Processo)) == 0, FALSE, TRUE)
  importacao$validacao_manual <- any(planilha$'Responsável pela Pesquisa' == 'VALIDAÇÃO MANUAL')

  # Define a coluna final das Unidades
  # Coluna inicial + quantidade de Unidades
  # Isto já somaria uma Unidade a mais, e excluindo UFSC GERAL, diminui-se 2
  importacao$coluna_final_unidades <- importacao$coluna_inicial + importacao$qtde_unidades - 2

  set_config(importacao = importacao)
  
  # Inicia filtragem da lista final
  
  log_secao("OBTENDO DADOS DA PLANILHA DE CONTROLE", "FILTROS")
  
  filtros <- filtro_planilha_controle(filtros, planilha, unidades, script_a_executar, importacao)
  
  log_secao("VERIFICANDO PROCESSO(S) A ANALISAR E ESTATÍSTICAS", "FILTROS")
  
  filtro_resultado <- filtro_processo_especifico(filtros, planilha, script_a_executar, importacao)
  planilha <- filtro_resultado$planilha
  filtros  <- filtro_resultado$filtros

  filtro_resultado <- filtro_estatisticas(filtros, stats, planilha, importacao)
  filtros <- filtro_resultado$filtros
  stats   <- filtro_resultado$stats
  
  log_secao("VERIFICANDO ORÇAMENTOS NÃO ENVIADOS", "FILTROS")
  
  filtro_resultado <- filtro_orcamentos_nao_enviados(filtros, stats, planilha, script_a_executar, importacao)
  filtros <- filtro_resultado$filtros
  stats   <- filtro_resultado$stats

  log_secao("SEM DEMANDA (QUANTIDADE IGUAL À ZERO)", "FILTROS")
  
  filtro_resultado <- filtro_sem_demanda(filtros, stats, planilha)
  filtros  <- filtro_resultado$filtros
  stats    <- filtro_resultado$stats
  planilha <- filtro_resultado$planilha
  
  log_secao(paste0("QUANTIDADE DEMANDADA MENOR QUE ", importacao$qtde_minima), "FILTROS")

  filtro_resultado <- filtro_abaixo_qtde(filtros, stats, planilha, importacao)
  filtros  <- filtro_resultado$filtros
  stats    <- filtro_resultado$stats
  planilha <- filtro_resultado$planilha
  
  log_secao(paste0("VALOR MÍNIMO (R$ ", importacao$valor_minimo, ") NÃO ATINGIDO"), "FILTROS")
  
  filtro_resultado <- filtro_valor_minimo(filtros, stats, planilha, importacao)
  filtros  <- filtro_resultado$filtros
  stats    <- filtro_resultado$stats
  planilha <- filtro_resultado$planilha
  
  log_secao("VERIFICANDO AJUSTES E EXCLUSÕES", "FILTROS")

  filtro_resultado <- filtro_ajustes_exclusoes(filtros, stats, planilha, importacao)
  filtros <- filtro_resultado$filtros
  stats   <- filtro_resultado$stats
  
  log_secao("EXCLUÍDOS MANUALMENTE NA PLANILHA (MARCADO NA COLUNA 'F')", "FILTROS")
  
  filtro_resultado <- filtro_excluidos(filtros, stats, planilha)

  # Retorna o resultado do último filtro, que já contém planilhas, filtros e stats
  return(filtro_resultado)
  
}

main_dados_limpar <- function(planilha) {
# Função para eliminar caracteres que podem causar problemas na importação no Solar
  
  planilha %>%
    mutate(
      # Código do item - remove quebras de linha e espaços
      `Código do item` = `Código do item` %>% 
        str_replace_all("[\r\n]", "") %>% 
        str_trim(),
      
      # Elimina caracteres não conversíveis nos textos das descrições dos itens e quebras de linha
      across(
        c(
          `Descrição Resumida`, 
          Especificação, 
          `Detalhamento (especificação complementar)`
        ),
        ~ .x %>%
          iconv(from = 'UTF-8', to = 'CP1252', sub = ' ') %>%
          str_replace_all("[\r\n]", " ")
      ),
      
      # Elimina espaços duplos das descrições dos itens
      `Detalhamento (especificação complementar)` = 
        `Detalhamento (especificação complementar)` %>% 
        str_squish()
    )
  
}

main_dados_montar <- function(planilha, unidades, info, script_a_executar) {
  
  main_dados_pivotar <- function(planilha, importacao) {
    
    tryCatch({
      # Pivotar os dados para formato longo (uma linha por unidade com demanda)
      planilha_longa <- planilha %>%
        pivot_longer(
          cols           = importacao$coluna_inicial:(importacao$coluna_inicial + importacao$qtde_unidades - 1),
          names_to       = "Unidade",
          values_to      = "Quantitativo",
          values_drop_na = TRUE) %>%
        # Filtrar apenas valores válidos (não zero, não X/x, não NA)
        filter(!is.na(Quantitativo),
               Quantitativo          != "0",
               tolower(Quantitativo) != "x") %>%
        # Converter quantidade para numérico
        mutate(qtd = suppressWarnings(as.numeric(gsub(",", ".", gsub("\\.", "", Quantitativo)))))
      
      return(planilha_longa)
    },
    error = function(e) {
      log_erro("Não foi possível pivotar os dados. Encerrando...",
               e,
               finalizar = TRUE)
    })
  }
  
  main_dados_adicionar_colunas <- function(planilha_longa, unidades, importacao) {
    planilha_longa <- planilha_longa %>%
      mutate(
        setorResponsavelPesquisa = ifelse(
          `Responsável pela Pesquisa` == "VALIDAÇÃO MANUAL" & !importacao$pos_processos,
          "VALIDAÇÃO MANUAL",
          unidades$sigla_solar[match(`Responsável pela Pesquisa`, unidades$Sigla)]
        ),
        setorOrigem        = unidades$sigla_solar[match(Unidade, unidades$Sigla)],
        codigoItem         = `Código do item`,
        descricaoResumida  = `Descrição Resumida`,
        item_compartilhado = `Item comp.`
      )
    
    if (importacao$pos_processos) {
      planilha_longa <- planilha_longa %>%
        mutate(
          n_item                 = `Nº Item`,
          processoFmt            = Processo,
          cpfResponsavelPesquisa = str_pad(unidades$cpf[match(`Responsável pela Pesquisa`, unidades$Sigla)], 11, pad = "0"),
          codigoImovel           = unidades$imovel[match(Unidade, unidades$Sigla)],
          unidadeMedida          = toupper(`Unidade de Medida`),
          especificacao          = Especificação,
          detalhamento           = `Detalhamento (especificação complementar)`
        )
    }
    
    return(planilha_longa)
  }
  
  main_dados_verificar_erros <- function(planilha_longa, importacao) {
    # Verificar erros de forma vetorizada
    msg_erro <- list()
    
    # Verifica quantitativos inválidos (NA após conversão)
    invalidos <- which(is.na(planilha_longa$qtd))
    invalidos_unidade <- planilha_longa$setorOrigem[invalidos]
    if (length(invalidos)) {
      msg_erro <- c(msg_erro, 
                    paste("Quantitativos inválidos nas linhas:", 
                          paste(planilha_longa$linha_planilha[invalidos], invalidos_unidade, sep = " da Unidade ", collapse = ", ")))
    }
    
    # Verificar setor responsável
    invalidos <- which(is.na(planilha_longa$setorResponsavelPesquisa))
    if (length(invalidos)) {
      msg_erro <- c(msg_erro, 
                    paste("Setor responsável pela pesquisa inválido nas linhas:",
                          paste(unique(planilha_longa$linha_planilha[invalidos]), collapse = ", ")))
    }
    
    # Verificar setor origem
    invalidos <- which(is.na(planilha_longa$setorOrigem))
    if (length(invalidos)) {
      msg_erro <- c(msg_erro, 
                    paste("Setor requerente inválido nas linhas:",
                          paste(unique(planilha_longa$linha_planilha[invalidos]), collapse = ", ")))
    }
    
    # Verificar código do item
    invalidos <- which(is.na(planilha_longa$codigoItem) | 
                       !str_detect(planilha_longa$codigoItem, "^\\d\\d\\d\\.\\d\\d\\.\\d\\d\\d\\d\\d\\d$"))
    if (length(invalidos)) {
      msg_erro <- c(msg_erro, 
                    paste("Código do item inválido nas linhas:",
                          paste(unique(planilha_longa$linha_planilha[invalidos]), collapse = ", ")))
    }
    
    # Verificar descrição resumida
    invalidos <- which(is.na(planilha_longa$descricaoResumida))
    if (length(invalidos)) {
      msg_erro <- c(msg_erro, 
                     paste("Descrição resumida inválida nas linhas:",
                           paste(unique(planilha_longa$linha_planilha[invalidos]), collapse = ", ")))
    }
    
    
    # Se houver processos, verificar campos adicionais
    if (importacao$pos_processos) {
  
      # Verificar processo
      invalidos <- which(is.na(planilha_longa$processoFmt) | 
                         !str_detect(planilha_longa$processoFmt, "^23080\\.\\d\\d\\d\\d\\d\\d/\\d\\d\\d\\d-\\d\\d$"))
      if (length(invalidos)) {
        msg_erro <- c(msg_erro, 
                       paste("Número de processo inválido nas linhas:",
                             paste(unique(planilha_longa$linha_planilha[invalidos]), collapse = ", ")))
      }
      
      # Verificar CPF
      invalidos <- which(is.na(planilha_longa$cpfResponsavelPesquisa) | 
                         !sapply(planilha_longa$cpfResponsavelPesquisa, utils_verificar_cpf))
      if (length(invalidos)) {
        msg_erro <- c(msg_erro, 
                      paste("CPF do responsável inválido nas linhas:",
                            paste(unique(planilha_longa$linha_planilha[invalidos]), collapse = ", ")))
      }
      
      # Verificar código imóvel
      invalidos <- which(is.na(planilha_longa$codigoImovel) | 
                         !str_detect(planilha_longa$codigoImovel, "^[0-9]{1,4}$"))
      if (length(invalidos)) {
        msg_erro <- c(msg_erro, 
                      paste("Código do imóvel inválido nas linhas:",
                            paste(unique(planilha_longa$linha_planilha[invalidos]), collapse = ", ")))
      }
      
      # Verificar unidade de medida
      invalidos <- which(is.na(planilha_longa$unidadeMedida) | 
                         !str_detect(planilha_longa$unidadeMedida, "^^[A-Za-z0-9]{1,10}$"))
      if (length(invalidos)) {
        msg_erro <- c(msg_erro, 
                      paste("Unidade de medida inválida nas linhas:",
                            paste(unique(planilha_longa$linha_planilha[invalidos]), collapse = ", ")))
      }
      
      # Verificar especificação
      invalidos <- which(is.na(planilha_longa$especificacao))
      if (length(invalidos)) {
        msg_erro <- c(msg_erro, 
                      paste("Especificação inválida nas linhas:",
                            paste(unique(planilha_longa$linha_planilha[invalidos]), collapse = ", ")))
      }
      
      # Verificar detalhamento
      invalidos <- which(is.na(planilha_longa$detalhamento))
      if (length(invalidos)) {
        msg_erro <- c(msg_erro, 
                      paste("Detalhamento inválido nas linhas:",
                            paste(unique(planilha_longa$linha_planilha[invalidos]), collapse = ", ")))
      }
    }
    
    return(msg_erro)
  }
  
  main_dados_criar_dataframe <- function(planilha_longa, importacao, info) {
  
    # Salva dados para conferência
    info$n_itens        <- length(unique(planilha_longa$linha_planilha))
    info$n_solicitacoes <- nrow(planilha_longa)
  
    config_json("conferencia", 
                list(ano            = info$ano, 
                     etapa          = info$etapa,
                     grupo          = info$grupo, 
                     nome_grupo     = info$nome_grupo,
                     n_itens        = info$n_itens, 
                     n_solicitacoes = info$n_solicitacoes))
    
    # Criar os dataframes finais
    if (importacao$pos_processos) {
      
      dados_importar <- planilha_longa %>%
        select(linha_planilha, qtd, processoFmt, setorResponsavelPesquisa, 
               cpfResponsavelPesquisa, setorOrigem, codigoImovel, codigoItem, 
               unidadeMedida, descricaoResumida, especificacao, detalhamento)
      
      dados_relatorio <- planilha_longa %>%
        select(n_item, linha_planilha, item_compartilhado, qtd, processoFmt, 
               setorResponsavelPesquisa, setorOrigem, codigoItem, descricaoResumida)
      
      return(list(importar  = as.data.frame(dados_importar), 
                  relatorio = as.data.frame(dados_relatorio), 
                  info      = info))
      
    } else {
      
      dados_relatorio <- planilha_longa %>%
        select(linha_planilha, item_compartilhado, qtd, setorResponsavelPesquisa, 
               setorOrigem, codigoItem, descricaoResumida)
      
      return(list(relatorio = as.data.frame(dados_relatorio), 
                  info      = info))
      
    }
    
  }
  
  #verifique se há processos distribuídos e se há VALIDAÇÃO MANUAL
  importacao <- get_config("importacao")
  
  importacao$pos_processos    <- ifelse(sum(!is.na(planilha$Processo)) == 0, FALSE, TRUE)
  importacao$validacao_manual <- any(planilha$'Responsável pela Pesquisa' == 'VALIDAÇÃO MANUAL')
  
  set_config(importacao = importacao)
  
  #verifica se o script chamado é para gerar arquivos para importação
  #se não foram gerados os processos (pos_processos = FALSE) avisa e encerra
  if(script_a_executar == "gerar" & !importacao$pos_processos) {
    log_erro("Você está tentando gerar os arquivos para importação no Solar, mas ainda não foram distribuídos os processos. Encerrando...",
             finalizar = TRUE)
  }
  
  #Pivotar e filtrar dados básicos
  planilha_longa <- main_dados_pivotar(planilha, importacao)
  
  #Adicionar colunas
  planilha_longa <- main_dados_adicionar_colunas(planilha_longa, unidades, importacao)
  
  #Verificar erros
  erros <- main_dados_verificar_erros(planilha_longa, importacao)
  if (length(erros) > 0) {
    log_erro("Houve problemas na captura dos dados.", unique(unlist(erros)))
    return(NULL)
  }
  
  #Criar dataframes finais (importar e relatório)
  dados_finais <- main_dados_criar_dataframe(planilha_longa, importacao, info)
  
  #Feedback visual do resultado
  log_info(paste0("ANO             : ", dados_finais$info$ano),
           paste0("ETAPA           : ", dados_finais$info$etapa),
           paste0("GRUPO           : ", dados_finais$info$grupo),
           paste0("NOME DO GRUPO   : ", dados_finais$info$nome_grupo),
           paste0("N. ITENS        : ", dados_finais$info$n_itens),
           paste0("N. SOLICITAÇÕES : ", dados_finais$info$n_solicitacoes),
           "-",
           sprintf("Pós geração dos processos : %s", ifelse(importacao$pos_processos, "Sim", "Não")),
           sprintf("Validação manual          : %s", ifelse(importacao$validacao_manual, "Sim", "Não")),
           sprintf("Relatório consolidado     : %s", ifelse(importacao$consolidado, "Sim", "Não")),
           cores = "verde")
  
  return(dados_finais)
}

main_lista_final_obter <- function() {
  importacao  <- get_config("importacao")
  
  pb <- log_barra_progresso("Obtendo informações da LISTA FINAL", 4)
  
  log_info("Planilha de inserção de demandas",
           paste0("→ Link : ", importacao$url_planilha),
           paste0("→ Aba  : ", importacao$aba_lista_final),
           estilo = "inicio")
  
  planilha_original <- main_lista_final_baixar(importacao)

  log_barra_progresso("Ajustando planilha", pb = pb)
  
  planilha_ajustada <- main_lista_final_limpar(planilha_original, importacao)
  
  log_barra_progresso("Obtendo informações da Etapa do Calendário", pb = pb)
  
  log_info(paste0("→ Aba  : ", importacao$aba_menu),
           estilo = "fim")
  
  planilha_info <- main_lista_final_info(importacao)
  
  log_barra_progresso("Obtendo dados das Unidades requerentes", pb = pb)
  
  unidades <- utils_unidades_requerentes_obter()

  log_barra_progresso(pb = pb)

  if (!is.null(planilha_ajustada) & !is.null(unidades)) {
    lista_final <- list(original = planilha_original,
                        ajustada = planilha_ajustada,
                        info     = planilha_info,
                        unidades = unidades)
    return(lista_final)
  } else {
    log_erro("Não foi possível obter os dados da Lista Final. Encerrando...",
             finalizar = TRUE)
  }
}

main_lista_final_baixar <- function(importacao) {
  gs4_deauth()

  tryCatch({
    planilha_original <- 
      range_read(importacao$url_planilha, 
                 sheet = importacao$aba_lista_final, 
                 skip = importacao$linha_inicial - 2,
                 col_names = TRUE,
                 col_types = paste0("cccccllllcccccnn", strrep("c", importacao$qtde_unidades), "cccccn"))
    return(planilha_original)
  },
  error = function(e) { 
    log_erro("Erro ao acessar planilha de inserção de demandas. Verifique o link informado. Encerrando...", 
             e,
             finalizar = TRUE)
  })
}

main_lista_final_limpar <- function(planilha_original, importacao) {
  tryCatch({
    #adiciona o número da linha da planilha original do Google Drive
    #altera o nome da última coluna para "Linha original"
    planilha_original$`Qtd. Mapa` <- seq.int(importacao$linha_inicial, nrow(planilha_original) + importacao$linha_inicial - 1)
    colnames(planilha_original)[ncol(planilha_original)] <- "linha_planilha"

    #ajusta número da linha da planilha original com 3 dígitos antecedidos por 0
    planilha_original <- planilha_original %>%
      mutate(linha_planilha = str_pad(linha_planilha, width = 3, pad = "0"))
    
    #exclui últimas colunas que não são de unidades
    planilha_original[(importacao$coluna_inicial + importacao$qtde_unidades):(ncol(planilha_original) - 1)] <- NULL
    
    return(planilha_original)
  },
  error = function(e) { 
    log_erro("Erro ao ajustar os dados das planilhas. Encerrando...", 
             e,
             finalizar = TRUE)
  })
}
  
main_lista_final_info <- function(importacao) {
  tryCatch({
    info_lista <- range_read(importacao$url_planilha, sheet = importacao$aba_menu, range = "G4:G7", col_names = FALSE)
    
    ano        <- as.numeric(unlist(info_lista)[1])
    etapa      <- as.character(unlist(info_lista)[2])
    grupo      <- as.character(unlist(info_lista)[3])
    nome_grupo <- as.character(unlist(info_lista)[4])
    
    if (ano        == "FALSE" | is.na(ano)        | ano        == "") { ano        <- "Não informado" }
    if (etapa      == "FALSE" | is.na(etapa)      | etapa      == "") { etapa      <- "Não informado" }
    if (grupo      == "FALSE" | is.na(grupo)      | grupo      == "") { grupo      <- "Não informado" }
    if (nome_grupo == "FALSE" | is.na(nome_grupo) | nome_grupo == "") { nome_grupo <- "Não informado" }
    
    log_info(c(paste0("ANO      : ", ano),
               paste0("ETAPA    : ", etapa),
               paste0("GRUPO    : ", grupo),
               paste0("MATERIAL : ", nome_grupo)),
             cores = "verde")

    info <- list(ano        = ano, 
                 etapa      = etapa, 
                 grupo      = grupo, 
                 nome_grupo = nome_grupo)
    
    return(info)
  },  
  error = function(e) { 
    log_erro("Erro ao acessar as informações da lista (ano, etapa e grupo). Verifique a aba Menu da LISTA FINAL. Encerrando...", 
             e,
             finalizar = TRUE)
  })
}

main_utils_grupo <- function(dados) {
  tryCatch({
    info_grupo <- dados$info$grupo
    
    info_grupo <- gsub("[\\\\/:*?\"<>|]", "", info_grupo)
    
    if (nchar(info_grupo) > 13) {
      grupo <- paste0(str_sub(trimws(info_grupo), 1, 13), "...")
    } else {
      grupo <- info_grupo
    }
    
    return(grupo)
  },
  error = function(e) { 
    log_erro("Não foi possível obter o nome do grupo. Utilizando nome genérico.", 
             e,
             alerta = TRUE)
    return("grupo_invalido")
  })
}

# GERAÇÃO DE ARQUIVOS A IMPORTAR====

gerar_comparar_detalhamentos <- function(dados) {
  pasta <- get_config("pasta")
  grupo <- main_utils_grupo(dados)
  
  # salva arquivo com os itens que tiveram caracteres ajustados na descrição complementar para o sistema Solar
  tryCatch({
    comparacao <- dados$importar %>%
      select(linha_planilha, codigoItem, descricaoResumida, detalhamento) %>%
      left_join(
        dados$original %>% 
          mutate(linha_planilha = str_pad(linha_planilha, 3, pad = "0")) %>%
          select(linha_planilha, `Detalhamento (especificação complementar)`),
        by = "linha_planilha") %>%
      filter(detalhamento != `Detalhamento (especificação complementar)`) %>%
      distinct(linha_planilha, .keep_all = TRUE) %>%
      select(linha_planilha, codigoItem, descricaoResumida, detalhamento)     
  },
  error = function(e) { 
    log_erro("Não foi possível verificar os detalhamentos a ajustar. ATENÇÃO: Isto não impede a importação dos pedidos.", 
             e, 
             alerta = TRUE) 
    comparacao <- NULL
  })
  
  if (!is.null(comparacao) && nrow(comparacao) > 0) {
    tryCatch({
      arquivo_descricao <- sprintf("Descrição a ajustar - Grupo %s.xlsx", 
                                   grupo)

      wb <- createWorkbook()
        addWorksheet(wb, grupo)
        writeData(wb, grupo, 
                  comparacao)
        setColWidths(wb, grupo, 
                     cols = c(1, 2, 3, 4),
                     widths = c("auto", "auto", "auto", 150))
        addStyle(wb, sheet = grupo,
                 style = createStyle(fgFill = 'yellow'),
                 cols = 4,
                 rows = 1:nrow(comparacao) + 1)
      saveWorkbook(wb, 
                   file.path(pasta$arquivos_importar, arquivo_descricao), 
                   overwrite = TRUE)
      
      log_erro(sprintf("%d descrições precisam de ajuste. Verifique o arquivo 'Descrição a ajustar'", nrow(comparacao)), 
               alerta = TRUE)
      
      print.data.frame(
        select(comparacao, linha_planilha, codigoItem, descricaoResumida),
        right = FALSE,
        row.names = FALSE)
    },
    error = function(e) { 
      log_erro("Não foi possível salvar o arquivo com as descrições a ajustar. Verifique o arquivo de log.", 
               e, 
               alerta = TRUE) 
    })
  }
}

gerar_salvar_arquivos_a_importar <- function(dados) {
  pasta     <- get_config("pasta")
  processos <- unique(dados$importar$processoFmt)
  
  # salva banco de dados conforme dividido por processos
  tryCatch({
    arquivos_gerados <- processos %>%
      map_chr(~ {
        arquivo_csv <- sprintf("Importar - %s - Processo %s.csv",
                               main_utils_grupo(dados),
                               substr(.x, 7, 12))
        
        dados$importar %>%
          filter(processoFmt == .x) %>%
          select(-linha_planilha) %>%
          write.csv2(
            file         = file.path(pasta$arquivos_importar, arquivo_csv),
            row.names    = FALSE,
            quote        = c(9, 10, 11),
            fileEncoding = "CP1252")
        
        arquivo_csv
      })
    
    log_info("ARQUIVOS GERADOS PARA IMPORTAÇÃO NO SOLAR", 
             arquivos_gerados,
             cores = "verde")
  }, 
  error = function(e) {
    log_erro("Não foi possível gravar os arquivos a importar. Encerrando...", 
             e,
             finalizar = TRUE)
  })
}

gerar_main <- function(dados) {
  if (is.null(dados) | is.null(dados$importar)) {
    log_erro("Não foi possível gerar arquivos. Dados importados não encontrados. Encerrando...",
             finalizar = TRUE)
  }
  
  log_secao("SALVANDO ARQUIVOS")
  
  gerar_salvar_arquivos_a_importar(dados)
  
  log_secao("VERIFICANDO DETALHAMENTOS A AJUSTAR")
  
  gerar_comparar_detalhamentos(dados)
  
  # finaliza e grava log
  config_finalizar(sucesso = TRUE)
}

# RELATÓRIO GERENCIAL====

relatorio_gerar_temporario <- function(dados) {
# Define o nome do arquivo .html a ser salvo,
# gera e salva o .html temporário, que depois será excluído após salvar em .pdf
# e passa o nome do arquivo de volta à função principal
  
  pasta      <- get_config("pasta")
  importacao <- get_config("importacao")

  # Define o nome do arquivo do relatório, incluindo o caminho  
  relatorio_arquivo = sprintf("Relatório Gerencial %s %s %s",
                              ifelse(importacao$pos_processos, "-", "Inicial -"),
                              main_utils_grupo(dados),
                              ifelse(importacao$pos_processos,
                                     ifelse(importacao$consolidado, 
                                            '- Consolidado',
                                            paste0("- ", gsub("/", "-", importacao$processo_para_relatorio))),
                                     ifelse(importacao$validacao_manual, "- PRELIMINAR", "")))
  
  tryCatch({
    
    # Gera o relatório em html
    rmarkdown::render("importacao.Rmd",
                      output_dir = pasta$relatorios,
                      intermediates_dir = pasta$relatorios,
                      knit_root_dir = pasta$relatorios,
                      output_file = paste0(relatorio_arquivo, ".html"),
                      clean = TRUE)
    
    log_info("Arquivo temporário html salvo:", relatorio_arquivo,
             cores = "verde")
    
    # Retorna o nome do arquivo à função principal
    return(relatorio_arquivo)
  },
  error = function(e) { 
    log_erro("Não foi possível gerar o relatório em HTML. Encerrando...", 
             e, 
             finalizar = TRUE)
  })
}

relatorio_preparar_ambiente <- function() { 
# Instala, se necessário, e ativa o pandoc
# para interpretar o arquivo .rmd e gerar o relatório em html
  
  tryCatch({
    # Se não existe pandoc, o instala
    if (!pandoc_available()) {
      log_info("Pandoc não disponível. Instalando...")
      pandoc::pandoc_install()
    }
    log_info("Pandoc instalado. Ativando...")
    pandoc::pandoc_activate()
  },
  error = function(e) { 
    log_erro("Não foi possível instalar o pacote Pandoc. Encerrando...", 
             e, 
             finalizar = TRUE)
  })
}

relatorio_salvar_pdf <- function(relatorio_arquivo) {
# Imprime em .pdf o relatório temporário .html gerado
# utilizando o Chrome
  
  pasta <- get_config("pasta")
  
  # Imprime em pdf
  tryCatch({
    chrome_print(input = file.path(pasta$relatorios, paste0(relatorio_arquivo, ".html")), 
                 output = file.path(pasta$relatorios, paste0(relatorio_arquivo, ".pdf")),
                 wait = 1,
                 verbose = 1)
    
    # Exclui .html temporário
    unlink(file.path(pasta$relatorios, paste0(relatorio_arquivo, ".html")))
  },
  error = function(e) { 
    log_erro("Não foi possível salvar o relatório final em PDF. Encerrando...", 
             e, 
             finalizar = TRUE)
  })
}

relatorio_planilha_dfd <- function(dados) {
# Cria e salva planilha auxiliar para edição do DFD  
  
  config       <- get_config()
  pasta        <- config$pasta
  processo_spa <- config$importacao$processo_para_relatorio

  # Gera a planilha somente se for escolhido um processo específico
  if (processo_spa != "todos") {
    
    # Planilha com as Unidades e números dos pedidos/SDs
    planilha_dfd <- dados$filtros$planilha_controle %>%
      filter(processo == processo_spa) %>%
      select(requerente, pedido, tipo, oficio, peculiaridades, etp, equipe_apoio)
    
    # Planilha com as informações dos itens
    planilha_itens <- dados$relatorio %>%
      mutate(
        n_item_num = suppressWarnings(as.numeric(n_item))
      ) %>%
      filter(!is.na(n_item_num)) %>%
      group_by(n_item_num) %>%
      summarise(
        linha_planilha = first(linha_planilha),
        codigoItem = first(codigoItem), 
        descricaoResumida = first(descricaoResumida),
        qtd_total = sum(qtd, na.rm = TRUE)
      ) %>%
      rename(n_item = n_item_num)
    
    tryCatch({
      
      # Cria o arquivo Excel
      wb <- createWorkbook()
        addWorksheet(wb, "DFD", tabColour = "blue")
        writeDataTable(wb, "DFD", planilha_dfd, tableStyle = "TableStyleMedium2")
        setColWidths(wb, "DFD", 
                     cols   = 1:7,
                     widths = 20)
        
        addWorksheet(wb, "Itens", tabColour = "green")
        writeDataTable(wb, "Itens", planilha_itens, tableStyle = "TableStyleMedium4")
        setColWidths(wb, "Itens", 
                     cols   = 1:5,
                     widths = c(15, 15, 15, 50, 15))
      
      # Salva o arquivo
      saveWorkbook(wb, 
                   file.path(pasta$relatorios, paste0("DFD - ", gsub("/", "-", processo_spa), ".xlsx")), 
                   overwrite = TRUE)
    },
    error = function(e) { 
      log_erro("Erro ao salvar o arquivo. Encerrando...", 
               e,
               finalizar = TRUE)
    },
    warning = function(w) { 
      log_erro("Erro ao salvar o arquivo. Encerrando...", 
               w,
               finalizar = TRUE)
    })
  }
}

relatorio_main <- function(dados) {
# FUNÇÃO PRINCIPAL PARA RELATÓRIO GERENCIAL
  
  # Se não houver dados, gera erro e finaliza
  if (is.null(dados) | is.null(dados$relatorio)) {
    log_erro("Não foi possível gerar relatórios. Dados importados não encontrados. Encerrando...",
             finalizar = TRUE)
  }
  
  log_secao("GERANDO RELATÓRIO DA LISTA FINAL")
  
  # Inicia barra de progresso
  pb <- log_barra_progresso("Iniciando...", 5)
  
  log_barra_progresso("Instalando pacote RMarkdown", pb = pb)
  
  relatorio_preparar_ambiente()
  
  log_barra_progresso("Produzindo relatório", pb = pb)

  relatorio_arquivo <- relatorio_gerar_temporario(dados)
  
  log_barra_progresso("Salvando em PDF", pb = pb)
  
  relatorio_salvar_pdf(relatorio_arquivo)
  
  log_barra_progresso("Gerando planilha para DFD", pb = pb)
  
  relatorio_planilha_dfd(dados)
  
  log_barra_progresso(pb = pb)
  
  #finaliza e grava log
  config_finalizar(sucesso = TRUE)
}

# RESUMO DOS PEDIDOS====

resumo_montar <- function(dados, lista_inicial) {
  
  tryCatch({
    # Processa os dados de pedidos usando dplyr
    pedidos_total <- dados$importar %>%
      count(processoFmt, setorOrigem, name = "N. itens") %>%
      left_join(
        distinct(dados$unidades, sigla_solar, email),
        by = c("setorOrigem" = "sigla_solar")
      ) %>%
      rename(
        Processo = processoFmt,
        Unidade  = setorOrigem,
        `E-mail` = email
      )
    
    extrair_todos_dados <- function(texto, padrao) {
      map(texto, ~str_extract_all(.x, padrao)[[1]]) %>% 
        flatten_chr()
    }
    
    # Processa todos os elementos da lista_inicial
    lista_processo  <- extrair_todos_dados(lista_inicial, "\\d{5}\\.\\d{6}/\\d{4}-\\d{2}")
    lista_unidade   <- extrair_todos_dados(lista_inicial, "[:alpha:].+[:alpha:]") %>% 
      str_replace("PRODEGESP/UFS", "PRODEGESP/UFSC")
    lista_protocolo <- extrair_todos_dados(lista_inicial, " \\d{6}/\\d{4}") %>% 
      str_trim()
    lista_qtd       <- extrair_todos_dados(lista_inicial, "\\s\\d{1,3}\\s") %>% 
      str_trim() %>% as.numeric()
    
    # Cria tibble com os dados da lista
    lista_pedidos <- tibble(
      Processo     = lista_processo,
      Unidade      = lista_unidade,
      `N. Pedido`  = lista_protocolo,
      `Para orçar` = lista_qtd
    )
    
    # Cruza as tabelas
    lista_final <- pedidos_total %>%
      left_join(lista_pedidos, by = c("Processo", "Unidade")) %>%
      mutate(
        `N. Pedido`  = coalesce(`N. Pedido`, sprintf("SD %s", Unidade)),
        `Para orçar` = as.numeric(`Para orçar`)
      ) %>%
      select(Processo, Unidade, `N. Pedido`, `N. itens`, `Para orçar`, `E-mail`) %>%
      arrange(Processo, desc(`Para orçar`), Unidade)
    
    # Inclui informações no objeto dados
    dados$resumo <- lista_final
    
    return(dados)
  },
  error = function(e) {
    log_erro("Aconteceu algum erro ao tratar os dados. Encerrando...",
             e,
             finalizar = TRUE)
  })
}

resumo_obter_pdf <- function() {
  
  pasta <- get_config("pasta")
  
  tryCatch({
    #cria lista de arquivos com os pedidos gerados
    pdf_arquivos <- Sys.glob(file.path(pasta$resumo_pedidos, "*.pdf*"))
    
    log_info(pdf_arquivos,
             cores = "verde")
    
    #captura o conteúdo dos PDFs dos pedidos
    pdf_dados <- lapply(pdf_arquivos, pdf_text)
    pdf_dados <- as.character(pdf_dados)
    
    #obtém dados dos PDFs
    lista_pedidos <- str_extract(pdf_dados, "(\\d\\d\\d\\d\\d\\.\\d\\d\\d\\d\\d\\d/\\d\\d\\d\\d-\\d\\d)(?s)(.*)(\\d\\d\\d\\d\\d\\d/\\d\\d\\d\\d)")
    
    return(lista_pedidos)
  },
  error = function(e) {
    log_erro("Não foi possível obter o conteúdo dos PDFs dos resumos. Encerrando...", 
             finalizar = TRUE)
  })
}

resumo_salvar <- function(dados) {
  #salvar os dados em arquivo Excel
  processos <- sort(unique(dados$importar$processoFmt))
  cores     <- brewer.pal(n = max(3, length(processos)), name = "Blues")
  grupo     <- main_utils_grupo(dados)
  pasta     <- get_config("pasta")
  
  tryCatch({
    wb <- createWorkbook()
      addWorksheet(wb, grupo)
      writeData(wb, grupo, dados$resumo)
      for (i in seq_along(processos)) {
        conditionalFormatting(wb, grupo,
                              cols  = 1:6,
                              rows  = 1:(nrow(dados$resumo) + 1),
                              rule  = paste0("$A1=\"", processos[i], "\""), 
                              style = createStyle(bgFill = cores[i])) 
      }
      setColWidths(wb, grupo, 
                   cols   = 1:6,
                   widths = "auto")
    saveWorkbook(wb, 
                 file.path(pasta$resumo_pedidos, paste0("Resumo - Grupo ", grupo, ".xlsx")), 
                 overwrite = TRUE)
  },
  error = function(e) { 
    log_erro("Erro ao salvar o arquivo. Encerrando...", 
             e,
             finalizar = TRUE)
  },
  warning = function(w) { 
    log_erro("Erro ao salvar o arquivo. Encerrando...", 
             w,
             finalizar = TRUE)
  })
}

resumo_main <- function(dados) {
  if (is.null(dados) | is.null(dados$importar)) {
    log_erro("Não foi possível gerar arquivos. Dados importados não encontrados. Encerrando...",
             finalizar = TRUE)
  }
  
  log_secao("OBTENDO DADOS DOS PEDIDOS")
  
  lista_inicial <- resumo_obter_pdf()
  
  log_secao("ORGANIZANDO DADOS")
  
  dados <- resumo_montar(dados, lista_inicial)

  log_secao("SALVANDO RESUMO DOS PEDIDOS")
  
  resumo_salvar(dados)
  
  #finaliza e grava log
  config_finalizar(sucesso = TRUE)
}

# SCRIPT GERAL====

main_inicializar <- function(pacotes, script_a_executar) {
# Função chamada pela função principal para inicializar ambiente
  
  # Complementa com as pastas necessárias aos scripts vinculados às listas finais
  pasta <- list()
  pasta$atual             = getwd()
  pasta$arquivos_importar = file.path(pasta$atual,
                                      Sys.getenv("IMPORTACAO_ARQUIVOS_A_IMPORTAR"))
  pasta$resumo_pedidos    = file.path(pasta$atual,
                                      Sys.getenv("IMPORTACAO_RESUMO_PEDIDOS"))
  pasta$relatorios        = file.path(pasta$atual,
                                      Sys.getenv("IMPORTACAO_RELATORIOS"))
  pasta$criar             = c(pasta$relatorios, 
                              pasta$arquivos_importar, 
                              pasta$resumo_pedidos)

  # Complementa com os pacotes necessários aos scripts vinculados às listas finais
  pacotes_compartilhados <- c("googlesheets4", "openxlsx2", "openxlsx", 
                              "stringr", "tidyr", "dplyr", "purrr")
  pacotes <- if (missing(pacotes)) pacotes_compartilhados else append(pacotes_compartilhados, pacotes)

  # Chama função compartilhada para inicializar configurações gerais
  config_inicializar(pacotes, pasta)
  
  log_secao("CONFIGURAÇÕES INICIAIS", "MAIN")
  
  # Carrega/define configurações dos scripts das listas finais
  main_ambiente(script_a_executar)
  
}

main_ambiente <- function(script_a_executar) {
# Função para definir/salvar configurações no ambiente atual para uso
# durante a execução do script
  
  pasta      <- get_config("pasta")
  importacao <- list()
  
  #Limpa informações para conferência da geração anterior
  config_json("conferencia", opcao = "remove")
  
  #Se o script for para gerar arquivo para importacao ou o resumo, não será consolidado
  #Se for para relatório e não houver número do processo, será consolidado
  importacao$processo_para_relatorio <- config_json("processo", opcao = "get")
  importacao$consolidado <- 
    if (script_a_executar != "relatorio" || 
        importacao$processo_para_relatorio == "todos" || 
        is.na(importacao$processo_para_relatorio)) TRUE else FALSE
  
  tryCatch({
    
    #Lê arquivo config.json para carregar configurações e gravar no ambiente atual
    R_config <- config_json(opcao = "all")

    celula_inicial <- trimws(R_config$celula)
    
    importacao$url_planilha <- R_config$link_planilha
    importacao$valor_minimo <- as.numeric(R_config$valor_minimo)
    importacao$qtde_minima  <- as.numeric(R_config$qtde_minima)
    
    #Separa a celula_inicial para saber 
    #a coluna (posição no alfabeto da letra da coluna) e o número da linha
    importacao$coluna_inicial <- which(letters[1:26] == tolower(substring(celula_inicial, 1, 1)))
    importacao$linha_inicial  <- as.numeric(substring(celula_inicial, 2, 2))
    
    importacao$qtde_unidades   <- as.numeric(R_config$unidades)
    importacao$aba_menu        <- trimws(R_config$aba_menu)
    importacao$aba_lista_final <- trimws(R_config$aba_lista_final)

    log_info("Pastas",
     c(sprintf("ATUAL               : %s", pasta$atual),
       sprintf("ARQUIVOS A IMPORTAR : %s", pasta$arquivos_importar),
       sprintf("RESUMO PEDIDOS      : %s", pasta$resumo_pedidos),
       sprintf("RELATÓRIOS          : %s", pasta$relatorios)),
     "-",
     "Configurações informadas",
     c(sprintf("Valor mínimo           : %s", importacao$valor_minimo),
       sprintf("Quantidade mínima      : %s", importacao$qtde_minima),
       sprintf("Célula inicial         : %s", celula_inicial),
       sprintf("Quantidade de Unidades : %s", importacao$qtde_unidades),
       sprintf("Aba Menu               : %s", importacao$aba_menu),
       sprintf("Aba LISTA FINAL        : %s", importacao$aba_lista_final),
       sprintf("Processo - Relatório   : %s", importacao$processo_para_relatorio),
       sprintf("Script a executar      : %s", script_a_executar)),
      cores = "verde")

    # Grava as configurações recuperadas do arquivo no ambiente atual
    set_config(importacao = importacao)
    
  },
  error = function(e) { 
    log_erro("Erro ao acessar o arquivo de configurações 'config.json'. Encerrando...", 
             e,
             finalizar = TRUE)
  })
}

importacao_main <- function() {
# FUNÇÃO PRINCIPAL
# Chamada ao carregar este script
# Faz definições iniciais, obtém dados 
# e passa para atividade a ser executada conforme escolha do usuário

  # Carrega funções compartilhadas
  source(file.path("..", "_common", "config.R"), chdir = TRUE)
  
  # Verifica que script foi solicitada a execução
  script_a_executar <- commandArgs(trailingOnly = TRUE)[1]
  if (is.na(script_a_executar)) script_a_executar <- "relatorio"
  
  # Define nome para log e pacotes conforme o script a executar
  switch(
    script_a_executar,
    "gerar" = {
      pacotes <- c()
    },
    "resumo" = {
      pacotes <- c("pdftools", "RColorBrewer")
    },
    "relatorio" = {
      pacotes <- c("rmarkdown", "knitr", "pandoc", 
                   "pivottabler", "pagedown", "kableExtra")
    }
  )

  cat("►►► Continuando configuração do ambiente para importação\n")
  
  # Passa configurações para inicializar script
  main_inicializar(
    pacotes           = pacotes,
    script_a_executar = script_a_executar
  )

  log_secao("OBTENDO DADOS DA LISTA FINAL")
  
  # Chama função geral que fará obtenção dos dados
  # Para verificar andamento, ver a função main_dados_obter()
  dados <- main_dados_obter(script_a_executar)
  
  # Após obtidos os dados, executa o script conforme definição do usuário
  switch(
    script_a_executar,
    "gerar"     = gerar_main(dados),
    "resumo"    = resumo_main(dados),
    "relatorio" = relatorio_main(dados)
  )
  
}

# Ao carregar script, inicia execução chamando função principal
importacao_main()