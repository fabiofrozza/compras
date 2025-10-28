fornecedores_arquivo_a_importar_montar <- function(dados, municipios) {
  
  #vincular código município na coluna cod_mun
  tryCatch({
    dados_a_importar <- merge(dados$fornecedores[ , -c(1, 16)], 
                              municipios[!duplicated(municipios$nome_mun), ], 
                              by = "nome_mun", 
                              all.x = TRUE)
    
    #preenche municipios com problema com São Paulo
    dados_a_importar$cod_mun[is.na(dados_a_importar$cod_mun)] <- 7107
  
    return(dados_a_importar)
  },
  error = function(e) {
    log_erro("Não foi possível montar o arquivo para importar. Encerrando...", 
             e,
             finalizar = TRUE) 
  })
}

fornecedores_arquivo_a_importar_salvar <- function(dados_a_importar, pregao) {
  
  pasta <- get_config("pasta")
  
  #salvar os dados em arquivo .csv para importar e .xls para conferência
  if (nrow(dados_a_importar > 0)) {
    tryCatch( 
      write.csv(dados_a_importar, 
                file.path(pasta$importar, paste0("PE_", pregao, ".csv")), 
                row.names = FALSE, 
                fileEncoding = "CP1252")
    ,
    error = function(e) { 
      log_erro("Não foi possível salvar o arquivo para importar. Encerrando...", 
               e,
               finalizar = TRUE)
    })
  }
}

fornecedores_arquivo_comparacao_salvar <- function(dados, dados_a_importar, pregao) {

  pasta <- get_config("pasta")
  
  comparacao <- data.frame(dados$originais[ , c(1:3)],
                           dados$originais$banco,   dados$para_comparacao$banco, 
                           dados$originais$agencia, dados$para_comparacao$agencia, 
                           dados$originais$conta,   dados$para_comparacao$conta)
  names(comparacao) <- c("Arquivo", "CNPJ", "Razão Social",
                         "Banco original", "Banco ajustado",
                         "Agência original", "Agência ajustada",
                         "Conta original", "Conta ajustada")

  tryCatch({
    wb <- createWorkbook()
    
      addWorksheet(wb, "Comparação")
      addWorksheet(wb, "Dados originais")
      addWorksheet(wb, "Dados ajustados")
      addWorksheet(wb, "Arquivo a importar")
      
      writeData(wb, "Comparação", comparacao,
                headerStyle = createStyle(textDecoration = "bold", halign = "center"))
      writeData(wb, "Dados originais",    dados$originais)
      writeData(wb, "Dados ajustados",    dados$fornecedores)
      writeData(wb, "Arquivo a importar", dados_a_importar)
      
      verde    <- createStyle(bgFill = "green")
      vermelho <- createStyle(bgFill = "red")
      
      conditionalFormatting(wb, 
                            "Comparação",
                            cols  = 5,
                            rows  = 2:(nrow(comparacao) + 1), 
                            rule  = "E2<>D2", 
                            style = vermelho) 
      
      conditionalFormatting(wb, 
                            "Comparação",
                            cols  = 7,
                            rows  = 2:(nrow(comparacao) + 1), 
                            rule  = "G2<>F2", 
                            style = vermelho) 

      conditionalFormatting(wb, 
                            "Comparação",
                            cols  = 9,
                            rows  = 2:(nrow(comparacao) + 1), 
                            rule  = "I2<>H2", 
                            style = vermelho) 

      conditionalFormatting(wb, 
                            "Comparação",
                            cols  = 5,
                            rows  = 2:(nrow(comparacao) + 1), 
                            rule  = "E2=D2", 
                            style = verde) 

      conditionalFormatting(wb, 
                            "Comparação",
                            cols  = 7,
                            rows  = 2:(nrow(comparacao) + 1), 
                            rule  = "G2=F2", 
                            style = verde) 
      
      conditionalFormatting(wb, 
                            "Comparação",
                            cols  = 9,
                            rows  = 2:(nrow(comparacao) + 1), 
                            rule  = "I2=H2", 
                            style = verde) 
      
      setColWidths(wb, 
                   "Comparação", 
                   cols = c(1:9),
                   widths = c("auto", "auto", 50, 16, 16, 16, 16, 16, 16))
      
      addStyle(wb, 
               "Comparação", 
               style = createStyle(halign = 'right'), 
               cols = 4:9, 
               rows = 2:(nrow(comparacao) + 1), 
               gridExpand = TRUE)

    saveWorkbook(wb, 
               file = file.path(pasta$importar, 
                                paste0("PE_", pregao, "_CONFERENCIA.xlsx")), 
               overwrite = TRUE)
  },
  error = function(e) { 
    log_erro(paste0("Erro ao salvar arquivo de conferência PE_", pregao, "_CONFERENCIA.xlsx. Verifique se não está em uso."), 
             e,
             alerta = TRUE)
  },
  warning = function(w) { 
    log_erro(paste0("Erro ao salvar arquivo de conferência PE_", pregao, "_CONFERENCIA.xlsx. Verifique se não está em uso."), 
             w)
  })
}

fornecedores_excel_fornecedores_ler <- function(excel_fornecedores) {
  excel_conteudo <- NULL
  excel_dados    <- NULL
  
  for (i in seq_along(excel_fornecedores)) {
    cat(sprintf(" - Lendo arquivo %s\n", basename(excel_fornecedores[i])))
  
    tryCatch({
      excel_conteudo[[i]] <- read_excel(excel_fornecedores[i], 
                                        col_names = c("dado", "informacao"))
    
      linha_inicial <- which(excel_conteudo[[i]]$dado == "CNPJ")
      
      excel_dados[[i]] <- excel_conteudo[[i]][linha_inicial:(linha_inicial + 14), ] 
    },
    error = function(e) { 
      log_erro(sprintf("Não foi possível ler ou há algum problema com o formato do arquivo %s", basename(excel_fornecedores[i])), 
               e,
               alerta = TRUE) 
    })
  }
  
  return(excel_dados)
}

fornecedores_excel_fornecedores_obter <- function(pregao) {
  
  pasta <- get_config("pasta")
  
  tryCatch({
    excel_fornecedores <- list.files(path = file.path(pasta$dados, pregao),
                                     pattern = "^[^~]*.xls*",
                                     full.names = TRUE)
    
    log_info("PREGÃO", pregao,
             "-",
             "RELATÓRIOS SICAF", excel_fornecedores,
             cores = "verde")
    
    return(excel_fornecedores)
  },
  error = function(e) { 
    log_erro("Não foi possível obter a lista de arquivos. Encerrando...", 
             e,
             finalizar = TRUE) 
  })
}

fornecedores_cnpj_verificar <- function(dados, arquivo_alerta) {
  #elimina linhas sem CNPJ, verifica dígitos do CNPJ e cria arquivo de alerta
  dados$fornecedores$CNPJ_OK <- sapply(dados$fornecedores$cnpj, utils_verificar_cnpj)
  dados$para_comparacao      <- dados$fornecedores
  
  if (any(is.na(dados$fornecedores$cnpj)) || any(dados$fornecedores$CNPJ_OK == FALSE)) {
    sem_cnpj    <- dados$fornecedores[is.na(dados$fornecedores$cnpj), c(1:3)]
    cnpj_errado <- dados$fornecedores[dados$fornecedores$CNPJ_OK == FALSE & !is.na(dados$fornecedores$cnpj), c(1:3)]
    
    if (nrow(sem_cnpj) > 0) {
      mensagem_sem_cnpj <- c("OS SEGUINTES ARQUIVOS NÃO POSSUEM CNPJ.", 
                             "=======================================", 
                             as.character(unlist(as.vector(t(sem_cnpj)))), 
                             "=======================================", 
                             "---> OS DEMAIS ARQUIVOS FORAM IMPORTADOS CORRETAMENTE.")
      log_info(mensagem_sem_cnpj, cores = "vermelho")
    } else { 
      mensagem_sem_cnpj <- NULL
    }
    
    if (nrow(cnpj_errado) > 0) {
      mensagem_cnpj_errado <- c("OS SEGUINTES FORNECEDORES TEM PROBLEMAS NO CNPJ",
                                "=======================================", 
                                as.character(unlist(as.vector(t(cnpj_errado)))),
                                "=======================================", 
                                "---> VERIFIQUE OS ARQUIVOS E TENTE NOVAMENTE.")
      log_info(mensagem_cnpj_errado, cores = "vermelho")
    } else {
      mensagem_cnpj_errado <- NULL
    }
    
    tryCatch({
      CON <- file(arquivo_alerta, "w")
        if (!is.null(mensagem_sem_cnpj)) writeLines(mensagem_sem_cnpj, CON)
        if (!is.null(mensagem_sem_cnpj) & !is.null(mensagem_cnpj_errado)) writeLines(" ", CON)
        if (!is.null(mensagem_cnpj_errado)) writeLines(mensagem_cnpj_errado, CON)
      close(CON)
    },
    error = function(e) { 
      log_erro("Não foi possível salvar o arquivo de alerta", 
               e,
               alerta = TRUE) 
    })
  
    dados$fornecedores <- dados$fornecedores[!is.na(dados$fornecedores$cnpj), ]
    dados$fornecedores <- dados$fornecedores[dados$fornecedores$CNPJ_OK == TRUE, ]
    
    log_erro("Alguns arquivos estão com problemas. Verifique e resolva ou exclua-os e tente novamente.", 
             alerta = TRUE)
  }
  
  return(dados)  
}

fornecedores_dados_limpar <- function(dados) {
  #substitui NA por - em complemento e bairro e "não informado" em razao social, endereço, município e contato
  dados$fornecedores$razao_social <- replace(dados$fornecedores$razao_social, is.na(dados$fornecedores$razao_social), "NAO INFORMADO")
  dados$fornecedores$complemento  <- replace(dados$fornecedores$complemento,  is.na(dados$fornecedores$complemento),  "-")
  dados$fornecedores$bairro       <- replace(dados$fornecedores$bairro,       is.na(dados$fornecedores$bairro),       "-")
  dados$fornecedores$endereco     <- replace(dados$fornecedores$endereco,     is.na(dados$fornecedores$endereco),     "NAO INFORMADO")
  dados$fornecedores$nome_mun     <- replace(dados$fornecedores$nome_mun,     is.na(dados$fornecedores$nome_mun),     "NAO INFORMADO")
  dados$fornecedores$nome_contato <- replace(dados$fornecedores$nome_contato, is.na(dados$fornecedores$nome_contato), "NAO INFORMADO")
  
  #preenche telefone se foi preenchido em ddd
  dados$fornecedores$telefone <- replace(dados$fornecedores$telefone, is.na(dados$fornecedores$telefone), dados$fornecedores$ddd[is.na(dados$fornecedores$telefone)])
  
  #elimina quebras de linha e torna tudo maiúsculo...
  for (i in 1:14) {
    dados$fornecedores[ , i] <- gsub("[\r\n]", "", dados$fornecedores[ , i])
    dados$fornecedores[ , i] <- toupper(dados$fornecedores[ , i])
  }
  #...exceto e-mail, que será minúsculo
  #...e apenas um
  #se for inválido, informa um provisório
  dados$fornecedores$e_mail   <- tolower(dados$fornecedores$e_mail)
    dados$fornecedores$e_mail <- str_extract(dados$fornecedores$e_mail, pattern = "\\b[-A-Za-z0-9_.%]+\\@[-A-Za-z0-9_.%]+\\.[A-Za-z]+")
    dados$fornecedores$e_mail <- replace(dados$fornecedores$e_mail, is.na(dados$fornecedores$e_mail), "inv@ali.do")
  
  #manter somente dígitos do CNPJ, CEP, DDD, telefone, banco, agência e conta
  dados$fornecedores$cnpj       <- gsub("[^[:digit:]]", "", dados$fornecedores$cnpj)
    dados$fornecedores$cnpj     <- str_pad(dados$fornecedores$cnpj, width = 14, side = "left", pad = "0")
    dados$fornecedores$cnpj     <- str_sub(dados$fornecedores$cnpj, 1, 14)
    dados$fornecedores$cnpj     <- gsub("00000000000000", NA, dados$fornecedores$cnpj)
  dados$fornecedores$cep        <- gsub("[^[:digit:]]", "", dados$fornecedores$cep)
    dados$fornecedores$cep[which(dados$fornecedores$cep == "")] <- "01001000"
    dados$fornecedores$cep      <- str_pad(dados$fornecedores$cep, width = 8, side = "left", pad = "0")
  dados$fornecedores$ddd        <- gsub("[^[:digit:]]", "", dados$fornecedores$ddd)
    dados$fornecedores$ddd      <- str_sub(dados$fornecedores$ddd, 1, 2)
  dados$fornecedores$telefone   <- gsub("[^[:digit:]]", "", dados$fornecedores$telefone)
    dados$fornecedores$telefone <- str_sub(dados$fornecedores$telefone, -9)
  dados$fornecedores$banco      <- gsub("[^[:digit:]]", "", dados$fornecedores$banco)
  dados$fornecedores$agencia    <- gsub("[^[:digit:]|^[:punct:]]", "", dados$fornecedores$agencia)
  dados$fornecedores$conta      <- gsub("[^[:digit:]^[:punct:]^X]", "", dados$fornecedores$conta)
    
  #formatar banco com 3 dígitos
  #preenche banco vazio ou 000 com 001
  #capturar apenas os dígitos da agência antes do hífen ou outra pontuação
  #formatar agencia com no máximo 4 dígitos
  #preenche agencia vazia ou 0 com 1
  dados$fornecedores$banco   <- str_pad(dados$fornecedores$banco, width = 3, side = "left", pad = "0")
  dados$fornecedores$banco[which(dados$fornecedores$banco == "000" | is.na(dados$fornecedores$banco))] <- "001"
  dados$fornecedores$agencia <- str_extract(dados$fornecedores$agencia, "[:digit:]+(?=[:punct:])|[:digit:]+")
  dados$fornecedores$agencia <- str_sub(dados$fornecedores$agencia, 1, 4)
  dados$fornecedores$agencia[which(as.numeric(dados$fornecedores$agencia) == 0 | is.na(dados$fornecedores$agencia))] <- "1"
  #se o banco for importado do Excel como "1.0" 
  for (k in 1:nrow(dados$fornecedores)) {
    if (dados$originais$banco[k] == "1.0" && dados$fornecedores$banco[k] == "010") {
      dados$fornecedores$banco[k] <- "001"
    }
  }
  
  #deixar municipios sem acentos e sem o estado
  dados$fornecedores$nome_mun <- gsub("*(\\s|/|-)+[A-z][A-z]$", "", dados$fornecedores$nome_mun)
  dados$fornecedores$nome_mun <- stri_trans_general(dados$fornecedores$nome_mun, 'latin-ascii')

  return(dados)
}

fornecedores_dados_obter <- function(excel_dados, excel_fornecedores) {
  #cria variáveis
  cnpj         <- NULL
  razao_social <- NULL
  endereco     <- NULL
  complemento  <- NULL
  bairro       <- NULL
  nome_mun     <- NULL
  cep          <- NULL
  ddd          <- NULL
  telefone     <- NULL
  nome_contato <- NULL
  e_mail       <- NULL
  banco        <- NULL
  agencia      <- NULL
  conta        <- NULL
  
  #obtém dados para as variáveis
  for (i in seq_along(excel_dados)) {
    cat(sprintf(" - Lendo arquivo %s\n", basename(excel_fornecedores[i])))
    
    tryCatch({
      cnpj[i]         <- excel_dados[[i]]$informacao[1]
        if (is.na(cnpj[i])) {
          log_erro(sprintf("Problemas no CNPJ do arquivo %s", basename(excel_fornecedores[i])))
        }
      razao_social[i] <- excel_dados[[i]]$informacao[2]
      endereco[i]     <- excel_dados[[i]]$informacao[3]
      complemento[i]  <- excel_dados[[i]]$informacao[4]
      bairro[i]       <- excel_dados[[i]]$informacao[5]
      nome_mun[i]     <- excel_dados[[i]]$informacao[6]
      cep[i]          <- excel_dados[[i]]$informacao[8]
      ddd[i]          <- excel_dados[[i]]$informacao[9]
      telefone[i]     <- excel_dados[[i]]$informacao[10]
      nome_contato[i] <- excel_dados[[i]]$informacao[11]
      e_mail[i]       <- excel_dados[[i]]$informacao[12]
      banco[i]        <- excel_dados[[i]]$informacao[13]
      agencia[i]      <- excel_dados[[i]]$informacao[14]
        if (is.na(agencia[i])) {
          log_erro(sprintf("Problemas na Agência do arquivo %s", basename(excel_fornecedores[i])))
        }
      conta[i]        <- excel_dados[[i]]$informacao[15]
        if (is.na(conta[i])) {
          log_erro(sprintf("Problemas na conta do arquivo %s", basename(excel_fornecedores[i])))
        }
    },
    error = function(e) { 
      log_erro(sprintf("Há algum problema nos dados do arquivo %s", basename(excel_fornecedores[i])), 
               e) 
    })
  }

  #cria dataframe
  tryCatch({
    fornecedores <- data.frame(excel_fornecedores, cnpj, razao_social, 
                               endereco, complemento, bairro, nome_mun, cep, 
                               ddd, telefone, nome_contato, e_mail, 
                               banco, agencia, conta)
    originais <- fornecedores
    
    return(list(fornecedores = fornecedores,
                originais    = originais))
  },
  error = function(e) { 
    log_erro("Não foi possível compilar todos os dados. Verifique os arquivos com erro e tente novamente. Encerrando...", 
             e,
             finalizar = TRUE) 
  })
}

fornecedores_municipios_obter <- function() {  
  
  pasta <- get_config("pasta")
  
  #recuperar dados dos municipios
  tryCatch({
    municipios <- read.csv(file.path(pasta$fontes, "municipios.csv"), 
                           header = TRUE, 
                           sep = ";", 
                           encoding = "UTF-8", 
                           colClasses = c("numeric", "character", "NULL"))
    #deixar municipios em minúsculas, sem acentos e sem o estado
    municipios$nome_mun  <- toupper(municipios$nome_mun)
    municipios$nome_mun  <- stri_trans_general(municipios$nome_mun, 'latin-ascii')
  
    return(municipios)
  },
  error = function(e) {
    log_erro("Não foi possível vincular os códigos dos municípios. Verifique se o arquivo municipios.csv está na pasta _fontes.", 
             e,
             finalizar = TRUE) 
  })
}

fornecedores_main <- function() {
  
  source(file.path("..", "_common", "config.R"), chdir = TRUE)
  
  pasta          = list()
  pasta$atual    = getwd()
  pasta$fontes   = file.path(pasta$atual, "_fontes")
  pasta$dados    = file.path(pasta$atual, Sys.getenv("FORNECEDORES_DADOS"))
  pasta$importar = file.path(pasta$atual, 
                             Sys.getenv("FORNECEDORES_PARA IMPORTAR"))
  pasta$criar    = c(pasta$dados, pasta$importar)
  
  pacotes = c("openxlsx", "readxl", "stringi", "stringr")
  
  config_inicializar(pacotes, pasta)
  
  pregao         <- commandArgs(trailingOnly = TRUE)
  arquivo_alerta <- file.path(pasta$importar, paste0("PE_", pregao, "_ERRO.txt"))
  if (file.exists(arquivo_alerta)) {
    unlink(arquivo_alerta, recursive = TRUE)
  }

  log_secao("OBTENDO LISTA DE ARQUIVOS")

  excel_fornecedores <- fornecedores_excel_fornecedores_obter(pregao)
  
  log_secao("OBTENDO CONTEÚDO DOS ARQUIVOS")
  
  excel_dados <- fornecedores_excel_fornecedores_ler(excel_fornecedores)
  
  log_secao("LENDO DADOS DOS FORNECEDORES")
  
  dados <- fornecedores_dados_obter(excel_dados, excel_fornecedores)

  log_secao("ANALISANDO E LIMPANDO INFORMAÇÕES")
  
  dados <- fornecedores_dados_limpar(dados)
  
  log_secao("OBTENDO INFORMAÇÕES DE MUNICÍPIOS")
  
  municipios <- fornecedores_municipios_obter()
  
  log_secao("VERIFICANDO CNPJs")
  
  dados <- fornecedores_cnpj_verificar(dados, arquivo_alerta)

  log_secao("MONTANDO ARQUIVO")

  dados_a_importar <- fornecedores_arquivo_a_importar_montar(dados, municipios)
  
  log_secao("SALVANDO ARQUIVOS")

  fornecedores_arquivo_a_importar_salvar(dados_a_importar, pregao)  
  fornecedores_arquivo_comparacao_salvar(dados, dados_a_importar, pregao)
  
  config_finalizar()
}

fornecedores_main()