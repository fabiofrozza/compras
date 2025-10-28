atas_configuracoes_iniciais <- function() {
  
  tryCatch({
    dados_pregao <- list()
    
    dados_pregao$pregao     <- config_json("n_pregao", opcao = "get")
    dados_pregao$ano_pregao <- config_json("ano_pregao", opcao = "get")
    dados_pregao$processo   <- config_json("processo", opcao = "get")
    dados_pregao$objeto     <- config_json("objeto", opcao = "get")
    dados_pregao$data       <- config_json("data", opcao = "get")
    dados_pregao$ata        <- config_json("n_ata", opcao = "get")
    dados_pregao$ano_ata    <- config_json("ano_ata", opcao = "get")
    
    log_info("DADOS DO PREGÃO", dados_pregao,
             cores = "verde")
    
    return(dados_pregao)
  },
  error = function(e) {
    log_erro("Não foi possível recuperar os dados do pregão. Encerrando...",
             e,
             finalizar = TRUE)
  })
}

atas_salvar <- function(dados_atas) {

  pasta <- get_config("pasta")
  
  #salva os dados em arquivo .xls e verifica se ocorre erro
  tryCatch({
    write.xlsx(dados_atas, 
               file = file.path(pasta$atas, "dados_atas.xlsx"),
               sheetName = "Dados_para_Atas")
    
    log_info("INFORMAÇÕES IMPORTADAS",
             "-",
             dados_atas[ , c(6, 8, 9)],
             cores = "verde")
  },
  error = function(e) {
    log_erro("Não foi possível salvar o arquivo. Verifique o log.",
             e,
             finalizar = TRUE)
  })
}

atas_sicaf_ler <- function(sicaf_arquivos, dados_pregao) {
  #prepara variáveis
  pregao       <- dados_pregao$pregao
  ano_pregao   <- dados_pregao$ano_pregao
  processo     <- dados_pregao$processo
  objeto       <- dados_pregao$objeto
  data         <- dados_pregao$data
  ano_ata      <- dados_pregao$ano_ata
  ata          <- NULL
  cnpj         <- NULL
  razao_social <- NULL
  endereco     <- NULL
  cep          <- NULL
  cidade       <- NULL
  estado       <- NULL
  telefone     <- NULL
  email        <- NULL
  cpf          <- NULL
  responsavel  <- NULL
  nome_arquivo <- NULL
  
  #inicia contadores
  msg_erro <- list()
  
  #captura dos dados dos SICAFs
  
  pb <- log_barra_progresso("Aguarde...", length(sicaf_arquivos))
  
  for (i in seq_along(sicaf_arquivos)) {
  
    ata[i] <- as.numeric(dados_pregao$ata) - 1 + i
    log_barra_progresso(sprintf("Verificando arquivo %s === Ata %s", basename(sicaf_arquivos[i]), ata[i]), pb = pb)
    
    cat(sprintf(" === Verificando arquivo %s === Ata %s", basename(sicaf_arquivos[i]), ata[i]), "\n")
  
    tryCatch({
      pdf_texto <- pdf_text(sicaf_arquivos[i])
    
      pdf_linha_unica <- str_c(pdf_texto, collapse = "\n")
      pdf_responsavel <- str_extract(pdf_linha_unica, "(?=Dados do Responsável Legal)(?s)(.*?)($)")
    
      cnpj[i] <- str_extract(pdf_linha_unica, "\\d\\d\\.\\d\\d\\d\\.\\d\\d\\d/\\d\\d\\d\\d-\\d\\d")
        if (is.na(cnpj[i])) {
          msg_erro[[length(msg_erro) + 1]] <- paste("CNPJ do pdf", sicaf_arquivos[i])
        } else {
          cat(sprintf("     CNPJ: %s", cnpj[i]), "\n")
        }
      
      razao_social_previo   <- str_trim(str_extract(pdf_linha_unica, "(?<=Social:)(?s)(.*?)(?=Nome Fantasia)"))
        razao_social_previo <- str_replace_all(razao_social_previo, "\\\\n|/", "")
        razao_social[i]     <- str_to_upper(str_squish(razao_social_previo))
        nome_arquivo[i]     <- str_replace_all(razao_social[i], "[\\\\/:*?\"<>|]", "")
        if (is.na(razao_social[i])) {
          msg_erro[[length(msg_erro) + 1]] <- paste("RAZÃO SOCIAL do pdf", sicaf_arquivos[i])
        } else {
          cat(sprintf("     Razão social: %s", razao_social[i]), "\n\n")
        }
        
      endereco[i] <- str_to_upper(str_trim(str_extract(pdf_linha_unica, "(?<=Endereço:)(?s)(.*?)(?=\\n|\\\\n)")))
        if (is.na(endereco[i])) {
          msg_erro[[length(msg_erro) + 1]] <- paste("ENDEREÇO do pdf", sicaf_arquivos[i])
        }
    
      cep[i] <- str_extract(pdf_linha_unica, "\\d\\d\\.\\d\\d\\d\\-\\d\\d\\d")
        if (is.na(cep[i])) {
          msg_erro[[length(msg_erro) + 1]] <- paste("CEP do pdf", sicaf_arquivos[i])
        }
    
      cidade[i] <- str_trim(str_extract(pdf_linha_unica, "(?<=Município / UF:)(?s)(.*?)(?=/)"))
        if (is.na(cidade[i])) {
          msg_erro[[length(msg_erro) + 1]] <- paste("CIDADE do pdf", sicaf_arquivos[i])
        }
      
      estado_previo   <- str_trim(str_extract(pdf_linha_unica, "(?<=Município / UF:)(?s)(.*?)(?=\\n|\\\\n)"))
        estado_previo <- str_extract(estado_previo, "(?<=/ )(.+)")
        estado[i] = case_match(
          estado_previo,
          "Acre"                ~ "AC",
          "Alagoas"             ~ "AL",
          "Amapá"               ~ "AP",
          "Amazonas"            ~ "AM",
          "Bahia"               ~ "BA",
          "Ceará"               ~ "CE",
          "Espírito Santo"      ~ "ES",
          "Goiás"               ~ "GO",
          "Maranhão"            ~ "MA",
          "Mato Grosso"         ~ "MT",
          "Mato Grosso do Sul"  ~ "MS",
          "Minas Gerais"        ~ "MG",
          "Pará"                ~ "PA",
          "Paraíba"             ~ "PB",
          "Paraná"              ~ "PR",
          "Pernambuco"          ~ "PE",
          "Piauí"               ~ "PI",
          "Rio de Janeiro"      ~ "RJ",
          "Rio Grande do Norte" ~ "RN",
          "Rio Grande do Sul"   ~ "RS",
          "Rondônia"            ~ "RO",
          "Roraima"             ~ "RR",
          "Santa Catarina"      ~ "SC",
          "São Paulo"           ~ "SP",
          "Sergipe"             ~ "SE",
          "Tocantins"           ~ "TO",
          "Distrito Federal"    ~ "DF")
          if (is.na(estado[i])) {
            msg_erro[[length(msg_erro) + 1]] <- paste("ESTADO do pdf", sicaf_arquivos[i])
          }
        
      telefone[i] <- str_trim(str_extract(pdf_linha_unica, "(?<=Telefone:)(?s)(.*?)(?=\\n|\\\\n|Telefone)"))
        if (is.na(telefone[i])) {
          msg_erro[[length(msg_erro) + 1]] <- paste("TELEFONE do pdf", sicaf_arquivos[i])
        }
      
      email_previo <- as.list(str_extract_all(pdf_linha_unica, "(?<=E-mail:)(?s)(.*?)(?=\\n|\\\\n)"))
        for (j in 1:length(email_previo[[1]])) {
          email_previo[[1]][j] <- str_to_lower(str_trim(email_previo[[1]][j]))
        }
        email[i] <- paste(unique(email_previo[[1]][nchar(email_previo[[1]]) != 0]), collapse = ", ")
        if (is.na(email[i])) {
          msg_erro[[length(msg_erro) + 1]] <- paste("E-MAIL do pdf", sicaf_arquivos[i])
        }
        
      cpf[i] <- str_trim(str_extract(pdf_responsavel, "(?<=Dados do Responsável Legal\\\\nCPF:|Dados do Responsável Legal\\nCPF:)(?s)(.*?)(\\d\\d\\d\\.\\d\\d\\d\\.\\d\\d\\d-\\d\\d)"))
        if (is.na(cpf[i])) {
          msg_erro[[length(msg_erro) + 1]] <- paste("CPF do pdf", sicaf_arquivos[i])
        }
      
      responsavel[i] <- str_trim(str_extract(pdf_responsavel, "(?<=Nome:)(?s)(.*?)(?=\\n|\\\\n)"))
        if (is.na(responsavel[i])) {
          msg_erro[[length(msg_erro) + 1]] <- paste("RESPONSÁVEL do pdf", sicaf_arquivos[i])
        }
    },
    error = function(e) {
      log_erro(paste("Não foi possível ler o arquivo", basename(sicaf_arquivos[i])), 
               e)
    })
  }
  log_barra_progresso(pb = pb)
  
  if (length(msg_erro) > 0) {
    log_erro("Houve problemas na captura dos dados:", 
             unlist(msg_erro),
             finalizar = TRUE)
  } 
  
  #cria dataframe
  dados_atas <- data.frame(pregao, ano_pregao, 
                           processo, objeto, data, 
                           ata, ano_ata, 
                           cnpj, razao_social, 
                           endereco, cep, cidade, estado, 
                           telefone, email, 
                           cpf, responsavel,
                           nome_arquivo)

  return(dados_atas)
}

atas_sicaf_obter <- function() {
  
  pasta <- get_config("pasta")
  
  tryCatch({
    #cria lista dos SICAFs em ordem alfabética
    sicaf_arquivos <- Sys.glob(file.path(pasta$sicaf, "*.pdf*"))
    
    log_info("RELATÓRIOS SICAF", sicaf_arquivos,
             cores = "verde")

    return(sicaf_arquivos)
  }, 
  error = function(e) {
    log_erro("Não foi possível obter os relatórios SICAF. Encerrando...",
             e,
             finalizar = TRUE)
  })
}

atas_verificar_duplicados <- function(dados_atas) {

  #verifica se há arquivos duplicados para emitir alerta
  duplicados <- dados_atas[dados_atas$cnpj %in% dados_atas$cnpj[duplicated(dados_atas$cnpj)], c(6,8:9)]
  
  if (nrow(duplicados) > 0) {
    log_erro("Há arquivos duplicados:", 
             duplicados,
             alerta = TRUE)
    
    return(duplicados)
  } else {
    return(NULL)
  }
}

atas_main <- function() {

  source(file.path("..", "_common", "config.R"), chdir = TRUE)
  
  pasta       = list()
  pasta$atual = getwd()
  pasta$atas  = file.path(pasta$atual, Sys.getenv("ATAS_ATAS"))
  pasta$sicaf = file.path(pasta$atual, Sys.getenv("ATAS_SICAF"))
  pasta$criar = c(pasta$atas, pasta$sicaf)
  
  pacotes     = c("openxlsx", "dplyr", "pdftools", "stringr")
  
  config_inicializar(pacotes, pasta)
  
  log_secao("RECUPERANDO DADOS DO PREGÃO")

  dados_pregao <- atas_configuracoes_iniciais()
  
  log_secao("ANALISANDO PASTA SICAF POR RELATÓRIOS")
  
  sicaf_arquivos <- atas_sicaf_obter()

  log_secao("CARREGANDO DADOS DOS RELATÓRIOS EM PDF")
  
  dados_atas <- atas_sicaf_ler(sicaf_arquivos, dados_pregao)
  
  log_secao("VERIFICANDO ARQUIVOS DUPLICADOS")
  
  duplicados <- atas_verificar_duplicados(dados_atas)
  
  log_secao("SALVANDO PLANILHA DADOS_ATAS.XLSX NA PASTA ATAS")
  
  atas_salvar(dados_atas)
  
  config_finalizar(sucesso = TRUE)
}

atas_main()