# ---- UTILITÁRIOS ----
# Funções utilitárias gerais para uso pelos scripts que não sejam
# de configuração ou logging

utils_catalogo <- function() {
  #' Obtém dados do Catálogo de Materiais
  #'
  #' @description
  #' Acessa e lê a planilha do Catálogo de Materiais através do link configurado
  #' nas opções. A função realiza a autenticação e obtém os dados dos grupos
  #' e suas descrições.
  #'
  #' @returns Data frame contendo os dados do catálogo de materiais,
  #' incluindo uma coluna adicional 'grupoDescricao' que concatena
  #' o grupo e sua descrição.
  #'
  #' @examples
  #' # Obtém os dados do catálogo
  #' catalogo <- utils_catalogo()
  #'
  #' @seealso \code{\link{get_config}}
  
  log_secao("OBTENDO GRUPOS DO CATÁLOGO DE MATERIAIS", "UTILS")
  
  url_catalogo <- get_config("url")$catalogo
  
  log_info("Planilha: Catálogo de Materiais",
           paste0("Link: ", url_catalogo),
           cores = "verde")
  
  gs4_deauth()
  
  tryCatch({
    catalogo <- range_read(url_catalogo) %>%
      mutate(grupoDescricao = paste(Grupo, "-", `Descrição Grupo`))
  },
  error = function(e) { 
    log_erro("Erro ao acessar Catálogo de Materiais. Verifique a conexão. 
             Encerrando...", 
             e,
             finalizar = TRUE) 
  })
  
  if (nrow(catalogo) == 0) {
    log_erro("Não há informações na planilha do Catálogo. Encerrando...", 
             finalizar = TRUE) 
  }
  
  return(catalogo)
  
}

utils_color <- function(color = "azul") {
  #' Define cores do console (somente Windows)
  #' 
  #' @description
  #' Controla as cores de texto e fundo do console do Windows durante a execução
  #' do script. A função verifica automaticamente se está rodando no Windows
  #' e se não está no modo silencioso antes de aplicar as cores.
  #' 
  #' Esta função centraliza todo o controle de cores usado pelas funções de log
  #' e configuração, garantindo consistência visual e facilitando manutenção.
  #' 
  #' Em sistemas não-Windows ou no modo silencioso, a função não faz nada,
  #' garantindo compatibilidade multiplataforma.
  #'
  #' @param color Character. Nome da cor/estilo a ser aplicado. 
  #' Opções disponíveis:
  #' \itemize{
  #'   \item{"inicial"}{: Fundo amarelo com texto azul (padrão do sistema)}
  #'   \item{"verde"}{: Texto verde}
  #'   \item{"vermelho"}{: Texto vermelho}
  #'   \item{"azul"}{: Texto azul (padrão)}
  #'   \item{"cinza_fundo"}{: Fundo cinza}
  #'   \item{"encerrar_erro"}{: Fundo vermelho com texto amarelo}
  #'   \item{"encerrar_alerta"}{: Fundo cinza com texto azul}  
  #'   \item{"encerrar_ok"}{: Fundo verde com texto amarelo}
  #' }
  #' Se color for NULL ou uma cor não reconhecida, usa "azul" como padrão.
  #'
  #' @details
  #' A função utiliza códigos de escape ANSI através do comando shell do Windows
  #' para alterar as cores do console. Os códigos seguem o padrão:
  #' \itemize{
  #'   \item{ESCXXm para cores de texto}
  #'   \item{ESCXXmESCYYm para combinações de fundo e texto}
  #' }
  #' 
  #' A verificação de plataforma garante que os comandos shell só sejam
  #' executados no Windows, evitando erros em outros sistemas operacionais.
  #'
  #' @returns 
  #' Invisível. A função não retorna valor, apenas modifica as cores do console
  #' como efeito colateral. Em sistemas não-Windows ou modo silencioso,
  #' não há efeito visível.
  #'
  #' @examples
  #' # Aplica cor verde ao texto
  #' utils_color("verde")
  #' cat("Este texto aparecerá em verde")
  #' 
  #' # Volta à cor padrão
  #' utils_color("azul")
  #' 
  #' # Aplica estilo de erro (fundo vermelho, texto amarelo)
  #' utils_color("erro_final")
  #' cat("ERRO: Mensagem importante")
  #' 
  #' # Restaura padrão do sistema
  #' utils_color("inicial")
  #'
  #' @seealso \code{\link{utils_silent}}, \code{\link{log_erro}}, 
  #' \code{\link{log_info}}
  #' 
  #' @references 
  #' \href{https://ss64.com/nt/syntax-ansi.html}{How-to: Use ANSI colours 
  #' in the terminal}
  
  if (.Platform$OS.type == "windows" && !(utils_silent())) {
    colors <- list(
      "inicial"         = "@ECHO \033[103m\033[34m", #fundo amarelo texto azul
      "verde"           = "@ECHO \033[32m", 
      "vermelho"        = "@ECHO \033[31m",
      "azul"            = "@ECHO \033[34m", #padrão
      "cinza_fundo"     = "@ECHO \033[47m", #fundo cinza texto padrão
      "encerrar_erro"   = "@ECHO \033[41m\033[93m", #fundo vermelho txt amarelo
      "encerrar_alerta" = "@ECHO \033[47m\033[34m", #fundo cinza texto azul
      "encerrar_ok"     = "ECHO \033[42m\033[93m" #fundo verde texto amarelo
    )
    if (!is.null(color) && color %in% names(colors)) {
      shell(colors[[color]])
    } else {
      shell(colors[["azul"]])
    }
  }
}

utils_corrigir_valor <- function(valor) {
  #' Corrige formato de valores numéricos
  #'
  #' @description
  #' Converte uma string representando um valor numérico para o formato correto,
  #' tratando separadores decimais e de milhares. Remove pontos de separação
  #' de milhares e converte vírgulas decimais em pontos.
  #'
  #' @param valor Character. String contendo o valor numérico a ser corrigido.
  #'
  #' @returns Numeric. O valor convertido para formato numérico.
  #'
  #' @examples
  #' # Corrige valor com vírgula decimal
  #' utils_corrigir_valor("1.234,56")  # Retorna 1234.56
  #' 
  #' # Corrige valor sem decimais
  #' utils_corrigir_valor("1.234")     # Retorna 1234
  
  as.numeric(gsub(",", ".", gsub("\\.", "", valor)))
  
}

utils_silent <- function() {
  #' Verifica se o script está executando em modo silencioso
  #'
  #' @description
  #' Determina se o script foi chamado com o argumento de linha de comando 
  #' "silent", que indica execução em segundo plano ou sem interface visual.
  #' 
  #' Esta função centraliza a verificação do modo silencioso, evitando
  #' duplicação de código em várias funções que precisam adaptar seu
  #' comportamento quando rodando sem interação do usuário.
  #' 
  #' Quando em modo silencioso, as funções do sistema tipicamente:
  #' \itemize{
  #'   \item{Não aplicam cores ao console}
  #'   \item{Usam barras de progresso em modo texto}
  #'   \item{Reduzem mensagens visuais}
  #'   \item{Evitam comandos que requerem interação}
  #' }
  #'
  #' @returns 
  #' Logical. TRUE se o script está rodando em modo silencioso 
  #' (argumento "silent" foi fornecido na linha de comando), 
  #' FALSE caso contrário.
  #'
  #' @details
  #' A função verifica os argumentos de linha de comando usando 
  #' \code{commandArgs(trailingOnly = TRUE)} e procura pela string "silent".
  #' 
  #' O modo silencioso é útil quando:
  #' \itemize{
  #'   \item{O script é executado via agendador de tarefas}
  #'   \item{Está rodando em um servidor sem interface gráfica}  
  #'   \item{Faz parte de um pipeline automatizado}
  #'   \item{Precisa rodar sem intervenção do usuário}
  #' }
  #'
  #' @examples
  #' # Verifica se está em modo silencioso
  #' if (utils_silent()) {
  #'   cat("Executando em modo silencioso\n")
  #' } else {
  #'   cat("Executando em modo interativo\n")
  #' }
  #' 
  #' # Uso típico em outras funções
  #' if (!utils_silent()) {
  #'   # Aplica cores ou mostra interface visual
  #'   utils_color("verde")
  #' }
  #' 
  #' # Exemplo de chamada na linha de comando que ativaria o modo:
  #' # Rscript meu_script.R silent
  #'
  #' @seealso \code{\link{utils_color}}, \code{\link{log_barra_progresso}},
  #' \code{\link{commandArgs}}
  
  "silent" %in% commandArgs(trailingOnly = TRUE)
  
}

utils_unidades_requerentes_obter <- function(origem = "importacao") {
  #' Obtém dados das unidades requerentes
  #'
  #' @description
  #' Acessa e lê a planilha de Unidades Requerentes através do link configurado
  #' nas opções. A função pode ler dados de duas abas diferentes da planilha,
  #' dependendo da origem especificada.
  #'
  #' @param origem Character. Define qual aba da planilha será lida.
  #' Valores possíveis: \code{"importacao"} (padrão) ou \code{"power bi"}.
  #'
  #' @returns Data frame contendo os dados das unidades requerentes.
  #' Se origem for "importacao", retorna colunas padronizadas incluindo
  #' sigla_solar, imovel, cpf e email.
  #' Se origem for "power bi", verifica duplicidade de siglas no Solar.
  #'
  #' @examples
  #' # Obtém dados para importação
  #' unidades <- utils_unidades_requerentes_obter()
  #' 
  #' # Obtém dados para Power BI
  #' unidades_bi <- utils_unidades_requerentes_obter("power bi")
  #'
  #' @seealso \code{\link{get_config}}
  
  log_secao("OBTENDO DADOS DAS UNIDADES REQUERENTES", "UTILS")
  
  url_unidades <- get_config("url")$unidades
  
  log_info("Planilha: Unidades - espelhada",
           paste0("Link: ", url_unidades),
           paste0("Aba : ", ifelse(origem == "importacao", 
                                   "Script Importação", 
                                   "Power BI")),
           cores = "verde")
  
  gs4_deauth()
  
  tryCatch({
    if (origem == "importacao") {
      
      unidades <- range_read(url_unidades, 
                             sheet = "Script Importação", 
                             col_types = "cccccc")
      
      colnames(unidades)[c(3,4,5,6)] <- c("sigla_solar", "imovel", 
                                          "cpf", "email")
      
    } else if (origem == "power bi") {
      
      unidades <- range_read(url_unidades, 
                             sheet = "Power BI")
      
      sigla_solar_duplicidade <- unidades %>%
        group_by(`Sigla no Solar`) %>%
        filter(n() > 1) %>%
        mutate(siglas_duplicidade = paste(Setor, "-", `Sigla no Solar`)) %>%
        arrange(`Sigla no Solar`) %>%
        pull(siglas_duplicidade)
      
      if (length(sigla_solar_duplicidade) > 0) {
        log_erro("Há Unidades com siglas no Solar repetidas.Isto causará problemas no Painel do Observatório.", 
                 "Mantenha somente siglas únicas na aba 'Power Bi'.")
        log_erro("As Unidades que precisam ser ajustadas são:", 
                 sigla_solar_duplicidade)
      }
      
    }
    return(unidades)
  },
  error = function(e) { 
    log_erro("Erro ao acessar planilha de Unidades requerentes. Encerrando...", 
             e,
             finalizar = TRUE) 
  })
}

utils_verificar_cnpj <- function(cnpj) {
  #' Valida número de CNPJ
  #'
  #' @description
  #' Verifica se um CNPJ é válido aplicando as regras de validação oficiais,
  #' incluindo verificação dos dígitos verificadores. A função aceita CNPJ
  #' com ou sem formatação (pontos, traços e barra).
  #'
  #' @param cnpj Character. Número do CNPJ a ser validado, pode conter
  #' caracteres de formatação.
  #'
  #' @returns Logical. TRUE se o CNPJ é válido, FALSE caso contrário.
  #' Retorna FALSE para CNPJs vazios, NA, com quantidade incorreta de dígitos,
  #' ou com dígitos verificadores inválidos.
  #'
  #' @examples
  #' # Valida CNPJ formatado
  #' utils_verificar_cnpj("11.222.333/0001-81")
  #' 
  #' # Valida CNPJ sem formatação
  #' utils_verificar_cnpj("11222333000181")
  #' 
  #' # Valida CNPJ inválido
  #' utils_verificar_cnpj("11.222.333/0001-00")
  
  # Verifica se o CNPJ é NA ou vazio ou NULL
  if (is.na(cnpj) || cnpj == "" || is.null(cnpj)) return(FALSE) 
  
  # Remove caracteres não numéricos
  cnpj_limpo <- gsub("[^0-9]", "", cnpj)
  
  # Verifica se o CNPJ tem 14 dígitos
  if (nchar(cnpj_limpo) != 14) return(FALSE)
  
  # Converte para vetor numérico
  digitos <- as.numeric(strsplit(cnpj_limpo, "")[[1]])
  
  # Verifica se todos os dígitos são iguais (ex: 00000000000000)
  if (length(unique(digitos)) == 1) return(FALSE)
  
  d1 <- NULL
  d1[1]  <- digitos[1]  * 5
  d1[2]  <- digitos[2]  * 4
  d1[3]  <- digitos[3]  * 3
  d1[4]  <- digitos[4]  * 2
  d1[5]  <- digitos[5]  * 9
  d1[6]  <- digitos[6]  * 8
  d1[7]  <- digitos[7]  * 7
  d1[8]  <- digitos[8]  * 6
  d1[9]  <- digitos[9]  * 5
  d1[10] <- digitos[10] * 4
  d1[11] <- digitos[11] * 3
  d1[12] <- digitos[12] * 2
  resto1 <- sum(d1, na.rm = TRUE) %% 11
  dv1 <- ifelse(resto1 < 2, 0, 11 - resto1)
  
  d2 <- NULL
  d2[1]  <- digitos[1]  * 6
  d2[2]  <- digitos[2]  * 5
  d2[3]  <- digitos[3]  * 4
  d2[4]  <- digitos[4]  * 3
  d2[5]  <- digitos[5]  * 2
  d2[6]  <- digitos[6]  * 9
  d2[7]  <- digitos[7]  * 8
  d2[8]  <- digitos[8]  * 7
  d2[9]  <- digitos[9]  * 6
  d2[10] <- digitos[10] * 5
  d2[11] <- digitos[11] * 4
  d2[12] <- digitos[12] * 3
  d2[13] <- dv1         * 2
  resto2 <- sum(d2, na.rm = TRUE) %% 11
  dv2 <- ifelse(resto2 < 2, 0, 11 - resto2)
  
  # Compara os dígitos verificadores calculados com os informados
  dv_informado <- substring(cnpj_limpo, 13, 14)
  dv_calculado <- paste0(dv1, dv2)
  
  return(dv_informado == dv_calculado)
}

utils_verificar_cpf <- function(cpf) {
  #' Valida número de CPF
  #'
  #' @description
  #' Verifica se um CPF é válido aplicando as regras de validação oficiais,
  #' incluindo verificação dos dígitos verificadores. A função aceita CPF
  #' com ou sem formatação (pontos e traço).
  #'
  #' @param cpf Character. Número do CPF a ser validado, pode conter
  #' caracteres de formatação.
  #'
  #' @returns Logical. TRUE se o CPF é válido, FALSE caso contrário.
  #' Retorna FALSE para CPFs vazios, NA, com quantidade incorreta de dígitos,
  #' ou com dígitos verificadores inválidos.
  #'
  #' @examples
  #' # Valida CPF formatado
  #' utils_verificar_cpf("123.456.789-09")
  #' 
  #' # Valida CPF sem formatação
  #' utils_verificar_cpf("12345678909")
  #' 
  #' # Valida CPF inválido
  #' utils_verificar_cpf("111.111.111-11")

  # Se for NA, vazio ou NULL, já retorna inválido
  if (is.na(cpf) || cpf == "" || is.null(cpf)) return(FALSE)  
  
  # Remove tudo que não for dígito e verifica se tem 11 caracteres
  cpf_limpo <- gsub("[^0-9]", "", cpf)
  if (nchar(cpf_limpo) != 11) return(FALSE)
  
  # Converte para vetor numérico
  digitos <- as.numeric(strsplit(cpf_limpo, "")[[1]])
  
  # Adicionar verificação de CPFs conhecidamente inválidos
  if (length(unique(digitos)) == 1) return(FALSE)
  
  d1    <- NULL
  d1[1] <- digitos[1]* 10
  d1[2] <- digitos[2]* 9
  d1[3] <- digitos[3]* 8
  d1[4] <- digitos[4]* 7
  d1[5] <- digitos[5]* 6
  d1[6] <- digitos[6]* 5
  d1[7] <- digitos[7]* 4
  d1[8] <- digitos[8]* 3
  d1[9] <- digitos[9]* 2
  resto1 <- do.call(sum, as.list(d1)) %% 11
  if (resto1 < 2) { 
    dv1 <- 0 
  } else { 
    dv1 <- 11 - resto1 
  }
  
  d2     <- NULL
  d2[1]  <- digitos[1]  * 11
  d2[2]  <- digitos[2]  * 10
  d2[3]  <- digitos[3]  * 9
  d2[4]  <- digitos[4]  * 8
  d2[5]  <- digitos[5]  * 7
  d2[6]  <- digitos[6]  * 6
  d2[7]  <- digitos[7]  * 5
  d2[8]  <- digitos[8]  * 4
  d2[9]  <- digitos[9]  * 3
  d2[10] <- dv1         * 2
  resto2 <- do.call(sum, as.list(d2)) %% 11
  if (resto2 < 2) { 
    dv2 <- 0 
  } else { 
    dv2 <- 11 - resto2 
  }
  
  dv_calculado <- paste0(as.character(dv1), as.character(dv2))
  dv_informado <- substr(cpf_limpo, 10, 11)

  return(dv_calculado == dv_informado)
}

utils_verificar_script <- function(scripts_permitidos, script_padrao) {
  #' Valida e obtém o script a ser executado via linha de comando
  #'
  #' @description
  #' Verifica os argumentos de linha de comando para determinar qual script
  #' deve ser executado. Se nenhum argumento for fornecido ou se o argumento
  #' não estiver na lista de scripts permitidos, retorna o script padrão.
  #'
  #' @param scripts_permitidos Character vector. Lista de nomes de scripts
  #' que são permitidos executar.
  #' @param script_padrao Character. Nome do script a ser usado como padrão
  #' quando nenhum argumento válido for fornecido.
  #'
  #' @returns Character. Nome do script a ser executado. Se nenhum argumento
  #' for fornecido ou se o argumento for inválido, retorna o script_padrao.
  #'
  #' @examples
  #' # Define scripts permitidos e padrão
  #' scripts <- c("script1.R", "script2.R", "script3.R")
  #' padrao <- "script1.R"
  #' 
  #' # Verifica script a executar
  #' script <- utils_verificar_script(scripts, padrao)
  #'
  #' @seealso \code{\link{log_erro}}
  
  argumentos <- commandArgs(trailingOnly = TRUE)
  
  if (length(argumentos) == 0) {
    log_erro(sprintf("Nenhum argumento fornecido. Usando valor padrão: '%s'", 
                     script_padrao),
             alerta = TRUE)
    return(script_padrao)
  }
  
  script_a_executar <- argumentos[1]
  
  if (!(script_a_executar %in% scripts_permitidos)) {
    log_erro(sprintf("Argumento inválido: '%s'. Usando valor padrão: '%s'", 
                     script_a_executar, script_padrao),
             alerta = TRUE)
    return(script_padrao)
  }
  
  return(script_a_executar)
  
}
