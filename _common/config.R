# encoding: UTF-8

# Para chamar este arquivo, utilize
# source(file.path("..", "_common", "config.R"), chdir = TRUE)

# ---- CONFIGURAÇÕES ----
# Funções para início dos ambientes para execução dos scripts e 
# para gravação e recuperação de variáveis de ambiente

config_main <- function() {
  #' Inicializa o ambiente de configuração principal
  #'
  #' @description
  #' Função principal que prepara o ambiente de configuração inicial. 
  #' Esta é a primeira função a ser executada, sendo chamada automaticamente 
  #' ao carregar o arquivo, e realiza a configuração básica necessária 
  #' para todas as outras funções.
  #'
  #' @details
  #' A função executa as seguintes tarefas em ordem:
  #' \itemize{
  #'   \item{Cria o ambiente \code{.config_env} como um novo ambiente vazio
  #'   (usando \code{emptyenv()} como pai) para armazenar todas as variáveis
  #'   globais utilizadas durante a execução dos scripts}
  #'   \item{Obtém o diretório de pacotes do R através da variável de ambiente
  #'   \code{R_LIBS_USER} e verifica se está definida}
  #'   \item{Cria o diretório de pacotes se não existir, usando 
  #'   \code{config_pasta()}}
  #'   \item{Verifica, instala se necessário e carrega o pacote \code{jsonlite}.
  #'   Esta dependência é crítica pois é usada para ler e gravar as 
  #'   configurações no arquivo \code{config.json}}
  #'   \item{Inicializa o ambiente com status inicial usando 
  #'   \code{set_config()}}
  #' }
  #'
  #' Em caso de erro, especialmente na instalação do jsonlite:
  #' \itemize{
  #'   \item{Altera a cor do console para vermelho em sistemas Windows}
  #'   \item{Exibe mensagem detalhada do erro}
  #'   \item{Apresenta mensagem visual de erro com ASCII art}
  #'   \item{Pausa a execução esperando input do usuário}
  #'   \item{Encerra o script com status de erro (1)}
  #' }
  #'
  #' @returns Invisível. A função não retorna valor, 
  #' mas como efeitos colaterais:
  #' \itemize{
  #'   \item{Cria e inicializa o ambiente global \code{.config_env}}
  #'   \item{Cria o diretório de pacotes se necessário}
  #'   \item{Instala e carrega o pacote jsonlite se necessário}
  #'   \item{Em caso de sucesso, define status.inicio = TRUE via 
  #'   \code{set_config()}}
  #'   \item{Em caso de erro, encerra a execução com status 1}
  #' }
  #'
  #' @examples
  #' # A função é chamada automaticamente ao carregar o arquivo
  #' # mas pode ser chamada manualmente se necessário
  #' config_main()
  #'
  #' @seealso 
  #' \code{\link{set_config}} para gerenciar configurações no ambiente
  #' \code{\link{config_pasta}} para criar diretórios
  #' \code{\link{utils_silent}} para verificar modo silencioso
  
  # Mensagem de início com pilha de chamadas
  cat("►►► INICIANDO SCRIPT R. AGUARDE...\n → ")
  cat(as.character(sys.calls()), sep = "\n → ")
  cat("►►► Inicializando ambiente de configuração...\n")
  
  # Cria ambiente .config_env para registro das variáveis utilizadas durante a
  # execução dos scripts
  if (!exists(".config_env")) {
    .config_env <<- new.env(parent = emptyenv())
  }  
  
  tryCatch({
    
    # Verifica se existe a pasta para instalação dos pacotes
    cat("►►► Verificando pacote para arquivo de configurações...\n")
    pasta_pacotes <- Sys.getenv("R_LIBS_USER")
    if (pasta_pacotes == "") {
      stop("Variável R_LIBS_USER não está definida")
    }
    
    config_pasta(pasta_pacotes)
      
    if (!require('jsonlite', lib.loc = pasta_pacotes)) {
      cat("►►► Instalando pacote jsonlite. Aguarde...\n")
      install.packages('jsonlite', 
                       lib = pasta_pacotes,
                       repos = "https://cloud.r-project.org/")
      library(jsonlite, lib.loc = pasta_pacotes)
    }
    
    set_config(.config = list(status = list(inicio = TRUE)))
    
    # Chama demais scripts compartilhados
    source("logging.R")
    source("utils.R")
    
  },
  error = function(e) {
    if (.Platform$OS.type == "windows" && !(utils_silent())) shell("@COLOR 4F")
    cat("=== OCORREU ALGUM ERRO NA INSTALAÇÃO DO PACOTE JSONLITE. 
        SEM ELE NÃO É POSSÍVEL CONTINUAR A EXECUÇÃO. ===\n")
    print(e)
    cat("\n\n\n=== CHAME O FÁBIO!!!\n")
    cat("          _\n")
    cat("         | |\n")
    cat("         | |===( )   //////\n")
    cat("         |_|   |||  | o o|\n")
    cat("                ||| ( c  )                  ____\n")
    cat("                 ||| \\= /                  ||   \\_\n")
    cat("                  ||||||                   ||     |\n")
    cat("                  ||||||                ...||__/|-\"\n")
    cat("                  ||||||             __|________|__\n")
    cat("                    |||             |______________|\n")
    cat("                    |||             || ||      || ||\n")
    cat("                    |||             || ||      || ||\n")
    cat("--------------------|||-------------||-||------||-||-------\n")
    cat("                    |__>            || ||      || ||\n")
    shell("@PAUSE")
    quit(status = 1)
  })
  
}

set_config <- function(..., .config = NULL) {
  #' Registra variáveis de ambiente na memória
  #'
  #' @description
  #' Registra uma ou mais variáveis de ambiente para serem utilizadas durante
  #' a execução do script por todas as funções.
  #' 
  #' Esta função armazena os valores somente na memória, no ambiente
  #' \code{.config_env}, e eles podem ser recuperados com a função 
  #' \code{get_config()}.
  #' 
  #' Para persistir configurações em um arquivo para uso entre diferentes 
  #' execuções, utilize a função \code{config_json()}.
  #'
  #' @param ... Argumentos nomeados a serem registrados, no formato 
  #' \code{nome = valor}.
  #' @param .config Lista. Se fornecido, substitui completamente a configuração
  #' existente pela lista informada.
  #'
  #' @returns Retorna invisivelmente (nenhum valor). 
  #' A função modifica o ambiente de configuração como efeito colateral.
  #'
  #' @examples
  #' # Registra uma variável de ambiente simples
  #' set_config(opcoes = list(primeira = "Primeira opção"))
  #'
  #' # Registra múltiplas variáveis de ambiente
  #' opcoes <- list(
  #'   primeira = "Primeira opção",
  #'   segunda = "Segunda opção"
  #' )
  #' configuracoes <- list(
  #'   config_a = "Configuração A",
  #'   config_b = "Configuração B"
  #' )
  #' set_config(opcoes = opcoes, configuracoes = configuracoes)
  #'
  #' # Substitui uma variável já registrada
  #' set_config(opcoes = list(primeira = "Outra primeira opção"))
  #'
  #' # Substitui toda a configuração
  #' config <- list(nova_configuracao = "Nova configuração")
  #' set_config(.config = config)
  #'
  #' @seealso \code{\link{get_config}}, \code{\link{config_json}}

  # Valida se o ambiente existe
  if (!exists(".config_env")) {
    stop("Ambiente .config_env não foi inicializado. 
         Execute config_main() primeiro.")
  }
  
  # Se .config for fornecido, substitui toda a configuração
  if (!is.null(.config)) {
    stopifnot(is.list(.config))
    assign("config", .config, envir = .config_env)
    return(invisible())
  }
  
  # Se não houver configuração existente, cria uma nova
  if (!exists("config", envir = .config_env)) {
    assign("config", list(), envir = .config_env)
  }
  
  # Obtém a configuração atual
  current_config <- get("config", envir = .config_env)
  
  # Processa os argumentos nomeados
  new_values <- list(...)
  
  # Atualiza a configuração com os novos valores
  for (name in names(new_values)) {
    current_config[[name]] <- new_values[[name]]
  }
  
  # Armazena a configuração atualizada
  assign("config", current_config, envir = .config_env)
  
  invisible()
}

get_config <- function(item = NULL) {
  #' Recupera variáveis de ambiente da memória
  #'
  #' @description
  #' Obtém as variáveis de ambiente registradas durante a execução do script.
  #' 
  #' Esta função recupera as variáveis armazenadas na memória no ambiente
  #' \code{.config_env}, que foram registradas com a função \code{set_config()}.
  #'
  #' @param item Character. Nome do item específico a ser recuperado. 
  #' Se NULL (padrão), retorna a lista com todas as configurações.
  #'
  #' @returns Uma lista com todas as configurações ou o valor específico do item
  #' solicitado. 
  #' Se o item solicitado não existir, retorna NULL.
  #'
  #' @examples
  #' # Primeiro, definimos algumas configurações
  #' set_config(
  #'   opcoes = list(a = 1, b = 2),
  #'   usuario = "teste"
  #' )
  #' 
  #' # Recupera todas as configurações
  #' config_completa <- get_config()
  #' 
  #' # Recupera um item específico
  #' opcoes_salvas <- get_config("opcoes")
  #'
  #' @seealso \code{\link{set_config}}, \code{\link{config_json}}

  # Valida se o ambiente existe
  if (!exists(".config_env")) {
    stop("Ambiente .config_env não foi inicializado. 
         Execute config_main() primeiro.")
  }
  
  if (!exists("config", envir = .config_env)) {
    stop("Configuração não inicializada. Use set_config() primeiro.")
  }
  
  config <- get("config", envir = .config_env)
  
  if (is.null(item)) return(config)
  
  if (item %in% names(config)) return(config[[item]])
  
}

config_json <- function(chave, valor, append = FALSE, opcao = "set", 
                        secao = NULL, max_tentativas = 3, intervalo = 0.5) {
  #' Gerencia configurações em arquivo JSON centralizado
  #'
  #' @description
  #' Função para gerenciar as configurações armazenadas em arquivo JSON 
  #' centralizado na pasta _common.
  #' Permite salvar, recuperar, adicionar e remover configurações do arquivo
  #' com sistema de retry para acesso simultâneo.
  #'
  #' @param chave Character. Nome da configuração a ser gerenciada.
  #' Ignorado se \code{opcao = "all"}
  #' @param valor Opcional. Valor a ser armazenado para a chave. 
  #' Necessário apenas quando opcao="set".
  #' @param append Logical. Se TRUE, adiciona o valor ao conteúdo existente
  #' da chave ao invés de substituí-lo. Padrão é FALSE.
  #' @param opcao Character. Operação a ser realizada:
  #' \itemize{
  #'   \item{"set"}{: Define um valor para a chave (padrão)}
  #'   \item{"get"}{: Recupera o valor da chave}
  #'   \item{"remove"}{: Remove a chave e seu valor}
  #'   \item{"all"}{: Retorna todas as configurações}
  #' }
  #' @param secao Character. Nome da seção do script no config centralizado.
  #' Se NULL, usa o nome do script atual.
  #' @param max_tentativas Numeric. Número máximo de tentativas em caso de erro.
  #' @param intervalo Numeric. Intervalo em segundos entre tentativas.
  
  # Define a seção (nome do script)
  if (is.null(secao)) {
    secao <- get_config("geral")$script_nome
  }
  
  # Caminho do arquivo centralizado
  arquivo_config <- get_config("geral")$config_centralizado
  
  # Verifica se o conteúdo está em cache e se não está desatualizado
  cache_key <- paste0("config_json_", secao)
  cache_time_key <- paste0(cache_key, "_timestamp")
  R_config_secao <- get_config(cache_key)
  cache_timestamp <- get_config(cache_time_key)
  
  # Verifica se o cache está atualizado
  file_timestamp <- if (file.exists(arquivo_config)) file.info(arquivo_config)$mtime else NULL
  cache_valid <- !is.null(R_config_secao) && 
                 !is.null(cache_timestamp) && 
                 !is.null(file_timestamp) && 
                 cache_timestamp >= file_timestamp
  
  # Função auxiliar para ler o arquivo com retry
  .ler_json <- function() {
    for (tentativa in 1:max_tentativas) {
      tryCatch({
        if (!file.exists(arquivo_config)) {
          if (!file.exists(arquivo_config)) {
            write_json(list(), arquivo_config, pretty = TRUE)
            cat("=== Arquivo config.json centralizado não localizado. 
                Gerado um novo vazio. ===\n")
          }
        } else if (file.info(arquivo_config)$size == 0) {
          write_json(list(), arquivo_config, pretty = TRUE)
          cat("=== Arquivo config.json estava vazio. 
              Inicializado com lista vazia. ===\n")
        }
        
        config_completo <- fromJSON(arquivo_config)
        
        # Se a seção não existe, cria uma vazia
        if (!(secao %in% names(config_completo))) {
          config_completo[[secao]] <- list()
        }
        
        return(config_completo)
        
      }, 
      error = function(e) {
        if (tentativa < max_tentativas) {
          Sys.sleep(intervalo)
        } else {
          stop(sprintf("Erro ao ler config.json após %d tentativas: %s", 
                       max_tentativas, e$message))
        }
      })
    }
  }
  
  # Função auxiliar para escrever no arquivo com retry
  .escrever_json <- function(config_completo) {
    for (tentativa in 1:max_tentativas) {
      tryCatch({
        # Escreve em arquivo temporário primeiro (escrita atômica)
        temp_file <- tempfile(tmpdir = dirname(arquivo_config), 
                              fileext = ".json")
        write_json(config_completo, temp_file, 
                   pretty = TRUE, auto_unbox = TRUE)
        
        # Move o arquivo temporário para o definitivo
        file.rename(temp_file, arquivo_config)
        
        return(TRUE)
        
      }, 
      error = function(e) {
        if (file.exists(temp_file)) file.remove(temp_file)
        
        if (tentativa < max_tentativas) {
          Sys.sleep(intervalo)
        } else {
          stop(sprintf("Erro ao salvar config.json após %d tentativas: %s", 
                       max_tentativas, e$message))
        }
      })
    }
  }
  
  # Se não está em cache ou está desatualizado, carrega do disco
  if (!cache_valid) {
    config_completo <- .ler_json()
    R_config_secao <- config_completo[[secao]]
    
    # Salva no cache com timestamp
    config_cache <- list()
    config_cache[[cache_key]] <- R_config_secao
    config_cache[[cache_time_key]] <- Sys.time()
    set_config(.config = modifyList(get_config(), config_cache))
  }
  
  modificado <- FALSE
  
  # Operações de leitura não alteram nada, apenas retornam o valor
  if (opcao == "all") return(R_config_secao)
  
  if (opcao == "get") return(R_config_secao[[chave]])
  
  if (opcao == "remove") {
    R_config_secao[[chave]] <- NULL
    modificado <- TRUE
  }
  
  if (opcao == "set") {
    # Tratamento especial para dataframes
    if (is.data.frame(valor)) {
      # Converte o dataframe para a sua representação visual, pois será exibido
      # na mensagem de erro.
      # Por não ser necessário recuperar o dataframe a partir do config.json
      # não é preciso convertê-lo para lista
      valor <- capture.output(print.data.frame(valor, 
                                               right = FALSE, 
                                               row.names = FALSE))
    }
    
    # Verifica se a chave existe e se devemos fazer append
    if (append && chave %in% names(R_config_secao)) {
      # Caso append=TRUE e a chave exista, adiciona ao conteúdo existente
      if (is.list(R_config_secao[[chave]]) && is.list(valor)) {
        # Se ambos são listas, combina recursivamente
        R_config_secao[[chave]] <- modifyList(R_config_secao[[chave]], valor)
      } else if (is.vector(R_config_secao[[chave]])) {
        # Se é um vetor, concatena os valores
        R_config_secao[[chave]] <- c(R_config_secao[[chave]], valor)
      } else {
        # Para outros tipos, cria uma lista com ambos valores
        if (!identical(class(R_config_secao[[chave]]), class(valor))) {
          warning(sprintf("Tipos diferentes ao fazer append: %s e %s", 
                          class(R_config_secao[[chave]]), class(valor)))
        }
        R_config_secao[[chave]] <- list(R_config_secao[[chave]], valor)
      }
    } else {
      # Caso append=FALSE ou chave não exista, substitui o valor
      R_config_secao[[chave]] <- valor
    }
    modificado <- TRUE
  }
  
  # Salva se houve modificação
  if (modificado) {
    # Lê configuração completa atualizada do disco
    config_completo <- .ler_json()
    
    # Atualiza a seção específica
    config_completo[[secao]] <- R_config_secao
    
    # Salva no disco
    .escrever_json(config_completo)
    
    # Atualiza o cache
    config_cache <- list()
    config_cache[[cache_key]] <- R_config_secao
    set_config(.config = modifyList(get_config(), config_cache))
  }
  
  invisible()
  
}

config_ambiente <- function (pasta = NULL) {
  #' Define variáveis de ambiente iniciais
  #'
  #' @description
  #' Configura as variáveis de ambiente necessárias para a execução do script.
  #' Define opções do R, URLs de serviços externos, estrutura de pastas e 
  #' status inicial do script.
  #' Esta função é chamada internamente por \code{config_inicializar()}.
  #'
  #' @param pasta Lista opcional. Se fornecido, deve conter o caminho atual em 
  #' \code{pasta$atual}.
  #' Se NULL (padrão), a pasta atual será obtida com \code{getwd()}.
  #' A função criará automaticamente os caminhos para as pastas \code{_fontes}, 
  #' pasta superior e \code{_common}.
  #'
  #' @details
  #' A função configura:
  #' \itemize{
  #'   \item{Opções do R para exibição de logs (width e max.print)}
  #'   \item{Estrutura de pastas do projeto 
  #'   (atual, fontes, superior, common, log)}
  #'   \item{Informações gerais 
  #'   (nome do script, hora início, tamanho mensagens)}
  #'   \item{URLs de serviços externos do arquivo \code{.Renviron}:
  #'     \itemize{
  #'       \item{url_unidades: Planilha de unidades requerentes}
  #'       \item{url_catalogo: Planilha de grupos do Catálogo de Materiais}
  #'       \item{url_controle: Planilha de Controle de Processos DCOM}
  #'       \item{url_saa: Planilha de Processos SAA/DCOM}
  #'       \item{url_api_catmat: API dos Dados Abertos do Compras.gov.br}
  #'     }
  #'   }
  #'   \item{Status inicial do script (início, erros, alerta)}
  #' }
  #'
  #' @returns Invisível. 
  #' As configurações são armazenadas via \code{set_config()} nas variáveis:
  #' \itemize{
  #'   \item{geral: Informações gerais do script}
  #'   \item{url: URLs dos serviços externos}
  #'   \item{status: Status de execução}
  #'   \item{pasta: Estrutura de pastas}
  #' }
  #'
  #' @examples
  #' # Uso básico - usa pasta atual
  #' config_ambiente()
  #'
  #' # Fornecendo pasta específica
  #' pasta <- list(atual = getwd())
  #' config_ambiente(pasta)
  #'
  #' @seealso \code{\link{config_inicializar}}, \code{\link{set_config}}  

  options(width = 10000)                    #tamanho da linha do log
  options(max.print = .Machine$integer.max) #máximo de linhas registradas no log
  
  if (is.null(pasta)) {
    pasta       <- list()
    pasta$atual <- getwd()
  }
  pasta$superior <- dirname(pasta$atual)
  pasta$common   <- file.path(pasta$superior, "_common")
  pasta$log      <- file.path(pasta$common, "log")

  geral <- 
    list(
      # busca o nome do script pelo nome da função chamadora
      script_nome         = toupper(
        strsplit(as.character(sys.call(1)), "_")[[1]][1]),
      # registra a hora de início do script
      tempo_inicio_script = Sys.time(),
      # define o tamanho do box dos logs/infos/erros
      tamanho_mensagens   = 100,
      config_centralizado = file.path(pasta$common, "config.json")
    )
  
  # Lê arquivo .Renviron na pasta _common para urls salvas em arquivo separado
  readRenviron(file.path(pasta$common, ".Renviron"))
  
  url <- 
    list(
      # planilha de unidades requerentes
      unidades   = Sys.getenv("url_unidades"),
      # planilha de grupos do Catálogo de Materiais
      catalogo   = Sys.getenv("url_catalogo"),
      # planilha de Controle de Processos DCOM
      controle   = Sys.getenv("url_controle"),
      # planilha de Processos SAA/DCOM
      saa        = Sys.getenv("url_saa"),
      # API dos Dados Abertos do Compras.gov.br para consulta aos Catmat e NCM
      api_catmat = Sys.getenv("url_api_catmat"),
      # Lista de CATMATs do Google Drive
      lista_catmat = Sys.getenv("url_lista_catmat")
    )
  
  status <- 
    list(
      inicio = TRUE,
      erros  = FALSE,
      alerta = FALSE
    )
  
  set_config(geral  = geral,
             url    = url,
             status = status,
             pasta  = pasta)
  
}

config_finalizar <- function(sucesso = FALSE) {
  #' Finaliza a execução do script
  #'
  #' @description
  #' Encerra a execução do script, exibindo o relatório de erros/alertas 
  #' se houver, fechando arquivos de log e definindo o 
  #' código de saída apropriado.
  #'
  #' @param sucesso Logical. Se TRUE, indica que o script concluiu sua tarefa
  #' principal mesmo com eventuais erros ou alertas. Padrão é FALSE.
  #'
  #' @details
  #' A função:
  #' \itemize{
  #'   \item{Verifica se houve erros ou alertas durante a execução}
  #'   \item{Exibe relatório detalhado em caso de erros/alertas}
  #'   \item{Fecha arquivos de log}
  #'   \item{Define o código de saída do script}
  #'   \item{Registra o resultado final no arquivo de configurações}
  #' }
  #'
  #' Os códigos de saída são:
  #' \itemize{
  #'   \item{0: Sucesso sem erros}
  #'   \item{1: Erro fatal}
  #'   \item{2: Sucesso com erros/alertas}
  #' }
  #'
  #' @returns A função encerra o script com o código de saída apropriado se não
  #' estiver rodando em um ambiente interativo como o RStudio.
  #'
  #' @seealso \code{\link{config_inicializar}}, \code{\link{log_erro}}

  config            <- get_config()
  logR              <- config$logR
  tamanho_mensagens <- config$geral$tamanho_mensagens
  status            <- config$status

  if (status$erros || status$alerta) {
    msg_aviso <- if (status$erros) "ERROS  " else "ALERTAS"
    cor_aviso <- if (status$erros) "encerrar_erro" else "encerrar_alerta"
    utils_color(cor_aviso)
  
    log_secao(sprintf("ENCERRANDO SCRIPT !!! COM %s", msg_aviso), "CONFIG")
    
    cat(strrep("▄", tamanho_mensagens),
        paste0("█▓▒", 
               strrep("░", (tamanho_mensagens - 26) / 2), 
               sprintf("RELATÓRIO DE %s", msg_aviso), 
               strrep("░", (tamanho_mensagens - 26) / 2), 
               "▒▓█"),
        paste0("█", strrep("▀", tamanho_mensagens - 2), "█"), sep = "\n")
    cat(unlist(as.vector(config_json("msg_erro", opcao = "get"))), 
        labels = "█ ", 
        fill = 1)
    cat("█", strrep("▄", tamanho_mensagens - 2), "█\n", sep = "")

  } else {
    
    utils_color("encerrar_ok")

    log_secao("ENCERRANDO SCRIPT SEM ERROS", "CONFIG")
    
  }
  
  if (sucesso && (status$erros || status$alerta)) {
    resultado_geracao <- "ambos"
    codigo_saida <- 2
  } else if (status$erros) {
    resultado_geracao <- "erro"
    codigo_saida <- 1
  } else {
    resultado_geracao <- "sucesso"
    codigo_saida <- 0
  }

  config_json("resultado_geracao", resultado_geracao)
  
  sink(type="message")
  sink()
  close(logR$con)

  log_secao("SCRIPT FINALIZADO", "CONFIG")
  
  if (!interactive() || Sys.getenv("RSTUDIO") != "1") {
    quit(status = codigo_saida)
  }
  
}

config_inicializar <- function(pacotes, pasta = NULL) {
  #' Inicializa o ambiente de execução do script
  #'
  #' @description
  #' Função principal chamada no início de cada script específico para preparar
  #' o ambiente de execução. Configura variáveis, cria pastas, inicia logs
  #' e carrega pacotes necessários.
  #'
  #' @param pasta Lista. Estrutura de pastas do projeto. Se não fornecido,
  #' será registrada a variável \code{pasta$atual} com a pasta raiz, além das 
  #' pastas para registro dos logs e das fontes 
  #' (arquivos necessários ao script).
  #' Se fornecido, deve ser uma lista, incluindo a \code{pasta$atual} e, 
  #' opcionalmente um vector chamado \code{pasta$criar}, com as pastas que
  #' devem ser criadas, caso não existam. 
  #' 
  #' @param pacotes Character vector. Lista de pacotes R necessários para
  #' a execução do script.
  #'
  #' @details
  #' A função executa as seguintes etapas:
  #' \itemize{
  #'   \item{Define variáveis de ambiente via \code{config_ambiente()}}
  #'   \item{Cria pastas necessárias via \code{config_pastas()}}
  #'   \item{Inicia gravação do log}
  #'   \item{Exibe informações do sistema e ambiente}
  #'   \item{Instala e carrega pacotes necessários via \code{config_pacotes()}}
  #' }
  #'
  #' @examples
  #' # Exemplo de como inicializar um script
  #' pasta <- list(atual = getwd())
  #' pasta$criar <- c(file.path(pasta$atual, "DOCS"))
  #' config_inicializar(
  #'   pasta = pasta,
  #'   pacotes = c("dplyr", "ggplot2")
  #' )
  #'
  #' @seealso \code{\link{config_ambiente}}, \code{\link{config_pastas}},
  #' \code{\link{config_pacotes}}

  # Define a cor do console
  utils_color("inicial")
  
  # Define variáveis de ambiente
  config_ambiente(pasta)
  
  # Verifica pastas de dependência obrigatoria
  pasta <- get_config("pasta")
  config_pasta(pasta$log)
  
  # Inicia gravação do log
  log_gravacao()
  
  # Verifica se há pastas a serem criadas
  if (!is.null(pasta$criar)) {
    config_pasta(pasta$criar)
    
     # Remove o vetor pasta$criar para não ficar registrado
     # nas configurações, pois não é necessário
    pasta$criar <- NULL
    set_config(pasta = pasta)
  }
  
  # Elimina eventual mensagens de erro anteriores
  config_json("msg_erro", opcao = "remove")
  config_json("config_novo", opcao = "remove")
  # Registra que o script está sendo executado, para eventual controle
  config_json("resultado_geracao", "running")
  
  if (utils_silent()) {
    log_info("Modo silencioso ativado. Script executado em segundo plano.")
  }

  log_secao(sprintf("INÍCIO SCRIPT ‹ %s ›", get_config("geral")$script_nome), 
            "CONFIG")
  
  log_info("Informações do sistema", Sys.info(),
           "-",
           "Informações da rede", config_rede(),
           "-",
           "Sessão R", R.version,
           cores = "verde")

  log_secao("CONFIGURAÇÕES INICIAIS", "CONFIG")
  
  log_info("PASTAS UTILIZADAS", pasta,
           "-",
           "LINHA DE COMANDO", commandArgs(),
           cores = "verde")
  
  # Carrega (ou instala) os pacotes necessários
  config_pacotes(pacotes)
  
}

config_pacotes <- function(pacotes) {
  #' Gerencia instalação e carregamento de pacotes
  #'
  #' @description
  #' Verifica se os pacotes necessários estão instalados, instala os que faltam
  #' e carrega todos para uso no script.
  #'
  #' @param pacotes Character vector. Lista de pacotes R necessários.
  #'
  #' @details
  #' A função:
  #' \itemize{
  #'   \item{Verifica se a pasta de instalação existe}
  #'   \item{Identifica quais pacotes precisam ser instalados}
  #'   \item{Instala os pacotes faltantes com suas dependências}
  #'   \item{Carrega todos os pacotes necessários}
  #' }
  #'
  #' A instalação é feita na pasta definida por 
  #' \code{Sys.getenv("R_LIBS_USER")}.
  #'
  #' @returns 
  #' \itemize{
  #'   \item{TRUE se todos os pacotes foram instalados e carregados com sucesso}
  #'   \item{FALSE se houve erro na instalação}
  #'   \item{Erro fatal se não for possível carregar os pacotes}
  #' }
  #'
  #' @examples
  #' config_pacotes(c("dplyr", "ggplot2", "jsonlite"))
  #'
  #' @seealso \code{\link{config_inicializar}}

  log_secao("PACOTES", "CONFIG")
  
  log_info(estilo = "inicio",
           "PACOTES SOLICITADOS", pacotes,
           "-")
  
  pasta_pacotes          <- Sys.getenv("R_LIBS_USER")
  pacotes_instalados     <- pacotes %in% 
    rownames(installed.packages(lib.loc = pasta_pacotes))
  pacotes_nao_instalados <- pacotes[!pacotes_instalados]
  
  if (any(pacotes_instalados == FALSE)) {
    log_info(estilo = "meio", 
             sprintf("Iniciando instalação dos pacotes %s em %s", 
                     pacotes_nao_instalados, pasta_pacotes))
    
    pb <- log_barra_progresso("Aguarde...", length(pacotes_nao_instalados))

    for (i in 1:length(pacotes_nao_instalados)) {
      log_barra_progresso(sprintf("Instalando pacote %s", 
                                  pacotes_nao_instalados[i]), pb = pb)

      install.packages(pacotes_nao_instalados[i], 
                       lib = pasta_pacotes, 
                       dependencies = TRUE, 
                       repos = "https://cloud.r-project.org/")
    }
    
    log_barra_progresso(pb = pb)

    pacotes_instalados     <- pacotes %in% 
      rownames(installed.packages(lib.loc = pasta_pacotes))
    pacotes_nao_instalados <- pacotes[!pacotes_instalados]
    
    if (any(pacotes_instalados == FALSE)) {
      log_erro(sprintf("Erro ao instalar os pacotes %s em %s", 
                       pacotes_nao_instalados, pasta_pacotes))
      return(FALSE)
    } else {
      log_info(estilo = "meio", 
               sprintf("Pacotes necessários instalados em %s", pasta_pacotes))
    }
  }

  pb <- log_barra_progresso("Aguarde...", length(pacotes))

  tryCatch({
    for (i in 1:length(pacotes)) {
      log_barra_progresso(sprintf("Carregando pacote %s", pacotes[i]), pb = pb)
      
      library(pacotes[i], 
              lib.loc = pasta_pacotes, 
              verbose = TRUE, 
              character.only = TRUE)
    }
    log_barra_progresso(pb = pb)
    log_info(paste0("Pacotes necessários carregados em ", pasta_pacotes),
             estilo = "meio")
  },
  error = function(e) {
    log_erro(paste("Não foi possível carregar os pacotes necessários em", 
                   pasta_pacotes), 
             e,
             finalizar = TRUE)
  })
  
  log_info(estilo = "fim")
}

config_pasta <- function(...) {
  #' Cria diretórios de forma segura
  #'
  #' @description
  #' Função utilitária para criar um ou mais diretórios de forma segura,
  #' com tratamento de erros e suporte a criação recursiva de pastas.
  #' É utilizada internamente por várias funções do sistema para garantir
  #' que as estruturas de diretórios necessárias existam.
  #'
  #' @param ... Caminhos dos diretórios a serem criados. Podem ser fornecidos:
  #' \itemize{
  #'   \item{Um ou mais caminhos como argumentos separados}
  #'   \item{Um vetor ou lista de caminhos}
  #'   \item{Uma combinação dos anteriores}
  #' }
  #' Caminhos NULL ou vazios ("") são ignorados silenciosamente.
  #'
  #' @details
  #' Para cada caminho fornecido, a função:
  #' \itemize{
  #'   \item{Verifica se o diretório já existe}
  #'   \item{Se não existir, cria o diretório e todos os subdiretórios 
  #'   necessários (equivalente a mkdir -p)}
  #'   \item{Exibe mensagem informando a criação}
  #'   \item{Em caso de erro, lança uma exceção com mensagem detalhada}
  #' }
  #'
  #' @returns Invisível. Como efeitos colaterais:
  #' \itemize{
  #'   \item{Cria os diretórios solicitados se não existirem}
  #'   \item{Emite warning se nenhum caminho for fornecido}
  #'   \item{Em caso de erro na criação, interrompe a execução com erro}
  #' }
  #'
  #' @examples
  #' # Criar um único diretório
  #' config_pasta("./dados")
  #'
  #' # Criar múltiplos diretórios
  #' config_pasta("./dados/brutos", "./dados/processados")
  #'
  #' # Usar com vetor de caminhos
  #' pastas <- c("./log", "./temp")
  #' config_pasta(pastas)
  #'
  #' @seealso 
  #' Usado por:
  #' \code{\link{config_main}} para criar pasta de pacotes
  #' \code{\link{config_ambiente}} para criar estrutura do projeto
  #'

  pastas <- unlist(list(...))

  if (length(pastas) == 0) {
    warning("Nenhuma pasta fornecida para config_pasta()")
    return(invisible())
  }
  
  for (pasta in pastas) {
    # Ignora valores NULL ou vazios
    if (is.null(pasta) || pasta == "") next
    
    tryCatch({
      if (!dir.exists(pasta)) {
        cat(sprintf("=== Criando pasta %s ===\n", pasta))
        dir.create(pasta, recursive = TRUE)
      }
    },
    error = function(e) {
      stop(sprintf("Não foi possível criar a pasta %s:\n%s\n", 
                   pasta, 
                   e$message))
    })
  }
  
}

config_rede <- function() {
  #' Identifica conexões de rede ativas
  #'
  #' @description
  #' Verifica as conexões de rede ativas no sistema executando o comando
  #' \code{ipconfig} e analisando seu resultado.
  #'
  #' @details
  #' A função:
  #' \itemize{
  #'   \item{Executa o comando ipconfig}
  #'   \item{Analisa as linhas contendo informações de DNS}
  #'   \item{Identifica conexões ativas pelos sufixos DNS}
  #'   \item{Remove conexões duplicadas}
  #' }
  #'
  #' @returns Character vector contendo os sufixos DNS das conexões ativas.
  #' Em caso de erro ou se nenhuma conexão for encontrada, retorna uma
  #' mensagem descritiva.
  #'
  #' @examples
  #' # Obtém lista de conexões ativas
  #' redes <- config_rede()
  #' print(redes)
  #'
  #' @seealso \code{\link{config_inicializar}}
  
  tryCatch({
    rede_info <- system("ipconfig", intern = TRUE, timeout = 5)
    
    linhas_dns <- grep("DNS", rede_info, useBytes = TRUE)
    redes      <- c()
    
    for (i in linhas_dns) {
      linha <- rede_info[i]
      
      sufixo <- trimws(unlist(strsplit(rede_info[i], ":", useBytes = TRUE))[2])

      if (sufixo != "" && !grepl("desconectada|disconnected", 
                                 rede_info[i - 1], 
                                 useBytes = TRUE)) redes <- c(redes, sufixo)
    }
    
    redes <- unique(redes)
    
    if (length(redes) == 0) redes <- "NENHUMA CONEXÃO LOCALIZADA"
    
    return(redes)
  },
  error = function(e) {
    return("Erro: NÃO FOI POSSÍVEL TESTAR A REDE")
  },
  warning = function(w) {
    return("Aviso: NÃO FOI POSSÍVEL TESTAR A REDE")
  })
  
}

# ---- INICIAR ----

# Executa a função config_main() ao carregar este arquivo
# para iniciar dependências críticas (json, logging e utils)
config_main()