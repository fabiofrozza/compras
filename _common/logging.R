# ---- LOGGING ----
# Funções para feedback visual de informações gerais e erros/alertas
# e registro de log

log_barra_progresso <- function(label = NULL, steps = NULL, pb = NULL) {
  #' Exibe um feedback visual de progresso
  #' 
  #' @description Exibe ao usuário o progresso da execução da atividade atual.
  #' 
  #' Utilize quando a atividade pode levar bastante tempo, para evitar que o 
  #' usuário confunda com travamento do script.
  #' 
  #' Quando o script é executado dentro de uma janela ou em segundo plano, a
  #' barra de progresso é criada no modo texto (\code{txtProgressBar}).
  #' Se não, é exibida a barra de progresso padrão do Windows 
  #' (\code{winProgressBar}).
  #' 
  #' @details
  #' \itemize{
  #' \item{O modo texto é escolhido caso o script seja chamado com o argumento
  #' \code{silent}.}
  #' \item{O tamanho da barra no modo texto é definida na função
  #' \code{config_ambiente()}.}
  #' \item{Para o título exibido na janela (se for o caso), é usado o nome do
  #' script definido na função \code{config_opcoes()}.}
  #' }
  #' 
  #' @note
  #' Embora seja utilizada para exibição do status de progresso da atividade,
  #' se o script estiver no modo \code{silent}, a barra de progresso será 
  #' registrada no log (pois é executada no modo texto).
  #' Isto, contudo, não interfere nem na execução do script 
  #' nem no registro do log.
  #' 
  #' @param label Texto. Informação a ser exibida ao usuário. 
  #' Argumento posicional (primeiro) e não é necessário nomeá-lo. 
  #' Utilize para criar a barra e atualizar o status.
  #' \strong{Não} use este parâmetro ao fechar a barra.
  #' 
  #' @param steps Numérico. Quantidade de passos que a atividade terá. 
  #' Argumento posicional (segundo) e não é necessário nomeá-lo. 
  #' Utilize \strong{apenas} para criar a barra.
  #' 
  #' @param pb Objeto. Nome da barra de progresso, definido ao chamar a função
  #' (veja Exemplos). 
  #' \strong{Sempre} nomeie este argumento, para evitar que seja
  #' interpretado como a quantidade de passos. 
  #' Utilize \strong{apenas} para atualizar e fechar
  #' 
  #' @usage 
  #' log_barra_progresso(label = NULL, steps = NULL, pb = NULL)
  #' 
  #' @examples
  #' #Inicia a barra de progresso de uma atividade com 10 passos
  #' #Veja que foi ela foi nomeada como "pb", e deve ser assim 
  #' #referenciada nas atualizações e fechamento
  #' pb <- log_barra_progresso("Aguarde...", 10)  
  #' 
  #' #Atualiza o status para uma nova atividade
  #' log_barra_progresso("Nova atividade...", pb = pb)
  #' 
  #' #Finaliza a barra de progresso e fecha a janela (se existente)
  #' log_barra_progresso(pb = pb)
  #' 
  #' @return Será exibida uma caixa de diálogo do Windows com a 
  #' barra de progresso da execução da tarefa (a não ser que o script esteja 
  #' sendo executado em segundo plano ou dentro de outra janela).
  #' 
  #' @seealso \code{\link{config_ambiente}}, \code{\link{config_opcoes}},
  #' \code{\link{config_inicializar}}
  
  if (utils_silent()) {
    if (is.null(pb)) { #se não informa pb (progress_bar) existente, cria
      tamanho <- get_config("geral")$tamanho_mensagens
      
      pb <- txtProgressBar(
        label = label, 
        min = 0, 
        max = steps,
        width = tamanho,
        char = "■")
      return(pb)
    }
    if (!is.null(label)) { #se informa label, atualiza
      valor_atual <- getTxtProgressBar(pb)
      setTxtProgressBar(
        pb, 
        value = valor_atual + 1, 
        label = label)
    } else { #fecha barra se não fornecido label
      close(pb)
    }
    
  } else {
    
    if (is.null(pb)) { #se não informa pb (progress_bar) existente, cria
      
      titulo <- get_config("geral")$script_nome
      
      pb <- winProgressBar(
        title = titulo, 
        label = label, 
        min = 0, 
        max = steps, 
        width = 500)
      
      return(pb)
    }
    
    if (!is.null(label)) { #se informa label, atualiza
      valor_atual <- getWinProgressBar(pb)
      
      setWinProgressBar(
        pb, 
        value = valor_atual + 1, 
        label = label)
      
    } else { #fecha barra se não fornecido label
      
      close(pb)
      
    }
  }
  
  invisible(NULL)
}

log_erro <- function(msg_erro = NULL, dados = NULL, 
                     titulo = "ERRO", alerta = FALSE, finalizar = FALSE) {
  #' Exibe e registra erros ou alertas
  #' 
  #' @description 
  #' Exibe uma mensagem formatada de erro ou alerta no console e a registra
  #' no log. Pode ser usada para finalizar o script em caso de um erro crítico.
  #' Se msg_erro for omitido, vazio ou nulo, será usada uma mensagem genérica.
  #' 
  #' @details
  #' \itemize{
  #' \item{O tamanho da caixa contendo as informações é definido na função
  #' \code{config_ambiente()}.}
  #' \item{Na primeira chamada da função, a variável de ambiente
  #' \code{status$inicio} é alterada de TRUE para FALSE e o nome do arquivo
  #' de log (\code{logR$nome}) é registrado no log para posterior controle.}
  #' \item{Também é registrado nas variáveis de ambiente se ocorrerem erros
  #' (\code{status$erros = TRUE}) ou alertas (\code{status$alerta = TRUE}).}
  #' \item{Se o script não estiver no modo \code{silent}, a cor da mensagem
  #' será vermelha (erros) ou laranja (alertas).}
  #' \item{As mensagens exibidas são registradas no arquivo de configurações
  #' por meio da função \code{config_json()}, que altera a variável
  #' \code{msg_erro}, criando-a ou adicionando ao seu conteúdo já existente.}
  #' }
  #' 
  #' @param msg_erro Obrigatória. Informações gerais sobre a ocorrência.
  #' Se não informada, será exibida uma mensagem genérica.
  #'
  #' @param dados Opcional. Informações extras com os detalhes do erro/alerta,
  #' a fim de auxiliar a identificação da sua causa
  #' Pode ser fornecido o texto diretamente ou os objetos (character, numeric,
  #' vector, list, dataframe ou error/warning capturados com tryCatch).
  #' Caso o objeto contenha mais de uma informação, cada uma será exibida em uma
  #' linha com uma seta no início.
  #' Se o objeto contiver a classe \code{error} ou \code{warning}, serão
  #' exibidas a mensagem de erro do sistema e a 
  #' linha de código causadora do erro.
  #'
  #' @param titulo Opcional. Se não informado, será exibido como "ERRO"
  #' (a não ser que alerta = TRUE).
  #' Sempre nomeie este argumento.
  #'
  #' @param alerta Opcional. Por padrão é \code{FALSE}, ou seja, a mensagem
  #' será tratada como um erro. Se for \code{TRUE}, será tratada com um alerta,
  #' ou seja, não será registrada a ocorrência de erro no arquivo de 
  #' configurações.
  #' Utilize para situações que não impedem a execução do script ou a geração
  #' dos arquivos desejados, mas que merecem atenção.
  #' Sempre nomeie este argumento.
  #' 
  #' @param finalizar Opcional. Se \code{TRUE}, o script será encerrado. 
  #' Utilize para erros que impedem a continuidade da execução do script.
  #' Sempre nomeie este argumento.
  #' 
  #' @usage log_erro(msg_erro = NULL, dados = NULL, 
  #' titulo = "ERRO", alerta = FALSE, finalizar = FALSE)
  #' 
  #' @examples
  #' #Mensagem de erro simples
  #' log_erro("Houve um erro e a atividade não foi executada")
  #' 
  #' #Mensagem de alerta para atenção
  #' log_erro("Tudo ocorreu como esperado, mas confira isto.",
  #'          arquivos_a_serem_verificados,
  #'          alerta = TRUE)
  #' 
  #' #Mensagem de erro com título escolhido pelo usuário
  #' log_erro("ERRO CATASTRÓFICO",
  #'          titulo = "A T E N Ç Ã O",
  #'          finalizar = TRUE)
  #' 
  #' #Mensagem de erro grave, capturada num bloco try/catch
  #' tryCatch({
  #'   #função para salvar o arquivo
  #' },
  #' error = function(e) {
  #'   log_erro("Não foi possível salvar o arquivo. Encerrando...", 
  #'            e,
  #'            finalizar = TRUE)
  #' })
  #'  
  #' @return Será exibida uma caixa destacada com as informações fornecidas, 
  #' além do carimbo de tempo, conforme a seguir:
  #' \preformatted{
  #' ▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄
  #' █▓▒░░░░░░ ERRO ░░░░░░░░▒▓█
  #' █ Erros identificados:   █
  #' █  x, y e z
  #' █▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄█ ‹17:42 | 7.9 secs›}
  #' As mensagens também serão registradas no arquivo de log.
  #' Caso utilizado o argumento \code{finalizar = TRUE}, chama a função
  #' \code{config_finalizar()} para encerrar o script. 
  #' 
  #' @seealso \code{\link{config_ambiente}}, \code{\link{get_config}}, 
  #' \code{\link{config_json}}, \code{\link{log_tempo_decorrido}}
  
  # Define mensagem padrão se msg_erro for omitido, vazio ou nulo
  if (is.null(msg_erro) || nchar(trimws(msg_erro)) == 0) {
    msg_erro <- "OCORREU UM ERRO E A CAUSA NÃO FOI INFORMADA (msg_erro VAZIA)"
  }
  
  tamanho_mensagens <- get_config("geral")$tamanho_mensagens
  tamanho_erro      <- max(tamanho_mensagens, nchar(msg_erro) + 4)
  
  if (alerta & titulo == "ERRO") titulo <- "ALERTA"

  status <- get_config("status")
  logR   <- get_config("logR")
  
  if (status$inicio) {
    config_json("msg_erro", logR$nome)
    status$inicio <- FALSE
  } 
  
  if (alerta) {
    status$alerta <- TRUE
  } else {
    status$erros <- TRUE
  }
  
  set_config("status" = status)
  
  titulo <- paste0(" ", toupper(trimws(titulo)), " ")
  
  espacamento <- (tamanho_erro - nchar(titulo) - 6) / 2
  if (espacamento %% 1 != 0) {
    titulo <- paste0(titulo, "░")
  }
  
  if (!utils_silent()) { 
    if (alerta) utils_color("cinza_fundo") else utils_color("vermelho") 
  }
  
  cat("\n", strrep("▄", tamanho_erro),
      "\n█▓▒", strrep("░", espacamento), 
      titulo, 
      strrep("░", espacamento), "▒▓█", "\n█ ", 
      msg_erro, 
      strrep(" ", tamanho_erro - nchar(msg_erro) - 3), "█\n", sep = "")
  
  config_json("msg_erro", msg_erro, append = TRUE)
  
  if (!(missing(dados))) {
    if ("error" %in% class(dados) | "warning" %in% class(dados)) {
      cat(sprintf("█ → Código: %s\n", deparse(dados$call)))
      cat(sprintf("█ → Erro  : %s\n", dados$message))
      
      config_json("msg_erro", sprintf("  → Código: %s", deparse(dados$call)), 
                  append = TRUE)
      config_json("msg_erro", sprintf("  → Erro  : %s", dados$message), 
                  append = TRUE)
      
    } else if (is.data.frame(dados)) {
      print.data.frame(dados, right = FALSE, row.names = FALSE)
      
      config_json("msg_erro", dados, append = TRUE)
      
    } else {
      if (is.list(dados)) dados <- as.vector(unlist(dados))
      cat(dados, labels = "█ ", fill = 1)
      
      config_json("msg_erro", dados, append = TRUE)
    }
  }
  cat("█", strrep("▄", tamanho_erro - 2), "█", 
      log_tempo_decorrido(), "\n", sep = "")
  
  utils_color("inicial")
  
  if (finalizar) config_finalizar()
  
}

log_gravacao <- function() {
  #' Inicia a gravação do log em arquivo 
  #'
  #' @description
  #' Gera um nome de arquivo de log único, o registra nas variáveis de ambiente
  #' e no arquivo de configurações, e inicia a captura de saídas do console
  #' para esse arquivo.
  #'  
  #' @details
  #' \itemize{
  #' \item{A função gera um nome do arquivo de log único utilizando data e 
  #' horário e dados do usuário e computador.}
  #' \item{A pasta onde será gravado o log é a pasta definida em 
  #' \code{config_ambiente()}, na variável \code{pasta$log} 
  #' obtida com \code{get_config()}.}
  #' \item{O nome do arquivo é registrado nas variáveis de ambiente
  #' (\code{logR$nome}) com a função \code{set_config()} 
  #' para posterior utilização.}
  #' \item{O nome do arquivo também é registrado no arquivo de configurações
  #' (\code{config.json}) com a função \code{config_json()}.}
  #' \item{A gravação do arquivo é iniciada e registrada na variável de ambiente
  #' \code{logR$con}, para ser utilizada posteriormente.}
  #' }
  #' 
  #' @param - Nenhum argumento necessário
  #' 
  #' @returns
  #' Inicia a gravação do log em um arquivo na pasta definida.
  #' 
  #' @usage log_gravacao()
  #' @examples log_gravacao()
  #' @seealso \code{\link{set_config}}, \code{\link{config_json}}
  
  pasta       <- get_config("pasta")
  script_nome <- get_config("geral")$script_nome
  
  tryCatch({
    logR      <- list()
    logR$nome <- sprintf("Log_%s_%s_%s_%s-%s_R.log",
                         toupper(strsplit(script_nome, " ")[[1]][1]),
                         format(as.POSIXct(Sys.time()), format = "%Y-%m-%d"),
                         format(as.POSIXct(Sys.time()), format = "%H-%M-%S"),
                         toupper(Sys.getenv("USERNAME")),
                         toupper(Sys.info()["nodename"]))
    logR$nome <- gsub("[[:blank:]]", "", logR$nome)
    logR$con  <- file(file.path(pasta$log, logR$nome), 
                      open = "wt", 
                      encoding = "UTF-8")
    
    config_json("arquivo_log_R", logR$nome)
    
    if (!"RStudio" %in% commandArgs()) {
      sink(logR$con, append = TRUE, split = TRUE)
      sink(logR$con, append = TRUE, type = "message")
    }
    
    set_config(logR = logR)
    
    log_info("Log iniciado em", 
             logR$nome, 
             cores = "vermelho")
  },
  error = function(e) {
    log_erro("Não foi possível iniciar a gravação do log. Encerrando...", 
             e,
             finalizar = TRUE)
  })
}

log_info <- function(..., estilo = "completo", cores = NULL) {
  #' Exibe uma caixa de informações formatada
  #' 
  #' @description No feedback visual ao usuário ou no registro do log,
  #' exibe informações úteis para verificar a função do script, 
  #' a atividade sendo executada e o seu sucesso ou eventuais erros.
  #' 
  #' @details
  #' \itemize{
  #' \item{O tamanho da caixa contendo as informações é definido na função
  #' \code{config_ambiente()}.}
  #' }
  #' 
  #' @param ... Informações a serem exibidas, cada uma em uma linha.
  #' Pode ser fornecido o texto diretamente ou os objetos (character, numeric, 
  #' vector, list ou dataframe).
  #' Caso o objeto contenha mais de uma informação, cada uma será exibida em uma
  #' linha com uma seta no início.
  #' Utilize um hífen ("-") para adicionar um separador.
  #' Se nada for informado, será apenas desenhada a caixa
  #' (ou a parte solicitada).
  #' 
  #' @param estilo Opcional. Opções disponíveis: \code{completo} (padrão),
  #' \code{inicio}, \code{meio} e \code{fim}.
  #' Se não informado ou for informada opção inexistente, será utilizado o 
  #' estilo \code{completo}.
  #' O estilo \code{completo} desenha a caixa inteira ao redor das informações
  #' fornecidas, enquanto os outros estilos desenham apenas a parte correspondente.
  #' Estas outras opções são úteis quando é necessário executar alguma função
  #' quando a caixa de informações já foi desenhada mas ainda não finalizada.
  #' Sempre nomeie este argumento.
  #' 
  #' @param cores Opcional. Opções disponíveis: \code{verde} e \code{vermelho}. 
  #' Se não informado (ou informada outra opção), será utilizada a cor padrão
  #' do console.
  #' Sempre nomeie este argumento.
  #' Se o script estiver sendo executado no modo \code{silent}, este argumento
  #' será ignorado.
  #' 
  #' @usage log_info(..., estilo = "completo", cores = NULL)
  #' 
  #' @examples
  #' #Box básico, informando um texto e um objeto
  #' log_info("Nome do script:", nome_do_script)
  #' 
  #' #Box com separador e cor diferenciada
  #' log_info("Pacotes necessários", 
  #'          c("openxlsx", "dplyr"),
  #'          cores = "verde")
  #'          
  #' #Iniciar um box com algumas informações e um separador, 
  #' #executar uma função e fechá-lo
  #' pastas <- c("main", "common")
  #' log_info("Serão criadas as seguintes pastas:", 
  #'          pastas, 
  #'          "-",
  #'          estilo = "inicio")
  #' dir.create(pastas)
  #' log_info(estilo = "fim")
  #'  
  #' @return Será exibida uma caixa com as informações fornecidas, além do 
  #' carimbo de tempo, fornecido pela função \code{log_tempo_decorrido()}, 
  #' conforme a seguir:
  #' 
  #' \preformatted{
  #' ╭───────────────────────────╮
  #' │ PACOTES SOLICITADOS       │
  #' │ → openxlsx                │
  #' ├───────────────────────────┤
  #' │ Pacotes carregados em ... │
  #' ╰───────────────────────────╯ ‹12:12 | 3.9 secs›}
  #' 
  #' @seealso \code{\link{config_ambiente}}, \code{\link{config_opcoes}}, 
  #' \code{\link{log_tempo_decorrido}}
  
  tamanho_mensagens <- get_config("geral")$tamanho_mensagens
  
  estilos_validos <- c("completo", "inicio", "meio", "fim")
  if (!(estilo %in% estilos_validos)) {
    estilo <- "completo"
  }
  
  desenhar_borda_superior <- estilo %in% c("completo", "inicio")
  desenhar_borda_inferior <- estilo %in% c("completo", "fim")
  
  if (desenhar_borda_superior) {
    utils_color(cores)
    cat("╭", strrep("─", tamanho_mensagens - 2), "╮\n", sep = "")
  }
  
  tryCatch({  
    for (i in list(...)) {
      if (is.list(i) & !is.data.frame(i)) i <- as.vector(unlist(i))
      if (is.data.frame(i)) {
        print.data.frame(i, right = FALSE, row.names = FALSE)
      } else if (length(i) != 1) {
        cat(paste0("│ → ", i, strrep(" ", ifelse(
          tamanho_mensagens - nchar(i) - 5 < 0, 0, 
          tamanho_mensagens - nchar(i) - 5)), "│\n"), sep = "")
      } else if (i == "-") {
        cat("├", strrep("─", tamanho_mensagens - 2), "┤\n", sep = "")
      } else {
        cat(paste0("│ ", i, strrep(" ", ifelse(
          tamanho_mensagens - nchar(i) - 3 < 0, 0, 
          tamanho_mensagens - nchar(i) - 3)), "│\n"))
      }
    }
  },
  error = function(e) { 
    log_erro("Alguma informação solicitada está indisponível.", 
             e) 
  })
  
  if (desenhar_borda_inferior) {
    cat("╰", strrep("─", tamanho_mensagens - 2), "╯", 
        log_tempo_decorrido(), "\n", sep = "")
    utils_color()
  }
  
}

log_secao <- function(subtitulo, titulo = NULL) {
  #' Exibe um cabeçalho de seção formatado
  #' 
  #' @description No feedback visual ao usuário ou no registro do log,
  #' é importante informar o que vai ser providenciado pelo código.
  #' 
  #' No caso do log, é essencial para verificar a origem de eventuais erros.
  #' 
  #' @details
  #' \itemize{
  #' \item{Todo o texto é convertido para maiúsculas.}
  #' \item{O tamanho da caixa contendo as informações é definida na função
  #' \code{config_ambiente()}.}
  #' \item{Tanto o título (caso não fornecido) quanto o tamanho da caixa 
  #' estão na variável \code{opcoes}, atribuídos pela função 
  #' \code{config_opcoes()}.}
  #' }
  #' 
  #' @param subtitulo Obrigatório. Descrição da seção que está sendo iniciada
  #' @param titulo Opcional. Título que será exibido no lado esquerdo,
  #' para fins de localização da origem da chamada.
  #' Caso não informado, será utilizado o nome do script definido em 
  #' \code{config_ambiente()}
  #' 
  #' @usage log_secao(subtitulo, titulo = NULL)
  #' 
  #' @examples
  #' log_secao("Lendo arquivos")
  #' log_secao("Configurações iniciais", "LOG")
  #' 
  #' @return Será exibida uma caixa com o título e o subtítulo, além do 
  #' carimbo de tempo, fornecido pela função \code{log_tempo_decorrido()}, 
  #' conforme a seguir:
  #' 
  #' \preformatted{
  #' ╭─────┬────────────────────────╮
  #' │ LOG │ CONFIGURAÇÕES INICIAIS │
  #' ╰─────┴────────────────────────╯ ‹08:08 | 1.1 secs›}
  #' 
  #' @seealso \code{\link{config_ambiente}}, \code{\link{config_opcoes}}, 
  #' \code{\link{log_tempo_decorrido}}
  
  geral <- get_config("geral")
  
  titulo <- if (is.null(titulo)) {
    toupper(trimws(geral$script_nome))
  } else {
    toupper(trimws(titulo)) 
  }
  subtitulo <- toupper(trimws(subtitulo))
  
  tamanho_secao <- max(geral$tamanho_mensagens, 
                       nchar(titulo) + nchar(subtitulo) + 7)
  
  linha_1_1 <- paste0("╭", strrep("─", nchar(titulo) + 2), "┬")
  linha_1_2 <- paste0(strrep("─", tamanho_secao - nchar(linha_1_1) - 1), "╮")
  
  linha_2_1 <- paste0("│ ", titulo, " │")
  espacamento <- (tamanho_secao - nchar(titulo) - 4 - nchar(subtitulo) - 1) / 2
  if (espacamento %% 1 != 0) subtitulo <- paste0(subtitulo, " ")
  espacamento <- strrep(" ", espacamento)
  linha_2_2 <- paste0(espacamento, subtitulo, espacamento, "│")
  
  linha_3_1 <- paste0("╰", strrep("─", nchar(titulo) + 2), "┴")
  linha_3_2 <- paste0(strrep("─", tamanho_secao - nchar(linha_1_1) - 1), "╯", 
                      log_tempo_decorrido())
  
  cat("\n",
      linha_1_1, linha_1_2, "\n",
      linha_2_1, linha_2_2, "\n",
      linha_3_1, linha_3_2, "\n",
      sep = "")
}

log_tempo_decorrido <- function() {
  #' Gera um marcador de tempo decorrido de execução
  #' 
  #' @description Informa o horário atual e o tempo decorrido desde o início 
  #' da execução do script, para eventual verificação de falhas.
  #' 
  #' Esta informação é importante tanto para feedback visual ao usuário
  #' quanto para registro no log.
  #' 
  #' @details O horário de início da execução do script é registrado nas 
  #' variáveis de ambiente pela função \code{config_ambiente()} e é recuperado 
  #' pela função \code{get_config()}.
  #' 
  #' @param - Nenhum parâmetro necessário
  #'
  #' @returns A função retorna no ponto onde é chamada um carimbo de tempo
  #' neste formato:
  #' 
  #' \code{‹12:00 | 5.2 mins›}
  #' 
  #' @usage log_tempo_decorrido()
  #'
  #' @examples log_tempo_decorrido()
  #' 
  #' @seealso \code{\link{config_ambiente}}, \code{\link{config_opcoes}}
  
  agora <- Sys.time()
  decorrido <- format(agora - get_config("geral")$tempo_inicio_script, 
                      digits = 2)
  return(paste0(" ‹", format(agora, "%H:%M"), " ◌ ", decorrido, "›"))
}
