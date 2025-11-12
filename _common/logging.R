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

  versao_texto <- !utils_is_windows() || utils_silent()

  # Criar nova barra
  if (is.null(pb)) {
    if (versao_texto) {
      return(
        txtProgressBar(
          label = label,
          min = 0,
          max = steps,
          width = get_config("skin")$tamanho_mensagens,
          char = "■"
        )
      )
    } else {
      return(
        winProgressBar(
          title = get_config("geral")$script_nome,
          label = label,
          min = 0,
          max = steps,
          width = 500
        )
      )
    }
  }

  # Atualizar barra existente
  if (!is.null(label)) {
    if (versao_texto) {
      setTxtProgressBar(pb, value = getTxtProgressBar(pb) + 1, label = label)
    } else {
      setWinProgressBar(pb, value = getWinProgressBar(pb) + 1, label = label)
    }

    return(invisible(NULL))
  }

  # Fechar barra
  close(pb)
  return(invisible(NULL))
}

log_erro <- function(
  msg_erro = NULL,
  dados = NULL,
  titulo = "ERRO",
  alerta = FALSE,
  finalizar = FALSE
) {
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
  #' \code{\link{config_json}}}

  # Define mensagem padrão se msg_erro for omitido, vazio ou nulo
  if (is.null(msg_erro) || nchar(trimws(msg_erro)) == 0) {
    msg_erro <- "OCORREU UM ERRO E A CAUSA NÃO FOI INFORMADA (msg_erro VAZIA)"
  }

  tamanho_mensagens <- get_config("skin")$tamanho_mensagens
  tamanho_erro <- max(tamanho_mensagens, nchar(msg_erro) + 4)

  if (alerta && titulo == "ERRO") titulo <- "ALERTA"

  status <- get_config("status")
  log_r <- get_config("log_r")

  if (status$inicio) {
    config_json("msg_erro", log_r$nome)
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

  if (alerta) utils_color("alert") else utils_color("error")

  cat("\n", strrep("▄", tamanho_erro),
    "\n█▓▒", strrep("░", espacamento),
    titulo,
    strrep("░", espacamento), "▒▓█", "\n█ ",
    msg_erro,
    strrep(" ", tamanho_erro - nchar(msg_erro) - 3), "█\n",
    sep = ""
  )

  config_json("msg_erro", msg_erro, append = TRUE)

  if (!(missing(dados))) {
    if ("error" %in% class(dados) || "warning" %in% class(dados)) {
      cat(sprintf("█ → Código: %s\n", deparse(dados$call)))
      cat(sprintf("█ → Erro  : %s\n", dados$message))

      config_json("msg_erro", sprintf("  → Código: %s", deparse(dados$call)),
        append = TRUE
      )
      config_json("msg_erro", sprintf("  → Erro  : %s", dados$message),
        append = TRUE
      )
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
    .log_tempo_decorrido(), "\n",
    sep = ""
  )

  utils_color("default")

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

  pasta <- get_config("pasta")
  script_nome <- get_config("geral")$script_nome

  tryCatch(
    {
      log_r <- list()
      log_r$nome <- sprintf(
        "Log_%s_%s_%s_%s-%s_R.log",
        toupper(strsplit(script_nome, " ")[[1]][1]),
        format(as.POSIXct(Sys.time()), format = "%Y-%m-%d"),
        format(as.POSIXct(Sys.time()), format = "%H-%M-%S"),
        toupper(Sys.getenv("USERNAME")),
        toupper(Sys.info()["nodename"])
      )
      log_r$nome <- gsub("[[:blank:]]", "", log_r$nome)
      log_r$con <- file(file.path(pasta$log, log_r$nome),
        open = "wt",
        encoding = "UTF-8"
      )

      config_json("arquivo_log_R", log_r$nome)

      if (!utils_is_interactive()) {
        sink(log_r$con, append = TRUE, split = TRUE)
        sink(log_r$con, append = TRUE, type = "message")
      }

      set_config(log_r = log_r)
    },
    error = function(e) {
      log_erro("Não foi possível iniciar a gravação do log. Encerrando...",
        e,
        finalizar = TRUE
      )
    }
  )

  log_info("Log iniciado em",
    log_r$nome,
    cores = "highlight2"
  )
}

log_info <- function(..., estilo = "completo", cores = NULL, timestamp = TRUE) {
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
  #' fornecidas, enquanto os outros estilos desenham apenas a
  #' parte correspondente.
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
  #'          cores = "highlight1")
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
  #' carimbo de tempo, fornecido pela função \code{.log_tempo_decorrido()},
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
  #' @seealso \code{\link{config_ambiente}}, \code{\link{config_opcoes}}}

  # Validação de estilo
  estilos_validos <- c("completo", "inicio", "meio", "fim")
  if (!(estilo %in% estilos_validos)) estilo <- "completo"

  # Processa conteúdo
  linhas <- tryCatch(
    .log_processar_conteudo(..., estilo = estilo),
    error = function(e) {
      log_erro("Alguma informação solicitada está indisponível.", e)
      list()
    }
  )

  # Configura renderização
  .log_exibir_box(
    tipo = "info",
    cores = cores,
    timestamp = timestamp,
    borda_superior = estilo %in% c("completo", "inicio"),
    borda_inferior = estilo %in% c("completo", "fim"),
    linhas = linhas
  )
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
  #' carimbo de tempo, fornecido pela função \code{.log_tempo_decorrido()},
  #' conforme a seguir:
  #'
  #' \preformatted{
  #' ╭─────┬────────────────────────╮
  #' │ LOG │ CONFIGURAÇÕES INICIAIS │
  #' ╰─────┴────────────────────────╯ ‹08:08 | 1.1 secs›}
  #'
  #' @seealso \code{\link{config_ambiente}}, \code{\link{config_opcoes}}}

  script_nome <- get_config("geral")$script_nome

  .log_exibir_box(
    tipo = "secao",
    titulo = toupper(trimws(ifelse(is.null(titulo), script_nome, titulo))),
    subtitulo = toupper(trimws(subtitulo))
  )
}

# ---- FUNÇÕES AUXILIARES INTERNAS ----

.log_tempo_decorrido <- function() {
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
  #' @usage .log_tempo_decorrido()
  #'
  #' @examples .log_tempo_decorrido()
  #'
  #' @seealso \code{\link{config_ambiente}}, \code{\link{config_opcoes}}

  agora <- Sys.time()
  decorrido <-
    format(agora - get_config("geral")$tempo_inicio_script,
      digits = 2
    )
  paste0(" ‹", format(agora, "%H:%M"), " ◌ ", decorrido, "›")
}

.log_calcular_espacamento <- function(tamanho_total, tamanho_conteudo) {
  espacamento_base <- (tamanho_total - tamanho_conteudo) / 2
  ajuste_direita <- if (espacamento_base %% 1 != 0) 1 else 0

  list(
    esquerda = floor(espacamento_base),
    direita = floor(espacamento_base) + ajuste_direita
  )
}

.log_montar_linha <- function(
  tipo = c("borda", "conteudo", "secao"),
  posicao = c("superior", "separador", "meio", "inferior"),
  conteudo = NULL,
  timestamp = TRUE
) {
  #' Monta uma linha do box
  #'
  #' @param tipo "borda", "conteudo", "secao"
  #' @param posicao "superior", "separador", "meio", "inferior"
  #' @param conteudo Texto ou lista de textos
  #' @param timestamp Se exibe ou não a marcação de tempo decorrido

  skin <- get_config("skin")
  borders <- skin$border
  tamanho <- skin$tamanho_mensagens
  margem <- skin$tamanho_margem

  if (tipo == "borda") {
    if (posicao == "superior") {
      left <- borders$corner_upper_left
      right <- borders$corner_upper_right
      bar <- borders$bar_horizontal_upper
    } else if (posicao == "inferior") {
      left <- borders$corner_lower_left
      right <- borders$corner_lower_right
      bar <- borders$bar_horizontal_lower
    } else if (posicao == "separador") {
      left <- borders$separator_left
      right <- borders$separator_right
      bar <- borders$bar_horizontal_upper
    }

    margem_espacamento <-
      rep(
        paste0(
          borders$bar_vertical_left,
          strtrim(strrep(" ", tamanho - 2), tamanho - 2),
          borders$bar_vertical_right,
          "\n"
        ),
        if (margem > 1) floor(margem / 2) else 0
      )

    return(
      c(
        if (posicao %in% c("inferior", "separador")) margem_espacamento,
        paste0(
          left,
          strtrim(strrep(bar, tamanho - 2), tamanho - 2),
          right,
          if (posicao %in% c("superior", "separador")) "\n"
        ),
        if (posicao %in% c("superior", "separador")) margem_espacamento,
        if (timestamp && posicao == "inferior") .log_tempo_decorrido(),
        if (posicao == "inferior") "\n"
      )
    )
  }

  if (tipo == "conteudo") {
    margem_espacamento <- strrep(" ", margem)
    prefixo <- if (!is.null(conteudo$arrow) && conteudo$arrow) {
      paste0(margem_espacamento, " → ")
    } else {
      margem_espacamento
    }
    texto <- conteudo$texto
    espacamento <-
      max(
        0,
        tamanho - nchar(texto) - nchar(prefixo) - nchar(margem_espacamento) - 2
      )

    return(
      paste0(
        borders$bar_vertical_left,
        prefixo,
        texto,
        strrep(" ", espacamento),
        margem_espacamento,
        borders$bar_vertical_right,
        "\n"
      )
    )
  }

  if (tipo == "secao") {
    # + 3 = left and right borders and separator
    tamanho <-
      max(
        tamanho,
        nchar(conteudo$titulo) + nchar(conteudo$subtitulo) + (margem * 4) + 3
      )
    largura_coluna1 <- nchar(conteudo$titulo) + (margem * 2)
    largura_coluna2 <- tamanho - largura_coluna1 - 3

    margem_espacamento <- ""
    bar <- ""
    conteudo_1 <- ""
    conteudo_2 <- ""
    quebra_de_linha_superior <- ""
    quebra_de_linha_inferior <- ""
    stamp <- ""

    if (posicao == "meio") {
      espacamento <-
        .log_calcular_espacamento(largura_coluna2, nchar(conteudo$subtitulo))

      left <- borders$bar_vertical_left
      conteudo_1 <- paste0(
        strrep(" ", margem),
        conteudo$titulo,
        strrep(" ", margem)
      )
      separador <- borders$bar_vertical_right
      conteudo_2 <- paste0(
        strrep(" ", espacamento$esquerda),
        conteudo$subtitulo,
        strrep(" ", espacamento$direita)
      )
      right <- borders$bar_vertical_right
      quebra_de_linha_inferior <- "\n"

      margem_espacamento <-
        rep(
          paste0(
            left,
            strrep(" ", largura_coluna1),
            separador,
            strrep(" ", largura_coluna2),
            right,
            "\n"
          ),
          if (margem > 1) floor(margem / 2) else 0
        )
    }

    if (posicao == "superior") {
      quebra_de_linha_superior <- "\n"
      left <- borders$corner_upper_left
      bar <- borders$bar_horizontal_upper
      separador <- borders$separator_upper
      right <- borders$corner_upper_right
      quebra_de_linha_inferior <- "\n"
    }
    if (posicao == "inferior") {
      left <- borders$corner_lower_left
      bar <- borders$bar_horizontal_lower
      separador <- borders$separator_lower
      right <- borders$corner_lower_right
      stamp <- if (timestamp) .log_tempo_decorrido()
      quebra_de_linha_inferior <- "\n"
    }

    return(
      c(
        margem_espacamento,
        paste0(
          quebra_de_linha_superior,
          left,
          strtrim(strrep(bar, largura_coluna1), largura_coluna1),
          conteudo_1,
          separador,
          strtrim(strrep(bar, largura_coluna2), largura_coluna2),
          conteudo_2,
          right,
          stamp,
          quebra_de_linha_inferior
        ),
        margem_espacamento
      )
    )
  }
}

.log_processar_conteudo <- function(..., estilo) {
  #' Processa conteúdo variável e retorna lista de linhas formatadas
  #'
  #' @param ... Conteúdo a ser processado
  #' @return Lista com elementos do tipo list(texto = "...", arrow = FALSE)

  linhas <- list()

  if (...length() == 0) {
    return(linhas)
  }

  margem <- get_config("skin")$tamanho_margem

  for (item in list(...)) {
    # Pula NULLs
    if (is.null(item)) next

    # Dataframes são processados externamente
    if (is.data.frame(item)) {
      linhas[[length(linhas) + 1]] <- list(
        tipo = "dataframe",
        conteudo = item
      )
      next
    }

    # Converte listas para vetor
    if (is.list(item)) {
      item <- as.vector(unlist(item))
    }

    # Múltiplos itens (com seta)
    if (length(item) > 1) {
      for (sub_item in item) {
        linhas[[length(linhas) + 1]] <- list(
          tipo = "conteudo",
          texto = paste0(sub_item, strrep(" ", margem)),
          arrow = TRUE
        )
      }
      next
    }

    # Separador
    if (item == "-") {
      linhas[[length(linhas) + 1]] <- list(
        tipo = "borda",
        posicao = "separador"
      )
      next
    }

    # Item único
    linhas[[length(linhas) + 1]] <- list(
      tipo = "conteudo",
      texto = paste0(item, strrep(" ", margem)),
      arrow = FALSE
    )
  }

  linhas
}

.log_exibir_box <- function(
  tipo,
  titulo = NULL,
  subtitulo = NULL,
  cores = NULL,
  timestamp = TRUE,
  borda_superior = TRUE,
  borda_inferior = TRUE,
  linhas = NULL
) {
  #' Sistema unificado de renderização de boxes
  #'
  #' @param tipo "info" ou "secao"
  #' @param titulo Texto do título (obrigatório para tipo "secao")
  #' @param subtitulo Texto do subtítulo (obrigatório para tipo "secao")
  #' @param cores Esquema de cores (opcional)
  #' @param timestamp Incluir timestamp (default TRUE)
  #' @param borda_superior Desenhar borda superior (default TRUE)
  #' @param borda_inferior Desenhar borda inferior (default TRUE)
  #' @param linhas Conteúdo processado (obrigatório para tipo "info")

  if (tipo == "info" && is.null(linhas)) {
    stop("Tipo 'info' requer o parâmetro 'linhas'")
  }
  if (tipo == "secao" && (is.null(titulo) || is.null(subtitulo))) {
    stop("Tipo 'secao' requer os parâmetros 'titulo' e 'subtitulo'")
  }

  if (!is.null(cores)) utils_color(cores)

  if (tipo == "info") {
    if (borda_superior) {
      cat(.log_montar_linha("borda", "superior"), sep = "")
    }

    for (linha in linhas) {
      if (linha$tipo == "dataframe") {
        print.data.frame(linha$conteudo, right = FALSE, row.names = FALSE)
        next
      }

      cat(.log_montar_linha(linha$tipo, linha$posicao, linha), sep = "")
    }

    if (borda_inferior) {
      cat(
        .log_montar_linha("borda", "inferior", timestamp = timestamp),
        sep = ""
      )
    }
  }

  if (tipo == "secao") {
    conteudo <- list(
      titulo = titulo,
      subtitulo = subtitulo
    )

    cat(.log_montar_linha("secao", "superior", conteudo), sep = "")

    cat(.log_montar_linha("secao", "meio", conteudo), sep = "")

    cat(.log_montar_linha("secao", "inferior", conteudo, timestamp), sep = "")
  }

  if (!is.null(cores)) utils_color("default")
}
