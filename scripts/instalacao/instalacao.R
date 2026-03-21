instalacao_main <- function() {
  source(file.path("..", "_common", "config.R"), chdir = TRUE)

  args <- commandArgs(trailingOnly = TRUE)
  modo <- if (length(args) >= 1) args[1] else "instalar"

  pacotes <- c(
    "openxlsx", "readxl", "stringi", "stringr",
    "googlesheets4", "openxlsx2", "tidyr", "dplyr",
    "rmarkdown", "knitr", "pandoc", "pivottabler",
    "pagedown", "kableExtra",
    "pdftools", "RColorBrewer"
  )

  if (modo == "atualizar") {
    # Inicializa sem carregar os pacotes para não travá-los no Windows
    config_inicializar("INSTALACAO", c())

    log_secao("ATUALIZANDO PACOTES")
    config_atualizar_pacotes(pacotes)
  } else {
    config_inicializar("INSTALACAO", pacotes)

    log_secao("TENTANDO OS PACOTES DE NOVO, SO PRA GARANTIR")
    config_pacotes(pacotes)
  }

  log_secao("PACOTES EXTRAS")

  tryCatch(
    {
      if (!pandoc_available()) {
        log_info("Pandoc nao disponivel. Instalando...")
        pandoc::pandoc_install()
      }
      log_info("Pandoc instalado. Ativando...")
      pandoc::pandoc_activate()
    },
    error = function(e) {
      log_erro(
        "Nao foi possivel instalar o pacote Pandoc.",
        e
      )
    }
  )

  config_finalizar()
}

instalacao_main()
