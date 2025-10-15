primeiro_uso_main <- function() {
  
  cat("=== INICIANDO SCRIPT. AGUARDE... ===\n")
  
  source(file.path("..", "_common", "config.R"), chdir = TRUE)
  
  pacotes <- c("openxlsx", "readxl", "stringi", "stringr",
               "googlesheets4", "openxlsx2", "tidyr", "dplyr",
               "rmarkdown", "knitr", "pandoc", "pivottabler", 
               "pagedown", "kableExtra",
               "pdftools", "RColorBrewer")
  
  config_inicializar(pacotes)

  log_secao("TENTANDO OS PACOTES DE NOVO, SÓ PRA GARANTIR")

  config_pacotes(pacotes)
  
  log_secao("PACOTES EXTRAS")
  
  #se não existe pandoc, o instala
  tryCatch({
    if (!pandoc_available()) {
      log_info("Pandoc não disponível. Instalando...")
      pandoc::pandoc_install()
    }
    log_info("Pandoc instalado. Ativando...")
    pandoc::pandoc_activate()
  },
  error = function(e) { 
    log_erro_tratar("Não foi possível instalar o pacote Pandoc.", 
                    e) 
  })
  
  #finaliza e grava log
  config_finalizar()

}

primeiro_uso_main()
