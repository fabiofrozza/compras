#NoEnv
#SingleInstance, force
SetWorkingDir %A_ScriptDir%
#Include, ..\_common\config.ahk
Menu, Tray, Icon, ..\_common\images\company.png

Gui, Add, Picture, y10 w120 h120, ..\_common\images\department.png
Gui, Add, GroupBox, x150 y10 w330 h120, Configuração Inicial - R
Gui, Add, Text, x170 y40 w300, Este assistente irá auxiliar na instalação do aplicativo R (caso já não instalado) e na instalação e configuração dos pacotes necessários para o funcionamento dos scripts.`n`nEscolha abaixo a opção desejada:

Gui, Add, Button, xp+135 y+40 w80 gExecutar vBtnExecutar, Iniciar
Gui, Add, Button, x+10 w80 gCancelar vBtnCancelar, Cancelar
Gui, Show, w500 h180, Configuração Inicial - R
Gui, Flash

Return

Executar:

    if !VerificarDependencias("primeiro_uso.R") {
        return
    }

    rpath := LocalizarRPath()

    ; If R is installed
    if (rpath != "") {

        MsgBox, 52, Configuração inicial, R localizado em: "%rpath%"`n`nSerá executado o script "primeiro_uso.R".`n`nIsto aqui pode demorar bastante, então vá tomar um café e volte daqui a meia hora 😁`n`nDeseja continuar?

        IfMsgBox, No
            return

        Gui, +Disabled
        GuiControl, , BtnExecutar, Aguarde...
        GuiControl, , BtnCancelar, Aguarde...

        RunWait, %rpath%\bin\Rscript.exe --vanilla primeiro_uso.R, , Max

        VerificarResultadoR(nomescript := "PRIMEIRO")

        ExitApp
    }

    ; If R isn´t installed
    arquivoR := GetFileName("R*.exe", "instalação do programa R")
        
    ; If there is installation file
    if (FileExist(arquivoR)) {
        MsgBox, 52, R não instalado, O programa R não está instalado.`n`nDeseja executar o instalador?
        IfMsgBox, No
            ExitApp

        MsgBox, 48, R será instalado em breve, Aguarde a execução do instalador do R.`n`nIsto pode levar alguns segundos... Tenha paciência...`n`nDurante a instalação, aceite todas as opções exibidas clicando em OK/Continuar/Próximo ou equivalente.

        Gui, +Disabled
        GuiControl, , BtnExecutar, Aguarde...
        GuiControl, , BtnCancelar, Aguarde...

        RunWait, %arquivoR%, , Max

        Gui, -Disabled
        GuiControl, , BtnExecutar, Iniciar
        GuiControl, , BtnCancelar, Cancelar

        MsgBox, 48, R instalado, Aplicativo R instalado com sucesso.`n`nO assistente será reiniciado.
        Gosub, Executar
    } 
    
    ; If there isn´t installation file
    MsgBox, 36, R não instalado, O programa R não está instalado e não há arquivo de instalação na pasta do script.`n`nDeseja tentar baixar automaticamente agora?
    IfMsgBox, No
        {
            MsgBox, 48, Atenção, Procure por "CRAN download R" na internet, baixe manualmente o arquivo de instalação e coloque-o na mesma pasta deste script.
            return
        }
    
    arquivoR := GetLatestR()
    if (arquivoR = "") {
        MsgBox, 16, Erro, Não foi possível localizar a última versão do aplicativo R para download.`n`nVerifique sua conexão e tente novamente.
        ExitApp
    }
    version := 
    RegExMatch(arquivoR, "[\d.]+", version)
    url := "https://cran.r-project.org/bin/windows/base/old/" . version . "/" . arquivoR

    TrayTip, Download do instalador do R, Baixando o instalador do R. Aguarde..., 60, 1

    Gui, +Disabled
    GuiControl, , BtnExecutar, Aguarde...
    GuiControl, , BtnCancelar, Aguarde...

    try {
        UrlDownloadToFile, %url%, %arquivoR%
    } catch e {
        TrayTip
        MsgBox, 16, Erro, Falha ao baixar o instalador.`n`nVerifique sua conexão com a internet e tente novamente.
        ExitApp
    }

    TrayTip
    
    MsgBox, 48, Download concluído, O instalador foi baixado com sucesso!`n`nO assistente continuará normalmente.
    Gosub, Executar

    Gui, -Disabled
    GuiControl, , BtnExecutar, Iniciar
    GuiControl, , BtnCancelar, Cancelar

    return

Cancelar:
    ExitApp
    return

GuiClose:
    ExitApp
    return
