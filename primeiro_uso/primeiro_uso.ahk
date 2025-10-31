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

    ;rpath := LocalizarRPath()
    rpath := ""
    if (rpath = "") {
        
        arquivoR := GetFileName("R*.exe")
            
        if (FileExist(arquivoR)) {
            MsgBox, 52, R não instalado, O programa R não está instalado.`n`nDeseja executar o instalador?
            IfMsgBox, No
                ExitApp

            MsgBox, 48, R será instalado em breve, O R será instalado em breve.`n`nAceite todas as opções exibidas.

            Gui, +Disabled
            GuiControl, , BtnExecutar, Aguarde...
            GuiControl, , BtnCancelar, Aguarde...

            RunWait, %arquivoR%, , Max

            Gui, -Disabled
            GuiControl, , BtnExecutar, Iniciar
            GuiControl, , BtnCancelar, Cancelar

            MsgBox, 48, R instalado, Aplicativo R instalado com sucesso.`n`nO assistente será reiniciado.
            Gosub, Executar
        } else {
            arquivoR := "R-4.5.0-win.exe"
            MsgBox, 36, R não instalado, O programa R não está instalado e não há arquivo de instalação na pasta do script.`n`nDeseja tentar baixar automaticamente agora?
            IfMsgBox, Yes
            {
                url := "https://cran.r-project.org/bin/windows/base/old/4.5.0/R-4.5.0-win.exe"

                TrayTip, Download do instalador do R, Baixando o instalador do R do endereço "https://cran.r-project.org/bin/windows/base/old/4.5.0/"`nAguarde..., 30, 1

                Gui, +Disabled
                GuiControl, , BtnExecutar, Aguarde...
                GuiControl, , BtnCancelar, Aguarde...

                UrlDownloadToFile, %url%, %arquivoR%

                Gui, -Disabled
                GuiControl, , BtnExecutar, Iniciar
                GuiControl, , BtnCancelar, Cancelar

                TrayTip

                if ErrorLevel {
                    MsgBox, 16, Erro, Falha ao baixar o instalador.`n`nVerifique sua conexão com a internet e tente novamente.
                } else {
                    MsgBox, 48, Download concluído, O instalador foi baixado com sucesso!`n`nO assistente continuará normalmente.
                    Gosub, Executar
                }
            }
            else
            {
                MsgBox, 48, Atenção, Procure pelo arquivo %arquivoR% na internet, baixe manualmente e coloque-o na mesma pasta deste script.
            }
        }
        return
    }

    MsgBox, 52, Configuração inicial, R localizado em: "%rpath%"`nSerá executado o script "primeiro_uso.R".`n`nIsto aqui pode demorar bastante, então vá tomar um café e volte daqui a meia hora 😁`n`nDeseja continuar?

    IfMsgBox, Yes
    {
        Gui, +Disabled
        GuiControl, , BtnExecutar, Aguarde...
        GuiControl, , BtnCancelar, Aguarde...

        RunWait, %rpath%\bin\Rscript.exe --vanilla primeiro_uso.R, , Max

        VerificarResultadoR(nomescript := "PRIMEIRO")

        ExitApp
    }

return

Cancelar:
    ExitApp
return

GuiClose:
    ExitApp
return
