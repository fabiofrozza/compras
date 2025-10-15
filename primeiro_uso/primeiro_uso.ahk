#NoEnv
#SingleInstance, force
SetWorkingDir %A_ScriptDir%
#Include, ..\_common\config.ahk
Menu, Tray, Icon, ..\_common\images\ufsc_ico.png

; --- Configuração da interface gráfica ---
Gui, Add, Picture, y10 w120 h120, ..\_common\images\DCOM.png
Gui, Add, GroupBox, x150 y10 w330 h120, Configuração Inicial - R
Gui, Add, Text, x170 y40 w300, Este assistente irá auxiliar na instalação do aplicativo R (caso já não instalado) e na instalação e configuração dos pacotes necessários para o funcionamento dos scripts.`n`nEscolha abaixo a opção desejada:

Gui, Add, Button, xp+135 y+40 w80 gExecutar vBtnExecutar, Iniciar
Gui, Add, Button, x+10 w80 gCancelar, Cancelar
Gui, Show, w500 h180, Configuração Inicial - R
Gui, Flash

Return

; --- Botão "Executar" ---
Executar:

    if !VerificarDependencias("primeiro_uso.R") {
        return
    }

    rpath := LocalizarRPath()  
    if (rpath = "") {
        MsgBox, 48, R não instalado, O programa R não está instalado.`n`nSerá executado o instalador.`n`nAceite todas as opções exibidas.
        if (FileExist("R-4.5.0-win.exe")) {
            Gui, +Disabled  ; Desabilita todos os controles
            GuiControl, , BtnExecutar, Aguarde...

            RunWait, R-4.5.0-win.exe, , Max

            ; Reabilita a GUI principal
            Gui, -Disabled
            GuiControl, , BtnExecutar, Iniciar

            MsgBox, 48, R instalado, Aplicativo R instalado com sucesso.`n`nO assistente será reiniciado.
            Gosub, Executar
        } else {
            MsgBox, 36, Erro, Arquivo de instalação "R-4.5.0-win.exe" não encontrado na pasta do script.`n`nDeseja tentar baixar automaticamente agora?
            IfMsgBox, Yes
            {
                TrayTip, Download do instalador do R, Baixando o instalador do R do endereço "https://cran.r-project.org/bin/windows/base/old/4.5.0/"`nAguarde..., 30, 1
                UrlDownloadToFile, https://cran.r-project.org/bin/windows/base/old/4.5.0/R-4.5.0-win.exe, R-4.5.0-win.exe
                TrayTip
                if (FileExist("R-4.5.0-win.exe")) {
                    MsgBox, 48, Download concluído, O instalador foi baixado com sucesso!`n`nO assistente continuará normalmente.
                    Gosub, Executar
                } else {
                    MsgBox, 16, Erro, Falha ao baixar o instalador.`n`nVerifique sua conexão com a internet e tente novamente.
                }
            }
            else
            {
                MsgBox, 48, Atenção, Procure pelo arquivo "R-4.5.0-win.exe" na internet, baixe manualmente e coloque-o na mesma pasta deste script.
            }
        }
        return
    }

    MsgBox, 52, Configuração inicial, R localizado em: "%rpath%"`nSerá executado o script "primeiro_uso.R".`n`nIsto aqui pode demorar bastante, então vá tomar um café e volte daqui a meia hora 😁`n`nDeseja continuar?

    IfMsgBox, Yes
    {
        ; Desabilita a GUI principal
        Gui, +Disabled  ; Desabilita todos os controles
        GuiControl, , BtnExecutar, Aguarde...

        ; Executa o script R
        RunWait, %rpath%\bin\Rscript.exe --vanilla primeiro_uso.R, , Max

        VerificarResultadoR(nomescript := "PRIMEIRO")

        ; Reabilita a GUI principal
        Gui, -Disabled

        ExitApp
    }

return

; --- Botão "Cancelar" ---
Cancelar:
    ExitApp
return

; --- Fechar a GUI ---
GuiClose:
    ExitApp
return
