#NoEnv
SetWorkingDir %A_ScriptDir%
#Include, ..\_common\config.ahk
Menu, Tray, Icon, ..\_common\images\company.png

Gui, Add, Picture, y20 w120 h120, ..\_common\images\department.png
Gui, Add, GroupBox, x150 y10 w330 h150, Atualização do PowerShell
Gui, Add, Text, x170 y40 w300, Esta é uma etapa opcional para atualizar o PowerShell no seu computador.`n`nO Windows já possui a versão 5 que é compatível com o script para importação das planilhas.`n`nDeseja instalar a nova versão do PowerShell no seu computador?
Gui, Add, Button, xp+130 y+35 w80 gInstalar vBtnInstalar, Instalar
Gui, Add, Button, x+10 w80 gCancelar vBtnCancelar, Cancelar
Gui, Show, w490 h220, Atualizar PowerShell
Gui, Flash

return

Instalar:
    EnvGet, A_LocalAppData, LocalAppData

    pastaDestino := A_LocalAppData . "\PowerShell\7"
    if !FileExist(pastaDestino)
        FileCreateDir, %pastaDestino%

    arquivoZip := GetFileName("*.zip")

    if !FileExist(arquivoZip) {
        MsgBox, 36, Erro, Arquivo de instalação do PowerShell não encontrado na pasta do script.`n`nDeseja tentar baixar automaticamente agora?
        IfMsgBox, No
        {
            MsgBox, 48, Atenção, Procure pelo arquivo de instalação do PowerShell (versão mínima 7.5, formato zip) na internet, baixe manualmente e coloque-o na mesma pasta deste script.
            Return
        }

        arquivoZip := "PowerShell-7.5.1-win-x64.zip"
        url := "https://github.com/PowerShell/PowerShell/releases/download/v7.5.1/PowerShell-7.5.1-win-x64.zip"

        TrayTip, Download do PowerShell, Baixando PowerShell do endereço "https://github.com/powershell/powershell/releases"`nAguarde..., 30, 1
            
        GuiControl, , BtnInstalar, Aguarde...
        GuiControl, , BtnCancelar, Aguarde...
        Gui, +Disabled

        UrlDownloadToFile, %url%, %arquivoZip%
                    
        Gui, -Disabled
        GuiControl, , BtnInstalar, Instalar
        GuiControl, , BtnCancelar, Cancelar

        TrayTip

        if ErrorLevel {
            MsgBox, 16, Erro, Falha ao baixar o PowerShell.`n`nVerifique sua conexão com a internet e tente novamente.
        } else {
            MsgBox, 48, Download concluído, PowerShell baixado com sucesso!`n`nO assistente continuará normalmente.
            Gosub, Instalar
        }
        
        Return
    }

    MsgBox, 36, Instalação, Será aberta uma nova janela onde serão descompactados os arquivos do PowerShell 7 na pasta "%pastaDestino%"`n`nAguarde a finalização e não feche a janela.`n`nDeseja continuar?
    IfMsgBox, No
        Return

    GuiControl, , BtnInstalar, Aguarde...
    GuiControl, , BtnCancelar, Aguarde...
    Gui, +Disabled

    RunWait, tar -xvf "%arquivoZip%" -C "%pastaDestino%", , Max

    if ErrorLevel {
        MsgBox, 16, Erro, Falha na instalação!`n`n(Código de erro: %ErrorLevel%)
    } else {
        MsgBox, 64, Sucesso, PowerShell 7 instalado em:`n"%pastaDestino%"`n`nEste assistente será encerrado...
    }
    
    ExitApp
return

Cancelar:
    ExitApp
return

GuiClose:
    ExitApp
return