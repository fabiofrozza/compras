#NoEnv
;#NoTrayIcon
SetWorkingDir %A_ScriptDir%

; --- Configuração da interface gráfica ---
Menu, Tray, Icon, ..\_common\images\ufsc_ico.png
Gui, Add, Picture, y20 w120 h120, ..\_common\images\DCOM.png
Gui, Add, GroupBox, x150 y10 w330 h150, Atualização do PowerShell
Gui, Add, Text, x170 y40 w300, Esta é uma etapa opcional para atualizar o PowerShell no seu computador.`n`nO Windows já possui a versão 5 que é compatível com o script para importação das planilhas.`n`nDeseja instalar a nova versão do PowerShell 7.5.1 no seu computador?
Gui, Add, Button, xp+130 y+35 w80 gInstalar, Instalar
Gui, Add, Button, x+10 w80 gCancelar, Cancelar
Gui, Show, w490 h220, Atualizar PowerShell
Gui, Flash

return

; --- Botão "Instalar" ---
Instalar:
    EnvGet, A_LocalAppData, LocalAppData
    ; 1. Cria a pasta de destino (se não existir)
    pastaDestino := A_LocalAppData . "\PowerShell\7"
    if !FileExist(pastaDestino)
        FileCreateDir, %pastaDestino%

    ; 2. Descompacta o ZIP (supondo que ele está na mesma pasta do script)
    arquivoZip := "PowerShell-7.5.1-win-x64.zip"
    if !FileExist(arquivoZip) {
        MsgBox, 36, Erro, Arquivo de instalação "%arquivoZip%" não encontrado na pasta do script.`n`nDeseja tentar baixar automaticamente agora?
        IfMsgBox, Yes
        {
            TrayTip, Download do PowerShell, Baixando PowerShell do endereço "https://github.com/powershell/powershell/releases"`nAguarde..., 30, 1
            UrlDownloadToFile, https://github.com/PowerShell/PowerShell/releases/download/v7.5.1/PowerShell-7.5.1-win-x64.zip, %arquivoZip%
            TrayTip
            if (FileExist(arquivoZip)) {
                MsgBox, 48, Download concluído, PowerShell baixado com sucesso!`n`nO assistente continuará normalmente.
                Gosub, Instalar
                Return
            } else {
                MsgBox, 16, Erro, Falha ao baixar o PowerShell.`n`nVerifique sua conexão com a internet e tente novamente.
                Return
            }
        }
        else
        {
            MsgBox, 48, Atenção, Procure pelo arquivo "%arquivoZip%" na internet, baixe manualmente e coloque-o na mesma pasta deste script.
            Return
        }
    }

    ; Usa o comando 'tar' do Windows (10/11) para descompactar
    MsgBox, 36, Instalação, Será aberta uma nova janela onde serão descompactados os arquivos do PowerShell 7 na pasta "%pastaDestino%"`n`nAguarde a finalização e não feche a janela.`n`nDeseja continuar?
    IfMsgBox, No
        Return

    ; Executa o comando tar para descompactar o arquivo
    RunWait, tar -xvf "%arquivoZip%" -C "%pastaDestino%", , Max

    ; Verifica se deu erro
    if ErrorLevel {
        MsgBox, 16, Erro, Falha na instalação!`n`n(Código de erro: %ErrorLevel%)
    } else {
        MsgBox, 64, Sucesso, PowerShell 7 instalado em:`n"%pastaDestino%"`n`nEste assistente será encerrado...
    }

    ExitApp
return

; --- Botão "Cancelar" ---
Cancelar:
    ExitApp
return

; --- Fechar a GUI ---
GuiClose:
    ExitApp
return