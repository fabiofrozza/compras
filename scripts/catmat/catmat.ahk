#NoEnv
#SingleInstance, Force
SetWorkingDir %A_ScriptDir%
#Include, ..\_common\config.ahk
Menu, Tray, Icon, ..\_common\images\company.png

; --- Interface Gráfica ---
nomeJanela := "Consulta de CATMAT"
Gui, Add, Picture, y40 w120 h120, ..\_common\images\department.png
Gui, Add, GroupBox, x150 y10 w420 h190, Consulta de CATMAT
Gui, Add, Text, x170 y40 w360, Cole o arquivo com a lista de itens do pregão na pasta TR.`n`nSelecione o arquivo (TR), o método de consulta e clique em Executar.

; --- Opções de Consulta ---
Gui, Add, GroupBox, x170 y100 w250 h80, Método de Consulta
Gui, Add, Button, x445 y110 w110 h65 gExecutarConsulta vBtnExecutar, Executar
Gui, Add, Radio, x180 y125 w230, API - Consulta direta aos Dados Abertos
Gui, Add, Radio, x180 y150 w230 vMetodoConsulta Checked, Lista de CATMATs - Consulta manual

; --- Listas de arquivos ---
Gui, Add, GroupBox, x10 y220 w270 h175, Lista de itens do TR
Gui, Add, ListBox, x20 y240 w250 h120 vListaTR
Gui, Add, Button, x120 y360 w150 gAbrirPastaTR, Abrir pasta TR

Gui, Add, GroupBox, x300 y220 w270 h175, Arquivos Gerados
Gui, Add, ListBox, x310 y240 w250 h120 vListaGerados
Gui, Add, Button, x375 y360 w185 gAbrirPastaGerados, Abrir pasta ARQUIVOS GERADOS

; Botões principais
Gui, Add, Button, x370 y+30 w90 gAtualizarListas, Atualizar
Gui, Add, Button, x+10 w90 gCancelar, Sair

Gui, Show, w580 h450, %nomeJanela%
Gui, Flash

AtualizarListasFunc() ; Chama a função ao iniciar
return

AtualizarListas:
    AtualizarListasFunc()
return

AtualizarListasFunc() {
    ; Limpa lista de arquivos TR
    GuiControl,, ListaTR, |
    trFiles := ""
    Loop, Files, %A_ScriptDir%\TR\*, F
        trFiles .= A_LoopFileName "|"
    GuiControl,, ListaTR, %trFiles%
    
    ; Limpa lista de arquivos gerados
    GuiControl,, ListaGerados, |
    gerados := ""
    Loop, Files, %A_ScriptDir%\ARQUIVOS GERADOS\*, F
        gerados .= A_LoopFileName "|"
    GuiControl,, ListaGerados, %gerados%
}
return

ExecutarConsulta:
    Gui, Submit, NoHide

    ; Obtém o método de consulta selecionado
    if (MetodoConsulta = 1) {
        tipo := "api"
        mensagem_tipo := "API de consulta direta aos Dados Abertos do Governo Federal"
    } else {
        tipo := "lista"
        mensagem_tipo := "lista de CATMATs baixada manualmente do Compras.gov.br"
    }

    ; Obtém o arquivo selecionado na lista TR
    GuiControlGet, arquivoTR, , ListaTR
    if (arquivoTR = "") {
        MsgBox, 48, Atenção, Selecione um arquivo na lista TR antes de executar.

        WinRestore, %nomeJanela%
        WinActivate, %nomeJanela%
        WinWaitActive, %nomeJanela%, , .5
        if ErrorLevel
        {
            WinShow, %nomeJanela%
            WinActivate, %nomeJanela%
        }

        return
    }

    ; Localiza o R instalado
    rpath := LocalizarRPath()  
    if (rpath = "") {
        MsgBox, 48, R não instalado, O programa R não está instalado.`n`nPor favor, instale a última versão para Windows em:`nhttps://cran.r-project.org/
        return
    }
    
    ; Verifica se os arquivos necessários existem na pasta _fontes
    if !FileExist(A_ScriptDir . "\_fontes\Lista CATMAT.xlsx") || !FileExist(A_ScriptDir . "\_fontes\Margens.xlsx") {
        MsgBox, 20, Arquivo não encontrado, Os arquivos "Lista CATMAT.xlsx" e/ou "Margens.xlsx" não foram encontrados na pasta "_fontes".`n`nA lista de CATMATs é baixada do Portal Compras.gov.br e a lista de margens é extraída manualmente da legislação.`n`nVerifique e execute o assistente novamente.`n`nDeseja aproveitar para pesquisar no Google por "lista catmat"?
        IfMsgBox, Yes
            Run, https://www.google.com/search?q=lista+catmat
        return
    }

    MsgBox, 52, Iniciando, R localizado em: "%rpath%"`n`nSerá efetuada a análise pela %mensagem_tipo%.`n`nDeseja continuar?
    IfMsgBox, No
        return

    ; Desabilita a GUI principal
    GuiControl,, BtnExecutar, Aguarde...`n`n⏳  ; Altera o texto do botão
    Gui, +Disabled  ; Desabilita todos os controles

    ; Passa o nome do arquivo selecionado como argumento
    RunWait, %rpath%\bin\Rscript.exe --vanilla catmat.R %tipo% "%arquivoTR%", , Max
    
    VerificarResultadoR("CATMAT")  ; Verifica o resultado da execução do script R

    ; Reabilita a GUI principal
    Gui, -Disabled
    GuiControl,, BtnExecutar, Executar  ; Restaura o texto do botão
    WinRestore, %nomeJanela%
    WinActivate, %nomeJanela%
    WinWaitActive, %nomeJanela%, , 2
    if ErrorLevel
    {
        WinShow, %nomeJanela%
        WinActivate, %nomeJanela%
    }

    AtualizarListasFunc()
    
return

; --- Cancelar ---
Cancelar:
    ExitApp
return

GuiClose:
    ExitApp
return

AbrirPastaTR:
    Run, explorer.exe "%A_ScriptDir%\TR"
return

AbrirPastaGerados:
    Run, explorer.exe "%A_ScriptDir%\ARQUIVOS GERADOS"
return