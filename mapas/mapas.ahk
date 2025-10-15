#NoEnv
#SingleInstance, force
SetWorkingDir %A_ScriptDir%
#Include, ..\_common\config.ahk
Menu, Tray, Icon, ..\_common\images\ufsc_ico.png

; --- Interface Gráfica ---
nomeJanela := "Mapas - Listas Prévias"
Gui, Add, Picture, y10 w120 h120, ..\_common\images\DCOM.png
Gui, Add, GroupBox, x150 y10 w420 h130, Listas Prévias - MAPAS
Gui, Add, Text, x170 y40 w360, Escolha a forma de geração da(s) lista(s) prévia(s):`n`n- Por Processo: gera uma lista prévia para cada processo`n- Combinada dos Grupos: gera uma única lista prévia de todos os Mapas ; +50 na largura
Gui, Add, Button, x170 y105 w175 gSeparar, Por Processo
Gui, Add, Button, x+30 w175 gCombinar, Combinada dos Grupos

; --- Listas de arquivos ---
Gui, Add, GroupBox, x10 y150 w270 h200, Pasta MAPAS
Gui, Add, ListBox, x20 y170 w250 h140 vListaMapas
Gui, Add, Button, x130 y315 w140 gAbrirPastaMapas, Abrir Pasta

Gui, Add, GroupBox, x300 y150 w270 h200, Listas prévias geradas
Gui, Add, ListBox, x310 y170 w250 h140 vListaXLS
Gui, Add, Button, x430 y315 w130 gAbrirPastaListas, Abrir Pasta

; Botão Atualizar Listas (ao lado esquerdo do Sair)
Gui, Add, Button, x370 y365 w90 gAtualizarListas, Atualizar listas
Gui, Add, Button, x470 y365 w90 gCancelar, Sair

Gui, Show, w580 h400, %nomeJanela%
Gui, Flash

AtualizarListasFunc() ; Chama a função ao iniciar
return

AtualizarListas:
    AtualizarListasFunc()
return

AtualizarListasFunc() {
    ; Limpa lista de arquivos MAPAS
    GuiControl,, ListaMapas, |
    mapas := ""
    Loop, Files, %A_ScriptDir%\MAPAS\*, F
        mapas .= A_LoopFileName "|"
    GuiControl,, ListaMapas, %mapas%

    ; Limpa lista de arquivos XLS
    GuiControl,, ListaXLS, |
    xls := ""
    Loop, Files, %A_ScriptDir%\LISTAS\*.xls*, F
        xls .= A_LoopFileName "|"
    GuiControl,, ListaXLS, %xls%
}
return

Separar:
    tipo := "processo"
    Gosub, Executar
return

Combinar:
    tipo := "grupo"
    Gosub, Executar
return

; --- Executar script ---
Executar:

    ; Verificar se existem arquivos na pasta MAPAS
    arquivosMapas := 0
    Loop, Files, %A_ScriptDir%\MAPAS\*, F
    {
        arquivosMapas++
        break  ; Só precisamos saber se existe pelo menos 1 arquivo
    }
    
    if (arquivosMapas = 0) {
        MsgBox, 16, Pasta vazia, A pasta MAPAS está vazia.`n`nBaixe os Mapas de Licitação dos processos escolhidos e tente novamente.
        return
    }

    rpath := LocalizarRPath()  
    if (rpath = "") {
        MsgBox, 48, R não instalado, O programa R não está instalado.`n`nPor favor, instale e tente novamente.
        ExitApp
    }

    TrayTip, Mapas - Listas Prévias, R localizado em: "%rpath%"`n`nAguarde a geração da lista (%tipo%)..., 10
    ;MsgBox, 64, Iniciando, R localizado em: "%rpath%"`n`nAguarde a geração da lista (%tipo%)...
    
    ; Desabilita a GUI principal
    Gui, +Disabled  ; Desabilita todos os controles

    RunWait, %rpath%\bin\Rscript.exe --vanilla mapas.R %tipo%, , Max

    TrayTip

    VerificarResultadoR("MAPAS")

    ; Reabilita a GUI principal
    Gui, -Disabled
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

AbrirPastaGenerica(pasta) {
    Run, explorer.exe "%pasta%"
}

AbrirPastaMapas:
    AbrirPastaGenerica(A_ScriptDir "\MAPAS")
return

AbrirPastaListas:
    AbrirPastaGenerica(A_ScriptDir "\LISTAS")
return
