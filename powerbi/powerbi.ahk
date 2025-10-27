#NoEnv
#SingleInstance, Force
SetWorkingDir %A_ScriptDir%
#Include, ../_common/config.ahk
Menu, Tray, Icon, ..\_common\images\company.png

if (A_Args.Length() > 0 && A_Args[1] = "silent") {
    rpath := LocalizarRPath()
    TrayTip, Power BI, A geração dos arquivos de dados está ocorrendo em segundo plano., 15
    RunWait, %rpath%\bin\Rscript.exe --vanilla powerbi.R todos silent, , Hide
    TrayTip, Power BI, Script finalizado. Verifique os arquivos e o log., 30
    ExitApp
}

; --- Interface Gráfica ---
nomeJanela := "Geração de Dados para Power BI"
Gui, Add, Picture, x20 y70 w120 h120, ..\_common\images\department.png
Gui, Add, GroupBox, x150 y10 w660 h200, Geração de Dados para Power BI
Gui, Add, Text, xp+20 yp+25 w470, Selecione o painel para gerar os dados para o Power BI (ou as opções de manutenção) e clique em Executar para mais informações.

; Grupos
Gui, Add, GroupBox, y+15 w290 h120, Painel do Observatório
Gui, Add, GroupBox, x+25 w180 h95, Manutenção

; --- Opções de Relatório ---
Gui, Add, Radio, xp-305 yp+25 w180 vTipoRelatorio Checked, Planejamento
Gui, Add, Radio, y+10, Licitação 
Gui, Add, Radio, y+10, Execução
Gui, Add, Radio, y+10 w270, Processos Administrativos e Alterações Contratuais
Gui, Add, Radio, xp+315 yp-70 w120, Renomear arquivos
Gui, Add, Radio, y+10, Agendar execução automática
Gui, Add, Radio, y+10, Todos os Painéis

; Adicione um controle Edit para exibir a saída do script R
Gui, Add, Edit, x10 y230 w800 h260 vOutputLog ReadOnly +VScroll +HScroll +0x100 -Wrap, Aguardando execução...  ; +0x100 = WS_VSCROLL (barra vertical)
Gui, Font, s10 cDarkBlue, Consolas
GuiControl, Font, OutputLog
Gui, Font

; Botões principais
Gui, Add, Button, x+-110 y+15 w110 gCancelar, Sair
Gui, Add, Button, x690 y80 w105 h110 gExecutarRelatorio vBtnExecutar HwndBtnExecutar, Executar

Gui, Show, w820 h540, %nomeJanela%
Gui, Flash
return

ExecutarRelatorio:
    Gui, Submit, NoHide
    
    if !VerificarDependencias("powerbi.R") {
        return
    }

    ; Determina qual script R será executado
    if (TipoRelatorio = 1) {
        R_script := "planejamento"
        mensagem := "Painel Visão Planejamento`n`nSerá acessada a planilha de controle para recuperar os dados dos processos, Unidades requerentes e situação do envio da documentação."
    } else if (TipoRelatorio = 2) {
        R_script := "licitacao"
        mensagem := "Painel Visão Licitação`n`nSerão obtidos os dados dos Mapas de Licitação da pasta POWERBI/Mapa de licitações."
    } else if (TipoRelatorio = 3) {
        R_script := "execucao"
        mensagem := "Painel Visão Execução`n`nSerão obtidos os dados dos relatórios de execução das AFs/Empenhos da pasta POWERBI/Execucao AF Empenho."
    } else if (TipoRelatorio = 4) {
        R_script := "paalteracoes"
        mensagem := "Painel Processos Administrativos e Alterações Contratuais`n`nSerá acessada a planilha de controle para recuperar os dados dos processos administrativos, trocas de marca, cancelamentos e reequilíbrios."
    } else if (TipoRelatorio = 5) {
        R_script := "renomear"
        mensagem := "Os arquivos das pastas Mapa de licitações e Execucao AF Empenho serão renomeados conforme o padrão TIPO - ANO - ETAPA - PROCESSO - PREGÃO."
    } else if (TipoRelatorio = 6) {
        mensagem := "Configura (ou exclui) a execução automática deste assistente no Agendador de Tarefas do Windows.`n`nO agendamento permite que, todos os dias, no horário escolhido, sejam gerados os dados para os quatro Painéis do Observatório."
    } else if (TipoRelatorio = 7) {
        R_script := "todos"
        mensagem := "Serão gerados dados para todos os Painéis do Observatório."
    }
    
    MsgBox, 52, Confirmação, %mensagem%`n`nDeseja continuar?
    IfMsgBox, No
        return
    
    rpath := LocalizarRPath()  
    if (rpath = "") {
        MsgBox, 52, R não instalado, O programa R não está instalado ou não foi detectado.`n`nÉ necessário instalar o R para continuar.`n`nDeseja abrir agora a página de download do R para Windows?
        IfMsgBox, Yes
        {
            ; Abre o site do CRAN com a versão mais recente para Windows
            Run, https://cran.r-project.org/bin/windows/base/        
        }
        return
    } 

    if (TipoRelatorio = 6) {

        nomeTarefa := "PowerBI_Automatico"

        ; Verifica se a tarefa já existe
        RunWait, schtasks /Query /TN %nomeTarefa%, , Hide UseErrorLevel
        if (ErrorLevel = 0) {
            MsgBox, 36, Tarefa já existe, A tarefa "%nomeTarefa%" já existe.`n`nPara configurar novamente, escolha 'Sim'.`nPara excluir o agendamento, escolha 'Não'
            IfMsgBox, No
            {
                RunWait, schtasks /Delete /TN %nomeTarefa% /F, , Hide UseErrorLevel
                MsgBox, 64, Sucesso, Tarefa "%nomeTarefa%" excluída com sucesso!

                return
            }
            ; Se Sim, continua e recria a tarefa
        }

        ; Solicita ao usuário o horário desejado
        InputBox, horaTarefa, Horário do agendamento, Informe o horário para execução diária do script (formato 24:00):, , 500, 150
        if (ErrorLevel || horaTarefa = "")
            return

        ; Validação do formato HH:MM (00:00 a 23:59)
        if !RegExMatch(horaTarefa, "^(?:[01]\d|2[0-3]):[0-5]\d$")
        {
            MsgBox, 16, Erro, Horário inválido!`n`nPor favor, informe no formato HH:MM (hora com dois dígitos, de 00 a 24, e minutos com dois dígitos).
            return
        }

        ; Cria uma tarefa agendada
        scriptPath := A_ScriptFullPath
        RunWait, schtasks /Create /F /TN %nomeTarefa% /TR "'%scriptPath%' silent" /SC DAILY /ST %horaTarefa%, , Hide UseErrorLevel

        if (ErrorLevel) {
            MsgBox, 16, Erro, Não foi possível criar a tarefa agendada.`n`nExecute novamente este assistente com as permissões de Administrador.
        } else {
            MsgBox, 64, Sucesso, Tarefa agendada criada com sucesso!`n`nTodos os dias, às %horaTarefa%, serão gerados os dados para os três Painéis do Observatório.
        }
        return
    }

    Gui, Font, s10 cDarkBlue, Consolas
    GuiControl, Font, OutputLog

    ; Limpa o log anterior
    GuiControl,, OutputLog, Preparando para executar o script R...`n

    ; Desabilita a GUI principal
    GuiControl,, BtnExecutar, Aguarde...`n`n⏳  ; Altera o texto do botão
    Gui, +Disabled  ; Desabilita todos os controles

    ; Cria um arquivo temporário para capturar a saída
    tempFile := A_Temp "\R_output.txt"
    if FileExist(tempFile)
        FileDelete, %tempFile%

    ; Remove o arquivo temporário se existir (duplo check)
    if FileExist(tempFile)
        FileDelete, %tempFile%
    
    ; Executa o script R em segundo plano
    Run, %ComSpec% /c ""%rpath%\bin\Rscript.exe" --vanilla powerbi.R %R_script% silent > "%tempFile%" 2>&1", , Hide, R_PID

    ; Inicia o monitoramento da saída do R
    SetTimer, MonitorarSaidaR, 500

return

Cancelar:
    ExitApp
return

GuiClose:
    ExitApp
return

MonitorarSaidaR:
    global tempFile, R_PID
    if FileExist(tempFile) {
        FileRead, output, *P65001 %tempFile%
        GuiControl,, OutputLog, %output%
        ; Rola automaticamente para o final
        SendMessage, 0x0115, 7, 0, Edit1, A  ; WM_VSCROLL = 0x0115, SB_BOTTOM = 7
    }
    ; Verifica se o processo R ainda existe
    Process, Exist, %R_PID%
    if (ErrorLevel = 0) {
        ; Processo terminou, faz leitura final e verificação de erros
        SetTimer, MonitorarSaidaR, Off

        if FileExist(tempFile) {
            FileRead, output, *P65001 %tempFile%
            GuiControl,, OutputLog, %output%
            ; Rola automaticamente para o final
            SendMessage, 0x0115, 7, 0, Edit1, A  ; WM_VSCROLL = 0x0115, SB_BOTTOM = 7
        }

        ; Verificação de erros
        resultado_geracao := VerificarResultadoR("POWER", true)
        resultado_geracao := resultado_geracao[1]  ; Obtém somente o resultado_geracao retornado

        if (resultado_geracao = "erro") {
            MsgBox, 16, Erro, Ocorreram erros na execução e não foram gerados os arquivos.`n`nVerifique o andamento do script.
            Gui, Font, s10 c800000, Consolas ; Vermelho escuro (hex RGB)
            GuiControl, Font, OutputLog
        } else if (resultado_geracao = "ambos") {
            MsgBox, 48, Alerta, Ocorreram erros na execução, mas os arquivos foram gerados`n`nVerifique o andamento do script.
            Gui, Font, s10 cFFA500, Consolas ; Laranja (hex RGB)
            GuiControl, Font, OutputLog
        } else if (resultado_geracao = "sucesso") {
            MsgBox, 64, Sucesso, Script finalizado com sucesso!
        } else {
            MsgBox, 16, Erro, Ocorreu um erro desconhecido na execução.`n`nVerifique o andamento do scrip.
            Gui, Font, s10 c800000, Consolas ; Vermelho escuro (hex RGB)
            GuiControl, Font, OutputLog
        }

        ; Reabilita a GUI principal
        Gui, -Disabled
        GuiControl,, BtnExecutar, Executar
        WinRestore, %nomeJanela%
        WinActivate, %nomeJanela%
        WinWaitActive, %nomeJanela%, , 2
        if ErrorLevel
        {
            WinShow, %nomeJanela%
            WinActivate, %nomeJanela%
        }

        ; Remove o arquivo temporário
        FileDelete, %tempFile%
    }
return
