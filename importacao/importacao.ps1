# FUNÇÕES

function OpenFileFolder {
    param (
        [System.Object]$ctr_
    )

    InterfaceMinimize

    if ($ctr_ -is [System.Windows.Forms.ToolStripMenuItem]) {
        $lst_ = $ctr_.GetCurrentParent().SourceControl
    } elseif ($ctr_ -is [System.Windows.Forms.ListView]) {
        $lst_ = $ctr_
    } else {
        return
    }

    $selection = $lst_.SelectedItems[0]

    if ($ctr_.Text -eq "Abrir pasta" -or $lst_.Empty) { 
        $fle = "" 
    } else {
        $fle = $selection.Text
    }

    switch ($lst_) {
        {@($lst_importar, $lst_descricao) -contains $_} {
            Start-Process "$fldToImport\$fle"
        }
        {@($lst_prints, $lst_resumos) -contains $_} {
            Start-Process "$fldResumo\$fle"
        }
        {@($lst_dfd, $lst_relatorio) -contains $_} {
            Start-Process "$fldRelatorio\$fle"
        }
    }
}

function RefreshCmbProcessos {
    param(
        $processosSPA
    )

    $cmb_processos.Items.Clear()

    if ($processosSPA.Count -eq 0) {

        [void]$cmb_processos.Items.Add("Nenhum processo disponível")
                
    } else {
    
        [void]$cmb_processos.Items.Add("Todos - Relatório consolidado")

        foreach ($processo in $processosSPA) {
            [void]$cmb_processos.Items.Add($processo)
        }
    }

    $cmb_processos.SelectedIndex = 0
}

function RefreshBtnGerar {

    $IsSettingsValid = IsSettingsValid

    if ($IsSettingsValid -and $installedR -and $txt_link_planilha.validLink) {
        $btn_gerar.Enabled = $true
    } else {
        $btn_gerar.Enabled = $false
    }

}

function IsSettingsValid {

    $validationRules = @{
        "valor_minimo" = @{
            pattern = "^[0-9]+$";
            errorMsg = "Deve ser um número inteiro"
        }
        "qtde_minima" = @{
            pattern = "^[0-9]+$";
            errorMsg = "Deve ser um número inteiro"
        }
        "celula" = @{
            pattern = "^[A-Z]+[0-9]+$";
            errorMsg = "Célula de planilha excel, com letra(s) + número (ex: Q6)"
        }
        "unidades" = @{
            pattern = "^[0-9]+$";
            errorMsg = "Deve ser um número inteiro"
        }
        "aba_menu" = @{
            pattern = "^[\w\d\s]+$";
            errorMsg = "Apenas letras, números e espaços"
        }
        "aba_lista_final" = @{
            pattern = "^[\w\d\s]+$";
            errorMsg = "Apenas letras, números e espaços"
        }
    }
    
    $isAllSettingsValid = $true
    
    foreach ($row in $lst_settings.Rows) {
        $value = $row.Cells["Valor"].Value
        $rule  = $validationRules[$row.Tag]
        
        if ($value -match $rule.pattern) {
            $row.Cells["Icone"].Value = $lst_image.Images["config"]
            $row.Cells["Valor"].ToolTipText = ""
        } else {
            $row.Cells["Icone"].Value = $lst_image.Images["config_erro"]
            $row.Cells["Valor"].ToolTipText = $rule.errorMsg
            $isAllSettingsValid = $false
        }
    }

    RefreshBtnTbpSettings $isAllSettingsValid

    return $isAllSettingsValid

}

function RefreshBtnTbpSettings {
    param (
        [bool]$isAllSettingsValid
    )
    
    if ($isAllSettingsValid) {
        $image     = "settings"
        $backColor = ""
    } else {
        $image     = "settings_erro"
        $backColor = $COLOR_ERROR
    }

    $btn_settings.Image     = InterfaceGetImage $image
    $tbp_settings.BackColor = $backColor

}

function RefreshLstConferencia {
    
    $lastRunResults = ConfigJSON "conferencia"

    $lst_conferencia.Items.Clear()

    if ($lastRunResults) {
        $ano           = $lastRunResults.ano
        $etapa         = $lastRunResults.etapa
        $grupo         = $lastRunResults.grupo
        $nomeGrupo     = $lastRunResults.nome_grupo
        $nItens        = $lastRunResults.n_itens
        $nSolicitacoes = $lastRunResults.n_solicitacoes

        [void]$lst_conferencia.Items.Add("Etapa: $etapa - $ano", "ufsc")
        [void]$lst_conferencia.Items.Add("Grupo: $grupo", "ufsc") 
        [void]$lst_conferencia.Items.Add("Material: $nomeGrupo", "ufsc")
        [void]$lst_conferencia.Items.Add("Itens: $nItens", "ufsc")
        [void]$lst_conferencia.Items.Add("Solicitações: $nSolicitacoes", "ufsc")

        InterfaceCustomProperty $lst_conferencia "Empty" $false
    } else {
        [void]$lst_conferencia.Items.Add("Informe o link, verifique as configurações, escolha a aba desejada e clique em Gerar...", "ufsc")
        InterfaceCustomProperty $lst_conferencia "Empty" $true
    }

}

function GetFiles {

    $files = @{
        lst_importar  = Get-ChildItem -Path $fldToImport\*.csv -Name;
        lst_descricao = Get-ChildItem -Path $fldToImport\Descrição*.xls* -Name;
        lst_prints    = Get-ChildItem -Path $fldResumo\*.pdf -Name;
        lst_resumos   = Get-ChildItem -Path $fldResumo\*.xls* -Name;
        lst_dfd       = Get-ChildItem -Path $fldRelatorio\*.xls* -Name;
        lst_relatorio = Get-ChildItem -Path $fldRelatorio\*.pdf -Name;
    }
    
    return $files
}

function CheckFiles {
    
    $files = GetFiles

    foreach ($lstName in $files.Keys) {
        $lst_ = (Get-Variable -Name $lstName).Value

        if ($null -eq $files[$lstName]) {
            $lst_ | Add-Member -MemberType NoteProperty -Name "Empty" -Value $true -Force
        } else {
            $lst_ | Add-Member -MemberType NoteProperty -Name "Empty" -Value $false -Force
        }
    }

    return $files
}

function ClearLinkStates {

    $config = ConfigJSON -option "all"

    $statusRScript = $config.resultado_geracao
    if ($statusRScript -ne "running") {
        ConfigJSON -key "resultado_geracao" -value "waiting"
    }
    ConfigJSON -key "conferencia" -option "remove"

    RefreshLstConferencia
    RefreshBtnInfo "waiting"
    RefreshBtnLog $statusRScript $config.arquivo_log_R

}

function TestLink {
    
    $url = $txt_link_planilha.Text

    if ([string]::IsNullOrEmpty($url)) {
        return @{
            isValid = $false;
            msg = "Informe o link da aba LISTA FINAL e aguarde";
            backColor = "";
        }
    }
    
    ShowMessage $url "Link informado" 200

    if (-not [Uri]::IsWellFormedUriString($url, 'Absolute')) {
        return @{
            isValid = $false;
            msg = "Link inválido";
            backColor = $COLOR_ERROR;
        }
    }
    
    try {
        $response = Invoke-WebRequest -Method Get -Uri $url -TimeoutSec 5

        $result = TestLinkResponse $response
        return $result
    } catch {
        return @{
            isValid = $false;
            msg = "Erro ao acessar o link informado. Veja erro no console.";
            backColor = $COLOR_ERROR;
            error = $_.Exception.Message;
        }
    }

}

function TestLinkResponse {
    param (
        $response
    )
    
    if ($response.RawContent -match "LISTA FINAL") {
        $grupoMateriais = $response.InputFields.value
        $processosSPA   = (Select-String $REGEX_PROCESSOS_SPA -InputObject $response.RawContent -AllMatches | 
                           ForEach-Object {$_.matches.Value} | Sort-Object -Unique)
        
        return @{
            isValid = $true;
            msg = $grupoMateriais;
            backColor = $COLOR_OK;
            processosSPA = $processosSPA;
        }
    } else {
        return @{
            isValid = $true;
            msg = "Este não parece ser um link de planilha de inserção de demandas";
            backColor = $COLOR_WARNING;
        }
    }

}

function CheckLink {

    if ($txt_link_planilha.currentLink -eq $txt_link_planilha.Text) {
        return
    }

    $lbl_info_grupo.Text = "Aguarde... Acessando o link informado..."

    $lbl_wait.Show()

        ClearLinkStates

        $result = TestLink
        
        RefreshLink $result
    
    $lbl_wait.Hide()

}

function RefreshLink {
    param(
        $result
    ) 

    if ($result.error) {
        ShowMessage $result.error "Erro ao acessar link:" -bgColor "Red" -fgColor "Yellow"
    }    

    $lbl_info_grupo.Text = $result.msg
    $pnl_link.BackColor  = $result.backColor
    
    InterfaceCustomProperty $txt_link_planilha "validLink" $result.isValid
    
    InterfaceCustomProperty $txt_link_planilha "currentLink" $txt_link_planilha.Text

    ConfigJSON -key "link_planilha" -value $txt_link_planilha.Text

    RefreshCmbProcessos $result.processosSPA

    RefreshBtnGerar
}

function RefreshBtnInfo {
    param(
        [string]$lastRunResults
    )

    switch ($lastRunResults) {
        "sucesso" {
            $image = "ok"
            $tag   = "A última geração foi bem sucedida. Verifique os arquivos abaixo."
        }
        "erro" {
            $image = "erro"
            $tag   = "Ocorreram erros na última geração e não foram gerados arquivos. Verifique o log."
        }
        "ambos" {
            $image = "alerta"
            $tag   = "Ocorreram problemas na última geração. Verifique os arquivos gerados e o log."    
        }
        "waiting" {
            $image = "waiting"
            $tag   = "Informe o link, verifique as configurações, escolha a aba desejada e clique em Gerar."
        }
        Default {
            $image = "erro"
            $tag   = "Ocorreu algum erro não identificado. Verifique o último log."
        }
    }

    $btn_info.Image = InterfaceGetImage $image
    $btn_info.Tag   = $tag
}

function RefreshBtnLog {
    param (
        [string]$lastRunResults,
        [string]$fleLogR
    )
    
    if ($null -eq $fleLogR -or $lastRunResults -eq "waiting" -or $lastRunResults -eq "sucesso") {
        $image = "log"
        $tag   = "Abrir pasta de logs"
    } else {
        $image = "log_erro"
        $tag   = "Abrir último log com erros"
    }

    $btn_log.Image = InterfaceGetImage $image 
    $btn_log.Tag   = $tag
}

Function ReadyToRunRScript {

    $RScripts = "importacao.R", "importacao.Rmd" 
    $errors   = New-Object System.Collections.ArrayList

    foreach ($script in $RScripts) {
        if (-not (Test-Path -Path "$fldRoot\$script")) {
            $errors.Add("Não foi localizado o script $script. Verifique a pasta.")
        }
    }

    if ($errors.Count -gt 0) {
        [void](ShowErrors $errors)
        return $false
    }

    if ($lbl_info_grupo.Text -eq "Este não parece ser um link de planilha de inserção de demandas") {
        $grupoMateriais = " informada?"

        $result = [System.Windows.Forms.MessageBox]::Show("Este não parece ser um link de planilha de inserção de demandas.`n`nGostaria de continuar mesmo assim?", "Alerta", 4, 48)
        if ($result -eq "No") {
            [void](ShowErrors "Verifique o link informado antes de continuar.")
            return $false
        }
    } else {
        $grupoMateriais = "`n`n" + $lbl_info_grupo.Text
    }
    
    $selectedTab = $tbc_files.SelectedTab.Name

    if (($selectedTab -eq "tbp_importar" -or $selectedTab -eq "tbp_resumo") -and $cmb_processos.SelectedItem -eq "Nenhum processo disponível") {
        $result = [System.Windows.Forms.MessageBox]::Show("Parece que ainda não foram gerados os processos deste grupo.`n`nDeseja continuar mesmo assim?", "Processos não gerados", 4, 48)
        if ($result -eq "No") {
            [void](ShowErrors "Gere os processos e informe-os na Lista Final antes de tentar novamente.")
            return $false
        }
    }

    if ($selectedTab -eq "tbp_resumo" -and $lst_prints.Empty) {
        [void](ShowErrors "Não é possível gerar o resumo.`nSalve os PDFs dos pedidos gerados na pasta e tente novamente.")
        return $false
    } 

    switch ($selectedTab) {
        "tbp_importar" { 
            $lst_        = $lst_importar
            $msgNotEmpty = "Já há arquivos para importação na pasta.`n`nGostaria de gerar novos arquivos?"
            $msgError    = "Apague os arquivos .csv existentes antes de continuar."
            $msgRun      = "os arquivos para importar a lista"
        }
        "tbp_resumo" {
            $lst_        = $lst_resumos
            $msgNotEmpty = "Já há um ou mais resumos gerados.`n`nGostaria de gerar novo resumo?"
            $msgError    = "Apague os resumos existentes antes de continuar."
            $msgRun      = "o resumo dos pedidos da lista"
        }
        "tbp_relatorio" {
            $lst_        = $lst_relatorio
            $msgNotEmpty = "Já há um ou mais relatórios gerados.`n`nGostaria de gerar novo relatório?"
            $msgError    = "Apague os relatórios existentes antes de continuar."
            $msgRun      = "o relatório gerencial da lista" 
        }
    }

    if (-not $lst_.Empty) {
        $result = [System.Windows.Forms.MessageBox]::Show($msgNotEmpty, "Alerta", 4, 48)
        if ($result -eq "No") {
            [void](ShowErrors $msgError)
            return $false   
        }
    }

    $result = [System.Windows.Forms.MessageBox]::Show("Deseja gerar $msgRun$grupoMateriais", "Confirmação", 4, 48)
    if ($result -eq "Yes") {
        return $true
    } else {
        return $false
    }
}

function ShowErrors {
    param (
        [string]$errors
    )
    
    $msgError = $errors -join "`n`n"
    [System.Windows.Forms.MessageBox]::Show($msgError, "Erro", 0, 16)

}

function DeleteFiles {
    param (
        [System.Windows.Forms.ToolStripMenuItem]$ctr_
    )

    $lst_ = $ctr_.GetCurrentParent().SourceControl
   
    if ($lst_.Empty) {
        return
    }

    switch ($lst_) {
        $lst_importar {
            $msg = "todos os arquivos a importar (.csv)"
            $fleToDelete = "$fldToImport\*.csv"
        }
        $lst_descricao {
            $msg = "todos os arquivos com descrição a ajustar"
            $fleToDelete = "$fldToImport\Descrição*.xlsx"
        }
        $lst_prints {
            $msg = "todas as listas dos pedidos gerados (.pdf)"
            $fleToDelete = "$fldResumo\*.pdf"
        }
        $lst_resumos {
            $msg = "todos os resumos gerados (.xls)"
            $fleToDelete = "$fldResumo\*.xls*"
        }
        $lst_dfd {
            $msg = "todos os DFDs gerados (.xls)"
            $fleToDelete = "$fldRelatorio\*.xls*"
        }
        $lst_relatorio {
            $msg = "todos os relatórios gerados (.pdf)"
            $fleToDelete = "$fldRelatorio\*.pdf"
        }
    }

    $result = [System.Windows.Forms.MessageBox]::Show("Deseja excluir ${msg}?", "Confirmação de exclusão", 4, 48)
    if ($result -eq "No") {
        return
    }

    try{
        Get-ChildItem -Path $fleToDelete -File | Remove-Item -Force -ErrorAction Stop
    } catch {
        $msgError = "Não foi possível excluir os arquivos.`n`nVerifique se não estão sendo utilizados."
        [System.Windows.Forms.MessageBox]::Show($msgError, "Erro", 0, 16)
    }

    RefreshLstFiles
}

function RefreshLstFiles {

    $files = CheckFiles

    $lists = @{
        "lst_importar" = "csv";
        "lst_descricao" = "excel_erro";
        "lst_prints" = "pdf";
        "lst_resumos" = "excel";
        "lst_dfd" = "excel";
        "lst_relatorio" = "pdf";
    }

    foreach ($lstName in $lists.Keys) {
        $lst_ = (Get-Variable -Name $lstName).Value
        $lst_.Items.Clear()

        $icon = $lists[$lstName]

        if ($lst_.Empty) {
            [void]$lst_.Items.Add("Não há arquivos gerados", $icon)
            continue
        }

        foreach ($file in $files[$lstName]) {
            [void]$lst_.Items.Add($file, $icon)
        }
    }

}

Function ExecuteRScript {

    if (IsUserRunningR) {
        return
    }

    RefreshLstFiles

    if (-not (ReadyToRunRScript)) {
        return
    }

    $btn_gerar.Text    = "Aguarde..."
    $btn_gerar.Enabled = $false
    $btn_exit.Enabled  = $false
    $lbl_wait.Show()

    InterfaceMinimize -HideTaskBar

    if ($cmb_processos.SelectedItem -eq "Todos - Relatório consolidado" -or $cmb_processos.SelectedItem -eq "Nenhum processo disponível") {
        $processoSPA = "todos"
    } else {
        $processoSPA = $cmb_processos.SelectedItem
    }
    ConfigJSON -key "processo" -value $processoSPA

    switch ($tbc_files.SelectedTab.Name) {
        "tbp_importar" {
            $RScript = "gerar"
            $msgPostRun = "Script finalizado com sucesso.`n`nConfira os arquivos .csv gerados."
        }
        "tbp_resumo" {
            $RScript = "resumo"
            $msgPostRun = "Script finalizado com sucesso.`n`nConfira a planilha com o resumo dos pedidos."
        }
        "tbp_relatorio" {
            $RScript = "relatorio"
            $msgPostRun = "Script finalizado com sucesso.`n`nConfira a relatório gerado."
        }
    }
    ConfigJSON -key "script_a_executar" -value $RScript

    ConfigJSON -show

    RunR -script "importacao.R" -arguments $RScript

    InterfaceRestore

    $btn_gerar.Text    = "Gerar"
    $btn_gerar.Enabled = $true
    $btn_exit.Enabled  = $true
    $lbl_wait.Hide()

    ConfigJSON -show

    CheckLastRunResults $msgPostRun
}

function CheckLastRunResults {
    param(
        [string]$msgPostRun
    )

    $config = ConfigJSON -option "all"

    $lastRunResults = $config.resultado_geracao 
    $fleLogR        = $config.arquivo_log_R
    $errors         = $config.msg_erro

    switch ($lastRunResults) 
    {
        "sucesso" {
            $title      = "Sucesso"
            $type       = 64
            $tbpToFocus = "tbp_status"
        }
        "erro" {
            $msgPostRun = "Não foram gerados arquivos pois foram identificados os seguintes erros:`n=======", ($errors -join "`n") -join "`n"
            $title      = "Erro"
            $type       = 16
            $tbpToFocus = "tbp_erros"
        }
        "ambos" {
            $msgPostRun = "Os arquivos foram gerados, MAS foram identificados os seguintes erros:`n=======", ($errors -join "`n") -join "`n"
            $title      = "Alerta"
            $type       = 48
            $tbpToFocus = "tbp_erros"
        }
        Default {
            $msgPostRun = "O script retornou algum erro não identificado.`n`nVerifique os logs."
            $title      = "Erro desconhecido"
            $type       = 16
            $tbpToFocus = "tbp_erros"
        }
    }

    [System.Windows.Forms.MessageBox]::Show($this, $msgPostRun, $title, 0, $type)

    RefreshInterfaceAfterRun $lastRunResults $errors $fleLogR $tbpToFocus
}

function RefreshInterfaceAfterRun {
    param(
        [string]$lastRunResults,
        $errors,
        [string]$fleLogR,
        [string]$tbpToFocus
    )

    $tbc_info.SelectedTab = $tbc_info.TabPages[$tbpToFocus]

    RefreshLstFiles

    RefreshLstConferencia
    RefreshBtnInfo $lastRunResults

    RefreshLstErrors $errors
    RefreshBtnLog $lastRunResults $fleLogR
}

function RefreshLstErrors {
    param (
        $errors = $null
    )
    
    if ($null -eq $errors) {
        $errors = ConfigJSON "msg_erro"
    }

    if ($errors -is [array]) {
        $lst_erros.Text = $errors -join "`r`n"
    } else {
        $lst_erros.Text = $errors
    }

}

function ContextMenuPnlLink {
    param(
        [System.Object]$mnuContext
    )

    $mnu_ = $mnuContext.Text
    $ctr_ = $mnuContext.GetCurrentParent().SourceControl
   
    switch ($mnu_) {
        "Abrir link" {
            if ($txt_link_planilha.validLink) { 
                InterfaceMinimize
                Start-Process $txt_link_planilha.Text 
            }
        }
        "Copiar" {
            Set-Clipboard -Value $ctr_.Text
        }
        "Colar" {
            $ctr_.Text = Get-Clipboard -Raw
            CheckLink
        }
        "Recortar" {
            Set-Clipboard -Value $ctr_.Text
            $ctr_.Clear()
            CheckLink
        }
        "Limpar" {
            $ctr_.Clear()
            CheckLink
        }
    }
}

function ContextMenuLstSettings {
    param(
        [System.Object]$mnuContext
    )

    $mnu_    = $mnuContext.Text
    $setting = $mnuContext.GetCurrentParent().SourceControl

    switch ($mnu_) {
        "Limpar" {
            $setting.CurrentRow.Cells["Valor"].Value = ""
            RefreshBtnGerar
        }
        "Copiar" {
            Set-Clipboard -Value $setting.CurrentRow.Cells["Valor"].Value
        }
        "Colar" {
            $setting.CurrentRow.Cells["Valor"].Value = Get-Clipboard -Raw
            RefreshBtnGerar
        }
        "Recortar" {
            Set-Clipboard -Value $setting.CurrentRow.Cells["Valor"].Value
            $setting.CurrentRow.Cells["Valor"].Value = ""
            RefreshBtnGerar
        }
    }
}

function ContextMenuPnlInfo {
    param(
        [System.Object]$mnuContext
    )

    $mnu_ = $mnuContext.Text
    $ctr_ = $mnuContext.GetCurrentParent().SourceControl
   
    switch ($mnu_) {
        "Limpar" {
            switch ($ctr_) {
                $lst_conferencia {
                    if ($lst_conferencia.Empty) {
                        return
                    }
                    $msg = "da última importação"
                }
                $lst_erros {
                    if ($lst_erros.Text -eq "") {
                        return
                    }
                    $msg = "de erros da última importação"
                }
            }
        
            $result = [System.Windows.Forms.MessageBox]::Show("Deseja limpar as informações ${msg}?", "Confirmação de exclusão", 4, 48)
            
            if ($result -eq "No") {
                return
            }

            switch ($ctr_) {
                $lst_conferencia {
                    ConfigJSON -key "conferencia" -option "remove"
                    ConfigJSON -key "resultado_geracao" -value "waiting"
                    RefreshLstConferencia
                    RefreshBtnInfo "waiting"
                }
                $lst_erros {
                    ConfigJSON -key "arquivo_log_R" -value $null
                    ConfigJSON -key "msg_erro" -option "remove"
                    $lst_erros.Clear()
                    RefreshBtnLog $null $null
                }
            }
        }
        "Copiar" {
            if ($ctr_ -is [System.Windows.Forms.ListView]) {
                $text = ($ctr_.Items | ForEach-Object { $_.Text }) -join "`r`n"
            } elseif ($ctr_ -is [System.Windows.Forms.TextBox]) {
                $text = $ctr_.Text
            }
            Set-Clipboard -Value $text
         }
    }
}

function RefreshLstSettings {
   
    $config = ConfigJSON -option "all"

    $settings = @(
        @{
            name = "Valor mínimo"; 
            key = "valor_minimo"; 
            value = $config.valor_minimo; 
            tip = "Valor mínimo total do item para inclusão no processo"
        },
        @{
            name = "Quantidade mínima";
            key = "qtde_minima";
            value = $config.qtde_minima;
            tip = "Quantidade mínima demandada do item para inclusão no processo"
        },
        @{
            name = "Célula inicial (quantitativos)";
            key = "celula";
            value = $config.celula;
            tip = "Célula da LISTA FINAL onde inicia a área de inserção dos quantitativos pelas Unidades (ex: Q6)"
        },
        @{
            name = "Quantidade de Unidades requerentes";
            key = "unidades";
            value = $config.unidades;
            tip = "Quantidade de colunas de Unidades requerentes (incluindo ocultas e UFSC GERAL)"
        },
        @{
            name = "Aba Menu";
            key = "aba_menu";
            value = $config.aba_menu;
            tip = "Nome da aba Menu na planilha (apenas letras, números e espaços)"
        },
        @{
            name = "Aba LISTA FINAL";
            key = "aba_lista_final";
            value = $config.aba_lista_final;
            tip = "Nome da aba LISTA FINAL na planilha (apenas letras, números e espaços)"
        }
    )

    $lst_settings.SuspendLayout()

        foreach ($setting in $settings) {
            $row = New-Object System.Windows.Forms.DataGridViewRow
            $row.CreateCells($lst_settings)
            $row.Cells[0].Value = $lst_image.Images["config"]
            $row.Cells[1].Value = $setting.name
            $row.Cells[1].ToolTipText = $setting.tip
            $row.Cells[2].Value = $setting.value
            $row.Tag = $setting.key
            [void]$lst_settings.Rows.Add($row)
        }

    $lst_settings.ResumeLayout()

    RefreshBtnGerar

}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$fldRoot      = $PSScriptRoot
$fldParent    = Join-Path -Path $PSScriptRoot -ChildPath .. -Resolve
$fldCommon    = "$fldParent\_common"
$fldToImport  = "$fldRoot\ARQUIVOS A IMPORTAR"
$fldRelatorio = "$fldRoot\RELATORIOS"
$fldResumo    = "$fldRoot\RESUMO PEDIDOS" 

$configMain = "config.psm1"
if (Test-Path -Path "$fldCommon\$configMain") {
    Unblock-File -Path "$fldCommon\$configMain"
    Import-Module "$fldCommon\$configMain"
} else {
    [System.Windows.Forms.MessageBox]::Show("Não foi localizado o arquivo '${configMain}'.`n`nNão é possível executar o script.", "Erro", 0, 16)
    [Environment]::Exit(1)
}

main "IMPORTACAO"

SetAppConstants @{
    # Regex Patterns
    REGEX_PROCESSOS_SPA = "23080\.\d{6}/\d{4}-\d{2}"
    
    COLOR_OK      = "LightGreen"
    COLOR_WARNING = "LightGoldenrodYellow"
    COLOR_ERROR   = "LightSalmon"

    # Tamanhos de UI
    PADDING_INNER = 10
    PADDING_OUTER = 15

}

CheckFolder @(
    $fldToImport,
    $fldRelatorio,
    $fldResumo
)    

# MAIN FORM

$frm_main, $pic_banner, $tip_ = InterfaceMainForm "Importação Planilhas - Google Drive/Solar" 900 620 "planilha"

$lst_image = InterfaceImageList

$lbl_wait = InterfaceSplashScreen -onlyLabel
$lbl_wait.Hide()

# CONTEXT MENUS

$items = @( 
    @("Abrir link", "google", {ContextMenuPnlLink $this}), 
    @("-"),
    @("Copiar", "copiar", {ContextMenuPnlLink $this}),
    @("Colar", "novo", {ContextMenuPnlLink $this}),
    @("Recortar", "recortar", {ContextMenuPnlLink $this}),
    @("-"),
    @("Limpar", "excluir", {ContextMenuPnlLink $this})
)
$mnu_context_link = InterfaceContextMenu $items

$items = @( 
    @("Abrir arquivo", "mover", {OpenFileFolder $this}), 
    @("Abrir pasta", "folder_bw", {OpenFileFolder $this}),
    @("Excluir arquivos", "excluir", {DeleteFiles $this})
)
$mnu_context_files = InterfaceContextMenu $items

$items = @( 
    @("Copiar", "copiar", {ContextMenuLstSettings $this}),
    @("Colar", "novo", {ContextMenuLstSettings $this}),
    @("Recortar", "recortar", {ContextMenuLstSettings $this}),
    @("-"),
    @("Limpar", "excluir", {ContextMenuLstSettings $this})
)
$mnu_context_settings = InterfaceContextMenu $items

$items = @( 
    @("Copiar", "copiar", {ContextMenuPnlInfo $this}),
    @("-"),
    @("Limpar", "excluir", {ContextMenuPnlInfo $this})
)
$mnu_context_info = InterfaceContextMenu $items

# PANEL LINK

$margin = @{
    top = 15;
    bottom = $pic_banner.Height + 15;
    left = 130;
    right = 110;
}    

$utilSpace = @{
    width  = $frm_main.Width - $margin.left - $margin.right;
    height = $frm_main.Height - $margin.top - 75 - $margin.bottom - ($PADDING_OUTER * 2)
}

$params = @{
    width = $utilSpace.width;
    height = 75;
    top = $margin.top;
    left = $margin.left;
}
$pnl_link = InterfacePanel @params

$params = @{
    size = 30;
    top = ($pnl_link.Height - 30) / 2;
    left = $PADDING_INNER;
    name = "drive";
    tag = "Abrir LISTA FINAL";
    function = {
        CheckLink
        if ($txt_link_planilha.validLink) { 
            InterfaceMinimize
            Start-Process $txt_link_planilha.Text 
        }
    };
}
$btn_google = InterfaceButtonImage @params

$params = @{
    type = "Textbox";
    name = "txt_link_planilha";
    labelText = "Link da aba LISTA FINAL";
    tag = "Link da aba LISTA FINAL";
    top = 25;
    left = $btn_google.Width + ($PADDING_INNER * 2);
    width = $pnl_link.Width - 50 - 35;
    mnuContext = $mnu_context_link;
    events = @{
        Click = {$this.SelectAll()};
        GotFocus = {$this.SelectAll()};
        LostFocus = {
            CheckLink
            ConfigJSON "link_planilha" $this.Text
        };
        KeyUp = {
            param($sender, $e)
            if ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::V) {
                CheckLink
                ConfigJSON "link_planilha" $this.Text
            }
            elseif ($e.KeyCode -eq [System.Windows.Forms.Keys]::Delete) {
                CheckLink
                ConfigJSON "link_planilha" $this.Text
            }
        }
    }
}
$txt_link_planilha, $lbl_link_planilha = InterfaceControl @params

$params = @{
    size = 24;
    top = $txt_link_planilha.Top + (($txt_link_planilha.Height - 24) / 2);
    left = InterfacePosition $txt_link_planilha "right" 5;
    name = "atualizar";
    tag = "Clique para atualizar o link";
    function = {
        InterfaceCustomProperty $txt_link_planilha "currentLink" $null
        CheckLink
    };
}
$btn_atualizar_link = InterfaceButtonImage @params

$params = @{
    labelText = "";
    width = $txt_link_planilha.Width;
    height = 22;
    top = $pnl_link.Height - 22;
    left = $txt_link_planilha.Left;
}
$lbl_info_grupo = InterfaceLabel @params

$controls = @(
    $btn_google,
    $lbl_link_planilha,
    $txt_link_planilha,
    $lbl_info_grupo,
    $btn_atualizar_link
)
$pnl_link.Controls.AddRange($controls)
$frm_main.Controls.AddRange(@($pnl_link))

# TABCONTROL INFO

$params = @{
    type = "control";
    width = $pnl_link.Width;
    height = $utilSpace.height * 0.45;
    top = InterfacePosition $pnl_link "bottom" $PADDING_OUTER;
    left = $pnl_link.Left;
    properties = @{
        drawMode = [System.Windows.Forms.TabDrawMode]::OwnerDrawFixed;
    };
    events = @{
        "DrawItem" = {
            param($sender, $e)
    
            $brush = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(173, 216, 230))
            $font = [System.Drawing.Font]::new($e.Font, [System.Drawing.FontStyle]::Bold)
            
            if ($e.Index -eq $sender.SelectedIndex) {
                $e.Graphics.FillRectangle($brush, $e.Bounds)
                $e.Graphics.DrawString($sender.TabPages[$e.Index].Text, $font, 
                                    [System.Drawing.Brushes]::Black, 
                                    [System.Drawing.PointF]::new($e.Bounds.X + 3, $e.Bounds.Y + 3))
            } else {
                $e.Graphics.FillRectangle([System.Drawing.Brushes]::Transparent, $e.Bounds)
                $e.Graphics.DrawString($sender.TabPages[$e.Index].Text, $e.Font, 
                                    [System.Drawing.Brushes]::Black, 
                                    [System.Drawing.PointF]::new($e.Bounds.X + 2, $e.Bounds.Y + 2))
            }
            
            $brush.Dispose()
            $font.Dispose()
        }
    }
}
$tbc_info = InterfaceTabs @params

$params = @{
    type = "pages";
    names = @("tbp_status", "tbp_settings", "tbp_erros");
    texts = @("Status", "Configurações", "Log e erros");
}
$tbp_status, $tbp_settings, $tbp_erros = InterfaceTabs @params

$frm_main.Controls.Add($tbc_info) 
$tbc_info.Controls.AddRange(@($tbp_status, $tbp_settings, $tbp_erros))

# TAB STATUS

$params = @{
    size = 30;
    top = $PADDING_OUTER;
    left = $PADDING_INNER;
    name = "erro";
    tag = "..."
} 
$btn_info = InterfaceButtonImage @params

$params = @{
    name = "lst_conferencia"; 
    width = $tbc_info.Width - ($btn_info.Width + ($PADDING_INNER * 2)) - $PADDING_OUTER;
    height = $tbc_info.Height - ($PADDING_OUTER * 3);
    top = $PADDING_INNER;
    left = $btn_info.Width + ($PADDING_INNER * 2);
    mnuContext = $mnu_context_info;
    properties = @{
        HideSelection = $true;
        Scrollable = $false
    };
    columns = [ordered]@{"Ultima Importação" = -2}
}
$lst_conferencia = InterfaceList @params

$tbp_status.Controls.AddRange(@($btn_info, $lst_conferencia))

# TAB CONFIGURATIONS

$params = @{
    size = $btn_info.Width;
    top = $btn_info.Top;
    left = $btn_info.Left;
    name = "settings"
}
$btn_settings = InterfaceButtonImage @params

$params = @{
    type = "DataGridView";
    name = "lst_settings";
    top = $lst_conferencia.Top;
    left = $lst_conferencia.Left;
    width = $lst_conferencia.Width;
    height = $lst_conferencia.Height;
    mnuContext = $mnu_context_settings;
    properties = @{
        BorderStyle                 = 'FixedSingle';
        AllowUserToAddRows          = $false;
        AllowUserToDeleteRows       = $false;
        AllowUserToResizeRows       = $false;
        AllowUserToResizeColumns    = $false;
        AllowUserToOrderColumns     = $false;
        ReadOnly                    = $false;
        RowHeadersVisible           = $false;
        EditMode                    = 'EditOnEnter';
        SelectionMode               = "FullRowSelect";
        MultiSelect                 = $false;
        ColumnHeadersHeightSizeMode = "DisableResizing";
        ColumnHeadersHeight         = 20;
        CellBorderStyle             = 'None';
        ColumnHeadersBorderStyle    = [System.Windows.Forms.DataGridViewHeaderBorderStyle]::Single;
    };
    events = @{
        "CellEndEdit" = {
            $row = $lst_settings.CurrentRow
            $key = $row.Tag
            $value = $row.Cells[2].Value  

            ConfigJSON -key $key -value $value
            RefreshBtnGerar
        };
        "MouseDown" = {
            param($sender, $e)

            if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Right) {
                $hitTestInfo = $lst_settings.HitTest($e.X, $e.Y)

                if ($hitTestInfo.RowIndex -ge 0) {
                    $lst_settings.ClearSelection()
                    $lst_settings.Rows[$hitTestInfo.RowIndex].Selected = $true
                    $lst_settings.CurrentCell = $lst_settings.Rows[$hitTestInfo.RowIndex].Cells[0]
                }
            }
        }
    }
}

$lst_settings = InterfaceList @params
$lst_settings.RowTemplate.Height = 15
$lst_settings.ColumnHeadersDefaultCellStyle.SelectionBackColor = $frm_main.BackColor

$coluna_imagem             = New-Object System.Windows.Forms.DataGridViewImageColumn
$coluna_imagem.HeaderText  = ""
$coluna_imagem.Name        = "Icone"
$coluna_imagem.Width       = 30
$coluna_imagem.ImageLayout = [System.Windows.Forms.DataGridViewImageCellLayout]::Normal

$coluna_nome            = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$coluna_nome.HeaderText = "Configuração"
$coluna_nome.Name       = "Configuracao"
$coluna_nome.ReadOnly   = $true
$coluna_nome.Width      = 250

$coluna_valor            = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$coluna_valor.HeaderText = "Valor"
$coluna_valor.Name       = "Valor"
$coluna_valor.ReadOnly   = $false
$coluna_valor.Width      = $lst_settings.Width - $coluna_imagem.Width - $coluna_nome.Width - 20

$lst_settings.Columns.AddRange($coluna_imagem, $coluna_nome, $coluna_valor)
$tbp_settings.Controls.AddRange(@($btn_settings, $lst_settings))

# TAB ERRORS

$params = @{
    size = $btn_info.Width;
    top = $btn_info.Top;
    left = $btn_info.Left;
    name = "erro"
}
$btn_erro = InterfaceButtonImage @params

$params = @{
    size = $btn_info.Width;
    top = InterfacePosition $lst_conferencia "bottom" (-$PADDING_OUTER * 2);
    left = $btn_info.Left;
    name = "log";
    tag = "...";
    function = {
        InterfaceMinimize
        $fleLogR = ConfigJSON "arquivo_log_R"
        try { 
            Start-Process "$fldLog\$fleLogR" 
        } 
        catch { 
            Start-Process "$fldLog"
        }
    }
}
$btn_log = InterfaceButtonImage @params

$params = @{
    type = "TextBox";
    name = "lst_erros";
    top = $lst_conferencia.Top;
    left = $lst_conferencia.Left;
    width = $lst_conferencia.Width;
    height = $lst_conferencia.Height;
    mnuContext = $mnu_context_info;
    properties = @{
        Text = $config.msg_erro;
        Multiline = $true;
        ReadOnly = $true;
        ScrollBars = "Vertical";
        BackColor = "White"
    }
}
$lst_erros = InterfaceControl @params

$tbp_erros.Controls.AddRange(@($btn_erro, $btn_log, $lst_erros))

# TABCONTROL FILES

$params = @{
    type = "control";
    width = $pnl_link.Width;
    height = $utilSpace.height * 0.55;
    top = InterfacePosition $tbc_info "bottom" $PADDING_OUTER;
    left = $pnl_link.Left;
    properties = @{
        DrawMode = [System.Windows.Forms.TabDrawMode]::OwnerDrawFixed
    };
    events = @{
        "DrawItem" = {
            param($sender, $e)
            
            $brush = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(173, 216, 230))
            $font = [System.Drawing.Font]::new($e.Font, [System.Drawing.FontStyle]::Bold)
            
            if ($e.Index -eq $sender.SelectedIndex) {
                $e.Graphics.FillRectangle($brush, $e.Bounds)
                $e.Graphics.DrawString($sender.TabPages[$e.Index].Text, $font, 
                                    [System.Drawing.Brushes]::Black, 
                                    [System.Drawing.PointF]::new($e.Bounds.X + 3, $e.Bounds.Y + 3))
            } else {
                $e.Graphics.FillRectangle([System.Drawing.Brushes]::Transparent, $e.Bounds)
                $e.Graphics.DrawString($sender.TabPages[$e.Index].Text, $e.Font, 
                                    [System.Drawing.Brushes]::Black, 
                                    [System.Drawing.PointF]::new($e.Bounds.X + 2, $e.Bounds.Y + 2))
            }
            
            $brush.Dispose()
            $font.Dispose()
        }
    }
}
$tbc_files = InterfaceTabs @params

$params = @{
    type = "pages";
    names = @("tbp_importar", "tbp_resumo", "tbp_relatorio");
    texts = @("Arquivos para importação", "Resumo dos pedidos", "Relatório gerencial");
}
$tbp_importar, $tbp_resumo, $tbp_relatorio = InterfaceTabs @params

$frm_main.Controls.Add($tbc_files) 
$tbc_files.Controls.AddRange(@($tbp_importar, $tbp_resumo, $tbp_relatorio))

# TAB IMPORTAR ARQUIVOS

$params = @{
    name = "lst_importar"; 
    width = ($tbc_files.Width / 2) - ($PADDING_INNER * 2);
    height = $tbc_files.Height - ($PADDING_OUTER * 3);
    top = $PADDING_INNER;
    left = $PADDING_INNER;
    mnuContext = $mnu_context_files;
    columns = [ordered]@{
        "Arquivos para importar" = -2
    };
    events = @{
        "ItemActivate" = {OpenFileFolder $lst_importar}
    }
}
$lst_importar = InterfaceList @params

$params = @{
    name = "lst_descricao"; 
    width = $lst_importar.Width;
    height = $lst_importar.Height;
    top = 10;
    left = InterfacePosition $lst_importar "right" $PADDING_OUTER;  
    mnuContext = $mnu_context_files;
    columns = [ordered]@{
        "Descrição de itens a ajustar" = -2
    };
    events = @{
        "ItemActivate" = {OpenFileFolder $lst_descricao}
    }
}
$lst_descricao = InterfaceList @params

$tbp_importar.Controls.AddRange(@($lst_importar, $lst_descricao))

# TAB RESUMO PEDIDOS

$params = @{
    name = "lst_prints"; 
    width = $lst_importar.Width;
    height = $lst_importar.Height;
    top = $lst_importar.Top;
    left = $lst_importar.Left;
    mnuContext = $mnu_context_files;
    columns = [ordered]@{
        "Prints da tela dos pedidos" = -2
    };
    events = @{
        "ItemActivate" = {OpenFileFolder $lst_prints}
    }
}
$lst_prints = InterfaceList @params

$params = @{
    name = "lst_resumos"; 
    width = $lst_descricao.Width;
    height = $lst_descricao.Height;
    top = $lst_descricao.Top;
    left = $lst_descricao.Left;  
    mnuContext = $mnu_context_files;
    columns = [ordered]@{
        "Resumos gerados" = -2
    };
    events = @{
        "ItemActivate" = {OpenFileFolder $lst_resumos}
    }
}
$lst_resumos = InterfaceList @params

$tbp_resumo.Controls.AddRange(@($lst_prints, $lst_resumos))

# TAB RELATÓRIOS GERENCIAIS

$params = @{
    type = "ComboBox";
    labelText = "Selecione o processo para o relatório";
    top = $PADDING_OUTER * 2;
    left = $lst_importar.Left;
    width = $lst_importar.Width - 70;
    tag = "Selecione um processo da lista.";
}
$cmb_processos, $lbl_cmb_processos = InterfaceControl @params

$params = @{
    name = "lst_dfd"; 
    width = $lst_importar.Width;
    height = $lst_importar.Height - ($cmb_processos.Top + $cmb_processos.Height + $PADDING_INNER);
    top = InterfacePosition $cmb_processos "bottom" ($PADDING_OUTER + 4);
    left = $lst_importar.Left;  
    mnuContext = $mnu_context_files;
    columns = [ordered]@{
        "Planilha auxiliar para DFD" = -2
    };
    events = @{
        "ItemActivate" = {OpenFileFolder $lst_dfd}
    }
}
$lst_dfd = InterfaceList @params

$params = @{
    name = "lst_relatorio"; 
    width = $lst_descricao.Width;
    height = $lst_descricao.Height;
    top = $lst_descricao.Top; 
    left = $lst_descricao.Left;  
    mnuContext = $mnu_context_files;
    columns = [ordered]@{
        "Relatórios gerenciais" = -2
    };
    events = @{
        "ItemActivate" = {OpenFileFolder $lst_relatorio}
    }
}
$lst_relatorio = InterfaceList @params

$tbp_relatorio.Controls.AddRange(@($lbl_cmb_processos, $cmb_processos, $lst_dfd, $lst_relatorio))

# FORM ICONS AND IMAGES

$params = @{
    size = 100;
    height = 50;
    top = InterfacePosition $pic_banner "top" $PADDING_OUTER 50;
    left = ($margin.left / 2) - (100 / 2);
    name = "dcom.tif";
    tag = "Ir para a página do DCOM";
    function = {
        InterfaceMinimize
        Start-Process "https://dcom.ufsc.br"
    }
}
$btn_dcom = InterfaceButtonImage @params

$params = @{
    size = 70;
    top = InterfacePosition $btn_dcom "top" $PADDING_OUTER -targetHeight 70;
    left = InterfacePosition $btn_dcom "centerHorizontal" -targetWidth 70;
    name = "ufsc";
    tag = "Ir para a página da UFSC";
    function = {
        InterfaceMinimize
        Start-Process "https://www.ufsc.br"
    }
}
$btn_ufsc = InterfaceButtonImage @params

$params = @{
    size = 50;
    top = $margin.top;
    left = InterfacePosition $btn_dcom "centerHorizontal" -targetWidth 50;
    name = "r";
    tag = "O programa R não está instalado.`r`nClique aqui para ir para a página de download.";
    function = {
        InterfaceMinimize
        Start-Process "https://cran.r-project.org/bin/windows/base/"
    }
}
$btn_r = InterfaceButtonImage @params

$params = @{
    size = 70;
    top = $margin.top;
    left = $frm_main.Width - [math]::Round($margin.right / 2) - (70 / 2);
    name = "calendario";
    tag = "Ir para a página do Calendário de Compras";
    function = {
        InterfaceMinimize
        Start-Process "http://compras.ufsc.br/calendario-de-compras"
    }
}
$btn_calendario = InterfaceButtonImage @params

$params = @{
    size = $btn_calendario.Width;
    top = InterfacePosition $tbc_info "bottom" (-$PADDING_INNER) $btn_calendario.Height;
    left = $btn_calendario.Left;
    name = "Logo-manual-icone";
    tag = "Ir para o Manual do DCOM";
    function = {
        InterfaceMinimize
        Start-Process "https://dcom.wiki.ufsc.br/index.php/METODOLOGIA_DAS_LISTAS#Importa%C3%A7%C3%A3o_das_listas"
    }
}
$btn_manual = InterfaceButtonImage @params

$params = @{
    text = "Atualizar";
    tag = "Atualizar arquivos";
    top = $frm_main.Height - 85;
    left = $margin.left;
    function = {
        RefreshLstFiles
    }
}
$btn_refresh = InterfaceButton @params

$params = @{
    text = "Sair";
    tag = "Sair do script";
    top = $btn_refresh.Top;
    left = InterfacePosition $tbc_files "right" -90;
    function = {
        InterfaceClose
    }
}
$btn_exit = InterfaceButton @params

$params = @{
    text = "Gerar";
    tag = "Gerar os arquivos conforme a aba selecionada";
    top = $btn_refresh.Top;
    left = InterfacePosition $btn_exit "left" $PADDING_INNER;
    function = {
        ExecuteRScript
    }
}
$btn_gerar = InterfaceButton @params

$controls = @(
    $btn_ufsc,
    $btn_dcom,
    $btn_r,
    $btn_calendario,
    $btn_manual,
    $lbl_wait,
    $btn_refresh,
    $btn_gerar,
    $btn_exit,
    $pic_banner
    )
$frm_main.Controls.AddRange($controls)

$lbl_wait.BringToFront()

# SHOW FORM

# SPLASH SCREEN

$frm_splash, $lbl_splash = InterfaceSplashScreen
$frm_splash.Controls.AddRange(@($lbl_splash))

InterfaceShowForm -title "IMPORTAÇÃO GOOGLE DRIVE/SOLAR - MANTER ABERTA" -start {

    [void]$frm_splash.Show()
    
    $txt_link_planilha.Text = ConfigJSON "link_planilha"
    CheckLink
    
    RefreshLstFiles
    RefreshLstErrors
    RefreshLstSettings
    
    ConfigJSON -show

    [void]$frm_splash.Close()

} -close {

    ConfigJSON -show

}
