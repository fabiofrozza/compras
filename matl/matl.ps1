# FUNCTIONS

function Abrir {
    $fleSelected = $lst_atas.SelectedItems[0].Text

    if (([string]::IsNullOrEmpty($fleSelected))) {
        return
    }

    $fldPregao = $lst_atas.Columns[0].Text.substring(14)

    InterfaceMinimize

    Start-Process "${fldData}\$fldPregao\$fleSelected"
}

function RefreshAtas {
    try {
        $pregao = $lst_resumo.Items[0].SubItems[1].Text
    }
    catch {
        $lst_atas.Items.Clear()
        $lst_atas.Columns[0].Text = "Atas"
        return
    }

    $fleToImport = "${fldImport}\PE_${pregao}.csv"
    $fleErrors   = "${fldImport}\PE_${pregao}_ERRO.txt"
    $atas        = Get-ChildItem -Path "${fldData}\$pregao" -File -Filter "*.xls*" -ErrorAction Stop | Where-Object { -not $_.Name.StartsWith('~') } | Select-Object -ExpandProperty Name | Sort-Object
    
    $lst_atas.SuspendLayout()
        $lst_atas.Items.Clear()
    
        $lst_atas.Columns[0].Text = "Atas" + $(if (-not [string]::IsNullOrEmpty($pregao)) {" - Pregão ${pregao}"})
    
        foreach ($ata in $atas) {
            $imageKey = "excel"
            $hasError = if (Test-Path $fleToImport) {"Não"} else {"-"}

            if (Test-Path ${fleErrors}) {
                $contentErrors += Get-Content -Path ${fleErrors} -Raw
                if ($contentErrors -match [regex]::Escape($ata)) {
                    $hasError = "Sim"
                    $imageKey = "excel_erro"
                } 
            } 

            $item = $lst_atas.Items.Add($ata, $imageKey)
            [void]$item.SubItems.Add($hasError)
        }
    $lst_atas.ResumeLayout()
}

function RefreshImportar {
    $fleToImport = Get-ChildItem -Path "${fldImport}" -File -Filter "*.csv" | Select-Object -ExpandProperty Name | Sort-Object
    $fleErrors   = Get-ChildItem -Path "${fldImport}" -File -Filter "*ERRO.txt" | Select-Object -ExpandProperty Name | Sort-Object

    $lst_importar.SuspendLayout()
        $lst_importar.Items.Clear()

        foreach ($fle in $fleToImport) {
            $hasError, $imageKey = if ($fleErrors -match ([System.IO.Path]::GetFileNameWithoutExtension($fle))) {
                @("Sim", "csv_erro")
            } else {
                $("Não", "csv")
            }

            $item = $lst_importar.Items.Add($fle, $imageKey)
            [void]$item.SubItems.Add($hasError)
        }
    $lst_importar.ResumeLayout()
}

function RefreshPregoes {
    $flesToImport = Get-ChildItem -Path "${fldImport}" -File -Filter "*.csv" | Select-Object -ExpandProperty Name | Sort-Object
    $fleErrors    = Get-ChildItem -Path "${fldImport}" -File -Filter "*ERRO.txt" | Select-Object -ExpandProperty Name | Sort-Object
    $pregoes      = Get-ChildItem -Path "$fldData" -Directory | Select-Object -ExpandProperty Name | Sort-Object

    $lst_pregoes.SuspendLayout()
        $lst_pregoes.Items.Clear()

        foreach ($pregao in $pregoes) {
            $imageKey = "folder"

            if ($flesToImport -match "PE_${pregao}.csv") {
                $imageKey = "folder_ok"
            }

            if ($fleErrors -match "PE_${pregao}_ERRO.txt") {
                $imageKey = "folder_erro"
            }

            [void]$lst_pregoes.Items.Add($pregao, $imageKey)
        }
    $lst_pregoes.ResumeLayout()
}

function ClearListSelection {
    param(
        [System.Object]$sender, 
        [System.Object]$e
    )

    $hitTest = $sender.HitTest($e.Location)

    if (-not ($null -eq $hitTest.Item)) {
        return
    }
    
    $sender.SelectedItems.Clear()

    switch ($sender.Name) {
        "lst_importar" {
            if ($lst_pregoes.SelectedItems.Count -gt 0) {
                $lst_pregoes.SelectedItems.Clear()
            }
        }
        "lst_pregoes" {
            if ($lst_importar.SelectedItems.Count -gt 0) {
                $lst_importar.SelectedItems.Clear()
            }
        }
    }
}

function RefreshResumo {
    param(
        [System.Object]$ctr_ = $null,
        [string]$pregao = $null
    )

    switch ($ctr_) {
        $lst_pregoes {
            if ($lst_pregoes.SelectedItems.Count -eq 0) {
                $lst_resumo.Items.Clear()
                $lst_atas.Items.Clear()
                return
            }
            
            $pregao = $lst_pregoes.SelectedItems[0].Text 
        }
        $lst_importar {
            if ($lst_importar.SelectedItems.Count -eq 0) {
                $lst_resumo.Items.Clear()
                $lst_atas.Items.Clear()
                return
            }

            $pregao = $lst_importar.SelectedItems[0].Text
            $pregao = $pregao.substring(3) -replace ".{4}$"

            if (-not(Test-Path -Path "$fldData\$pregao")) {
                $lst_resumo.Items.Clear()
                return
            }
        }
    }

    $lst_resumo.Items.Clear()

    if ([string]::IsNullOrEmpty($pregao)) {
        return
    }

    $pregoes      = Get-ChildItem -Path "$fldData" -Directory | Select-Object -ExpandProperty FullName | Sort-Object | ForEach-Object { $_+"\" }
    $flesToImport = Get-ChildItem -Path "${fldImport}" -File -Filter "*.csv" | Select-Object -ExpandProperty Name | Sort-Object
    $fleErrors    = Get-ChildItem -Path "${fldImport}" -File -Filter "*ERRO.txt" | Select-Object -ExpandProperty Name | Sort-Object

    if (-not($flesToImport -match "PE_${pregao}.csv" -or $pregoes -match "${pregao}")) {
        return
    }

    $lst_resumo.SuspendLayout()
        $item = $lst_resumo.Items.Add("Pregão", "folder")
        [void]$item.SubItems.Add($pregao)

        try {
            $qtyAtas = (Get-ChildItem "$fldData\$pregao" -File -Filter "*.xls*" -ErrorAction Stop | Where-Object { -not $_.Name.StartsWith('~') } | Measure-Object).Count
        }
        catch {
            $qtyAtas = "-"
        }
        finally {
            $item = $lst_resumo.Items.Add("Qtde. Atas", "excel")
            [void]$item.SubItems.Add($qtyAtas)
        }

        if ($flesToImport -match "PE_${pregao}.csv") {
            $hasFleToImport = "Sim"
        } else {
            $hasFleToImport = "Não"
        }
        $item = $lst_resumo.Items.Add("Gerado?", "folder_ok")
        [void]$item.SubItems.Add($hasFleToImport)

        if ($fleErrors -match "PE_${pregao}_ERRO.txt") {
            $hasError = "Sim"
        } else {
            if ($hasFleToImport -eq "Sim") {
                $hasError = "Não"
            } else {
                $hasError = "-"
            }
        }
        $item = $lst_resumo.Items.Add("Com erro?", "folder_erro")
        [void]$item.SubItems.Add($hasError)
    $lst_resumo.ResumeLayout()
}

function Conferir {
    $itemSelected = $lst_importar.SelectedItems[0]

    if ($null -eq $itemSelected) {
        return
    }

    $fleSelected = $itemSelected.Text
    $fleBaseName = ([System.IO.Path]::GetFileNameWithoutExtension($fleSelected))
    $fleToCheck  = "${fldImport}\${fleBaseName}"
    $hasError    = $lst_importar.SelectedItems[0].SubItems[1].Text

    InterfaceMinimize

    Start-Process "${fleToCheck}_CONFERENCIA.xlsx"
    if ($hasError -eq "Sim") {
        Start-Process "${fleToCheck}_ERRO.txt"
    }
}

function Excluir {
    param(
        [System.Object]$lst_
    )
   
    $selection = $lst_.SelectedItems[0].Text

    if (([string]::IsNullOrEmpty($selection))) {
        return
    }

    switch ($lst_.Name) {
        "lst_pregoes"  { 
            $msg = "a pasta $selection e todo o seu conteúdo" 
        }
        "lst_atas"     { 
            $msg = "o arquivo $selection"
        }
        "lst_importar" {
            $msg = "o arquivo $selection e arquivos relacionados"
        }
    }

    $result = [System.Windows.Forms.MessageBox]::Show("Deseja excluir ${msg}?", "Confirmação de exclusão", 4, 48)

    if ($result -eq "No") {
        return
    }

    $msgError = "Não foi possível completar a exclusão.`n`nVerifique se você tem as permissões necessárias ou se não há arquivos abertos."

    switch ($lst_.Name) {
        "lst_pregoes"  { 
            $toDelete = "$fldData\${selection}"

            try {
                Remove-Item -Path "$toDelete" -Force -Recurse -ErrorAction Stop
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show($msgError, "Erro", 0, 16)
            }

            RefreshPregoes
            $lst_atas.Items.Clear()
            $lst_resumo.Items.Clear()
        }
        "lst_atas"     { 
            $pregao    = $lst_atas.Columns[0].Text
            $fldPregao = $pregao.substring(14)
            $toDelete  = "$fldData\$fldPregao\$selection"

            try {
                Get-ChildItem -Path "$toDelete" -File | Remove-Item -Force -ErrorAction Stop                    
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show($msgError, "Erro", 0, 16)
            }

            RefreshAtas
            RefreshResumo
        }
        "lst_importar" { 
            $toDelete            = "${fldImport}\$selection" 
            $toDeleteBasename    = ([System.IO.Path]::GetFileNameWithoutExtension($toDelete))
            $toDeleteConferencia = "${fldImport}\${toDeleteBasename}_CONFERENCIA.xlsx"
            $toDeleteError       = "${fldImport}\${toDeleteBaseName}_ERRO.txt"

            try {
                Get-ChildItem -Path "$toDelete" -File | Remove-Item -Force -ErrorAction Stop  
                Get-ChildItem -Path "$toDeleteConferencia" -File | Remove-Item -Force -ErrorAction Stop  
                if (Test-Path "$toDeleteError") {
                    Get-ChildItem -Path "$toDeleteError" -File | Remove-Item -Force -ErrorAction Stop  
                }
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show($msgError, "Erro", 0, 16)
            }

            RefreshImportar
            RefreshPregoes
            $lst_resumo.Items.Clear()
            $lst_atas.Items.Clear()
        }
    }
}

function Gerar {
    try {
        $pregao = $lst_resumo.Items[0].SubItems[1].Text
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Selecione um pregão antes de gerar o arquivo.", "Atenção", 0, 48)
        return
    }

    if (-not $installedR) {
        [System.Windows.Forms.MessageBox]::Show("Para gerar os arquivos, é necessário instalar o programa R, disponível no site https://cran.r-project.org/bin/windows/base/", "Erro", 0, 16) 
        Start-Process "https://cran.r-project.org/bin/windows/base/"
        return
    }

    $qtyAtas = $lst_resumo.Items[1].SubItems[1].Text
    if ($qtyAtas -eq 0 -or $qtyAtas -eq "-") {
        [System.Windows.Forms.MessageBox]::Show("A pasta do pregão $pregao está vazia ou não disponível.", "Erro", 0, 16)
        return
    }
    
    $hasFleToImport = $lst_resumo.Items[2].SubItems[1].Text
    if ($hasFleToImport -eq "Sim") {
        $result = [System.Windows.Forms.MessageBox]::Show("Já foi gerado o arquivo para importar do pregão $pregao.`n`nGostaria de gerar novamente?", "Confirmação", 4, 48)
        if ($result -eq "No") {
            return
        }
    }

    $result = [System.Windows.Forms.MessageBox]::Show("Certifique-se de que todos os arquivos com dados dos fornecedores estejam salvos e fechados.`n`nCertifique-se também de que todos os arquivos gerados anteriormente estejam fechados e eventualmente movidos ou excluídos`n`nGostaria de continuar com a geração?", "Confirmação", 4, 48)
    if ($result -eq "No") {
        return
    }
    
    $btn_new.Text = "Aguarde..."

    InterfaceMinimize -hideTaskBar

    RunR -script "matl.R" -arguments @($pregao)

    InterfaceRestore

    $fleToImport    = "${fldImport}\PE_${pregao}.csv"
    $hasFleToImport = Test-Path -Path $fleToImport

    $fleErrors    = "${fldImport}\PE_${pregao}_ERRO.txt"
    $hasFleErrors = Test-Path -Path $fleErrors

    $errors          = New-Object System.Collections.ArrayList
    $msgErrorScriptR = ConfigJSON -key "msg_erro" -option "get"
    
    if ($hasFleErrors) {
        foreach ($lineError in Get-Content -Path $fleErrors -Encoding UTF8) {
            $errors.Add($lineError)
        }    
    }
    
    if (-not [string]::IsNullOrEmpty($msgErrorScriptR)) {
        Add-Content -Path $fleErrors -Value $msgErrorScriptR

        $errors.Add("`n=================")
        $errors.Add("ERROS DO SCRIPT")
        $errors.Add("=================")
        foreach ($lineError in $msgErrorScriptR) {
            $errors.Add($lineError)
        }
    }

    $hasFleErrors = Test-Path -Path $fleErrors

    if ($hasFleToImport) {
        $msg   = "O arquivo a importar foi gerado com sucesso."
        $title = "Sucesso"
        $type  = 64
        if ($hasFleErrors) {
            $msg   = "O arquivo a importar foi gerado. Porém, há algumas Atas com problemas."
            $title = "Atenção"
            $type  = 48
        }
    } else {
        $msg   = "Não foi gerado o arquivo para importação.`n`nConfira o log para verificar o que causou o erro."
        $title = "Erro"
        $type  = 16
    }

    if (-not [string]::IsNullOrEmpty($errors)) {
        $errors.Add("`n==================================`n")
        $errors.Add($msg)
        $msg = $errors -join "`n"
    }

    [System.Windows.Forms.MessageBox]::Show($this, $msg, $title, 0, $type)

    RefreshImportar
    RefreshResumo $this $pregao
    RefreshAtas
    RefreshPregoes

    $btn_new.Text = "Novo"
}

function Mover {
    $fleSelected    = $lst_importar.SelectedItems[0].Text
    $fleBaseName    = ([System.IO.Path]::GetFileNameWithoutExtension($fleSelected))
    $fleToImport    = "${fldImport}\${fleSelected}"
    $fleErrors      = "${fldImport}\${fleBaseName}_ERRO.txt"
    $fleConferencia = "${fldImport}\${fleBaseName}_CONFERENCIA.xlsx"
    $hasFleErrors   = Test-Path -Path $fleErrors

    $fldDocuments   = [Environment]::GetFolderPath("MyDocuments").ToLower().Replace("\onedrive", "") + "\MATL_Cadastro"
    $fldDocOriginal = [Environment]::GetFolderPath("MyDocuments").ToLower().Replace("\onedrive", "")

    if ([string]::IsNullOrEmpty($fleSelected)) {
        return
    }

    $result = [System.Windows.Forms.MessageBox]::Show("Deseja mover ${fleSelected} para a pasta ${fldDocuments}?", "Confirmação", 4, 48)
    if ($result -eq "No") {
        return
    }

    if (-not(Test-Path -Path $fldDocuments)) {
        try {
            New-Item -ItemType Directory -Path $fldDocuments -Force -ErrorAction Stop | Out-Null
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Não foi possível criar a pasta ${fldDocuments}.`n`nVerifique se você tem as permissões necessárias.", "Erro", 0, 16)
            [System.Windows.Forms.MessageBox]::Show("Será feita tentativa de mover para a pasta ${fldDocOriginal}.`n`nSe ainda assim ocorrer um erro, mova os arquivos diretamente no Explorador de Arquivos.", "Alerta", 0, 48)
            $fldDocuments = $fldDocOriginal
        }
    }

    try {
        Move-Item -Path "${fleToImport}" -Destination "${fldDocuments}" -Force -ErrorAction Stop  
        Move-Item -Path "$fleConferencia" -Destination "${fldDocuments}" -Force -ErrorAction Stop  
        if ($hasFleErrors) {
            Move-Item -Path "$fleErrors" -Destination "${fldDocuments}" -Force -ErrorAction Stop  
        }
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Não foi possível mover todos os arquivos.`n`nVerifique se você tem as permissões necessárias ou se não há arquivos abertos.", "Erro", 0, 16)
        if (${fldDocuments} -ne $fldDocOriginal) {
            [System.Windows.Forms.MessageBox]::Show("Será feita tentativa de mover para a pasta ${fldDocOriginal}.`n`nSe ainda assim ocorrer um erro, mova os arquivos diretamente no Explorador de Arquivos.", "Alerta", 0, 48)
            try {
                Move-Item -Path "$fleToImport" -Destination "${fldDocOriginal}" -Force -ErrorAction Stop  
                Move-Item -Path "$fleConferencia" -Destination "${fldDocOriginal}" -Force -ErrorAction Stop  
                if ($hasFleErrors) {
                    Move-Item -Path "$fleErrors" -Destination "${fldDocOriginal}" -Force -ErrorAction Stop  
                }
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("Não foi possível mover todos os arquivos.`n`nMova-os diretamente no Explorador de Arquivos.", "Erro", 0, 16)
            }
        }
    }

    RefreshImportar
    RefreshAtas
    RefreshPregoes
    RefreshResumo
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$fldRoot   = $PSScriptRoot
$fldParent = Join-Path -Path $PSScriptRoot -ChildPath .. -Resolve
$fldCommon = "$fldParent\_common"
$fldData   = "${fldRoot}\DADOS"
$fldImport = "${fldRoot}\PARA IMPORTAR"

$configMain = "config.psm1"
if (Test-Path -Path "$fldCommon\$configMain") {
    Unblock-File -Path "$fldCommon\$configMain"
    Import-Module "$fldCommon\$configMain" -Global
} else {
    [System.Windows.Forms.MessageBox]::Show("Não foi localizado o arquivo '${configMain}'.`n`nNão é possível executar o script.", "Erro", 0, 16)
    [Environment]::Exit(1)
}    

main "MATL"

# MAIN FORM

$frm_main, $pic_banner, $tip_ = InterfaceMainForm "MATL - Cadastro Fornecedores" 800 500 "matl"

$padding = 15
$innerPadding = 10

$lst_image = InterfaceImageList

# PANEL RESUMO

$pnl_resumo, $lbl_pnl_resumo = InterfacePanel -top 35 -left 120 -width 260 -height 160 -labelText "Resumo   "

$params = @{
    name = "lst_resumo"; 
    width = InterfacePosition $pnl_resumo "width" $padding;
    height = 130;
    top = $padding;
    left = $padding;
    sorting = "None";
    columns = [ordered]@{
        'Status' = 120;
        'Informação' = -2
    }
}
$lst_resumo = InterfaceList @params

$pnl_resumo.Controls.AddRange(@($lst_resumo))
$frm_main.Controls.AddRange(@($pnl_resumo, $lbl_pnl_resumo))
$lbl_pnl_resumo.BringToFront()

# PANEL IMPORTAR

$params = @{
    top = $pnl_resumo.Top;
    left = InterfacePosition $pnl_resumo "right" 25;
    width = 310;
    height = $pnl_resumo.Height;
    labelText = "Arquivos a importar"
}
$pnl_importar, $lbl_pnl_importar = InterfacePanel @params

$params = @{
    name = "lst_importar"; 
    width = 250;
    height = $lst_resumo.Height;
    top = $padding;
    left = $padding;
    events = @{
        MouseUp = {
            RefreshResumo $this
            RefreshAtas
        };
        MouseDown = {
            param($sender, $e)
            ClearListSelection $sender $e
        }
    };
    columns = [ordered]@{
        "MATL - Importar" = 180;
        "Erros" = 40
    }
}
$lst_importar = InterfaceList @params

$params = @{
    hover = $true;
    name = "excluir";
    top = InterfacePosition $lst_importar "bottom" -20;
    left = InterfacePosition $lst_importar "right" $innerPadding;
    function = {Excluir $lst_importar};
    tag = "Excluir arquivo"
}
$btn_delete_importar = InterfaceButtonImage @params

$params = @{
    hover = $true;
    name = "mover";
    top = InterfacePosition $btn_delete_importar "top" $innerPadding;
    left = $btn_delete_importar.Left;
    function = {Mover};
    tag = "Mover o arquivo para a pasta Documentos"
}
$btn_move = InterfaceButtonImage @params

$params = @{
    hover = $true;
    name = "check";
    top = InterfacePosition $btn_move "top" $innerPadding
    left = $btn_delete_importar.Left;
    function = {Conferir};
    tag = "Conferir informações do arquivo a importar"
}
$btn_check = InterfaceButtonImage @params

$pnl_importar.Controls.AddRange(@($lst_importar, $btn_check, $btn_move, $btn_delete_importar))
$frm_main.Controls.AddRange(@($pnl_importar, $lbl_pnl_importar))
$lbl_pnl_importar.BringToFront()

# PANEL PREGÕES

$params = @{
    top = InterfacePosition $pnl_resumo "bottom" 30;
    left = $pnl_resumo.Left;
    width = $pnl_importar.Left + $pnl_importar.Width - $pnl_resumo.Left;
    height = 160;
    labelText = "Pregões"
}
$pnl_pregoes, $lbl_pnl_pregoes = InterfacePanel @params

$params = @{
    name = "lst_pregoes"; 
    width = $lst_resumo.Width;
    height = 130;
    top = $padding;
    left = $padding;
    view = "LargeIcon";
    events = @{
        MouseUp = {
            RefreshResumo $this
            RefreshAtas
        };
        MouseDown = {
            param($sender, $e)
            ClearListSelection $sender $e
        }
    };
    columns = [ordered]@{"Pregão" = 175}
}    
$lst_pregoes = InterfaceList @params

$params = @{
    hover = $true;
    name = "excluir";
    top = InterfacePosition $lst_pregoes "bottom" -20;
    left = InterfacePosition $lst_pregoes "right" $innerPadding;
    function = {Excluir $lst_pregoes};
    tag = "Excluir pasta"
}
$btn_delete_pregao = InterfaceButtonImage @params

$params = @{
    labelText = "Conteúdo da pasta"; 
    name = "lst_atas"; 
    width = $lst_importar.Width;
    height = $lst_pregoes.Height - 15;
    top = $padding + 15;
    left = $pnl_importar.Left + $lst_importar.Left - $pnl_pregoes.Left;
    columns = [ordered]@{
        "Atas" = 180;
        "Erros" = 40
    }
}
$lst_atas, $lbl_lst_atas = InterfaceList @params

$params = @{
    hover = $true;
    name = "excluir";
    top = InterfacePosition $lst_atas "bottom" -20;
    left = InterfacePosition $lst_atas "right" $innerPadding;
    function = {Excluir $lst_atas};
    tag = "Excluir arquivo"
}
$btn_delete_ata = InterfaceButtonImage @params

$params = @{
    hover = $true;
    name = "excel_bw";
    top = InterfacePosition $btn_delete_ata "top" $innerPadding;
    left = $btn_delete_ata.Left;
    function = {Abrir};
    tag = "Abrir a planilha com os dados do fornecedor"
}
$btn_open = InterfaceButtonImage @params

$controls = @(
    $lst_pregoes,
    $btn_delete_pregao,
    $lbl_lst_atas,
    $lst_atas,
    $btn_open,
    $btn_delete_ata
)
$pnl_pregoes.Controls.AddRange($controls)
$frm_main.Controls.AddRange(@($pnl_pregoes, $lbl_pnl_pregoes))
$lbl_pnl_pregoes.BringToFront()

# FORM ICONS AND IMAGES

$params = @{
    size = 70;
    top = 260;
    left = 25;
    name = "ufsc";
    function = {
        InterfaceMinimize
        Start-Process "https://www.ufsc.br"
    };
    tag = "Ir para a página da UFSC."
}
$btn_ufsc = InterfaceButtonImage @params

$params = @{
    size = 100;
    height = 50;
    top = InterfacePosition $btn_ufsc "bottom";
    left = 10;
    name = "dcom.tif";
    function = {
        InterfaceMinimize
        Start-Process "https://dcom.ufsc.br"
    };
    tag = "Ir para a página do DCOM."
}
$btn_dcom = InterfaceButtonImage @params

$params = @{
    size = 40;
    top = InterfacePosition $pnl_pregoes "bottom" -40;
    left = InterfacePosition $pnl_pregoes "right" $padding;
    name = "folder";
    function = {
        InterfaceMinimize
        Start-Process "$fldData"
    };
    tag = "Abrir a pasta com os dados dos fornecedores."
}
$btn_folder = InterfaceButtonImage @params

$params = @{
    size = 40;
    top = $pnl_importar.Top;
    left = $btn_folder.Left;
    name = "r";
    function = {
        InterfaceMinimize
        Start-Process "https://cran.r-project.org/bin/windows/base/"
    };
    tag = "O programa R não está instalado.`r`nClique aqui para ir para a página de download."
}
$btn_r = InterfaceButtonImage @params

$params = @{
    text = "Sair";
    top = 415;
    left = InterfacePosition $pnl_pregoes "right" -targetWidth 90;
    function = {InterfaceClose};
    tag = "Sair do script" 
}
$btn_exit = InterfaceButton @params

$params = @{
    text = "Novo";
    image = "Gerar"
    top = $btn_exit.Top;
    left = InterfacePosition $btn_exit "left" $padding;
    function = {Gerar};
    tag = "Gerar novo arquivo para importar"    
}
$btn_new = InterfaceButton @params 

$params = @{
    text = "Atualizar";
    top = $btn_exit.Top;
    left = $pnl_pregoes.Left;
    function = {
        RefreshPregoes
        RefreshResumo
        RefreshImportar
        RefreshAtas
    };
    tag = "Atualizar arquivos e informações"    
}
$btn_refresh = InterfaceButton @params

$controls = @(
    $btn_ufsc,
    $btn_dcom,
    $btn_r,
    $btn_folder,
    $btn_refresh,
    $btn_new,
    $btn_exit,
    $pic_banner
)
$frm_main.Controls.AddRange($controls)

# SHOW FORM

InterfaceShowForm -title "MATL - CADASTRO FORNECEDORES - MANTER ABERTA" -start {
    RefreshImportar
    RefreshPregoes

    CheckRInstallation
} 