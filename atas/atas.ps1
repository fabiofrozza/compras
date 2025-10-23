# FUNCTIONS

function AtualizarFornecedores {
    $sicafs = Get-ChildItem -Path "$fldRoot\SICAF\*.pdf" -Name | Sort-Object
    
    $lst_sicaf.SuspendLayout()
        $lst_sicaf.Rows.Clear()
        $i = [int]$txt_n_ata.Value
        foreach ($sicaf in $sicafs) {
            [void]$lst_sicaf.Rows.Add($i, $sicaf)
            $i = $i + 1
        }
    $lst_sicaf.ResumeLayout()

    if ($lst_sicaf.RowCount -eq 0) {
        $btn_erro_sicaf.Show()
        $btn_delete.Enabled = $false
        $btn_delete.Image = InterfaceGetImage "excluir_desativado"
    } else {
        $btn_erro_sicaf.Hide()
        $btn_delete.Enabled = $true
    }
}

function ChecarDados {

    if ($txt_processo.MaskCompleted -eq $false -or $txt_processo.Text -eq "") {
        $btn_erro_processo.Show()
    } else {
        $btn_erro_processo.Hide()
    }

    if ($txt_objeto.Text -eq "") {
        $txt_objeto.Width = (InterfacePosition $pnl_pregao "width" $PADDING_OUTER) - ($PADDING_OUTER * 2) - $btn_erro_objeto.Width
        $btn_erro_objeto.Show()
    } else {
        $btn_erro_objeto.Hide()
        $txt_objeto.Width = (InterfacePosition $pnl_pregao "width" $PADDING_OUTER)
        $txt_objeto.Text = $txt_objeto.Text.Replace("`r`n", " ")
        $txt_objeto.Refresh()
    }

    foreach ($ctr_ in $pnl_pregao.Controls) {
        if ($ctr_ -is [System.Windows.Forms.NumericUpDown]) {
            if ([string]::IsNullOrWhiteSpace($ctr_.Text) -or $ctr_.Value -lt $ctr_.Minimum) {
                $ctr_.Value = $ctr_.Minimum
                $ctr_.Text  = $ctr_.Value
            }
        }
    }
}

function DataPorExtenso {

    $date  = $txt_data_seletor.Value
    $day   = $date.ToString("dd")
    $month = $date.ToString("MMMM")
    $year  = $date.ToString("yyyy")

    $txt_data_extenso.Text = "$day de $month de $year"
}  
  
function DeleteFiles {
    $result = [System.Windows.Forms.MessageBox]::Show("Deseja excluir todos os relatórios SICAF?", "Confirmação de exclusão", 4, 48)

    if ($result -eq "Yes") {
        Get-ChildItem -Path "$fldRoot\SICAF\*.pdf" -File | Remove-Item -Force 
        AtualizarFornecedores
    }
}

function Gerar {

    if (IsUserRunningR) {
        return
    }

    $errors = New-Object System.Collections.ArrayList

    if (!$installedR) {
        $errors.Add("instale o programa R, disponível no site https://cran.r-project.org/bin/windows/base/")
        Start-Process "https://cran.r-project.org/bin/windows/base/"
    }

    if ($txt_processo.MaskCompleted -eq $false -or $txt_objeto.Text -eq "") {
        $errors.Add("Preencha todos as informações do pregão.")
    }

    if ($lst_sicaf.RowCount -eq 0) {
        $errors.Add("Não há relatórios na pasta SICAF.")
    }

    if ($errors.Count -eq 0) {
        $btn_gerar.Text    = "Aguarde..."
        $btn_gerar.Enabled = $false
        $btn_clear.Enabled = $false
        $btn_exit.Enabled  = $false

        ConfigJSON -show

        InterfaceMinimize -HideTaskBar

        RunR -script "atas.R"

        InterfaceRestore

        switch (ConfigJSON -key "resultado_geracao" -option "get") 
        {
            "sucesso" {
                $msg   = "Script finalizado.`n`nConfira a planilha e gere as Atas no Word."
                $title = "Sucesso"
                $type  = 64
            }
            "erro" {
                $msg   = "Não foram salvos os dados, pois foram verificados os seguintes erros:`n======", (ConfigJSON -key "msg_erro" -option "get") -join "`n"
                $title = "Erro"
                $type  = 16
            }
            "ambos" {
                $msg   = "Os dados foram salvos, mas foram verificados os seguintes problemas:`n======", (ConfigJSON -key "msg_erro" -option "get") -join "`n"
                $title = "Alerta"
                $type  = 48
            }
            Default {
                $msg   = "O script retornou algum erro não identificado.`n`nVerifique os logs."
                $title = "Erro desconhecido"
                $type  = 16
            }
        }
    } else {
        $msg   = $errors -join "`n"
        $title = "Erro"
        $type  = 16
    }

    $btn_gerar.Text    = "Gerar"
    $btn_gerar.Enabled = $true
    $btn_clear.Enabled = $true
    $btn_exit.Enabled  = $true    

    [System.Windows.Forms.MessageBox]::Show($this, $msg, $title, 0, $type)

    ConfigJSON -show
}

function CleanPregao {
    $result = [System.Windows.Forms.MessageBox]::Show("Deseja limpar todos os dados?", "Confirmação", 4, 48)

    if ($result -eq "Yes") {
        $txt_n_pregao.Text      = 1
        $txt_ano_pregao.Value   = Get-Date -Format "yyyy"
        $txt_processo.Text      = ""
        $txt_objeto.Text        = ""
        $txt_data_seletor.Value = Get-Date
        $txt_n_ata.Text         = 1
        $txt_ano_ata.Value      = Get-Date -Format "yyyy"

        ChecarDados
        DataPorExtenso
        AtualizarFornecedores
        SalvarDados
    }
}

function SalvarDados {

    ConfigJSON -key "n_pregao" -value ([int]$txt_n_pregao.Value)
    ConfigJSON -key "ano_pregao" -value ([int]$txt_ano_pregao.Value)
    ConfigJSON -key "processo" -value $txt_processo.Text
    ConfigJSON -key "objeto" -value $txt_objeto.Text
    ConfigJSON -key "data" -value $txt_data_extenso.Text
    ConfigJSON -key "n_ata" -value ([int]$txt_n_ata.Value)
    ConfigJSON -key "ano_ata" -value ([int]$txt_ano_ata.Value)
    ConfigJSON -key "data_seletor" -value $txt_data_seletor.Value

}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$fldRoot   = $PSScriptRoot
$fldParent = Join-Path -Path $PSScriptRoot -ChildPath .. -Resolve
$fldCommon = "$fldParent\_common"
$fldAtas   = "$fldRoot\ATAS"
$fldSicaf  = "$fldRoot\SICAF"

$configMain = "config.psm1"
if (Test-Path -Path "$fldCommon\$configMain") {
    Unblock-File -Path "$fldCommon\$configMain"
    Import-Module "$fldCommon\$configMain"
} else {
    [System.Windows.Forms.MessageBox]::Show("Não foi localizado o arquivo '${configMain}'.`n`nNão é possível executar o script.", "Erro", 0, 16)
    [Environment]::Exit(1)
}

main "ATAS"

# MAIN FORM

InterfaceConstants -frmWidth 1000 -frmHeight 500 -marginTop 35 -marginLeft 120 -marginRight 85

$frm_main, $pic_banner, $tip_ = InterfaceMainForm "Atas" "atas"

# PANEL INFORMAÇÕES DO PREGÃO

$params = @{
    labelText = "Informações do Pregão";
    top = $MARGIN_TOP;
    left = $MARGIN_LEFT;
    width = $UTIL_AREA_WIDTH * 0.55;
    height = $UTIL_AREA_HEIGHT;
}
$pnl_pregao, $lbl_pregao = InterfacePanel @params

$utilAreaPnlPregao = $pnl_pregao.Height - ($PADDING_OUTER * 2)
$offsetY = 30
$offsetX = 30

$params = @{
    type = "NumericUpDown";
    labelText = "Número";
    top = $PADDING_OUTER + 20;
    left = $PADDING_OUTER;
    width = 100;
    tag = "Informe o número do pregão";
    min = 1;
    max = 999999999;
    events = @{
        LostFocus = {
            ChecarDados
        }
    }
}
$txt_n_pregao, $lbl_n_pregao = InterfaceControl @params

$params = @{
    type = "NumericUpDown";
    labelText = "Ano";
    top = $txt_n_pregao.Top;
    left = InterfacePosition $txt_n_pregao "right" $offsetX;
    width = 60;
    tag = "Informe o ano do pregão";
    min = 2020;
    max = (Get-Date).Year + 1;
    events = @{
        LostFocus = {
            ChecarDados
        }
    }
}
$txt_ano_pregao, $lbl_ano_pregao = InterfaceControl @params

$params = @{
    type = "MaskedTextBox";
    labelText = "Processo SPA";
    top = InterfacePosition $txt_n_pregao "bottom" $offsetY;
    left = $txt_n_pregao.Left;
    width = 150;
    tag = "Digite o número do processo no SPA";
    properties = @{
        Mask = "\2\3\0\8\0\.000000\/0000\-00"
    };
    events = @{
        TextChanged = {
            ChecarDados
        }
        LostFocus = {
            ChecarDados
        }
    }
}
$txt_processo, $lbl_processo = InterfaceControl @params

$params = @{
    size = 20;
    name = "erro";
    top = InterfacePosition $txt_processo "center" -targetWidth 20;
    left = InterfacePosition $txt_processo "right" 10;
    tag = "Número do processo no SPA não preenchido corretamente.`nO formato correto é 23080.000000/0000-00"
}
$btn_erro_processo = InterfaceButtonImage @params

$params = @{
    type = "TextBox";
    labelText = "Objeto (após 'REGISTRAR PREÇO...')";
    top = InterfacePosition $txt_processo "bottom" $offsetY;
    left = $txt_n_pregao.Left;
    width = InterfacePosition $pnl_pregao "width" $PADDING_OUTER;
    height = 65;
    tag = "Informe o objeto do pregão conforme o Edital.`nNa Ata, esta informação será incluída após o trecho `"REGISTRAR PREÇO...`".";
    properties = @{
        Multiline = $true;
        ScrollBars = "Vertical"
    };
    events = @{
        LostFocus = {
            ChecarDados
        }
    }
}
$txt_objeto, $lbl_objeto = InterfaceControl @params

$params = @{
    size = 20;
    name = "erro";
    top = InterfacePosition $txt_objeto "center" -targetWidth 20;
    left = (InterfacePosition $pnl_pregao "innerCorner" -targetWidth 20 -offset $PADDING_OUTER).left;
    tag = "Objeto não preenchido.`nO campo deve conter o objeto do pregão conforme o Edital.`nNa Ata, esta informação será incluída após o trecho 'REGISTRAR PREÇO...'"
}
$btn_erro_objeto = InterfaceButtonImage @params

$params = @{
    type = "DateTimePicker";
    labelText = "Data Pregão DOU";
    top = InterfacePosition $txt_objeto "bottom" $offsetY;
    left = $txt_n_pregao.Left;
    width = 150;
    tag = "Informe a data da publicação do pregão no DOU.";
    properties = @{
        Format = "Custom";
        CustomFormat = "dd/MMMM/yyyy"
    };
    events = @{
        CloseUp = {
            DataPorExtenso
        };
        TextChanged = {
            DataPorExtenso
        }
    }
}
$txt_data_seletor, $lbl_date = InterfaceControl @params

$params = @{
    top = $txt_data_seletor.Top + 25;
    left = $txt_n_pregao.Left;
    width = 300;
    height = 20;
}
$txt_data_extenso = InterfaceLabel @params

$params = @{
    type = "NumericUpDown";
    labelText = "Primeira Ata";
    top = InterfacePosition $txt_data_extenso "bottom" $offsetY;
    left = $txt_n_pregao.Left;
    width = 80;
    tag = "Informe o número da primeira Ata a ser gerada.`nAs próximas serão numeradas automaticamente, conforme os relatórios SICAF presentes na pasta.";
    min = 1;
    max = 10000;
    events = @{
        Click = {
            AtualizarFornecedores
        }
        LostFocus = {
            ChecarDados
            AtualizarFornecedores
        }
    }
}
$txt_n_ata, $lbl_n_ata = InterfaceControl @params

$params = @{
    type = "NumericUpDown";
    labelText = "Ano Ata";
    top = $txt_n_ata.Top;
    left = InterfacePosition $txt_n_ata "right" $offsetX;
    width = 60;
    tag = "Informe o ano das Atas.";
    min = $txt_ano_pregao.Minimum;
    max = $txt_ano_pregao.Maximum;
    events = @{
        LostFocus = {
            ChecarDados
        }
    }
}
$txt_ano_ata, $lbl_ano_ata = InterfaceControl @params

$position = InterfacePosition $pnl_pregao "innerCorner" -targetWidth 20 -targetHeight 20 -offset $PADDING_OUTER;
$params = @{
    hover = $true;
    size = 20;
    name = "excluir";
    top = $position.top;
    left = $position.left;
    function = {CleanPregao};
    tag = "Clique aqui para limpar as informações."
}
$btn_clear = InterfaceButtonImage @params

$controls = @(
    $lbl_n_pregao, $txt_n_pregao,
    $lbl_ano_pregao, $txt_ano_pregao,
    $lbl_processo, $txt_processo, $btn_erro_processo,
    $lbl_objeto, $txt_objeto, $btn_erro_objeto,
    $lbl_date, $txt_data_seletor, $txt_data_extenso,
    $lbl_n_ata, $txt_n_ata,
    $lbl_ano_ata, $txt_ano_ata,
    $btn_clear
)
$pnl_pregao.Controls.AddRange($controls)
$frm_main.Controls.AddRange(@($pnl_pregao, $lbl_pregao))
$lbl_pregao.BringToFront()

# PANEL FORNECEDORES

$params = @{
    labelText = "Fornecedores";
    top = $MARGIN_TOP;
    left = InterfacePosition $pnl_pregao "right" $PADDING_OUTER;
    width = $UTIL_AREA_WIDTH * 0.45 - $PADDING_OUTER;
    height = $UTIL_AREA_HEIGHT
}
$pnl_fornecedores, $lbl_fornecedores = InterfacePanel @params

$params = @{
    type = "DataGridView";
    labelText = "Relatórios SICAF";
    top = $PADDING_OUTER + 20;
    left = $PADDING_OUTER;
    width = InterfacePosition $pnl_fornecedores "width" -offset $PADDING_OUTER;
    height = 260;
    properties = @{
        BorderStyle              = "Fixed3D";
        ColumnCount              = 2;
        ColumnHeadersVisible     = $true;
        RowHeadersVisible        = $false;
        AllowUserToOrderColumns  = $true;
        AllowUserToResizeColumns = $false;
        AllowUserToResizeRows    = $false;
        AllowUserToAddRows       = $false;
        ReadOnly                 = $true;
        ScrollBars               = "Vertical"
    };
    columns = [ordered]@{
        "Ata" = (InterfacePosition $pnl_fornecedores "width" -offset $PADDING_OUTER) * .2;
        "Fornecedor/Arquivo" = (InterfacePosition $pnl_fornecedores "width" -offset $PADDING_OUTER) * .77
    }
}
$lst_sicaf, $lbl_sicaf = InterfaceList @params

$position = InterfacePosition $pnl_fornecedores "innerCorner" -targetWidth 20 -targetHeight 20 -offset $PADDING_OUTER
$params = @{
    hover = $true;
    size = 20;
    name = "excluir";
    top = $position.top;
    left = $position.left;
    function = {DeleteFiles};
    tag = "Clique aqui para excluir todos os relatórios exibidos acima."
}
$btn_delete = InterfaceButtonImage @params

$params = @{
    hover = $true;
    size = 20;
    name = "folder_bw";
    top = $btn_delete.Top;
    left = InterfacePosition $btn_delete "left" 20;
    function = {
        InterfaceMinimize
        Start-Process "$fldRoot/SICAF"
    };
    tag = "Clique aqui para abrir a pasta onde devem ser salvos os relatórios de credenciamento do SICAF."
}
$btn_folder_sicaf = InterfaceButtonImage  @params

$params = @{
    hover = $true;
    size = 20;
    name = "atualizar_bw";
    top = $btn_delete.Top;
    left = InterfacePosition $btn_folder_sicaf "left" 20;
    function = {AtualizarFornecedores};
    tag = "Clique aqui para atualizar a lista de relatórios de credenciamento do SICAF."
}
$btn_refresh_sicaf = InterfaceButtonImage @params

$params = @{
    size = 20;
    name = "erro";
    top = $btn_delete.Top;
    left = $lst_sicaf.Left;
    tag = "Não há relatórios na pasta SICAF."
}
$btn_erro_sicaf = InterfaceButtonImage @params
$btn_erro_sicaf.Hide()

$controls = @(
    $lbl_sicaf, $lst_sicaf, 
    $btn_refresh_sicaf, 
    $btn_folder_sicaf, 
    $btn_delete, 
    $btn_erro_sicaf
)
$pnl_fornecedores.Controls.AddRange($controls)
$frm_main.Controls.AddRange(@($pnl_fornecedores, $lbl_fornecedores))
$lbl_fornecedores.BringToFront()

$txt_n_pregao.Text     = ConfigJSON -key "n_pregao" -option "get"
$txt_ano_pregao.Text   = ConfigJSON -key "ano_pregao" -option "get"
$txt_processo.Text     = ConfigJSON -key "processo" -option "get"
$txt_objeto.Text       = ConfigJSON -key "objeto" -option "get"
$txt_data_extenso.Text = ConfigJSON -key "data" -option "get"
$txt_n_ata.Text        = ConfigJSON -key "n_ata" -option "get"
$txt_ano_ata.Text      = ConfigJSON -key "ano_ata" -option "get"
try {
    $txt_data_seletor.Value = ConfigJSON -key "data_seletor" -option "get"
}
catch {
    $txt_ano_ata = Get-Date
}

# FORM ICONS AND BUTTONS

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
    top = $pnl_fornecedores.Top;
    left = InterfacePosition $pnl_fornecedores "right" $PADDING_OUTER;
    name = "word";
    function = {
        InterfaceMinimize
        Get-ChildItem "$fldRoot/ATAS/MODELO SCRIPT*" | ForEach-Object { Start-Process $_ }
    };
    tag = "Clique aqui para abrir o modelo do Word para geração das Atas."
}
$btn_word = InterfaceButtonImage @params

$offsetY = 10

$params = @{
    size = 40;
    top = InterfacePosition $btn_word "bottom" $offsetY;
    left = $btn_word.Left;
    name = "excel";
    function = {
        InterfaceMinimize
        Start-Process "$fldRoot/ATAS/dados_atas.xlsx"
    };
    tag = "Clique aqui para abrir a planilha com os dados obtidos (se existir)." 
}
$btn_excel = InterfaceButtonImage @params

$params = @{
    size = 40;
    top = InterfacePosition $btn_excel "bottom" $offsetY;
    left = $btn_word.Left;
    name = "folder";
    function = {
        InterfaceMinimize
        Invoke-Item "$fldRoot/ATAS/"
    };
    tag = "Clique aqui para abrir a pasta onde são salvos a planilha com os dados e as Atas geradas." 
}
$btn_folder_atas = InterfaceButtonImage @params

$params = @{
    size = 40;
    top = InterfacePosition $pnl_fornecedores "bottom" -targetWidth 40;
    left = $btn_word.Left;
    name = "r";
    function = {
        InterfaceMinimize
        Start-Process "https://cran.r-project.org/bin/windows/base/"
    };
    tag = "O programa R não está instalado.`r`nClique aqui para ir para a página de download.";
}
$btn_r = InterfaceButtonImage @params

$params = @{
    text = "Sair";
    top = InterfacePosition $pnl_fornecedores "bottom" 30;
    left = InterfacePosition $pnl_fornecedores "right" -targetWidth 90;
    function = {
        SalvarDados
        InterfaceClose
    };
    tag = "Sair do script"
}
$btn_exit = InterfaceButton @params

$params = @{
    text = "Gerar";
    top = $btn_exit.Top;
    left = InterfacePosition $btn_exit "left" 20;
    function = { 
        SalvarDados
        Gerar
    };
    tag = "Clique aqui para gerar a planilha com os dados dos relatórios de credenciamento SICAF."    
}
$btn_gerar = InterfaceButton @params 

$controls = @(
    $btn_ufsc, 
    $btn_dcom, 
    $btn_r, 
    $btn_word, 
    $btn_excel, 
    $btn_folder_atas, 
    $btn_gerar, 
    $btn_exit, 
    $pic_banner
)
$frm_main.Controls.AddRange($controls)

# SHOW FORM

InterfaceShowForm -title "ATAS" -start {
    AtualizarFornecedores
    ChecarDados
    DataPorExtenso
    ConfigJSON -show
    CheckRInstallation
} -close {
    SalvarDados
}