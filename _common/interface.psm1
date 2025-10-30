function InterfaceConstants {
    param (
        [int]$frmWidth,
        [int]$frmHeight,
        [int]$marginLeft,
        [int]$marginRight,
        [int]$marginTop
    )
    
    $config = GetEnvConfig
    $msgNotPersonalized = "(Para personalizar os dados da empresa, leia o arquivo README)"

    SetScriptConstants @{
        # Company info (for icons)
        COMPANY_NAME    = if ([string]::IsNullOrEmpty($config.COMPANY_NAME))    { "empresa $msgNotPersonalized" }      else { $config.COMPANY_NAME }
        COMPANY_SITE    = if ([string]::IsNullOrEmpty($config.COMPANY_SITE))    { "www.empresa.com.br" }               else { $config.COMPANY_SITE }
        DEPARTMENT_NAME = if ([string]::IsNullOrEmpty($config.DEPARTMENT_NAME)) { "departamento $msgNotPersonalized" } else { $config.DEPARTMENT_NAME }
        DEPARTMENT_SITE = if ([string]::IsNullOrEmpty($config.DEPARTMENT_SITE)) { "www.departamento.com.br" }          else { $config.DEPARTMENT_SITE }
        MANUAL_SITE     = if ([string]::IsNullOrEmpty($config.MANUAL_SITE))     { "www.departamento.com.br/manual" }   else { $config.MANUAL_SITE }
        LISTS_PAGE_SITE = if ([string]::IsNullOrEmpty($config.LISTS_PAGE_SITE)) { "www.departamento.com.br/lists" }    else { $config.LISTS_PAGE_SITE }

        # UI control sizes
        FRM_MAIN_WIDTH    = $frmWidth
        FRM_MAIN_HEIGHT   = $frmHeight
        PIC_BANNER_HEIGHT = 75

        BTN_WIDTH  = 90
        BTN_HEIGHT = 30

        BTN_IMAGE_SMALL_WIDTH  = 20
        BTN_IMAGE_SMALL_HEIGHT = 20
        BTN_IMAGE_BIG_WIDTH    = 50
        BTN_IMAGE_BIG_HEIGHT   = 50

        # UI colors
        COLOR_OK      = "LightGreen"
        COLOR_SUCCESS = "LightGreen"
        COLOR_WARNING = "LightGoldenrodYellow"
        COLOR_ERROR   = "LightSalmon"
    
        # UI paddings
        PADDING_INNER = 10
        PADDING_OUTER = 15
    }    

    SetScriptConstants @{
        # UI margins
        MARGIN_TOP    = $marginTop
        MARGIN_BOTTOM = $PIC_BANNER_HEIGHT + $PADDING_OUTER
        MARGIN_LEFT   = $marginLeft
        MARGIN_RIGHT  = $marginRight
    }

    SetScriptConstants @{
        # UI util area
        UTIL_AREA_WIDTH  = $FRM_MAIN_WIDTH - $MARGIN_LEFT - $MARGIN_RIGHT
        UTIL_AREA_HEIGHT = $FRM_MAIN_HEIGHT - $MARGIN_TOP - $MARGIN_BOTTOM
    }
}

function InterfaceMainForm {
    param (
        [string]$title,
        [string]$icon
    )

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    #[System.Windows.Forms.Application]::EnableVisualStyles()

    $frm_                 = New-Object System.Windows.Forms.Form
    $frm_.Text            = $title
    $frm_.ClientSize      = New-Object System.Drawing.Size($FRM_MAIN_WIDTH, $FRM_MAIN_HEIGHT)  
    $frm_.StartPosition   = "CenterScreen"
    $frm_.MaximizeBox     = $false
    $frm_.KeyPreview      = $true
    $frm_.Icon            = New-Object System.Drawing.Icon ("$fldImages\${icon}.ico")
    $frm_.FormBorderStyle = 'Fixed3D'

    return @($frm_, (InterfaceBanner), (InterfaceTip))
}

function InterfaceTip {

    $tip_                = New-Object System.Windows.Forms.ToolTip
    $tip_.AutomaticDelay = 100
    $tip_.ToolTipIcon    = 1

    return $tip_
}

function InterfaceBanner {

    $pic_          = New-Object Windows.Forms.PictureBox
    $pic_.Image    = InterfaceGetImage "banner"
    $pic_.Width    = $FRM_MAIN_WIDTH
    $pic_.Height   = $PIC_BANNER_HEIGHT
    $pic_.Left     = 0
    $pic_.Top      = $FRM_MAIN_HEIGHT - $pic_.Height
    $pic_.SizeMode = "StretchImage"

    return $pic_
}

function InterfaceButton {
    param (
        [string]$text,
        [string]$image = $text.ToLower(),
        [string]$tag,
        [int]$top,
        [int]$left,
        [scriptblock]$function,
        [int]$width = $BTN_WIDTH,
        [int]$height = $BTN_HEIGHT
    )

    $btn_        = New-Object System.Windows.Forms.Button
    $btn_.Width  = $width
    $btn_.Height = $height
    $btn_.Top    = $top
    $btn_.Left   = $left

    $btn_.Image             = (New-Object System.Drawing.Bitmap (InterfaceGetImage $image), $BTN_IMAGE_SMALL_WIDTH, $BTN_IMAGE_SMALL_HEIGHT)
    $btn_.TextImageRelation = "ImageBeforeText"
    $btn_.ImageAlign        = "TopLeft"
    $btn_.TextAlign         = "MiddleRight"
    
    $btn_.Text = $text
    $btn_.Tag  = $tag
    
    $btn_.Add_Click($function)
    $btn_.Add_MouseEnter({Tip $this})
    $btn_.Add_MouseLeave({Tip $this -close})

    return $btn_
}

function InterfaceButtonImage {
    param (
        [switch]$hover,
        [string]$name,
        [string]$image = $name,
        [int]$top,
        [int]$left,
        [int]$size = $BTN_IMAGE_SMALL_WIDTH,
        [int]$height = $size,
        [string]$tag,
        [scriptblock]$function
    )

    $btn_        = New-Object Windows.Forms.PictureBox
    $btn_.Top    = $top
    $btn_.Left   = $left
    $btn_.Width  = $size
    $btn_.Height = $height

    $btn_.Name  = $name
    $btn_.Image = InterfaceGetImage $image

    $btn_.SizeMode = "Zoom"
    
    if ($function) {
        $btn_.Add_Click($function)
        $btn_.Cursor = "Hand"
        $btn_.Tag    = $tag
    } else {
        if ($tag) {
            $btn_.Cursor = "Help"
            $btn_.Tag    = $tag
        } else {
            $btn_.Cursor = "Default"
        }
    }

    $btn_.Add_MouseEnter({
        Tip $this
        if ($hover) { 
            InterfaceButtonHover $this -mouseEnter 
        }
    }.GetNewClosure())

    $btn_.Add_MouseLeave({
        Tip $this -close
        if ($hover) { 
            InterfaceButtonHover $this 
        }
    }.GetNewClosure())

    return $btn_
}

function InterfaceImageList {
    
    $lst_           = New-Object System.Windows.Forms.ImageList 
    $lst_.ImageSize = "$BTN_IMAGE_SMALL_WIDTH,$BTN_IMAGE_SMALL_HEIGHT" 
    $images = @{
        "excel"       = "excel"
        "excel_erro"  = "excel_erro"
        "csv"         = "csv"
        "csv_erro"    = "csv_erro"
        "pdf"         = "pdf"
        "company"     = "company"
        "config"      = "settings"
        "config_erro" = "erro"
        "folder"      = "folder"
        "folder_ok"   = "folder_ok"
        "folder_erro" = "folder_erro"   
    }
    foreach ($key in $images.Keys) {
        $image = InterfaceGetImage $images[$key]

        $lst_.Images.Add($key, $image)
    }

    return $lst_
}

function InterfaceList {
    param (
        [string]$type = "ListView",
        [string]$labelText,
        [string]$name,
        [int]$width,
        [int]$height,
        [int]$top,
        [int]$left,
        [string]$view = "Details",
        [string]$sorting = "Ascending",
        [System.Windows.Forms.ContextMenuStrip]$mnuContext,
        [hashtable]$events = @{},
        [hashtable]$properties = @{},
        [System.Collections.Specialized.OrderedDictionary]$columns
    )
    
    switch ($type) {
        "ListView" {
            $lst_                = New-Object System.Windows.Forms.ListView
            $lst_.View           = $view
            $lst_.Sorting        = $sorting
            $lst_.HeaderStyle    = "Nonclickable"
            $lst_.BorderStyle    = "FixedSingle"
            $lst_.HideSelection  = $false
            $lst_.MultiSelect    = $false
            $lst_.FullRowSelect  = $true
            $lst_.SmallImageList = $lst_image
            $lst_.LargeImageList = $lst_image

            foreach ($column in $columns.Keys) {
                [void]$lst_.Columns.Add($column, $columns[$column])
            }

            if ($lst_.Columns.Count -gt 1) {
                $lst_.Columns[1].TextAlign = "Right"
            }
        }
        "DataGridView" {
            $lst_                             = New-Object System.Windows.Forms.DataGridView
            $lst_.ColumnHeadersHeightSizeMode = 'DisableResizing'

            foreach($column in $columns.Keys) {
                $addColumn            = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
                $addColumn.HeaderText = $column
                $addColumn.Name       = $column
                $addColumn.ReadOnly   = $true
                $addColumn.Width      = $columns[$column]

                $lst_.Columns.AddRange($addColumn)
            }
        }
    }

    $lst_.Name = $name

    $lst_.Width  = $width
    $lst_.Height = $height
    $lst_.Top    = $top
    $lst_.Left   = $left
    
    if ($mnuContext) {
        $lst_.ContextMenuStrip = $mnuContext
    }
    
    foreach ($property in $properties.Keys) {
        $lst_.$property = $properties[$property]
    }

    foreach ($event_ in $events.Keys) {
        $lst_."Add_$event_"($events[$event_])
    }

    if (-not $labelText) {
        return $lst_
    }

    $lbl_ = InterfaceLabel -labelText $labelText -top ($top - 18) -left $left -width $width -height 15 

    return @($lst_, $lbl_)
}

function InterfaceLabel {
    param (
        [string]$labelText,
        [int]$top,
        [int]$left,
        [int]$width,
        [int]$height
    )

    $lbl_        = New-Object System.Windows.Forms.Label
    $lbl_.Top    = $top
    $lbl_.Left   = $left
    $lbl_.Width  = $width
    $lbl_.Height = $height
    $lbl_.Text   = $labelText

    return $lbl_
} 

function InterfacePanel {
    param (
        [string]$labelText,
        [int]$top,
        [int]$left,
        [int]$width,
        [int]$height
    )

    $pnl_             = New-Object System.Windows.Forms.Panel
    $pnl_.Top         = $top
    $pnl_.Left        = $left
    $pnl_.Width       = $width
    $pnl_.Height      = $height
    $pnl_.BorderStyle = "FixedSingle"

    if (-not $labelText) { 
        return $pnl_ 
    }

    $lbl_ = InterfaceLabel -labelText $labelText -top ($top - 8) -left ($left + 5) -width ($labelText.Length * 7) -height 20

    return @($pnl_, $lbl_)
}

function InterfaceContextMenu {
    param (
        [object[]]$items 
    )
    
    $mnu_ = New-Object System.Windows.Forms.ContextMenuStrip
    
    foreach ($item in $items) {
        
        $text = $item[0]
        if ($text -eq "-") {
            [void]$mnu_.Items.Add("-")
            continue
        }
        $icon = $item[1]
        $function = $item[2]

        $mnu_.Items.Add($text, (InterfaceGetImage $icon)).Add_Click($function)
    }
    
    return $mnu_
}

function InterfaceControl {
    param (
        [string]$type,
        [string]$name = "",
        [string]$labelText = "",
        [int]$top,
        [int]$left,
        [int]$width,
        [int]$height = 20,
        [int]$min = 0,
        [int]$max = 999999999,
        [hashtable]$properties = @{},
        [string]$tag,
        [hashtable]$events = @{},
        [System.Windows.Forms.ContextMenuStrip]$mnuContext
    )

    switch ($type) {
        "NumericUpDown" {
            $ctr_ = New-Object System.Windows.Forms.NumericUpDown
            $ctr_.Minimum = $min
            $ctr_.Maximum = $max
        }
        "MaskedTextBox" {
            $ctr_ = New-Object System.Windows.Forms.MaskedTextBox
        }
        "TextBox" {
            $ctr_ = New-Object System.Windows.Forms.TextBox
        }
        "DateTimePicker" {
            $ctr_ = New-Object System.Windows.Forms.DateTimePicker
        }
        "ComboBox" {
            $ctr_ = New-Object System.Windows.Forms.ComboBox
            $ctr_.DropDownStyle = "DropDownList"
        }
    }
    
    $ctr_.Top    = $top
    $ctr_.Left   = $left
    $ctr_.Width  = $width
    $ctr_.Height = $height
    $ctr_.Tag    = $tag
    if ($name) {
        $ctr_.Name = $name
    }

    if ($mnuContext) {
        $ctr_.ContextMenuStrip = $mnuContext
    }

    foreach ($property in $properties.Keys) {
        $ctr_.$property = $properties[$property]
    }

    $ctr_.Add_MouseEnter({Tip $this})
    $ctr_.Add_MouseLeave({Tip $this -close})

    foreach ($event_ in $events.Keys) {
        $ctr_."Add_$event_"($events[$event_])
    }

    if (-not $labelText) {
        return $ctr_
    }

    $lbl_ = InterfaceLabel -labelText $labelText -top ($top - 20) -left $left -width $width -height 20

    return @($ctr_, $lbl_)
}

function InterfaceShowForm {
    param (
        [string]$title,
        [scriptblock]$start,
        [scriptblock]$close
    )

    $frm_main.Add_Load({
        ShowMessage "Iniciando interface"
        (Get-Host).UI.RawUI.WindowTitle = "$title - MANTER ESTA JANELA ABERTA"
        & $start
    })

    $frm_main.Add_Shown({
        $this.Activate()
        $this.WindowState = [System.Windows.Forms.FormWindowState]::Normal
        ShowMessage "Exibindo interface"
    })

    $frm_main.Add_Closing({
        (Get-Host).UI.RawUI.WindowTitle = "FECHE ESTA JANELA"
        if ($close) {
            & $close
        }
    }.GetNewClosure())
    
    [void]$frm_main.ShowDialog()
}
    
function InterfaceButtonHover {
    param (
        [System.Windows.Forms.PictureBox]$btn_,
        [switch]$mouseEnter
    )
    
    $image = if ($mouseEnter) {"$($btn_.Name)_hover"} else {$btn_.Name}

    $btn_.Image = InterfaceGetImage $image
}

function InterfaceMinimize {
    param (
        [switch]$hideTaskBar
    )

    $frm_main.WindowState = [System.Windows.Forms.FormWindowState]::Minimized

    if ($hideTaskBar) {
        $frm_main.ShowInTaskbar = $false
    }
}

function InterfaceRestore {

    $frm_main.WindowState = [System.Windows.Forms.FormWindowState]::Normal

    if (-not $frm_main.ShowInTaskbar) {
        $frm_main.ShowInTaskbar = $true
    } 
}

function Tip {
    param (
        [System.Windows.Forms.Control]$ctr_,
        [switch]$close
    )
    
    if ($close) {
        $tip_.Hide($ctr_)
    } else {
        $tip_.Show($ctr_.Tag, $ctr_)
    }
    
}

function InterfaceClose {
    [void]$frm_main.Close()    
}

function InterfaceGetImage {
    param(
        [string]$image
    )
    
    $extension = if ($image -like "*.*") { "" } else { ".png" }

    $fullPath = Join-Path $fldImages "$image$extension"
    $key = $fullPath.ToLower()
    
    if ($global:imageCache.ContainsKey($key)) {
        return $global:imageCache[$key]
    }
    
    try {
        if (Test-Path $fullPath) {
            $img_ = [System.Drawing.Image]::FromFile($fullPath)
            $global:imageCache[$key] = $img_
            return $img_
        } else {
            Write-Warning "Imagem não encontrada: $fullPath"
            return $null
        }
    }
    catch {
        Write-Error "Erro ao carregar imagem $fullPath : $($_.Exception.Message)"
        return $null
    }
}

function InterfacePosition {
    param (
        [System.Windows.Forms.Control]$ctr_,
        [string]$option,
        [int]$offset = 0,
        [int]$targetWidth = 0,
        [int]$targetHeight = 0
    )
    
    switch ($option) {
        "top" {
            if (-not $targetWidth) { $targetWidth = $ctr_.Height}

            $position = $ctr_.Top - $offset - $targetWidth
        }
        "bottom" {
            $position = $ctr_.Bottom + $offset - $targetWidth
        }
        "left" {
            if (-not $targetWidth) { $targetWidth = $ctr_.Width}

            $position = $ctr_.Left - $offset - $targetWidth
        }
        "right" {
            $position = $ctr_.Right + $offset - $targetWidth
        }
        "center" {
            $position = $ctr_.Top + ($ctr_.Height - $targetHeight) / 2
        }
        "centerHorizontal" {
            $position = $ctr_.Left + ($ctr_.Width / 2) - ($targetWidth / 2)
        }
        "width" {
            $position = $ctr_.Width - ($offset * 2)
        }
        "innerCorner" {
            $position = @{
                left = $ctr_.Width - $targetWidth - $offset;
                top = $ctr_.Height - $targetHeight - $offset
            }
        }
    }

    return $position

}

function InterfaceTabs {
    param(
        [string]$type,
        [string[]]$names,
        [string[]]$texts,
        [int]$width,
        [int]$height,
        [int]$top,
        [int]$left,
        [hashtable]$properties = @{},
        [hashtable]$events = @{}
    )

    switch ($type) {
        "control" {
            $tbc_        = New-Object System.Windows.Forms.TabControl
            $tbc_.Width  = $width
            $tbc_.Height = $height
            $tbc_.Top    = $top
            $tbc_.Left   = $left

            foreach ($property in $properties.Keys) {
                $tbc_.$property = $properties[$property]
            }

            foreach ($event_ in $events.Keys) {
                $tbc_."Add_$event_"($events[$event_])
            }

            return $tbc_
        }
        "pages" {
            $tbps_ = @()
    
            for ($i = 0; $i -lt $names.Count; $i++) {
                $tbp_          = New-Object System.Windows.Forms.TabPage
                $tbp_.TabIndex = $i
                $tbp_.Text     = $texts[$i]
                $tbp_.Name     = $names[$i]
                $tbps_        += $tbp_
            }
    
            return $tbps_
        }
    }
}

function InterfaceCustomProperty {
    param (
        [System.Windows.Forms.Control]$ctr_,
        [string]$name,
        $value
    )

    $ctr_ | Add-Member -MemberType NoteProperty -Name $name -Value $value -Force
}

function InterfaceSplashScreen {
    param (
        [switch]$onlyLabel,
        [int]$width = 400,
        [int]$height = 150
    )
    
    $lbl_             = New-Object System.Windows.Forms.Label
    $lbl_.Width       = $width
    $lbl_.Height      = $height
    $lbl_.BackColor   = "#0c5c94"
    $lbl_.ForeColor   = "White"
    $lbl_.BorderStyle = 'Fixed3D'
    $lbl_.TextAlign   = [System.Drawing.ContentAlignment]::MiddleCenter
    $lbl_.Font        = New-Object System.Drawing.Font('Segoe UI Light', 24)
    $lbl_.Text        = "⏳`nAguarde..."

    if ($onlyLabel) {
        $lbl_.Top  = ($FRM_MAIN_HEIGHT - $height) / 2 
        $lbl_.Left = ($FRM_MAIN_WIDTH - $width) / 2 

        return $lbl_
    }

    $frm_                 = New-Object System.Windows.Forms.Form
    $frm_.Width           = $width
    $frm_.Height          = $height
    $frm_.StartPosition   = "CenterScreen"
    $frm_.FormBorderStyle = 'None'
    $frm_.ShowInTaskbar   = $false

    return @($frm_, $lbl_)

}

function InterfaceCompanyButtons {

    $params = @{
        size = $BTN_IMAGE_BIG_WIDTH * 0.75;
        top = $FRM_MAIN_HEIGHT - $MARGIN_BOTTOM - ($BTN_IMAGE_BIG_WIDTH * 0.75);
        left = ($MARGIN_LEFT - ($BTN_IMAGE_BIG_WIDTH * 0.75)) / 2;
        name = "department";
        function = {
            InterfaceMinimize
            Start-Process $DEPARTMENT_SITE
        };
        tag = "Ir para a página de $DEPARTMENT_NAME"
    }
    $btn_department = InterfaceButtonImage @params

    $params = @{
        size = $BTN_IMAGE_BIG_WIDTH;
        top = InterfacePosition $btn_department "top" $PADDING_OUTER $BTN_IMAGE_BIG_WIDTH;
        left = ($MARGIN_LEFT - $BTN_IMAGE_BIG_WIDTH) / 2;
        name = "company";
        function = {
            InterfaceMinimize
            Start-Process $COMPANY_SITE
        };
        tag = "Ir para a página de $COMPANY_NAME"
    }
    $btn_company = InterfaceButtonImage @params

    return @($btn_company, $btn_department)

}

function InterfaceRButton {
    param(
        [Parameter(Mandatory = $false)]
        [ValidateSet('TopLeft', 'BottomRight')]
        [string]$position = 'TopLeft',

        [int]$size = $BTN_IMAGE_BIG_WIDTH
    )

    switch ($position) {
        "TopLeft" {
            $top  = $MARGIN_TOP;
            $left = ($MARGIN_LEFT - $BTN_IMAGE_BIG_WIDTH) / 2;
        }
        "BottomRight" {
            $top  = $FRM_MAIN_HEIGHT - $MARGIN_BOTTOM - $size;
            $left = $FRM_MAIN_WIDTH - ($MARGIN_RIGHT + $size) / 2;
        }
    }
    $params = @{
        size = $size;
        top = $top;
        left = $left;
        name = "r";
        tag = "O programa R não está instalado.`r`nClique aqui para ir para a página de download.";
        function = {
            InterfaceMinimize
            Start-Process "https://cran.r-project.org/bin/windows/base/"
        }
    }

    $btn_r = InterfaceButtonImage @params

    return $btn_r

}