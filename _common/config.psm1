function ShowMessage {
    param (
        $msg,
        $title = $null,
        [int]$width = 100,
        $bgColor = "DarkBlue",
        $fgColor = "White"
    )
        
    $lines = $msg -split "`r`n"
    $lines = $lines | ForEach-Object {
        if ($_.Length -gt $width) {
            $_.Substring(0, $width - 3) + "..."
        } else {
            $_
        }
    }

    $titleWidth = ($title | Measure-Object -Property Length -Maximum).Maximum + 2
    $msgWidth   = ($lines | Measure-Object -Property Length -Maximum).Maximum + 2
    $lineWidth  = ($msgWidth, $titleWidth | Measure-Object -Maximum).Maximum

    Write-Host "╭$('─' * $lineWidth)╮" -BackgroundColor $bgColor -ForegroundColor $fgColor
    if (![string]::IsNullOrEmpty($title)) {
        Write-Host "│ $title $(' ' * ($lineWidth - $title.Length - 2))│" -BackgroundColor $bgColor -ForegroundColor $fgColor
        Write-Host "├$('─' * $lineWidth)┤" -BackgroundColor $bgColor -ForegroundColor $fgColor
    }
    foreach ($line in $lines) {
        Write-Host "│ $line$(' ' * ($lineWidth - $line.Length - 2)) │" -BackgroundColor $bgColor -ForegroundColor $fgColor
    }
    Write-Host "╰$('─' * $lineWidth)╯`n" -BackgroundColor $bgColor -ForegroundColor $fgColor
    
}

function CheckRInstallation {

    $pathRegister = @(
        "HKLM:\SOFTWARE\R-core\R\",
        "HKCU:\SOFTWARE\R-core\R\"
    )
    $commonFolders = @(
        "$env:LOCALAPPDATA\Programs\R",
        "C:\Program Files\R",
        "C:\Program Files (x86)\R",
        "C:\R"
    )
    $fleExeR = $null

    foreach ($path in $pathRegister) {
        if (Test-Path -Path $path) {
            $instalacoes = Get-ChildItem -Path $path -ErrorAction SilentlyContinue | Sort-Object Name -Descending
            foreach ($inst in $instalacoes) {
                try {
                    $RFolder = Get-ItemPropertyValue -Path $inst.PSPath -name InstallPath -ErrorAction Stop
                    $possivelExe = Join-Path $RFolder "bin\Rscript.exe"
                    if (Test-Path $possivelExe) {
                        $fleExeR = $possivelExe
                        break
                    }
                } catch {}
            }
            if ($fleExeR) { break }
        }
    }

    if (-not $fleExeR) {
        foreach ($folder in $commonFolders) {
            if (Test-Path $folder) {
                $possiveis = Get-ChildItem -Path $folder -Directory | Sort-Object Name -Descending
                foreach ($dir in $possiveis) {
                    $possivelExe = Join-Path $dir.FullName "bin\Rscript.exe"
                    if (Test-Path $possivelExe) {
                        $fleExeR = $possivelExe
                        break
                    }
                }
            }
            if ($fleExeR) { break }
        }
    }

    if ($fleExeR -and (Test-Path $fleExeR)) {
        ShowMessage "R instalado em ${fleExeR}" 
        $installedR = $true
    } else {
        ShowMessage "R não instalado" -bgColor "Red" -fgColor "White"
        $installedR = $false
        [System.Windows.Forms.MessageBox]::Show(
            "Para gerar os arquivos, será necessário instalar o programa R, disponível no site https://cran.r-project.org/bin/windows/base/", "Erro", 0, 16
        ) 
        Start-Process "https://cran.r-project.org/bin/windows/base/"
    }

    Set-Variable -Name installedR -Scope Global -Value $installedR
    Set-Variable -Name fleExeR -Scope Global -Value $fleExeR

    if (-not $installedR) {
        $btn_r.Show()
    } else {
        $btn_r.Hide()
    }
} 

function CheckFolder {
    param(
        [string[]]$paths
    )

    foreach ($path in $paths) {
        if (!(Test-Path -Path $path)) {
            New-Item -Path $path -ItemType "Directory" -Force | Out-Null
        }
    }
}

function GetConfigJSON {

    if (-not (Get-Variable -Name scriptName -ErrorAction SilentlyContinue)) {
        throw "A variável scriptName não está definida. Execute a função main primeiro."
    }

    if ([string]::IsNullOrWhiteSpace($scriptName)) {
        throw "A variável scriptName está vazia. Execute a função main primeiro."
    }

    try {
        if (Test-Path $fleConfig) {
            $config = Get-Content $fleConfig -Raw -Encoding UTF8 -ErrorAction Stop | ConvertFrom-Json
            if ($null -eq $config -or $config -eq "") {
                $config = [pscustomobject]@{}
            }
        } else {
            $config = [pscustomobject]@{}
        }
    } catch {
        Write-Warning "Erro ao ler o arquivo de configuração. Criando novo..."
        $config = [pscustomobject]@{}
    }

    if (-not ($config.PSObject.Properties.Name -contains $scriptName)) {
        $config | Add-Member -NotePropertyName $scriptName -NotePropertyValue ([pscustomobject]@{}) -Force
    }

    return $config
}

function ConfigJSON {
    param (
        [string]$key = $null,
        $value = $null,

        [Parameter(Mandatory = $false)]
        [ValidateSet("all", "remove", "get", "set")]
        [string]$option,

        [switch]$append,
        [switch]$show
    )

    $config = GetConfigJSON

    if ($show) { 
        ShowMessage ($config.$scriptName | ConvertTo-Json) "Configurações do Script $scriptName"
        return $null 
    }

    if ($option -eq "all") {
        return $config.$scriptName
    } 
    
    $containsKey = $config.$scriptName.PSObject.Properties.Name -contains $key

    if ($option -eq "remove") {
        if ($containsKey) {
            $config.$scriptName.PSObject.Properties.Remove($key)
            SaveConfigJSON $config
            return
        } else {
            return $null
        }
    } 
    
    # If $value is null, consider option as "get"
    if ([string]::IsNullOrEmpty($option) -and $null -eq $value) {

        if ($containsKey) {
            $valueReturn = $config.$scriptName.$key
    
            if ($valueReturn -is [array]) {
                return $valueReturn -join "`r`n"
            }
            return $valueReturn
        }
        return $null
    }
    
    # If has $key and $value, option is "set"
    if ($append -and $containsKey) {
        
        $valueExistent = $config.$scriptName.$key
        
        if ($valueExistent -is [array] -and $value -is [array]) {
            $config.$scriptName | Add-Member -NotePropertyName $key -NotePropertyValue (@($valueExistent) + @($value)) -Force
        }
        elseif ($valueExistent -is [hashtable] -and $value -is [hashtable]) {
            foreach ($key in $value.Keys) {
                $valueExistent[$key] = $value[$key]
            }
            $config.$scriptName | Add-Member -NotePropertyName $key -NotePropertyValue $valueExistent -Force
        }
        else {
            $config.$scriptName | Add-Member -NotePropertyName $key -NotePropertyValue @($valueExistent, $value) -Force
        }
    } else {
        $config.$scriptName | Add-Member -NotePropertyName $key -NotePropertyValue $value -Force
    }

    SaveConfigJSON $config
    return
    
}

function SaveConfigJSON {
    param (
        [pscustomobject]$config
    )
    
    try {
        $json = $config | ConvertTo-Json -Depth 10
        $utf8NoBom = New-Object System.Text.UTF8Encoding $False
        [System.IO.File]::WriteAllText($fleConfig, $json, $utf8NoBom)
    } catch {
        Write-Error "Falha ao salvar o arquivo de configuração: $_"
    }
}

function RunR {
    param (
        [string]$script,
        [array]$arguments = $null
    )

    $line1 = "Executando R em ${fleExeR}"
    $line2 = "Script: ${script}"
    $line3 = "Argumentos: $($arguments -join ', ')"

    ShowMessage ($line1, $line2, $line3) "Iniciando Script R" -fgColor Black -bgColor White

    $argumentsList = @("`"$script`"") + $(if ($null -ne $arguments) { $arguments | ForEach-Object { "`"$_`"" } } else { @() }) -join " "

    Start-Process -FilePath "${fleExeR}" -ArgumentList "`"--vanilla`"", $argumentsList -NoNewWindow -Wait

}

function IsUserRunningR {

    $statusScriptR = ConfigJSON -key "resultado_geracao" -option "get"

    if ($statusScriptR -ne "running") {
        return $false
    }

    $fleLogR     = ConfigJSON -key "arquivo_log_R" -option "get"
    $currentUser = ($fleLogR -split "_")[4]

    $msg = "O usuário $currentUser está executando o script.`n`nÉ recomendado aguardar o fim da execução.`n`nDeseja continuar mesmo assim?"
    $result = [System.Windows.Forms.MessageBox]::Show($msg, "Alerta", 4, 16)
    if ($result -eq "No") {
        return $true
    } else {
        return $false
    }

}

function SetScriptConstants {
    <#
    .SYNOPSIS
        Defines global variables for the application
    
    .DESCRIPTION
        Creates global variables from a hashtable
    
    .PARAMETER constants
        Hashtable containing the names and values of the constants
    
    .EXAMPLE
        SetScriptConstants @{
            REGEX_EMAIL = '^[\w-\.]+@([\w-]+\.)+[\w-]{2,4}$'
            BTN_WIDTH = 120
        }
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$constants
    )
    
    foreach ($key in $constants.Keys) {
        $params = @{
            Name  = $key
            Value = $constants[$key]
            Scope = 'Global'
            Force = $true
        }
        
        New-Variable @params -ErrorAction Stop
    }
}

function GetEnvConfig {
    <#
    .SYNOPSIS
        Obtém configurações do arquivo .Renviron
    
    .DESCRIPTION
        Lê as configurações do arquivo .Renviron localizado na pasta _common.
        Pode retornar todas as configurações ou uma configuração específica.
    
    .PARAMETER ConfigName
        Nome da configuração específica a ser retornada.
        Se não for fornecido, retorna todas as configurações.
    
    .EXAMPLE
        GetEnvConfig
        Retorna todas as configurações do arquivo .Renviron
    
    .EXAMPLE
        GetEnvConfig -ConfigName "DATABASE_URL"
        Retorna o valor da configuração DATABASE_URL
    #>
    param(
        [Parameter(Mandatory = $false)]
        [string]$configName
    )

    $envFile = ".Renviron"
    $envPath = Join-Path $fldCommon $envFile
    
    if (-not (Test-Path $envPath)) {
        Write-Error "Arquivo não encontrado em: $envPath"
        return $null
    }

    try {
        $configs = @{}
        Get-Content $envPath | ForEach-Object {
            if ($_ -match '^\s*([^#][^=]+)=(.*)$') {
                $key = $matches[1].Trim()
                $value = $matches[2].Trim()
                $configs[$key] = $value
            }
        }

        if ([string]::IsNullOrEmpty($configName)) {
            return $configs
        }
        
        if ($configs.ContainsKey($configName)) {
            return $configs[$configName]
        }
        else {
            Write-Warning "Configuração '$configName' não encontrada no arquivo $envFile"
            return $null
        }
    }
    catch {
        Write-Error "Erro ao ler o arquivo ${envFile}: $_"
        return $null
    }
}

function main {
    param(
        [string]$scriptName
    )

    Set-Variable -Name fldSource -Scope Global -Value (Join-Path $fldRoot "_fontes")
    Set-Variable -Name fldImages -Scope Global -Value (Join-Path $fldCommon "images")
    Set-Variable -Name fldLog -Scope Global -Value (Join-Path $fldCommon "log")

    Set-Variable -Name fleConfig -Scope Global -Value (Join-Path $fldCommon "config.json")
    $global:imageCache = @{}

    Set-Variable -Name scriptName -Scope Global -Value $scriptName.ToUpper()
    $userName   = ($env:USERNAME).ToUpper()
    $hostName   = (hostname).ToUpper()
    $date       = Get-Date -f "yyyy-MM-dd"
    $time       = Get-Date -f "HH-mm-ss"
    $fleLog     = ("Log_${scriptName}_${date}_${time}_${userName}-${hostName}_pwsh.log").Replace(" ", "")
    $Transcript = Join-Path -Path $fldLog "$fleLog" 

    $DebugPreference       = "Continue"
    $InformationPreference = "Continue"

    CheckFolder -Paths @($fldLog)

    Write-Host "`n ───────────────────────────────────────────"
    Write-Host "  Script PowerShell iniciado"
    Write-Host " ───────────────────────────────────────────`n"
    
    Start-Transcript -Path $Transcript
    
    Import-Module (Join-Path $fldCommon "interface.psm1") -Global

    ShowMessage "Log iniciado - $fleLog" -fgColor "Black"
    
    ShowMessage ($PSVersionTable | Format-Table | Out-String) "PowerShell"
    
    ShowMessage (Get-ExecutionPolicy -List | Format-Table | Out-String) "Permissões"
    
    ShowMessage (Get-ChildItem -Recurse -Path $fldRoot | Format-Table | Out-String) "Arquivos"
    
    CheckRInstallation
    
}