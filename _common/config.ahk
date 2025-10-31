GetFileName(pattern) {

    Loop, Files, %pattern%
    {
        fileName := A_LoopFileName
        break
    }

    return %fileName%
}

VerificarDependencias(itens*) {
    for index, item in itens {
        if !FileExist(item) {
            MsgBox, 16, Erro, O arquivo ou pasta '%item%' não foi localizado.
            return false
        }
    }
    return true
}

VerificarResultadoR(nomeScript := "", only_check := false) {

    FileRead, json, *P65001 ..\_common\config.json
    msg_erro := ""
    resultado_geracao := ""
    
    ; Busca dentro da chave específica do script
    if RegExMatch(json, """" nomeScript """\s*:\s*\{([^}]*)\}", scriptSection) {
        if RegExMatch(scriptSection1, """resultado_geracao""\s*:\s*""(.*?)""", match) {
            resultado_geracao := match1
        }
        if RegExMatch(scriptSection1, """msg_erro""\s*:\s*\[(.*?)\]", match) {
            msg_erro := match1
        }
    }

    ; Se only_check := true, apenas verifica e retorna o resultado e a mensagem de erro, sem mostrar mensagens
    if (only_check)
        return [resultado_geracao, msg_erro]

    ; Se only_check for false (padrão), continua para mostrar mensagens

    ; Mostra mensagens apropriadas com base no resultado
    ; e abre a pasta de logs se houver erro desconhecido

    if (resultado_geracao = "erro") {
        MsgBox, 16, Erro, Ocorreram os seguintes erros na execução e não foram gerados os arquivos.`n`n%msg_erro%`n`nPara mais informações, verifique o arquivo mais recente na pasta _fontes/log.
    } else if (resultado_geracao = "ambos") {
        MsgBox, 48, Alerta, Ocorreram erros na execução, mas os arquivos foram gerados`n`n%msg_erro%`n`nPara mais informações, verifique o arquivo mais recente na pasta _fontes/log
    } else if (resultado_geracao = "sucesso") {
        MsgBox, 64, Sucesso, Script finalizado com sucesso!
    } else {
        MsgBox, 16, Erro, Ocorreu um erro desconhecido na execução.`n`nSerá aberta a pasta dos logs.`n`nVerifique o arquivo mais recente.
        Run, ..\_common\log
    }

    Return
}

LocalizarRPath() {
    ; --- 1. Procura no Registro (HKLM e HKCU) de forma combinada ---
    hives := ["HKEY_LOCAL_MACHINE", "HKEY_CURRENT_USER"]
    for _, hive in hives
    {
        versoes := []
        Loop, Reg, % hive "\SOFTWARE\R-core\R", K  ; O "K" garante que apenas subchaves sejam lidas
        {
            versoes.Push(A_LoopRegName)
        }

        if (versoes.Length() = 0)
            continue

        versoes := OrdenaVersoes(versoes, true) ; Ordena decrescente (mais novo primeiro)

        for _, v in versoes  ; Loop simplificado
        {
            RegRead, regPath, % hive "\SOFTWARE\R-core\R\" v, InstallPath
            if (regPath && FileExist(regPath "\bin\Rscript.exe"))
                return regPath
        }
    }

    ; --- 2. Procura nas pastas padrão do Windows ---
    EnvGet, AppsUsuario, LocalAppData
    EnvGet, ArquivosProgramas, ProgramFiles
    EnvGet, ArquivosProgramasX86, ProgramFiles(x86) ; Lida com sistemas 64-bit

    rdirs := [AppsUsuario "\Programs\R", ArquivosProgramas "\R"]
    if (ArquivosProgramasX86) ; Adiciona a pasta x86 apenas se ela existir
        rdirs.Push(ArquivosProgramasX86 "\R")
    rdirs.Push("C:\R")

    for _, dir in rdirs 
    {
        if !InStr(FileExist(dir), "D") ; Checagem mais robusta se o diretório existe
            continue

        versoesNomes := []
        versoesPaths := {} ; Usar um objeto (mapa) para associar versão -> path

        ; Procura por pastas que correspondem ao padrão "R-versão"
        Loop, Files, % dir "\R-*", D 
        {
            versaoNum := SubStr(A_LoopFileName, 3) ; Extrai "4.3.1" de "R-4.3.1" (mais rápido que RegEx)
            versoesNomes.Push(versaoNum)
            versoesPaths[versaoNum] := A_LoopFileFullPath
        }

        if (versoesNomes.Length() > 0)
        {
            versoesNomes := OrdenaVersoes(versoesNomes, true) ; Ordena decrescente
            
            for _, vNum in versoesNomes
            {
                vdir := versoesPaths[vNum]
                if FileExist(vdir "\bin\rRscript.exe")
                    return vdir ; Encontrou o mais novo nesta pasta, retorna e encerra a função
            }
        }
    }
    
    ; --- 3. MELHORIA: Fallback procurando no PATH do sistema ---
    if rpath := BuscaRNoPath()
        return rpath

    return "" ; Se não encontrou nada, retorna vazio
}

OrdenaVersoes(versoes, reverse := false) {
    ; Este método usa o comando Sort nativo do AHK com a opção "V",
    ; que é otimizada para classificar números de versão.
    
    versoesString := ""
    For _, v in versoes
        versoesString .= v "`n"

    ; 1. Primeiro, construímos a string de opções em uma variável.
    options := "V"
    if (reverse)
        options .= " R" ; Adiciona o " R" para ordenação reversa.

    ; 2. Em seguida, usamos a variável com o comando Sort.
    ;    A sintaxe %options% é a forma correta de usar uma variável neste parâmetro.
    Sort, versoesString, %options%
    
    ; StrSplit remove entradas vazias por padrão, o que lida com a nova linha final
    return StrSplit(versoesString, "`n", "`r")
}

BuscaRNoPath() {
    ; Procura por um diretório bin do R no PATH do sistema.
    
    ; 1. Primeiro, o conteúdo da variável de ambiente PATH é obtido e armazenado em uma variável local.
    EnvGet, systemPath, PATH

    ; 2. Em seguida, usamos essa variável no comando Loop, Parse.
    ;    O comando agora opera sobre o conteúdo de 'systemPath'.
    Loop, Parse, systemPath, ; (ponto e vírgula é o delimitador)
    {
        ; 3. Regex aprimorado para capturar o diretório base diretamente em 'match1'.
        if RegExMatch(A_LoopField, "i)^(.*\\R-[0-9\.]+[\\/])bin$", match)
        {
            if FileExist(A_LoopField "\Rscript.exe")
            {
                ; Retorna diretamente o grupo capturado (o caminho base, ex: "C:\Program Files\R\R-4.3.1")
                return match1
            }
        }
    }
    return "" ; Não encontrou no PATH
}