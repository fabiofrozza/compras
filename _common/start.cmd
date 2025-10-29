@CHCP 65001 >NUL
@TITLE PREPARANDO PARA EXECUTAR POWERSHELL
@FOR /F %%a IN ('ECHO PROMPT $e^| CMD') DO @SET "esc=%%a"
@SET "separador=───────────────────────────────────────────"

@ECHO %esc%[34m                                                        
@ECHO   ██████╗ ██████╗        ███╗   ██╗ █████╗ ███╗   ███╗███████╗
@ECHO  ██╔════╝██╔═══██╗       ████╗  ██║██╔══██╗████╗ ████║██╔════╝
@ECHO  ██║     ██║   ██║       ██╔██╗ ██║███████║██╔████╔██║█████╗  
@ECHO  ██║     ██║   ██║       ██║╚██╗██║██╔══██║██║╚██╔╝██║██╔══╝  
@ECHO  ╚██████╗╚██████╔╝██╗    ██║ ╚████║██║  ██║██║ ╚═╝ ██║███████╗
@ECHO   ╚═════╝ ╚═════╝ ╚═╝    ╚═╝  ╚═══╝╚═╝  ╚═╝╚═╝     ╚═╝╚══════╝%esc%[0m

@SET "script="
@FOR %%f IN (*.ps1) DO @IF NOT DEFINED script @SET "script=%%f"
@IF NOT DEFINED script (
    @ECHO  %separador%
    @ECHO  Nenhum arquivo .ps1 encontrado na pasta.
    @ECHO  Não é possível executar o script. Encerrando...
    @ECHO  %separador%
    @PAUSE
    @EXIT /B 1
)
@FOR %%f IN (*.ps1) DO @IF NOT "%%f"=="%script%" (
    @ECHO  %separador%
    @ECHO  Mais de um arquivo .ps1 encontrado. Usando o primeiro: "%script%"
    @ECHO  Certifique-se de que haja apenas um arquivo .ps1 na pasta.
    @ECHO  %separador%
    @PAUSE
)

@PROMPT=$s$e[34m$v$_$s$d$s$b$s$t$h$h$h$_$s$p$_$_$s$e[0m%separador%$_$s$e[34m$sLocalizando versão do Powershell$e[0m$_$s%separador%$_$s

:: Verifica qual versão do PowerShell está disponível
@SET "ps="
@SET "found="

:: 1. Verifica pwsh.exe no PATH
%ComSpec% /C "WHERE pwsh.exe" >NUL 2>&1
@IF %ERRORLEVEL% EQU 0 (
    @SET "ps=pwsh.exe"
    @SET "found=1"
)

:: 2. Se não encontrou, verifica no caminho específico do PowerShell 7
@IF NOT DEFINED found (
    @IF EXIST "C:\Program Files\PowerShell\7\pwsh.exe" (
        @SET "ps=C:\Program Files\PowerShell\7\pwsh.exe"
        @SET "found=1"
    )
)

:: 3. Se não encontrou, verifica no caminho local (no caso de Desktop Gerenciado)
@IF NOT DEFINED found (
    @IF EXIST "%LocalAppData%\PowerShell\7\pwsh.exe" (
        @SET "ps=%LocalAppData%\PowerShell\7\pwsh.exe"
        @SET "found=1"
    )
)

:: 4. Se ainda não encontrou, usa o powershell.exe padrão
@IF NOT DEFINED found (
    @SET "ps=powershell.exe"
)

@SET "texto=Desbloqueia script"
@PROMPT=$s%separador%$_$s$e[34m$s%texto%$e[0m$_$s%separador%$_$s
"%ps%" -Command Unblock-File -Path "'%script%'"

@SET "texto=Exibe permissões atuais"
@PROMPT=$s%separador%$_$s$e[34m$s%texto%$e[0m$_$s%separador%$_$s
"%ps%" -Command Get-ExecutionPolicy -List

@SET "texto=Define permissões"
@PROMPT=$s%separador%$_$s$e[34m$s%texto%$e[0m$_$s%separador%$_$s
"%ps%" -Command Set-ExecutionPolicy Bypass -Scope CurrentUser -Force

@SET "texto=Exibe novamente para controle"
@PROMPT=$s%separador%$_$s$e[34m$s%texto%$e[0m$_$s%separador%$_$s
"%ps%" -Command Get-ExecutionPolicy -List

@SET "texto=Executa script"
@PROMPT=$s%separador%$_$s$e[34m$s%texto%$e[0m$_$s%separador%$_$s
"%ps%" -ExecutionPolicy Bypass -File "%script%" -WindowStyle Normal -NoExit

@ECHO OFF
ECHO %esc%[41m%esc%[93m

SET "texto=Script PowerShell encerrado"
ECHO  %separador%
ECHO   %texto%
ECHO  %separador%

:: Gera um número entre 1 e 4
SET /A choice=(%RANDOM%*4/32768)+1

IF %choice% EQU 1 (
    GOTO Mensagem1
) ELSE IF %choice% EQU 2 (
    GOTO Mensagem2
) ELSE IF %choice% EQU 3 (
    GOTO Mensagem3
) ELSE IF %choice% EQU 4 (
    GOTO Mensagem4
)

:Mensagem1
ECHO:
ECHO              /´¯/'   '/´¯¯`·¸
ECHO           /'/   /   /      /¨¯\
ECHO          ('(   ´   ´     ¯~/'  ')
ECHO           \                '    /
ECHO            ''   \          _.·´
ECHO             \             (
ECHO:
ECHO              Se você não fechar esta janela,
ECHO                                   você não é brother...
ECHO:
GOTO Fim

:Mensagem2
ECHO:
ECHO                       ,---.           ,---.
ECHO                      / /"`.\.--"""--./,'"\ \
ECHO                      \ \    _       _    / /
ECHO                       `./  / __   __ \  \,'
ECHO                        /    /_O)_(_O\    \
ECHO                        ^|  .-'  ___  `-.  ^|
ECHO                     .--^|       \_/       ^|--.
ECHO                   ,'    \   \   ^|   /   /    `.
ECHO                  /       `.  `--^--'  ,'       \
ECHO               .-"""""-.    `--.___.--'     .-"""""-.
ECHO  .-----------/         \------------------/         \--------------.
ECHO  ^| .---------\         /----------------- \         /------------. ^|
ECHO  ^| ^|          `-`--`--'                    `--'--'-'             ^| ^|
ECHO  ^| ^|                                                             ^| ^|
ECHO  ^| ^|                                                             ^| ^|
ECHO  ^| ^|   _,.-'~'-.,__,.-'~'-.,__,.-'~'-.,__,.-'~'-.,__,.-'~'-.,_   ^| ^|
ECHO  ^| ^|                O ursinho quer te dizer                      ^| ^|
ECHO  ^| ^|             que você já pode                                ^| ^|
ECHO  ^| ^|                       fechar esta janela :)                 ^| ^|
ECHO  ^| ^|  _,.-'~'-.,__,.-'~'-.,__,.-'~'-.,__,.-'~'-.,__,.-'~'-.,_    ^| ^|
ECHO  ^| ^|                                                             ^| ^|
ECHO  ^| ^|                                                             ^| ^|
ECHO  ^| ^|                                                             ^| ^|
ECHO  ^| ^|_____________________________________________________________^| ^|
ECHO  ^|_________________________________________________________________^|
ECHO                     )__________^|__^|__________(
ECHO                    ^|            ^|^|            ^|
ECHO                    ^|____________^|^|____________^|
ECHO                      ),-----.(      ),-----.(
ECHO                    ,'   ==.   \    /  .==    `.
ECHO                   /            )  (            \
ECHO                   `==========='    `===========' hjw
ECHO:
GOTO Fim

:Mensagem3
ECHO:
ECHO                                      _
ECHO    Esta janela você                ,:'/   _..._
ECHO             já pode fechar        // ( `""-.._.'
ECHO       e o doguinho então         \^| /    6\___
ECHO   levar para passear :)          ^|    6       4
ECHO                                  ^|            /
ECHO                                  \_       .--'
ECHO                                  (_'---'`)
ECHO                                  / `'---`()
ECHO                                ,'        ^|
ECHO                ,            .'`          ^|
ECHO                )\       _.-'             ;
ECHO               / ^|    .'`   _            /
ECHO             /` /   .'       '.        , ^|
ECHO            /  /   /           \   ;   ^| ^|
ECHO            ^|  \  ^|            ^|  .^|   ^| ^|
ECHO             \  `"|           /.-' |   | |
ECHO              '-..-\       _.;.._  ^|   ^|.;-.
ECHO                    \    ^<`..^_  )) ^|  .;-. ))
ECHO                    (__.  `  ))-'  \_    ))'
ECHO                        `'--"`  jgs  `"""`
ECHO:
GOTO Fim

:Mensagem4
ECHO:
ECHO                    _....._
ECHO                _.:`.--^|--.`:._
ECHO              .: .'\o  ^| o /'. '.
ECHO             // '.  \ o^|  /  o '.\
ECHO            //'._o'. \ ^|o/ o_.-'o\\
ECHO            ^|^| o '-.'.\^|/.-' o   ^|^|
ECHO            ^|^|--o--o--^>^|
ECHO:
ECHO                Chiudi questa finestra e 
ECHO                                goditi una fetta di pizza...
ECHO:
GOTO Fim

:Fim
PAUSE