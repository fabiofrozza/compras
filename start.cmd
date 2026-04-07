@CHCP 65001 >NUL
@TITLE PREPARANDO SERVIDOR LOCAL. NÃO FECHE ESTA JANELA
@FOR /F %%a IN ('ECHO PROMPT $e^| CMD') DO @SET "esc=%%a"
@SET "separador=───────────────────────────────────────────"

@ECHO %esc%[34m                                                        
@ECHO   ██████╗ ██████╗        ███╗   ██╗ █████╗ ███╗   ███╗███████╗
@ECHO  ██╔════╝██╔═══██╗       ████╗  ██║██╔══██╗████╗ ████║██╔════╝
@ECHO  ██║     ██║   ██║       ██╔██╗ ██║███████║██╔████╔██║█████╗  
@ECHO  ██║     ██║   ██║       ██║╚██╗██║██╔══██║██║╚██╔╝██║██╔══╝  
@ECHO  ╚██████╗╚██████╔╝██╗    ██║ ╚████║██║  ██║██║ ╚═╝ ██║███████╗
@ECHO   ╚═════╝ ╚═════╝ ╚═╝    ╚═╝  ╚═══╝╚═╝  ╚═╝╚═╝     ╚═╝╚══════╝%esc%[0m

@PROMPT=$s$e[34m$v$_$s$d$s$b$s$t$h$h$h$_$s$p$_$_$s$e[0m%separador%$_$s$e[34m$sIniciando servidor...$e[0m$_$s%separador%$_$s$e[30m
ECHO OFF

@TITLE SERVIDOR EM EXECUÇÃO. NÃO FECHE ESTA JANELA

IF NOT EXIST ".\node_modules\" (
    ECHO  %esc%[97m%separador%
    ECHO.%esc%[34m  Instalando dependências pela primeira vez...%esc%[0m
    ECHO. %separador%
    ECHO  %esc%[93m%esc%[41m  Isso pode demorar alguns minutos. Aguarde...%esc%[0m
    ECHO  %separador%
    CALL bin\npm install
    ECHO  %separador%
)

SET "texto=Aguarde..."
ECHO  %esc%[97m%separador%& ECHO.%esc%[34m  %texto%%esc%[0m& ECHO. %separador%%esc%[93m%esc%[41m
ECHO   A primeira execução pode demorar...
ECHO   A permissão ao firewall é opcional (se for Desktop Gerenciado, não permita)%esc%[0m
ECHO  %separador%%esc%[93m%esc%[44m
".\bin\node.exe" ".\src\server.js"

ECHO %esc%[41m%esc%[93m

SET "texto=Encerrando..."
ECHO  %separador%
ECHO   %texto%
ECHO  %separador%

@TITLE ESTA JANELA JÁ PODE SER FECHADA

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