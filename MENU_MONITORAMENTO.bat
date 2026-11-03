@echo off
setlocal
title NoFly - Monitoramento de Sedes
cd /d "%~dp0"

:MENU
cls
color 0A
echo ================================================================================
echo             NoFly - MONITORAMENTO DE SEDES (ROTINA DE VISTORIAS)
echo ================================================================================
echo.
echo   [1] Abrir Mapa de Vistorias no Navegador (index.html)
echo   [2] Baixar Relatorio de Vistorias do KoboToolbox (Excel Consolidado)
echo   [3] Atualizar Base do Mapa Web (Processar shapefiles Sede.shp e AREAS.shp)
echo   [4] Publicar Alteracoes no GitHub (Atualizar App do Celular)
echo   [5] Abrir Projeto no QGIS (MONITORAMENTO QUEIROZ.qgz)
echo   [6] Acessar Painel KoboToolbox na Nuvem (Ver Dados Online)
echo   [7] Abrir Painel de Analise de Frequencia e Calor (analise_frequencia.html)
echo   [8] Abrir Relatorio Executivo Online (relatorio.html)
echo   [9] Abrir Manual de Gestao e Geotecnologia (manual.html)
echo   [S] Gerar Relatorio por Sede com Fotos (relatorio_sede.html - salvar PDF)
echo   [C] Abrir Guia da Equipe de Campo (guia_campo.html)
echo   [R] Instalar/atualizar bibliotecas Python (primeiro uso)
echo   [0] Sair
echo.
echo ================================================================================
echo.
choice /C 123456789CRS0 /N /M "Pressione 1, 2, 3, 4, 5, 6, 7, 8, 9, S, C, R ou 0: "

if errorlevel 13 goto OP_SAIR
if errorlevel 12 goto OP_RELATORIO_SEDE
if errorlevel 11 goto OP_DEPS
if errorlevel 10 goto OP_CAMPO
if errorlevel 9 goto OP_MANUAL
if errorlevel 8 goto OP_RELATORIO
if errorlevel 7 goto OP_ANALISE
if errorlevel 6 goto OP_KOBO
if errorlevel 5 goto OP_QGIS
if errorlevel 4 goto OP_GITHUB
if errorlevel 3 goto OP_UPDATE
if errorlevel 2 goto OP_EXCEL
if errorlevel 1 goto OP_OPEN_WEB
goto MENU

:OP_RELATORIO_SEDE
cls
echo ================================================================================
echo             NoFly - RELATORIO POR SEDE (LEVANTAMENTO, TRATATIVAS E FOTOS)
echo ================================================================================
echo.
echo Link para qualquer pessoa (celular ou computador):
echo https://geotecnologiaclealco.github.io/MONITORAMENTO_SEDES/relatorio_sede.html
echo.
echo No relatorio: "Atualizar dados" busca vistorias e fotos novas do Kobo;
echo escolha a sede e clique em "Salvar PDF".
echo.
start "" "https://geotecnologiaclealco.github.io/MONITORAMENTO_SEDES/relatorio_sede.html"
echo Pronto!
timeout /t 4 > nul
goto MENU

:OP_CAMPO
cls
echo ================================================================================
echo             NoFly - GUIA DA EQUIPE DE CAMPO (AVALIADORES)
echo ================================================================================
echo.
echo Abrindo o guia de campo no navegador...
start "" "guia_campo.html"
echo Pronto!
timeout /t 3 > nul
goto MENU

:OP_MANUAL
cls
echo ================================================================================
echo             NoFly - MANUAL DE GESTAO E GEOTECNOLOGIA
echo ================================================================================
echo.
echo Abrindo o manual interativo no navegador...
start "" "manual.html"
echo Pronto!
timeout /t 3 > nul
goto MENU

:OP_RELATORIO
cls
echo ================================================================================
echo             NoFly - RELATORIO ONLINE DE MONITORAMENTO DE SEDES
echo ================================================================================
echo.
echo Link oficial para Celular, Computador e WhatsApp:
echo https://geotecnologiaclealco.github.io/MONITORAMENTO_SEDES/relatorio.html
echo.
echo Abrindo o relatorio online no seu navegador...
start "" "https://geotecnologiaclealco.github.io/MONITORAMENTO_SEDES/relatorio.html"
echo Pronto!
timeout /t 3 > nul
goto MENU

:OP_ANALISE
cls
echo Abrindo Painel de Analise de Frequencia e Mapa de Calor...
start "" "analise_frequencia.html"
echo Pronto!
timeout /t 2 > nul
goto MENU

:OP_OPEN_WEB
cls
echo Abrindo o Mapa de Vistorias no navegador...
start "" "index.html"
echo Pronto!
timeout /t 2 > nul
goto MENU

:OP_EXCEL
cls
color 0A
echo ================================================================================
echo   [2] BAIXANDO RELATORIO DE VISTORIAS DO KOBOTOOLBOX...
echo ================================================================================
echo.
python "scripts\baixar_planilha_visitas_kobo.py"
echo.
echo ================================================================================
pause
goto MENU

:OP_UPDATE
cls
color 0B
echo ================================================================================
echo   [3] ATUALIZANDO BASE DO MAPA WEB (SEDES E AREAS)...
echo ================================================================================
echo.
python "scripts\atualizar_base_web.py"
echo.
echo ================================================================================
pause
goto MENU

:OP_GITHUB
cls
color 0B
echo ================================================================================
echo   [4] PUBLICANDO ATUALIZACOES NO GITHUB (APP DO CELULAR)...
echo ================================================================================
echo.
rem So o branch main vai para o site; em outro branch o envio nao atualiza nada
for /f "tokens=*" %%b in ('git rev-parse --abbrev-ref HEAD') do set "BRANCH=%%b"
if /i not "%BRANCH%"=="main" (
  color 0C
  echo [ERRO] Este computador esta no branch "%BRANCH%", nao no "main".
  echo        Publicar daqui NAO atualiza o site. Volte para o main e tente de novo.
  pause
  goto MENU
)
echo [*] Baixando atualizacoes feitas online (botao "Atualizar dados")...
git pull --rebase --autostash -q origin main
if errorlevel 1 (
  color 0C
  echo [ERRO] Nao foi possivel sincronizar com o GitHub. Verifique a internet e tente novamente.
  pause
  goto MENU
)
echo [*] Sincronizando vistorias do KoboToolbox e atualizando base web...
python "scripts\atualizar_base_web.py"
if errorlevel 1 (
  color 0C
  echo.
  echo [ERRO] A atualizacao da base falhou. Nada foi publicado.
  pause
  goto MENU
)
echo.
echo [*] Verificando arquivos alterados...
rem A planilha completa (com matriculas) fica so neste computador
git rm --cached --ignore-unmatch -q "relatorios e formularios/Visitas_Monitoramento_Sedes_Kobo.xlsx" >nul 2>&1
git add -A
git status -s
git diff --cached --quiet
if not errorlevel 1 (
  echo.
  echo [INFO] Nenhuma alteracao para publicar.
  pause
  goto MENU
)
for /f "tokens=*" %%i in ('powershell -NoProfile -Command "Get-Date -Format 'dd/MM/yyyy HH:mm'"') do set "DATAHORA=%%i"
git commit -q -m "Atualizacao da base NoFly - %DATAHORA%"
echo.
echo [*] Enviando atualizacoes para o GitHub...
git push origin main
if errorlevel 1 (
  color 0C
  echo.
  echo ================================================================================
  echo   [ERRO] O envio para o GitHub FALHOU. O app do celular NAO foi atualizado.
  echo   Verifique a internet e o login do Git e tente a opcao [4] novamente.
  echo ================================================================================
  pause
  goto MENU
)
echo.
echo ================================================================================
echo   [CONCLUIDO] Projeto publicado no GitHub em %DATAHORA%.
echo   Em alguns minutos o app do celular recebe a nova versao automaticamente.
echo ================================================================================
pause
goto MENU

:OP_DEPS
cls
echo ================================================================================
echo   [R] INSTALANDO BIBLIOTECAS PYTHON (requirements.txt)...
echo ================================================================================
python -m pip install -r requirements.txt
if not exist "scripts\.env" (
  echo.
  echo [ATENCAO] Falta o arquivo scripts\.env com o token do KoboToolbox.
  echo           Copie scripts\.env.example para scripts\.env e cole o token.
)
pause
goto MENU

:OP_QGIS
cls
echo Abrindo o projeto MONITORAMENTO QUEIROZ.qgz no QGIS...
start "" "banco de dados geograficos\MONITORAMENTO QUEIROZ.qgz"
echo Projeto aberto!
timeout /t 2 > nul
goto MENU

:OP_KOBO
cls
echo Abrindo tabela de dados no KoboToolbox na nuvem...
start "" "https://kf.kobotoolbox.org/#/forms/a6xVWc5GTZEWFzSceegC4G/data/table"
echo Abrindo navegador!
timeout /t 2 > nul
goto MENU

:OP_SAIR
cls
endlocal
exit
