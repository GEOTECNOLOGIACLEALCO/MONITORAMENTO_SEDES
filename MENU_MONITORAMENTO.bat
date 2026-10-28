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
echo   [0] Sair
echo.
echo ================================================================================
echo.
choice /C 12345670 /N /M "Pressione 1, 2, 3, 4, 5, 6, 7 ou 0: "

if errorlevel 8 goto OP_SAIR
if errorlevel 7 goto OP_ANALISE
if errorlevel 6 goto OP_KOBO
if errorlevel 5 goto OP_QGIS
if errorlevel 4 goto OP_GITHUB
if errorlevel 3 goto OP_UPDATE
if errorlevel 2 goto OP_EXCEL
if errorlevel 1 goto OP_OPEN_WEB
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
echo [*] Sincronizando vistorias do KoboToolbox e atualizando base web...
python "scripts\atualizar_base_web.py"
echo.
echo [*] Verificando arquivos alterados...
git add -A
git status -s
echo.
echo [*] Enviando atualizacoes para o GitHub...
git commit -m "Atualizacao do sistema de vistoria de sedes" > nul 2>&1
git push origin main
echo.
echo ================================================================================
echo   [CONCLUIDO] Projeto sincronizado com o GitHub com sucesso!
echo   Basta tocar em 'Limpar Cache' no celular para carregar as mudancas.
echo ================================================================================
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
