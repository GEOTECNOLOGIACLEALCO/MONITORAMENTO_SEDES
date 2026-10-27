@echo off
setlocal
title Clealco - Monitoramento de Sedes
cd /d "%~dp0"

:MENU
cls
color 0A
echo ================================================================================
echo             CLEALCO - MONITORAMENTO DE SEDES (ROTINA DE VISTORIAS)
echo ================================================================================
echo.
echo   [1] Abrir Mapa de Vistorias no Navegador (index.html)
echo   [2] Gerar Relatorio de Vistorias em Excel (Visitas_Monitoramento_Sede.xlsx)
echo   [3] Atualizar Base do Mapa (Processar shapefiles de Sedes e Areas)
echo   [4] Publicar Alteracoes no GitHub (Atualizar App do Celular)
echo   [5] Abrir Projeto no QGIS (MONITORAMENTO QUEIROZ.qgz)
echo   [6] Abrir Pasta de Fotos das Vistorias (Fotos_Visitas)
echo   [0] Sair
echo.
echo ================================================================================
echo.
choice /C 1234560 /N /M "Pressione 1, 2, 3, 4, 5, 6 ou 0: "

if errorlevel 7 goto OP_SAIR
if errorlevel 6 goto OP_FOTOS
if errorlevel 5 goto OP_QGIS
if errorlevel 4 goto OP_GITHUB
if errorlevel 3 goto OP_UPDATE
if errorlevel 2 goto OP_EXCEL
if errorlevel 1 goto OP_OPEN_WEB
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
echo   [2] GERANDO RELATORIO DE VISTORIAS EM EXCEL...
echo ================================================================================
echo.
python exportar_planilha_visitas.py
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
python atualizar_base_web.py
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
echo   Basta tocar em 'Limpar Cache e Atualizar' no celular para carregar as mudancas.
echo ================================================================================
pause
goto MENU

:OP_QGIS
cls
echo Abrindo o projeto MONITORAMENTO QUEIROZ.qgz no QGIS...
start "" "MONITORAMENTO QUEIROZ.qgz"
echo Projeto aberto!
timeout /t 2 > nul
goto MENU

:OP_FOTOS
cls
color 0E
echo ================================================================================
echo   [6] ABRINDO PASTA DE FOTOS DAS VISTORIAS...
echo ================================================================================
echo.
if not exist "Fotos_Visitas" mkdir "Fotos_Visitas"
start "" "Fotos_Visitas"
echo Pasta Fotos_Visitas aberta no Explorer!
timeout /t 2 > nul
goto MENU

:OP_SAIR
cls
endlocal
exit
