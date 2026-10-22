@echo off
chcp 65001 >nul
setlocal
title Instalar LaTeX IHR
set "PASTA=%~dp0"
if "%PASTA:~-1%"=="\" set "PASTA=%PASTA:~0,-1%"
if not "%PASTA:~0,2%"=="\\" goto unidade

echo Instalando o suplemento LaTeX IHR (catálogo: %PASTA%)
echo.
reg add "HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{5e63d59a-2fb9-41c3-8b5e-79456743365e}" /v Id /t REG_SZ /d "{5e63d59a-2fb9-41c3-8b5e-79456743365e}" /f >nul || goto erro
reg add "HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{5e63d59a-2fb9-41c3-8b5e-79456743365e}" /v Url /t REG_SZ /d "%PASTA%" /f >nul || goto erro
reg add "HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{5e63d59a-2fb9-41c3-8b5e-79456743365e}" /v Flags /t REG_DWORD /d 1 /f >nul || goto erro
call "%~dp0atualizar-modelos.bat" <nul >nul 2>&1
echo Modelos de documento instalados (Arquivo ^> Novo ^> Pessoal).
echo.
echo Pronto! Agora:
echo   1. Feche e abra o Word de novo.
echo   2. Página Inicial ^> Suplementos ^> Avançado
echo      (ou Inserir ^> Meus Suplementos)
echo   3. Aba PASTA COMPARTILHADA ^> LaTeX IHR ^> Adicionar.
echo   4. O botão "Editor LaTeX" aparece na guia Página Inicial.
echo.
pause
exit /b 0

:unidade
echo Este instalador precisa ser aberto pelo endereço de rede da pasta,
echo começando com duas barras, por exemplo: \\MAQUINA\LaTeX-IHR
echo Ele foi aberto por uma letra de unidade: %PASTA%
echo Abra a pasta pelo endereço \\... no Explorador de Arquivos e execute de novo.
echo.
pause
exit /b 1

:erro
echo Não foi possível gravar a configuração do Office no registro do Windows.
echo Avise o responsável pelo suplemento.
pause
exit /b 1
