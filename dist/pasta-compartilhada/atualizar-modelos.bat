@echo off
chcp 65001 >nul
setlocal
title Atualizar modelos do Word (LaTeX IHR)
set "PASTA=%~dp0"
if "%PASTA:~-1%"=="\" set "PASTA=%PASTA:~0,-1%"
rem Pasta de modelos pessoais do Word (Arquivo > Novo > Pessoal)
set "DEST="
for /f "tokens=2,*" %%a in ('reg query "HKCU\Software\Microsoft\Office\16.0\Word\Options" /v PersonalTemplates 2^>nul ^| find "PersonalTemplates"') do set "DEST=%%b"
if defined DEST call set "DEST=%DEST%"
if defined DEST goto copiar
for /f "usebackq delims=" %%d in (`powershell -NoProfile -Command "[Environment]::GetFolderPath('MyDocuments')"`) do set "DOCS=%%d"
if not defined DOCS set "DOCS=%USERPROFILE%\Documents"
set "DEST=%DOCS%\Modelos Personalizados do Office"
reg add "HKCU\Software\Microsoft\Office\16.0\Word\Options" /v PersonalTemplates /t REG_EXPAND_SZ /d "%DEST%" /f >nul

:copiar
if not exist "%DEST%" mkdir "%DEST%"
echo Copiando os modelos para: %DEST%
copy /y "%PASTA%\Modelos\*.dotx" "%DEST%\" >nul || goto erro
echo.
echo Pronto! No Word: Arquivo ^> Novo ^> Pessoal mostra os modelos:
for %%f in ("%PASTA%\Modelos\*.dotx") do echo   - %%~nf
echo.
echo Documentos criados a partir deles ja abrem o painel LaTeX IHR ajustado ao modelo.
pause
exit /b 0

:erro
echo Nao foi possivel copiar os modelos. Verifique se a pasta de rede esta acessivel.
pause
exit /b 1
