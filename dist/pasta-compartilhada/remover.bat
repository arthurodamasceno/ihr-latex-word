@echo off
chcp 65001 >nul
title Remover LaTeX IHR
reg delete "HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{5e63d59a-2fb9-41c3-8b5e-79456743365e}" /f >nul 2>&1
echo Catálogo do LaTeX IHR removido deste usuário.
echo No Word, o suplemento some depois de fechar e abrir o programa.
echo (Se ele ainda aparecer: Página Inicial ^> Suplementos ^> Mais Suplementos ^> Meus Suplementos ^> ... ^> Remover.)
pause
