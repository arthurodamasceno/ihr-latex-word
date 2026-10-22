@echo off
chcp 65001 >nul
title Limpar cache do LaTeX IHR
echo Use se o painel ficar em branco, travado ou mostrando uma versão antiga.
echo Feche o Word antes de continuar e esteja na rede da empresa (ou VPN).
pause
rmdir /s /q "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef" 2>nul
echo Cache dos suplementos do Office limpo. Pode abrir o Word.
echo Se o botão sumir, adicione de novo pela aba PASTA COMPARTILHADA.
pause
