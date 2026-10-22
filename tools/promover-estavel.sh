#!/bin/sh
# Copia a versão atual (canal de teste, raiz do site) para o canal estável (stable/).
# Uso: sh tools/promover-estavel.sh   (depois: commit + push)
set -e
cd "$(dirname "$0")/.."
rm -rf stable && mkdir stable
cp -r taskpane.html boot.js panel.html taskpane.css taskpane.js commands.html assets vendor lib stable/
grep -o "V = '[0-9.]*'" stable/boot.js
