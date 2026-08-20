#!/usr/bin/env bash
# Publica esta pasta como um repositorio publico no seu GitHub.
#
# Uso:
#   1) Crie um repositorio PUBLICO e VAZIO em https://github.com/new
#      (sugestao de nome: skill-agravo-de-instrumento, sem README, sem .gitignore, sem licenca)
#   2) Rode, de dentro desta pasta:
#        bash publicar-no-github.sh <seu-usuario> <nome-do-repositorio>
set -e
USUARIO="$1"
REPO="$2"
if [ -z "$USUARIO" ] || [ -z "$REPO" ]; then
  echo "Uso: bash publicar-no-github.sh <seu-usuario> <nome-do-repositorio>"
  exit 1
fi
cd "$(dirname "$0")"
git init -b main
git add -A
git commit -m "Skill de Agravo de Instrumento municipal: instrucoes, referencias e gerador .docx"
git remote add origin "https://github.com/$USUARIO/$REPO.git"
git push -u origin main
echo
echo "Pronto: https://github.com/$USUARIO/$REPO"
