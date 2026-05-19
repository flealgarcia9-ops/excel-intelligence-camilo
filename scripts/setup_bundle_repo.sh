#!/usr/bin/env bash
set -euo pipefail

# Unpack the prepared bundle and create a local git repo with a single commit.
# Then optionally push to a remote repository if provided.

# Path to the bundle you want to commit (adjust if needed)
BUNDLE_PATH="/Users/camilomartinez/Documents/tareacam/tareacam_millonario_bundle.zip"

# Directory where the repo will live (created next to the current location)
WORK_DIR="$(pwd)/tareacam_millonario_bundle_repo"

echo "[Bundle] Unpacking bundle to: $WORK_DIR"
rm -rf "$WORK_DIR" 2>/dev/null || true
mkdir -p "$WORK_DIR"
unzip -o "$BUNDLE_PATH" -d "$WORK_DIR" >/dev/null

cd "$WORK_DIR"

echo "[Git] Initializing local repository..."
git init
git add .
git commit -m "feat(bundle): agregar Quien quiere ser millonario (CLI, GUI, web, tests)"

echo "[Git] Local commit realizado."
echo -n "¿Deseas subir a un remoto? (s/N): "
read -r RESP
RESP=${RESP:-n}
if [[ "$RESP" =~ ^[SsYy]$ ]]; then
  echo -n "Ingresa la URL del remoto (ej. https://github.com/tu-usuario/tareacam_millonario_bundle.git): "
  read -r REMOTE
  git remote add origin "$REMOTE"
  git branch -M main
  git push -u origin main
  echo "Push completado. Revisa el remoto en: $REMOTE"
else
  echo "Repositorio local listo en: $WORK_DIR. Puedes subirlo manualmente cuando lo desees with git push."
fi

# Optional: si tienes gh CLI instalado y autenticado, puedes usarlo para crear el remoto automáticamente:
# gh repo create tareacam_millonario_bundle --public --source=. --remote=origin
# git push -u origin main
