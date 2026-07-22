#!/bin/bash
# Iniciador — Mac / Linux
set -e
cd "$(dirname "$0")"

# Verificar Python 3
if ! command -v python3 &>/dev/null; then
  echo "❌  Python 3 no encontrado. Instalar desde https://www.python.org"
  exit 1
fi

# Crear entorno virtual si no existe
if [ ! -d ".venv" ]; then
  echo "⚙️   Creando entorno virtual..."
  python3 -m venv .venv
fi

# Activar
source .venv/bin/activate

# Instalar dependencias
echo "📦  Instalando dependencias..."
pip install --quiet --upgrade pip
pip install --quiet -r requirements.txt

# Abrir navegador después de 4 segundos
(sleep 4 && python3 -c "import webbrowser; webbrowser.open('http://localhost:8080')") &

# Iniciar servidor
echo ""
echo "✅  Iniciando servidor en http://localhost:8080"
echo "   Ctrl+C para detener"
echo ""
python3 watcher.py
