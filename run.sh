#!/usr/bin/env bash
# ──────────────────────────────────────────────────────────────────────────────
# Multifamily Demand Index — Auto-Launcher (Linux / macOS)
#
# Double-click this file or run: bash run.sh
# It will automatically:
#   1. Create a Python virtual environment (first run only)
#   2. Install dependencies (first run only)
#   3. Launch the Streamlit app in your default browser
# ──────────────────────────────────────────────────────────────────────────────
set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

VENV_DIR=".venv"
PYTHON=""

# ── Find Python 3 ────────────────────────────────────────────────────────────
for candidate in python3 python; do
    if command -v "$candidate" &>/dev/null; then
        version=$("$candidate" --version 2>&1 | grep -oP '\d+\.\d+')
        major=$(echo "$version" | cut -d. -f1)
        minor=$(echo "$version" | cut -d. -f2)
        if [[ "$major" -ge 3 ]] && [[ "$minor" -ge 9 ]]; then
            PYTHON="$candidate"
            break
        fi
    fi
done

if [[ -z "$PYTHON" ]]; then
    echo "❌ Python 3.9+ is required but not found."
    echo "   Please install Python from https://www.python.org/downloads/"
    read -rp "Press Enter to exit..."
    exit 1
fi

echo "🐍 Using $($PYTHON --version)"

# ── Create virtual environment if needed ─────────────────────────────────────
if [[ ! -d "$VENV_DIR" ]]; then
    echo "📦 Creating virtual environment…"
    "$PYTHON" -m venv "$VENV_DIR"
fi

# ── Activate and install ─────────────────────────────────────────────────────
source "$VENV_DIR/bin/activate"

# Re-install whenever requirements.txt changes (hash stored in .installed)
REQ_HASH=$(md5sum requirements.txt 2>/dev/null || md5 -q requirements.txt 2>/dev/null)
STORED_HASH=$(cat "$VENV_DIR/.installed" 2>/dev/null || echo "")

if [[ "$REQ_HASH" != "$STORED_HASH" ]]; then
    echo "📥 Installing dependencies…"
    pip install --upgrade pip -q
    pip install -r requirements.txt -q
    echo "$REQ_HASH" > "$VENV_DIR/.installed"
    echo "✅ Dependencies installed"
fi

# ── Suppress Streamlit email prompt ─────────────────────────────────────────
mkdir -p "$HOME/.streamlit"
if [[ ! -f "$HOME/.streamlit/credentials.toml" ]]; then
    printf '[general]\nemail = ""\n' > "$HOME/.streamlit/credentials.toml"
fi

# ── Launch ───────────────────────────────────────────────────────────────────
echo ""
echo "🚀 Launching Multifamily Demand Index App…"
echo "   (Close this terminal window to stop the app)"
echo ""

streamlit run app.py --server.headless=false --browser.gatherUsageStats=false
