#!/usr/bin/env bash
# ---------------------------------------------------------------------------
# pptx_translator installer for macOS / Linux
#
# Just double-nothing: open a Terminal in this folder and run:
#     bash install.sh
#
# It creates a private, self-contained Python environment in a ".venv" folder
# and installs everything the tool needs. It does not touch the rest of your
# computer.
# ---------------------------------------------------------------------------
set -e

cd "$(dirname "$0")"

echo ""
echo "  pptx_translator - setup (macOS / Linux)"
echo "  ---------------------------------------"

# 1. Find a Python 3 interpreter.
if command -v python3 >/dev/null 2>&1; then
  PY=python3
elif command -v python >/dev/null 2>&1; then
  PY=python
else
  echo ""
  echo "  Python 3 was not found on your computer."
  echo "  Please install it from https://www.python.org/downloads/ and run this"
  echo "  script again."
  exit 1
fi

echo "  Using: $($PY --version)"

# 2. Create the virtual environment.
echo "  Creating a private Python environment in ./.venv ..."
$PY -m venv .venv

# 3. Install dependencies into it.
echo "  Installing required packages (this can take a minute) ..."
./.venv/bin/python -m pip install --upgrade pip >/dev/null
./.venv/bin/python -m pip install -r requirements.txt

# 4. Quick self-test.
echo "  Verifying the installation ..."
./.venv/bin/python -m pptx_translator --help >/dev/null

echo ""
echo "  All set!  To use the tool, run these commands from this folder:"
echo ""
echo "    Extract text:"
echo "      ./.venv/bin/python -m pptx_translator extract \"MyDeck.pptx\" -o \"MyDeck.json\""
echo ""
echo "    Patch translations back in:"
echo "      ./.venv/bin/python -m pptx_translator patch \"MyDeck.pptx\" \"MyDeck.json\" -o \"MyDeck_translated.pptx\""
echo ""
