#!/bin/bash
# Double-click this file on Mac to run the TikTok data pipeline.

# Move to the folder where this script lives (handles spaces in path)
cd "$(dirname "$0")"

echo ""
echo "=== TikTok Report Builder ==="
echo ""

# ---------------------------------------------------------------
# Find a Python 3 with pandas available.
# Checks (in order): conda base, common venv names, system python3.
# ---------------------------------------------------------------
find_python() {
    # 1. Conda environments
    for conda_root in "$HOME/miniconda3" "$HOME/anaconda3" "$HOME/opt/anaconda3" "$HOME/opt/miniconda3"; do
        if [ -x "$conda_root/bin/python" ]; then
            if "$conda_root/bin/python" -c "import pandas" 2>/dev/null; then
                echo "$conda_root/bin/python"
                return
            fi
        fi
    done

    # 2. Local virtual environments (venv / .venv in project folder)
    for venv in "venv" ".venv"; do
        if [ -x "$venv/bin/python" ]; then
            if "$venv/bin/python" -c "import pandas" 2>/dev/null; then
                echo "$venv/bin/python"
                return
            fi
        fi
    done

    # 3. System / Homebrew python3
    for py in "python3" "python"; do
        if command -v "$py" &>/dev/null; then
            if "$py" -c "import pandas" 2>/dev/null; then
                echo "$(command -v $py)"
                return
            fi
        fi
    done

    echo ""
}

PYTHON=$(find_python)

if [ -z "$PYTHON" ]; then
    echo "ERROR: Could not find a Python installation with pandas."
    echo ""
    echo "Please install the required packages by running this in Terminal:"
    echo "   pip3 install pandas openpyxl"
    echo ""
    read -p "Press Enter to close..."
    exit 1
fi

echo "Using Python: $PYTHON"
echo ""

# Run the pipeline
"$PYTHON" merge_pipeline.py
EXIT_CODE=$?

echo ""
if [ $EXIT_CODE -eq 0 ]; then
    echo "Report saved to output/"
    # Open the output folder in Finder
    open output/
else
    echo "Something went wrong (exit code $EXIT_CODE). Check the messages above."
fi

echo ""
read -p "Press Enter to close this window..."
