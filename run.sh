#!/bin/bash
# This script prepares the environment and runs the Pi-Temps-Monitor application.
# It automatically creates a virtual environment and installs dependencies if needed.

# --- Configuration ---
VENV_DIR="env"
PYTHON_SCRIPT="main_app.py"
REQUIREMENTS_FILE="requirements.txt"

# --- Script Body ---
# Get the directory where this script is located to ensure paths are correct.
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" &> /dev/null && pwd )"

# Define the full paths for key files and directories.
VENV_PATH="$SCRIPT_DIR/$VENV_DIR"
ACTIVATE_SCRIPT="$VENV_PATH/bin/activate"
PYTHON_SCRIPT_PATH="$SCRIPT_DIR/$PYTHON_SCRIPT"
REQUIREMENTS_PATH="$SCRIPT_DIR/$REQUIREMENTS_FILE"

# 1. Check for and create the virtual environment if it doesn't exist.
if [ ! -d "$VENV_PATH" ]; then
    echo "Virtual environment not found. Creating it now at '$VENV_PATH'..."
    python3 -m venv "$VENV_PATH"
    if [ $? -ne 0 ]; then
        echo "Error: Failed to create the virtual environment."
        exit 1
    fi
fi

# 2. Activate the virtual environment.
source "$ACTIVATE_SCRIPT"
echo "Virtual environment activated."

# 3. Check for a flag file to see if dependencies are already installed.
#    This avoids running 'pip install' every single time.
FLAG_FILE="$VENV_PATH/.requirements_installed"
if [ ! -f "$FLAG_FILE" ]; then
    echo "Dependencies not installed. Installing from '$REQUIREMENTS_FILE'..."
    pip install -r "$REQUIREMENTS_PATH"
    
    # Check if pip install was successful.
    if [ $? -eq 0 ]; then
        # Create the flag file to mark installation as complete.
        touch "$FLAG_FILE"
        echo "Dependencies installed successfully."
    else
        echo "Error: Failed to install dependencies."
        deactivate
        exit 1
    fi
fi

# 4. Run the Python application.
echo "---------------------------------"
echo "Starting Pi-Temps-Monitor application..."
python "$PYTHON_SCRIPT_PATH"
echo "---------------------------------"
echo "Application finished."

# 5. Deactivate the virtual environment.
deactivate
echo "Virtual environment deactivated."
