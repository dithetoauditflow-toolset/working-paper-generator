#main.py
from cli_ui import main

if __name__ == "__main__":
    """
    Entry point for the Working Paper Generator CLI application.

    This script launches the new CLI-based interface that provides
    a beautiful command-line experience with Tkinter dialogs for
    file selection, consultant name input, and directory selection.

    The CLI interface features:
    - Professional colored output with boxes and borders
    - Progress bars and status updates
    - File status display with metadata
    - Beautiful results tables
    - Comprehensive error handling

    Returns:
        None
    """
    main()
