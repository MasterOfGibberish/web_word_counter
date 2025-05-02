#!/usr/bin/env python
"""
Web Word Counter Launcher Script
Provides a GUI interface to the web word counter tool
"""

import os
import sys
import tkinter as tk
import subprocess

def has_gui_dependencies():
    """Check if GUI dependencies are installed"""
    try:
        import tkinter
        from tkinter import ttk
        try:
            from ttkthemes import ThemedTk
            return True
        except ImportError:
            # Can still run with normal tkinter
            return True
    except ImportError:
        return False

def launch_gui():
    """Launch the GUI interface"""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    gui_path = os.path.join(script_dir, 'gui.py')
    
    if os.path.exists(gui_path):
        try:
            if sys.platform == 'win32':
                # On Windows, use pythonw.exe to avoid console window
                python_exec = os.path.join(os.path.dirname(sys.executable), 'pythonw.exe')
                if not os.path.exists(python_exec):
                    python_exec = sys.executable
            else:
                python_exec = sys.executable
                
            subprocess.Popen([python_exec, gui_path])
            return True
        except Exception as e:
            print(f"Error launching GUI: {e}")
            return False
    else:
        print(f"GUI file not found: {gui_path}")
        return False

def launch_cli():
    """Print help for command line usage"""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    main_path = os.path.join(script_dir, 'main.py')
    
    print("Web Word Counter - Command Line Mode")
    print("====================================")
    print(f"Usage: python {main_path} URL [options]")
    print("\nOptions:")
    print("  -l, --limit NUM       Maximum number of pages to crawl (default: 10)")
    print("  -o, --output FILE     Output filename (default: website_text.xlsx)")
    print("  -w, --wait SECONDS    Wait time in seconds for page loading (default: 10)")
    print("  --format FORMAT       Output format: 'excel' or 'docx' (default: excel)")
    print("  --overwrite           Overwrite existing output file if it exists")
    print("\nExample:")
    print(f"  python {main_path} https://example.com -l 20 -o report.xlsx")
    
    print("\nTo use the GUI interface, run without arguments:")
    print(f"  python {os.path.basename(__file__)}")
    
def main():
    """Main entry point"""
    if len(sys.argv) > 1:
        # If arguments provided, pass to main.py
        script_dir = os.path.dirname(os.path.abspath(__file__))
        main_path = os.path.join(script_dir, 'main.py')
        
        if os.path.exists(main_path):
            cmd = [sys.executable, main_path] + sys.argv[1:]
            try:
                subprocess.run(cmd)
            except KeyboardInterrupt:
                print("\nProcess interrupted by user.")
        else:
            print(f"Error: Main script not found at {main_path}")
            return
    else:
        # No arguments, try to launch GUI
        if has_gui_dependencies():
            if not launch_gui():
                # If GUI fails, show CLI help
                launch_cli()
        else:
            print("GUI dependencies not found. Please install tkinter and ttkthemes.")
            print("You can run in command-line mode or install the required packages:")
            print("  pip install ttkthemes")
            launch_cli()

if __name__ == "__main__":
    main() 