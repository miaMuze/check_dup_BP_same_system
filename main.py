"""
BP Duplicate Checker - Main Entry Point
=======================================
This is the main entry point for the BP Duplicate Checker application.
Run this file directly or use the compiled .exe version.

Usage:
    python main.py
"""

import sys
import os

# Ensure the src directory is in the path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from src.gui_app import main


if __name__ == "__main__":
    main()
