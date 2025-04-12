#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Altium Rule Generator - Main Application Entry
==============================================

This is the main entry point for the Altium Rule Generator application.
It initializes the application, sets up logging, and launches the GUI.
"""

import sys
import logging
import os
from PyQt5.QtWidgets import QApplication
from PyQt5.QtCore import Qt

# Add the src directory to the Python path so imports work correctly
current_dir = os.path.dirname(os.path.abspath(__file__))
parent_dir = os.path.dirname(current_dir)
sys.path.insert(0, parent_dir)

# Import directly from the modules without src prefix
from gui.main_window import MainWindow
from utils.config import ConfigManager
from themes.theme_manager import ThemeManager

# Setup logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("altium_rule_generator.log"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

def main():
    """Main application entry point"""
    logger.info("Starting Altium Rule Generator application")
    
    # Enable High DPI scaling
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)
    
    # Create application instance
    app = QApplication(sys.argv)
    app.setApplicationName("Altium Rule Generator")
    app.setOrganizationName("AltiumTools")
    
    # Load configuration
    config = ConfigManager()
    
    # Initialize theme manager
    theme_manager = ThemeManager(app)  # Remove the config argument
    theme_manager.apply_theme(config.get("theme", "light")) # Pass theme name from config
    
    # Create and show the main window
    main_window = MainWindow(config, theme_manager)
    main_window.show()
    
    # Start the application event loop
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
