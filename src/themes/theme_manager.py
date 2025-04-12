#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Theme Manager
============

Manages application themes and styles.
Handles loading, applying, and switching between themes.
"""

import os
import logging
import sys
from PyQt5.QtWidgets import QApplication
from PyQt5.QtCore import QObject, QFile, QTextStream, QDir

logger = logging.getLogger(__name__)

# Determine the base path for resources (assuming styles are in ../styles relative to this file)
# This might need adjustment depending on how the application is run/packaged
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
STYLES_DIR = os.path.join(BASE_DIR, '..', 'styles')

class ThemeManager(QObject):
    """Manages application themes and styling"""

    # Available themes and their corresponding QSS file names
    THEMES = {
        "light": "light.qss",
        "dark": "dark.qss"
    }

    def __init__(self, app, config_manager):
        """Initialize the theme manager"""
        super().__init__()
        self.app = app
        self.config = config_manager
        self.current_theme = self.config.get("theme", "light")

        logger.info(f"Theme manager initialized with theme: {self.current_theme}")

    def get_available_themes(self):
        """Get list of available theme names"""
        return list(self.THEMES.keys())

    def get_current_theme(self):
        """Get the current theme name"""
        return self.current_theme

    def set_theme(self, theme_name):
        """Set and apply a new theme"""
        if theme_name in self.THEMES:
            self.current_theme = theme_name
            self.config.set("theme", theme_name)
            self.apply_theme()
            logger.info(f"Theme changed to: {theme_name}")
            return True
        else:
            logger.error(f"Unknown theme: {theme_name}")
            return False

    def apply_theme(self):
        """Apply the current theme to the application by loading from a QSS file."""
        theme_file_name = self.THEMES.get(self.current_theme)
        if not theme_file_name:
            logger.error(f"Theme '{self.current_theme}' definition not found.")
            self.app.setStyleSheet("") # Apply empty stylesheet as fallback
            return False

        # Construct the full path to the QSS file
        qss_file_path = os.path.join(STYLES_DIR, theme_file_name)
        qss_file_path = QDir.toNativeSeparators(qss_file_path) # Ensure correct path separators

        logger.info(f"Attempting to load theme file: {qss_file_path}")

        try:
            file = QFile(qss_file_path)
            if not file.exists():
                logger.error(f"Theme file not found: {qss_file_path}")
                # Try alternative path if running from build directory
                alt_styles_dir = os.path.join(os.path.dirname(sys.executable), 'styles') if getattr(sys, 'frozen', False) else STYLES_DIR
                alt_qss_file_path = QDir.toNativeSeparators(os.path.join(alt_styles_dir, theme_file_name))
                logger.info(f"Attempting alternative path: {alt_qss_file_path}")
                file = QFile(alt_qss_file_path)
                if not file.exists():
                     logger.error(f"Alternative theme file also not found: {alt_qss_file_path}")
                     self.app.setStyleSheet("") # Apply empty stylesheet as fallback
                     return False

            if file.open(QFile.ReadOnly | QFile.Text):
                stream = QTextStream(file)
                stylesheet = stream.readAll()
                file.close()
                self.app.setStyleSheet(stylesheet)
                logger.info(f"Applied theme '{self.current_theme}' from {file.fileName()}")
                return True
            else:
                logger.error(f"Could not open theme file {file.fileName()}: {file.errorString()}")
                self.app.setStyleSheet("") # Apply empty stylesheet as fallback
                return False
        except Exception as e:
            logger.exception(f"Error applying theme '{self.current_theme}' from file {qss_file_path}: {e}", exc_info=True)
            self.app.setStyleSheet("") # Apply empty stylesheet as fallback
            return False
