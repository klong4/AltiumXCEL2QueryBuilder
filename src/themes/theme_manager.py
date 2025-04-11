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
from PyQt5.QtWidgets import QApplication
from PyQt5.QtCore import QObject, QFile, QTextStream

logger = logging.getLogger(__name__)

class ThemeManager(QObject):
    """Manages application themes and styling"""
    
    # Available themes
    THEMES = {
        "light": "Light Mode",
        "dark": "Dark Mode"
    }
    
    def __init__(self, app, config_manager):
        """Initialize the theme manager"""
        super().__init__()
        self.app = app
        self.config = config_manager
        self.current_theme = self.config.get("theme", "light")
        
        # Get the path to the styles directory
        self.base_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
        self.styles_dir = os.path.join(self.base_dir, "styles")
        
        logger.info(f"Theme manager initialized with theme: {self.current_theme}")
        
    def get_available_themes(self):
        """Get list of available themes"""
        return self.THEMES
    
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
        """Apply the current theme to the application"""
        theme_file = os.path.join(self.styles_dir, f"{self.current_theme}.qss")
        
        try:
            with open(theme_file, "r") as f:
                stylesheet = f.read()
                
            # Apply stylesheet to application
            self.app.setStyleSheet(stylesheet)
            logger.info(f"Applied theme from file: {theme_file}")
            return True
        except Exception as e:
            logger.error(f"Error loading theme file {theme_file}: {str(e)}")
            # Apply empty stylesheet as fallback
            self.app.setStyleSheet("")
            return False
