#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import logging
from pathlib import Path
from PyQt5.QtWidgets import QApplication
from PyQt5.QtCore import QFile, QTextStream

from utils.config import config_manager
from .dark import DarkTheme # Import theme class
from .light import LightTheme # Import theme class

BASE_DIR = Path(__file__).resolve().parent.parent
STYLES_DIR = os.path.join(BASE_DIR, '..', 'styles')

class ThemeManager:
    """Manages application themes."""

    def __init__(self, app: QApplication):
        """
        Initializes the ThemeManager.

        Args:
            app (QApplication): The main application instance.
        """
        self.app = app
        # Store theme instances
        self.THEMES = {
            "light": LightTheme(),
            "dark": DarkTheme(),
        }
        self.current_theme_name = config_manager.get("theme", "dark") # Default to dark theme
        self.apply_theme(self.current_theme_name)

    def apply_theme(self, theme_name: str):
        """
        Applies the specified theme to the application.

        Args:
            theme_name (str): The name of the theme to apply (e.g., "light", "dark").
        """
        if theme_name not in self.THEMES:
            logging.warning(f"Theme '{theme_name}' not found. Using default theme.")
            theme_name = "light" # Fallback to default

        theme_instance = self.THEMES.get(theme_name)

        if theme_instance:
            theme_instance.apply(self.app) # Call the theme's apply method
            self.current_theme_name = theme_name
            config_manager.set("theme", theme_name) # Save the selected theme
            logging.info(f"Applied theme: {theme_name}")
        else:
            # This case should ideally not happen if fallback works
            logging.error(f"Failed to get theme instance for: {theme_name}")

    def get_current_theme(self) -> str:
        """
        Returns the name of the currently applied theme.

        Returns:
            str: The name of the current theme.
        """
        return self.current_theme_name

    def get_available_themes(self) -> list:
        """
        Returns a list of available theme names.

        Returns:
            list: A list of strings representing the available theme names.
        """
        return list(self.THEMES.keys())
