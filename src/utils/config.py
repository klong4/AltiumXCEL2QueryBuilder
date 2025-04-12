#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import json
import os
from pathlib import Path
import logging

class ConfigManager:
    """Manages application configuration using a JSON file."""

    DEFAULT_CONFIG = {
        "theme": "dark",
        "recent_files": [],
        "max_recent_files": 10,
        "window_geometry": None,
        "window_state": None,
        "last_export_dir": str(Path.home()),
        "last_import_dir": str(Path.home()),
        "auto_load_last_file": False,
        "last_opened_file": None,
        "default_rule_name": "GeneratedRule",
        "default_rule_priority": 1,
        "default_rule_enabled": True,
        "default_rule_comment": "Generated by AltiumXCEL2QueryBuilder"
    }

    def __init__(self, config_file_name="config.json"):
        """Initializes the ConfigManager."""
        self.config_dir = Path.home() / ".AltiumXCEL2QueryBuilder"
        self.config_file_path = self.config_dir / config_file_name
        self.config = {}
        self._load_config()

    def _load_config(self):
        """Loads the configuration from the JSON file."""
        try:
            self.config_dir.mkdir(parents=True, exist_ok=True)

            if self.config_file_path.exists():
                with open(self.config_file_path, 'r', encoding='utf-8') as f:
                    loaded_config = json.load(f)
                    self.config = self.DEFAULT_CONFIG.copy()
                    self.config.update(loaded_config)
            else:
                self.config = self.DEFAULT_CONFIG.copy()
                self._save_config()
        except (json.JSONDecodeError, IOError) as e:
            logging.error(f"Error loading configuration: {e}. Using default configuration.")
            self.config = self.DEFAULT_CONFIG.copy()
            self._save_config() # Attempt to save a valid default config

    def _save_config(self):
        """Saves the current configuration to the JSON file."""
        try:
            self.config_dir.mkdir(parents=True, exist_ok=True)
            with open(self.config_file_path, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=4)
        except IOError as e:
            logging.error(f"Error saving configuration: {e}")

    def get(self, key, default=None):
        """Gets a configuration value."""
        return self.config.get(key, default)

    def set(self, key, value):
        """Sets a configuration value and saves the configuration."""
        self.config[key] = value
        self._save_config()

    def add_recent_file(self, file_path):
        """Adds a file path to the list of recent files."""
        recent_files = self.get("recent_files", [])
        if file_path in recent_files:
            recent_files.remove(file_path)
        recent_files.insert(0, file_path)
        max_files = self.get("max_recent_files", 10)
        self.set("recent_files", recent_files[:max_files])

    def get_recent_files(self):
        """Gets the list of recent files."""
        return self.get("recent_files", [])

    def clear_recent_files(self):
        """Clears the list of recent files."""
        self.set("recent_files", [])

# Global instance
config_manager = ConfigManager()
