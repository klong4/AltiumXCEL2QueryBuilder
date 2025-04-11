#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Configuration Manager
====================

Handles application configuration, preferences and settings.
Provides methods to load and save configurations.
"""

import os
import json
import logging
from pathlib import Path

logger = logging.getLogger(__name__)

class ConfigManager:
    """Manages application configuration settings"""
    
    # Default configuration values
    DEFAULT_CONFIG = {
        "theme": "light",
        "recent_files": [],
        "default_units": "mm",  # Options: mm, mil, inch
        "max_recent_files": 5,
        "export_path": "",
        "auto_save": True,
        "check_updates": True,
        "last_directory": "",
    }
    
    def __init__(self, config_file="settings.json"):
        """Initialize the configuration manager"""
        self.config_dir = os.path.join(os.path.expanduser("~"), ".altium_rule_generator")
        self.config_file = os.path.join(self.config_dir, config_file)
        self.config = {}
        
        # Ensure config directory exists
        if not os.path.exists(self.config_dir):
            os.makedirs(self.config_dir)
            
        # Load configuration
        self.load_config()
    
    def load_config(self):
        """Load configuration from file or use defaults"""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r') as f:
                    loaded_config = json.load(f)
                    # Merge with defaults to ensure all keys exist
                    self.config = {**self.DEFAULT_CONFIG, **loaded_config}
                    logger.info("Configuration loaded successfully")
            else:
                self.config = self.DEFAULT_CONFIG.copy()
                logger.info("Default configuration loaded")
                # Save default configuration
                self.save_config()
        except Exception as e:
            logger.error(f"Error loading configuration: {str(e)}")
            self.config = self.DEFAULT_CONFIG.copy()
    
    def save_config(self):
        """Save current configuration to file"""
        try:
            with open(self.config_file, 'w') as f:
                json.dump(self.config, f, indent=4)
            logger.info("Configuration saved successfully")
            return True
        except Exception as e:
            logger.error(f"Error saving configuration: {str(e)}")
            return False
    
    def get(self, key, default=None):
        """Get a configuration value"""
        return self.config.get(key, default)
    
    def set(self, key, value):
        """Set a configuration value"""
        self.config[key] = value
        return self.save_config()
    
    def add_recent_file(self, file_path):
        """Add a file to recent files list"""
        if not file_path:
            return False
            
        recent_files = self.get("recent_files", [])
        
        # Remove if already exists (to move to top)
        if file_path in recent_files:
            recent_files.remove(file_path)
            
        # Add to beginning of list
        recent_files.insert(0, file_path)
        
        # Limit list size
        max_files = self.get("max_recent_files", 5)
        if len(recent_files) > max_files:
            recent_files = recent_files[:max_files]
            
        self.set("recent_files", recent_files)
        return True
    
    def update_last_directory(self, directory):
        """Update the last used directory"""
        if os.path.isdir(directory):
            self.set("last_directory", directory)
            return True
        return False
