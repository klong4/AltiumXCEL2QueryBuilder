#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Preferences Dialog
===============

Dialog for configuring application preferences.
"""
import logging # Add logging import
import os # Add os import
from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QFormLayout, QGroupBox,
    QLabel, QLineEdit, QComboBox, QCheckBox, QPushButton,
    QDialogButtonBox, QFileDialog, QTabWidget, QWidget, QMessageBox # Add QMessageBox
)
from PyQt5.QtCore import Qt, QSettings, pyqtSignal

# Assuming config_manager provides get/set methods similar to QSettings
# Assuming theme_manager provides get_available_themes, get_current_theme, apply_theme
# from utils.config import ConfigManager # Example import
# from themes.theme_manager import ThemeManager # Example import

# Remove UnitType import if not used directly here
# from models.rule_model import UnitType

logger = logging.getLogger(__name__)

class PreferencesDialog(QDialog):
    """Dialog for configuring application preferences"""
    
    settings_changed = pyqtSignal() # Signal emitted when settings are applied
    
    def __init__(self, config_manager, theme_manager, parent=None):
        """Initialize the preferences dialog."""
        super().__init__(parent)
        self.config_manager = config_manager
        self.theme_manager = theme_manager
        self._changed_settings = {} # Store pending changes

        self.setWindowTitle("Preferences")
        self.setMinimumWidth(500)
        
        self._init_ui()
        self._load_settings()
        
        logger.debug("PreferencesDialog initialized.")

    def _init_ui(self):
        """Initialize the user interface components."""
        main_layout = QVBoxLayout(self)
        
        self.tab_widget = QTabWidget()
        main_layout.addWidget(self.tab_widget)
        
        # Create tabs
        self.general_tab = QWidget()
        self.appearance_tab = QWidget()
        self.paths_tab = QWidget()
        
        self.tab_widget.addTab(self.general_tab, "General")
        self.tab_widget.addTab(self.appearance_tab, "Appearance")
        self.tab_widget.addTab(self.paths_tab, "Paths")
        
        # Initialize content for each tab
        self._init_general_tab()
        self._init_appearance_tab()
        self._init_paths_tab()
        
        # Dialog buttons
        self.button_box = QDialogButtonBox(
            QDialogButtonBox.Ok | QDialogButtonBox.Cancel | QDialogButtonBox.Apply
        )
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)
        self.button_box.button(QDialogButtonBox.Apply).clicked.connect(self._apply_settings)
        # Disable Apply button initially
        self.button_box.button(QDialogButtonBox.Apply).setEnabled(False)
        
        main_layout.addWidget(self.button_box)

    def _init_general_tab(self):
        """Initialize the General settings tab."""
        layout = QFormLayout(self.general_tab)
        
        # Example setting: Auto-check for updates (conceptual)
        self.check_updates_checkbox = QCheckBox("Automatically check for updates on startup")
        self.check_updates_checkbox.stateChanged.connect(lambda: self._mark_as_changed("check_for_updates", self.check_updates_checkbox.isChecked()))
        layout.addRow(self.check_updates_checkbox)

        # Example setting: Default Unit Type (if applicable globally)
        # self.default_unit_combo = QComboBox()
        # for unit in UnitType:
        #     self.default_unit_combo.addItem(unit.name.capitalize(), unit)
        # self.default_unit_combo.currentIndexChanged.connect(lambda: self._mark_as_changed("default_unit", self.default_unit_combo.currentData().name))
        # layout.addRow("Default Unit:", self.default_unit_combo)

        # Add more general settings as needed
        layout.addRow(QLabel("More general settings can be added here."))


    def _init_appearance_tab(self):
        """Initialize the Appearance settings tab."""
        layout = QFormLayout(self.appearance_tab)
        
        # Theme selection
        self.theme_combo = QComboBox()
        available_themes = self.theme_manager.get_available_themes()
        for theme_name in available_themes:
            self.theme_combo.addItem(theme_name.capitalize(), theme_name)
        self.theme_combo.currentIndexChanged.connect(self._on_theme_changed)
        layout.addRow("Theme:", self.theme_combo)

        # Font selection (conceptual)
        # self.font_button = QPushButton("Select Font...")
        # self.font_label = QLabel("Default Font: System Default") # Placeholder
        # layout.addRow(self.font_label, self.font_button)

        layout.addRow(QLabel("More appearance settings can be added here."))

    def _init_paths_tab(self):
        """Initialize the Paths settings tab."""
        layout = QFormLayout(self.paths_tab)

        # Default Import Directory
        import_dir_layout = QHBoxLayout()
        self.import_dir_edit = QLineEdit()
        self.import_dir_edit.setReadOnly(True)
        import_dir_button = QPushButton("Browse...")
        import_dir_button.clicked.connect(self._browse_import_dir)
        import_dir_layout.addWidget(self.import_dir_edit)
        import_dir_layout.addWidget(import_dir_button)
        layout.addRow("Default Import Directory:", import_dir_layout)

        # Default Export Directory
        export_dir_layout = QHBoxLayout()
        self.export_dir_edit = QLineEdit()
        self.export_dir_edit.setReadOnly(True)
        export_dir_button = QPushButton("Browse...")
        export_dir_button.clicked.connect(self._browse_export_dir)
        export_dir_layout.addWidget(self.export_dir_edit)
        export_dir_layout.addWidget(export_dir_button)
        layout.addRow("Default Export Directory:", export_dir_layout)

        # Add more path settings as needed
        layout.addRow(QLabel("More path settings can be added here."))

    def _load_settings(self):
        """Load current settings from the config manager into the UI."""
        logger.debug("Loading settings into Preferences dialog.")
        # General Tab
        check_updates = self.config_manager.get("check_for_updates", True) # Default to True
        self.check_updates_checkbox.setChecked(check_updates)
        # default_unit_str = self.config_manager.get("default_unit", UnitType.MIL.name)
        # try:
        #     default_unit = UnitType[default_unit_str]
        #     index = self.default_unit_combo.findData(default_unit)
        #     if index != -1: self.default_unit_combo.setCurrentIndex(index)
        # except KeyError:
        #     logger.warning(f"Invalid default unit '{default_unit_str}' found in config.")

        # Appearance Tab
        current_theme = self.theme_manager.get_current_theme()
        index = self.theme_combo.findData(current_theme)
        if index != -1:
            self.theme_combo.setCurrentIndex(index)
        else:
            logger.warning(f"Current theme '{current_theme}' not found in available themes.")
            # Optionally set to a default theme index (e.g., 0)
            if self.theme_combo.count() > 0:
                self.theme_combo.setCurrentIndex(0)


        # Paths Tab
        import_dir = self.config_manager.get("default_import_directory", "")
        self.import_dir_edit.setText(import_dir)
        export_dir = self.config_manager.get("default_export_directory", "")
        self.export_dir_edit.setText(export_dir)

        self._changed_settings.clear() # Clear pending changes after loading
        self.button_box.button(QDialogButtonBox.Apply).setEnabled(False) # Disable Apply button

    def _save_settings(self):
        """Save the changed settings to the config manager."""
        logger.info(f"Saving {len(self._changed_settings)} changed settings.")
        if not self._changed_settings:
            return False # No changes to save

        for key, value in self._changed_settings.items():
            self.config_manager.set(key, value)
            logger.debug(f"Set config: {key} = {value}")

        # Special handling for theme application
        if "theme" in self._changed_settings:
            try:
                self.theme_manager.apply_theme(self._changed_settings["theme"])
                logger.info(f"Applied theme: {self._changed_settings['theme']}")
            except Exception as e:
                 logger.error(f"Error applying theme '{self._changed_settings['theme']}' during save: {e}")
                 # Optionally inform the user
                 QMessageBox.warning(self, "Theme Error", f"Could not apply theme '{self._changed_settings['theme']}'.")


        self._changed_settings.clear() # Clear pending changes after saving
        self.button_box.button(QDialogButtonBox.Apply).setEnabled(False) # Disable Apply button
        self.settings_changed.emit() # Signal that settings were potentially changed
        return True

    def _mark_as_changed(self, key: str, value: any):
        """Marks a setting as changed and enables the Apply button."""
        # Check if the value actually changed from the stored config value
        current_value = self.config_manager.get(key)
        # Handle potential type differences if necessary (e.g., bool vs str)
        if value != current_value:
            logger.debug(f"Setting marked as changed: {key} = {value} (was {current_value})")
            self._changed_settings[key] = value
            self.button_box.button(QDialogButtonBox.Apply).setEnabled(True)
        else:
            # If user changes back to original value, remove from changed list
            if key in self._changed_settings:
                logger.debug(f"Setting change reverted: {key}")
                del self._changed_settings[key]
                # Disable Apply button only if no other changes are pending
                if not self._changed_settings:
                    self.button_box.button(QDialogButtonBox.Apply).setEnabled(False)


    def _on_theme_changed(self):
        """Handle theme combo box change."""
        selected_theme = self.theme_combo.currentData()
        self._mark_as_changed("theme", selected_theme)

    def _apply_settings(self):
        """Apply the currently changed settings without closing the dialog."""
        logger.info("Applying preference changes.")
        self._save_settings()
        # Keep the dialog open

    def _browse_directory(self, line_edit: QLineEdit, config_key: str):
        """Opens a directory selection dialog and updates the line edit and config."""
        current_dir = line_edit.text() or self.config_manager.get("last_directory", "") # Start from current or last used
        directory = QFileDialog.getExistingDirectory(
            self,
            "Select Directory",
            current_dir
        )
        if directory:
            line_edit.setText(directory)
            self._mark_as_changed(config_key, directory)
            # Optionally update last_directory as well
            self.config_manager.set("last_directory", directory)

    def _browse_import_dir(self):
        """Browse for the default import directory."""
        self._browse_directory(self.import_dir_edit, "default_import_directory")

    def _browse_export_dir(self):
        """Browse for the default export directory."""
        self._browse_directory(self.export_dir_edit, "default_export_directory")

    def accept(self):
        """Apply settings and close the dialog."""
        logger.debug("Preferences accepted (OK clicked).")
        self._save_settings()
        super().accept()

    def reject(self):
        """Discard changes and close the dialog."""
        logger.debug("Preferences rejected (Cancel clicked).")
        if self._changed_settings:
             reply = QMessageBox.question(self, "Discard Changes?",
                                          "You have unsaved changes. Are you sure you want to discard them?",
                                          QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
             if reply == QMessageBox.No:
                 return # Do not close if user cancels discard

        super().reject()

