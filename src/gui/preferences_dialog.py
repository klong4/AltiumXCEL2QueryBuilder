#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Preferences Dialog
===============

Dialog for configuring application preferences.
"""

import os
import logging
from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QFormLayout, QGroupBox,
    QLabel, QLineEdit, QComboBox, QCheckBox, QPushButton,
    QDialogButtonBox, QFileDialog, QTabWidget
)
from PyQt5.QtCore import Qt, QSettings, pyqtSignal

from models.rule_model import UnitType

logger = logging.getLogger(__name__)

class PreferencesDialog(QDialog):
    """Dialog for configuring application preferences"""
    
    settings_changed = pyqtSignal()
    
    def __init__(self, config_manager, theme_manager, parent=None):
        """Initialize preferences dialog"""
        super().__init__(parent)
        self.config = config_manager
        self.theme_manager = theme_manager
        self.original_values = {}
        self.has_changes = False
        
        # Set window properties
        self.setWindowTitle("Preferences")
        self.setMinimumSize(500, 400)
        self.setModal(True)
        
        # Initialize UI
        self._init_ui()
        
        # Load current settings
        self._load_settings()
        
        logger.info("Preferences dialog initialized")
    
    def _init_ui(self):
        """Initialize the UI components"""
        # Main layout
        main_layout = QVBoxLayout()
        self.setLayout(main_layout)
        
        # Create tabs for different settings categories
        self.tabs = QTabWidget()
        main_layout.addWidget(self.tabs)
        
        # General tab
        self.general_tab = QWidget()
        self.tabs.addTab(self.general_tab, "General")
        self._init_general_tab()
        
        # Appearance tab
        self.appearance_tab = QWidget()
        self.tabs.addTab(self.appearance_tab, "Appearance")
        self._init_appearance_tab()
        
        # Paths tab
        self.paths_tab = QWidget()
        self.tabs.addTab(self.paths_tab, "Paths")
        self._init_paths_tab()
        
        # Button box
        self.button_box = QDialogButtonBox(
            QDialogButtonBox.Ok | QDialogButtonBox.Cancel | QDialogButtonBox.Apply
        )
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)
        
        # Get the Apply button
        self.apply_button = self.button_box.button(QDialogButtonBox.Apply)
        self.apply_button.clicked.connect(self._apply_settings)
        self.apply_button.setEnabled(False)  # Disabled until changes are made
        
        main_layout.addWidget(self.button_box)
    
    def _init_general_tab(self):
        """Initialize the general settings tab"""
        layout = QVBoxLayout()
        self.general_tab.setLayout(layout)
        
        # Units group
        units_group = QGroupBox("Units")
        units_layout = QFormLayout()
        units_group.setLayout(units_layout)
        
        self.default_unit_combo = QComboBox()
        self.default_unit_combo.addItem("mil", UnitType.MIL.value)
        self.default_unit_combo.addItem("mm", UnitType.MM.value)
        self.default_unit_combo.addItem("inch", UnitType.INCH.value)
        self.default_unit_combo.currentIndexChanged.connect(self._mark_as_changed)
        units_layout.addRow("Default Unit:", self.default_unit_combo)
        
        layout.addWidget(units_group)
        
        # Autosave group
        autosave_group = QGroupBox("Autosave")
        autosave_layout = QFormLayout()
        autosave_group.setLayout(autosave_layout)
        
        self.autosave_checkbox = QCheckBox("Enable autosave")
        self.autosave_checkbox.stateChanged.connect(self._mark_as_changed)
        autosave_layout.addRow("", self.autosave_checkbox)
        
        layout.addWidget(autosave_group)
        
        # Add stretch to push widgets to the top
        layout.addStretch(1)
    
    def _init_appearance_tab(self):
        """Initialize the appearance settings tab"""
        layout = QVBoxLayout()
        self.appearance_tab.setLayout(layout)
        
        # Theme group
        theme_group = QGroupBox("Theme")
        theme_layout = QFormLayout()
        theme_group.setLayout(theme_layout)
        
        self.theme_combo = QComboBox()
        # Add themes from theme manager
        for theme_id, theme_name in self.theme_manager.get_available_themes().items():
            self.theme_combo.addItem(theme_name, theme_id)
        self.theme_combo.currentIndexChanged.connect(self._mark_as_changed)
        theme_layout.addRow("Theme:", self.theme_combo)
        
        layout.addWidget(theme_group)
        
        # UI Elements group
        ui_group = QGroupBox("UI Elements")
        ui_layout = QFormLayout()
        ui_group.setLayout(ui_layout)
        
        self.statusbar_checkbox = QCheckBox("Show status bar")
        self.statusbar_checkbox.stateChanged.connect(self._mark_as_changed)
        ui_layout.addRow("", self.statusbar_checkbox)
        
        self.toolbar_checkbox = QCheckBox("Show toolbar")
        self.toolbar_checkbox.stateChanged.connect(self._mark_as_changed)
        ui_layout.addRow("", self.toolbar_checkbox)
        
        layout.addWidget(ui_group)
        
        # Add stretch to push widgets to the top
        layout.addStretch(1)
    
    def _init_paths_tab(self):
        """Initialize the paths settings tab"""
        layout = QVBoxLayout()
        self.paths_tab.setLayout(layout)
        
        # Directories group
        directories_group = QGroupBox("Default Directories")
        directories_layout = QFormLayout()
        directories_group.setLayout(directories_layout)
        
        # Import directory
        import_dir_layout = QHBoxLayout()
        self.import_dir_edit = QLineEdit()
        self.import_dir_edit.textChanged.connect(self._mark_as_changed)
        import_dir_layout.addWidget(self.import_dir_edit)
        
        self.import_dir_button = QPushButton("Browse...")
        self.import_dir_button.clicked.connect(self._browse_import_dir)
        import_dir_layout.addWidget(self.import_dir_button)
        directories_layout.addRow("Import Directory:", import_dir_layout)
        
        # Export directory
        export_dir_layout = QHBoxLayout()
        self.export_dir_edit = QLineEdit()
        self.export_dir_edit.textChanged.connect(self._mark_as_changed)
        export_dir_layout.addWidget(self.export_dir_edit)
        
        self.export_dir_button = QPushButton("Browse...")
        self.export_dir_button.clicked.connect(self._browse_export_dir)
        export_dir_layout.addWidget(self.export_dir_button)
        directories_layout.addRow("Export Directory:", export_dir_layout)
        
        layout.addWidget(directories_group)
        
        # Add stretch to push widgets to the top
        layout.addStretch(1)
    
    def _load_settings(self):
        """Load current settings into the dialog"""
        # General tab settings
        default_unit = self.config.get("default_unit", UnitType.MIL.value)
        index = self.default_unit_combo.findData(default_unit)
        if index >= 0:
            self.default_unit_combo.setCurrentIndex(index)
        self.original_values["default_unit"] = default_unit
        
        autosave = self.config.get("autosave", False)
        self.autosave_checkbox.setChecked(autosave)
        self.original_values["autosave"] = autosave
        
        # Appearance tab settings
        theme = self.theme_manager.get_current_theme()
        index = self.theme_combo.findData(theme)
        if index >= 0:
            self.theme_combo.setCurrentIndex(index)
        self.original_values["theme"] = theme
        
        show_statusbar = self.config.get("show_statusbar", True)
        self.statusbar_checkbox.setChecked(show_statusbar)
        self.original_values["show_statusbar"] = show_statusbar
        
        show_toolbar = self.config.get("show_toolbar", True)
        self.toolbar_checkbox.setChecked(show_toolbar)
        self.original_values["show_toolbar"] = show_toolbar
        
        # Paths tab settings
        import_dir = self.config.get("import_directory", "")
        self.import_dir_edit.setText(import_dir)
        self.original_values["import_directory"] = import_dir
        
        export_dir = self.config.get("export_directory", "")
        self.export_dir_edit.setText(export_dir)
        self.original_values["export_directory"] = export_dir
        
        # Reset has_changes flag
        self.has_changes = False
        self.apply_button.setEnabled(False)
    
    def _save_settings(self):
        """Save settings from the dialog"""
        # General tab settings
        default_unit = self.default_unit_combo.currentData()
        self.config.set("default_unit", default_unit)
        
        autosave = self.autosave_checkbox.isChecked()
        self.config.set("autosave", autosave)
        
        # Appearance tab settings
        theme = self.theme_combo.currentData()
        if theme != self.original_values["theme"]:
            self.theme_manager.set_theme(theme)
        
        show_statusbar = self.statusbar_checkbox.isChecked()
        self.config.set("show_statusbar", show_statusbar)
        
        show_toolbar = self.toolbar_checkbox.isChecked()
        self.config.set("show_toolbar", show_toolbar)
        
        # Paths tab settings
        import_dir = self.import_dir_edit.text()
        if import_dir and os.path.exists(import_dir):
            self.config.set("import_directory", import_dir)
        
        export_dir = self.export_dir_edit.text()
        if export_dir and os.path.exists(export_dir):
            self.config.set("export_directory", export_dir)
        
        # Reset the apply button
        self.has_changes = False
        self.apply_button.setEnabled(False)
        
        # Update original values
        self._load_settings()
        
        # Emit signal that settings have changed
        self.settings_changed.emit()
        
        logger.info("Preferences saved")
    
    def _mark_as_changed(self):
        """Mark the dialog as having unsaved changes"""
        self.has_changes = True
        self.apply_button.setEnabled(True)
    
    def _apply_settings(self):
        """Apply settings without closing the dialog"""
        self._save_settings()
    
    def _browse_import_dir(self):
        """Browse for import directory"""
        directory = QFileDialog.getExistingDirectory(
            self,
            "Select Import Directory",
            self.import_dir_edit.text() or os.path.expanduser("~")
        )
        
        if directory:
            self.import_dir_edit.setText(directory)
            self._mark_as_changed()
    
    def _browse_export_dir(self):
        """Browse for export directory"""
        directory = QFileDialog.getExistingDirectory(
            self,
            "Select Export Directory",
            self.export_dir_edit.text() or os.path.expanduser("~")
        )
        
        if directory:
            self.export_dir_edit.setText(directory)
            self._mark_as_changed()
    
    def accept(self):
        """Handle dialog acceptance (OK button)"""
        if self.has_changes:
            self._save_settings()
        super().accept()
    
    def reject(self):
        """Handle dialog rejection (Cancel button)"""
        if self.has_changes:
            # Ask for confirmation if there are unsaved changes
            from PyQt5.QtWidgets import QMessageBox
            reply = QMessageBox.question(
                self,
                "Unsaved Changes",
                "You have unsaved changes. Are you sure you want to discard them?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No
            )
            
            if reply == QMessageBox.No:
                return
        
        super().reject()

