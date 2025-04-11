#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Main Window
==========

The main application window for Altium Rule Generator.
Contains the main UI layout and controls.
"""

import os
import sys
import logging
from PyQt5.QtWidgets import (QMainWindow, QTabWidget, QAction, QFileDialog,
                             QMenu, QToolBar, QStatusBar, QMessageBox,
                             QDockWidget, QVBoxLayout, QHBoxLayout, QWidget)
from PyQt5.QtCore import Qt, QSize
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import QInputDialog

logger = logging.getLogger(__name__)

class MainWindow(QMainWindow):
    """Main application window"""
    
    def __init__(self, config_manager, theme_manager, parent=None):
        """Initialize the main window"""
        super().__init__(parent)
        self.config = config_manager
        self.theme_manager = theme_manager
        
        # Set window properties
        self.setWindowTitle("Altium Rule Generator")
        self.setMinimumSize(1000, 700)
        
        # Initialize UI components
        self._init_ui()
        
        # Restore window geometry
        self._restore_geometry()
        
        logger.info("Main window initialized")
    
    def _init_ui(self):
        """Initialize UI components"""
        # Create central widget with tab container
        self.tab_widget = QTabWidget()
        self.setCentralWidget(self.tab_widget)
        
        # Create menu bar
        self._create_menus()
        
        # Create toolbar
        self._create_toolbar()
        
        # Create status bar
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("Ready", 5000)
        
        # Add tabs to the tab widget
        self._add_tabs()

    def _add_tabs(self):
        """Add tabs to the main tab widget"""
        # Pivot Table Tab
        from gui.pivot_table_widget import PivotTableWidget
        pivot_tab = PivotTableWidget()
        self.tab_widget.addTab(pivot_tab, "Pivot Table")

        # Rule Editor Tab
        from gui.rule_editor_widget import RuleEditorWidget
        rule_editor_tab = RuleEditorWidget()
        self.tab_widget.addTab(rule_editor_tab, "Rule Editor")

        # Additional tabs can be added here as needed

    def _create_menus(self):
        """Create application menus"""
        # File menu
        self.file_menu = self.menuBar().addMenu("&File")
        
        # Import submenu
        import_menu = QMenu("&Import", self)
        self.file_menu.addMenu(import_menu)
        
        # Import actions
        import_excel_action = QAction("Import Excel File", self)
        import_excel_action.triggered.connect(self._import_excel)
        import_menu.addAction(import_excel_action)
        
        import_rul_action = QAction("Import RUL File", self)
        import_rul_action.triggered.connect(self._import_rul)
        import_menu.addAction(import_rul_action)
        
        # Export submenu
        export_menu = QMenu("&Export", self)
        self.file_menu.addMenu(export_menu)
        
        # Export actions
        export_excel_action = QAction("Export to Excel", self)
        export_excel_action.triggered.connect(self._export_excel)
        export_menu.addAction(export_excel_action)
        
        export_rul_action = QAction("Export RUL File", self)
        export_rul_action.triggered.connect(self._export_rul)
        export_menu.addAction(export_rul_action)
        
        self.file_menu.addSeparator()
        
        # Exit action
        exit_action = QAction("E&xit", self)
        exit_action.triggered.connect(self.close)
        self.file_menu.addAction(exit_action)
        
        # Edit menu
        self.edit_menu = self.menuBar().addMenu("&Edit")
        
        # Preferences action
        preferences_action = QAction("&Preferences", self)
        preferences_action.triggered.connect(self._show_preferences)
        self.edit_menu.addAction(preferences_action)
        
        # View menu
        self.view_menu = self.menuBar().addMenu("&View")
        
        # Theme submenu
        theme_menu = QMenu("&Theme", self)
        self.view_menu.addMenu(theme_menu)
        
        # Theme actions
        for theme_id, theme_name in self.theme_manager.get_available_themes().items():
            theme_action = QAction(theme_name, self)
            theme_action.setCheckable(True)
            theme_action.setChecked(theme_id == self.theme_manager.get_current_theme())
            theme_action.setData(theme_id)
            theme_action.triggered.connect(lambda checked, theme=theme_id: self.theme_manager.set_theme(theme))
            theme_menu.addAction(theme_action)
        
        # Help menu
        self.help_menu = self.menuBar().addMenu("&Help")
        
        # About action
        about_action = QAction("&About", self)
        about_action.triggered.connect(self._show_about)
        self.help_menu.addAction(about_action)
    
    def _create_toolbar(self):
        """Create main toolbar"""
        self.toolbar = QToolBar("Main Toolbar")
        self.toolbar.setMovable(False)
        self.addToolBar(Qt.TopToolBarArea, self.toolbar)
        
        # Import Excel button
        import_excel_action = QAction("Import Excel", self)
        import_excel_action.triggered.connect(self._import_excel)
        self.toolbar.addAction(import_excel_action)
        
        # Import RUL button
        import_rul_action = QAction("Import RUL", self)
        import_rul_action.triggered.connect(self._import_rul)
        self.toolbar.addAction(import_rul_action)
        
        self.toolbar.addSeparator()
        
        # Export Excel button
        export_excel_action = QAction("Export Excel", self)
        export_excel_action.triggered.connect(self._export_excel)
        self.toolbar.addAction(export_excel_action)
        
        # Export RUL button
        export_rul_action = QAction("Export RUL", self)
        export_rul_action.triggered.connect(self._export_rul)
        self.toolbar.addAction(export_rul_action)
    
    def _init_tabs(self):
        """Initialize tab widgets for different rule types"""
        # Placeholder for tabs - will be implemented in more detail later
        rule_types = [
            "Electrical Clearance",
            "Short-Circuit",
            "Un-Routed Net",
            "Un-Connected Pin",
            "Modified Polygon",
            "Creepage Distance"
        ]
        
        # Create placeholder tabs
        for rule_type in rule_types:
            tab = QWidget()
            layout = QVBoxLayout()
            tab.setLayout(layout)
            self.tab_widget.addTab(tab, rule_type)
        
        logger.info("Initialized tabs for rule types")
    
    def _restore_geometry(self):
        """Restore window size and position from settings"""
        # This will be implemented later to save/restore window state
        pass
    
    def _import_excel(self):
        """Import data from Excel file"""
        # Get file path
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Import Excel File",
            self.config.get("last_directory", ""),
            "Excel Files (*.xlsx *.xls);;All Files (*)"
        )
        
        if not file_path:
            return  # User cancelled
            
        # Update last directory
        self.config.update_last_directory(os.path.dirname(file_path))
        
        try:
            # Import Excel file
            from services.excel_importer import ExcelImporter
            excel_importer = ExcelImporter()
            
            # Get sheet names
            sheet_names = excel_importer.get_sheet_names(file_path)
            
            # If multiple sheets, ask user which one to import
            sheet_name = None
            if len(sheet_names) > 1:
                sheet_name, ok = QInputDialog.getItem(self, "Select Sheet", "Sheet Name:", sheet_names, 0, False)
                
                if not ok:
                    return  # User cancelled
            else:
                sheet_name = sheet_names[0]

            # Import raw Excel data for preview
            raw_df = excel_importer.import_file(file_path, sheet_name)

            # Show preview dialog
            from src.gui.excel_preview_dialog import ExcelPreviewDialog
            preview_dialog = ExcelPreviewDialog(raw_df, sheet_name, self)

            # If user cancels preview, abort import
            if not preview_dialog.exec_():
                return

            # Get processed dataframe and options from preview
            processed_df = preview_dialog.get_processed_dataframe()
            import_options = preview_dialog.get_import_options()

            # Allow user to set range of data to import
            start_row, ok_start = QInputDialog.getInt(self, "Set Start Row", "Enter the start row (1-based index):", 1, 1, processed_df.shape[0])
            if not ok_start:
                return

            end_row, ok_end = QInputDialog.getInt(self, "Set End Row", "Enter the end row (1-based index):", processed_df.shape[0], 1, processed_df.shape[0])
            if not ok_end:
                return

            # Adjust for 0-based indexing in pandas
            processed_df = processed_df.iloc[start_row-1:end_row]

            # Update tabs dynamically based on imported Excel data
            self._update_tabs_with_excel_data(processed_df)

            self.status_bar.showMessage(f"Successfully imported {os.path.basename(file_path)}", 5000)
            QMessageBox.information(self, "Import Successful", 
                                  f"Successfully imported {os.path.basename(file_path)}.\n\n"
                                  f"Sheet: {sheet_name}\n"
                                  f"Rows: {processed_df.shape[0]}\n",
                                  f"Columns: {processed_df.shape[1]}\n\n"
                                  f"Implementation of displaying the data is coming soon.")
            
        except Exception as e:
            error_msg = f"Error importing Excel file: {str(e)}"
            logger.error(error_msg)
            QMessageBox.critical(self, "Import Error", error_msg)
            self.status_bar.showMessage("Import failed", 5000)
    
    def _import_rul(self):
        """Import data from RUL file"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Import RUL File",
            self.config.get("last_directory", ""),
            "RUL Files (*.RUL *.rul);;All Files (*)"
        )
        
        if file_path:
            self.config.update_last_directory(os.path.dirname(file_path))
            self.config.add_recent_file(file_path)
            self.status_bar.showMessage(f"Importing RUL file: {os.path.basename(file_path)}", 5000)
            logger.info(f"Importing RUL file: {file_path}")
            # TODO: Implement actual import logic
    
    def _export_excel(self):
        """Export data to Excel file"""
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Export to Excel",
            self.config.get("last_directory", ""),
            "Excel Files (*.xlsx);;All Files (*)"
        )
        
        if file_path:
            self.config.update_last_directory(os.path.dirname(file_path))
            self.status_bar.showMessage(f"Exporting to Excel file: {os.path.basename(file_path)}", 5000)
            logger.info(f"Exporting to Excel file: {file_path}")
            # TODO: Implement actual export logic
    
    def _export_rul(self):
        """Export data to RUL file"""
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Export RUL File",
            self.config.get("last_directory", ""),
            "RUL Files (*.RUL);;All Files (*)"
        )
        
        if file_path:
            self.config.update_last_directory(os.path.dirname(file_path))
            self.status_bar.showMessage(f"Exporting to RUL file: {os.path.basename(file_path)}", 5000)
            logger.info(f"Exporting to RUL file: {file_path}")
            # TODO: Implement actual export logic
    
    def _show_preferences(self):
        """Show preferences dialog"""
        # TODO: Implement preferences dialog
        logger.info("Showing preferences dialog")
    
    def _show_about(self):
        """Show about dialog"""
        QMessageBox.about(
            self,
            "About Altium Rule Generator",
            """<h1>Altium Rule Generator</h1>
            <p>Version 1.0.0</p>
            <p>A tool for managing Altium Designer rule files.</p>
            <p>Converts between Excel pivot tables and Altium .RUL files.</p>"""
        )
    
    def closeEvent(self, event):
        """Handle window close event"""
        # TODO: Implement check for unsaved changes
        event.accept()    def _update_tabs_with_excel_data(self, dataframe):
        """Update tabs with data from the imported Excel file"""
        # Clear existing tabs except the default ones
        while self.tab_widget.count() > 2:  # Assuming first two tabs are default
            self.tab_widget.removeTab(2)

        # Check if 'Rule Set' column exists, if not, create a default one
        if 'Rule Set' not in dataframe.columns:
            logger.warning("No 'Rule Set' column found in imported data. Creating a default rule set.")
            dataframe['Rule Set'] = "Default Rule Set"

        # Create a tab for each unique rule set in the dataframe
        rule_sets = dataframe['Rule Set'].unique()
        for rule_set in rule_sets:
            tab = QWidget()
            layout = QVBoxLayout()
            tab.setLayout(layout)

            # Filter data for the current rule set
            rule_data = dataframe[dataframe['Rule Set'] == rule_set]

            # Create a PivotTableWidget to display the data
            from gui.pivot_table_widget import PivotTableWidget
            pivot_widget = PivotTableWidget()
            pivot_widget.set_pivot_data(rule_data)

            layout.addWidget(pivot_widget)
            self.tab_widget.addTab(tab, rule_set)
