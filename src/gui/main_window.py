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
                             QDockWidget, QVBoxLayout, QHBoxLayout, QWidget,
                             QShortcut, QApplication, QInputDialog)
from PyQt5.QtGui import QFont, QIcon, QKeySequence
from PyQt5.QtCore import Qt, QSize

from models.excel_data import ExcelPivotData
from models.rule_model import RuleType, UnitType

logger = logging.getLogger(__name__)

class MainWindow(QMainWindow):
    """Main application window"""
    
    def __init__(self, config_manager, theme_manager, parent=None):
        """Initialize the main window"""
        super().__init__(parent)
        self.config = config_manager
        self.theme_manager = theme_manager
        self.has_unsaved_changes = False

        # Set default font for the application
        default_font = QFont("Arial", 10)
        QApplication.setFont(default_font)

        # Ensure button text is visible
        self.setStyleSheet("QPushButton { text-align: center; }")

        # Set up icons path
        self.icons_path = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), 
                                      "resources", "icons")
        if not os.path.exists(self.icons_path):
            os.makedirs(self.icons_path, exist_ok=True)
            logger.info(f"Created icons directory at {self.icons_path}")

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
        
        # Create UI components
        self._create_menus()
        self._create_toolbar()
        
        # Create status bar
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("Ready", 5000)
        
        # Add tabs to the tab widget
        self._add_tabs()
        
        # Set up additional keyboard shortcuts
        self._setup_shortcuts()

    def _add_tabs(self):
        """Add tabs to the main tab widget"""
        # Pivot Table Tab
        from gui.pivot_table_widget import PivotTableWidget
        self.pivot_tab = PivotTableWidget()
        self.pivot_tab.data_changed.connect(self._on_data_changed)
        self.tab_widget.addTab(self.pivot_tab, "Pivot Table")

        # Rule Editor Tab
        from gui.rule_editor_widget import RulesManagerWidget
        self.rule_editor_tab = RulesManagerWidget()
        self.rule_editor_tab.rules_changed.connect(self._on_data_changed)
        self.tab_widget.addTab(self.rule_editor_tab, "Rule Editor")

        # Connect pivot data updates between tabs
        self.rule_editor_tab.pivot_data_updated.connect(self._on_rule_pivot_updated)
        
    def _setup_shortcuts(self):
        """Set up additional keyboard shortcuts beyond those in menus/toolbars"""
        # Tab navigation shortcuts
        next_tab_shortcut = QShortcut(QKeySequence("Ctrl+Tab"), self)
        next_tab_shortcut.activated.connect(self._next_tab)
        
        prev_tab_shortcut = QShortcut(QKeySequence("Ctrl+Shift+Tab"), self)
        prev_tab_shortcut.activated.connect(self._prev_tab)
    
    def _next_tab(self):
        """Switch to the next tab"""
        current = self.tab_widget.currentIndex()
        if current < self.tab_widget.count() - 1:
            self.tab_widget.setCurrentIndex(current + 1)
        else:
            self.tab_widget.setCurrentIndex(0)
    
    def _prev_tab(self):
        """Switch to the previous tab"""
        current = self.tab_widget.currentIndex()
        if current > 0:
            self.tab_widget.setCurrentIndex(current - 1)
        else:
            self.tab_widget.setCurrentIndex(self.tab_widget.count() - 1)

    def _create_menus(self):
        """Create application menus"""
        self._create_file_menu()
        self._create_edit_menu()
        self._create_view_menu()
        self._create_help_menu()
    
    def _create_file_menu(self):
        """Create file menu and actions"""
        self.file_menu = self.menuBar().addMenu("&File")
        
        # Import submenu
        import_menu = QMenu("&Import", self)
        import_menu.setIcon(QIcon(os.path.join(self.icons_path, "import.png")))
        self.file_menu.addMenu(import_menu)
        
        # Import actions
        self._add_action(import_menu, "Import Excel File", "excel.png", "Ctrl+I", 
                         "Import data from Excel file (Ctrl+I)", self._import_excel)
        
        self._add_action(import_menu, "Import RUL File", "rul.png", "Ctrl+R",
                         "Import data from Altium RUL file (Ctrl+R)", self._import_rul)
        
        # Export submenu
        export_menu = QMenu("&Export", self)
        export_menu.setIcon(QIcon(os.path.join(self.icons_path, "export.png")))
        self.file_menu.addMenu(export_menu)
        
        # Export actions
        self._add_action(export_menu, "Export to Excel", "excel.png", "Ctrl+E",
                         "Export data to Excel file (Ctrl+E)", self._export_excel)
        
        self._add_action(export_menu, "Export RUL File", "rul.png", "Ctrl+S",
                         "Export data to Altium RUL file (Ctrl+S)", self._export_rul)
        
        self.file_menu.addSeparator()
        
        # Exit action
        self._add_action(self.file_menu, "E&xit", "exit.png", "Alt+F4", 
                         "Exit the application (Alt+F4)", self.close)
    
    def _create_edit_menu(self):
        """Create edit menu and actions"""
        self.edit_menu = self.menuBar().addMenu("&Edit")
        
        # Preferences action
        self._add_action(self.edit_menu, "&Preferences", "settings.png", "Ctrl+P",
                         "Open application preferences (Ctrl+P)", self._show_preferences)
    
    def _create_view_menu(self):
        """Create view menu and actions"""
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
    
    def _create_help_menu(self):
        """Create help menu and actions"""
        self.help_menu = self.menuBar().addMenu("&Help")
        
        # About action
        self._add_action(self.help_menu, "&About", "about.png", None,
                         "Show information about the application", self._show_about)
    
    def _add_action(self, parent, text, icon_name, shortcut, tooltip, callback):
        """Helper method to create and add actions"""
        action = QAction(text, self)
        
        if icon_name:
            action.setIcon(QIcon(os.path.join(self.icons_path, icon_name)))
        
        if shortcut:
            action.setShortcut(QKeySequence(shortcut))
        
        action.setToolTip(tooltip)
        action.triggered.connect(callback)
        parent.addAction(action)
        return action

    def _create_toolbar(self):
        """Create main toolbar"""
        self.toolbar = QToolBar("Main Toolbar")
        self.toolbar.setMovable(False)
        self.addToolBar(Qt.TopToolBarArea, self.toolbar)
        
        # Set toolbar to show both icon and text
        self.toolbar.setToolButtonStyle(Qt.ToolButtonTextUnderIcon)
        
        # Set toolbar icon size to be larger
        self.toolbar.setIconSize(QSize(32, 32))
        
        # Add toolbar actions
        self._add_toolbar_action("excel_import.png", "Import Excel", "Ctrl+I", 
                                 "Import data from Excel file (Ctrl+I)", self._import_excel)
        
        self._add_toolbar_action("rul_import.png", "Import RUL", "Ctrl+R",
                                 "Import data from Altium RUL file (Ctrl+R)", self._import_rul)
        
        self.toolbar.addSeparator()
        
        self._add_toolbar_action("excel_export.png", "Export Excel", "Ctrl+E",
                                 "Export data to Excel file (Ctrl+E)", self._export_excel)
        
        self._add_toolbar_action("rul_export.png", "Export RUL", "Ctrl+S",
                                 "Export data to Altium RUL file (Ctrl+S)", self._export_rul)
    
    def _add_toolbar_action(self, icon_name, text, shortcut, tooltip, callback):
        """Helper method to create and add toolbar actions"""
        action = QAction(QIcon(os.path.join(self.icons_path, icon_name)), text, self)
        if shortcut:
            action.setShortcut(QKeySequence(shortcut))
        action.setToolTip(tooltip)
        action.triggered.connect(callback)
        self.toolbar.addAction(action)
        return action
    
    def _save_geometry(self):
        """Save window size and position to settings"""
        try:
            # Save window geometry
            geometry = self.saveGeometry()
            self.config.set("window_geometry", geometry.toBase64().data().decode())
            
            # Save window state
            state = self.saveState()
            self.config.set("window_state", state.toBase64().data().decode())
            
            # Save current tab index
            current_tab = self.tab_widget.currentIndex()
            self.config.set("current_tab", current_tab)
            
            logger.info("Saved window geometry and state")
        except Exception as e:
            logger.error(f"Error saving window geometry: {str(e)}")
    
    def _restore_geometry(self):
        """Restore window size and position from settings"""
        try:
            # Restore window geometry if it exists
            if self.config.contains("window_geometry"):
                geometry_bytes = self.config.get("window_geometry", "")
                if geometry_bytes:
                    from PyQt5.QtCore import QByteArray
                    geometry = QByteArray.fromBase64(geometry_bytes.encode())
                    self.restoreGeometry(geometry)
            
            # Restore window state if it exists
            if self.config.contains("window_state"):
                state_bytes = self.config.get("window_state", "")
                if state_bytes:
                    from PyQt5.QtCore import QByteArray
                    state = QByteArray.fromBase64(state_bytes.encode())
                    self.restoreState(state)
            
            # Restore current tab if it exists
            if self.config.contains("current_tab"):
                current_tab = int(self.config.get("current_tab", 0))
                if 0 <= current_tab < self.tab_widget.count():
                    self.tab_widget.setCurrentIndex(current_tab)
            
            logger.info("Restored window geometry and state")
        except Exception as e:
            logger.error(f"Error restoring window geometry: {str(e)}")
            # If there's an error, use default position and size
            screen_geometry = QApplication.desktop().availableGeometry()
            window_geometry = self.geometry()
            x = (screen_geometry.width() - window_geometry.width()) // 2
            y = (screen_geometry.height() - window_geometry.height()) // 2
            self.move(x, y)
    
    def _import_excel(self):
        """Import data from Excel file"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Import Excel File",
            self.config.get("last_directory", ""),
            "Excel Files (*.xlsx *.xls);;All Files (*)"
        )
        
        if not file_path:
            return  # User cancelled
            
        # Update last directory        self.config.update_last_directory(os.path.dirname(file_path))
        
        try:
            self._process_excel_import(file_path)
        except Exception as e:
            error_msg = f"Error importing Excel file: {str(e)}"
            logger.error(error_msg)
            QMessageBox.critical(self, "Import Error", error_msg)
            self.status_bar.showMessage("Import failed", 5000)
    
    def _process_excel_import(self, file_path):
        """Process Excel import from the given file path"""
        from services.excel_importer import ExcelImporter
        excel_importer = ExcelImporter()
        
        # Get sheet names
        sheet_names = excel_importer.get_sheet_names(file_path)
        
        # If multiple sheets, ask user which one to import
        sheet_name = self._get_sheet_selection(sheet_names)
        if not sheet_name:
            return  # User cancelled

        # Import raw Excel data for preview
        raw_df = excel_importer.import_file(file_path, sheet_name)

        # Show preview dialog using the existing helper method
        processed_df, import_options = self._show_excel_preview(raw_df, sheet_name)
        
        # If user cancels preview, abort import
        if processed_df is None or import_options is None:
            return
        
        if processed_df is None or processed_df.empty:
            QMessageBox.warning(self, "Import Warning", "No data to import after processing.")
            return

        # Validate dataframe
        if not self._validate_dataframe(processed_df):
            return

        # Get row range from user if needed
        start_row = import_options.get("start_row", 1)
        end_row = min(import_options.get("end_row", processed_df.shape[0]), processed_df.shape[0])
        
        # Adjust for 0-based indexing in pandas
        processed_df = processed_df.iloc[start_row-1:end_row]

        # Update tabs dynamically based on imported Excel data
        self._update_tabs_with_excel_data(processed_df)

        self.status_bar.showMessage(f"Successfully imported {os.path.basename(file_path)}", 5000)
        QMessageBox.information(self, "Import Successful", 
                              f"Successfully imported {os.path.basename(file_path)}.\n\n"
                              f"Sheet: {sheet_name}\n"
                              f"Rows: {processed_df.shape[0]}\n"
                              f"Columns: {processed_df.shape[1]}")
    
    def _get_sheet_selection(self, sheet_names):
        """Get sheet selection from user"""
        if len(sheet_names) > 1:
            sheet_name, ok = QInputDialog.getItem(self, "Select Sheet", "Sheet Name:", 
                                                 sheet_names, 0, False)
            if not ok:
                return None
            return sheet_name
        return sheet_names[0]
    
    def _show_excel_preview(self, raw_df, sheet_name):
        """Show Excel preview dialog and return processed dataframe and options"""
        from gui.excel_preview_dialog import ExcelPreviewDialog
        preview_dialog = ExcelPreviewDialog(raw_df, sheet_name, self)

        # If user cancels preview, abort import
        if not preview_dialog.exec_():
            return None, None

        # Get processed dataframe and options from preview
        processed_df = preview_dialog.get_processed_dataframe()
        import_options = preview_dialog.get_import_options()
        
        return processed_df, import_options
    
    def _validate_dataframe(self, df):
        """Validate the dataframe to ensure it has data"""
        if df.empty:
            QMessageBox.warning(self, "Import Warning", "The imported Excel file contains no data.")
            return False

        if df.shape[0] == 0:
            QMessageBox.warning(self, "Import Warning", "The imported Excel file contains no rows.")
            return False

        if df.shape[1] == 0:
            QMessageBox.warning(self, "Import Warning", "The imported Excel file contains no columns.")
            return False
            
        return True
    
    def _get_row_range(self, df):
        """Get row range selection from user"""
        start_row, ok_start = QInputDialog.getInt(self, "Set Start Row", 
                                                "Enter the start row (1-based index):", 
                                                1, 1, df.shape[0])
        if not ok_start:
            return None, None

        end_row, ok_end = QInputDialog.getInt(self, "Set End Row", 
                                            "Enter the end row (1-based index):", 
                                            df.shape[0], 1, df.shape[0])
        if not ok_end:
            return None, None
            
        return start_row, end_row
    
    def _import_rul(self):
        """Import data from RUL file"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Import RUL File",
            self.config.get("last_directory", ""),
            "RUL Files (*.RUL *.rul);;All Files (*)"
        )
        
        if not file_path:
            return  # User cancelled
        
        # Update last directory and add to recent files
        self.config.update_last_directory(os.path.dirname(file_path))
        self.config.add_recent_file(file_path)
        self.status_bar.showMessage(f"Importing RUL file: {os.path.basename(file_path)}", 5000)
        logger.info(f"Importing RUL file: {file_path}")
        
        try:
            self._process_rul_import(file_path)
        except Exception as e:
            error_msg = f"Error importing RUL file: {str(e)}"
            logger.error(error_msg)
            QMessageBox.critical(self, "Import Error", error_msg)
            self.status_bar.showMessage("Import failed", 5000)
    
    def _process_rul_import(self, file_path):
        """Process RUL import from the given file path"""
        from services.rule_generator import RuleGenerator, RuleGeneratorError
        rule_generator = RuleGenerator()
        
        try:
            # Parse RUL file
            success = rule_generator.parse_rul_file(file_path)
            
            if not success:
                QMessageBox.warning(
                    self, 
                    "Import Warning", 
                    "The RUL file was parsed but no valid rules were found."
                )
                return
            
            # Get all rules
            all_rules = rule_generator.rule_manager.rules
            
            if not all_rules:
                QMessageBox.warning(
                    self, 
                    "Import Warning", 
                    "No rules were found in the RUL file."
                )
                return
            
            # Get clearance rules
            clearance_rules = rule_generator.get_rules_by_type(RuleType.CLEARANCE)
            
            if not clearance_rules:
                # If no clearance rules, show warning but continue with other rule types
                logger.warning("No clearance rules found in the RUL file.")
                QMessageBox.information(
                    self, 
                    "Import Information", 
                    "No clearance rules were found in the RUL file. Other rule types will be imported."
                )
            
            # Update the tabs with the imported rules
            self._update_tabs_with_rules(all_rules, clearance_rules)
            
            # Show success message
            self._show_rul_import_success(file_path, all_rules, clearance_rules)
            
        except RuleGeneratorError as e:
            error_msg = f"Error parsing RUL file: {str(e)}"
            logger.error(error_msg)
            QMessageBox.critical(self, "Import Error", error_msg)
            self.status_bar.showMessage("Import failed", 5000)
    
    def _update_tabs_with_rules(self, all_rules, clearance_rules):
        """Update tabs based on imported rules"""
        # Clear existing tabs except the default ones
        while self.tab_widget.count() > 2:  # Assuming first two tabs are default
            self.tab_widget.removeTab(2)
        
        # Create a tab for clearance rules if any exist
        if clearance_rules:
            # Create pivot data from clearance rules
            pivot_data = ExcelPivotData.from_clearance_rules(clearance_rules)
            
            if pivot_data:
                # Create a new tab for clearance rules
                tab = QWidget()
                layout = QVBoxLayout()
                tab.setLayout(layout)
                
                # Create pivot table widget
                from gui.pivot_table_widget import PivotTableWidget
                pivot_widget = PivotTableWidget()
                pivot_widget.set_pivot_data(pivot_data)
                
                # Connect the data_changed signal to a handler
                pivot_widget.data_changed.connect(self._on_pivot_data_changed)

                layout.addWidget(pivot_widget)
                self.tab_widget.addTab(tab, "Clearance Rules")
                
                logger.info(f"Added tab for clearance rules with {len(clearance_rules)} rules")
        
        # Switch to the newly added tab
        if self.tab_widget.count() > 2:
            self.tab_widget.setCurrentIndex(2)
    
    def _show_rul_import_success(self, file_path, all_rules, clearance_rules):
        """Show success message for RUL import"""
        self.status_bar.showMessage(f"Successfully imported {os.path.basename(file_path)}", 5000)
        
        rule_types = set(rule.rule_type.value for rule in all_rules)
        rule_types_str = ", ".join(rule_types)
        
        QMessageBox.information(
            self, 
            "Import Successful", 
            f"Successfully imported {len(all_rules)} rules from {os.path.basename(file_path)}.\n\n"
            f"Rule types: {rule_types_str}\n"
            f"Clearance rules: {len(clearance_rules)}"
        )
    
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
        
        if not file_path:
            return  # User cancelled
            
        # Ensure file has .RUL extension
        if not file_path.upper().endswith(".RUL"):
            file_path += ".RUL"
        
        self.config.update_last_directory(os.path.dirname(file_path))
        
        try:
            self._process_rul_export(file_path)
        except Exception as e:
            error_msg = f"Error exporting to RUL file: {str(e)}"
            logger.error(error_msg)
            QMessageBox.critical(self, "Export Error", error_msg)
            self.status_bar.showMessage("Export failed", 5000)
    
    def _process_rul_export(self, file_path):
        """Process RUL export to the given file path"""
        # Get the pivot data from the active tab
        pivot_data = self._get_pivot_data_from_active_tab()
        
        if not pivot_data:
            QMessageBox.warning(
                self,
                "Export Warning",
                "No pivot data available for export. Please import or create data first."
            )
            return
        
        # Create rules from pivot data
        rules = self._create_rules_from_pivot_data(pivot_data)
        
        if not rules:
            QMessageBox.warning(
                self,
                "Export Warning",
                "No rules could be created from the pivot data. Ensure the table has valid data."
            )
            return
        
        # Export rules to file
        self._save_rules_to_rul_file(rules, file_path)
    
    def _get_pivot_data_from_active_tab(self):
        """Get pivot data from the active tab"""
        current_tab_index = self.tab_widget.currentIndex()
        current_tab = self.tab_widget.widget(current_tab_index)
        
        # If current tab is the default pivot table tab
        if current_tab_index == 0:
            return self.pivot_tab.get_pivot_data()
        
        # Look for PivotTableWidget in the layout
        for i in range(current_tab.layout().count()):
            widget = current_tab.layout().itemAt(i).widget()
            if widget and widget.__class__.__name__ == "PivotTableWidget":
                return widget.get_pivot_data()
        
        return None
    
    def _create_rules_from_pivot_data(self, pivot_data):
        """Create rules from pivot data based on rule type"""
        if pivot_data.rule_type == RuleType.CLEARANCE:
            return pivot_data.to_clearance_rules()
        elif pivot_data.rule_type == RuleType.SHORT_CIRCUIT:
            return pivot_data.to_short_circuit_rules()
        elif pivot_data.rule_type == RuleType.UNROUTED_NET:
            return pivot_data.to_unrouted_net_rules()
        else:
            # Default to clearance rules if rule type not set or recognized
            logger.warning(f"Unknown rule type: {pivot_data.rule_type}. Using clearance rules.")
            return pivot_data.to_clearance_rules()
    
    def _save_rules_to_rul_file(self, rules, file_path):
        """Save rules to RUL file"""
        from services.rule_generator import RuleGenerator, RuleGeneratorError
        rule_generator = RuleGenerator()
        
        # Add rules to rule generator
        rule_generator.add_rules(rules)
        
        # Save to file
        try:
            success = rule_generator.save_to_file(file_path)
            
            if success:
                self.status_bar.showMessage(f"Successfully exported to {os.path.basename(file_path)}", 5000)
                QMessageBox.information(
                    self,
                    "Export Successful",
                    f"Successfully exported {len(rules)} rules to {os.path.basename(file_path)}."
                )
                # Mark data as saved
                self._on_data_saved()
            else:
                QMessageBox.warning(
                    self,
                    "Export Warning",
                    "The file was saved but there may have been issues. Check the log for details."
                )
        except RuleGeneratorError as e:
            raise Exception(f"Error saving RUL file: {str(e)}")

    def _show_preferences(self):
        """Show preferences dialog"""
        try:
            from gui.preferences_dialog import PreferencesDialog
            dialog = PreferencesDialog(self.config, self.theme_manager, self)
            
            # Connect the settings_changed signal
            dialog.settings_changed.connect(self._on_preferences_changed)
            
            # Show the dialog
            dialog.exec_()
            
            logger.info("Preferences dialog shown")
        except Exception as e:
            error_msg = f"Error showing preferences dialog: {str(e)}"
            logger.error(error_msg)
            QMessageBox.critical(self, "Error", error_msg)
    
    def _on_preferences_changed(self):
        """Handle preferences changing"""
        # Update UI based on preferences
        show_statusbar = self.config.get("show_statusbar", True)
        self.status_bar.setVisible(show_statusbar)
        
        show_toolbar = self.config.get("show_toolbar", True)
        self.toolbar.setVisible(show_toolbar)
        
        logger.info("Applied preferences changes")
    
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
        
    def _on_data_changed(self):
        """Handle data changes in the application"""
        self.has_unsaved_changes = True
        # Update window title to indicate unsaved changes
        self.setWindowTitle("*Altium Rule Generator")
    
    def _on_data_saved(self):
        """Handle data saved in the application"""
        self.has_unsaved_changes = False
        # Update window title to remove unsaved indicator
        self.setWindowTitle("Altium Rule Generator")
    
    def _on_rule_pivot_updated(self, pivot_data):
        """Handle pivot data updates from rule editor"""
        if pivot_data and self.pivot_tab:
            # Update the pivot table tab with the new data
            self.pivot_tab.set_pivot_data(pivot_data)
            
            # Switch to the pivot table tab
            self.tab_widget.setCurrentIndex(0)
            
            logger.info("Updated pivot table from rule editor")
    
    def closeEvent(self, event):
        """Handle window close event"""
        if self.has_unsaved_changes:
            # Ask user if they want to save changes
            reply = QMessageBox.question(
                self,
                "Unsaved Changes",
                "You have unsaved changes. Do you want to save before closing?",
                QMessageBox.Save | QMessageBox.Discard | QMessageBox.Cancel,
                QMessageBox.Save
            )
            
            if reply == QMessageBox.Save:
                # Determine what to save based on active tab
                current_tab = self.tab_widget.currentIndex()
                if current_tab == 0 or current_tab == 1:  # Pivot Table or Rule Editor tab
                    self._export_rul()
                
                # Only close if save was successful (check if still has unsaved changes)
                if self.has_unsaved_changes:
                    event.ignore()
                    return
            elif reply == QMessageBox.Cancel:
                # User canceled, don't close
                event.ignore()
                return
                
        # Save window geometry before closing
        self._save_geometry()
        
        # Accept the close event
        event.accept()
        
    def _update_tabs_with_excel_data(self, dataframe):
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
            
            # Convert DataFrame to ExcelPivotData format
            try:
                # Create a pivot data object with the rule type set to Clearance by default
                pivot_data = ExcelPivotData(RuleType.CLEARANCE)
                
                # Load the DataFrame into the pivot data structure
                # Default to MIL units if not specified
                unit = UnitType.MIL
                if 'Unit' in rule_data.columns:
                    unit_str = rule_data['Unit'].iloc[0] if not rule_data['Unit'].empty else 'mil'
                    try:
                        unit = UnitType.from_string(unit_str)
                    except ValueError:
                        logger.warning(f"Invalid unit type: {unit_str}, using MIL")
                
                # Load the dataframe into the pivot data
                success = pivot_data.load_dataframe(rule_data, unit)
                
                if success:
                    # Set the pivot data in the widget
                    pivot_widget.set_pivot_data(pivot_data)
                    logger.info(f"Successfully loaded data for rule set: {rule_set}")
                else:
                    logger.warning(f"Failed to load pivot data for rule set: {rule_set}")
            except Exception as e:
                logger.error(f"Error converting DataFrame to pivot data: {str(e)}")

            layout.addWidget(pivot_widget)
            self.tab_widget.addTab(tab, str(rule_set))
            
        # Switch to the first new tab if any were created
        if self.tab_widget.count() > 2:
            self.tab_widget.setCurrentIndex(2)
