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
import re # Import re for regex matching
from datetime import datetime # Import datetime
from typing import List, Optional # For type hinting
from PyQt5.QtWidgets import (QMainWindow, QTabWidget, QAction, QFileDialog,
                             QMenu, QToolBar, QStatusBar, QMessageBox,
                             QDockWidget, QVBoxLayout, QHBoxLayout, QWidget,
                             QShortcut, QApplication, QInputDialog, QActionGroup)
from PyQt5.QtGui import QFont, QIcon, QKeySequence
from PyQt5.QtCore import Qt, QSize, pyqtSignal

from models.excel_data import ExcelPivotData
from models.rule_model import RuleType, UnitType, BaseRule # Import BaseRule
# Import RuleManager directly
from models.rule_model import RuleManager

from gui.preferences_dialog import PreferencesDialog # Add import for PreferencesDialog
# Import RulesManagerWidget instead of RuleEditorWidget
from gui.rule_editor_widget import RulesManagerWidget
from services.rule_generator import RuleGenerator, RuleGeneratorError # Add import for RuleGenerator

logger = logging.getLogger(__name__)

class MainWindow(QMainWindow):
    """Main application window"""
    # Signal emitted when a tab's data changes significantly enough to warrant a save prompt
    unsaved_changes_changed = pyqtSignal(bool)

    def __init__(self, config_manager, theme_manager, parent=None):
        """Initialize the main window"""
        super().__init__(parent)
        self.config = config_manager
        self.theme_manager = theme_manager
        # self.has_unsaved_changes = False # Removed, will use signal/slot or check tab state
        self.pivot_tab = None
        # Rename instance variable
        self.rules_manager_tab = None

        # Set default font for the application
        default_font = QFont("Arial", 10)
        QApplication.setFont(default_font)

        # Ensure button text is visible
        # self.setStyleSheet("QPushButton { text-align: center; }") # Remove or comment out if theme handles this

        # Apply base stylesheet including menu padding
        self.setStyleSheet("""
            QPushButton { 
                text-align: center; 
            }
            QMenu::item {
                padding: 5px 30px 5px 20px; /* Increased right padding (Top, Right, Bottom, Left) */
            }
            QMenuBar::item {
                padding: 5px 10px; /* Adjust spacing for top-level menu bar items */
            }
        """)

        # Set up icons path
        self.icons_path = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), 
                                      "resources", "icons")
        if not os.path.exists(self.icons_path):
            os.makedirs(self.icons_path, exist_ok=True)
            logger.info(f"Created icons directory at {self.icons_path}")

        self.setWindowTitle("Altium Rule Generator")
        self.setMinimumSize(1200, 800)
        
        # Initialize UI components
        self._init_ui()
        
        # Restore window geometry
        self._restore_geometry()
        
        logger.info("Main window initialized")
    
    def _init_ui(self):
        """Initialize UI components"""
        # Create central widget with tab container
        self.tab_widget = QTabWidget()
        self.tab_widget.setTabsClosable(True) # Make tabs closable
        self.tab_widget.tabCloseRequested.connect(self._close_tab) # Connect close signal
        self.setCentralWidget(self.tab_widget)
        
        # Create UI components
        self._create_menus()
        self._create_toolbar()
        
        # Create status bar
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("Ready", 5000)
        
        # Remove automatic tab adding
        # self._add_tabs()
        
        # Set up additional keyboard shortcuts
        self._setup_shortcuts()

    def _close_tab(self, index):
        """Close the tab at the given index."""
        widget = self.tab_widget.widget(index)
        if widget:
            # Check if the widget being closed has unsaved changes
            prompt_save = False
            tab_name = self.tab_widget.tabText(index)
            # Check if RulesManagerWidget implements has_unsaved_changes
            if hasattr(widget, 'has_unsaved_changes') and widget.has_unsaved_changes():
                reply = QMessageBox.question(self, f'Unsaved Changes in {tab_name}',
                                             f'The tab "{tab_name}" has unsaved changes. Do you want to save them before closing?',
                                             QMessageBox.Save | QMessageBox.Discard | QMessageBox.Cancel,
                                             QMessageBox.Cancel)

                if reply == QMessageBox.Save:
                    # Attempt to save based on tab type
                    save_successful = False
                    # Use renamed variable
                    if widget == self.rules_manager_tab:
                        save_successful = self._export_rul() # Returns True on success/cancel, False on failure
                    elif widget == self.pivot_tab:
                        # Decide what saving means for pivot table (e.g., export Excel?)
                        # save_successful = self._export_excel() # Example
                        QMessageBox.information(self, "Save Pivot Table", "Saving the Pivot Table directly is not implemented. You can export it manually.")
                        save_successful = True # Allow closing without forcing save for now
                    else:
                         QMessageBox.warning(self, "Save Error", "Cannot determine how to save this tab.")
                         save_successful = False # Prevent closing if save logic unknown

                    if not save_successful:
                        return # Abort closing if save failed or was cancelled during save dialog
                elif reply == QMessageBox.Cancel:
                    return # Abort closing
                # If Discard, proceed to close without saving

            # --- Proceed with closing the tab ---

            # Disconnect signals
            if widget == self.pivot_tab:
                try:
                    if hasattr(self.pivot_tab, 'model') and self.pivot_tab.model:
                         self.pivot_tab.model.data_changed.disconnect(self._on_data_changed)
                    # Disconnect rules_generated signal
                    if hasattr(self.pivot_tab, 'rules_generated'):
                        self.pivot_tab.rules_generated.disconnect(self._handle_generated_rules)
                except TypeError: # Signal already disconnected
                    pass
                self.pivot_tab = None
                logger.info("Pivot Table tab closed.")
            # Use renamed variable
            elif widget == self.rules_manager_tab:
                try:
                    # Use correct signal names from RulesManagerWidget
                    self.rules_manager_tab.rules_changed.disconnect(self._on_data_changed)
                    self.rules_manager_tab.pivot_data_updated.disconnect(self._on_rule_pivot_updated)
                except TypeError:
                    pass
                # Use renamed variable
                self.rules_manager_tab = None
                logger.info("Rule Manager tab closed.")

            # Remove the tab
            self.tab_widget.removeTab(index)
            widget.deleteLater() # Ensure the widget is properly deleted
            logger.info(f"Closed tab: {tab_name}")
            # Update overall unsaved changes status after closing a tab
            self._check_unsaved_changes()

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
        # Restore the Export to Excel action
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
        
        # Update action text and callback if desired, but keep callback name for now
        self._add_action(self.view_menu, "Show &Rule Manager", "settings.png", None,
                         "Open the Rule Manager tab", self._show_rule_editor_tab) # Keep callback name for now
        
        self.view_menu.addSeparator()

        # Theme selection submenu
        self.theme_menu = self.view_menu.addMenu("&Themes")
        theme_group = QActionGroup(self)
        theme_group.setExclusive(True) # Only one theme can be active

        light_theme_action = self._add_action(self.theme_menu, "&Light", None, None,
                                              "Switch to Light Theme", lambda: self._change_theme("light"),
                                              checkable=True)
        dark_theme_action = self._add_action(self.theme_menu, "&Dark", None, None,
                                             "Switch to Dark Theme", lambda: self._change_theme("dark"),
                                             checkable=True)

        theme_group.addAction(light_theme_action)
        theme_group.addAction(dark_theme_action)

        # Set the initially checked theme action
        current_theme = self.theme_manager.get_current_theme()
        if current_theme == "dark":
            dark_theme_action.setChecked(True)
        else:
            light_theme_action.setChecked(True) # Default to light

    def _change_theme(self, theme_name):
        """Applies the selected theme."""
        try:
            self.theme_manager.apply_theme(theme_name)
            logger.info(f"Theme changed to: {theme_name}")
            # Update checkmarks in the theme menu
            if self.theme_menu: # Check if theme_menu exists before accessing actions
                for action in self.theme_menu.actions():
                    if action.isCheckable():
                        action.setChecked(action.data() == theme_name)
            else:
                logger.warning("Theme menu not found, cannot update checkmarks.")
        except Exception as e:
            logger.error(f"Error changing theme to {theme_name}: {e}")
            QMessageBox.warning(self, "Theme Error", f"Could not apply theme '{theme_name}'.")

    def _create_help_menu(self):
        """Create help menu and actions"""
        self.help_menu = self.menuBar().addMenu("&Help")
        
        # About action
        self._add_action(self.help_menu, "&About", "about.png", None,
                         "Show information about the application", self._show_about)
    
    def _show_about(self):
        """Show the About dialog."""
        QMessageBox.about(self, "About Altium Rule Generator",
                          "<b>Altium Rule Generator</b><br>" 
                          "Version 1.0.0<br><br>" 
                          "This application helps generate Altium Designer rules " 
                          "from structured data (e.g., Excel).<br><br>" 
                          "Author: Karl Long (klong4)<br>" 
                          "Company: eControls<br>"
                          "Email: klong@econtrols.com<br>"
                          "URL: <a href='https://github.com/klong4/AltiumXCEL2QueryBuilder'>https://github.com/klong4/AltiumXCEL2QueryBuilder</a><br><br>" 
                          "Copyright © 2025 Karl Long (klong4) / eControls")
        logger.info("Showed About dialog.")

    def _add_action(self, parent, text, icon_name, shortcut, tooltip, callback, checkable=False):
        """Helper method to create and add actions"""
        action = QAction(text, self)
        
        if icon_name:
            action.setIcon(QIcon(os.path.join(self.icons_path, icon_name)))
        
        if shortcut:
            action.setShortcut(QKeySequence(shortcut))
        
        action.setToolTip(tooltip)
        action.triggered.connect(callback)
        action.setCheckable(checkable)  # Set checkable state
        parent.addAction(action)
        return action

    def _create_toolbar(self):
        """Create main toolbar"""
        self.toolbar = QToolBar("Main Toolbar")
        self.toolbar.setObjectName("MainToolBar") # Set object name
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
        
        # Add Export to Excel toolbar action
        self._add_toolbar_action("excel_export.png", "Export Excel", "Ctrl+E",
                                 "Export pivot data to Excel file (Ctrl+E)", self._export_excel)

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
            # Use 'in' operator to check for key existence in the config dictionary
            if "window_geometry" in self.config.config:
                geometry_bytes = self.config.get("window_geometry", "")
                if geometry_bytes:
                    from PyQt5.QtCore import QByteArray
                    geometry = QByteArray.fromBase64(geometry_bytes.encode())
                    self.restoreGeometry(geometry)
            
            # Restore window state if it exists
            if "window_state" in self.config.config:
                state_bytes = self.config.get("window_state", "")
                if state_bytes:
                    from PyQt5.QtCore import QByteArray
                    state = QByteArray.fromBase64(state_bytes.encode())
                    self.restoreState(state)
            
            # Restore current tab if it exists
            if "current_tab" in self.config.config:
                current_tab = int(self.config.get("current_tab", 0))
                if 0 <= current_tab < self.tab_widget.count():
                    self.tab_widget.setCurrentIndex(current_tab)
            
            logger.info("Restored window geometry and state")
        except Exception as e:
            # Log the error with traceback for better debugging
            logger.error(f"Error restoring window geometry: {str(e)}", exc_info=True)
            # If there's an error, use default position and size
            # Ensure QApplication instance exists before accessing desktop()
            app = QApplication.instance()
            if app:
                screen_geometry = app.desktop().availableGeometry()
                window_geometry = self.geometry()
                x = (screen_geometry.width() - window_geometry.width()) // 2
                y = (screen_geometry.height() - window_geometry.height()) // 2
                self.move(x, y)
            else:
                logger.warning("QApplication instance not found, cannot center window.")

    def _get_file_path_dialog(self, dialog_type: str, title: str, 
                              directory_key: str = "last_directory", 
                              file_filter: str = "All Files (*)") -> str:
        """Helper method to show a file dialog (open or save) and return the selected path."""
        last_dir = self.config.get(directory_key, "")
        
        if dialog_type == "open":
            file_path, _ = QFileDialog.getOpenFileName(self, title, last_dir, file_filter)
        elif dialog_type == "save":
            file_path, _ = QFileDialog.getSaveFileName(self, title, last_dir, file_filter)
        else:
            logger.error(f"Invalid dialog type specified: {dialog_type}")
            return None

        if file_path:
            # Update the last used directory in the config
            self.config.set(directory_key, os.path.dirname(file_path))
            # Ensure .RUL extension for save dialog if needed (could be more generic)
            if dialog_type == "save" and "RUL Files" in file_filter and not file_path.upper().endswith('.RUL'):
                 file_path += '.RUL'
            return file_path
        else:
            return None # User cancelled

    def _import_excel(self):
        """Import data from Excel file"""
        file_path = self._get_file_path_dialog(
            dialog_type="open",
            title="Import Excel File",
            file_filter="Excel Files (*.xlsx *.xls);;All Files (*)"
        )
        
        if not file_path:
            return  # User cancelled
            
        # Update last directory - Handled by _get_file_path_dialog now
        # self.config.update_last_directory(os.path.dirname(file_path)) 
        
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
        
        try:
            # Get sheet names
            sheet_names = excel_importer.get_sheet_names(file_path)
        except Exception as e:
            error_msg = f"Error reading sheet names from {os.path.basename(file_path)}: {str(e)}"
            logger.error(error_msg)
            QMessageBox.critical(self, "Import Error", error_msg)
            return
        
        # If multiple sheets, ask user which one to import
        sheet_name = self._get_sheet_selection(sheet_names)
        if not sheet_name:
            logger.info("Excel import cancelled during sheet selection.")
            return  # User cancelled

        try:
            # Import raw Excel data for preview
            raw_df = excel_importer.import_file(file_path, sheet_name)
        except Exception as e:
            error_msg = f"Error reading data from sheet '{sheet_name}' in {os.path.basename(file_path)}: {str(e)}"
            logger.error(error_msg)
            QMessageBox.critical(self, "Import Error", error_msg)
            return

        # Show preview dialog using the existing helper method
        processed_df, import_options = self._show_excel_preview(raw_df, sheet_name)
        
        # If user cancels preview, abort import
        if processed_df is None or import_options is None:
            logger.info("Excel import cancelled during preview.")
            return
        
        if processed_df.empty:
            QMessageBox.warning(self, "Import Warning", "No data to import after processing.")
            logger.warning("Excel import resulted in empty dataframe after processing.")
            return

        # --- Create or Update Pivot Table Tab ---
        if self.pivot_tab is None:
            try:
                from gui.pivot_table_widget import PivotTableWidget
                self.pivot_tab = PivotTableWidget()
                # Connect signals
                if hasattr(self.pivot_tab, 'model') and self.pivot_tab.model:
                    self.pivot_tab.model.data_changed.connect(self._on_data_changed)
                else:
                    logger.warning("PivotTableWidget created without a model, cannot connect data_changed signal.")
                # Connect the rules_generated signal
                self.pivot_tab.rules_generated.connect(self._handle_generated_rules)

                index = self.tab_widget.addTab(self.pivot_tab, "Pivot Table")
                self.tab_widget.setCurrentIndex(index)
                logger.info("Pivot Table tab created.")
            except Exception as e:
                logger.error(f"Failed to create Pivot Table tab: {e}", exc_info=True)
                QMessageBox.critical(self, "Error", f"Could not create the Pivot Table tab: {e}")
                self.pivot_tab = None # Ensure it's None if creation failed
                return # Stop import if tab creation fails
        else:
            # Find the index of the existing pivot tab and switch to it
            for i in range(self.tab_widget.count()):
                if self.tab_widget.widget(i) == self.pivot_tab:
                    self.tab_widget.setCurrentIndex(i)
                    break
            logger.info("Switched to existing Pivot Table tab.")
            # Ensure signal is connected even if tab existed
            try:
                self.pivot_tab.rules_generated.disconnect(self._handle_generated_rules) # Disconnect first to avoid duplicates
            except TypeError:
                pass # Signal was not connected
            self.pivot_tab.rules_generated.connect(self._handle_generated_rules)


        # Load data into the pivot tab
        try:
            # Create an ExcelPivotData object and load the DataFrame
            pivot_data_obj = ExcelPivotData()
            # Assuming the unit needs to be determined or defaulted, e.g., UnitType.MIL
            # You might need to get the unit from the import options or elsewhere
            from models.rule_model import UnitType # Add import if not already present
            if pivot_data_obj.load_dataframe(processed_df, unit=UnitType.MIL): # Pass the DataFrame here
                self.pivot_tab.set_pivot_data(pivot_data_obj) # Pass the ExcelPivotData object
                logger.info(f"Loaded data ({processed_df.shape[0]}x{processed_df.shape[1]}) into Pivot Table tab.")
                # Update rule editor if it exists
                if self.rules_manager_tab and hasattr(self.pivot_tab, 'get_pivot_data'):
                    # get_pivot_data should return the ExcelPivotData object
                    current_pivot_data = self.pivot_tab.get_pivot_data() 
                    if current_pivot_data:
                        self.rules_manager_tab.update_pivot_data(current_pivot_data)
                        logger.info("Updated Rule Manager with new pivot data.")
            else:
                 raise ValueError("Failed to load DataFrame into ExcelPivotData object.")
        except Exception as e:
            error_msg = f"Error loading data into pivot table: {str(e)}"
            logger.error(error_msg, exc_info=True)
            QMessageBox.critical(self, "Load Error", error_msg)
            # Optionally close the tab if loading fails critically
            # self._close_tab(self.tab_widget.indexOf(self.pivot_tab))
            return

        # --- Update Status and Show Message --- 
        self.status_bar.showMessage(f"Successfully imported {os.path.basename(file_path)}", 5000)
        QMessageBox.information(self, "Import Successful", 
                              f"Successfully imported {os.path.basename(file_path)}.\n\n"
                              f"Sheet: {sheet_name}\n"
                              f"Rows: {processed_df.shape[0]}\n"
                              f"Columns: {processed_df.shape[1]}")
        self._check_unsaved_changes() # Check unsaved status after import
    
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
        """Import data from Altium RUL file"""
        file_path = self._get_file_path_dialog(
            dialog_type="open",
            title="Import RUL File",
            file_filter="RUL Files (*.RUL *.rul);;All Files (*)"
        )

        if not file_path:
            return # User cancelled

        # Update last directory - Handled by _get_file_path_dialog now
        # self.config.update_last_directory(os.path.dirname(file_path))

        try:
            # Ensure rule manager tab exists before importing RUL
            # Use renamed variable
            if self.rules_manager_tab is None:
                self._show_rule_editor_tab() # Create it if it doesn't exist
                # Use renamed variable
                if self.rules_manager_tab is None: # Check again if creation failed
                    raise RuleGeneratorError("Could not create or find the Rule Manager tab.")

            # Get the current rule manager or create a new one
            # Use renamed variable
            rule_manager = self.rules_manager_tab.get_rule_manager()
            if rule_manager is None:
                rule_manager = RuleManager()
                logger.info("Created a new RuleManager for RUL import.")

            # Use RuleGenerator to parse the file into the manager
            # (Assuming RuleGenerator needs a method like parse_rul_file)
            # This part might need adjustment based on RuleGenerator's capabilities
            try:
                # Placeholder: Assume RuleGenerator has a method to load rules into a manager
                # You might need to implement this method in RuleGenerator
                # generator = RuleGenerator()
                # generator.rule_manager = rule_manager # Assign the manager
                # Example: generator.parse_rul_file(file_path) # This method needs to exist
                # For now, let's assume RuleManager has an import method for simplicity here.
                # This depends heavily on how RUL parsing is implemented.
                if not hasattr(rule_manager, 'import_from_rul'):
                     raise NotImplementedError("RuleManager does not have an 'import_from_rul' method.")

                rule_manager.import_from_rul(file_path) # Assumes this method exists and handles parsing
                logger.info(f"Rules imported into RuleManager from {file_path}")

            except Exception as parse_error:
                 raise RuleGeneratorError(f"Failed to parse RUL file: {parse_error}") from parse_error


            # Set the updated rule manager back to the tab
            # Use renamed variable
            self.rules_manager_tab.set_rule_manager(rule_manager)

            self.status_bar.showMessage(f"Successfully imported {os.path.basename(file_path)}", 5000)
            QMessageBox.information(self, "Import Successful", f"Successfully imported RUL file: {file_path}")
            self._check_unsaved_changes() # Update unsaved status

        except RuleGeneratorError as rge:
             error_msg = f"Error importing RUL file: {str(rge)}"
             logger.error(error_msg)
             QMessageBox.critical(self, "Import Error", error_msg)
             self.status_bar.showMessage("RUL import failed", 5000)
        except Exception as e:
            error_msg = f"An unexpected error occurred during RUL import: {str(e)}"
            logger.error(error_msg, exc_info=True)
            QMessageBox.critical(self, "Import Error", error_msg)
            self.status_bar.showMessage("RUL import failed", 5000)


    def _export_excel(self):
        """Export pivot data to Excel file"""
        if self.pivot_tab is None:
            QMessageBox.warning(self, "Export Error", "No pivot table data to export.")
            return

        file_path = self._get_file_path_dialog(
            dialog_type="save",
            title="Export Pivot Table to Excel",
            file_filter="Excel Files (*.xlsx);;All Files (*)"
        )

        if not file_path:
            return # User cancelled

        # Update last directory - Handled by _get_file_path_dialog now
        # self.config.update_last_directory(os.path.dirname(file_path))

        try:
            # Call the export method on the pivot table widget
            if self.pivot_tab.export_to_excel(file_path):
                 self.status_bar.showMessage(f"Successfully exported pivot data to {os.path.basename(file_path)}", 5000)
                 QMessageBox.information(self, "Export Successful", f"Successfully exported pivot data to:\\n{file_path}")
            # Error handling should ideally be within pivot_tab.export_to_excel
            # else:
            #     QMessageBox.critical(self, "Export Error", "Failed to export pivot data to Excel.")
            #     self.status_bar.showMessage("Excel export failed", 5000)

        except Exception as e:
            error_msg = f"An error occurred during Excel export: {str(e)}"
            logger.error(error_msg, exc_info=True)
            QMessageBox.critical(self, "Export Error", error_msg)
            self.status_bar.showMessage("Excel export failed", 5000)

    def _export_rul(self):
        """Export rules to Altium RUL file"""
        # Use renamed variable
        if self.rules_manager_tab is None or not hasattr(self.rules_manager_tab, 'get_rule_manager'):
            QMessageBox.warning(self, "Export Error", "Rule Manager tab is not open or does not support exporting.")
            logger.warning("Attempted to export RUL when Rule Manager tab is not available.")
            return False # Indicate failure/not applicable

        # Use renamed variable
        rule_manager = self.rules_manager_tab.get_rule_manager()
        if not rule_manager or not rule_manager.rules:
            QMessageBox.warning(self, "Export Error", "No rules available in the Rule Manager to export.")
            logger.warning("Attempted to export RUL with no rules loaded in the manager.")
            return False # Indicate failure/nothing to save

        suggested_filename = "generated_rules.RUL"
        # You could potentially base the suggested name on an imported file if tracked

        file_path = self._get_file_path_dialog(
            dialog_type="save",
            title="Export Rules to Altium RUL File",
            file_filter="RUL Files (*.RUL);;All Files (*)"
        )

        if not file_path:
            logger.info("RUL export cancelled by user.")
            # Returning True here because the user explicitly cancelled, not an error state.
            # This prevents the closeEvent from aborting if the user cancels the save dialog.
            return True # User cancelled the dialog

        # Ensure the filename ends with .RUL (case-insensitive check)
        if not file_path.lower().endswith('.rul'):
            file_path += '.RUL'

        # Update last directory - Handled by _get_file_path_dialog now
        # self.config.update_last_directory(os.path.dirname(file_path))
        self.status_bar.showMessage(f"Exporting rules to {os.path.basename(file_path)}...", 3000)
        logger.info(f"Exporting rules to RUL file: {file_path}")

        try:
            rule_generator = RuleGenerator()
            # Assign the manager from the editor to the generator instance
            rule_generator.rule_manager = rule_manager
            # Generate the RUL content using the assigned manager
            rul_content = rule_generator.generate_rul_content()

            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(rul_content)

            self.status_bar.showMessage(f"Successfully exported to {os.path.basename(file_path)}", 5000)
            logger.info(f"Successfully exported rules to {file_path}")
            QMessageBox.information(self, "Export Successful", f"Rules successfully exported to:\\n{file_path}")

            # Mark the tab as saved - Needs implementation in RulesManagerWidget
            # if hasattr(self.rules_manager_tab, 'mark_saved'):
            #     self.rules_manager_tab.mark_saved()
            self._check_unsaved_changes() # Update window title

            return True # Indicate success

        except RuleGeneratorError as rge:
            error_msg = f"Error generating RUL content: {str(rge)}"
            logger.error(error_msg, exc_info=True)
            QMessageBox.critical(self, "Export Error", error_msg)
            self.status_bar.showMessage("Export failed", 5000)
            return False # Indicate failure
        except IOError as ioe:
            error_msg = f"Error writing RUL file '{os.path.basename(file_path)}': {str(ioe)}"
            logger.error(error_msg, exc_info=True)
            QMessageBox.critical(self, "Export Error", error_msg)
            self.status_bar.showMessage("Export failed", 5000)
            return False # Indicate failure
        except Exception as e:
            error_msg = f"An unexpected error occurred during RUL export: {str(e)}"
            logger.error(error_msg, exc_info=True)
            QMessageBox.critical(self, "Export Error", error_msg)
            self.status_bar.showMessage("Export failed", 5000)
            return False # Indicate failure

    def _show_preferences(self):
        """Show the preferences dialog."""
        # Pass the config manager and theme manager to the dialog
        dialog = PreferencesDialog(self.config, self.theme_manager, self)
        # Execute the dialog modally
        dialog.exec_()
        # Optionally, apply changes immediately if the dialog signals it
        # For example, if the theme changed, re-apply it:
        # self.theme_manager.apply_theme(self.config.get("theme", "light"))
        logger.info("Preferences dialog closed.")

    def _handle_generated_rules(self, generated_rules: List[BaseRule]):
        """Handles the rules generated by the PivotTableWidget.
        Loads the generated rules into the Rule Editor tab.

        Args:
            rules (list[BaseRule]): The list of generated rule objects.
        """
        if not generated_rules:
            logger.warning("Received empty list of generated rules.")
            return

        logger.info(f"Received {len(generated_rules)} generated rules from pivot table.")

        try:
            # Ensure the Rule Manager tab exists or create it
            self._show_rule_editor_tab()

            if self.rules_manager_tab:
                # Pass the generated rules to the Rule Manager tab using the new method
                self.rules_manager_tab.set_and_load_rules(generated_rules)
                # Switch to the Rule Manager tab
                self.tab_widget.setCurrentWidget(self.rules_manager_tab) # Corrected: self.tabs -> self.tab_widget
                logger.info("Loaded generated rules into Rule Manager tab and switched view.")
            else:
                # Error already logged in _show_rule_editor_tab if creation failed
                logger.error("Failed to show or access the Rule Manager tab after attempting creation.")
                # QMessageBox might have already been shown in _show_rule_editor_tab

        except Exception as e:
            logger.error(f"Error handling generated rules: {e}", exc_info=True)
            QMessageBox.critical(self, "Error", f"An unexpected error occurred while handling generated rules: {e}")

    def _show_rule_editor_tab(self):
        """Creates or shows the Rule Manager tab."""
        if self.rules_manager_tab is None:
            try:
                logger.info("Creating Rule Manager tab.")
                self.rules_manager_tab = RulesManagerWidget(self) # Pass self as parent
                # Connect signals
                self.rules_manager_tab.unsaved_changes_changed.connect(self._update_window_title)
                # Add other necessary signal connections here

                tab_index = self.tab_widget.addTab(self.rules_manager_tab, "Rule Manager") # Corrected: self.tabs -> self.tab_widget
                logger.info(f"Rule Manager tab created at index {tab_index}.")
                # Rules are now loaded externally via set_and_load_rules, not here.
            except Exception as e:
                logger.error(f"Failed to create Rule Manager tab: {e}", exc_info=True)
                QMessageBox.critical(self, "Error", f"Could not create the Rule Manager tab: {e}")
                self.rules_manager_tab = None # Ensure it's None if creation failed
                return # Stop here if creation failed

        # Switch to the tab if it exists (even if just created)
        if self.rules_manager_tab:
            self.tab_widget.setCurrentWidget(self.rules_manager_tab) # Corrected: self.tabs -> self.tab_widget
            logger.debug("Switched to Rule Manager tab.")
        else:
            logger.warning("Attempted to switch to Rule Manager tab, but it's None.")

    def _update_window_title(self, has_unsaved_changes: Optional[bool] = None):
        """Updates the window title to indicate unsaved changes."""
        base_title = "Altium Rule Generator"
        unsaved_marker = " [*]"

        # Determine if any tab has unsaved changes
        # If has_unsaved_changes is provided by signal, use it directly for that tab
        # Otherwise, check all relevant tabs
        # For now, let's assume only Rule Manager tracks this
        is_unsaved = False
        if self.rules_manager_tab and self.rules_manager_tab.has_unsaved_changes():
            is_unsaved = True
        # Add checks for other tabs if they implement has_unsaved_changes()
        # elif self.pivot_table_tab and self.pivot_table_tab.has_unsaved_changes():
        #     is_unsaved = True

        new_title = base_title + (unsaved_marker if is_unsaved else "")

        if self.windowTitle() != new_title:
            self.setWindowTitle(new_title)
            logger.debug(f"Window title updated: {new_title}")

    def closeEvent(self, event):
        """Handle the window close event."""
        # Check for unsaved changes before closing
        unsaved = False
        if self.rules_manager_tab and self.rules_manager_tab.has_unsaved_changes():
            unsaved = True
        # Add checks for other tabs

        if unsaved:
            reply = QMessageBox.question(self,
                                         "Unsaved Changes",
                                         "There are unsaved changes. Do you want to save before exiting?",
                                         QMessageBox.Save | QMessageBox.Discard | QMessageBox.Cancel,
                                         QMessageBox.Cancel)

            if reply == QMessageBox.Save:
                # Attempt to save rules (assuming only rule manager needs saving)
                if self.rules_manager_tab:
                    # Ideally, trigger the save action of the relevant widget
                    # For now, let's assume _save_rules shows a dialog
                    self.rules_manager_tab._save_rules() # This might show its own dialog
                    # Re-check if changes were actually saved
                    if self.rules_manager_tab.has_unsaved_changes():
                        event.ignore() # Ignore close event if save failed or was cancelled
                        return
                    else:
                        event.accept() # Accept close event if save was successful
                else:
                    event.accept() # No rule manager tab to save
            elif reply == QMessageBox.Discard:
                event.accept() # Discard changes and close
            else: # Cancel
                event.ignore() # Ignore the close event
                return
        else:
            event.accept() # No unsaved changes, close normally

        # Save window state if closing is accepted
        if event.isAccepted():
            self._save_geometry() # Renamed from _save_settings
            logger.info("Application closing.")
            super().closeEvent(event)

    # Add placeholder methods if needed by other parts, assuming they exist in widgets:
    # def _on_data_saved(self): ... # Might be called after successful export

    def _export_rul(self):
        """Export the rules from the Rule Editor tab to an Altium RUL file."""
        # Use renamed variable
        if self.rules_manager_tab is None or not hasattr(self.rules_manager_tab, 'get_rule_manager'):
            QMessageBox.warning(self, "Export Error", "Rule Manager tab is not open or does not support exporting.")
            logger.warning("Attempted to export RUL when Rule Manager tab is not available.")
            return False # Indicate failure/not applicable

        # Use renamed variable
        rule_manager = self.rules_manager_tab.get_rule_manager()
        if not rule_manager or not rule_manager.rules:
            QMessageBox.warning(self, "Export Error", "No rules available in the Rule Manager to export.")
            logger.warning("Attempted to export RUL with no rules loaded in the manager.")
            return False # Indicate failure/nothing to save

        suggested_filename = "generated_rules.RUL"
        # You could potentially base the suggested name on an imported file if tracked

        file_path = self._get_file_path_dialog(
            dialog_type="save",
            title="Export Rules to Altium RUL File",
            file_filter="RUL Files (*.RUL);;All Files (*)"
        )

        if not file_path:
            logger.info("RUL export cancelled by user.")
            # Returning True here because the user explicitly cancelled, not an error state.
            # This prevents the closeEvent from aborting if the user cancels the save dialog.
            return True # User cancelled the dialog

        # Ensure the filename ends with .RUL (case-insensitive check)
        if not file_path.lower().endswith('.rul'):
            file_path += '.RUL'

        # Update last directory - Handled by _get_file_path_dialog now
        # self.config.update_last_directory(os.path.dirname(file_path))
        self.status_bar.showMessage(f"Exporting rules to {os.path.basename(file_path)}...", 3000)
        logger.info(f"Exporting rules to RUL file: {file_path}")

        try:
            rule_generator = RuleGenerator()
            # Assign the manager from the editor to the generator instance
            rule_generator.rule_manager = rule_manager
            # Generate the RUL content using the assigned manager
            rul_content = rule_generator.generate_rul_content()

            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(rul_content)

            self.status_bar.showMessage(f"Successfully exported to {os.path.basename(file_path)}", 5000)
            logger.info(f"Successfully exported rules to {file_path}")
            QMessageBox.information(self, "Export Successful", f"Rules successfully exported to:\\n{file_path}")

            # Mark the tab as saved - Needs implementation in RulesManagerWidget
            # if hasattr(self.rules_manager_tab, 'mark_saved'):
            #     self.rules_manager_tab.mark_saved()
            self._check_unsaved_changes() # Update window title

            return True # Indicate success

        except RuleGeneratorError as rge:
            error_msg = f"Error generating RUL content: {str(rge)}"
            logger.error(error_msg, exc_info=True)
            QMessageBox.critical(self, "Export Error", error_msg)
            self.status_bar.showMessage("Export failed", 5000)
            return False # Indicate failure
        except IOError as ioe:
            error_msg = f"Error writing RUL file '{os.path.basename(file_path)}': {str(ioe)}"
            logger.error(error_msg, exc_info=True)
            QMessageBox.critical(self, "Export Error", error_msg)
            self.status_bar.showMessage("Export failed", 5000)
            return False # Indicate failure
        except Exception as e:
            error_msg = f"An unexpected error occurred during RUL export: {str(e)}"
            logger.error(error_msg, exc_info=True)
            QMessageBox.critical(self, "Export Error", error_msg)
            self.status_bar.showMessage("Export failed", 5000)
            return False # Indicate failure

    def _on_data_changed(self, *args, **kwargs):
        """Slot to handle data changes from tabs (e.g., pivot table, rule editor)."""
        logger.debug("Data changed signal received.")
        self._check_unsaved_changes()

    def _on_rule_pivot_updated(self, pivot_data: ExcelPivotData):
        """Slot to handle pivot data updates from the rule editor."""
        if self.pivot_tab:
            self.pivot_tab.set_pivot_data(pivot_data)
            logger.info("Pivot table updated with data from rule editor.")
        else:
            logger.warning("Received pivot data update from rule editor, but pivot tab does not exist.")

    def _show_rule_editor_tab(self):
        """Creates or shows the Rule Manager tab."""
        if self.rules_manager_tab is None:
            try:
                logger.info("Creating Rule Manager tab.")
                self.rules_manager_tab = RulesManagerWidget(self) # Pass self as parent
                # Connect signals
                self.rules_manager_tab.unsaved_changes_changed.connect(self._update_window_title)
                # Add other necessary signal connections here

                tab_index = self.tab_widget.addTab(self.rules_manager_tab, "Rule Manager") # Corrected: self.tabs -> self.tab_widget
                logger.info(f"Rule Manager tab created at index {tab_index}.")
                # Rules are now loaded externally via set_and_load_rules, not here.
            except Exception as e:
                logger.error(f"Failed to create Rule Manager tab: {e}", exc_info=True)
                QMessageBox.critical(self, "Error", f"Could not create the Rule Manager tab: {e}")
                self.rules_manager_tab = None # Ensure it's None if creation failed
                return # Stop here if creation failed

        # Switch to the tab if it exists (even if just created)
        if self.rules_manager_tab:
            self.tab_widget.setCurrentWidget(self.rules_manager_tab) # Corrected: self.tabs -> self.tab_widget
            logger.debug("Switched to Rule Manager tab.")
        else:
            logger.warning("Attempted to switch to Rule Manager tab, but it's None.")

    def _check_unsaved_changes(self):
        """Checks all open tabs for unsaved changes and updates the window title."""
        has_changes = False
        # Iterate through all widgets in the tab widget
        for i in range(self.tab_widget.count()):
            widget = self.tab_widget.widget(i)
            if hasattr(widget, 'has_unsaved_changes') and widget.has_unsaved_changes():
                has_changes = True
                break # Found unsaved changes, no need to check further
        
        # Update window title if unsaved changes exist
        base_title = "Altium Rule Generator"
        if has_changes:
            self.setWindowTitle(f"{base_title}*")
        else:
            self.setWindowTitle(base_title)
        
        # Emit signal if needed (though direct title update might be sufficient)
        # self.unsaved_changes_changed.emit(has_changes)
        logger.debug(f"Unsaved changes status checked: {has_changes}")

    def closeEvent(self, event):
        """Handle the window close event."""
        # Check for unsaved changes before closing
        unsaved = False
        if self.rules_manager_tab and self.rules_manager_tab.has_unsaved_changes():
            unsaved = True
        # Add checks for other tabs

        if unsaved:
            reply = QMessageBox.question(self,
                                         "Unsaved Changes",
                                         "There are unsaved changes. Do you want to save before exiting?",
                                         QMessageBox.Save | QMessageBox.Discard | QMessageBox.Cancel,
                                         QMessageBox.Cancel)

            if reply == QMessageBox.Save:
                # Attempt to save rules (assuming only rule manager needs saving)
                if self.rules_manager_tab:
                    # Ideally, trigger the save action of the relevant widget
                    # For now, let's assume _save_rules shows a dialog
                    self.rules_manager_tab._save_rules() # This might show its own dialog
                    # Re-check if changes were actually saved
                    if self.rules_manager_tab.has_unsaved_changes():
                        event.ignore() # Ignore close event if save failed or was cancelled
                        return
                    else:
                        event.accept() # Accept close event if save was successful
                else:
                    event.accept() # No rule manager tab to save
            elif reply == QMessageBox.Discard:
                event.accept() # Discard changes and close
            else: # Cancel
                event.ignore() # Ignore the close event
                return
        else:
            event.accept() # No unsaved changes, close normally

        # Save window state if closing is accepted
        if event.isAccepted():
            self._save_geometry() # Renamed from _save_settings
            logger.info("Application closing.")
            super().closeEvent(event)