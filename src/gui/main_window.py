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
from PyQt5.QtWidgets import (QMainWindow, QTabWidget, QAction, QFileDialog,
                             QMenu, QToolBar, QStatusBar, QMessageBox,
                             QDockWidget, QVBoxLayout, QHBoxLayout, QWidget,
                             QShortcut, QApplication, QInputDialog, QActionGroup)
from PyQt5.QtGui import QFont, QIcon, QKeySequence
from PyQt5.QtCore import Qt, QSize, pyqtSignal

from models.excel_data import ExcelPivotData
from models.rule_model import RuleType, UnitType, BaseRule
from models.rule_model import RuleManager

from gui.preferences_dialog import PreferencesDialog # Add import for PreferencesDialog
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
        self.rule_editor_tab = None

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
                padding: 5px 20px 5px 20px; /* Top, Right, Bottom, Left padding */
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
            if hasattr(widget, 'has_unsaved_changes') and widget.has_unsaved_changes():
                reply = QMessageBox.question(self, f'Unsaved Changes in {tab_name}',
                                             f'The tab "{tab_name}" has unsaved changes. Do you want to save them before closing?',
                                             QMessageBox.Save | QMessageBox.Discard | QMessageBox.Cancel,
                                             QMessageBox.Cancel)

                if reply == QMessageBox.Save:
                    # Attempt to save based on tab type
                    save_successful = False
                    if widget == self.rule_editor_tab:
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
                except TypeError: # Signal already disconnected
                    pass
                self.pivot_tab = None
                logger.info("Pivot Table tab closed.")
            elif widget == self.rule_editor_tab:
                try:
                    self.rule_editor_tab.rules_changed.disconnect(self._on_data_changed)
                    self.rule_editor_tab.pivot_data_updated.disconnect(self._on_rule_pivot_updated)
                except TypeError:
                    pass
                self.rule_editor_tab = None
                logger.info("Rule Editor tab closed.")
            
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
        
        # Add action to show Rule Editor
        self._add_action(self.view_menu, "Show &Rule Editor", "settings.png", None,
                         "Open the Rule Editor tab", self._show_rule_editor_tab)
        
        self.view_menu.addSeparator()
        
        # Theme submenu
        theme_menu = QMenu("&Theme", self)
        self.view_menu.addMenu(theme_menu)
        
        # Theme actions
        theme_group = QActionGroup(self)  # Ensure only one theme can be selected at a time
        theme_group.setExclusive(True)

        # Iterate directly over the list of theme IDs (names)
        for theme_id in self.theme_manager.get_available_themes():
            # Use theme_id for display name (e.g., capitalize) or keep as is
            theme_display_name = theme_id.capitalize() 
            theme_action = QAction(theme_display_name, self)
            theme_action.setCheckable(True)
            theme_action.setChecked(theme_id == self.theme_manager.get_current_theme())
            theme_action.setData(theme_id) # Store the theme_id
            # Connect using the theme_id captured by the lambda
            theme_action.triggered.connect(lambda checked, theme=theme_id: self.theme_manager.set_theme(theme))
            theme_group.addAction(theme_action)
            theme_menu.addAction(theme_action)

    def _show_rule_editor_tab(self):
        """Creates and shows the Rule Editor tab if it doesn't exist."""
        if self.rule_editor_tab is None:
            try:
                from gui.rule_editor_widget import RulesManagerWidget
                self.rule_editor_tab = RulesManagerWidget()
                # Connect signals only when creating the tab
                self.rule_editor_tab.rules_changed.connect(self._on_data_changed)
                self.rule_editor_tab.pivot_data_updated.connect(self._on_rule_pivot_updated)
                # Connect to pivot data if pivot tab exists
                if self.pivot_tab and hasattr(self.pivot_tab, 'get_pivot_data'):
                    pivot_data = self.pivot_tab.get_pivot_data()
                    if pivot_data:
                        self.rule_editor_tab.update_pivot_data(pivot_data)
                
                index = self.tab_widget.addTab(self.rule_editor_tab, "Rule Editor")
                self.tab_widget.setCurrentIndex(index)
                logger.info("Rule Editor tab created and shown.")
            except Exception as e:
                logger.error(f"Failed to create Rule Editor tab: {e}", exc_info=True)
                QMessageBox.critical(self, "Error", f"Could not create the Rule Editor tab: {e}")
                self.rule_editor_tab = None # Ensure it's None if creation failed
        else:
            # Find the index of the existing rule editor tab and switch to it
            for i in range(self.tab_widget.count()):
                if self.tab_widget.widget(i) == self.rule_editor_tab:
                    self.tab_widget.setCurrentIndex(i)
                    break
            logger.info("Switched to existing Rule Editor tab.")

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
                if self.rule_editor_tab and hasattr(self.pivot_tab, 'get_pivot_data'):
                    # get_pivot_data should return the ExcelPivotData object
                    current_pivot_data = self.pivot_tab.get_pivot_data() 
                    if current_pivot_data:
                        self.rule_editor_tab.update_pivot_data(current_pivot_data)
                        logger.info("Updated Rule Editor with new pivot data.")
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
        from models.rule_model import RuleManager # Import RuleManager
        rule_generator = RuleGenerator()
        
        try:
            # Parse RUL file - assume parse_rul_file populates rule_generator.rule_manager
            success = rule_generator.parse_rul_file(file_path)

            if not success:
                # Check if rule_manager exists and has rules even if success is False (might parse some)
                if not rule_generator.rule_manager or not rule_generator.rule_manager.rules:
                     QMessageBox.warning(
                         self,
                         "Import Warning",
                         "The RUL file could not be parsed or no valid rules were found."
                     )
                     logger.warning(f"Parsing RUL file {file_path} failed or yielded no rules.")
                     return
                else:
                     logger.warning(f"Parsing RUL file {file_path} reported failure, but some rules were extracted.")

            # Get the RuleManager instance containing the parsed rules
            rule_manager = rule_generator.rule_manager # This is the manager with parsed rules

            if not rule_manager or not rule_manager.rules:
                QMessageBox.warning(
                    self,
                    "Import Warning",
                    "No rules were found or extracted from the RUL file."
                )
                logger.warning(f"No rules found in RuleManager after parsing {file_path}.")
                return

            # --- Create or Update Rule Editor Tab ---
            if self.rule_editor_tab is None:
                # Use the existing _show_rule_editor_tab method to create/show
                self._show_rule_editor_tab()
                # Check if creation was successful (it might fail)
                if self.rule_editor_tab is None:
                     logger.error("Failed to create Rule Editor tab during RUL import.")
                     # QMessageBox shown in _show_rule_editor_tab
                     return
            else:
                 # Find the index of the existing rule editor tab and switch to it
                 for i in range(self.tab_widget.count()):
                     if self.tab_widget.widget(i) == self.rule_editor_tab:
                         self.tab_widget.setCurrentIndex(i)
                         break
                 logger.info("Switched to existing Rule Editor tab for RUL import.")

            # Load the parsed rules into the Rule Editor tab
            try:
                # Load the manager containing the parsed rules directly
                self.rule_editor_tab.load_rules(rule_manager) # Load the manager with parsed rules
                logger.info(f"Loaded {len(rule_manager.rules)} parsed rules into the Rule Editor tab.")

                # Optionally, switch focus to the Rule Editor tab
                for i in range(self.tab_widget.count()):
                    if self.tab_widget.widget(i) == self.rule_editor_tab:
                        self.tab_widget.setCurrentIndex(i)
                        break

                # Mark the rule editor as having unsaved changes
                if hasattr(self.rule_editor_tab, 'mark_unsaved_changes'):
                    self.rule_editor_tab.mark_unsaved_changes()
                self._check_unsaved_changes() # Update window title

            except Exception as e:
                error_msg = f"Error loading parsed rules into rule editor: {str(e)}"
                logger.error(error_msg, exc_info=True)
                QMessageBox.critical(self, "Load Error", error_msg)
        except RuleGeneratorError as rge: # Specific exception for rule generation issues
            error_msg = f"Error parsing RUL file: {str(rge)}"
            logger.error(error_msg, exc_info=True) # Include traceback for parsing errors
            QMessageBox.critical(self, "RUL Parse Error", error_msg)
            self.status_bar.showMessage("RUL import failed", 5000)
        except Exception as e: # General exception handler
            error_msg = f"An unexpected error occurred during RUL import: {str(e)}"
            logger.error(error_msg, exc_info=True)
            QMessageBox.critical(self, "Import Error", error_msg)
            self.status_bar.showMessage("Import failed", 5000)

    def _export_excel(self):
        """Export the current pivot table data to an Excel file."""
        if self.pivot_tab is None or self.pivot_tab.model is None:
            QMessageBox.warning(self, "Export Error", "No pivot table data available to export.")
            logger.warning("Attempted to export Excel with no pivot table tab or model.")
            return False

        try:
            # Get the ExcelPivotData object from the model
            pivot_data_obj = self.pivot_tab.model.get_updated_pivot_data()
            if pivot_data_obj is None or pivot_data_obj.pivot_df is None:
                 QMessageBox.warning(self, "Export Error", "Could not retrieve pivot data structure.")
                 logger.warning("get_updated_pivot_data() returned None or pivot_df was None.")
                 return False

            df = pivot_data_obj.pivot_df # Get the DataFrame
            if df.empty:
                QMessageBox.warning(self, "Export Error", "Pivot table data is empty.")
                logger.warning("Attempted to export empty pivot table data to Excel.")
                return False
        except Exception as e:
            error_msg = f"Error retrieving data from pivot table model: {str(e)}"
            logger.error(error_msg, exc_info=True)
            QMessageBox.critical(self, "Export Error", error_msg)
            return False

        # Suggest a filename based on the original import or a default
        suggested_filename = "pivot_export.xlsx"
        # You could potentially store the original filename when importing
        # if hasattr(self.pivot_tab, 'source_filename') and self.pivot_tab.source_filename:
        #     base, _ = os.path.splitext(os.path.basename(self.pivot_tab.source_filename))
        #     suggested_filename = f"{base}_pivot_export.xlsx"

        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Export Pivot Table to Excel",
            os.path.join(self.config.get("last_directory", ""), suggested_filename),
            "Excel Files (*.xlsx);;All Files (*)"
        )

        if not file_path:
            logger.info("Excel export cancelled by user.")
            return False # User cancelled

        # Ensure the filename ends with .xlsx
        if not file_path.lower().endswith('.xlsx'):
            file_path += '.xlsx'

        self.config.update_last_directory(os.path.dirname(file_path))
        self.status_bar.showMessage(f"Exporting data to {os.path.basename(file_path)}...", 3000)
        logger.info(f"Exporting pivot table data to Excel: {file_path}")

        try:
            # Use pandas to export the DataFrame
            # Make sure 'openpyxl' is installed (add to requirements.txt if needed)
            df.to_excel(file_path, index=False) # index=False is common for exports like this
            self.status_bar.showMessage(f"Successfully exported to {os.path.basename(file_path)}", 5000)
            logger.info(f"Successfully exported pivot data to {file_path}")
            QMessageBox.information(self, "Export Successful", f"Data successfully exported to:\\n{file_path}")
            return True
        except Exception as e:
            error_msg = f"Error exporting data to Excel file '{os.path.basename(file_path)}': {str(e)}"
            logger.error(error_msg, exc_info=True)
            QMessageBox.critical(self, "Export Error", error_msg)
            self.status_bar.showMessage("Export failed", 5000)
            return False

    def _handle_generated_rules(self, rules: list):
        """
        Handles the rules generated by the PivotTableWidget.
        Loads the generated rules into the Rule Editor tab.

        Args:
            rules (list[BaseRule]): The list of generated rule objects.
        """
        if not rules:
            logger.info("No rules were generated by the pivot table.")
            QMessageBox.information(self, "No Rules Generated", "No rules were generated based on the current pivot table configuration.")
            return

        logger.info(f"Received {len(rules)} generated rules from pivot table.")

        # Ensure the rule editor tab exists
        self._show_rule_editor_tab()

        if self.rule_editor_tab:
            # Ask the user if they want to replace or append rules
            # reply = QMessageBox.question(self, "Add Generated Rules",
            #                              "Do you want to replace existing rules in the Rule Editor or append the newly generated rules?",
            #                              QMessageBox.Replace | QMessageBox.Append | QMessageBox.Cancel,
            #                              QMessageBox.Append) # Default to Append

            msgBox = QMessageBox(self)
            msgBox.setIcon(QMessageBox.Question)
            msgBox.setWindowTitle("Add Generated Rules")
            msgBox.setText("How do you want to add the generated rules to the Rule Editor?")
            replaceButton = msgBox.addButton("Replace", QMessageBox.ActionRole)
            appendButton = msgBox.addButton("Append", QMessageBox.ActionRole)
            cancelButton = msgBox.addButton(QMessageBox.Cancel)
            msgBox.setDefaultButton(appendButton) # Set Append as default

            msgBox.exec()
            clicked_button = msgBox.clickedButton()

            if clicked_button == replaceButton:
                # Call set_rules on the rule_model instead of table_model
                self.rule_editor_tab.rule_model.set_rules(rules) 
                logger.info("Replaced rules in Rule Editor with generated rules.")
            elif clicked_button == appendButton:
                self.rule_editor_tab.add_rules(rules)
                logger.info("Appended generated rules to Rule Editor.")
            elif clicked_button == cancelButton: # Check if the cancel button was clicked
                logger.info("User cancelled adding generated rules.")
                return
            else: # Should not happen, but handle just in case
                logger.warning("Unexpected button clicked in add rules dialog.")
                return
            
            # Switch to the rule editor tab
            self._show_rule_editor_tab() # This will switch if it already exists
        else:
            logger.error("Failed to show or access the Rule Editor tab to add generated rules.")
            QMessageBox.critical(self, "Error", "Could not access the Rule Editor tab to add the generated rules.")

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
        """Handle window close event"""
        if self._check_unsaved_changes():
            reply = QMessageBox.question(self, 'Unsaved Changes',
                                         "You have unsaved changes. Do you want to save before exiting?",
                                         QMessageBox.Save | QMessageBox.Discard | QMessageBox.Cancel,
                                         QMessageBox.Cancel) # Default to Cancel

            if reply == QMessageBox.Save:
                # Attempt to save ALL tabs with changes? Or just the current one?
                # Let's try saving only the Rule Editor if it has changes, as it's the primary output.
                save_successful = True # Assume success unless a save fails
                if self.rule_editor_tab and hasattr(self.rule_editor_tab, 'has_unsaved_changes') and self.rule_editor_tab.has_unsaved_changes():
                    save_successful = self._export_rul() # export_rul returns True on success/cancel, False on error
                
                # Add logic here if other tabs need saving (e.g., Pivot Table -> Excel)
                # elif self.pivot_tab and hasattr(self.pivot_tab, 'has_unsaved_changes') and self.pivot_tab.has_unsaved_changes():
                #    save_successful = self._export_excel()

                if not save_successful:
                    event.ignore() # Prevent closing if save failed
                    return
                # If save was successful or not needed for certain tabs, proceed to accept
                event.accept()

            elif reply == QMessageBox.Discard:
                event.accept()
            else: # Cancel
                event.ignore()
                return
        else:
            event.accept()
        
        # Save geometry before closing if accepted
        if event.isAccepted():
            self._save_geometry()
            logger.info("Application closing.")

    # Add placeholder methods if needed by other parts, assuming they exist in widgets:
    # def _on_data_saved(self): ... # Might be called after successful export

    def _export_rul(self):
        """Export the rules from the Rule Editor tab to an Altium RUL file."""
        if self.rule_editor_tab is None or not hasattr(self.rule_editor_tab, 'get_rule_manager'):
            QMessageBox.warning(self, "Export Error", "Rule Editor tab is not open or does not support exporting.")
            logger.warning("Attempted to export RUL when Rule Editor tab is not available.")
            return False # Indicate failure/not applicable

        rule_manager = self.rule_editor_tab.get_rule_manager()
        if not rule_manager or not rule_manager.rules:
            QMessageBox.warning(self, "Export Error", "No rules available in the Rule Editor to export.")
            logger.warning("Attempted to export RUL with no rules loaded in the editor.")
            return False # Indicate failure/nothing to save

        suggested_filename = "generated_rules.RUL"
        # You could potentially base the suggested name on an imported file if tracked

        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Export Rules to Altium RUL File",
            os.path.join(self.config.get("last_directory", ""), suggested_filename),
            "Altium Rules Files (*.RUL);;All Files (*)"
        )

        if not file_path:
            logger.info("RUL export cancelled by user.")
            # Returning True here because the user explicitly cancelled, not an error state.
            # This prevents the closeEvent from aborting if the user cancels the save dialog.
            return True # User cancelled the dialog

        # Ensure the filename ends with .RUL (case-insensitive check)
        if not file_path.lower().endswith('.rul'):
            file_path += '.RUL'

        self.config.update_last_directory(os.path.dirname(file_path))
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

            # Mark the tab as saved
            if hasattr(self.rule_editor_tab, 'mark_saved'):
                self.rule_editor_tab.mark_saved()
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
        """Show the application preferences dialog."""
        logging.info("Preferences action triggered.")
        try:
            # Pass the config manager and theme manager to the dialog
            pref_dialog = PreferencesDialog(self.config, self.theme_manager, self)
            pref_dialog.exec_() # Show the dialog modally
            # Changes (like theme) are applied by the dialog itself via the theme_manager
            logger.info("Preferences dialog closed.")
        except Exception as e:
            logger.error(f"Failed to open preferences dialog: {e}", exc_info=True)
            QMessageBox.critical(self, "Error", f"Could not open preferences dialog: {e}")


    def _show_about(self):
        """Show the About dialog box."""
        logging.info("About action triggered.")
        app_version = "Unknown" # Default version
        try:
            # Attempt to read version from setup.py
            setup_path = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), 'setup.py')
            if os.path.exists(setup_path):
                with open(setup_path, 'r', encoding='utf-8') as f:
                    setup_content = f.read()
                    # Use regex to find version = '...' or version = "..."
                    match = re.search(r"^version\s*=\s*['\"]([^'\"]*)['\"]", setup_content, re.MULTILINE)
                    if match:
                        app_version = match.group(1)
                    else:
                        logger.warning("Could not find version pattern in setup.py")
            else:
                logger.warning(f"setup.py not found at expected path: {setup_path}")
        except Exception as e:
            logger.warning(f"Could not read version info from setup.py: {e}")

        QMessageBox.about(
            self,
            "About Altium Rule Generator",
            f"<b>Altium XCEL to Query Builder</b>\\n\\n"
            f"Version: {app_version}\\n\\n"
            "This application helps generate Altium Designer rule queries "
            "from structured data, typically imported from Excel files.\\n\\n"
            "Provide feedback or report issues at: [Your GitHub/Contact Link Here]\\n\\n" # Added feedback link placeholder
            f"(c) {datetime.now().year} Your Name/Company" # Replace with actual copyright, use current year
        )

    def _check_unsaved_changes(self):
        """Checks if any open tab has unsaved changes and updates window title."""
        has_changes = False
        # Iterate through widgets in the tab widget
        for i in range(self.tab_widget.count()):
            widget = self.tab_widget.widget(i)
            if widget and hasattr(widget, 'has_unsaved_changes') and widget.has_unsaved_changes():
                has_changes = True
                break # Found one, no need to check further
        
        # Update window title
        title = "Altium Rule Generator"
        if has_changes:
            title += " *"
        self.setWindowTitle(title)
        logger.debug(f"Overall unsaved changes status: {has_changes}")
        return has_changes

    def closeEvent(self, event):
        """Handle window close event"""
        if self._check_unsaved_changes():
            reply = QMessageBox.question(self, 'Unsaved Changes',
                                         "You have unsaved changes. Do you want to save before exiting?",
                                         QMessageBox.Save | QMessageBox.Discard | QMessageBox.Cancel,
                                         QMessageBox.Cancel) # Default to Cancel

            if reply == QMessageBox.Save:
                # Attempt to save ALL tabs with changes? Or just the current one?
                # Let's try saving only the Rule Editor if it has changes, as it's the primary output.
                save_successful = True # Assume success unless a save fails
                if self.rule_editor_tab and hasattr(self.rule_editor_tab, 'has_unsaved_changes') and self.rule_editor_tab.has_unsaved_changes():
                    save_successful = self._export_rul() # export_rul returns True on success/cancel, False on error
                
                # Add logic here if other tabs need saving (e.g., Pivot Table -> Excel)
                # elif self.pivot_tab and hasattr(self.pivot_tab, 'has_unsaved_changes') and self.pivot_tab.has_unsaved_changes():
                #    save_successful = self._export_excel()

                if not save_successful:
                    event.ignore() # Prevent closing if save failed
                    return
                # If save was successful or not needed for certain tabs, proceed to accept
                event.accept()

            elif reply == QMessageBox.Discard:
                event.accept()
            else: # Cancel
                event.ignore()
                return
        else:
            event.accept()
        
        # Save geometry before closing if accepted
        if event.isAccepted():
            self._save_geometry()
            logger.info("Application closing.")

    # Add placeholder methods if needed by other parts, assuming they exist in widgets:
    # def _on_data_saved(self): ... # Might be called after successful export

    def _export_rul(self):
        """Export the rules from the Rule Editor tab to an Altium RUL file."""
        if self.rule_editor_tab is None or not hasattr(self.rule_editor_tab, 'get_rule_manager'):
            QMessageBox.warning(self, "Export Error", "Rule Editor tab is not open or does not support exporting.")
            logger.warning("Attempted to export RUL when Rule Editor tab is not available.")
            return False # Indicate failure/not applicable

        rule_manager = self.rule_editor_tab.get_rule_manager()
        if not rule_manager or not rule_manager.rules:
            QMessageBox.warning(self, "Export Error", "No rules available in the Rule Editor to export.")
            logger.warning("Attempted to export RUL with no rules loaded in the editor.")
            return False # Indicate failure/nothing to save

        suggested_filename = "generated_rules.RUL"
        # You could potentially base the suggested name on an imported file if tracked

        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Export Rules to Altium RUL File",
            os.path.join(self.config.get("last_directory", ""), suggested_filename),
            "Altium Rules Files (*.RUL);;All Files (*)"
        )

        if not file_path:
            logger.info("RUL export cancelled by user.")
            # Returning True here because the user explicitly cancelled, not an error state.
            # This prevents the closeEvent from aborting if the user cancels the save dialog.
            return True # User cancelled the dialog

        # Ensure the filename ends with .RUL (case-insensitive check)
        if not file_path.lower().endswith('.rul'):
            file_path += '.RUL'

        self.config.update_last_directory(os.path.dirname(file_path))
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

            # Mark the tab as saved
            if hasattr(self.rule_editor_tab, 'mark_saved'):
                self.rule_editor_tab.mark_saved()
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
