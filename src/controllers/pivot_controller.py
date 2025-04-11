#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Pivot Controller
==============

Controller for managing pivot table data and interactions.
"""

import os
import logging
from typing import Dict, List, Optional, Union, Tuple

from PyQt5.QtWidgets import QWidget, QFileDialog, QMessageBox, QInputDialog

from gui.pivot_table_widget import PivotTableWidget
from gui.excel_preview_dialog import ExcelPreviewDialog
from models.excel_data import ExcelPivotData
from models.rule_model import UnitType, RuleType, BaseRule
from services.excel_importer import ExcelImporter, ExcelImportError
from services.excel_formatter import ExcelFormatter

logger = logging.getLogger(__name__)

class PivotController:
    """Controller for pivot table operations"""
      def __init__(self, pivot_widget: PivotTableWidget, config_manager):
        """Initialize pivot controller"""
        self.pivot_widget = pivot_widget
        self.config = config_manager
        self.excel_importer = ExcelImporter()
        self.excel_formatter = ExcelFormatter()
        self.current_file_path = None
        self.current_sheet_name = None
        
        # Connect signals
        self.pivot_widget.data_changed.connect(self._on_data_changed)
        
        logger.info("Pivot controller initialized")
    
    def import_excel_file(self, parent_widget: QWidget = None) -> bool:
        """Import Excel file and display in pivot widget"""
        try:
            # Get file path
            file_path, _ = QFileDialog.getOpenFileName(
                parent_widget,
                "Import Excel File",
                self.config.get("last_directory", ""),
                "Excel Files (*.xlsx *.xls);;All Files (*)"
            )
            
            if not file_path:
                # User cancelled
                return False
            
            # Update last directory
            self.config.update_last_directory(os.path.dirname(file_path))
              # Get sheet names
            sheet_names = self.excel_importer.get_sheet_names(file_path)
            
            # If multiple sheets, ask user which one to import
            sheet_name = None
            if len(sheet_names) > 1:
                sheet_name, ok = QInputDialog.getItem(self, "Select Sheet", "Sheet Name:", sheet_names, 0, False)
                
                if not ok:
                    # User cancelled
                    return False
            else:
                sheet_name = sheet_names[0]
              # Import raw Excel data for preview
            raw_df = self.excel_importer.import_file(file_path, sheet_name)
            
            # Show preview dialog
            preview_dialog = ExcelPreviewDialog(raw_df, sheet_name, parent_widget)
            
            # If user cancels preview, abort import
            if not preview_dialog.exec_():
                return False
            
            # Get processed dataframe and options from preview
            processed_df = preview_dialog.get_processed_dataframe()
            import_options = preview_dialog.get_import_options()
            
            # Import as pivot data with the processed dataframe
            rule_type = RuleType.CLEARANCE  # Default to clearance rules
            
            # Create pivot data
            pivot_data = ExcelPivotData(rule_type)
            
            # Set unit from import options
            unit = import_options["unit"]
            
            # Process dataframe if needed for pivot structure
            if import_options["use_first_column_as_index"]:
                # First column is net class names
                index_col = processed_df.columns[0]
                processed_df.set_index(index_col, inplace=True)
                processed_df.reset_index(inplace=True)  # Reset but keep the column
            
            # Load dataframe into pivot data
            success = pivot_data.load_dataframe(processed_df, unit)
            
            if not success:
                raise ExcelImportError("Failed to load DataFrame into pivot data structure")
            
            # Set in widget
            self.pivot_widget.set_pivot_data(pivot_data)
            
            # Store current file info
            self.current_file_path = file_path
            self.current_sheet_name = sheet_name
            self.config.add_recent_file(file_path)
            
            logger.info(f"Imported Excel file: {file_path}, sheet: {sheet_name}")
            return True
        
        except ExcelImportError as e:
            error_msg = f"Error importing Excel file: {str(e)}"
            logger.error(error_msg)
            
            if parent_widget:
                QMessageBox.critical(parent_widget, "Import Error", error_msg)
            
            return False
        
        except Exception as e:
            error_msg = f"Unexpected error importing Excel file: {str(e)}"
            logger.error(error_msg)
            
            if parent_widget:
                QMessageBox.critical(parent_widget, "Import Error", error_msg)
            
            return False
    
    def export_excel_file(self, parent_widget: QWidget = None) -> bool:
        """Export pivot data to Excel file"""
        try:
            # Get pivot data
            pivot_data = self.pivot_widget.get_pivot_data()
            
            if not pivot_data:
                error_msg = "No pivot data to export"
                logger.error(error_msg)
                
                if parent_widget:
                    QMessageBox.warning(parent_widget, "Export Error", error_msg)
                
                return False
            
            # Get file path
            file_path, _ = QFileDialog.getSaveFileName(
                parent_widget,
                "Export to Excel",
                self.config.get("last_directory", ""),
                "Excel Files (*.xlsx);;All Files (*)"
            )
            
            if not file_path:
                # User cancelled
                return False
            
            # Ensure file has .xlsx extension
            if not file_path.lower().endswith('.xlsx'):
                file_path += '.xlsx'
            
            # Update last directory
            self.config.update_last_directory(os.path.dirname(file_path))
              # Get DataFrame
            df = pivot_data.to_dataframe()
            
            # Ask user for formatting preset
            preset_options = self.excel_formatter.get_available_presets()
            preset_names = list(preset_options.values())
            preset_ids = list(preset_options.keys())
            
            selected_preset_name, ok = QInputDialog.getItem(
                parent_widget,
                "Select Excel Formatting",
                "Choose a formatting style for the Excel export:",
                preset_names,
                0,  # Default to first option
                False  # Non-editable
            )
            
            if ok and selected_preset_name:
                # Find the preset ID by name
                preset_index = preset_names.index(selected_preset_name)
                preset_id = preset_ids[preset_index]
                self.excel_formatter.set_preset(preset_id)
                
                # Get sheet title
                sheet_title = f"Clearance Rules - {pivot_data.unit.value}"
                
                # Apply formatting and export
                success = self.excel_formatter.format_workbook(
                    file_path=file_path,
                    df=df,
                    sheet_name="Clearance",
                    unit=pivot_data.unit,
                    title=sheet_title
                )
                
                if not success:
                    # Fall back to basic export if formatting fails
                    df.to_excel(file_path, sheet_name="Clearance", index=False)
                    logger.warning("Fell back to basic Excel export after formatting failed")
            else:
                # User cancelled or basic export
                df.to_excel(file_path, sheet_name="Clearance", index=False)
            
            logger.info(f"Exported to Excel file: {file_path}")
            
            if parent_widget:
                QMessageBox.information(
                    parent_widget,
                    "Export Complete",
                    f"Successfully exported to Excel file:\n{file_path}"
                )
            
            return True
        
        except Exception as e:
            error_msg = f"Error exporting to Excel file: {str(e)}"
            logger.error(error_msg)
            
            if parent_widget:
                QMessageBox.critical(parent_widget, "Export Error", error_msg)
            
            return False
    
    def get_rules_from_pivot(self, rule_type: RuleType = None) -> List[BaseRule]:
        """
        Get rules from current pivot data
        If rule_type is provided, convert to that type, otherwise use the type in the pivot data
        """
        try:
            # Get pivot data
            pivot_data = self.pivot_widget.get_pivot_data()
            
            if not pivot_data:
                logger.warning("No pivot data to convert to rules")
                return []
            
            # Use provided rule type or the one from pivot data
            if rule_type is None:
                rule_type = pivot_data.rule_type
            
            # Convert based on rule type
            if rule_type == RuleType.CLEARANCE:
                return pivot_data.to_clearance_rules()
            elif rule_type == RuleType.SHORT_CIRCUIT:
                return pivot_data.to_short_circuit_rules()
            elif rule_type == RuleType.UNROUTED_NET:
                return pivot_data.to_unrouted_net_rules()
            else:
                logger.warning(f"Unsupported rule type for conversion: {rule_type.value}")
                return []
        
        except Exception as e:
            error_msg = f"Error converting pivot data to rules: {str(e)}"
            logger.error(error_msg)
            return []
    
    def set_pivot_from_clearance_rules(self, rules: List[BaseRule]) -> bool:
        """Set pivot data from clearance rules"""
        try:
            # Check if rules are valid
            if not rules:
                logger.warning("No rules to convert to pivot data")
                return False
            
            # Create pivot data from rules
            pivot_data = ExcelPivotData.from_clearance_rules(rules)
            
            if not pivot_data:
                logger.warning("Failed to create pivot data from rules")
                return False
            
            # Set in widget
            self.pivot_widget.set_pivot_data(pivot_data)
            
            # Clear current file info
            self.current_file_path = None
            self.current_sheet_name = None
            
            logger.info(f"Set pivot data from {len(rules)} clearance rules")
            return True
        
        except Exception as e:
            error_msg = f"Error setting pivot data from rules: {str(e)}"
            logger.error(error_msg)
            return False
    
    def _on_data_changed(self):
        """Handle pivot data changes"""
        logger.debug("Pivot data changed")
