#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Pivot Table Widget
================

Custom widget for displaying and editing pivot table data for rule generation.
"""

import logging # Add logging import
import numpy as np # Add numpy import
import pandas as pd # Add pandas import
import os # Add os import

from typing import Dict, List, Optional, Union, Tuple
from PyQt5.QtWidgets import (
    QWidget, QTableView, QHeaderView, QVBoxLayout, QHBoxLayout,
    QLabel, QComboBox, QPushButton, QGroupBox, QFormLayout,
    QLineEdit, QSpinBox, QDoubleSpinBox, QMessageBox, QFileDialog
)
from PyQt5.QtCore import Qt, QAbstractTableModel, QModelIndex, QVariant, pyqtSignal
from PyQt5.QtGui import QColor, QBrush

from models.excel_data import ExcelPivotData
from models.rule_model import UnitType, RuleType
# Import RuleGeneratorError if needed for specific exception handling
from services.rule_generator import RuleGenerator, RuleGeneratorError


logger = logging.getLogger(__name__)

class PivotTableModel(QAbstractTableModel):
    """Model for pivot table data to be displayed in a QTableView"""

    data_changed = pyqtSignal()

    def __init__(self, parent=None):
        """Initialize pivot table model"""
        super().__init__(parent)
        self.pivot_data = None
        self.headers = []
        self.index_column = []
        self.data_array = np.array([])
        self.editable = True

    def set_pivot_data(self, pivot_data: ExcelPivotData):
        """Set the pivot data to display"""
        self.beginResetModel()
        self.pivot_data = pivot_data

        if pivot_data is not None and pivot_data.pivot_df is not None:
            self.headers = pivot_data.column_index
            self.index_column = pivot_data.row_index
            # Convert to numeric, coercing errors to NaN, but keep original object type if possible
            try:
                # Attempt conversion using pandas.to_numeric for better handling
                numeric_values = pivot_data.pivot_df.apply(pd.to_numeric, errors='coerce').values
                self.data_array = numeric_values
            except Exception as e:
                logger.warning(f"Could not convert all pivot data to numeric using pd.to_numeric: {e}. Keeping original types.")
                # Fallback: Keep original data but ensure it's a NumPy array
                self.data_array = np.array(pivot_data.values, dtype=object) # Use dtype=object to preserve strings
        else:
            self.headers = []
            self.index_column = []
            self.data_array = np.array([])

        self.endResetModel()
        logger.info(f"Pivot table model updated with {len(self.index_column)} rows and {len(self.headers)} columns")

    def rowCount(self, parent=None):
        """Return number of rows in the model"""
        return len(self.index_column) if self.index_column else 0

    def columnCount(self, parent=None):
        """Return number of columns in the model"""
        # +1 for the index column
        return len(self.headers) + 1 if self.headers else 0

    def data(self, index, role=Qt.DisplayRole):
        """Return data for the given index and role"""
        if not index.isValid():
            return QVariant()

        row, col = index.row(), index.column()

        # Check if row and column are valid
        if row < 0 or row >= self.rowCount() or col < 0 or col >= self.columnCount():
            return QVariant()

        # Handle display role
        if role == Qt.DisplayRole or role == Qt.EditRole:
            # First column is the index column
            if col == 0:
                return str(self.index_column[row]) if row < len(self.index_column) else QVariant()

            # Data columns
            data_col = col - 1
            # Check bounds for data_array
            if row < self.data_array.shape[0] and data_col < self.data_array.shape[1]:
                value = self.data_array[row, data_col]
                # Handle potential NaN or None values gracefully for display
                if pd.isna(value) or value is None:
                    return "" # Display empty string for NaN/None
                # Format numeric values nicely, keep strings as is
                if isinstance(value, (int, float, np.number)):
                     # Basic formatting, could be refined (e.g., specific precision)
                     # Check if it's a floating point number before formatting
                     if isinstance(value, (float, np.floating)):
                         return f"{value:.3f}" # Format floats to 3 decimal places
                     else:
                         return str(value) # Integers as strings
                return str(value) # Return string representation for other types (like 'D', 'F')
            else:
                logger.warning(f"Data index out of bounds: row={row}, data_col={data_col}")
                return QVariant() # Out of bounds

        # Handle background color role
        elif role == Qt.BackgroundRole:
            # First column has different color (e.g., slightly darker gray)
            if col == 0:
                 return QBrush(QColor("#404040")) # Darker gray for index column

            # Alternating row colors for data cells
            if row % 2 == 0:
                 return QBrush(QColor("#323232")) # Dark theme color
            else:
                 return QBrush(QColor("#2d2d2d")) # Slightly lighter dark theme color

        # Handle text alignment role
        elif role == Qt.TextAlignmentRole:
            # Center-align all cells
            return Qt.AlignCenter

        return QVariant()

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        """Return header data for the given section and orientation"""
        if role != Qt.DisplayRole:
            return QVariant()

        if orientation == Qt.Horizontal:
            # First column header (Index Name - e.g., "NetClass")
            if section == 0:
                 # Try to get the index name from pivot_data if available
                 index_name = "Index" # Default
                 if self.pivot_data and self.pivot_data.pivot_df is not None and self.pivot_data.pivot_df.index.name:
                     index_name = self.pivot_data.pivot_df.index.name
                 return str(index_name)

            # Data column headers
            header_index = section - 1
            if 0 <= header_index < len(self.headers):
                return str(self.headers[header_index])

        elif orientation == Qt.Vertical:
             # Use row number as vertical header (optional, could use index_column values too)
             return str(section) # Simple row number

        return QVariant()

    def flags(self, index):
        """Return item flags for the given index"""
        flags = super().flags(index)

        # Make first column read-only
        if index.column() == 0:
            # Remove editable flag if present
            flags &= ~Qt.ItemIsEditable
        # Make other cells editable if self.editable is True
        elif self.editable:
             flags |= Qt.ItemIsEditable

        return flags

    def setData(self, index, value, role=Qt.EditRole):
        """Set data for the given index and role"""
        if not index.isValid() or role != Qt.EditRole:
             return False

        row, col = index.row(), index.column()

        # Only allow editing data cells, not the index column
        if col == 0:
             logger.warning("Attempted to edit read-only index column.")
             return False

        # Update the data array
        data_col = col - 1
        if row >= self.data_array.shape[0] or data_col >= self.data_array.shape[1]:
            logger.error(f"setData index out of bounds: row={row}, col={col} (data_col={data_col})")
            return False

        original_value = self.data_array[row, data_col]
        new_value = None

        # Keep string variables as strings, try converting others to float
        if isinstance(value, str):
            stripped_value = value.strip() # Keep case for D/F initially
            if stripped_value.upper() in ['D', 'F']:
                 # Store D/F as strings, but maybe handle them differently later?
                 # For now, store as string. Conversion happens in get_updated_pivot_data or generation
                 new_value = stripped_value
            else:
                 # Try converting other strings to float
                 try:
                     if stripped_value == "": # Treat empty string as NaN
                         new_value = np.nan
                     else:
                         new_value = float(stripped_value)
                 except ValueError:
                     logger.warning(f"Could not convert input '{value}' to float for cell ({row}, {data_col}). Keeping as string.")
                     new_value = value # Keep original string if conversion fails and it wasn't D/F
        else:
             # Handle non-string input (e.g., from spinbox, already numeric)
             try:
                 new_value = float(value)
             except (ValueError, TypeError):
                 logger.warning(f"Could not convert input '{value}' (type: {type(value)}) to float for cell ({row}, {data_col}).")
                 # Decide fallback: keep original, set NaN, or keep as string? Let's try keeping original.
                 new_value = original_value # Revert to original if conversion fails


        # Check if the value actually changed (handle NaN comparison)
        # Use pd.isna for robust NaN check
        if (new_value == original_value) or (pd.isna(new_value) and pd.isna(original_value)):
            return False # No actual change

        self.data_array[row, data_col] = new_value
        self.dataChanged.emit(index, index, [role]) # Emit signal for the specific cell
        self.data_changed.emit() # Emit custom signal indicating general data change
        logger.debug(f"Data changed at ({row}, {data_col}) from {original_value} to {new_value}")
        return True

    def replace_variables_in_data(self, variables: Dict[str, float]) -> bool:
        """Replace string variables ('D', 'F') in the data array with numeric values."""
        if not variables or self.data_array is None or self.data_array.size == 0:
            return False

        modified = False
        rows, cols = self.data_array.shape

        # Create lists of QModelIndex for changed cells
        changed_indexes = []

        for r in range(rows):
             for c in range(cols):
                 current_value = self.data_array[r, c]
                 # Check if it's a string and matches a variable key (case-insensitive)
                 if isinstance(current_value, str):
                     upper_val = current_value.upper()
                     if upper_val in variables:
                         new_value = variables[upper_val]
                         # Check if value actually changes (handle potential type diff)
                         try:
                             # Compare numerically if possible
                             if float(self.data_array[r, c]) != new_value:
                                 self.data_array[r, c] = new_value
                                 modified = True
                                 index = self.index(r, c + 1) # +1 for model column
                                 changed_indexes.append(index)
                                 logger.debug(f"Replaced variable '{upper_val}' at ({r}, {c}) with {new_value}")
                         except (ValueError, TypeError):
                             # If current value wasn't numeric, compare directly
                             if self.data_array[r, c] != str(new_value): # Compare as string if needed
                                 self.data_array[r, c] = new_value
                                 modified = True
                                 index = self.index(r, c + 1) # +1 for model column
                                 changed_indexes.append(index)
                                 logger.debug(f"Replaced variable '{upper_val}' at ({r}, {c}) with {new_value}")


        if modified:
             # Emit dataChanged for the affected range
             if changed_indexes:
                 # Find min/max row/col for the signal range (more efficient)
                 min_row = min(idx.row() for idx in changed_indexes)
                 max_row = max(idx.row() for idx in changed_indexes)
                 min_col = min(idx.column() for idx in changed_indexes)
                 max_col = max(idx.column() for idx in changed_indexes)
                 top_left = self.index(min_row, min_col)
                 bottom_right = self.index(max_row, max_col)
                 self.dataChanged.emit(top_left, bottom_right, [Qt.DisplayRole, Qt.EditRole])
                 self.data_changed.emit() # Emit custom signal
                 logger.info(f"Replaced variables {list(variables.keys())} in model data.")

        return modified

    def get_updated_pivot_data(self) -> Optional[ExcelPivotData]:
        """Get updated pivot data from the model"""
        if self.pivot_data is None:
             logger.warning("Cannot get updated pivot data, original pivot_data is None.")
             return None

        # Create new pivot data object based on the original type
        updated_pivot_data = ExcelPivotData(
            rule_type=self.pivot_data.rule_type
            # unit=self.pivot_data.unit # Removed: unit is handled internally
        )
        # Manually set the unit after initialization if needed, based on original data
        if self.pivot_data and hasattr(self.pivot_data, 'unit'):
            updated_pivot_data.unit = self.pivot_data.unit
        
        updated_pivot_data.row_index = self.index_column[:] # Copy lists
        updated_pivot_data.column_index = self.headers[:]
        # Ensure the data array is copied and potentially converted if D/F were stored as strings
        # For now, assume data_array holds floats or NaNs after edits/replacements
        updated_pivot_data.values = self.data_array.copy() # Copy numpy array

        # Reconstruct DataFrame (optional but good for consistency)
        try:
            # Use the index_column as the DataFrame index
            # Ensure data shape matches headers and index
            num_rows = len(self.index_column)
            num_cols = len(self.headers)
            current_data_shape = updated_pivot_data.values.shape

            if current_data_shape == (num_rows, num_cols):
                df = pd.DataFrame(updated_pivot_data.values, index=self.index_column, columns=self.headers)
            elif current_data_shape[0] == num_rows and current_data_shape[1] > num_cols:
                 # If data has more columns than headers, slice the data
                 logger.warning(f"Data shape {current_data_shape} has more columns than headers ({num_cols}). Slicing data.")
                 df = pd.DataFrame(updated_pivot_data.values[:, :num_cols], index=self.index_column, columns=self.headers)
            elif current_data_shape[0] > num_rows and current_data_shape[1] == num_cols:
                 # If data has more rows than index, slice the data
                 logger.warning(f"Data shape {current_data_shape} has more rows than index ({num_rows}). Slicing data.")
                 df = pd.DataFrame(updated_pivot_data.values[:num_rows, :], index=self.index_column, columns=self.headers)
            else:
                # If shapes don't match in a way we can't easily fix, raise the original error after logging
                logger.error(f"Shape mismatch: Data shape {current_data_shape}, Index length {num_rows}, Headers length {num_cols}")
                # Re-raise the original error or a more specific one
                raise ValueError(f"Shape of passed values is {current_data_shape}, indices imply ({num_rows}, {num_cols})")

            # Assign the index name if available
            index_name = "Index"
            if self.pivot_data.pivot_df is not None and self.pivot_data.pivot_df.index.name:
                index_name = self.pivot_data.pivot_df.index.name
            df.index.name = index_name
            updated_pivot_data.pivot_df = df
            logger.debug("Reconstructed DataFrame for updated pivot data.")
        except Exception as e:
            logger.error(f"Error reconstructing DataFrame for updated pivot data: {e}", exc_info=True)
            updated_pivot_data.pivot_df = None # Set to None if reconstruction fails

        return updated_pivot_data


class PivotTableWidget(QWidget):
    """Widget to display and edit pivot table data"""

    rules_generated = pyqtSignal(list) # Emits list[BaseRule]

    def __init__(self, parent=None):
        """Initialize pivot table widget"""
        super().__init__(parent)

        self.pivot_data: Optional[ExcelPivotData] = None
        self.model = PivotTableModel(self)
        # Connect the model's data_changed signal to the main window's handler if needed
        # self.model.data_changed.connect(...) # Connect this in main_window after creating the widget

        self._init_ui()

    def _init_ui(self):
        """Initialize the user interface"""
        layout = QVBoxLayout(self)

        # --- Table View ---
        self.table_view = QTableView()
        self.table_view.setModel(self.model)
        self.table_view.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table_view.verticalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        layout.addWidget(self.table_view)

        # --- Options Group ---
        options_group = QGroupBox("Rule Generation Options")
        options_layout = QFormLayout()

        self.rule_type_combo = QComboBox()
        # Populate with RuleType enum values
        for rule_type in RuleType:
            # Use rule_type.value for display text, rule_type enum member as data
            self.rule_type_combo.addItem(rule_type.value, rule_type)
        options_layout.addRow("Rule Type:", self.rule_type_combo)

        self.unit_combo = QComboBox()
        # Populate with UnitType enum values
        for unit_type in UnitType:
            # Use unit_type.value for display text, unit_type enum member as data
            self.unit_combo.addItem(unit_type.value, unit_type)
        options_layout.addRow("Units:", self.unit_combo)

        self.rule_prefix_input = QLineEdit("Rule_") # Default prefix
        options_layout.addRow("Rule Name Prefix:", self.rule_prefix_input)

        options_group.setLayout(options_layout)
        layout.addWidget(options_group)

        # --- Action Buttons ---
        button_layout = QHBoxLayout()
        self.generate_button = QPushButton("Generate Rules")
        self.generate_button.clicked.connect(self._generate_rules)
        button_layout.addWidget(self.generate_button)

        self.export_button = QPushButton("Export to Excel")
        # Connect to the internal _export_to_excel method
        self.export_button.clicked.connect(self.export_to_excel) # Connect to public method
        button_layout.addWidget(self.export_button)

        layout.addLayout(button_layout)

        self.setLayout(layout)

    def set_pivot_data(self, pivot_data: ExcelPivotData):
        """Set the pivot data and update the view and controls"""
        self.pivot_data = pivot_data
        self.model.set_pivot_data(pivot_data)

        # Update unit combo based on loaded data
        if pivot_data and pivot_data.unit:
            # Find the index corresponding to the UnitType enum member
            index = self.unit_combo.findData(pivot_data.unit)
            if index >= 0:
                self.unit_combo.setCurrentIndex(index)
            else:
                logger.warning(f"Unit type {pivot_data.unit} from loaded data not found in combo box.")
        else:
             # Default to MIL if not specified
             index = self.unit_combo.findData(UnitType.MIL)
             if index >= 0: self.unit_combo.setCurrentIndex(index)


        # Update rule type combo based on loaded data
        if pivot_data and pivot_data.rule_type:
            # Find the index corresponding to the RuleType enum member
            index = self.rule_type_combo.findData(pivot_data.rule_type)
            if index >= 0:
                self.rule_type_combo.setCurrentIndex(index)
                # Set default prefix based on rule type
                self.rule_prefix_input.setText(f"{pivot_data.rule_type.value}_")
            else:
                 logger.warning(f"Rule type {pivot_data.rule_type} from loaded data not found in combo box.")
                 # Default to first item if not found? Or leave as is? Let's default.
                 if self.rule_type_combo.count() > 0:
                     self.rule_type_combo.setCurrentIndex(0)
                     default_type = self.rule_type_combo.currentData()
                     self.rule_prefix_input.setText(f"{default_type.value}_")

        else:
             # Default to Clearance if not specified
             index = self.rule_type_combo.findData(RuleType.CLEARANCE)
             if index >= 0:
                 self.rule_type_combo.setCurrentIndex(index)
                 self.rule_prefix_input.setText(f"{RuleType.CLEARANCE.value}_")


    # Slot removed, connection should be made in main_window
    # def _on_model_data_changed(self):
    #     pass

    def _generate_rules(self):
        """Generate rules based on the current pivot table data and options"""
        logger.info("Generate Rules button clicked.")

        # 1. Get updated pivot data from the model
        updated_pivot_data = self.model.get_updated_pivot_data()
        if updated_pivot_data is None or updated_pivot_data.values is None:
            QMessageBox.warning(self, "Generation Error", "No valid pivot data available to generate rules.")
            logger.warning("Rule generation attempted with no valid pivot data.")
            return

        # 2. Get selected options
        selected_rule_type: RuleType = self.rule_type_combo.currentData()
        selected_unit: UnitType = self.unit_combo.currentData()
        rule_prefix = self.rule_prefix_input.text().strip()

        if not selected_rule_type:
            QMessageBox.warning(self, "Generation Error", "Please select a valid Rule Type.")
            logger.warning("Rule generation attempted without a selected rule type.")
            return

        # Update the pivot_data object with current selections (important for conversion)
        updated_pivot_data.rule_type = selected_rule_type
        updated_pivot_data.unit = selected_unit

        logger.info(f"Generating rules of type: {selected_rule_type.value}, Unit: {selected_unit.value}, Prefix: '{rule_prefix}'")

        # 3. Generate rules based on type using methods from ExcelPivotData
        generated_rules = []
        try:
            # Use the methods defined in ExcelPivotData
            if selected_rule_type == RuleType.CLEARANCE:
                # Assuming to_clearance_rules exists and handles potential errors
                generated_rules = updated_pivot_data.to_clearance_rules(rule_name_prefix=rule_prefix)
            elif selected_rule_type == RuleType.WIDTH:
                 # Assuming to_width_rules exists
                 generated_rules = updated_pivot_data.to_width_rules(rule_name_prefix=rule_prefix)
            # Add elif blocks for other rule types as they are implemented in ExcelPivotData
            # elif selected_rule_type == RuleType.ROUTING_VIA_STYLE:
            #     generated_rules = updated_pivot_data.to_routing_via_style_rules(rule_name_prefix=rule_prefix)
            else:
                QMessageBox.warning(self, "Not Implemented", f"Rule generation for '{selected_rule_type.value}' is not yet implemented in the data model.")
                logger.warning(f"Rule generation not implemented for type: {selected_rule_type.value}")
                return

            # Check if rule generation failed (returned None) or produced an empty list
            if generated_rules is None:
                 # This indicates an error occurred within the to_..._rules method
                 QMessageBox.critical(self, "Generation Error", f"An error occurred while generating {selected_rule_type.value} rules. Check logs for details.")
                 logger.error(f"Rule generation method for {selected_rule_type.value} returned None.")
                 return # Don't emit signal
            elif not generated_rules:
                 # No error, but no rules were generated (e.g., empty data)
                 QMessageBox.information(self, "No Rules Generated", "No rules were generated based on the current data and options.")
                 logger.info("Rule generation resulted in an empty list.")
                 # Optionally emit empty list? Or just do nothing? Let's do nothing.
                 return

            logger.info(f"Successfully generated {len(generated_rules)} rules.")

            # 4. Emit the signal with the list of generated BaseRule objects
            self.rules_generated.emit(generated_rules)
            QMessageBox.information(self, "Generation Successful", f"Successfully generated {len(generated_rules)} rules. Check the 'Rule Manager' tab.")

        except AttributeError as ae:
             # Handle cases where the generation method (e.g., to_width_rules) doesn't exist
             error_msg = f"Rule generation method for '{selected_rule_type.value}' is not implemented."
             logger.error(error_msg, exc_info=False) # No need for full traceback here
             QMessageBox.warning(self, "Not Implemented", error_msg)
        except RuleGeneratorError as rge:
             # Catch specific errors from the generation process if defined
             error_msg = f"Error during rule generation: {str(rge)}"
             logger.error(error_msg, exc_info=True)
             QMessageBox.critical(self, "Generation Error", error_msg)
        except Exception as e:
            # Catch any other unexpected errors
            error_msg = f"An unexpected error occurred during rule generation: {str(e)}"
            logger.error(error_msg, exc_info=True)
            QMessageBox.critical(self, "Generation Error", error_msg)

    # Renamed from _export_to_excel to avoid conflict if main_window calls it directly
    # Made public so main_window can call it if needed, though button connects here.
    def export_to_excel(self, file_path: Optional[str] = None) -> bool:
        """Export the current pivot table data to an Excel file.
        If file_path is not provided, a save dialog is shown.
        Returns True on success, False on failure or cancellation.
        """
        logger.info(f"Exporting pivot data to Excel...")
        updated_pivot_data = self.model.get_updated_pivot_data()

        if updated_pivot_data is None or updated_pivot_data.pivot_df is None:
            QMessageBox.warning(self, "Export Error", "No valid pivot data available to export.")
            logger.warning("Excel export failed: No valid pivot data.")
            return False # Indicate failure

        if not file_path:
            # Show save file dialog if no path is given
            last_dir = "" # Ideally get from config
            default_filename = "pivot_export.xlsx"
            file_path, _ = QFileDialog.getSaveFileName(
                self,
                "Export Pivot Table to Excel",
                os.path.join(last_dir, default_filename), # Suggest a filename
                "Excel Files (*.xlsx);;All Files (*)"
            )
            if not file_path:
                logger.info("Excel export cancelled by user.")
                return False # User cancelled

        # Ensure the filename ends with .xlsx
        if not file_path.lower().endswith('.xlsx'):
            file_path += '.xlsx'

        try:
            # Use pandas to export the DataFrame
            # Ensure index is included as it's meaningful (Net Classes)
            updated_pivot_data.pivot_df.to_excel(file_path, index=True)
            logger.info(f"Successfully exported pivot data to {file_path}")
            # Show success message only if dialog was used (file_path was initially None)
            # If called programmatically, the caller might show the message.
            # if not file_path: # This logic is flawed, check if the button triggered it?
            QMessageBox.information(self, "Export Successful", f"Successfully exported pivot data to:\n{file_path}")

            # Mark as saved (if tracking unsaved changes for pivot table)
            # self.mark_saved() # Needs implementation if tracking changes
            return True # Indicate success

        except Exception as e:
            error_msg = f"An error occurred during Excel export: {str(e)}"
            logger.error(error_msg, exc_info=True)
            QMessageBox.critical(self, "Export Error", error_msg)
            return False # Indicate failure

    def has_unsaved_changes(self) -> bool:
        """Checks if the pivot table data has unsaved changes."""
        # Basic check: Compare current model data with original loaded data
        # This is a simplified check. A more robust way would be to track edits via signals.
        if self.pivot_data is None:
            return False # No data loaded, so no changes

        # Get current data without reconstructing DataFrame if possible
        current_values = self.model.data_array
        current_index = self.model.index_column
        current_headers = self.model.headers

        if current_values is None or current_index is None or current_headers is None:
             # This case might indicate an issue, log it.
             logger.warning("Could not retrieve current data parts for unsaved changes check.")
             return False # Error getting current data parts

        # Compare numpy arrays (handle NaNs correctly)
        original_values = self.pivot_data.values

        # Ensure original_values is also float for comparison, handle potential errors
        try:
            # Make sure original_values is a numpy array first
            if not isinstance(original_values, np.ndarray):
                 original_values = np.array(original_values) # Attempt conversion
            original_values_float = original_values.astype(float)
        except (ValueError, TypeError, AttributeError) as e:
             # If original cannot be float, assume change if current is different type/shape
             logger.warning(f"Original pivot data values could not be cast to float for comparison: {e}. Comparing directly.")
             # Basic comparison if cast fails
             if not np.array_equal(original_values, current_values):
                 logger.debug("Unsaved changes detected: Direct array comparison failed after cast error.")
                 return True
             # If direct comparison also passes, proceed to check index/headers


        # Check shape first
        if original_values_float.shape != current_values.shape:
             logger.debug("Unsaved changes detected: Shape mismatch")
             return True # Shape mismatch means changes

        # Use numpy.array_equal with equal_nan=True for robust comparison
        if not np.array_equal(original_values_float, current_values, equal_nan=True):
             logger.debug("Unsaved changes detected: Data values differ")
             return True # Data values differ

        # Compare headers and index lists
        if self.pivot_data.row_index != current_index or \
           self.pivot_data.column_index != current_headers:
            logger.debug("Unsaved changes detected: Index or headers differ")
            return True

        return False # No changes detected

    # Add mark_saved if needed, similar to RulesManagerWidget
    # def mark_saved(self):
    #     # Reset unsaved changes flag/state if implementing robust tracking
    #     # This might involve storing the initial state more formally
    #     logger.debug("Pivot table marked as saved (unsaved changes flag reset).")
    #     # Emit a signal if the main window needs to know
    #     # self.model.data_changed.emit() # Or a specific saved signal
