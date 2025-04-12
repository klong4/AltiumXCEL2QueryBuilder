#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Pivot Table Widget
================

Custom widget for displaying and editing pivot table data for rule generation.
"""

import logging
from typing import Dict, List, Optional, Union, Tuple
import pandas as pd
import numpy as np

from PyQt5.QtWidgets import (
    QWidget, QTableView, QHeaderView, QVBoxLayout, QHBoxLayout, 
    QLabel, QComboBox, QPushButton, QGroupBox, QFormLayout,
    QLineEdit, QSpinBox, QDoubleSpinBox, QMessageBox, QFileDialog
)
from PyQt5.QtCore import Qt, QAbstractTableModel, QModelIndex, QVariant, pyqtSignal
from PyQt5.QtGui import QColor, QBrush

from models.excel_data import ExcelPivotData
from models.rule_model import UnitType, RuleType

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
            self.data_array = pivot_data.values
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
        return len(self.headers) + 1 if self.headers else 0  # +1 for the index column
    
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
                return str(self.index_column[row])
              # Data columns
            data_col = col - 1
            if data_col < len(self.headers):
                value = self.data_array[row, data_col]
                # Format numeric values
                if isinstance(value, (int, float)):
                    if np.isnan(value):
                        return ""
                    return str(value)
                elif pd.isna(value):
                    return ""
                else:
                    return str(value)
        
        # Handle background color role
        elif role == Qt.BackgroundRole:
            # First column has different color
            if col == 0:
                return QBrush(QColor("#383838"))
            
            # Alternating row colors
            if row % 2 == 0:
                return QBrush(QColor("#323232"))
            else:
                return QBrush(QColor("#2d2d2d"))
        
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
            # First column header
            if section == 0:
                return "Net Class"
            
            # Data column headers
            if section - 1 < len(self.headers):
                return str(self.headers[section - 1])
        
        return QVariant()
    
    def flags(self, index):
        """Return item flags for the given index"""
        flags = super().flags(index)
        
        # Make first column read-only
        if index.column() == 0:
            return flags
        
        # Make other cells editable
        if self.editable:
            flags |= Qt.ItemIsEditable
        
        return flags
    
    def setData(self, index, value, role=Qt.EditRole):
        """Set data for the given index and role"""
        if not index.isValid() or role != Qt.EditRole:
            return False
        
        row, col = index.row(), index.column()
        
        # Only allow editing data cells, not the index column
        if col == 0:
            return False
        
        # Update the data array
        data_col = col - 1
        original_value = self.data_array[row, data_col]
        
        # Keep string variables as strings, try converting others to float
        if isinstance(value, str) and value.strip().upper() in ['D', 'F']:
             new_value = value.strip().upper()
        else:
            try:
                if value == "":
                    new_value = np.nan
                else:
                    new_value = float(value)
            except (ValueError, TypeError):
                logger.warning(f"Invalid value for cell ({row}, {col}): {value}, keeping original.")
                # Optionally show a message to the user or revert
                # For now, just don't update if conversion fails
                return False 

        if new_value != original_value:
            self.data_array[row, data_col] = new_value
            self.dataChanged.emit(index, index)
            self.data_changed.emit() # Emit custom signal if needed elsewhere
            return True
        return False

    def replace_variables_in_data(self, variables: Dict[str, float]) -> bool:
        """Replace string variables ('D', 'F') in the data array with numeric values."""
        modified = False
        rows, cols = self.data_array.shape
        
        # Create lists of QModelIndex for changed cells
        top_left_list = []
        bottom_right_list = []

        for r in range(rows):
            for c in range(cols):
                current_value = self.data_array[r, c]
                if isinstance(current_value, str):
                    var_upper = current_value.strip().upper()
                    if var_upper in variables:
                        new_value = variables[var_upper]
                        if self.data_array[r, c] != new_value: # Check if value actually changes
                            self.data_array[r, c] = new_value
                            modified = True
                            # Store index for dataChanged signal (column needs +1 for view)
                            model_index = self.index(r, c + 1)
                            top_left_list.append(model_index)
                            bottom_right_list.append(model_index)

        if modified:
            # Emit dataChanged for all modified cells efficiently
            # It's often simpler and sometimes more efficient to just reset the model
            # if many cells change, but let's try emitting specific signals first.
            # self.dataChanged.emit(min_index, max_index) # This requires finding min/max row/col
            # For simplicity, let's emit for each cell or reset.
            # Resetting the model is easiest if many cells might change:
            self.beginResetModel()
            self.endResetModel()
            self.data_changed.emit() # Emit custom signal
            logger.info(f"Replaced variables {list(variables.keys())} in pivot data model.")

        return modified

    def get_updated_pivot_data(self) -> ExcelPivotData:
        """Get updated pivot data from the model"""
        if self.pivot_data is None:
            return None
        
        # Create new pivot data with updated values
        updated_pivot_data = ExcelPivotData(self.pivot_data.rule_type)
        updated_pivot_data.row_index = self.index_column
        updated_pivot_data.column_index = self.headers
        updated_pivot_data.values = self.data_array
        updated_pivot_data.unit = self.pivot_data.unit
        
        # Reconstruct DataFrame
        df = pd.DataFrame(index=range(len(self.index_column)), columns=["NetClass"] + self.headers)
        df["NetClass"] = self.index_column
        
        for row_idx, row_name in enumerate(self.index_column):
            for col_idx, col_name in enumerate(self.headers):
                df.loc[row_idx, col_name] = self.data_array[row_idx, col_idx]
        
        updated_pivot_data.pivot_df = df
        
        return updated_pivot_data


class PivotTableWidget(QWidget):
    """Widget to display and edit pivot table data"""
    
    rules_generated = pyqtSignal(list)
    
    def __init__(self, parent=None):
        """Initialize pivot table widget"""
        super().__init__(parent)
        
        self.pivot_data = None
        self.model = PivotTableModel(self)
        self.model.data_changed.connect(self._on_model_data_changed)
        
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
            self.rule_type_combo.addItem(rule_type.value, rule_type)
        options_layout.addRow("Rule Type:", self.rule_type_combo)
        
        self.unit_combo = QComboBox()
        # Populate with UnitType enum values
        for unit_type in UnitType:
            self.unit_combo.addItem(unit_type.value, unit_type)
        options_layout.addRow("Units:", self.unit_combo)
        
        self.rule_prefix_input = QLineEdit("Rule_")
        options_layout.addRow("Rule Name Prefix:", self.rule_prefix_input)
        
        options_group.setLayout(options_layout)
        layout.addWidget(options_group)
        
        # --- Action Buttons --- 
        button_layout = QHBoxLayout()
        self.generate_button = QPushButton("Generate Rules")
        self.generate_button.clicked.connect(self._generate_rules)
        button_layout.addWidget(self.generate_button)
        
        self.export_button = QPushButton("Export to Excel")
        self.export_button.clicked.connect(self._export_to_excel)
        button_layout.addWidget(self.export_button)
        
        layout.addLayout(button_layout)
        
        self.setLayout(layout)

    def set_pivot_data(self, pivot_data: ExcelPivotData):
        """Set the pivot data and update the view"""
        self.pivot_data = pivot_data
        self.model.set_pivot_data(pivot_data)
        
        # Update unit combo based on loaded data
        if pivot_data and pivot_data.unit:
            index = self.unit_combo.findData(pivot_data.unit)
            if index >= 0:
                self.unit_combo.setCurrentIndex(index)
        
        # Update rule type combo based on loaded data
        if pivot_data and pivot_data.rule_type:
            index = self.rule_type_combo.findData(pivot_data.rule_type)
            if index >= 0:
                self.rule_type_combo.setCurrentIndex(index)
                # Set default prefix based on rule type
                self.rule_prefix_input.setText(f"{pivot_data.rule_type.value}_")

    def _on_model_data_changed(self):
        """Handle data changes in the model"""
        # Potentially update internal state or trigger validation if needed
        pass

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

        # 3. Generate rules based on type
        generated_rules = []
        try:
            if selected_rule_type == RuleType.CLEARANCE:
                generated_rules = updated_pivot_data.to_clearance_rules(rule_name_prefix=rule_prefix)
            # Add elif blocks for other rule types as they are implemented
            # elif selected_rule_type == RuleType.WIDTH:
            #     generated_rules = updated_pivot_data.to_width_rules(rule_name_prefix=rule_prefix)
            else:
                QMessageBox.warning(self, "Not Implemented", f"Rule generation for '{selected_rule_type.value}' is not yet implemented.")
                logger.warning(f"Rule generation not implemented for type: {selected_rule_type.value}")
                return

            if not generated_rules:
                 # Check if None was returned due to an issue during conversion vs just no rules generated
                 if generated_rules is None: # Explicit check for None which might indicate an error state in conversion
                     QMessageBox.warning(self, "Generation Warning", "Rule generation resulted in an unexpected state. Check logs.")
                     logger.warning("Rule generation function returned None unexpectedly.")
                 else: # Empty list means no rules applicable
                     QMessageBox.information(self, "Generation Info", "No applicable rules were generated based on the current data and options.")
                     logger.info("Rule generation completed, but no rules were applicable.")
                 # Don't emit signal if no rules or error
                 return
                 
            logger.info(f"Successfully generated {len(generated_rules)} rules.")
            
            # 4. Emit the signal
            self.rules_generated.emit(generated_rules)
            QMessageBox.information(self, "Generation Successful", f"Successfully generated {len(generated_rules)} rules. Check the 'Rule Editor' tab.")

        except Exception as e:
            error_msg = f"An error occurred during rule generation: {str(e)}"
            logger.error(error_msg, exc_info=True)
            QMessageBox.critical(self, "Generation Error", error_msg)

    def _export_to_excel(self):
        """Export the current pivot table data to an Excel file"""
        # ... existing code ...
