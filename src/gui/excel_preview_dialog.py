#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Excel Preview Dialog
=================

Dialog for previewing and manipulating Excel data before import.
"""

import logging
from typing import Dict, List, Optional, Tuple, Any
import pandas as pd
import numpy as np

from PyQt5.QtWidgets import (
    QDialog, QTableView, QHeaderView, QVBoxLayout, QHBoxLayout, 
    QLabel, QComboBox, QPushButton, QGroupBox, QFormLayout,
    QDialogButtonBox, QCheckBox, QSpinBox, QFileDialog, QSplitter,
    QLineEdit, QSpacerItem, QSizePolicy, QMessageBox
)
from PyQt5.QtCore import Qt, QAbstractTableModel, QModelIndex, QVariant, pyqtSignal
from PyQt5.QtGui import QColor, QBrush

from models.rule_model import UnitType, RuleType

logger = logging.getLogger(__name__)

class ExcelPreviewModel(QAbstractTableModel):
    """Model for previewing Excel data in a table view"""
    
    data_changed = pyqtSignal()
    
    def __init__(self, parent=None):
        """Initialize Excel preview model"""
        super().__init__(parent)
        self.df = pd.DataFrame() # Explicitly initialize as empty DataFrame
        self.editable = True
    
    def set_dataframe(self, df: pd.DataFrame):
        """Set the dataframe to display"""
        self.beginResetModel()
        self.df = df
        self.endResetModel()
        logger.info(f"Excel preview model updated with {df.shape[0]} rows and {df.shape[1]} columns")
    
    def rowCount(self, parent=None):
        """Return number of rows in the model"""
        return self.df.shape[0] if self.df is not None else 0
    
    def columnCount(self, parent=None):
        """Return number of columns in the model"""
        return self.df.shape[1] if self.df is not None else 0
    
    def data(self, index, role=Qt.DisplayRole):
        """Return data for the given index and role"""
        if not index.isValid() or self.df is None:
            return QVariant()
        
        row, col = index.row(), index.column()
        
        # Check if row and column are valid
        if row < 0 or row >= self.rowCount() or col < 0 or col >= self.columnCount():
            return QVariant()
        
        # Handle display role
        if role == Qt.DisplayRole or role == Qt.EditRole:
            value = self.df.iloc[row, col]
            
            # Format numeric values
            if isinstance(value, (int, float)):
                if pd.isna(value) or np.isnan(value):
                    return ""
                return str(value)
            elif pd.isna(value):
                return ""
            else:
                return str(value)
        
        # Handle background color role
        elif role == Qt.BackgroundRole:
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
        if role != Qt.DisplayRole or self.df is None:
            return QVariant()
        
        if orientation == Qt.Horizontal:
            return str(self.df.columns[section])
        
        if orientation == Qt.Vertical:
            return str(section)
        
        return QVariant()
    
    def flags(self, index):
        """Return item flags for the given index"""
        flags = super().flags(index)
        
        # Make cells editable if editable is True
        if self.editable:
            flags |= Qt.ItemIsEditable
        
        return flags
    
    def setData(self, index, value, role=Qt.EditRole):
        """Set data for the given index and role"""
        if not index.isValid() or role != Qt.EditRole or self.df is None:
            return False
        
        row, col = index.row(), index.column()
        
        # Convert input to appropriate type
        try:
            # If cell is currently numeric, try to convert to float
            current_value = self.df.iloc[row, col]
            
            if isinstance(current_value, (int, float)):
                if value == "":
                    # Empty string becomes NaN
                    self.df.iloc[row, col] = np.nan
                else:
                    try:
                        self.df.iloc[row, col] = float(value)
                    except ValueError:
                        # If conversion fails, use as string
                        self.df.iloc[row, col] = value
            else:
                # For non-numeric cells, use the value as-is
                self.df.iloc[row, col] = value
            
            # Emit signal
            self.dataChanged.emit(index, index)
            self.data_changed.emit()
            return True
        except Exception as e:
            logger.warning(f"Invalid value for cell ({row}, {col}): {value}, error: {str(e)}")
            return False
    
    def get_dataframe(self) -> pd.DataFrame:
        """Get the current dataframe"""
        # Return an empty DataFrame if self.df is None
        return self.df.copy() if self.df is not None else pd.DataFrame()


class ExcelPreviewDialog(QDialog):
    """Dialog for previewing and manipulating Excel data before import"""
    
    def __init__(self, df: pd.DataFrame, sheet_name: str, parent=None):
        """Initialize Excel preview dialog"""
        super().__init__(parent)
        self.df = df
        self.sheet_name = sheet_name
        self.unit = UnitType.MIL
        self.skip_rows = 0
        self.end_row = 11 # -1 means no end row limit
        self.use_first_row_as_header = False # Default changed to False
        self.use_first_column_as_index = True
        
        self.setWindowTitle(f"Preview Excel File - {sheet_name}")
        self.resize(800, 600)
        
        # Initialize UI
        self._init_ui()
        
        # Load data
        self._load_data()
        
        logger.info(f"Excel preview dialog initialized for sheet: {sheet_name}")
    
    def _init_ui(self):
        """Initialize the UI components"""
        # Main layout
        main_layout = QVBoxLayout()
        self.setLayout(main_layout)

        # Top section with controls and variable replacement
        top_section_layout = QHBoxLayout()
        main_layout.addLayout(top_section_layout)

        # --- Import Options Group --- 
        options_group = QGroupBox("Import Options")
        options_layout = QFormLayout()
        options_group.setLayout(options_layout)
        top_section_layout.addWidget(options_group)

        # Skip rows option
        self.skip_rows_spin = QSpinBox()
        self.skip_rows_spin.setRange(0, 100)
        self.skip_rows_spin.setValue(self.skip_rows)
        self.skip_rows_spin.valueChanged.connect(self._on_options_changed)
        options_layout.addRow("Skip Rows:", self.skip_rows_spin)

        # End row option
        self.end_row_spin = QSpinBox()
        self.end_row_spin.setRange(-1, 1000000) # Allow large number of rows, -1 for no limit
        self.end_row_spin.setValue(self.end_row)
        self.end_row_spin.setToolTip("Specify the last row to include (-1 for no limit). Applied after skipping rows.")
        self.end_row_spin.valueChanged.connect(self._on_options_changed)
        options_layout.addRow("End Row:", self.end_row_spin)

        # Use first row as header option
        self.header_checkbox = QCheckBox()
        self.header_checkbox.setChecked(self.use_first_row_as_header)
        self.header_checkbox.stateChanged.connect(self._on_options_changed)
        options_layout.addRow("Use First Row as Header:", self.header_checkbox)

        # Use first column as index option
        self.index_col_checkbox = QCheckBox()
        self.index_col_checkbox.setChecked(self.use_first_column_as_index)
        self.index_col_checkbox.stateChanged.connect(self._on_options_changed)
        options_layout.addRow("Use First Column as Index:", self.index_col_checkbox)

        # --- Variable Replacement Group --- 
        variable_group = QGroupBox("Replace Variables (in Preview)")
        variable_layout = QFormLayout()
        variable_group.setLayout(variable_layout)
        top_section_layout.addWidget(variable_group)

        self.d_var_input = QLineEdit()
        self.d_var_input.setPlaceholderText("Enter numeric value for D")
        variable_layout.addRow("D =", self.d_var_input)

        self.f_var_input = QLineEdit()
        self.f_var_input.setPlaceholderText("Enter numeric value for F")
        variable_layout.addRow("F =", self.f_var_input)

        self.replace_button = QPushButton("Replace Variables Now")
        self.replace_button.clicked.connect(self._replace_variables)
        variable_layout.addRow(self.replace_button)

        # --- Table View --- 
        self.table_view = QTableView()
        self.model = ExcelPreviewModel(self)
        self.table_view.setModel(self.model)
        self.table_view.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table_view.verticalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        main_layout.addWidget(self.table_view)

        # --- Dialog Buttons --- 
        self.button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)
        main_layout.addWidget(self.button_box)

    def _load_data(self):
        """Load and process data based on current options"""
        try:
            processed_df = self.df.copy()
            
            # Skip rows
            if self.skip_rows > 0 and self.skip_rows < len(processed_df):
                processed_df = processed_df.iloc[self.skip_rows:]
            
            # Set header
            header_row_present = False
            if self.use_first_row_as_header and not processed_df.empty:
                processed_df.columns = processed_df.iloc[0]
                processed_df = processed_df[1:]
                header_row_present = True
            
            # Apply end row limit (relative to the data *after* skipping and potential header removal)
            if self.end_row != -1 and self.end_row >= 0:
                # Calculate the actual number of data rows to keep
                rows_to_keep = self.end_row
                if rows_to_keep < len(processed_df):
                     processed_df = processed_df.iloc[:rows_to_keep]

            # Set index (optional, might not be needed for preview model)
            # if self.use_first_column_as_index and not processed_df.empty:
            #     processed_df = processed_df.set_index(processed_df.columns[0])
            
            # Reset index for display consistency in the table model
            processed_df = processed_df.reset_index(drop=True)
            
            self.current_processed_df = processed_df # Store the processed df
            self.model.set_dataframe(processed_df)
            logger.info("Preview data reloaded with current options.")
        except Exception as e:
            logger.error(f"Error processing data for preview: {str(e)}")
            QMessageBox.warning(self, "Processing Error", f"Could not apply options: {str(e)}")
            # Fallback to original data if processing fails
            self.current_processed_df = self.df.copy()
            self.model.set_dataframe(self.current_processed_df)

    def _on_options_changed(self):
        """Handle changes in import options"""
        self.skip_rows = self.skip_rows_spin.value()
        self.end_row = self.end_row_spin.value()
        self.use_first_row_as_header = self.header_checkbox.isChecked()
        self.use_first_column_as_index = self.index_col_checkbox.isChecked()
        self._load_data() # Reload data with new options

    def _replace_variables(self):
        """Replace 'D' and 'F' variables in the preview table with user-provided values."""
        if self.current_processed_df is None:
            QMessageBox.warning(self, "No Data", "No data loaded in the preview.")
            return

        d_value_str = self.d_var_input.text().strip()
        f_value_str = self.f_var_input.text().strip()

        variables = {}
        try:
            if d_value_str:
                variables['D'] = float(d_value_str)
            if f_value_str:
                variables['F'] = float(f_value_str)
        except ValueError:
            QMessageBox.warning(self, "Invalid Input", "Please enter valid numeric values for D and F.")
            return

        if not variables:
            QMessageBox.information(self, "No Variables", "Please enter values for D or F to replace.")
            return

        modified_count = 0
        df_copy = self.current_processed_df.copy()

        for col in df_copy.columns:
            # Apply replacement only to object columns (likely strings)
            # Use pd.api.types.is_object_dtype for robust type checking
            if pd.api.types.is_object_dtype(df_copy[col]):
                for var, value in variables.items():
                    # Case-insensitive replacement
                    # Ensure comparison happens with strings
                    mask = df_copy[col].astype(str).str.upper() == str(var).upper()
                    if mask.any():
                        df_copy.loc[mask, col] = value
                        modified_count += mask.sum()
        
        # Convert columns back to numeric if possible after replacement
        for col in df_copy.columns:
             # Use pd.api.types.is_object_dtype for robust type checking
             if pd.api.types.is_object_dtype(df_copy[col]):
                 # Attempt conversion, ignore errors for non-numeric strings
                 df_copy[col] = pd.to_numeric(df_copy[col], errors='ignore')

        if modified_count > 0:
            # Update the model with the modified DataFrame
            self.current_processed_df = df_copy
            self.model.set_dataframe(self.current_processed_df)
            QMessageBox.information(self, "Replacement Complete", f"Replaced {modified_count} instance(s) of variables {list(variables.keys())} in the preview.")
        else:
            QMessageBox.information(self, "No Changes", "No instances of the specified variables were found in the preview data.")

    def get_processed_dataframe(self) -> pd.DataFrame:
        """Return the final processed dataframe after preview and potential modifications"""
        # Ensure the latest state of the dataframe (after potential variable replacement) is returned
        return self.current_processed_df.copy() if hasattr(self, 'current_processed_df') and self.current_processed_df is not None else pd.DataFrame()

    def get_import_options(self) -> Dict[str, Any]:
        """Return the selected import options"""
        return {
            "skip_rows": self.skip_rows,
            "header": 0 if self.use_first_row_as_header else None,
            "index_col": 0 if self.use_first_column_as_index else None,
            "nrows": self.end_row if self.end_row != -1 else None # Add nrows based on end_row for potential use in pandas read_excel
            # Add other relevant options if needed
        }
