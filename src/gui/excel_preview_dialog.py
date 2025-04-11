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
    QDialogButtonBox, QCheckBox, QSpinBox, QFileDialog, QSplitter
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
        self.df = None
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
        self.use_first_row_as_header = True
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

        # Controls layout
        controls_layout = QHBoxLayout()
        main_layout.addLayout(controls_layout)

        # Sheet options
        options_group = QGroupBox("Import Options")
        options_layout = QFormLayout()
        options_group.setLayout(options_layout)

        # Skip rows option
        self.skip_rows_spin = QSpinBox()
        self.skip_rows_spin.setRange(0, 100)
        self.skip_rows_spin.setValue(self.skip_rows)
        self.skip_rows_spin.valueChanged.connect(self._on_skip_rows_changed)
        options_layout.addRow("Skip Rows:", self.skip_rows_spin)

        # Use first row as header
        self.header_checkbox = QCheckBox("Use First Row as Headers")
        self.header_checkbox.setChecked(self.use_first_row_as_header)
        self.header_checkbox.toggled.connect(self._on_header_option_changed)
        options_layout.addRow("", self.header_checkbox)

        # Start row option
        self.start_row_spin = QSpinBox()
        self.start_row_spin.setRange(1, 1000)  # Example range, adjust as needed
        self.start_row_spin.setValue(1)
        self.start_row_spin.valueChanged.connect(self._on_start_row_changed)
        options_layout.addRow("Start Row:", self.start_row_spin)

        # End row option
        self.end_row_spin = QSpinBox()
        self.end_row_spin.setRange(1, 1000)  # Example range, adjust as needed
        self.end_row_spin.setValue(1000)
        self.end_row_spin.valueChanged.connect(self._on_end_row_changed)
        options_layout.addRow("End Row:", self.end_row_spin)

        controls_layout.addWidget(options_group)

        # Unit selection
        unit_group = QGroupBox("Unit")
        unit_layout = QFormLayout()
        unit_group.setLayout(unit_layout)

        self.unit_combo = QComboBox()
        self.unit_combo.addItem("mil", UnitType.MIL.value)
        self.unit_combo.addItem("mm", UnitType.MM.value)
        self.unit_combo.addItem("inch", UnitType.INCH.value)
        self.unit_combo.currentIndexChanged.connect(self._on_unit_changed)

        unit_layout.addRow("Unit:", self.unit_combo)
        controls_layout.addWidget(unit_group)

        # Add refresh button
        self.refresh_button = QPushButton("Refresh Preview")
        self.refresh_button.clicked.connect(self._load_data)
        controls_layout.addWidget(self.refresh_button)

        # Add stretcher to push controls to the left
        controls_layout.addStretch(1)

        # Table view
        self.table_view = QTableView()
        self.table_model = ExcelPreviewModel()
        self.table_view.setModel(self.table_model)

        # Configure table view
        self.table_view.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table_view.verticalHeader().setVisible(True)
        self.table_view.setAlternatingRowColors(True)

        main_layout.addWidget(self.table_view)

        # Dialog buttons
        self.button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)
        main_layout.addWidget(self.button_box)

        # Update the preview when start and stop positions are changed
        self.start_row_spin.valueChanged.connect(self._update_preview)
        self.end_row_spin.valueChanged.connect(self._update_preview)

    def _load_data(self):
        """Load data into the preview table"""
        try:
            # Make a copy of the original dataframe
            processed_df = self.df.copy()
            
            # Skip rows if needed
            if self.skip_rows > 0:
                processed_df = processed_df.iloc[self.skip_rows:].reset_index(drop=True)
            
            # Use first row as header if needed
            if self.use_first_row_as_header and len(processed_df) > 0:
                # Get the first row as headers
                headers = processed_df.iloc[0].values
                # Remove the first row (now used as headers)
                processed_df = processed_df.iloc[1:].reset_index(drop=True)
                # Set the column names
                processed_df.columns = headers
            
            # Update model
            self.table_model.set_dataframe(processed_df)
            
            # Resize columns for better visibility
            self.table_view.resizeColumnsToContents()
            self.table_view.resizeRowsToContents()
            
            # Force the table to update
            self.table_view.update()
            
            logger.info("Excel data loaded in preview")
        except Exception as e:
            logger.error(f"Error loading Excel data in preview: {str(e)}")
    
    def _on_skip_rows_changed(self, value):
        """Handle skip rows change"""
        self.skip_rows = value
        logger.info(f"Skip rows changed to: {value}")
        # Automatically refresh the preview
        self._load_data()
    
    def _on_header_option_changed(self, checked):
        """Handle header option change"""
        self.use_first_row_as_header = checked
        logger.info(f"Use first row as header changed to: {checked}")
        # Automatically refresh the preview
        self._load_data()
    
    def _on_start_row_changed(self, value):
        """Handle start row change"""
        self.start_row = value
        logger.info(f"Start row changed to: {value}")
        # We don't refresh here as this is used during import, not preview

    def _on_end_row_changed(self, value):
        """Handle end row change"""
        self.end_row = value
        logger.info(f"End row changed to: {value}")
        # We don't refresh here as this is used during import, not preview

    def _on_unit_changed(self, index):
        """Handle unit change"""
        unit_str = self.unit_combo.currentData()
        try:
            self.unit = UnitType(unit_str)
            logger.info(f"Unit changed to: {self.unit.value}")
        except ValueError:
            logger.error(f"Invalid unit: {unit_str}")

    def _update_preview(self):
        """Update the preview table based on start and stop rows."""
        try:
            # Adjust the dataframe based on start and stop rows
            start_row = self.start_row_spin.value() - 1  # Convert to 0-based index
            end_row = self.end_row_spin.value()

            # Ensure valid range
            if start_row < 0 or end_row > len(self.df) or start_row >= end_row:
                logger.warning("Invalid start or end row range for preview update.")
                return

            # Slice the dataframe
            preview_df = self.df.iloc[start_row:end_row]

            # Update the model with the sliced dataframe
            self.table_model.set_dataframe(preview_df)

            # Resize columns for better visibility
            self.table_view.resizeColumnsToContents()
            self.table_view.resizeRowsToContents()

            logger.info("Preview updated with selected row range.")
        except Exception as e:
            logger.error(f"Error updating preview: {str(e)}")
    
    def get_processed_dataframe(self) -> pd.DataFrame:
        """Get the processed dataframe"""
        return self.table_model.get_dataframe()
    
    def get_import_options(self) -> Dict[str, Any]:
        """Get the import options"""
        return {
            "unit": self.unit,
            "skip_rows": self.skip_rows,
            "use_first_row_as_header": self.use_first_row_as_header,
            "use_first_column_as_index": self.use_first_column_as_index
        }
