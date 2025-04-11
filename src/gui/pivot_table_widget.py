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
        
        if pivot_data and pivot_data.pivot_df is not None:
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
        
        # Convert input to proper type
        try:
            # Try to convert to float
            if value == "":
                # Empty string becomes NaN
                value = np.nan
            else:
                value = float(value)
            
            # Update the data array
            data_col = col - 1
            self.data_array[row, data_col] = value
            
            # Emit signal
            self.dataChanged.emit(index, index)
            self.data_changed.emit()
            return True
        except (ValueError, TypeError) as e:
            logger.warning(f"Invalid value for cell ({row}, {col}): {value}, error: {str(e)}")
            return False
    
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
    """Widget for displaying and editing pivot table data"""
    
    data_changed = pyqtSignal()
    
    def __init__(self, parent=None):
        """Initialize pivot table widget"""
        super().__init__(parent)
        self.pivot_data = None
        self.rule_type = RuleType.CLEARANCE
        self.unit = UnitType.MIL
        
        # Initialize UI
        self._init_ui()
        
        logger.info("Pivot table widget initialized")
    
    def _init_ui(self):
        """Initialize the UI components"""
        # Main layout
        main_layout = QVBoxLayout()
        self.setLayout(main_layout)
        
        # Controls layout
        controls_layout = QHBoxLayout()
        main_layout.addLayout(controls_layout)
        
        # Rule type selection
        rule_type_group = QGroupBox("Rule Type")
        rule_type_layout = QFormLayout()
        rule_type_group.setLayout(rule_type_layout)
        
        self.rule_type_combo = QComboBox()
        self.rule_type_combo.addItem("Electrical Clearance", RuleType.CLEARANCE.value)
        self.rule_type_combo.addItem("Short Circuit", RuleType.SHORT_CIRCUIT.value)
        self.rule_type_combo.addItem("Un-Routed Net", RuleType.UNROUTED_NET.value)
        self.rule_type_combo.addItem("Un-Connected Pin", RuleType.UNCONNECTED_PIN.value)
        self.rule_type_combo.addItem("Modified Polygon", RuleType.MODIFIED_POLYGON.value)
        self.rule_type_combo.addItem("Creepage Distance", RuleType.CREEPAGE_DISTANCE.value)
        self.rule_type_combo.currentIndexChanged.connect(self._on_rule_type_changed)
        
        rule_type_layout.addRow("Rule Type:", self.rule_type_combo)
        controls_layout.addWidget(rule_type_group)
        
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
        
        # Table actions
        actions_group = QGroupBox("Actions")
        actions_layout = QVBoxLayout()
        actions_group.setLayout(actions_layout)
        
        self.add_row_button = QPushButton("Add Net Class")
        self.add_row_button.clicked.connect(self._on_add_row)
        actions_layout.addWidget(self.add_row_button)
        
        self.remove_row_button = QPushButton("Remove Net Class")
        self.remove_row_button.clicked.connect(self._on_remove_row)
        actions_layout.addWidget(self.remove_row_button)
        
        controls_layout.addWidget(actions_group)
        
        # Add stretcher to push controls to the left
        controls_layout.addStretch(1)
        
        # Table view
        self.table_view = QTableView()
        self.table_model = PivotTableModel()
        self.table_view.setModel(self.table_model)
        
        # Configure table view
        self.table_view.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table_view.verticalHeader().setVisible(False)
        self.table_view.setAlternatingRowColors(True)
        
        # Connect signals
        self.table_model.data_changed.connect(self._on_data_changed)
        
        main_layout.addWidget(self.table_view)
    
    def set_pivot_data(self, pivot_data: ExcelPivotData):
        """Set the pivot data to display"""
        if pivot_data is None:
            logger.warning("Attempted to set None pivot data")
            return
        
        self.pivot_data = pivot_data
        
        if pivot_data:
            # Update rule type and unit combos
            if pivot_data.rule_type:
                index = self.rule_type_combo.findData(pivot_data.rule_type.value)
                if index >= 0:
                    self.rule_type_combo.setCurrentIndex(index)
            
            if pivot_data.unit:
                index = self.unit_combo.findData(pivot_data.unit.value)
                if index >= 0:
                    self.unit_combo.setCurrentIndex(index)
            
            # Set model data
            self.table_model.set_pivot_data(pivot_data)
            
            # Resize columns for better visibility
            self.table_view.resizeColumnsToContents()
            self.table_view.resizeRowsToContents()
            
            # Force the table to update
            self.table_view.update()
            
            logger.info("Pivot data set in widget")
            logger.debug(f"Pivot data matrix shape: {pivot_data.values.shape if pivot_data.values is not None else 'None'}")
        else:
            self.table_model.set_pivot_data(None)
            logger.warning("Empty pivot data provided")
    
    def get_pivot_data(self) -> ExcelPivotData:
        """Get the current pivot data from the model"""
        return self.table_model.get_updated_pivot_data()
    
    def _on_rule_type_changed(self, index):
        """Handle rule type change"""
        rule_type_str = self.rule_type_combo.currentData()
        try:
            self.rule_type = RuleType(rule_type_str)
            
            if self.pivot_data:
                self.pivot_data.rule_type = self.rule_type
                logger.info(f"Rule type changed to: {self.rule_type.value}")
                self.data_changed.emit()
        except ValueError:
            logger.error(f"Invalid rule type: {rule_type_str}")
    
    def _on_unit_changed(self, index):
        """Handle unit change"""
        unit_str = self.unit_combo.currentData()
        try:
            self.unit = UnitType(unit_str)
            
            if self.pivot_data:
                # TODO: Convert values if needed
                self.pivot_data.unit = self.unit
                logger.info(f"Unit changed to: {self.unit.value}")
                self.data_changed.emit()
        except ValueError:
            logger.error(f"Invalid unit: {unit_str}")
    
    def _on_data_changed(self):
        """Handle data changes in the model"""
        # Update pivot data with changes from model
        self.pivot_data = self.table_model.get_updated_pivot_data()
        
        # Emit signal
        self.data_changed.emit()
        logger.debug("Pivot data changed in widget")
    
    def _on_add_row(self):
        """Add a new row (net class) to the table"""
        # TODO: Implement add row functionality
        QMessageBox.information(self, "Not Implemented", "Add row functionality not yet implemented")
    
    def _on_remove_row(self):
        """Remove selected row (net class) from the table"""
        # TODO: Implement remove row functionality
        QMessageBox.information(self, "Not Implemented", "Remove row functionality not yet implemented")
