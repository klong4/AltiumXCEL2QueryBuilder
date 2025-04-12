#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Rule Editor Widget
================

Custom widget for displaying and editing Altium rule properties.
"""

import logging
from typing import Dict, List, Optional, Union, Tuple

from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QFormLayout, QLabel,
    QLineEdit, QComboBox, QSpinBox, QDoubleSpinBox, QCheckBox,
    QGroupBox, QPushButton, QTableView, QHeaderView, QMessageBox
)
from PyQt5.QtCore import Qt, QAbstractTableModel, QModelIndex, QVariant, pyqtSignal
from PyQt5.QtGui import QColor, QBrush

from models.rule_model import (
    UnitType, RuleType, RuleScope, BaseRule,
    ClearanceRule, ShortCircuitRule, UnRoutedNetRule,
    # Add other specific rule types if they exist, e.g.:
    # PowerPlaneConnectRule, PolygonConnectRule, etc.
)

logger = logging.getLogger(__name__)

class RuleEditorWidget(QWidget):
    """Base widget for editing rules"""
    
    rule_changed = pyqtSignal(BaseRule)
    
    def __init__(self, parent=None):
        """Initialize rule editor widget"""
        super().__init__(parent)
        self.rule = None
        
        # Initialize UI
        self._init_ui()
        
        logger.info("Rule editor widget initialized")
    
    def _init_ui(self):
        """Initialize the UI components"""
        # Main layout
        self.main_layout = QVBoxLayout()
        self.setLayout(self.main_layout)
        
        # Common properties group
        common_group = QGroupBox("Rule Properties")
        common_layout = QFormLayout()
        common_group.setLayout(common_layout)
        
        # Rule name
        self.name_edit = QLineEdit()
        self.name_edit.textChanged.connect(self._on_property_changed)
        common_layout.addRow("Rule Name:", self.name_edit)
        
        # Rule enabled
        self.enabled_checkbox = QCheckBox("Enabled")
        self.enabled_checkbox.setChecked(True)
        self.enabled_checkbox.stateChanged.connect(self._on_property_changed)
        common_layout.addRow("", self.enabled_checkbox)
        
        # Rule priority
        self.priority_spin = QSpinBox()
        self.priority_spin.setRange(1, 100)
        self.priority_spin.setValue(1)
        self.priority_spin.valueChanged.connect(self._on_property_changed)
        common_layout.addRow("Priority:", self.priority_spin)
        
        # Rule comment
        self.comment_edit = QLineEdit()
        self.comment_edit.textChanged.connect(self._on_property_changed)
        common_layout.addRow("Comment:", self.comment_edit)

        self.main_layout.addWidget(common_group)

        # GroupBox for rule-specific properties (layout cleared/populated dynamically)
        self.rule_specific_group = QGroupBox("Rule Specific Properties")
        self.rule_specific_layout = QFormLayout() # Use QFormLayout for consistency
        self.rule_specific_group.setLayout(self.rule_specific_layout)
        self.main_layout.addWidget(self.rule_specific_group)

        # Add stretch to push widgets to the top
        self.main_layout.addStretch(1)

    def set_rule(self, rule: BaseRule):
        """Set the rule to edit"""
        self.rule = rule
        
        if rule:
            # Update common properties
            self.name_edit.setText(rule.name)
            self.enabled_checkbox.setChecked(rule.enabled)
            self.priority_spin.setValue(rule.priority)
            self.comment_edit.setText(rule.comment)
            
            # Update rule-specific properties
            self._update_rule_specific_properties()

            # Ensure the specific properties group is visible AFTER updating
            self.rule_specific_group.setVisible(True) # <-- Add this line

            logger.info(f"Rule set in editor: {rule.name} ({rule.rule_type.value})")
        else:
            # Clear fields if no rule is selected
            self.name_edit.clear()
            self.enabled_checkbox.setChecked(True)
            self.priority_spin.setValue(1)
            self.comment_edit.clear()
            self._clear_layout(self.rule_specific_layout)
            self.rule_specific_group.setVisible(False)

    def _clear_layout(self, layout):
        """Removes all widgets from a layout."""
        if layout is not None:
            while layout.count():
                item = layout.takeAt(0)
                widget = item.widget()
                if widget is not None:
                    widget.deleteLater()
                else:
                    # Recursively clear nested layouts if necessary
                    sub_layout = item.layout()
                    if sub_layout is not None:
                        self._clear_layout(sub_layout)

    def _update_rule_specific_properties(self):
        """Dynamically populate the rule-specific properties section based on rule type.
        NOTE: This is primarily a fallback for unknown rule types where a specific
        editor subclass wasn't created. Subclasses override this method.
        """
        # If this instance is actually a specific subclass, let its override handle this.
        if type(self) is not RuleEditorWidget:
            return

        self._clear_layout(self.rule_specific_layout)
        self.rule_specific_group.setVisible(True)

        if not self.rule:
            self.rule_specific_group.setVisible(False)
            return

        # --- Clearance Rule --- 
        if isinstance(self.rule, ClearanceRule):
            self.rule_specific_group.setTitle("Clearance Rule Properties")

            # Min Clearance
            self.min_clearance_spin = QDoubleSpinBox()
            self.min_clearance_spin.setSuffix(f" {self.rule.unit.value}")
            self.min_clearance_spin.setDecimals(3) # Adjust precision as needed
            self.min_clearance_spin.setRange(0, 10000) # Set appropriate range
            self.min_clearance_spin.setValue(self.rule.min_clearance)
            self.min_clearance_spin.valueChanged.connect(self._on_property_changed)
            self.rule_specific_layout.addRow("Minimum Clearance:", self.min_clearance_spin)

            # Scope 1
            self.scope1_combo = QComboBox()
            self.scope1_combo.addItems([scope.value for scope in RuleScope])
            self.scope1_combo.setCurrentText(self.rule.scope1.value)
            self.scope1_combo.currentTextChanged.connect(self._on_property_changed)
            self.rule_specific_layout.addRow("Scope 1:", self.scope1_combo)

            # Scope 2
            self.scope2_combo = QComboBox()
            self.scope2_combo.addItems([scope.value for scope in RuleScope])
            self.scope2_combo.setCurrentText(self.rule.scope2.value)
            self.scope2_combo.currentTextChanged.connect(self._on_property_changed)
            self.rule_specific_layout.addRow("Scope 2:", self.scope2_combo)

        # --- Short Circuit Rule --- 
        elif isinstance(self.rule, ShortCircuitRule):
            self.rule_specific_group.setTitle("Short Circuit Rule Properties")
            # Short circuit rules often just need scope (handled by common props?) or might have specific flags
            # Add specific widgets if ShortCircuitRule has unique properties
            # Example: self.allow_short_circuit_checkbox = QCheckBox("Allow Short Circuit")
            # self.allow_short_circuit_checkbox.setChecked(self.rule.allow_short_circuit) # Assuming property exists
            # self.allow_short_circuit_checkbox.stateChanged.connect(self._on_property_changed)
            # self.rule_specific_layout.addRow("Options:", self.allow_short_circuit_checkbox)
            pass # No specific properties defined in model yet

        # --- Un-Routed Net Rule --- 
        elif isinstance(self.rule, UnRoutedNetRule):
            self.rule_specific_group.setTitle("Un-Routed Net Rule Properties")
            # Un-routed net rules typically just need scope (handled by common props?)
            # Add specific widgets if UnRoutedNetRule has unique properties
            pass # No specific properties defined in model yet

        # --- Add other rule types here --- 
        # elif isinstance(self.rule, PowerPlaneConnectRule):
        #    # ... add widgets for PowerPlaneConnectRule ...
        # elif isinstance(self.rule, PolygonConnectRule):
        #    # ... add widgets for PolygonConnectRule ...

        else:
            # Handle unknown or base rule type (maybe show nothing specific)
            self.rule_specific_group.setTitle("Rule Specific Properties (N/A)")
            self.rule_specific_group.setVisible(False)

    def _on_property_changed(self, value=None):
        """Update the internal rule object when a property changes in the UI."""
        if not self.rule:
            return

        sender = self.sender()

        try:
            # Update common properties
            if sender == self.name_edit:
                self.rule.name = self.name_edit.text()
            elif sender == self.enabled_checkbox:
                self.rule.enabled = self.enabled_checkbox.isChecked()
            elif sender == self.priority_spin:
                self.rule.priority = self.priority_spin.value()
            elif sender == self.comment_edit:
                self.rule.comment = self.comment_edit.text()

            # Emit signal that the rule has changed
            self.rule_changed.emit(self.rule)
            logger.debug(f"Rule property changed (Common): {self.rule.name}")

        except Exception as e:
            logger.error(f"Error updating rule property: {e}", exc_info=True)
            # Optionally show an error to the user
            # QMessageBox.warning(self, "Error", f"Failed to update property: {e}")


class ClearanceRuleEditor(RuleEditorWidget):
    """Widget for editing clearance rules"""
    
    def __init__(self, parent=None):
        """Initialize clearance rule editor widget"""
        super().__init__(parent)
        
        # Initialize rule-specific UI components
        self._init_rule_specific_ui()
    
    def _init_rule_specific_ui(self):
        """Initialize the rule-specific UI components"""
        # Clear existing components
        while self.rule_specific_layout.count():
            item = self.rule_specific_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        
        # Update group title
        self.rule_specific_group.setTitle("Clearance Rule Properties")
        
        # Minimum clearance
        self.clearance_spin = QDoubleSpinBox()
        self.clearance_spin.setRange(0.001, 10000.0)
        self.clearance_spin.setValue(10.0)
        self.clearance_spin.setDecimals(3)
        self.clearance_spin.valueChanged.connect(self._on_property_changed)
        self.rule_specific_layout.addRow("Minimum Clearance:", self.clearance_spin)
        
        # Unit
        self.unit_combo = QComboBox()
        self.unit_combo.addItem("mil", UnitType.MIL.value)
        self.unit_combo.addItem("mm", UnitType.MM.value)
        self.unit_combo.addItem("inch", UnitType.INCH.value)
        self.unit_combo.currentIndexChanged.connect(self._on_property_changed)
        self.rule_specific_layout.addRow("Unit:", self.unit_combo)
        
        # Source scope
        self.source_scope_group = QGroupBox("Source Scope")
        source_scope_layout = QFormLayout()
        self.source_scope_group.setLayout(source_scope_layout)
        
        self.source_scope_combo = QComboBox()
        self.source_scope_combo.addItem("All", "All")
        self.source_scope_combo.addItem("Net Class", "NetClass")
        self.source_scope_combo.addItem("Net Classes", "NetClasses")
        self.source_scope_combo.addItem("Custom", "Custom")
        self.source_scope_combo.currentIndexChanged.connect(self._on_source_scope_type_changed)
        source_scope_layout.addRow("Type:", self.source_scope_combo)
        
        self.source_scope_edit = QLineEdit()
        self.source_scope_edit.setEnabled(False)
        self.source_scope_edit.textChanged.connect(self._on_property_changed)
        source_scope_layout.addRow("Items:", self.source_scope_edit)
        
        self.rule_specific_layout.addRow("", self.source_scope_group)
        
        # Target scope
        self.target_scope_group = QGroupBox("Target Scope")
        target_scope_layout = QFormLayout()
        self.target_scope_group.setLayout(target_scope_layout)
        
        self.target_scope_combo = QComboBox()
        self.target_scope_combo.addItem("All", "All")
        self.target_scope_combo.addItem("Net Class", "NetClass")
        self.target_scope_combo.addItem("Net Classes", "NetClasses")
        self.target_scope_combo.addItem("Custom", "Custom")
        self.target_scope_combo.currentIndexChanged.connect(self._on_target_scope_type_changed)
        target_scope_layout.addRow("Type:", self.target_scope_combo)
        
        self.target_scope_edit = QLineEdit()
        self.target_scope_edit.setEnabled(False)
        self.target_scope_edit.textChanged.connect(self._on_property_changed)
        target_scope_layout.addRow("Items:", self.target_scope_edit)
        
        self.rule_specific_layout.addRow("", self.target_scope_group)
    
    def _update_rule_specific_properties(self):
        """Update UI with rule-specific properties"""
        if not isinstance(self.rule, ClearanceRule):
            logger.error(f"Expected ClearanceRule, got {type(self.rule).__name__}")
            return
        
        # Update minimum clearance
        self.clearance_spin.setValue(self.rule.min_clearance)
        
        # Update unit
        index = self.unit_combo.findData(self.rule.unit.value)
        if index >= 0:
            self.unit_combo.setCurrentIndex(index)
        
        # Update source scope
        source_scope_type = self.rule.source_scope.scope_type
        source_index = self.source_scope_combo.findData(source_scope_type)
        if source_index >= 0:
            self.source_scope_combo.setCurrentIndex(source_index)
        
        if source_scope_type != "All":
            self.source_scope_edit.setEnabled(True)
            self.source_scope_edit.setText(";".join(self.rule.source_scope.items))
        else:
            self.source_scope_edit.setEnabled(False)
            self.source_scope_edit.clear()
        
        # Update target scope
        target_scope_type = self.rule.target_scope.scope_type
        target_index = self.target_scope_combo.findData(target_scope_type)
        if target_index >= 0:
            self.target_scope_combo.setCurrentIndex(target_index)
        
        if target_scope_type != "All":
            self.target_scope_edit.setEnabled(True)
            self.target_scope_edit.setText(";".join(self.rule.target_scope.items))
        else:
            self.target_scope_edit.setEnabled(False)
            self.target_scope_edit.clear()

    def _on_property_changed(self, value=None):
        """Handle property changes from UI elements."""
        if not self.rule or not isinstance(self.rule, ClearanceRule):
            # Let base class handle if rule type is wrong or null
            super()._on_property_changed(value)
            return

        sender = self.sender()
        rule_modified = False

        # Handle specific properties
        if sender == self.clearance_spin:
            if self.rule.min_clearance != self.clearance_spin.value():
                self.rule.min_clearance = self.clearance_spin.value()
                rule_modified = True
        elif sender == self.unit_combo:
            unit_val = self.unit_combo.currentData()
            if self.rule.unit.value != unit_val:
                try:
                    self.rule.unit = UnitType(unit_val)
                    rule_modified = True
                except ValueError:
                    logger.error(f"Invalid unit selected: {unit_val}")
        elif sender == self.source_scope_combo or sender == self.source_scope_edit:
            scope_type = self.source_scope_combo.currentData()
            items = []
            if scope_type != "All":
                items_str = self.source_scope_edit.text()
                items = items_str.split(";") if items_str else []
            new_scope = RuleScope(scope_type, items)
            # Use __dict__ comparison for RuleScope if __eq__ is not defined
            if self.rule.source_scope.__dict__ != new_scope.__dict__:
                self.rule.source_scope = new_scope
                rule_modified = True
        elif sender == self.target_scope_combo or sender == self.target_scope_edit:
            scope_type = self.target_scope_combo.currentData()
            items = []
            if scope_type != "All":
                items_str = self.target_scope_edit.text()
                items = items_str.split(";") if items_str else []
            new_scope = RuleScope(scope_type, items)
            # Use __dict__ comparison for RuleScope if __eq__ is not defined
            if self.rule.target_scope.__dict__ != new_scope.__dict__:
                self.rule.target_scope = new_scope
                rule_modified = True
        else:
            # Not a specific property, let the base class handle common properties
            super()._on_property_changed(value)
            return # Base class emits signal

        # If a specific property was modified, emit the signal
        if rule_modified:
            logger.debug(f"Rule property changed (Clearance Specific): {self.rule.name}")
            self.rule_changed.emit(self.rule)

    def _on_source_scope_type_changed(self, index):
        """Handle source scope type changes"""
        scope_type = self.source_scope_combo.currentData()
        
        if scope_type == "All":
            self.source_scope_edit.setEnabled(False)
            self.source_scope_edit.clear()
        else:
            self.source_scope_edit.setEnabled(True)
        
        self._on_property_changed()
    
    def _on_target_scope_type_changed(self, index):
        """Handle target scope type changes"""
        scope_type = self.target_scope_combo.currentData()
        
        if scope_type == "All":
            self.target_scope_edit.setEnabled(False)
            self.target_scope_edit.clear()
        else:
            self.target_scope_edit.setEnabled(True)
        
        self._on_property_changed()


class ShortCircuitRuleEditor(RuleEditorWidget):
    """Widget for editing short circuit rules"""
    
    def __init__(self, parent=None):
        """Initialize short circuit rule editor widget"""
        super().__init__(parent)
        
        # Initialize rule-specific UI components
        self._init_rule_specific_ui()
    
    def _init_rule_specific_ui(self):
        """Initialize the rule-specific UI components"""
        # Clear existing components
        while self.rule_specific_layout.count():
            item = self.rule_specific_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        
        # Update group title
        self.rule_specific_group.setTitle("Short Circuit Rule Properties")
        
        # Scope
        self.scope_group = QGroupBox("Scope")
        scope_layout = QFormLayout()
        self.scope_group.setLayout(scope_layout)
        
        self.scope_combo = QComboBox()
        self.scope_combo.addItem("All", "All")
        self.scope_combo.addItem("Net Class", "NetClass")
        self.scope_combo.addItem("Net Classes", "NetClasses")
        self.scope_combo.addItem("Custom", "Custom")
        self.scope_combo.currentIndexChanged.connect(self._on_scope_type_changed)
        scope_layout.addRow("Type:", self.scope_combo)
        
        self.scope_edit = QLineEdit()
        self.scope_edit.setEnabled(False)
        self.scope_edit.textChanged.connect(self._on_property_changed)
        scope_layout.addRow("Items:", self.scope_edit)
        
        self.rule_specific_layout.addRow("", self.scope_group)
    
    def _update_rule_specific_properties(self):
        """Update UI with rule-specific properties"""
        if not isinstance(self.rule, ShortCircuitRule):
            logger.error(f"Expected ShortCircuitRule, got {type(self.rule).__name__}")
            return
        
        # Update scope
        scope_type = self.rule.scope.scope_type
        scope_index = self.scope_combo.findData(scope_type)
        if scope_index >= 0:
            self.scope_combo.setCurrentIndex(scope_index)
        
        if scope_type != "All":
            self.scope_edit.setEnabled(True)
            self.scope_edit.setText(";".join(self.rule.scope.items))
        else:
            self.scope_edit.setEnabled(False)
            self.scope_edit.clear()
    
    def _on_property_changed(self, value=None):
        """Handle property changes from UI elements."""
        if not self.rule or not isinstance(self.rule, ShortCircuitRule):
             # Let base class handle if rule type is wrong or null
            super()._on_property_changed(value)
            return

        sender = self.sender()
        rule_modified = False

        # Handle specific properties (Scope)
        if sender == self.scope_combo or sender == self.scope_edit:
            scope_type = self.scope_combo.currentData()
            items = []
            if scope_type != "All":
                items_str = self.scope_edit.text()
                items = items_str.split(";") if items_str else []
            new_scope = RuleScope(scope_type, items)
            # Use __dict__ comparison for RuleScope if __eq__ is not defined
            if self.rule.scope.__dict__ != new_scope.__dict__:
                self.rule.scope = new_scope
                rule_modified = True
        else:
            # Not a specific property, let the base class handle common properties
            super()._on_property_changed(value)
            return # Base class emits signal

        # If a specific property was modified, emit the signal
        if rule_modified:
            logger.debug(f"Rule property changed (ShortCircuit Specific): {self.rule.name}")
            self.rule_changed.emit(self.rule)

    def _on_scope_type_changed(self, index):
        """Handle scope type changes"""
        scope_type = self.scope_combo.currentData()
        
        if scope_type == "All":
            self.scope_edit.setEnabled(False)
            self.scope_edit.clear()
        else:
            self.scope_edit.setEnabled(True)
        
        self._on_property_changed()


class UnRoutedNetRuleEditor(RuleEditorWidget):
    """Widget for editing unrouted net rules"""
    
    def __init__(self, parent=None):
        """Initialize unrouted net rule editor widget"""
        super().__init__(parent)
        
        # Initialize rule-specific UI components
        self._init_rule_specific_ui()
    
    def _init_rule_specific_ui(self):
        """Initialize the rule-specific UI components"""
        # Clear existing components
        while self.rule_specific_layout.count():
            item = self.rule_specific_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        
        # Update group title
        self.rule_specific_group.setTitle("Unrouted Net Rule Properties")
        
        # Scope
        self.scope_group = QGroupBox("Scope")
        scope_layout = QFormLayout()
        self.scope_group.setLayout(scope_layout)
        
        self.scope_combo = QComboBox()
        self.scope_combo.addItem("All", "All")
        self.scope_combo.addItem("Net Class", "NetClass")
        self.scope_combo.addItem("Net Classes", "NetClasses")
        self.scope_combo.addItem("Custom", "Custom")
        self.scope_combo.currentIndexChanged.connect(self._on_scope_type_changed)
        scope_layout.addRow("Type:", self.scope_combo)
        
        self.scope_edit = QLineEdit()
        self.scope_edit.setEnabled(False)
        self.scope_edit.textChanged.connect(self._on_property_changed)
        scope_layout.addRow("Items:", self.scope_edit)
        
        self.rule_specific_layout.addRow("", self.scope_group)
    
    def _update_rule_specific_properties(self):
        """Update UI with rule-specific properties"""
        if not isinstance(self.rule, UnRoutedNetRule):
            logger.error(f"Expected UnRoutedNetRule, got {type(self.rule).__name__}")
            return
        
        # Update scope
        scope_type = self.rule.scope.scope_type
        scope_index = self.scope_combo.findData(scope_type)
        if scope_index >= 0:
            self.scope_combo.setCurrentIndex(scope_index)
        
        if scope_type != "All":
            self.scope_edit.setEnabled(True)
            self.scope_edit.setText(";".join(self.rule.scope.items))
        else:
            self.scope_edit.setEnabled(False)
            self.scope_edit.clear()

    def _on_property_changed(self, value=None):
        """Handle property changes from UI elements."""
        if not self.rule or not isinstance(self.rule, UnRoutedNetRule):
             # Let base class handle if rule type is wrong or null
            super()._on_property_changed(value)
            return

        sender = self.sender()
        rule_modified = False

        # Handle specific properties (Scope)
        if sender == self.scope_combo or sender == self.scope_edit:
            scope_type = self.scope_combo.currentData()
            items = []
            if scope_type != "All":
                items_str = self.scope_edit.text()
                items = items_str.split(";") if items_str else []
            new_scope = RuleScope(scope_type, items)
            # Use __dict__ comparison for RuleScope if __eq__ is not defined
            if self.rule.scope.__dict__ != new_scope.__dict__:
                self.rule.scope = new_scope
                rule_modified = True
        else:
            # Not a specific property, let the base class handle common properties
            super()._on_property_changed(value)
            return # Base class emits signal

        # If a specific property was modified, emit the signal
        if rule_modified:
            logger.debug(f"Rule property changed (UnRoutedNet Specific): {self.rule.name}")
            self.rule_changed.emit(self.rule)

    def _on_scope_type_changed(self, index):
        """Handle scope type changes"""
        scope_type = self.scope_combo.currentData()
        
        if scope_type == "All":
            self.scope_edit.setEnabled(False)
            self.scope_edit.clear()
        else:
            self.scope_edit.setEnabled(True)
        
        self._on_property_changed()


class RuleTableModel(QAbstractTableModel):
    """Model for displaying rules in a table"""
    
    def __init__(self, parent=None):
        """Initialize rule table model"""
        super().__init__(parent)
        self.rules = []
        self.column_headers = ["Name", "Type", "Enabled", "Priority"]
    
    def set_rules(self, rules: List[BaseRule]):
        """Set the rules to display"""
        self.beginResetModel()
        self.rules = rules
        self.endResetModel()
        logger.info(f"Rule table model updated with {len(rules)} rules")
    
    def rowCount(self, parent=None):
        """Return number of rows in the model"""
        return len(self.rules)
    
    def columnCount(self, parent=None):
        """Return number of columns in the model"""
        return len(self.column_headers)
    
    def data(self, index, role=Qt.DisplayRole):
        """Return data for the given index and role"""
        if not index.isValid() or not self.rules:
            return QVariant()
        
        row, col = index.row(), index.column()
        
        # Check if row and column are valid
        if row < 0 or row >= self.rowCount() or col < 0 or col >= self.columnCount():
            return QVariant()
        
        rule = self.rules[row]
        
        # Handle display role
        if role == Qt.DisplayRole:
            if col == 0:
                return rule.name
            elif col == 1:
                return rule.rule_type.value
            elif col == 2:
                return "Yes" if rule.enabled else "No"
            elif col == 3:
                return str(rule.priority)
        
        # Handle background color role - alternate row colors
        elif role == Qt.BackgroundRole:
            if row % 2 == 0:
                return QBrush(QColor("#323232"))
            else:
                return QBrush(QColor("#2d2d2d"))
        
        # Handle text alignment role
        elif role == Qt.TextAlignmentRole:
            return Qt.AlignCenter
        
        return QVariant()
    
    def headerData(self, section, orientation, role=Qt.DisplayRole):
        """Return header data for the given section and orientation"""
        if role != Qt.DisplayRole:
            return QVariant()
        
        if orientation == Qt.Horizontal and section < len(self.column_headers):
            return self.column_headers[section]
        
        return QVariant()


class RulesManagerWidget(QWidget):
    """Widget for managing multiple rules"""
    
    rules_changed = pyqtSignal()
    pivot_data_updated = pyqtSignal(object)  # Emits ExcelPivotData
    
    def __init__(self, parent=None):
        """Initialize rules manager widget"""
        super().__init__(parent)
        self.rule_manager = None
        self.current_rule = None
        self.current_editor = None
        
        # Initialize UI
        self._init_ui()
        
        # Create empty rule manager
        self.set_rule_manager(None)
        
        logger.info("Rules manager widget initialized")

    def add_rules(self, new_rules: List[BaseRule]):
        """Appends new rules to the existing list in the table model."""
        if not self.rule_model: # Check rule_model instead of table_model
            logger.warning("Rule table model not initialized. Cannot add rules.")
            return

        current_rules = self.rule_model.rules # Get rules from rule_model
        combined_rules = current_rules + new_rules
        self.rule_model.set_rules(combined_rules) # Set rules on rule_model
        logger.info(f"Appended {len(new_rules)} rules. Total rules: {len(combined_rules)}")

    def _init_ui(self):
        """Initialize the UI components"""
        # Main layout
        main_layout = QHBoxLayout()
        self.setLayout(main_layout)
        
        # Left panel - rule list
        left_panel = QVBoxLayout()
        
        # Rule type selection for new rules
        rule_type_group = QGroupBox("New Rule Type")
        rule_type_layout = QFormLayout()
        rule_type_group.setLayout(rule_type_layout)
        
        self.rule_type_combo = QComboBox()
        self.rule_type_combo.addItem("Electrical Clearance", RuleType.CLEARANCE.value)
        self.rule_type_combo.addItem("Short Circuit", RuleType.SHORT_CIRCUIT.value)
        self.rule_type_combo.addItem("Unrouted Net", RuleType.UNROUTED_NET.value)
        rule_type_layout.addRow("Rule Type:", self.rule_type_combo)
        
        # Add button for new rules
        self.add_rule_button = QPushButton("Add New Rule")
        self.add_rule_button.clicked.connect(self._on_add_rule)
        rule_type_layout.addRow("", self.add_rule_button)
        
        left_panel.addWidget(rule_type_group)
        
        # Rule list
        list_group = QGroupBox("Rules")
        list_layout = QVBoxLayout()
        list_group.setLayout(list_layout)
        
        # Rule table
        self.rule_table = QTableView()
        self.rule_model = RuleTableModel()
        self.rule_table.setModel(self.rule_model)
        
        # Configure table
        self.rule_table.setSelectionBehavior(QTableView.SelectRows)
        self.rule_table.setSelectionMode(QTableView.SingleSelection)
        self.rule_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.rule_table.verticalHeader().setVisible(False)
        self.rule_table.setAlternatingRowColors(True)
        
        # Connect selection change
        self.rule_table.selectionModel().selectionChanged.connect(self._on_rule_selection_changed)
        
        list_layout.addWidget(self.rule_table)
        
        # Buttons for rule management
        button_layout = QHBoxLayout()
        
        self.delete_rule_button = QPushButton("Delete Rule")
        self.delete_rule_button.clicked.connect(self._on_delete_rule)
        self.delete_rule_button.setEnabled(False)
        button_layout.addWidget(self.delete_rule_button)
        
        self.to_pivot_button = QPushButton("Update Pivot")
        self.to_pivot_button.clicked.connect(self._on_update_pivot)
        button_layout.addWidget(self.to_pivot_button)
        
        list_layout.addLayout(button_layout)
        
        left_panel.addWidget(list_group)
        main_layout.addLayout(left_panel, 1)
        
        # Right panel - rule editor
        self.rule_editor_container = QVBoxLayout()
        
        # Default message
        self.no_rule_label = QLabel("Select a rule to edit or add a new rule.")
        self.no_rule_label.setAlignment(Qt.AlignCenter)
        self.rule_editor_container.addWidget(self.no_rule_label)
        
        main_layout.addLayout(self.rule_editor_container, 2)
    
    def set_rule_manager(self, rule_manager):
        """Set the rule manager to use"""
        if rule_manager is None:
            from models.rule_model import RuleManager
            self.rule_manager = RuleManager()
        else:
            self.rule_manager = rule_manager
        
        # Update rule table
        self.rule_model.set_rules(self.rule_manager.rules)
        
        # Clear current rule selection
        self.current_rule = None
        self._update_rule_editor()
        
        logger.info(f"Rule manager set with {len(self.rule_manager.rules)} rules")
    
    def get_rule_manager(self):
        """Get the current rule manager"""
        return self.rule_manager
    
    def _update_rule_editor(self):
        """Update the rule editor based on the current rule"""
        # Clear the editor container
        while self.rule_editor_container.count():
            item = self.rule_editor_container.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        
        if self.current_rule:
            # Create appropriate editor for the rule type
            self.current_editor = create_rule_editor(self.current_rule.rule_type, self)
            self.current_editor.set_rule(self.current_rule)
            self.current_editor.rule_changed.connect(self._on_rule_changed)
            self.rule_editor_container.addWidget(self.current_editor)
            
            # Enable delete button
            self.delete_rule_button.setEnabled(True)
            
            logger.info(f"Editing rule: {self.current_rule.name}")
        else:
            # Show default message
            self.no_rule_label = QLabel("Select a rule to edit or add a new rule.")
            self.no_rule_label.setAlignment(Qt.AlignCenter)
            self.rule_editor_container.addWidget(self.no_rule_label)
            
            # Disable delete button
            self.delete_rule_button.setEnabled(False)
    
    def _on_rule_selection_changed(self, selected, deselected):
        """Handle rule selection changes"""
        indexes = selected.indexes()
        
        if indexes:
            # Get the selected row
            row = indexes[0].row()
            
            # Set the current rule
            if 0 <= row < len(self.rule_manager.rules):
                self.current_rule = self.rule_manager.rules[row]
                self._update_rule_editor()
    
    def _on_add_rule(self):
        """Handle add rule button"""
        # Get the selected rule type
        rule_type_str = self.rule_type_combo.currentData()
        try:
            rule_type = RuleType(rule_type_str)
        except ValueError:
            logger.error(f"Invalid rule type: {rule_type_str}")
            return
        
        # Create a new rule based on the type
        rule = None
        name = f"New_{rule_type.value}_Rule_{len(self.rule_manager.rules) + 1}"
        
        if rule_type == RuleType.CLEARANCE:
            rule = ClearanceRule(name=name)
        elif rule_type == RuleType.SHORT_CIRCUIT:
            rule = ShortCircuitRule(name=name)
        elif rule_type == RuleType.UNROUTED_NET:
            rule = UnRoutedNetRule(name=name)
        else:
            logger.error(f"Unsupported rule type: {rule_type.value}")
            return
        
        # Add the rule to the manager
        self.rule_manager.add_rule(rule)
        
        # Update rule table
        self.rule_model.set_rules(self.rule_manager.rules)
        
        # Select the new rule
        self.current_rule = rule
        self._update_rule_editor()
        
        # Select the new rule in the table
        row = self.rule_manager.rules.index(rule)
        self.rule_table.selectRow(row)
        
        # Emit rules changed signal
        self.rules_changed.emit()
        
        logger.info(f"Added new rule: {rule.name} ({rule.rule_type.value})")
    
    def _on_delete_rule(self):
        """Handle delete rule button"""
        if not self.current_rule:
            return
        
        # Confirm deletion
        reply = QMessageBox.question(
            self,
            "Confirm Deletion",
            f"Are you sure you want to delete the rule '{self.current_rule.name}'?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if reply != QMessageBox.Yes:
            return
        
        # Get the index of the current rule
        current_index = self.rule_manager.rules.index(self.current_rule)
        
        # Remove the rule from the manager
        self.rule_manager.rules.remove(self.current_rule)
        
        # Update rule table
        self.rule_model.set_rules(self.rule_manager.rules)
        
        # Select a new rule if available
        self.current_rule = None
        if self.rule_manager.rules:
            # Select the same index if possible, otherwise the last rule
            new_index = min(current_index, len(self.rule_manager.rules) - 1)
            self.current_rule = self.rule_manager.rules[new_index]
            self.rule_table.selectRow(new_index)
        
        # Update the editor
        self._update_rule_editor()
        
        # Emit rules changed signal
        self.rules_changed.emit()
        
        logger.info(f"Deleted rule at index {current_index}")
    
    def _on_rule_changed(self, rule):
        """Handle when a rule is modified in the editor"""
        if not rule:
            return
        
        # Ensure the rule is in the manager
        if rule not in self.rule_manager.rules:
            return
        
        # Update the rule table model to reflect changes
        self.rule_model.set_rules(self.rule_manager.rules)
        
        # Emit rules changed signal
        self.rules_changed.emit()
        
        logger.debug(f"Rule updated: {rule.name}")
    
    def _on_update_pivot(self):
        """Convert the rules to pivot data and emit the pivot_data_updated signal"""
        try:
            # Check if we have any clearance rules to convert
            clearance_rules = []
            for rule in self.rule_manager.rules:
                if rule.rule_type == RuleType.CLEARANCE:
                    clearance_rules.append(rule)
            
            if not clearance_rules:
                QMessageBox.warning(
                    self,
                    "No Clearance Rules",
                    "No clearance rules found to convert to pivot data.\n"
                    "Only clearance rules can be converted to pivot table format."
                )
                return
            
            # Import the ExcelPivotData class
            from models.excel_data import ExcelPivotData
            
            # Create pivot data from clearance rules
            pivot_data = ExcelPivotData.from_clearance_rules(clearance_rules)
            
            if not pivot_data:
                QMessageBox.warning(
                    self,
                    "Conversion Failed",
                    "Failed to convert rules to pivot data. Check the log for details."
                )
                return
            
            # Emit the pivot data updated signal with the new pivot data
            self.pivot_data_updated.emit(pivot_data)
            
            # Show success message
            QMessageBox.information(
                self,
                "Success",
                f"Successfully converted {len(clearance_rules)} clearance rules to pivot data.\n"
                "The pivot table has been updated."
            )
            
            logger.info(f"Converted {len(clearance_rules)} rules to pivot data")
        
        except Exception as e:
            error_msg = f"Error converting rules to pivot data: {str(e)}"
            logger.error(error_msg)
            QMessageBox.critical(self, "Conversion Error", error_msg)


def create_rule_editor(rule_type: RuleType, parent=None) -> RuleEditorWidget:
    """Factory function to create appropriate rule editor based on rule type"""
    if rule_type == RuleType.CLEARANCE:
        return ClearanceRuleEditor(parent)
    elif rule_type == RuleType.SHORT_CIRCUIT:
        return ShortCircuitRuleEditor(parent)
    elif rule_type == RuleType.UNROUTED_NET:
        return UnRoutedNetRuleEditor(parent)
    # Add more rule types as needed
    else:
        logger.warning(f"No specific editor for rule type: {rule_type.value}, using base editor")
        return RuleEditorWidget(parent)
