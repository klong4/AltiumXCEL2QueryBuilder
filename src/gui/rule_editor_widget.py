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
    
    HEADERS = ["Name", "Type", "Enabled", "Priority"]

    def __init__(self, parent=None):
        """Initialize rule table model"""
        super().__init__(parent)
        self._rules: List[BaseRule] = []
        logger.debug("RuleTableModel initialized")
    
    def set_rules(self, rules: List[BaseRule]):
        """Set the list of rules for the model"""
        self.beginResetModel()
        self._rules = sorted(rules, key=lambda r: r.priority) # Keep sorted by priority maybe? Or name?
        self.endResetModel()
        logger.debug(f"Model updated with {len(self._rules)} rules.")
    
    def rowCount(self, parent=QModelIndex()) -> int:
        """Return number of rules"""
        return len(self._rules) if not parent.isValid() else 0
    
    def columnCount(self, parent=QModelIndex()) -> int:
        """Return number of columns"""
        return len(self.HEADERS) if not parent.isValid() else 0
    
    def data(self, index: QModelIndex, role: int = Qt.DisplayRole) -> QVariant:
        """Return data for a given index and role"""
        if not index.isValid() or not (0 <= index.row() < len(self._rules)):
            return QVariant()

        rule = self._rules[index.row()]
        column = index.column()

        if role == Qt.DisplayRole:
            if column == 0: # Name
                return QVariant(rule.name)
            elif column == 1: # Type
                return QVariant(rule.rule_type.value) # Display enum value (string)
            elif column == 2: # Enabled
                return QVariant("Yes" if rule.enabled else "No")
            elif column == 3: # Priority
                return QVariant(rule.priority)
        elif role == Qt.ToolTipRole:
             return QVariant(f"Comment: {rule.comment}" if rule.comment else "No comment")
        # Add other roles like Qt.BackgroundRole if needed

        return QVariant() # Default return
    
    def headerData(self, section: int, orientation: Qt.Orientation, role: int = Qt.DisplayRole) -> QVariant:
        """Return header data"""
        if orientation == Qt.Horizontal and role == Qt.DisplayRole:
            if 0 <= section < len(self.HEADERS):
                return QVariant(self.HEADERS[section])
        return QVariant() # Default return

    def get_rule_at_row(self, row: int) -> Optional[BaseRule]:
        """Get the rule object corresponding to a specific row."""
        if 0 <= row < len(self._rules):
            return self._rules[row]
        return None


class RulesManagerWidget(QWidget):
    """Widget for managing multiple rules"""
    
    rules_changed = pyqtSignal()
    pivot_data_updated = pyqtSignal(object)  # Emits ExcelPivotData
    
    def __init__(self, parent=None):
        """Initialize rules manager widget"""
        super().__init__(parent)
        self.rule_manager = None
        self.current_editor_widget = None # To hold the active editor
        self._init_ui()
        logger.info("Rules manager widget initialized")

    def add_rules(self, new_rules: List[BaseRule]):
        """Add new rules to the manager and update the view"""
        if self.rule_manager:
            for rule in new_rules:
                self.rule_manager.add_rule(rule)
            self.table_model.set_rules(self.rule_manager.get_all_rules())
            self.rules_changed.emit() # Emit signal when rules are added
            logger.info(f"Added {len(new_rules)} rules.")
        else:
            logger.warning("Rule manager not set. Cannot add rules.")

    def _init_ui(self):
        """Initialize the UI components"""
        main_layout = QHBoxLayout()
        self.setLayout(main_layout)

        # Left side: Rule list and buttons
        left_panel = QWidget()
        left_layout = QVBoxLayout()
        left_panel.setLayout(left_layout)

        # Rule table view
        self.table_view = QTableView()
        self.table_model = RuleTableModel()
        self.table_view.setModel(self.table_model)
        self.table_view.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table_view.setSelectionBehavior(QTableView.SelectRows)
        self.table_view.setSelectionMode(QTableView.SingleSelection)
        self.table_view.selectionModel().selectionChanged.connect(self._on_rule_selection_changed)
        left_layout.addWidget(self.table_view)

        # Buttons layout
        button_layout = QHBoxLayout()
        self.add_button = QPushButton("Add Rule")
        self.add_button.clicked.connect(self._on_add_rule)
        self.delete_button = QPushButton("Delete Rule")
        self.delete_button.clicked.connect(self._on_delete_rule)
        self.update_pivot_button = QPushButton("Update Pivot Table")
        self.update_pivot_button.clicked.connect(self._on_update_pivot)

        button_layout.addWidget(self.add_button)
        button_layout.addWidget(self.delete_button)
        button_layout.addStretch()
        button_layout.addWidget(self.update_pivot_button)
        left_layout.addLayout(button_layout)

        # Right side: Rule editor (placeholder layout)
        self.editor_layout = QVBoxLayout() # Layout to hold the current editor
        editor_container = QWidget()
        editor_container.setLayout(self.editor_layout)

        # Add panels to main layout
        main_layout.addWidget(left_panel, 1) # Give list more space initially
        main_layout.addWidget(editor_container, 1) # Editor takes equal space

        # Initially, no editor is shown
        self._show_editor(None)


    def set_rule_manager(self, rule_manager):
        """Set the rule manager instance"""
        self.rule_manager = rule_manager
        self.table_model.set_rules(self.rule_manager.get_all_rules())
        logger.info("Rule manager set and table updated.")

    def get_rule_manager(self):
        """Get the rule manager instance"""
        return self.rule_manager

    def _show_editor(self, rule: Optional[BaseRule]):
        """Creates or updates the editor widget for the given rule."""
        # Remove the existing editor widget if it exists
        if self.current_editor_widget:
            self.editor_layout.removeWidget(self.current_editor_widget)
            self.current_editor_widget.deleteLater()
            self.current_editor_widget = None

        if rule:
            # Create the appropriate editor for the rule type
            try:
                # Disconnect previous editor's signal if any (safety measure)
                # This might not be strictly necessary if deleteLater works reliably
                # if self.current_editor_widget:
                #     try:
                #         self.current_editor_widget.rule_changed.disconnect(self._on_rule_changed)
                #     except TypeError: # Signal has no slots to disconnect
                #         pass

                self.current_editor_widget = create_rule_editor(rule.rule_type, self)
                if self.current_editor_widget:
                    self.current_editor_widget.set_rule(rule)
                    self.current_editor_widget.rule_changed.connect(self._on_rule_changed)
                    self.editor_layout.addWidget(self.current_editor_widget)
                    logger.debug(f"Showing editor for rule: {rule.name}")
                else:
                    logger.warning(f"No specific editor found for rule type {rule.rule_type}. Cannot display editor.")
                    # Optionally show a placeholder or message widget
                    placeholder = QLabel(f"No editor available for rule type: {rule.rule_type.name}")
                    placeholder.setAlignment(Qt.AlignCenter)
                    self.current_editor_widget = placeholder # Assign placeholder to be removed later
                    self.editor_layout.addWidget(self.current_editor_widget)

            except Exception as e:
                logger.error(f"Error creating/showing rule editor: {e}", exc_info=True)
                QMessageBox.critical(self, "Error", f"Could not create editor for rule '{rule.name}':\\n{e}")
        else:
            # Show a placeholder if no rule is selected
            placeholder = QLabel("Select a rule to edit its properties.")
            placeholder.setAlignment(Qt.AlignCenter)
            self.current_editor_widget = placeholder # Assign placeholder to be removed later
            self.editor_layout.addWidget(self.current_editor_widget)
            logger.debug("No rule selected, showing placeholder.")


    def _on_rule_selection_changed(self, selected, deselected):
        """Handle rule selection changes in the table view"""
        indexes = selected.indexes()
        if indexes:
            row = indexes[0].row()
            selected_rule = self.table_model.get_rule_at_row(row) # Assuming this method exists in RuleTableModel
            if selected_rule:
                logger.info(f"Rule selected: {selected_rule.name}")
                self._show_editor(selected_rule)
            else:
                logger.warning(f"Could not retrieve rule at selected row {row}.")
                self._show_editor(None)
        else:
            # No selection
            logger.info("Rule selection cleared.")
            self._show_editor(None)

    def _on_add_rule(self):
        """Handle add rule button click"""
        # Example: Add a default ClearanceRule
        # In a real app, you might show a dialog to choose rule type and initial name
        if self.rule_manager:
            default_rule = ClearanceRule(name="New Clearance Rule")
            self.rule_manager.add_rule(default_rule)
            self.table_model.set_rules(self.rule_manager.get_all_rules()) # Refresh model
            
            # Select the newly added rule in the table
            new_row_index = self.table_model.rowCount() - 1
            if new_row_index >= 0:
                qt_index = self.table_model.index(new_row_index, 0)
                self.table_view.setCurrentIndex(qt_index)
                # Selection change should trigger _on_rule_selection_changed and show editor

            self.rules_changed.emit() # Emit signal
            logger.info(f"Added new rule: {default_rule.name}")
        else:
            QMessageBox.warning(self, "Warning", "Rule manager not available.")
            logger.warning("Cannot add rule: Rule manager not set.")

    def _on_delete_rule(self):
        """Handle delete rule button click"""
        selected_indexes = self.table_view.selectionModel().selectedRows()
        if not selected_indexes:
            QMessageBox.information(self, "Delete Rule", "Please select a rule to delete.")
            return

        row = selected_indexes[0].row()
        rule_to_delete = self.table_model.get_rule_at_row(row) # Assuming this method exists

        if rule_to_delete and self.rule_manager:
            reply = QMessageBox.question(self, "Confirm Delete",
                                         f"Are you sure you want to delete the rule '{rule_to_delete.name}'?",
                                         QMessageBox.Yes | QMessageBox.No, QMessageBox.No)

            if reply == QMessageBox.Yes:
                self.rule_manager.delete_rule(rule_to_delete.name)
                self.table_model.set_rules(self.rule_manager.get_all_rules()) # Refresh model
                self._show_editor(None) # Clear editor after deletion
                self.rules_changed.emit() # Emit signal
                logger.info(f"Deleted rule: {rule_to_delete.name}")
        elif not self.rule_manager:
             QMessageBox.warning(self, "Warning", "Rule manager not available.")
             logger.warning("Cannot delete rule: Rule manager not set.")
        else:
            logger.error(f"Could not retrieve rule at row {row} for deletion.")


    def _on_rule_changed(self, rule: BaseRule):
        """Handle changes made in the rule editor"""
        if self.rule_manager and rule:
            # The rule object passed *is* the one from the manager,
            # modifications are already applied by the editor's _on_property_changed.
            # We just need to update the table view potentially and emit signal.
            
            # Find the row corresponding to the rule to update the view
            row = self.rule_manager.get_rule_index(rule.name) # Assuming RuleManager has this
            if row is not None and row >= 0:
                # Notify the model that data has changed for this row
                start_index = self.table_model.index(row, 0)
                end_index = self.table_model.index(row, self.table_model.columnCount() - 1)
                self.table_model.dataChanged.emit(start_index, end_index)
                logger.debug(f"Table view notified of change for rule: {rule.name}")

            self.rules_changed.emit() # Emit signal that rules (potentially) changed
            logger.info(f"Rule '{rule.name}' updated via editor.")
        elif not self.rule_manager:
             logger.warning("Cannot process rule change: Rule manager not set.")


    def _on_update_pivot(self):
        """Handle update pivot table button click"""
        if self.rule_manager:
            try:
                # Assuming rule_manager can generate pivot data
                pivot_data = self.rule_manager.generate_pivot_data() # Placeholder method
                if pivot_data:
                    self.pivot_data_updated.emit(pivot_data)
                    logger.info("Pivot data updated and signal emitted.")
                else:
                    logger.warning("Pivot data generation returned nothing.")
            except AttributeError:
                 logger.error("Rule manager does not have 'generate_pivot_data' method.")
                 QMessageBox.critical(self, "Error", "Feature not implemented in Rule Manager.")
            except Exception as e:
                 logger.error(f"Error generating pivot data: {e}", exc_info=True)
                 QMessageBox.critical(self, "Error", f"Failed to generate pivot data:\\n{e}")
        else:
            QMessageBox.warning(self, "Warning", "Rule manager not available.")
            logger.warning("Cannot update pivot: Rule manager not set.")


def create_rule_editor(rule_type: RuleType, parent=None) -> Optional[RuleEditorWidget]:
    """Factory function to create appropriate rule editor based on rule type"""
    editor_class = None
    if rule_type == RuleType.CLEARANCE:
        editor_class = ClearanceRuleEditor
    elif rule_type == RuleType.SHORT_CIRCUIT:
        editor_class = ShortCircuitRuleEditor
    elif rule_type == RuleType.UNROUTED_NET:
        editor_class = UnRoutedNetRuleEditor
    # Add other rule types here
    # elif rule_type == RuleType.POWER_PLANE_CONNECT:
    #     editor_class = PowerPlaneConnectRuleEditor
    # ... etc ...

    if editor_class:
        return editor_class(parent)
    else:
        logger.warning(f"No specific editor class defined for rule type: {rule_type}")
        # Return a generic base editor or None if preferred
        # return RuleEditorWidget(parent) # Or return None
        return None # Return None if no specific editor exists
