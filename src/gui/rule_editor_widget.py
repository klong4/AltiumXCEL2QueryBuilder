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
    ClearanceRule, ShortCircuitRule, UnRoutedNetRule
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
        
        # Placeholder for rule-specific properties
        self.rule_specific_group = QGroupBox("Rule Specific Properties")
        self.rule_specific_layout = QFormLayout()
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
            
            logger.info(f"Rule set in editor: {rule.name} ({rule.rule_type.value})")
        else:
            # Clear properties
            self.name_edit.clear()
            self.enabled_checkbox.setChecked(True)
            self.priority_spin.setValue(1)
            self.comment_edit.clear()
            
            logger.warning("Attempted to set None rule")
    
    def get_rule(self) -> BaseRule:
        """Get the edited rule"""
        if not self.rule:
            return None
        
        # Update common properties
        self.rule.name = self.name_edit.text()
        self.rule.enabled = self.enabled_checkbox.isChecked()
        self.rule.priority = self.priority_spin.value()
        self.rule.comment = self.comment_edit.text()
        
        # Update rule-specific properties
        self._update_rule_from_ui()
        
        return self.rule
    
    def _update_rule_specific_properties(self):
        """Update UI with rule-specific properties (to be overridden by subclasses)"""
        pass
    
    def _update_rule_from_ui(self):
        """Update rule with values from UI (to be overridden by subclasses)"""
        pass
    
    def _on_property_changed(self):
        """Handle property changes in the UI"""
        if self.rule:
            # Update rule with current values
            rule = self.get_rule()
            
            # Emit signal
            self.rule_changed.emit(rule)
            logger.debug(f"Rule '{rule.name}' properties changed")


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
    
    def _update_rule_from_ui(self):
        """Update rule with values from UI"""
        if not isinstance(self.rule, ClearanceRule):
            logger.error(f"Expected ClearanceRule, got {type(self.rule).__name__}")
            return
        
        # Update minimum clearance
        self.rule.min_clearance = self.clearance_spin.value()
        
        # Update unit
        unit_str = self.unit_combo.currentData()
        try:
            self.rule.unit = UnitType(unit_str)
        except ValueError:
            logger.error(f"Invalid unit: {unit_str}")
        
        # Update source scope
        source_scope_type = self.source_scope_combo.currentData()
        source_items = []
        
        if source_scope_type != "All":
            source_items_str = self.source_scope_edit.text()
            if source_items_str:
                source_items = source_items_str.split(";")
        
        self.rule.source_scope = RuleScope(source_scope_type, source_items)
        
        # Update target scope
        target_scope_type = self.target_scope_combo.currentData()
        target_items = []
        
        if target_scope_type != "All":
            target_items_str = self.target_scope_edit.text()
            if target_items_str:
                target_items = target_items_str.split(";")
        
        self.rule.target_scope = RuleScope(target_scope_type, target_items)
    
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
    
    def _update_rule_from_ui(self):
        """Update rule with values from UI"""
        if not isinstance(self.rule, ShortCircuitRule):
            logger.error(f"Expected ShortCircuitRule, got {type(self.rule).__name__}")
            return
        
        # Update scope
        scope_type = self.scope_combo.currentData()
        scope_items = []
        
        if scope_type != "All":
            scope_items_str = self.scope_edit.text()
            if scope_items_str:
                scope_items = scope_items_str.split(";")
        
        self.rule.scope = RuleScope(scope_type, scope_items)
    
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
    
    def _update_rule_from_ui(self):
        """Update rule with values from UI"""
        if not isinstance(self.rule, UnRoutedNetRule):
            logger.error(f"Expected UnRoutedNetRule, got {type(self.rule).__name__}")
            return
        
        # Update scope
        scope_type = self.scope_combo.currentData()
        scope_items = []
        
        if scope_type != "All":
            scope_items_str = self.scope_edit.text()
            if scope_items_str:
                scope_items = scope_items_str.split(";")
        
        self.rule.scope = RuleScope(scope_type, scope_items)
    
    def _on_scope_type_changed(self, index):
        """Handle scope type changes"""
        scope_type = self.scope_combo.currentData()
        
        if scope_type == "All":
            self.scope_edit.setEnabled(False)
            self.scope_edit.clear()
        else:
            self.scope_edit.setEnabled(True)
        
        self._on_property_changed()


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
