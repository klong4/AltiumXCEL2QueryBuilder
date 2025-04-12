#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Rule Edit Dialog
================

Dialog for editing the parameters of an Altium design rule.
"""

import logging
from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QFormLayout, QLineEdit, QCheckBox, QSpinBox,
    QComboBox, QDialogButtonBox, QLabel, QWidget, QHBoxLayout, QListWidget,
    QPushButton, QStackedWidget, QDoubleSpinBox, QMessageBox
)
from PyQt5.QtCore import Qt

from models.rule_model import BaseRule, RuleType, UnitType, RuleScope, ClearanceRule, SingleScopeRule

logger = logging.getLogger(__name__)

class RuleEditDialog(QDialog):
    """Dialog window for editing a single rule."""

    def __init__(self, rule: BaseRule, parent=None):
        """Initialize the rule edit dialog."""
        super().__init__(parent)
        self.rule = rule # Store the original rule object
        self.updated_rule_data = {} # Store changes

        self.setWindowTitle(f"Edit Rule: {rule.name}")
        self.setMinimumWidth(500)

        layout = QVBoxLayout(self)

        # --- General Properties ---
        self.general_layout = QFormLayout()
        self.name_edit = QLineEdit(rule.name)
        self.enabled_check = QCheckBox()
        self.enabled_check.setChecked(rule.enabled)
        self.priority_spin = QSpinBox()
        self.priority_spin.setRange(1, 100) # Example range
        self.priority_spin.setValue(rule.priority)
        self.comment_edit = QLineEdit(rule.comment)

        self.general_layout.addRow("Name:", self.name_edit)
        self.general_layout.addRow("Enabled:", self.enabled_check)
        self.general_layout.addRow("Priority:", self.priority_spin)
        self.general_layout.addRow("Comment:", self.comment_edit)
        layout.addLayout(self.general_layout)

        # --- Rule Type Specific Properties ---
        self.specific_props_widget = QWidget() # Container for specific props
        self.specific_layout = QFormLayout(self.specific_props_widget)
        layout.addWidget(self.specific_props_widget)

        self._setup_specific_properties()

        # --- Dialog Buttons ---
        self.button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        self.button_box.accepted.connect(self._on_accept)
        self.button_box.rejected.connect(self.reject)
        layout.addWidget(self.button_box)

    def _setup_specific_properties(self):
        """Set up input fields based on the rule type."""
        # Clear existing specific properties widgets if any
        while self.specific_layout.count():
            item = self.specific_layout.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()

        if isinstance(self.rule, ClearanceRule):
            self._setup_clearance_properties()
        elif isinstance(self.rule, SingleScopeRule):
            self._setup_single_scope_properties()
        # Add elif blocks for other rule types here
        # elif isinstance(self.rule, WidthRule):
        #     self._setup_width_properties()
        else:
            self.specific_layout.addRow(QLabel(f"Editing for rule type '{self.rule.rule_type.name}' not fully implemented."))

    def _setup_clearance_properties(self):
        """Setup fields specific to ClearanceRule."""
        rule: ClearanceRule = self.rule # Type hint for clarity

        self.min_clearance_spin = QDoubleSpinBox()
        self.min_clearance_spin.setDecimals(4) # Adjust precision as needed
        self.min_clearance_spin.setRange(0, 10000) # Example range
        self.min_clearance_spin.setValue(rule.min_clearance)

        self.unit_combo = QComboBox()
        for unit in UnitType:
            self.unit_combo.addItem(unit.value, unit) # Store enum member as data
        self.unit_combo.setCurrentIndex(self.unit_combo.findData(rule.unit))

        self.source_scope_widget = self._create_scope_widget(rule.source_scope)
        self.target_scope_widget = self._create_scope_widget(rule.target_scope)

        self.specific_layout.addRow("Min Clearance:", self.min_clearance_spin)
        self.specific_layout.addRow("Unit:", self.unit_combo)
        self.specific_layout.addRow("Source Scope:", self.source_scope_widget)
        self.specific_layout.addRow("Target Scope:", self.target_scope_widget)

    def _setup_single_scope_properties(self):
        """Setup fields specific to rules with a single scope."""
        rule: SingleScopeRule = self.rule # Type hint

        self.scope_widget = self._create_scope_widget(rule.scope)
        self.specific_layout.addRow("Scope:", self.scope_widget)

    def _create_scope_widget(self, scope: RuleScope) -> QWidget:
        """Creates a widget for editing a RuleScope."""
        widget = QWidget()
        layout = QHBoxLayout(widget)
        layout.setContentsMargins(0, 0, 0, 0)

        scope_type_combo = QComboBox()
        scope_type_combo.addItems(["All", "NetClass", "NetClasses", "Custom"])
        scope_type_combo.setCurrentText(scope.scope_type)

        # Use QStackedWidget to show relevant input based on scope type
        stacked_widget = QStackedWidget()
        all_label = QLabel("(Applies to all objects)") # Placeholder for 'All'
        netclass_edit = QLineEdit()
        netclasses_list = QListWidget() # Simple list for now
        netclasses_list.setToolTip("Enter one NetClass per line")
        custom_edit = QLineEdit()

        stacked_widget.addWidget(all_label)         # Index 0: All
        stacked_widget.addWidget(netclass_edit)     # Index 1: NetClass
        stacked_widget.addWidget(netclasses_list)   # Index 2: NetClasses
        stacked_widget.addWidget(custom_edit)       # Index 3: Custom

        # Populate based on initial scope
        if scope.scope_type == "NetClass" and scope.items:
            netclass_edit.setText(scope.items[0])
            stacked_widget.setCurrentIndex(1)
        elif scope.scope_type == "NetClasses":
            netclasses_list.addItems(scope.items)
            stacked_widget.setCurrentIndex(2)
        elif scope.scope_type == "Custom" and scope.items:
            custom_edit.setText(scope.items[0])
            stacked_widget.setCurrentIndex(3)
        else: # Default to 'All'
            scope_type_combo.setCurrentText("All")
            stacked_widget.setCurrentIndex(0)

        # Connect signal to change stacked widget page
        def update_stacked_widget(index):
            text = scope_type_combo.itemText(index)
            if text == "All":
                stacked_widget.setCurrentIndex(0)
            elif text == "NetClass":
                stacked_widget.setCurrentIndex(1)
            elif text == "NetClasses":
                stacked_widget.setCurrentIndex(2)
            elif text == "Custom":
                stacked_widget.setCurrentIndex(3)

        scope_type_combo.currentIndexChanged.connect(update_stacked_widget)
        # Initial call to set the correct widget
        update_stacked_widget(scope_type_combo.currentIndex())

        layout.addWidget(scope_type_combo)
        layout.addWidget(stacked_widget)

        # Store references to easily retrieve data later
        widget.setProperty("scope_type_combo", scope_type_combo)
        widget.setProperty("stacked_widget", stacked_widget)
        widget.setProperty("netclass_edit", netclass_edit)
        widget.setProperty("netclasses_list", netclasses_list)
        widget.setProperty("custom_edit", custom_edit)

        return widget

    def _get_scope_from_widget(self, scope_widget: QWidget) -> RuleScope:
        """Retrieves RuleScope data from the scope editor widget."""
        scope_type_combo = scope_widget.property("scope_type_combo")
        stacked_widget = scope_widget.property("stacked_widget")
        netclass_edit = scope_widget.property("netclass_edit")
        netclasses_list = scope_widget.property("netclasses_list")
        custom_edit = scope_widget.property("custom_edit")

        scope_type = scope_type_combo.currentText()
        items = []

        if scope_type == "NetClass":
            items = [netclass_edit.text().strip()]
        elif scope_type == "NetClasses":
            items = [netclasses_list.item(i).text().strip() for i in range(netclasses_list.count()) if netclasses_list.item(i).text().strip()]
        elif scope_type == "Custom":
            items = [custom_edit.text().strip()]
        # 'All' has no items

        return RuleScope(scope_type=scope_type, items=items)

    def _on_accept(self):
        """Validate and store updated data before accepting."""
        # --- Validate Name (Example) ---
        new_name = self.name_edit.text().strip()
        if not new_name:
            QMessageBox.warning(self, "Validation Error", "Rule name cannot be empty.")
            return # Prevent dialog from closing

        # --- Collect General Properties ---
        self.updated_rule_data = {
            "name": new_name,
            "enabled": self.enabled_check.isChecked(),
            "priority": self.priority_spin.value(),
            "comment": self.comment_edit.text().strip()
        }

        # --- Collect Specific Properties ---
        if isinstance(self.rule, ClearanceRule):
            self.updated_rule_data["min_clearance"] = self.min_clearance_spin.value()
            self.updated_rule_data["unit"] = self.unit_combo.currentData() # Get stored enum
            self.updated_rule_data["source_scope"] = self._get_scope_from_widget(self.source_scope_widget)
            self.updated_rule_data["target_scope"] = self._get_scope_from_widget(self.target_scope_widget)
        elif isinstance(self.rule, SingleScopeRule):
            self.updated_rule_data["scope"] = self._get_scope_from_widget(self.scope_widget)
        # Add elif for other rule types

        logger.info(f"Rule '{self.rule.name}' edit accepted. New data: {self.updated_rule_data}")
        self.accept() # Close the dialog

    def get_updated_data(self) -> dict:
        """Return the collected updated rule data."""
        return self.updated_rule_data

# Example usage (for testing)
if __name__ == '__main__':
    import sys
    from PyQt5.QtWidgets import QApplication

    app = QApplication(sys.argv)

    # Create a dummy rule to edit
    dummy_scope = RuleScope("NetClass", ["PowerNets"])
    dummy_rule = ClearanceRule(
        name="TestClearance",
        min_clearance=8.5,
        unit=UnitType.MIL,
        source_scope=RuleScope("All"),
        target_scope=dummy_scope
    )
    # dummy_rule = ShortCircuitRule(name="TestShort", scope=RuleScope("Custom", ["IsPad or IsVia"]))


    dialog = RuleEditDialog(dummy_rule)
    if dialog.exec_() == QDialog.Accepted:
        print("Dialog Accepted. Updated data:")
        updated_data = dialog.get_updated_data()
        print(updated_data)
        # Here you would update the original rule object
        # e.g., dummy_rule.name = updated_data['name'] ... etc.
    else:
        print("Dialog Cancelled.")

    # sys.exit(app.exec_()) # Keep running if needed for testing
