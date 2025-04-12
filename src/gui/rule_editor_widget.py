#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Rule Editor Widget
==================

Widget for viewing, editing, adding, and deleting Altium design rules.
"""

import logging
from typing import Dict, List, Optional, Union, Tuple # Add List
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QTreeView, QPushButton, QMessageBox,
    QAbstractItemView, QMenu, QListWidget, QListWidgetItem, QGroupBox, QLabel,
    QFileDialog, QDialog # Added QDialog
)
from PyQt5.QtCore import Qt, pyqtSignal, QAbstractItemModel, QModelIndex, QVariant
from PyQt5.QtGui import QStandardItemModel, QStandardItem, QColor, QBrush

# Import RuleManager and specific rule types
from models.rule_model import BaseRule, RuleManager, RuleType, UnitType, RuleScope, ClearanceRule, SingleScopeRule
from services.rule_generator import RuleGeneratorError # Keep this for the except block
# Import the new dialog
from .rule_edit_dialog import RuleEditDialog

logger = logging.getLogger(__name__)

class RulesManagerWidget(QWidget):
    """Widget to manage (view, edit, add, delete) Altium rules."""

    # Signals
    rules_updated = pyqtSignal(list) # Emits the current list of rules when changed
    unsaved_changes_changed = pyqtSignal(bool) # Emits True if there are unsaved changes

    def __init__(self, parent=None):
        """Initialize rules manager widget"""
        super().__init__(parent)
        # self._rule_manager = parent # Removed: Manage rules internally
        self._rules: List[BaseRule] = [] # Initialize internal rule list
        self._unsaved_changes = False

        self._init_ui()
        # self.load_rules() # Removed: Load rules explicitly via set_and_load_rules

    def _init_ui(self):
        """Initialize the user interface."""
        layout = QVBoxLayout(self)

        # --- Rule View (Using QTreeView for potential hierarchy) ---
        self.rules_list_widget = QListWidget()
        self.rules_list_widget.setAlternatingRowColors(True)
        self.rules_list_widget.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.rules_list_widget.setContextMenuPolicy(Qt.CustomContextMenu)
        self.rules_list_widget.customContextMenuRequested.connect(self._show_context_menu)
        layout.addWidget(self.rules_list_widget)

        # --- Action Buttons ---
        button_layout = QHBoxLayout()
        self.add_button = QPushButton("Add Rule")
        self.edit_button = QPushButton("Edit Rule")
        self.delete_button = QPushButton("Delete Rule")
        self.clear_button = QPushButton("Clear All Rules")
        self.save_button = QPushButton("Save Rules")
        self.export_button = QPushButton("Export Rules")

        self.add_button.clicked.connect(self._add_rule)
        self.edit_button.clicked.connect(self._edit_rule)
        self.delete_button.clicked.connect(self._delete_rule)
        self.clear_button.clicked.connect(self._clear_rules)
        self.save_button.clicked.connect(self._save_rules)
        self.export_button.clicked.connect(self._export_rules)

        button_layout.addWidget(self.add_button)
        button_layout.addWidget(self.edit_button)
        button_layout.addWidget(self.delete_button)
        button_layout.addWidget(self.clear_button)
        button_layout.addWidget(self.save_button)
        button_layout.addWidget(self.export_button)
        button_layout.addStretch()
        layout.addLayout(button_layout)

        # Connect selection change to enable/disable buttons
        self.rules_list_widget.selectionModel().selectionChanged.connect(self._update_button_states)
        self._update_button_states() # Initial state

        # --- Rule Details ---
        self.details_group = QGroupBox("Rule Details")
        layout.addWidget(self.details_group)
        self.details_layout = QVBoxLayout()
        self.details_group.setLayout(self.details_layout)

        # Connect selection change to load rule details
        self.rules_list_widget.itemSelectionChanged.connect(self._on_selection_changed)

    def _update_button_states(self):
        """Enable/disable buttons based on selection state."""
        has_selection = bool(self.rules_list_widget.selectedItems())
        self.edit_button.setEnabled(has_selection)
        self.delete_button.setEnabled(has_selection)
        # Enable/disable other buttons based on selection or other criteria
        self.clear_button.setEnabled(self.rules_list_widget.count() > 0)
        self.save_button.setEnabled(self.rules_list_widget.count() > 0)
        self.export_button.setEnabled(self.rules_list_widget.count() > 0)

    def _show_context_menu(self, position):
        """Show context menu for the rule list."""
        # Placeholder: Implement context menu actions (e.g., Edit, Delete)
        context_menu = QMenu(self)
        edit_action = context_menu.addAction("Edit")
        delete_action = context_menu.addAction("Delete")
        # Add more actions as needed

        action = context_menu.exec_(self.rules_list_widget.mapToGlobal(position))

        if action == edit_action:
            self._edit_rule()
        elif action == delete_action:
            self._delete_rule()
        # Handle other actions

    def set_and_load_rules(self, rules: List[BaseRule]):
        """Set the internal rules list and load them into the list widget."""
        self.rules_list_widget.clear()
        if rules is not None:
            logger.info(f"Loading {len(rules)} rules into the editor view.")
            # Store the actual rule objects, making a copy
            self._rules = list(rules)
            for rule in self._rules:
                item = QListWidgetItem(f"{rule.name} ({rule.rule_type.value})")
                # Store the rule object with the item for later retrieval
                item.setData(Qt.UserRole, rule)
                self.rules_list_widget.addItem(item)
        else:
            logger.warning("Received None or empty list, clearing rules view.")
            self._rules = [] # Ensure _rules is an empty list

        self._update_rule_details(None) # Clear details view
        self._set_unsaved_changes(False) # Reset unsaved changes flag after loading
        logger.debug(f"Rules loaded, unsaved changes set to {self._unsaved_changes}")

    def _on_selection_changed(self):
        """Handle selection changes in the rules list."""
        selected_items = self.rules_list_widget.selectedItems()
        if selected_items:
            # Get the rule object stored in the item's data
            selected_rule = selected_items[0].data(Qt.UserRole)
            self._update_rule_details(selected_rule)
        else:
            self._update_rule_details(None)

    def _update_rule_details(self, rule: Optional[BaseRule]):
        """Update the details view with the selected rule's information."""
        # Clear previous widgets in the details layout
        for i in reversed(range(self.details_layout.count())):
            widget = self.details_layout.itemAt(i).widget()
            if widget is not None:
                widget.deleteLater()

        if rule is None:
            self.details_layout.addWidget(QLabel("Select a rule to view details."))
            return

        # --- General Rule Properties ---
        self.details_layout.addWidget(QLabel(f"Name: {rule.name}"))
        self.details_layout.addWidget(QLabel(f"Type: {rule.rule_type.name}"))
        # Display scope based on rule type
        if isinstance(rule, ClearanceRule):
            self.details_layout.addWidget(QLabel(f"Source Scope: {rule.source_scope.to_query_string()}"))
            self.details_layout.addWidget(QLabel(f"Target Scope: {rule.target_scope.to_query_string()}"))
            self.details_layout.addWidget(QLabel(f"Min Clearance: {rule.min_clearance} {rule.unit.value}"))
        elif isinstance(rule, SingleScopeRule):
            self.details_layout.addWidget(QLabel(f"Scope: {rule.scope.to_query_string()}"))
        else:
            self.details_layout.addWidget(QLabel("Scope: (Not applicable or unknown)"))

        # Remove generic value display, handled by specific types now
        # value_str = "..." # Placeholder for complex value
        # if hasattr(rule, 'value'):
        #     value_str = str(rule.value)
        # elif hasattr(rule, 'clearance'): # Example for ClearanceRule
        #     value_str = str(rule.clearance)
        # self.details_layout.addWidget(QLabel(f"Value: {value_str}"))

        # --- Rule Type Specific Properties (Add more details here if needed) ---
        if isinstance(rule, ClearanceRule):
            # Already added clearance value above
            pass # Add more specific details if necessary
        # Add elif blocks for other rule types (Width, ViaStyle, etc.)
        # elif isinstance(rule, WidthRule):
        #    # ... WidthRule specific fields ...
        elif not isinstance(rule, SingleScopeRule): # If not Clearance or SingleScope
            self.details_layout.addWidget(QLabel(f"Details view not fully implemented for rule type: {type(rule).__name__}"))

        # No need to set layout again, just add widgets
        # self.details_group.setLayout(self.details_layout)

    def _add_rule(self):
        """Add a new default rule."""
        # For now, add a default ClearanceRule
        # TODO: Allow selecting rule type to add
        new_rule_name = f"New_Rule_{len(self._rules) + 1}"
        new_rule = ClearanceRule(name=new_rule_name)
        self._rules.append(new_rule)

        item = QListWidgetItem(f"{new_rule.name} ({new_rule.rule_type.value})")
        item.setData(Qt.UserRole, new_rule)
        self.rules_list_widget.addItem(item)
        self.rules_list_widget.setCurrentItem(item) # Select the new rule
        self._set_unsaved_changes(True)
        logger.info(f"Added new rule: {new_rule_name}")

    def _edit_rule(self):
        """Open the RuleEditDialog for the selected rule."""
        selected_items = self.rules_list_widget.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "Edit Rule", "Please select a rule to edit.")
            return

        selected_item = selected_items[0]
        rule_to_edit = selected_item.data(Qt.UserRole)

        if not isinstance(rule_to_edit, BaseRule):
            logger.error(f"Invalid data found for selected item: {rule_to_edit}")
            QMessageBox.critical(self, "Error", "Could not retrieve rule data for editing.")
            return

        # Create and show the dialog
        dialog = RuleEditDialog(rule_to_edit, self)
        if dialog.exec_() == QDialog.Accepted:
            updated_data = dialog.get_updated_data()
            logger.info(f"Applying updated data for rule '{rule_to_edit.name}': {updated_data}")

            # Update the original rule object directly
            try:
                rule_to_edit.name = updated_data.get('name', rule_to_edit.name)
                rule_to_edit.enabled = updated_data.get('enabled', rule_to_edit.enabled)
                rule_to_edit.priority = updated_data.get('priority', rule_to_edit.priority)
                rule_to_edit.comment = updated_data.get('comment', rule_to_edit.comment)

                if isinstance(rule_to_edit, ClearanceRule):
                    rule_to_edit.min_clearance = updated_data.get('min_clearance', rule_to_edit.min_clearance)
                    rule_to_edit.unit = updated_data.get('unit', rule_to_edit.unit)
                    rule_to_edit.source_scope = updated_data.get('source_scope', rule_to_edit.source_scope)
                    rule_to_edit.target_scope = updated_data.get('target_scope', rule_to_edit.target_scope)
                elif isinstance(rule_to_edit, SingleScopeRule):
                    rule_to_edit.scope = updated_data.get('scope', rule_to_edit.scope)
                # Add elif blocks for other specific rule types if they exist

                # Update the list widget item text
                selected_item.setText(f"{rule_to_edit.name} ({rule_to_edit.rule_type.value})")
                # Update the details view if this item is still selected
                self._update_rule_details(rule_to_edit)
                # Mark changes as unsaved
                self._set_unsaved_changes(True)
                logger.info(f"Rule '{rule_to_edit.name}' updated successfully.")
            except Exception as e:
                logger.error(f"Error applying updated rule data: {e}", exc_info=True)
                QMessageBox.critical(self, "Update Error", f"Failed to apply rule changes: {e}")
        else:
            logger.info(f"Edit cancelled for rule '{rule_to_edit.name}'.")

    def _delete_rule(self):
        """Delete the selected rule(s)."""
        selected_items = self.rules_list_widget.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "Delete Rule", "Please select rule(s) to delete.")
            return

        reply = QMessageBox.question(self, "Confirm Deletion",
                                     f"Are you sure you want to delete {len(selected_items)} rule(s)?",
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)

        if reply == QMessageBox.Yes:
            rows_to_delete = []
            for item in selected_items:
                rule_to_delete = item.data(Qt.UserRole)
                if rule_to_delete in self._rules:
                    self._rules.remove(rule_to_delete)
                rows_to_delete.append(self.rules_list_widget.row(item))

            # Remove items from list widget (iterate backwards to avoid index issues)
            for row in sorted(rows_to_delete, reverse=True):
                self.rules_list_widget.takeItem(row)

            logger.info(f"Deleted {len(selected_items)} rules. Remaining: {len(self._rules)}")
            self._update_rule_details(None) # Clear details view
            self._set_unsaved_changes(True)

    def _clear_rules(self):
        """Clear all rules from the manager."""
        if not self._rules:
            QMessageBox.information(self, "Clear Rules", "There are no rules to clear.")
            return

        reply = QMessageBox.question(self, "Confirm Clear All",
                                     "Are you sure you want to clear all rules? This cannot be undone.",
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)

        if reply == QMessageBox.Yes:
            self._rules.clear()
            self.rules_list_widget.clear()
            self._update_rule_details(None)
            self._set_unsaved_changes(True)
            logger.info("Cleared all rules.")

    def _save_rules(self):
        """Save the current rules to a file (internal format, e.g., JSON/Pickle)."""
        if not self._rules:
            QMessageBox.warning(self, "Save Rules", "There are no rules to save.")
            return

        # Use self._rules directly
        rules_to_save = self._rules
        # TODO: Implement actual saving logic (e.g., to JSON or Pickle)
        # For now, just simulate saving
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Save Rules As", "", "Rule Files (*.json);;All Files (*)"
        )
        if file_path:
            try:
                # Placeholder: Replace with actual serialization
                with open(file_path, 'w') as f:
                    import json
                    # Need a way to serialize rule objects (e.g., custom encoder)
                    # json.dump([rule.to_dict() for rule in rules_to_save], f, indent=4)
                    f.write(f"Placeholder: {len(rules_to_save)} rules saved.")
                logger.info(f"Saved {len(rules_to_save)} rules to {file_path}")
                self._set_unsaved_changes(False)
                QMessageBox.information(self, "Save Successful", f"Rules saved to:\n{file_path}")
            except Exception as e:
                logger.error(f"Error saving rules to {file_path}: {e}", exc_info=True)
                QMessageBox.critical(self, "Save Error", f"Failed to save rules: {e}")

    def _export_rules(self):
        """Export the current rules to a .RUL file."""
        if not self._rules:
            QMessageBox.warning(self, "Export Error", "No rules to export.")
            return

        rules_to_export = self._rules
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Export Rules As", "", "Altium Rule Files (*.RUL);;All Files (*)"
        )
        if file_path:
            try:
                # Create a RuleManager instance to handle export
                rule_manager = RuleManager()
                for rule in rules_to_export:
                    rule_manager.add_rule(rule) # Add rules to the manager

                # Use the RuleManager's export method
                rule_manager.export_rules_to_file(file_path)

                logger.info(f"Exported {len(rules_to_export)} rules to {file_path}")
                # Exporting doesn't necessarily mean changes are 'saved' internally
                # self._set_unsaved_changes(False) # Decide if export should reset this
                QMessageBox.information(self, "Export Successful", f"Rules exported to:\n{file_path}")
            except RuleGeneratorError as rge: # Keep this catch if RuleManager raises it indirectly or for other potential errors
                logger.error(f"Rule Generator Error during export: {rge}", exc_info=True)
                QMessageBox.critical(self, "Export Error", f"Rule Generation Error: {rge}")
            except Exception as e:
                logger.error(f"Error exporting rules to {file_path}: {e}", exc_info=True)
                QMessageBox.critical(self, "Export Error", f"An unexpected error occurred during export: {str(e)}")

    def _import_rules(self):
        """Import rules from a .RUL file."""
        # TODO: Implement import functionality
        pass

    def _set_unsaved_changes(self, changed: bool):
        """Set the unsaved changes flag and emit signal if state changes."""
        if self._unsaved_changes != changed:
            self._unsaved_changes = changed
            self.unsaved_changes_changed.emit(changed)
            logger.debug(f"Unsaved changes status set to: {changed}")

    def has_unsaved_changes(self) -> bool:
        """Check if there are unsaved changes."""
        # This relies on the _unsaved_changes flag which should be set correctly
        # by methods like _add_rule, _delete_rule, _edit_rule, _clear_rules, _save_rule_details
        return self._unsaved_changes

    def get_current_rules(self) -> List[BaseRule]:
        """Return the current list of rules being managed."""
        return self._rules
