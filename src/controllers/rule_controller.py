#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Rule Controller
=============

Controller for managing rule data and interactions.
"""

import os
import logging
from typing import Dict, List, Optional, Union, Tuple

from PyQt5.QtWidgets import QWidget, QFileDialog, QMessageBox

from gui.rule_editor_widget import RuleEditorWidget, create_rule_editor
from models.rule_model import (
    UnitType, RuleType, RuleManager, BaseRule,
    ClearanceRule, ShortCircuitRule, UnRoutedNetRule
)
from services.rule_generator import RuleGenerator, RuleGeneratorError

logger = logging.getLogger(__name__)

class RuleController:
    """Controller for rule operations"""
    
    def __init__(self, config_manager):
        """Initialize rule controller"""
        self.config = config_manager
        self.rule_generator = RuleGenerator()
        self.current_file_path = None
        
        logger.info("Rule controller initialized")
    
    def add_rules(self, rules: List[BaseRule]) -> bool:
        """Add rules to the rule generator"""
        try:
            if not rules:
                logger.warning("No rules to add")
                return False
            
            self.rule_generator.add_rules(rules)
            logger.info(f"Added {len(rules)} rules")
            return True
        
        except Exception as e:
            error_msg = f"Error adding rules: {str(e)}"
            logger.error(error_msg)
            return False
    
    def clear_rules(self) -> bool:
        """Clear all rules from the rule generator"""
        try:
            self.rule_generator.clear_rules()
            logger.info("Cleared all rules")
            return True
        
        except Exception as e:
            error_msg = f"Error clearing rules: {str(e)}"
            logger.error(error_msg)
            return False
    
    def get_rules_by_type(self, rule_type: RuleType) -> List[BaseRule]:
        """Get all rules of a specific type"""
        try:
            rules = self.rule_generator.get_rules_by_type(rule_type)
            logger.info(f"Got {len(rules)} rules of type {rule_type.value}")
            return rules
        
        except Exception as e:
            error_msg = f"Error getting rules by type: {str(e)}"
            logger.error(error_msg)
            return []
    
    def get_all_rules(self) -> List[BaseRule]:
        """Get all rules"""
        try:
            return self.rule_generator.rule_manager.rules
        
        except Exception as e:
            error_msg = f"Error getting all rules: {str(e)}"
            logger.error(error_msg)
            return []
    
    def import_rul_file(self, parent_widget: QWidget = None) -> bool:
        """Import .RUL file and populate rule generator"""
        try:
            # Get file path
            file_path, _ = QFileDialog.getOpenFileName(
                parent_widget,
                "Import RUL File",
                self.config.get("last_directory", ""),
                "RUL Files (*.RUL *.rul);;All Files (*)"
            )
            
            if not file_path:
                # User cancelled
                return False
            
            # Update last directory
            self.config.update_last_directory(os.path.dirname(file_path))
            
            # Parse RUL file
            success = self.rule_generator.parse_rul_file(file_path)
            
            if not success:
                error_msg = "Failed to parse RUL file"
                logger.error(error_msg)
                
                if parent_widget:
                    QMessageBox.critical(parent_widget, "Import Error", error_msg)
                
                return False
            
            # Store current file info
            self.current_file_path = file_path
            self.config.add_recent_file(file_path)
            
            num_rules = len(self.rule_generator.rule_manager.rules)
            logger.info(f"Imported RUL file: {file_path} with {num_rules} rules")
            
            if parent_widget:
                QMessageBox.information(
                    parent_widget,
                    "Import Complete",
                    f"Successfully imported {num_rules} rules from RUL file:\n{file_path}"
                )
            
            return True
        
        except RuleGeneratorError as e:
            error_msg = f"Error importing RUL file: {str(e)}"
            logger.error(error_msg)
            
            if parent_widget:
                QMessageBox.critical(parent_widget, "Import Error", error_msg)
            
            return False
        
        except Exception as e:
            error_msg = f"Unexpected error importing RUL file: {str(e)}"
            logger.error(error_msg)
            
            if parent_widget:
                QMessageBox.critical(parent_widget, "Import Error", error_msg)
            
            return False
    
    def export_rul_file(self, parent_widget: QWidget = None) -> bool:
        """Export rules to .RUL file"""
        try:
            # Check if rules exist
            rules = self.rule_generator.rule_manager.rules
            if not rules:
                error_msg = "No rules to export"
                logger.error(error_msg)
                
                if parent_widget:
                    QMessageBox.warning(parent_widget, "Export Error", error_msg)
                
                return False
            
            # Get file path
            file_path, _ = QFileDialog.getSaveFileName(
                parent_widget,
                "Export RUL File",
                self.config.get("last_directory", ""),
                "RUL Files (*.RUL);;All Files (*)"
            )
            
            if not file_path:
                # User cancelled
                return False
            
            # Ensure file has .RUL extension
            if not file_path.upper().endswith('.RUL'):
                file_path += '.RUL'
            
            # Update last directory
            self.config.update_last_directory(os.path.dirname(file_path))
            
            # Generate and save RUL file
            success = self.rule_generator.save_to_file(file_path)
            
            if not success:
                error_msg = "Failed to save RUL file"
                logger.error(error_msg)
                
                if parent_widget:
                    QMessageBox.critical(parent_widget, "Export Error", error_msg)
                
                return False
            
            # Store current file info
            self.current_file_path = file_path
            
            logger.info(f"Exported RUL file: {file_path} with {len(rules)} rules")
            
            if parent_widget:
                QMessageBox.information(
                    parent_widget,
                    "Export Complete",
                    f"Successfully exported {len(rules)} rules to RUL file:\n{file_path}"
                )
            
            return True
        
        except RuleGeneratorError as e:
            error_msg = f"Error exporting RUL file: {str(e)}"
            logger.error(error_msg)
            
            if parent_widget:
                QMessageBox.critical(parent_widget, "Export Error", error_msg)
            
            return False
        
        except Exception as e:
            error_msg = f"Unexpected error exporting RUL file: {str(e)}"
            logger.error(error_msg)
            
            if parent_widget:
                QMessageBox.critical(parent_widget, "Export Error", error_msg)
            
            return False
    
    def create_rule_editor(self, rule_type: RuleType, parent=None) -> RuleEditorWidget:
        """Create appropriate rule editor for the given rule type"""
        return create_rule_editor(rule_type, parent)
    
    def create_new_rule(self, rule_type: RuleType, name: str = None) -> BaseRule:
        """Create a new rule of the specified type"""
        try:
            if rule_type == RuleType.CLEARANCE:
                return ClearanceRule(
                    name=name or "New_Clearance_Rule",
                    enabled=True,
                    comment="",
                    priority=1,
                    min_clearance=10.0,
                    unit=UnitType.MIL,
                    source_scope=None,
                    target_scope=None
                )
            elif rule_type == RuleType.SHORT_CIRCUIT:
                return ShortCircuitRule(
                    name=name or "New_ShortCircuit_Rule",
                    enabled=True,
                    comment="",
                    priority=1,
                    scope=None
                )
            elif rule_type == RuleType.UNROUTED_NET:
                return UnRoutedNetRule(
                    name=name or "New_UnroutedNet_Rule",
                    enabled=True,
                    comment="",
                    priority=1,
                    scope=None
                )
            else:
                logger.warning(f"Unsupported rule type for creation: {rule_type.value}")
                return None
        
        except Exception as e:
            error_msg = f"Error creating new rule: {str(e)}"
            logger.error(error_msg)
            return None
