#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Rule Generator Service
=====================

Service for generating Altium .RUL files from rule models.
Also handles parsing .RUL files back into rule models.
"""

import os
import re
import logging
from typing import Dict, List, Optional, Union, Tuple, Any

from models.rule_model import (
    UnitType, RuleType, RuleScope, RuleManager,
    BaseRule, ClearanceRule, ShortCircuitRule, UnRoutedNetRule
)

logger = logging.getLogger(__name__)

class RuleGeneratorError(Exception):
    """Exception raised for errors during rule generation or parsing"""
    pass

class RuleGenerator:
    """Service for generating and parsing Altium .RUL files"""
    
    def __init__(self):
        """Initialize rule generator"""
        self.rule_manager = RuleManager()
    
    def add_rule(self, rule: BaseRule):
        """Add a rule to the rule manager"""
        self.rule_manager.add_rule(rule)
    
    def add_rules(self, rules: List[BaseRule]):
        """Add multiple rules to the rule manager"""
        for rule in rules:
            self.rule_manager.add_rule(rule)
    
    def clear_rules(self):
        """Clear all rules from the rule manager"""
        self.rule_manager.rules = []
        logger.info("Cleared all rules from rule manager")
    
    def get_rules_by_type(self, rule_type: RuleType) -> List[BaseRule]:
        """Get all rules of a specific type"""
        return self.rule_manager.get_rules_by_type(rule_type)
    
    def generate_rul_content(self) -> str:
        """Generate .RUL file content from rules"""
        try:
            # Generate RUL content
            rul_content = self.rule_manager.to_rul_format()
            logger.info(f"Generated RUL content with {len(self.rule_manager.rules)} rules")
            
            return rul_content
        
        except Exception as e:
            error_msg = f"Error generating RUL content: {str(e)}"
            logger.error(error_msg)
            raise RuleGeneratorError(error_msg)
    
    def save_to_file(self, file_path: str) -> bool:
        """Save .RUL file to the specified path"""
        try:
            # Generate RUL content
            rul_content = self.generate_rul_content()
            
            # Write to file
            with open(file_path, 'w') as f:
                f.write(rul_content)
            
            logger.info(f"Saved RUL file to: {file_path}")
            return True
        
        except Exception as e:
            error_msg = f"Error saving RUL file: {str(e)}"
            logger.error(error_msg)
            raise RuleGeneratorError(error_msg)
    
    def parse_rul_file(self, file_path: str) -> bool:
        """Parse .RUL file and populate rule manager"""
        try:
            # Check if file exists
            if not os.path.exists(file_path):
                raise RuleGeneratorError(f"File not found: {file_path}")
            
            # Read file content
            with open(file_path, 'r') as f:
                rul_content = f.read()
            
            # Parse RUL content
            success = self.parse_rul_content(rul_content)
            
            if success:
                logger.info(f"Successfully parsed RUL file: {file_path}")
            else:
                logger.warning(f"Parsed RUL file with warnings: {file_path}")
            
            return success
        
        except Exception as e:
            error_msg = f"Error parsing RUL file: {str(e)}"
            logger.error(error_msg)
            raise RuleGeneratorError(error_msg)
    
    def parse_rul_content(self, rul_content: str) -> bool:
        """Parse .RUL file content and populate rule manager"""
        try:
            # Clear existing rules
            self.clear_rules()
            
            # Extract rule blocks
            rule_blocks = self._extract_rule_blocks(rul_content)
            
            if not rule_blocks:
                logger.warning("No rule blocks found in RUL content")
                return False
            
            logger.info(f"Found {len(rule_blocks)} rule blocks in RUL content")
            
            # Parse each rule block
            for block in rule_blocks:
                rule = self._parse_rule_block(block)
                if rule:
                    self.rule_manager.add_rule(rule)
            
            return True
        
        except Exception as e:
            error_msg = f"Error parsing RUL content: {str(e)}"
            logger.error(error_msg)
            raise RuleGeneratorError(error_msg)
    
    def _extract_rule_blocks(self, rul_content: str) -> List[str]:
        """Extract rule blocks from RUL content"""
        # Pattern to match a rule block: starts with "Rule" and ends with "}"
        pattern = r'Rule\s*{[^}]*}'
        
        # Find all matches
        blocks = re.findall(pattern, rul_content, re.DOTALL)
        
        return blocks
    
    def _parse_rule_block(self, block: str) -> Optional[BaseRule]:
        """Parse a rule block into a rule object"""
        try:
            # Extract rule properties
            name = self._extract_property(block, 'Name')
            enabled_str = self._extract_property(block, 'Enabled')
            enabled = enabled_str.lower() == 'true' if enabled_str else True
            comment = self._extract_property(block, 'Comment')
            priority_str = self._extract_property(block, 'Priority')
            priority = int(priority_str) if priority_str and priority_str.isdigit() else 1
            rule_kind = self._extract_property(block, 'RuleKind')
            
            if not name or not rule_kind:
                logger.warning("Rule block missing required properties (Name or RuleKind)")
                return None
            
            # Create rule object based on rule kind
            if rule_kind == RuleType.CLEARANCE.value:
                return self._parse_clearance_rule(
                    block, name, enabled, comment, priority
                )
            elif rule_kind == RuleType.SHORT_CIRCUIT.value:
                return self._parse_short_circuit_rule(
                    block, name, enabled, comment, priority
                )
            elif rule_kind == RuleType.UNROUTED_NET.value:
                return self._parse_unrouted_net_rule(
                    block, name, enabled, comment, priority
                )
            else:
                logger.warning(f"Unsupported rule kind: {rule_kind}")
                return None
        
        except Exception as e:
            logger.error(f"Error parsing rule block: {str(e)}")
            return None
    
    def _parse_clearance_rule(self, block: str, name: str, enabled: bool, 
                            comment: str, priority: int) -> Optional[ClearanceRule]:
        """Parse a clearance rule block"""
        try:
            # Extract clearance properties
            min_clearance_str = self._extract_property(block, 'MinimumClearance')
            min_clearance_type = self._extract_property(block, 'MinimumClearanceType')
            source_scope_str = self._extract_property(block, 'SourceScope')
            target_scope_str = self._extract_property(block, 'TargetScope')
            
            if not min_clearance_str:
                logger.warning("Clearance rule missing MinimumClearance property")
                return None
            
            # Parse min clearance
            try:
                min_clearance = float(min_clearance_str)
            except ValueError:
                logger.warning(f"Invalid MinimumClearance value: {min_clearance_str}")
                min_clearance = 10.0
            
            # Parse unit type
            try:
                unit = UnitType.from_string(min_clearance_type) if min_clearance_type else UnitType.MIL
            except ValueError:
                logger.warning(f"Invalid MinimumClearanceType: {min_clearance_type}, using MIL")
                unit = UnitType.MIL
            
            # Parse scopes
            source_scope = self._parse_scope(source_scope_str) if source_scope_str else RuleScope("All")
            target_scope = self._parse_scope(target_scope_str) if target_scope_str else RuleScope("All")
            
            # Create rule
            return ClearanceRule(
                name=name,
                enabled=enabled,
                comment=comment,
                priority=priority,
                min_clearance=min_clearance,
                unit=unit,
                source_scope=source_scope,
                target_scope=target_scope
            )
        
        except Exception as e:
            logger.error(f"Error parsing clearance rule: {str(e)}")
            return None
    
    def _parse_short_circuit_rule(self, block: str, name: str, enabled: bool, 
                                comment: str, priority: int) -> Optional[ShortCircuitRule]:
        """Parse a short circuit rule block"""
        try:
            # Extract scope
            scope_str = self._extract_property(block, 'Scope')
            
            # Parse scope
            scope = self._parse_scope(scope_str) if scope_str else RuleScope("All")
            
            # Create rule
            return ShortCircuitRule(
                name=name,
                enabled=enabled,
                comment=comment,
                priority=priority,
                scope=scope
            )
        
        except Exception as e:
            logger.error(f"Error parsing short circuit rule: {str(e)}")
            return None
    
    def _parse_unrouted_net_rule(self, block: str, name: str, enabled: bool, 
                               comment: str, priority: int) -> Optional[UnRoutedNetRule]:
        """Parse an unrouted net rule block"""
        try:
            # Extract scope
            scope_str = self._extract_property(block, 'Scope')
            
            # Parse scope
            scope = self._parse_scope(scope_str) if scope_str else RuleScope("All")
            
            # Create rule
            return UnRoutedNetRule(
                name=name,
                enabled=enabled,
                comment=comment,
                priority=priority,
                scope=scope
            )
        
        except Exception as e:
            logger.error(f"Error parsing unrouted net rule: {str(e)}")
            return None
    
    def _extract_property(self, block: str, property_name: str) -> str:
        """Extract a property value from a rule block"""
        # Pattern to match property: "PropertyName = 'Value'" or "PropertyName = Value"
        pattern = r'\s+' + re.escape(property_name) + r'\s*=\s*[\'"]?([^\'"\n}]*)[\'"]?\s*'
        
        match = re.search(pattern, block)
        if match:
            return match.group(1).strip()
        return ""
    
    def _parse_scope(self, scope_str: str) -> RuleScope:
        """Parse a scope string into a RuleScope object"""
        if not scope_str or scope_str == "All":
            return RuleScope("All")
        
        # Check for InNetClass pattern
        net_class_match = re.match(r'InNetClass\([\'"]([^\'"]*)[\'"]', scope_str)
        if net_class_match:
            return RuleScope("NetClass", [net_class_match.group(1)])
        
        # Check for InNetClasses pattern
        net_classes_match = re.match(r'InNetClasses\([\'"]([^\'"]*)[\'"]', scope_str)
        if net_classes_match:
            classes = net_classes_match.group(1).split(';')
            return RuleScope("NetClasses", classes)
        
        # Check for simple quoted string (custom scope)
        quoted_match = re.match(r'[\'"]([^\'"]*)[\'"]', scope_str)
        if quoted_match:
            items = quoted_match.group(1).split(';')
            return RuleScope("Custom", items)
        
        # Default to All
        logger.warning(f"Could not parse scope: {scope_str}, using All")
        return RuleScope("All")
