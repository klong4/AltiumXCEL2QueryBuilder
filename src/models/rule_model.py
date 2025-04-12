#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Rule Model
=========

Defines the data models for different Altium rule types.
"""

import logging
import re
import uuid # Import the uuid module
from enum import Enum
from typing import Dict, List, Optional, Union, Tuple

logger = logging.getLogger(__name__)

class UnitType(Enum):
    """Unit types for measurements"""
    MIL = "mil"
    MM = "mm"
    INCH = "inch"
    
    @staticmethod
    def from_string(unit_str: str) -> 'UnitType':
        """Convert string to UnitType"""
        unit_str = unit_str.lower().strip()
        if unit_str in ('mil', 'mils'):
            return UnitType.MIL
        elif unit_str in ('mm', 'millimeter', 'millimeters'):
            return UnitType.MM
        elif unit_str in ('inch', 'inches', 'in'):
            return UnitType.INCH
        else:
            raise ValueError(f"Unknown unit type: {unit_str}. Valid unit types are: 'mil', 'mils', 'mm', 'millimeter', 'millimeters', 'inch', 'inches', 'in'.")
    
    @staticmethod
    def convert(value: float, from_unit: 'UnitType', to_unit: 'UnitType') -> float:
        """Convert value between unit types"""
        # Convert to mils as base unit
        MM_TO_MIL_CONVERSION = 39.37007874015748  # More precise conversion factor
        if from_unit == UnitType.MM:
            value_mils = value * MM_TO_MIL_CONVERSION
        elif from_unit == UnitType.INCH:
            value_mils = value * 1000
        else:  # already in mils
            value_mils = value
        
        # Convert from mils to target unit
        if to_unit == UnitType.MM:
            return value_mils / 39.3701
        elif to_unit == UnitType.INCH:
            return value_mils / 1000
        else:  # target is mils
            return value_mils


class RuleType(Enum):
    """Types of Altium design rules"""
    CLEARANCE = "Clearance"
    SHORT_CIRCUIT = "ShortCircuit"
    UNROUTED_NET = "UnroutedNet"
    UNCONNECTED_PIN = "UnconnectedPin"
    MODIFIED_POLYGON = "ModifiedPolygon"
    CREEPAGE_DISTANCE = "CreepageDistance"


class RuleScope:
    """Defines the scope of a rule"""
    
    def __init__(self, scope_type: str, items: List[str] = None):
        """Initialize rule scope"""
        self.scope_type = scope_type  # All, NetClasses, NetClass, Custom
        self.items = items or []
    
    def to_dict(self) -> Dict:
        """Convert to dictionary"""
        return {
            "scope_type": self.scope_type,
            "items": self.items
        }
    
    @classmethod
    def from_dict(cls, data: Dict) -> 'RuleScope':
        """Create from dictionary"""
        return cls(
            scope_type=data.get("scope_type", "All"),
            items=data.get("items", [])
        )
    
    def to_query_string(self) -> str:
        """Convert to RUL file format query string"""
        if self.scope_type == "All":
            return "All"
        elif self.scope_type == "NetClass":
            # Altium query for a single net class
            return f"InNetClass('{self.items[0]}')" if self.items else "All"
        elif self.scope_type == "NetClasses":
            # Altium query for multiple net classes (ORed together)
            class_queries = [f"InNetClass('{item}')" for item in self.items]
            return ' OR '.join(class_queries)
        elif self.scope_type == "Custom":
            # Assume custom scope is already a valid Altium query string
            return self.items[0] if self.items else "All"
        else:
            logger.warning(f"Unknown scope type '{self.scope_type}' for RUL format, defaulting to All")
            return "All"


class BaseRule:
    """Base class for all rules"""
    
    def __init__(self, rule_type: RuleType, name: str, enabled: bool = True, 
                 comment: str = "", priority: int = 1):
        """Initialize base rule"""
        self.rule_type = rule_type
        self.name = name
        self.enabled = enabled
        self.comment = comment
        self.priority = priority
    
    def to_dict(self) -> Dict:
        """Convert to dictionary"""
        return {
            "rule_type": self.rule_type.value,
            "name": self.name,
            "enabled": self.enabled,
            "comment": self.comment,
            "priority": self.priority
        }
    
    @classmethod
    def from_dict(cls, data: Dict) -> 'BaseRule':
        """Create from dictionary"""
        return cls(
            rule_type=RuleType(data.get("rule_type")),
            name=data.get("name", ""),
            enabled=data.get("enabled", True),
            comment=data.get("comment", ""),
            priority=data.get("priority", 1)
        )
    
    def to_rul_format(self) -> str:
        """Convert to a single pipe-delimited RUL file line"""
        # Base properties common to all rules
        properties = {
            "NAME": self.name,
            "ENABLED": str(self.enabled).upper(),  # TRUE/FALSE
            "PRIORITY": str(self.priority),
            "COMMENT": self.comment,
            "RULEKIND": self.rule_type.value,
            # Generate a unique ID for the rule using UUID4
            "UNIQUEID": str(uuid.uuid4()).upper() # Generate and format UUID
        }
        # Subclasses will add their specific properties
        return self._build_rul_line(properties)

    def _build_rul_line(self, properties: Dict[str, str]) -> str:
        """Helper to build the pipe-delimited string from properties"""
        # Filter out empty values, except for potentially required ones if any
        line_parts = [f"{key}={value}" for key, value in properties.items() if value is not None and value != ""]
        return '|'.join(line_parts)


class ClearanceRule(BaseRule):
    """Electrical clearance rule"""
    
    def __init__(self, name: str, enabled: bool = True, comment: str = "", 
                 priority: int = 1, min_clearance: float = 10.0, 
                 unit: UnitType = UnitType.MIL, source_scope: RuleScope = None, 
                 target_scope: RuleScope = None):
        """Initialize clearance rule"""
        super().__init__(RuleType.CLEARANCE, name, enabled, comment, priority)
        self.min_clearance = min_clearance
        self.unit = unit
        self.source_scope = source_scope or RuleScope("All")
        self.target_scope = target_scope or RuleScope("All")
    
    def to_dict(self) -> Dict:
        """Convert to dictionary"""
        data = super().to_dict()
        data.update({
            "min_clearance": self.min_clearance,
            "unit": self.unit.value,
            "source_scope": self.source_scope.to_dict(),
            "target_scope": self.target_scope.to_dict()
        })
        return data
    
    @classmethod
    def from_dict(cls, data: Dict) -> 'ClearanceRule':
        """Create from dictionary"""
        base_rule = super(ClearanceRule, cls).from_dict(data)
        
        source_scope_data = data.get("source_scope", {})
        target_scope_data = data.get("target_scope", {})
        
        return cls(
            name=base_rule.name,
            enabled=base_rule.enabled,
            comment=base_rule.comment,
            priority=base_rule.priority,
            min_clearance=data.get("min_clearance", 10.0),
            unit=UnitType(data.get("unit", "mil")),
            source_scope=RuleScope.from_dict(source_scope_data),
            target_scope=RuleScope.from_dict(target_scope_data)
        )
    
    def to_rul_format(self) -> str:
        """Convert to a single pipe-delimited RUL file line for Clearance"""
        # Start with base properties
        properties = {
            "NAME": self.name,
            "ENABLED": str(self.enabled).upper(),
            "PRIORITY": str(self.priority),
            "COMMENT": self.comment,
            "RULEKIND": self.rule_type.value,
            # Clearance specific properties
            "SCOPE1EXPRESSION": self.source_scope.to_rul_format(),
            "SCOPE2EXPRESSION": self.target_scope.to_rul_format(),
            # Use GAP for the clearance value, append unit
            "GAP": f"{self.min_clearance}{self.unit.value}",
            # Add other common clearance defaults if needed
            "NETSCOPE": "DifferentNets", # Common default
            "LAYERKIND": "SameLayer",   # Common default
        }
        return self._build_rul_line(properties)


class ShortCircuitRule(BaseRule):
    """Short circuit rule"""
    
    def __init__(self, name: str, enabled: bool = True, comment: str = "", 
                 priority: int = 1, scope: RuleScope = None):
        """Initialize short circuit rule"""
        super().__init__(RuleType.SHORT_CIRCUIT, name, enabled, comment, priority)
        self.scope = scope or RuleScope("All")
    
    def to_dict(self) -> Dict:
        """Convert to dictionary"""
        data = super().to_dict()
        data.update({
            "scope": self.scope.to_dict()
        })
        return data
    
    @classmethod
    def from_dict(cls, data: Dict) -> 'ShortCircuitRule':
        """Create from dictionary"""
        base_rule = super(ShortCircuitRule, cls).from_dict(data)
        
        scope_data = data.get("scope", {})
        
        return cls(
            name=base_rule.name,
            enabled=base_rule.enabled,
            comment=base_rule.comment,
            priority=base_rule.priority,
            scope=RuleScope.from_dict(scope_data)
        )
    
    def to_rul_format(self) -> str:
        """Convert to a single pipe-delimited RUL file line for ShortCircuit"""
        properties = {
            "NAME": self.name,
            "ENABLED": str(self.enabled).upper(),
            "PRIORITY": str(self.priority),
            "COMMENT": self.comment,
            "RULEKIND": self.rule_type.value,
            # ShortCircuit specific properties
            "SCOPE1EXPRESSION": self.scope.to_rul_format(), # Use SCOPE1EXPRESSION for single scope rules
            "SCOPE2EXPRESSION": self.scope.to_rul_format(), # Often the same for ShortCircuit
            "ALLOWED": "FALSE", # Common default for ShortCircuit
        }
        return self._build_rul_line(properties)


class UnRoutedNetRule(BaseRule):
    """Unrouted net rule"""
    
    def __init__(self, name: str, enabled: bool = True, comment: str = "", 
                 priority: int = 1, scope: RuleScope = None):
        """Initialize unrouted net rule"""
        super().__init__(RuleType.UNROUTED_NET, name, enabled, comment, priority)
        self.scope = scope or RuleScope("All")
    
    def to_dict(self) -> Dict:
        """Convert to dictionary"""
        data = super().to_dict()
        data.update({
            "scope": self.scope.to_dict()
        })
        return data
    
    @classmethod
    def from_dict(cls, data: Dict) -> 'UnRoutedNetRule':
        """Create from dictionary"""
        base_rule = super(UnRoutedNetRule, cls).from_dict(data)
        
        scope_data = data.get("scope", {})
        
        return cls(
            name=base_rule.name,
            enabled=base_rule.enabled,
            comment=base_rule.comment,
            priority=base_rule.priority,
            scope=RuleScope.from_dict(scope_data)
        )
    
    def to_rul_format(self) -> str:
        """Convert to a single pipe-delimited RUL file line for UnRoutedNet"""
        properties = {
            "NAME": self.name,
            "ENABLED": str(self.enabled).upper(),
            "PRIORITY": str(self.priority),
            "COMMENT": self.comment,
            "RULEKIND": self.rule_type.value,
            # UnRoutedNet specific properties
            "SCOPE1EXPRESSION": self.scope.to_rul_format(), # Use SCOPE1EXPRESSION
            "CHECKBADCONNECTIONS": "TRUE", # Common default
        }
        return self._build_rul_line(properties)


# Add additional rule types later...


class RuleManager:
    """Manages a collection of rules"""
    
    def __init__(self):
        """Initialize rule manager"""
        self.rules = []
    
    def add_rule(self, rule: BaseRule):
        """Add a rule to the collection"""
        self.rules.append(rule)
        logger.info(f"Added rule: {rule.name} ({rule.rule_type.value})")
    
    def remove_rule(self, rule_index: int):
        """Remove a rule by index"""
        if 0 <= rule_index < len(self.rules):
            rule = self.rules.pop(rule_index)
            logger.info(f"Removed rule: {rule.name} ({rule.rule_type.value})")
            return True
        return False
    
    def get_rules_by_type(self, rule_type: RuleType) -> List[BaseRule]:
        """Get all rules of a specific type"""
        return [rule for rule in self.rules if rule.rule_type == rule_type]
    
    def to_rul_format(self) -> str:
        """Convert all rules to RUL file format (pipe-delimited lines)"""
        # Generate one line per rule
        rul_lines = [rule.to_rul_format() for rule in self.rules]
        # Join lines with newline characters appropriate for the target system (e.g., \r\n for Windows)
        return "\r\n".join(rul_lines)
    
    def from_rul_content(self, rul_content: str) -> bool:
        """Parse rules from RUL file content (pipe-delimited lines)

        Args:
            rul_content (str): The content of a .RUL file

        Returns:
            bool: True if parsing was successful, False otherwise
        """
        try:
            # Clear existing rules
            self.rules = []
            
            # Extract rule blocks
            rule_blocks = self._extract_rule_blocks(rul_content)
            
            if not rule_blocks:
                logger.error("No rule blocks found in RUL content. Ensure the RUL file contains valid rule definitions.")
                return False
            
            logger.info(f"Found {len(rule_blocks)} rule blocks in RUL content")
            
            # Parse each rule block
            successful_rules = 0
            for block in rule_blocks:
                rule = self._parse_rule_block(block)
                if rule:
                    self.add_rule(rule)
                    successful_rules += 1
            
            if successful_rules == 0:
                logger.warning("No valid rules were found in the RUL content")
                return False
                
            logger.info(f"Successfully parsed {successful_rules} rules from RUL content")
            return True
            
        except Exception as e:
            error_msg = f"Error parsing RUL content: {str(e)}"
            logger.error(error_msg)
            return False
    
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
                try:
                    unit = UnitType.from_string(min_clearance_type) if min_clearance_type else UnitType.MIL
                except ValueError:
                    logger.warning(f"Invalid MinimumClearanceType: {min_clearance_type}, defaulting to MIL")
                    unit = UnitType.MIL
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
            value = match.group(1).strip()
            # Sanitize the value to handle special characters or escape sequences
            sanitized_value = re.sub(r'[^\w\s\-.,;:]', '', value)  # Allow alphanumeric, spaces, and common symbols
            return sanitized_value
        return ""
    
    def _parse_scope(self, scope_str: str) -> RuleScope:
        # Check for InNetClass pattern
        net_class_match = re.match(r'InNetClass\([\'"]([^\'"]*)[\'"]\)', scope_str)
        if net_class_match:
            net_class_name = net_class_match.group(1).strip()
            # Validate net class name for invalid characters
            if re.match(r'^[a-zA-Z0-9_\-]+$', net_class_name):  # Allow alphanumeric, underscores, and hyphens
                return RuleScope("NetClass", [net_class_name])
            else:
                logger.warning(f"Invalid net class name: {net_class_name}")
        else:
            logger.warning(f"Invalid InNetClass format: {scope_str}")
        # Check for InNetClass pattern
        net_class_match = re.match(r'InNetClass\([\'"]([^\'"]*)[\'"]', scope_str)
        if net_class_match:
            return RuleScope("NetClass", [net_class_match.group(1)])
        
        # Check for InNetClasses pattern
        net_classes_match = re.match(r'InNetClasses\([\'"]([^\'"]*)[\'"]', scope_str)
        if net_classes_match:
            classes = net_classes_match.group(1).split(';')
            return RuleScope("NetClasses", classes)
        
        logger.warning(f"Could not parse scope: '{scope_str}'. Defaulting to 'All'. Ensure the scope string is valid.")
        quoted_match = re.match(r'[\'"]([^\'"]*)[\'"]', scope_str)
        if quoted_match:
            items = quoted_match.group(1).split(';')
            return RuleScope("Custom", items)
        
        # Default to All
        logger.warning(f"Could not parse scope: {scope_str}, using All")
        return RuleScope("All")
    
    def to_dict(self) -> List[Dict]:
        """Convert all rules to dictionary format"""
        return [rule.to_dict() for rule in self.rules]
