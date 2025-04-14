#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Rule Model
=========

Defines the data models for different Altium rule types.
"""

import logging
import re
import uuid
from enum import Enum
from typing import Dict, List, Optional, Union, Tuple, Type, Any

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
        unit_map = {
            'mil': UnitType.MIL, 'mils': UnitType.MIL,
            'mm': UnitType.MM, 'millimeter': UnitType.MM, 'millimeters': UnitType.MM,
            'inch': UnitType.INCH, 'inches': UnitType.INCH, 'in': UnitType.INCH
        }
        
        if unit_str in unit_map:
            return unit_map[unit_str]
        
        valid_units = ', '.join([f"'{u}'" for u in unit_map.keys()])
        raise ValueError(f"Unknown unit type: {unit_str}. Valid unit types are: {valid_units}.")
    
    @staticmethod
    def convert(value: float, from_unit: 'UnitType', to_unit: 'UnitType') -> float:
        """Convert value between unit types"""
        # Conversion factors
        MM_TO_MIL = 39.37007874015748
        INCH_TO_MIL = 1000
        
        # Convert to mils as base unit
        if from_unit == UnitType.MM:
            value_mils = value * MM_TO_MIL
        elif from_unit == UnitType.INCH:
            value_mils = value * INCH_TO_MIL
        else:  # already in mils
            value_mils = value
        
        # Convert from mils to target unit
        if to_unit == UnitType.MM:
            return value_mils / MM_TO_MIL
        elif to_unit == UnitType.INCH:
            return value_mils / INCH_TO_MIL
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
            return f"InNetClass('{self.items[0]}')" if self.items else "All"
        elif self.scope_type == "NetClasses":
            class_queries = [f"InNetClass('{item}')" for item in self.items]
            return ' OR '.join(class_queries)
        elif self.scope_type == "Custom":
            return self.items[0] if self.items else "All"
        else:
            logger.warning(f"Unknown scope type '{self.scope_type}' for RUL format, defaulting to All")
            return "All"
    
    # Alias for to_query_string to prevent refactoring errors
    def to_rul_format(self) -> str:
        """Alias for to_query_string"""
        return self.to_query_string()


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
    
    def get_base_rul_properties(self) -> Dict[str, Any]:
        """Return a dictionary of base properties for RUL format."""
        # Base properties common to most rules, using defaults from sample
        return {
            "SELECTION": "FALSE",
            "LAYER": "UNKNOWN", # Default, can be overridden by specific rules if needed
            "LOCKED": "FALSE",
            "POLYGONOUTLINE": "FALSE",
            "USERROUTED": "TRUE",
            "KEEPOUT": "FALSE",
            "UNIONINDEX": "0", # Use string '0' as seen in sample
            "RULEKIND": self.rule_type.value,
            "NETSCOPE": "DifferentNets", # Default, can be overridden
            "LAYERKIND": "SameLayer", # Default, can be overridden
            "SCOPE1EXPRESSION": "All", # Default, must be overridden by subclasses needing it
            "SCOPE2EXPRESSION": "All", # Default, must be overridden by subclasses needing it
            "NAME": self.name,
            "ENABLED": str(self.enabled).upper(),
            "PRIORITY": str(self.priority),
            "COMMENT": self.comment,
            # Generate 8-char uppercase hex ID, closer to sample
            "UNIQUEID": uuid.uuid4().hex[:8].upper(),
            "DEFINEDBYLOGICALDOCUMENT": "FALSE"
        }

    def to_rul_format(self) -> str:
        """Convert to a single pipe-delimited RUL file line.
        Subclasses should override this, call super().get_base_rul_properties(),
        update the dictionary, and then call self._build_rul_line()."""
        properties = self.get_base_rul_properties()
        # Base rule itself doesn't have enough info, subclasses must implement fully
        logger.warning(f"Direct call to BaseRule.to_rul_format for rule '{self.name}'. Subclass implementation missing?")
        return self._build_rul_line(properties) # Return basic line for safety

    def _build_rul_line(self, properties: Dict[str, Any]) -> str:
        """Helper to build the pipe-delimited string from properties."""
        # Convert all values to string before joining
        # Filter out None or empty string values *after* potential overrides
        line_parts = [f"{key}={str(value)}" for key, value in properties.items() if value is not None and str(value) != '']
        # Sort alphabetically by key for consistency
        line_parts.sort()
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
        
        return cls(
            name=base_rule.name,
            enabled=base_rule.enabled,
            comment=base_rule.comment,
            priority=base_rule.priority,
            min_clearance=data.get("min_clearance", 10.0),
            unit=UnitType(data.get("unit", "mil")),
            source_scope=RuleScope.from_dict(data.get("source_scope", {})),
            target_scope=RuleScope.from_dict(data.get("target_scope", {}))
        )
    
    def to_rul_format(self) -> str:
        """Convert to a single pipe-delimited RUL file line for Clearance"""
        properties = self.get_base_rul_properties() # Get base properties dict
        properties.update({
            "SCOPE1EXPRESSION": self.source_scope.to_rul_format(),
            "SCOPE2EXPRESSION": self.target_scope.to_rul_format(),
            "GAP": f"{self.min_clearance}{self.unit.value}",
            "GENERICCLEARANCE": f"{self.min_clearance}{self.unit.value}", # Add missing
            # Add other missing ones with defaults or from instance if available
            # TODO: Consider adding these attributes to the class if they need to be configurable
            "IGNOREPADTOPADCLEARANCEINFOOTPRINT": "FALSE", # Default based on sample majority
            "OBJECTCLEARANCES": " ", # Default based on sample (space or empty)
            # NETSCOPE and LAYERKIND use defaults from base, override if needed
            # "NETSCOPE": "DifferentNets", # Example override if needed
        })
        return self._build_rul_line(properties) # Build the final string


class SingleScopeRule(BaseRule):
    """Base class for rules with a single scope"""
    
    def __init__(self, rule_type: RuleType, name: str, enabled: bool = True, 
                 comment: str = "", priority: int = 1, scope: RuleScope = None):
        """Initialize rule with a single scope"""
        super().__init__(rule_type, name, enabled, comment, priority)
        self.scope = scope or RuleScope("All")
    
    def to_dict(self) -> Dict:
        """Convert to dictionary"""
        data = super().to_dict()
        data.update({
            "scope": self.scope.to_dict()
        })
        return data


class ShortCircuitRule(SingleScopeRule):
    """Short circuit rule"""
    
    def __init__(self, name: str, enabled: bool = True, comment: str = "", 
                 priority: int = 1, scope: RuleScope = None):
        """Initialize short circuit rule"""
        super().__init__(RuleType.SHORT_CIRCUIT, name, enabled, comment, priority, scope)
    
    @classmethod
    def from_dict(cls, data: Dict) -> 'ShortCircuitRule':
        """Create from dictionary"""
        base_rule = BaseRule.from_dict(data)
        
        return cls(
            name=base_rule.name,
            enabled=base_rule.enabled,
            comment=base_rule.comment,
            priority=base_rule.priority,
            scope=RuleScope.from_dict(data.get("scope", {}))
        )
    
    def to_rul_format(self) -> str:
        """Convert to a single pipe-delimited RUL file line for ShortCircuit"""
        properties = self.get_base_rul_properties() # Get base properties dict
        properties.update({
            "SCOPE1EXPRESSION": self.scope.to_rul_format(),
            "SCOPE2EXPRESSION": self.scope.to_rul_format(), # Short circuit uses same scope twice
            "ALLOWED": "FALSE",
            # NETSCOPE and LAYERKIND use defaults from base
        })
        return self._build_rul_line(properties)


class UnRoutedNetRule(SingleScopeRule):
    """Unrouted net rule"""
    
    def __init__(self, name: str, enabled: bool = True, comment: str = "", 
                 priority: int = 1, scope: RuleScope = None):
        """Initialize unrouted net rule"""
        super().__init__(RuleType.UNROUTED_NET, name, enabled, comment, priority, scope)
    
    @classmethod
    def from_dict(cls, data: Dict) -> 'UnRoutedNetRule':
        """Create from dictionary"""
        base_rule = BaseRule.from_dict(data)
        
        return cls(
            name=base_rule.name,
            enabled=base_rule.enabled,
            comment=base_rule.comment,
            priority=base_rule.priority,
            scope=RuleScope.from_dict(data.get("scope", {}))
        )
    
    def to_rul_format(self) -> str:
        """Convert to a single pipe-delimited RUL file line for UnRoutedNet"""
        properties = self.get_base_rul_properties() # Get base properties dict
        properties.update({
            "SCOPE1EXPRESSION": self.scope.to_rul_format(),
            "SCOPE2EXPRESSION": "All", # Use default 'All' for scope 2 in this rule type
            "CHECKBADCONNECTIONS": "TRUE",
            # NETSCOPE and LAYERKIND use defaults from base
        })
        return self._build_rul_line(properties)


class RuleManager:
    """Manages a collection of rules"""
    
    def __init__(self):
        """Initialize rule manager"""
        self.rules = []
    
    def add_rule(self, rule: BaseRule):
        """Add a rule to the collection"""
        self.rules.append(rule)
        logger.info(f"Added rule: {rule.name} ({rule.rule_type.value})")
    
    def remove_rule(self, rule_index: int) -> bool:
        """Remove a rule by index"""
        if 0 <= rule_index < len(self.rules):
            rule = self.rules.pop(rule_index)
            logger.info(f"Removed rule: {rule.name} ({rule.rule_type.value})")
            return True
        return False

    def delete_rule(self, rule_name: str) -> bool:
        """Remove a rule by its name."""
        initial_length = len(self.rules)
        self.rules = [rule for rule in self.rules if rule.name != rule_name]
        if len(self.rules) < initial_length:
            logger.info(f"Deleted rule: {rule_name}")
            return True
        logger.warning(f"Rule not found for deletion: {rule_name}")
        return False

    def get_rule_index(self, rule_name: str) -> Optional[int]:
        """Get the index of a rule by its name."""
        for i, rule in enumerate(self.rules):
            if rule.name == rule_name:
                return i
        return None

    def get_rules_by_type(self, rule_type: RuleType) -> List[BaseRule]:
        """Get all rules of a specific type"""
        return [rule for rule in self.rules if rule.rule_type == rule_type]
    
    def to_rul_format(self) -> str:
        """Convert all rules to RUL file format (pipe-delimited lines)"""
        rul_lines = [rule.to_rul_format() for rule in self.rules]
        return "\r\n".join(rul_lines)
    
    def export_rules_to_file(self, file_path: str):
        """Export all rules to a .RUL file."""
        try:
            rul_content = self.to_rul_format()
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(rul_content)
            logger.info(f"Successfully exported {len(self.rules)} rules to {file_path}")
        except IOError as e:
            logger.error(f"Error writing RUL file to {file_path}: {e}", exc_info=True)
            raise
        except Exception as e:
            logger.error(f"An unexpected error occurred during RUL export: {e}", exc_info=True)
            raise

    def generate_pivot_data(self):
        """Placeholder for generating pivot data from rules."""
        logger.warning("generate_pivot_data is not yet implemented.")
        return None

    def from_rul_content(self, rul_content: str) -> bool:
        """Parse rules from RUL file content."""
        try:
            self.rules = []
            rule_blocks = self._extract_rule_blocks(rul_content)
            
            if not rule_blocks:
                logger.error("No rule blocks found in RUL content.")
                return False
            
            logger.info(f"Found {len(rule_blocks)} rule blocks in RUL content")
            
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
            logger.error(f"Error parsing RUL content: {str(e)}")
            return False
    
    def _extract_rule_blocks(self, rul_content: str) -> List[str]:
        """Extract rule blocks from RUL content"""
        pattern = r'Rule\s*{[^}]*}'
        return re.findall(pattern, rul_content, re.DOTALL)
    
    def _parse_rule_block(self, block: str) -> Optional[BaseRule]:
        """Parse a rule block into a rule object"""
        try:
            properties = self._extract_rule_properties(block)
            if not properties.get('Name') or not properties.get('RuleKind'):
                logger.warning("Rule block missing required properties (Name or RuleKind)")
                return None
            
            rule_factories = {
                RuleType.CLEARANCE.value: self._create_clearance_rule,
                RuleType.SHORT_CIRCUIT.value: self._create_short_circuit_rule,
                RuleType.UNROUTED_NET.value: self._create_unrouted_net_rule
            }
            
            rule_kind = properties.get('RuleKind')
            if rule_kind in rule_factories:
                return rule_factories[rule_kind](properties)
            
            logger.warning(f"Unsupported rule kind: {rule_kind}")
            return None
        
        except Exception as e:
            logger.error(f"Error parsing rule block: {str(e)}")
            return None
    
    def _extract_rule_properties(self, block: str) -> Dict[str, str]:
        """Extract all properties from a rule block"""
        properties = {}
        # Extract property pattern
        property_pattern = r'\s+(\w+)\s*=\s*[\'"]?([^\'"\n}]*)[\'"]?\s*'
        
        for match in re.finditer(property_pattern, block):
            key, value = match.groups()
            properties[key] = value.strip()
        
        return properties
    
    def _create_clearance_rule(self, properties: Dict[str, str]) -> Optional[ClearanceRule]:
        """Create a clearance rule from properties"""
        try:
            name = properties.get('Name', '')
            enabled = properties.get('Enabled', 'TRUE').upper() == 'TRUE'
            comment = properties.get('Comment', '')
            priority = int(properties.get('Priority', '1'))
            
            # Parse clearance value and unit
            clearance_str = properties.get('MinimumClearance', '10.0')
            try:
                min_clearance = float(clearance_str)
            except ValueError:
                logger.warning(f"Invalid MinimumClearance value: {clearance_str}")
                min_clearance = 10.0
            
            unit_str = properties.get('MinimumClearanceType', 'mil')
            try:
                unit = UnitType.from_string(unit_str)
            except ValueError:
                logger.warning(f"Invalid unit type: {unit_str}, defaulting to MIL")
                unit = UnitType.MIL
            
            # Parse scopes
            source_scope = self._parse_scope(properties.get('SourceScope', ''))
            target_scope = self._parse_scope(properties.get('TargetScope', ''))
            
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
            logger.error(f"Error creating clearance rule: {str(e)}")
            return None
    
    def _create_short_circuit_rule(self, properties: Dict[str, str]) -> Optional[ShortCircuitRule]:
        """Create a short circuit rule from properties"""
        try:
            name = properties.get('Name', '')
            enabled = properties.get('Enabled', 'TRUE').upper() == 'TRUE'
            comment = properties.get('Comment', '')
            priority = int(properties.get('Priority', '1'))
            
            scope = self._parse_scope(properties.get('Scope', ''))
            
            return ShortCircuitRule(
                name=name,
                enabled=enabled,
                comment=comment,
                priority=priority,
                scope=scope
            )
        except Exception as e:
            logger.error(f"Error creating short circuit rule: {str(e)}")
            return None
    
    def _create_unrouted_net_rule(self, properties: Dict[str, str]) -> Optional[UnRoutedNetRule]:
        """Create an unrouted net rule from properties"""
        try:
            name = properties.get('Name', '')
            enabled = properties.get('Enabled', 'TRUE').upper() == 'TRUE'
            comment = properties.get('Comment', '')
            priority = int(properties.get('Priority', '1'))
            
            scope = self._parse_scope(properties.get('Scope', ''))
            
            return UnRoutedNetRule(
                name=name,
                enabled=enabled,
                comment=comment,
                priority=priority,
                scope=scope
            )
        except Exception as e:
            logger.error(f"Error creating unrouted net rule: {str(e)}")
            return None
    
    def _parse_scope(self, scope_str: str) -> RuleScope:
        """Parse a scope string into a RuleScope object"""
        if not scope_str or scope_str.strip().lower() == 'all':
            return RuleScope("All")
        
        # Check for InNetClass pattern
        net_class_match = re.search(r'InNetClass\([\'"]([^\'"]*)[\'"]', scope_str)
        if net_class_match:
            net_class_name = net_class_match.group(1).strip()
            if net_class_name and re.match(r'^[a-zA-Z0-9_\-]+$', net_class_name):
                return RuleScope("NetClass", [net_class_name])
        
        # Check for multiple net classes (OR pattern)
        if ' OR ' in scope_str:
            class_matches = re.findall(r'InNetClass\([\'"]([^\'"]*)[\'"]', scope_str)
            if class_matches:
                return RuleScope("NetClasses", class_matches)
        
        # Check for quoted custom scope
        quoted_match = re.search(r'[\'"]([^\'"]*)[\'"]', scope_str)
        if quoted_match:
            return RuleScope("Custom", [quoted_match.group(1)])
        
        # Default to custom with the provided string
        return RuleScope("Custom", [scope_str])
    
    def to_dict(self) -> List[Dict]:
        """Convert all rules to dictionary format"""
        return [rule.to_dict() for rule in self.rules]
