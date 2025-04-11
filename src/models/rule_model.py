#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Rule Model
=========

Defines the data models for different Altium rule types.
"""

import logging
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
            raise ValueError(f"Unknown unit type: {unit_str}")
    
    @staticmethod
    def convert(value: float, from_unit: 'UnitType', to_unit: 'UnitType') -> float:
        """Convert value between unit types"""
        # Convert to mils as base unit
        if from_unit == UnitType.MM:
            value_mils = value * 39.3701
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
    
    def to_rul_format(self) -> str:
        """Convert to RUL file format"""
        if self.scope_type == "All":
            return "All"
        elif self.scope_type == "NetClasses":
            return f"InNetClasses('{';'.join(self.items)}')"
        elif self.scope_type == "NetClass":
            return f"InNetClass('{self.items[0]}')" if self.items else "All"
        elif self.scope_type == "Custom":
            return f"'{';'.join(self.items)}'"
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
    
    def to_rul_format(self) -> List[str]:
        """Convert to RUL file format"""
        rul_lines = []
        rul_lines.append(f"Rule")
        rul_lines.append(f"{{")
        rul_lines.append(f"    Name = '{self.name}'")
        rul_lines.append(f"    Enabled = '{str(self.enabled).lower()}'")
        if self.comment:
            rul_lines.append(f"    Comment = '{self.comment}'")
        rul_lines.append(f"    Priority = {self.priority}")
        # Subclasses will add specific rule parameters
        return rul_lines


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
    
    def to_rul_format(self) -> List[str]:
        """Convert to RUL file format"""
        rul_lines = super().to_rul_format()
        rul_lines.append(f"    RuleKind = '{self.rule_type.value}'")
        rul_lines.append(f"    MinimumClearance = {self.min_clearance}")
        rul_lines.append(f"    MinimumClearanceType = '{self.unit.value}'")
        rul_lines.append(f"    SourceScope = {self.source_scope.to_rul_format()}")
        rul_lines.append(f"    TargetScope = {self.target_scope.to_rul_format()}")
        rul_lines.append("}")
        return rul_lines


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
    
    def to_rul_format(self) -> List[str]:
        """Convert to RUL file format"""
        rul_lines = super().to_rul_format()
        rul_lines.append(f"    RuleKind = '{self.rule_type.value}'")
        rul_lines.append(f"    Scope = {self.scope.to_rul_format()}")
        rul_lines.append("}")
        return rul_lines


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
    
    def to_rul_format(self) -> List[str]:
        """Convert to RUL file format"""
        rul_lines = super().to_rul_format()
        rul_lines.append(f"    RuleKind = '{self.rule_type.value}'")
        rul_lines.append(f"    Scope = {self.scope.to_rul_format()}")
        rul_lines.append("}")
        return rul_lines


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
        """Convert all rules to RUL file format"""
        rul_lines = []
        
        # Add header
        rul_lines.append("# Altium Designer Rules")
        rul_lines.append("#")
        rul_lines.append("# Auto-generated file")
        rul_lines.append("")
        
        # Add all rules
        for rule in self.rules:
            rul_lines.extend(rule.to_rul_format())
            rul_lines.append("")
        
        return "\n".join(rul_lines)
    
    def from_rul_content(self, rul_content: str) -> bool:
        """Parse rules from RUL file content"""
        # TODO: Implement RUL file parsing
        # This requires a more complex parser to handle the RUL file format
        logger.warning("RUL file parsing not yet implemented")
        return False
    
    def to_dict(self) -> List[Dict]:
        """Convert all rules to dictionary format"""
        return [rule.to_dict() for rule in self.rules]
