#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Excel Data Model
===============

Defines the data models for Excel import/export operations.
Handles pivot table data structures and conversion to rule models.
"""

import logging
from typing import Dict, List, Optional, Union, Tuple
import pandas as pd
import numpy as np

from models.rule_model import (
    UnitType, RuleType, RuleScope, 
    BaseRule, ClearanceRule, ShortCircuitRule, UnRoutedNetRule
)

logger = logging.getLogger(__name__)

class ExcelPivotData:
    """Represents data from an Excel pivot table"""
    
    def __init__(self, rule_type: RuleType = None):
        """Initialize Excel pivot data"""
        self.rule_type = rule_type
        self.pivot_df = None
        self.column_index = None
        self.row_index = None
        self.values = None
        self.unit = UnitType.MIL
    
    def load_dataframe(self, df: pd.DataFrame, unit: UnitType = UnitType.MIL) -> bool:
        """Load data from pandas DataFrame"""
        try:
            # Store the DataFrame
            self.pivot_df = df
            self.unit = unit
            
            # Check if the DataFrame has the expected structure
            if df.empty or df.shape[0] < 2 or df.shape[1] < 2:
                logger.error("DataFrame is too small to be a valid pivot table")
                return False
              # Extract column and row indexes
            self.column_index = list(df.columns)[1:]  # Skip the first column (row headers)
            self.row_index = df.iloc[:, 0].tolist()  # First column values
            
            # Extract values as a numpy array (excluding first column)
            self.values = df.iloc[:, 1:].to_numpy()
            
            # Debug info
            logger.info(f"Loaded DataFrame with {len(self.row_index)} rows and {len(self.column_index)} columns")
            logger.info(f"Column index: {self.column_index}")
            logger.info(f"Row index first few: {self.row_index[:5] if len(self.row_index) > 5 else self.row_index}")
            logger.info(f"Values shape: {self.values.shape}")
            logger.info(f"Values sample: {str(self.values[:2, :2]) if self.values.size > 0 else 'Empty'}")
            return True
        except Exception as e:
            logger.error(f"Error loading DataFrame: {str(e)}")
            return False
    
    def to_clearance_rules(self, rule_name_prefix: str = "Clearance_") -> List[ClearanceRule]:
        """Convert pivot data to clearance rules"""
        if self.pivot_df is None or self.values is None:
            logger.error("No data loaded")
            return []
        
        rules = []
        
        # For each cell in the pivot table
        for row_idx, row_name in enumerate(self.row_index):
            for col_idx, col_name in enumerate(self.column_index):
                # Get clearance value
                clearance_value = self.values[row_idx, col_idx]
                
                # Skip NaN or empty values
                if pd.isna(clearance_value) or clearance_value == "":
                    continue
                
                # Convert to float if possible
                try:
                    clearance_value = float(clearance_value)
                except (ValueError, TypeError):
                    logger.warning(f"Invalid clearance value at {row_name}/{col_name}: {clearance_value}")
                    continue
                
                # Create rule name
                rule_name = f"{rule_name_prefix}{row_name}_{col_name}"
                
                # Create rule scopes
                source_scope = RuleScope("NetClass", [row_name])
                target_scope = RuleScope("NetClass", [col_name])
                
                # Create the rule
                rule = ClearanceRule(
                    name=rule_name,
                    enabled=True,
                    comment=f"Clearance between {row_name} and {col_name}",
                    min_clearance=clearance_value,
                    unit=self.unit,
                    source_scope=source_scope,
                    target_scope=target_scope
                )
                
                rules.append(rule)
        
        logger.info(f"Created {len(rules)} clearance rules from pivot data")
        return rules
    
    def to_short_circuit_rules(self, rule_name_prefix: str = "ShortCircuit_") -> List[ShortCircuitRule]:
        """Convert pivot data to short circuit rules"""
        if self.pivot_df is None or self.values is None:
            logger.error("No data loaded")
            return []
        
        rules = []
        
        # Extract unique net classes from both row and column indexes
        net_classes = set(self.row_index) | set(self.column_index)
        
        # Create a rule for each net class
        for net_class in net_classes:
            rule_name = f"{rule_name_prefix}{net_class}"
            scope = RuleScope("NetClass", [net_class])
            
            rule = ShortCircuitRule(
                name=rule_name,
                enabled=True,
                comment=f"Short circuit rule for {net_class}",
                scope=scope
            )
            
            rules.append(rule)
        
        logger.info(f"Created {len(rules)} short circuit rules from pivot data")
        return rules
    
    def to_unrouted_net_rules(self, rule_name_prefix: str = "UnroutedNet_") -> List[UnRoutedNetRule]:
        """Convert pivot data to unrouted net rules"""
        if self.pivot_df is None or self.values is None:
            logger.error("No data loaded")
            return []
        
        rules = []
        
        # Extract unique net classes from both row and column indexes
        net_classes = set(self.row_index) | set(self.column_index)
        
        # Create a rule for each net class
        for net_class in net_classes:
            rule_name = f"{rule_name_prefix}{net_class}"
            scope = RuleScope("NetClass", [net_class])
            
            rule = UnRoutedNetRule(
                name=rule_name,
                enabled=True,
                comment=f"Unrouted net rule for {net_class}",
                scope=scope
            )
            
            rules.append(rule)
        
        logger.info(f"Created {len(rules)} unrouted net rules from pivot data")
        return rules
    
    @staticmethod
    def from_clearance_rules(rules: List[ClearanceRule]) -> Optional['ExcelPivotData']:
        """Create pivot data from clearance rules"""
        if not rules:
            logger.error("No rules provided")
            return None
        
        # Extract all unique net classes from rules
        net_classes = set()
        for rule in rules:
            # Assume source and target scopes are NetClass type with single item
            if rule.source_scope.scope_type == "NetClass" and rule.source_scope.items:
                net_classes.add(rule.source_scope.items[0])
            if rule.target_scope.scope_type == "NetClass" and rule.target_scope.items:
                net_classes.add(rule.target_scope.items[0])
        
        net_classes = sorted(list(net_classes))
        
        # Create empty DataFrame with net classes as rows and columns
        df = pd.DataFrame(
            index=net_classes,
            columns=["NetClass"] + net_classes
        )
        
        # Set row headers
        df["NetClass"] = net_classes
        
        # Fill with NaN initially
        for col in net_classes:
            df[col] = np.nan
        
        # Fill in clearance values from rules
        for rule in rules:
            if (rule.source_scope.scope_type == "NetClass" and rule.source_scope.items and
                rule.target_scope.scope_type == "NetClass" and rule.target_scope.items):
                
                source_class = rule.source_scope.items[0]
                target_class = rule.target_scope.items[0]
                
                if source_class in net_classes and target_class in net_classes:
                    df.loc[source_class, target_class] = rule.min_clearance
        
        # Create ExcelPivotData instance
        pivot_data = ExcelPivotData(RuleType.CLEARANCE)
        pivot_data.load_dataframe(df, unit=rules[0].unit if rules else UnitType.MIL)
        
        return pivot_data
    
    def to_dataframe(self) -> pd.DataFrame:
        """Convert pivot data to DataFrame for export"""
        if self.pivot_df is not None:
            return self.pivot_df
        
        # If we don't have a pivot_df but have the components, reconstruct it
        if self.row_index is not None and self.column_index is not None and self.values is not None:
            # Create DataFrame with the right shape
            df = pd.DataFrame(
                index=range(len(self.row_index)),
                columns=["NetClass"] + self.column_index
            )
            
            # Set row headers
            df["NetClass"] = self.row_index
            
            # Set values
            for row_idx, row_name in enumerate(self.row_index):
                for col_idx, col_name in enumerate(self.column_index):
                    df.loc[row_idx, col_name] = self.values[row_idx, col_idx]
            
            return df
        
        logger.error("Cannot convert to DataFrame, missing data")
        return pd.DataFrame()
