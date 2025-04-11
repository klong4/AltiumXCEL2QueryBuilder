#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Excel Importer Service
=====================

Service for importing Excel files with pivot tables and converting to rule models.
"""

import os
import logging
from typing import Dict, List, Optional, Union, Tuple
import pandas as pd

from models.excel_data import ExcelPivotData
from models.rule_model import (
    UnitType, RuleType, RuleScope, BaseRule, 
    ClearanceRule, ShortCircuitRule, UnRoutedNetRule
)

logger = logging.getLogger(__name__)

class ExcelImportError(Exception):
    """Exception raised for errors during Excel import"""
    pass

class ExcelImporter:
    """Service for importing Excel files"""
    def __init__(self):
        """Initialize Excel importer"""
        self.last_file_path = None
        self.last_sheet_name = None
        self.detected_unit = UnitType.MIL
    
    def get_sheet_names(self, file_path: str) -> List[str]:
        """Get the sheet names from an Excel file"""
        try:
            # Check if file exists
            if not os.path.exists(file_path):
                raise ExcelImportError(f"File not found: {file_path}")
            
            # Read Excel file
            xls = pd.ExcelFile(file_path)
            sheet_names = xls.sheet_names
            
            if not sheet_names:
                raise ExcelImportError(f"No sheets found in Excel file: {file_path}")
            
            logger.info(f"Found {len(sheet_names)} sheets in Excel file: {file_path}")
            return sheet_names
            
        except Exception as e:
            error_msg = f"Error getting sheet names from Excel file: {str(e)}"
            logger.error(error_msg)
            raise ExcelImportError(error_msg)
            
    def import_file(self, file_path: str, sheet_name: Optional[str] = None) -> pd.DataFrame:
        """Import Excel file and return as DataFrame"""
        try:
            # Check if file exists
            if not os.path.exists(file_path):
                raise ExcelImportError(f"File not found: {file_path}")
            
            # Read Excel file
            if sheet_name:
                df = pd.read_excel(file_path, sheet_name=sheet_name)
            else:
                # Try to read the first sheet
                xls = pd.ExcelFile(file_path)
                if not xls.sheet_names:
                    raise ExcelImportError(f"No sheets found in Excel file: {file_path}")
                
                sheet_name = xls.sheet_names[0]
                df = pd.read_excel(file_path, sheet_name=sheet_name)
            
            # Ensure the DataFrame is not empty before returning
            if df.empty:
                raise ExcelImportError("The imported Excel file contains no data.")
            
            # Store successful import details
            self.last_file_path = file_path
            self.last_sheet_name = sheet_name
            
            logger.info(f"Imported Excel file: {file_path}, sheet: {sheet_name}")
            return df
        
        except pd.errors.EmptyDataError:
            error_msg = f"Excel file is empty: {file_path}"
            logger.error(error_msg)
            raise ExcelImportError(error_msg)
        
        except Exception as e:
            error_msg = f"Error importing Excel file: {str(e)}"
            logger.error(error_msg)
            raise ExcelImportError(error_msg)
    
    def detect_pivot_structure(self, df: pd.DataFrame) -> Tuple[bool, str]:
        """
        Detect if DataFrame has a valid pivot table structure
        Returns (is_valid, message)
        """
        # Check if DataFrame is empty
        if df.empty:
            return False, "DataFrame is empty"

        # Check if DataFrame has at least 2 rows and 2 columns
        if df.shape[0] < 2 or df.shape[1] < 2:
            return False, "DataFrame is too small to be a valid pivot table"

        # Check if 'Rule Set' column exists
        if 'Rule Set' not in df.columns:
            return False, "Missing 'Rule Set' column in the DataFrame"

        # Check if first column contains string values (net class names)
        first_col = df.iloc[:, 0]
        if not all(isinstance(x, str) for x in first_col if pd.notna(x)):
            return False, "First column should contain net class names (string values)"

        # Check if column headers are also string values
        headers = df.columns[1:]  # Skip the first column header
        if not all(isinstance(x, str) for x in headers if pd.notna(x)):
            return False, "Column headers should be net class names (string values)"

        # Try to detect the unit type
        self._detect_unit_type(df)

        return True, f"Valid pivot table structure detected with unit: {self.detected_unit.value}"
    
    def _detect_unit_type(self, df: pd.DataFrame) -> UnitType:
        """Try to detect the unit type based on values in the DataFrame"""
        # Get all numeric values in the DataFrame
        numeric_values = []
        for col in df.columns[1:]:  # Skip the first column
            for val in df[col]:
                if isinstance(val, (int, float)) and pd.notna(val):
                    numeric_values.append(val)
        
        if not numeric_values:
            # Default to mil if no numeric values
            self.detected_unit = UnitType.MIL
            return self.detected_unit
        
        # Check the range of values to guess the unit
        avg_value = sum(numeric_values) / len(numeric_values)
        
        if avg_value < 1.0:
            # Very small values are likely in inches
            self.detected_unit = UnitType.INCH
        elif avg_value < 50.0:
            # Medium values are likely in mm
            self.detected_unit = UnitType.MM
        else:
            # Larger values are likely in mils
            self.detected_unit = UnitType.MIL
        
        logger.info(f"Detected unit type: {self.detected_unit.value}")
        return self.detected_unit
    
    def import_as_pivot_data(self, file_path: str, sheet_name: Optional[str] = None,
                           rule_type: RuleType = RuleType.CLEARANCE) -> ExcelPivotData:
        """
        Import Excel file and convert to pivot data structure
        """
        try:
            # Import Excel file
            df = self.import_file(file_path, sheet_name)
            
            # Detect pivot structure
            is_valid, message = self.detect_pivot_structure(df)
            if not is_valid:
                raise ExcelImportError(f"Invalid pivot table structure: {message}")
            
            # Create pivot data
            pivot_data = ExcelPivotData(rule_type)
            success = pivot_data.load_dataframe(df, self.detected_unit)
            
            if not success:
                raise ExcelImportError("Failed to load DataFrame into pivot data structure")
            
            logger.info(f"Successfully imported {file_path} as pivot data for {rule_type.value}")
            return pivot_data
        
        except Exception as e:
            error_msg = f"Error importing as pivot data: {str(e)}"
            logger.error(error_msg)
            raise ExcelImportError(error_msg)
    
    def import_as_clearance_rules(self, file_path: str, sheet_name: Optional[str] = None,
                                rule_name_prefix: str = "Clearance_") -> List[ClearanceRule]:
        """
        Import Excel file and convert to clearance rules
        """
        try:
            # Import as pivot data
            pivot_data = self.import_as_pivot_data(file_path, sheet_name, RuleType.CLEARANCE)
            
            # Convert to clearance rules
            rules = pivot_data.to_clearance_rules(rule_name_prefix)
            
            logger.info(f"Converted pivot data to {len(rules)} clearance rules")
            return rules
        
        except Exception as e:
            error_msg = f"Error importing as clearance rules: {str(e)}"
            logger.error(error_msg)
            raise ExcelImportError(error_msg)
    
    def import_as_short_circuit_rules(self, file_path: str, sheet_name: Optional[str] = None,
                                    rule_name_prefix: str = "ShortCircuit_") -> List[ShortCircuitRule]:
        """
        Import Excel file and convert to short circuit rules
        """
        try:
            # Import as pivot data
            pivot_data = self.import_as_pivot_data(file_path, sheet_name, RuleType.SHORT_CIRCUIT)
            
            # Convert to short circuit rules
            rules = pivot_data.to_short_circuit_rules(rule_name_prefix)
            
            logger.info(f"Converted pivot data to {len(rules)} short circuit rules")
            return rules
        
        except Exception as e:
            error_msg = f"Error importing as short circuit rules: {str(e)}"
            logger.error(error_msg)
            raise ExcelImportError(error_msg)
    
    def import_as_unrouted_net_rules(self, file_path: str, sheet_name: Optional[str] = None,
                                   rule_name_prefix: str = "UnroutedNet_") -> List[UnRoutedNetRule]:
        """
        Import Excel file and convert to unrouted net rules
        """
        try:
            # Import as pivot data
            pivot_data = self.import_as_pivot_data(file_path, sheet_name, RuleType.UNROUTED_NET)
            
            # Convert to unrouted net rules
            rules = pivot_data.to_unrouted_net_rules(rule_name_prefix)
            
            logger.info(f"Converted pivot data to {len(rules)} unrouted net rules")
            return rules
        
        except Exception as e:
            error_msg = f"Error importing as unrouted net rules: {str(e)}"
            logger.error(error_msg)
            raise ExcelImportError(error_msg)
    
    def get_sheet_names(self, file_path: str) -> List[str]:
        """Get sheet names from Excel file"""
        try:
            if not os.path.exists(file_path):
                raise ExcelImportError(f"File not found: {file_path}")
            
            xls = pd.ExcelFile(file_path)
            return xls.sheet_names
        
        except Exception as e:
            error_msg = f"Error getting sheet names: {str(e)}"
            logger.error(error_msg)
            raise ExcelImportError(error_msg)
