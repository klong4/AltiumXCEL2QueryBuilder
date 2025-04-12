#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import logging
from datetime import datetime
from typing import List, Dict, Any, Optional
import pandas as pd
import uuid # Import uuid for generating unique IDs

from models.rule_model import RuleManager

class RuleGeneratorError(Exception):
    """Custom exception class for RuleGenerator errors."""

    def __init__(self, message: str):
        """
        Initializes the RuleGeneratorError.

        Args:
            message (str): Error message to be displayed.
        """
        super().__init__(message)
        logging.error(message)

class RuleGenerator:
    """Generates Altium Designer rule files (.RUL) from structured data."""

    def __init__(self, rule_manager: RuleManager):
        """
        Initializes the RuleGenerator.

        Args:
            rule_manager (RuleManager): The manager containing the rule data.
        """
        self.rule_manager = rule_manager

    def _format_scope_expression(self, scope: str, type: str) -> str:
        """Formats the scope expression based on type and scope."""
        if not type or not scope:
            return "All" # Default if type or scope is missing

        # Simple mapping for now, can be expanded
        type_lower = str(type).lower() # Ensure type is string and lower case
        scope_str = str(scope) # Ensure scope is string

        if type_lower == 'netclass':
            return f"InNetClass('{scope_str}')"
        elif type_lower == 'net':
            return f"InNet('{scope_str}')"
        elif type_lower == 'layer':
            return f"OnLayer('{scope_str}')"
        elif type_lower == 'room':
             return f"WithinRoom('{scope_str}')"
        # Add more mappings as needed (e.g., IsPad, IsVia, etc.)
        else:
            # Fallback for unknown types or simple keywords
            scope_lower = scope_str.lower()
            if scope_lower == 'all':
                return 'All'
            elif scope_lower == 'iskeepout':
                 return 'IsKeepOut'
            elif scope_lower == 'onmid': # Added based on sample
                 return 'OnMid'
            # Add more simple keywords
            else:
                 # Attempt a generic format, might need refinement
                 # If type is provided but not matched above, use it
                 if type_lower:
                     return f"{type}('{scope_str}')" # e.g., IsPad('P1') - needs verification
                 else:
                     # If only scope is provided and not matched, return it directly?
                     # This might be incorrect, defaulting to All might be safer
                     # return scope_str
                     return 'All' # Safer default


    def generate_rul_content(self, rules_data: List[Dict[str, Any]]) -> str:
        """
        Generates the content for the .RUL file based on the provided rules data,
        matching the single-line pipe-delimited format.

        Args:
            rules_data (List[Dict[str, Any]]): A list of dictionaries, each representing a rule.
                                                Expected keys: 'Name', 'Priority', 'Enabled',
                                                'Object Scope 1', 'Object Type 1',
                                                'Object Scope 2', 'Object Type 2',
                                                'Value', 'Unit', 'Comment', 'RuleKind',
                                                'NetScope', 'LayerKind'.

        Returns:
            str: The generated content for the .RUL file.
        """
        rul_lines = []
        # No header needed based on the sample

        for i, rule in enumerate(rules_data):
            try:
                # --- Extract and Sanitize Data ---
                name = rule.get('Name', f'Rule_{i+1}')
                priority = int(rule.get('Priority', i + 1))
                enabled = rule.get('Enabled', True)
                scope1_val = rule.get('Object Scope 1', 'All')
                type1 = rule.get('Object Type 1', '')
                scope2_val = rule.get('Object Scope 2', 'All')
                type2 = rule.get('Object Type 2', '')
                value = rule.get('Value', 0)
                unit = rule.get('Unit', 'mil') # Default to mil based on sample
                comment = rule.get('Comment', '')
                rule_kind = rule.get('RuleKind', 'Clearance') # Default to Clearance
                net_scope = rule.get('NetScope', 'DifferentNets')
                layer_kind = rule.get('LayerKind', 'SameLayer')
                layer = rule.get('Layer', 'UNKNOWN') # Get Layer if specified
                ignore_pad_clearance = rule.get('IgnorePadToPadClearance', False) # Default based on sample

                # Basic validation/sanitization
                name = ''.join(c for c in str(name) if c.isalnum() or c in ['_', '-', ' ', '(', ')', '.']) # Allow more chars in name
                enabled_str = 'TRUE' if enabled else 'FALSE' # Use uppercase TRUE/FALSE
                value_str = str(value)
                comment_str = str(comment)
                ignore_pad_clearance_str = 'TRUE' if ignore_pad_clearance else 'FALSE'

                # --- Format Expressions ---
                scope1_expr = self._format_scope_expression(scope1_val, type1)
                scope2_expr = self._format_scope_expression(scope2_val, type2)

                # --- Generate Unique ID (using 8 random hex chars, closer to sample) ---
                unique_id = uuid.uuid4().hex[:8].upper()

                # --- Construct Rule Dictionary (Order matters for comparison with sample) ---
                rule_dict = {
                    "SELECTION": "FALSE",
                    "LAYER": layer,
                    "LOCKED": "FALSE",
                    "POLYGONOUTLINE": "FALSE",
                    "USERROUTED": "TRUE",
                    "KEEPOUT": "FALSE",
                    "UNIONINDEX": 0,
                    "RULEKIND": rule_kind,
                    "NETSCOPE": net_scope,
                    "LAYERKIND": layer_kind,
                    "SCOPE1EXPRESSION": scope1_expr,
                    "SCOPE2EXPRESSION": scope2_expr,
                    "NAME": name,
                    "ENABLED": enabled_str,
                    "PRIORITY": priority,
                    "COMMENT": comment_str,
                    "UNIQUEID": unique_id,
                    "DEFINEDBYLOGICALDOCUMENT": "FALSE",
                }

                # Add RuleKind specific fields
                if rule_kind == 'Clearance':
                    rule_dict["GAP"] = f"{value_str}{unit}"
                    rule_dict["GENERICCLEARANCE"] = f"{value_str}{unit}"
                    rule_dict["IGNOREPADTOPADCLEARANCEINFOOTPRINT"] = ignore_pad_clearance_str
                    rule_dict["OBJECTCLEARANCES"] = " " # Default empty, complex structure
                # Add other RuleKind specific fields here if needed
                # elif rule_kind == 'Width':
                #     rule_dict["MINLIMIT"] = f"{value_str}{unit}"
                #     rule_dict["MAXLIMIT"] = f"{value_str}{unit}" # Example
                #     rule_dict["PREFEREDWIDTH"] = f"{value_str}{unit}" # Example

                # --- Format as Single Line ---
                rule_line = "|".join([f"{k}={v}" for k, v in rule_dict.items()])
                rul_lines.append(rule_line)

            except Exception as e:
                logging.error(f"Error processing rule '{rule.get('Name', 'N/A')}': {e}", exc_info=True)
                # Optionally skip the rule or add a placeholder comment

        # Join lines with newline
        return "\n".join(rul_lines) + "\n" # Add trailing newline


    def generate_and_save_rul(self, output_path: str, rules_data: Optional[List[Dict[str, Any]]] = None):
        """
        Generates the .RUL file content and saves it to the specified path.

        Args:
            output_path (str): The full path where the .RUL file should be saved.
            rules_data (Optional[List[Dict[str, Any]]]): The rule data to use.
                                                        If None, uses data from the rule_manager.
        """
        if rules_data is None:
            rules_data = self.rule_manager.to_dict()

        if not rules_data:
            logging.warning("No rules data provided or found in the model. Cannot generate RUL file.")
            return

        rul_content = self.generate_rul_content(rules_data)

        try:
            # Ensure we write with an encoding that supports potential special characters
            # Use 'utf-8' and let Python handle line endings based on OS ('\n')
            # Altium likely handles standard \n line endings correctly.
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(rul_content)

            logging.info(f"Successfully generated and saved RUL file to: {output_path}")
        except IOError as e:
            logging.error(f"Error saving RUL file to {output_path}: {e}")
            raise # Re-raise the exception for the caller to handle

    def generate_from_dataframe(self, df: pd.DataFrame) -> str:
         """
         Generates RUL content directly from a pandas DataFrame.

         Args:
             df (pd.DataFrame): DataFrame containing the rule data.
                                Expected columns match the keys in generate_rul_content.

         Returns:
             str: The generated content for the .RUL file.
         """
         # Convert DataFrame rows to list of dictionaries
         # Handle potential NaN values from Excel/Pandas
         df_filled = df.fillna('') # Replace NaN with empty strings
         rules_data = df_filled.to_dict('records')
         return self.generate_rul_content(rules_data)
