"""Rule generation and execution engine."""
from pathlib import Path
from typing import Dict, Any, List, Optional
import openpyxl
from loguru import logger
from langchain_openai import ChatOpenAI

class RuleEngine:
    """Engine for generating and executing migration rules."""
    
    def __init__(self, llm_provider: str = "openai"):
        """Initialize the rule engine."""
        self.llm = ChatOpenAI(
            model_name="gpt-4",
            temperature=0.7
        )
        logger.debug(f"Initialized rule engine with {llm_provider}")

    async def generate_rules(
        self,
        source_file: Path,
        target_file: Path,
        source_sheet: str,
        target_sheet: str
    ) -> List[Dict[str, Any]]:
        """Generate migration rules by analyzing example files."""
        try:
            # Analyze source and target structures
            source_structure = self._analyze_sheet(source_file, source_sheet)
            target_structure = self._analyze_sheet(target_file, target_sheet)
            
            # Generate rules based on analysis
            rules = await self._generate_rules_from_analysis(
                source_structure,
                target_structure
            )
            
            return rules
            
        except Exception as e:
            logger.error(f"Rule generation failed: {str(e)}")
            return []

    def _analyze_sheet(self, file_path: Path, sheet_name: str) -> Dict[str, Any]:
        """Analyze sheet structure and content."""
        try:
            wb = openpyxl.load_workbook(file_path, read_only=True)
            ws = wb[sheet_name]
            
            analysis = {
                "sheet_name": sheet_name,
                "headers": [],
                "data_types": {},
                "sample_data": []
            }
            
            # Get headers
            for cell in ws[1]:
                if cell.value:
                    analysis["headers"].append(str(cell.value))
            
            # Analyze data types and get samples
            for col_idx, header in enumerate(analysis["headers"], 1):
                col_letter = openpyxl.utils.get_column_letter(col_idx)
                values = []
                for row in range(2, min(7, ws.max_row + 1)):  # Sample first 5 data rows
                    cell = ws[f"{col_letter}{row}"]
                    if cell.value:
                        values.append(str(cell.value))
                
                if values:
                    analysis["data_types"][header] = self._infer_data_type(values)
                    analysis["sample_data"].append({
                        "header": header,
                        "samples": values
                    })
            
            wb.close()
            return analysis
            
        except Exception as e:
            logger.error(f"Sheet analysis failed: {str(e)}")
            raise

    def _infer_data_type(self, values: List[str]) -> str:
        """Infer data type from sample values."""
        try:
            # Try numeric
            float(values[0])
            return "numeric"
        except ValueError:
            pass
        
        # Check date format
        date_indicators = ["/", "-", ":", "AM", "PM"]
        if any(ind in values[0] for ind in date_indicators):
            return "datetime"
        
        # Check boolean
        bool_values = ["true", "false", "yes", "no", "0", "1"]
        if all(v.lower() in bool_values for v in values):
            return "boolean"
        
        return "text"

    async def _generate_rules_from_analysis(
        self,
        source: Dict[str, Any],
        target: Dict[str, Any]
    ) -> List[Dict[str, Any]]:
        """Generate rules by comparing source and target structures."""
        rules = []
        
        # Direct field mappings
        for target_header in target["headers"]:
            # Find best matching source field
            source_field = self._find_matching_field(
                target_header,
                source["headers"],
                source["data_types"]
            )
            
            if source_field:
                rules.append({
                    "type": "field_mapping",
                    "source_field": source_field,
                    "target_field": target_header,
                    "transformation": self._get_transformation_rule(
                        source["data_types"].get(source_field),
                        target_header
                    )
                })
        
        # Add any necessary data transformations
        transformation_rules = await self._generate_transformation_rules(
            source,
            target
        )
        rules.extend(transformation_rules)
        
        return rules

    def _find_matching_field(
        self,
        target_field: str,
        source_fields: List[str],
        source_types: Dict[str, str]
    ) -> str:
        """Find best matching source field for target field."""
        # Direct match
        if target_field in source_fields:
            return target_field
        
        # Fuzzy match based on common variations
        target_normalized = target_field.lower().replace(" ", "").replace("_", "")
        for source_field in source_fields:
            source_normalized = source_field.lower().replace(" ", "").replace("_", "")
            if source_normalized == target_normalized:
                return source_field
            
            # Handle common field variations
            if target_normalized.endswith("id") and source_normalized.endswith("id"):
                if target_normalized[:-2] in source_normalized:
                    return source_field
            
            if target_normalized.endswith("name") and source_normalized.endswith("name"):
                if target_normalized[:-4] in source_normalized:
                    return source_field
        
        return ""

    def _get_transformation_rule(self, source_type: str, target_field: str) -> Dict[str, Any]:
        """Get appropriate transformation rule based on field types."""
        transformation = {
            "type": "direct",  # Default to direct copy
            "params": {}
        }
        
        # Add type-specific transformations
        if source_type == "datetime":
            transformation.update({
                "type": "datetime_format",
                "params": {
                    "format": self._infer_date_format(target_field)
                }
            })
        elif source_type == "numeric":
            transformation.update({
                "type": "numeric_format",
                "params": {
                    "decimal_places": 2 if "amount" in target_field.lower() else 0,
                    "thousands_separator": True
                }
            })
        
        return transformation

    def _infer_date_format(self, field_name: str) -> str:
        """Infer date format based on field name."""
        if any(word in field_name.lower() for word in ["time", "timestamp"]):
            return "%Y-%m-%d %H:%M:%S"
        return "%Y-%m-%d"

    async def _generate_transformation_rules(
        self,
        source: Dict[str, Any],
        target: Dict[str, Any]
    ) -> List[Dict[str, Any]]:
        """Generate complex transformation rules."""
        rules = []
        
        # Look for calculated fields
        for target_header in target["headers"]:
            if not self._find_matching_field(target_header, source["headers"], source["data_types"]):
                # This might be a calculated field
                rule = await self._infer_calculation_rule(
                    target_header,
                    source,
                    target
                )
                if rule:
                    rules.append(rule)
        
        return rules

    async def _infer_calculation_rule(
        self,
        target_field: str,
        source: Dict[str, Any],
        target: Dict[str, Any]
    ) -> Optional[Dict[str, Any]]:
        """Infer calculation rule for a target field."""
        try:
            # Use LLM to suggest calculation
            prompt = f"""Given a target field '{target_field}' and available source fields:
            {', '.join(source['headers'])}
            
            Suggest a calculation rule to derive the target field.
            Consider the sample data:
            {source['sample_data']}
            
            Respond with a JSON object containing:
            - type: calculation
            - formula: the calculation formula
            - description: explanation of the calculation
            """
            
            response = await self.llm.ainvoke(prompt)
            
            # Parse and validate response
            if isinstance(response, dict) and "formula" in response:
                return {
                    "type": "calculation",
                    "target_field": target_field,
                    "formula": response["formula"],
                    "description": response.get("description", ""),
                    "source_fields": self._extract_source_fields(response["formula"])
                }
            
            return None
            
        except Exception as e:
            logger.error(f"Failed to infer calculation rule: {str(e)}")
            return None

    def _extract_source_fields(self, formula: str) -> List[str]:
        """Extract source field names from a formula."""
        # Simple extraction - look for field-like patterns
        import re
        field_pattern = r'\[([^\]]+)\]'  # Fields in [brackets]
        return re.findall(field_pattern, formula)