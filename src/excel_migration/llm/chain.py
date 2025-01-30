"""LangChain integration for Excel migrations."""
from typing import Any, Dict, Optional
from langchain.chat_models import ChatOpenAI, ChatAnthropic
from langchain.prompts import ChatPromptTemplate
from langchain.schema import BaseLanguageModel
from langchain.chains import LLMChain
import logging

logger = logging.getLogger(__name__)

class LLMProvider:
    """Factory for creating LLM instances."""
    
    @staticmethod
    def create_llm(provider: str, **kwargs) -> BaseLanguageModel:
        """Create an LLM instance based on provider name."""
        if provider.lower() == "openai":
            return ChatOpenAI(**kwargs)
        elif provider.lower() == "anthropic":
            return ChatAnthropic(**kwargs)
        else:
            raise ValueError(f"Unsupported LLM provider: {provider}")

class TransformationChain:
    """Chain for handling Excel data transformations."""
    
    def __init__(self, llm: BaseLanguageModel):
        self.llm = llm
        self._setup_prompts()

    def _setup_prompts(self):
        """Set up prompt templates."""
        self.transform_prompt = ChatPromptTemplate.from_messages([
            ("system", """You are an expert at analyzing and transforming Excel data.
            Given source data and a transformation rule, output the transformed value.
            Only output the final transformed value, no explanations.
            If the transformation is not possible, output 'ERROR: ' followed by a brief reason."""),
            ("user", """Source Data:
            {source_data}
            
            Transformation Rule:
            {transformation_rule}
            
            Additional Context:
            {context}""")
        ])

        self.validation_prompt = ChatPromptTemplate.from_messages([
            ("system", """You are an expert at validating Excel data.
            Given a value and validation rules, determine if the value is valid.
            Output only 'VALID' or 'INVALID: ' followed by a brief reason."""),
            ("user", """Value to Validate:
            {value}
            
            Validation Rules:
            {validation_rules}
            
            Additional Context:
            {context}""")
        ])

    def transform_value(self, source_data: Dict[str, Any], 
                       transformation_rule: str,
                       context: Optional[Dict[str, Any]] = None) -> Any:
        """Transform source data according to rule."""
        try:
            chain = LLMChain(llm=self.llm, prompt=self.transform_prompt)
            result = chain.run(
                source_data=str(source_data),
                transformation_rule=transformation_rule,
                context=str(context or {})
            )

            if result.startswith("ERROR: "):
                logger.error(f"Transformation failed: {result}")
                return None

            return self._parse_result(result)

        except Exception as e:
            logger.error(f"Error in transformation chain: {str(e)}")
            return None

    def validate_value(self, value: Any, 
                      validation_rules: str,
                      context: Optional[Dict[str, Any]] = None) -> bool:
        """Validate a value according to rules."""
        try:
            chain = LLMChain(llm=self.llm, prompt=self.validation_prompt)
            result = chain.run(
                value=str(value),
                validation_rules=validation_rules,
                context=str(context or {})
            )

            return not result.startswith("INVALID: ")

        except Exception as e:
            logger.error(f"Error in validation chain: {str(e)}")
            return False

    def _parse_result(self, result: str) -> Any:
        """Parse LLM output into appropriate type."""
        result = result.strip()
        
        # Try to convert to number if possible
        try:
            if '.' in result:
                return float(result)
            return int(result)
        except ValueError:
            pass

        # Handle boolean values
        if result.lower() == 'true':
            return True
        if result.lower() == 'false':
            return False

        # Return as string if no other type matches
        return result

class FormulaAnalysisChain:
    """Chain for analyzing Excel formulas."""
    
    def __init__(self, llm: BaseLanguageModel):
        self.llm = llm
        self._setup_prompts()

    def _setup_prompts(self):
        """Set up prompt templates."""
        self.formula_prompt = ChatPromptTemplate.from_messages([
            ("system", """You are an expert at analyzing Excel formulas.
            Given a formula and its context, explain its logic and suggest improvements.
            Focus on understanding complex calculations and business rules."""),
            ("user", """Formula:
            {formula}
            
            Context:
            {context}
            
            Task:
            {task}""")
        ])

    def analyze_formula(self, formula: str, 
                       context: Dict[str, Any],
                       task: str = "Explain this formula's logic") -> str:
        """Analyze an Excel formula."""
        try:
            chain = LLMChain(llm=self.llm, prompt=self.formula_prompt)
            return chain.run(
                formula=formula,
                context=str(context),
                task=task
            )
        except Exception as e:
            logger.error(f"Error in formula analysis chain: {str(e)}")
            return f"Error analyzing formula: {str(e)}"