"""LangChain integration for Excel migrations."""
from typing import Any, Dict, Optional, List
from langchain_openai import ChatOpenAI
from langchain_community.chat_models import ChatAnthropic
from langchain.prompts import ChatPromptTemplate
from langchain_core.language_models import BaseLanguageModel
from langchain.chains import LLMChain
from langchain.chains.base import Chain
from langchain.memory import ConversationBufferMemory
from langchain.callbacks.base import BaseCallbackHandler
import logging

from .agents import MultiAgentSystem, ExcelAgent, AgentFactory

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

class ChainManager:
    """Manager for LangChain components."""
    
    def __init__(self, llm: BaseLanguageModel):
        self.llm = llm
        self.memory = ConversationBufferMemory(
            memory_key="chat_history",
            return_messages=True
        )
        self.multi_agent_system = MultiAgentSystem(llm)
        self.transformation_chain = self._create_transformation_chain()
        self.validation_chain = self._create_validation_chain()
        self.formula_chain = self._create_formula_chain()
    
    def _create_transformation_chain(self) -> Chain:
        """Create the transformation chain."""
        prompt = ChatPromptTemplate.from_messages([
            ("system", """You are an expert at transforming Excel data.
            Given source data and transformation rules, output the transformed data.
            Consider the context and previous transformations in your decisions."""),
            ("human", """Source Data: {source_data}
            Transformation Rules: {rules}
            Context: {context}
            Previous Transformations: {chat_history}
            """)
        ])
        
        return LLMChain(
            llm=self.llm,
            prompt=prompt,
            memory=self.memory,
            verbose=True
        )
    
    def _create_validation_chain(self) -> Chain:
        """Create the validation chain."""
        prompt = ChatPromptTemplate.from_messages([
            ("system", """You are an expert at validating Excel data.
            Given data and validation rules, determine if the data is valid.
            Provide detailed feedback on any validation failures."""),
            ("human", """Data to Validate: {data}
            Validation Rules: {rules}
            Context: {context}
            Previous Validations: {chat_history}
            """)
        ])
        
        return LLMChain(
            llm=self.llm,
            prompt=prompt,
            memory=self.memory,
            verbose=True
        )
    
    def _create_formula_chain(self) -> Chain:
        """Create the formula analysis chain."""
        prompt = ChatPromptTemplate.from_messages([
            ("system", """You are an expert at analyzing Excel formulas.
            Given a formula, explain its logic and suggest optimizations.
            Consider the context and previous analyses in your suggestions."""),
            ("human", """Formula: {formula}
            Context: {context}
            Previous Analyses: {chat_history}
            """)
        ])
        
        return LLMChain(
            llm=self.llm,
            prompt=prompt,
            memory=self.memory,
            verbose=True
        )

class ExcelProcessor:
    """Processor for Excel-specific operations using LangChain."""
    
    def __init__(self, chain_manager: ChainManager):
        self.chain_manager = chain_manager
        self.agents: Dict[str, ExcelAgent] = {}
    
    def get_or_create_agent(self, agent_type: str) -> ExcelAgent:
        """Get an existing agent or create a new one."""
        if agent_type not in self.agents:
            self.agents[agent_type] = AgentFactory.create_agent(
                agent_type,
                self.chain_manager.llm
            )
        return self.agents[agent_type]
    
    async def process_transformation(self, data: Any, rules: Dict[str, Any],
                                  context: Optional[Dict[str, Any]] = None) -> Any:
        """Process a data transformation."""
        try:
            # Use multi-agent system for complex transformations
            if self._is_complex_transformation(rules):
                return await self.chain_manager.multi_agent_system.process_task(
                    task=f"Transform data according to rules: {rules}",
                    context={"data": data, **(context or {})}
                )
            
            # Use transformation chain for simple transformations
            return await self.chain_manager.transformation_chain.arun(
                source_data=str(data),
                rules=str(rules),
                context=str(context or {})
            )
            
        except Exception as e:
            logger.error(f"Transformation failed: {str(e)}")
            return None
    
    async def validate_data(self, data: Any, rules: Dict[str, Any],
                          context: Optional[Dict[str, Any]] = None) -> bool:
        """Validate data against rules."""
        try:
            result = await self.chain_manager.validation_chain.arun(
                data=str(data),
                rules=str(rules),
                context=str(context or {})
            )
            return result.lower().startswith("valid")
            
        except Exception as e:
            logger.error(f"Validation failed: {str(e)}")
            return False
    
    async def analyze_formula(self, formula: str,
                            context: Optional[Dict[str, Any]] = None) -> str:
        """Analyze an Excel formula."""
        try:
            return await self.chain_manager.formula_chain.arun(
                formula=formula,
                context=str(context or {})
            )
            
        except Exception as e:
            logger.error(f"Formula analysis failed: {str(e)}")
            return f"Error analyzing formula: {str(e)}"
    
    def _is_complex_transformation(self, rules: Dict[str, Any]) -> bool:
        """Determine if a transformation requires the multi-agent system."""
        # Add logic to determine complexity based on rules
        return len(rules.get("steps", [])) > 1 or "llm_prompt" in rules

class ProcessingCallback(BaseCallbackHandler):
    """Callback handler for monitoring LangChain operations."""
    
    def on_llm_start(self, serialized: Dict[str, Any], prompts: List[str], **kwargs):
        """Log when LLM starts processing."""
        logger.info(f"Starting LLM operation with {len(prompts)} prompts")
    
    def on_llm_end(self, response, **kwargs):
        """Log when LLM completes processing."""
        logger.info("LLM operation completed")
    
    def on_chain_start(self, serialized: Dict[str, Any], inputs: Dict[str, Any], **kwargs):
        """Log when a chain starts processing."""
        logger.info(f"Starting chain operation: {serialized.get('name', 'Unknown chain')}")
    
    def on_chain_end(self, outputs: Dict[str, Any], **kwargs):
        """Log when a chain completes processing."""
        logger.info("Chain operation completed")
    
    def on_tool_start(self, serialized: Dict[str, Any], input_str: str, **kwargs):
        """Log when a tool starts processing."""
        logger.info(f"Starting tool operation: {serialized.get('name', 'Unknown tool')}")
    
    def on_tool_end(self, output: str, **kwargs):
        """Log when a tool completes processing."""
        logger.info("Tool operation completed")
    
    def on_agent_action(self, action, **kwargs):
        """Log when an agent takes an action."""
        logger.info(f"Agent taking action: {action}")
    
    def on_agent_finish(self, finish, **kwargs):
        """Log when an agent finishes processing."""
        logger.info("Agent finished processing")