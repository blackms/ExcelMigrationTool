"""LangChain agents for Excel data processing."""
from typing import Any, Dict, List, Optional
from langchain.agents import AgentType, initialize_agent
from langchain.agents.tools import Tool
from langchain.chains import LLMChain
from langchain.prompts import ChatPromptTemplate
from langchain_core.language_models import BaseLanguageModel
from langchain.tools import BaseTool
import logging

logger = logging.getLogger(__name__)

class ExcelTools:
    """Collection of tools for Excel data processing."""
    
    @staticmethod
    def create_formula_analyzer() -> Tool:
        """Create a tool for analyzing Excel formulas."""
        return Tool(
            name="analyze_formula",
            description="Analyze an Excel formula and explain its logic",
            func=lambda formula: f"Formula analysis: {formula}..."
        )
    
    @staticmethod
    def create_data_validator() -> Tool:
        """Create a tool for validating Excel data."""
        return Tool(
            name="validate_data",
            description="Validate Excel data against specified rules",
            func=lambda data, rules: f"Validation result for {data} against {rules}..."
        )
    
    @staticmethod
    def create_text_transformer() -> Tool:
        """Create a tool for transforming text data."""
        return Tool(
            name="transform_text",
            description="Transform text data according to specified rules",
            func=lambda text, rules: f"Transformed text: {text}..."
        )

class ExcelAgent:
    """Agent for handling complex Excel operations."""
    
    def __init__(self, llm: BaseLanguageModel):
        """Initialize with a language model."""
        self.llm = llm
        self.tools = self._setup_tools()
        self.agent = self._create_agent()
    
    def _setup_tools(self) -> List[Tool]:
        """Set up the tools available to the agent."""
        return [
            ExcelTools.create_formula_analyzer(),
            ExcelTools.create_data_validator(),
            ExcelTools.create_text_transformer()
        ]
    
    def _create_agent(self):
        """Create the agent with tools."""
        return initialize_agent(
            tools=self.tools,
            llm=self.llm,
            agent=AgentType.CHAT_CONVERSATIONAL_REACT_DESCRIPTION,
            verbose=True
        )
    
    async def process_task(self, task: str, context: Dict[str, Any]) -> Any:
        """Process a task using the agent."""
        try:
            result = await self.agent.arun(
                input=task,
                context=str(context)
            )
            return result
        except Exception as e:
            logger.error(f"Agent task processing failed: {str(e)}")
            return None

class MultiAgentSystem:
    """System for coordinating multiple agents."""
    
    def __init__(self, llm: BaseLanguageModel):
        """Initialize with a language model."""
        self.llm = llm
        self.agents = self._setup_agents()
        self._setup_coordination_chain()
    
    def _setup_agents(self) -> Dict[str, ExcelAgent]:
        """Set up specialized agents."""
        return {
            "formula": ExcelAgent(self.llm),
            "validation": ExcelAgent(self.llm),
            "transformation": ExcelAgent(self.llm)
        }
    
    def _setup_coordination_chain(self):
        """Set up the coordination chain."""
        self.coordinator_prompt = ChatPromptTemplate.from_messages([
            ("system", """You are a coordinator for Excel data processing tasks.
            Analyze the task and determine which specialized agent should handle it.
            Available agents: formula, validation, transformation.
            Output the agent name and subtask in format: AGENT:subtask"""),
            ("user", "{task}")
        ])
        
        self.coordinator = LLMChain(
            llm=self.llm,
            prompt=self.coordinator_prompt
        )
    
    async def process_task(self, task: str, context: Optional[Dict[str, Any]] = None) -> Any:
        """Process a task using the appropriate agent."""
        try:
            # Determine which agent should handle the task
            coordination_result = await self.coordinator.arun(task=task)
            agent_name, subtask = coordination_result.split(":", 1)
            
            if agent_name.lower().strip() not in self.agents:
                raise ValueError(f"Unknown agent: {agent_name}")
            
            # Execute the task with the chosen agent
            agent = self.agents[agent_name.lower().strip()]
            return await agent.process_task(subtask.strip(), context or {})
            
        except Exception as e:
            logger.error(f"Multi-agent task processing failed: {str(e)}")
            return None

    async def analyze_task(self, context: Dict[str, Any]) -> Dict[str, Any]:
        """Analyze a task context and provide insights."""
        try:
            # Extract relevant information from context
            sheet_analysis = context.get("sheet_analysis", {})
            mapping = context.get("mapping", {})
            
            # Create analysis prompt
            analysis_prompt = ChatPromptTemplate.from_messages([
                ("system", """You are an expert at analyzing Excel data structures and migrations.
                Analyze the provided sheet structure and mapping to provide insights for migration.
                Consider data types, formulas, and potential transformation needs."""),
                ("user", """Sheet Analysis: {sheet_analysis}
                Mapping: {mapping}
                Context: {context}""")
            ])
            
            analysis_chain = LLMChain(
                llm=self.llm,
                prompt=analysis_prompt
            )
            
            # Get analysis
            result = await analysis_chain.arun(
                sheet_analysis=str(sheet_analysis),
                mapping=str(mapping),
                context=str(context)
            )
            
            # Parse and structure the analysis
            return {
                "insights": result,
                "recommendations": self._extract_recommendations(result),
                "warnings": self._extract_warnings(result)
            }
            
        except Exception as e:
            logger.error(f"Task analysis failed: {str(e)}")
            return {
                "insights": "Analysis failed",
                "recommendations": [],
                "warnings": [str(e)]
            }
    
    def _extract_recommendations(self, analysis: str) -> List[str]:
        """Extract recommendations from analysis text."""
        # Simple extraction - split on newlines and look for recommendation-like statements
        recommendations = []
        for line in analysis.split('\n'):
            line = line.strip().lower()
            if any(word in line for word in ['recommend', 'suggest', 'should', 'could']):
                recommendations.append(line)
        return recommendations
    
    def _extract_warnings(self, analysis: str) -> List[str]:
        """Extract warnings from analysis text."""
        # Simple extraction - split on newlines and look for warning-like statements
        warnings = []
        for line in analysis.split('\n'):
            line = line.strip().lower()
            if any(word in line for word in ['warning', 'caution', 'careful', 'note']):
                warnings.append(line)
        return warnings

class AgentFactory:
    """Factory for creating specialized agents."""
    
    @staticmethod
    def create_agent(agent_type: str, llm: BaseLanguageModel) -> ExcelAgent:
        """Create a specialized agent."""
        agent = ExcelAgent(llm)
        
        # Add specialized tools based on agent type
        if agent_type == "formula":
            agent.tools.append(Tool(
                name="advanced_formula_analysis",
                description="Advanced analysis of Excel formulas including optimization suggestions",
                func=lambda formula: f"Advanced analysis: {formula}..."
            ))
        elif agent_type == "validation":
            agent.tools.append(Tool(
                name="advanced_validation",
                description="Advanced data validation with custom rules and error reporting",
                func=lambda data, rules: f"Advanced validation: {data}..."
            ))
        elif agent_type == "transformation":
            agent.tools.append(Tool(
                name="advanced_transformation",
                description="Advanced data transformation with custom rules and formatting",
                func=lambda data, rules: f"Advanced transformation: {data}..."
            ))
        else:
            raise ValueError(f"Unknown agent type: {agent_type}")
        
        return agent