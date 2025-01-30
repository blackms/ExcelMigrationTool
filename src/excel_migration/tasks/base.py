"""Base task implementations for Excel migration framework."""
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional
from loguru import logger

from ..core.interfaces import (
    Task, TaskHandler, TaskProcessor, RuleGenerator,
    SheetAnalyzer, DataExtractor, ImageProcessor
)

@dataclass
class MigrationTask:
    """Base migration task implementation."""
    source_file: Path
    target_file: Path
    task_type: str
    description: str
    context: Dict[str, Any]
    example_source: Optional[Path] = None
    example_target: Optional[Path] = None
    screenshots: List[Path] = None

    def __post_init__(self):
        """Initialize task with defaults."""
        self.screenshots = self.screenshots or []
        self.context = self.context or {}
        logger.info(f"Initialized {self.task_type} task: {self.description}")

class TaskRegistry:
    """Registry for task handlers."""
    
    def __init__(self):
        self._handlers: Dict[str, TaskHandler] = {}
    
    def register(self, task_type: str, handler: TaskHandler):
        """Register a handler for a task type."""
        self._handlers[task_type] = handler
        logger.debug(f"Registered handler for task type: {task_type}")
    
    async def get_handler(self, task: Task) -> Optional[TaskHandler]:
        """Get appropriate handler for a task."""
        for handler in self._handlers.values():
            if await handler.can_handle(task):
                return handler
        return None

class BaseTaskHandler(TaskHandler):
    """Base implementation of task handler."""
    
    def __init__(self, processor: TaskProcessor):
        self.processor = processor
        logger.debug(f"Initialized {self.__class__.__name__}")
    
    async def can_handle(self, task: Task) -> bool:
        """Check if this handler can process the task."""
        return await self.processor.validate(task)
    
    async def handle(self, task: Task) -> bool:
        """Handle the task."""
        try:
            logger.info(f"Processing task: {task.description}")
            return await self.processor.process(task)
        except Exception as e:
            logger.exception(f"Task handling failed: {str(e)}")
            return False

class ExampleBasedRuleGenerator(RuleGenerator):
    """Generate rules by analyzing example files."""
    
    def __init__(self, sheet_analyzer: SheetAnalyzer, llm_provider: Any):
        self.sheet_analyzer = sheet_analyzer
        self.llm_provider = llm_provider
        logger.debug("Initialized example-based rule generator")
    
    async def generate_rules(self, source: Path, target: Path) -> List[Dict[str, Any]]:
        """Generate rules by analyzing source and target files."""
        try:
            # Analyze source and target sheets
            source_analysis = await self.sheet_analyzer.analyze_sheet(source)
            target_analysis = await self.sheet_analyzer.analyze_sheet(target)
            
            # Generate rules using LLM
            rules = await self.llm_provider.generate_rules(
                source_analysis,
                target_analysis
            )
            
            logger.info(f"Generated {len(rules)} rules from example files")
            return rules
            
        except Exception as e:
            logger.exception(f"Rule generation failed: {str(e)}")
            return []

class MultimodalSheetAnalyzer(SheetAnalyzer):
    """Analyze sheets using both direct access and image processing."""
    
    def __init__(self, image_processor: ImageProcessor, data_extractor: DataExtractor):
        self.image_processor = image_processor
        self.data_extractor = data_extractor
        logger.debug("Initialized multimodal sheet analyzer")
    
    async def analyze_sheet(self, sheet_path: Path) -> Dict[str, Any]:
        """Analyze a sheet using multiple approaches."""
        try:
            analysis = {
                "path": str(sheet_path),
                "direct_analysis": await self._analyze_direct(sheet_path),
                "image_analysis": await self._analyze_image(sheet_path),
                "extracted_data": await self._extract_data(sheet_path)
            }
            
            logger.info(f"Completed multimodal analysis of {sheet_path}")
            return analysis
            
        except Exception as e:
            logger.exception(f"Sheet analysis failed: {str(e)}")
            return {"error": str(e)}
    
    async def _analyze_direct(self, sheet_path: Path) -> Dict[str, Any]:
        """Analyze sheet through direct file access."""
        # Implement direct Excel file analysis
        pass
    
    async def _analyze_image(self, sheet_path: Path) -> Dict[str, Any]:
        """Analyze sheet through image processing."""
        # Convert sheet to image and process
        pass
    
    async def _extract_data(self, sheet_path: Path) -> Dict[str, Any]:
        """Extract data from sheet image."""
        # Extract data from sheet image
        pass

class TaskBasedProcessor(TaskProcessor):
    """Process tasks with rule generation and multimodal analysis."""
    
    def __init__(self, 
                 rule_generator: RuleGenerator,
                 sheet_analyzer: SheetAnalyzer,
                 llm_provider: Any):
        self.rule_generator = rule_generator
        self.sheet_analyzer = sheet_analyzer
        self.llm_provider = llm_provider
        logger.debug("Initialized task-based processor")
    
    async def process(self, task: Task) -> bool:
        """Process a task."""
        try:
            # Generate rules if example files are provided
            if task.example_source and task.example_target:
                rules = await self.rule_generator.generate_rules(
                    task.example_source,
                    task.example_target
                )
                task.context["generated_rules"] = rules
            
            # Analyze source file
            source_analysis = await self.sheet_analyzer.analyze_sheet(task.source_file)
            task.context["source_analysis"] = source_analysis
            
            # Process screenshots if available
            if task.screenshots:
                screenshot_analyses = []
                for screenshot in task.screenshots:
                    analysis = await self._process_screenshot(screenshot)
                    screenshot_analyses.append(analysis)
                task.context["screenshot_analyses"] = screenshot_analyses
            
            # Get task-specific insights from LLM
            insights = await self.llm_provider.analyze_task(task)
            task.context["llm_insights"] = insights
            
            logger.info(f"Completed processing task: {task.description}")
            return True
            
        except Exception as e:
            logger.exception(f"Task processing failed: {str(e)}")
            return False
    
    async def validate(self, task: Task) -> bool:
        """Validate task requirements."""
        try:
            # Validate file existence
            if not task.source_file.exists():
                logger.error(f"Source file not found: {task.source_file}")
                return False
            
            # Validate screenshots if provided
            if task.screenshots:
                for screenshot in task.screenshots:
                    if not screenshot.exists():
                        logger.error(f"Screenshot not found: {screenshot}")
                        return False
            
            # Validate example files if provided
            if task.example_source and not task.example_source.exists():
                logger.error(f"Example source file not found: {task.example_source}")
                return False
            if task.example_target and not task.example_target.exists():
                logger.error(f"Example target file not found: {task.example_target}")
                return False
            
            return True
            
        except Exception as e:
            logger.exception(f"Task validation failed: {str(e)}")
            return False
    
    async def _process_screenshot(self, screenshot: Path) -> Dict[str, Any]:
        """Process a screenshot for additional insights."""
        try:
            # Analyze screenshot using image processor
            image_analysis = await self.image_processor.process_image(screenshot)
            
            # Extract any table data
            table_data = await self.image_processor.extract_table(image_analysis["image"])
            
            return {
                "path": str(screenshot),
                "analysis": image_analysis,
                "table_data": table_data
            }
            
        except Exception as e:
            logger.exception(f"Screenshot processing failed: {str(e)}")
            return {"error": str(e)}