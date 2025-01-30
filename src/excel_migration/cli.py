"""Command-line interface for Excel migration framework."""
import argparse
from pathlib import Path
from typing import List, Optional
import sys
import asyncio
from loguru import logger

from .tasks.base import MigrationTask, TaskRegistry, TaskBasedProcessor
from .core.interfaces import RuleGenerator, SheetAnalyzer
from .llm.agents import MultiAgentSystem
from .vision.processor import SheetImageProcessor

def setup_logging(log_level: str = "INFO", log_file: Optional[str] = None):
    """Configure logging with loguru."""
    # Remove default handler
    logger.remove()
    
    # Add console handler with custom format
    logger.add(
        sys.stderr,
        format="<green>{time:YYYY-MM-DD HH:mm:ss}</green> | "
               "<level>{level: <8}</level> | "
               "<cyan>{name}</cyan>:<cyan>{function}</cyan>:<cyan>{line}</cyan> | "
               "<level>{message}</level>",
        level=log_level
    )
    
    # Add file handler if specified
    if log_file:
        logger.add(
            log_file,
            rotation="10 MB",
            retention="1 week",
            compression="zip",
            level=log_level
        )

def parse_args() -> argparse.Namespace:
    """Parse command line arguments."""
    parser = argparse.ArgumentParser(
        description="Excel Migration Framework CLI",
        formatter_class=argparse.RawDescriptionHelpFormatter
    )
    
    # File arguments
    parser.add_argument(
        "source",
        type=Path,
        help="Source Excel file path"
    )
    
    parser.add_argument(
        "target",
        type=Path,
        help="Target Excel file path"
    )
    
    # Sheet selection
    parser.add_argument(
        "--source-sheets",
        nargs="+",
        help="Specific sheets to process from source file"
    )
    
    parser.add_argument(
        "--target-sheets",
        nargs="+",
        help="Corresponding target sheet names"
    )
    
    parser.add_argument(
        "--example-source-sheets",
        nargs="+",
        help="Specific sheets to use from example source file"
    )
    
    parser.add_argument(
        "--example-target-sheets",
        nargs="+",
        help="Corresponding sheets from example target file"
    )
    
    # Task configuration
    parser.add_argument(
        "--task-type",
        choices=["migrate", "analyze", "validate"],
        default="migrate",
        help="Type of task to perform"
    )
    
    parser.add_argument(
        "--example-source",
        type=Path,
        help="Example source file for rule generation"
    )
    
    parser.add_argument(
        "--example-target",
        type=Path,
        help="Example target file for rule generation"
    )
    
    # Visual analysis
    parser.add_argument(
        "--screenshots",
        type=Path,
        nargs="+",
        help="Screenshots of Excel sheets for additional analysis"
    )
    
    parser.add_argument(
        "--screenshot-sheet-mapping",
        nargs="+",
        help="Mapping of screenshots to sheet names (format: screenshot.png:SheetName)"
    )
    
    # Rules and LLM
    parser.add_argument(
        "--rules",
        type=Path,
        help="JSON file containing migration rules"
    )
    
    parser.add_argument(
        "--llm-provider",
        choices=["openai", "anthropic"],
        default="openai",
        help="LLM provider to use"
    )
    
    parser.add_argument(
        "--model",
        help="Specific model to use with the LLM provider"
    )
    
    # Logging and debug
    parser.add_argument(
        "--log-level",
        choices=["DEBUG", "INFO", "WARNING", "ERROR"],
        default="INFO",
        help="Logging level"
    )
    
    parser.add_argument(
        "--log-file",
        type=str,
        help="Log file path"
    )
    
    parser.add_argument(
        "--config",
        type=Path,
        help="Configuration file path"
    )
    
    parser.add_argument(
        "--cache-dir",
        type=Path,
        help="Directory for caching results"
    )
    
    parser.add_argument(
        "--no-cache",
        action="store_true",
        help="Disable caching"
    )
    
    parser.add_argument(
        "--debug",
        action="store_true",
        help="Enable debug mode"
    )
    
    args = parser.parse_args()
    
    # Validate sheet mappings
    if args.source_sheets and args.target_sheets:
        if len(args.source_sheets) != len(args.target_sheets):
            parser.error("Number of source and target sheets must match")
    
    if args.example_source_sheets and args.example_target_sheets:
        if len(args.example_source_sheets) != len(args.example_target_sheets):
            parser.error("Number of example source and target sheets must match")
    
    if args.screenshot_sheet_mapping:
        try:
            screenshot_mappings = [mapping.split(":") for mapping in args.screenshot_sheet_mapping]
            args.screenshot_map = {screenshot: sheet for screenshot, sheet in screenshot_mappings}
        except ValueError:
            parser.error("Invalid screenshot mapping format. Use 'screenshot.png:SheetName'")
    
    return args

async def run_task(args: argparse.Namespace) -> bool:
    """Run a migration task with the provided arguments."""
    try:
        # Set up components
        image_processor = SheetImageProcessor()
        llm_system = MultiAgentSystem(
            provider=args.llm_provider,
            model=args.model
        )
        
        # Create task processor
        processor = TaskBasedProcessor(
            rule_generator=RuleGenerator(),
            sheet_analyzer=SheetAnalyzer(image_processor),
            llm_provider=llm_system
        )
        
        # Create task registry
        registry = TaskRegistry()
        
        # Prepare sheet mappings
        sheet_mapping = {}
        if args.source_sheets and args.target_sheets:
            sheet_mapping = dict(zip(args.source_sheets, args.target_sheets))
        
        example_sheet_mapping = {}
        if args.example_source_sheets and args.example_target_sheets:
            example_sheet_mapping = dict(zip(args.example_source_sheets, args.example_target_sheets))
        
        # Create task
        task = MigrationTask(
            source_file=args.source,
            target_file=args.target,
            task_type=args.task_type,
            description=f"Migrate from {args.source} to {args.target}",
            context={
                "llm_provider": args.llm_provider,
                "model": args.model,
                "debug": args.debug,
                "sheet_mapping": sheet_mapping,
                "example_sheet_mapping": example_sheet_mapping,
                "screenshot_mapping": getattr(args, "screenshot_map", {})
            },
            example_source=args.example_source,
            example_target=args.example_target,
            screenshots=args.screenshots
        )
        
        # Get handler
        handler = await registry.get_handler(task)
        if not handler:
            logger.error(f"No handler found for task type: {args.task_type}")
            return False
        
        # Execute task
        success = await handler.handle(task)
        
        if success:
            logger.success("Task completed successfully")
        else:
            logger.error("Task failed")
        
        return success
        
    except Exception as e:
        logger.exception(f"Task execution failed: {str(e)}")
        return False

def main():
    """Main entry point for the CLI."""
    args = parse_args()
    
    # Set up logging
    setup_logging(args.log_level, args.log_file)
    
    if args.debug:
        logger.debug("Arguments: {}", vars(args))
    
    # Run task
    success = asyncio.run(run_task(args))
    
    # Exit with appropriate code
    sys.exit(0 if success else 1)

if __name__ == "__main__":
    main()