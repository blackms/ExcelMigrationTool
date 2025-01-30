"""Command-line interface for Excel migration framework."""
import argparse
from pathlib import Path
from typing import List, Optional
import sys
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
    
    parser.add_argument(
        "--screenshots",
        type=Path,
        nargs="+",
        help="Screenshots of Excel sheets for additional analysis"
    )
    
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
    
    return parser.parse_args()

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
        
        # Create task
        task = MigrationTask(
            source_file=args.source,
            target_file=args.target,
            task_type=args.task_type,
            description=f"Migrate from {args.source} to {args.target}",
            context={
                "llm_provider": args.llm_provider,
                "model": args.model,
                "debug": args.debug
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