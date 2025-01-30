"""Example script demonstrating basic usage of the Excel Migration Framework."""
import os
from pathlib import Path
import logging
from dotenv import load_dotenv

from excel_migration.core.models import MigrationContext
from excel_migration.core.processor import ExcelMigrationProcessor
from excel_migration.rules.engine import RuleEngine

# Load environment variables from .env file
load_dotenv()

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

def main():
    """Run a simple Excel migration example."""
    try:
        # Get API key from environment
        api_key = os.getenv('OPENAI_API_KEY')
        if not api_key:
            raise ValueError("OPENAI_API_KEY environment variable is required")

        # Initialize the rule engine with OpenAI
        rule_engine = RuleEngine(
            llm_provider="openai",
            api_key=api_key,
            model_name="gpt-4",
            temperature=0.7
        )

        # Get the current directory
        current_dir = Path(__file__).parent

        # Load rules from JSON file
        rules = rule_engine.load_rules(current_dir / "rules.json")

        # Create migration context
        context = MigrationContext(
            source_file=str(current_dir / "source.xlsx"),
            target_file=str(current_dir / "target.xlsx"),
            rules=rules,
            sheet_mapping={
                "Sheet1": "Output1",
                "Sheet2": "Output2"
            }
        )

        # Execute migration
        processor = ExcelMigrationProcessor(context)
        success = processor.process()

        if success:
            logger.info("Migration completed successfully!")
            logger.info(f"Output file created: {context.target_file}")
        else:
            logger.error("Migration failed!")

    except Exception as e:
        logger.error(f"Error during migration: {str(e)}")
        raise

if __name__ == "__main__":
    main()