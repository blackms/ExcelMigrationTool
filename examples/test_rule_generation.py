"""Test script demonstrating rule generation from example files."""
import asyncio
from pathlib import Path
from loguru import logger
import json
import os
from dotenv import load_dotenv
from langchain_openai import ChatOpenAI

from excel_migration.tasks.base import MigrationTask, TaskBasedProcessor, SheetMapping
from excel_migration.llm.agents import MultiAgentSystem
from excel_migration.rules.engine import RuleEngine
from excel_migration.rules.executor import RuleExecutor
from excel_migration.core.analyzers import ExcelSheetAnalyzer
from excel_migration.vision.processor import SheetImageProcessor

# Load environment variables
load_dotenv()

async def generate_and_test_rules():
    """Generate rules from example files and test them."""
    try:
        # Setup paths
        data_dir = Path(__file__).parent / "data"
        source_file = data_dir / "source.xlsx"
        target_file = data_dir / "target.xlsx"
        rules_file = data_dir / "generated_rules.json"

        # Initialize LLM
        llm = ChatOpenAI(
            model_name=os.getenv("OPENAI_MODEL", "gpt-4"),
            temperature=float(os.getenv("OPENAI_TEMPERATURE", "0.7"))
        )

        # Initialize components
        llm_system = MultiAgentSystem(llm)
        image_processor = SheetImageProcessor()
        rule_engine = RuleEngine(llm_provider="openai")
        sheet_analyzer = ExcelSheetAnalyzer(image_processor)
        rule_executor = RuleExecutor()
        
        # Create task for rule generation
        generation_task = MigrationTask(
            source_file=source_file,
            target_file=target_file,
            task_type="analyze",
            description="Generate rules from example files",
            context={
                "llm_provider": "openai",
                "model": os.getenv("OPENAI_MODEL", "gpt-4"),
                "rule_executor": rule_executor
            },
            sheet_mappings=[
                SheetMapping(
                    source_sheet="CustomerData",
                    target_sheet="CustomerSummary"
                ),
                SheetMapping(
                    source_sheet="Transactions",
                    target_sheet="TransactionSummary"
                )
            ],
            # Use the same files as examples
            example_source=source_file,
            example_target=target_file,
            example_sheet_mappings=[
                SheetMapping(
                    source_sheet="CustomerData",
                    target_sheet="CustomerSummary"
                ),
                SheetMapping(
                    source_sheet="Transactions",
                    target_sheet="TransactionSummary"
                )
            ]
        )

        # Generate rules
        logger.info("🔍 Analyzing example files to generate rules...")
        processor = TaskBasedProcessor(
            rule_generator=rule_engine,
            sheet_analyzer=sheet_analyzer,
            llm_provider=llm_system
        )

        success = await processor.process(generation_task)
        if not success:
            logger.error("❌ Rule generation failed!")
            return

        # Save generated rules
        rules = generation_task.context.get("generated_rules", [])
        if not rules:
            logger.error("❌ No rules were generated!")
            return

        with open(rules_file, 'w') as f:
            json.dump(rules, f, indent=2)
        logger.info(f"✨ Generated rules saved to: {rules_file}")

        # Print example rules
        logger.info("\n🔍 Example of generated rules:")
        for i, rule in enumerate(rules[:2], 1):
            logger.info(f"\nRule {i}:")
            logger.info(json.dumps(rule, indent=2))

        # Test the generated rules
        logger.info("\n🧪 Testing generated rules...")
        
        # Create new output file for testing
        test_output = data_dir / "test_output.xlsx"
        
        test_task = MigrationTask(
            source_file=source_file,
            target_file=test_output,
            task_type="migrate",
            description="Test generated rules",
            context={
                "llm_provider": "openai",
                "model": os.getenv("OPENAI_MODEL", "gpt-4"),
                "rules": rules,
                "rule_executor": rule_executor
            },
            sheet_mappings=[
                SheetMapping(
                    source_sheet="CustomerData",
                    target_sheet="CustomerSummary"
                ),
                SheetMapping(
                    source_sheet="Transactions",
                    target_sheet="TransactionSummary"
                )
            ]
        )

        success = await processor.process(test_task)
        if success:
            logger.info(f"✅ Rules tested successfully! Output saved to: {test_output}")
        else:
            logger.error("❌ Rule testing failed!")

        # Compare results
        logger.info("\n📊 Rule Analysis:")
        logger.info("Generated rules include:")
        
        rule_types = {}
        for rule in rules:
            rule_type = rule.get("type", "unknown")
            rule_types[rule_type] = rule_types.get(rule_type, 0) + 1
        
        for rule_type, count in rule_types.items():
            logger.info(f"- {count} {rule_type} rules")

        # Provide summary
        logger.info("\n📝 Summary:")
        logger.info(f"- Source sheets: {', '.join(m.source_sheet for m in test_task.sheet_mappings)}")
        logger.info(f"- Target sheets: {', '.join(m.target_sheet for m in test_task.sheet_mappings)}")
        logger.info(f"- Total rules generated: {len(rules)}")
        logger.info(f"- Rules file: {rules_file}")
        logger.info(f"- Test output: {test_output}")

    except Exception as e:
        logger.exception(f"❌ Error during rule generation and testing: {str(e)}")

def main():
    """Run the example."""
    # Configure logging
    logger.remove()
    logger.add(
        lambda msg: print(msg),
        format="<green>{time:HH:mm:ss}</green> | {message}",
        level="INFO"
    )

    # Print header
    logger.info("=" * 60)
    logger.info("🚀 Excel Migration Framework - Rule Generation Example")
    logger.info("=" * 60)

    # Run example
    asyncio.run(generate_and_test_rules())

if __name__ == "__main__":
    main()