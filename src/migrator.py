import argparse
import os
from pathlib import Path
from loguru import logger
from dedaexcelai import migrate_excel

# Configure logger
logger.remove()  # Remove default handler
logger.add(
    "migration.log",
    rotation="1 day",
    retention="7 days",
    level="DEBUG",
    format="{time:YYYY-MM-DD HH:mm:ss} | {level} | {message}",
    colorize=False  # No colors in file
)
logger.add(
    lambda msg: print(msg),
    level="INFO",
    format="<level>{level: <8}</level> | <cyan>{message}</cyan>",
    colorize=True  # Enable colors in console
)

def parse_args():
    parser = argparse.ArgumentParser(description='Excel Migration Tool')
    parser.add_argument('--input', '-i', required=True, help='Input Excel file path')
    parser.add_argument('--output', '-o', required=True, help='Output Excel file path')
    parser.add_argument('--template', '-t', default=str(Path(__file__).parent.parent / 'template' / 'template.xlsx'),
                       help='Template Excel file path (default: template/template.xlsx)')
    parser.add_argument('--verbose', '-v', action='store_true', help='Enable verbose logging')
    return parser.parse_args()

def main():
    args = parse_args()
    
    # Adjust log level if verbose flag is set
    if args.verbose:
        logger.remove()
        logger.add(
            "migration.log",
            rotation="1 day",
            retention="7 days",
            level="DEBUG",
            format="{time:YYYY-MM-DD HH:mm:ss} | {level} | {message}",
            colorize=False
        )
        logger.add(
            lambda msg: print(msg),
            level="DEBUG",
            format="<blue>{time:HH:mm:ss}</blue> | <level>{level: <8}</level> | <cyan>{message}</cyan>",
            colorize=True
        )
    
    # Validate input file exists
    if not os.path.exists(args.input):
        logger.error(f"Input file '{args.input}' does not exist")
        return 1
        
    # Validate template file exists
    if not os.path.exists(args.template):
        logger.error(f"Template file '{args.template}' does not exist")
        return 1
    
    # Execute migration
    success = migrate_excel(args.input, args.output, args.template)
    
    if success:
        logger.success("Migration completed successfully!")
        logger.info(f"Output file: {args.output}")
        return 0
    else:
        logger.error("Migration failed!")
        return 1

if __name__ == '__main__':
    exit(main())
