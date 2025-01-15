import argparse
import os
from pathlib import Path
from dedaexcelai import migrate_excel
from dedaexcelai.logger import setup_logging, get_logger

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
    
    # Setup logging based on verbose flag
    setup_logging(args.verbose)
    logger = get_logger()
    
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
