"""Command line interface for Excel migration."""
import argparse
from dedaexcelai import migrate_excel, get_logger

logger = get_logger()

def main():
    """Main entry point."""
    parser = argparse.ArgumentParser(description='Migrate Excel file to new format.')
    parser.add_argument('-i', '--input', required=True, help='Input Excel file path')
    parser.add_argument('-o', '--output', required=True, help='Output Excel file path')
    parser.add_argument('-v', '--verbose', action='store_true', help='Enable verbose logging')
    parser.add_argument('--openai-key', help='OpenAI API key for startup days analysis')
    
    args = parser.parse_args()
    
    if migrate_excel(args.input, args.output, args.openai_key):
        logger.info("Migration completed successfully")
    else:
        logger.error("Migration failed")
        exit(1)

if __name__ == '__main__':
    main()
