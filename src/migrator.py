#!/usr/bin/env python3
import argparse
import os
from dedaexcelai import migrate_excel, get_logger
from typing import Optional
from dedaexcelai.logger import setup_logging

logger = get_logger()

def main():
    """Main entry point for the Excel migration tool"""
    # Get the default template path relative to this script
    default_template = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'template', 'template.xlsx')
    
    parser = argparse.ArgumentParser(description='Excel Migration Tool')
    parser.add_argument('-i', '--input', required=True, help='Input Excel file')
    parser.add_argument('-o', '--output', required=True, help='Output Excel file')
    parser.add_argument('-t', '--template', default=default_template, help='Template Excel file (default: template/template.xlsx)')
    parser.add_argument('-v', '--verbose', action='store_true', help='Enable verbose logging')
    parser.add_argument('--openai-key', help='OpenAI API key for GG Startup analysis. If not provided, will try OPENAI_API_KEY environment variable')
    
    args = parser.parse_args()
    
    # Configure logging based on verbose flag
    setup_logging(args.verbose)
    
    # Validate input file exists
    if not os.path.exists(args.input):
        logger.error(f"Input file does not exist: {args.input}")
        return 1
    
    # Validate template file exists
    if not os.path.exists(args.template):
        logger.error(f"Template file does not exist: {args.template}")
        return 1
    
    # Get OpenAI API key from argument or environment
    openai_key: Optional[str] = args.openai_key or os.getenv('OPENAI_API_KEY')
    
    # Perform migration
    success = migrate_excel(args.input, args.output, args.template, openai_key)
    
    if not success:
        return 1
    return 0

if __name__ == "__main__":
    exit(main())
