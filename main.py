#!/usr/bin/env python3

import os
import sys
import logging
from subprocess import call

# Setup basic logging
logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')

def file_exists(filepath):
    """Check if a file exists and is readable."""
    return os.path.isfile(filepath) and os.access(filepath, os.R_OK)

def is_file_empty(filepath):
    """Check if the file is empty (0 bytes)."""
    return os.stat(filepath).st_size == 0

def main():
    # Documenting supported file types
    supported_types = {
        'docx': '_docx_meta.py',
        'pdf': '_pdf_meta.py',
        'pptx': '_pptx_meta.py',  # Ensure this script exists and is correctly implemented
    }

    if len(sys.argv) < 3:
        logging.error("Usage: metad {docx,pdf,pptx} filepath [additional arguments]\n To get help for a specific tool use command metad {docx,pdf,pptx} -h")
        sys.exit(1)

    filetype, next_arg = sys.argv[1], sys.argv[2]

    if filetype not in supported_types:
        logging.error(f"Unsupported file type '{filetype}'. Supported types are: {', '.join(supported_types.keys())}.")
        sys.exit(1)

    extractor_tool = supported_types[filetype]

    if next_arg == '-h':
        command = ['python', extractor_tool, '-h']
    else:
        additional_args = ['-a'] if len(sys.argv) == 3 else sys.argv[3:]
        if not file_exists(next_arg):
            logging.error(f"The file '{next_arg}' does not exist or is not readable.")
            sys.exit(1)
        if is_file_empty(next_arg):
            logging.error(f"The file '{next_arg}' is empty.")
            sys.exit(1)
        command = ['python', extractor_tool, next_arg] + additional_args

    result = call(command)
    if result != 0:
        logging.error("An error occurred while executing the extractor tool.")
    sys.exit(result)

if __name__ == "__main__":
    main()
