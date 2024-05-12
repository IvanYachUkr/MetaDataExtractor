#!/usr/bin/env python3

import os
import sys
import logging
from subprocess import call

# Setup basic logging
logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')

def sanitize_path(filepath):
    """Sanitize the file path to prevent directory traversal."""
    # Prevent directory traversal by collapsing any .. and preventing absolute paths.
    return os.path.normpath(os.path.join(os.getcwd(), filepath))

def file_exists(filepath):
    """Check if a file exists and is readable."""
    try:
        return os.path.isfile(filepath) and os.access(filepath, os.R_OK)
    except TypeError:
        logging.error(f"Invalid file path: {filepath}")
        return False

def is_file_empty(filepath):
    """Check if the file is empty (0 bytes)."""
    try:
        return os.stat(filepath).st_size == 0
    except OSError as e:
        logging.error(f"Error checking if file is empty: {e}")
        return True  # Treat error as empty file to prevent usage

def main():
    # Documenting supported file types
    supported_types = {
        'docx': '_docx_meta.py',
        'pdf': '_pdf_meta.py',
        'pptx': '_pptx_meta.py',  
    }

    if len(sys.argv) < 3:
        logging.error("Usage: metad {docx,pdf,pptx} filepath [additional arguments]\nTo get help for a specific tool use command: metad {docx,pdf,pptx} -h")
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

        # Sanitize the file path to prevent directory traversal attacks
        sanitized_path = sanitize_path(next_arg)
        if not file_exists(sanitized_path):
            logging.error(f"The file '{sanitized_path}' does not exist or is not readable.")
            sys.exit(1)
        if is_file_empty(sanitized_path):
            logging.error(f"The file '{sanitized_path}' is empty.")
            sys.exit(1)
        command = ['python', extractor_tool, sanitized_path] + additional_args

    result = call(command)
    if result != 0:
        logging.error("An error occurred while executing the extractor tool.")
    sys.exit(result)

if __name__ == "__main__":
    main()
