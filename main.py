#!/usr/bin/env python3

import os
import sys
from subprocess import call

def file_exists(filepath):
    """Check if a file exists and is readable."""
    return os.path.isfile(filepath) and os.access(filepath, os.R_OK)

def main():
    # Check for minimum required arguments: the script name, filetype, and at least one additional argument (filepath or -h)
    if len(sys.argv) < 3:
        print("Usage: metad {docx,pdf,pptx} filepath [additional arguments]")
        sys.exit(1)

    filetype = sys.argv[1]
    next_arg = sys.argv[2]

    # Supported file types
    supported_types = ['docx', 'pdf', 'pptx']

    # Determine the extractor tool based on the file type
    if filetype not in supported_types:
        print(f"Error: Unsupported file type '{filetype}'. Supported types are 'docx', 'pdf', and 'pptx'.")
        sys.exit(1)

    extractor_tool = f"_{filetype}_meta.py"

    # If the next argument is -h, call the corresponding extractor's help directly
    if next_arg == '-h':
        command = ['python', extractor_tool, '-h']
    else:
        # If only filetype and filepath are provided, append '-a' as the default option
        if len(sys.argv) == 3:
            additional_args = ['-a']
        else:
            additional_args = sys.argv[3:]

        # Ensure the filepath exists and is accessible
        if not file_exists(next_arg):
            print(f"Error: The file '{next_arg}' does not exist or is not readable.")
            sys.exit(1)

        # Prepare the command to call the appropriate extractor tool with the filepath and any additional (or default) arguments
        command = ['python', extractor_tool, next_arg] + additional_args

    # Call the extractor tool
    sys.exit(call(command))

if __name__ == "__main__":
    main()
