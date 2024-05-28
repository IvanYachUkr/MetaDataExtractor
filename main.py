
import os
import sys
import logging
import subprocess

# Setup basic logging
logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')

def sanitize_path(filepath):
    """Sanitize the file path to prevent directory traversal."""
    abs_path = os.path.abspath(filepath)
    # Ensure the file path is within the current working directory
    if not abs_path.startswith(os.getcwd()):
        raise ValueError("Invalid file path: potential directory traversal attack.")
    return abs_path

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

def is_batch_file(filepath):
    """Check if the file is a Windows batch file."""
    return filepath.lower().endswith(('.bat', '.cmd'))

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

        try:
            # Sanitize the file path to prevent directory traversal attacks
            sanitized_path = sanitize_path(next_arg)
        except ValueError as e:
            logging.error(str(e))
            sys.exit(2)

        if not file_exists(sanitized_path):
            logging.error(f"The file '{sanitized_path}' does not exist or is not readable.")
            sys.exit(3)
        if is_file_empty(sanitized_path):
            logging.error(f"The file '{sanitized_path}' is empty.")
            sys.exit(4)

        # On Windows, avoid direct execution of batch files to prevent security risks
        if os.name == 'nt' and is_batch_file(sanitized_path):
            logging.error("Execution of batch files is not allowed for security reasons.")
            sys.exit(5)

        command = ['python', extractor_tool, sanitized_path] + additional_args

    try:
        result = subprocess.run(command, check=True)
    except subprocess.CalledProcessError as e:
        logging.error(f"An error occurred while executing the extractor tool: {e}")
        sys.exit(e.returncode)
    except Exception as e:
        logging.error(f"An unexpected error occurred: {e}")
        sys.exit(1)

    sys.exit(0)

if __name__ == "__main__":
    main()
