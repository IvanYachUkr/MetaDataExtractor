#!/usr/bin/env python3

import argparse
from PyPDF2 import PdfReader

class PdfMetadataHelper:
    """
    Class responsible for providing documentation for the command-line tool.
    """
    @staticmethod
    def print_help():
        help_text = """
        Usage: _pdf_meta.py [options] file

        Options:
        -a, --all       Extract all metadata.
        -t, --textprops Extract text-related properties (title, subject, page count).
        -m, --mainprops Extract main properties (author, creator, creation date, modification date, subject, title).
        -d, --appdata   Extract application data (producer, creator).
        -s, --simple    Extract author, creation date, and modification date only.
        -h, --help      Display this help message and exit.

        Examples:
        _pdf_meta.py -a my_document.docx
        _pdf_meta.py -t -m my_document.docx
        """
        print(help_text)

def extract_pdf_data(pdf_path):
    reader = PdfReader(pdf_path)
    metadata = dict(reader.metadata)
    metadata["/Pages"] = f"{len(reader.pages)}"
    return dict(metadata)

def filter_metadata(metadata, options):
    # Define metadata groupings
    text_md = ['/Title', '/Subject', '/Pages']
    main_md = ['/Author', '/Creator', '/CreationDate', '/ModDate', '/Subject',
               '/Title']
    app_md = ['/Producer', '/Creator']
    simple_md = ['/Author', '/CreationDate', '/ModDate']

    filtered_metadata = dict()

    # Helper function to add selected props to filtered_metadata
    def add_props(props_list):
        for prop in props_list:
            if prop in metadata:
                filtered_metadata[prop] = metadata[prop]
    
    # Filter based on selected options
    if 'all' in options:
        filtered_metadata = metadata
    else:
        if 'textprops' in options:
            add_props(text_md)
        if 'mainprops' in options:
            add_props(main_md)
        if 'appdata' in options:
            add_props(app_md)
        if 'simple' in options:
            add_props(simple_md)
    
    return filtered_metadata

def main():
    # Setup argument parser
    parser = argparse.ArgumentParser(
            description="Extract metadata from a PDF file.", add_help=False)
    parser.add_argument(
            "-a", "--all", help="Extract all metadata.", action="store_true")
    parser.add_argument(
            "-t", "--textprops", help="Extract text-related properties.",
            action="store_true")
    parser.add_argument(
            "-m", "--mainprops", help="Extract main properties.",
            action="store_true")
    parser.add_argument(
            "-d", "--appdata", help="Extract application data.",
            action="store_true")
    parser.add_argument(
            "-s", "--simple",
            help="Extract author, creation date, and modification date only.",
            action="store_true")
    parser.add_argument(
            "-h", "--help", action="store_true",
            help="Display help message and exit.")
    parser.add_argument("file", nargs='?', help="Path to the PDF file.")

    # Parse arguments
    args = parser.parse_args()

    # Handle help option
    if args.help or not args.file: # also display help if no file provided
        PdfMetadataHelper.print_help()
        return

    # Collect selected options
    selected_options = []
    if args.all:
        selected_options.append('all')
    if args.textprops:
        selected_options.append('textprops')
    if args.mainprops:
        selected_options.append('mainprops')
    if args.appdata:
        selected_options.append('appdata')
    if args.simple:
        selected_options.append('simple')

    # Extract and filter metadata
    metadata = extract_pdf_data(args.file)
    filtered_metadata = filter_metadata(metadata, selected_options)
    
    # Print filtered metadata in a more readable format
    for key, value in filtered_metadata.items():
        print(f"{key}: {value if value is not None else 'Not Available'}")

if __name__ == "__main__":
    main()
