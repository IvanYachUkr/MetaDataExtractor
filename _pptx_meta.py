#!/usr/bin/env python3

import argparse
import zipfile
from xml.etree import ElementTree as ET
from pptx import Presentation
from datetime import datetime


class PptxMetadataHelper:
    """Class responsible for providing documentation for the command-line tool."""

    @staticmethod
    def print_help():
        help_text = """
        Usage: metad file options

        Options:
        -a, --all       Extract all metadata.
        -t, --textprops Extract text-related properties (words, paragraphs, slides, notes, hidden slides, multimedia clips).
        -m, --mainprops Extract main properties (author, last modified by, revision, creation date, modification date).
        -d, --appdata   Extract application data (application name, presentation format, total editing time, template).
        -s, --simple    Extract author, creation date, and modification date only.
        -h, --help      Display this help message and exit.

        Examples:
        metad pptx example.pptx -a -t
        metad pptx help
        """
        print(help_text)

def extract_xml_data(pptx_path):
    metadata = {}
    with zipfile.ZipFile(pptx_path, 'r') as pptx:
        # Extract core.xml and app.xml and parse their contents
        with pptx.open('docProps/core.xml') as core_xml, pptx.open('docProps/app.xml') as app_xml:
            core_tree, app_tree = ET.parse(core_xml), ET.parse(app_xml)
            core_root, app_root = core_tree.getroot(), app_tree.getroot()

            # Namespace for core.xml and app.xml
            ns = {
                'cp': 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
                'dc': 'http://purl.org/dc/elements/1.1/',
                'dcterms': 'http://purl.org/dc/terms/',
                'xsi': 'http://www.w3.org/2001/XMLSchema-instance',
                'ep': 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties'
            }
            
            # Extract metadata from core.xml
            metadata['Author'] = core_root.find('dc:creator', ns).text
            metadata['LastModifiedBy'] = core_root.find('cp:lastModifiedBy', ns).text
            metadata['Revision'] = core_root.find('cp:revision', ns).text
            metadata['Created'] = core_root.find('dcterms:created', ns).text
            metadata['Modified'] = core_root.find('dcterms:modified', ns).text
            
            # Extract metadata from app.xml
            metadata['Template'] = app_root.find('ep:Template', ns).text if app_root.find('ep:Template', ns) is not None else None
            metadata['TotalTime'] = app_root.find('ep:TotalTime', ns).text if app_root.find('ep:TotalTime', ns) is not None else None
            metadata['Words'] = app_root.find('ep:Words', ns).text
            metadata['Paragraphs'] = app_root.find('ep:Paragraphs', ns).text
            metadata['Slides'] = app_root.find('ep:Slides', ns).text
            metadata['Notes'] = app_root.find('ep:Notes', ns).text
            metadata['HiddenSlides'] = app_root.find('ep:HiddenSlides', ns).text
            metadata['MMClips'] = app_root.find('ep:MMClips', ns).text
            metadata['Application'] = app_root.find('ep:Application', ns).text
            metadata['PresentationFormat'] = app_root.find('ep:PresentationFormat', ns).text
            metadata['AppVersion'] = app_root.find('ep:AppVersion', ns).text

    return metadata

def filter_metadata(metadata, options):
    filtered_metadata = {}
    
    # Define metadata groupings
    text_related_props = ['Words', 'Paragraphs', 'Slides', 'Notes', 'HiddenSlides', 'MMClips']
    main_props = ['Author', 'LastModifiedBy', 'Revision', 'Created', 'Modified']
    app_data_props = ['Application', 'PresentationFormat', 'TotalTime', 'Template', 'AppVersion']
    simple_props = ['Author', 'Created', 'Modified']
    
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
            add_props(text_related_props)
        if 'mainprops' in options:
            add_props(main_props)
        if 'appdata' in options:
            add_props(app_data_props)
        if 'simple' in options:
            add_props(simple_props)
    
    return filtered_metadata

def main():
    # Setup argument parser
    parser = argparse.ArgumentParser(description="Extract metadata from a PPTX file.", add_help=False)
    parser.add_argument("-a", "--all", help="Extract all metadata.", action="store_true")
    parser.add_argument("-t", "--textprops", help="Extract text-related properties.", action="store_true")
    parser.add_argument("-m", "--mainprops", help="Extract main properties.", action="store_true")
    parser.add_argument("-d", "--appdata", help="Extract application data.", action="store_true")
    parser.add_argument("-s", "--simple", help="Extract author, creation date, and modification date only.", action="store_true")
    parser.add_argument("-h", "--help", action="store_true", help="Display help message and exit.")
    parser.add_argument("file", nargs='?', help="Path to the PPTX file.")

    # Parse arguments
    args = parser.parse_args()

    # Handle help option
    if args.help or not args.file:  # Display help if -h is used or no file is provided
        PptxMetadataHelper.print_help()
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
    metadata = extract_xml_data(args.file)
    filtered_metadata = filter_metadata(metadata, selected_options)
    
    # Print filtered metadata in a more readable format
    for key, value in filtered_metadata.items():
        print(f"{key}: {value if value is not None else 'Not Available'}")

if __name__ == "__main__":
    main()
