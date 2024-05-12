#!/usr/bin/env python3

import argparse
import zipfile
import defusedxml.ElementTree as ET
from docx import Document
from datetime import datetime
import os
import re
import sys

class DocxMetadataHelper:
    """Class responsible for providing documentation for the command-line tool."""

    @staticmethod
    def print_help():
        help_text = """
        Usage: _docx_meta.py [options] file

        Options:
        -a, --all       Extract all metadata.
        -t, --textprops Extract text-related properties (pages, words, characters, keywords, language).
        -m, --mainprops Extract main properties (author, description, last modified by, revision, title, subject, creation date, modification date, language).
        -d, --appdata   Extract application data (application name and version).
        -s, --simple    Extract author, creation date, and modification date only.
        -h, --help      Display this help message and exit.

        Examples:
        _docx_meta.py -a my_document.docx
        _docx_meta.py -t -m my_document.docx
        """
        print(help_text)

def extract_xml_data(docx_path):
    metadata = {}
    with zipfile.ZipFile(docx_path, 'r') as docx:
        with docx.open('docProps/core.xml') as core_xml, docx.open('docProps/app.xml') as app_xml:
            core_tree, app_tree = ET.parse(core_xml), ET.parse(app_xml)
            core_root, app_root = core_tree.getroot(), app_tree.getroot()

            ns = {
                'cp': 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
                'dc': 'http://purl.org/dc/elements/1.1/',
                'dcterms': 'http://purl.org/dc/terms/',
                'xsi': 'http://www.w3.org/2001/XMLSchema-instance',
                'ep': 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties'
            }
            
            metadata['Author'] = core_root.find('dc:creator', ns).text if core_root.find('dc:creator', ns) is not None else 'Not Available'
            metadata['Description'] = core_root.find('dc:description', ns).text if core_root.find('dc:description', ns) is not None else 'Not Available'
            metadata['LastModifiedBy'] = core_root.find('cp:lastModifiedBy', ns).text if core_root.find('cp:lastModifiedBy', ns) is not None else 'Not Available'
            metadata['Revision'] = core_root.find('cp:revision', ns).text if core_root.find('cp:revision', ns) is not None else 'Not Available'
            metadata['Title'] = core_root.find('dc:title', ns).text if core_root.find('dc:title', ns) is not None else 'Not Available'
            metadata['Subject'] = core_root.find('dc:subject', ns).text if core_root.find('dc:subject', ns) is not None else 'Not Available'
            metadata['Created'] = core_root.find('dcterms:created', ns).text if core_root.find('dcterms:created', ns) is not None else 'Not Available'
            metadata['Modified'] = core_root.find('dcterms:modified', ns).text if core_root.find('dcterms:modified', ns) is not None else 'Not Available'
            metadata['Language'] = core_root.find('dc:language', ns).text if core_root.find('dc:language', ns) is not None else 'Not Available'
            
            metadata['Template'] = app_root.find('ep:Template', ns).text if app_root.find('ep:Template', ns) is not None else 'Not Available'
            metadata['TotalTime'] = app_root.find('ep:TotalTime', ns).text if app_root.find('ep:TotalTime', ns) is not None else 'Not Available'
            metadata['Pages'] = app_root.find('ep:Pages', ns).text if app_root.find('ep:Pages', ns) is not None else 'Not Available'
            metadata['Words'] = app_root.find('ep:Words', ns).text if app_root.find('ep:Words', ns) is not None else 'Not Available'
            metadata['Characters'] = app_root.find('ep:Characters', ns).text if app_root.find('ep:Characters', ns) is not None else 'Not Available'
            metadata['Application'] = app_root.find('ep:Application', ns).text if app_root.find('ep:Application', ns) is not None else 'Not Available'
            metadata['AppVersion'] = app_root.find('ep:AppVersion', ns).text if app_root.find('ep:AppVersion', ns) is not None else 'Not Available'

    return metadata

def filter_metadata(metadata, options):
    filtered_metadata = {}
    text_related_props = ['Pages', 'Words', 'Characters', 'Keywords', 'Language']
    main_props = ['Author', 'Description', 'LastModifiedBy', 'Revision', 'Title', 'Subject', 'Created', 'Modified', 'Language']
    app_data_props = ['Application', 'AppVersion']
    simple_props = ['Author', 'Created', 'Modified']

    def add_props(props_list):
        for prop in props_list:
            if prop in metadata:
                filtered_metadata[prop] = metadata[prop]

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
    parser = argparse.ArgumentParser(description="Extract metadata from a DOCX file.", add_help=False)
    parser.add_argument("-a", "--all", help="Extract all metadata.", action="store_true")
    parser.add_argument("-t", "--textprops", help="Extract text-related properties.", action="store_true")
    parser.add_argument("-m", "--mainprops", help="Extract main properties.", action="store_true")
    parser.add_argument("-d", "--appdata", help="Extract application data.", action="store_true")
    parser.add_argument("-s", "--simple", help="Extract author, creation date, and modification date only.", action="store_true")
    parser.add_argument("-h", "--help", action="store_true", help="Display help message and exit.")
    parser.add_argument("file", nargs='?', help="Path to the DOCX file.")

    args = parser.parse_args()

    if args.help or not args.file:
        DocxMetadataHelper.print_help()
        return

    args.file = args.file
    try:
        filepath = args.file 
        metadata = extract_xml_data(filepath)
    except (zipfile.BadZipFile, zipfile.LargeZipFile, KeyError, ET.ParseError, ValueError) as e:
        print("Failed to process the document. Ensure the file is a valid DOCX and path is safe.")
        sys.exit(1)

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

    filtered_metadata = filter_metadata(metadata, selected_options)
    
    for key, value in filtered_metadata.items():
        print(f"{key}: {value if value is not None else 'Not Available'}")

if __name__ == "__main__":
    main()

