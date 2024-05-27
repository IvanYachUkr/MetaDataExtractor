# MetaDataExtractor

MetaDataExtractor is a Python-based tool designed to extract metadata from DOCX, PDF, and PPTX files. This user manual provides detailed instructions on how to install, configure, and use the tool effectively.

## Table of Contents
- [Installation](#installation)
- [Usage](#usage)
  - [Command-Line Arguments](#command-line-arguments)
  - [Examples](#examples)
- [Modules](#modules)
  - [main.py](#mainpy)
  - [_docx_meta.py](#docx_metapy)
  - [_pdf_meta.py](#pdf_metapy)
  - [_pptx_meta.py](#pptx_metapy)
- [Contributing](#contributing)
- [License](#license)

## Installation

1. **Clone the Repository**
   ``` 
   git clone https://github.com/IvanYachUkr/MetaDataExtractor.git
   cd MetaDataExtractor
Create and Activate a Virtual Environment

 
 
python3 -m venv venv
source venv/bin/activate  # On Windows use `venv\Scripts\activate`
Install Dependencies

 
 
pip install -r requirements.txt
Usage
Command-Line Arguments
-a, --all: Extract all metadata.
-t, --textprops: Extract text-related properties (pages, words, characters, keywords, language).
-m, --mainprops: Extract main properties (author, description, last modified by, revision, title, subject, creation date, modification date, language).
-d, --appdata: Extract application data (application name and version).
-s, --simple: Extract author, creation date, and modification date only.
-h, --help: Display help message and exit.

Examples
Extract All Metadata

 
 
python main.py -a path/to/document
Extract Text Properties

 
 
python main.py -t path/to/document
Extract Main Properties

 
 
python main.py -m path/to/document
Extract Application Data

 
 
python main.py -d path/to/document
Extract Simple Properties

 
 
python main.py -s path/to/document
Modules
main.py
The main module initializes the command-line interface and handles user inputs to extract metadata from specified document files. It uses the following arguments to determine the type of metadata to extract:

-a, --all: Calls all metadata extraction functions.
-t, --textprops: Calls functions to extract text-related properties.
-m, --mainprops: Calls functions to extract main properties.
-d, --appdata: Calls functions to extract application data.
-s, --simple: Calls functions to extract basic properties (author, creation date, modification date).
_docx_meta.py
This module handles metadata extraction from DOCX files. It uses XML parsing to extract the following metadata:

Author
Description
Last Modified By
Revision
Title
Subject
Creation Date
Modification Date
Language
Application Name
Application Version
Pages
Words
Characters
Usage Example:

from _docx_meta import extract_docx_metadata
metadata = extract_docx_metadata("path/to/document.docx")
print(metadata)
_pdf_meta.py
This module handles metadata extraction from PDF files. It extracts the following metadata:

Author
Title
Subject
Keywords
Creator
Producer
Creation Date
Modification Date
Usage Example:

from _pdf_meta import extract_pdf_metadata
metadata = extract_pdf_metadata("path/to/document.pdf")
print(metadata)
_pptx_meta.py
This module handles metadata extraction from PPTX files. It uses XML parsing to extract the following metadata:

Author
Title
Subject
Keywords
Revision Number
Last Modified By
Creation Date
Modification Date
Application Name
Application Version
Total Time
Slides Count
Notes Count
Hidden Slides Count
Usage Example:
 
from _pptx_meta import extract_pptx_metadata
metadata = extract_pptx_metadata("path/to/document.pptx")
print(metadata)
Contributing
Contributions are welcome! Please read the contributing guidelines for more details.

License
This project is licensed under the MIT License. See the LICENSE file for more details.
