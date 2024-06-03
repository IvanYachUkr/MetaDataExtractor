
# MetaDataExtractor

MetaDataExtractor is a Python-based tool designed to extract metadata from DOCX, PDF, and PPTX files. This user manual provides detailed instructions on how to install, configure, and use the tool effectively.

## Table of Contents
- [Installation](#installation)
- [Usage](#usage)
  - [Command-Line Arguments](#command-line-arguments)
  - [Examples](#examples)
- [Modules](#modules)
  - [main.py](#mainpy)
  - [_docx_meta.py](#_docx_metapy)
  - [_pdf_meta.py](#_pdf_metapy)
  - [_pptx_meta.py](#_pptx_metapy)
- [Contributing](#contributing)
- [License](#license)

## Installation

1. **Clone the Repository**
   ```bash
   git clone https://github.com/IvanYachUkr/MetaDataExtractor.git
   cd MetaDataExtractor
   ```

2. **Create and Activate a Virtual Environment**
   ```bash
   python3 -m venv venv
   source venv/bin/activate  # On Windows use `venv\Scripts\activate`
   ```

3. **Install Dependencies**
   ```bash
   pip install -r requirements.txt
   pip install -r optional_requirements.txt
   ```

## Usage

### Command-Line Arguments

- `-a, --all`: Extract all metadata.
- `-t, --textprops`: Extract text-related properties (pages, words, characters, keywords, language).
- `-m, --mainprops`: Extract main properties (author, description, last modified by, revision, title, subject, creation date, modification date, language).
- `-d, --appdata`: Extract application data (application name and version).
- `-s, --simple`: Extract author, creation date, and modification date only.
- `-h, --help`: Display help message and exit.

### Examples

1. **Extract All Metadata**
   ```bash
   ./metad.sh docx path/to/document -a
   ```

2. **Extract Text Properties**
   ```bash
   ./metad.sh docx path/to/document -t
   ```

3. **Extract Main Properties**
   ```bash
   ./metad.sh docx path/to/document -m
   ```

4. **Extract Application Data**
   ```bash
   ./metad.sh docx path/to/document -d
   ```

5. **Extract Simple Properties**
   ```bash
   ./metad.sh docx path/to/document -s
   ```

## Modules

### main.py

The main module initializes the command-line interface and handles user inputs to extract metadata from specified document files. It uses the following arguments to determine the type of metadata to extract:

- `-a, --all`: Calls all metadata extraction functions.
- `-t, --textprops`: Calls functions to extract text-related properties.
- `-m, --mainprops`: Calls functions to extract main properties.
- `-d, --appdata`: Calls functions to extract application data.
- `-s, --simple`: Calls functions to extract basic properties (author, creation date, modification date).

### _docx_meta.py

This module handles metadata extraction from DOCX files. It uses XML parsing to extract the following metadata:

- Author
- Description
- Last Modified By
- Revision
- Title
- Subject
- Creation Date
- Modification Date
- Language
- Application Name
- Application Version
- Pages
- Words
- Characters

**Usage Example:**
```python
./metad.sh docx path/to/document.docx
```

### _pdf_meta.py

This module handles metadata extraction from PDF files. It extracts the following metadata:

- Author
- Title
- Subject
- Keywords
- Creator
- Producer
- Creation Date
- Modification Date

**Usage Example:**
```python
./metad.sh pdf path/to/document.pdf
```

### _pptx_meta.py

This module handles metadata extraction from PPTX files. It uses XML parsing to extract the following metadata:

- Author
- Title
- Subject
- Keywords
- Revision Number
- Last Modified By
- Creation Date
- Modification Date
- Application Name
- Application Version
- Total Time
- Slides Count
- Notes Count
- Hidden Slides Count

**Usage Example:**
```python
./metad.sh pptx path/to/document.pptx
```


## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for more details.
```
