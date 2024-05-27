# Functional Requirements for Metadata Extraction System

## 1. Overview

This document outlines the functional requirements for a software system designed to extract metadata from supported document types. The system aims to facilitate secure, efficient, and user-friendly extraction of metadata.

## 2. System Overview

### 2.1 Purpose

The purpose of this system is to extract and display metadata from specified types of documents using a command-line interface (CLI).

### 2.2 Scope

The system supports metadata extraction from the following document types:
- DOCX (Microsoft Word Document)
- PDF (Portable Document Format)
- PPTX (Microsoft PowerPoint Presentation)

The system is limited to these document types and is designed to allow for future extensions.

## 3. Functional Requirements

### 3.1 Metadata Extraction

**FR1.1**: The system must support metadata extraction for three specific document types: DOCX, PDF, and PPTX.

**FR1.2**: The system must extract metadata based on the command-line input which specifies the document type, the file path, and optionally, the specific type of metadata desired.

**FR1.3**: For DOCX files, the system must safely parse using `defusedxml.ElementTree` to prevent XML External Entity (XXE) attacks and similar XML-related vulnerabilities.

**FR1.4**: For PDF files, the system must use `PyPDF2.PdfReader` to read and extract metadata safely and accurately.

**FR1.5**: For PPTX files, the system must use `defusedxml.ElementTree` for secure XML parsing to prevent security vulnerabilities related to XML processing.

### 3.2 Command-Line Interface (CLI)

**FR2.1**: The system must provide a CLI for users to specify the operation mode, which includes choosing the document type and providing the file path.

**FR2.2**: The system must provide detailed help and usage instructions when the `-h` or `--help` option is used.

**FR2.3**: If invoked with insufficient or incorrect command-line arguments, the system must show usage information and exit with a status code of 1.

### 3.3 File Handling and Validation

**FR3.1**: The system must sanitize input file paths using `os.path.normpath` and other relevant functions to prevent directory traversal attacks. If the sanitized path leads to a file outside the intended directory, the system must log an error and exit with a status code of 1.

**FR3.2**: The system must check if the specified file exists and is readable. If not, it must log an error and exit with status code 1.

**FR3.3**: If the file is empty (0 bytes), the system must log an error stating "The file '{filename}' is empty." and exit with status code 1.

### 3.4 Security and Safety

**FR4.1**: The system must adhere to secure XML parsing practices with `defusedxml.ElementTree` for DOCX and PPTX files to mitigate risks such as XXE attacks.

**FR4.2**: The system must handle all operational errors and exceptions gracefully, ensuring that sensitive error information or system details are not exposed to the user.

**FR4.3**: The system must execute metadata extraction scripts securely using `subprocess.call` with a list to avoid shell injection vulnerabilities.

**FR4.4**: The system must ensure that dynamically constructed paths or commands are securely composed to prevent injection attacks and other common vulnerabilities.

### 3.5 Metadata Extraction Options by Document Type

#### 3.5.1 DOCX Metadata Extraction Options

**FR5.1**: For DOCX, the system should allow users to extract:
- All metadata (`-a` or `--all`)
- Text-related properties: Words, Paragraphs (`-t` or `--textprops`)
- Main properties: Author, Last Modified By, Revision, Creation Date, Modification Date (`-m` or `--mainprops`)
- Application data: Application Name, App Version, Template, Total Editing Time (`-d` or `--appdata`)
- Simple properties: Author, Creation Date, Modification Date (`-s` or `--simple`)

#### 3.5.2 PDF Metadata Extraction Options

**FR5.2**: For PDF, the system should support extracting:
- All metadata (`-a` or `--all`)
- Text-related properties: Title, Subject, Number of Pages (`-t` or `--textprops`)
- Main properties: Author, Creator, Creation Date, Modification Date, Subject, Title (`-m` or `--mainprops`)
- Application data: Producer, Creator (`-d` or `--appdata`)
- Simple properties: Author, Creation Date, Modification Date (`-s` or `--simple`)

#### 3.5.3 PPTX Metadata Extraction Options

**FR5.3**: For PPTX, the system should enable users to extract:
- All metadata (`-a` or `--all`)
- Text-related properties: Words, Paragraphs, Slides, Notes, Hidden Slides, Multimedia Clips (`-t` or `--textprops`)
- Main properties: Author, Last Modified By, Revision, Creation Date, Modification Date (`-m` or `--mainprops`)
- Application data: Application Name, Presentation Format, Total Editing Time, Template, App Version (`-d` or `--appdata`)
- Simple properties: Author, Creation Date, Modification Date (`-s` or `--simple`)

### 3.6 Error Handling

**FR6.1**: The system must handle errors related to file parsing, missing metadata fields, and invalid files.

**FR6.2**: For DOCX and PPTX files, errors related to invalid XML structure and file integrity issues must be handled appropriately.

**FR6.3**: For PDF files, errors related to invalid or missing metadata fields must be handled gracefully.

### 3.7 Help Documentation

**FR7.1**: The system must provide comprehensive help documentation accessible through the `-h` or `--help` argument.

**FR7.2**: Detailed CLI usage instructions and options must be provided for each supported document type.

This document captures all the functional requirements necessary to build a secure and efficient metadata extraction system with robust error handling and user-friendly CLI.
