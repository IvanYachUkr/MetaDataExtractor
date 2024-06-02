## MetaD Software Application Functional Requirements

### Introduction
The MetaD software application is a command-line tool designed to extract metadata from various document formats, including DOCX, PDF, and PPTX files. The tool supports extracting different types of metadata based on user-defined options and ensures security by addressing common vulnerabilities.

### Functional Requirements

#### 1. General Requirements

1.1 The application shall support the following file types for metadata extraction:
- DOCX
- PDF
- PPTX

1.2 The application shall be executable via the command line.

1.3 The application shall log errors and important events to the standard output using a consistent logging format.

1.4 The application shall sanitize file paths to prevent directory traversal attacks by:
- Converting file paths to absolute paths.
- Ensuring file paths are within the current working directory.

1.5 The application shall check if a file exists and is readable before processing it by verifying:
- The file is present in the file system.
- The file has read permissions.

1.6 The application shall check if a file is empty (0 bytes) before processing it.

1.7 On Windows, the application shall prevent the execution of batch files (*.bat and *.cmd) to avoid security risks by:
- Checking the file extension.
- Exiting with an error if a batch file is detected.

1.8 The application shall exit with appropriate exit codes based on different error scenarios to facilitate debugging and automation.

1.9 The application shall limit the size of the files and their metadata to prevent zip bombs:
- The extraction size limit shall be 20% of the system's RAM.
- The metadata size limit shall be 10% of the system's RAM.
- If `psutil` is not available, the extraction size limit shall default to 1 GB, and the metadata size limit shall default to 200 MB.

1.10 The application shall limit overall RAM usage to 10% of the system's RAM when it can be done:
- Use the `resource` module on Unix systems.
- Use `psutil` for real-time monitoring on systems where `resource` is unavailable.
- Log a warning if neither method is available.

1.11 The application shall use `defusedxml.ElementTree` to mitigate XML-related attacks such as XML bombs and XML injection.

1.12 The application shall avoid command line injection by:
- Constructing commands as lists of arguments.
- Not using `shell=True`.

#### 2. Command-Line Interface

2.1 The application shall accept the following command-line arguments:
- File type (`docx`, `pdf`, `pptx`)
- File path
- Additional arguments for metadata extraction options

2.2 The application shall display a help message when the `-h` or `--help` option is used, providing usage instructions and examples.

2.3 The application shall process the additional arguments for specific metadata extraction options based on the file type.

#### 3. Metadata Extraction for DOCX Files

3.1 The application shall extract metadata from DOCX files using `_docx_meta.py`.

3.2 The application shall support the following metadata extraction options for DOCX files:
- `-a`, `--all`: Extract all metadata.
- `-t`, `--textprops`: Extract text-related properties (pages, words, characters, keywords, language).
- `-m`, `--mainprops`: Extract main properties (author, description, last modified by, revision, title, subject, creation date, modification date, language).
- `-d`, `--appdata`: Extract application data (application name and version).
- `-s`, `--simple`: Extract author, creation date, and modification date only.

3.3 The application shall limit the size of the DOCX file and its metadata to prevent zip bombs:
- The extraction size limit shall be 20% of the system's RAM.
- The metadata size limit shall be 10% of the system's RAM.
- If `psutil` is not available, the extraction size limit shall default to 1 GB, and the metadata size limit shall default to 200 MB.

3.4 The application shall handle errors related to:
- Invalid ZIP files.
- Exceeding decompression size limits.
- XML parsing issues.

#### 4. Metadata Extraction for PDF Files

4.1 The application shall extract metadata from PDF files using `_pdf_meta.py`.

4.2 The application shall support the following metadata extraction options for PDF files:
- `-a`, `--all`: Extract all metadata.
- `-t`, `--textprops`: Extract text-related properties (title, subject, page count).
- `-m`, `--mainprops`: Extract main properties (author, creator, creation date, modification date, subject, title).
- `-d`, `--appdata`: Extract application data (producer, creator).
- `-s`, `--simple`: Extract author, creation date, and modification date only.

4.3 The application shall limit the size of the PDF file and its metadata to prevent large files from causing issues:
- The extraction size limit shall be 20% of the system's RAM.
- The metadata size limit shall be 10% of the system's RAM.
- If `psutil` is not available, the extraction size limit shall default to 1 GB, and the metadata size limit shall default to 200 MB.

4.4 The application shall handle errors related to:
- File reading issues.
- Metadata extraction failures.
- File size limit exceedances.

#### 5. Metadata Extraction for PPTX Files

5.1 The application shall extract metadata from PPTX files using `_pptx_meta.py`.

5.2 The application shall support the following metadata extraction options for PPTX files:
- `-a`, `--all`: Extract all metadata.
- `-t`, `--textprops`: Extract text-related properties (words, paragraphs, slides, notes, hidden slides, multimedia clips).
- `-m`, `--mainprops`: Extract main properties (author, last modified by, revision, creation date, modification date).
- `-d`, `--appdata`: Extract application data (application name, presentation format, total editing time, template).
- `-s`, `--simple`: Extract author, creation date, and modification date only.

5.3 The application shall limit the size of the PPTX file and its metadata to prevent zip bombs:
- The extraction size limit shall be 20% of the system's RAM.
- The metadata size limit shall be 10% of the system's RAM.
- If `psutil` is not available, the extraction size limit shall default to 1 GB, and the metadata size limit shall default to 200 MB.

5.4 The application shall handle errors related to:
- Invalid ZIP files.
- Exceeding decompression size limits.
- XML parsing issues.

#### 6. Error Handling

6.1 The application shall exit with code `1` for general errors, including invalid arguments, file path issues, and unexpected errors.

6.2 The application shall exit with specific exit codes for different error scenarios:
- Exit code `2` for unsupported file types.
- Exit code `3` for invalid file paths (directory traversal).
- Exit code `4` for non-existent or unreadable files.
- Exit code `5` for empty files.

6.3 The application shall log detailed error messages to the standard output to help diagnose issues, including the specific error encountered and potential solutions.


#### 7. Help and Documentation

7.1 The application shall implement a general help option that can be triggered using the `-h` or `--help` argument.

7.2 Each file type class (DOCX, PDF, PPTX) shall implement specific help messages that detail the available options and usage examples:
- The DOCX help message shall include usage instructions and examples for extracting DOCX metadata.
- The PDF help message shall include usage instructions and examples for extracting PDF metadata.
- The PPTX help message shall include usage instructions and examples for extracting PPTX metadata.

7.3 The help messages shall be clear, concise, and provide sufficient information for users to understand how to use the application effectively.
