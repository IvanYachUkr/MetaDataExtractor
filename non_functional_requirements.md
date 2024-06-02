## MetaD Software Application Non-Functional Requirements

### Introduction
The MetaD software application is a command-line tool designed to extract metadata from DOCX, PDF, and PPTX files. 
### Non-Functional Requirements

#### 1. Performance Requirements

1.1 **Processing Time**:
- The application shall extract metadata from files and provide the results within a reasonable time frame, typically within a few seconds for standard-sized documents (up to 10 MB).

1.2 **Resource Utilization**:
- The application shall efficiently use system resources (CPU, memory) to ensure that it does not significantly degrade the performance of the host system.

1.3 **Scalability**:
- The application shall be able to handle multiple files in succession without significant performance degradation.

#### 2. Security Requirements

2.1 **File Path Sanitization**:
- The application shall sanitize all file paths to prevent directory traversal attacks. This includes converting file paths to absolute paths and ensuring they are within the current working directory.

2.2 **Zip Bomb Prevention**:
- The application shall try to limit the overall RAM usage by a default amount of 10% of the system's RAM.
- As an additional safety measure, the application shall limit the size of extracted files and metadata:
  - The extraction size limit shall be 20% of the system's RAM.
  - The metadata size limit shall be 10% of the system's RAM.
  - If `psutil` is not available, the extraction size limit shall default to 1 GB, and the metadata size limit shall default to 200 MB.
- These extraction and metadata size limits serve as the main security measure if it is not possible to set a global RAM limit.

2.3 **XML Bomb and Injection Mitigation**:
- The application shall use `defusedxml.ElementTree` for XML parsing to mitigate XML-related attacks such as XML bombs and XML injection.

2.4 **Command Line Injection Prevention**:
- The application shall construct commands as lists of arguments and avoid using `shell=True` to prevent command line injection attacks.

2.5 **Executable File Execution Prevention**:
- On Windows, the application shall prevent the execution of batch files (`*.bat` and `*.cmd`) to avoid security risks.

2.6 **File Existence and Readability Checks**:
- The application shall verify that files exist and are readable before processing them.

2.7 **Empty File Handling**:
- The application shall check if files are empty (0 bytes) and prevent processing if they are.

#### 3. Usability Requirements

3.1 **Command-Line Interface**:
- The application shall provide a user-friendly command-line interface with clear and concise error messages.

3.2 **Help and Documentation**:
- The application shall provide a general help option (`-h` or `--help`) and specific help messages for each file type class (DOCX, PDF, PPTX). These help messages shall include usage instructions and examples.

3.3 **Error Messages**:
- The application shall log detailed error messages to the standard output to help diagnose issues, including the specific error encountered and potential solutions.

#### 4. Reliability Requirements

4.1 **Error Handling**:
- The application shall handle errors gracefully, providing appropriate error messages and exiting with specific exit codes for different error scenarios.

4.2 **Consistency**:
- The application shall ensure consistent behavior and output across different environments and document types.

4.3 **Robustness**:
- The application shall be robust against invalid inputs, corrupted files, and other unexpected conditions.

#### 5. Maintainability Requirements

5.1 **Code Quality**:
- The application code shall follow best practices for readability, modularity, and documentation to facilitate maintenance and future enhancements. The code shall adhere to the PEP 8 style guide.

5.2 **Logging**:
- The application shall log errors and important events to aid in debugging and maintaining the application.

5.3 **Modular Design**:
- The application shall be designed in a modular fashion, with separate modules for different file types and functionalities, making it easier to update and maintain.

#### 6. Portability Requirements

6.1 **Platform Independence**:
- The application shall be platform-independent and run on any operating system that supports Python. Examples of supported operating systems include Windows, macOS, iPadOS, and Linux.

6.2 **Dependency Management**:
- The application shall manage dependencies appropriately, ensuring that all required libraries are available and properly documented.

#### 7. Documentation Requirements

7.1 **Documentation**:
- The application shall provide documentation for all supported file types (DOCX, PDF, PPTX). This documentation shall include:
  - Usage instructions
  - Examples of command-line options
  - Detailed descriptions of available metadata extraction options
