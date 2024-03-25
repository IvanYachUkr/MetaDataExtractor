### Functional Requirements for Metadata Extraction System

#### General Operations
- **FR1.1**: The system shall allow users to extract metadata from specified document files via a command-line interface.
- **FR1.2**: The system shall support multiple document formats, including but not limited to DOCX, PPTX, and PDF.

#### Metadata Extraction
- **FR2.1**: The system shall provide the functionality to extract all available metadata from the document.
- **FR2.2**: The system shall enable users to extract specific categories of metadata, including:
  - **FR2.2.1**: Text-related properties (e.g., word count, paragraph count).
  - **FR2.2.2**: Main document properties (e.g., author, title, creation date).
  - **FR2.2.3**: Application-specific data (e.g., application name, version).
- **FR2.3**: The system shall offer a simplified metadata extraction mode, fetching only key metadata elements such as author, creation date, and modification date.

#### User Interaction
- **FR3.1**: The system shall provide a help option detailing usage instructions and available commands.
- **FR3.2**: The system shall display clear error messages for unsupported file formats, missing files, or incorrect command usage.

#### Extensibility
- **FR4.1**: The system shall be designed to facilitate easy integration of additional file format support in the future, such as PDF.

#### Output and Display
- **FR5.1**: The system shall present the extracted metadata in a structured and readable format to the user.
- **FR5.2**: In cases where metadata is not available or applicable, the system shall indicate this clearly in the output.

#### Data Handling
- **FR6.1**: The system shall ensure the integrity of the original document, with metadata extraction processes being read-only.
- **FR6.2**: The system shall handle exceptions and errors gracefully, ensuring stability and continuity of service.

#### Performance and Efficiency
- **FR7.1**: The system shall optimize metadata extraction processes to minimize processing time and resource consumption.
