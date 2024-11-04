import subprocess
import sys
import pytest
from unittest.mock import patch, MagicMock
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QApplication, QPushButton, QLabel, QComboBox, QTextEdit, QCheckBox
import os
import logging

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
import app

# ================== Logging Configuration ==================

# Define the results directory relative to this script
results_dir = os.path.join(os.path.dirname(__file__), 'ui_test_results')
os.makedirs(results_dir, exist_ok=True)

# Log file paths
all_log_file = os.path.join(results_dir, 'all_messages.log')
error_log_file = os.path.join(results_dir, 'error_messages.log')

# Create a custom logger
logger = logging.getLogger('ui_test_logger')
logger.setLevel(logging.DEBUG)  # Capture all levels of logs

# Formatter for all logs
all_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')

# Formatter for console output
console_formatter = logging.Formatter('%(message)s')

# File handler for all messages
fh_all = logging.FileHandler(all_log_file, mode='w', encoding='utf-8')
fh_all.setLevel(logging.DEBUG)
fh_all.setFormatter(all_formatter)
logger.addHandler(fh_all)

# File handler for error messages
fh_error = logging.FileHandler(error_log_file, mode='w', encoding='utf-8')
fh_error.setLevel(logging.ERROR)
fh_error.setFormatter(all_formatter)
logger.addHandler(fh_error)

# Console handler for final status
ch = logging.StreamHandler(sys.stdout)
ch.setLevel(logging.INFO)
ch.setFormatter(console_formatter)
logger.addHandler(ch)

# Flag to track if any errors occurred
error_occurred = False

# ================== Pytest Hooks ==================

def pytest_sessionfinish(session, exitstatus):
    """
    Hook to execute after all tests are run.
    Writes "No errors" to error log if no errors occurred.
    Prints final status to the console.
    """
    global error_occurred
    if exitstatus == 0:
        logger.info("Tests completed successfully.")
        with open(error_log_file, 'a', encoding='utf-8') as ef:
            ef.write("No errors\n")
    else:
        logger.error("Errors occurred during tests.")
        error_occurred = True

# ================== Fixtures ==================

@pytest.fixture
def qtbot_app(qtbot):
    """Fixture to create the application and return the qtbot and the main window."""
    test_app = app.MetadataExtractorApp()
    test_app.show()
    qtbot.addWidget(test_app)
    return qtbot, test_app

@pytest.fixture
def mock_subprocess_run():
    """Fixture to mock subprocess.run."""
    with patch('subprocess.run') as mock_run:
        yield mock_run

# ================== Test Metadata Outputs ==================

# Sample metadata outputs
ALL_METADATA_DOCX = """Author: Учетная запись Майкрософт;Ivan Iachnyk
Description: Not Available
LastModifiedBy: Учетная запись Майкрософт
Revision: 3
Title: Not Available
Subject: Not Available
Created: 2024-03-16T20:21:00Z
Modified: 2024-03-16T20:26:00Z
Language: Not Available
Template: Normal.dotm
TotalTime: 5
Pages: 2
Words: 573
Characters: 3272
Application: Microsoft Office Word
AppVersion: 15.0000
"""

ALL_METADATA_PPTX = """Author: Учетная запись Майкрософт;Ivan Iachnyk
LastModifiedBy: Учетная запись Майкрософт
Revision: 1
Created: 2024-03-17T12:03:51Z
Modified: 2024-03-17T12:06:22Z
Template: Integral
TotalTime: 2
Words: 7
Paragraphs: 7
Slides: 2
Notes: 0
HiddenSlides: 0
MMClips: 0
Application: Microsoft Office PowerPoint
PresentationFormat: Widescreen
AppVersion: 15.0000
"""

ALL_METADATA_PDF = """/Producer: pdfTeX-1.40.26
/Author: john
/Subject: Test PDF file
/Title: Lorem Ipsum
/Creator: TeX
/CreationDate: D:20240411151305+03'00'
/ModDate: D:20240411151305+03'00'
/Trapped: /False
/PTEX.Fullbanner: This is pdfTeX, Version 3.141592653-2.6-1.40.26 (TeX Live 2024/Arch Linux) kpathsea version 6.4.0
/Pages: 1
"""

# ================== Test Functions ==================

def test_select_file_type(qtbot_app):
    qtbot, window = qtbot_app
    filetype_combo = window.findChild(QComboBox, "filetype_combo")
    assert filetype_combo is not None
    logger.info("test_select_file_type: filetype_combo found.")

    # Test selecting each file type
    for index, filetype in enumerate(["DOCX", "PPTX", "PDF"]):
        filetype_combo.setCurrentIndex(index)
        qtbot.wait(100)  # Wait for the event to process
        current = filetype_combo.currentText()
        assert current == filetype
        logger.info(f"test_select_file_type: Selected file type '{current}' successfully.")

def test_browse_file(qtbot_app, tmp_path):
    qtbot, window = qtbot_app
    select_file_btn = window.findChild(QPushButton, "select_file_btn")
    file_path_label = window.findChild(QLabel, "file_path_label")

    # Create dummy files
    docx_file = tmp_path / "example.docx"
    pptx_file = tmp_path / "example.pptx"
    pdf_file = tmp_path / "example.pdf"
    docx_file.touch()
    pptx_file.touch()
    pdf_file.touch()
    logger.info("test_browse_file: Dummy files created.")

    # Mock QFileDialog.getOpenFileName to return the docx_file
    with patch('PyQt5.QtWidgets.QFileDialog.getOpenFileName', return_value=(str(docx_file), 'DOCX Files (*.docx)')):
        qtbot.mouseClick(select_file_btn, Qt.LeftButton)
        qtbot.wait(100)  # Wait for the event to process
        assert file_path_label.text() == str(docx_file)
        logger.info(f"test_browse_file: Selected DOCX file '{docx_file}' successfully.")

    # Mock for pptx
    with patch('PyQt5.QtWidgets.QFileDialog.getOpenFileName', return_value=(str(pptx_file), 'PPTX Files (*.pptx)')):
        window.filetype_combo.setCurrentText("PPTX")
        qtbot.mouseClick(select_file_btn, Qt.LeftButton)
        qtbot.wait(100)  # Wait for the event to process
        assert file_path_label.text() == str(pptx_file)
        logger.info(f"test_browse_file: Selected PPTX file '{pptx_file}' successfully.")

    # Mock for pdf
    with patch('PyQt5.QtWidgets.QFileDialog.getOpenFileName', return_value=(str(pdf_file), 'PDF Files (*.pdf)')):
        window.filetype_combo.setCurrentText("PDF")
        qtbot.mouseClick(select_file_btn, Qt.LeftButton)
        qtbot.wait(100)  # Wait for the event to process
        assert file_path_label.text() == str(pdf_file)
        logger.info(f"test_browse_file: Selected PDF file '{pdf_file}' successfully.")

def test_metadata_options(qtbot_app):
    qtbot, window = qtbot_app

    option_all = window.findChild(QCheckBox, "option_all")
    option_textprops = window.findChild(QCheckBox, "option_textprops")
    option_mainprops = window.findChild(QCheckBox, "option_mainprops")
    option_appdata = window.findChild(QCheckBox, "option_appdata")
    option_simple = window.findChild(QCheckBox, "option_simple")

    assert option_all is not None
    assert option_textprops is not None
    assert option_mainprops is not None
    assert option_appdata is not None
    assert option_simple is not None
    logger.info("test_metadata_options: All metadata checkboxes found.")

    # Ensure 'option_all' is enabled and visible
    assert option_all.isEnabled()
    assert option_all.isVisible()
    logger.info("test_metadata_options: 'All Metadata' checkbox is enabled and visible.")

    # Select 'All Metadata' and verify others are unchecked and disabled
    option_all.setChecked(True)
    qtbot.wait(100)  # Wait for the event to process

    assert option_all.isChecked(), "Failed to check 'All Metadata' checkbox."
    assert not option_textprops.isChecked()
    assert not option_mainprops.isChecked()
    assert not option_appdata.isChecked()
    assert not option_simple.isChecked()
    assert not option_textprops.isEnabled()
    assert not option_mainprops.isEnabled()
    assert not option_appdata.isEnabled()
    assert not option_simple.isEnabled()
    logger.info("test_metadata_options: 'All Metadata' selected; other options unchecked and disabled.")

    # Uncheck 'All Metadata' and verify others are enabled
    option_all.setChecked(False)
    qtbot.wait(100)  # Wait for the event to process

    assert not option_all.isChecked(), "'All Metadata' checkbox should be unchecked."
    assert option_textprops.isEnabled()
    assert option_mainprops.isEnabled()
    assert option_appdata.isEnabled()
    assert option_simple.isEnabled()
    logger.info("test_metadata_options: 'All Metadata' unchecked; other options enabled.")

    # Select individual options
    option_textprops.setChecked(True)
    option_mainprops.setChecked(True)
    qtbot.wait(100)  # Wait for the events to process

    assert option_textprops.isChecked(), "'Text Properties' checkbox should be checked."
    assert option_mainprops.isChecked(), "'Main Properties' checkbox should be checked."
    logger.info("test_metadata_options: Individual metadata options selected.")

def test_extract_metadata_docx(qtbot_app, mock_subprocess_run, tmp_path, capsys):
    qtbot, window = qtbot_app
    select_file_btn = window.findChild(QPushButton, "select_file_btn")
    file_path_label = window.findChild(QLabel, "file_path_label")
    extract_btn = window.findChild(QPushButton, "extract_btn")
    results_text = window.findChild(QTextEdit, "results_text")
    option_all = window.findChild(QCheckBox, "option_all")

    # Create a dummy docx file
    docx_file = tmp_path / "example.docx"
    docx_file.touch()
    window.filetype_combo.setCurrentText("DOCX")
    logger.info(f"test_extract_metadata_docx: Created dummy DOCX file '{docx_file}'.")

    # Mock QFileDialog to select the docx file
    with patch('PyQt5.QtWidgets.QFileDialog.getOpenFileName', return_value=(str(docx_file), 'DOCX Files (*.docx)')):
        qtbot.mouseClick(select_file_btn, Qt.LeftButton)
        qtbot.wait(100)  # Wait for the event to process
        assert file_path_label.text() == str(docx_file)
        logger.info(f"test_extract_metadata_docx: Selected DOCX file '{docx_file}'.")

    # Select 'All Metadata'
    option_all.setChecked(True)
    qtbot.wait(100)  # Wait for the event to process
    assert option_all.isChecked(), "Failed to check 'All Metadata' checkbox."
    logger.info("test_extract_metadata_docx: 'All Metadata' option selected.")

    # Mock subprocess.run to return ALL_METADATA_DOCX
    mock_subprocess_run.return_value = MagicMock(stdout=ALL_METADATA_DOCX, stderr="", returncode=0)
    logger.info("test_extract_metadata_docx: Mocked subprocess.run for DOCX.")

    # Click 'Extract Metadata'
    qtbot.mouseClick(extract_btn, Qt.LeftButton)
    qtbot.wait(500)  # Wait for the extraction to process
    logger.info("test_extract_metadata_docx: 'Extract Metadata' button clicked.")

    # Verify that subprocess.run was called with correct arguments
    extractor_script = window.get_extractor_script("docx")
    expected_command = [
        sys.executable, extractor_script,
        "--all", str(docx_file)
    ]
    mock_subprocess_run.assert_called_with(
        expected_command,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        encoding='utf-8',
        check=True
    )
    logger.info("test_extract_metadata_docx: subprocess.run called with expected arguments.")

    # Verify that results_text contains the expected metadata
    extracted_metadata = results_text.toPlainText()
    assert extracted_metadata == ALL_METADATA_DOCX
    logger.info("test_extract_metadata_docx: Extracted metadata matches expected DOCX metadata.")

    # Log the extracted metadata
    logger.info("Extracted DOCX Metadata:")
    logger.info(extracted_metadata)

def test_extract_metadata_pptx(qtbot_app, mock_subprocess_run, tmp_path, capsys):
    qtbot, window = qtbot_app
    select_file_btn = window.findChild(QPushButton, "select_file_btn")
    file_path_label = window.findChild(QLabel, "file_path_label")
    extract_btn = window.findChild(QPushButton, "extract_btn")
    results_text = window.findChild(QTextEdit, "results_text")
    option_all = window.findChild(QCheckBox, "option_all")

    # Create a dummy pptx file
    pptx_file = tmp_path / "example.pptx"
    pptx_file.touch()
    window.filetype_combo.setCurrentText("PPTX")
    logger.info(f"test_extract_metadata_pptx: Created dummy PPTX file '{pptx_file}'.")

    # Mock QFileDialog to select the pptx file
    with patch('PyQt5.QtWidgets.QFileDialog.getOpenFileName', return_value=(str(pptx_file), 'PPTX Files (*.pptx)')):
        qtbot.mouseClick(select_file_btn, Qt.LeftButton)
        qtbot.wait(100)  # Wait for the event to process
        assert file_path_label.text() == str(pptx_file)
        logger.info(f"test_extract_metadata_pptx: Selected PPTX file '{pptx_file}'.")

    # Select 'All Metadata'
    option_all.setChecked(True)
    qtbot.wait(100)  # Wait for the event to process
    assert option_all.isChecked(), "Failed to check 'All Metadata' checkbox."
    logger.info("test_extract_metadata_pptx: 'All Metadata' option selected.")

    # Mock subprocess.run to return ALL_METADATA_PPTX
    mock_subprocess_run.return_value = MagicMock(stdout=ALL_METADATA_PPTX, stderr="", returncode=0)
    logger.info("test_extract_metadata_pptx: Mocked subprocess.run for PPTX.")

    # Click 'Extract Metadata'
    qtbot.mouseClick(extract_btn, Qt.LeftButton)
    qtbot.wait(500)  # Wait for the extraction to process
    logger.info("test_extract_metadata_pptx: 'Extract Metadata' button clicked.")

    # Verify that subprocess.run was called with correct arguments
    extractor_script = window.get_extractor_script("pptx")
    expected_command = [
        sys.executable, extractor_script,
        "--all", str(pptx_file)
    ]
    mock_subprocess_run.assert_called_with(
        expected_command,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        encoding='utf-8',
        check=True
    )
    logger.info("test_extract_metadata_pptx: subprocess.run called with expected arguments.")

    # Verify that results_text contains the expected metadata
    extracted_metadata = results_text.toPlainText()
    assert extracted_metadata == ALL_METADATA_PPTX
    logger.info("test_extract_metadata_pptx: Extracted metadata matches expected PPTX metadata.")

    # Log the extracted metadata
    logger.info("Extracted PPTX Metadata:")
    logger.info(extracted_metadata)

def test_extract_metadata_pdf(qtbot_app, mock_subprocess_run, tmp_path, capsys):
    qtbot, window = qtbot_app
    select_file_btn = window.findChild(QPushButton, "select_file_btn")
    file_path_label = window.findChild(QLabel, "file_path_label")
    extract_btn = window.findChild(QPushButton, "extract_btn")
    results_text = window.findChild(QTextEdit, "results_text")
    option_all = window.findChild(QCheckBox, "option_all")

    # Create a dummy pdf file
    pdf_file = tmp_path / "example.pdf"
    pdf_file.touch()
    window.filetype_combo.setCurrentText("PDF")
    logger.info(f"test_extract_metadata_pdf: Created dummy PDF file '{pdf_file}'.")

    # Mock QFileDialog to select the pdf file
    with patch('PyQt5.QtWidgets.QFileDialog.getOpenFileName', return_value=(str(pdf_file), 'PDF Files (*.pdf)')):
        qtbot.mouseClick(select_file_btn, Qt.LeftButton)
        qtbot.wait(100)  # Wait for the event to process
        assert file_path_label.text() == str(pdf_file)
        logger.info(f"test_extract_metadata_pdf: Selected PDF file '{pdf_file}'.")

    # Select 'All Metadata'
    option_all.setChecked(True)
    qtbot.wait(100)  # Wait for the event to process
    assert option_all.isChecked(), "Failed to check 'All Metadata' checkbox."
    logger.info("test_extract_metadata_pdf: 'All Metadata' option selected.")

    # Mock subprocess.run to return ALL_METADATA_PDF
    mock_subprocess_run.return_value = MagicMock(stdout=ALL_METADATA_PDF, stderr="", returncode=0)
    logger.info("test_extract_metadata_pdf: Mocked subprocess.run for PDF.")

    # Click 'Extract Metadata'
    qtbot.mouseClick(extract_btn, Qt.LeftButton)
    qtbot.wait(500)  # Wait for the extraction to process
    logger.info("test_extract_metadata_pdf: 'Extract Metadata' button clicked.")

    # Verify that subprocess.run was called with correct arguments
    extractor_script = window.get_extractor_script("pdf")
    expected_command = [
        sys.executable, extractor_script,
        "--all", str(pdf_file)
    ]
    mock_subprocess_run.assert_called_with(
        expected_command,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        encoding='utf-8',
        check=True
    )
    logger.info("test_extract_metadata_pdf: subprocess.run called with expected arguments.")

    # Verify that results_text contains the expected metadata
    extracted_metadata = results_text.toPlainText()
    assert extracted_metadata == ALL_METADATA_PDF
    logger.info("test_extract_metadata_pdf: Extracted metadata matches expected PDF metadata.")

    # Log the extracted metadata
    logger.info("Extracted PDF Metadata:")
    logger.info(extracted_metadata)

def test_copy_to_clipboard(qtbot_app, tmp_path):
    qtbot, window = qtbot_app
    copy_btn = window.findChild(QPushButton, "copy_btn")
    results_text = window.findChild(QTextEdit, "results_text")

    # Set some text in results_text
    results_text.setPlainText(ALL_METADATA_DOCX)
    qtbot.wait(100)  # Wait for the state to update
    logger.info("test_copy_to_clipboard: Set metadata text for copying.")

    # Mock clipboard
    clipboard = QApplication.clipboard()
    clipboard.clear()

    # Mock QMessageBox.information
    with patch('PyQt5.QtWidgets.QMessageBox.information') as mock_info:
        # Click 'Copy to Clipboard'
        qtbot.mouseClick(copy_btn, Qt.LeftButton)
        qtbot.wait(100)  # Wait for the event to process

        # Verify clipboard contains the expected text
        assert clipboard.text() == ALL_METADATA_DOCX
        logger.info("test_copy_to_clipboard: Clipboard contains the expected metadata.")

        # Verify that the information message was shown
        mock_info.assert_called_once_with(
            window,
            "Copied",
            "Metadata copied to clipboard."
        )
        logger.info("test_copy_to_clipboard: Information message displayed.")

def test_extract_no_file_selected(qtbot_app, mock_subprocess_run):
    qtbot, window = qtbot_app
    extract_btn = window.findChild(QPushButton, "extract_btn")

    # Ensure no file is selected
    file_path_label = window.findChild(QLabel, "file_path_label")
    file_path_label.setText("")
    logger.info("test_extract_no_file_selected: No file selected.")

    # Mock QMessageBox.warning
    with patch('PyQt5.QtWidgets.QMessageBox.warning') as mock_warning:
        qtbot.mouseClick(extract_btn, Qt.LeftButton)
        qtbot.wait(100)  # Wait for the event to process
        mock_warning.assert_called_once_with(
            window,
            "Invalid File",
            "Please select a valid file."
        )
        logger.info("test_extract_no_file_selected: Warning message displayed for no file selected.")

def test_extract_no_options_selected(qtbot_app, mock_subprocess_run, tmp_path):
    qtbot, window = qtbot_app
    extract_btn = window.findChild(QPushButton, "extract_btn")
    select_file_btn = window.findChild(QPushButton, "select_file_btn")
    file_path_label = window.findChild(QLabel, "file_path_label")

    # Create a dummy docx file
    docx_file = tmp_path / "example.docx"
    docx_file.touch()
    window.filetype_combo.setCurrentText("DOCX")
    logger.info(f"test_extract_no_options_selected: Created dummy DOCX file '{docx_file}'.")

    # Mock QFileDialog to select the docx file
    with patch('PyQt5.QtWidgets.QFileDialog.getOpenFileName', return_value=(str(docx_file), 'DOCX Files (*.docx)')):
        qtbot.mouseClick(select_file_btn, Qt.LeftButton)
        qtbot.wait(100)  # Wait for the event to process
        assert file_path_label.text() == str(docx_file)
        logger.info(f"test_extract_no_options_selected: Selected DOCX file '{docx_file}'.")

    # Ensure no options are selected
    option_all = window.findChild(QCheckBox, "option_all")
    option_textprops = window.findChild(QCheckBox, "option_textprops")
    option_mainprops = window.findChild(QCheckBox, "option_mainprops")
    option_appdata = window.findChild(QCheckBox, "option_appdata")
    option_simple = window.findChild(QCheckBox, "option_simple")

    option_all.setChecked(False)
    option_textprops.setChecked(False)
    option_mainprops.setChecked(False)
    option_appdata.setChecked(False)
    option_simple.setChecked(False)
    qtbot.wait(100)  # Wait for the state to update
    logger.info("test_extract_no_options_selected: No metadata options selected.")

    # Mock QMessageBox.warning
    with patch('PyQt5.QtWidgets.QMessageBox.warning') as mock_warning:
        qtbot.mouseClick(extract_btn, Qt.LeftButton)
        qtbot.wait(100)  # Wait for the event to process
        mock_warning.assert_called_once_with(
            window,
            "No Options Selected",
            "Please select at least one metadata option."
        )
        logger.info("test_extract_no_options_selected: Warning message displayed for no options selected.")

def test_copy_no_data(qtbot_app):
    qtbot, window = qtbot_app
    copy_btn = window.findChild(QPushButton, "copy_btn")
    results_text = window.findChild(QTextEdit, "results_text")

    # Ensure results_text is empty
    results_text.setPlainText("")
    qtbot.wait(100)  # Wait for the state to update
    logger.info("test_copy_no_data: Results text is empty.")

    # Mock QMessageBox.warning
    with patch('PyQt5.QtWidgets.QMessageBox.warning') as mock_warning:
        qtbot.mouseClick(copy_btn, Qt.LeftButton)
        qtbot.wait(100)  # Wait for the event to process
        mock_warning.assert_called_once_with(
            window,
            "No Data",
            "There is no metadata to copy."
        )
        logger.info("test_copy_no_data: Warning message displayed for no data to copy.")
