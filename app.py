# # # # app.py
# # #
# # # import sys
# # # import os
# # # import subprocess
# # # from PyQt5.QtWidgets import (
# # #     QApplication, QWidget, QLabel, QPushButton, QFileDialog,
# # #     QVBoxLayout, QHBoxLayout, QComboBox, QMessageBox, QTextEdit,
# # #     QGroupBox, QCheckBox, QScrollArea
# # # )
# # # from PyQt5.QtCore import Qt
# # #
# # # class MetadataExtractorApp(QWidget):
# # #     def __init__(self):
# # #         super().__init__()
# # #         self.setWindowTitle("Metadata Extractor")
# # #         self.setGeometry(100, 100, 800, 600)
# # #         self.init_ui()
# # #
# # #     def init_ui(self):
# # #         # Main layout
# # #         main_layout = QVBoxLayout()
# # #
# # #         # File selection group
# # #         file_group = QGroupBox("Select File")
# # #         file_layout = QHBoxLayout()
# # #
# # #         self.file_path_label = QLabel("No file selected")
# # #         self.file_path_label.setWordWrap(True)
# # #         select_file_btn = QPushButton("Browse")
# # #         select_file_btn.clicked.connect(self.browse_file)
# # #
# # #         file_layout.addWidget(self.file_path_label)
# # #         file_layout.addWidget(select_file_btn)
# # #         file_group.setLayout(file_layout)
# # #
# # #         # File type selection
# # #         filetype_layout = QHBoxLayout()
# # #         filetype_label = QLabel("Select File Type:")
# # #         self.filetype_combo = QComboBox()
# # #         self.filetype_combo.addItems(["DOCX", "PPTX", "PDF"])
# # #         filetype_layout.addWidget(filetype_label)
# # #         filetype_layout.addWidget(self.filetype_combo)
# # #         filetype_layout.addStretch()
# # #
# # #         # Metadata options
# # #         options_group = QGroupBox("Metadata Options")
# # #         options_layout = QVBoxLayout()
# # #
# # #         self.option_all = QCheckBox("All Metadata")
# # #         self.option_textprops = QCheckBox("Text Properties")
# # #         self.option_mainprops = QCheckBox("Main Properties")
# # #         self.option_appdata = QCheckBox("Application Data")
# # #         self.option_simple = QCheckBox("Simple Properties")
# # #
# # #         options_layout.addWidget(self.option_all)
# # #         options_layout.addWidget(self.option_textprops)
# # #         options_layout.addWidget(self.option_mainprops)
# # #         options_layout.addWidget(self.option_appdata)
# # #         options_layout.addWidget(self.option_simple)
# # #
# # #         options_group.setLayout(options_layout)
# # #
# # #         # Extract button
# # #         extract_btn = QPushButton("Extract Metadata")
# # #         extract_btn.clicked.connect(self.extract_metadata)
# # #
# # #         # Results display
# # #         results_group = QGroupBox("Extracted Metadata")
# # #         results_layout = QVBoxLayout()
# # #         self.results_text = QTextEdit()
# # #         self.results_text.setReadOnly(True)
# # #         results_layout.addWidget(self.results_text)
# # #         results_group.setLayout(results_layout)
# # #
# # #         # Add widgets to main layout
# # #         main_layout.addWidget(file_group)
# # #         main_layout.addLayout(filetype_layout)
# # #         main_layout.addWidget(options_group)
# # #         main_layout.addWidget(extract_btn)
# # #         main_layout.addWidget(results_group)
# # #
# # #         self.setLayout(main_layout)
# # #
# # #     def browse_file(self):
# # #         filetype = self.filetype_combo.currentText().lower()
# # #         if filetype == "pdf":
# # #             filter_str = "PDF Files (*.pdf)"
# # #         elif filetype == "docx":
# # #             filter_str = "DOCX Files (*.docx)"
# # #         elif filetype == "pptx":
# # #             filter_str = "PPTX Files (*.pptx)"
# # #         else:
# # #             filter_str = "All Files (*)"
# # #
# # #         file_path, _ = QFileDialog.getOpenFileName(
# # #             self, "Select File", "", filter_str
# # #         )
# # #         if file_path:
# # #             self.file_path_label.setText(file_path)
# # #
# # #     def extract_metadata(self):
# # #         file_path = self.file_path_label.text()
# # #         if not file_path or not os.path.isfile(file_path):
# # #             QMessageBox.warning(self, "Invalid File", "Please select a valid file.")
# # #             return
# # #
# # #         filetype = self.filetype_combo.currentText().lower()
# # #         extractor_script = self.get_extractor_script(filetype)
# # #         if not extractor_script:
# # #             QMessageBox.critical(
# # #                 self, "Error", f"No extractor script found for {filetype.upper()} files."
# # #             )
# # #             return
# # #
# # #         # Determine selected options
# # #         options = []
# # #         if self.option_all.isChecked():
# # #             options.append("--all")
# # #         else:
# # #             if self.option_textprops.isChecked():
# # #                 options.append("--textprops")
# # #             if self.option_mainprops.isChecked():
# # #                 options.append("--mainprops")
# # #             if self.option_appdata.isChecked():
# # #                 options.append("--appdata")
# # #             if self.option_simple.isChecked():
# # #                 options.append("--simple")
# # #
# # #         # Prepare command
# # #         command = [
# # #             sys.executable, extractor_script
# # #         ] + options + [file_path]
# # #
# # #         try:
# # #             # Execute the extractor script with UTF-8 encoding
# # #             result = subprocess.run(
# # #                 command,
# # #                 stdout=subprocess.PIPE,
# # #                 stderr=subprocess.PIPE,
# # #                 text=True,
# # #                 encoding='utf-8',  # Specify UTF-8 encoding
# # #                 check=True
# # #             )
# # #             metadata = result.stdout
# # #             self.results_text.setPlainText(metadata)
# # #         except subprocess.CalledProcessError as e:
# # #             error_message = e.stderr if e.stderr else "An unknown error occurred."
# # #             QMessageBox.critical(self, "Extraction Error", error_message)
# # #         except UnicodeDecodeError as e:
# # #             QMessageBox.critical(self, "Decoding Error", f"Failed to decode output: {str(e)}")
# # #         except Exception as e:
# # #             QMessageBox.critical(self, "Error", str(e))
# # #
# # #     def get_extractor_script(self, filetype):
# # #         script_map = {
# # #             "docx": "_docx_meta.py",
# # #             "pptx": "_pptx_meta.py",
# # #             "pdf": "_pdf_meta.py"
# # #         }
# # #         script = script_map.get(filetype)
# # #         if script and os.path.isfile(script):
# # #             return os.path.abspath(script)
# # #         return None
# # #
# # # def main():
# # #     app = QApplication(sys.argv)
# # #     window = MetadataExtractorApp()
# # #     window.show()
# # #     sys.exit(app.exec_())
# # #
# # # if __name__ == "__main__":
# # #     main()
# #
# #
# #
# # # app.py
# #
# # import sys
# # import os
# # import subprocess
# # from PyQt5.QtWidgets import (
# #     QApplication, QWidget, QLabel, QPushButton, QFileDialog,
# #     QVBoxLayout, QHBoxLayout, QComboBox, QMessageBox, QTextEdit,
# #     QGroupBox, QCheckBox, QScrollArea, QSpacerItem, QSizePolicy
# # )
# # from PyQt5.QtCore import Qt
# # from PyQt5.QtGui import QIcon, QFont
# #
# # class MetadataExtractorApp(QWidget):
# #     def __init__(self):
# #         super().__init__()
# #         self.setWindowTitle("Metadata Extractor")
# #         self.setWindowIcon(QIcon("icon.png"))  # Optional: Add an icon.png in your project directory
# #         self.setGeometry(100, 100, 900, 700)
# #         self.init_ui()
# #
# #     def init_ui(self):
# #         # Apply a simple stylesheet for better visuals
# #         self.setStyleSheet("""
# #             QWidget {
# #                 font-family: Arial, sans-serif;
# #                 font-size: 14px;
# #             }
# #             QGroupBox {
# #                 font-weight: bold;
# #                 margin-top: 10px;
# #             }
# #             QPushButton {
# #                 padding: 5px 10px;
# #             }
# #             QTextEdit {
# #                 background-color: #f0f0f0;
# #             }
# #         """)
# #
# #         # Main layout
# #         main_layout = QVBoxLayout()
# #
# #         # File type selection
# #         filetype_layout = QHBoxLayout()
# #         filetype_label = QLabel("Select File Type:")
# #         self.filetype_combo = QComboBox()
# #         self.filetype_combo.addItems(["DOCX", "PPTX", "PDF"])
# #         self.filetype_combo.currentTextChanged.connect(self.update_file_filter)
# #         filetype_layout.addWidget(filetype_label)
# #         filetype_layout.addWidget(self.filetype_combo)
# #         filetype_layout.addStretch()
# #
# #         # File selection group
# #         file_group = QGroupBox("Select File")
# #         file_layout = QHBoxLayout()
# #
# #         self.file_path_label = QLabel("No file selected")
# #         self.file_path_label.setWordWrap(True)
# #         select_file_btn = QPushButton("Browse")
# #         select_file_btn.clicked.connect(self.browse_file)
# #
# #         file_layout.addWidget(self.file_path_label)
# #         file_layout.addWidget(select_file_btn)
# #         file_group.setLayout(file_layout)
# #
# #         # Metadata options
# #         options_group = QGroupBox("Metadata Options")
# #         options_layout = QVBoxLayout()
# #
# #         self.option_all = QCheckBox("All Metadata")
# #         self.option_all.stateChanged.connect(self.toggle_options)
# #
# #         self.option_textprops = QCheckBox("Text Properties")
# #         self.option_mainprops = QCheckBox("Main Properties")
# #         self.option_appdata = QCheckBox("Application Data")
# #         self.option_simple = QCheckBox("Simple Properties")
# #
# #         options_layout.addWidget(self.option_all)
# #         options_layout.addWidget(self.option_textprops)
# #         options_layout.addWidget(self.option_mainprops)
# #         options_layout.addWidget(self.option_appdata)
# #         options_layout.addWidget(self.option_simple)
# #
# #         options_group.setLayout(options_layout)
# #
# #         # Extract button
# #         extract_btn = QPushButton("Extract Metadata")
# #         extract_btn.setStyleSheet("background-color: #4CAF50; color: white;")
# #         extract_btn.clicked.connect(self.extract_metadata)
# #
# #         # Results display
# #         results_group = QGroupBox("Extracted Metadata")
# #         results_layout = QVBoxLayout()
# #
# #         # Horizontal layout for QTextEdit and Copy button
# #         results_display_layout = QHBoxLayout()
# #         self.results_text = QTextEdit()
# #         self.results_text.setReadOnly(True)
# #         self.results_text.setFont(QFont("Courier", 10))
# #         results_display_layout.addWidget(self.results_text)
# #
# #         copy_btn = QPushButton("Copy to Clipboard")
# #         copy_btn.setToolTip("Copy all extracted metadata to clipboard")
# #         copy_btn.clicked.connect(self.copy_to_clipboard)
# #         results_display_layout.addWidget(copy_btn)
# #
# #         results_layout.addLayout(results_display_layout)
# #         results_group.setLayout(results_layout)
# #
# #         # Add widgets to main layout
# #         main_layout.addLayout(filetype_layout)
# #         main_layout.addWidget(file_group)
# #         main_layout.addWidget(options_group)
# #         main_layout.addWidget(extract_btn)
# #         main_layout.addWidget(results_group)
# #
# #         self.setLayout(main_layout)
# #
# #     def update_file_filter(self):
# #         """Update the file dialog filter based on selected file type."""
# #         # This method can be expanded if additional logic is needed when file type changes
# #         pass
# #
# #     def toggle_options(self, state):
# #         """Disable other options when 'All Metadata' is selected."""
# #         if state == Qt.Checked:
# #             self.option_textprops.setChecked(False)
# #             self.option_mainprops.setChecked(False)
# #             self.option_appdata.setChecked(False)
# #             self.option_simple.setChecked(False)
# #
# #             self.option_textprops.setEnabled(False)
# #             self.option_mainprops.setEnabled(False)
# #             self.option_appdata.setEnabled(False)
# #             self.option_simple.setEnabled(False)
# #         else:
# #             self.option_textprops.setEnabled(True)
# #             self.option_mainprops.setEnabled(True)
# #             self.option_appdata.setEnabled(True)
# #             self.option_simple.setEnabled(True)
# #
# #     def browse_file(self):
# #         filetype = self.filetype_combo.currentText().lower()
# #         if filetype == "pdf":
# #             filter_str = "PDF Files (*.pdf)"
# #         elif filetype == "docx":
# #             filter_str = "DOCX Files (*.docx)"
# #         elif filetype == "pptx":
# #             filter_str = "PPTX Files (*.pptx)"
# #         else:
# #             filter_str = "All Files (*)"
# #
# #         file_path, _ = QFileDialog.getOpenFileName(
# #             self, "Select File", "", filter_str
# #         )
# #         if file_path:
# #             self.file_path_label.setText(file_path)
# #             self.results_text.clear()  # Clear previous results when a new file is selected
# #
# #     def extract_metadata(self):
# #         file_path = self.file_path_label.text()
# #         if not file_path or not os.path.isfile(file_path):
# #             QMessageBox.warning(self, "Invalid File", "Please select a valid file.")
# #             return
# #
# #         filetype = self.filetype_combo.currentText().lower()
# #         extractor_script = self.get_extractor_script(filetype)
# #         if not extractor_script:
# #             QMessageBox.critical(
# #                 self, "Error", f"No extractor script found for {filetype.upper()} files."
# #             )
# #             return
# #
# #         # Determine selected options
# #         options = []
# #         if self.option_all.isChecked():
# #             options.append("--all")
# #         else:
# #             if self.option_textprops.isChecked():
# #                 options.append("--textprops")
# #             if self.option_mainprops.isChecked():
# #                 options.append("--mainprops")
# #             if self.option_appdata.isChecked():
# #                 options.append("--appdata")
# #             if self.option_simple.isChecked():
# #                 options.append("--simple")
# #
# #         if not options:
# #             QMessageBox.warning(self, "No Options Selected", "Please select at least one metadata option.")
# #             return
# #
# #         # Prepare command
# #         command = [
# #             sys.executable, extractor_script
# #         ] + options + [file_path]
# #
# #         try:
# #             # Execute the extractor script with UTF-8 encoding
# #             result = subprocess.run(
# #                 command,
# #                 stdout=subprocess.PIPE,
# #                 stderr=subprocess.PIPE,
# #                 text=True,
# #                 encoding='utf-8',  # Specify UTF-8 encoding
# #                 check=True
# #             )
# #             metadata = result.stdout
# #             self.results_text.setPlainText(metadata)
# #         except subprocess.CalledProcessError as e:
# #             error_message = e.stderr if e.stderr else "An unknown error occurred."
# #             QMessageBox.critical(self, "Extraction Error", error_message)
# #         except UnicodeDecodeError as e:
# #             QMessageBox.critical(self, "Decoding Error", f"Failed to decode output: {str(e)}")
# #         except Exception as e:
# #             QMessageBox.critical(self, "Error", str(e))
# #
# #     def copy_to_clipboard(self):
# #         metadata = self.results_text.toPlainText()
# #         if metadata:
# #             clipboard = QApplication.clipboard()
# #             clipboard.setText(metadata)
# #             QMessageBox.information(self, "Copied", "Metadata copied to clipboard.")
# #         else:
# #             QMessageBox.warning(self, "No Data", "There is no metadata to copy.")
# #
# #     def get_extractor_script(self, filetype):
# #         script_map = {
# #             "docx": "_docx_meta.py",
# #             "pptx": "_pptx_meta.py",
# #             "pdf": "_pdf_meta.py"
# #         }
# #         script = script_map.get(filetype)
# #         if script and os.path.isfile(script):
# #             return os.path.abspath(script)
# #         return None
# #
# # def main():
# #     app = QApplication(sys.argv)
# #     window = MetadataExtractorApp()
# #     window.show()
# #     sys.exit(app.exec_())
# #
# # if __name__ == "__main__":
# #     main()
#
#
# # app.py
# #
# # import sys
# # import os
# # import subprocess
# # from PyQt5.QtWidgets import (
# #     QApplication, QWidget, QLabel, QPushButton, QFileDialog,
# #     QVBoxLayout, QHBoxLayout, QComboBox, QMessageBox, QTextEdit,
# #     QGroupBox, QCheckBox, QSpacerItem, QSizePolicy
# # )
# # from PyQt5.QtCore import Qt
# # from PyQt5.QtGui import QIcon, QFont
# #
# # class MetadataExtractorApp(QWidget):
# #     def __init__(self):
# #         super().__init__()
# #         self.setWindowTitle("Metadata Extractor")
# #         self.setWindowIcon(QIcon("icon.png"))  # Optional: Add an icon.png in your project directory
# #         self.setGeometry(100, 100, 900, 700)
# #         self.init_ui()
# #
# #     def init_ui(self):
# #         # Apply a simple stylesheet for better visuals
# #         self.setStyleSheet("""
# #             QWidget {
# #                 font-family: Arial, sans-serif;
# #                 font-size: 14px;
# #             }
# #             QGroupBox {
# #                 font-weight: bold;
# #                 margin-top: 10px;
# #             }
# #             QPushButton {
# #                 padding: 8px 16px;
# #                 font-size: 14px;
# #             }
# #             QTextEdit {
# #                 background-color: #f0f0f0;
# #                 font-family: Courier, monospace;
# #                 font-size: 12px;
# #             }
# #             QCheckBox {
# #                 margin-left: 10px;
# #                 font-size: 14px;
# #             }
# #             QLabel {
# #                 font-size: 14px;
# #             }
# #             QComboBox {
# #                 font-size: 14px;
# #             }
# #         """)
# #
# #         # Main layout
# #         main_layout = QVBoxLayout()
# #         main_layout.setContentsMargins(20, 20, 20, 20)
# #         main_layout.setSpacing(15)
# #
# #         # File type selection
# #         filetype_layout = QHBoxLayout()
# #         filetype_label = QLabel("Select File Type:")
# #         self.filetype_combo = QComboBox()
# #         self.filetype_combo.addItems(["DOCX", "PPTX", "PDF"])
# #         self.filetype_combo.currentTextChanged.connect(self.update_file_filter)
# #         filetype_layout.addWidget(filetype_label)
# #         filetype_layout.addWidget(self.filetype_combo)
# #         filetype_layout.addStretch()
# #
# #         # File selection group
# #         file_group = QGroupBox("Select File")
# #         file_layout = QHBoxLayout()
# #
# #         self.file_path_label = QLabel("No file selected")
# #         self.file_path_label.setWordWrap(True)
# #         select_file_btn = QPushButton("Browse")
# #         select_file_btn.clicked.connect(self.browse_file)
# #
# #         file_layout.addWidget(self.file_path_label)
# #         file_layout.addWidget(select_file_btn)
# #         file_group.setLayout(file_layout)
# #
# #         # Metadata options
# #         options_group = QGroupBox("Metadata Options")
# #         options_layout = QVBoxLayout()
# #         options_layout.setSpacing(10)  # Increase spacing between checkboxes
# #
# #         self.option_all = QCheckBox("All Metadata")
# #         self.option_all.stateChanged.connect(self.toggle_options)
# #
# #         self.option_textprops = QCheckBox("Text Properties")
# #         self.option_mainprops = QCheckBox("Main Properties")
# #         self.option_appdata = QCheckBox("Application Data")
# #         self.option_simple = QCheckBox("Simple Properties")
# #
# #         options_layout.addWidget(self.option_all)
# #         options_layout.addWidget(self.option_textprops)
# #         options_layout.addWidget(self.option_mainprops)
# #         options_layout.addWidget(self.option_appdata)
# #         options_layout.addWidget(self.option_simple)
# #
# #         options_group.setLayout(options_layout)
# #
# #         # Extract button
# #         extract_btn = QPushButton("Extract Metadata")
# #         extract_btn.setStyleSheet("background-color: #4CAF50; color: white;")
# #         extract_btn.setMinimumHeight(40)  # Increase button height for better visibility
# #         extract_btn.clicked.connect(self.extract_metadata)
# #
# #         # Results display
# #         results_group = QGroupBox("Extracted Metadata")
# #         results_layout = QVBoxLayout()
# #
# #         # Horizontal layout for QTextEdit and Copy button
# #         results_display_layout = QHBoxLayout()
# #         self.results_text = QTextEdit()
# #         self.results_text.setReadOnly(True)
# #         self.results_text.setFont(QFont("Courier", 10))
# #         results_display_layout.addWidget(self.results_text)
# #
# #         copy_btn = QPushButton("Copy to Clipboard")
# #         copy_btn.setToolTip("Copy all extracted metadata to clipboard")
# #         copy_btn.setStyleSheet("background-color: #2196F3; color: white;")
# #         copy_btn.setMinimumHeight(40)  # Increase button height for better visibility
# #         copy_btn.clicked.connect(self.copy_to_clipboard)
# #         results_display_layout.addWidget(copy_btn)
# #
# #         results_layout.addLayout(results_display_layout)
# #         results_group.setLayout(results_layout)
# #
# #         # Add widgets to main layout
# #         main_layout.addLayout(filetype_layout)
# #         main_layout.addWidget(file_group)
# #         main_layout.addWidget(options_group)
# #         main_layout.addWidget(extract_btn)
# #         main_layout.addWidget(results_group)
# #
# #         # Add stretch to push everything to the top
# #         main_layout.addStretch()
# #
# #         self.setLayout(main_layout)
# #
# #     def update_file_filter(self):
# #         """Update the file dialog filter based on selected file type."""
# #         # Currently, the filter is updated dynamically in browse_file.
# #         # This method is kept for potential future enhancements.
# #         pass
# #
# #     def toggle_options(self, state):
# #         """Disable other options when 'All Metadata' is selected."""
# #         if state == Qt.Checked:
# #             self.option_textprops.setChecked(False)
# #             self.option_mainprops.setChecked(False)
# #             self.option_appdata.setChecked(False)
# #             self.option_simple.setChecked(False)
# #
# #             self.option_textprops.setEnabled(False)
# #             self.option_mainprops.setEnabled(False)
# #             self.option_appdata.setEnabled(False)
# #             self.option_simple.setEnabled(False)
# #         else:
# #             self.option_textprops.setEnabled(True)
# #             self.option_mainprops.setEnabled(True)
# #             self.option_appdata.setEnabled(True)
# #             self.option_simple.setEnabled(True)
# #
# #     def browse_file(self):
# #         filetype = self.filetype_combo.currentText().lower()
# #         if filetype == "pdf":
# #             filter_str = "PDF Files (*.pdf)"
# #         elif filetype == "docx":
# #             filter_str = "DOCX Files (*.docx)"
# #         elif filetype == "pptx":
# #             filter_str = "PPTX Files (*.pptx)"
# #         else:
# #             filter_str = "All Files (*)"
# #
# #         file_path, _ = QFileDialog.getOpenFileName(
# #             self, "Select File", "", filter_str
# #         )
# #         if file_path:
# #             self.file_path_label.setText(file_path)
# #             self.results_text.clear()  # Clear previous results when a new file is selected
# #
# #     def extract_metadata(self):
# #         file_path = self.file_path_label.text()
# #         if not file_path or not os.path.isfile(file_path):
# #             QMessageBox.warning(self, "Invalid File", "Please select a valid file.")
# #             return
# #
# #         filetype = self.filetype_combo.currentText().lower()
# #         extractor_script = self.get_extractor_script(filetype)
# #         if not extractor_script:
# #             QMessageBox.critical(
# #                 self, "Error", f"No extractor script found for {filetype.upper()} files."
# #             )
# #             return
# #
# #         # Determine selected options
# #         options = []
# #         if self.option_all.isChecked():
# #             options.append("--all")
# #         else:
# #             if self.option_textprops.isChecked():
# #                 options.append("--textprops")
# #             if self.option_mainprops.isChecked():
# #                 options.append("--mainprops")
# #             if self.option_appdata.isChecked():
# #                 options.append("--appdata")
# #             if self.option_simple.isChecked():
# #                 options.append("--simple")
# #
# #         if not options:
# #             QMessageBox.warning(self, "No Options Selected", "Please select at least one metadata option.")
# #             return
# #
# #         # Prepare command
# #         command = [
# #             sys.executable, extractor_script
# #         ] + options + [file_path]
# #
# #         try:
# #             # Execute the extractor script with UTF-8 encoding
# #             result = subprocess.run(
# #                 command,
# #                 stdout=subprocess.PIPE,
# #                 stderr=subprocess.PIPE,
# #                 text=True,
# #                 encoding='utf-8',  # Specify UTF-8 encoding
# #                 check=True
# #             )
# #             metadata = result.stdout
# #             self.results_text.setPlainText(metadata)
# #         except subprocess.CalledProcessError as e:
# #             error_message = e.stderr if e.stderr else "An unknown error occurred."
# #             QMessageBox.critical(self, "Extraction Error", error_message)
# #         except UnicodeDecodeError as e:
# #             QMessageBox.critical(self, "Decoding Error", f"Failed to decode output: {str(e)}")
# #         except Exception as e:
# #             QMessageBox.critical(self, "Error", str(e))
# #
# #     def copy_to_clipboard(self):
# #         metadata = self.results_text.toPlainText()
# #         if metadata:
# #             clipboard = QApplication.clipboard()
# #             clipboard.setText(metadata)
# #             QMessageBox.information(self, "Copied", "Metadata copied to clipboard.")
# #         else:
# #             QMessageBox.warning(self, "No Data", "There is no metadata to copy.")
# #
# #     def get_extractor_script(self, filetype):
# #         script_map = {
# #             "docx": "_docx_meta.py",
# #             "pptx": "_pptx_meta.py",
# #             "pdf": "_pdf_meta.py"
# #         }
# #         script = script_map.get(filetype)
# #         if script and os.path.isfile(script):
# #             return os.path.abspath(script)
# #         return None
# #
# # def main():
# #     app = QApplication(sys.argv)
# #     window = MetadataExtractorApp()
# #     window.show()
# #     sys.exit(app.exec_())
# #
# # if __name__ == "__main__":
# #     main()
#
#
# #
# # import sys
# # import os
# # import subprocess
# # from PyQt5.QtWidgets import (
# #     QApplication, QWidget, QLabel, QPushButton, QFileDialog,
# #     QVBoxLayout, QHBoxLayout, QComboBox, QMessageBox, QTextEdit,
# #     QGroupBox, QCheckBox, QSpacerItem, QSizePolicy
# # )
# # from PyQt5.QtCore import Qt
# # from PyQt5.QtGui import QIcon, QFont
# #
# # class MetadataExtractorApp(QWidget):
# #     def __init__(self):
# #         super().__init__()
# #         self.setWindowTitle("Metadata Extractor")
# #         self.setWindowIcon(QIcon("icon.png"))  # Optional: Add an icon.png in your project directory
# #         self.setGeometry(100, 100, 900, 700)
# #         self.init_ui()
# #
# #     def init_ui(self):
# #         # Apply a simple stylesheet for better visuals
# #         self.setStyleSheet("""
# #             QWidget {
# #                 font-family: Arial, sans-serif;
# #                 font-size: 14px;
# #             }
# #             QGroupBox {
# #                 font-weight: bold;
# #                 margin-top: 10px;
# #             }
# #             QPushButton {
# #                 padding: 8px 16px;
# #                 font-size: 14px;
# #             }
# #             QTextEdit {
# #                 background-color: #f0f0f0;
# #                 font-family: Courier, monospace;
# #                 font-size: 12px;
# #             }
# #             QCheckBox {
# #                 margin-left: 10px;
# #                 font-size: 14px;
# #             }
# #             QLabel {
# #                 font-size: 14px;
# #             }
# #             QComboBox {
# #                 font-size: 14px;
# #             }
# #         """)
# #
# #         # Main layout
# #         main_layout = QVBoxLayout()
# #         main_layout.setContentsMargins(20, 20, 20, 20)
# #         main_layout.setSpacing(15)
# #
# #         # File type selection
# #         filetype_layout = QHBoxLayout()
# #         filetype_label = QLabel("Select File Type:")
# #         self.filetype_combo = QComboBox()
# #         self.filetype_combo.addItems(["DOCX", "PPTX", "PDF"])
# #         self.filetype_combo.currentTextChanged.connect(self.update_file_filter)
# #         self.filetype_combo.setObjectName("filetypeComboBox")  # Added object name
# #         filetype_layout.addWidget(filetype_label)
# #         filetype_layout.addWidget(self.filetype_combo)
# #         filetype_layout.addStretch()
# #
# #         # File selection group
# #         file_group = QGroupBox("Select File")
# #         file_layout = QHBoxLayout()
# #
# #         self.file_path_label = QLabel("No file selected")
# #         self.file_path_label.setWordWrap(True)
# #         self.file_path_label.setObjectName("filePathLabel")  # Added object name
# #         select_file_btn = QPushButton("Browse")
# #         select_file_btn.clicked.connect(self.browse_file)
# #         select_file_btn.setObjectName("browseButton")  # Added object name
# #
# #         file_layout.addWidget(self.file_path_label)
# #         file_layout.addWidget(select_file_btn)
# #         file_group.setLayout(file_layout)
# #
# #         # Metadata options
# #         options_group = QGroupBox("Metadata Options")
# #         options_layout = QVBoxLayout()
# #         options_layout.setSpacing(10)  # Increase spacing between checkboxes
# #
# #         self.option_all = QCheckBox("All Metadata")
# #         self.option_all.stateChanged.connect(self.toggle_options)
# #         self.option_all.setObjectName("optionAllCheckBox")  # Added object name
# #
# #         self.option_textprops = QCheckBox("Text Properties")
# #         self.option_textprops.setObjectName("optionTextPropsCheckBox")  # Added object name
# #         self.option_mainprops = QCheckBox("Main Properties")
# #         self.option_mainprops.setObjectName("optionMainPropsCheckBox")  # Added object name
# #         self.option_appdata = QCheckBox("Application Data")
# #         self.option_appdata.setObjectName("optionAppDataCheckBox")  # Added object name
# #         self.option_simple = QCheckBox("Simple Properties")
# #         self.option_simple.setObjectName("optionSimpleCheckBox")  # Added object name
# #
# #         options_layout.addWidget(self.option_all)
# #         options_layout.addWidget(self.option_textprops)
# #         options_layout.addWidget(self.option_mainprops)
# #         options_layout.addWidget(self.option_appdata)
# #         options_layout.addWidget(self.option_simple)
# #
# #         options_group.setLayout(options_layout)
# #
# #         # Extract button
# #         extract_btn = QPushButton("Extract Metadata")
# #         extract_btn.setStyleSheet("background-color: #4CAF50; color: white;")
# #         extract_btn.setMinimumHeight(40)  # Increase button height for better visibility
# #         extract_btn.clicked.connect(self.extract_metadata)
# #         extract_btn.setObjectName("extractButton")  # Added object name
# #
# #         # Results display
# #         results_group = QGroupBox("Extracted Metadata")
# #         results_layout = QVBoxLayout()
# #
# #         # Horizontal layout for QTextEdit and Copy button
# #         results_display_layout = QHBoxLayout()
# #         self.results_text = QTextEdit()
# #         self.results_text.setReadOnly(True)
# #         self.results_text.setFont(QFont("Courier", 10))
# #         self.results_text.setObjectName("resultsTextEdit")  # Added object name
# #         results_display_layout.addWidget(self.results_text)
# #
# #         copy_btn = QPushButton("Copy to Clipboard")
# #         copy_btn.setToolTip("Copy all extracted metadata to clipboard")
# #         copy_btn.setStyleSheet("background-color: #2196F3; color: white;")
# #         copy_btn.setMinimumHeight(40)  # Increase button height for better visibility
# #         copy_btn.clicked.connect(self.copy_to_clipboard)
# #         copy_btn.setObjectName("copyButton")  # Added object name
# #         results_display_layout.addWidget(copy_btn)
# #
# #         results_layout.addLayout(results_display_layout)
# #         results_group.setLayout(results_layout)
# #
# #         # Add widgets to main layout
# #         main_layout.addLayout(filetype_layout)
# #         main_layout.addWidget(file_group)
# #         main_layout.addWidget(options_group)
# #         main_layout.addWidget(extract_btn)
# #         main_layout.addWidget(results_group)
# #
# #         # Add stretch to push everything to the top
# #         main_layout.addStretch()
# #
# #         self.setLayout(main_layout)
# #
# #     def update_file_filter(self):
# #         """Update the file dialog filter based on selected file type."""
# #         # Currently, the filter is updated dynamically in browse_file.
# #         # This method is kept for potential future enhancements.
# #         pass
# #
# #     def toggle_options(self, state):
# #         """Disable other options when 'All Metadata' is selected."""
# #         if state == Qt.Checked:
# #             self.option_textprops.setChecked(False)
# #             self.option_mainprops.setChecked(False)
# #             self.option_appdata.setChecked(False)
# #             self.option_simple.setChecked(False)
# #
# #             self.option_textprops.setEnabled(False)
# #             self.option_mainprops.setEnabled(False)
# #             self.option_appdata.setEnabled(False)
# #             self.option_simple.setEnabled(False)
# #         else:
# #             self.option_textprops.setEnabled(True)
# #             self.option_mainprops.setEnabled(True)
# #             self.option_appdata.setEnabled(True)
# #             self.option_simple.setEnabled(True)
# #
# #     def browse_file(self):
# #         filetype = self.filetype_combo.currentText().lower()
# #         if filetype == "pdf":
# #             filter_str = "PDF Files (*.pdf)"
# #         elif filetype == "docx":
# #             filter_str = "DOCX Files (*.docx)"
# #         elif filetype == "pptx":
# #             filter_str = "PPTX Files (*.pptx)"
# #         else:
# #             filter_str = "All Files (*)"
# #
# #         file_path, _ = QFileDialog.getOpenFileName(
# #             self, "Select File", "", filter_str
# #         )
# #         if file_path:
# #             self.file_path_label.setText(file_path)
# #             self.results_text.clear()  # Clear previous results when a new file is selected
# #
# #     def extract_metadata(self):
# #         file_path = self.file_path_label.text()
# #         if not file_path or not os.path.isfile(file_path):
# #             QMessageBox.warning(self, "Invalid File", "Please select a valid file.")
# #             return
# #
# #         filetype = self.filetype_combo.currentText().lower()
# #         extractor_script = self.get_extractor_script(filetype)
# #         if not extractor_script:
# #             QMessageBox.critical(
# #                 self, "Error", f"No extractor script found for {filetype.upper()} files."
# #             )
# #             return
# #
# #         # Determine selected options
# #         options = []
# #         if self.option_all.isChecked():
# #             options.append("--all")
# #         else:
# #             if self.option_textprops.isChecked():
# #                 options.append("--textprops")
# #             if self.option_mainprops.isChecked():
# #                 options.append("--mainprops")
# #             if self.option_appdata.isChecked():
# #                 options.append("--appdata")
# #             if self.option_simple.isChecked():
# #                 options.append("--simple")
# #
# #         if not options:
# #             QMessageBox.warning(self, "No Options Selected", "Please select at least one metadata option.")
# #             return
# #
# #         # Prepare command
# #         command = [
# #             sys.executable, extractor_script
# #         ] + options + [file_path]
# #
# #         try:
# #             # Execute the extractor script with UTF-8 encoding
# #             result = subprocess.run(
# #                 command,
# #                 stdout=subprocess.PIPE,
# #                 stderr=subprocess.PIPE,
# #                 text=True,
# #                 encoding='utf-8',  # Specify UTF-8 encoding
# #                 check=True
# #             )
# #             metadata = result.stdout
# #             self.results_text.setPlainText(metadata)
# #         except subprocess.CalledProcessError as e:
# #             error_message = e.stderr if e.stderr else "An unknown error occurred."
# #             QMessageBox.critical(self, "Extraction Error", error_message)
# #         except UnicodeDecodeError as e:
# #             QMessageBox.critical(self, "Decoding Error", f"Failed to decode output: {str(e)}")
# #         except Exception as e:
# #             QMessageBox.critical(self, "Error", str(e))
# #
# #     def copy_to_clipboard(self):
# #         metadata = self.results_text.toPlainText()
# #         if metadata:
# #             clipboard = QApplication.clipboard()
# #             clipboard.setText(metadata)
# #             QMessageBox.information(self, "Copied", "Metadata copied to clipboard.")
# #         else:
# #             QMessageBox.warning(self, "No Data", "There is no metadata to copy.")
# #
# #     def get_extractor_script(self, filetype):
# #         script_map = {
# #             "docx": "_docx_meta.py",
# #             "pptx": "_pptx_meta.py",
# #             "pdf": "_pdf_meta.py"
# #         }
# #         script = script_map.get(filetype)
# #         if script and os.path.isfile(script):
# #             return os.path.abspath(script)
# #         return None
# #
# # def main():
# #     app = QApplication(sys.argv)
# #     window = MetadataExtractorApp()
# #     window.show()
# #     sys.exit(app.exec_())
# #
# # if __name__ == "__main__":
# #     main()
#
#
#
#
#
#
#
#
# import sys
# import os
# import subprocess
# from PyQt5.QtWidgets import (
#     QApplication, QWidget, QLabel, QPushButton, QFileDialog,
#     QVBoxLayout, QHBoxLayout, QComboBox, QMessageBox, QTextEdit,
#     QGroupBox, QCheckBox, QSpacerItem, QSizePolicy
# )
# from PyQt5.QtCore import Qt
# from PyQt5.QtGui import QIcon, QFont
#
# class MetadataExtractorApp(QWidget):
#     def __init__(self):
#         super().__init__()
#         self.setWindowTitle("Metadata Extractor")
#         self.setWindowIcon(QIcon("icon.png"))  # Optional: Add an icon.png in your project directory
#         self.setGeometry(100, 100, 900, 700)
#         self.init_ui()
#
#     def init_ui(self):
#         # Apply a simple stylesheet for better visuals
#         self.setStyleSheet("""
#             QWidget {
#                 font-family: Arial, sans-serif;
#                 font-size: 14px;
#             }
#             QGroupBox {
#                 font-weight: bold;
#                 margin-top: 10px;
#             }
#             QPushButton {
#                 padding: 8px 16px;
#                 font-size: 14px;
#             }
#             QTextEdit {
#                 background-color: #f0f0f0;
#                 font-family: Courier, monospace;
#                 font-size: 12px;
#             }
#             QCheckBox {
#                 margin-left: 10px;
#                 font-size: 14px;
#             }
#             QLabel {
#                 font-size: 14px;
#             }
#             QComboBox {
#                 font-size: 14px;
#             }
#         """)
#
#         # Main layout
#         main_layout = QVBoxLayout()
#         main_layout.setContentsMargins(20, 20, 20, 20)
#         main_layout.setSpacing(15)
#
#         # File type selection
#         filetype_layout = QHBoxLayout()
#         filetype_label = QLabel("Select File Type:")
#         filetype_label.setObjectName("filetype_label")
#         self.filetype_combo = QComboBox()
#         self.filetype_combo.setObjectName("filetype_combo")
#         self.filetype_combo.addItems(["DOCX", "PPTX", "PDF"])
#         self.filetype_combo.currentTextChanged.connect(self.update_file_filter)
#         filetype_layout.addWidget(filetype_label)
#         filetype_layout.addWidget(self.filetype_combo)
#         filetype_layout.addStretch()
#
#         # File selection group
#         file_group = QGroupBox("Select File")
#         file_group.setObjectName("file_group")
#         file_layout = QHBoxLayout()
#
#         self.file_path_label = QLabel("No file selected")
#         self.file_path_label.setObjectName("file_path_label")
#         self.file_path_label.setWordWrap(True)
#         select_file_btn = QPushButton("Browse")
#         select_file_btn.setObjectName("select_file_btn")
#         select_file_btn.clicked.connect(self.browse_file)
#
#         file_layout.addWidget(self.file_path_label)
#         file_layout.addWidget(select_file_btn)
#         file_group.setLayout(file_layout)
#
#         # Metadata options
#         options_group = QGroupBox("Metadata Options")
#         options_group.setObjectName("options_group")
#         options_layout = QVBoxLayout()
#         options_layout.setSpacing(10)  # Increase spacing between checkboxes
#
#         self.option_all = QCheckBox("All Metadata")
#         self.option_all.setObjectName("option_all")
#         self.option_all.stateChanged.connect(self.toggle_options)
#
#         self.option_textprops = QCheckBox("Text Properties")
#         self.option_textprops.setObjectName("option_textprops")
#         self.option_mainprops = QCheckBox("Main Properties")
#         self.option_mainprops.setObjectName("option_mainprops")
#         self.option_appdata = QCheckBox("Application Data")
#         self.option_appdata.setObjectName("option_appdata")
#         self.option_simple = QCheckBox("Simple Properties")
#         self.option_simple.setObjectName("option_simple")
#
#         options_layout.addWidget(self.option_all)
#         options_layout.addWidget(self.option_textprops)
#         options_layout.addWidget(self.option_mainprops)
#         options_layout.addWidget(self.option_appdata)
#         options_layout.addWidget(self.option_simple)
#
#         options_group.setLayout(options_layout)
#
#         # Extract button
#         extract_btn = QPushButton("Extract Metadata")
#         extract_btn.setObjectName("extract_btn")
#         extract_btn.setStyleSheet("background-color: #4CAF50; color: white;")
#         extract_btn.setMinimumHeight(40)  # Increase button height for better visibility
#         extract_btn.clicked.connect(self.extract_metadata)
#
#         # Results display
#         results_group = QGroupBox("Extracted Metadata")
#         results_group.setObjectName("results_group")
#         results_layout = QVBoxLayout()
#
#         # Horizontal layout for QTextEdit and Copy button
#         results_display_layout = QHBoxLayout()
#         self.results_text = QTextEdit()
#         self.results_text.setObjectName("results_text")
#         self.results_text.setReadOnly(True)
#         self.results_text.setFont(QFont("Courier", 10))
#         results_display_layout.addWidget(self.results_text)
#
#         copy_btn = QPushButton("Copy to Clipboard")
#         copy_btn.setObjectName("copy_btn")
#         copy_btn.setToolTip("Copy all extracted metadata to clipboard")
#         copy_btn.setStyleSheet("background-color: #2196F3; color: white;")
#         copy_btn.setMinimumHeight(40)  # Increase button height for better visibility
#         copy_btn.clicked.connect(self.copy_to_clipboard)
#         results_display_layout.addWidget(copy_btn)
#
#         results_layout.addLayout(results_display_layout)
#         results_group.setLayout(results_layout)
#
#         # Add widgets to main layout
#         main_layout.addLayout(filetype_layout)
#         main_layout.addWidget(file_group)
#         main_layout.addWidget(options_group)
#         main_layout.addWidget(extract_btn)
#         main_layout.addWidget(results_group)
#
#         # Add stretch to push everything to the top
#         main_layout.addStretch()
#
#         self.setLayout(main_layout)
#
#     def update_file_filter(self):
#         """Update the file dialog filter based on selected file type."""
#         # Currently, the filter is updated dynamically in browse_file.
#         # This method is kept for potential future enhancements.
#         pass
#
#     def toggle_options(self, state):
#         """Disable other options when 'All Metadata' is selected."""
#         if state == Qt.Checked:
#             self.option_textprops.setChecked(False)
#             self.option_mainprops.setChecked(False)
#             self.option_appdata.setChecked(False)
#             self.option_simple.setChecked(False)
#
#             self.option_textprops.setEnabled(False)
#             self.option_mainprops.setEnabled(False)
#             self.option_appdata.setEnabled(False)
#             self.option_simple.setEnabled(False)
#         else:
#             self.option_textprops.setEnabled(True)
#             self.option_mainprops.setEnabled(True)
#             self.option_appdata.setEnabled(True)
#             self.option_simple.setEnabled(True)
#
#     def browse_file(self):
#         filetype = self.filetype_combo.currentText().lower()
#         if filetype == "pdf":
#             filter_str = "PDF Files (*.pdf)"
#         elif filetype == "docx":
#             filter_str = "DOCX Files (*.docx)"
#         elif filetype == "pptx":
#             filter_str = "PPTX Files (*.pptx)"
#         else:
#             filter_str = "All Files (*)"
#
#         file_path, _ = QFileDialog.getOpenFileName(
#             self, "Select File", "", filter_str
#         )
#         if file_path:
#             self.file_path_label.setText(file_path)
#             self.results_text.clear()  # Clear previous results when a new file is selected
#
#     def extract_metadata(self):
#         file_path = self.file_path_label.text()
#         if not file_path or not os.path.isfile(file_path):
#             QMessageBox.warning(self, "Invalid File", "Please select a valid file.")
#             return
#
#         filetype = self.filetype_combo.currentText().lower()
#         extractor_script = self.get_extractor_script(filetype)
#         if not extractor_script:
#             QMessageBox.critical(
#                 self, "Error", f"No extractor script found for {filetype.upper()} files."
#             )
#             return
#
#         # Determine selected options
#         options = []
#         if self.option_all.isChecked():
#             options.append("--all")
#         else:
#             if self.option_textprops.isChecked():
#                 options.append("--textprops")
#             if self.option_mainprops.isChecked():
#                 options.append("--mainprops")
#             if self.option_appdata.isChecked():
#                 options.append("--appdata")
#             if self.option_simple.isChecked():
#                 options.append("--simple")
#
#         if not options:
#             QMessageBox.warning(self, "No Options Selected", "Please select at least one metadata option.")
#             return
#
#         # Prepare command
#         command = [
#             sys.executable, extractor_script
#         ] + options + [file_path]
#
#         try:
#             # Execute the extractor script with UTF-8 encoding
#             result = subprocess.run(
#                 command,
#                 stdout=subprocess.PIPE,
#                 stderr=subprocess.PIPE,
#                 text=True,
#                 encoding='utf-8',  # Specify UTF-8 encoding
#                 check=True
#             )
#             metadata = result.stdout
#             self.results_text.setPlainText(metadata)
#         except subprocess.CalledProcessError as e:
#             error_message = e.stderr if e.stderr else "An unknown error occurred."
#             QMessageBox.critical(self, "Extraction Error", error_message)
#         except UnicodeDecodeError as e:
#             QMessageBox.critical(self, "Decoding Error", f"Failed to decode output: {str(e)}")
#         except Exception as e:
#             QMessageBox.critical(self, "Error", str(e))
#
#     def copy_to_clipboard(self):
#         metadata = self.results_text.toPlainText()
#         if metadata:
#             clipboard = QApplication.clipboard()
#             clipboard.setText(metadata)
#             QMessageBox.information(self, "Copied", "Metadata copied to clipboard.")
#         else:
#             QMessageBox.warning(self, "No Data", "There is no metadata to copy.")
#
#     def get_extractor_script(self, filetype):
#         script_map = {
#             "docx": "_docx_meta.py",
#             "pptx": "_pptx_meta.py",
#             "pdf": "_pdf_meta.py"
#         }
#         script = script_map.get(filetype)
#         if script and os.path.isfile(script):
#             return os.path.abspath(script)
#         return None
#
# def main():
#     app = QApplication(sys.argv)
#     window = MetadataExtractorApp()
#     window.show()
#     sys.exit(app.exec_())
#
# if __name__ == "__main__":
#     main()











import sys
import os
import subprocess
from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QPushButton, QFileDialog,
    QVBoxLayout, QHBoxLayout, QComboBox, QMessageBox, QTextEdit,
    QGroupBox, QCheckBox, QSpacerItem, QSizePolicy
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QIcon, QFont

class MetadataExtractorApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Metadata Extractor")
        self.setWindowIcon(QIcon("icon.png"))  # Optional: Add an icon.png in your project directory
        self.setGeometry(100, 100, 900, 700)
        self.init_ui()

    def init_ui(self):
        # Apply a simple stylesheet for better visuals
        self.setStyleSheet("""
            QWidget {
                font-family: Arial, sans-serif;
                font-size: 14px;
            }
            QGroupBox {
                font-weight: bold;
                margin-top: 10px;
            }
            QPushButton {
                padding: 8px 16px;
                font-size: 14px;
            }
            QTextEdit {
                background-color: #f0f0f0;
                font-family: Courier, monospace;
                font-size: 12px;
            }
            QCheckBox {
                margin-left: 10px;
                font-size: 14px;
            }
            QLabel {
                font-size: 14px;
            }
            QComboBox {
                font-size: 14px;
            }
        """)

        # Main layout
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(15)

        # File type selection
        filetype_layout = QHBoxLayout()
        filetype_label = QLabel("Select File Type:")
        filetype_label.setObjectName("filetype_label")
        self.filetype_combo = QComboBox()
        self.filetype_combo.setObjectName("filetype_combo")
        self.filetype_combo.addItems(["DOCX", "PPTX", "PDF"])
        self.filetype_combo.currentTextChanged.connect(self.update_file_filter)
        filetype_layout.addWidget(filetype_label)
        filetype_layout.addWidget(self.filetype_combo)
        filetype_layout.addStretch()

        # File selection group
        file_group = QGroupBox("Select File")
        file_group.setObjectName("file_group")
        file_layout = QHBoxLayout()

        self.file_path_label = QLabel("No file selected")
        self.file_path_label.setObjectName("file_path_label")
        self.file_path_label.setWordWrap(True)
        select_file_btn = QPushButton("Browse")
        select_file_btn.setObjectName("select_file_btn")
        select_file_btn.clicked.connect(self.browse_file)

        file_layout.addWidget(self.file_path_label)
        file_layout.addWidget(select_file_btn)
        file_group.setLayout(file_layout)

        # Metadata options
        options_group = QGroupBox("Metadata Options")
        options_group.setObjectName("options_group")
        options_layout = QVBoxLayout()
        options_layout.setSpacing(10)  # Increase spacing between checkboxes

        self.option_all = QCheckBox("All Metadata")
        self.option_all.setObjectName("option_all")
        self.option_all.stateChanged.connect(self.toggle_options)

        self.option_textprops = QCheckBox("Text Properties")
        self.option_textprops.setObjectName("option_textprops")
        self.option_mainprops = QCheckBox("Main Properties")
        self.option_mainprops.setObjectName("option_mainprops")
        self.option_appdata = QCheckBox("Application Data")
        self.option_appdata.setObjectName("option_appdata")
        self.option_simple = QCheckBox("Simple Properties")
        self.option_simple.setObjectName("option_simple")

        options_layout.addWidget(self.option_all)
        options_layout.addWidget(self.option_textprops)
        options_layout.addWidget(self.option_mainprops)
        options_layout.addWidget(self.option_appdata)
        options_layout.addWidget(self.option_simple)

        options_group.setLayout(options_layout)

        # Extract button
        extract_btn = QPushButton("Extract Metadata")
        extract_btn.setObjectName("extract_btn")
        extract_btn.setStyleSheet("background-color: #4CAF50; color: white;")
        extract_btn.setMinimumHeight(40)  # Increase button height for better visibility
        extract_btn.clicked.connect(self.extract_metadata)

        # Results display
        results_group = QGroupBox("Extracted Metadata")
        results_group.setObjectName("results_group")
        results_layout = QVBoxLayout()

        # Horizontal layout for QTextEdit and Copy button
        results_display_layout = QHBoxLayout()
        self.results_text = QTextEdit()
        self.results_text.setObjectName("results_text")
        self.results_text.setReadOnly(True)
        self.results_text.setFont(QFont("Courier", 10))
        results_display_layout.addWidget(self.results_text)

        copy_btn = QPushButton("Copy to Clipboard")
        copy_btn.setObjectName("copy_btn")
        copy_btn.setToolTip("Copy all extracted metadata to clipboard")
        copy_btn.setStyleSheet("background-color: #2196F3; color: white;")
        copy_btn.setMinimumHeight(40)  # Increase button height for better visibility
        copy_btn.clicked.connect(self.copy_to_clipboard)
        results_display_layout.addWidget(copy_btn)

        results_layout.addLayout(results_display_layout)
        results_group.setLayout(results_layout)

        # Add widgets to main layout
        main_layout.addLayout(filetype_layout)
        main_layout.addWidget(file_group)
        main_layout.addWidget(options_group)
        main_layout.addWidget(extract_btn)
        main_layout.addWidget(results_group)

        # Add stretch to push everything to the top
        main_layout.addStretch()

        self.setLayout(main_layout)

    def update_file_filter(self):
        """Update the file dialog filter based on selected file type."""
        # Currently, the filter is updated dynamically in browse_file.
        # This method is kept for potential future enhancements.
        pass

    def toggle_options(self, state):
        """Disable other options when 'All Metadata' is selected."""
        if state == Qt.Checked:
            self.option_textprops.setChecked(False)
            self.option_mainprops.setChecked(False)
            self.option_appdata.setChecked(False)
            self.option_simple.setChecked(False)

            self.option_textprops.setEnabled(False)
            self.option_mainprops.setEnabled(False)
            self.option_appdata.setEnabled(False)
            self.option_simple.setEnabled(False)
        else:
            self.option_textprops.setEnabled(True)
            self.option_mainprops.setEnabled(True)
            self.option_appdata.setEnabled(True)
            self.option_simple.setEnabled(True)

    def browse_file(self):
        filetype = self.filetype_combo.currentText().lower()
        if filetype == "pdf":
            filter_str = "PDF Files (*.pdf)"
        elif filetype == "docx":
            filter_str = "DOCX Files (*.docx)"
        elif filetype == "pptx":
            filter_str = "PPTX Files (*.pptx)"
        else:
            filter_str = "All Files (*)"

        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select File", "", filter_str
        )
        if file_path:
            self.file_path_label.setText(file_path)
            self.results_text.clear()  # Clear previous results when a new file is selected

    def extract_metadata(self):
        file_path = self.file_path_label.text()
        if not file_path or not os.path.isfile(file_path):
            QMessageBox.warning(self, "Invalid File", "Please select a valid file.")
            return

        filetype = self.filetype_combo.currentText().lower()
        extractor_script = self.get_extractor_script(filetype)
        if not extractor_script:
            QMessageBox.critical(
                self, "Error", f"No extractor script found for {filetype.upper()} files."
            )
            return

        # Determine selected options
        options = []
        if self.option_all.isChecked():
            options.append("--all")
        else:
            if self.option_textprops.isChecked():
                options.append("--textprops")
            if self.option_mainprops.isChecked():
                options.append("--mainprops")
            if self.option_appdata.isChecked():
                options.append("--appdata")
            if self.option_simple.isChecked():
                options.append("--simple")

        if not options:
            QMessageBox.warning(self, "No Options Selected", "Please select at least one metadata option.")
            return

        # Prepare command
        command = [
            sys.executable, extractor_script
        ] + options + [file_path]

        try:
            # Execute the extractor script with UTF-8 encoding
            result = subprocess.run(
                command,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                encoding='utf-8',  # Specify UTF-8 encoding
                check=True
            )
            metadata = result.stdout
            self.results_text.setPlainText(metadata)
        except subprocess.CalledProcessError as e:
            error_message = e.stderr if e.stderr else "An unknown error occurred."
            QMessageBox.critical(self, "Extraction Error", error_message)
        except UnicodeDecodeError as e:
            QMessageBox.critical(self, "Decoding Error", f"Failed to decode output: {str(e)}")
        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))

    def copy_to_clipboard(self):
        metadata = self.results_text.toPlainText()
        if metadata:
            clipboard = QApplication.clipboard()
            clipboard.setText(metadata)
            QMessageBox.information(self, "Copied", "Metadata copied to clipboard.")
        else:
            QMessageBox.warning(self, "No Data", "There is no metadata to copy.")

    def get_extractor_script(self, filetype):
        script_map = {
            "docx": "_docx_meta.py",
            "pptx": "_pptx_meta.py",
            "pdf": "_pdf_meta.py"
        }
        script = script_map.get(filetype)
        if script and os.path.isfile(script):
            return os.path.abspath(script)
        return None

def main():
    app_instance = QApplication(sys.argv)
    window = MetadataExtractorApp()
    window.show()
    sys.exit(app_instance.exec_())

if __name__ == "__main__":
    main()
