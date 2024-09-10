import pytest
from unittest.mock import patch, MagicMock
import zipfile
import defusedxml.ElementTree as ET
from _docx_meta import extract_xml_data, filter_metadata, check_decompression_size

# Test Cases for check_decompression_size

@patch('_docx_meta.zipfile.ZipFile')
def test_check_decompression_size_valid(mock_zipfile):
    # Arrange
    mock_zip = MagicMock()
    mock_zip.infolist.return_value = [MagicMock(file_size=1000), MagicMock(file_size=2000)]
    mock_zipfile.return_value.__enter__.return_value = mock_zip
    
    # Act
    result = check_decompression_size("test.docx", 5000)
    
    # Assert
    assert result == 3000  # Total uncompressed size

@patch('_docx_meta.zipfile.ZipFile')
def test_check_decompression_size_exceeds_limit(mock_zipfile):
    # Arrange
    mock_zip = MagicMock()
    mock_zip.infolist.return_value = [MagicMock(file_size=1000), MagicMock(file_size=5000)]
    mock_zipfile.return_value.__enter__.return_value = mock_zip
    
    # Act & Assert
    with pytest.raises(ValueError, match="Extraction size exceeds the allowed limit"):
        check_decompression_size("test.docx", 4000)

@patch('_docx_meta.zipfile.ZipFile')
def test_check_decompression_size_empty_zip(mock_zipfile):
    # Arrange
    mock_zip = MagicMock()
    mock_zip.infolist.return_value = []
    mock_zipfile.return_value.__enter__.return_value = mock_zip
    
    # Act
    result = check_decompression_size("test.docx", 5000)
    
    # Assert
    assert result == 0  # Empty ZIP file should return size 0

@patch('_docx_meta.zipfile.ZipFile')
def test_check_decompression_size_zero_file_size(mock_zipfile):
    # Arrange
    mock_zip = MagicMock()
    mock_zip.infolist.return_value = [MagicMock(file_size=0), MagicMock(file_size=0)]
    mock_zipfile.return_value.__enter__.return_value = mock_zip
    
    # Act
    result = check_decompression_size("test.docx", 5000)
    
    # Assert
    assert result == 0  # ZIP files with zero size should return 0

# Parameterized test for extract_xml_data for valid and invalid cases
@pytest.mark.parametrize("mock_core_find, mock_app_find, expected_metadata", [
    (
        # Valid case with all fields
        {
            'dc:creator': MagicMock(text="AuthorName"),
            'dc:description': MagicMock(text="DescriptionText"),
            'cp:lastModifiedBy': MagicMock(text="ModifiedByUser"),
            'cp:revision': MagicMock(text="1"),
            'dc:title': MagicMock(text="TitleText"),
            'dc:subject': MagicMock(text="SubjectText"),
            'dcterms:created': MagicMock(text="2024-09-09T15:06:23Z"),
            'dcterms:modified': MagicMock(text="2024-09-10T15:06:23Z"),
            'dc:language': MagicMock(text="en-US")
        },
        {
            'ep:Template': MagicMock(text="TemplateName"),
            'ep:TotalTime': MagicMock(text="10"),
            'ep:Pages': MagicMock(text="50"),
            'ep:Words': MagicMock(text="1000"),
            'ep:Characters': MagicMock(text="5000"),
            'ep:Application': MagicMock(text="WordProcessorApp"),
            'ep:AppVersion': MagicMock(text="1.0")
        },
        {
            "Author": "AuthorName",
            "Description": "DescriptionText",
            "LastModifiedBy": "ModifiedByUser",
            "Revision": "1",
            "Title": "TitleText",
            "Subject": "SubjectText",
            "Created": "2024-09-09T15:06:23Z",
            "Modified": "2024-09-10T15:06:23Z",
            "Language": "en-US",
            "Template": "TemplateName",
            "TotalTime": "10",
            "Pages": "50",
            "Words": "1000",
            "Characters": "5000",
            "Application": "WordProcessorApp",
            "AppVersion": "1.0"
        }
    ),
    (
        # Case with missing fields in XML
        {
            'dc:creator': MagicMock(text="AuthorName"),
            'dc:description': None,  # Missing description
            'cp:lastModifiedBy': MagicMock(text="ModifiedByUser"),
            'cp:revision': MagicMock(text="1"),
            'dc:title': None,  # Missing title
            'dc:subject': MagicMock(text="SubjectText"),
            'dcterms:created': MagicMock(text="2024-09-09T15:06:23Z"),
            'dcterms:modified': None,  # Missing modified date
            'dc:language': MagicMock(text="en-US")
        },
        {
            'ep:Template': MagicMock(text="TemplateName"),
            'ep:TotalTime': None,  # Missing total time
            'ep:Pages': MagicMock(text="50"),
            'ep:Words': MagicMock(text="1000"),
            'ep:Characters': MagicMock(text="5000"),
            'ep:Application': MagicMock(text="WordProcessorApp"),
            'ep:AppVersion': MagicMock(text="1.0")
        },
        {
            "Author": "AuthorName",
            "Description": "Not Available",
            "LastModifiedBy": "ModifiedByUser",
            "Revision": "1",
            "Title": "Not Available",
            "Subject": "SubjectText",
            "Created": "2024-09-09T15:06:23Z",
            "Modified": "Not Available",
            "Language": "en-US",
            "Template": "TemplateName",
            "TotalTime": "Not Available",
            "Pages": "50",
            "Words": "1000",
            "Characters": "5000",
            "Application": "WordProcessorApp",
            "AppVersion": "1.0"
        }
    ),
])
@patch('_docx_meta.zipfile.ZipFile')
@patch('_docx_meta.ET.parse')
def test_extract_xml_data_valid(mock_et_parse, mock_zipfile, mock_core_find, mock_app_find, expected_metadata):
    # Arrange
    mock_zip = MagicMock()
    mock_zip.open.return_value.__enter__.side_effect = [MagicMock(), MagicMock()]
    mock_zipfile.return_value.__enter__.return_value = mock_zip

    core_xml_mock = MagicMock()
    app_xml_mock = MagicMock()
    mock_et_parse.side_effect = [core_xml_mock, app_xml_mock]

    core_root = MagicMock()
    app_root = MagicMock()

    core_xml_mock.getroot.return_value = core_root
    app_xml_mock.getroot.return_value = app_root

    core_root.find.side_effect = lambda tag, ns: mock_core_find.get(tag, None)
    app_root.find.side_effect = lambda tag, ns: mock_app_find.get(tag, None)

    # Act
    metadata = extract_xml_data("test.docx", 1000000, 500000)

    # Assert
    assert metadata == expected_metadata


@patch('_docx_meta.zipfile.ZipFile')
def test_extract_xml_data_invalid_zip(mock_zipfile):
    # Arrange
    mock_zipfile.side_effect = zipfile.BadZipFile

    # Act & Assert
    with pytest.raises(ValueError, match="is not a valid ZIP file"):
        extract_xml_data("invalid.docx", 1000000, 500000)

@patch('_docx_meta.zipfile.ZipFile')
def test_extract_xml_data_missing_file(mock_zipfile):
    # Arrange
    mock_zip = MagicMock()
    mock_zip.open.side_effect = KeyError("Missing file in ZIP")
    mock_zipfile.return_value.__enter__.return_value = mock_zip

    # Act & Assert
    with pytest.raises(ValueError, match="Error accessing file in the DOCX archive"):
        extract_xml_data("test.docx", 1000000, 500000)

@patch('_docx_meta.zipfile.ZipFile')
@patch('_docx_meta.ET.parse')
def test_extract_xml_data_malformed_xml(mock_et_parse, mock_zipfile):
    # Arrange
    mock_zip = MagicMock()
    mock_zip.open.return_value.__enter__.side_effect = [MagicMock(), MagicMock()]
    mock_zipfile.return_value.__enter__.return_value = mock_zip

    # Simulate a parse error
    mock_et_parse.side_effect = ET.ParseError("Malformed XML")

    # Act & Assert
    with pytest.raises(ValueError, match="Error parsing XML data"):
        extract_xml_data("test.docx", 1000000, 500000)


# Parameterized tests for filter_metadata with different options
@pytest.mark.parametrize("metadata, options, expected", [
    (
        {
            "Author": "AuthorName",
            "Pages": "50",
            "Words": "1000",
            "Language": "en-US",
        },
        ['textprops'],
        {"Pages": "50", "Words": "1000", "Language": "en-US"}
    ),
    (
        {
            "Author": "AuthorName",
            "Created": "2024-09-09T15:06:23Z",
            "Modified": "2024-09-10T15:06:23Z",
        },
        ['simple'],
        {"Author": "AuthorName", "Created": "2024-09-09T15:06:23Z", "Modified": "2024-09-10T15:06:23Z"}
    ),
    (
        {
            "Author": "AuthorName",
            "Description": "DescriptionText",
            "LastModifiedBy": "ModifiedByUser",
            "Revision": "1",
            "Title": "TitleText",
            "Pages": "50",
            "Words": "1000",
            "Characters": "5000",
        },
        ['all'],
        {
            "Author": "AuthorName",
            "Description": "DescriptionText",
            "LastModifiedBy": "ModifiedByUser",
            "Revision": "1",
            "Title": "TitleText",
            "Pages": "50",
            "Words": "1000",
            "Characters": "5000",
        }
    )
])
def test_filter_metadata(metadata, options, expected):
    result = filter_metadata(metadata, options)
    assert result == expected
