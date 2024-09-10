import pytest
from unittest.mock import patch, MagicMock
import zipfile
import defusedxml.ElementTree as ET
from _pptx_meta import extract_xml_data, filter_metadata, check_decompression_size

# Test Cases for check_decompression_size

@patch('_pptx_meta.zipfile.ZipFile')
def test_check_decompression_size_valid(mock_zipfile):
    # Arrange
    mock_zip = MagicMock()
    mock_zip.infolist.return_value = [MagicMock(file_size=1000), MagicMock(file_size=2000)]
    mock_zipfile.return_value.__enter__.return_value = mock_zip
    
    # Act
    result = check_decompression_size("test.pptx", 5000)
    
    # Assert
    assert result == 3000  # Total uncompressed size

@patch('_pptx_meta.zipfile.ZipFile')
def test_check_decompression_size_exceeds_limit(mock_zipfile):
    # Arrange
    mock_zip = MagicMock()
    mock_zip.infolist.return_value = [MagicMock(file_size=1000), MagicMock(file_size=5000)]
    mock_zipfile.return_value.__enter__.return_value = mock_zip
    
    # Act & Assert
    with pytest.raises(ValueError, match="Extraction size exceeds the allowed limit"):
        check_decompression_size("test.pptx", 4000)

@patch('_pptx_meta.zipfile.ZipFile')
def test_check_decompression_size_empty_zip(mock_zipfile):
    # Arrange
    mock_zip = MagicMock()
    mock_zip.infolist.return_value = []
    mock_zipfile.return_value.__enter__.return_value = mock_zip
    
    # Act
    result = check_decompression_size("test.pptx", 5000)
    
    # Assert
    assert result == 0  # Empty ZIP file should return size 0

@patch('_pptx_meta.zipfile.ZipFile')
def test_check_decompression_size_zero_file_size(mock_zipfile):
    # Arrange
    mock_zip = MagicMock()
    mock_zip.infolist.return_value = [MagicMock(file_size=0), MagicMock(file_size=0)]
    mock_zipfile.return_value.__enter__.return_value = mock_zip
    
    # Act
    result = check_decompression_size("test.pptx", 5000)
    
    # Assert
    assert result == 0  # ZIP files with zero size should return 0


# Parameterized test for extract_xml_data for valid and invalid cases
@pytest.mark.parametrize("mock_core_find, mock_app_find, expected_metadata", [
    (
        # Valid case with all fields
        {
            'dc:creator': MagicMock(text="AuthorName"),
            'cp:lastModifiedBy': MagicMock(text="ModifiedByUser"),
            'cp:revision': MagicMock(text="1"),
            'dcterms:created': MagicMock(text="2024-09-09T15:15:08Z"),
            'dcterms:modified': MagicMock(text="2024-09-10T15:15:08Z"),
        },
        {
            'ep:Template': MagicMock(text="TemplateName"),
            'ep:TotalTime': MagicMock(text="10"),
            'ep:Words': MagicMock(text="1000"),
            'ep:Paragraphs': MagicMock(text="500"),
            'ep:Slides': MagicMock(text="20"),
            'ep:Notes': MagicMock(text="5"),
            'ep:HiddenSlides': MagicMock(text="1"),
            'ep:MMClips': MagicMock(text="3"),
            'ep:Application': MagicMock(text="PresentationApp"),
            'ep:PresentationFormat': MagicMock(text="16:9"),
            'ep:AppVersion': MagicMock(text="1.0"),
        },
        {
            "Author": "AuthorName",
            "LastModifiedBy": "ModifiedByUser",
            "Revision": "1",
            "Created": "2024-09-09T15:15:08Z",
            "Modified": "2024-09-10T15:15:08Z",
            "Template": "TemplateName",
            "TotalTime": "10",
            "Words": "1000",
            "Paragraphs": "500",
            "Slides": "20",
            "Notes": "5",
            "HiddenSlides": "1",
            "MMClips": "3",
            "Application": "PresentationApp",
            "PresentationFormat": "16:9",
            "AppVersion": "1.0"
        }
    ),
    (
        # Case with missing fields in XML
        {
            'dc:creator': MagicMock(text="AuthorName"),
            'cp:lastModifiedBy': MagicMock(text="ModifiedByUser"),
            'cp:revision': None,  # Missing revision
            'dcterms:created': MagicMock(text="2024-09-09T15:15:08Z"),
            'dcterms:modified': None,  # Missing modified date
        },
        {
            'ep:Template': MagicMock(text="TemplateName"),
            'ep:TotalTime': None,  # Missing total time
            'ep:Words': MagicMock(text="1000"),
            'ep:Paragraphs': None,  # Missing paragraphs
            'ep:Slides': MagicMock(text="20"),
            'ep:Notes': MagicMock(text="5"),
            'ep:HiddenSlides': None,  # Missing hidden slides
            'ep:MMClips': MagicMock(text="3"),
            'ep:Application': MagicMock(text="PresentationApp"),
            'ep:PresentationFormat': None,  # Missing presentation format
            'ep:AppVersion': MagicMock(text="1.0"),
        },
        {
            "Author": "AuthorName",
            "LastModifiedBy": "ModifiedByUser",
            "Revision": "Not Available",
            "Created": "2024-09-09T15:15:08Z",
            "Modified": "Not Available",
            "Template": "TemplateName",
            "TotalTime": "Not Available",
            "Words": "1000",
            "Paragraphs": "Not Available",
            "Slides": "20",
            "Notes": "5",
            "HiddenSlides": "Not Available",
            "MMClips": "3",
            "Application": "PresentationApp",
            "PresentationFormat": "Not Available",
            "AppVersion": "1.0"
        }
    ),
])
@patch('_pptx_meta.zipfile.ZipFile')
@patch('_pptx_meta.ET.parse')
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
    metadata = extract_xml_data("test.pptx", 1000000, 500000)

    # Assert
    assert metadata == expected_metadata


@patch('_pptx_meta.zipfile.ZipFile')
def test_extract_xml_data_invalid_zip(mock_zipfile):
    # Arrange
    mock_zipfile.side_effect = zipfile.BadZipFile

    # Act & Assert
    with pytest.raises(ValueError, match="is not a valid ZIP file"):
        extract_xml_data("invalid.pptx", 1000000, 500000)

@patch('_pptx_meta.zipfile.ZipFile')
def test_extract_xml_data_missing_file(mock_zipfile):
    # Arrange
    mock_zip = MagicMock()
    mock_zip.open.side_effect = KeyError("Missing file in ZIP")
    mock_zipfile.return_value.__enter__.return_value = mock_zip

    # Act & Assert
    with pytest.raises(ValueError, match="Error accessing file in the PPTX archive"):
        extract_xml_data("test.pptx", 1000000, 500000)

@patch('_pptx_meta.zipfile.ZipFile')
@patch('_pptx_meta.ET.parse')
def test_extract_xml_data_malformed_xml(mock_et_parse, mock_zipfile):
    # Arrange
    mock_zip = MagicMock()
    mock_zip.open.return_value.__enter__.side_effect = [MagicMock(), MagicMock()]
    mock_zipfile.return_value.__enter__.return_value = mock_zip

    # Simulate a parse error
    mock_et_parse.side_effect = ET.ParseError("Malformed XML")

    # Act & Assert
    with pytest.raises(ValueError, match="Error parsing XML data"):
        extract_xml_data("test.pptx", 1000000, 500000)


# Parameterized tests for filter_metadata with different options
@pytest.mark.parametrize("metadata, options, expected", [
    (
        {
            "Words": "1000",
            "Paragraphs": "500",
            "Slides": "20",
            "Notes": "5",
            "HiddenSlides": "1",
            "MMClips": "3",
        },
        ['textprops'],
        {
            "Words": "1000",
            "Paragraphs": "500",
            "Slides": "20",
            "Notes": "5",
            "HiddenSlides": "1",
            "MMClips": "3",
        }
    ),
    (
        {
            "Author": "AuthorName",
            "Created": "2024-09-09T15:15:08Z",
            "Modified": "2024-09-10T15:15:08Z",
        },
        ['simple'],
        {
            "Author": "AuthorName",
            "Created": "2024-09-09T15:15:08Z",
            "Modified": "2024-09-10T15:15:08Z"
        }
    ),
    (
        {
            "Author": "AuthorName",
            "LastModifiedBy": "ModifiedByUser",
            "Revision": "1",
            "Created": "2024-09-09T15:15:08Z",
            "Modified": "2024-09-10T15:15:08Z",
            "Template": "TemplateName",
            "TotalTime": "10",
            "Slides": "20",
            "Application": "PresentationApp"
        },
        ['all'],
        {
            "Author": "AuthorName",
            "LastModifiedBy": "ModifiedByUser",
            "Revision": "1",
            "Created": "2024-09-09T15:15:08Z",
            "Modified": "2024-09-10T15:15:08Z",
            "Template": "TemplateName",
            "TotalTime": "10",
            "Slides": "20",
            "Application": "PresentationApp"
        }
    )
])
def test_filter_metadata(metadata, options, expected):
    result = filter_metadata(metadata, options)
    assert result == expected
