import unittest
from unittest.mock import patch, MagicMock
from io import BytesIO
from PyPDF2 import PdfReader
import _pdf_meta
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter

def create_in_memory_pdf_with_metadata(metadata=None):
    """Create an in-memory PDF file with optional metadata embedded."""
    packet = BytesIO()
    can = canvas.Canvas(packet, pagesize=letter)
    
    # Draw something on the PDF
    can.drawString(100, 750, "Hello, this is a test PDF.")
    can.showPage()
    
    # Add metadata to the PDF
    if metadata:
        can.setAuthor(metadata.get('/Author', ''))
        can.setTitle(metadata.get('/Title', ''))
        can.setProducer(metadata.get('/Producer', ''))

    can.save()

    # Move to the beginning of the BytesIO buffer
    packet.seek(0)
    
    return packet

class TestPdfMeta(unittest.TestCase):

    @patch('os.path.getsize')
    def test_extract_pdf_data_within_limit(self, mock_getsize):
        """Test successful extraction of PDF metadata within limits using in-memory PDF."""
        mock_getsize.return_value = 500 * 1024 * 1024  # 500 MB
        
        # Create an in-memory PDF with metadata
        pdf_file = create_in_memory_pdf_with_metadata(metadata={
            '/Author': 'Test Author',
            '/Title': 'Test Title',
            '/Producer': 'Test Producer',
            '/CreationDate': 'D:20240909150118',  
            '/ModDate': 'D:20240910150118'         
        })
        
        result = _pdf_meta.extract_pdf_data(pdf_file, 1 * 1024 * 1024 * 1024, 200 * 1024 * 1024)

        # Assertions for the extracted metadata
        self.assertEqual(result['/Author'], 'Test Author')
        self.assertEqual(result['/Title'], 'Test Title')

    @patch('os.path.getsize')
    def test_check_file_size_within_limit(self, mock_getsize):
        """Test if file size is within the extraction limit."""
        mock_getsize.return_value = 500 * 1024 * 1024  # 500 MB
        result = _pdf_meta.check_file_size('dummy.pdf', 1 * 1024 * 1024 * 1024)  # 1 GB limit
        self.assertEqual(result, 500 * 1024 * 1024)
    
    @patch('os.path.getsize')
    def test_check_file_size_exceeds_limit(self, mock_getsize):
        """Test if file size exceeds the extraction limit."""
        mock_getsize.return_value = 1.5 * 1024 * 1024 * 1024  # 1.5 GB
        with self.assertRaises(ValueError) as context:
            _pdf_meta.check_file_size('dummy.pdf', 1 * 1024 * 1024 * 1024)  # 1 GB limit
        self.assertIn('File size exceeds the allowed limit', str(context.exception))

    def test_filter_metadata_all(self):
        """Test filtering when 'all' option is selected."""
        metadata = {
            '/Author': 'Test Author',
            '/Title': 'Test Title',
            '/Producer': 'Test Producer',
            '/CreationDate': 'D:20240909150118',  
            '/ModDate': 'D:20240910150118'        
        }
        options = ['all']
        result = _pdf_meta.filter_metadata(metadata, options)
        self.assertEqual(result, metadata)
    
    def test_filter_metadata_textprops(self):
        """Test filtering for text-related properties."""
        metadata = {
            '/Author': 'Test Author',
            '/Title': 'Test Title',
            '/Pages': '10',
            '/Subject': 'Test Subject',
            '/CreationDate': 'D:20240909150118',  
            '/ModDate': 'D:20240910150118'        
        }
        options = ['textprops']
        result = _pdf_meta.filter_metadata(metadata, options)
        self.assertEqual(result, {
            '/Title': 'Test Title',
            '/Pages': '10',
            '/Subject': 'Test Subject'
        })
    
    def test_filter_metadata_mainprops(self):
        """Test filtering for main properties."""
        metadata = {
            '/Author': 'Test Author',
            '/Title': 'Test Title',
            '/CreationDate': 'D:20240909150118',  
            '/ModDate': 'D:20240910150118',        
            '/Producer': 'Test Producer',
            '/Pages': '10'
        }
        options = ['mainprops']
        result = _pdf_meta.filter_metadata(metadata, options)
        self.assertEqual(result, {
            '/Author': 'Test Author',
            '/Title': 'Test Title',
            '/CreationDate': 'D:20240909150118',
            '/ModDate': 'D:20240910150118'
        })

    def test_filter_metadata_simple(self):
        """Test filtering for simple properties."""
        metadata = {
            '/Author': 'Test Author',
            '/CreationDate': 'D:20240909150118',  
            '/ModDate': 'D:20240910150118',        
            '/Producer': 'Test Producer'
        }
        options = ['simple']
        result = _pdf_meta.filter_metadata(metadata, options)
        self.assertEqual(result, {
            '/Author': 'Test Author',
            '/CreationDate': 'D:20240909150118',
            '/ModDate': 'D:20240910150118'
        })

if __name__ == '__main__':
    unittest.main()
