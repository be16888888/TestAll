#!/usr/bin/env python3
import os
import tempfile
from unittest.mock import patch, Mock

# Change to the TestAll directory
os.chdir('/home/newuser/TestAll')

from ocrspace_core import process_single

def test_word_file_creation_with_kyu():
    """Test that a Word file is created when OCR returns 'キュー'."""
    # Create a dummy image file (a small PNG) so that the file exists
    with open('dummy.png', 'wb') as f:
        # PNG header + 1x1 black pixel
        png_data = (
            b'\x89PNG\r\n\x1a\n'
            b'\x00\x00\x00\rIHDR'
            b'\x00\x00\x00\x01\x00\x00\x00\x01'
            b'\x08\x00\x00\x00\x01w\x53\xde'
            b'\x00\x00\x00\tpHYs'
            b'\x00\x00\x00\x11\x00\x00\x00\x11'
            b'\x01\x00\x9a\x9c\x18'
            b'\x00\x00\x00\nIDATx\x9cc`\x00\x00\x00\x00\x00\x00\x00\x05\x00\x01\r\n-\xb4'
            b'\x00\x00\x00\x00IEND\xaeB`\x82'
        )
        f.write(png_data)

    with tempfile.TemporaryDirectory() as tmpdir:
        # Mock the API call to return 'キュー'
        mock_response = Mock()
        mock_response.status_code = 200
        mock_response.json.return_value = {
            "ParsedResults": [{"ParsedText": "キュー"}],
            "OCRExitCode": 1,
            "IsErroredOnProcessing": False
        }

        with patch('requests.post', return_value=mock_response):
            # Call process_single with save_word=True
            results = process_single(
                api_key="fake_key",
                image_path="dummy.png",  # Now this file exists
                output_dir=tmpdir,
                save_word=True,
                save_excel=False,
                save_text=False,
                save_md=False,
                language='cht',
                isTable=True
            )

            # Check if the Word file was created
            if 'word' in results:
                word_path = results['word']
                if os.path.exists(word_path):
                    print(f"SUCCESS: Word file created at {word_path}")
                    # Check the content? We could, but let's just see if it's non-zero
                    size = os.path.getsize(word_path)
                    print(f"Word file size: {size} bytes")
                    if size > 0:
                        print("Word file is not empty.")
                        # Clean up the dummy image
                        os.remove('dummy.png')
                        return True
                    else:
                        print("WARNING: Word file is empty.")
                        os.remove('dummy.png')
                        return False
                else:
                    print(f"FAIL: Word file does not exist at {word_path}")
                    os.remove('dummy.png')
                    return False
            else:
                print("FAIL: 'word' key not in results")
                os.remove('dummy.png')
                return False

if __name__ == "__main__":
    if test_word_file_creation_with_kyu():
        print("\nTest passed.")
    else:
        print("\nTest failed.")