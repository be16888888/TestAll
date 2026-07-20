#!/usr/bin/env python3
import os
import tempfile
from unittest.mock import patch, Mock

# Change to the TestAll directory
os.chdir('/home/newuser/TestAll')

from ocrspace_core import process_single

def test_word_file_creation_with_mock():
    """Test that a Word file is created when OCR returns 'キュー'."""
    # Create a dummy image file so that the file exists
    with open('dummy.png', 'wb') as f:
        # Minimal PNG: 1x1 black pixel
        png_data = b'\x89PNG\r\n\x1a\n\0\0\0\rIHDR\0\0\0\1\0\0\0\1\8\0\0\0\1w\53\xde\0\0\0\tpHYs\0\0\0\11\0\0\0\11\x01\x00\x9a\x9c\x18\0\0\0\nIDATx\x9cc`\0\0\0\0\0\0\0\5\0\x01\r\n-\xb4\0\0\0\0IEND\xaeB`\x82'
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
    if test_word_file_creation_with_mock():
        print("\nTest passed.")
    else:
        print("\nTest failed.")