#!/usr/bin/env python3
import os
import tempfile
from unittest.mock import patch, Mock

# Change to the TestAll directory
os.chdir('/home/newuser/TestAll')

from ocrspace_core import process_single

def test_word_file_creation():
    """Test that a Word file is created when OCR returns 'キュー'."""
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
                image_path="dummy.png",  # This won't be used because we mock the request
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
                        return True
                    else:
                        print("WARNING: Word file is empty.")
                        return False
                else:
                    print(f"FAIL: Word file does not exist at {word_path}")
                    return False
            else:
                print("FAIL: 'word' key not in results")
                return False

if __name__ == "__main__":
    if test_word_file_creation():
        print("\nTest passed.")
    else:
        print("\nTest failed.")