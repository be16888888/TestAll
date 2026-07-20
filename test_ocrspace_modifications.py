#!/usr/bin/env python3
"""Quick test of the modified OCR.Space API function."""

import os
from unittest.mock import patch, Mock

# Change to the TestAll directory to find WebOcrAPI.json
os.chdir('/home/newuser/TestAll')

from ocrspace_core import call_ocrspace_api

def test_no_text_detected_triggers_retry():
    """Test that '[No text detected]' raises an exception (to trigger retry in process_single)."""
    mock_response = Mock()
    mock_response.status_code = 200
    mock_response.json.return_value = {
        "ParsedResults": [{"ParsedText": "[No text detected]"}],
        "OCRExitCode": 1,
        "IsErroredOnProcessing": False
    }
    
    with patch('requests.post', return_value=mock_response):
        try:
            result = call_ocrspace_api("fake_key", "dummy.png")
            print("FAIL: Expected exception for '[No text detected]'")
            return False
        except Exception as e:
            if "OCR detected no text" in str(e):
                print("PASS: '[No text detected]' correctly triggers retry exception")
                return True
            else:
                print(f"FAIL: Unexpected exception: {e}")
                return False

def test_normal_text_works():
    """Test that normal text is returned correctly."""
    mock_response = Mock()
    mock_response.status_code = 200
    mock_response.json.return_value = {
        "ParsedResults": [{"ParsedText": "Normal OCR text"}],
        "OCRExitCode": 1,
        "IsErroredOnProcessing": False
    }
    
    with patch('requests.post', return_value=mock_response):
        try:
            result = call_ocrspace_api("fake_key", "dummy.png")
            if result == "Normal OCR text":
                print("PASS: Normal text returned correctly")
                return True
            else:
                print(f"FAIL: Expected 'Normal OCR text', got '{result}'")
                return False
        except Exception as e:
            print(f"FAIL: Unexpected exception: {e}")
            return False

def test_kaji_character_no_longer_triggers_retry():
    """Test that the Japanese character 'キュー' does NOT trigger a false retry."""
    mock_response = Mock()
    mock_response.status_code = 200
    mock_response.json.return_value = {
        "ParsedResults": [{"ParsedText": "キュー"}],
        "OCRExitCode": 1,
        "IsErroredOnProcessing": False
    }
    
    with patch('requests.post', return_value=mock_response):
        try:
            result = call_ocrspace_api("fake_key", "dummy.png")
            if result == "キュー":
                print("PASS: 'キュー' returned as-is (no false retry)")
                return True
            else:
                print(f"FAIL: Expected 'キュー', got '{result}'")
                return False
        except Exception as e:
            print(f"FAIL: Unexpected exception for 'キュー': {e}")
            return False

if __name__ == "__main__":
    print("Testing OCR.Space API modifications in TestAll directory...\n")
    
    tests = [
        test_no_text_detected_triggers_retry,
        test_normal_text_works,
        test_kaji_character_no_longer_triggers_retry,
    ]
    
    passed = 0
    for test in tests:
        if test():
            passed += 1
        print()
    
    print(f"Results: {passed}/{len(tests)} tests passed")
    if passed == len(tests):
        print("All tests passed! ✅")
    else:
        print("Some tests failed. ❌")