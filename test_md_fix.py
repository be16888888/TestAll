#!/usr/bin/env python3
"""Quick test of the Markdown table detection fix."""

import os
import sys
os.chdir('/home/newuser/TestAll')
sys.path.insert(0, '/home/newuser/TestAll')

from ocrspace_core import split_text_around_first_table

# Simulated OCR text from the user's sample
test_text = """長及商业系统里
(6庫)進、銷貨庫存表
115年3月17日
填單者:\ 辨識不完全

品項	數量	單價	金額	成本	毛利	備註
青江菜 | 100 | 20 | 2000 | 1500 | 500 | 
蘿蔔加工 | 50 | 30 | 1500 | 1000 | 500 |
甜豆加工 | 46.5 | 1 | 45. | 177 | 0.24 | 9

辨識不完全"""

print("Testing split_text_around_first_table...")
before, table_block, after = split_text_around_first_table(test_text)

print(f"BEFORE:\n{before}\n")
print(f"TABLE_BLOCK:\n{table_block}\n")
print(f"AFTER:\n{after}\n")

# Check if table contains all data rows
if "甜豆加工" in table_block and "9" in table_block:
    print("✅ PASS: Last row correctly included in table block")
else:
    print("❌ FAIL: Last row missing from table block")

# Check before/after separation
if "填單者" in before and "辨識不完全" in after:
    print("✅ PASS: Header/footer correctly separated")
else:
    print("❌ FAIL: Header/footer separation issue")