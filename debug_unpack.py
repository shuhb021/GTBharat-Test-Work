#!/usr/bin/env python
"""Debug script to find the unpacking error"""
import sys
import traceback
import logging
from pathlib import Path

logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

# Find test files
test_dir = Path(r"e:\GT WORK @\GT Work\far-automation\Testing Dataset\Testing_FAR")
bs_file = list(test_dir.glob("*.xlsb"))[0]
pl_file = list(test_dir.glob("*.xlsb"))[0]  # same file for now

print(f"Testing with BS: {bs_file}")
print(f"Testing with PL: {pl_file}")

try:
    from src.core.excel_parser import parse_balance_sheet, parse_profit_loss
    
    print("\n=== Parsing Balance Sheet ===")
    bs_result = parse_balance_sheet(str(bs_file))
    print(f"✓ BS parsed: {type(bs_result)}, keys={list(bs_result.keys()) if isinstance(bs_result, dict) else 'N/A'}")
    
    print("\n=== Parsing Profit & Loss ===")
    pl_result = parse_profit_loss(str(pl_file))
    print(f"✓ PL parsed: {type(pl_result)}, keys={list(pl_result.keys()) if isinstance(pl_result, dict) else 'N/A'}")
    
except Exception as e:
    print(f"\n❌ ERROR: {type(e).__name__}: {e}")
    traceback.print_exc()
    sys.exit(1)

print("\n✅ All parsing successful!")
