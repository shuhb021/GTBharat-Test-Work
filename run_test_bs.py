import sys
sys.path.insert(0, '.')
from src.core.excel_parser import extract_bs_summary

bs_data = [
    {'particulars': 'Total equity', 'cy': 1000, 'py': 800},
    {'particulars': 'Current liabilities', 'cy': 200, 'py': 150, 'note': '11'},
    {'particulars': 'Non current liabilities', 'cy': 500, 'py': 400, 'note': '12'},
]
print('Calling extract_bs_summary...')
summary = extract_bs_summary(bs_data)
print('Summary keys:', list(summary.keys()))
print('ltb_cy, ltb_py:', summary.get('ltb_cy'), summary.get('ltb_py'))
