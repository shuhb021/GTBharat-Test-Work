import sys
import traceback
from pathlib import Path

sys.path.insert(0, '.')

from src.core.data_engine import DataEngine

# Find first BS/PL file in Testing Dataset/Testing_FAR
test_dir = Path('Testing Dataset') / 'Testing_FAR'
files = list(test_dir.glob('*.xlsb')) + list(test_dir.glob('*.xlsx'))
if not files:
    print('No test files found in', test_dir)
    sys.exit(1)

bs_file = str(files[0])
pl_file = str(files[0])
print('Using BS/PL file:', bs_file)

form_data = {
    'bs_file': bs_file,
    'pl_file': pl_file,
    'notes_files': [],
    'rounding_unit': 'Lakhs',
    'round_off': True,
    'cy_from_date': '2025-04-01',
    'cy_to_date': '2025-03-31',
    'py_from_date': '2024-04-01',
    'py_to_date': '2024-03-31',
    'auto_annualize': True,
    'report_type': 'FAR',
    'client_name': ''
}

engine = DataEngine()
try:
    payload = engine.process_far_data(form_data, progress_callback=lambda p,m: print(p,m))
    print('Payload keys:', list(payload.keys()))
except Exception as e:
    traceback.print_exc()
    print('Error:', e)
    sys.exit(1)

print('Done')
