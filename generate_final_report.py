import sys
import os
sys.stdout.reconfigure(encoding='utf-8')

# Ensure the project root is in the path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from src.core.data_engine import DataEngine
from src.core.excel_export import export_far_workbook
from src.core.chart_generator import generate_all_charts
import tempfile
import shutil

filepath = r"e:\GT WORK @\GT Work\far-automation\Testing Dataset\Testing_FAR\1. 04 Eviden IT_FS_Mar 25 Version 29 ( In Lacs).xlsx"
output_path = r"e:\GT WORK @\GT Work\far-automation\Testing Dataset\Testing_FAR\Eviden_IT_FAR_2025-26.xlsx"

form_data = {
    'firm_name': 'Walker Chandiok & Co LLP',
    'client_name': 'Eviden IT Services Private Limited',
    'financial_year': '2024-25',
    'round_off': True,
    'rounding_unit': 'Lakhs',
    'notes_required': True,
    'bs_file': filepath,
    'pl_file': filepath,
    'notes_files': [filepath],
}

if __name__ == '__main__':
    # Clear cache to guarantee fresh generation
    shutil.rmtree('.far_cache', ignore_errors=True)

    print("Step 1: Running DataEngine to parse and reconcile values...")
    engine = DataEngine()
    d = engine.process_far_data(form_data)

    print("Step 2: Generating visual charts...")
    temp_dir = tempfile.mkdtemp(prefix='far_charts_final_')
    chart_paths = generate_all_charts(
        d.get('bs_result', []),
        d.get('pl_result', []),
        d.get('ratios', []),
        d.get('cy_year', 2025),
        d.get('py_year', 2024),
        d.get('rounding_unit', 'Lakhs'),
        temp_dir
    )

    print("Step 3: Exporting corrected FAR Excel workbook...")
    try:
        export_far_workbook(
            output_path,
            d.get('bs_data', []),
            d.get('pl_data', []),
            d.get('notes_data', {}),
            d.get('ratios', []),
            {}, 
            {},
            'Eviden IT Services Private Limited',
            'Walker Chandiok & Co LLP',
            '2024-25',
            'Lakhs',
            d.get('meta', {}).get('cy_year', 2025),
            d.get('meta', {}).get('py_year', 2024),
            chart_paths=chart_paths,
            bs_footers=d.get('bs_footers', []),
            pl_footers=d.get('pl_footers', []),
            notes_footers=d.get('notes_footers', {}),
            meta=d.get('meta', {})
        )
        print(f"\nSUCCESS: Corrected and reconciled FAR Excel file generated at:\n{output_path}")
    except PermissionError:
        fallback_path = output_path.replace(".xlsx", "_Corrected.xlsx")
        print(f"\nWARNING: Permission denied writing to {output_path} (it is probably open in Excel). Saving to fallback path:")
        export_far_workbook(
            fallback_path,
            d.get('bs_data', []),
            d.get('pl_data', []),
            d.get('notes_data', {}),
            d.get('ratios', []),
            {}, 
            {},
            'Eviden IT Services Private Limited',
            'Walker Chandiok & Co LLP',
            '2024-25',
            'Lakhs',
            d.get('meta', {}).get('cy_year', 2025),
            d.get('meta', {}).get('py_year', 2024),
            chart_paths=chart_paths,
            bs_footers=d.get('bs_footers', []),
            pl_footers=d.get('pl_footers', []),
            notes_footers=d.get('notes_footers', {}),
            meta=d.get('meta', {})
        )
        print(f"SUCCESS: Corrected and reconciled FAR Excel file generated at fallback path:\n{fallback_path}")

    # Clean up temp charts
    shutil.rmtree(temp_dir, ignore_errors=True)

