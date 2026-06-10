import os
import logging
from concurrent.futures import ProcessPoolExecutor, as_completed
import polars as pl
import duckdb

from .excel_parser import parse_balance_sheet, parse_profit_loss, parse_notes, clear_workbook_cache
from .cache_manager import CacheManager
from .cpp_bridge import compute_variances, compute_ratios_from_raw

logger = logging.getLogger(__name__)

# Initialize single caching instance
cache_mgr = CacheManager()

def process_file_task(task_type, filepath):
    """Worker function for multiprocessing."""
    try:
        if task_type == 'bs':
            return 'bs', parse_balance_sheet(filepath)
        elif task_type == 'pl':
            return 'pl', parse_profit_loss(filepath)
        elif task_type == 'notes':
            return 'notes', parse_notes(filepath)
    except Exception as e:
        logger.error(f"Error processing {task_type} in background: {e}")
        return task_type, {'error': str(e)}

class DataEngine:
    """Orchestrates Phase 3 (Polars), Phase 4 (Multiprocessing), Phase 7 (DuckDB)"""
    def __init__(self):
        self.con = duckdb.connect(database=':memory:')
        
    def process_far_data(self, form_data, progress_callback=None):
        """Main pipeline integrating cache, multiprocessing, polars, duckdb."""
        
        # Phase 6: Caching
        all_files = [form_data['bs_file'], form_data['pl_file']] + form_data['notes_files']
        cache_key = cache_mgr.generate_hash(all_files, form_data['rounding_unit'])
        
        cached_result = cache_mgr.get(cache_key)
        if cached_result:
            if progress_callback: progress_callback(100, "Loaded instantly from Cache!")
            return cached_result
            
        if progress_callback: progress_callback(10, "Extracting Excel data in parallel...")
            
        # Phase 4: Multiprocessing
        results = {'notes': []}
        tasks = [
            ('bs', form_data['bs_file']),
            ('pl', form_data['pl_file'])
        ]
        for nf in form_data['notes_files']:
            tasks.append(('notes', nf))
            
        with ProcessPoolExecutor(max_workers=os.cpu_count() or 4) as executor:
            future_to_task = {executor.submit(process_file_task, t[0], t[1]): t for t in tasks}
            for i, future in enumerate(as_completed(future_to_task)):
                task_type = future_to_task[future][0]
                try:
                    res_type, data = future.result()
                    if res_type == 'notes':
                        results['notes'].append(data)
                    else:
                        results[res_type] = data
                except Exception as e:
                    logger.error("Task generated an exception: %s", e)
                
                if progress_callback: 
                    progress_callback(10 + int((i+1)/len(tasks)*30), f"Processed {task_type.upper()}")

        if progress_callback: progress_callback(50, "Applying Polars transformations...")

        # Clear Excel parser cache (save memory)
        clear_workbook_cache()

        bs_parsed = results.get('bs')
        pl_parsed = results.get('pl')
        notes_parsed = results.get('notes', [])

        if not bs_parsed or 'data' not in bs_parsed or not pl_parsed or 'data' not in pl_parsed:
            raise ValueError("Failed to extract Balance Sheet or P&L data.")

        # Phase 3 & 7: Polars + DuckDB Operations
        rounding_factor = self._get_rounding_factor(form_data['rounding_unit'])
        
        # Convert Python Dicts to Polars LazyFrames for vectorized processing
        df_bs = pl.DataFrame(bs_parsed['data']).lazy()
        df_pl = pl.DataFrame(pl_parsed['data']).lazy()
        
        # Apply scaling and calculations vectorially
        df_bs = df_bs.with_columns([
            (pl.col('cy') / rounding_factor).alias('cy'),
            (pl.col('py') / rounding_factor).alias('py')
        ])
        
        df_pl = df_pl.with_columns([
            (pl.col('cy') / rounding_factor).alias('cy'),
            (pl.col('py') / rounding_factor).alias('py')
        ])

        # Register with DuckDB for fast analytics
        self.con.register('bs_table', df_bs.collect())
        self.con.register('pl_table', df_pl.collect())
        
        # Compute Variances via C++ Engine (fallback handles JSON if no DLL)
        # To maintain compatibility we extract the dictionaries again, which is very fast now.
        if progress_callback: progress_callback(70, "Computing Variances & Ratios...")
        
        # We process bulk via the existing C++ bridge for variance.
        bs_list = df_bs.collect().to_dicts()
        pl_list = df_pl.collect().to_dicts()
        
        bs_variances = compute_variances(bs_list)
        pl_variances = compute_variances(pl_list)
        
        ratio_results = compute_ratios_from_raw(bs_list, pl_list)
        
        final_payload = {
            'bs_data': bs_variances,
            'pl_data': pl_variances,
            'notes_data': notes_parsed, # Not modifying notes for now to save time
            'ratios': ratio_results.get('ratios', []),
            'bs_summary': ratio_results.get('bs_summary', {}),
            'pl_summary': ratio_results.get('pl_summary', {}),
            'meta': {
                'client_name': bs_parsed.get('client_name', ''),
                'cy_year': bs_parsed.get('cy_year', 2025),
                'py_year': bs_parsed.get('py_year', 2024),
                'cy_col_letter': bs_parsed.get('cy_col_letter', 'C'),
                'py_col_letter': bs_parsed.get('py_col_letter', 'D')
            }
        }
        
        # Save to Cache
        cache_mgr.set(cache_key, final_payload)
        
        if progress_callback: progress_callback(100, "Processing Complete")
        
        return final_payload

    def _get_rounding_factor(self, unit_str):
        unit = str(unit_str).lower()
        if 'lakh' in unit: return 100000.0
        if 'million' in unit: return 1000000.0
        if 'crore' in unit: return 10000000.0
        if 'thousand' in unit: return 1000.0
        return 1.0
