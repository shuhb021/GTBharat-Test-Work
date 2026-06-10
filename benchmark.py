import time
import tracemalloc
import logging
from src.core.data_engine import DataEngine
from src.core.cache_manager import cache_mgr, clear_workbook_cache

logging.basicConfig(level=logging.INFO)

def run_benchmark(files_dict, use_cache=False):
    """
    Run the FarAutomation pipeline and measure performance.
    files_dict must be formatted as:
    {'bs_file': path, 'pl_file': path, 'notes_files': [path], 'rounding_unit': 'Lakhs'}
    """
    
    if not use_cache:
        # Clear existing disk cache to simulate fresh upload
        for f in cache_mgr.cache_dir.glob("*.pkl"):
            f.unlink()
            
    clear_workbook_cache()
    
    engine = DataEngine()
    
    tracemalloc.start()
    start_time = time.perf_counter()
    
    # Run the pipeline
    payload = engine.process_far_data(files_dict)
    
    end_time = time.perf_counter()
    current_mem, peak_mem = tracemalloc.get_traced_memory()
    tracemalloc.stop()
    
    duration = end_time - start_time
    peak_mb = peak_mem / 1024 / 1024
    
    print(f"--- Benchmark Results {'(CACHED)' if use_cache else '(FRESH)'} ---")
    print(f"Total Runtime: {duration:.4f} seconds")
    print(f"Peak Memory: {peak_mb:.2f} MB")
    print(f"Extracted BS rows: {len(payload['bs_data'])}")
    print(f"Computed Ratios: {len(payload['ratios'])}")
    print("-" * 40)
    
    return duration, peak_mb

if __name__ == "__main__":
    import sys
    # Example usage: python benchmark.py <path_to_excel_file>
    if len(sys.argv) > 1:
        test_file = sys.argv[1]
        test_data = {
            'bs_file': test_file,
            'pl_file': test_file,
            'notes_files': [],
            'rounding_unit': 'Lakhs',
            'notes_required': False,
            'round_off': True
        }
        
        print("Running Cold Start Benchmark...")
        run_benchmark(test_data, use_cache=False)
        
        print("\nRunning Warm Cache Benchmark...")
        run_benchmark(test_data, use_cache=True)
    else:
        print("Please provide an Excel file path for benchmarking.")
