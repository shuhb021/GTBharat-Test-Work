import os
from distutils.ccompiler import new_compiler
import sys

def build_dll():
    compiler = new_compiler()
    
    src_files = [
        r'cpp_engine\src\variance_calc.cpp',
        r'cpp_engine\src\ratio_engine.cpp',
        r'cpp_engine\src\data_processor.cpp'
    ]
    
    include_dirs = [r'cpp_engine\src']
    
    print("Compiling C++ files...")
    # Compile objects
    objects = compiler.compile(
        src_files, 
        include_dirs=include_dirs,
        extra_preargs=['/std:c++17', '/EHsc', '/O2']
    )
    
    print("Linking DLL...")
    # Link into DLL
    compiler.link_shared_lib(
        objects, 
        'far_engine_fast',
        output_dir=r'cpp_engine\build',
        export_symbols=['compute_variances', 'compute_ratios', 'compute_ratios_from_raw', 'process_bulk_data', 'free_result', 'compute_variances_fast']
    )
    
    print("Build complete: cpp_engine/build/far_engine_fast.dll")

if __name__ == '__main__':
    build_dll()
