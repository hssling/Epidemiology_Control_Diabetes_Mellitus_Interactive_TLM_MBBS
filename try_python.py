#!/usr/bin/env python3
"""
Simple Python detection and environment test script
"""
import sys
import os
import platform

def main():
    print("=== Python Environment Detection ===")
    print(f"Python Version: {sys.version}")
    print(f"Platform: {platform.platform()}")
    print(f"Architecture: {platform.architecture()}")
    print(f"Python Executable: {sys.executable}")
    print(f"Current Working Directory: {os.getcwd()}")

    # Check for required libraries
    required_libs = ['pptx', 'pandas', 'matplotlib', 'numpy', 'seaborn']
    print("\n=== Library Availability Check ===")

    for lib in required_libs:
        try:
            __import__(lib)
            print(f"✅ {lib}: Available")
        except ImportError:
            print(f"❌ {lib}: Not found")

    print("\n=== Environment Summary ===")
    print("This script successfully ran Python!")
    print("To generate the PowerPoint presentation:")
    print("1. Install missing libraries: pip install python-pptx pandas matplotlib numpy seaborn")
    print("2. Navigate to TLM_Diabetes_Mellitus folder")
    print("3. Run: python create_improved_pptx_with_npcdcs.py")

if __name__ == "__main__":
    main()
