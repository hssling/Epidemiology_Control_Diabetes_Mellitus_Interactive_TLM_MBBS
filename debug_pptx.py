#!/usr/bin/env python3
"""
Debug script to check PowerPoint creation
"""

import os
print(f'Current directory: {os.getcwd()}')

try:
    # Check existing PPTX files
    print('Files before running main script:')
    pptx_files_before = [f for f in os.listdir('.') if f.endswith('.pptx')]
    for f in pptx_files_before:
        print(f'  {f}')

    # Try to import the main functions
    print('\nTrying to import main functions...')
    from create_improved_pptx_with_npcdcs import main

    # Run main
    print('Running main function...')
    main()

    # Check files after
    print('\nFiles after running main script:')
    pptx_files_after = [f for f in os.listdir('.') if f.endswith('.pptx')]
    for f in pptx_files_after:
        print(f'  {f}')

    # Check for new files
    new_files = set(pptx_files_after) - set(pptx_files_before)
    if new_files:
        print(f'\nNew PPTX files created: {new_files}')
        for file in new_files:
            size = os.path.getsize(file)
            print(f'  {file}: {size} bytes')
    else:
        print('\nNo new PPTX files created')

except Exception as e:
    import traceback
    print(f'Error occurred: {e}')
    traceback.print_exc()
