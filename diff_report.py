# -*- coding: utf-8 -*-
"""
Created on Sat Sep  5 11:29:51 2026

@author: jmc53
"""

import difflib
import argparse
import sys

def generate_diff(file1_path, file2_path, output_path):
    try:
        # Read the lines from both files
        with open(file1_path, 'r', encoding='utf-8') as f1, \
             open(file2_path, 'r', encoding='utf-8') as f2:
            
            # Generate a unified diff
            diff = difflib.unified_diff(
                f1.readlines(), 
                f2.readlines(), 
                fromfile=file1_path, 
                tofile=file2_path
            )
            
        # Write the diff to the output file
        with open(output_path, 'w', encoding='utf-8') as out_file:
            out_file.writelines(diff)
            
        print(f"Success! Diff saved to: {output_path}")

    except FileNotFoundError as e:
        print(f"Error: {e}")
        sys.exit(1)
    except Exception as e:
        print(f"An error occurred: {e}")
        sys.exit(1)

# =============================================================================
# if __name__ == "__main__":
#     # Set up command-line arguments
#     parser = argparse.ArgumentParser(description="Generate a unified diff file between two text files.")
#     parser.add_argument("original", help="Path to the original file")
#     parser.add_argument("modified", help="Path to the modified file")
#     parser.add_argument("-o", "--output", default="changes.diff", help="Output diff file name (default: changes.diff)")
#     
#     args = parser.parse_args()
#     
#     generate_diff(args.original, args.modified, args.output)
# =============================================================================

file1 = "O:\Data\Programming\Qwen\Diagram_Tool\R3a\Qwen_html_20260903_1ixwkzv5d - R3a.html"
file2 = "O:\Data\Programming\Qwen\Diagram_Tool\R4\Qwen_html_20260903_1ixwkzv5d.html"

generate_diff(file1, file2, "O:\Data\Programming\Qwen\Diagram_Tool\R4\Qwen_html_20260903_1ixwkzv5d.diff")