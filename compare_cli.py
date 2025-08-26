#!/usr/bin/env python3
"""Small CLI wrapper to run the project's comparison runner without the GUI.
Usage:
  .venv\Scripts\python.exe compare_cli.py fileA fileB [excel|pdf]
This script prints a short summary and writes a `comparison_result.txt` file with a repr of the result.
"""
import sys
import os
import traceback

try:
    from src.core.runner import run_compare
except Exception:  # helpful message if run outside repo
    print("ERROR: can't import project modules. Run this from the project root where `src` lives.")
    raise


def main(argv):
    if len(argv) < 3:
        print("Usage: python compare_cli.py <file1> <file2> [excel|pdf]")
        return 2

    a, b = argv[1], argv[2]
    file_type = argv[3] if len(argv) > 3 else None

    if not os.path.exists(a):
        print(f"File not found: {a}")
        return 3
    if not os.path.exists(b):
        print(f"File not found: {b}")
        return 3

    print(f"Running comparison: {a} vs {b} (type={file_type or 'auto'})")

    try:
        result = run_compare([a, b], file_type=file_type)
        print("Comparison finished.")
        # Write a small summary file so coworkers can send it back easily
        out_path = os.path.join(os.getcwd(), 'comparison_result.txt')
        with open(out_path, 'w', encoding='utf-8') as fh:
            fh.write('Result repr:\n')
            try:
                fh.write(repr(result))
            except Exception:
                fh.write('Could not repr() result; type: ' + str(type(result)))

        print('Wrote summary to', out_path)
        print('If the tool produced a saved file (e.g. an xlsx or pdf) the returned object or path may point to it.')
        return 0

    except Exception:
        print('Error during comparison:')
        traceback.print_exc()
        return 1


if __name__ == '__main__':
    sys.exit(main(sys.argv))
