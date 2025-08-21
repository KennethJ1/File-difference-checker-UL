"""Runner module to dispatch comparisons by file type.
Provides run_compare(file_list, options) which auto-detects type by extension and calls
compare_excel_files or compare_pdfs with normalized options and progress callback.
"""
from typing import List, Dict, Any, Tuple
import os

from .excel_diff import compare_excel_files, save_comparison_result
from .pdf_diff import compare_pdfs


EXT_TO_TYPE = {
    ".xlsx": "excel",
    ".xls": "excel",
    ".pdf": "pdf",
}


def _detect_type(file_list: List[str], override: str = None) -> str:
    if override:
        return override.lower()
    # prefer checking first provided file
    _, ext = os.path.splitext(file_list[0])
    return EXT_TO_TYPE.get(ext.lower(), "unknown")


def run_compare(file_list: List[str], options: Dict[str, Any] = None, file_type: str = None, progress_cb=None, return_meta: bool = False) -> Any:
    """Run comparison for provided files.
    - file_list: list of paths (first two used)
    - options: engine-specific options dict
    - file_type: optional override: 'excel' or 'pdf'
    - progress_cb: callable(percent:int)
    - return_meta: if True return (result, meta)

    Returns: path or Workbook (for excel when not saved), or (result, meta) when return_meta True.
    """
    opts = dict(options) if options else {}
    if progress_cb and callable(progress_cb):
        opts["progress_cb"] = progress_cb
    opts["return_meta"] = return_meta

    detected = _detect_type(file_list, override=file_type)

    if detected == "excel":
        # ensure compare_excel_files expects list of two paths
        res = compare_excel_files(file_list[:2], options=opts)
    elif detected == "pdf":
        res = compare_pdfs(file_list[:2], options=opts)
    else:
        raise ValueError("Unsupported file type for comparison. Provide file_type override.")

    return res
