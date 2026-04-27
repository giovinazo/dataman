# -*- coding: utf-8 -*-
"""DataMan 모듈 OS 자동 분기 어댑터.

macOS는 dataman_mac (textutil + LibreOffice 폴백) 사용,
그 외(Windows·Linux)는 dataman (pywin32 COM) 사용.
"""
import os
import platform
import sys

_PARENT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _PARENT not in sys.path:
    sys.path.insert(0, _PARENT)

if platform.system() == "Darwin":
    import dataman_mac as _dm
else:
    import dataman as _dm

# 추출 함수 재노출
extract_pdf = _dm.extract_pdf
extract_hwp_direct = _dm.extract_hwp_direct
extract_hwpx_direct = _dm.extract_hwpx_direct
extract_txt = _dm.extract_txt
extract_docx = _dm.extract_docx
extract_doc = _dm.extract_doc
extract_text = _dm.extract_text

# 유틸리티 재노출
get_file_type = _dm.get_file_type
SUPPORTED_EXTENSIONS = _dm.SUPPORTED_EXTENSIONS
parse_filename_metadata = _dm.parse_filename_metadata
clean_text = _dm.clean_text
_build_entry = _dm._build_entry
_extract_one = _dm._extract_one
run = _dm.run

# HWPX 헬퍼
_hwpx_local = _dm._hwpx_local
_hwpx_paragraph_text = _dm._hwpx_paragraph_text
_hwpx_table_to_markdown = _dm._hwpx_table_to_markdown

# 환경 정보
OCR_AVAILABLE = _dm.OCR_AVAILABLE
PLATFORM = platform.system()
