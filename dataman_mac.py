# -*- coding: utf-8 -*-
# ──────────────────────────────────────────────────
# 프로그램명: DataMan for Mac (데이터맨 - 문서 텍스트 추출 도구 macOS 버전)
# 버전: 1.1-mac
# 저작자: 허재영
# 창작연도: 2025
# 최종 수정일: 2026-04-27
# Copyright (c) 2025-2026 허재영. All rights reserved.
# v1.1 변경: HWPX 표 마크다운 보존 / JSONL 스트리밍 출력 / 중단 시 부분 결과 저장
# ──────────────────────────────────────────────────
"""
DataMan for Mac - 문서 텍스트 추출 도구 (macOS)
====================================================================
HWP/HWPX/PDF/TXT/DOC/DOCX -> JSON 통합 전처리 스크립트

실행:
  python dataman_mac.py                          (폴더 선택 대화상자)
  python dataman_mac.py --input "폴더경로"       (직접 지정)

처리 규칙:
  1. 선택한 폴더와 하위 모든 폴더를 포함하여 탐색한다.
  2. 추출이 어려운 파일은 에러 처리 후 다음 파일로 넘어간다.
  3. CLI 실행 시 매 파일마다 진행상황을 보고한다.
  4. 추출 결과는 JSON 1개 파일로 저장한다.
  5. 추출 완료 후 "추출로그.txt" 파일을 생성한다.

macOS 특이사항:
  - DOC 추출: textutil (내장) 또는 LibreOffice (headless) 사용
  - OCR: Tesseract (brew install tesseract) 필요
  - 단축키: Cmd+Q (종료), Cmd+O (폴더 열기), Esc (중지)
"""

import atexit
import json
import os
import re
import shutil
import struct
import subprocess
import sys
import tempfile
import threading
import time
import tkinter as tk
import argparse
import zipfile
import zlib
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from tkinter import ttk, messagebox, filedialog
from typing import Callable, Optional
from xml.etree import ElementTree

import olefile

# defusedxml이 있으면 사용, 없으면 안전한 파서 직접 구성
try:
    from defusedxml.ElementTree import fromstring as _safe_fromstring
    _HAS_DEFUSEDXML = True
except ImportError:
    _HAS_DEFUSEDXML = False


# ── 공통 유틸리티 (인라인) ─────────────────────────
# 이하 코드는 shared/ 모듈에서 추출. 단독 실행을 위해 포함.

# ── macOS 네이티브 테마 적용 ────────────────────

def apply_macos_theme(root: tk.Tk) -> ttk.Style:
    style = ttk.Style(root)
    for theme in ("aqua", "clam", "default"):
        if theme in style.theme_names():
            style.theme_use(theme)
            break
    return style


# ── BoundedText ──────────────────────────────────

class BoundedText(tk.Text):
    MAX_LINES = 10000

    def __init__(self, master=None, max_lines: int = MAX_LINES, **kwargs):
        super().__init__(master, **kwargs)
        self._max_lines = max_lines

    def insert(self, index, chars, *args):
        super().insert(index, chars, *args)
        self._trim_lines()

    def _trim_lines(self):
        line_count = int(self.index("end-1c").split(".")[0])
        if line_count > self._max_lines:
            overflow = line_count - self._max_lines
            self.delete("1.0", f"{overflow + 1}.0")


# ── ETACalculator ────────────────────────────────

class ETACalculator:
    def __init__(self, total: int):
        self._total = max(total, 1)
        self._start_time = time.time()

    def update(self, current: int) -> str:
        if current <= 0:
            return ""
        elapsed = time.time() - self._start_time
        rate = current / elapsed
        remaining_items = self._total - current
        if rate <= 0 or remaining_items <= 0:
            return "거의 완료"
        remaining_secs = remaining_items / rate
        return self._format_time(remaining_secs)

    @staticmethod
    def _format_time(seconds: float) -> str:
        seconds = int(seconds)
        if seconds < 5:
            return "거의 완료"
        if seconds < 60:
            return f"약 {seconds}초"
        minutes = seconds // 60
        secs = seconds % 60
        if secs == 0:
            return f"약 {minutes}분"
        return f"약 {minutes}분 {secs}초"


# ── 종료 핸들러 ──────────────────────────────────

def setup_close_handler(
    root: tk.Tk,
    cleanup_fn: Optional[Callable] = None,
    confirm_if_running: Optional[Callable[[], bool]] = None,
):
    def _on_closing():
        if confirm_if_running and confirm_if_running():
            result = messagebox.askyesno(
                "확인",
                "작업이 진행 중입니다. 프로그램을 종료하시겠습니까?"
            )
            if not result:
                return
        if cleanup_fn:
            try:
                cleanup_fn()
            except Exception:
                pass
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", _on_closing)
    if cleanup_fn:
        atexit.register(cleanup_fn)


# ── 키보드 단축키 (macOS: Command 키) ────────────

def bind_shortcuts(
    root: tk.Tk,
    on_quit: Optional[Callable] = None,
    on_open: Optional[Callable] = None,
    on_stop: Optional[Callable] = None,
):
    if on_quit:
        def _quit(event=None):
            on_quit()
        root.bind("<Command-q>", _quit)
    if on_open:
        def _open(event=None):
            on_open()
        root.bind("<Command-o>", _open)
    if on_stop:
        def _stop(event=None):
            on_stop()
        root.bind("<Escape>", _stop)


# ── 설정 저장/복원 ──────────────────────────────

SETTINGS_FILE = os.path.join(os.path.expanduser("~"), ".audit_tools.json")


def _load_all() -> dict:
    if not os.path.exists(SETTINGS_FILE):
        return {}
    try:
        with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except (json.JSONDecodeError, OSError):
        return {}


def _save_all(data: dict) -> None:
    try:
        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except OSError:
        pass


def load_settings(app_name: str) -> dict:
    all_settings = _load_all()
    return all_settings.get(app_name, {})


def save_settings(app_name: str, settings: dict) -> None:
    all_settings = _load_all()
    all_settings[app_name] = settings
    _save_all(all_settings)


# ── XML XXE 방지 ─────────────────────────────────

_DEFAULT_MAX_XML_SIZE = 50 * 1024 * 1024  # 50 MB
DEFAULT_MAX_ZIP_SIZE = 200 * 1024 * 1024  # 200 MB (HWPX Zip bomb 방지)


def safe_parse_xml(
    data: bytes,
    max_size: int = _DEFAULT_MAX_XML_SIZE,
) -> ElementTree.Element:
    if len(data) > max_size:
        raise ValueError(f"XML 크기 한도 초과: {len(data):,} 바이트 > {max_size:,} 바이트")
    if _HAS_DEFUSEDXML:
        return _safe_fromstring(data)
    parser = ElementTree.XMLParser()
    parser.feed(data)
    return parser.close()


# ── 텍스트 처리 유틸리티 ─────────────────────────

_MULTI_SPACE = re.compile(r'[ \t]+')
_MULTI_NEWLINE = re.compile(r'\n{3,}')
_CONTROL_CHARS = re.compile(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]')
_SURROGATES = re.compile(r'[\ud800-\udfff]')


def clean_text(text: str) -> tuple[str, bool]:
    if not text:
        return text, False
    original = text
    text = _CONTROL_CHARS.sub('', text)
    text = _SURROGATES.sub('', text)
    text = _MULTI_SPACE.sub(' ', text)
    text = _MULTI_NEWLINE.sub('\n\n', text)
    text = text.strip()
    changed = text != original
    return text, changed


# ── 공통 유틸리티 끝 ──────────────────────────────

# ============================================================
# 상수
# ============================================================

APP_NAME = "dataman"

# HWP 바이너리 레코드 태그
HWPTAG_PARA_TEXT = 67

# HWP 헤더 압축 플래그 오프셋
HWP_HEADER_COMPRESSED_OFFSET = 36

SUPPORTED_EXTENSIONS = [".hwp", ".hwpx", ".pdf", ".txt", ".docx", ".doc"]

# ============================================================
# 라이브러리 확인
# ============================================================

try:
    import fitz  # PyMuPDF
except ImportError:
    print("오류: PyMuPDF가 설치되어 있지 않습니다. -> pip install PyMuPDF")
    sys.exit(1)

DOCX_AVAILABLE = False
try:
    import docx
    DOCX_AVAILABLE = True
except ImportError:
    pass

# ── macOS: textutil (내장 CLI) 으로 DOC 추출 ────
TEXTUTIL_AVAILABLE = False
TEXTUTIL_PATH = "/usr/bin/textutil"
if os.path.isfile(TEXTUTIL_PATH):
    TEXTUTIL_AVAILABLE = True
else:
    _textutil_found = shutil.which("textutil")
    if _textutil_found:
        TEXTUTIL_PATH = _textutil_found
        TEXTUTIL_AVAILABLE = True

# LibreOffice headless: DOC 추출 보조 fallback
LIBREOFFICE_AVAILABLE = False
LIBREOFFICE_PATH = ""
_lo_paths = [
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    "/usr/local/bin/soffice",
]
for _lp in _lo_paths:
    if os.path.isfile(_lp):
        LIBREOFFICE_PATH = _lp
        LIBREOFFICE_AVAILABLE = True
        break
if not LIBREOFFICE_AVAILABLE:
    _lo_found = shutil.which("libreoffice") or shutil.which("soffice")
    if _lo_found:
        LIBREOFFICE_PATH = _lo_found
        LIBREOFFICE_AVAILABLE = True

# ── OCR (Tesseract) ──────────────────────────────
OCR_AVAILABLE = False
OCR_UNAVAIL_REASON = ""
try:
    import pytesseract
    from PIL import Image
    import io
    _tesseract_paths = [
        "/opt/homebrew/bin/tesseract",       # Apple Silicon Homebrew
        "/usr/local/bin/tesseract",          # Intel Homebrew
        "/opt/local/bin/tesseract",          # MacPorts
    ]
    try:
        pytesseract.get_tesseract_version()
        OCR_AVAILABLE = True
    except Exception:
        for _path in _tesseract_paths:
            if os.path.isfile(_path):
                pytesseract.pytesseract.tesseract_cmd = _path
                try:
                    pytesseract.get_tesseract_version()
                    OCR_AVAILABLE = True
                    break
                except Exception:
                    continue
        if not OCR_AVAILABLE:
            OCR_UNAVAIL_REASON = "Tesseract OCR을 찾을 수 없습니다. (brew install tesseract)"
except ImportError:
    OCR_UNAVAIL_REASON = "pytesseract 또는 Pillow 미설치"


# ============================================================
# 출력 경로 설정 (스크립트와 같은 경로에 저장)
# ============================================================

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = SCRIPT_DIR
OCR_LANG = "kor+eng"
PARALLEL_WORKERS = min(os.cpu_count() or 4, 8)


def safe_print(msg, overwrite=False):
    """콘솔 유니코드 출력 (macOS UTF-8 기본)."""
    prefix = "\r" if overwrite else ""
    try:
        print(f"{prefix}{msg}", flush=True)
    except UnicodeEncodeError:
        print(f"{prefix}{msg}".encode(sys.stdout.encoding or "utf-8", errors="replace").decode(
            sys.stdout.encoding or "utf-8", errors="replace"), flush=True)


# ============================================================
# 메타데이터 추출
# ============================================================

def parse_filename_metadata(filename):
    """파일명에서 번호, 기관명, 제목을 분리"""
    name = os.path.splitext(filename)[0]
    parts = name.split('_', 2)
    result = {"file_no": "", "inst_name": "", "title": name}
    if len(parts) >= 3:
        if parts[0].isdigit():
            result["file_no"] = parts[0]
            result["inst_name"] = parts[1]
            result["title"] = parts[2]
        else:
            result["inst_name"] = parts[0]
            result["title"] = '_'.join(parts[1:])
    elif len(parts) == 2:
        if parts[0].isdigit():
            result["file_no"] = parts[0]
            result["inst_name"] = parts[1]
            result["title"] = parts[1]
        else:
            result["inst_name"] = parts[0]
            result["title"] = parts[1]
    return result


# ============================================================
# 텍스트 추출 함수들 (PDF/TXT/DOCX/DOC)
# ============================================================

def ocr_pdf_page(page, dpi=300):
    if not OCR_AVAILABLE:
        return ""
    try:
        mat = fitz.Matrix(dpi / 72, dpi / 72)
        pix = page.get_pixmap(matrix=mat)
        img = Image.open(io.BytesIO(pix.tobytes("png")))
        return pytesseract.image_to_string(img, lang=OCR_LANG)
    except (OSError, RuntimeError):
        return ""


def ocr_full_pdf(pdf_path, dpi=300):
    doc = None
    try:
        doc = fitz.open(pdf_path)
        texts = []
        for i in range(doc.page_count):
            try:
                t = ocr_pdf_page(doc[i], dpi)
                if t:
                    texts.append(t)
            except (OSError, RuntimeError):
                continue
        pages = doc.page_count
        cleaned, _ = clean_text('\n'.join(texts))
        return {"text": cleaned, "pages": pages}
    except (OSError, RuntimeError) as e:
        return {"text": "", "pages": 0, "error": str(e)}
    finally:
        if doc:
            doc.close()


def extract_pdf(path):
    doc = None
    try:
        doc = fitz.open(path)
        texts = []
        total = doc.page_count
        for i in range(total):
            try:
                t = doc[i].get_text("text")
                if t:
                    texts.append(t)
            except (OSError, RuntimeError):
                continue
        raw = '\n'.join(texts)
        cleaned, dup = clean_text(raw)
        cpp = len(cleaned) / max(total, 1)
        is_scanned = (len(cleaned) == 0) or (cpp < 50 and total >= 2)
        ocr_applied = False
        if is_scanned and OCR_AVAILABLE:
            ocr_res = ocr_full_pdf(path)
            if ocr_res.get("text"):
                cleaned = ocr_res["text"]
                cpp = len(cleaned) / max(total, 1)
                ocr_applied = True
        return {"text": cleaned, "pages": total, "chars_per_page": round(cpp, 1),
                "is_scanned": is_scanned, "ocr_applied": ocr_applied, "dup_fixed": dup}
    except (OSError, RuntimeError) as e:
        return {"text": "", "pages": 0, "chars_per_page": 0,
                "is_scanned": False, "ocr_applied": False, "dup_fixed": False, "error": str(e)}
    finally:
        if doc:
            doc.close()


def extract_txt(path):
    for enc in ["utf-8", "utf-8-sig", "cp949", "euc-kr", "utf-16", "latin-1"]:
        try:
            with open(path, "r", encoding=enc) as f:
                raw = f.read()
            cleaned, dup = clean_text(raw)
            return {"text": cleaned, "pages": raw.count('\n') + 1,
                    "chars_per_page": 0, "is_scanned": False, "ocr_applied": False, "dup_fixed": dup}
        except (UnicodeDecodeError, UnicodeError):
            continue
    return {"text": "", "pages": 0, "chars_per_page": 0,
            "is_scanned": False, "ocr_applied": False, "dup_fixed": False,
            "error": "인코딩 판별 실패"}


def extract_docx(path):
    if not DOCX_AVAILABLE:
        return {"text": "", "pages": 0, "chars_per_page": 0,
                "is_scanned": False, "ocr_applied": False, "dup_fixed": False,
                "error": "python-docx 미설치"}
    try:
        document = docx.Document(path)
        paras = [p.text for p in document.paragraphs if p.text.strip()]
        for table in document.tables:
            for row in table.rows:
                cells = [c.text.strip() for c in row.cells if c.text.strip()]
                if cells:
                    paras.append("\t".join(cells))
        cleaned, dup = clean_text('\n'.join(paras))
        return {"text": cleaned, "pages": 0,
                "chars_per_page": 0, "is_scanned": False, "ocr_applied": False, "dup_fixed": dup}
    except OSError as e:
        return {"text": "", "pages": 0, "chars_per_page": 0,
                "is_scanned": False, "ocr_applied": False, "dup_fixed": False, "error": str(e)}
    except Exception as e:
        return {"text": "", "pages": 0, "chars_per_page": 0,
                "is_scanned": False, "ocr_applied": False, "dup_fixed": False, "error": str(e)}


# ============================================================
# DOC 추출: textutil (macOS 내장) + LibreOffice fallback
# ============================================================

class TextutilExtractor:
    """macOS textutil을 사용하여 DOC 파일 텍스트 추출"""

    def __init__(self):
        self.available = TEXTUTIL_AVAILABLE
        self.path = TEXTUTIL_PATH

    def start(self):
        """textutil은 상태 없는 CLI 도구이므로 시작 불필요"""
        if not self.available:
            raise RuntimeError("textutil을 찾을 수 없습니다.")

    def extract(self, doc_path):
        """DOC에서 텍스트 추출. 반환: (텍스트, 페이지수)"""
        abs_path = os.path.abspath(doc_path)
        result = subprocess.run(
            [self.path, "-convert", "txt", "-stdout", abs_path],
            capture_output=True, timeout=60
        )
        if result.returncode != 0:
            stderr = result.stderr.decode("utf-8", errors="replace").strip()
            raise RuntimeError(f"textutil 실패: {stderr}")
        raw = result.stdout.decode("utf-8", errors="replace")
        return raw, 0

    def restart(self):
        """textutil은 상태가 없으므로 재시작 불필요"""
        pass

    def quit(self):
        """textutil은 상태가 없으므로 정리 불필요"""
        pass


def _extract_doc_libreoffice(doc_path):
    """LibreOffice 헤드리스 모드로 DOC -> TXT 변환 (fallback)"""
    abs_path = os.path.abspath(doc_path)
    with tempfile.TemporaryDirectory() as tmpdir:
        result = subprocess.run(
            [LIBREOFFICE_PATH, "--headless", "--convert-to", "txt:Text",
             "--outdir", tmpdir, abs_path],
            capture_output=True, timeout=120
        )
        if result.returncode != 0:
            stderr = result.stderr.decode("utf-8", errors="replace").strip()
            raise RuntimeError(f"LibreOffice 변환 실패: {stderr}")
        basename = os.path.splitext(os.path.basename(abs_path))[0] + ".txt"
        txt_path = os.path.join(tmpdir, basename)
        if not os.path.isfile(txt_path):
            txt_files = [f for f in os.listdir(tmpdir) if f.endswith(".txt")]
            if txt_files:
                txt_path = os.path.join(tmpdir, txt_files[0])
            else:
                raise RuntimeError("LibreOffice 변환 결과 파일을 찾을 수 없음")
        with open(txt_path, "r", encoding="utf-8") as f:
            return f.read()


def extract_doc(path, extractor=None):
    """DOC 파일 텍스트 추출 (macOS: textutil > LibreOffice 순서)"""
    # 1) textutil 시도 (extractor 사용)
    if extractor and extractor.available:
        try:
            raw, pages = extractor.extract(path)
            cleaned, dup = clean_text(raw)
            return {"text": cleaned, "pages": pages,
                    "chars_per_page": round(len(cleaned) / max(pages, 1), 1) if pages else 0,
                    "is_scanned": False, "ocr_applied": False, "dup_fixed": dup}
        except Exception as e1:
            if LIBREOFFICE_AVAILABLE:
                try:
                    raw = _extract_doc_libreoffice(path)
                    cleaned, dup = clean_text(raw)
                    return {"text": cleaned, "pages": 0,
                            "chars_per_page": 0,
                            "is_scanned": False, "ocr_applied": False, "dup_fixed": dup}
                except Exception as e2:
                    return {"text": "", "pages": 0, "chars_per_page": 0,
                            "is_scanned": False, "ocr_applied": False, "dup_fixed": False,
                            "error": f"textutil: {e1} / LibreOffice: {e2}"}
            return {"text": "", "pages": 0, "chars_per_page": 0,
                    "is_scanned": False, "ocr_applied": False, "dup_fixed": False,
                    "error": str(e1)}

    # extractor 없이 직접 시도
    if TEXTUTIL_AVAILABLE:
        try:
            abs_path = os.path.abspath(path)
            result = subprocess.run(
                [TEXTUTIL_PATH, "-convert", "txt", "-stdout", abs_path],
                capture_output=True, timeout=60
            )
            if result.returncode != 0:
                raise RuntimeError(result.stderr.decode("utf-8", errors="replace"))
            raw = result.stdout.decode("utf-8", errors="replace")
            cleaned, dup = clean_text(raw)
            return {"text": cleaned, "pages": 0,
                    "chars_per_page": 0,
                    "is_scanned": False, "ocr_applied": False, "dup_fixed": dup}
        except Exception:
            pass  # fall through to LibreOffice

    if LIBREOFFICE_AVAILABLE:
        try:
            raw = _extract_doc_libreoffice(path)
            cleaned, dup = clean_text(raw)
            return {"text": cleaned, "pages": 0,
                    "chars_per_page": 0,
                    "is_scanned": False, "ocr_applied": False, "dup_fixed": dup}
        except Exception as e:
            return {"text": "", "pages": 0, "chars_per_page": 0,
                    "is_scanned": False, "ocr_applied": False, "dup_fixed": False,
                    "error": str(e)}

    return {"text": "", "pages": 0, "chars_per_page": 0,
            "is_scanned": False, "ocr_applied": False, "dup_fixed": False,
            "error": "DOC 추출 도구 없음 (textutil/LibreOffice 필요)"}


def extract_text(path, word_extractor=None):
    """확장자에 따라 적절한 추출 함수 호출 (HWP 제외)"""
    ext = os.path.splitext(path)[1].lower()
    if ext == ".pdf":
        return extract_pdf(path)
    elif ext == ".txt":
        return extract_txt(path)
    elif ext == ".docx":
        return extract_docx(path)
    elif ext == ".doc":
        return extract_doc(path, word_extractor)
    return {"text": "", "pages": 0, "chars_per_page": 0,
            "is_scanned": False, "ocr_applied": False, "dup_fixed": False,
            "error": f"미지원 형식: {ext}"}


def extract_hwp_direct(path):
    """HWP(OLE2)에서 순수 Python으로 텍스트 추출 (olefile + zlib)"""
    ole = None
    try:
        ole = olefile.OleFileIO(path)

        # 압축 여부 확인: FileHeader offset, bit 0
        header = ole.openstream("FileHeader").read()
        is_compressed = bool(header[HWP_HEADER_COMPRESSED_OFFSET] & 1)

        paragraphs = []
        section_idx = 0
        while True:
            stream_name = f"BodyText/Section{section_idx}"
            if not ole.exists(stream_name):
                break
            raw = ole.openstream(stream_name).read()
            if is_compressed:
                try:
                    raw = zlib.decompress(raw, -15)
                except zlib.error:
                    section_idx += 1
                    continue

            # 바이너리 레코드 파싱
            pos = 0
            while pos < len(raw) - 4:
                header_val = struct.unpack_from("<I", raw, pos)[0]
                tag_id = header_val & 0x3FF
                # level = (header_val >> 10) & 0x3FF
                size = (header_val >> 20) & 0xFFF
                pos += 4
                if size == 0xFFF:
                    if pos + 4 > len(raw):
                        break
                    size = struct.unpack_from("<I", raw, pos)[0]
                    pos += 4

                if pos + size > len(raw):
                    break

                if tag_id == HWPTAG_PARA_TEXT:
                    data = raw[pos:pos + size]
                    chars = []
                    i = 0
                    while i < len(data) - 1:
                        code = struct.unpack_from("<H", data, i)[0]
                        i += 2
                        if code == 0:
                            break
                        # 제어 문자 확장 영역 (inline 제어): 8바이트 추가 건너뜀
                        if code in (1, 2, 3, 11, 12, 14, 15, 16, 17, 18, 21, 22, 23):
                            i += 14  # 확장 제어문자: 총 16바이트(코드2 + 추가14)
                            continue
                        # 탭, 줄바꿈 유지
                        if code == 9:
                            chars.append('\t')
                        elif code in (10, 13):
                            chars.append('\n')
                        elif code < 0x20:
                            continue
                        else:
                            chars.append(chr(code))
                    text = ''.join(chars).strip()
                    if text:
                        paragraphs.append(text)

                pos += size
            section_idx += 1

        text = '\n'.join(paragraphs)
        cleaned, dup_fixed = clean_text(text)
        return {
            "text": cleaned, "pages": section_idx,
            "chars_per_page": round(len(cleaned) / max(section_idx, 1), 1),
            "is_scanned": False, "ocr_applied": False, "dup_fixed": dup_fixed
        }
    except OSError as e:
        return {
            "text": "", "pages": 0, "chars_per_page": 0,
            "is_scanned": False, "ocr_applied": False, "dup_fixed": False,
            "error": f"HWP 직접 추출 실패: {e}"
        }
    except Exception as e:
        return {
            "text": "", "pages": 0, "chars_per_page": 0,
            "is_scanned": False, "ocr_applied": False, "dup_fixed": False,
            "error": f"HWP 직접 추출 실패: {e}"
        }
    finally:
        if ole:
            try:
                ole.close()
            except Exception:
                pass


def _hwpx_local(elem):
    """{namespace}tag 에서 local tag만 추출"""
    return elem.tag.split('}')[-1] if '}' in elem.tag else elem.tag


def _hwpx_paragraph_text(p_elem):
    """hp:p 안의 hp:t 텍스트만 수집. 자손 hp:tbl 내부는 표로 별도 처리되므로 제외"""
    texts = []

    def walk(node):
        local = _hwpx_local(node)
        if local == 'tbl':
            return
        if local == 't' and node.text:
            texts.append(node.text)
        for child in node:
            walk(child)

    walk(p_elem)
    return ''.join(texts)


def _hwpx_cell_text(tc_elem):
    """hp:tc 셀 안의 모든 hp:p 텍스트를 공백으로 연결.
    셀 내 nested tbl은 텍스트로 평탄화하지 않고 무시 (드문 케이스)."""
    p_texts = []

    def walk(node):
        local = _hwpx_local(node)
        if local == 'p':
            txt = _hwpx_paragraph_text(node)
            if txt:
                p_texts.append(txt)
            return  # 셀 안의 inner p 자손은 이미 모두 _hwpx_paragraph_text에서 흡수
        if local == 'tbl':
            return  # nested table 무시
        for child in node:
            walk(child)

    walk(tc_elem)
    cell = ' '.join(p_texts)
    # 마크다운 표 안전화: pipe·줄바꿈 제거
    return cell.replace('|', '\\|').replace('\n', ' ').strip()


def _hwpx_table_to_markdown(tbl_elem):
    """hp:tbl 을 마크다운 표로 변환. tr/tc 직속 자식만 사용 (중첩 표 영향 차단)."""
    rows = []
    for tr in list(tbl_elem):
        if _hwpx_local(tr) != 'tr':
            continue
        cells = []
        for tc in list(tr):
            if _hwpx_local(tc) != 'tc':
                continue
            cells.append(_hwpx_cell_text(tc))
        if cells:
            rows.append(cells)
    if not rows:
        return ''
    n_cols = max(len(r) for r in rows)
    lines = []
    header = rows[0] + [''] * (n_cols - len(rows[0]))
    lines.append('| ' + ' | '.join(header) + ' |')
    lines.append('| ' + ' | '.join(['---'] * n_cols) + ' |')
    for r in rows[1:]:
        padded = r + [''] * (n_cols - len(r))
        lines.append('| ' + ' | '.join(padded) + ' |')
    return '\n'.join(lines)


def _hwpx_walk_blocks(elem, blocks):
    """root/sec 부터 재귀 순회. p는 텍스트, tbl은 마크다운 표로 변환하여 blocks 누적.
    한/글 HWPX는 표가 <p><run><tbl> 형태로 paragraph 안에 들어있으므로,
    p 안에 tbl이 있으면 텍스트 + 표를 순서대로 모두 추출한다."""
    local = _hwpx_local(elem)
    if local == 'p':
        # p 직속 텍스트 (자손 tbl 내부는 _hwpx_paragraph_text가 자동 제외)
        txt = _hwpx_paragraph_text(elem)
        if txt:
            blocks.append(txt)
        # p 안에 tbl 자손이 있으면 표로 별도 추출
        for child in elem:
            if any(_hwpx_local(d) == 'tbl' for d in child.iter()) or _hwpx_local(child) == 'tbl':
                _hwpx_walk_blocks(child, blocks)
        return
    if local == 'tbl':
        md = _hwpx_table_to_markdown(elem)
        if md:
            blocks.append(md)
        return
    for child in elem:
        _hwpx_walk_blocks(child, blocks)


def extract_hwpx_direct(path):
    """HWPX에서 COM 없이 직접 텍스트 추출 (ZIP+XML, 표 마크다운 보존, Zip Bomb/XXE 방지)"""
    try:
        blocks = []
        table_count = 0
        with zipfile.ZipFile(path, 'r') as zf:
            # Zip bomb 방지: 전체 해제 크기 점검
            total_uncompressed = sum(info.file_size for info in zf.infolist())
            if total_uncompressed > DEFAULT_MAX_ZIP_SIZE:
                raise ValueError(
                    f"HWPX 해제 크기 한도 초과: {total_uncompressed:,} > {DEFAULT_MAX_ZIP_SIZE:,} 바이트"
                )

            section_files = sorted([
                n for n in zf.namelist()
                if 'section' in n.lower() and n.endswith('.xml')
            ])
            for section_file in section_files:
                data = zf.read(section_file)
                root = safe_parse_xml(data)
                # 섹션 내 표 개수 집계
                for sub in root.iter():
                    if _hwpx_local(sub) == 'tbl':
                        # 외곽 tbl만 카운트하기 위해 부모에 tbl이 있는지 빠르게 체크
                        # 단순화: 전체 tbl count (중첩 포함) -- 통계용이라 OK
                        table_count += 1
                _hwpx_walk_blocks(root, blocks)
        text = '\n\n'.join(blocks)
        cleaned, dup_fixed = clean_text(text)
        return {
            "text": cleaned, "pages": len(section_files),
            "chars_per_page": 0, "is_scanned": False,
            "ocr_applied": False, "dup_fixed": dup_fixed,
            "tables": table_count,
        }
    except (OSError, zipfile.BadZipFile, ValueError) as e:
        return {
            "text": "", "pages": 0, "chars_per_page": 0,
            "is_scanned": False, "ocr_applied": False, "dup_fixed": False,
            "error": f"HWPX 직접 추출 실패: {e}"
        }
    except ElementTree.ParseError as e:
        return {
            "text": "", "pages": 0, "chars_per_page": 0,
            "is_scanned": False, "ocr_applied": False, "dup_fixed": False,
            "error": f"HWPX XML 파싱 실패: {e}"
        }


def get_file_type(filename):
    ext = os.path.splitext(filename)[1].lower()
    return {".pdf": "pdf", ".txt": "txt", ".docx": "docx", ".doc": "doc",
            ".hwp": "hwp", ".hwpx": "hwpx"}.get(ext, "unknown")


def _build_entry(file_path, result, input_folder):
    """파일 1건의 추출 결과를 표준 entry dict로 변환 (JSON/JSONL 공통)"""
    filename = os.path.basename(file_path)
    rel_path = os.path.relpath(file_path, input_folder)
    parent_folder = os.path.basename(os.path.dirname(file_path))
    meta = parse_filename_metadata(filename)
    orig_type = get_file_type(filename)
    ext = os.path.splitext(file_path)[1].lower()

    if ext == ".hwpx":
        method = "hwpx_direct"
    elif ext == ".hwp":
        method = "hwp_direct"
    elif ext == ".doc":
        method = "textutil" if TEXTUTIL_AVAILABLE else "libreoffice"
    else:
        method = "ocr" if result.get("ocr_applied") else orig_type

    entry = {
        "filename": filename, "file_type": orig_type,
        "file_no": meta["file_no"], "inst_name": meta["inst_name"],
        "title": meta["title"], "text": result.get("text", ""),
        "pages": result.get("pages", 0),
        "text_length": len(result.get("text", "")),
        "extraction_method": method,
        "dup_fixed": result.get("dup_fixed", False),
        "folder": parent_folder,
        "rel_path": rel_path,
    }
    if "tables" in result:
        entry["tables"] = result["tables"]
    if result.get("error"):
        entry["error"] = result["error"]
    return entry


def _extract_one(file_path):
    """단일 파일 추출 (병렬 처리용, macOS에서는 DOC도 포함)"""
    ext = os.path.splitext(file_path)[1].lower()
    try:
        if ext == ".hwpx":
            return extract_hwpx_direct(file_path)
        elif ext == ".hwp":
            return extract_hwp_direct(file_path)
        elif ext == ".pdf":
            return extract_pdf(file_path)
        elif ext == ".txt":
            return extract_txt(file_path)
        elif ext == ".docx":
            return extract_docx(file_path)
        elif ext == ".doc":
            return extract_doc(file_path)  # textutil은 thread-safe
        return {"text": "", "pages": 0, "chars_per_page": 0,
                "is_scanned": False, "ocr_applied": False, "dup_fixed": False,
                "error": f"미지원 형식: {ext}"}
    except Exception as e:
        return {"text": "", "pages": 0, "chars_per_page": 0,
                "is_scanned": False, "ocr_applied": False, "dup_fixed": False,
                "error": str(e)}


# ============================================================
# 메인 처리
# ============================================================

def run(input_folder, output_dir=None, log_callback=None, progress_callback=None,
        done_callback=None, stop_event=None, output_format='json'):
    """통합 전처리 실행 (하위 폴더 항상 포함)
    output_dir: 결과 저장 폴더 (None이면 OUTPUT_DIR 사용)
    log_callback(msg): GUI 로그 출력
    progress_callback(done, total): 진행률
    done_callback(summary_text): 완료 시 호출
    stop_event: threading.Event -- set()하면 중단
    output_format: 'json' (단일 통합 JSON) 또는 'jsonl' (라인별 JSON, 스트리밍/중단 안전)
    """
    if stop_event is None:
        stop_event = threading.Event()

    save_dir = output_dir if output_dir else OUTPUT_DIR
    output_format = (output_format or 'json').lower()
    if output_format not in ('json', 'jsonl'):
        output_format = 'json'

    def _log(msg, overwrite=False):
        if log_callback:
            log_callback(msg)
        else:
            safe_print(msg, overwrite=overwrite)

    now = datetime.now()
    timestamp = now.strftime('%Y-%m-%d %H:%M:%S')
    ts_file = now.strftime('%Y%m%d_%H%M%S')
    folder_name = os.path.basename(os.path.abspath(input_folder))
    out_ext = 'jsonl' if output_format == 'jsonl' else 'json'
    OUTPUT_JSON = os.path.join(save_dir, f"dataman_{folder_name}_{ts_file}.{out_ext}")
    EXTRACT_LOG = os.path.join(save_dir, f"추출로그_{folder_name}_{ts_file}.txt")

    _log(f"{'='*50}")
    _log(f"  DataMan for Mac")
    _log(f"  HWP/HWPX/PDF/TXT/DOC/DOCX -> {out_ext.upper()}")
    _log(f"{'='*50}")
    _log(f"  대상 폴더 : {input_folder}")
    _log(f"  저장 경로 : {OUTPUT_JSON}")
    _log(f"  출력 형식 : {output_format} ({'스트리밍' if output_format == 'jsonl' else '단일 통합'})")
    _log(f"  추출 로그 : {EXTRACT_LOG}")
    _log(f"  OCR      : {'가능 (' + OCR_LANG + ')' if OCR_AVAILABLE else '불가 (' + OCR_UNAVAIL_REASON + ')'}")
    _doc_tool = "textutil" if TEXTUTIL_AVAILABLE else ("LibreOffice" if LIBREOFFICE_AVAILABLE else "불가")
    _log(f"  DOC 추출 : {_doc_tool}")
    _log(f"{'='*50}")

    # -- 파일 수집 (항상 재귀) --
    all_files = []
    for root, dirs, files in os.walk(input_folder):
        for f in sorted(files):
            if os.path.splitext(f)[1].lower() in SUPPORTED_EXTENSIONS:
                all_files.append(os.path.join(root, f))
    all_files.sort()

    if not all_files:
        _log(f"오류: 지원 파일이 없습니다. ({', '.join(SUPPORTED_EXTENSIONS)})")
        if done_callback:
            done_callback("지원 파일 없음")
        return

    # 형식별 집계
    type_counts = {}
    has_doc = False
    for f in all_files:
        ext = os.path.splitext(f)[1].lower()
        ft = get_file_type(os.path.basename(f)).upper()
        type_counts[ft] = type_counts.get(ft, 0) + 1
        if ext == ".doc":
            has_doc = True

    total = len(all_files)
    type_strs = [f"{ft} {cnt}개" for ft, cnt in sorted(type_counts.items())]
    _log(f"대상 파일: {total}개 ({', '.join(type_strs)})")

    # -- DOC 추출 도구 준비 --
    word_extractor = None
    if has_doc:
        if TEXTUTIL_AVAILABLE:
            word_extractor = TextutilExtractor()
            try:
                word_extractor.start()
                _log("textutil 준비 완료.")
            except Exception as e:
                _log(f"textutil 시작 실패: {e}")
                word_extractor = None
        elif LIBREOFFICE_AVAILABLE:
            _log("textutil 미사용 -- LibreOffice를 대체 사용합니다.")
        else:
            _log("경고: textutil/LibreOffice 미설치 -- DOC 파일은 건너뜁니다.")

    # cleanup 등록
    def _cleanup_extractor():
        if word_extractor:
            try:
                word_extractor.quit()
            except Exception:
                pass
    atexit.register(_cleanup_extractor)

    # -- 텍스트 추출 (병렬) --
    # macOS: textutil은 thread-safe이므로 DOC 포함 전체 병렬 처리
    parallel_indices = list(range(total))

    workers = min(PARALLEL_WORKERS, len(parallel_indices)) if parallel_indices else 1
    _log(f"텍스트 추출 시작 ({total}건, 워커 {workers}개)")
    _log(f"{'='*50}")

    extraction_results = [None] * total
    start_time = time.time()
    done_count = 0

    # ETA 계산기
    eta_calc = ETACalculator(total)

    # JSONL 모드: 라인별 즉시 기록 (스트리밍·중단 안전)
    jsonl_fp = None
    if output_format == 'jsonl':
        jsonl_fp = open(OUTPUT_JSON, 'w', encoding='utf-8')

    try:
        # 병렬 추출 (HWP/HWPX/PDF/TXT/DOCX/DOC 모두 포함)
        if parallel_indices:
            with ThreadPoolExecutor(max_workers=workers) as pool:
                future_to_idx = {
                    pool.submit(_extract_one, all_files[i]): i
                    for i in parallel_indices
                }
                for future in as_completed(future_to_idx):
                    if stop_event.is_set():
                        _log("사용자에 의해 중단되었습니다.")
                        pool.shutdown(wait=False, cancel_futures=True)
                        break
                    idx = future_to_idx[future]
                    done_count += 1
                    filename = os.path.basename(all_files[idx])
                    ft = get_file_type(filename).upper()
                    result = future.result()
                    extraction_results[idx] = result

                    # JSONL 즉시 기록 (스트리밍·부분 저장 자동)
                    if jsonl_fp is not None:
                        try:
                            entry = _build_entry(all_files[idx], result, input_folder)
                            jsonl_fp.write(json.dumps(entry, ensure_ascii=False) + '\n')
                            jsonl_fp.flush()
                        except OSError as e:
                            _log(f"JSONL 기록 실패: {e}")

                    # 진행률 표시 (ETACalculator 사용)
                    pct = done_count * 100 // total
                    eta_str = eta_calc.update(done_count)
                    if eta_str:
                        eta_str = f" {eta_str}"
                    progress = f"[{done_count}/{total} {pct}%{eta_str}]"

                    text_len = len(result.get("text", ""))
                    if result.get("error"):
                        _log(f"  {progress} [X] [{ft}] {filename[:45]}", overwrite=True)
                    elif text_len == 0:
                        _log(f"  {progress} [!] [{ft}] {filename[:45]}  텍스트 없음", overwrite=True)
                    else:
                        _log(f"  {progress} [O] [{ft}] {filename[:45]}  {text_len:,}자", overwrite=True)
                    if progress_callback:
                        progress_callback(done_count, total)
    finally:
        if jsonl_fp is not None:
            try:
                jsonl_fp.close()
            except Exception:
                pass

    # 중단 처리: 부분 결과 저장
    if stop_event.is_set():
        if output_format == 'json':
            partial_results = []
            for i, fp in enumerate(all_files):
                res = extraction_results[i]
                if res is not None:
                    partial_results.append(_build_entry(fp, res, input_folder))
            if partial_results:
                partial_path = os.path.join(
                    save_dir, f"dataman_partial_{folder_name}_{ts_file}.json"
                )
                try:
                    with open(partial_path, 'w', encoding='utf-8') as f:
                        json.dump(partial_results, f, ensure_ascii=False, indent=2)
                    _log(f"부분 결과 저장: {partial_path} ({len(partial_results)}건)")
                except OSError as e:
                    _log(f"부분 결과 저장 실패: {e}")
        else:
            _log(f"부분 결과 (JSONL): {OUTPUT_JSON} ({done_count}건 기록)")
        if done_callback:
            done_callback("사용자 중단")
        return

    # 결과 조립 (원래 파일 순서 유지)
    results = []
    log_lines = []
    stats = {
        "success": 0, "empty": 0, "scanned": 0,
        "ocr_ok": 0, "ocr_fail": 0, "error": 0,
        "dup_fixed": 0, "total_chars": 0,
    }
    for idx, file_path in enumerate(all_files):
        rel_path = os.path.relpath(file_path, input_folder)
        result = extraction_results[idx]

        if result is None:
            # 중단으로 미처리된 파일
            result = {"text": "", "pages": 0, "chars_per_page": 0,
                      "is_scanned": False, "ocr_applied": False, "dup_fixed": False,
                      "error": "중단됨"}

        entry = _build_entry(file_path, result, input_folder)
        results.append(entry)

        text_len = entry["text_length"]
        stats["total_chars"] += text_len
        if result.get("dup_fixed"):
            stats["dup_fixed"] += 1

        # 로그 기록 (파일 순서)
        if result.get("error"):
            stats["error"] += 1
            log_lines.append(f"[X] {rel_path} -- {result['error']}")
        elif result.get("ocr_applied"):
            stats["ocr_ok"] += 1
            extra = " [중복보정]" if result.get("dup_fixed") else ""
            log_lines.append(f"[OCR] {rel_path} -- {text_len:,}자{extra}")
        elif result.get("is_scanned") and text_len == 0:
            if OCR_AVAILABLE:
                stats["ocr_fail"] += 1
                log_lines.append(f"[!] {rel_path} -- OCR 실패")
            else:
                stats["scanned"] += 1
                log_lines.append(f"[!] {rel_path} -- 스캔본 (OCR 미설치)")
        elif text_len == 0:
            stats["empty"] += 1
            log_lines.append(f"[!] {rel_path} -- 텍스트 없음")
        else:
            stats["success"] += 1
            extra = " [중복보정]" if result.get("dup_fixed") else ""
            log_lines.append(f"[O] {rel_path} -- {text_len:,}자{extra}")

    # -- extractor 종료 --
    if word_extractor:
        try:
            word_extractor.quit()
        except Exception:
            pass

    # -- 결과 저장 (JSONL은 이미 라인별 기록 완료) --
    if output_format == 'json':
        with open(OUTPUT_JSON, 'w', encoding='utf-8') as f:
            json.dump(results, f, ensure_ascii=False, indent=2)

    # -- 최종 집계 --
    elapsed = time.time() - start_time
    file_size = os.path.getsize(OUTPUT_JSON) / 1024 / 1024

    summary_lines = [
        f"전처리 완료!",
        f"  총 파일 수       : {total}개",
        f"  텍스트 추출 성공 : {stats['success']}개",
    ]
    if stats["ocr_ok"] > 0:
        summary_lines.append(f"  OCR 추출 성공    : {stats['ocr_ok']}개")
    if stats["ocr_fail"] > 0:
        summary_lines.append(f"  OCR 추출 실패    : {stats['ocr_fail']}개")
    if stats["scanned"] > 0:
        summary_lines.append(f"  스캔본(OCR없음)  : {stats['scanned']}개")
    if stats["empty"] > 0:
        summary_lines.append(f"  텍스트 없음      : {stats['empty']}개")
    if stats["error"] > 0:
        summary_lines.append(f"  오류 발생        : {stats['error']}개")
    if stats["dup_fixed"] > 0:
        summary_lines.append(f"  글자중복 보정    : {stats['dup_fixed']}개")
    summary_lines += [
        f"  총 글자 수       : {stats['total_chars']:,}자",
        f"  평균 글자 수     : {stats['total_chars'] // max(total, 1):,}자/파일",
        f"  소요 시간        : {elapsed:.1f}초",
        f"",
        f"  JSON 저장 : {OUTPUT_JSON} ({file_size:.1f} MB)",
        f"  추출 로그 : {EXTRACT_LOG}",
    ]

    # -- 최종 출력 --
    _log(f"{'='*50}")
    for line in summary_lines:
        _log(f"  {line}" if line and not line.startswith("  ") else line)
    _log(f"{'='*50}")

    # -- 추출로그.txt 저장 --
    with open(EXTRACT_LOG, 'w', encoding='utf-8') as f:
        f.write(f"추출로그 - {timestamp}\n")
        f.write(f"대상 폴더: {input_folder}\n")
        f.write("=" * 60 + "\n\n")
        f.write("[파일별 결과]\n")
        for line in log_lines:
            f.write(line + "\n")
        f.write("\n" + "=" * 60 + "\n")
        f.write("[요약]\n")
        for line in summary_lines:
            f.write(line + "\n")

    if done_callback:
        done_callback('\n'.join(summary_lines))


# ============================================================
# GUI
# ============================================================


class DataManGUI:

    def __init__(self):
        self.root = tk.Tk()
        self.root.title("DataMan for Mac - 문서 텍스트 추출")
        self.root.geometry("680x480")
        self.root.resizable(True, True)
        self.root.minsize(520, 380)

        # threading.Event로 실행 상태 관리
        self._running_event = threading.Event()
        self._stop_event = threading.Event()

        # macOS 네이티브 테마 적용
        self._style = apply_macos_theme(self.root)

        self._build_ui()

        # 설정 복원
        self._load_settings()

        # 키보드 단축키 (macOS: Command)
        bind_shortcuts(
            self.root,
            on_quit=self._on_quit,
            on_open=self._browse,
            on_stop=self._on_stop,
        )

        # 종료 핸들러
        setup_close_handler(
            self.root,
            cleanup_fn=self._cleanup,
            confirm_if_running=lambda: self._running_event.is_set(),
        )

    def _build_ui(self):
        # -- 폴더 선택 --
        folder_frame = ttk.LabelFrame(self.root, text="대상 폴더", padding=8)
        folder_frame.pack(fill="x", padx=10, pady=(10, 4))

        self.folder_var = tk.StringVar()
        ttk.Entry(folder_frame, textvariable=self.folder_var).pack(side="left", fill="x", expand=True, padx=(0, 6))
        ttk.Button(folder_frame, text="찾아보기...", command=self._browse).pack(side="left")

        # -- 저장 위치 --
        output_frame = ttk.LabelFrame(self.root, text="저장 위치 (비우면 대상 폴더에 저장)", padding=8)
        output_frame.pack(fill="x", padx=10, pady=(4, 4))

        self.output_var = tk.StringVar()
        ttk.Entry(output_frame, textvariable=self.output_var).pack(side="left", fill="x", expand=True, padx=(0, 6))
        ttk.Button(output_frame, text="찾아보기...", command=self._browse_output).pack(side="left")

        # -- 출력 형식 --
        fmt_frame = ttk.Frame(self.root, padding=(10, 0))
        fmt_frame.pack(fill="x")
        ttk.Label(fmt_frame, text="출력 형식:").pack(side="left")
        self.format_var = tk.StringVar(value="json")
        ttk.Radiobutton(fmt_frame, text="JSON (단일 통합)", variable=self.format_var,
                        value="json").pack(side="left", padx=(8, 0))
        ttk.Radiobutton(fmt_frame, text="JSONL (라인별·중단 안전)", variable=self.format_var,
                        value="jsonl").pack(side="left", padx=(8, 0))

        # -- 버튼 + 상태 --
        ctrl_frame = ttk.Frame(self.root, padding=(10, 4))
        ctrl_frame.pack(fill="x")

        self.start_btn = ttk.Button(ctrl_frame, text="추출 시작", command=self._start)
        self.start_btn.pack(side="left")

        self.stop_btn = ttk.Button(ctrl_frame, text="중지", command=self._on_stop, state="disabled")
        self.stop_btn.pack(side="left", padx=(6, 0))

        self.status_label = ttk.Label(ctrl_frame, text="대기 중", foreground="gray")
        self.status_label.pack(side="left", padx=12)

        _doc_status = "O" if TEXTUTIL_AVAILABLE else ("LO" if LIBREOFFICE_AVAILABLE else "X")
        env_txt = f"OCR: {'O' if OCR_AVAILABLE else 'X'} | DOC: {_doc_status} | HWP/HWPX: 직접파싱"
        ttk.Label(ctrl_frame, text=env_txt, foreground="gray").pack(side="right")

        # -- 프로그레스바 --
        self.progress = ttk.Progressbar(self.root, mode="determinate", maximum=100)
        self.progress.pack(fill="x", padx=10, pady=4)

        # -- 로그 (BoundedText) --
        log_frame = ttk.LabelFrame(self.root, text="로그", padding=4)
        log_frame.pack(fill="both", expand=True, padx=10, pady=(2, 10))

        self.log_text = BoundedText(log_frame, font=("Menlo", 11), wrap="word", state="disabled")
        scrollbar = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        self.log_text.pack(fill="both", expand=True)

        self.log_text.tag_configure("ok", foreground="green")
        self.log_text.tag_configure("err", foreground="red")

        ttk.Label(self.root, text="문의: giovinazo@yahoo.co.kr",
                  foreground="gray", font=("TkDefaultFont", 9)).pack(fill=tk.X, padx=10, pady=(0, 4), anchor=tk.E)

    def _browse(self):
        folder = filedialog.askdirectory(title="전처리할 문서 폴더를 선택하세요")
        if folder:
            self.folder_var.set(folder)

    def _browse_output(self):
        folder = filedialog.askdirectory(title="추출 결과를 저장할 폴더를 선택하세요")
        if folder:
            self.output_var.set(folder)

    def _append_log(self, msg):
        self.log_text.configure(state="normal")
        tag = None
        if "[X]" in msg or "오류" in msg:
            tag = "err"
        elif "[O]" in msg or "완료" in msg:
            tag = "ok"
        self.log_text.insert("end", msg + "\n", tag)
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _log_callback(self, msg):
        self.root.after(0, self._append_log, msg)

    def _progress_callback(self, done, total):
        pct = done * 100 // max(total, 1)
        self.root.after(0, self._update_progress, pct, done, total)

    def _update_progress(self, pct, done, total):
        self.progress["value"] = pct
        self.status_label.configure(text=f"처리 중... {done}/{total} ({pct}%)", foreground="blue")

    def _done_callback(self, summary):
        self.root.after(0, self._on_done)

    def _on_done(self):
        self._running_event.clear()
        self._stop_event.clear()
        self.start_btn.configure(state="normal")
        self.stop_btn.configure(state="disabled")
        self.progress["value"] = 100
        self.status_label.configure(text="완료", foreground="green")
        self._append_log("")
        self._append_log("[주의] 추출 결과에 개인정보(성명, 직위, 징계내역 등)가 "
                         "포함되어 있을 수 있습니다.")
        self._append_log("[주의] 외부 AI 서비스에 전달 시 개인정보보호법 위반이 "
                         "될 수 있으므로, 반드시 익명화 처리 후 사용하십시오.")

    def _on_stop(self):
        """Escape 키 또는 중지 버튼으로 작업 중단"""
        if self._running_event.is_set():
            self._stop_event.set()
            self.status_label.configure(text="중단 요청...", foreground="orange")

    def _on_quit(self):
        """Cmd+Q 종료"""
        self.root.event_generate("<<CloseWindow>>")
        self.root.destroy()

    def _cleanup(self):
        """종료 시 정리 작업"""
        self._stop_event.set()
        self._save_settings()

    def _load_settings(self):
        """이전 설정 복원"""
        settings = load_settings(APP_NAME)
        if settings.get("last_folder"):
            self.folder_var.set(settings["last_folder"])
        if settings.get("last_output_folder"):
            self.output_var.set(settings["last_output_folder"])
        if settings.get("output_format") in ("json", "jsonl"):
            self.format_var.set(settings["output_format"])
        if settings.get("geometry"):
            try:
                self.root.geometry(settings["geometry"])
            except Exception:
                pass

    def _save_settings(self):
        """현재 설정 저장"""
        settings = {
            "last_folder": self.folder_var.get().strip(),
            "last_output_folder": self.output_var.get().strip(),
            "output_format": self.format_var.get(),
            "geometry": self.root.geometry(),
        }
        save_settings(APP_NAME, settings)

    def _start(self):
        folder = self.folder_var.get().strip()
        if not folder or not os.path.isdir(folder):
            self._append_log("오류: 유효한 폴더를 선택하세요.")
            return
        if self._running_event.is_set():
            return

        self._running_event.set()
        self._stop_event.clear()
        self.start_btn.configure(state="disabled")
        self.stop_btn.configure(state="normal")
        self.progress["value"] = 0
        self.status_label.configure(text="시작 중...", foreground="blue")

        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

        output_dir = self.output_var.get().strip()
        if not output_dir:
            output_dir = folder
        if not os.path.isdir(output_dir):
            os.makedirs(output_dir, exist_ok=True)

        output_format = self.format_var.get() or "json"
        t = threading.Thread(
            target=self._run_worker,
            args=(folder, output_dir, output_format),
            daemon=True,
        )
        t.start()

    def _run_worker(self, folder, output_dir, output_format):
        try:
            run(folder,
                output_dir=output_dir,
                log_callback=self._log_callback,
                progress_callback=self._progress_callback,
                done_callback=self._done_callback,
                stop_event=self._stop_event,
                output_format=output_format)
        except Exception as e:
            self._log_callback(f"오류 발생: {e}")
            self.root.after(0, self._on_done)

    def mainloop(self):
        self.root.mainloop()


# ============================================================
# 엔트리 포인트
# ============================================================

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="DataMan for Mac - 문서 텍스트 추출 도구 (하위 폴더 자동 포함)")
    parser.add_argument("--input", "-i", help="대상 폴더 경로 (CLI 모드)")
    parser.add_argument("--output", "-o", help="결과 출력 폴더 경로 (기본: 스크립트 위치)")
    parser.add_argument("--format", "-f", choices=["json", "jsonl"], default="json",
                        help="출력 형식 -- json: 단일 통합 / jsonl: 라인별 스트리밍·중단 안전 (기본 json)")
    parser.add_argument("--ocr-lang", default="kor+eng",
                        help="Tesseract OCR 언어 (기본: kor+eng)")
    parser.add_argument("--cli", action="store_true", help="CLI 모드 (GUI 없이)")
    args = parser.parse_args()

    # OCR 언어 설정 반영
    OCR_LANG = args.ocr_lang

    # 출력 경로 변경
    if args.output:
        OUTPUT_DIR = os.path.abspath(args.output)
        if not os.path.isdir(OUTPUT_DIR):
            os.makedirs(OUTPUT_DIR, exist_ok=True)

    if args.input or args.cli:
        # CLI 모드
        folder = args.input
        if not folder:
            try:
                root = tk.Tk()
                root.withdraw()
                folder = filedialog.askdirectory(title="전처리할 문서 폴더를 선택하세요")
                root.destroy()
                if not folder:
                    print("폴더가 선택되지 않았습니다.")
                    sys.exit(0)
            except Exception:
                print("사용법: python dataman_mac.py --input \"폴더경로\" [--format jsonl]")
                sys.exit(1)
        if not os.path.isdir(folder):
            print(f"오류: 폴더를 찾을 수 없습니다.\n경로: {folder}")
            sys.exit(1)
        run(folder, output_format=args.format)
    else:
        # GUI 모드 (기본)
        app = DataManGUI()
        app.mainloop()
