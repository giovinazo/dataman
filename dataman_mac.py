# -*- coding: utf-8 -*-
# ──────────────────────────────────────────────────
# 프로그램명: DataMan for Mac (데이터맨 - 문서 텍스트 추출 도구 macOS 버전)
# 버전: 1.2-mac
# 저작자: 허재영
# 창작연도: 2025
# 최종 수정일: 2026-09-23
# Copyright (c) 2025-2026 허재영. All rights reserved.
# v1.1 변경: HWPX 표 마크다운 보존 / JSONL 스트리밍 출력 / 중단 시 부분 결과 저장
# v1.2 변경: PDF 쪽 단위 OCR 판정(스캔·깨진 글자층·글자층 부실) / macOS Vision 기본 엔진
#            (Tesseract 예비) / 작은 이미지 확대 인식 / 이미지·엑셀·hwx·zip 지원 /
#            경로 NFC·본문해시·쪽별 추출 방식 기록 / 실행 후 점검표 자동 출력
# ──────────────────────────────────────────────────
"""
DataMan for Mac - 문서 텍스트 추출 도구 (macOS)
====================================================================
HWP/HWPX/HWX/PDF/TXT/DOC/DOCX/XLSX/XLS/JPG/PNG/ZIP -> JSON 통합 전처리 스크립트

실행:
  python dataman_mac.py                          (폴더 선택 대화상자)
  python dataman_mac.py --input "폴더경로"       (직접 지정)

처리 규칙:
  1. 선택한 폴더와 하위 모든 폴더를 포함하여 탐색한다.
  2. 추출이 어려운 파일은 에러 처리 후 다음 파일로 넘어간다.
  3. CLI 실행 시 매 파일마다 진행상황을 보고한다.
  4. 추출 결과는 JSON 1개 파일로 저장한다.
  5. 추출 완료 후 "추출로그.txt"와 "점검표.md" 파일을 생성한다.

PDF 쪽 단위 OCR 판정 (v1.2):
  - 스캔 쪽: 글자층 50자 미만 + 이미지 있음 → OCR 결과로 교체
  - 깨진 글자층: garbled() 판정 → OCR 결과로 교체
  - 글자층 부실: 50자 미만(도형 글자 등) 또는 큰 이미지(30% 이상) 위 한글 100자 미만
    → OCR 한글이 글자층의 1.1배+5자를 넘을 때만 교체
  - 작은 이미지 쪽(이미지가 쪽 면적 70% 미만): 이미지 상자를 잘라 확대 인식,
    한글이 기존의 1.1배+5자를 넘을 때만 채택 (Vision 전용)
  - 절반 이상 OCR한 문서: 나머지 쪽도 OCR 한글이 1.1배+5자를 넘으면 교체

macOS 특이사항:
  - DOC 추출: textutil (내장) 또는 LibreOffice (headless) 사용
  - OCR: macOS Vision(pyobjc-framework-Vision 필요) 기본,
         없으면 Tesseract (brew install tesseract)
  - 단축키: Cmd+Q (종료), Cmd+O (폴더 열기), Esc (중지)
"""

import atexit
import hashlib
import json
import multiprocessing
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
import unicodedata
import zipfile
import zlib
from concurrent.futures import ProcessPoolExecutor, as_completed
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


def nfc(s):
    """경로·파일명 NFC 정규화 (macOS 파일명은 NFD로 들어오는 경우가 많음)"""
    return unicodedata.normalize("NFC", s) if isinstance(s, str) else s


def text_hash(text: str) -> str:
    """본문해시: 같은 본문을 묶어 집계할 때 쓰는 md5 앞 16자리"""
    return hashlib.md5(text.encode("utf-8")).hexdigest()[:16] if text else ""


# ── 공통 유틸리티 끝 ──────────────────────────────

# ============================================================
# 상수
# ============================================================

APP_NAME = "dataman"

# HWP 바이너리 레코드 태그
HWPTAG_PARA_TEXT = 67

# HWP 헤더 압축 플래그 오프셋
HWP_HEADER_COMPRESSED_OFFSET = 36

IMAGE_EXTENSIONS = [".jpg", ".jpeg", ".png"]
EXCEL_EXTENSIONS = [".xlsx", ".xlsm", ".xls"]
SUPPORTED_EXTENSIONS = ([".hwp", ".hwpx", ".hwx", ".pdf", ".txt", ".docx", ".doc"]
                        + IMAGE_EXTENSIONS + EXCEL_EXTENSIONS + [".zip"])

# PDF 쪽 단위 OCR 판정 기준 (스캔이 섞인 공문서 묶음 약 1.1만 건으로 검증한 값)
PAGE_MIN_CHARS = 50          # 쪽 글자층이 이보다 적으면 스캔·부실 후보
WEAK_IMAGE_COVERAGE = 0.3    # 큰 이미지 기준(쪽 면적 비율)
WEAK_MAX_HANGUL = 100        # 큰 이미지 위 글자층 한글이 이보다 적으면 부실 후보
ZOOM_COVERAGE = 0.7          # 이미지가 쪽 면적의 이 비율 미만이면 확대 인식 대상
OCR_DPI = 250                # 쪽 렌더링 해상도
OCR_MAX_SIDE = 5000          # 렌더링 한 변 상한(초대형 스캔 대비)
ZOOM_CLIP_SIDE = 3500        # 확대 인식 시 이미지 상자 렌더링 한 변

# zip: 임시 폴더에 풀어 내부 파일을 각각 추출 (원본 폴더에는 쓰지 않음)
ZIP_MAX_TOTAL = 4 * 1024 * 1024 * 1024   # 압축 1개당 해제 크기 상한 4GB
ZIP_MAX_DEPTH = 3                        # 중첩 zip 재귀 깊이
ZIP_DONE_PREFIX = "압축해제_"            # 옆에 이 폴더가 있으면 이미 풀린 zip으로 보고 건너뜀(auto)

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

# ── OCR (Vision 기본, Tesseract 예비) ─────────────
# Tesseract는 한글을 "세 금 계 산 서"처럼 음절마다 띄어 검색이 안 되므로 예비로만 쓴다.
VISION_AVAILABLE = False
try:
    import Vision
    import Quartz
    from Foundation import NSData
    try:
        from objc import autorelease_pool as _autorelease_pool
    except ImportError:
        from contextlib import nullcontext as _autorelease_pool
    VISION_AVAILABLE = True
except ImportError:
    pass

OPENPYXL_AVAILABLE = False
try:
    import openpyxl
    OPENPYXL_AVAILABLE = True
except ImportError:
    pass

XLRD_AVAILABLE = False
try:
    import xlrd
    XLRD_AVAILABLE = True
except ImportError:
    pass

OCR_AVAILABLE = False          # Tesseract 사용 가능 여부 (아래에서 판정 후 엔진 기준으로 재설정)
OCR_UNAVAIL_REASON = ""
try:
    import io
    from PIL import Image
    Image.MAX_IMAGE_PIXELS = 1_000_000_000   # 초대형 스캔 이미지의 DecompressionBombError 방지
except ImportError:
    Image = None
try:
    import pytesseract
    if Image is None:
        raise ImportError
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

TESSERACT_AVAILABLE = OCR_AVAILABLE
OCR_ENGINE = "none"


def set_ocr_engine(engine: str = "auto") -> str:
    """OCR 엔진 선택: auto(Vision > Tesseract) / vision / tesseract / none.
    쓸 수 없는 엔진을 고르면 auto 규칙으로 대체한다. 결정된 엔진 이름을 돌려준다."""
    global OCR_ENGINE, OCR_AVAILABLE
    engine = (engine or "auto").lower()
    if engine == "vision" and VISION_AVAILABLE:
        OCR_ENGINE = "vision"
    elif engine == "tesseract" and TESSERACT_AVAILABLE:
        OCR_ENGINE = "tesseract"
    elif engine == "none":
        OCR_ENGINE = "none"
    else:
        OCR_ENGINE = ("vision" if VISION_AVAILABLE
                      else "tesseract" if TESSERACT_AVAILABLE else "none")
    OCR_AVAILABLE = OCR_ENGINE != "none"
    return OCR_ENGINE


set_ocr_engine("auto")


def ocr_status_text() -> str:
    if OCR_ENGINE == "vision":
        return "macOS Vision" + (" (예비: Tesseract)" if TESSERACT_AVAILABLE else "")
    if OCR_ENGINE == "tesseract":
        return f"Tesseract ({OCR_LANG})" + ("" if VISION_AVAILABLE else " - Vision 미사용(pyobjc 없음)")
    return "불가 (" + (OCR_UNAVAIL_REASON or "OCR 끔") + ")"


# ============================================================
# 출력 경로 설정 (스크립트와 같은 경로에 저장)
# ============================================================

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = SCRIPT_DIR
OCR_LANG = "kor+eng"
PARALLEL_WORKERS = min(os.cpu_count() or 4, 6)   # Vision 병렬 검증치: spawn 프로세스 6개(초당 약 5쪽)


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

_HANGUL_RE = re.compile(r"[가-힣]")
_HIRAGANA_RE = re.compile(r"[぀-ゟ]")
_PUNCT = set("!\"#$%&'()*+,-./:;<=>?@[\\]^_`{|}~")


def hangul_count(text) -> int:
    return len(_HANGUL_RE.findall(text or ""))


def garbled(t) -> bool:
    """글자층이 깨졌는지 판정 (글꼴 대응표 손상으로 엉뚱한 글자가 나오는 PDF 등).
    대량 스캔 문서 재처리에서 검증한 규칙."""
    s = re.sub(r"\s+", "", t or "")
    L = len(s)
    if L < 30:
        return False
    ctrl = sum(1 for c in s if ord(c) < 0x20)
    s = "".join(c for c in s if ord(c) >= 0x20)     # 글자 사이 구분용 제어문자는 걷어내고 판정
    L = len(s)
    if L < 30:
        return ctrl > 0
    hang = len(_HANGUL_RE.findall(s))
    if ctrl and hang / L >= 0.3:                     # 한글 본문 + 제어문자 구분자 → 정상
        return False
    if ctrl / (L + ctrl) > 0.02:
        return True
    hira = len(_HIRAGANA_RE.findall(s))
    if hira >= 5 and hira / L >= 0.05:               # 히라가나가 고르게 있음 → 실제 일본어 문서
        return False
    ok = hang + sum(1 for c in s if c.isascii())
    if (L - ok) / L > 0.3:                           # 한글·영문 외 이상 문자가 많음
        return True
    if hang / L < 0.05:
        if sum(1 for c in s if c in _PUNCT) / L > 0.25:
            return True
        words = re.findall(r"[A-Za-z]{3,}", t)
        if len(words) >= 5:
            nov = sum(1 for w in words if not re.search(r"[aeiouyAEIOUY]", w))
            if nov / len(words) > 0.4:
                return True
    return False


# ── 렌더링·OCR 엔진 ──────────────────────────────

def _render_page(page, dpi=OCR_DPI, clip=None, side=None):
    """쪽(또는 clip 영역)을 렌더링. side를 주면 긴 변을 그 크기로, 어느 경우든 OCR_MAX_SIDE 이하"""
    rect = clip if clip is not None else page.rect
    longest = max(rect.width, rect.height)
    if longest <= 0:
        return None
    z = side / longest if side else dpi / 72
    if longest * z > OCR_MAX_SIDE:
        z = OCR_MAX_SIDE / longest
    return page.get_pixmap(matrix=fitz.Matrix(z, z), clip=clip, alpha=False)


def _too_small(pix) -> bool:
    """Vision은 한 변 2px 이하 이미지를 거부(Code 13)하므로 미리 건너뜀"""
    return pix is None or min(pix.width, pix.height) < 3


def _vision_observations(png_bytes):
    """Vision 문자인식 → [(x0, y0, x1, y1, 글자)] 정규화 좌표(원점 좌상단, 0~1)"""
    with _autorelease_pool():
        data = NSData.dataWithBytes_length_(png_bytes, len(png_bytes))
        src = Quartz.CGImageSourceCreateWithData(data, None)
        cg = Quartz.CGImageSourceCreateImageAtIndex(src, 0, None)
        req = Vision.VNRecognizeTextRequest.alloc().init()
        req.setRecognitionLevel_(0)                       # accurate
        req.setRecognitionLanguages_(["ko-KR", "en-US"])
        req.setUsesLanguageCorrection_(True)
        handler = Vision.VNImageRequestHandler.alloc().initWithCGImage_options_(cg, None)
        ok, err = handler.performRequests_error_([req], None)
        if not ok:
            raise RuntimeError(f"Vision: {err}")
        out = []
        for o in req.results() or []:
            c = o.topCandidates_(1)
            if not c:
                continue
            b = o.boundingBox()                           # 원점 좌하단
            out.append((b.origin.x, 1 - b.origin.y - b.size.height,
                        b.origin.x + b.size.width, 1 - b.origin.y, str(c[0].string())))
        return out


def _vision_lines(obs) -> str:
    """같은 높이의 조각을 한 줄로 묶어 위→아래, 왼→오 순서로 합침"""
    items = sorted((((y0 + y1) / 2, y1 - y0, x0, s) for x0, y0, x1, y1, s in obs),
                   key=lambda t: t[0])
    lines, cur, cy, chh = [], [], None, None
    for y, hh, x, s in items:
        if cur and abs(y - cy) > 0.5 * max(hh, chh):
            lines.append(cur)
            cur = []
        if not cur:
            cy, chh = y, hh
        cur.append((x, s))
    if cur:
        lines.append(cur)
    return "\n".join("  ".join(s for _, s in sorted(l)) for l in lines)


def ocr_pixmap(pix) -> str:
    """렌더링된 이미지 1장을 현재 OCR 엔진으로 인식"""
    if _too_small(pix):
        return ""
    if OCR_ENGINE == "vision":
        return _vision_lines(_vision_observations(pix.tobytes("png")))
    if OCR_ENGINE == "tesseract":
        img = Image.open(io.BytesIO(pix.tobytes("png")))
        return pytesseract.image_to_string(img, lang=OCR_LANG)
    return ""


def ocr_pdf_page(page, dpi=OCR_DPI):
    """PDF 쪽 1개 OCR (오류는 호출한 쪽에서 처리)"""
    if not OCR_AVAILABLE:
        return ""
    return ocr_pixmap(_render_page(page, dpi))


# ── 작은 이미지 확대 인식 (Vision 전용) ──────────

def _image_coverages(page):
    """쪽 안 이미지들이 쪽 면적에서 차지하는 비율 목록. 교집합은 (r & b)로 구한다
    (Rect.intersect()는 원래 사각형을 바꾸므로 쓰지 않는다)."""
    A = page.rect.get_area()
    if A <= 0:
        return []
    out = []
    for info in page.get_image_info():
        r = fitz.Rect(info["bbox"]) & page.rect
        if not r.is_empty:
            out.append(r.get_area() / A)
    return out


def _is_small_image_page(page) -> bool:
    cov = [c for c in _image_coverages(page) if c > 0.02]
    return bool(cov) and max(cov) < ZOOM_COVERAGE


def _zoom_text(page, page_obs, min_frac=0.02) -> str:
    """이미지 상자는 잘라 확대해 따로 읽고, 상자 밖 글자는 쪽 전체 인식(page_obs)에서 가져와
    위→아래로 합친다."""
    A = page.rect.get_area()
    boxes = []
    for info in page.get_image_info():
        r = fitz.Rect(info["bbox"]) & page.rect
        if r.is_empty or r.get_area() < min_frac * A or min(r.width, r.height) < 20:
            continue
        if any((r & b).get_area() > 0.8 * r.get_area() for b in boxes):
            continue
        boxes.append(r)
    W, Hh = page.rect.width, page.rect.height
    blocks = []                                           # (top, left, text)
    for x0, y0, x1, y1, s in page_obs:
        cx, cy = (x0 + x1) / 2 * W, (y0 + y1) / 2 * Hh
        if any(b.contains(fitz.Point(cx, cy)) for b in boxes):
            continue
        blocks.append((y0 * Hh, x0 * W, s))
    for b in boxes:
        try:
            cp = _render_page(page, clip=b, side=ZOOM_CLIP_SIDE)
            t = "" if _too_small(cp) else _vision_lines(_vision_observations(cp.tobytes("png")))
        except RuntimeError:
            t = ""
        if t.strip():
            blocks.append((b.y0, b.x0, t))
    blocks.sort(key=lambda t: (round(t[0] / 6), t[1]))    # 약 6pt 단위로 같은 줄 판정
    lines, cur, cy = [], [], None
    for y, x, s in blocks:
        key = round(y / 6)
        if cur and key != cy:
            lines.append("  ".join(cur))
            cur = []
        cy = key
        cur.append(s)
    if cur:
        lines.append("  ".join(cur))
    return "\n".join(lines)


# ── PDF 쪽 단위 판정·추출 ────────────────────────

def _page_kind(page, text):
    """OCR 대상 쪽 판정: 'scan'(스캔) / 'garbled'(깨진 글자층) / 'weak'(글자층 부실) / None"""
    L = len(text.strip())
    has_img = bool(page.get_images())
    if L < PAGE_MIN_CHARS and has_img:
        return "scan"
    if garbled(text):
        return "garbled"
    if L < PAGE_MIN_CHARS:
        # 이미지 없이 도형으로 그린 글자 등. 완전 백지는 건너뜀
        return "weak" if (L > 0 or page.get_drawings()) else None
    if (has_img and hangul_count(text) < WEAK_MAX_HANGUL
            and max(_image_coverages(page) or [0]) >= WEAK_IMAGE_COVERAGE):
        return "weak"
    return None


def _ocr_page_best(page, text, kind):
    """판정된 쪽을 OCR해 채택할 본문을 돌려줌. 반환: (본문 또는 None(글자층 유지), 확대 인식 여부)"""
    pix = _render_page(page)
    if _too_small(pix):
        return None, False
    obs = None
    if OCR_ENGINE == "vision":
        obs = _vision_observations(pix.tobytes("png"))
        a = _vision_lines(obs)
    else:
        a = ocr_pixmap(pix)
    if kind == "weak":      # 글자층이 있는 쪽은 OCR이 확실히 나을 때만 교체
        adopt = (hangul_count(a) > hangul_count(text) * 1.1 + 5
                 or (not text.strip() and bool(a.strip())))
    else:
        adopt = bool(a.strip())
    best = a if adopt else None
    zoomed = False
    if obs is not None and _is_small_image_page(page):
        b = _zoom_text(page, obs)
        ref = best if best is not None else text
        if hangul_count(b) > hangul_count(ref) * 1.1 + 5:
            best, zoomed = b, True
    return best, zoomed


def _empty_result(error=None, **extra):
    r = {"text": "", "pages": 0, "chars_per_page": 0,
         "is_scanned": False, "ocr_applied": False, "dup_fixed": False}
    r.update(extra)
    if error:
        r["error"] = error
    return r


def extract_pdf(path):
    """PDF 추출: 쪽마다 글자층을 읽고, 스캔·깨짐·부실 쪽만 OCR로 보완"""
    doc = None
    try:
        doc = fitz.open(path)
        total = doc.page_count
        texts = [""] * total
        ocr_pages, zoom_pages, garbled_pages, cand_pages, ocr_errors = [], [], [], [], []

        def _try(i, page, kind):
            try:
                best, zoomed = _ocr_page_best(page, texts[i], kind)
            except Exception as e:          # 쪽 1개 실패는 기록만 하고 계속
                ocr_errors.append(f"p{i + 1}: {str(e)[:80]}")
                return
            if best is not None:
                texts[i] = best
                ocr_pages.append(i + 1)
                if zoomed:
                    zoom_pages.append(i + 1)

        for i in range(total):
            try:
                page = doc[i]
                texts[i] = page.get_text("text") or ""
                kind = _page_kind(page, texts[i])
            except (OSError, RuntimeError, ValueError) as e:
                ocr_errors.append(f"p{i + 1}: {str(e)[:80]}")
                continue
            if kind:
                cand_pages.append(i + 1)
                if kind == "garbled":
                    garbled_pages.append(i + 1)
                if OCR_AVAILABLE:
                    _try(i, page, kind)
        # 절반 이상 OCR한 스캔 위주 문서: 나머지 쪽도 OCR이 확실히 나을 때만 교체
        # (글자 일부만 글자층에 있고 나머지는 도형으로 그린 쪽 대비)
        if OCR_AVAILABLE and total >= 2 and len(ocr_pages) * 2 >= total:
            for i in range(total):
                if i + 1 in cand_pages or not texts[i].strip():
                    continue
                try:
                    page = doc[i]
                except (OSError, RuntimeError, ValueError):
                    continue
                cand_pages.append(i + 1)
                _try(i, page, "weak")
            cand_pages.sort()
            ocr_pages.sort()
            zoom_pages.sort()
        cleaned, dup = clean_text('\n'.join(t for t in texts if t))
        cpp = len(cleaned) / max(total, 1)
        tag = "vision" if OCR_ENGINE == "vision" else "ocr"
        if not ocr_pages:
            method = "pdf"
        elif len(ocr_pages) == total:
            method = tag
        else:
            method = f"pdf+{tag}"
        r = {"text": cleaned, "pages": total, "chars_per_page": round(cpp, 1),
             "is_scanned": bool(cand_pages), "ocr_applied": bool(ocr_pages), "dup_fixed": dup,
             "method": method, "scan_pages": cand_pages, "garbled_pages": garbled_pages}
        if cand_pages:
            r["ocr_engine"] = OCR_ENGINE
        if ocr_pages:
            r["vision_pages" if OCR_ENGINE == "vision" else "ocr_pages"] = ocr_pages
        if zoom_pages:
            r["zoom_pages"] = zoom_pages
        if ocr_errors:
            r["ocr_errors"] = ocr_errors
        return r
    except (OSError, RuntimeError, ValueError) as e:
        return _empty_result(str(e))
    finally:
        if doc:
            doc.close()


def ocr_full_pdf(pdf_path, dpi=OCR_DPI):
    """PDF 전 쪽 OCR (하위 호환용)"""
    doc = None
    try:
        doc = fitz.open(pdf_path)
        texts = []
        for i in range(doc.page_count):
            try:
                t = ocr_pdf_page(doc[i], dpi)
                if t:
                    texts.append(t)
            except Exception:
                continue
        cleaned, _ = clean_text('\n'.join(texts))
        return {"text": cleaned, "pages": doc.page_count}
    except (OSError, RuntimeError) as e:
        return {"text": "", "pages": 0, "error": str(e)}
    finally:
        if doc:
            doc.close()


# ── 이미지·엑셀·hwx·zip ─────────────────────────

def extract_image(path):
    """jpg·png: 원본 해상도(한 변 OCR_MAX_SIDE 이하)로 렌더링해 OCR"""
    if not OCR_AVAILABLE:
        return _empty_result("OCR 엔진 없음 (이미지 파일)")
    doc = None
    try:
        doc = fitz.open(path)
        page = doc[0]
        info = page.get_image_info()
        native = max(info[0]["width"], info[0]["height"]) if info else max(page.rect.width, page.rect.height)
        pix = _render_page(page, side=min(native, OCR_MAX_SIDE))
        if _too_small(pix):
            return _empty_result("이미지가 너무 작음 (한 변 2px 이하)", pages=1)
        cleaned, dup = clean_text(ocr_pixmap(pix))
        tag = "vision" if OCR_ENGINE == "vision" else "ocr"
        return {"text": cleaned, "pages": 1, "chars_per_page": len(cleaned),
                "is_scanned": True, "ocr_applied": bool(cleaned), "dup_fixed": dup,
                "method": tag, "ocr_engine": OCR_ENGINE,
                ("vision_pages" if tag == "vision" else "ocr_pages"): [1] if cleaned else []}
    except Exception as e:
        return _empty_result(f"이미지 OCR 실패: {str(e)[:120]}")
    finally:
        if doc:
            doc.close()


def _excel_rows_text(rows, drop_tail=("",)):
    out = []
    for row in rows:
        cells = ["" if c is None else str(c).replace("\n", " ").strip() for c in row]
        while cells and cells[-1] in drop_tail:
            cells.pop()
        if cells:
            out.append(" | ".join(cells))
    return out


def extract_excel(path):
    """xlsx·xlsm(openpyxl)·xls(xlrd): 시트별로 '## [시트] 이름' 머리 + 행마다 ' | ' 구분"""
    ext = os.path.splitext(path)[1].lower()
    try:
        parts, sheets = [], 0
        if ext in (".xlsx", ".xlsm"):
            if not OPENPYXL_AVAILABLE:
                return _empty_result("openpyxl 미설치")
            wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
            try:
                for ws in wb.worksheets:
                    sheets += 1
                    parts.append(f"## [시트] {ws.title}")
                    parts += _excel_rows_text(ws.iter_rows(values_only=True))
            finally:
                wb.close()
        else:
            if not XLRD_AVAILABLE:
                return _empty_result("xlrd 미설치")
            bk = xlrd.open_workbook(path)
            for ws in bk.sheets():
                sheets += 1
                parts.append(f"## [시트] {ws.name}")
                parts += _excel_rows_text(([c.value for c in ws.row(r)] for r in range(ws.nrows)),
                                          drop_tail=("", "0.0"))
        cleaned, dup = clean_text("\n".join(parts))
        return {"text": cleaned, "pages": sheets, "chars_per_page": 0, "is_scanned": False,
                "ocr_applied": False, "dup_fixed": dup, "method": "excel", "tables": sheets}
    except Exception as e:
        return _empty_result(f"엑셀 추출 실패: {str(e)[:120]}")


_OLE_SIG = bytes.fromhex("D0CF11E0A1B11AE1")


def extract_hwx(path):
    """hwx(결재문서): 안에 든 OLE HWP를 찾아 추출, 없으면 평문 한글 조각을 모음"""
    try:
        with open(path, "rb") as f:
            raw = f.read()
        text, method = "", "hwx_none"
        off = raw.find(_OLE_SIG)
        if off >= 0:
            r = extract_hwp_direct(raw[off:])          # olefile은 바이트열도 받음
            text = r.get("text", "")
            if text:
                method = "hwx_embedded_hwp"
        if not text:
            cand = re.findall(rb"(?:[\xea-\xed][\x80-\xbf]{2}){3,}", raw)
            text = nfc(b" ".join(cand).decode("utf-8", "ignore"))
            method = "hwx_rawscan" if text else "hwx_none"
        cleaned, dup = clean_text(text)
        return {"text": cleaned, "pages": 0, "chars_per_page": 0, "is_scanned": False,
                "ocr_applied": False, "dup_fixed": dup, "method": method}
    except Exception as e:
        return _empty_result(f"hwx 추출 실패: {str(e)[:120]}")


def _zip_member_name(info) -> str:
    """zipfile이 cp437로 잘못 읽은 한글 파일명을 cp949(또는 utf-8)로 복원"""
    name = info.filename
    if info.flag_bits & 0x800:                          # UTF-8 표시가 있으면 그대로
        return nfc(name)
    try:
        raw = name.encode("cp437")
    except UnicodeEncodeError:
        return nfc(name)
    for enc in ("cp949", "utf-8"):
        try:
            return nfc(raw.decode(enc))
        except UnicodeDecodeError:
            continue
    return nfc(name)


def _expand_zip(zip_path, dest_dir, depth=0):
    """zip을 dest_dir(임시 폴더)에 풀고 [(실제 경로, zip 안 상대경로)]와 오류 목록을 돌려줌.
    중첩 zip은 ZIP_MAX_DEPTH까지 재귀. 경로 탈출(..)·맥 부속 파일은 건너뜀."""
    members, errors = [], []
    try:
        zf = zipfile.ZipFile(zip_path)
    except UnicodeDecodeError:
        raise ValueError("파일명 인코딩 손상(UTF-8 표시가 붙은 비UTF-8 이름), 수동 해제 필요")
    with zf:
        infos = [i for i in zf.infolist() if not i.is_dir()]
        total = sum(i.file_size for i in infos)
        if total > ZIP_MAX_TOTAL:
            raise ValueError(f"zip 해제 크기 한도 초과: {total:,} 바이트")
        for info in infos:
            parts = [p.replace("\\", "_").strip()[:150] for p in _zip_member_name(info).split("/")]
            parts = [p for p in parts if p not in ("", ".", "..")]
            if (not parts or "__MACOSX" in parts or parts[-1].startswith("._")
                    or parts[-1] == ".DS_Store"):
                continue
            target = os.path.join(dest_dir, *parts)
            if os.path.exists(target):
                continue
            os.makedirs(os.path.dirname(target), exist_ok=True)
            rel = "/".join(parts)
            try:
                with zf.open(info) as src, open(target, "wb") as dst:
                    shutil.copyfileobj(src, dst)
            except Exception as e:                       # 암호 걸린 항목 등
                errors.append((rel, str(e)[:120]))
                continue
            ext = os.path.splitext(target)[1].lower()
            if ext == ".zip":
                if depth >= ZIP_MAX_DEPTH:
                    errors.append((rel, "중첩 zip 깊이 초과"))
                    continue
                try:
                    sub, sub_err = _expand_zip(target, target + "_풀림", depth + 1)
                except Exception as e:
                    errors.append((rel, f"중첩 zip 해제 실패: {str(e)[:100]}"))
                    continue
                members += [(p, f"{rel}/{r}") for p, r in sub]
                errors += [(f"{rel}/{r}", m) for r, m in sub_err]
            elif ext in SUPPORTED_EXTENSIONS:
                members.append((target, rel))
    return members, errors


def extract_zip(path):
    """zip 1개를 통째로 추출 (단일 파일 호출·MCP용). 폴더 일괄 처리(run)에서는
    zip 안 파일마다 별도 레코드를 만든다."""
    tmp = tempfile.mkdtemp(prefix="dataman_zip_")
    try:
        members, errors = _expand_zip(path, tmp)
        parts, n_ok = [], 0
        for real, rel in members:
            r = _extract_one(real)
            if r.get("text"):
                n_ok += 1
                parts.append(f"### [압축 내부] {rel}\n{r['text']}")
            if r.get("error"):
                errors.append((rel, r["error"]))
        cleaned, dup = clean_text("\n\n".join(parts))
        res = {"text": cleaned, "pages": len(members), "chars_per_page": 0, "is_scanned": False,
               "ocr_applied": False, "dup_fixed": dup, "method": "zip",
               "members": [rel for _, rel in members]}
        if errors:
            res["zip_errors"] = [f"{r}: {m}" for r, m in errors]
        return res
    except Exception as e:
        return _empty_result(f"zip 해제 실패: {str(e)[:120]}")
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


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
    elif ext in IMAGE_EXTENSIONS:
        return extract_image(path)
    elif ext in EXCEL_EXTENSIONS:
        return extract_excel(path)
    elif ext == ".hwx":
        return extract_hwx(path)
    elif ext == ".zip":
        return extract_zip(path)
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
            ".hwp": "hwp", ".hwpx": "hwpx", ".hwx": "hwx",
            ".jpg": "jpg", ".jpeg": "jpeg", ".png": "png",
            ".xlsx": "xlsx", ".xlsm": "xlsm", ".xls": "xls", ".zip": "zip"}.get(ext, "unknown")


# 추출 결과에서 entry로 그대로 옮기는 선택 필드 (있을 때만)
_OPTIONAL_FIELDS = ("tables", "ocr_engine", "vision_pages", "ocr_pages", "zoom_pages",
                    "garbled_pages", "ocr_errors", "members", "zip_errors")


def _build_entry(file_path, result, input_folder, rel_path=None, archive=None):
    """파일 1건의 추출 결과를 표준 entry dict로 변환 (JSON/JSONL 공통).
    rel_path: zip 안 파일처럼 실제 경로와 다른 상대경로를 쓸 때 지정
    archive: zip 안 파일이면 그 zip의 상대경로"""
    if rel_path is None:
        rel_path = os.path.relpath(file_path, input_folder)
        parent_folder = os.path.basename(os.path.dirname(file_path))
    else:
        parent_folder = (os.path.basename(os.path.dirname(rel_path))
                         or os.path.basename(os.path.abspath(input_folder)))
    rel_path = nfc(rel_path)
    filename = nfc(os.path.basename(rel_path))
    meta = parse_filename_metadata(filename)
    orig_type = get_file_type(filename)
    ext = os.path.splitext(filename)[1].lower()

    if result.get("method"):
        method = result["method"]
    elif ext == ".hwpx":
        method = "hwpx_direct"
    elif ext == ".hwp":
        method = "hwp_direct"
    elif ext == ".doc":
        method = "textutil" if TEXTUTIL_AVAILABLE else "libreoffice"
    else:
        method = "ocr" if result.get("ocr_applied") else orig_type

    text = result.get("text", "")
    entry = {
        "filename": filename, "file_type": orig_type,
        "file_no": meta["file_no"], "inst_name": nfc(meta["inst_name"]),
        "title": nfc(meta["title"]), "text": text,
        "pages": result.get("pages", 0),
        "text_length": len(text),
        "extraction_method": method,
        "dup_fixed": result.get("dup_fixed", False),   # 공백 정리 여부 (필드명은 하위 호환 유지)
        "folder": nfc(parent_folder),
        "rel_path": rel_path,
        "본문해시": text_hash(text),
    }
    for k in _OPTIONAL_FIELDS:
        if result.get(k):
            entry[k] = result[k]
    if archive:
        entry["archive"] = nfc(archive)
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
        elif ext in IMAGE_EXTENSIONS:
            return extract_image(file_path)
        elif ext in EXCEL_EXTENSIONS:
            return extract_excel(file_path)
        elif ext == ".hwx":
            return extract_hwx(file_path)
        elif ext == ".zip":
            return extract_zip(file_path)
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

def _worker_init(engine, lang):
    """spawn 방식 작업 프로세스 초기화: 부모의 OCR 설정을 넘겨받는다"""
    global OCR_LANG
    OCR_LANG = lang
    set_ocr_engine(engine)
    try:
        fitz.TOOLS.mupdf_display_errors(False)
    except Exception:
        pass


def _collect_tasks(input_folder, zip_mode, tmp_holder):
    """대상 파일 수집. 반환: (작업 목록, 건너뛴 zip 목록)
    작업 = {"path": 실제 경로, "rel": 상대경로, "archive": zip 상대경로 또는 None,
            "preset": 미리 정해진 결과(해제 실패 등) 또는 None}"""
    tasks, skipped = [], []
    for root, dirs, files in os.walk(input_folder):
        dirs.sort()
        for f in sorted(files):
            if f.startswith("._") or f == ".DS_Store":
                continue
            ext = os.path.splitext(f)[1].lower()
            if ext not in SUPPORTED_EXTENSIONS:
                continue
            p = os.path.join(root, f)
            rel = os.path.relpath(p, input_folder)
            if ext != ".zip":
                tasks.append({"path": p, "rel": rel, "archive": None, "preset": None})
                continue
            done_dir = os.path.join(root, ZIP_DONE_PREFIX + os.path.splitext(f)[0])
            if zip_mode == "skip" or (zip_mode == "auto" and os.path.isdir(done_dir)):
                skipped.append(rel)
                continue
            if tmp_holder[0] is None:
                tmp_holder[0] = tempfile.mkdtemp(prefix="dataman_zip_")
            dest = os.path.join(tmp_holder[0], f"z{len(tasks):06d}")
            try:
                members, errors = _expand_zip(p, dest)
            except Exception as e:
                tasks.append({"path": p, "rel": rel, "archive": None,
                              "preset": _empty_result(f"zip 해제 실패: {str(e)[:120]}")})
                continue
            for real, mrel in members:
                tasks.append({"path": real, "rel": os.path.join(rel, mrel),
                              "archive": rel, "preset": None})
            for mrel, msg in errors:
                tasks.append({"path": None, "rel": os.path.join(rel, mrel), "archive": rel,
                              "preset": _empty_result(f"zip 항목 해제 실패: {msg}")})
    tasks.sort(key=lambda t: nfc(t["rel"]))
    return tasks, skipped


def build_checklist(entries, skipped_zips=(), title=""):
    """점검표(마크다운) 줄 목록: 본문 0자·50자 미만·섞인 스캔 쪽·깨진 글자층·오류·같은 본문"""
    empty = [e for e in entries if not e["text"]]
    short = [e for e in entries if 0 < e["text_length"] < 50]
    mixed = [e for e in entries if "+" in e["extraction_method"]]
    full_ocr = [e for e in entries if e["extraction_method"] in ("vision", "ocr")
                and e["file_type"] == "pdf"]
    garb = [e for e in entries if e.get("garbled_pages")]
    errs = [e for e in entries if e.get("error") or e.get("ocr_errors") or e.get("zip_errors")]
    zoom = [e for e in entries if e.get("zoom_pages")]
    hashes = {}
    for e in entries:
        if e["본문해시"]:
            hashes[e["본문해시"]] = hashes.get(e["본문해시"], 0) + 1
    dup_groups = [c for c in hashes.values() if c > 1]
    ocr_pg = sum(len(e.get("vision_pages") or e.get("ocr_pages") or []) for e in entries)
    methods = {}
    for e in entries:
        methods[e["extraction_method"]] = methods.get(e["extraction_method"], 0) + 1

    def pages_str(lst):
        s = ",".join(map(str, lst[:15]))
        return s + (f" 외 {len(lst) - 15}쪽" if len(lst) > 15 else "")

    L = [f"# DataMan 점검표{(' - ' + title) if title else ''}", "",
         f"- 생성: {datetime.now():%Y-%m-%d %H:%M}",
         f"- OCR 엔진: {ocr_status_text()}",
         f"- 레코드 {len(entries):,}건 · 글자 {sum(e['text_length'] for e in entries):,}자",
         f"- OCR로 읽은 쪽 {ocr_pg:,}쪽 (확대 인식 채택 "
         f"{sum(len(e.get('zoom_pages', [])) for e in entries):,}쪽)",
         f"- 같은 본문 2건 이상 묶음 {len(dup_groups):,}개 ({sum(dup_groups):,}건) "
         f"→ 건수 집계는 `본문해시`로 묶을 것",
         "", "## 요약", "",
         "| 항목 | 건수 |", "|---|---|",
         f"| 본문 0자 | {len(empty):,} |",
         f"| 본문 50자 미만 | {len(short):,} |",
         f"| 섞인 스캔(일부 쪽만 OCR) | {len(mixed):,} |",
         f"| 전 쪽 OCR PDF | {len(full_ocr):,} |",
         f"| 깨진 글자층 쪽이 있는 파일 | {len(garb):,} |",
         f"| 작은 이미지 확대 인식 파일 | {len(zoom):,} |",
         f"| 오류 | {len(errs):,} |",
         f"| 건너뛴 zip(이미 압축해제 폴더 있음) | {len(skipped_zips):,} |",
         "", "## 추출 방식", ""]
    L += [f"- {k}: {v:,}건" for k, v in sorted(methods.items(), key=lambda kv: -kv[1])]
    L += ["", f"## 본문 0자 ({len(empty):,}건)", ""]
    L += [f"- {e['rel_path']}" + (f" ({e['error'][:60]})" if e.get("error") else "") for e in empty]
    L += ["", f"## 본문 50자 미만 ({len(short):,}건)", ""]
    L += [f"- {e['rel_path']} : {e['text'][:40]!r}" for e in short]
    L += ["", f"## 섞인 스캔 쪽 ({len(mixed):,}건, OCR 쪽 번호)", ""]
    L += [f"- {e['rel_path']} : {pages_str(e.get('vision_pages') or e.get('ocr_pages') or [])}"
          f" / 전체 {e['pages']}쪽" for e in mixed]
    L += ["", f"## 깨진 글자층 ({len(garb):,}건, 쪽 번호)", ""]
    L += [f"- {e['rel_path']} : {pages_str(e['garbled_pages'])}" for e in garb]
    L += ["", f"## 오류 ({len(errs):,}건)", ""]
    for e in errs:
        msg = e.get("error") or "; ".join((e.get("ocr_errors") or e.get("zip_errors") or [])[:3])
        L.append(f"- {e['rel_path']} : {str(msg)[:120]}")
    if skipped_zips:
        L += ["", f"## 건너뛴 zip ({len(skipped_zips):,}건)", ""]
        L += [f"- {nfc(z)}" for z in skipped_zips]
    return L


def run(input_folder, output_dir=None, log_callback=None, progress_callback=None,
        done_callback=None, stop_event=None, output_format='json',
        workers=None, zip_mode='auto'):
    """통합 전처리 실행 (하위 폴더 항상 포함)
    output_dir: 결과 저장 폴더 (None이면 OUTPUT_DIR 사용)
    log_callback(msg): GUI 로그 출력
    progress_callback(done, total): 진행률
    done_callback(summary_text): 완료 시 호출
    stop_event: threading.Event -- set()하면 중단
    output_format: 'json' (단일 통합 JSON) 또는 'jsonl' (라인별 JSON, 스트리밍/중단 안전)
    workers: 작업 프로세스 수 (None이면 PARALLEL_WORKERS)
    zip_mode: 'auto'(옆에 압축해제_ 폴더가 있으면 건너뜀) / 'expand'(항상 풂) / 'skip'(zip 무시)
    """
    if stop_event is None:
        stop_event = threading.Event()

    save_dir = output_dir if output_dir else OUTPUT_DIR
    output_format = (output_format or 'json').lower()
    if output_format not in ('json', 'jsonl'):
        output_format = 'json'
    if zip_mode not in ('auto', 'expand', 'skip'):
        zip_mode = 'auto'

    def _log(msg, overwrite=False):
        if log_callback:
            log_callback(msg)
        else:
            safe_print(msg, overwrite=overwrite)

    now = datetime.now()
    timestamp = now.strftime('%Y-%m-%d %H:%M:%S')
    ts_file = now.strftime('%Y%m%d_%H%M%S')
    folder_name = nfc(os.path.basename(os.path.abspath(input_folder)))
    out_ext = 'jsonl' if output_format == 'jsonl' else 'json'
    OUTPUT_JSON = os.path.join(save_dir, f"dataman_{folder_name}_{ts_file}.{out_ext}")
    EXTRACT_LOG = os.path.join(save_dir, f"추출로그_{folder_name}_{ts_file}.txt")
    CHECKLIST = os.path.join(save_dir, f"점검표_{folder_name}_{ts_file}.md")

    _log(f"{'='*50}")
    _log(f"  DataMan for Mac v1.2")
    _log(f"  HWP/HWPX/HWX/PDF/TXT/DOC/DOCX/XLSX/XLS/이미지/ZIP -> {out_ext.upper()}")
    _log(f"{'='*50}")
    _log(f"  대상 폴더 : {input_folder}")
    _log(f"  저장 경로 : {OUTPUT_JSON}")
    _log(f"  출력 형식 : {output_format} ({'스트리밍' if output_format == 'jsonl' else '단일 통합'})")
    _log(f"  추출 로그 : {EXTRACT_LOG}")
    _log(f"  점검표    : {CHECKLIST}")
    _log(f"  OCR      : {ocr_status_text()}")
    _doc_tool = "textutil" if TEXTUTIL_AVAILABLE else ("LibreOffice" if LIBREOFFICE_AVAILABLE else "불가")
    _log(f"  DOC 추출 : {_doc_tool}")
    _log(f"  zip 처리 : {zip_mode}")
    _log(f"{'='*50}")

    # -- 파일 수집 (항상 재귀, zip은 임시 폴더에 풀어 내부 파일을 작업으로 추가) --
    tmp_holder = [None]
    try:
        tasks, skipped_zips = _collect_tasks(input_folder, zip_mode, tmp_holder)
        _run_tasks(input_folder, tasks, skipped_zips, save_dir, output_format, workers,
                   OUTPUT_JSON, EXTRACT_LOG, CHECKLIST, folder_name, ts_file, timestamp,
                   _log, progress_callback, done_callback, stop_event)
    finally:
        if tmp_holder[0]:
            shutil.rmtree(tmp_holder[0], ignore_errors=True)


def _run_tasks(input_folder, tasks, skipped_zips, save_dir, output_format, workers,
               OUTPUT_JSON, EXTRACT_LOG, CHECKLIST, folder_name, ts_file, timestamp,
               _log, progress_callback, done_callback, stop_event):
    if skipped_zips:
        _log(f"건너뛴 zip {len(skipped_zips)}개 (옆에 {ZIP_DONE_PREFIX} 폴더가 이미 있음)")
    if not tasks:
        _log(f"오류: 지원 파일이 없습니다. ({', '.join(SUPPORTED_EXTENSIONS)})")
        if done_callback:
            done_callback("지원 파일 없음")
        return

    # 형식별 집계
    type_counts = {}
    has_doc = False
    for t in tasks:
        ft = get_file_type(t["rel"]).upper()
        type_counts[ft] = type_counts.get(ft, 0) + 1
        if ft == "DOC":
            has_doc = True
    n_zip_members = sum(1 for t in tasks if t["archive"])

    total = len(tasks)
    type_strs = [f"{ft} {cnt}개" for ft, cnt in sorted(type_counts.items())]
    _log(f"대상 파일: {total}개 ({', '.join(type_strs)})"
         + (f" / 그중 zip 내부 {n_zip_members}개" if n_zip_members else ""))
    if has_doc and not (TEXTUTIL_AVAILABLE or LIBREOFFICE_AVAILABLE):
        _log("경고: textutil/LibreOffice 미설치 -- DOC 파일은 추출되지 않습니다.")

    # -- 텍스트 추출 (spawn 방식 작업 프로세스: Vision·PyMuPDF를 프로세스마다 따로 씀) --
    run_idx = [i for i, t in enumerate(tasks) if t["preset"] is None]
    n_workers = max(1, min(workers or PARALLEL_WORKERS, len(run_idx) or 1))
    _log(f"텍스트 추출 시작 ({total}건, 작업 프로세스 {n_workers}개)")
    _log(f"{'='*50}")

    extraction_results = [None] * total
    for i, t in enumerate(tasks):
        if t["preset"] is not None:
            extraction_results[i] = t["preset"]
    start_time = time.time()
    done_count = 0
    eta_calc = ETACalculator(total)

    # JSONL 모드: 라인별 즉시 기록 (스트리밍·중단 안전)
    jsonl_fp = None
    if output_format == 'jsonl':
        jsonl_fp = open(OUTPUT_JSON, 'w', encoding='utf-8')

    def _emit(idx, result):
        nonlocal done_count
        done_count += 1
        t = tasks[idx]
        filename = nfc(os.path.basename(t["rel"]))
        ft = get_file_type(filename).upper()
        if jsonl_fp is not None:
            try:
                entry = _build_entry(t["path"] or t["rel"], result, input_folder,
                                     rel_path=t["rel"] if t["archive"] or t["path"] is None else None,
                                     archive=t["archive"])
                jsonl_fp.write(json.dumps(entry, ensure_ascii=False) + '\n')
                jsonl_fp.flush()
            except OSError as e:
                _log(f"JSONL 기록 실패: {e}")
        pct = done_count * 100 // total
        eta_str = eta_calc.update(done_count)
        progress = f"[{done_count}/{total} {pct}%{(' ' + eta_str) if eta_str else ''}]"
        text_len = len(result.get("text", ""))
        n_ocr = len(result.get("vision_pages") or result.get("ocr_pages") or [])
        ocr_note = f" (OCR {n_ocr}쪽)" if n_ocr and ft == "PDF" else ""
        if result.get("error"):
            _log(f"  {progress} [X] [{ft}] {filename[:45]}", overwrite=True)
        elif text_len == 0:
            _log(f"  {progress} [!] [{ft}] {filename[:45]}  텍스트 없음", overwrite=True)
        else:
            _log(f"  {progress} [O] [{ft}] {filename[:45]}  {text_len:,}자{ocr_note}", overwrite=True)
        if progress_callback:
            progress_callback(done_count, total)

    try:
        for i, t in enumerate(tasks):
            if t["preset"] is not None:
                _emit(i, t["preset"])
        if run_idx:
            ctx = multiprocessing.get_context("spawn")
            with ProcessPoolExecutor(max_workers=n_workers, mp_context=ctx,
                                     initializer=_worker_init,
                                     initargs=(OCR_ENGINE, OCR_LANG)) as pool:
                future_to_idx = {pool.submit(_extract_one, tasks[i]["path"]): i for i in run_idx}
                for future in as_completed(future_to_idx):
                    if stop_event.is_set():
                        _log("사용자에 의해 중단되었습니다.")
                        pool.shutdown(wait=False, cancel_futures=True)
                        break
                    idx = future_to_idx[future]
                    try:
                        result = future.result()
                    except Exception as e:           # 작업 프로세스 비정상 종료 등
                        result = _empty_result(f"작업 프로세스 오류: {str(e)[:120]}")
                    extraction_results[idx] = result
                    _emit(idx, result)
    finally:
        if jsonl_fp is not None:
            try:
                jsonl_fp.close()
            except Exception:
                pass

    def _entry(i, res):
        t = tasks[i]
        return _build_entry(t["path"] or t["rel"], res, input_folder,
                            rel_path=t["rel"] if t["archive"] or t["path"] is None else None,
                            archive=t["archive"])

    # 중단 처리: 부분 결과 저장
    if stop_event.is_set():
        if output_format == 'json':
            partial_results = [_entry(i, r) for i, r in enumerate(extraction_results) if r is not None]
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

    # 결과 조립 (파일 순서 유지)
    results = []
    log_lines = []
    stats = {
        "success": 0, "empty": 0, "scanned": 0,
        "ocr_ok": 0, "ocr_fail": 0, "error": 0,
        "dup_fixed": 0, "total_chars": 0,
    }
    for idx in range(total):
        result = extraction_results[idx]
        if result is None:
            result = _empty_result("중단됨")
        entry = _entry(idx, result)
        results.append(entry)
        rel_path = entry["rel_path"]

        text_len = entry["text_length"]
        stats["total_chars"] += text_len
        if result.get("dup_fixed"):
            stats["dup_fixed"] += 1
        extra = " [공백정리]" if result.get("dup_fixed") else ""

        # 로그 기록 (파일 순서)
        if result.get("error"):
            stats["error"] += 1
            log_lines.append(f"[X] {rel_path} -- {result['error']}")
        elif result.get("ocr_applied"):
            stats["ocr_ok"] += 1
            pg = result.get("vision_pages") or result.get("ocr_pages") or []
            log_lines.append(f"[OCR] {rel_path} -- {text_len:,}자 ({entry['extraction_method']}, "
                             f"{len(pg)}쪽){extra}")
        elif result.get("is_scanned") and text_len == 0:
            if OCR_AVAILABLE:
                stats["ocr_fail"] += 1
                log_lines.append(f"[!] {rel_path} -- OCR 실패")
            else:
                stats["scanned"] += 1
                log_lines.append(f"[!] {rel_path} -- 스캔본 (OCR 미사용)")
        elif text_len == 0:
            stats["empty"] += 1
            log_lines.append(f"[!] {rel_path} -- 텍스트 없음")
        else:
            stats["success"] += 1
            log_lines.append(f"[O] {rel_path} -- {text_len:,}자{extra}")

    # -- 결과 저장 (JSONL은 이미 라인별 기록 완료, 파일 순서로 다시 정렬해 씀) --
    if output_format == 'json':
        with open(OUTPUT_JSON, 'w', encoding='utf-8') as f:
            json.dump(results, f, ensure_ascii=False, indent=2)
    else:
        with open(OUTPUT_JSON, 'w', encoding='utf-8') as f:
            for e in results:
                f.write(json.dumps(e, ensure_ascii=False) + '\n')

    # -- 최종 집계 --
    elapsed = time.time() - start_time
    file_size = os.path.getsize(OUTPUT_JSON) / 1024 / 1024

    summary_lines = [
        f"전처리 완료!",
        f"  총 파일 수       : {total}개",
        f"  텍스트 추출 성공 : {stats['success']}개",
    ]
    if stats["ocr_ok"] > 0:
        summary_lines.append(f"  OCR 보완 추출    : {stats['ocr_ok']}개")
    if stats["ocr_fail"] > 0:
        summary_lines.append(f"  OCR 추출 실패    : {stats['ocr_fail']}개")
    if stats["scanned"] > 0:
        summary_lines.append(f"  스캔본(OCR없음)  : {stats['scanned']}개")
    if stats["empty"] > 0:
        summary_lines.append(f"  텍스트 없음      : {stats['empty']}개")
    if stats["error"] > 0:
        summary_lines.append(f"  오류 발생        : {stats['error']}개")
    if stats["dup_fixed"] > 0:
        summary_lines.append(f"  공백 정리        : {stats['dup_fixed']}개")
    if skipped_zips:
        summary_lines.append(f"  건너뛴 zip       : {len(skipped_zips)}개")
    summary_lines += [
        f"  총 글자 수       : {stats['total_chars']:,}자",
        f"  평균 글자 수     : {stats['total_chars'] // max(total, 1):,}자/파일",
        f"  소요 시간        : {elapsed:.1f}초",
        f"",
        f"  JSON 저장 : {OUTPUT_JSON} ({file_size:.1f} MB)",
        f"  추출 로그 : {EXTRACT_LOG}",
        f"  점검표    : {CHECKLIST}",
    ]

    # -- 점검표 (자동 출력) --
    checklist = build_checklist(results, skipped_zips, folder_name)
    try:
        with open(CHECKLIST, 'w', encoding='utf-8') as f:
            f.write("\n".join(checklist) + "\n")
    except OSError as e:
        _log(f"점검표 저장 실패: {e}")

    # -- 최종 출력 --
    _log(f"{'='*50}")
    for line in summary_lines:
        _log(f"  {line}" if line and not line.startswith("  ") else line)
    _log(f"{'='*50}")
    _log("[점검표 요약]")
    for line in checklist:
        if line.startswith("| ") and not line.startswith("| 항목"):
            _log("  " + line.strip("| ").replace(" | ", " : "))
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
        _ocr_status = {"vision": "Vision", "tesseract": "Tesseract"}.get(OCR_ENGINE, "X")
        env_txt = f"OCR: {_ocr_status} | DOC: {_doc_status} | HWP/HWPX: 직접파싱"
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
    parser.add_argument("--ocr-engine", choices=["auto", "vision", "tesseract", "none"],
                        default="auto",
                        help="OCR 엔진 (기본 auto: Vision, 없으면 Tesseract)")
    parser.add_argument("--workers", type=int, default=None,
                        help=f"작업 프로세스 수 (기본 {PARALLEL_WORKERS})")
    parser.add_argument("--zip", dest="zip_mode", choices=["auto", "expand", "skip"],
                        default="auto",
                        help="zip 처리: auto(옆에 압축해제_ 폴더 있으면 건너뜀) / expand / skip")
    parser.add_argument("--cli", action="store_true", help="CLI 모드 (GUI 없이)")
    args = parser.parse_args()

    # OCR 언어 설정 반영
    OCR_LANG = args.ocr_lang
    set_ocr_engine(args.ocr_engine)

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
        run(folder, output_format=args.format, workers=args.workers, zip_mode=args.zip_mode)
    else:
        # GUI 모드 (기본)
        app = DataManGUI()
        app.mainloop()
