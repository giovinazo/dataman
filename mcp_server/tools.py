# -*- coding: utf-8 -*-
"""DataMan MCP 도구 구현 (4개 MVP).

scan_folder       -- 폴더 사전 스캔 (빠름)
extract_summary   -- 단일 파일 메타·미리보기
extract_text      -- 단일 파일 본문 추출 (페이지 범위 옵션)
extract_folder    -- 폴더 일괄 추출 + progress notifications + JSONL 출력
"""
import json
import os
import time
import zipfile
from datetime import datetime
from typing import Optional

from . import adapter


# ─────────────────────────────────────────────────────────
# scan_folder
# ─────────────────────────────────────────────────────────

def scan_folder(path: str) -> dict:
    """폴더를 재귀 스캔하여 형식별 파일 개수·총 용량·예상 처리시간을 반환."""
    if not os.path.isdir(path):
        return {"error": f"폴더가 아니거나 존재하지 않음: {path}"}

    by_format = {}
    total = 0
    total_bytes = 0
    samples = []

    for root, _, files in os.walk(path):
        for f in files:
            ext = os.path.splitext(f)[1].lower()
            if ext not in adapter.SUPPORTED_EXTENSIONS:
                continue
            fp = os.path.join(root, f)
            try:
                size = os.path.getsize(fp)
            except OSError:
                size = 0
            ft = ext.lstrip(".")
            by_format[ft] = by_format.get(ft, 0) + 1
            total += 1
            total_bytes += size
            if len(samples) < 5:
                samples.append(os.path.relpath(fp, path))

    # 추정 처리 시간 (실측치 기반: 평균 약 0.13초/파일, OCR 미적용)
    est_seconds = round(total * 0.13, 1)

    return {
        "folder": path,
        "total": total,
        "by_format": by_format,
        "total_bytes": total_bytes,
        "total_mb": round(total_bytes / 1024 / 1024, 2),
        "estimated_seconds": est_seconds,
        "samples": samples,
        "supported_formats": [e.lstrip(".") for e in adapter.SUPPORTED_EXTENSIONS],
    }


# ─────────────────────────────────────────────────────────
# extract_summary
# ─────────────────────────────────────────────────────────

def extract_summary(path: str, preview_chars: int = 200) -> dict:
    """단일 파일의 메타데이터·미리보기만 반환 (LLM 컨텍스트 절약용)."""
    if not os.path.isfile(path):
        return {"error": f"파일 없음: {path}"}

    ext = os.path.splitext(path)[1].lower()
    if ext not in adapter.SUPPORTED_EXTENSIONS:
        return {"error": f"미지원 형식: {ext}"}

    result = adapter._extract_one(path)
    text = result.get("text", "")

    summary = {
        "filename": os.path.basename(path),
        "file_type": ext.lstrip("."),
        "pages": result.get("pages", 0),
        "tables": result.get("tables", 0),  # HWPX만 채워짐
        "text_length": len(text),
        "preview": text[:preview_chars] + ("..." if len(text) > preview_chars else ""),
        "is_scanned": result.get("is_scanned", False),
        "ocr_applied": result.get("ocr_applied", False),
    }
    if result.get("error"):
        summary["error"] = result["error"]
    return summary


# ─────────────────────────────────────────────────────────
# extract_text (페이지 범위)
# ─────────────────────────────────────────────────────────

def extract_text(path: str, page_start: Optional[int] = None,
                 page_end: Optional[int] = None,
                 max_chars: int = 50000) -> dict:
    """단일 파일 본문 추출. PDF는 page_start·page_end 범위 한정 가능.
    HWP/HWPX/DOC/DOCX/TXT는 전체 추출 후 max_chars로 절단.
    """
    if not os.path.isfile(path):
        return {"error": f"파일 없음: {path}"}

    ext = os.path.splitext(path)[1].lower()
    if ext not in adapter.SUPPORTED_EXTENSIONS:
        return {"error": f"미지원 형식: {ext}"}

    if ext == ".pdf" and (page_start is not None or page_end is not None):
        # PDF만 페이지 범위 추출 지원
        try:
            import fitz
            doc = fitz.open(path)
            try:
                total_pages = doc.page_count
                start = max(1, page_start or 1)
                end = min(total_pages, page_end or total_pages)
                texts = []
                for i in range(start - 1, end):
                    try:
                        t = doc[i].get_text("text")
                        if t:
                            texts.append(t)
                    except (OSError, RuntimeError):
                        continue
                cleaned, _ = adapter.clean_text("\n".join(texts))
                truncated = len(cleaned) > max_chars
                if truncated:
                    cleaned = cleaned[:max_chars]
                return {
                    "filename": os.path.basename(path),
                    "page_start": start,
                    "page_end": end,
                    "total_pages": total_pages,
                    "text": cleaned,
                    "text_length": len(cleaned),
                    "truncated": truncated,
                }
            finally:
                doc.close()
        except OSError as e:
            return {"error": f"PDF 추출 실패: {e}"}

    # 그 외: 전체 추출 후 max_chars 절단
    result = adapter._extract_one(path)
    text = result.get("text", "")
    truncated = len(text) > max_chars
    if truncated:
        text = text[:max_chars]
    out = {
        "filename": os.path.basename(path),
        "file_type": ext.lstrip("."),
        "pages": result.get("pages", 0),
        "text": text,
        "text_length": len(text),
        "truncated": truncated,
    }
    if "tables" in result:
        out["tables"] = result["tables"]
    if result.get("error"):
        out["error"] = result["error"]
    return out


# ─────────────────────────────────────────────────────────
# extract_folder (progress notifications + JSONL)
# ─────────────────────────────────────────────────────────

async def extract_folder(path: str, output_dir: Optional[str] = None,
                         skip_ocr: bool = False, ctx=None) -> dict:
    """폴더 일괄 추출. progress notifications로 타임아웃 회피, JSONL 즉시 flush.

    ctx: MCP Context (None이면 progress 보고 생략)
    """
    if not os.path.isdir(path):
        return {"error": f"폴더 없음: {path}"}

    # 파일 수집
    all_files = []
    for root, _, files in os.walk(path):
        for f in sorted(files):
            if os.path.splitext(f)[1].lower() in adapter.SUPPORTED_EXTENSIONS:
                all_files.append(os.path.join(root, f))
    all_files.sort()

    if not all_files:
        return {"error": "지원 파일 없음", "total": 0}

    # 출력 경로
    save_dir = output_dir or path
    os.makedirs(save_dir, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    folder_name = os.path.basename(os.path.abspath(path))
    out_path = os.path.join(save_dir, f"dataman_{folder_name}_{ts}.jsonl")

    # OCR skip 옵션 처리: dataman의 OCR_AVAILABLE을 임시로 끄지 않고,
    # 추출 후 OCR 적용된 결과만 텍스트 비움 (단순화)
    stats = {"success": 0, "empty": 0, "error": 0, "ocr_applied": 0,
             "total_chars": 0, "tables_total": 0}
    start = time.time()
    total = len(all_files)

    with open(out_path, "w", encoding="utf-8") as fp:
        for i, fpath in enumerate(all_files, 1):
            if ctx is not None:
                try:
                    await ctx.report_progress(
                        i - 1, total,
                        f"[{i}/{total}] {os.path.basename(fpath)[:40]}"
                    )
                except Exception:
                    pass

            result = adapter._extract_one(fpath)

            if skip_ocr and result.get("ocr_applied"):
                # OCR 결과 폐기 (텍스트만 비움)
                result = dict(result)
                result["text"] = ""
                result["ocr_applied"] = False
                result["error"] = "OCR skip"

            entry = adapter._build_entry(fpath, result, path)
            fp.write(json.dumps(entry, ensure_ascii=False) + "\n")
            fp.flush()

            text_len = entry["text_length"]
            stats["total_chars"] += text_len
            if "tables" in entry:
                stats["tables_total"] += entry["tables"]
            if entry.get("error"):
                stats["error"] += 1
            elif text_len == 0:
                stats["empty"] += 1
            else:
                stats["success"] += 1
            if result.get("ocr_applied"):
                stats["ocr_applied"] += 1

    if ctx is not None:
        try:
            await ctx.report_progress(total, total, "완료")
        except Exception:
            pass

    return {
        "jsonl_path": out_path,
        "total": total,
        "elapsed_seconds": round(time.time() - start, 1),
        **stats,
    }
