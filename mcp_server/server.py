# -*- coding: utf-8 -*-
"""DataMan MCP 서버 entry point.

stdio transport로 동작. Claude Desktop·Code에서 mcpServers에 등록 후 사용.

실행:
    python -m mcp_server.server

Claude Desktop 등록 (~/Library/Application Support/Claude/claude_desktop_config.json):
{
  "mcpServers": {
    "dataman": {
      "command": "python3",
      "args": ["-m", "mcp_server.server"],
      "cwd": "/path/to/dataman"
    }
  }
}
"""
from typing import Optional

from mcp.server.fastmcp import Context, FastMCP

from . import tools

mcp = FastMCP("dataman")


@mcp.tool()
def scan_folder(path: str) -> dict:
    """폴더를 재귀 스캔하여 형식별 파일 개수·총 용량·예상 처리시간을 반환한다.

    Args:
        path: 스캔할 폴더 절대 경로

    Returns:
        total, by_format, total_mb, estimated_seconds, samples 등
    """
    return tools.scan_folder(path)


@mcp.tool()
def extract_summary(path: str, preview_chars: int = 200) -> dict:
    """단일 문서의 메타데이터·미리보기만 추출한다 (LLM 컨텍스트 절약용).

    Args:
        path: 문서 파일 절대 경로 (HWP/HWPX/PDF/DOC/DOCX/TXT)
        preview_chars: 미리보기 글자 수 (기본 200)

    Returns:
        filename, file_type, pages, tables, text_length, preview 등
    """
    return tools.extract_summary(path, preview_chars=preview_chars)


@mcp.tool()
def extract_text(path: str,
                 page_start: Optional[int] = None,
                 page_end: Optional[int] = None,
                 max_chars: int = 50000) -> dict:
    """단일 문서의 본문 텍스트를 추출한다.

    PDF는 page_start/page_end로 페이지 범위 지정 가능 (1-based, 끝 페이지 포함).
    HWP/HWPX/DOC/DOCX/TXT는 전체 추출 후 max_chars로 절단.
    HWPX는 표가 마크다운 형식(| 셀1 | 셀2 |)으로 본문에 포함되어 있다.

    Args:
        path: 문서 파일 절대 경로
        page_start: PDF 시작 페이지 (1-based, 옵션)
        page_end: PDF 종료 페이지 (옵션)
        max_chars: 반환 텍스트 최대 글자 수 (기본 50,000)

    Returns:
        text, text_length, truncated, pages 등
    """
    return tools.extract_text(path,
                              page_start=page_start,
                              page_end=page_end,
                              max_chars=max_chars)


@mcp.tool()
async def extract_folder(path: str,
                         output_dir: Optional[str] = None,
                         skip_ocr: bool = False,
                         ctx: Context = None) -> dict:
    """폴더 안 모든 지원 문서를 일괄 추출하여 JSONL로 저장한다.

    처리 도중 progress 알림을 보내 LLM 클라이언트 타임아웃을 방지한다.
    JSONL은 라인별 즉시 flush되므로 중단되어도 진행분이 보존된다.

    Args:
        path: 추출 대상 폴더 절대 경로 (재귀)
        output_dir: 결과 JSONL 저장 폴더 (None이면 path 자체)
        skip_ocr: True면 스캔 PDF에 OCR 적용 안 함 (속도 우선)

    Returns:
        jsonl_path: 출력 파일 경로
        total, success, empty, error, ocr_applied, tables_total, total_chars, elapsed_seconds
    """
    return await tools.extract_folder(path,
                                      output_dir=output_dir,
                                      skip_ocr=skip_ocr,
                                      ctx=ctx)


def main():
    mcp.run()


if __name__ == "__main__":
    main()
