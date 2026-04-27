# 프로그램 저작권 등록 설명서

## 1. 프로그램 개요

| 항목 | 내용 |
|------|------|
| 프로그램명 | DataMan (데이터맨 - 문서 텍스트 추출 도구) |
| 버전 | 1.1 |
| 저작자 | 허재영 |
| 창작연도 | 2025년 |
| 최종 수정일 | 2026년 4월 27일 |
| 개발 언어 | Python 3 |
| 코드 규모 | 약 1,500줄 (단일 파일) |
| 실행 환경 | Windows / macOS (Python 3.10 이상) |
| 사용자 인터페이스 | GUI (tkinter) + CLI (명령줄) |

### v1.1 변경사항 (2026-04-27)

- HWPX 표 구조 마크다운 보존: 한/글 표(hp:tbl)를 `| 셀1 | 셀2 |` 형식으로 본문에 통합 추출
- JSONL 스트리밍 출력 옵션 추가 (`--format jsonl` / GUI 라디오 버튼)
- 중단 시 부분 결과 자동 저장 (JSON 모드는 `dataman_partial_*.json`, JSONL은 라인별 즉시 기록)
- 추출 결과에 `tables` 필드 추가 (HWPX 표 개수)

### 프로그램 목적 및 용도

다양한 형식의 문서 파일(HWP, HWPX, PDF, TXT, DOC, DOCX)에서 텍스트를 추출하여 하나의 JSON 파일로 통합하는 데이터 전처리 도구이다. 대량의 문서를 일괄 처리하여 텍스트 분석, AI 학습 데이터 구축, 검색 시스템 구축 등 후속 작업에 활용할 수 있는 구조화된 데이터를 생성한다.

---

## 2. 주요 기능

### 2.1 문서 형식별 텍스트 추출

| 문서 형식 | 추출 방식 | 특징 |
|----------|----------|------|
| HWP (한글 97-2007) | OLE2 바이너리 직접 파싱 | 외부 프로그램 불필요, zlib 압축 해제 |
| HWPX (한글 2010+) | ZIP+XML 파싱 | XXE 공격 방지, Zip Bomb 보안 처리 |
| PDF | PyMuPDF 텍스트 추출 | 스캔 PDF 자동 감지 및 OCR 폴백 |
| TXT | 다중 인코딩 자동 감지 | UTF-8, CP949, EUC-KR, UTF-16, Latin-1 |
| DOC (Word 97-2003) | MS Word COM 자동화 | Windows 전용, 스레드 안전 래퍼 |
| DOCX (Word 2007+) | python-docx | 문단 및 표 텍스트 추출 |

### 2.2 실행 모드

**GUI 모드** (기본):
- 폴더 및 출력 경로 선택 대화상자
- 실시간 진행 바 및 ETA 표시
- 색상 구분 로그 (녹색=성공, 빨간색=오류)
- 설정 자동 저장 (마지막 사용 폴더, 창 크기)

**CLI 모드**:
```
python dataman.py --input "폴더경로"
python dataman.py --input "폴더경로" --output "출력경로"
python dataman.py --input "폴더경로" --format jsonl       # 스트리밍·중단 안전
python dataman.py --ocr-lang "kor"
```

### 2.3 처리 기능

- 재귀적 하위 폴더 탐색
- 병렬 처리 (CPU 코어 수 기반, 최대 8개 워커)
- 파일별 오류 격리 (실패 시 다음 파일로 계속 진행)
- 파일명 메타데이터 자동 추출 (파일번호, 기관명, 제목)
- 텍스트 정리 (제어문자 제거, 공백 정규화)
- 처리 완료 후 추출 로그 파일 자동 생성

### 2.4 출력 형식

**JSON 출력** (`dataman_[폴더명]_[타임스탬프].json`) -- 단일 통합 파일:
```json
[
  {
    "filename": "파일명.hwp",
    "file_type": "hwp",
    "file_no": "31932",
    "inst_name": "기관명",
    "title": "제목",
    "text": "추출된 본문 텍스트",
    "pages": 10,
    "text_length": 5000,
    "extraction_method": "hwp_direct",
    "tables": 5
  }
]
```

**JSONL 출력** (`dataman_[폴더명]_[타임스탬프].jsonl`) -- 라인별 스트리밍 (v1.1):
```
{"filename": "파일1.hwpx", "file_type": "hwpx", ...}
{"filename": "파일2.pdf", "file_type": "pdf", ...}
```
- 처리 즉시 라인 단위 기록 → 중단되어도 done까지 저장됨
- RAG 임베딩 등 후속 파이프라인에 직접 스트리밍 투입 가능

**부분 결과 저장** (v1.1) -- 중단 시 자동 생성:
- JSON 모드: `dataman_partial_[폴더명]_[타임스탬프].json`
- JSONL 모드: 본 출력 파일이 자동으로 부분 저장됨

**추출 로그** (`추출로그_[폴더명]_[타임스탬프].txt`):
- 파일별 처리 결과: [O]=성공, [!]=빈 파일/스캔, [X]=오류
- 요약 통계: 파일 수, 문자 수, OCR 통계, 소요 시간

---

## 3. 프로그램 구조

### 3.1 주요 클래스

| 클래스명 | 역할 |
|---------|------|
| DataManGUI | 메인 GUI 애플리케이션 (파일 선택, 진행률, 로그 표시) |
| WordTextExtractor | MS Word COM 객체 관리 (DOC 추출용, 재사용·자동복구) |
| BoundedText | 줄 수 제한 텍스트 위젯 (메모리 관리) |
| ETACalculator | 작업 예상 소요시간 계산기 |

### 3.2 주요 추출 함수

| 함수명 | 역할 |
|--------|------|
| extract_hwp_direct() | HWP 바이너리 직접 파싱 (OLE2 + zlib) |
| extract_hwpx_direct() | HWPX ZIP+XML 파싱 (보안 처리 포함) |
| extract_pdf() | PDF 텍스트 추출 + OCR 폴백 |
| extract_txt() | TXT 다중 인코딩 추출 |
| extract_doc() | DOC COM 자동화 추출 |
| extract_docx() | DOCX 문단/표 추출 |

### 3.3 데이터 처리 흐름

```
입력 폴더 선택 (GUI 또는 CLI)
    ↓
하위 폴더 포함 파일 목록 수집
    ↓
파일 형식별 분류 (확장자 기반)
    ↓
병렬 텍스트 추출 (DOC만 순차 처리)
    ↓
텍스트 정리 및 메타데이터 추출
    ↓
JSON 파일 저장 + 추출 로그 생성
```

---

## 4. 기술적 특징

### 4.1 독창적 구현 요소

- **HWP 바이너리 직접 파싱**: 한/글 프로그램 없이 OLE2 컨테이너에서 HWPTAG_PARA_TEXT 레코드를 직접 파싱하여 텍스트를 추출한다. 압축 여부를 FileHeader에서 확인하고 zlib로 해제하는 과정을 순수 Python으로 구현하였다.
- **HWPX 보안 파싱**: XML External Entity(XXE) 공격 방지를 위해 defusedxml을 활용하고, Zip Bomb 방지를 위해 비압축 크기를 200MB로 제한한다.
- **스캔 PDF 자동 감지**: 페이지당 추출 문자 수가 50자 미만이면 스캔 PDF로 판단하고 OCR로 자동 전환한다.
- **DOC 추출 스레드 안전 래퍼**: MS Word COM 객체가 스레드 안전하지 않으므로, 순차 처리 및 연결 실패 시 자동 재시작 메커니즘을 구현하였다.

### 4.2 사용 기술

- **GUI**: tkinter + ttk (순수 Python 표준 라이브러리)
- **HWP 파싱**: olefile (OLE2), zlib (압축 해제), struct (바이너리 파싱)
- **HWPX 파싱**: zipfile, xml.etree.ElementTree, defusedxml (선택)
- **PDF 처리**: PyMuPDF (fitz)
- **OCR**: pytesseract + Pillow (선택)
- **DOCX 처리**: python-docx
- **DOC 처리**: pywin32 (win32com, Windows 전용)
- **병렬 처리**: concurrent.futures (ThreadPoolExecutor)

---

## 5. 저작권 정보

- **저작자**: 허재영
- **창작연도**: 2025년
- **최종 수정**: 2026년 2월
- **저작물 유형**: 컴퓨터프로그램저작물
- **권리 범위**: 본 프로그램의 소스코드, GUI 디자인, 문서 파싱 알고리즘, 데이터 처리 로직 일체
