# DataON CoreTrustSeal 신청서 URL Dead Link 검사 작업 문서

**작업일**: 2026-05-12  
**대상 문서**: `DataON CoreTrustSeal Application_Update_250912.docx`  
**대상 문서 경로**: `CTS 갱신 준비_2026/2605 신청서 작성 준비 및 역할 분담/`  
**결과 파일**: `DataON_URL_Check_Results.xlsx`  
**검사 스크립트**: `check_urls.py`

---

## 1. 작업 개요

CoreTrustSeal 갱신 신청서에 포함된 모든 각주 URL(96개 항목)을 대상으로 dead link 여부를 자동 검사하고 그 결과를 엑셀 파일로 정리하였다. 봇차단 우회를 위해 실제 브라우저 헤더를 모사한 다중 User-Agent 전략을 사용하였다.

---

## 2. URL 추출 방법

### 2-1. DOCX 언팩

`.docx` 파일은 ZIP 아카이브 구조이므로 PowerShell `Expand-Archive`로 압축 해제하였다.

```powershell
$docxPath = "...\DataON CoreTrustSeal Application_Update_250912.docx"
Copy-Item $docxPath "temp_docx.zip"
Expand-Archive -Path "temp_docx.zip" -DestinationPath "docx_unpacked"
```

언팩 결과 구조:
```
docx_unpacked/
├── word/
│   ├── document.xml      # 본문
│   ├── footnotes.xml     # 각주
│   └── _rels/
│       └── document.xml.rels
├── docProps/
└── [Content_Types].xml
```

### 2-2. 각주 XML 파싱

`footnotes.xml`을 PowerShell XML 파서로 파싱하여 각주 번호(`w:id`)별 텍스트 및 URL을 추출하였다.

```powershell
[xml]$footnotesXml = Get-Content "docx_unpacked\word\footnotes.xml" -Encoding UTF8
$ns = @{ w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main" }
$footnotes = Select-Xml -Xml $footnotesXml -Namespace $ns -XPath "//w:footnote[@w:id]"
```

각 각주에서 `https?://` 패턴으로 URL을 정규식 추출하였다.  
총 **각주 1번~96번** (일부 빈 각주 제외)에서 **96개 항목** 수집.

---

## 3. URL 검사 방법 (`check_urls.py`)

### 3-1. 봇차단 우회 전략

단순 `requests.get()`은 서버의 봇 필터링에 의해 403 응답을 받을 수 있으므로, 다음 두 가지 브라우저 헤더 프로파일을 순차 시도하였다.

| 순서 | User-Agent | 설명 |
|------|-----------|------|
| 1차 | Chrome 124 / Windows | 주 시도 |
| 2차 | Firefox 124 / macOS | 1차 실패 시 대체 |

공통 헤더:
- `Accept`, `Accept-Language` (ko-KR 우선)
- `Accept-Encoding: gzip, deflate, br`
- `Sec-Fetch-*` 계열 (브라우저 탐색 모드 모사)

### 3-2. SSL 오류 처리

SSL 인증서 오류(`SSLError`) 발생 시 `verify=False`로 재시도하여 인증서 문제와 실제 dead link를 구분하였다.

### 3-3. 중복 URL 처리

동일 URL이 여러 각주에 등장할 경우 최초 검사 결과를 캐시하여 재사용(서버 부하 최소화, 검사 시간 단축).

### 3-4. 요청 간 딜레이

서버 과부하 방지를 위해 새 URL 검사마다 **0.8초** 딜레이 적용.

---

## 4. 검사 결과

### 4-1. 전체 통계

| 결과 | 건수 | 비율 |
|------|------|------|
| ✅ 정상 | 78 | 81.3% |
| ❌ Dead URL (404/410) | 15 | 15.6% |
| ⚠️ 오류 (접속불가/타임아웃) | 3 | 3.1% |
| **합계** | **96** | 100% |

### 4-2. Dead URL 목록

| 각주 번호 | HTTP 코드 | URL | 비고 |
|-----------|-----------|-----|------|
| 1, 21, 62 | 404 | `https://dataon.kisti.re.kr/intro/intro01.do?mm=MENU_00102&sm=MENU00144` | DataON 소개 페이지 |
| 4 | 404 | `https://dataon.kisti.re.kr/intro/intro03.do?mm=MENU_00102&sm=MENU00146` | DataON 개발 연혁 |
| 5 | 404 | `https://dataon.kisti.re.kr/intro/intro02.do?mode=1&mm=MENU_00102&sm=MENU00145` | DataON 서비스 소개 |
| 6 | 404 | `https://dataon.kisti.re.kr/intro/intro07.do?mm=MENU_00102&sm=MENU00200` | 협력 기관 |
| 33, 44 | 404 | `https://dataon.kisti.re.kr/intro/intro08.do` | 조직도 |
| 48 | 404 | `https://kacademy.kisti.re.kr/online-edu/free/620b7cf8be9d081049637166` | KISTI 과학데이터 교육 |
| 50 | 404 | `https://www.scidatacon.org/IDW-2022/sessions/293/` | IDW2022 세션 |
| 51 | 404 | `https://www.rd-alliance.org/about-rda/our-leadership/rda-technical-advisory-board.html` | RDA TAB |
| 66 | 404 | `https://dataon.kisti.re.kr/intro/intro01.do` | 데이터 수집 범위 |
| 67, 94 | 404 | `https://dataon.kisti.re.kr/intro/intro09.do?mm=MENU_00102&sm=MENU00307` | 메타데이터 스키마 |
| 93 | **410** | `https://www.tta.or.kr/tta/ttaSearchView.do?key=77&rep=1&searchStandardNo=TTAK.KO-10.0976&searchCate=TTAS` | TTA 표준 (영구 삭제) |

### 4-3. 오류(접속 불가) 목록

| 각주 번호 | URL | 오류 유형 |
|-----------|-----|-----------|
| 18 | `https://dmp.kisti.re.kr/` | 연결 시간 초과 |
| 70b | `https://dmp.kisti.re.kr/main.do` | 연결 시간 초과 |
| 89b | `https://dmp.kisti.re.kr/main.do` | 연결 시간 초과 |

> **비고**: `dmp.kisti.re.kr`는 검사 당일 서버 자체가 응답하지 않는 상태였다. 실제 폐기 여부는 별도 확인 필요.

---

## 5. 주요 발견 사항 및 해석

### DataON 소개 메뉴 URL 대거 404
`dataon.kisti.re.kr/intro/introXX.do` 형식의 URL이 다수 404로 확인되었다. 이는 DataON 사이트가 **메뉴 구조를 개편**하면서 기존 URL 체계가 변경된 것으로 추정된다. 신청서 작성 당시(2025년 9월)에는 유효했던 URL이 현재는 접근 불가능한 상태이다.

### TTA 표준 참조 URL 영구 삭제 (410)
각주 93번 TTA 표준 검색 URL이 410 응답(Gone)을 반환하였다. 410은 리소스가 서버에서 **의도적으로 제거**되었음을 의미하며 404와 달리 복구 가능성이 없다.

### RDA 링크 일부 사망
각주 51번(RDA Technical Advisory Board 페이지)이 404이나, 각주 52번(RDA OA-Members)은 정상 접근된다. RDA 웹사이트 내부 구조 개편으로 일부 페이지가 이동된 것으로 보인다.

### DOI 링크는 전부 정상
`http://doi.org/10.22711/N` 형식의 DOI 링크는 전부 DataON 내 실제 문서로 리다이렉트되어 정상 접근되었다.

---

## 6. 출력 파일 구성

**파일**: `DataON_URL_Check_Results.xlsx`

| 열 | 내용 |
|----|------|
| A: URL | 각주에서 추출한 전체 URL |
| B: 검사 결과 | 정상 / Dead URL / 오류 / 리다이렉트(정상) / 봇차단 등 |
| C: 서버 반환 코드 | HTTP 상태 코드 (200, 404, 410, N/A 등) |
| D: 엑셀 문서 내 URL 각주 번호 | 원본 문서 내 각주 번호 |
| E: 봇차단 우회를 통한 정상 접근 여부 | 예 / 아니오 / 해당없음 |
| F: 오류의 의미 | 상태 코드 또는 예외에 대한 한국어 설명 |
| G: DataON 내 페이지 여부 | 예 (DataON) / 예 (KISTI) / 아니오 |

**셀 색상 규칙**:
- 🟢 초록: 정상 또는 리다이렉트
- 🔴 빨강: Dead URL 또는 오류
- 🟡 노랑: 경고 (봇차단, 인증필요 등)

---

## 7. 사용 도구 및 환경

| 항목 | 내용 |
|------|------|
| 언어 | Python 3.13, PowerShell 5.1 |
| 주요 라이브러리 | `requests` 2.32.5, `openpyxl` 3.1.5 |
| OS | Windows 11 Pro |
| 검사 방식 | HTTP GET + 다중 User-Agent (봇우회) |
| SSL 처리 | SSLError 발생 시 verify=False 재시도 |
| 요청 타임아웃 | 15초 |
| 요청 간 딜레이 | 0.8초 |
| 최대 리다이렉트 | 10회 |
