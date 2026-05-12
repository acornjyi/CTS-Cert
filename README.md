# CTS-Cert — DataON CoreTrustSeal 갱신 작업

DataON의 CoreTrustSeal(CTS) 재인증을 위한 작업 스크립트 및 결과물 저장소입니다.

## 구성

```
.
├── check_urls.py                    # Dead URL 검사 스크립트
├── DataON_URL_Check_Results.xlsx   # URL 검사 결과 (엑셀)
├── Docs/
│   └── URL_검사_작업_문서.md       # 작업 방법 및 결과 문서
└── README.md
```

## Dead URL 검사 (`check_urls.py`)

`DataON CoreTrustSeal Application_Update_250912.docx`에 포함된 각주 URL 96개를 대상으로 dead link 여부를 자동 검사합니다.

### 실행 방법

```bash
pip install requests openpyxl
python check_urls.py
```

결과 파일 `DataON_URL_Check_Results.xlsx`가 동일 디렉토리에 생성됩니다.

### 주요 기능

- 브라우저 User-Agent 모사를 통한 봇차단 우회
- SSL 인증서 오류 자동 우회 재시도
- 중복 URL 캐싱으로 서버 부하 최소화
- HTTP 상태 코드별 결과 분류 및 색상 표시

### 검사 결과 요약 (2026-05-12 기준)

| 결과 | 건수 |
|------|------|
| ✅ 정상 | 78 |
| ❌ Dead URL (404/410) | 15 |
| ⚠️ 오류 (접속불가) | 3 |
| **합계** | **96** |

주요 Dead URL: `dataon.kisti.re.kr/intro/introXX.do` 계열 (사이트 개편으로 인한 URL 변경 추정), TTA 표준 참조 URL(410), RDA TAB 페이지(404)

자세한 내용은 [Docs/URL_검사_작업_문서.md](Docs/URL_검사_작업_문서.md) 참고.

## 관련 링크

- [DataON](https://dataon.kisti.re.kr)
- [CoreTrustSeal](https://www.coretrustseal.org/)
