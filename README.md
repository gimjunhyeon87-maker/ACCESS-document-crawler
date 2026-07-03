# ACCESS Document Crawler

U.S. Department of Commerce ACCESS 시스템에서 한국 관련 CVD(상계관세) 조사 문서를 
자동으로 수집하는 Selenium 기반 크롤링 스크립트입니다.

## 배경

미국 상무부 ACCESS 시스템에 등록된 24개 한국 CVD 케이스(C-580-602 ~ C-580-989)의 
조사 문서를 수동으로 다운로드하려면 수일이 소요됩니다. 이를 자동화하여 
약 23,000건의 PDF 문서를 효율적으로 수집했습니다.

## 주요 기능

- Selenium으로 키보드·마우스 동작 자동화
- 케이스별 검색 및 PDF 일괄 다운로드
- 체크포인트 저장으로 중단 후 재시작 가능
- 바코드 기반 중복 파일 감지 및 건너뜀
- 손상 파일 복구 및 중복 제거 검증 스크립트
- Excel 메타데이터 자동 정리 (케이스번호, 제출일, 제출자, 문서유형 등)

## 기술 스택

- Python 3.x
- Selenium
- pandas
- openpyxl

## 설치

pip install selenium pandas openpyxl

## 사용 방법

1. ChromeDriver 설치 (Chrome 버전에 맞게)
2. `config.py`에서 다운로드 경로 및 케이스 번호 설정
3. 스크립트 실행 후 브라우저에서 직접 로그인 (CAPTCHA)
4. 로그인 완료 후 터미널에서 Enter → 자동 다운로드 시작

## 파일 구조

```
ACCESS-document-crawler/
├── crawler/
│   ├── main.py           # 메인 크롤러
│   ├── downloader.py     # PDF 다운로드 로직
│   ├── metadata.py       # 메타데이터 정리
│   └── validator.py      # 손상 파일 복구 및 중복 제거
├── config.py             # 케이스 번호, 경로 설정
├── requirements.txt
└── README.md
```

> 본 프로젝트는 실제 업무에서 직접 설계·구현한 자동화 스크립트입니다.
