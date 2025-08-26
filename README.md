# 📊 수입신고 RISK 분석 시스템

수입신고 RAW 데이터를 분석하여 다양한 Risk를 식별하고 리포트를 생성하는 Streamlit 웹 애플리케이션입니다.

## 🚀 주요 기능

### 📋 분석 유형
- **Summary**: 전체적인 분석 요약 및 통계
- **8% 환급 검토**: 8% 이상 관세율에 대한 환급 검토 대상 분석(FTA 세율은 고려하지 않음.)
- **0% Risk**: 낮은 관세율 Risk 분석(예: 세번이 잘못되어 CIT 0%로 가지 않았을까?)
- **세율 Risk**: 세번부호 불일치 위험 분석 (동일 규격1인데 2가지 이상의 HS CODE로 분류)
- **단가 Risk**: 단가 변동성 위험 분석 (동일 규격인데 단가 차이가 나는 건)

### 💾 출력 형식
- **Excel 파일**: 모든 분석 결과를 시트별로 정리
- **Word 문서**: 분석 결과 요약 보고서

## 📦 설치 및 실행

### 1. 패키지 설치
```bash
pip install -r requirements.txt
```

### 2. 로컬 실행
```bash
streamlit run streamlit_app.py
```

### 3. 웹 브라우저에서 접속
자동으로 브라우저가 열리거나 `http://localhost:8501`로 직접 접속하세요.

## 🌐 배포 옵션

### Streamlit Cloud (무료)
1. GitHub에 코드 업로드
2. [share.streamlit.io](https://share.streamlit.io)에서 배포
3. GitHub 저장소 연결

### Heroku
1. `Procfile` 생성:
```
web: sh setup.sh && streamlit run streamlit_app.py --server.port=$PORT --server.address=0.0.0.0
```

2. `setup.sh` 생성:
```bash
mkdir -p ~/.streamlit/

echo "\
[general]\n\
email = \"your-email@domain.com\"\n\
" > ~/.streamlit/credentials.toml

echo "\
[server]\n\
headless = true\n\
enableCORS=false\n\
port = $PORT\n\
" > ~/.streamlit/config.toml
```

### Railway/Render
`requirements.txt`와 함께 배포하면 자동으로 인식됩니다.

## 📊 사용법

1. **파일 업로드**: 분석할 수입신고 데이터가 포함된 엑셀 파일을 업로드
2. **분석 옵션 선택**: 사이드바에서 원하는 분석 유형 선택
3. **분석 실행**: '분석 시작' 버튼 클릭
4. **결과 확인**: 탭에서 분석 결과 확인
5. **파일 다운로드**: Excel 및 Word 형태로 결과 다운로드

## 📁 파일 구조

```
Streamlit202508/
├── streamlit_app.py          # 메인 애플리케이션
├── requirements.txt          # 패키지 의존성
├── .streamlit/
│   └── config.toml          # Streamlit 설정
├── README.md                # 이 파일
└── app-new202505-v3.py     # 원본 tkinter 버전 (참고용)
```

## 🔧 시스템 요구사항

- Python 3.7 이상
- 메모리: 최소 512MB (권장: 1GB 이상)
- 업로드 파일 크기: 최대 1000MB

## 📝 주의사항

- 대용량 파일 처리 시 시간이 소요될 수 있습니다
- 인터넷 연결이 필요합니다 (배포된 버전 사용 시)
- 데이터 보안을 위해 민감한 정보는 로컬에서만 처리하는 것을 권장합니다

## 🆘 문제 해결

### 일반적인 오류
1. **패키지 설치 오류**: `pip install --upgrade pip` 후 재시도
2. **메모리 부족**: 작은 데이터셋으로 테스트 후 점진적으로 크기 증가
3. **업로드 오류**: 파일 형식이 .xlsx 또는 .xls인지 확인

### 성능 최적화
- 대용량 파일은 필요한 분석만 선택하여 실행
- 브라우저 캐시 정리로 성능 개선 가능

## 🔄 업데이트 이력

- **v1.0**: 초기 Streamlit 버전 릴리스
- 기존 tkinter 기반 데스크톱 앱을 웹 애플리케이션으로 전환
