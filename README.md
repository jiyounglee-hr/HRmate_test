# 인사팀 대시보드

인사팀을 위한 데이터 시각화 대시보드입니다.

## 기능

- 현재 인원 현황 표시
- 부서별 인원 통계
- 직급별 인원 통계
- 연도별 인원 추이

## 설치 방법

1. Python 3.8 이상이 설치되어 있어야 합니다.
2. 필요한 패키지 설치:
```bash
pip install -r requirements.txt
```

## 실행 방법

다음 명령어로 애플리케이션을 실행합니다:
```bash
streamlit run app.py
```

## 데이터 업데이트

1. Excel 파일을 최신 데이터로 업데이트
2. 앱 새로고침

## 주의사항

- Excel 파일 형식을 유지해주세요
- 데이터 업데이트 시 열 이름을 변경하지 마세요

## 데이터 형식

다음 열들이 포함된 엑셀 파일을 준비해주세요:
- 부서: 부서명
- 직급: 직급 정보
- 입사일: YYYY-MM-DD 형식
- 생년월일: YYYY-MM-DD 형식

위 열들이 모두 없어도 실행은 가능하며, 있는 데이터만 시각화됩니다. 