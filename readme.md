# PDF Metrics Analyzer

PDF Metrics Analyzer는 PDF 테스트 보고서에서 성능 지표를 추출하고 분석하는 도구입니다. PDF 파일을 처리하여 FPS, 대역폭, RTT(Round-Trip Time)와 같은 핵심 지표를 추출하고 폴더별로 포괄적인 보고서를 생성합니다.

## 주요 기능

- 성능 지표가 포함된 다수의 PDF 보고서 분석
- FPS, 대역폭, RTT 값 추출
- 폴더 기반 평균 지표 생성
- 요약 테이블이 포함된 상세 마크다운 보고서 생성
- 자동 보고서 생성

## 설치 방법

### Windows 사용자

1. [Releases](../../releases) 페이지에서 최신 버전을 다운로드하세요
2. `PDFMetricsAnalyzer.exe` 파일을 원하는 위치에 저장하세요 - 설치 필요 없음
3. 실행 파일과 같은 위치에 `data` 폴더를 생성하세요 (실행 시 없으면 자동 생성됨)
4. PDF 파일을 `data` 폴더에 넣으세요

### 개발자를 위한 설치

소스에서 스크립트를 실행하려면:

1. Python 3.8 이상이 설치되어 있는지 확인하세요
2. 이 저장소를 복제하세요:
   ```
   git clone https://github.com/yourusername/pdf-metrics-analyzer.git
   cd pdf-metrics-analyzer
   ```
3. 필요한 종속성을 설치하세요:
   ```
   pip install -r requirements.txt
   ```
4. 스크립트를 실행하세요:
   ```
   python main.py
   ```

## 사용 방법

1. PDF 파일을 `data` 폴더에 넣으세요
2. `PDFMetricsAnalyzer.exe`를 더블클릭하여 실행하세요
3. 프로그램이 자동으로 데이터 폴더의 모든 PDF 파일을 처리합니다
4. 처리가 완료되면 `reports` 폴더에 생성된 보고서를 확인하세요

## 보고서 형식

생성된 보고서는 다음을 포함합니다:
- FPS, 대역폭, RTT에 대한 폴더 평균이 포함된 요약 테이블
- 각 폴더의 각 파일에 대한 지표의 상세 내역
- 보고서는 `reports/folder_metrics_report.md` 파일로 저장됩니다

## 소스에서 빌드하기

직접 실행 파일을 빌드하려면:

1. PyInstaller 설치:
   ```
   pip install pyinstaller
   ```
2. 실행 파일 빌드:
   ```
   pyinstaller --onefile --name=PDFMetricsAnalyzer main.py
   ```
3. 실행 파일은 `dist` 폴더에서 사용 가능합니다

## 라이선스

이 프로젝트는 MIT 라이선스에 따라 사용할 수 있습니다. 자세한 내용은 LICENSE 파일을 참조하세요.