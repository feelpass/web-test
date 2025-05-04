# PDF Performance Report Generator

PDF 파일에서 성능 데이터를 추출하고 분석하여 리포트를 생성하는 프로그램입니다.

## 기능

- PDF 파일에서 성능 데이터 자동 추출
- 마크다운 및 HTML 형식의 리포트 생성
- Excel 파일 생성 (상세 데이터 및 평균값)
- 성능 차트 생성 (지역별, 통신사별)
- 사용하기 쉬운 그래픽 사용자 인터페이스 (GUI)

## 설치 방법

### 1. 실행 파일 다운로드 (일반 사용자)

1. [Releases](releases) 페이지에서 최신 버전의 실행 파일을 다운로드합니다.
2. 다운로드한 파일을 원하는 위치에 압축 해제합니다.
3. `PDFReportGenerator.app` (macOS) 또는 `PDFReportGenerator.exe` (Windows)를 실행합니다.

### 2. 소스코드로 실행 (개발자)

1. 필요한 패키지 설치:
```bash
pip install -r requirements.txt
```

2. GUI 버전 실행:
```bash
python main_gui.py
```

3. 명령줄 버전 실행:
```bash
# 기본 실행 (폴더 지정 필수)
python main.py --folder <폴더경로>

# 모든 옵션 사용
python main.py --folder <폴더경로> --excel --plots
```

### 명령줄 인자 설명

`main.py`는 다음과 같은 명령줄 인자를 지원합니다:

- `--folder`: (필수) PDF 파일이 있는 폴더 경로를 지정합니다.
  ```bash
  python main.py --folder ./data
  ```

- `--excel`: (선택) Excel 리포트를 생성합니다. 상세 데이터와 평균값 파일이 생성됩니다.
  ```bash
  python main.py --folder ./data --excel
  ```

- `--plots`: (선택) 성능 차트를 생성합니다. 지역별, 통신사별 성능 분석 차트가 생성됩니다.
  ```bash
  python main.py --folder ./data --plots
  ```

모든 옵션을 함께 사용할 수 있습니다:
```bash
python main.py --folder ./data --excel --plots
```

## 실행 파일 빌드 방법

### macOS
```bash
# GUI 앱 빌드 (앱 번들 생성)
pyinstaller --windowed --onedir --name "PDFReportGenerator" --add-data "main.py:." --noconfirm main_gui.py
```

### Windows
```bash
# GUI 앱 빌드 (실행 파일 생성)
pyinstaller --windowed --onefile --name "PDFReportGenerator" --add-data "main.py;." main_gui.py
```

생성된 실행 파일은 `dist` 폴더에서 찾을 수 있습니다:
- macOS: `dist/PDFReportGenerator.app`
- Windows: `dist/PDFReportGenerator.exe`

## 사용 방법

1. 프로그램을 실행합니다.
2. "폴더 선택" 버튼을 클릭하여 PDF 파일이 있는 폴더를 선택합니다.
3. 원하는 출력 옵션을 선택합니다:
   - Excel 파일 생성
   - 성능 차트 생성
4. "리포트 생성" 버튼을 클릭합니다.
5. 생성된 파일은 `reports` 폴더에서 확인할 수 있습니다.

## 생성되는 파일

- `reports/folder_metrics_report.md`: 마크다운 형식의 리포트
- `reports/folder_metrics_report.html`: HTML 형식의 리포트
- `reports/metrics_details_[날짜시간].xlsx`: 상세 데이터 Excel 파일
- `reports/metrics_averages_[날짜시간].xlsx`: 평균값 데이터 Excel 파일
- `reports/plots/`: 성능 차트 이미지 파일들

## 주의사항

- 프로그램은 선택한 폴더와 하위 폴더의 모든 PDF 파일을 처리합니다.
- 언더스코어(_)로 시작하는 폴더는 처리하지 않습니다.
- 처리 시간은 PDF 파일의 수와 크기에 따라 달라질 수 있습니다.