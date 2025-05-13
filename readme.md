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

- **GUI에서 평균 계산 방식 옵션**: FPS, BW, RTT 각각에 대해 원하는 평균 계산 방식을 콤보박스로 선택 가능하며, 선택한 옵션이 리포트/엑셀/파일명에 모두 반영됩니다.

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

- `reports/folder_metrics_report_fps-minmax_bw-min_rtt-none_20240613_153000.md`: 마크다운 리포트 (옵션이 파일명에 반영됨)
- `reports/folder_metrics_report_fps-minmax_bw-min_rtt-none_20240613_153000.html`: HTML 리포트
- `reports/metrics_details_fps-minmax_bw-min_rtt-none_20240613_153000.xlsx`: 상세 데이터 Excel 파일
- `reports/metrics_averages_fps-minmax_bw-min_rtt-none_20240613_153000.xlsx`: 평균값 및 min/max 포함 Excel 파일
- `reports/plots/`: 성능 차트 이미지 파일들

## 주의사항

- 리포트/엑셀 파일명, 리포트 상단 표, 옵션 설명, min/max 표 등에서 실제 사용된 평균 계산 옵션을 반드시 확인하세요.
- 프로그램은 선택한 폴더와 하위 폴더의 모든 PDF 파일을 처리합니다.
- 언더스코어(_)로 시작하는 폴더는 처리하지 않습니다.
- 처리 시간은 PDF 파일의 수와 크기에 따라 달라질 수 있습니다.

## PDF 데이터 폴더 구조 가이드

이 프로젝트의 PDF 파일들은 아래와 같은 폴더 구조로 정리되어야 합니다.

## 폴더 구조 예시

```
20250503/London/런던시내_빅벤앞/Ireland/ee/4G/sloto/fold6
```

- **20250503** : 측정 날짜 (YYYYMMDD, 8자리)
- **London** : 도시명 (영문)
- **런던시내_빅벤앞** : 지역/세부장소 (한글 또는 영문, 자유롭게)
- **Ireland** : 리전/국가명 (영문)
- **ee** : 통신사명 (영문)
- **4G** : 네트워크 타입 (4G, 5G, WIFI 등)
- **sloto** : 게임/앱 이름
- **s25u** : 디바이스명 또는 추가 구분자

## PDF 파일 위치
- 위 구조의 마지막 폴더(예: `s25u`) 안에 PDF 파일들을 넣어주세요.
- 예시: `20250503/London/런던시내_빅벤앞/Ireland/ee/4G/sloto/s25u/측정결과1.pdf`

## 폴더 구조 규칙
- 날짜 폴더(YYYYMMDD)는 8자리 숫자여야 합니다.
- 각 폴더명은 공백 대신 언더스코어(_)를 사용할 수 있습니다.
- 네트워크 타입 폴더는 반드시 4G, 5G, WIFI 중 하나여야 합니다.
- PDF 파일은 반드시 마지막(가장 하위) 폴더에 위치해야 합니다.

## 예시 트리
```
20250503/
└── London/
    └── 런던시내_빅벤앞/
        └── Ireland/
            └── ee/
                └── 4G/
                    └── sloto/
                        └── s25u/
                            ├── 측정결과1.pdf
                            └── 측정결과2.pdf
```

## 참고
- 폴더 구조가 올바르지 않으면 리포트 생성 시 정보가 누락될 수 있습니다.
- 폴더명은 분석 목적에 맞게 자유롭게 추가/변경할 수 있으나, 날짜~네트워크 타입까지는 반드시 포함되어야 합니다.

이제 data 폴더에 위와 같은 구조로 PDF 파일을 정리하면 됩니다!  
추가로 안내가 필요하면 말씀해 주세요.

## 주요 업데이트 (2024-06)

- **평균 계산 방식 옵션화**: FPS, BW, RTT 각각에 대해 평균 계산 시 '최솟값/최댓값 제외', '최솟값만 제외', '최댓값만 제외', '모두 포함' 중 선택 가능
- **옵션별 파일명 자동 반영**: 리포트/엑셀 파일명에 평균 계산 옵션이 명확히 표시됨 (예: `folder_metrics_report_fps-minmax_bw-min_rtt-none_20240613_153000.md`)
- **옵션 정보 리포트 표기**: 리포트 상단에 실제 사용된 옵션이 표와 설명으로 기록됨
- **폴더별 min/max 값 표기**: 리포트와 metrics_averages.xlsx에 폴더별 min/max 값이 표로 추가됨
- **옵션 config 저장/복원**: GUI에서 선택한 평균 옵션이 config에 저장되어, 프로그램 재실행 시 자동 복원됨