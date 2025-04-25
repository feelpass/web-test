# PDF Metrics Analyzer

PDF Metrics Analyzer is a tool for extracting and analyzing performance metrics from PDF test reports. It processes PDF files to extract key metrics such as FPS, bandwidth, and round-trip time (RTT), and generates comprehensive reports summarizing the data by folder.

## Features

- Analyze multiple PDF reports containing performance metrics
- Extract FPS, bandwidth, and RTT values
- Generate folder-based average metrics
- Create detailed markdown reports with summary tables
- User-friendly GUI interface

## Installation

### Windows Users

1. Download the latest release from the [Releases](../../releases) page
2. Run the `PDFMetricsAnalyzer.exe` file - no installation required
3. Create a `data` folder in the same directory as the executable (if not already present)
4. Place your PDF files in the `data` folder or use the app to select a different folder

### For Developers

If you want to run the script from source:

1. Ensure you have Python 3.8 or later installed
2. Clone this repository:
   ```
   git clone https://github.com/yourusername/pdf-metrics-analyzer.git
   cd pdf-metrics-analyzer
   ```
3. Install the required dependencies:
   ```
   pip install -r requirements.txt
   ```
4. Create a `data` folder for your PDF files (this folder is excluded from the repository)
5. Run the application:
   ```
   python app.py
   ```

## Usage

1. Launch the application
2. Click "Select PDF Directory" to choose a folder containing your PDF test reports
3. Click "Process PDF Files" to begin analysis
4. Once processing is complete, click "Open Latest Report" to view the results

## Report Format

The generated report includes:
- Summary table with folder averages for FPS, bandwidth, and RTT
- Detailed breakdown of metrics for each file in each folder

## Building from Source

To build the executable yourself:

1. Install PyInstaller:
   ```
   pip install pyinstaller
   ```
2. Generate the icon:
   ```
   python create_icon.py
   ```
3. Build the executable:
   ```
   pyinstaller --onefile --windowed --icon=app_icon.ico --name=PDFMetricsAnalyzer --add-data "main.py;." app.py
   ```
4. The executable will be available in the `dist` folder

## Note on Data Files

- The `data` folder is excluded from Git version control
- You'll need to create the folder and add your PDF files after cloning the repository
- This keeps the repository clean and avoids storing large binary files in Git

## License

This project is available under the MIT License. See the LICENSE file for more information.

PDF 보고서 프로세서 사용 방법

1. 'data' 폴더에 분석할 PDF 파일을 넣으세요.
2. PDF_Report_Processor.exe를 더블클릭하여 실행하세요.
3. 처리가 완료되면 'reports' 폴더에서 결과 보고서를 찾을 수 있습니다.

주의사항:
- 보고서는 마크다운(.md) 형식으로 생성됩니다.
- 