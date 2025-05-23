import sys
import os
from datetime import datetime
from PySide6.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QCheckBox,
    QProgressBar,
    QMessageBox,
    QTextEdit,
    QScrollBar,
    QFileDialog,
    QLineEdit,
    QComboBox,
)
from PySide6.QtCore import Qt, QThread, Signal
from main import (
    process_pdf_files,
    generate_folder_report,
    export_to_excel,
    generate_performance_plots,
    get_last_folder,
    save_last_folder,
)


class ProcessorThread(QThread):
    finished = Signal(tuple)
    error = Signal(str)
    progress = Signal(str)
    log = Signal(str)
    progress_count = Signal(int, int)  # (current, total)

    def __init__(self, excel_enabled, plots_enabled, root_dir=".", avg_mode=None):
        super().__init__()
        self.excel_enabled = excel_enabled
        self.plots_enabled = plots_enabled
        self.root_dir = root_dir
        self.avg_mode = avg_mode or {"fps": "minmax", "bw": "minmax", "rtt": "minmax"}

    def run(self):
        try:
            # PDF 파일 처리
            self.progress.emit("PDF 파일 처리 중...")

            def progress_callback(current, total):
                self.progress_count.emit(current, total)

            folder_data = process_pdf_files(
                self.root_dir,
                log_callback=self.log.emit,
                progress_callback=progress_callback,
            )

            if not folder_data:
                self.error.emit("PDF 파일 처리 중 오류가 발생했습니다.")
                return

            # reports_dir 가져오기
            reports_dir = folder_data.get("_config", {}).get("reports_dir")
            if not reports_dir:
                self.error.emit("리포트 디렉토리 정보를 찾을 수 없습니다.")
                return

            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            reports_generated = []

            # 마크다운 및 HTML 리포트 생성
            self.progress.emit("마크다운 및 HTML 리포트 생성 중...")
            fixed_md, timestamped_md, fixed_html, timestamped_html = (
                generate_folder_report(
                    folder_data, log_callback=self.log.emit, exclude_mode=self.avg_mode
                )
            )
            if fixed_md:
                reports_generated.append("마크다운/HTML 리포트")

            # Excel 리포트 생성
            if self.excel_enabled:
                self.progress.emit("Excel 리포트 생성 중...")
                details_excel, averages_excel = export_to_excel(
                    folder_data,
                    reports_dir=reports_dir,
                    timestamp=timestamp,
                    log_callback=self.log.emit,
                    exclude_mode=self.avg_mode,
                )
                if details_excel:
                    reports_generated.append("Excel 리포트")

            # 성능 차트 생성
            if self.plots_enabled:
                self.progress.emit("성능 차트 생성 중...")
                generate_performance_plots(
                    folder_data,
                    reports_dir=reports_dir,
                    timestamp=timestamp,
                    log_callback=self.log.emit,
                )
                reports_generated.append("성능 차트")

            self.finished.emit((True, reports_generated, reports_dir))

        except Exception as e:
            self.error.emit(str(e))


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDF 리포트 생성기")
        self.setMinimumSize(800, 600)

        # 메인 위젯 및 레이아웃 설정
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        layout = QVBoxLayout(main_widget)

        # 상단 설명 레이블
        description = QLabel(
            "이 프로그램은 PDF 파일에서 성능 메트릭을 추출하고 리포트를 생성합니다."
        )
        description.setWordWrap(True)
        layout.addWidget(description)

        # 폴더 선택 영역
        folder_layout = QHBoxLayout()
        self.folder_path = QLineEdit()
        self.folder_path.setReadOnly(True)
        self.folder_path.setPlaceholderText("PDF 파일을 검색할 폴더를 선택하세요")

        folder_layout.addWidget(self.folder_path)

        select_folder_button = QPushButton("폴더 선택")
        select_folder_button.clicked.connect(self.select_folder)
        folder_layout.addWidget(select_folder_button)
        layout.addLayout(folder_layout)

        # 옵션 체크박스들
        options_layout = QHBoxLayout()

        self.excel_checkbox = QCheckBox("Excel 리포트 생성")
        self.excel_checkbox.setChecked(True)
        options_layout.addWidget(self.excel_checkbox)

        self.plots_checkbox = QCheckBox("성능 차트 생성")
        self.plots_checkbox.setChecked(False)
        self.plots_checkbox.setVisible(False)
        options_layout.addWidget(self.plots_checkbox)

        # 평균 계산 방식 콤보박스 추가 (FPS, BW, RTT 각각)
        self.fps_avg_combo = QComboBox()
        self.fps_avg_combo.addItem("최솟값/최댓값 제외", "minmax")
        self.fps_avg_combo.addItem("최솟값만 제외", "min")
        self.fps_avg_combo.addItem("최댓값만 제외", "max")
        self.fps_avg_combo.addItem("모두 포함", "none")
        options_layout.addWidget(QLabel("FPS 평균:"))
        options_layout.addWidget(self.fps_avg_combo)

        self.bw_avg_combo = QComboBox()
        self.bw_avg_combo.addItem("최솟값/최댓값 제외", "minmax")
        self.bw_avg_combo.addItem("최솟값만 제외", "min")
        self.bw_avg_combo.addItem("최댓값만 제외", "max")
        self.bw_avg_combo.addItem("모두 포함", "none")
        options_layout.addWidget(QLabel("BW 평균:"))
        options_layout.addWidget(self.bw_avg_combo)

        self.rtt_avg_combo = QComboBox()
        self.rtt_avg_combo.addItem("최솟값/최댓값 제외", "minmax")
        self.rtt_avg_combo.addItem("최솟값만 제외", "min")
        self.rtt_avg_combo.addItem("최댓값만 제외", "max")
        self.rtt_avg_combo.addItem("모두 포함", "none")
        options_layout.addWidget(QLabel("RTT 평균:"))
        options_layout.addWidget(self.rtt_avg_combo)

        layout.addLayout(options_layout)

        # 마지막으로 선택한 폴더와 옵션 로드 (콤보박스 생성 이후에 해야 함)
        last_folder, last_avg_mode = get_last_folder()
        if last_folder:
            self.folder_path.setText(last_folder)
        # 평균 옵션도 불러와서 콤보박스에 반영
        if last_avg_mode:
            idx = self.fps_avg_combo.findData(last_avg_mode.get("fps", "minmax"))
            if idx >= 0:
                self.fps_avg_combo.setCurrentIndex(idx)
            idx = self.bw_avg_combo.findData(last_avg_mode.get("bw", "minmax"))
            if idx >= 0:
                self.bw_avg_combo.setCurrentIndex(idx)
            idx = self.rtt_avg_combo.findData(last_avg_mode.get("rtt", "minmax"))
            if idx >= 0:
                self.rtt_avg_combo.setCurrentIndex(idx)

        # 로그 출력 영역
        self.log_output = QTextEdit()
        self.log_output.setReadOnly(True)
        self.log_output.setStyleSheet(
            """
            QTextEdit {
                background-color: #1E1E1E;
                color: #FFFFFF;
                font-family: 'Courier New', monospace;
                font-size: 12px;
                padding: 10px;
            }
        """
        )
        layout.addWidget(self.log_output)

        # 진행 상태 표시줄
        self.progress_label = QLabel("대기 중...")
        layout.addWidget(self.progress_label)
        self.progress_bar = QProgressBar()
        self.progress_bar.setMinimum(0)
        self.progress_bar.setValue(0)
        layout.addWidget(self.progress_bar)

        # 시작 버튼
        self.start_button = QPushButton("리포트 생성 시작")
        self.start_button.clicked.connect(self.start_processing)
        layout.addWidget(self.start_button)

        # 프로세서 스레드 초기화
        self.processor_thread = None

    def select_folder(self):
        folder = QFileDialog.getExistingDirectory(
            self,
            "PDF 파일을 검색할 폴더 선택",
            self.folder_path.text()
            or os.path.expanduser("~"),  # 현재 선택된 폴더나 홈 디렉토리에서 시작
            QFileDialog.ShowDirsOnly | QFileDialog.DontResolveSymlinks,
        )
        if folder:
            self.folder_path.setText(folder)
            # 폴더와 옵션 저장
            avg_mode = {
                "fps": self.fps_avg_combo.currentData(),
                "bw": self.bw_avg_combo.currentData(),
                "rtt": self.rtt_avg_combo.currentData(),
            }
            save_last_folder(folder, avg_mode)  # 선택한 폴더와 옵션 저장

    def start_processing(self):
        if not self.folder_path.text():
            QMessageBox.warning(self, "경고", "폴더를 선택해주세요.")
            return

        self.start_button.setEnabled(False)
        self.log_output.clear()
        self.progress_label.setText("처리 중...")
        self.progress_bar.setValue(0)
        self.progress_bar.setMaximum(0)  # Indeterminate until we know total

        # 평균 계산 방식 값 추출 (dict)
        avg_mode = {
            "fps": self.fps_avg_combo.currentData(),
            "bw": self.bw_avg_combo.currentData(),
            "rtt": self.rtt_avg_combo.currentData(),
        }
        # 폴더와 옵션 저장
        save_last_folder(self.folder_path.text(), avg_mode)

        # 새 프로세서 스레드 생성
        self.processor_thread = ProcessorThread(
            self.excel_checkbox.isChecked(),
            self.plots_checkbox.isChecked(),
            self.folder_path.text(),
            avg_mode,
        )

        # 시그널 연결
        self.processor_thread.progress.connect(self.update_progress)
        self.processor_thread.error.connect(self.handle_error)
        self.processor_thread.finished.connect(self.handle_completion)
        self.processor_thread.log.connect(self.append_log)
        self.processor_thread.progress_count.connect(self.update_progress_count)

        # 스레드 시작
        self.processor_thread.start()

    def append_log(self, message):
        self.log_output.append(message)
        # 자동 스크롤
        scrollbar = self.log_output.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())

    def update_progress(self, message):
        self.progress_label.setText(message)

    def handle_error(self, error_message):
        self.progress_label.setText("오류 발생")
        QMessageBox.critical(
            self, "오류", f"처리 중 오류가 발생했습니다:\n{error_message}"
        )
        self.start_button.setEnabled(True)

    def handle_completion(self, result):
        success, reports, reports_dir = result
        self.start_button.setEnabled(True)

        if success and reports:
            report_types = ", ".join(reports)
            QMessageBox.information(
                self,
                "완료",
                f"처리가 완료되었습니다.\n\n"
                f"생성된 리포트: {report_types}\n"
                f"저장 위치: {reports_dir}",
            )
        else:
            QMessageBox.warning(self, "경고", "처리된 파일이 없습니다.")

    def update_progress_count(self, current, total):
        self.progress_label.setText(f"{total}개 중 {current}개 처리 중...")
        self.progress_bar.setMaximum(total)
        self.progress_bar.setValue(current)


def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")  # 모던한 스타일 적용
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
