import os
import sys
import tkinter as tk
from tkinter import messagebox, scrolledtext, ttk
import threading
import subprocess
import importlib.util
from datetime import datetime
import webbrowser

# Check if we're running as a bundled executable or as a Python script
if getattr(sys, "frozen", False):
    # Running as compiled executable
    application_path = os.path.dirname(sys.executable)
    main_script_path = os.path.join(application_path, "main.py")

    # Redirect stdout and stderr to log file when running as exe
    log_file = os.path.join(
        application_path,
        f"pdf_analyzer_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt",
    )
    sys.stdout = open(log_file, "w")
    sys.stderr = open(log_file, "w")
else:
    # Running as script
    application_path = os.path.dirname(os.path.abspath(__file__))
    main_script_path = os.path.join(application_path, "main.py")

# Create default directories if they don't exist
data_dir = os.path.join(application_path, "data")
reports_dir = os.path.join(application_path, "reports")

if not os.path.exists(data_dir):
    os.makedirs(data_dir, exist_ok=True)
    print(f"Created data directory: {data_dir}")

if not os.path.exists(reports_dir):
    os.makedirs(reports_dir, exist_ok=True)
    print(f"Created reports directory: {reports_dir}")

# Import main.py functions directly if possible
try:
    spec = importlib.util.spec_from_file_location("main", main_script_path)
    main_module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(main_module)
    find_pdf_files = main_module.find_pdf_files
    extract_text_from_pdf = main_module.extract_text_from_pdf
    parse_pdf_content = main_module.parse_pdf_content
    get_folder_path = main_module.get_folder_path
    generate_folder_report = main_module.generate_folder_report
    main_function = main_module.main
    direct_import = True
except:
    # Fallback to subprocess if import fails
    direct_import = False


class PDFMetricsApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Metrics Analyzer")
        self.root.geometry("1000x700")
        self.root.minsize(800, 600)

        # Configure the grid layout
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=0)  # Header
        self.root.rowconfigure(1, weight=1)  # Content area
        self.root.rowconfigure(2, weight=0)  # Footer buttons

        # Create header frame
        self.header_frame = tk.Frame(root, bg="#4a7abc", padx=10, pady=10)
        self.header_frame.grid(row=0, column=0, sticky="ew")

        self.title_label = tk.Label(
            self.header_frame,
            text="PDF Metrics Analyzer",
            font=("Arial", 16, "bold"),
            bg="#4a7abc",
            fg="white",
        )
        self.title_label.pack(side=tk.LEFT)

        # Create content frame with two panes
        self.content_frame = tk.Frame(root, padx=10, pady=10)
        self.content_frame.grid(row=1, column=0, sticky="nsew")
        self.content_frame.columnconfigure(0, weight=1)
        self.content_frame.columnconfigure(1, weight=1)
        self.content_frame.rowconfigure(0, weight=0)  # Labels
        self.content_frame.rowconfigure(1, weight=1)  # Text areas

        # Progress section
        self.progress_label = tk.Label(
            self.content_frame,
            text="PDF Processing Log",
            font=("Arial", 12, "bold"),
            anchor="w",
        )
        self.progress_label.grid(row=0, column=0, sticky="w", padx=5, pady=5)

        self.console = scrolledtext.ScrolledText(
            self.content_frame, wrap=tk.WORD, background="#f0f0f0", height=20
        )
        self.console.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)
        self.console.config(state=tk.DISABLED)

        # Results section
        self.results_label = tk.Label(
            self.content_frame,
            text="Report Preview",
            font=("Arial", 12, "bold"),
            anchor="w",
        )
        self.results_label.grid(row=0, column=1, sticky="w", padx=5, pady=5)

        self.report_display = scrolledtext.ScrolledText(
            self.content_frame, wrap=tk.WORD, background="#ffffff", height=20
        )
        self.report_display.grid(row=1, column=1, sticky="nsew", padx=5, pady=5)
        self.report_display.config(state=tk.DISABLED)

        # Progress bar
        self.progress_bar = ttk.Progressbar(self.content_frame, mode="indeterminate")
        self.progress_bar.grid(
            row=2, column=0, columnspan=2, sticky="ew", padx=5, pady=5
        )

        # Status label
        self.status_label = tk.Label(
            self.content_frame,
            text="Ready to process PDF files from data folder.",
            font=("Arial", 10),
            anchor="w",
        )
        self.status_label.grid(
            row=3, column=0, columnspan=2, sticky="ew", padx=5, pady=5
        )

        # Create buttons frame
        self.buttons_frame = tk.Frame(root, padx=10, pady=10)
        self.buttons_frame.grid(row=2, column=0, sticky="ew")

        self.process_button = tk.Button(
            self.buttons_frame,
            text="Process PDF Files",
            command=self.process_pdfs,
            bg="#2196F3",
            fg="white",
            padx=20,
            pady=10,
            font=("Arial", 12),
        )
        self.process_button.pack(side=tk.LEFT, padx=5)

        self.open_report_button = tk.Button(
            self.buttons_frame,
            text="Open Full Report",
            command=self.open_report,
            bg="#FF9800",
            fg="white",
            padx=20,
            pady=10,
            font=("Arial", 12),
            state=tk.DISABLED,
        )
        self.open_report_button.pack(side=tk.LEFT, padx=5)

        self.open_folder_button = tk.Button(
            self.buttons_frame,
            text="Open Data Folder",
            command=self.open_data_folder,
            bg="#4CAF50",
            fg="white",
            padx=20,
            pady=10,
            font=("Arial", 12),
        )
        self.open_folder_button.pack(side=tk.LEFT, padx=5)

        self.exit_button = tk.Button(
            self.buttons_frame,
            text="Exit",
            command=self.root.destroy,
            bg="#f44336",
            fg="white",
            padx=20,
            pady=10,
            font=("Arial", 12),
        )
        self.exit_button.pack(side=tk.RIGHT, padx=5)

        # Initialize variables
        self.selected_dir = data_dir  # Always use the data directory
        self.latest_report = ""
        self.processing = False

        # Handle window close
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        # Redirect stdout to the console
        self.redirect_stdout()

        # Update status with information about data directory
        self.status_label.config(
            text=f"Ready to process PDF files from: {self.selected_dir}"
        )

        # Auto-check for PDF files in data directory
        self.check_pdf_files()

    def check_pdf_files(self):
        """Check if there are PDF files in the data directory and update status."""
        try:
            if os.path.exists(self.selected_dir):
                pdf_count = len(
                    [f for f in os.listdir(self.selected_dir) if f.endswith(".pdf")]
                )
                if pdf_count > 0:
                    self.status_label.config(
                        text=f"Found {pdf_count} PDF files in data folder. Ready to process."
                    )
                else:
                    self.status_label.config(
                        text="No PDF files found in data folder. Please add PDF files before processing."
                    )
                    self.update_console(
                        "No PDF files found in data folder. Please add PDF files before processing.\n"
                    )
        except Exception as e:
            self.status_label.config(text=f"Error checking data folder: {str(e)}")

    def redirect_stdout(self):
        class ConsoleRedirector:
            def __init__(self, text_widget):
                self.text_widget = text_widget
                self.buffer = ""

            def write(self, string):
                self.buffer += string
                self.text_widget.config(state=tk.NORMAL)
                self.text_widget.insert(tk.END, string)
                self.text_widget.see(tk.END)
                self.text_widget.config(state=tk.DISABLED)

            def flush(self):
                pass

        # Only redirect if not already redirected to a file (when bundled)
        if not getattr(sys, "frozen", False):
            self.old_stdout = sys.stdout
            sys.stdout = ConsoleRedirector(self.console)

    def restore_stdout(self):
        if hasattr(self, "old_stdout"):
            sys.stdout = self.old_stdout

    def update_console(self, message):
        self.console.config(state=tk.NORMAL)
        self.console.insert(tk.END, message)
        self.console.see(tk.END)
        self.console.config(state=tk.DISABLED)

    def update_report_display(self, content):
        self.report_display.config(state=tk.NORMAL)
        self.report_display.delete(1.0, tk.END)
        self.report_display.insert(tk.END, content)
        self.report_display.see(tk.END)
        self.report_display.config(state=tk.DISABLED)

    def process_pdfs(self):
        if self.processing:
            messagebox.showwarning(
                "Processing in Progress", "PDF processing is already running."
            )
            return

        # Always use the data directory
        if not os.path.exists(self.selected_dir):
            messagebox.showwarning(
                "No Data Directory",
                "The data directory does not exist. It will be created, but you need to add PDF files to it.",
            )
            os.makedirs(self.selected_dir, exist_ok=True)
            return

        self.processing = True
        self.process_button.config(state=tk.DISABLED)
        self.status_label.config(text="Processing PDF files... Please wait.")

        # Clear console and report display
        self.console.config(state=tk.NORMAL)
        self.console.delete(1.0, tk.END)
        self.console.config(state=tk.DISABLED)

        self.report_display.config(state=tk.NORMAL)
        self.report_display.delete(1.0, tk.END)
        self.report_display.config(state=tk.DISABLED)

        # Start progress bar
        self.progress_bar.start(10)

        # Start processing in a separate thread
        threading.Thread(target=self._process_thread, daemon=True).start()

    def _process_thread(self):
        try:
            self.update_console("Starting PDF processing...\n")

            # Create reports directory if it doesn't exist
            os.makedirs("reports", exist_ok=True)

            if direct_import:
                # Use direct import method
                original_dir = os.getcwd()
                os.chdir(self.selected_dir)

                # Find PDF files
                pdf_files = find_pdf_files(".")
                self.update_console(f"Found {len(pdf_files)} PDF files to process.\n")

                if not pdf_files:
                    self.update_console("\nNo PDF files found in the data directory.\n")
                    self.update_console(
                        "Please add PDF files to the 'data' folder and try again.\n"
                    )
                    return

                # Group data by folder
                from collections import defaultdict

                folder_data = defaultdict(list)

                # Process each PDF file
                for i, pdf_path in enumerate(pdf_files):
                    try:
                        self.update_console(f"Processing: {pdf_path}\n")

                        # Extract text from PDF
                        text = extract_text_from_pdf(pdf_path)

                        if text:
                            # Parse content to get metrics
                            data = parse_pdf_content(text, pdf_path)

                            # Get the folder path
                            folder_path = get_folder_path(pdf_path)

                            # Add data to the folder's collection
                            folder_data[folder_path].append(data)

                            # Update status with progress percentage
                            progress_pct = int((i + 1) / len(pdf_files) * 100)
                            self.root.after(
                                0,
                                lambda p=progress_pct: self.status_label.config(
                                    text=f"Processing: {p}% complete ({i + 1}/{len(pdf_files)} files)"
                                ),
                            )
                        else:
                            self.update_console(
                                f"Warning: No text extracted from {pdf_path}\n"
                            )

                    except Exception as e:
                        self.update_console(f"Error processing {pdf_path}: {e}\n")

                # Change back to original directory for report generation
                os.chdir(original_dir)

                # Generate report of folder metrics
                if folder_data:
                    report_path = generate_folder_report(folder_data)
                    self.latest_report = os.path.abspath(report_path)
                    self.update_console(
                        f"\nGenerated folder metrics report: {self.latest_report}\n"
                    )

                    # Read report content and display in the report panel
                    try:
                        with open(self.latest_report, "r") as f:
                            report_content = f.read()
                            self.root.after(
                                0, lambda: self.update_report_display(report_content)
                            )
                    except Exception as e:
                        self.update_console(f"Error reading report: {e}\n")
                else:
                    self.update_console(
                        "No data was processed successfully. Report not generated.\n"
                    )
            else:
                # Use subprocess method
                process = subprocess.Popen(
                    [sys.executable, main_script_path],
                    stdout=subprocess.PIPE,
                    stderr=subprocess.STDOUT,
                    universal_newlines=True,
                    cwd=self.selected_dir,
                )

                output_lines = []
                while True:
                    line = process.stdout.readline()
                    if not line and process.poll() is not None:
                        break
                    output_lines.append(line)
                    self.update_console(line)

                # Find the latest report file
                report_dir = os.path.join(self.selected_dir, "reports")
                if os.path.exists(report_dir):
                    report_files = [
                        os.path.join(report_dir, f)
                        for f in os.listdir(report_dir)
                        if f.endswith(".md")
                    ]
                    if report_files:
                        self.latest_report = max(report_files, key=os.path.getmtime)

                        # Read report content and display in the report panel
                        try:
                            with open(self.latest_report, "r") as f:
                                report_content = f.read()
                                self.root.after(
                                    0,
                                    lambda: self.update_report_display(report_content),
                                )
                        except Exception as e:
                            self.update_console(f"Error reading report: {e}\n")

            self.update_console("\nProcessing complete!\n")

            # Enable the open report button if we have a report
            if self.latest_report:
                self.root.after(
                    0, lambda: self.open_report_button.config(state=tk.NORMAL)
                )

        except Exception as e:
            self.update_console(f"Error during processing: {str(e)}\n")
        finally:
            # Stop progress bar
            self.root.after(0, self.progress_bar.stop)

            # Re-enable buttons
            self.root.after(0, lambda: self.process_button.config(state=tk.NORMAL))

            # Update status
            if self.latest_report:
                self.root.after(
                    0,
                    lambda: self.status_label.config(
                        text=f"Processing complete. Report generated at: {self.latest_report}"
                    ),
                )
            else:
                self.root.after(
                    0,
                    lambda: self.status_label.config(
                        text="Processing complete. No report was generated."
                    ),
                )

            self.processing = False

    def open_report(self):
        if not self.latest_report or not os.path.exists(self.latest_report):
            messagebox.showwarning(
                "Report Not Found",
                "No report file found. Please process PDF files first.",
            )
            return

        # Open the report file with the default application
        if sys.platform == "win32":
            os.startfile(self.latest_report)
        elif sys.platform == "darwin":  # macOS
            subprocess.call(["open", self.latest_report])
        else:  # Linux
            subprocess.call(["xdg-open", self.latest_report])

    def open_data_folder(self):
        """Open the data folder in file explorer"""
        if not os.path.exists(self.selected_dir):
            os.makedirs(self.selected_dir, exist_ok=True)

        if sys.platform == "win32":
            os.startfile(self.selected_dir)
        elif sys.platform == "darwin":  # macOS
            subprocess.call(["open", self.selected_dir])
        else:  # Linux
            subprocess.call(["xdg-open", self.selected_dir])

    def on_closing(self):
        if self.processing:
            if messagebox.askyesno(
                "Quit", "Processing is in progress. Are you sure you want to quit?"
            ):
                self.restore_stdout()
                self.root.destroy()
        else:
            self.restore_stdout()
            self.root.destroy()


if __name__ == "__main__":
    # Create the main window
    root = tk.Tk()
    app = PDFMetricsApp(root)

    # Set icon if available
    icon_path = os.path.join(application_path, "app_icon.ico")
    if os.path.exists(icon_path):
        try:
            root.iconbitmap(icon_path)
        except:
            pass

    # Start the application
    root.mainloop()
