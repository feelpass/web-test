import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import threading
import subprocess
import importlib.util
from datetime import datetime

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
        self.root.geometry("800x600")
        self.root.minsize(600, 400)

        # Configure the grid layout
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=0)  # Header
        self.root.rowconfigure(1, weight=1)  # Console output
        self.root.rowconfigure(2, weight=0)  # Buttons

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

        # Create console output area
        self.console_frame = tk.Frame(root, padx=10, pady=10)
        self.console_frame.grid(row=1, column=0, sticky="nsew")

        self.console_frame.columnconfigure(0, weight=1)
        self.console_frame.rowconfigure(0, weight=0)
        self.console_frame.rowconfigure(1, weight=1)

        self.status_label = tk.Label(
            self.console_frame,
            text="Ready to process PDF files.",
            font=("Arial", 10),
            anchor="w",
        )
        self.status_label.grid(row=0, column=0, sticky="ew")

        self.console = scrolledtext.ScrolledText(
            self.console_frame, wrap=tk.WORD, background="#f0f0f0", height=20
        )
        self.console.grid(row=1, column=0, sticky="nsew")
        self.console.config(state=tk.DISABLED)

        # Create buttons frame
        self.buttons_frame = tk.Frame(root, padx=10, pady=10)
        self.buttons_frame.grid(row=2, column=0, sticky="ew")

        self.select_dir_button = tk.Button(
            self.buttons_frame,
            text="Select PDF Directory",
            command=self.select_directory,
            bg="#4CAF50",
            fg="white",
            padx=10,
            pady=5,
        )
        self.select_dir_button.pack(side=tk.LEFT, padx=5)

        self.process_button = tk.Button(
            self.buttons_frame,
            text="Process PDF Files",
            command=self.process_pdfs,
            bg="#2196F3",
            fg="white",
            padx=10,
            pady=5,
        )
        self.process_button.pack(side=tk.LEFT, padx=5)

        self.open_report_button = tk.Button(
            self.buttons_frame,
            text="Open Latest Report",
            command=self.open_report,
            bg="#FF9800",
            fg="white",
            padx=10,
            pady=5,
            state=tk.DISABLED,
        )
        self.open_report_button.pack(side=tk.LEFT, padx=5)

        self.exit_button = tk.Button(
            self.buttons_frame,
            text="Exit",
            command=self.root.destroy,
            bg="#f44336",
            fg="white",
            padx=10,
            pady=5,
        )
        self.exit_button.pack(side=tk.RIGHT, padx=5)

        # Initialize variables
        self.selected_dir = data_dir  # Default to the data directory
        self.latest_report = ""
        self.processing = False

        # Update status with default directory
        self.status_label.config(text=f"Default directory: {self.selected_dir}")

        # Handle window close
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        # Redirect stdout to the console
        self.redirect_stdout()

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

    def select_directory(self):
        selected = filedialog.askdirectory(
            title="Select Directory Containing PDF Files", initialdir=self.selected_dir
        )
        if selected:
            self.selected_dir = selected
            self.status_label.config(text=f"Selected directory: {self.selected_dir}")
            self.update_console(f"Selected directory: {self.selected_dir}\n")

    def update_console(self, message):
        self.console.config(state=tk.NORMAL)
        self.console.insert(tk.END, message)
        self.console.see(tk.END)
        self.console.config(state=tk.DISABLED)

    def process_pdfs(self):
        if self.processing:
            messagebox.showwarning(
                "Processing in Progress", "PDF processing is already running."
            )
            return

        if not self.selected_dir:
            messagebox.showwarning(
                "No Directory Selected",
                "Please select a directory containing PDF files first.",
            )
            return

        self.processing = True
        self.process_button.config(state=tk.DISABLED)
        self.select_dir_button.config(state=tk.DISABLED)
        self.status_label.config(text="Processing PDF files... Please wait.")

        # Clear console
        self.console.config(state=tk.NORMAL)
        self.console.delete(1.0, tk.END)
        self.console.config(state=tk.DISABLED)

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
                    self.update_console(
                        "\nNo PDF files found in the selected directory.\n"
                    )
                    self.update_console(
                        "Please add PDF files to the directory or select a different directory.\n"
                    )
                    return

                # Group data by folder
                from collections import defaultdict

                folder_data = defaultdict(list)

                # Process each PDF file
                for pdf_path in pdf_files:
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
                        f"Generated folder metrics report: {self.latest_report}\n"
                    )
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

                while True:
                    line = process.stdout.readline()
                    if not line and process.poll() is not None:
                        break
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

            self.update_console("\nProcessing complete!\n")

            # Enable the open report button if we have a report
            if self.latest_report:
                self.root.after(
                    0, lambda: self.open_report_button.config(state=tk.NORMAL)
                )

        except Exception as e:
            self.update_console(f"Error during processing: {str(e)}\n")
        finally:
            # Re-enable buttons
            self.root.after(0, lambda: self.process_button.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.select_dir_button.config(state=tk.NORMAL))
            self.root.after(
                0, lambda: self.status_label.config(text="Processing complete.")
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
