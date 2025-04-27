import os
import re
import glob
from collections import defaultdict
import PyPDF2
import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime


def find_pdf_files(root_dir="."):
    """Find all PDF files in all subdirectories without specific naming restrictions."""
    pdf_files = []

    for dirpath, _, filenames in os.walk(root_dir):
        for filename in filenames:
            if filename.lower().endswith(".pdf"):
                pdf_files.append(os.path.join(dirpath, filename))

    return pdf_files


def extract_text_from_pdf(pdf_path):
    """Extract text content from a PDF file."""
    try:
        with open(pdf_path, "rb") as file:
            reader = PyPDF2.PdfReader(file)
            text = ""
            for page in reader.pages:
                text += page.extract_text()
        return text
    except Exception as e:
        print(f"Error extracting text from {pdf_path}: {e}")
        return ""


def parse_pdf_content(text, pdf_path):
    """Parse PDF text content with improved data extraction and error indication."""
    data = {}

    # Extract filename
    filename = os.path.basename(pdf_path)
    data["filename"] = filename

    # Extract region from filename if available (optional)
    region_match = re.search(r"\(([^)]+)\)", filename)
    if region_match:
        data["region"] = region_match.group(1)
    else:
        data["region"] = "Unknown"  # Set default value if pattern not found

    # Extract timestamp from filename if available (optional)
    timestamp_match = re.search(r"(\d{4}_\d{2}_\d{2}_\d{2}_\d{2}_\d{2})", filename)
    if timestamp_match:
        data["timestamp"] = timestamp_match.group(1).replace("_", "-")
    else:
        # Use file modification date as fallback
        try:
            mod_time = os.path.getmtime(pdf_path)
            data["timestamp"] = datetime.fromtimestamp(mod_time).strftime(
                "%Y-%m-%d-%H-%M-%S"
            )
        except:
            data["timestamp"] = "Unknown"

    # Debug information about the file being processed
    print(f"\n==== Processing file: {filename} ====")

    # Playtime extraction - enhanced pattern matching
    playtime_found = False

    # List of pattern strategies to try
    playtime_patterns = [
        r"Play\s*Time\s*\n\s*(\d+\.?\d*)\s*s",  # Standard pattern: "Play Time" followed by newline and digits with 's'
        r"Play\s+Time\s*[\r\n\s]*(\d+\.?\d*)\s*s",  # More flexible with various whitespace
        r"Play Time.*?(\d+\.\d+)\s*s",  # Any content between "Play Time" and the number
        r"Play\s*Time[^0-9]*(\d+\.?\d+)",  # Anything except digits between "Play Time" and number
        r"Play\s*Time.*?(\d+\.?\d+)",  # Most general pattern
        r"play\s*time.*?(\d+\.?\d+)\s*s",  # Case insensitive variation
        r"duration.*?(\d+\.?\d+)\s*s",  # Alternative term "duration"
        r"length.*?(\d+\.?\d+)\s*s",  # Alternative term "length"
        r"time.*?(\d+\.?\d+)\s*s",  # Last resort - any "time" followed by a number and "s"
    ]

    # Try each pattern strategy
    for i, pattern in enumerate(playtime_patterns):
        playtime_match = re.search(pattern, text, re.DOTALL | re.IGNORECASE)
        if playtime_match:
            try:
                # Extract the value from the matched group
                data["playtime"] = float(playtime_match.group(1))
                print(f"Found Playtime with pattern {i+1}: {data['playtime']} s")
                playtime_found = True
                break
            except (ValueError, IndexError) as e:
                print(f"Error parsing playtime with pattern {i+1}: {e}")

    # If still not found, try looking directly at play time context
    if not playtime_found:
        print("Playtime patterns failed, examining Play Time section context...")

        # Get 200 characters surrounding "Play Time" mention
        play_time_idx = text.find("Play Time")
        if play_time_idx == -1:
            play_time_idx = text.lower().find("play time")

        if play_time_idx != -1:
            context_start = max(0, play_time_idx - 20)
            context_end = min(len(text), play_time_idx + 180)
            context = text[context_start:context_end]
            print(f"Play Time context: '{context}'")

            # Look for any number in this context
            number_match = re.search(r"(\d+\.?\d*)\s*s", context)
            if number_match:
                try:
                    data["playtime"] = float(number_match.group(1))
                    print(f"Extracted Playtime from context: {data['playtime']} s")
                    playtime_found = True
                except (ValueError, IndexError) as e:
                    print(f"Error parsing playtime from context: {e}")

    # As a very last resort, search for any number followed by "s" or "seconds" in the first 10% of the document
    if not playtime_found:
        print("Trying to find any time-like values in the document...")
        search_section = text[: int(len(text) * 0.1)]  # First 10% of document
        time_matches = re.findall(
            r"(\d+\.?\d*)\s*(?:s|sec|seconds)", search_section, re.IGNORECASE
        )
        if time_matches:
            try:
                # Use the first reasonable value found (between 1 and 3600 seconds)
                for potential_time in time_matches:
                    time_val = float(potential_time)
                    if 1 <= time_val <= 3600:
                        data["playtime"] = time_val
                        print(f"Found potential playtime value: {data['playtime']} s")
                        playtime_found = True
                        break
            except (ValueError, IndexError) as e:
                print(f"Error parsing potential time values: {e}")

    # If still not found, set to -1 to indicate an error
    if not playtime_found:
        data["playtime"] = -1  # Error value
        print(f"Could not find Playtime, setting error value (-1)")

    # Extract FPS information
    # FPS pattern looks for "Avg : XX.XX" under the FPS section
    fps_patterns = [
        r"FPS\s*.*?\s*Avg\s*:\s*(\d+\.?\d*)",  # Standard pattern
        r"FPS.*?average.*?(\d+\.?\d*)",  # Alternative "average" wording
        r"Frames\s*Per\s*Second.*?(\d+\.?\d*)",  # Full "frames per second" term
        r"Frame\s*Rate.*?(\d+\.?\d*)",  # Alternative "frame rate" term
        r"Average\s*FPS.*?(\d+\.?\d*)",  # "Average FPS" pattern
        r"fps.*?(\d+\.?\d*)",  # Simple "fps" mention
    ]

    fps_found = False

    for i, pattern in enumerate(fps_patterns):
        fps_match = re.search(pattern, text, re.DOTALL | re.IGNORECASE)
        if fps_match:
            try:
                data["fps"] = float(fps_match.group(1))
                print(f"Found FPS with pattern {i+1}: {data['fps']}")
                fps_found = True
                break
            except (ValueError, IndexError) as e:
                print(f"Error parsing FPS with pattern {i+1}: {e}")

    # If not found, use error value (-1)
    if not fps_found:
        data["fps"] = -1  # Error value
        print(f"Could not find FPS, setting error value (-1)")

    # Bandwidth patterns
    bandwidth_patterns = [
        r"Bandwidth\s*.*?\s*Avg\s*:\s*(\d+\.?\d*)\s*Mbps",  # Standard pattern
        r"Bandwidth.*?(\d+\.?\d*)\s*Mbps",  # Simpler pattern
        r"Average\s*Bandwidth.*?(\d+\.?\d*)\s*Mbps",  # "Average Bandwidth" term
        r"Network\s*Speed.*?(\d+\.?\d*)\s*Mbps",  # Alternative "Network Speed" term
        r"Data\s*Rate.*?(\d+\.?\d*)\s*Mbps",  # Alternative "Data Rate" term
        r"(\d+\.?\d*)\s*Mbps",  # Last resort - any Mbps value
    ]

    bw_found = False

    for i, pattern in enumerate(bandwidth_patterns):
        bandwidth_match = re.search(pattern, text, re.DOTALL | re.IGNORECASE)
        if bandwidth_match:
            try:
                data["bandwidth"] = float(bandwidth_match.group(1))
                print(f"Found Bandwidth with pattern {i+1}: {data['bandwidth']} Mbps")
                bw_found = True
                break
            except (ValueError, IndexError) as e:
                print(f"Error parsing Bandwidth with pattern {i+1}: {e}")

    # If not found, use error value (-1)
    if not bw_found:
        data["bandwidth"] = -1  # Error value
        print(f"Could not find Bandwidth, setting error value (-1)")

    # RTT patterns
    rtt_patterns = [
        r"Round Trip Time\s*.*?\s*Avg\s*:\s*(\d+\.?\d*)\s*ms",  # Standard pattern
        r"RTT.*?(\d+\.?\d*)\s*ms",  # RTT abbreviation
        r"Round\s*Trip.*?(\d+\.?\d*)\s*ms",  # Partial "Round Trip" term
        r"Latency.*?(\d+\.?\d*)\s*ms",  # Alternative "Latency" term
        r"Ping.*?(\d+\.?\d*)\s*ms",  # Alternative "Ping" term
        r"Response\s*Time.*?(\d+\.?\d*)\s*ms",  # Alternative "Response Time" term
    ]

    rtt_found = False

    for i, pattern in enumerate(rtt_patterns):
        rtt_match = re.search(pattern, text, re.DOTALL | re.IGNORECASE)
        if rtt_match:
            try:
                data["rtt"] = float(rtt_match.group(1))
                print(f"Found RTT with pattern {i+1}: {data['rtt']} ms")
                rtt_found = True
                break
            except (ValueError, IndexError) as e:
                print(f"Error parsing RTT with pattern {i+1}: {e}")

    # If not found, use error value (-1)
    if not rtt_found:
        data["rtt"] = -1  # Error value
        print(f"Could not find RTT, setting error value (-1)")

    print(f"==== Finished processing file: {filename} ====\n")
    return data


def get_folder_path(pdf_path):
    """Extract folder path from the full path."""
    return os.path.dirname(pdf_path)


def generate_folder_report(folder_data):
    """Generate a report of average metrics by folder with improved table formatting and averages in appendix."""
    # Create output directory
    os.makedirs("reports", exist_ok=True)

    # Prepare report content
    report_content = "# Folder Metrics Report\n\n"
    report_content += (
        f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n"
    )
    report_content += "## Summary by Folder\n\n"

    # Sort folders for consistent output
    sorted_folders = sorted(folder_data.keys())

    # Calculate column widths for summary table
    folder_col_width = max(
        max(len(folder) for folder in sorted_folders), len("Folder Path")
    )
    files_col_width = max(5, len("Number of Files"))
    playtime_col_width = max(
        15, len("Avg Playtime (s)")
    )  # Increased for error messages
    fps_col_width = max(15, len("Average FPS"))  # Increased for error messages
    bw_col_width = max(
        15, len("Average Bandwidth (Mbps)")
    )  # Increased for error messages
    rtt_col_width = max(15, len("Average RTT (ms)"))  # Increased for error messages

    # Create header row with dynamic width
    report_content += f"| {'Folder Path'.ljust(folder_col_width)} | {'Number of Files'.ljust(files_col_width)} | {'Avg Playtime (s)'.ljust(playtime_col_width)} | {'Average FPS'.ljust(fps_col_width)} | {'Average Bandwidth (Mbps)'.ljust(bw_col_width)} | {'Average RTT (ms)'.ljust(rtt_col_width)} |\n"
    report_content += f"| {'-' * folder_col_width} | {'-' * files_col_width} | {'-' * playtime_col_width} | {'-' * fps_col_width} | {'-' * bw_col_width} | {'-' * rtt_col_width} |\n"

    # Add data rows with more robust average calculations
    for folder in sorted_folders:
        files = folder_data[folder]
        num_files = len(files)

        # For all metrics, exclude -1 (error) values and include only valid values
        valid_playtimes = [
            file.get("playtime")
            for file in files
            if isinstance(file.get("playtime"), (int, float))
            and file.get("playtime") > 0
        ]

        valid_fps = [
            file.get("fps")
            for file in files
            if isinstance(file.get("fps"), (int, float)) and file.get("fps") > 0
        ]

        valid_bandwidths = [
            file.get("bandwidth")
            for file in files
            if isinstance(file.get("bandwidth"), (int, float))
            and file.get("bandwidth") > 0
        ]

        valid_rtts = [
            file.get("rtt")
            for file in files
            if isinstance(file.get("rtt"), (int, float)) and file.get("rtt") > 0
        ]

        # Count error files for each metric
        error_playtime_count = sum(
            1
            for file in files
            if isinstance(file.get("playtime"), (int, float))
            and file.get("playtime") == -1
        )

        error_fps_count = sum(
            1
            for file in files
            if isinstance(file.get("fps"), (int, float)) and file.get("fps") == -1
        )

        error_bw_count = sum(
            1
            for file in files
            if isinstance(file.get("bandwidth"), (int, float))
            and file.get("bandwidth") == -1
        )

        error_rtt_count = sum(
            1
            for file in files
            if isinstance(file.get("rtt"), (int, float)) and file.get("rtt") == -1
        )

        # Format playtime average and error information
        if valid_playtimes:
            avg_playtime = sum(valid_playtimes) / len(valid_playtimes)
            if error_playtime_count > 0:
                avg_playtime_str = (
                    f"{avg_playtime:.2f} (errors: {error_playtime_count} files)"
                )
            else:
                avg_playtime_str = f"{avg_playtime:.2f}"
        else:
            if error_playtime_count > 0:
                avg_playtime_str = f"Error (all {error_playtime_count} files)"
            else:
                avg_playtime_str = "N/A"

        # Format FPS average and error information
        if valid_fps:
            avg_fps = sum(valid_fps) / len(valid_fps)
            if error_fps_count > 0:
                avg_fps_str = f"{avg_fps:.2f} (errors: {error_fps_count} files)"
            else:
                avg_fps_str = f"{avg_fps:.2f}"
        else:
            if error_fps_count > 0:
                avg_fps_str = f"Error (all {error_fps_count} files)"
            else:
                avg_fps_str = "N/A"

        # Format bandwidth average and error information
        if valid_bandwidths:
            avg_bandwidth = sum(valid_bandwidths) / len(valid_bandwidths)
            if error_bw_count > 0:
                avg_bandwidth_str = (
                    f"{avg_bandwidth:.2f} (errors: {error_bw_count} files)"
                )
            else:
                avg_bandwidth_str = f"{avg_bandwidth:.2f}"
        else:
            if error_bw_count > 0:
                avg_bandwidth_str = f"Error (all {error_bw_count} files)"
            else:
                avg_bandwidth_str = "N/A"

        # Format RTT average and error information
        if valid_rtts:
            avg_rtt = sum(valid_rtts) / len(valid_rtts)
            if error_rtt_count > 0:
                avg_rtt_str = f"{avg_rtt:.2f} (errors: {error_rtt_count} files)"
            else:
                avg_rtt_str = f"{avg_rtt:.2f}"
        else:
            if error_rtt_count > 0:
                avg_rtt_str = f"Error (all {error_rtt_count} files)"
            else:
                avg_rtt_str = "N/A"

        # Format the row with proper alignment
        report_content += (
            f"| {folder.ljust(folder_col_width)} | {str(num_files).ljust(files_col_width)} | "
            f"{avg_playtime_str.ljust(playtime_col_width)} | {avg_fps_str.ljust(fps_col_width)} | "
            f"{avg_bandwidth_str.ljust(bw_col_width)} | {avg_rtt_str.ljust(rtt_col_width)} |\n"
        )

    # Add appendix with detailed metrics for each folder
    report_content += "\n\n## Appendix: Detailed Metrics by Folder\n\n"

    for folder in sorted_folders:
        files = folder_data[folder]
        report_content += f"### {folder}\n\n"

        # Sort files by name for consistent output
        sorted_files = sorted(files, key=lambda x: x.get("filename", ""))

        # Calculate column widths for each folder's table
        filename_width = max(
            max(len(file.get("filename", "Unknown")) for file in sorted_files),
            len("Filename"),
        )
        playtime_width = max(15, len("Playtime (s)"))  # Increased for error messages
        fps_width = max(15, len("FPS"))  # Increased for error messages
        bw_width = max(15, len("Bandwidth (Mbps)"))  # Increased for error messages
        rtt_width = max(15, len("RTT (ms)"))  # Increased for error messages

        # Create header with dynamic width
        report_content += f"| {'Filename'.ljust(filename_width)} | {'Playtime (s)'.ljust(playtime_width)} | {'FPS'.ljust(fps_width)} | {'Bandwidth (Mbps)'.ljust(bw_width)} | {'RTT (ms)'.ljust(rtt_width)} |\n"
        report_content += f"| {'-' * filename_width} | {'-' * playtime_width} | {'-' * fps_width} | {'-' * bw_width} | {'-' * rtt_width} |\n"

        # For all metrics, exclude -1 (error) values and include only valid values (recalculated for this folder)
        valid_playtimes = [
            file.get("playtime")
            for file in sorted_files
            if isinstance(file.get("playtime"), (int, float))
            and file.get("playtime") > 0
        ]

        valid_fps = [
            file.get("fps")
            for file in sorted_files
            if isinstance(file.get("fps"), (int, float)) and file.get("fps") > 0
        ]

        valid_bandwidths = [
            file.get("bandwidth")
            for file in sorted_files
            if isinstance(file.get("bandwidth"), (int, float))
            and file.get("bandwidth") > 0
        ]

        valid_rtts = [
            file.get("rtt")
            for file in sorted_files
            if isinstance(file.get("rtt"), (int, float)) and file.get("rtt") > 0
        ]

        # Add data rows
        for file_data in sorted_files:
            filename = file_data.get("filename", "Unknown")
            playtime = file_data.get("playtime", "N/A")
            fps = file_data.get("fps", "N/A")
            bandwidth = file_data.get("bandwidth", "N/A")
            rtt = file_data.get("rtt", "N/A")

            # Display -1 values as "Error (extraction failed)" for all metrics
            playtime_str = (
                "Error (extraction failed)"
                if isinstance(playtime, (int, float)) and playtime == -1
                else (
                    f"{playtime:.2f}"
                    if isinstance(playtime, (int, float))
                    else str(playtime)
                )
            )
            fps_str = (
                "Error (extraction failed)"
                if isinstance(fps, (int, float)) and fps == -1
                else (f"{fps:.2f}" if isinstance(fps, (int, float)) else str(fps))
            )
            bandwidth_str = (
                "Error (extraction failed)"
                if isinstance(bandwidth, (int, float)) and bandwidth == -1
                else (
                    f"{bandwidth:.2f}"
                    if isinstance(bandwidth, (int, float))
                    else str(bandwidth)
                )
            )
            rtt_str = (
                "Error (extraction failed)"
                if isinstance(rtt, (int, float)) and rtt == -1
                else (f"{rtt:.2f}" if isinstance(rtt, (int, float)) else str(rtt))
            )

            # Format the row with proper alignment
            report_content += (
                f"| {filename.ljust(filename_width)} | {playtime_str.ljust(playtime_width)} | "
                f"{fps_str.ljust(fps_width)} | {bandwidth_str.ljust(bw_width)} | {rtt_str.ljust(rtt_width)} |\n"
            )

        # Add average row at the bottom of each folder table
        # Calculate averages for valid data
        avg_playtime_str = (
            f"{sum(valid_playtimes) / len(valid_playtimes):.2f}"
            if valid_playtimes
            else "N/A"
        )
        avg_fps_str = f"{sum(valid_fps) / len(valid_fps):.2f}" if valid_fps else "N/A"
        avg_bw_str = (
            f"{sum(valid_bandwidths) / len(valid_bandwidths):.2f}"
            if valid_bandwidths
            else "N/A"
        )
        avg_rtt_str = (
            f"{sum(valid_rtts) / len(valid_rtts):.2f}" if valid_rtts else "N/A"
        )

        # Add a separator row
        report_content += f"| {'-' * filename_width} | {'-' * playtime_width} | {'-' * fps_width} | {'-' * bw_width} | {'-' * rtt_width} |\n"

        # Add the average row
        report_content += (
            f"| {'AVERAGE'.ljust(filename_width)} | {avg_playtime_str.ljust(playtime_width)} | "
            f"{avg_fps_str.ljust(fps_width)} | {avg_bw_str.ljust(bw_width)} | {avg_rtt_str.ljust(rtt_width)} |\n"
        )

        report_content += "\n"

    # Write report to markdown file only
    md_filename = "reports/folder_metrics_report.md"

    with open(md_filename, "w", encoding="utf-8") as f:
        f.write(report_content)

    print(f"Generated folder metrics report: {md_filename}")
    return md_filename


def main():
    """Main function to process PDF files and generate folder-based report."""
    print("Starting PDF report processing...")

    # Find all PDF files in current directory and all subdirectories
    pdf_files = find_pdf_files(".")
    print(f"Found {len(pdf_files)} PDF files to process.")

    # Group data by folder
    folder_data = defaultdict(list)

    # Process each PDF file
    for pdf_path in pdf_files:
        try:
            print(f"Processing: {pdf_path}")

            # Extract text from PDF
            text = extract_text_from_pdf(pdf_path)
            text = re.sub(r"\s*\.\s*", ".", text)
            print(text)
            if text:
                # Parse content to get metrics
                data = parse_pdf_content(text, pdf_path)

                # Get the folder path
                folder_path = get_folder_path(pdf_path)

                # Add data to the folder's collection
                folder_data[folder_path].append(data)
            else:
                print(f"Warning: No text extracted from {pdf_path}")

        except Exception as e:
            print(f"Error processing {pdf_path}: {e}")

    # Generate report of folder metrics
    if folder_data:
        report_path = generate_folder_report(folder_data)
        print(f"Generated folder metrics report: {report_path}")
    else:
        print("No data was processed successfully. Report not generated.")

    print("\nProcessing complete! Press Enter to exit...")
    input()  # Keep console window open until user presses Enter


if __name__ == "__main__":
    main()
