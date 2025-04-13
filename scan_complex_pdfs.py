import os
import sys
import time
import argparse
import logging
from concurrent.futures import ProcessPoolExecutor, as_completed

import camelot
import pandas as pd

# ------------------------
# Global Configurations
# ------------------------
LOG_LEVEL = logging.INFO
logging.basicConfig(
    level=LOG_LEVEL,
    format='%(asctime)s - %(levelname)s - [%(processName)s] - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
log = logging.getLogger(__name__)

# Default Camelot parameters
DEFAULT_CAMELOT_FLAVOR = 'lattice'
DEFAULT_CAMELOT_LINE_SCALE = 40


# ------------------------
# Helper Functions
# ------------------------
def is_complex_table(table, pdf_basename: str, table_index: int) -> bool:
    """
    Determines if a table is complex by checking its cell attributes.
    
    A table is considered complex if any cell in the table has either a horizontal span or vertical span.
    
    Args:
        table: A Camelot table object.
        pdf_basename (str): Base name of the PDF file (for logging).
        table_index (int): Index of the table in the PDF.

    Returns:
        bool: True if the table is considered complex, False otherwise.
    """
    try:
        if not hasattr(table, "cells") or not table.cells or not isinstance(table.cells, list):
            log.debug(f"Table {table_index} in '{pdf_basename}' lacks valid cells.")
            return False

        first_row = table.cells[0]
        if not (isinstance(first_row, list) and first_row):
            log.debug(f"Table {table_index} in '{pdf_basename}' has irregular or empty cells.")
            return False

        for r_idx, row in enumerate(table.cells):
            if not isinstance(row, list):
                continue  # Skip non-list rows
            for c_idx, cell in enumerate(row):
                if hasattr(cell, "hspan") and hasattr(cell, "vspan"):
                    if cell.hspan or cell.vspan:
                        log.info(f"Complex cell found in '{pdf_basename}', "
                                 f"Table {table_index}, Cell ({r_idx},{c_idx}).")
                        return True
        return False

    except Exception as error:
        log.warning(f"Error checking table complexity for Table {table_index} in '{pdf_basename}': {error}")
        return False


def analyze_pdf_for_complex_tables(pdf_path: str, flavor: str, line_scale: int) -> tuple:
    """
    Analyzes all tables in a PDF, logs the analysis (including printing table content),
    and returns detailed information.
    
    Returns:
        tuple: (complex_table_indices: list[int], total_table_count: int, table_details: list[dict])
            - complex_table_indices: 0-based indices of tables detected as complex.
            - total_table_count: Total number of tables detected in the PDF.
              (Returns -1 if there is a processing error.)
            - table_details: A list of dictionaries. Each dictionary represents a table with keys:
                "index": table index,
                "complex": boolean flag,
                "content": string representation of the table.
    """
    pdf_basename = os.path.basename(pdf_path)
    log.info(f"Worker starting analysis for: {pdf_basename}")

    try:
        tables = camelot.read_pdf(
            pdf_path,
            pages='all',
            flavor=flavor,
            suppress_stdout=True,
            line_scale=line_scale
        )
        total_tables = len(tables)
        log.info(f"Camelot found {total_tables} table(s) in: {pdf_basename}")

        if total_tables == 0:
            log.info(f"Finished analysis (no tables found): {pdf_basename}")
            return ([], 0, [])

        complex_indices = []
        table_details = []
        for table_index, table in enumerate(tables):
            log.info(f"Analyzing Table {table_index} in {pdf_basename}...")
            flag_complex = is_complex_table(table, pdf_basename, table_index)

            try:
                table_content = table.df.to_string()
            except Exception as e:
                table_content = f"Unable to convert table to string: {e}"

            # Log the table details into the output log
            if flag_complex:
                log.info(f"*** COMPLEX TABLE DETECTED *** File='{pdf_basename}', Table Index={table_index}")
                complex_indices.append(table_index)
                log.info(f"[COMPLEX] Table {table_index} in {pdf_basename}:\n{table_content}")
            else:
                log.info(f"[SIMPLE] Table {table_index} in {pdf_basename}:\n{table_content}")

            # Append details for output file later
            table_details.append({
                "index": table_index,
                "complex": flag_complex,
                "content": table_content
            })

        log.info(f"Finished analysis for: {pdf_basename} (Complex tables found: {complex_indices})")
        return (complex_indices, total_tables, table_details)

    except FileNotFoundError:
        log.error(f"File not found during processing: '{pdf_path}'")
        return ([], -1, [])
    except ImportError as imp_err:
        log.critical(f"Import error processing '{pdf_basename}': {imp_err}. Check dependencies?", exc_info=True)
        return ([], -1, [])
    except Exception as error:
        error_str = str(error).lower()
        if "skip-no-text" in error_str:
            log.warning(f"Skipping '{pdf_basename}' (likely no text detected by Camelot).")
        elif "password" in error_str:
            log.warning(f"Skipping '{pdf_basename}' (likely password protected).")
        elif "ghostscript" in error_str or "error" in error_str:
            log.error(f"Parsing/Ghostscript error processing '{pdf_basename}': {error}",
                      exc_info=(LOG_LEVEL <= logging.DEBUG))
        else:
            log.error(f"Failed processing '{pdf_basename}': {error}",
                      exc_info=(LOG_LEVEL <= logging.DEBUG))
        return ([], -1, [])


def scan_pdf_directory(pdf_directory: str) -> list:
    """
    Recursively scans the given directory for PDF files.
    
    Args:
        pdf_directory (str): The directory to scan.
    
    Returns:
        list: Absolute paths of PDF files.
    """
    pdf_files = []
    try:
        for root, _, files in os.walk(pdf_directory):
            for file in files:
                if file.lower().endswith(".pdf"):
                    pdf_files.append(os.path.join(root, file))
    except Exception as error:
        log.critical(f"Error scanning directory '{pdf_directory}': {error}", exc_info=True)
        sys.exit(1)
    return pdf_files


def write_results_to_file(results: dict, output_file: str) -> None:
    """
    Writes the detailed results from PDF analysis to an output file.
    
    The output for each PDF will include:
      - The PDF's full path.
      - Total tables detected.
      - The indices of complex tables.
      - A detailed list of each table with a flag and its content.
    
    Args:
        results (dict): Dictionary mapping each PDF file path to a dictionary with keys:
            "complex_indices", "total_tables", and "table_details".
        output_file (str): Output file path.
    """
    if not results:
        log.info("No PDFs containing complex table information were found. Output file not created.")
        return

    try:
        output_path = os.path.abspath(output_file)
        with open(output_path, 'w', encoding='utf-8') as f:
            for pdf_path, details in sorted(results.items()):
                f.write(f"PDF File: {pdf_path}\n")
                f.write(f"Total Tables Detected: {details['total_tables']}\n")
                f.write(f"Complex Table Indices: {details['complex_indices']}\n")
                f.write("Table Details:\n")
                for table in details["table_details"]:
                    flag = "COMPLEX" if table["complex"] else "SIMPLE"
                    f.write(f"  Table {table['index']} [{flag}]:\n")
                    # Indent each line of the table content for readability.
                    for line in table["content"].splitlines():
                        f.write("      " + line + "\n")
                    f.write("\n")
                f.write("-" * 80 + "\n\n")
        log.info(f"Detailed analysis saved to: '{output_path}'")
    except IOError as io_err:
        log.error(f"IOError writing output file '{output_file}': {io_err}")
    except Exception as error:
        log.error(f"Unexpected error writing output file '{output_file}': {error}", exc_info=True)


# ------------------------
# Main Execution
# ------------------------
def main():
    parser = argparse.ArgumentParser(
        description="Scan PDF files, log table analysis details (with table content) and output detailed information.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter
    )
    parser.add_argument("pdf_directory", help="Directory containing PDF files to scan.")
    parser.add_argument("-o", "--output", default="results.txt",
                        help="Output file for detailed PDF table analysis.")
    parser.add_argument("-w", "--workers", type=int, default=os.cpu_count(),
                        help="Number of worker processes.")
    parser.add_argument("-n", "--limit", type=int, default=None,
                        help="Limit the number of PDFs to process.")
    parser.add_argument("--flavor", default=DEFAULT_CAMELOT_FLAVOR, choices=['lattice', 'stream'],
                        help="Camelot flavor.")
    parser.add_argument("--line_scale", type=int, default=DEFAULT_CAMELOT_LINE_SCALE,
                        help="Camelot line_scale.")
    parser.add_argument("--log", default="INFO", choices=['DEBUG', 'INFO', 'WARNING', 'ERROR', 'CRITICAL'],
                        help="Logging level.")

    args = parser.parse_args()

    # Set logging level based on argument
    effective_log_level = getattr(logging, args.log.upper(), logging.INFO)
    logging.getLogger().setLevel(effective_log_level)
    global LOG_LEVEL
    LOG_LEVEL = effective_log_level

    log.info("=" * 20 + " Starting PDF Complex Table Scan " + "=" * 20)
    log.info(f"Source Directory: {os.path.abspath(args.pdf_directory)}")
    log.info(f"Output File: {os.path.abspath(args.output)}")
    log.info(f"Worker Processes: {args.workers}")
    if args.limit:
        log.info(f"Processing Limit: First {args.limit} PDF file(s)")
    log.info(f"Camelot Settings: flavor='{args.flavor}', line_scale={args.line_scale}")
    log.info(f"Logging Level: {args.log.upper()}")
    log.info("-" * 66)

    # Validate the source directory
    if not os.path.isdir(args.pdf_directory):
        log.critical(f"CRITICAL ERROR: Source directory not found: {args.pdf_directory}")
        sys.exit(1)

    # Scan for PDF files
    log.info("Scanning for PDF files...")
    pdf_files = scan_pdf_directory(args.pdf_directory)
    total_pdf_files = len(pdf_files)
    log.info(f"Found {total_pdf_files} PDF file(s).")

    # Apply processing limit, if any
    if args.limit:
        if args.limit <= 0:
            log.warning("Processing limit is non-positive. Exiting.")
            sys.exit(0)
        pdf_files = pdf_files[:args.limit]
        log.info(f"Processing limited to first {len(pdf_files)} file(s).")

    if not pdf_files:
        log.info("No PDF files to process.")
        sys.exit(0)

    # Process PDFs using a process pool executor
    detailed_pdf_results = {}  # Will map each PDF path to a dict containing detailed analysis.
    processed_count = 0
    total_tables_detected = 0
    processing_errors = 0
    start_time = time.time()

    log.info(f"Submitting {len(pdf_files)} PDF(s) to process pool...")
    with ProcessPoolExecutor(max_workers=args.workers) as executor:
        future_to_pdf = {
            executor.submit(
                analyze_pdf_for_complex_tables,
                pdf_path,
                args.flavor,
                args.line_scale
            ): pdf_path for pdf_path in pdf_files
        }

        log.info("Waiting for results from workers...")
        for future in as_completed(future_to_pdf):
            processed_count += 1
            pdf_path = future_to_pdf[future]
            pdf_basename = os.path.basename(pdf_path)
            try:
                complex_indices, table_count, table_details = future.result()

                if table_count >= 0:  # Successfully processed
                    total_tables_detected += table_count
                    detailed_pdf_results[pdf_path] = {
                        "complex_indices": sorted(complex_indices),
                        "total_tables": table_count,
                        "table_details": table_details
                    }
                    if not complex_indices:
                        log.info(f"'{pdf_basename}' does not contain any complex tables.")
                else:
                    log.warning(f"Worker indicated processing error for '{pdf_basename}'.")
                    processing_errors += 1
            except Exception as error:
                log.error(f"Error retrieving result for '{pdf_basename}': {error}",
                          exc_info=(LOG_LEVEL <= logging.DEBUG))
                processing_errors += 1

            # Progress reporting
            log_interval = max(100, len(pdf_files) // 100) if len(pdf_files) > 100 else 50
            if processed_count % log_interval == 0 or processed_count == len(pdf_files):
                elapsed_time = time.time() - start_time
                pdfs_per_sec = processed_count / elapsed_time if elapsed_time > 0 else 0
                log.info(f"Progress: {processed_count}/{len(pdf_files)} files analyzed "
                         f"({pdfs_per_sec:.1f} files/sec). "
                         f"Complex PDFs: {sum(1 for v in detailed_pdf_results.values() if v['complex_indices'])}. "
                         f"Total tables: {total_tables_detected}. "
                         f"Errors: {processing_errors}.")

    # Processing summary
    end_time = time.time()
    total_time = end_time - start_time
    successful_pdfs = processed_count - processing_errors
    avg_time_per_pdf = total_time / processed_count if processed_count > 0 else 0

    log.info("=" * 25 + " Processing Summary " + "=" * 25)
    log.info(f"Total PDFs submitted: {len(pdf_files)}")
    log.info(f"Total PDFs analyzed: {processed_count}")
    log.info(f"PDFs with processing errors: {processing_errors}")
    log.info(f"Successfully analyzed PDFs: {successful_pdfs}")
    log.info(f"Total tables detected: {total_tables_detected}")
    log.info(f"Total processing time: {total_time:.2f} seconds")
    log.info(f"Average time per PDF: {avg_time_per_pdf:.3f} seconds")
    log.info("=" * 66)

    # Write detailed results to output file
    write_results_to_file(detailed_pdf_results, args.output)
    log.info("Script finished.")


if __name__ == "__main__":
    main()