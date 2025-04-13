import os
import sys
import time
import argparse
import logging
from concurrent.futures import ProcessPoolExecutor, as_completed
from docx import Document

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

# ------------------------
# Helper Functions for DOCX
# ------------------------
def is_complex_table_docx(table, docx_basename: str, table_index: int) -> bool:
    """
    Determines if a table in a DOCX file is considered complex.
    A table is marked as complex if any cell shows evidence of either a horizontal merge (<w:gridSpan>)
    or a vertical merge (<w:vMerge>).

    Args:
        table: A docx.table table object.
        docx_basename (str): The base name of the DOCX file for logging.
        table_index (int): Index of the table in the document.

    Returns:
        bool: True if the table is considered complex, False otherwise.
    """
    try:
        for row_index, row in enumerate(table.rows):
            for col_index, cell in enumerate(row.cells):
                tc = cell._tc  # underlying XML element for the cell
                # Check for horizontal merge via <w:gridSpan>
                grid_span = tc.xpath('.//w:gridSpan')
                if grid_span:
                    log.info(f"Complex cell (horizontal merge) detected in '{docx_basename}', "
                             f"Table {table_index}, Cell ({row_index},{col_index}).")
                    return True
                # Check for vertical merge via <w:vMerge>
                vmerge = tc.xpath('.//w:vMerge')
                if vmerge:
                    log.info(f"Complex cell (vertical merge) detected in '{docx_basename}', "
                             f"Table {table_index}, Cell ({row_index},{col_index}).")
                    return True
        return False
    except Exception as error:
        log.warning(f"Error checking complexity for Table {table_index} in '{docx_basename}': {error}")
        return False


def analyze_docx_for_complex_tables(docx_path: str) -> tuple:
    """
    Analyzes all tables in a DOCX file, logs table details and returns analysis information.
    
    Since DOCX files are flow-based (not paginated), the page number is always returned as 'N/A'.
    
    Returns:
        tuple: (complex_table_indices: list[int], total_table_count: int, table_details: list[dict])
            - complex_table_indices: 0-based indices of tables detected as complex.
            - total_table_count: Total number of tables in the document.
              (Returns -1 in case of a processing error.)
            - table_details: A list of dictionaries for each table with keys:
                "index": table index,
                "page": page number (always 'N/A'),
                "complex": boolean flag indicating complexity,
                "content": string representation of the table.
    """
    docx_basename = os.path.basename(docx_path)
    log.info(f"Starting analysis for DOCX file: {docx_basename}")

    try:
        document = Document(docx_path)
    except Exception as error:
        log.error(f"Error opening DOCX file '{docx_path}': {error}")
        return ([], -1, [])

    tables = document.tables
    total_tables = len(tables)
    log.info(f"Found {total_tables} table(s) in: {docx_basename}")

    if total_tables == 0:
        log.info(f"No tables found in {docx_basename}.")
        return ([], 0, [])

    complex_indices = []
    table_details = []
    for table_index, table in enumerate(tables):
        # For DOCX files, the page number is not available
        page_number = 'N/A'
        log.info(f"Analyzing Table {table_index} in {docx_basename}...")
        flag_complex = is_complex_table_docx(table, docx_basename, table_index)

        # Create a text representation of the table by joining cell texts in each row
        table_rows = []
        for row in table.rows:
            row_cells = [cell.text.strip() for cell in row.cells]
            table_rows.append(" | ".join(row_cells))
        table_content = "\n".join(table_rows)

        if flag_complex:
            log.info(f"*** COMPLEX TABLE DETECTED *** File='{docx_basename}', Table Index={table_index}, Page={page_number}")
            complex_indices.append(table_index)
            log.info(f"[COMPLEX] Table {table_index} on page {page_number} in {docx_basename}:\n{table_content}")
        else:
            log.info(f"[SIMPLE] Table {table_index} on page {page_number} in {docx_basename}:\n{table_content}")

        table_details.append({
            "index": table_index,
            "page": page_number,
            "complex": flag_complex,
            "content": table_content
        })

    log.info(f"Finished analysis for: {docx_basename} (Complex tables found: {complex_indices})")
    return (complex_indices, total_tables, table_details)


def scan_docx_directory(docx_directory: str) -> list:
    """
    Recursively scans the given directory for DOCX files.
    
    Args:
        docx_directory (str): The directory to scan.
    
    Returns:
        list: Absolute paths of DOCX files.
    """
    docx_files = []
    try:
        for root, _, files in os.walk(docx_directory):
            for file in files:
                if file.lower().endswith(".docx"):
                    docx_files.append(os.path.join(root, file))
    except Exception as error:
        log.critical(f"Error scanning directory '{docx_directory}': {error}", exc_info=True)
        sys.exit(1)
    return docx_files


def write_results_to_file(results: dict, output_file: str) -> None:
    """
    Writes the detailed results from DOCX analysis to an output file.
    
    The output for each DOCX file will include:
      - The DOCX file's full path.
      - Total tables detected.
      - The indices of complex tables.
      - A detailed list of each table with a flag, its page (always 'N/A'), and its content.
    
    Args:
        results (dict): Dictionary mapping each DOCX file path to a dictionary with keys:
            "complex_indices", "total_tables", and "table_details".
        output_file (str): Output file path.
    """
    if not results:
        log.info("No DOCX files containing complex table information were found. Output file not created.")
        return

    try:
        output_path = os.path.abspath(output_file)
        with open(output_path, 'w', encoding='utf-8') as f:
            for docx_path, details in sorted(results.items()):
                f.write(f"DOCX File: {docx_path}\n")
                f.write(f"Total Tables Detected: {details['total_tables']}\n")
                f.write(f"Complex Table Indices: {details['complex_indices']}\n")
                f.write("Table Details:\n")
                for table in details["table_details"]:
                    flag = "COMPLEX" if table["complex"] else "SIMPLE"
                    f.write(f"  Table {table['index']} [{flag}] (Page: {table['page']}):\n")
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
        description="Scan DOCX files, log table analysis details (with table content) and output detailed information.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter
    )
    parser.add_argument("docx_directory", help="Directory containing DOCX files to scan.")
    parser.add_argument("-o", "--output", default="docx_results.txt",
                        help="Output file for detailed DOCX table analysis.")
    parser.add_argument("-w", "--workers", type=int, default=os.cpu_count(),
                        help="Number of worker processes.")
    parser.add_argument("-n", "--limit", type=int, default=None,
                        help="Limit the number of DOCX files to process.")
    parser.add_argument("--log", default="INFO", choices=['DEBUG', 'INFO', 'WARNING', 'ERROR', 'CRITICAL'],
                        help="Logging level.")

    args = parser.parse_args()

    # Set logging level based on argument.
    effective_log_level = getattr(logging, args.log.upper(), logging.INFO)
    logging.getLogger().setLevel(effective_log_level)
    global LOG_LEVEL
    LOG_LEVEL = effective_log_level

    log.info("=" * 20 + " Starting DOCX Complex Table Scan " + "=" * 20)
    log.info(f"Source Directory: {os.path.abspath(args.docx_directory)}")
    log.info(f"Output File: {os.path.abspath(args.output)}")
    log.info(f"Worker Processes: {args.workers}")
    if args.limit:
        log.info(f"Processing Limit: First {args.limit} DOCX file(s)")
    log.info(f"Logging Level: {args.log.upper()}")
    log.info("-" * 66)

    # Validate the source directory.
    if not os.path.isdir(args.docx_directory):
        log.critical(f"CRITICAL ERROR: Source directory not found: {args.docx_directory}")
        sys.exit(1)

    # Scan for DOCX files.
    log.info("Scanning for DOCX files...")
    docx_files = scan_docx_directory(args.docx_directory)
    total_docx_files = len(docx_files)
    log.info(f"Found {total_docx_files} DOCX file(s).")

    # Apply processing limit, if provided.
    if args.limit:
        if args.limit <= 0:
            log.warning("Processing limit is non-positive. Exiting.")
            sys.exit(0)
        docx_files = docx_files[:args.limit]
        log.info(f"Processing limited to first {len(docx_files)} file(s).")

    if not docx_files:
        log.info("No DOCX files to process.")
        sys.exit(0)

    # Process DOCX files using a process pool executor.
    detailed_docx_results = {}  # Maps each DOCX path to its analysis details.
    processed_count = 0
    total_tables_detected = 0
    processing_errors = 0
    start_time = time.time()

    log.info(f"Submitting {len(docx_files)} DOCX file(s) to process pool...")
    with ProcessPoolExecutor(max_workers=args.workers) as executor:
        future_to_docx = {
            executor.submit(analyze_docx_for_complex_tables, docx_path): docx_path
            for docx_path in docx_files
        }

        log.info("Waiting for results from workers...")
        for future in as_completed(future_to_docx):
            processed_count += 1
            docx_path = future_to_docx[future]
            docx_basename = os.path.basename(docx_path)
            try:
                complex_indices, table_count, table_details = future.result()

                if table_count >= 0:  # Successfully processed.
                    total_tables_detected += table_count
                    detailed_docx_results[docx_path] = {
                        "complex_indices": sorted(complex_indices),
                        "total_tables": table_count,
                        "table_details": table_details
                    }
                    if not complex_indices:
                        log.info(f"'{docx_basename}' does not contain any complex tables.")
                else:
                    log.warning(f"Worker indicated processing error for '{docx_basename}'.")
                    processing_errors += 1
            except Exception as error:
                log.error(f"Error retrieving result for '{docx_basename}': {error}",
                          exc_info=(LOG_LEVEL <= logging.DEBUG))
                processing_errors += 1

            # Progress reporting.
            log_interval = max(100, len(docx_files) // 100) if len(docx_files) > 100 else 50
            if processed_count % log_interval == 0 or processed_count == len(docx_files):
                elapsed_time = time.time() - start_time
                docs_per_sec = processed_count / elapsed_time if elapsed_time > 0 else 0
                log.info(f"Progress: {processed_count}/{len(docx_files)} files analyzed "
                         f"({docs_per_sec:.1f} files/sec). "
                         f"Complex DOCX files: {sum(1 for v in detailed_docx_results.values() if v['complex_indices'])}. "
                         f"Total tables: {total_tables_detected}. "
                         f"Errors: {processing_errors}.")

    # Processing summary.
    end_time = time.time()
    total_time = end_time - start_time
    successful_docs = processed_count - processing_errors
    avg_time_per_doc = total_time / processed_count if processed_count > 0 else 0

    log.info("=" * 25 + " Processing Summary " + "=" * 25)
    log.info(f"Total DOCX files submitted: {len(docx_files)}")
    log.info(f"Total DOCX files analyzed: {processed_count}")
    log.info(f"DOCX files with processing errors: {processing_errors}")
    log.info(f"Successfully analyzed DOCX files: {successful_docs}")
    log.info(f"Total tables detected: {total_tables_detected}")
    log.info(f"Total processing time: {total_time:.2f} seconds")
    log.info(f"Average time per DOCX file: {avg_time_per_doc:.3f} seconds")
    log.info("=" * 66)

    # Write detailed results to output file.
    write_results_to_file(detailed_docx_results, args.output)
    log.info("Script finished.")


if __name__ == "__main__":
    main()