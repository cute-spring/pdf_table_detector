import sys
import json
import re
from docx import Document
from docx.shared import Pt
from docx.table import _Cell, Table
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import requests
import logging
import time
# Import for concurrency
import concurrent.futures
from concurrent.futures import ThreadPoolExecutor, as_completed

# --- Configuration ---
DEFAULT_OLLAMA_URL = "http://localhost:11434/api/generate"
DEFAULT_OLLAMA_MODEL = "phi4" # Or your preferred model
DEFAULT_TABLE_STYLE_FALLBACK = 'Table Normal' # Usually safe fallback
# --- CONCURRENCY CONFIG ---
# Adjust based on your Ollama server's capacity and network conditions
MAX_LLM_WORKERS = 4 # Number of tables to process in parallel

# Setup logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# --- Helper Functions (extract_table..., parse_markdown_table, insert_table_after) ---
# These functions remain largely the same as in the previous version.
# We will focus on refactoring simplify_table_with_ollama slightly and the main process_docx loop.

def extract_table_to_textual_representation(table: Table):
    """
    Extracts table data into a simple textual representation. (No changes needed for speed here)
    """
    text_repr = []
    # Using list comprehensions MIGHT be marginally faster, but unlikely to be noticeable
    # compared to LLM calls or file I/O. Readability is good here.
    logging.debug(f"Extracting table with {len(table.rows)} rows and {len(table.columns)} columns.")
    try:
        grid = table._cells
        num_rows = len(table.rows)
        num_cols = len(table.columns)

        for i in range(num_rows):
            row_cells = []
            for j in range(num_cols):
                try:
                    cell_idx = i * num_cols + j
                    if cell_idx < len(grid):
                        cell = grid[cell_idx]
                        cell_text = cell.text.strip().replace('\n', ' ')
                        row_cells.append(cell_text)
                    else:
                         logging.warning(f"Cell index {cell_idx} out of bounds at ({i},{j})")
                         row_cells.append("(cell-missing?)")
                except Exception as cell_e:
                    logging.error(f"Error processing cell ({i},{j}): {cell_e}")
                    row_cells.append("(error reading cell)")
            text_repr.append(" | ".join(row_cells))
    except Exception as e:
        logging.error(f"Error extracting table content: {e}")
        return None
    return "\n".join(text_repr)


# --- LLM Function (Remains mostly the same, called by worker threads) ---
def simplify_table_with_ollama(table_text,
                              model_name=DEFAULT_OLLAMA_MODEL,
                              ollama_url=DEFAULT_OLLAMA_URL,
                              timeout=120):
    """
    Sends extracted table text to Ollama LLM for simplification. (No changes needed for speed here)
    """
    # This function is inherently I/O bound (network) and CPU bound (on Ollama server).
    # Optimizing the request structure itself won't speed up the external process.
    # Concurrency is the key.
    logging.info(f"Sending table data to Ollama model '{model_name}'...") # Shortened log

    prompt = f"""
You are an expert data processor. Below is a textual representation of a table extracted from a document.
The table may contain merged cells. In the representation, cells that were part of a vertical or horizontal merge might contain repeated text identical to the top-left cell of the merge area. Some might appear as empty strings if the source cell was empty.

Your task is to:
1. Analyze the structure and content of the table representation.
2. Intelligently "unmerge" all cells. Identify regions with repeated text (or potentially adjacent empty cells following text) that indicate a merge.
3. Fill in all cells that were part of an original merge by replicating the data from the conceptual top-left cell of that merged region. Ensure the text is consistent downwards for vertical merges and rightwards for horizontal merges.
4. Ensure the final table has a consistent number of columns in every row and conceptually no merged cells.
5. Output ONLY the resulting table content as a standard Markdown table. Do not include any explanations, introductions, or summaries before or after the Markdown table itself.

Original table representation:
--- START TABLE ---
{table_text}
--- END TABLE ---

Unmerged Markdown table:
"""

    data = {"model": model_name, "prompt": prompt, "stream": False}

    try:
        response = requests.post(ollama_url, json=data, timeout=timeout)
        response.raise_for_status()
        result = response.json()
        simplified_table_md = result.get('response', '').strip()

        if not simplified_table_md:
            logging.error("LLM response was empty.")
            return None

        logging.info("Received non-empty response from Ollama.")
        # Less verbose check
        if not simplified_table_md.startswith('|') and simplified_table_md.count('\n') < 1:
             logging.warning("LLM response might not be valid Markdown.")
             logging.debug(f"LLM Raw Response:\n{simplified_table_md}")

        return simplified_table_md

    except requests.exceptions.Timeout:
        logging.error(f"Ollama request timed out after {timeout} seconds.")
        return None
    except requests.exceptions.RequestException as e:
        logging.error(f"Ollama communication error: {e}")
        return None
    except Exception as e: # Catch JSONDecodeError etc.
        logging.error(f"Unexpected error during LLM interaction: {e}")
        return None


def parse_markdown_table(markdown_string):
    """
    Parses a Markdown table string into a list of lists (rows and cells). (No changes needed for speed here)
    """
    # This function is CPU-bound but likely very fast already (regex, string splits).
    # Micro-optimizations are unlikely to yield significant gains.
    logging.info("Parsing simplified Markdown table...")
    lines = markdown_string.strip().split('\n')
    data = []
    header_separator_found = False
    num_cols = 0
    # Regex can be pre-compiled if this function were called millions of times, but not necessary here.
    sep_pattern = re.compile(r"^\s*\|?\s*[-:|]+\s*\|?(\s*[-:|]+\s*\|?\s*)*$")

    for i, line in enumerate(lines):
        line = line.strip()
        if not line: continue # Skip empty lines efficiently

        is_separator = sep_pattern.match(line)

        if not is_separator and (not line.startswith('|') or not line.endswith('|')):
            logging.debug(f"Skipping line {i+1} (not table data/separator): {line}")
            continue

        if is_separator:
            if not header_separator_found:
                 logging.debug(f"Found header separator at line {i+1}")
                 header_separator_found = True
                 if data: # Header row should be the last element added
                     num_cols = len(data[-1])
                     logging.debug(f"Determined {num_cols} columns from header.")
                 # simplified inferring logic
            continue

        # Extract cells - list comprehension can be slightly faster
        cells = [cell.strip() for cell in line.strip('|').split('|')]

        if not cells: continue # Skip lines that parse to empty cells list

        if num_cols == 0: # First real data/header row
            num_cols = len(cells)
            logging.debug(f"Set column count to {num_cols} from first row.")

        # Row consistency check
        if len(cells) != num_cols:
            logging.warning(f"Line {i+1}: Inconsistent column count ({len(cells)} vs {num_cols}). Adjusting.")
            if len(cells) < num_cols: cells.extend([""] * (num_cols - len(cells)))
            else: cells = cells[:num_cols]
        data.append(cells)


    if not data:
        logging.warning("Could not parse any data rows from Markdown.")
        return None
    if num_cols == 0 and data: # If only one row was added
         num_cols = len(data[0])
         logging.debug(f"Set column count to {num_cols} from single row.")

    logging.info(f"Parsed {len(data)} rows, {num_cols} columns.")
    return data

def insert_table_after(document, ref_table_element, data, original_style_name=None):
    """
    Inserts a new table populated with data after ref_table_element. (No changes needed for speed here)
    """
    # This involves OXML manipulation (python-docx API and direct OXML insertion).
    # These operations are inherently somewhat slow but necessary. Optimizing them
    # significantly would require deeper changes, possibly bypassing python-docx API.
    # The current structure with fallbacks is robust.

    if not data or not data[0]:
        logging.error("No data provided to insert_table_after.")
        return None

    num_rows = len(data)
    num_cols = len(data[0])
    if num_cols == 0:
        logging.error("Cannot insert table with 0 columns.")
        return None
    logging.info(f"Preparing to insert table: {num_rows} rows, {num_cols} cols.")

    new_table = None
    applied_style = None
    new_table_element = None
    safe_rows = max(1, num_rows)
    safe_cols = max(1, num_cols)

    # --- Try creating table object ---
    # (Keeping the try/except structure for style handling robustness)
    # ... (Style application logic remains the same) ...
    try:
        if original_style_name:
            try:
                new_table = document.add_table(rows=safe_rows, cols=safe_cols, style=original_style_name)
                applied_style = original_style_name
                logging.info(f"Created table obj with style '{applied_style}'.")
            except KeyError: new_table = None # Fall through
            except Exception: new_table = None # Fall through
        if new_table is None and DEFAULT_TABLE_STYLE_FALLBACK:
             try:
                new_table = document.add_table(rows=safe_rows, cols=safe_cols, style=DEFAULT_TABLE_STYLE_FALLBACK)
                applied_style = DEFAULT_TABLE_STYLE_FALLBACK
                logging.info(f"Created table obj with style '{applied_style}'.")
             except KeyError: new_table = None # Fall through
             except Exception: new_table = None # Fall through
        if new_table is None:
             new_table = document.add_table(rows=safe_rows, cols=safe_cols)
             applied_style = "[Document Default]"
             logging.info(f"Created table obj with {applied_style}.")

    except Exception as table_creation_e:
        logging.error(f"Critical failure during table object creation: {table_creation_e}", exc_info=True)
        return None
    # --- End table creation ---

    if new_table is None: # Should be redundant after above, but safety check
        logging.error("Table object is None after creation attempts.")
        return None

    new_table_element = new_table._element

    # --- Populate and Insert ---
    try:
        new_table.autofit = True
        # --- Populate Table Data ---
        if num_rows == 0 and len(new_table.rows) > 0: # Clear placeholder
             for cell_obj in new_table.rows[0].cells:
                 for para in list(cell_obj.paragraphs):
                    p_element = para._element
                    if p_element.getparent() is not None: p_element.getparent().remove(p_element)
                 cell_obj.add_paragraph("")
        elif num_rows > 0:
            logging.debug(f"Populating {num_rows} rows...")
            # Directly access rows and cells by index might be marginally faster
            # than iterating through `new_table.rows` and `row.cells` properties
            # but python-docx likely optimizes this internally. Let's keep clarity.
            all_rows = new_table.rows
            if len(all_rows) < num_rows:
                logging.warning(f"Table created with fewer rows ({len(all_rows)}) than requested ({num_rows}). Populating available rows.")
                num_rows = len(all_rows) # Adjust loop range

            # Pre-compile regex for cleaner code, minimal speed impact here
            clean_pattern = re.compile(r'[*_`]')
            for i in range(num_rows):
                row_data = data[i]
                row_obj = all_rows[i]
                row_cells = row_obj.cells
                if not row_cells: continue # Skip rows that somehow lack cells
                num_cells_in_row = len(row_cells)

                for j, cell_text in enumerate(row_data):
                    if j < num_cells_in_row:
                        cleaned_text = clean_pattern.sub('', str(cell_text))
                        cell_obj = row_cells[j]
                        # Optimisation: Instead of iterating list(paras), get paras once? Unlikely much diff.
                        for para in list(cell_obj.paragraphs):
                            p_element = para._element
                            if p_element.getparent() is not None: p_element.getparent().remove(p_element)
                        cell_obj.add_paragraph(cleaned_text)
                    else:
                         # Only log once per row if columns mismatch
                         if j == num_cells_in_row:
                             logging.warning(f"Row {i}: Data columns ({len(row_data)}) > table columns ({num_cells_in_row}). Truncating.")
                         break # Stop processing cells for this row

        logging.info("Finished populating table data.")

        # --- OXML Insertion ---
        logging.debug("Performing OXML insertion...")
        parent_element = ref_table_element.getparent()
        if parent_element is None: raise RuntimeError("Original table has no parent")
        original_table_index = parent_element.index(ref_table_element)

        # Insert table then paragraph
        parent_element.insert(original_table_index + 1, new_table_element)
        p = OxmlElement("w:p")
        parent_element.insert(original_table_index + 1, p)

        logging.info("OXML insertion complete.")
        return new_table

    except Exception as e_insert_populate:
        logging.error(f"Error during populate/insert: {e_insert_populate}", exc_info=True)
        # Cleanup attempt (same as before)
        if new_table_element is not None:
            current_parent = new_table_element.getparent()
            if current_parent is not None:
                try: current_parent.remove(new_table_element); logging.info("Cleaned up table element.")
                except Exception as cl_e: logging.error(f"Cleanup error: {cl_e}")
            else: logging.warning("No parent for cleanup.")
        return None

# --- New function to handle threaded execution ---
def process_single_table_llm(table_index, table_text, model, url, timeout):
    """Worker function: Calls LLM and returns index and result/error."""
    try:
        simplified_md = simplify_table_with_ollama(table_text, model, url, timeout)
        if simplified_md:
            return table_index, simplified_md, None # Success: index, data, no error
        else:
            return table_index, None, "LLM returned empty response" # Failure: index, no data, error msg
    except Exception as e:
        logging.error(f"Exception in LLM worker for table index {table_index}: {e}", exc_info=True)
        return table_index, None, str(e) # Failure: index, no data, error msg


# --- REFACTORED Main Workflow ---
def process_docx(input_path, output_path, ollama_model=DEFAULT_OLLAMA_MODEL, ollama_url=DEFAULT_OLLAMA_URL):
    """
    Processes DOCX using concurrent LLM calls for speed.
    """
    start_time = time.time()
    logging.info(f"Starting processing for DOCX: '{input_path}'")
    try:
        document = Document(input_path)
    except Exception as e:
        logging.error(f"Failed to open DOCX file '{input_path}': {e}")
        return

    if not document.tables:
        logging.info("No tables found in the document.")
        return

    logging.info(f"Found {len(document.tables)} tables. Preparing for concurrent processing.")
    tables_processed_count = 0
    tables_failed_before_insertion = 0 # Track failures *before* insertion attempt
    tables_insertion_failed = 0    # Track failures *during* insertion

    # --- Stage 1: Extract data and Submit LLM tasks ---
    original_tables_data = {} # Store info needed *after* LLM call
    llm_tasks = []
    futures_map = {} # Map future object back to table index

    with ThreadPoolExecutor(max_workers=MAX_LLM_WORKERS) as executor:
        logging.info(f"Submitting {len(document.tables)} table(s) to LLM using up to {MAX_LLM_WORKERS} workers...")
        for i, table in enumerate(document.tables):
            if table is None or not hasattr(table, '_element'):
                logging.warning(f"Skipping invalid table object at index {i}.")
                continue

            logging.debug(f"Table {i+1}: Extracting text...")
            table_text_repr = extract_table_to_textual_representation(table)

            if table_text_repr is None:
                logging.error(f"Table {i+1}: Failed text extraction. Skipping LLM call.")
                tables_failed_before_insertion += 1
                continue

            # Store data needed for insertion *later*
            style_name = None
            try:
                # Simplified style detection logic
                if hasattr(table, 'style') and table.style and hasattr(table.style, 'name') and table.style.name:
                     s_name = table.style.name
                     norm_s_name = s_name.lower().strip()
                     if norm_s_name not in ["none", "table normal", "normal table", "grid table"]:
                         style_name = s_name
                         logging.debug(f"Table {i+1}: Detected style '{style_name}'.")
            except Exception as e:
                 logging.warning(f"Table {i+1}: Error getting style name ({e}). Using fallback.")

            original_tables_data[i] = {
                'element': table._element, # Store OXML element reference
                'style_name': style_name
            }

            # Submit LLM task to the executor
            future = executor.submit(process_single_table_llm, i, table_text_repr, ollama_model, ollama_url, 120)
            llm_tasks.append(future)
            futures_map[future] = i # Map future -> index for easy lookup

        logging.info(f"Submitted {len(llm_tasks)} tasks to LLM executor.")

        # --- Stage 2: Process completed LLM tasks and Insert sequentially ---
        logging.info("Waiting for LLM responses and processing results...")
        llm_results = {} # Store successful LLM results mapped by index

        for future in as_completed(llm_tasks):
            table_index = futures_map[future] # Get original index
            try:
                idx, simplified_md, error_msg = future.result()
                if error_msg:
                    logging.error(f"Table {table_index+1}: LLM processing failed: {error_msg}")
                    tables_failed_before_insertion += 1
                elif simplified_md:
                    logging.info(f"Table {table_index+1}: Received LLM result.")
                    llm_results[idx] = simplified_md # Store successful result
                else:
                    # This case should ideally be covered by error_msg, but for safety:
                    logging.error(f"Table {table_index+1}: LLM worker returned unexpected empty result without error.")
                    tables_failed_before_insertion += 1

            except Exception as e:
                logging.error(f"Table {table_index+1}: Error retrieving result from future: {e}", exc_info=True)
                tables_failed_before_insertion += 1


    # --- Stage 3: Sequential Insertion ---
    logging.info("Starting sequential insertion of successfully processed tables...")
    # Process in original table order for deterministic insertion relative to original positions
    insertion_order = sorted(llm_results.keys())
    logging.debug(f"Insertion order: {insertion_order}")

    for table_index in insertion_order:
        logging.info(f"--- Processing Insertion for Original Table Index {table_index+1} ---")
        simplified_markdown = llm_results[table_index]
        original_data = original_tables_data.get(table_index)

        if not original_data:
            logging.error(f"Table {table_index+1}: Missing original data for insertion. Skipping.")
            # This shouldn't happen if extraction succeeded, but good safety check
            tables_insertion_failed += 1
            continue

        original_table_element = original_data['element']
        original_style_name = original_data['style_name']

        # Parse Markdown
        logging.debug("Parsing Markdown...")
        simplified_data = parse_markdown_table(simplified_markdown)
        if not simplified_data:
            logging.error(f"Table {table_index+1}: Failed to parse Markdown from LLM. Skipping insertion.")
            tables_failed_before_insertion += 1 # Count as pre-insertion failure
            continue

        # Insert Table
        logging.info(f"Table {table_index+1}: Calling insert_table_after...")
        inserted_table_obj = insert_table_after(
            document,
            original_table_element,
            simplified_data,
            original_style_name
        )

        if inserted_table_obj is None:
            logging.error(f"Table {table_index+1}: insert_table_after FAILED.")
            tables_insertion_failed += 1
        elif isinstance(inserted_table_obj, Table):
             logging.info(f"Table {table_index+1}: insert_table_after SUCCEEDED.")
             tables_processed_count += 1
        else:
             logging.error(f"Table {table_index+1}: insert_table_after returned unexpected type {type(inserted_table_obj)}.")
             tables_insertion_failed += 1
        logging.info(f"--- Finished Insertion for Original Table Index {table_index+1} ---")


    # --- Final Save ---
    end_time = time.time()
    duration = end_time - start_time
    logging.info(f"Processing finished in {duration:.2f} seconds.")
    logging.info(f"Summary: Processed={tables_processed_count}, Pre-Insert Fail={tables_failed_before_insertion}, Insert Fail={tables_insertion_failed}")

    try:
        logging.info(f"Attempting to save document to '{output_path}'...")
        document.save(output_path)
        logging.info(f"Successfully saved updated document.")
    except Exception as e:
        logging.error(f"Failed to save the final DOCX file '{output_path}': {e}", exc_info=True)


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Simplify complex tables in DOCX using Ollama LLM (Concurrent), reusing original table styles.")
    parser.add_argument("input_docx", help="Path to the input DOCX file.")
    parser.add_argument("output_docx", help="Path to save the modified DOCX file.")
    parser.add_argument("--model", default=DEFAULT_OLLAMA_MODEL, help=f"Ollama model name (default: {DEFAULT_OLLAMA_MODEL}).")
    parser.add_argument("--url", default=DEFAULT_OLLAMA_URL, help=f"Ollama API URL (default: {DEFAULT_OLLAMA_URL}).")
    # Add argument for concurrency level
    parser.add_argument("--workers", type=int, default=MAX_LLM_WORKERS, help=f"Number of parallel LLM workers (default: {MAX_LLM_WORKERS}).")
    parser.add_argument("--debug", action='store_true', help="Enable debug logging.")

    args = parser.parse_args()

    # Set global concurrency level from args
    MAX_LLM_WORKERS = args.workers

    if args.debug:
        logging.getLogger().setLevel(logging.DEBUG)
        logging.debug("Debug logging enabled.")
    else:
        logging.getLogger().setLevel(logging.INFO)

    process_docx(
        input_path=args.input_docx,
        output_path=args.output_docx,
        ollama_model=args.model,
        ollama_url=args.url
        # Concurrency is handled via the global MAX_LLM_WORKERS set from args.workers
    )