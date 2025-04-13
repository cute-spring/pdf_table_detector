import sys
import json
import re
from docx import Document
from docx.shared import Pt
from docx.table import _Cell
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import requests
import logging

# --- Configuration ---
# Default Ollama API endpoint
DEFAULT_OLLAMA_URL = "http://localhost:11434/api/generate"
# Default Ollama model to use
DEFAULT_OLLAMA_MODEL = "llama3" # Or "mistral", "llama2", etc.
# Index of the table to process (0 for the first table, 1 for the second, etc.)
TARGET_TABLE_INDEX = 0
# Style to apply to the new table
NEW_TABLE_STYLE = 'Table Grid'

# Setup logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# --- Helper Functions ---

def extract_table_to_textual_representation(table):
    """
    Extracts table data into a simple textual representation, handling merged cells implicitly.
    Attempts to create a basic structure that the LLM can interpret.
    Note: This is a best-effort extraction due to python-docx limitations with easily querying merge info.
    """
    text_repr = []
    logging.info(f"Extracting table with {len(table.rows)} rows and {len(table.columns)} columns.")
    for i, row in enumerate(table.rows):
        row_cells = []
        try:
            # Using row.cells provides the visible cells in the row.
            # Merged cells (except the top-left one) aren't directly listed here.
            # We capture the text from the available cells.
            cell_texts = [cell.text.strip().replace('\n', ' ') for cell in row.cells]
            # Attempt to add placeholders for missing cells due to horizontal merges, very basic
            expected_cols = len(table.columns)
            actual_cells = len(cell_texts)
            # This logic is basic and might not perfectly represent all merges
            row_cells.extend(cell_texts)
            if actual_cells < expected_cols:
                 # Placeholder logic - assumes missing cells are at the end. Not robust.
                 row_cells.extend(["(merged/empty?)"] * (expected_cols - actual_cells))
            elif actual_cells > expected_cols:
                 # Unlikely but possible with odd structures
                 row_cells = row_cells[:expected_cols]

            text_repr.append(" | ".join(row_cells))
        except Exception as e:
            logging.error(f"Error processing row {i}: {e}")
            text_repr.append(" | ".join([f"(error reading cell {j})" for j in range(len(table.columns))]))

    return "\n".join(text_repr)

import requests
import json
import logging

# Constants remain useful for defining defaults easily
DEFAULT_OLLAMA_URL = "http://localhost:11434/api/generate"
DEFAULT_OLLAMA_MODEL = "phi4"

# Setup logging if not already done elsewhere
# logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def simplify_table_with_ollama(table_text,
                              model_name=DEFAULT_OLLAMA_MODEL,
                              ollama_url=DEFAULT_OLLAMA_URL,
                              timeout=120): # Added timeout as a parameter
    """
    Sends the extracted table text to the Ollama LLM for simplification,
    following a similar structure to the compare_document_versions example.

    Args:
        table_text (str): The textual representation of the table to simplify.
        model_name (str): The Ollama model to use.
        ollama_url (str): The URL of the Ollama API endpoint.
        timeout (int): Request timeout in seconds.

    Returns:
        str: The simplified table in Markdown format, or None if an error occurs.
    """
    logging.info(f"Sending table data to Ollama model '{model_name}' at {ollama_url}")

    # Keep the detailed, task-specific prompt
    prompt = f"""
You are an expert data processor. Below is a textual representation of a table extracted from a document.
This table likely contains merged cells. Merged cells might be represented by repeated content, empty placeholders like '(merged/empty?)', or just missing content in some rows compared to the expected column count.

Your task is to:
1. Analyze the structure and content of the table.
2. 'Unmerge' all cells. Fill in the cells that were part of a merge by replicating the data from the top-left cell of the original merged region both downwards and rightwards as appropriate.
3. Ensure the final table has a consistent number of columns in every row and no merged cells.
4. Output ONLY the simplified table content as a standard Markdown table. Do not include any explanations, apologies, or introductory text before or after the Markdown table.

Original table representation:
--- START TABLE ---
{table_text}
--- END TABLE ---

Simplified Markdown table:
"""

    # Use 'data' for payload variable name consistency with example
    data = {
        "model": model_name,
        "prompt": prompt,
        "stream": False  # Get the full response at once
    }

    try:
        response = requests.post(ollama_url, json=data, timeout=timeout)
        # Use raise_for_status for concise checking of HTTP errors (4xx, 5xx)
        response.raise_for_status()

        # Process the successful response
        result = response.json()
        simplified_table_md = result.get('response', '').strip() # Use .strip() to clean whitespace

        if not simplified_table_md:
            logging.error("LLM response was empty.")
            return None # Return None on logical failure (empty response)

        logging.info("Received non-empty response from Ollama.")

        # Optional: Keep the basic Markdown check as a warning
        if not simplified_table_md.startswith('|') and not simplified_table_md.count('|') > 1:
             logging.warning("LLM response might not be a Markdown table. Check output carefully.")
             logging.debug(f"LLM Raw Response:\n{simplified_table_md}")

        return simplified_table_md # Return the successful result

    # Keep specific exception handling for better diagnostics
    except requests.exceptions.Timeout:
        logging.error(f"Request to Ollama timed out after {timeout} seconds.")
        return None
    except requests.exceptions.ConnectionError:
        logging.error(f"Could not connect to Ollama at {ollama_url}. Is it running?")
        return None
    except requests.exceptions.RequestException as e:
        # Catches other request errors (like HTTPError handled by raise_for_status)
        logging.error(f"Error communicating with Ollama: {e}")
        return None
    except json.JSONDecodeError as e:
        logging.error(f"Error decoding JSON response from Ollama: {e}")
        logging.debug(f"Raw response content: {response.text if 'response' in locals() else 'Response object not available'}")
        return None
    except Exception as e:
        # Catch any other unexpected errors during the process
        logging.error(f"An unexpected error occurred during LLM interaction: {e}")
        return None

# --- Example Usage (within the context of the larger script) ---
# In your main process_docx function, you would call it like this:
#
# simplified_markdown = simplify_table_with_ollama(
#     table_text_repr,
#     model_name=ollama_model, # Pass parameters from args or defaults
#     ollama_url=ollama_url
# )
# if not simplified_markdown:
#     logging.error("Failed to get simplified table from LLM.")
#     return
# # ... continue processing ...

def parse_markdown_table(markdown_string):
    """
    Parses a Markdown table string into a list of lists (rows and cells).
    """
    logging.info("Parsing simplified Markdown table from LLM response.")
    lines = markdown_string.strip().split('\n')
    data = []

    # Find the start of the actual table data (skip potential blank lines or separators)
    first_data_row_index = 0
    for i, line in enumerate(lines):
        line = line.strip()
        if not line:
            continue # Skip blank lines
        if line.startswith('|') and line.endswith('|'):
            # Check if it's the separator line (e.g., |---|---|)
            if re.match(r"^\s*\|?\s*[-:]+\s*\|?(\s*[-:]+\s*\|?\s*)*$", line):
                first_data_row_index = i + 1 # Data starts after separator
                continue
            # Assume this is a data or header row
            if first_data_row_index == 0: # Found first content row (might be header)
                first_data_row_index = i

            cells = [cell.strip() for cell in line.strip('|').split('|')]
            data.append(cells)

    if not data:
        logging.warning("Could not parse any data rows from the Markdown.")
        return None

    # Validate column consistency
    if len(data) > 1:
        num_cols = len(data[0])
        for i, row in enumerate(data[1:], 1):
            if len(row) != num_cols:
                logging.warning(f"Inconsistent number of columns found in parsed Markdown. Row {i+1} has {len(row)} cells, expected {num_cols}. Truncating/padding.")
                # Attempt basic correction: Pad with empty strings or truncate
                if len(row) < num_cols:
                    row.extend([""] * (num_cols - len(row)))
                else:
                    data[i] = row[:num_cols]

    logging.info(f"Parsed {len(data)} rows and {len(data[0]) if data else 0} columns.")
    return data

def insert_table_after(document, ref_table_element, data):
    """
    Inserts a new table populated with data immediately after the reference table element.
    Handles cases where the primary desired style (NEW_TABLE_STYLE) is not found.
    """
    if not data or not data[0]:
        logging.error("No data provided to create the new table.")
        return None

    num_rows = len(data)
    num_cols = len(data[0])
    logging.info(f"Preparing to create a new table with {num_rows} rows and {num_cols} columns.")

    new_table = None # Initialize new_table to None

    try:
        # Attempt 1: Try creating the table with the desired style
        logging.info(f"Attempting to apply style: '{NEW_TABLE_STYLE}'")
        new_table = document.add_table(rows=num_rows, cols=num_cols, style=NEW_TABLE_STYLE)
        logging.info(f"Successfully created table with style '{NEW_TABLE_STYLE}'.")
    except KeyError:
        logging.warning(f"Table style '{NEW_TABLE_STYLE}' not found in the document.")
        try:
            # Attempt 2: Fallback to 'Table Normal' (usually present)
            fallback_style = 'Table Normal'
            logging.info(f"Attempting fallback style: '{fallback_style}'")
            new_table = document.add_table(rows=num_rows, cols=num_cols, style=fallback_style)
            logging.info(f"Successfully created table with fallback style '{fallback_style}'.")
        except KeyError:
            logging.warning(f"Fallback style '{fallback_style}' also not found.")
            try:
                # Attempt 3: Fallback to using the document's default style (omit style argument)
                logging.info("Attempting to create table with document's default style.")
                new_table = document.add_table(rows=num_rows, cols=num_cols)
                logging.info("Successfully created table with default style.")
            except Exception as e:
                # If even default creation fails, log error and exit function
                logging.error(f"Failed to create table even with document's default style: {e}")
                return None
        except Exception as e:
             # Catch potential errors during fallback style creation
            logging.error(f"Failed to create table with fallback style '{fallback_style}': {e}")
            return None
    except Exception as e:
         # Catch potential errors during primary style creation (other than KeyError)
        logging.error(f"Failed to create table with primary style '{NEW_TABLE_STYLE}': {e}")
        return None

    # --- If table creation succeeded one way or another, proceed ---
    if new_table is None:
        # This should theoretically not be reached if the fallbacks work, but acts as a safeguard
        logging.error("Table creation failed through all attempts.")
        return None

    # Apply properties and populate the table
    new_table.autofit = True

    logging.info("Populating new table data.")
    for i, row_data in enumerate(data):
        # Ensure row exists (should be created by add_table)
        if i >= len(new_table.rows):
             logging.warning(f"Row index {i} out of bounds for newly created table ({len(new_table.rows)} rows). Skipping row population.")
             continue
        row_cells = new_table.rows[i].cells
        for j, cell_text in enumerate(row_data):
             # Check cell bounds - safety precaution
             if j < len(row_cells):
                cleaned_text = re.sub(r'[*_`]', '', cell_text)
                cell_obj = row_cells[j]
                # Clear existing default paragraph before adding text
                for para in cell_obj.paragraphs:
                    p_element = para._element
                    p_element.getparent().remove(p_element)
                # Add new paragraph with text
                para = cell_obj.add_paragraph(cleaned_text)
             else:
                 logging.warning(f"Data column index {j} exceeds table column count {len(row_cells)} in row {i}. Skipping cell.")


    logging.info("Populated new table data.")

    # --- XML Manipulation for Insertion ---
    # Find the paragraph element immediately following the original table element
    original_table_element = ref_table_element
    new_table_element = new_table._element # Get the OXML element of the newly created table

    parent_element = original_table_element.getparent()
    if parent_element is None:
        logging.error("Could not find the parent element of the original table. Cannot insert.")
        # Remove the table added at the end if we can't move it.
        new_table_element.getparent().remove(new_table_element)
        return None

    # Find the original table's position within the parent
    original_table_index = parent_element.index(original_table_element)

    # Insert the new table's element right after the original table
    parent_element.insert(original_table_index + 1, new_table_element)

    # Add a blank paragraph between the tables for spacing (optional)
    p = OxmlElement("w:p")
    # Insert the paragraph *before* the new table (which is now at index+1 relative to original index)
    # This correctly places it between the original table and the new table
    parent_element.insert(original_table_index + 1, p)

    logging.info("Moved new table immediately after the original table in the document structure, separated by a paragraph.")

    return new_table # Return the created and inserted table object

# --- Main Workflow ---
def process_docx(input_path, output_path, ollama_model=DEFAULT_OLLAMA_MODEL, ollama_url=DEFAULT_OLLAMA_URL):
    """
    Main function to process the DOCX file, iterating through ALL tables.
    """
    logging.info(f"Starting processing for DOCX: '{input_path}'")
    try:
        document = Document(input_path)
    except Exception as e:
        logging.error(f"Failed to open DOCX file '{input_path}': {e}")
        return

    if not document.tables:
        logging.info("No tables found in the document. Saving unchanged.")
        # Optional: Save a copy anyway or just exit
        # try:
        #     document.save(output_path)
        # except Exception as e:
        #     logging.error(f"Failed to save the (unchanged) DOCX file '{output_path}': {e}")
        return

    logging.info(f"Found {len(document.tables)} tables in the document. Processing all.")
    tables_processed_count = 0
    tables_failed_count = 0

    # --- Iterate through all tables ---
    # Use list(document.tables) to create a static copy in case insertion affects the live list ordering (unlikely but safer)
    all_original_tables = list(document.tables)
    for i, original_table in enumerate(all_original_tables):
        logging.info(f"--- Processing Table {i+1} of {len(all_original_tables)} ---")

        # Ensure we have a valid table object before getting _element
        if original_table is None or not hasattr(original_table, '_element'):
             logging.warning(f"Skipping table at index {i} as it seems invalid or corrupted.")
             tables_failed_count += 1
             continue

        original_table_element = original_table._element # Get element for insertion reference

        # 1. Extract Table Content
        table_text_repr = extract_table_to_textual_representation(original_table)
        if not table_text_repr:
            logging.error(f"Failed to extract content from table {i+1}. Skipping this table.")
            tables_failed_count += 1
            continue # Skip to the next table

        logging.debug(f"Extracted Text Representation (Table {i+1}):\n{table_text_repr}")

        # 2. Simplify Table via LLM
        simplified_markdown = simplify_table_with_ollama(table_text_repr, ollama_model, ollama_url)
        if not simplified_markdown:
            logging.error(f"Failed to get simplified table from LLM for table {i+1}. Skipping this table.")
            tables_failed_count += 1
            continue # Skip to the next table

        logging.debug(f"Simplified Markdown received (Table {i+1}):\n{simplified_markdown}")

        # 3. Parse LLM Response (Markdown Table)
        simplified_data = parse_markdown_table(simplified_markdown)
        if not simplified_data:
            logging.error(f"Failed to parse Markdown table from LLM response for table {i+1}. Skipping this table.")
            tables_failed_count += 1
            continue # Skip to the next table

        # 4. Insert Simplified Table
        inserted_table = insert_table_after(document, original_table_element, simplified_data)

        if inserted_table is None:
            logging.error(f"Failed to insert the new simplified table for original table {i+1}.")
            # Don't necessarily count this as a *processing* failure unless insertion is critical
            # Decide if you want to increment tables_failed_count here
            # tables_failed_count += 1 # Optional: uncomment if insertion failure is critical
            # continue # Continue processing other tables even if insertion fails for one
        else:
            logging.info(f"Successfully inserted simplified table after original table {i+1}.")
            tables_processed_count += 1

        logging.info(f"--- Finished Processing Table {i+1} ---")
        # Add a small visual separator in logs if processing many tables
        if i < len(all_original_tables) - 1:
            logging.info("-" * 20)


    # --- End of Loop ---

    # 5. Save Updated Document
    logging.info(f"Finished processing all tables. Processed successfully: {tables_processed_count}, Failed/Skipped: {tables_failed_count}.")
    try:
        document.save(output_path)
        logging.info(f"Successfully saved updated document with processed tables to '{output_path}'")
    except Exception as e:
        logging.error(f"Failed to save the final updated DOCX file '{output_path}': {e}")

if __name__ == "__main__":
    # --- Configuration from Command Line (Optional) ---
    # Example usage: python script_name.py input.docx output.docx --model phi4
    import argparse
    parser = argparse.ArgumentParser(description="Simplify complex tables in DOCX using Ollama LLM.")
    parser.add_argument("input_docx", help="Path to the input DOCX file.")
    parser.add_argument("output_docx", help="Path to save the modified DOCX file.")
    # Remove the --index argument
    # parser.add_argument("--index", type=int, default=TARGET_TABLE_INDEX, help=f"Index of the table to process (default: {TARGET_TABLE_INDEX}).")
    parser.add_argument("--model", default=DEFAULT_OLLAMA_MODEL, help=f"Ollama model name (default: {DEFAULT_OLLAMA_MODEL}).")
    parser.add_argument("--url", default=DEFAULT_OLLAMA_URL, help=f"Ollama API URL (default: {DEFAULT_OLLAMA_URL}).")
    parser.add_argument("--debug", action='store_true', help="Enable debug logging.")

    args = parser.parse_args()

    if args.debug:
        logging.getLogger().setLevel(logging.DEBUG)

    # Remove TARGET_TABLE_INDEX global variable if not used elsewhere
    # TARGET_TABLE_INDEX = 0 # This is no longer needed for the primary logic

    # --- Run the process ---
    # Remove the table_index argument from the call
    process_docx(
        input_path=args.input_docx,
        output_path=args.output_docx,
        # table_index=args.index, # Remove this line
        ollama_model=args.model,
        ollama_url=args.url
    )