import requests
import json
import time
import os
import re
import sys
import docx
from docx.document import Document
# NOTE: Importing specific XML types for isinstance checks
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table, _Row, _Cell # _Cell is needed for type hinting set_cell_text
from docx.text.paragraph import Paragraph
from urllib.parse import urlparse, parse_qs, urlencode, urlunparse

# --- Google Gemini Setup ---
try:
    from google import genai
    from google.genai import types
    # IMPORTANT: Use environment variables or a more secure method for API keys
    GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "YOUR_API_KEY_HERE") # Replace fallback with your key if not using env var
    if GEMINI_API_KEY == "YOUR_API_KEY_HERE":
        print("Warning: GEMINI_API_KEY not set as environment variable. Using placeholder.")
    # Configure the client (adjust model and options as needed)
    gemini_model = genai.Client(api_key="AIzaSyA_OdmrUqEp1levy3LhwjAI1m-d7nU6hSE", http_options={'api_version':'v1alpha'})
    print(f"Gemini AI client initialized.") # Model used: {gemini_model.model_name}") # Check attribute if needed

except ImportError:
    print("Error: Google GenAI library not installed. pip install google-generativeai")
    gemini_model = None # Set model to None
except Exception as e:
    print(f"Error initializing Gemini AI client: {e}")
    gemini_model = None # Set model to None

# --- Force UTF-8 ---
try:
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')
except AttributeError:
    print("Warning: Could not reconfigure stdout/stderr encoding (not supported on this platform/version?).")


# --- TestRail API Client (from testCaseGetter.py) ---
class TestRailAPIClient:
    def __init__(self, url: str, email: str, api_key: str):
        if not url.strip().startswith(('http://', 'https://')):
            raise ValueError(f"Invalid TestRail URL: '{url}'. Must start with http:// or https://")
        self.base_url = url.strip().rstrip('/')
        self.email = email
        self.api_key = api_key
        self.session = requests.Session()
        self.session.auth = (email, api_key)
        self.session.headers.update({'Content-Type': 'application/json'})
        print(f"TestRailAPIClient initialized for URL: {self.base_url}")

    def _make_request(self, method: str, full_url: str, **kwargs) -> requests.Response | None:
        """Internal helper to make requests. Returns Response object on success, None on error."""
        try:
            timeout = kwargs.pop('timeout', 60)
            response = self.session.request(method, full_url, timeout=timeout, **kwargs)
            response.raise_for_status()
            return response
        except requests.exceptions.HTTPError as http_err:
            status_code = http_err.response.status_code
            try:
                # Try to decode error response for better logging
                response_text = http_err.response.json()
            except json.JSONDecodeError:
                 response_text = http_err.response.text[:200].encode(sys.stdout.encoding, errors='replace').decode(sys.stdout.encoding)

            error_map = {
                400: "Info: Received 400 Bad Request. Check endpoint/parameters.",
                401: "Error: Received 401 Unauthorized. Check email/API key.",
                403: "Error: Received 403 Forbidden. Check permissions or API key.",
                404: "Info: Received 404 Not Found. Resource likely doesn't exist.",
                429: "Error: Received 429 Too Many Requests (Rate Limit). Consider adding delays."
            }
            message = error_map.get(status_code, f"HTTP error during request.")
            print(f"{message} URL: {method} {full_url} - Status: {status_code}. Response: {response_text}...")
            return None
        except requests.exceptions.Timeout:
            print(f"Error: Request timed out for {method} {full_url}")
            return None
        except requests.exceptions.RequestException as e:
            print(f"Error during network request to {method} {full_url}: {e}")
            return None

    def send_get_request(self, endpoint: str):
        """Sends a GET request for single objects (like a specific test case)."""
        print(f"--- Sending SINGLE GET request ---")
        clean_endpoint = endpoint.lstrip('/')
        full_url = f"{self.base_url}/index.php?/api/v2/{clean_endpoint}"
        print(f"  Target: {full_url}")
        response = self._make_request('GET', full_url)

        if response:
            try:
                 if response.status_code == 204 or not response.content:
                      print(f"  Success: Received status {response.status_code} with no content.")
                      return None
                 try:
                     response_data = response.json()
                 except json.JSONDecodeError:
                     print("  Warning: Initial JSON decode failed. Trying UTF-8 decoding.")
                     response_data = json.loads(response.content.decode('utf-8'))

                 print(f"  Success: Received single object data.")
                 return response_data
            except (json.JSONDecodeError, UnicodeDecodeError) as decode_err:
                print(f"  Error: Could not decode response from {response.url}. Status: {response.status_code}. Error: {decode_err}")
                print(f"  Response text (first 500 bytes): {response.content[:500]}...")
                return None
        else:
            print(f"  Failed to get response for single request.")
            return None

# --- Helper Functions ---

def set_cell_text(cell: _Cell, text: str):
    """ Safely sets the text in a cell, replacing existing content. Handles multi-line text. """
    if not isinstance(cell, _Cell):
        print(f"  Error: Invalid object passed to set_cell_text. Expected a Cell, got {type(cell)}.")
        return
    try:
        # Clear existing content more robustly
        cell.text = "" # Clear simple text
        # Remove all but the first paragraph
        while len(cell.paragraphs) > 1:
            p_to_remove = cell.paragraphs[-1]._element
            if p_to_remove.getparent() is not None:
                p_to_remove.getparent().remove(p_to_remove)
            else: # Cannot find parent, maybe already removed or structure issue
                break

        # Ensure at least one paragraph exists
        if not cell.paragraphs:
             cell.add_paragraph()

        # Clear runs in the first paragraph
        first_para = cell.paragraphs[0]
        for run in first_para.runs:
             run.clear()
        first_para.text = "" # Ensure text is cleared too

        # Add the new text, handling newlines as paragraph breaks
        lines = str(text).split('\n')
        first_para.text = lines[0]
        for line in lines[1:]:
            # Add subsequent lines as new paragraphs within the same cell
            cell.add_paragraph(line)

    except Exception as e:
        print(f"  Warn: Exception setting cell text (text: '{str(text)[:50]}...'): {e}")
        # Fallback just in case
        cell.text = str(text) # Direct assignment as last resort

def parse_setup_equipment(setup_text: str) -> list[str]:
    """ Extracts equipment names from the custom_tc_setup string (improved). """
    equipment = []
    if not setup_text:
        return []

    lines = setup_text.splitlines()
    in_equipment_section = False
    equipment_line_regex = re.compile(r"^(?:\|\||>\s*|-\s*|\*\s*|)(?:.*?\|)?\s*([^|#\n]+?)\s*(?:\|.*)?$", re.IGNORECASE)
    excluded_keywords = {'refer to protocol', 'as needed', 'equipment/material', 'quantity', 'part/list number', 'notes:', ''}

    for line in lines:
        line = line.strip()
        if '#**Equipment & Materials**#' in line:
            in_equipment_section = True
            continue
        if in_equipment_section and line.startswith('#**'):
            in_equipment_section = False
            break
        if not in_equipment_section or not line:
            continue

        match = equipment_line_regex.match(line)
        if match:
            potential_name = match.group(1).strip().rstrip('.')
            cleaned_name = ' '.join(potential_name.split())
            if cleaned_name.lower() not in excluded_keywords and len(cleaned_name) > 1:
                 if not ('quantity' in cleaned_name.lower() and 'part' in cleaned_name.lower()):
                    equipment.append(cleaned_name)

    unique_equipment = list(dict.fromkeys(equipment))
    return unique_equipment

# Note: extract_requirement_text_from_expected is not used if AI handles summary text extraction
# Keep it if needed as a fallback or for direct parsing approaches.

def get_gemini_analysis(test_case_data: dict) -> dict | None:
    """ Sends test case data to Gemini for analysis and returns the structured JSON response. """
    if not gemini_model:
        print("Error: Gemini AI model not available.")
        return None
    if not test_case_data:
        print("Error: No test case data provided for Gemini analysis.")
        return None

    type_id_map = { 6: "Functional", 8: "Performance", 7: "Inspection" }
    test_type = type_id_map.get(test_case_data.get('type_id'), "Unknown")

    prompt_data = {
        "test_case_id": test_case_data.get("id"),
        "title": test_case_data.get("title"),
        "requirements": test_case_data.get("refs"),
        "test_type_id": test_case_data.get("type_id"),
        "identified_test_type": test_type,
        "setup_section": test_case_data.get("custom_tc_setup"),
        "steps_expected": test_case_data.get("custom_steps_separated", [])
    }

    # System prompt remains the same as before
    system_prompt = """You are an expert Hardware Verification Engineer specializing in medical devices, specifically ICU Medical infusion pumps and TestRail documentation. Your task is to analyze TestRail test case data and determine the structure needed for a corresponding '.docx' Record File used during test execution.

    Focus on identifying:
    1.  **Test Type:** Confirm the likely test type (Functional, Performance, Inspection) based on the provided 'identified_test_type'.
    2.  **Equipment List:** Extract the names of all necessary equipment from the 'setup_section'.
    3.  **General Information:** Extract the Test Case ID, Test ID (often in refs or custom_test_id), and formulate the main title line combining ID, title, and refs.
    4.  **Results Table Headers:** Analyze the 'steps_expected' list (both 'content' and 'expected' fields). Identify *all* specific data points, conditions, or checks that the test executor needs to explicitly record or verify. Create a list of concise, descriptive column headers for the main results table(s).
        *   Always include standard columns: "No.", "Infuser Serial Number", "Sign Name/Date". Include "Battery Serial Number" if batteries are mentioned. Include "Stopwatch Cal_ID / Calibration Due Date" if a stopwatch is mentioned.
        *   Look for phrases like "Record the results", "Verify that [condition]", "Record the value", "Check [parameter]", "Time recorded", etc.
        *   If multiple distinct sets of results are recorded in different steps (e.g., testing different rates or conditions separately), create a separate list of headers for *each* distinct results table needed. Title each table appropriately (e.g., "TABLE 1. REQUIREMENT [REQ_ID] CONDITION VERIFICATION", "TABLE 2. REQUIREMENT [REQ_ID] CONDITION VERIFICATION"). Infer REQ_ID association if possible.
        *   For 'Performance' tests, specifically look for numerical values to be recorded and potentially Upper/Lower Specification Limits (USL/LSL) mentioned in 'expected'. Headers might be like "Measured Value (unit)", "USL", "LSL".
        *   For 'Inspection' tests, headers might be simpler like "Document Reviewed", "Section/Clause", "Requirement Met (Pass/Fail)".
    5.  **Requirements Summary:** Analyze the 'expected' fields. For each requirement ID listed in 'requirements', extract the core verification condition text (the part after "...as a 'Pass' if..."). Create a list of objects, each containing the 'req_id' and its corresponding 'verification_text'. Determine the 'step' and 'sub_step' associated with the verification if mentioned in the 'expected' text.

    **Output Format:** Return the analysis strictly as a JSON object with the following structure:

    ```json
    {
      "test_type": "Functional | Performance | Inspection | Unknown",
      "general_info": {
        "title_line": "string", // e.g., "C12345: 'Title...' (REQ-001, REQ-002)."
        "test_case_id": "string", // e.g., "12345"
        "test_id": "string" // e.g., "REQ-001.1" or "N/A"
      },
      "equipment": [ // List of equipment names found in setup
        "string",
        ...
      ],
      "results_tables": [ // List of objects, one for each distinct results table needed
        {
          "title": "string", // e.g., "TABLE 1. REQUIREMENT CADD-SYSDI-1137 CONDITION VERIFICATION"
          "headers": ["string", ...], // List of column header strings
          "num_rows": number // Typically 30 for Performance, 29 for Functional, 1 or more for Inspection based on steps
        },
        ...
      ],
      "requirements_summary": [ // List of objects, one for each requirement verified
        {
          "step": "string", // Step number from 'expected' text (e.g., "1", "2") or "N/A"
          "sub_step": "string", // Sub-step number, if applicable (e.g., "14") or "N/A"
          "req_id": "string", // e.g., "CADD-SYSDI-1137"
          "verification_text": "string" // The core condition text
        },
        ...
      ]
    }
    ```

    Be precise and adhere strictly to the JSON format. If information isn't clearly present, use reasonable defaults like "N/A" or base `num_rows` on typical test types (Functional: 29, Performance: 30, Inspection: 1). Combine related verification checks into single headers where logical (e.g., "Verify 'NCB DIS' and battery charge displayed").
    """

    user_content = f"""Analyze the following TestRail test case data and generate the JSON structure for its corresponding record file:

    ```json
    {json.dumps(prompt_data, indent=2)}
    ```
    """

    print("--- Sending request to Gemini AI for analysis ---")

    try:
        # Use the generate_content method of the model object
        response = gemini_model.models.generate_content(
            model="gemini-2.5-pro-exp-03-25",
            config=types.GenerateContentConfig(system_instruction=system_prompt),
            contents=user_content
             # system_instruction=system_prompt # Pass system prompt if supported this way by your SDK version
        )

        # Access response text
        if hasattr(response, 'text'):
            response_text = response.text
        elif hasattr(response, 'parts') and response.parts:
             response_text = "".join(part.text for part in response.parts)
        elif hasattr(response, 'candidates') and response.candidates: # Handle candidate structure
              try:
                  response_text = response.candidates[0].content.parts[0].text
              except (IndexError, AttributeError) as e:
                   print(f"Error accessing text through candidates structure: {e}")
                   print(f"Raw Gemini Response: {response}")
                   return None
        else:
             print("Error: Could not extract text from Gemini response. Structure unexpected.")
             print(f"Raw Gemini Response: {response}")
             return None

        # Clean the response text
        cleaned_response_text = re.sub(r'^```json\s*', '', response_text.strip(), flags=re.IGNORECASE)
        cleaned_response_text = re.sub(r'\s*```$', '', cleaned_response_text)

        print("--- Received Gemini AI analysis ---")
        # print(cleaned_response_text) # Debug

        analysis_data = json.loads(cleaned_response_text)
        print("  Successfully parsed Gemini JSON response.")
        return analysis_data

    except json.JSONDecodeError as json_err:
        print(f"Error: Could not decode JSON response from Gemini AI: {json_err}")
        print(f"Raw Gemini Response Text:\n{response_text}")
        return None
    except Exception as e:
        print(f"Error during Gemini AI request or processing: {e}")
        if hasattr(response, 'prompt_feedback'):
             print(f"Prompt Feedback: {response.prompt_feedback}")
        # print(f"Raw Gemini Response Object: {response}") # More detailed debug
        return None

# --- Document Population Logic ---

class RecordFileGenerator:
    """ Populates a Word template based on TestRail data and AI analysis. """

    def __init__(self, template_path: str, testrail_data: dict, ai_analysis: dict):
        if not os.path.exists(template_path):
            raise FileNotFoundError(f"Template file not found: {template_path}")
        self.doc = docx.Document(template_path)
        self.testrail_data = testrail_data
        self.ai_analysis = ai_analysis
        print(f"RecordFileGenerator initialized with template: {template_path}")

        # --- Define FIXED Element Indices & Coordinates (Updated based on parser dump) ---
        self.TITLE_PARA_INDEX = 0
        self.GENERAL_INFO_TABLE_INDEX = 1 # The actual index in doc.tables

        # Coordinates within General Info Table (Table at index 1)
        # Adjusted to target the start of visually merged cells based on dump/PDF
        self.TC_ID_CELL = (11, 18) # Row 11, Col 18
        self.TEST_ID_CELL = (12, 18) # Row 12, Col 18

        # --- Constants for finding other tables DYNAMICALLY ---
        self.EQUIPMENT_TABLE_SEARCH_HEADER = "Equipment and Material Description"
        # Using more specific search to avoid accidental matches
        self.RESULTS_TABLE_1_SEARCH_TEXT = "TABLE 1. REQUIREMENTS" # Will likely find element 8 -> table 9
        self.REQUIREMENTS_SUMMARY_TABLE_SEARCH_TEXT = "TABLE 2. REQUIREMENTS" # Will likely find element 10 -> table 11

        # --- Constants for Equipment Table Structure ---
        self.EQUIPMENT_START_DATA_ROW = 1
        self.EQUIPMENT_END_DATA_ROW = 10 # Last *template* row for potential matching
        self.EQUIPMENT_NAME_COL = 0
        self.EQUIPMENT_QTY_COL = 1
        self.EQUIPMENT_PART_COL = 2
        self.EQUIPMENT_LOT_COL = 3
        self.EQUIPMENT_CAL_COL = 4

        # --- Constants for Results Table Structure (Specific to Functional Template's Table 1 (Element 9)) ---
        self.RESULTS_TABLE_DATA_START_ROW = 3 # Data starts at row index 3 (Sample '1')

        # --- Constants for Requirements Summary Table Structure ---
        self.REQ_SUMMARY_START_ROW = 1
        self.REQ_SUMMARY_STEP_COL = -1 # No dedicated column in this template
        self.REQ_SUMMARY_SUBSTEP_COL = -1 # No dedicated column in this template
        self.REQ_SUMMARY_REQID_COL = 0
        self.REQ_SUMMARY_VERIF_COL = 1
        self.REQ_SUMMARY_RESULT_COL = 2
        self.REQ_SUMMARY_SIGN_COL = 3

    def _find_table_by_preceding_paragraph_text(self, search_text_regex: str) -> Table | None:
        """ Finds a table immediately following a paragraph matching the regex pattern. """
        try:
            pattern = re.compile(search_text_regex, re.IGNORECASE)
            # Iterate through the raw body elements to maintain order
            body_elements = list(self.doc.element.body)
            for i, element in enumerate(body_elements):
                 if isinstance(element, CT_P):
                     # Map XML element back to Paragraph object to get text
                     para = None
                     for p in self.doc.paragraphs:
                         if p._element == element:
                             para = p
                             break
                     if para and pattern.search(para.text):
                         # Check if the *next* element is a table
                         if i + 1 < len(body_elements) and isinstance(body_elements[i+1], CT_Tbl):
                             # Map table element back to Table object
                             table_element = body_elements[i+1]
                             for t in self.doc.tables:
                                 if t._element == table_element:
                                     print(f"  -> Found table using preceding paragraph regex: '{search_text_regex}'")
                                     return t
            print(f"  -> Warning: Could not find table using preceding paragraph regex: '{search_text_regex}'")
            return None
        except Exception as e:
            print(f"  -> Error finding table by preceding paragraph regex '{search_text_regex}': {e}")
            return None

    def _find_table_by_header_content(self, header_text_keyword: str) -> Table | None:
        """ Finds a table by checking the content of its first row cells. """
        try:
            for table in self.doc.tables:
                if len(table.rows) > 0:
                    header_row = table.rows[0]
                    for cell in header_row.cells:
                        # Handle potential None text or complex cell content
                        cell_text_parts = []
                        for p in cell.paragraphs:
                            cell_text_parts.append(p.text)
                        cell_text = "".join(cell_text_parts).strip()

                        if header_text_keyword.lower() in cell_text.lower():
                            print(f"  -> Found table using header content keyword: '{header_text_keyword}'")
                            return table
            print(f"  -> Warning: Could not find table using header content keyword: '{header_text_keyword}'")
            return None
        except Exception as e:
            print(f"  -> Error finding table by header keyword '{header_text_keyword}': {e}")
            return None

    def populate_title(self):
        """ Populates the main title paragraph (Element 0). """
        print("--- Populating Title ---")
        try:
            title_text = self.ai_analysis.get('general_info', {}).get('title_line', "[Title Missing]")
            if self.TITLE_PARA_INDEX < len(self.doc.paragraphs):
                para0 = self.doc.paragraphs[self.TITLE_PARA_INDEX]
                # Clear existing runs before setting new text
                for run in para0.runs:
                   run.clear()
                para0.text = title_text # Set text directly
                print(f"  Populated title: '{title_text[:100]}...'")
            else:
                print(f"  Error: Paragraph index {self.TITLE_PARA_INDEX} out of range.")
        except Exception as e:
            print(f"  Error populating title: {e}")

    def populate_general_info(self):
        """ Populates fields in the General Information table (Table at index 1). """
        print("--- Populating General Info Table ---")
        try:
            # Access table by its known index (less prone to failure than dynamic find here)
            if self.GENERAL_INFO_TABLE_INDEX < len(self.doc.tables):
                 table0 = self.doc.tables[self.GENERAL_INFO_TABLE_INDEX]
                 gen_info = self.ai_analysis.get('general_info', {})
                 tc_id = gen_info.get('test_case_id', 'N/A')
                 test_id = gen_info.get('test_id', 'N/A')

                 # Use updated cell coordinates
                 try:
                     set_cell_text(table0.cell(self.TC_ID_CELL[0], self.TC_ID_CELL[1]), tc_id)
                     print(f"  Populated Test Case ID at {self.TC_ID_CELL}: {tc_id}")
                 except IndexError:
                      print(f"  Error: Cell index {self.TC_ID_CELL} out of range for Test Case ID in Table {self.GENERAL_INFO_TABLE_INDEX}.")
                 try:
                      set_cell_text(table0.cell(self.TEST_ID_CELL[0], self.TEST_ID_CELL[1]), test_id)
                      print(f"  Populated Test ID at {self.TEST_ID_CELL}: {test_id}")
                 except IndexError:
                      print(f"  Error: Cell index {self.TEST_ID_CELL} out of range for Test ID in Table {self.GENERAL_INFO_TABLE_INDEX}.")
            else:
                 print(f"  Error: Table index {self.GENERAL_INFO_TABLE_INDEX} out of range for General Info.")
        except Exception as e:
            print(f"  Error populating General Info table: {e}")

    def populate_equipment(self):
        """ Populates the Equipment table (Found dynamically). """
        print("--- Populating Equipment Table ---")
        try:
            table1 = self._find_table_by_header_content(self.EQUIPMENT_TABLE_SEARCH_HEADER)
            if not table1:
                print("  Error: Equipment table not found using header search.")
                return

            equipment_list = self.ai_analysis.get('equipment', [])
            if not equipment_list:
                print("  Warning: Equipment list not found in AI analysis, attempting to parse from setup text.")
                equipment_list = parse_setup_equipment(self.testrail_data.get('custom_tc_setup', ''))

            print(f"  Equipment to populate: {equipment_list}")
            found_items_in_table = []
            remaining_equipment = list(equipment_list) # Copy list to modify

            for r in range(self.EQUIPMENT_START_DATA_ROW, min(self.EQUIPMENT_END_DATA_ROW + 1, len(table1.rows))):
                try:
                    cell_obj = table1.cell(r, self.EQUIPMENT_NAME_COL)
                    cell_text = "".join([p.text for p in cell_obj.paragraphs]).strip()

                    if not cell_text: continue

                    matched_eq = None
                    for eq_name in remaining_equipment:
                        # Clean both strings thoroughly for comparison
                        clean_cell_text = re.sub(r'\W+', '', cell_text).lower() # Remove non-alphanumeric
                        clean_eq_name = re.sub(r'\W+', '', eq_name).lower() # Remove non-alphanumeric

                        # Use startswith after cleaning for potentially more reliable matching
                        if clean_cell_text.startswith(clean_eq_name):
                            matched_eq = eq_name
                            # print(f"DEBUG: Match Found! EQ: '{eq_name}' vs CELL: '{cell_text}' (Cleaned: '{clean_eq_name}' vs '{clean_cell_text}')") # Debug Print
                            break # Found match for this row

                    if matched_eq:
                        try:
                            set_cell_text(table1.cell(r, self.EQUIPMENT_QTY_COL), "[Enter Qty]")
                            set_cell_text(table1.cell(r, self.EQUIPMENT_PART_COL), "[Enter Part#]")
                            set_cell_text(table1.cell(r, self.EQUIPMENT_LOT_COL), "[Enter Lot/SN]")
                            set_cell_text(table1.cell(r, self.EQUIPMENT_CAL_COL), "[Enter Cal/SW]")
                            found_items_in_table.append(matched_eq)
                            remaining_equipment.remove(matched_eq) # Remove from list once matched
                            print(f"  Populated placeholders in row {r} for equipment: '{cell_text}' (matched '{matched_eq}')")
                        except IndexError:
                             print(f"  Error: Column index out of range populating placeholders in row {r} for '{matched_eq}'.")
                except IndexError:
                    print(f"  Warning: Cell index ({r}, {self.EQUIPMENT_NAME_COL}) out of range in Equipment table.")
                    continue

            print(f"  Finished populating equipment placeholders. Items matched: {len(found_items_in_table)}/{len(equipment_list)}")
            if remaining_equipment:
                print(f"  Warning: Equipment items not matched to template rows: {remaining_equipment}")

        except Exception as e:
            print(f"  Error populating Equipment table: {e}")
            import traceback
            traceback.print_exc()

    def populate_results_tables(self):
        """ Populates the main results table(s) based on AI analysis (Found dynamically). """
        print("--- Populating Results Tables ---")
        results_tables_data = self.ai_analysis.get('results_tables', [])
        if not results_tables_data:
            print("  No results table structures found in AI analysis.")
            return

        # Assuming only one results table based on AI output for TC 983037
        # If AI starts generating multiple, this logic needs a loop and better finding
        if len(results_tables_data) > 1:
            print("  Warning: AI generated multiple results tables, but script currently handles only the first.")

        table_data = results_tables_data[0]
        table_title = table_data.get("title", "Results Table 1")
        headers = table_data.get("headers", [])
        num_rows = table_data.get("num_rows", 1)

        if not headers:
            print(f"    Warning: No headers found for table '{table_title}'. Skipping.")
            return

        # Find the table using the preceding paragraph text
        target_table = self._find_table_by_preceding_paragraph_text(self.RESULTS_TABLE_1_SEARCH_TEXT)

        if not target_table:
            print(f"  Error: Could not find Results Table 1 using search text '{self.RESULTS_TABLE_1_SEARCH_TEXT}'.")
            return

        print(f"  Populating Results Table 1 (Found Dynamically). Title: '{table_title}'")

        try:
            # --- Set Table Title (in preceding paragraph) ---
            table_element = target_table._element
            parent = table_element.getparent()
            title_para = None
            if parent is not None:
                try:
                    table_xml_index = parent.index(table_element)
                    if table_xml_index > 0:
                        preceding_element = parent[table_xml_index - 1]
                        if isinstance(preceding_element, CT_P):
                            # Map back to Paragraph object
                            for p in self.doc.paragraphs:
                                if p._element == preceding_element:
                                    title_para = p
                                    break
                except ValueError: pass # Ignore if index not found

            if title_para:
                print(f"    Setting title in preceding paragraph: '{table_title}'")
                for run in title_para.runs: run.clear()
                title_para.text = table_title
                if title_para.runs: title_para.runs[0].font.bold = True
            else:
                print(f"    Warning: No paragraph found directly preceding the results table. Cannot set title paragraph.")

            # --- Populate Headers (Handles multi-row headers in template by overwriting Row 0 only) ---
            if len(target_table.rows) > 0:
                 header_row = target_table.rows[0] # Target the first row for AI headers
                 if len(header_row.cells) >= len(headers):
                      print(f"    Setting headers in row 0: {headers}")
                      for c_idx, header_text in enumerate(headers):
                           set_cell_text(header_row.cells[c_idx], header_text)
                      # Clear any extra template cells in this specific header row
                      for c_idx in range(len(headers), len(header_row.cells)):
                           set_cell_text(header_row.cells[c_idx], "")
                 else:
                      print(f"    Error: Results table has fewer columns ({len(header_row.cells)}) than headers ({len(headers)}). Cannot set headers.")
                      return # Stop processing this table
            else:
                 print(f"    Error: Results table has no rows. Cannot set headers.")
                 return # Stop processing this table

            # --- Manage Data Rows (Corrected logic for start row and removal) ---
            # Functional template has 3 header rows, data starts index 3.
            num_template_header_rows = self.RESULTS_TABLE_DATA_START_ROW # 3
            template_data_rows = len(target_table.rows) - num_template_header_rows
            rows_to_add = num_rows - template_data_rows # num_rows is *data* rows needed
            print(f"    Target data rows: {num_rows}. Template has: {template_data_rows}. Need to add: {rows_to_add}")

            if rows_to_add > 0:
                 print(f"    Adding {rows_to_add} data rows...")
                 for _ in range(rows_to_add):
                     target_table.add_row()

            # Populate data rows (from RESULTS_TABLE_DATA_START_ROW up to needed rows)
            num_cols = len(headers)
            for r_idx_offset in range(num_rows): # Iterate 0 to num_rows-1
                actual_row_index = self.RESULTS_TABLE_DATA_START_ROW + r_idx_offset
                if actual_row_index >= len(target_table.rows):
                    print(f"    Error: Row index {actual_row_index} is out of bounds after adding rows.")
                    break
                row = target_table.rows[actual_row_index]

                if len(row.cells) < num_cols:
                    print(f"    Warning: Row {actual_row_index} has fewer cells ({len(row.cells)}) than headers ({num_cols}). Skipping pop.")
                    continue

                for c_idx in range(num_cols):
                    header_lower = headers[c_idx].lower()
                    cell = row.cells[c_idx]
                    # Check if it's the 'No.' column (usually first)
                    if c_idx == 0 and ("no." in header_lower or "sample" in header_lower):
                        set_cell_text(cell, str(r_idx_offset + 1)) # Populate 1-based sample number
                    elif "pass" in header_lower and "fail" in header_lower:
                        set_cell_text(cell, "Pass ☐ / Fail ☐")
                    elif "sign" in header_lower or "/date" in header_lower:
                        set_cell_text(cell, "")
                    else:
                        set_cell_text(cell, "") # Clear other data cells

            # Remove extra template data rows if num_rows < template_data_rows
            if rows_to_add < 0:
                 rows_to_remove_count = abs(rows_to_add)
                 print(f"    Removing {rows_to_remove_count} extra template data rows...")
                 for _ in range(rows_to_remove_count):
                      # Always remove the last row if it's beyond the needed data rows + header rows
                      if len(target_table.rows) > (num_rows + num_template_header_rows):
                           row_element_to_remove = target_table.rows[-1]._element
                           if row_element_to_remove.getparent() is not None:
                               row_element_to_remove.getparent().remove(row_element_to_remove)
                           else: break # Stop if parent gone
                      else: break # Stop if table size matches target

        except Exception as e:
            print(f"    Error processing Results Table 1: {e}")
            import traceback
            traceback.print_exc()

    def populate_requirements_summary(self):
        """ Populates the Requirements Summary table (Found dynamically). """
        print("--- Populating Requirements Summary Table ---")
        summary_data = self.ai_analysis.get('requirements_summary', [])
        if not summary_data:
            print("  No requirements summary data found in AI analysis.")
            return

        target_table = self._find_table_by_preceding_paragraph_text(self.REQUIREMENTS_SUMMARY_TABLE_SEARCH_TEXT)
        if not target_table:
            print(f"  Error: Could not find Requirements Summary table using search text '{self.REQUIREMENTS_SUMMARY_TABLE_SEARCH_TEXT}'.")
            return

        num_reqs_to_populate = len(summary_data)
        template_rows_available = len(target_table.rows) - self.REQ_SUMMARY_START_ROW
        print(f"  Populating {num_reqs_to_populate} requirements. Template has {template_rows_available} data rows available.")

        for i, req_info in enumerate(summary_data):
            row_index = self.REQ_SUMMARY_START_ROW + i
            row = None # Initialize row

            if row_index < len(target_table.rows):
                row = target_table.rows[row_index]
            else:
                # Add row if needed
                print(f"    Warning: Not enough rows in summary table template for requirement {i+1}. Adding row.")
                try:
                    row = target_table.add_row()
                except Exception as add_e:
                    print(f"    Error adding row to summary table: {add_e}")
                    break # Stop processing summary if adding rows fails

            if row: # Proceed only if row exists or was added successfully
                print(f"    Populating row {row_index} for Req: {req_info.get('req_id', 'N/A')}")
                try:
                    # Get data from AI
                    step_text = req_info.get('step', '') if req_info.get('step', 'N/A') != 'N/A' else ''
                    sub_step_text = req_info.get('sub_step', '') if req_info.get('sub_step', 'N/A') != 'N/A' else ''
                    req_id_text = req_info.get('req_id', '')
                    verif_text = req_info.get('verification_text', '')

                    # Prepend step/substep to verification text if those columns are missing
                    prefix = ""
                    if step_text and self.REQ_SUMMARY_STEP_COL < 0:
                        prefix += f"Step {step_text}"
                    if sub_step_text and self.REQ_SUMMARY_SUBSTEP_COL < 0:
                         prefix += f", Sub-step {sub_step_text}" if prefix else f"Sub-step {sub_step_text}"
                    if prefix:
                         verif_text = f"({prefix}): {verif_text}"

                    # Populate cells using column indices, checking if index is valid (>= 0)
                    if self.REQ_SUMMARY_STEP_COL >= 0: set_cell_text(row.cells[self.REQ_SUMMARY_STEP_COL], step_text)
                    if self.REQ_SUMMARY_SUBSTEP_COL >= 0: set_cell_text(row.cells[self.REQ_SUMMARY_SUBSTEP_COL], sub_step_text)
                    if self.REQ_SUMMARY_REQID_COL >= 0: set_cell_text(row.cells[self.REQ_SUMMARY_REQID_COL], req_id_text)
                    if self.REQ_SUMMARY_VERIF_COL >= 0: set_cell_text(row.cells[self.REQ_SUMMARY_VERIF_COL], verif_text)
                    if self.REQ_SUMMARY_RESULT_COL >= 0: set_cell_text(row.cells[self.REQ_SUMMARY_RESULT_COL], "Pass ☐ / Fail ☐")
                    if self.REQ_SUMMARY_SIGN_COL >= 0: set_cell_text(row.cells[self.REQ_SUMMARY_SIGN_COL], "")

                except IndexError:
                    print(f"    Error: Column index out of range while populating summary row {row_index}.")
                except Exception as cell_e:
                    print(f"    Error populating cells in summary row {row_index}: {cell_e}")


        # Clear Extra Template Rows
        if num_reqs_to_populate < template_rows_available:
            rows_to_clear_count = template_rows_available - num_reqs_to_populate
            print(f"  Clearing {rows_to_clear_count} extra rows in summary table template.")
            start_clear_row = self.REQ_SUMMARY_START_ROW + num_reqs_to_populate
            for r_idx in range(start_clear_row, len(target_table.rows)):
                 row = target_table.rows[r_idx]
                 for cell in row.cells:
                     set_cell_text(cell, "")

    def save(self, output_path: str):
        """ Saves the populated document. """
        print(f"--- Saving populated document to: {output_path} ---")
        try:
            output_dir = os.path.dirname(output_path)
            if output_dir: os.makedirs(output_dir, exist_ok=True)
            self.doc.save(output_path)
            print("  Save successful.")
        except PermissionError:
             print(f"  Error: Permission denied saving to '{output_path}'. Check if the file is open or permissions are correct.")
        except Exception as e:
            print(f"  Error saving document: {e}")
            import traceback
            traceback.print_exc()

# --- Main Execution ---
def main():
    # --- Configuration ---
    TESTRAIL_URL = "https://icumed.testrail.io"
    # IMPORTANT: Replace with your actual credentials (or use environment variables)
    TESTRAIL_EMAIL = os.environ.get("TESTRAIL_EMAIL", "jairo.rodriguez@icumed.com")
    TESTRAIL_API_KEY = os.environ.get("TESTRAIL_API_KEY", "uttDDNwG5EMQeiW71KaN-3l9Ndf4mBqWT0xV7TF5F")

    test_case_id_to_process = 983037 # Functional Example 2

    # Template file paths (Adjust as necessary)
    template_dir = r'C:\Users\JR00093781\OneDrive - ICU Medical Inc\Documents\TestRecordGenerator\Templates'
    template_file_functional = os.path.join(template_dir, 'Functional_Template_Base.docx')
    template_file_performance = os.path.join(template_dir, 'Performance_Template_Base.docx') # Assuming it exists
    template_file_inspection = os.path.join(template_dir, 'Inspection_Template_Base.docx') # Assuming it exists

    # Select template based on test type (Manual selection for now)
    # TODO: Add logic to select template based on ai_analysis['test_type'] later
    template_file_to_use = template_file_functional

    output_dir = r"C:\Users\JR00093781\OneDrive - ICU Medical Inc\Documents\TestRecordGenerator\v2\Generated Record Files"
    safe_tc_id = str(test_case_id_to_process)
    output_file_name = os.path.join(output_dir, f"Record File TC {safe_tc_id} - Generated.docx")

    # --- Initial Checks ---
    if not gemini_model:
        print("Error: Gemini AI Client failed to initialize. Exiting.")
        sys.exit(1)
    if "YOUR_TESTRAIL_KEY" in TESTRAIL_API_KEY or "your_email@example.com" in TESTRAIL_EMAIL:
         print("Error: Please configure TestRail credentials.")
         # sys.exit(1) # Comment out exit for testing if using hardcoded values
    if not os.path.exists(template_file_to_use):
         print(f"Error: Selected template file not found: {template_file_to_use}")
         sys.exit(1)

    # --- Processing ---
    print(f"--- Starting Record File Generation for Test Case ID: {test_case_id_to_process} ---")

    tr_client = TestRailAPIClient(TESTRAIL_URL, TESTRAIL_EMAIL, TESTRAIL_API_KEY)
    endpoint = f"get_case/{test_case_id_to_process}"
    testrail_data = tr_client.send_get_request(endpoint)

    if not testrail_data:
        print(f"Error: Failed to fetch TestRail data for TC {test_case_id_to_process}. Exiting.")
        sys.exit(1)

    ai_analysis = get_gemini_analysis(testrail_data)

    if not ai_analysis:
        print("Error: Failed to get analysis from Gemini AI. Exiting.")
        sys.exit(1)

    print("\n--- Received AI Analysis (Structure) ---")
    print(json.dumps(ai_analysis, indent=2))

    try:
        generator = RecordFileGenerator(template_file_to_use, testrail_data, ai_analysis)
        generator.populate_title()
        generator.populate_general_info()
        generator.populate_equipment()
        generator.populate_results_tables()
        generator.populate_requirements_summary()
        generator.save(output_file_name)
        print(f"\n--- Record File Generation Completed for TC {test_case_id_to_process} ---")
        print(f"Output saved to: {output_file_name}")
    except FileNotFoundError as fnf_err:
         print(f"\nError initializing generator: {fnf_err}")
    except Exception as gen_err:
         print(f"\nAn critical error occurred during record file generation: {gen_err}")
         import traceback
         traceback.print_exc()

if __name__ == "__main__":
    main()