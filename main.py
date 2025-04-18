import requests
import json
import time
import os
import re
import sys
import docx
from docx.document import Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table, _Row
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
    gemini_model = None
except Exception as e:
    print(f"Error initializing Gemini AI client: {e}")
    gemini_model = None

# --- Force UTF-8 ---
try:
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')
except AttributeError:
    print("Warning: Could not reconfigure stdout/stderr encoding (not supported on this platform/version?).")


# --- TestRail API Client (from testCaseGetter.py) ---
class TestRailAPIClient:
    # ... (Paste the __init__ and _make_request methods here) ...
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
            # print(f"DEBUG: Requesting {method} {full_url} with params {kwargs.get('params')}") # Debug
            response = self.session.request(method, full_url, timeout=timeout, **kwargs)
            response.raise_for_status() # Raise HTTPError for bad responses (4xx or 5xx)
            return response # Return the full response object
        except requests.exceptions.HTTPError as http_err:
            status_code = http_err.response.status_code
            response_text = http_err.response.text[:200].encode(sys.stdout.encoding, errors='replace').decode(sys.stdout.encoding) # Handle potential encoding issues in error text
            if status_code == 400:
                 print(f"Info: Received 400 Bad Request for {method} {full_url}. Check endpoint/parameters. Response: {response_text}...")
            elif status_code == 401:
                 print(f"Error: Received 401 Unauthorized for {method} {full_url}. Check email/API key.")
            elif status_code == 403:
                 print(f"Error: Received 403 Forbidden for {method} {full_url}. Check permissions or API key.")
            elif status_code == 404:
                 print(f"Info: Received 404 Not Found for {method} {full_url}. Resource likely doesn't exist.")
            elif status_code == 429:
                 print(f"Error: Received 429 Too Many Requests (Rate Limit) for {method} {full_url}. Try increasing delay between calls if applicable.")
            else:
                 print(f"HTTP error during request to {method} {full_url}: {http_err} - Status: {status_code}. Response: {response_text}...")
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
        # Ensure the base URL structure is correct for TestRail API v2
        full_url = f"{self.base_url}/index.php?/api/v2/{clean_endpoint}"
        print(f"  Target: {full_url}")
        response = self._make_request('GET', full_url)

        if response:
            try:
                 if response.status_code == 204 or not response.content:
                      print(f"  Success: Received status {response.status_code} with no content.")
                      return None
                 # Attempt to decode JSON, handling potential encoding issues
                 try:
                     response_data = response.json()
                 except json.JSONDecodeError:
                     # Try decoding explicitly with utf-8, common for web APIs
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
    # Add send_paginated_get_request if needed later, but get_case is single

# --- Helper Functions (from initialFiller.py, potentially modified) ---

def set_cell_text(cell: docx.table._Cell, text: str):
    """ Safely sets the text in a cell, replacing existing content. """
    if not isinstance(cell, docx.table._Cell):
        print(f"  Error: Invalid object passed to set_cell_text. Expected a Cell, got {type(cell)}.")
        return
    try:
        # Clear existing content more robustly
        cell.text = "" # Clear simple text
        for p in cell.paragraphs:
             # Remove all runs within the paragraph
             for r in p.runs:
                 r.clear()
             # If paragraph is now empty, remove it (except the first one)
             if not p.text and not p.runs and cell.paragraphs.index(p) > 0 :
                  p_element = p._element
                  if p_element.getparent() is not None:
                       p_element.getparent().remove(p_element)

        # Ensure there's at least one paragraph to add text to
        if not cell.paragraphs:
             cell.add_paragraph()

        # Add the new text to the first paragraph
        # Replace newline characters with paragraph breaks for multi-line text
        lines = str(text).split('\n')
        first_paragraph = cell.paragraphs[0]
        first_paragraph.text = lines[0]
        # Add subsequent lines as new paragraphs
        for line in lines[1:]:
            cell.add_paragraph(line)

    except Exception as e:
        print(f"  Warn: Exception setting cell text (text: '{str(text)[:50]}...'): {e}")
        # Fallback just in case
        if cell.paragraphs:
             cell.paragraphs[0].text = str(text)
        else:
             cell.text = str(text)

def parse_setup_equipment(setup_text: str) -> list[str]:
    """ Extracts equipment names from the custom_tc_setup string (improved). """
    equipment = []
    if not setup_text:
        return []

    lines = setup_text.splitlines()
    in_equipment_section = False
    # Regex to find potential equipment lines (more flexible than simple split)
    # Looks for markdown-like table rows or simple list items after the header
    # Handles variations like '||', '> ', '-', '*' at the start
    # Captures the *last* meaningful part of the line as potential equipment name
    # Excludes common headers/placeholders
    equipment_line_regex = re.compile(r"^(?:\|\||>\s*|-\s*|\*\s*|)(?:.*?\|)?\s*([^|#\n]+?)\s*(?:\|.*)?$", re.IGNORECASE)
    excluded_keywords = {'refer to protocol', 'as needed', 'equipment/material', 'quantity', 'part/list number', 'notes:', ''}

    for line in lines:
        line = line.strip()
        if '#**Equipment & Materials**#' in line:
            in_equipment_section = True
            continue
        if in_equipment_section and line.startswith('#**'): # Reached next section
            in_equipment_section = False
            break
        if not in_equipment_section or not line:
            continue

        match = equipment_line_regex.match(line)
        if match:
            potential_name = match.group(1).strip().rstrip('.')
            # Further clean and check against exclusion list
            cleaned_name = ' '.join(potential_name.split()) # Normalize whitespace
            if cleaned_name.lower() not in excluded_keywords and len(cleaned_name) > 1: # Avoid single characters
                 # Heuristic: Avoid lines that look like headers based on typical content
                 if not ('quantity' in cleaned_name.lower() and 'part' in cleaned_name.lower()):
                    equipment.append(cleaned_name)

    # Remove duplicates preserving order (Python 3.7+)
    unique_equipment = list(dict.fromkeys(equipment))
    # print(f"Debug: Unique equipment extracted: {unique_equipment}")
    return unique_equipment

def extract_requirement_text_from_expected(expected_text: str, req_id: str) -> str:
    """ Extracts the verification text for a specific requirement ID from the 'expected' field. """
    if not expected_text or not req_id:
        return f"[Verification text for {req_id} not found - missing input]"

    # Normalize req_id for comparison
    normalized_req_id = req_id.strip().upper()
    lines = expected_text.splitlines()

    for line in lines:
        line = line.strip()
        # Look for lines starting with a number/bullet and containing the req_id (case-insensitive)
        match = re.match(r"^\s*(?:\d+\.|[-*+]\s+)\s*Record\s+(.+?)\s+as\s+a\s+'Pass'\s+if\s+(.*)", line, re.IGNORECASE)
        if match:
            found_req = match.group(1).strip().upper()
            verification_condition = match.group(2).strip()
            # Check if the found req_id matches the one we're looking for
            if normalized_req_id in found_req.replace('-', ''): # Allow for CADD-SYSDI-XXX or CADDSYSDIXXX
                print(f"  -> Found verification text for {req_id}: '{verification_condition[:60]}...'")
                return verification_condition

    print(f"  -> Warning: Verification text pattern not found for Req {req_id} in expected results.")
    return f"[Verification text for {req_id} not clearly identified in expected results]"

def get_gemini_analysis(test_case_data: dict) -> dict | None:
    """ Sends test case data to Gemini for analysis and returns the structured JSON response. """
    if not gemini_model:
        print("Error: Gemini AI model not available.")
        return None
    if not test_case_data:
        print("Error: No test case data provided for Gemini analysis.")
        return None

    # --- Determine Test Type (Mapping Placeholder - Adjust based on your actual IDs) ---
    type_id_map = {
        6: "Functional",
        8: "Performance",
        7: "Inspection",
        # Add other mappings as needed
    }
    test_type = type_id_map.get(test_case_data.get('type_id'), "Unknown")

    # --- Select relevant data for the prompt ---
    prompt_data = {
        "test_case_id": test_case_data.get("id"),
        "title": test_case_data.get("title"),
        "requirements": test_case_data.get("refs"),
        "test_type_id": test_case_data.get("type_id"),
        "identified_test_type": test_type,
        "setup_section": test_case_data.get("custom_tc_setup"),
        "steps_expected": test_case_data.get("custom_steps_separated", [])
        # Add other fields if Gemini needs them (e.g., custom_test_id)
    }

    # --- Craft the System Prompt ---
    # This is critical and may need refinement based on Gemini's performance
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

    # --- Prepare the user content ---
    user_content = f"""Analyze the following TestRail test case data and generate the JSON structure for its corresponding record file:

    ```json
    {json.dumps(prompt_data, indent=2)}
    ```
    """

    print("--- Sending request to Gemini AI for analysis ---")
    # print(f"System Prompt: {system_prompt}") # Debug
    # print(f"User Content (first 500 chars): {user_content[:500]}...") # Debug

    try:
        # Using the newer client API structure if applicable
        response = gemini_model.models.generate_content(
            model="gemini-2.5-pro-exp-03-25",
            config=types.GenerateContentConfig(system_instruction=system_prompt),
            contents=user_content
             # system_instruction=system_prompt # Pass system prompt if supported this way by your SDK version
        )


        # --- Process the response ---
        # Accessing response text might differ slightly based on SDK version
        # Check response object structure via print(response) or dir(response) if needed
        if hasattr(response, 'text'):
            response_text = response.text
        elif hasattr(response, 'parts') and response.parts:
             # Handle potential streaming or multi-part responses if applicable
             response_text = "".join(part.text for part in response.parts)
        # elif hasattr(response, 'candidates') and response.candidates:
             # Handle response structure with candidates
             # response_text = response.candidates[0].content.parts[0].text
        else:
             print("Error: Could not extract text from Gemini response. Structure might be unexpected.")
             print(f"Raw Gemini Response: {response}")
             return None


        # Clean the response text - Gemini might wrap it in ```json ... ```
        cleaned_response_text = re.sub(r'^```json\s*', '', response_text.strip(), flags=re.IGNORECASE)
        cleaned_response_text = re.sub(r'\s*```$', '', cleaned_response_text)

        print("--- Received Gemini AI analysis ---")
        # print(cleaned_response_text) # Debug: Print the cleaned JSON string

        # Parse the JSON response
        analysis_data = json.loads(cleaned_response_text)
        print("  Successfully parsed Gemini JSON response.")
        return analysis_data

    except json.JSONDecodeError as json_err:
        print(f"Error: Could not decode JSON response from Gemini AI: {json_err}")
        print(f"Raw Gemini Response Text:\n{response_text}")
        return None
    except Exception as e:
        print(f"Error during Gemini AI request or processing: {e}")
        # Print detailed error info if available
        if hasattr(response, 'prompt_feedback'):
             print(f"Prompt Feedback: {response.prompt_feedback}")
        # Log the raw response for debugging
        # print(f"Raw Gemini Response Object: {response}")
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

        # --- Define Element Indices (CRITICAL - Update based on initialParser.py output for your template) ---
        # Indices for Functional_Template_Base.docx based on parser dump
        self.TITLE_PARA_INDEX = 0
        self.GENERAL_INFO_TABLE_INDEX = 1 # Was 0
        self.EQUIPMENT_TABLE_INDEX = 5 # Was 1
        self.FIRST_RESULTS_TABLE_INDEX = 9 # Was 2 (This is the main 32-row table)
        # This template seems to only have ONE main results table based on the dump structure
        # The next table (index 11) is the Requirements Summary
        # Set unused indices to something invalid like -100 if needed, or adjust logic
        self.SECOND_RESULTS_TABLE_INDEX = -100 # No second results table apparent
        self.THIRD_RESULTS_TABLE_INDEX = -100 # No third results table apparent
        # The table at index 11 is the Requirement Summary
        self.REQUIREMENTS_SUMMARY_TABLE_INDEX = 11 # Was -1

        # Example Cell Coordinates within Tables (Update based on template)
        # General Info Table (Table 1)
        self.TC_ID_CELL = (11, 23) # Row 11, Col 23 (Placeholder 'CXXXXX') - Was (7, 3)
        self.TEST_ID_CELL = (12, 23) # Row 12, Col 23 (Empty placeholder) - Was (8, 3)

        # Equipment Table (Table 5) - Define rows/cols for data entry placeholders
        self.EQUIPMENT_START_DATA_ROW = 1 # Data starts row 1 ('CADD Solis Infuser') - Was 1 (Correct)
        self.EQUIPMENT_END_DATA_ROW = 10 # Last empty template row is 10 - Was 9
        self.EQUIPMENT_NAME_COL = 0    # Column 0 for description - Was 0 (Correct)
        self.EQUIPMENT_QTY_COL = 1     # Column 1 for Quantity - Was 1 # Update from 4
        self.EQUIPMENT_PART_COL = 2    # Column 2 for Part Number - Was 2 # Update from 7
        self.EQUIPMENT_LOT_COL = 3     # Column 3 for Lot/SN/Cal_ID - Was 3 # Update from 14
        self.EQUIPMENT_CAL_COL = 4     # Column 4 for Cal Due/SW Ver - Was 4 # Update from 23

        # Requirements Summary Table (Table 11)
        self.REQ_SUMMARY_START_ROW = 1 # Data starts row 1 - Was 1 (Correct)
        # This template's summary table structure (Element 11): Requirement | Verification Steps | Result | Signature/Date
        # Map AI output to these columns. Step/Substep aren't separate columns here.
        self.REQ_SUMMARY_STEP_COL = -1 # Indicate Step is not a direct column - Was 0
        self.REQ_SUMMARY_SUBSTEP_COL = -1 # Indicate Substep is not a direct column - Was 1
        self.REQ_SUMMARY_REQID_COL = 0 # Column 0 for Requirement ID - Was 2
        self.REQ_SUMMARY_VERIF_COL = 1 # Column 1 for Verification Steps - Was 3
        self.REQ_SUMMARY_RESULT_COL = 2 # Column 2 for Result (Pass/Fail) - Was 4
        self.REQ_SUMMARY_SIGN_COL = 3 # Column 3 for Signature/Date - Was 5

    def _find_table_by_preceding_title(self, title_keyword: str) -> Table | None:
        """ Attempts to find a table immediately following a paragraph containing the title keyword. """
        # This is more robust than index but assumes structure Para -> Table
        try:
            for i, element in enumerate(self.doc.element.body):
                 if isinstance(element, CT_P):
                     # Need to get the Paragraph object to read text
                     para = None
                     for p in self.doc.paragraphs:
                         if p._element == element:
                             para = p
                             break
                     if para and title_keyword.upper() in para.text.upper():
                         # Check if the *next* element is a table
                         if i + 1 < len(self.doc.element.body) and isinstance(self.doc.element.body[i+1], CT_Tbl):
                             # Find the corresponding Table object
                             table_element = self.doc.element.body[i+1]
                             for t in self.doc.tables:
                                 if t._element == table_element:
                                     print(f"  -> Found table for '{title_keyword}' using preceding title method.")
                                     return t
            print(f"  -> Warning: Could not find table using preceding title: '{title_keyword}'")
            return None
        except Exception as e:
            print(f"  -> Error finding table by title '{title_keyword}': {e}")
            return None

    def populate_title(self):
        """ Populates the main title paragraph. """
        print("--- Populating Title ---")
        try:
            if 'general_info' in self.ai_analysis and 'title_line' in self.ai_analysis['general_info']:
                 title_text = self.ai_analysis['general_info']['title_line']
                 if self.TITLE_PARA_INDEX < len(self.doc.paragraphs):
                     para0 = self.doc.paragraphs[self.TITLE_PARA_INDEX]
                     # Clear existing runs before setting new text
                     for run in para0.runs:
                        run.clear()
                     para0.text = title_text # Set text directly
                     # Optional: Re-apply bold or specific formatting if needed
                     # para0.runs[0].font.bold = True
                     print(f"  Populated title: '{title_text[:100]}...'")
                 else:
                      print(f"  Error: Paragraph index {self.TITLE_PARA_INDEX} out of range for title.")
            else:
                print("  Warning: Title line not found in AI analysis.")
        except Exception as e:
            print(f"  Error populating title: {e}")

    def populate_general_info(self):
        """ Populates fields in the General Information table. """
        print("--- Populating General Info Table ---")
        try:
            if self.GENERAL_INFO_TABLE_INDEX < len(self.doc.tables):
                 table0 = self.doc.tables[self.GENERAL_INFO_TABLE_INDEX]
                 gen_info = self.ai_analysis.get('general_info', {})
                 tc_id = gen_info.get('test_case_id', 'N/A')
                 test_id = gen_info.get('test_id', 'N/A')

                 # Use defined cell coordinates
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

                 # Add other static fields if needed (e.g., Product Name if derivable)

            else:
                 print(f"  Error: Table index {self.GENERAL_INFO_TABLE_INDEX} out of range for General Info.")
        except Exception as e:
            print(f"  Error populating General Info table: {e}")

    def populate_equipment(self):
        """ Populates the Equipment and Material section table. """
        print("--- Populating Equipment Table ---")
        try:
            if self.EQUIPMENT_TABLE_INDEX < len(self.doc.tables):
                 table1 = self.doc.tables[self.EQUIPMENT_TABLE_INDEX]
                 equipment_list = self.ai_analysis.get('equipment', [])
                 if not equipment_list:
                     # Fallback to parsing setup text if AI fails
                     print("  Warning: Equipment list not found in AI analysis, attempting to parse from setup text.")
                     equipment_list = parse_setup_equipment(self.testrail_data.get('custom_tc_setup', ''))

                 print(f"  Equipment to populate: {equipment_list}")
                 found_items_in_table = []

                 # Iterate through the expected data rows in the template table
                 for r in range(self.EQUIPMENT_START_DATA_ROW, min(self.EQUIPMENT_END_DATA_ROW + 1, len(table1.rows))):
                     try:
                         cell_obj = table1.cell(r, self.EQUIPMENT_NAME_COL)
                         cell_text = "".join([p.text for p in cell_obj.paragraphs]).strip()

                         if not cell_text:
                             continue # Skip empty template rows

                         # Attempt to match equipment from the list to this row
                         for eq_name in equipment_list:
                             clean_cell_text = cell_text.lower().strip().rstrip('.')
                             clean_eq_name = eq_name.lower().strip().rstrip('.')

                             # Use 'in' for partial matches (e.g., "CADD Solis Infuser." matches "CADD Solis Infuser")
                             # Check startswith as well for more precise matching if needed
                             # if clean_cell_text.startswith(clean_eq_name): # Stricter
                             if clean_eq_name in clean_cell_text or clean_cell_text in clean_eq_name: # More lenient
                                 # Found a match - populate placeholder cells for this row
                                 try:
                                     set_cell_text(table1.cell(r, self.EQUIPMENT_QTY_COL), "[Enter Qty]")
                                     set_cell_text(table1.cell(r, self.EQUIPMENT_PART_COL), "[Enter Part#]")
                                     set_cell_text(table1.cell(r, self.EQUIPMENT_LOT_COL), "[Enter Lot/SN]")
                                     set_cell_text(table1.cell(r, self.EQUIPMENT_CAL_COL), "[Enter Cal/SW]")
                                     found_items_in_table.append(eq_name)
                                     print(f"  Populated placeholders in row {r} for equipment: '{cell_text}' (matched '{eq_name}')")
                                     # Remove from list to avoid re-matching (be careful if names are substrings)
                                     # equipment_list.remove(eq_name) # Potential issue if iterating and modifying
                                     break # Move to next row in template
                                 except IndexError:
                                      print(f"  Error: Column index out of range while populating placeholders in row {r} for '{eq_name}'.")
                                      break
                     except IndexError:
                         print(f"  Warning: Cell index ({r}, {self.EQUIPMENT_NAME_COL}) out of range in Equipment table.")
                         continue

                 print(f"  Finished populating equipment placeholders. Items matched: {len(set(found_items_in_table))}/{len(equipment_list)}")
                 unmatched = [eq for eq in equipment_list if eq not in found_items_in_table]
                 if unmatched:
                     print(f"  Warning: Equipment items not matched to template rows: {unmatched}")
                     # TODO: Optionally, add unmatched items as new rows if template allows

            else:
                 print(f"  Error: Table index {self.EQUIPMENT_TABLE_INDEX} out of range for Equipment.")
        except Exception as e:
            print(f"  Error populating Equipment table: {e}")
            import traceback
            traceback.print_exc() # Print full traceback for debugging

    def populate_results_tables(self):
        """ Populates the main results table(s) based on AI analysis. """
        print("--- Populating Results Tables ---")
        results_tables_data = self.ai_analysis.get('results_tables', [])
        if not results_tables_data:
            print("  No results table structures found in AI analysis.")
            return

        # --- Basic Indexing Approach (Fragile) ---
        # Assumes AI provides tables in the order they appear in the doc template
        current_table_index = self.FIRST_RESULTS_TABLE_INDEX
        for i, table_data in enumerate(results_tables_data):
            print(f"  Populating Results Table {i+1} (Attempting index {current_table_index})...")
            table_title = table_data.get("title", f"Results Table {i+1}")
            headers = table_data.get("headers", [])
            num_rows = table_data.get("num_rows", 1) # Default to 1 row if not specified

            if not headers:
                print(f"    Warning: No headers found for table '{table_title}'. Skipping.")
                current_table_index += 1
                continue

            try:
                if current_table_index < len(self.doc.tables):
                    target_table = self.doc.tables[current_table_index]
                    print(f"    Found table at index {current_table_index}. Title: '{table_title}'")

                    # --- Populate Title (if a paragraph precedes the table) ---
                    # Try to find the paragraph element just before the table
                    table_element = target_table._element
                    parent = table_element.getparent()
                    preceding_element = None
                    if parent is not None:
                        try:
                            table_xml_index = parent.index(table_element)
                            if table_xml_index > 0:
                                preceding_element = parent[table_xml_index - 1]
                        except ValueError:
                            print("    Warning: Could not determine table index within parent.")

                    if preceding_element is not None and isinstance(preceding_element, CT_P):
                         # Find the corresponding Paragraph object
                         title_para = None
                         for p in self.doc.paragraphs:
                              if p._element == preceding_element:
                                  title_para = p
                                  break
                         if title_para:
                              print(f"    Setting title in preceding paragraph: '{table_title}'")
                              # Clear existing runs
                              for run in title_para.runs:
                                  run.clear()
                              title_para.text = table_title
                              # Optional: Apply formatting (e.g., bold)
                              if title_para.runs: title_para.runs[0].font.bold = True
                         else:
                              print(f"    Warning: Found preceding element for Table {current_table_index}, but couldn't map to Paragraph object.")
                    else:
                         print(f"    Warning: No paragraph found directly preceding Table {current_table_index}. Cannot set title paragraph automatically.")


                    # --- Populate Headers ---
                    if len(target_table.rows) > 0:
                         header_row = target_table.rows[0]
                         if len(header_row.cells) >= len(headers):
                              print(f"    Setting headers: {headers}")
                              for c_idx, header_text in enumerate(headers):
                                   set_cell_text(header_row.cells[c_idx], header_text)
                              # Clear any extra cells in the template header row
                              for c_idx in range(len(headers), len(header_row.cells)):
                                   set_cell_text(header_row.cells[c_idx], "")
                         else:
                              print(f"    Error: Table {current_table_index} has fewer columns ({len(header_row.cells)}) than headers specified by AI ({len(headers)}). Cannot set headers.")
                              # TODO: Add columns if possible/needed (complex)
                              continue # Skip populating this table
                    else:
                         print(f"    Error: Table {current_table_index} has no rows. Cannot set headers.")
                         continue # Skip populating this table

                    # --- Manage Data Rows ---
                    # Assumes template has at least one data row after the header
                    template_data_rows = len(target_table.rows) - 1
                    rows_to_add = num_rows - template_data_rows
                    print(f"    Target data rows: {num_rows}. Template has: {template_data_rows}. Need to add: {rows_to_add}")

                    # Add rows if needed
                    if rows_to_add > 0:
                         print(f"    Adding {rows_to_add} rows...")
                         for _ in range(rows_to_add):
                             # Add row (inherits table structure)
                             new_row = target_table.add_row()
                             # Optional: Copy formatting from a template row if needed (more complex)

                    # Clear/Populate data rows (from 1 up to num_rows)
                    num_cols = len(headers) # Use number of headers set
                    for r_idx in range(1, min(num_rows + 1, len(target_table.rows))):
                         row = target_table.rows[r_idx]
                         if len(row.cells) < num_cols:
                              print(f"    Warning: Row {r_idx} in Table {current_table_index} has fewer cells ({len(row.cells)}) than headers ({num_cols}). Skipping population for this row.")
                              continue

                         for c_idx in range(num_cols):
                             header_lower = headers[c_idx].lower()
                             cell = row.cells[c_idx]
                             if "no." in header_lower or "sample" in header_lower : # Check for sample number column
                                 set_cell_text(cell, str(r_idx)) # Populate sample number
                             elif "pass" in header_lower and "fail" in header_lower: # Check for Pass/Fail column
                                 set_cell_text(cell, "Pass ☐ / Fail ☐") # Add checkbox placeholder
                             elif "sign" in header_lower or "date" in header_lower:
                                 set_cell_text(cell, "") # Clear signature fields
                             else:
                                 # Clear other data cells for execution input
                                 set_cell_text(cell, "") # Or use "[Enter Data]"

                    # Remove extra template rows if num_rows < template_data_rows
                    if rows_to_add < 0:
                         rows_to_remove_count = abs(rows_to_add)
                         print(f"    Removing {rows_to_remove_count} extra template rows...")
                         # Remove from the end upwards
                         for _ in range(rows_to_remove_count):
                              if len(target_table.rows) > num_rows + 1: # +1 for header
                                   row_element_to_remove = target_table.rows[-1]._element
                                   if row_element_to_remove.getparent() is not None:
                                       row_element_to_remove.getparent().remove(row_element_to_remove)
                                   else:
                                       print("    Warning: Could not remove extra row, parent not found.")
                                       break # Stop trying if parent is gone
                              else:
                                   print("    Warning: Tried to remove more rows than available.")
                                   break

                else:
                    print(f"    Error: Table index {current_table_index} out of range for Results Table {i+1}.")
                    # Stop trying to populate further results tables if index is wrong
                    break

                current_table_index += 1 # Move to the next expected table index

            except Exception as e:
                print(f"    Error processing Results Table {i+1} (Index {current_table_index}): {e}")
                import traceback
                traceback.print_exc()
                current_table_index += 1 # Increment index even on error to try next table


    def populate_requirements_summary(self):
        """ Populates the Requirements Summary table. """
        print("--- Populating Requirements Summary Table ---")
        summary_data = self.ai_analysis.get('requirements_summary', [])
        if not summary_data:
            print("  No requirements summary data found in AI analysis.")
            return

        # --- Find the table ---
        # Try finding by index first (assuming last or second last)
        target_table = None
        if self.REQUIREMENTS_SUMMARY_TABLE_INDEX < 0: # Use negative index
            try:
                summary_index = len(self.doc.tables) + self.REQUIREMENTS_SUMMARY_TABLE_INDEX
                if 0 <= summary_index < len(self.doc.tables):
                     target_table = self.doc.tables[summary_index]
                     print(f"  Found summary table using negative index {self.REQUIREMENTS_SUMMARY_TABLE_INDEX} (maps to {summary_index}).")
                else:
                    print(f"  Warning: Negative index {self.REQUIREMENTS_SUMMARY_TABLE_INDEX} resulted in invalid table index {summary_index}.")
            except IndexError:
                 print(f"  Warning: Could not find summary table using negative index {self.REQUIREMENTS_SUMMARY_TABLE_INDEX}.")
        elif self.REQUIREMENTS_SUMMARY_TABLE_INDEX < len(self.doc.tables):
            target_table = self.doc.tables[self.REQUIREMENTS_SUMMARY_TABLE_INDEX]
            print(f"  Found summary table using positive index {self.REQUIREMENTS_SUMMARY_TABLE_INDEX}.")
        else:
             print(f"  Error: Index {self.REQUIREMENTS_SUMMARY_TABLE_INDEX} out of range for summary table.")

        # Fallback: Try finding by preceding title 'TABLE X. REQUIREMENTS...'
        if not target_table:
            target_table = self._find_table_by_preceding_title("TABLE.*REQUIREMENTS")
            if not target_table:
                 print(f"  Error: Could not find Requirements Summary table by index or title keyword.")
                 return # Cannot proceed without the table

        # --- Populate Rows ---
        num_reqs_to_populate = len(summary_data)
        template_rows_available = len(target_table.rows) - self.REQ_SUMMARY_START_ROW # Exclude header
        print(f"  Populating {num_reqs_to_populate} requirements. Template has {template_rows_available} data rows available.")

        for i, req_info in enumerate(summary_data):
            row_index = self.REQ_SUMMARY_START_ROW + i
            if row_index < len(target_table.rows):
                row = target_table.rows[row_index]
                print(f"    Populating row {row_index} for Req: {req_info.get('req_id', 'N/A')}")
                try:
                    # Populate cells based on defined column indices
                    step_text = req_info.get('step', '') if req_info.get('step', 'N/A') != 'N/A' else ''
                    sub_step_text = req_info.get('sub_step', '') if req_info.get('sub_step', 'N/A') != 'N/A' else ''
                    req_id_text = req_info.get('req_id', '')
                    verif_text = req_info.get('verification_text', '')

                    set_cell_text(row.cells[self.REQ_SUMMARY_STEP_COL], step_text)
                    # Handle potential missing Sub-step column
                    if self.REQ_SUMMARY_SUBSTEP_COL < len(row.cells):
                        set_cell_text(row.cells[self.REQ_SUMMARY_SUBSTEP_COL], sub_step_text)
                    set_cell_text(row.cells[self.REQ_SUMMARY_REQID_COL], req_id_text)
                    set_cell_text(row.cells[self.REQ_SUMMARY_VERIF_COL], verif_text)
                    # Set Pass/Fail placeholder
                    set_cell_text(row.cells[self.REQ_SUMMARY_RESULT_COL], "Pass ☐ / Fail ☐")
                    # Clear signature cell if it exists
                    if self.REQ_SUMMARY_SIGN_COL < len(row.cells):
                        set_cell_text(row.cells[self.REQ_SUMMARY_SIGN_COL], "")

                except IndexError:
                    print(f"    Error: Column index out of range while populating summary row {row_index}.")
                    continue # Try next requirement
            else:
                # Need to add a new row (less common for summary, but possible)
                print(f"    Warning: Not enough rows in summary table template for requirement {i+1}. Adding row.")
                try:
                    new_row = target_table.add_row()
                     # Populate the newly added row
                    step_text = req_info.get('step', '') if req_info.get('step', 'N/A') != 'N/A' else ''
                    sub_step_text = req_info.get('sub_step', '') if req_info.get('sub_step', 'N/A') != 'N/A' else ''
                    req_id_text = req_info.get('req_id', '')
                    verif_text = req_info.get('verification_text', '')
                    set_cell_text(new_row.cells[self.REQ_SUMMARY_STEP_COL], step_text)
                    if self.REQ_SUMMARY_SUBSTEP_COL < len(new_row.cells): set_cell_text(new_row.cells[self.REQ_SUMMARY_SUBSTEP_COL], sub_step_text)
                    set_cell_text(new_row.cells[self.REQ_SUMMARY_REQID_COL], req_id_text)
                    set_cell_text(new_row.cells[self.REQ_SUMMARY_VERIF_COL], verif_text)
                    set_cell_text(new_row.cells[self.REQ_SUMMARY_RESULT_COL], "Pass ☐ / Fail ☐")
                    if self.REQ_SUMMARY_SIGN_COL < len(new_row.cells): set_cell_text(new_row.cells[self.REQ_SUMMARY_SIGN_COL], "")
                except IndexError:
                     print(f"    Error: Column index out of range while populating ADDED summary row for req {i+1}.")
                except Exception as add_e:
                     print(f"    Error adding row to summary table: {add_e}")
                     break # Stop trying to add rows


        # --- Clear Extra Template Rows ---
        if num_reqs_to_populate < template_rows_available:
            rows_to_clear_count = template_rows_available - num_reqs_to_populate
            print(f"  Clearing {rows_to_clear_count} extra rows in summary table template.")
            start_clear_row = self.REQ_SUMMARY_START_ROW + num_reqs_to_populate
            for r_idx in range(start_clear_row, len(target_table.rows)):
                 row = target_table.rows[r_idx]
                 for cell in row.cells:
                     set_cell_text(cell, "") # Clear all cells in extra rows

    def save(self, output_path: str):
        """ Saves the populated document. """
        print(f"--- Saving populated document to: {output_path} ---")
        try:
            # Ensure output directory exists
            output_dir = os.path.dirname(output_path)
            if output_dir:
                os.makedirs(output_dir, exist_ok=True)
            self.doc.save(output_path)
            print("  Save successful.")
        except Exception as e:
            print(f"  Error saving document: {e}")
            import traceback
            traceback.print_exc()


# --- Main Execution ---
def main():
    # --- Configuration ---
    # TestRail Config
    TESTRAIL_URL = "https://icumed.testrail.io"  # Replace with your TestRail URL
    TESTRAIL_EMAIL = "jairo.rodriguez@icumed.com"  # Use env var or replace
    TESTRAIL_API_KEY = "uttDDNwG5EMQeiW71KaN-3l9Ndf4mBqWT0xV7TF5F" # Use env var or replace

    # Input/Output Config
    # test_case_id_to_process = 979567 # Example 1 (Functional)
    test_case_id_to_process = 983037 # Example 2 (Functional)
    # test_case_id_to_process = 937646 # Example from initialFiller (Inspection-like?) Requires manual analysis
    # test_case_id_to_process = 985186 # Example from dump (Inspection) Requires manual analysis


    # IMPORTANT: Select the correct template based on test type (Manual for now)
    # You'll need to create these templates or adjust paths
    template_file_functional = r'C:\Users\JR00093781\OneDrive - ICU Medical Inc\Documents\TestRecordGenerator\Templates\Functional_Template_Base.docx'
    template_file_performance = r'C:\Users\JR00093781\OneDrive - ICU Medical Inc\Documents\TestRecordGenerator\Templates\Performance_Template_Base.docx'
    template_file_inspection = r'C:\Users\JR00093781\OneDrive - ICU Medical Inc\Documents\TestRecordGenerator\Templates\Inspection_Template_Base.docx'
    # Choose ONE template for this run (improve selection later)
    template_file_to_use = template_file_functional # Adjust as needed for the test case ID

    output_dir = r"C:\Users\JR00093781\OneDrive - ICU Medical Inc\Documents\TestRecordGenerator\v2\Generated Record Files"
    # Sanitize test case ID for filename
    safe_tc_id = str(test_case_id_to_process)
    output_file_name = os.path.join(output_dir, f"Record File TC {safe_tc_id} - Generated.docx")

    # --- Check Configuration ---
    if "your_email@example.com" in TESTRAIL_EMAIL or "YOUR_TESTRAIL_KEY" in TESTRAIL_API_KEY or "YOUR_API_KEY_HERE" in GEMINI_API_KEY:
         print("Error: Please configure TestRail and Gemini API credentials before running.")
         sys.exit(1)
    if not os.path.exists(template_file_to_use):
         print(f"Error: Selected template file not found: {template_file_to_use}")
         print("Please create the template file or correct the path.")
         # Optional: Offer to use the template from initialFiller example if found
         fallback_template = r'C:\Users\JR00093781\OneDrive - ICU Medical Inc\Documents\TestRecordGenerator\Test Case C938309 - Records file TEMPLATE.docx'
         if os.path.exists(fallback_template):
             print(f"Attempting to use fallback template: {fallback_template}")
             template_file_to_use = fallback_template
         else:
             sys.exit(1)


    # --- Processing ---
    print(f"--- Starting Record File Generation for Test Case ID: {test_case_id_to_process} ---")

    # 1. Fetch TestRail Data
    tr_client = TestRailAPIClient(TESTRAIL_URL, TESTRAIL_EMAIL, TESTRAIL_API_KEY)
    endpoint = f"get_case/{test_case_id_to_process}"
    testrail_data = tr_client.send_get_request(endpoint)

    if not testrail_data:
        print(f"Error: Failed to fetch TestRail data for TC {test_case_id_to_process}. Exiting.")
        sys.exit(1)

    # print("\n--- Fetched TestRail Data (Snippet) ---")
    # print(json.dumps(testrail_data, indent=2)[:1000] + "\n...") # Print snippet for debug

    # 2. Get AI Analysis
    ai_analysis = get_gemini_analysis(testrail_data)

    if not ai_analysis:
        print("Error: Failed to get analysis from Gemini AI. Exiting.")
        sys.exit(1)

    print("\n--- Received AI Analysis (Structure) ---")
    print(json.dumps(ai_analysis, indent=2)) # Print full AI structure


    # 3. Generate Record File
    try:
        generator = RecordFileGenerator(template_file_to_use, testrail_data, ai_analysis)

        # Execute population steps
        generator.populate_title()
        generator.populate_general_info()
        generator.populate_equipment()
        generator.populate_results_tables()
        generator.populate_requirements_summary()

        # Save the document
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