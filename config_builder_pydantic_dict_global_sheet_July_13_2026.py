#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
dynamic_config_builder.py
==========================
Generic Network Configuration Automation Generation Engine
(Zero Restrictions, Native Character Alignment, With-Lock Release, Pure Blind Box Whitelist Version)

Design Logic:
  1. Excel TAB Name = Sheet calling name in Jinja2 template (strictly case-sensitive).
  2. Excel Column Name = Variable name in Jinja2 template (strictly case-sensitive).
  3. Python does not alter any names (no lowercase/underscore conversion), copying them exactly as-is.

Data Structure:
  - 'Node' TAB        : Flat — fields accessed directly as node.FieldName
  - All other TABs    : Always List, regardless of how many rows per device.
                        Even if a device has only 1 row, it is stored as a list with 1 element.

Common Parameters Feature:
  - Any sheet can have optional common parameters placed ABOVE the NodeName header row.
  - Format: one header row immediately followed by one value row (no blank line between them).
  - Common parameters can start from ANY row and ANY column above the NodeName row.
  - Python automatically merges common parameters into EVERY device row of that sheet.
  - If a device row has the same column name as a common parameter, device data wins.
  - No special variable needed in templates — common values appear directly in node.SheetName[0].

Global Sheets Feature:
  - Any sheet WITHOUT a NodeName column is treated as a Global Sheet.
  - Global sheets store data that is shared across ALL devices (e.g. SERVER, QOS, POLICY).
  - Format: first non-empty row = headers, second non-empty row = values.
  - Global sheets are passed directly into every template render as top-level variables.
  - Template access: {{ SERVER.NTP1 }}  {{ QOS.SAP_OUT }}  {{ POLICY.TO_FW }}
  - Global sheets are completely independent — they do NOT affect node data or common params.
  - Sheet names and column names are case-sensitive and used exactly as-is.

Template Access Rules:
  - Node TAB fields    : node.SystemIP / node.NodeName      (direct access)
  - Other sheet data   : node.VPRN1000                      (always a list)
  - Single value       : {% set v = node.VPRN1000[0] %} then v.Port_ID
  - Loop all rows      : {% for row in node.VPRN1000 %}
  - If check           : {% if node.VPRN1000 %}             (empty list [] = False)
  - Global sheet value : {{ SERVER.NTP1 }} / {{ QOS.SAP_OUT }} / {{ POLICY.TO_FW }}

Requirements:
  - In 'Node' tab: Must have 'NodeName' and 'TemplateName' (case-sensitive).
  - Device sheets: Must have 'NodeName' to bind the data to the specific device.
  - Global sheets: Must NOT have 'NodeName' column. First row = headers, second row = values.
"""

from pathlib import Path
import pandas as pd
from jinja2 import Environment, FileSystemLoader, TemplateNotFound
from pydantic import BaseModel, ConfigDict

# ============================================================
# 1. Path Settings
# ============================================================
BASE_DIR      = Path(__file__).resolve().parent
TEMPLATES_DIR = BASE_DIR / "templates"
OUTPUT_DIR    = BASE_DIR / "Configuration-Files"
INPUT_XLSX    = BASE_DIR / "Spreadsheet" / "Spreadsheet v1.xlsx"

# ============================================================
# 2. Generic Dynamic Data Container Model
# ============================================================
class DynamicDataContainer(BaseModel):
    """A completely dynamic Pydantic model that accepts any arbitrary column names present in Excel"""
    model_config = ConfigDict(extra="allow")

# ============================================================
# 3. Data Cleaning Tools
# ============================================================
def clean_value(val):
    """Safely clean Excel cell data. Strip trailing/leading spaces, keep numbers/strings, discard NaN."""
    if pd.isna(val):
        return ""
    val_str = str(val).strip()
    return val_str

def row_to_clean_dict(row_series) -> dict:
    """Convert a Pandas row into a dictionary with cleaned keys and values (strip column name spaces)."""
    return {str(k).strip(): clean_value(v) for k, v in row_series.items()}

# ============================================================
# 4. Common Parameters Reader
# ============================================================
def read_common_params(xl_file, sheet_name: str) -> dict:
    """
    Reads optional common parameters placed ABOVE the NodeName header row in a sheet.

    Rules:
      - Common parameters can start from ANY row and ANY column above the NodeName row.
      - Format: header row immediately followed by value row (no blank line between them).
      - The function automatically finds the last two non-empty rows above NodeName.
      - If no valid common parameters exist, returns an empty dict {} silently — no errors.

    Merge behavior (called from build_network_universe):
      - Common params are merged INTO every device data row of that sheet.
      - Device row values always override common params if same column name exists.
      - No special _common variable in templates — values appear directly in node.SheetName[0].

    Example Excel layout:
      Row 1: (empty)
      Row 2: (empty)   VPRN_ID   SAP_IN   Customer_ID    ← common header (any column ok)
      Row 3: (empty)   1000      1000      1000           ← common values
      Row 4: (empty)
      Row 5: NodeName  Port_ID   InterfaceIP ...          ← device data starts here
      Row 6: R10       1/1/c12/1 100.30.1.18 ...

    After merge, template accesses:
      {% set v = node.VPRN1000[0] %}
      vprn {{ v.VPRN_ID }}      ← from common params
      port {{ v.Port_ID }}      ← from device data
    """
    try:
        df_raw = pd.read_excel(xl_file, sheet_name=sheet_name, header=None, dtype=str)
    except Exception:
        return {}

    if df_raw.empty:
        return {}

    # Find the NodeName header row index
    header_idx = None
    for idx, row in df_raw.iterrows():
        row_values = [str(val).strip() for val in row.dropna()]
        if "NodeName" in row_values:
            header_idx = idx
            break

    # Need at least 2 rows above NodeName for header+value pair
    if header_idx is None or header_idx < 2:
        return {}

    # Scan all rows above NodeName, collect non-empty rows only
    # A row is non-empty if it has at least one non-blank cell
    non_empty_rows = []
    for idx in range(header_idx):
        row = df_raw.iloc[idx]
        values = [str(v).strip() for v in row if pd.notna(v) and str(v).strip()]
        if values:
            non_empty_rows.append(df_raw.iloc[idx])

    # Need at least 2 non-empty rows: one for header, one for values
    if len(non_empty_rows) < 2:
        return {}

    # Take the last two non-empty rows above NodeName:
    # second-to-last = common parameter headers
    # last           = common parameter values
    header_row = non_empty_rows[-2]
    value_row  = non_empty_rows[-1]

    # Build the common params dict, skipping any NaN columns
    common = {}
    for h, v in zip(header_row, value_row):
        key = str(h).strip() if pd.notna(h) else ""
        val = str(v).strip() if pd.notna(v) else ""
        if key and val:
            common[key] = val

    return common

# ============================================================
# 4b. Global Sheet Reader
# ============================================================
def read_global_sheets(xl_file: pd.ExcelFile, all_sheets: list) -> dict:
    """
    Scans all sheets in the workbook and identifies Global Sheets.
    A Global Sheet is any sheet that does NOT contain a 'NodeName' column.
    These sheets store data shared across ALL devices (e.g. SERVER, QOS, POLICY).

    Format rules:
      - First non-empty row  = column headers
      - Second non-empty row = values
      - Any additional rows are ignored
      - Sheet name and column names are used exactly as-is (case-sensitive)
      - Blank cells are read as empty string ""

    Returns a dict of dicts:
      {
          "SERVER": { "NTP1": "10.1.1.1", "NTP2": "10.1.1.2", "SYSLOG": "10.1.1.3" },
          "QOS":    { "SAP_OUT": "1000",  "SAP_IN": "1001" },
          "POLICY": { "TO_FW": "TO-SCADA-FW", "FROM_FW": "FROM-SCADA-FW" },
      }

    Template access (top-level variables, no node. prefix needed):
      {{ SERVER.NTP1 }}
      {{ QOS.SAP_OUT }}
      {{ POLICY.TO_FW }}

    Important:
      - Global sheets are completely independent from device data.
      - They do NOT affect node context, common params, or VPRN/VPLS data in any way.
      - Sheets with NodeName column are device sheets — they are skipped here.
      - 'Node' sheet is always skipped.
    """
    global_context = {}

    for sheet_name in all_sheets:
        # Always skip the Node TAB — it is the device baseline, not a global sheet
        if sheet_name == "Node":
            continue

        try:
            df_raw = pd.read_excel(xl_file, sheet_name=sheet_name, header=None, dtype=str)
        except Exception:
            continue

        if df_raw.empty:
            continue

        # Check if this sheet has a NodeName column — if so, it is a device sheet, skip it
        has_nodename = False
        for _, row in df_raw.iterrows():
            row_values = [str(val).strip() for val in row.dropna()]
            if "NodeName" in row_values:
                has_nodename = True
                break

        if has_nodename:
            # This is a device sheet — handled by read_sheet_safely, skip here
            continue

        # ── This sheet has no NodeName → treat as Global Sheet ──────────────
        # Collect all non-empty rows
        non_empty_rows = []
        for idx in range(len(df_raw)):
            row = df_raw.iloc[idx]
            values = [str(v).strip() for v in row if pd.notna(v) and str(v).strip()]
            if values:
                non_empty_rows.append(df_raw.iloc[idx])

        # Need at least 2 non-empty rows: one header row + one value row
        if len(non_empty_rows) < 2:
            continue

        # First non-empty row = headers, second non-empty row = values
        header_row = non_empty_rows[0]
        value_row  = non_empty_rows[1]

        # Build the global dict for this sheet
        sheet_data = {}
        for h, v in zip(header_row, value_row):
            key = str(h).strip() if pd.notna(h) else ""
            val = str(v).strip() if pd.notna(v) else ""
            if key:
                sheet_data[key] = val

        if sheet_data:
            global_context[sheet_name.strip()] = sheet_data
            #print(f"  Global sheet found: [{sheet_name}] → {sheet_data}")

    return global_context

# ============================================================
# 5. Native Character Smart Sheet Parser (Using xlfile handle)
# ============================================================
def read_excel_with_dynamic_header(xl_file, sheet_name, keyword="NodeName", **kwargs):
    """
    Dynamically locates the header row containing the specified keyword,
    regardless of how many leading empty rows exist in the sheet.
    """
    # 1. Read the raw data without parsed headers (treat headers as regular rows initially)
    df_raw = pd.read_excel(xl_file, sheet_name=sheet_name, header=None, dtype=str)

    if df_raw.empty:
        return pd.DataFrame()

    header_idx = None

    # 2. Scan row by row to find the anchor keyword (e.g., NodeName)
    for idx, row in df_raw.iterrows():
        # Convert row values to stripped strings and filter out NaN values
        row_values = [str(val).strip() for val in row.dropna()]
        if keyword in row_values:
            header_idx = idx
            break

    # 3. If the actual header row is successfully located
    if header_idx is not None:
        # Slice the DataFrame: keep rows below the header row as data
        df_clean = df_raw.iloc[header_idx + 1:].copy()
        # Promote the found anchor row to be the DataFrame columns
        df_clean.columns = df_raw.iloc[header_idx].values
        # Reset index to clean up row numbering
        df_clean.reset_index(drop=True, inplace=True)
        return df_clean
    else:
        # Fallback to standard reading if the target keyword is nowhere to be found
        print(
            f"⚠️ WARNING: Anchor keyword '{keyword}' not found in sheet [{sheet_name}]. Attempting fallback read..."
        )
        return pd.read_excel(xl_file, sheet_name=sheet_name)


def read_sheet_safely(xlfile: pd.ExcelFile, sheet_name: str):
    """
    Smart Blind Reader: Strictly maintains native TAB names.
    Reads safely from an already opened xlfile via 'with' to avoid frequent file openings and deadlocks.

    Data structure rules:
      - 'Node' TAB  → Flat  (handled separately in build_network_universe)
      - All others  → Always List, even if only 1 row per device.
                      Template: {% set v = node.SheetName[0] %} for single value
                                {% for row in node.SheetName %}  for looping
                                {% if node.SheetName %}          for existence check
    """
    try:
        df = read_excel_with_dynamic_header(xlfile, sheet_name=sheet_name, keyword="NodeName", dtype=str)
    except Exception:
        return None

    # Clean leading/trailing spaces from column names, but keep the character case
    df.columns = [str(col).strip() for col in df.columns]
    if "NodeName" not in df.columns:
        # No NodeName column — this is a Global Sheet, not a device sheet
        # It will be handled by read_global_sheets(), skip it here silently
        return None

    df = df.dropna(axis=1, how="all").dropna(subset=["NodeName"])

    # ── Node TAB: keep as Flat for the baseline device lookup ──────────────
    if sheet_name == "Node":
        flat_result = {}
        for _, row in df.iterrows():
            clean_data = row_to_clean_dict(row)
            node_name  = clean_data.get("NodeName")
            flat_result[node_name] = DynamicDataContainer(**clean_data)
        #if "R1" in flat_result:
            #print(flat_result["R1"])
        return {"type": "flat", "data": flat_result}

    # ── All other TABs: ALWAYS List, regardless of row count per device ────
    # Even a device with only 1 row is stored as a list with 1 element.
    # This makes template access consistent: always use for loop or [0] index.
    grouped_result = {}
    for node_name, group in df.groupby("NodeName", sort=False):
        clean_node_name = str(node_name).strip()
        items = [DynamicDataContainer(**row_to_clean_dict(row)) for _, row in group.iterrows()]
        grouped_result[clean_node_name] = items
    #if "R16" in grouped_result:
        #print(grouped_result['R16'])
    return {"type": "list", "data": grouped_result}

# ============================================================
# 6. Native Assembly Engine (With Context Manager & Dynamic Whitelist)
# ============================================================
def build_network_universe(xlsx_path: Path) -> tuple[dict[str, dict], list[dict], dict]:
    """
    Scan the entire Excel, making no modifications to TAB names, aligning them exactly as-is.
    Now returns three values:
      1. master_nodes_context : per-device context dict
      2. pure_nodes_list      : baseline node list (Node sheet only)
      3. global_context       : global sheets dict (SERVER, QOS, POLICY, etc.)
    """

    # Use 'with' to encapsulate ExcelFile. Once out of this block, the file lock is released instantly!
    with pd.ExcelFile(xlsx_path) as xl:
        all_sheets = xl.sheet_names
        #print(all_sheets)

        if "Node" not in all_sheets:
            raise ValueError("Excel must contain a sheet named 'Node' as the baseline device list!")

        sheet_warehouse = {}
        for sheet in all_sheets:
            parsed_sheet = read_sheet_safely(xl, sheet)
            if parsed_sheet:
                sheet_warehouse[sheet] = parsed_sheet
        #print(sheet_warehouse["VPLS"]['data']['R15'][0].NodeName)

        # Read common parameters for all non-Node sheets while Excel is still open
        # Returns {} silently for sheets with no common parameters
        common_params_warehouse = {}
        for sheet in all_sheets:
            if sheet == "Node":
                continue
            common = read_common_params(xl, sheet)
            if common:
                common_params_warehouse[sheet] = common
                #print(f"  Common params found in [{sheet}]: {common}")

        # ── Read all Global Sheets (no NodeName column) ─────────────────────
        # These are completely separate from device data and common params.
        # Examples: SERVER, QOS, POLICY, NTP, SNMP etc.
        global_context = read_global_sheets(xl, all_sheets)
        if global_context:
            print(f"  Global sheets loaded: {list(global_context.keys())}")

    # --- Excel file is closed safely here. You can modify/save Excel anytime now ---

    # Get baseline router list (Node TAB is always Flat)
    nodes_base_data = sheet_warehouse["Node"]["data"]

    # Truly dynamic, zero-restriction global node list generation
    # Extract the original Node sheet blind box dict before any business data is mixed in
    pure_nodes_list = [base_obj.model_dump() for base_obj in nodes_base_data.values()]

    master_nodes_context = {}
    for node_name, base_obj in nodes_base_data.items():
        node_context = base_obj.model_dump()

        if "TemplateName" not in node_context or not node_context["TemplateName"]:
            print(f"⚠ WARNING: Router '{node_name}' is missing 'TemplateName' in Excel, skipped.")
            continue

        # Traverse all dynamic Tabs
        # Node TAB is skipped — its data is already in node_context via base_obj.model_dump()
        # All other TABs are always List type now
        for sheet_name, sheet_info in sheet_warehouse.items():
            if sheet_name == "Node":
                continue

            jinja_key = sheet_name.strip()

            # Get device rows for this router (empty list if no data in this sheet)
            device_rows = sheet_info["data"].get(node_name, [])

            # If this sheet has common parameters, merge them into every device row
            # Common params act as default values — device row values always take priority
            # After merge, templates access everything via node.SheetName[0].FieldName
            # No separate _common variable needed in templates
            if sheet_name in common_params_warehouse and device_rows:
                common = common_params_warehouse[sheet_name]
                merged_rows = []
                for row in device_rows:
                    # Merge: common params as base, device row data overrides
                    # {**common} sets the defaults, {**row.model_dump()} overrides with device data
                    merged_data = {**common, **row.model_dump()}
                    merged_rows.append(DynamicDataContainer(**merged_data))
                node_context[jinja_key] = merged_rows
            else:
                node_context[jinja_key] = device_rows

        master_nodes_context[node_name] = node_context

    # Returns three values: per-device context, baseline node list, global sheets
    return master_nodes_context, pure_nodes_list, global_context

# ============================================================
# 7. Converts Pydantic objects, lists, and dictionaries into Python types.
# ============================================================
def make_pure_dict(data):
    """
    Recursively converts Pydantic objects, lists, and dictionaries into
    standard Python types (dict/list) for seamless Jinja2 template access.
    """
    # Check if the data is a Pydantic model (has 'model_dump')
    if hasattr(data, 'model_dump'):
        # Convert Pydantic object to a dictionary and recurse on values
        return {k: make_pure_dict(v) for k, v in data.model_dump().items()}

    # If the data is a list, recurse through every item
    elif isinstance(data, list):
        return [make_pure_dict(i) for i in data]

    # If the data is a dictionary, recurse through every value
    elif isinstance(data, dict):
        return {k: make_pure_dict(v) for k, v in data.items()}

    # Return primitive types (strings, integers, booleans, None) as is
    return data

# ============================================================
# 8. Configuration Rendering Execution Engine
# ============================================================
def run_generator():
    OUTPUT_DIR.mkdir(exist_ok=True)
    for old_file in OUTPUT_DIR.glob("*.cfg"):
        old_file.unlink()

    print("⚡ [Automation Engine] Parsing original Excel sheets automatically...")
    # Receive three variables: full context, pure global node list, and global sheets context
    master_context, global_nodes_list, global_context = build_network_universe(INPUT_XLSX)
    all_data_dicts = [make_pure_dict(ctx) for ctx in master_context.values()]
    #pure_global_nodes = [make_pure_dict(node) for node in global_nodes_list]
    #node_lookup = { node["NodeName"]: node  for node in pure_global_nodes }
    node_lookup = {node["NodeName"]: node for node in all_data_dicts}
    #print(node_lookup)
    #print(node_lookup['R1']['MDA'][0])

    # global_context is already a plain dict of dicts — no make_pure_dict needed
    # Example: { "SERVER": {"NTP1": "10.1.1.1", ...}, "QOS": {"SAP_OUT": "1000", ...} }
    #print(global_context['SERVER'])

    env = Environment(
        loader=FileSystemLoader(str(TEMPLATES_DIR)),
        trim_blocks=True,
        lstrip_blocks=True,
        keep_trailing_newline=True,
    )

    print(f"⚡ [Automation Engine] Starting to render configurations for {len(master_context)} devices...")
    success_count = 0
    failed        = []

    for node_name, context_obj in master_context.items():

        node_dict = make_pure_dict(context_obj)
        #if node_name == "R10":
            #print(node_dict['VPRN1001'])

        template_file = node_dict.get("TemplateName")
        try:
            template = env.get_template(template_file)
        except TemplateNotFound:
            print(f"❌ {node_name}: TemplateFile '{template_file}' not found")
            failed.append(node_name)
            continue
        try:
            config_text = template.render(
                node=node_dict,
                node_list=all_data_dicts,
                node_lookup=node_lookup,
                # ── Global sheets injected as top-level template variables ──
                # Each global sheet becomes a direct variable in the template.
                # Example: SERVER sheet → {{ SERVER.NTP1 }}
                #          QOS sheet   → {{ QOS.SAP_OUT }}
                #          POLICY sheet→ {{ POLICY.TO_FW }}
                # Adding or removing global sheets in Excel requires no Python changes.
                **global_context
            )
            out_path = OUTPUT_DIR / f"{node_name}.cfg"
            out_path.write_text(config_text, encoding="utf-8")
            success_count += 1
            print(f"  -> Generated successfully: {out_path.name} associated with template: '{template_file}'")
        except Exception as e:
            print(f"❌ Router {node_name} configuration rendering failed: {e}")
            failed.append(node_name)

    print(f"\n{'='*55}")
    print(f"🎉 Finished!  Success: {success_count}  Failed: {len(failed)}")
    if failed:
        print(f"   Failed devices: {failed}")

if __name__ == "__main__":
    run_generator()
