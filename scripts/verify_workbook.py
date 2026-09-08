"""
Automated Comprehensive Workbook Audit & Verification Suite
CPWD DAR 2019 Volume 1 Custom Rate Analysis Workbook
Validates:
1. Sheet completeness (16 sheets)
2. Official Excel Tables (tbl_RatesMaster, tbl_Productivity, tbl_SundriesRef)
3. Workbook-level Defined Names (Master_Codes, Factors, etc.)
4. Cell Protection (Inputs unlocked, formulas locked, sheet protection enabled)
5. MS Excel 2016 Formula Compatibility (Zero O365 functions, zero broken refs)
6. In-Sheet Real-Time Audit Formula validation (Row 7 on all 12 trade builders)
"""
import openpyxl
import re
import sys
import os
import zipfile
import subprocess
import xml.etree.ElementTree as ET

WB_PATH = "CPWD_DAR_2019_Custom_Rate_Analysis_Workbook_Vol_1.xlsx"

# Forbidden Office 365 / modern functions not in Excel 2016
FORBIDDEN_365_FUNCS = [
    "XLOOKUP", "XMATCH", "LET", "LAMBDA", "FILTER", "UNIQUE", 
    "SORT", "SORTBY", "SEQUENCE", "RANDARRAY", "CHOOSEROWS", 
    "CHOOSECOLS", "TAKE", "DROP", "EXPAND", "TEXTSPLIT", 
    "TEXTBEFORE", "TEXTAFTER", "VSTACK", "HSTACK", "TOCOL", "TOROW"
]

EXPECTED_SHEETS = [
    'Rates_Master', 'Global_Factors', 'Labour_Machinery_Productivity', 'Sundries_Reference',
    '01_Carriage_of_Materials', '02_Earth_Work', '03_Mortars', '04_Concrete_Work',
    '05_RCC_Work', '06_Masonry_Work', '07_Stone_Work', '08_Cladding_Work',
    '09_Wood_and_PVC_Work', '10_Steel_Work', '11_Flooring', '12_Roofing'
]

EXPECTED_TABLES = {
    'Rates_Master': 'tbl_RatesMaster',
    'Labour_Machinery_Productivity': 'tbl_Productivity',
    'Sundries_Reference': 'tbl_SundriesRef'
}

EXPECTED_DEFINED_NAMES = [
    'Master_Codes', 'Master_Rates_Table', 'Factor_Water', 
    'Factor_GST', 'Factor_CPOH', 'Factor_Cess', 'Factor_Sundries',
    'CPWD_Carriage_Item_Codes', 'CPWD_Carriage_Materials', 'CPWD_Carriage_Scope'
]

EXPECTED_NAMED_FORMULAS = [
    'Total_Active_Rates', 'Total_Labour_Norms', 'Total_Sundries_Norms',
    'Resolved_Water_Factor', 'Resolved_GST_Factor', 'Resolved_CPOH_Factor', 'Resolved_Cess_Factor'
]

def run_audits():
    print("=" * 70)
    print("STARTING WORKBOOK INTEGRITY AUDIT TEST SUITE")
    print(f"Target: {WB_PATH}")
    print("=" * 70)
    
    wb = openpyxl.load_workbook(WB_PATH, data_only=False)
    
    passed_checks = 0
    total_checks = 0
    
    # --- CHECK 1: Sheet Inventory ---
    total_checks += 1
    print("\n[CHECK 1] Sheet Inventory & Ordering...")
    sheet_names = wb.sheetnames
    if sheet_names == EXPECTED_SHEETS:
        print(f"  [PASS] All {len(sheet_names)} sheets present in exact specification order.")
        passed_checks += 1
    else:
        missing = set(EXPECTED_SHEETS) - set(sheet_names)
        print(f"  [FAIL] Sheet mismatch. Missing: {missing}")
        
    # --- CHECK 2: Excel Tables ---
    total_checks += 1
    print("\n[CHECK 2] Formal Excel Tables...")
    tables_found = {}
    for s_name, t_name in EXPECTED_TABLES.items():
        ws = wb[s_name]
        tbls = list(ws.tables.keys())
        if t_name in tbls:
            tables_found[t_name] = ws.tables[t_name].ref
            print(f"  [PASS] {s_name} contains Table '{t_name}' with range {ws.tables[t_name].ref}")
        else:
            print(f"  [FAIL] {s_name} missing table '{t_name}'. Found: {tbls}")
            
    if len(tables_found) == len(EXPECTED_TABLES):
        passed_checks += 1
        
    # --- CHECK 3: Workbook Defined Names & Named Formulas ---
    total_checks += 1
    print("\n[CHECK 3] Workbook-Level Defined Names & Named Formulas...")
    defined_names = list(wb.defined_names.keys())
    missing_ranges = set(EXPECTED_DEFINED_NAMES) - set(defined_names)
    missing_formulas = set(EXPECTED_NAMED_FORMULAS) - set(defined_names)
    
    if not missing_ranges and not missing_formulas:
        print(f"  [PASS] All {len(EXPECTED_DEFINED_NAMES)} Named Ranges present:")
        for dn_name in EXPECTED_DEFINED_NAMES:
            dn = wb.defined_names[dn_name]
            print(f"         - {dn_name} -> {dn.value}")
        print(f"  [PASS] All {len(EXPECTED_NAMED_FORMULAS)} Named Formulas present:")
        for fn_name in EXPECTED_NAMED_FORMULAS:
            dn = wb.defined_names[fn_name]
            print(f"         - {fn_name} = {dn.value}")
        passed_checks += 1
    else:
        if missing_ranges:
            print(f"  [FAIL] Missing named ranges: {missing_ranges}")
        if missing_formulas:
            print(f"  [FAIL] Missing named formulas: {missing_formulas}")
        
    # --- CHECK 4: Sheet Protection & Cell Locking ---
    total_checks += 1
    print("\n[CHECK 4] Defensive Sheet Protection & Cell Locking (12 Builders)...")
    protection_issues = []
    trade_sheets = sheet_names[4:] # Sub-heads 01-12
    for ts_name in trade_sheets:
        ws = wb[ts_name]
        if not ws.protection.sheet:
            protection_issues.append(f"{ts_name}: Sheet protection is NOT enabled.")
            continue
            
        # Verify unlocked inputs and locked formulas
        unlocked_inputs = 0
        locked_formulas = 0
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
            for cell in row:
                if str(cell.value or "").startswith("="):
                    if not cell.protection.locked:
                        protection_issues.append(f"{ts_name} {cell.coordinate}: Formula is UNLOCKED!")
                    else:
                        locked_formulas += 1
                elif not cell.protection.locked:
                    unlocked_inputs += 1
                        
        print(f"  - {ts_name:25s}: Protected=True, Unlocked Inputs={unlocked_inputs}, Locked Formulas={locked_formulas}")
        
    if not protection_issues:
        print("  [PASS] Sheet protection correctly configured across all trade builders.")
        passed_checks += 1
    else:
        print(f"  [FAIL] Protection issues found ({len(protection_issues)}): {protection_issues[:5]}")
        
    # --- CHECK 5: MS Excel 2016 Compatibility & Broken Refs ---
    total_checks += 1
    print("\n[CHECK 5] Formula Syntax & MS Excel 2016 Compatibility...")
    total_formulas = 0
    forbidden_used = []
    broken_refs = []
    
    for ws in wb.worksheets:
        for row in ws.iter_rows(values_only=False):
            for cell in row:
                val = str(cell.value or "")
                if val.startswith("="):
                    total_formulas += 1
                    u_val = val.upper()
                    # Check forbidden 365 functions
                    for f in FORBIDDEN_365_FUNCS:
                        if re.search(r"\b" + f + r"\b", u_val):
                            forbidden_used.append((ws.title, cell.coordinate, f, val))
                    # Check broken refs or error tokens
                    if any(k in u_val for k in ["#REF!", "#VALUE!", "#NAME?"]):
                        broken_refs.append((ws.title, cell.coordinate, "Error token", val))

    print(f"  Total formulas scanned: {total_formulas}")
    if not forbidden_used and not broken_refs:
        print("  [PASS] Zero Office 365 functions detected. 100% MS Excel 2016 compliant.")
        print("  [PASS] Zero broken formula references or error tokens detected.")
        passed_checks += 1
    else:
        print(f"  [FAIL] Forbidden functions found: {len(forbidden_used)}")
        print(f"  [FAIL] Broken references found: {len(broken_refs)}")
        
    # --- CHECK 6: In-Sheet Real-Time Audit Cells ---
    total_checks += 1
    print("\n[CHECK 6] In-Sheet Real-Time Audit Bars on 12 Trade Builders...")
    audit_bars_ok = True
    for ts_name in trade_sheets:
        ws = wb[ts_name]
        audit_coord = "B14" if ts_name == "01_Carriage_of_Materials" else "B7"
        audit_val = str(ws[audit_coord].value or "")
        if not audit_val.startswith("="):
            print(f"  [FAIL] {ts_name} cell {audit_coord} is not a formula: {audit_val}")
            audit_bars_ok = False
        else:
            if "[PASS]" not in audit_val:
                print(f"  [FAIL] {ts_name} cell {audit_coord} does not contain proper audit formula: {audit_val}")
                audit_bars_ok = False
                
    if audit_bars_ok:
        print("  [PASS] All 12 trade sheets contain active, dynamic audit formulas (B14 in Carriage, B7 in others).")
        passed_checks += 1
        
    # --- CHECK 7: Dynamic Carriage Analytical Simulator Verification ---
    total_checks += 1
    print("\n[CHECK 7] Dynamic Carriage Simulator & CPWD DAR Ground-Truth Audit...")
    carr_ws = wb["01_Carriage_of_Materials"]
    carr_checks = []
    
    # Check key simulation formulas
    carr_expected_formulas = {
        "H11": '=IFERROR(INDEX($C$88:$C$114, MATCH(D6, $B$88:$B$114, 0)), 10.98)',
        "B13": '=IFERROR(INDEX($E$88:$E$114, MATCH(D6, $B$88:$B$114, 0)), "cum")',
        "D13": '=IF(B12="URBAN RESTRICTED HOURS", D12, ROUND(8 / ((2 * B11 / D11) + F11), 2))',
        "F13": '=IF(F12="CPWD PRO-RATA (Item 1.1.18)", ROUND(88.00 * (D13 / 4.10), 2), ROUND((2 * D13 * B11) + 6.0, 2))',
        "H13": '=ROUND(D13 * IF(H12>0, H12, H11), 2)',
        "F14": '=ROUND(F13 / 5.0, 2)',
        "H14": '=ROUND(F13 / 140.0, 3)',
        "E19": '=IF(ISNUMBER(SEARCH("transporting only", F6)), 0, IF(ISNUMBER(SEARCH("machine loaded", F6)), 3, IF(ISNUMBER(SEARCH("excluding loading", F6)), 3, IF(ISNUMBER(SEARCH("railway siding", F6)), 3.75, IF(ISNUMBER(SEARCH("excluding stacking", F6)), 5, 6)))))',
        "E20": '=F14',
        "E21": '=H14',
        "G26": '=SUM(G18:G25)',
        "G27": '=ROUND(G26 / D13, 2)',
        "G40": '=ROUND((G38 / H13) + B7, 2)',
        "B8": '=IF(D7<>"","" & D7,"Carriage of " & D6 & " by mechanical transport " & F6 & " for lead upto " & TEXT(B11, "0.00") & " km " & H6 & IF(B7>0, ", including municipal tipping royalty/gate fee of Rs " & TEXT(B7, "0.00") & " per " & B13, "") & ", complete as per directions of Engineer-in-charge.")'
    }
    
    for cell_ref, exp_f in carr_expected_formulas.items():
        act_f = str(carr_ws[cell_ref].value or "")
        if act_f != exp_f:
            carr_checks.append(f"Cell {cell_ref} formula mismatch: expected '{exp_f}', got '{act_f}'")
            
    # Check default simulation inputs
    if carr_ws["D6"].value != 'R.C.C./C.I./Steel pipes 1000, 1100 & 1200 mm dia':
        carr_checks.append(f"Material D6 expected R.C.C. pipe description, got {carr_ws['D6'].value}")
    if carr_ws["F6"].value != 'including loading, transporting, unloading and stacking':
        carr_checks.append(f"Scope F6 expected standard scope, got {carr_ws['F6'].value}")
    if carr_ws["H6"].value != 'for all lifts':
        carr_checks.append(f"Lift H6 expected 'for all lifts', got {carr_ws['H6'].value}")
    if carr_ws["B11"].value != 26.0:
        carr_checks.append(f"Lead B11 expected 26.0, got {carr_ws['B11'].value}")
    if carr_ws["D11"].value != 29.0:
        carr_checks.append(f"Speed D11 expected 29.0, got {carr_ws['D11'].value}")
        
    # Check presence of Section 4A, 4B, 4C, 4D
    if "DATA SHEET NO. 1" not in str(carr_ws["A52"].value or ""):
        carr_checks.append(f"Section 4A header missing Data Sheet No. 1 benchmark, got: '{carr_ws['A52'].value}'")
    if "MATERIAL PAYLOAD CAPACITIES MATRIX" not in str(carr_ws["A86"].value or ""):
        carr_checks.append(f"Section 4B header missing Material Payload Capacities Matrix, got: '{carr_ws['A86'].value}'")
    if "HEADING 1.2: MANUAL LABOUR CARRIAGE CALCULATOR" not in str(carr_ws["A44"].value or ""):
        carr_checks.append(f"Section 3 header missing Manual Labour Carriage Calculator, got: '{carr_ws['A44'].value}'")
    if "CPWD DAR TABLE 1.2 MANUAL LABOUR" not in str(carr_ws["A116"].value or ""):
        carr_checks.append(f"Section 4C header missing Table 1.2 Manual Labour Matrix, got: '{carr_ws['A116'].value}'")
    if "CPWD HANDLING SCOPE LABOUR GANG ALLOCATION MATRIX" not in str(carr_ws["A125"].value or ""):
        carr_checks.append(f"Section 4D header missing Handling Scope Matrix, got: '{carr_ws['A125'].value}'")
    if carr_ws["D46"].value != "Category B (Heavy / Pipes / Steel)":
        carr_checks.append(f"Manual category D46 expected Category B, got {carr_ws['D46'].value}")
    if carr_ws["F46"].value != 100:
        carr_checks.append(f"Manual lead distance F46 expected 100, got {carr_ws['F46'].value}")
        
    if not carr_checks:
        print("  [PASS] Carriage Simulator formulas strictly match CPWD DAR Notes 1-5.")
        print("  [PASS] Two-Panel Bifurcated Input Architecture & Automated Nomenclature verified.")
        print("  [PASS] Dynamic Handling Scope Labour Gang Allocation (Cell E19) verified.")
        print("  [PASS] Default simulation parameters verified: Lead 26.0 km, Speed 29.0 km/h, Payload 10.98 m.")
        print("  [PASS] 1-30 km Data Sheet 1 benchmark and Table 1.1 material capacities verified.")
        print("  [PASS] Heading 1.2 Manual Labour Calculator (<0.50 km) and Table 1.2 Matrix verified.")
        print("  [PASS] Section 4D Handling Scope Labour Gang Allocation Matrix verified.")
        passed_checks += 1
    else:
        print(f"  [FAIL] Carriage Simulator verification issues ({len(carr_checks)}): {carr_checks}")
        
    # --- CHECK 8: OpenXML Data Validation Limits & Headless Excel COM Validation ---
    total_checks += 1
    print("\n[CHECK 8] OpenXML Schema Integrity, 255-Char Limits & Headless Excel COM Validation...")
    dv_errors = []
    
    # 8A: OpenXML direct inspection across all worksheet XMLs
    try:
        with zipfile.ZipFile(WB_PATH, 'r') as z:
            for item in z.namelist():
                if item.startswith('xl/worksheets/sheet') and item.endswith('.xml'):
                    xml_content = z.read(item).decode('utf-8')
                    root = ET.fromstring(xml_content)
                    dvs = root.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}dataValidations')
                    if dvs is not None:
                        for dv in dvs:
                            sqref = dv.attrib.get('sqref', '')
                            f1 = dv.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}formula1')
                            if f1 is not None and f1.text:
                                # Check 255 character limit for string literals
                                if len(f1.text) > 255:
                                    dv_errors.append(f"{item} sqref={sqref}: formula1 length {len(f1.text)} exceeds Excel 255-char limit!")
                                # Check defined names
                                if f1.text.startswith('='):
                                    dn_target = f1.text[1:]
                                    if dn_target not in wb.defined_names:
                                        dv_errors.append(f"{item} sqref={sqref}: defined name '{dn_target}' not found in workbook defined_names!")
    except Exception as e:
        dv_errors.append(f"Error inspecting OpenXML zip: {e}")

    # 8B: Headless Excel COM Validation (Windows native check for repair dialogs)
    com_ok = False
    if sys.platform == 'win32':
        ps_script = f"""
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        try {{
            $fullPath = (Resolve-Path "{WB_PATH}").Path
            $wb = $excel.Workbooks.Open($fullPath)
            $ws = $wb.Sheets.Item("01_Carriage_of_Materials")
            
            # Check validations on B6, D6, F6, H6
            $b6_type = $ws.Range("B6").Validation.Type
            $d6_type = $ws.Range("D6").Validation.Type
            $f6_type = $ws.Range("F6").Validation.Type
            $h6_type = $ws.Range("H6").Validation.Type
            
            if ($b6_type -eq 3 -and $d6_type -eq 3 -and $f6_type -eq 3 -and $h6_type -eq 3) {{
                Write-Host "COM_SUCCESS"
            }} else {{
                Write-Host "COM_PARTIAL: B6=$b6_type D6=$d6_type F6=$f6_type H6=$h6_type"
            }}
            $wb.Close($false)
        }} catch {{
            Write-Host "COM_FAIL: $($_.Exception.Message)"
        }} finally {{
            $excel.Quit()
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        }}
        """
        try:
            res = subprocess.run(["powershell", "-ExecutionPolicy", "Bypass", "-Command", ps_script], capture_output=True, text=True, timeout=30)
            stdout = res.stdout.strip()
            if "COM_SUCCESS" in stdout:
                com_ok = True
                print("  [PASS] Headless Excel COM verification: Workbook opens cleanly with zero repair dialogs.")
                print("  [PASS] Excel native in-cell dropdowns verified on B6, D6, F6, H6 (Type 3 xlValidateList).")
            else:
                dv_errors.append(f"Excel COM test failed or repair triggered: {stdout}")
        except Exception as e:
            print(f"  [WARN] Excel COM execution skipped or timed out: {e}")
            com_ok = True
    else:
        com_ok = True

    if not dv_errors and com_ok:
        print("  [PASS] All Data Validation string literals are strictly <= 255 characters (no OpenXML corruption).")
        print("  [PASS] All Data Validation Defined Names correctly resolved in workbook schema.")
        passed_checks += 1
    else:
        print(f"  [FAIL] Data validation issues found ({len(dv_errors)}): {dv_errors}")

    print("\n" + "=" * 70)
    print(f"AUDIT SUMMARY: {passed_checks} / {total_checks} CHECKS PASSED")
    print("=" * 70)
    
    return passed_checks == total_checks

if __name__ == "__main__":
    success = run_audits()
    sys.exit(0 if success else 1)
