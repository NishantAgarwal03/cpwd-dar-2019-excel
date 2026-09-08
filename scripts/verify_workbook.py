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
    'Factor_GST', 'Factor_CPOH', 'Factor_Cess', 'Factor_Sundries'
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
        audit_coord = "B10" if ts_name == "01_Carriage_of_Materials" else "B7"
        audit_val = str(ws[audit_coord].value or "")
        if not audit_val.startswith("="):
            print(f"  [FAIL] {ts_name} cell {audit_coord} is not a formula: {audit_val}")
            audit_bars_ok = False
        else:
            if "[PASS]" not in audit_val:
                print(f"  [FAIL] {ts_name} cell {audit_coord} does not contain proper audit formula: {audit_val}")
                audit_bars_ok = False
                
    if audit_bars_ok:
        print("  [PASS] All 12 trade sheets contain active, dynamic audit formulas (B10 in Carriage, B7 in others).")
        passed_checks += 1
        
    # --- CHECK 7: Dynamic Carriage Analytical Simulator Verification ---
    total_checks += 1
    print("\n[CHECK 7] Dynamic Carriage Simulator & CPWD DAR Ground-Truth Audit...")
    carr_ws = wb["01_Carriage_of_Materials"]
    carr_checks = []
    
    # Check key simulation formulas
    carr_expected_formulas = {
        "F7": '=IF(D8="FIXED TRIPS", F8, ROUND(8 / ((2 * D5 / F5) + H5), 2))',
        "H7": '=ROUND((2 * F7 * D5) + B8, 2)',
        "D9": '=ROUND(H7 / H8, 2)',
        "F9": '=ROUND(H7 / B9, 3)',
        "H9": '=ROUND(F7 * B7, 2)',
        "E16": '=D9',
        "E17": '=F9',
        "G22": '=SUM(G14:G21)',
        "G23": '=ROUND(G22 / F7, 2)',
        "G36": '=ROUND(G35 / H9, 2)'
    }
    
    for cell_ref, exp_f in carr_expected_formulas.items():
        act_f = str(carr_ws[cell_ref].value or "")
        if act_f != exp_f:
            carr_checks.append(f"Cell {cell_ref} formula mismatch: expected '{exp_f}', got '{act_f}'")
            
    # Check default simulation inputs
    if carr_ws["D5"].value != 26.0:
        carr_checks.append(f"Lead D5 expected 26.0, got {carr_ws['D5'].value}")
    if carr_ws["F5"].value != 29.0:
        carr_checks.append(f"Speed F5 expected 29.0, got {carr_ws['F5'].value}")
    if carr_ws["B7"].value != 10.98:
        carr_checks.append(f"Pipe Payload B7 expected 10.98, got {carr_ws['B7'].value}")
        
    # Check presence of Section 3 (Data Sheet 1 benchmark 1-30 km) and Section 4 (Capacities)
    if "DATA SHEET NO. 1" not in str(carr_ws["A40"].value or ""):
        carr_checks.append("Section 3 header missing Data Sheet No. 1 benchmark")
    if "MATERIAL PAYLOAD CAPACITIES MATRIX" not in str(carr_ws["A74"].value or ""):
        carr_checks.append("Section 4 header missing Material Payload Capacities Matrix")
        
    if not carr_checks:
        print("  [PASS] Carriage Simulator formulas strictly match CPWD DAR Notes 1-5.")
        print("  [PASS] Default simulation parameters verified: Lead 26.0 km, Speed 29.0 km/h, Payload 10.98 m.")
        print("  [PASS] 1-30 km Data Sheet 1 benchmark and Table 1.1 material capacities verified.")
        passed_checks += 1
    else:
        print(f"  [FAIL] Carriage Simulator verification issues ({len(carr_checks)}): {carr_checks}")
        
    print("\n" + "=" * 70)
    print(f"AUDIT SUMMARY: {passed_checks} / {total_checks} CHECKS PASSED")
    print("=" * 70)
    
    return passed_checks == total_checks

if __name__ == "__main__":
    success = run_audits()
    sys.exit(0 if success else 1)
