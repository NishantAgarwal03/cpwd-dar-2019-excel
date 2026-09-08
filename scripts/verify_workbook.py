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
        
    # --- CHECK 3: Workbook Defined Names ---
    total_checks += 1
    print("\n[CHECK 3] Workbook-Level Defined Names (Named Ranges)...")
    defined_names = list(wb.defined_names.keys())
    missing_dns = set(EXPECTED_DEFINED_NAMES) - set(defined_names)
    if not missing_dns:
        print(f"  [PASS] All {len(EXPECTED_DEFINED_NAMES)} defined names present:")
        for dn_name in EXPECTED_DEFINED_NAMES:
            dn = wb.defined_names[dn_name]
            print(f"         - {dn_name} -> {dn.value}")
        passed_checks += 1
    else:
        print(f"  [FAIL] Missing defined names: {missing_dns}")
        
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
        
    # --- CHECK 6: In-Sheet Real-Time Audit Cells (Row 7) ---
    total_checks += 1
    print("\n[CHECK 6] In-Sheet Real-Time Audit Bars (Row 7 on 12 Trade Builders)...")
    audit_bars_ok = True
    for ts_name in trade_sheets:
        ws = wb[ts_name]
        b7 = str(ws["B7"].value or "")
        if not b7.startswith("="):
            print(f"  [FAIL] {ts_name} cell B7 is not a formula: {b7}")
            audit_bars_ok = False
        else:
            if "[PASS]" not in b7:
                print(f"  [FAIL] {ts_name} cell B7 does not contain proper audit formula: {b7}")
                audit_bars_ok = False
                
    if audit_bars_ok:
        print("  [PASS] All 12 trade sheets contain active, dynamic audit formulas in Row 7.")
        passed_checks += 1
        
    print("\n" + "=" * 70)
    print(f"AUDIT SUMMARY: {passed_checks} / {total_checks} CHECKS PASSED")
    print("=" * 70)
    
    return passed_checks == total_checks

if __name__ == "__main__":
    success = run_audits()
    sys.exit(0 if success else 1)
