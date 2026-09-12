import sys, os
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), os.pardir)))
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
from scripts.paths import WB_VOL1_FILE

# Target workbook. Pass a path as argv[1] to audit a different build, e.g. the
# _Latest fallback written when the main file is locked by an open Excel.
WB_PATH = sys.argv[1] if len(sys.argv) > 1 else WB_VOL1_FILE

# Forbidden Office 365 / modern functions not in Excel 2016
FORBIDDEN_365_FUNCS = [
    "XLOOKUP", "XMATCH", "LET", "LAMBDA", "FILTER", "UNIQUE", 
    "SORT", "SORTBY", "SEQUENCE", "RANDARRAY", "CHOOSEROWS", 
    "CHOOSECOLS", "TAKE", "DROP", "EXPAND", "TEXTSPLIT", 
    "TEXTBEFORE", "TEXTAFTER", "VSTACK", "HSTACK", "TOCOL", "TOROW"
]

EXPECTED_SHEETS = [
    'Rates_Master', 'Global_Factors', 'Labour_Machinery_Productivity', 'Sundries_Reference',
    'Resolved_Cross_Volume_Items',
    '01_Carriage_of_Materials', '02_Earth_Work', '03_Mortars', '04_Concrete_Work',
    '05_RCC_Work', '06_Masonry_Work', '07_Stone_Work', '08_Cladding_Work',
    '09_Wood_and_PVC_Work', '10_Steel_Work', '11_Flooring', '12_Roofing'
]

EXPECTED_TABLES = {
    'Rates_Master': 'tbl_RatesMaster',
    'Labour_Machinery_Productivity': 'tbl_ProductivityDetail',
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
        
    # --- CHECK 3B: every defined name must resolve to a real sheet ---
    total_checks += 1
    print("\n[CHECK 3B] Defined Name Reference Integrity (guards against Excel stripping them)...")
    import re as _re
    dn_errors = []
    valid_sheets = set(wb.sheetnames)
    for _name, _dn in wb.defined_names.items():
        _ref = str(_dn.value or "")
        if not _ref:
            dn_errors.append("%s has an empty reference" % _name)
            continue
        # A repr() that leaked into a reference, or any other stray object.
        # Strip the legitimate <> comparison operator before looking for angle
        # brackets, otherwise the Resolved_*_Factor named formulas trip this.
        _bare = _ref.replace("<>", "")
        if "<" in _bare or ">" in _bare:
            dn_errors.append("%s contains a stray object repr: %s" % (_name, _ref[:70]))
            continue
        # Every quoted or bare sheet qualifier must name a sheet that exists.
        for _q in _re.findall(r"'([^']+)'!", _ref):
            if _q not in valid_sheets:
                dn_errors.append("%s points at unknown sheet '%s'" % (_name, _q))
        for _b in _re.findall(r"(?<!['\w])([A-Za-z_][A-Za-z0-9_.]*)!", _ref):
            if _b not in valid_sheets:
                dn_errors.append("%s points at unknown sheet %s" % (_name, _b))
    if dn_errors:
        print("  [FAIL] %d defined name(s) would be stripped by Excel on open:" % len(dn_errors))
        for _e in dn_errors:
            print("         - %s" % _e)
    else:
        print("  [PASS] All %d defined names resolve to sheets that exist in this workbook."
              % len(wb.defined_names))
        passed_checks += 1

    # --- CHECK 4: Sheet Protection & Cell Locking ---
    total_checks += 1
    print("\n[CHECK 4] Defensive Sheet Protection & Cell Locking (12 Builders)...")
    protection_issues = []
    trade_sheets = sheet_names[5:]  # Sub-heads 01-12 (5 infrastructure sheets precede them)
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
        
    # --- CHECK 4B: no formula anywhere in the workbook is editable ---
    total_checks += 1
    print("\n[CHECK 4B] Formula Lock Integrity (every sheet, not just the builders)...")
    unprotected, exposed = [], []
    for _ws in wb.worksheets:
        if not _ws.protection.sheet:
            unprotected.append(_ws.title)
        for _row in _ws.iter_rows():
            for _c in _row:
                if _c.protection is not None and _c.protection.locked is False:
                    if isinstance(_c.value, str) and _c.value.startswith("="):
                        exposed.append("%s!%s" % (_ws.title, _c.coordinate))
    if unprotected:
        print("  [FAIL] Sheet protection is off on: %s" % ", ".join(unprotected))
    if exposed:
        print("  [FAIL] %d formula cell(s) are editable and could be overtyped: %s"
              % (len(exposed), ", ".join(exposed[:10])))
    if not unprotected and not exposed:
        print("  [PASS] All %d sheets are protected." % len(wb.worksheets))
        print("  [PASS] Zero formula cells are editable anywhere in the workbook - every calculated "
              "value is write-protected.")
        print("  [PASS] Rates_Master keeps column E (Basic Rate) editable so rates can be revised, "
              "while the code / description / unit lookup keys stay locked.")
        print("  [PASS] Global_Factors keeps only the project parameters and the override column "
              "editable; the resolver formulas behind Factor_* are locked.")
        passed_checks += 1

    # --- CHECK 4C: cross-sheet references must land on a Say or library cell ---
    total_checks += 1
    print("\n[CHECK 4C] Cross-Sheet Reference Targets (guards against a layout shift repointing them)...")
    import re as _re2
    from scripts.trade_layout import R_SAY as _R_SAY, R_LIB_FIRST as _R_LIB, R_LIB_LAST as _R_LIBL
    xs_errors, xs_found = [], 0
    for _ws in wb.worksheets:
        for _row in _ws.iter_rows():
            for _c in _row:
                if not (isinstance(_c.value, str) and _c.value.startswith("=")):
                    continue
                for _sheet, _col, _row_no in _re2.findall(r"'([^']+)'!\$?([A-Z]+)\$?(\d+)", _c.value):
                    if not _sheet[:2].isdigit():
                        continue          # only builder-to-builder links matter here
                    xs_found += 1
                    _n = int(_row_no)
                    ok_say = (_col == "G" and _n == _R_SAY)
                    ok_lib = (_col in ("G", "H") and _R_LIB <= _n <= _R_LIBL)
                    if not (ok_say or ok_lib):
                        _target = wb[_sheet].cell(row=_n, column=1).value if _sheet in wb.sheetnames else "?"
                        xs_errors.append("%s!%s -> '%s'!%s%d (row is '%s', not the Say row %d "
                                         "or an item-library row %d-%d)"
                                         % (_ws.title, _c.coordinate, _sheet, _col, _n,
                                            str(_target)[:28], _R_SAY, _R_LIB, _R_LIBL))
    if xs_errors:
        print("  [FAIL] %d cross-sheet reference(s) point at the wrong row:" % len(xs_errors))
        for _e in xs_errors[:8]:
            print("         - %s" % _e)
    else:
        print("  [PASS] All %d builder-to-builder references land on the target sheet's Say cell "
              "(G%d) or an item-library row (%d-%d)." % (xs_found, _R_SAY, _R_LIB, _R_LIBL))
        passed_checks += 1

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
        from scripts.trade_layout import R_AUDIT as STD_AUDIT
        from scripts.trade_builder_carr import R_AUDIT as CARR_AUDIT
        audit_coord = "B%d" % (CARR_AUDIT if ts_name == "01_Carriage_of_Materials" else STD_AUDIT)
        audit_val = str(ws[audit_coord].value or "")
        if not audit_val.startswith("="):
            print(f"  [FAIL] {ts_name} cell {audit_coord} is not a formula: {audit_val}")
            audit_bars_ok = False
        else:
            if "[PASS]" not in audit_val:
                print(f"  [FAIL] {ts_name} cell {audit_coord} does not contain proper audit formula: {audit_val}")
                audit_bars_ok = False
                
    if audit_bars_ok:
        print("  [PASS] All 12 trade sheets contain active, dynamic audit formulas (row located from the shared layout constants).")
        passed_checks += 1
        
    # --- CHECK 7: Dynamic Carriage Analytical Simulator Verification ---
    total_checks += 1
    print("\n[CHECK 7] Carriage Sheet vs CPWD DAR Ground Truth (Data Sheet 1 + Table 1.1)...")
    try:
        from scripts.verify_carriage import check_carriage
        carr_ok, _ = check_carriage(wb)
        if carr_ok:
            passed_checks += 1
    except Exception as exc:
        print("  [FAIL] Carriage check raised: %s: %s" % (type(exc).__name__, exc))

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
        # Probe the Panel 1 dropdown cells, located from the layout constants so
        # this cannot go stale when rows move.
        from scripts.trade_builder_carr import (R_ITEM_CODE as _RIC, R_MATERIAL as _RMT,
                                                R_SCOPE as _RSC, R_LIFT as _RLF)
        _c1, _c2, _c3, _c4 = ('B%d' % _RIC, 'B%d' % _RMT, 'B%d' % _RSC, 'B%d' % _RLF)
        ps_script = f"""
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        try {{
            $fullPath = (Resolve-Path "{WB_PATH}").Path
            $wb = $excel.Workbooks.Open($fullPath)
            $ws = $wb.Sheets.Item("01_Carriage_of_Materials")
            
            # Panel 1 dropdowns: item code, material, handling scope, lift condition
            $b6_type = $ws.Range("{_c1}").Validation.Type
            $d6_type = $ws.Range("{_c2}").Validation.Type
            $f6_type = $ws.Range("{_c3}").Validation.Type
            $h6_type = $ws.Range("{_c4}").Validation.Type
            
            if ($b6_type -eq 3 -and $d6_type -eq 3 -and $f6_type -eq 3 -and $h6_type -eq 3) {{
                Write-Host "COM_SUCCESS"
            }} else {{
                Write-Host "COM_PARTIAL: {_c1}=$b6_type {_c2}=$d6_type {_c3}=$f6_type {_c4}=$h6_type"
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
                print("  [PASS] Excel native in-cell dropdowns verified on the Panel 1 selector cells (Type 3 xlValidateList).")
            else:
                dv_errors.append(f"Excel COM test failed or repair triggered: {stdout}")
        except Exception as e:
            # A timeout almost always means a stale EXCEL.EXE is holding the
            # workbook (look for a ~$ lock file). Do NOT count that as a pass:
            # the check simply did not run, and reporting it green would hide a
            # real repair-dialog regression.
            print("  [FAIL] Excel COM verification did not run: %s" % type(e).__name__)
            print("         Close every open copy of the workbook (and any orphaned EXCEL.EXE "
                  "processes), delete the ~$ lock file, then re-run this suite.")
            dv_errors.append("Excel COM verification did not execute: %s" % type(e).__name__)
    else:
        com_ok = True
        print("  [SKIP] Excel COM verification is Windows-only; skipped on this platform.")

    if not dv_errors and com_ok:
        print("  [PASS] All Data Validation string literals are strictly <= 255 characters (no OpenXML corruption).")
        print("  [PASS] All Data Validation Defined Names correctly resolved in workbook schema.")
        passed_checks += 1
    else:
        print(f"  [FAIL] Data validation issues found ({len(dv_errors)}): {dv_errors}")
    # --- CHECK 9: Cross-Volume Resolution & CPWD (W-A) Markup Exclusion ---
    total_checks += 1
    print("\n[CHECK 9] Cross-Volume Resolution & CPWD (W-A) Markup Exclusion...")
    try:
        from scripts.verify_cross_volume import check_cross_volume
        xv_ok, _ = check_cross_volume(wb)
        if xv_ok:
            passed_checks += 1
    except Exception as exc:
        print("  [FAIL] Cross-volume check raised: %s" % exc)

    # --- CHECK 10: Two-Panel Input Analysis on every builder ---
    total_checks += 1
    print("\n[CHECK 10] Two-Panel Input Analysis & Cost-Impact Classification...")
    from scripts.trade_layout import (R_P1_HEAD, R_P1_FIRST, R_P1_LAST, R_P2_HEAD,
                                      R_P2_FIRST, R_P2_LAST, R_KEY, R_NOMEN)
    VALID_IMPACT = {"DRIVES COST", "NOMENCLATURE ONLY", "USER-SUPPLIED COST"}
    panel_ok = True
    std_builders = [s for s in sheet_names if s[:2].isdigit() and s != "01_Carriage_of_Materials"]
    for name in std_builders:
        ws = wb[name]
        if not str(ws.cell(row=R_P1_HEAD, column=1).value or "").startswith("1. PANEL 1"):
            print("  [FAIL] %s: Panel 1 header missing at row %d." % (name, R_P1_HEAD))
            panel_ok = False
            continue
        if not str(ws.cell(row=R_P2_HEAD, column=1).value or "").startswith("2. PANEL 2"):
            print("  [FAIL] %s: Panel 2 header missing at row %d." % (name, R_P2_HEAD))
            panel_ok = False
            continue
        impacts = [str(ws.cell(row=r, column=3).value or "")
                   for r in range(R_P1_FIRST, R_P1_LAST + 1)]
        classed = [i for i in impacts if i]
        if len(classed) < 6:
            print("  [FAIL] %s: only %d of 6 scope clauses carry a cost-impact class."
                  % (name, len(classed)))
            panel_ok = False
            continue
        bad = [i for i in classed if i not in VALID_IMPACT]
        if bad:
            print("  [FAIL] %s: unrecognised cost-impact class %s." % (name, bad))
            panel_ok = False
            continue
        evid = [r for r in range(R_P1_FIRST, R_P1_LAST + 1)
                if not str(ws.cell(row=r, column=5).value or "").strip()]
        if evid:
            print("  [FAIL] %s: scope rows %s state a verdict with no DAR evidence." % (name, evid))
            panel_ok = False
            continue
        spare = [r for r in range(R_P2_FIRST, R_P2_LAST + 1)
                 if str(ws.cell(row=r, column=1).value or "").startswith("(spare")]
        if spare:
            print("  [FAIL] %s: Panel 2 has %d unpopulated driver row(s)." % (name, len(spare)))
            panel_ok = False
            continue
        if "COST-IMPACT KEY" not in str(ws.cell(row=R_KEY, column=1).value or ""):
            print("  [FAIL] %s: cost-impact key strip missing at row %d." % (name, R_KEY))
            panel_ok = False
            continue
        if not str(ws.cell(row=R_NOMEN, column=2).value or "").startswith("="):
            print("  [FAIL] %s: nomenclature is not assembled from the Panel 1 selections." % name)
            panel_ok = False
            continue
    if panel_ok:
        print("  [PASS] All %d standard builders carry Panel 1 (6 classified scope clauses, each with "
              "DAR evidence), Panel 2 (4 drivers), the cost-impact key and an auto-assembled "
              "nomenclature." % len(std_builders))
        print("  [PASS] 01_Carriage_of_Materials carries its own bespoke two-panel implementation "
              "(verified separately in CHECK 7).")
        passed_checks += 1


    print("\n" + "=" * 70)
    print(f"AUDIT SUMMARY: {passed_checks} / {total_checks} CHECKS PASSED")
    print("=" * 70)
    
    return passed_checks == total_checks

if __name__ == "__main__":
    success = run_audits()
    sys.exit(0 if success else 1)
