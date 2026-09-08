import sys, os
sys.path.insert(0, os.path.abspath("."))
# -*- coding: utf-8 -*-
import openpyxl
import os, time
from scripts.styles import get_workbook_styles
from scripts.infra_sheets import (
    build_rates_master,
    build_global_factors,
    build_labour_productivity,
    build_sundries_reference
)
from scripts.trade_builder import build_standard_trade
from scripts.trade_builder_carr import build_carriage_trade
from scripts.trade_configs import get_all_trade_configs
from scripts.scope_inputs import merge_scope_metadata
from scripts.cross_volume import build_resolved_cross_volume

def generate_full_workbook():
    t0 = time.time()
    print("================================================================")
    print("CPWD DAR 2019 VOLUME 1 CUSTOM RATE ANALYSIS WORKBOOK GENERATOR")
    print("Strict MS Excel 2016 Compatibility & excel-estimator-design Standard")
    print("================================================================")
    
    wb = openpyxl.Workbook()
    default_sheet = wb.active
    
    styles = get_workbook_styles()
    
    # 1. Infrastructure Sheets
    print("--- Generating 5 Shared Infrastructure Sheets ---")
    build_rates_master(wb, styles)
    build_global_factors(wb, styles)
    build_labour_productivity(wb, styles)
    build_sundries_reference(wb, styles)
    build_resolved_cross_volume(wb, styles)
    
    # Remove default sheet
    if default_sheet in wb.worksheets:
        wb.remove(default_sheet)
        
    # 2. Trade Builder Sheets (01 to 12)
    print("--- Generating 12 Dedicated Trade Builder Sheets ---")
    configs = merge_scope_metadata(get_all_trade_configs())
    
    # 01 Carriage
    build_carriage_trade(wb, configs['01_Carriage_of_Materials'], styles)
    
    # 02 to 12 Standard Trade Builders (02 carries an empty MATERIAL block:
    # Earth Work items in the DAR are labour-and-plant only.)
    trade_keys = [
        '02_Earth_Work',
        '03_Mortars',
        '04_Concrete_Work',
        '05_RCC_Work',
        '06_Masonry_Work',
        '07_Stone_Work',
        '08_Cladding_Work',
        '09_Wood_and_PVC_Work',
        '10_Steel_Work',
        '11_Flooring',
        '12_Roofing'
    ]
    
    for key in trade_keys:
        build_standard_trade(wb, configs[key], styles)
        
    out_path = 'CPWD_DAR_2019_Custom_Rate_Analysis_Workbook_Vol_1.xlsx'
    print(f"Saving complete workbook to {out_path}...")
    try:
        wb.save(out_path)
    except PermissionError:
        alt_path = 'CPWD_DAR_2019_Custom_Rate_Analysis_Workbook_Vol_1_Latest.xlsx'
        print(f"WARNING: '{out_path}' is currently open in Excel. Saving to '{alt_path}' instead.")
        wb.save(alt_path)
        out_path = alt_path
        
    elapsed = time.time() - t0
    file_size_mb = os.path.getsize(out_path) / (1024 * 1024)
    print(f"SUCCESS! Generated {len(wb.sheetnames)} sheets in {elapsed:.2f}s ({file_size_mb:.2f} MB)")
    print(f"Sheet names: {wb.sheetnames}")
    print("================================================================")

if __name__ == '__main__':
    generate_full_workbook()
