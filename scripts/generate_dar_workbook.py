# CPWD DAR 2019 Volume 1 Custom Rate Analysis Workbook Generator
# Strict MS Excel 2016 Compatibility
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
import json, os

