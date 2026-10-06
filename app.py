# Decompiled with PyLingual (https://pylingual.io)
# Internal filename: 'main.py'
# Bytecode version: 3.10.b1 (3439)
# Source timestamp: 1970-01-01 00:00:00 UTC (0)

import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
import subprocess
import platform

def format_sales_orders_custom(input_path, output_path, selected_items, report_title, remove_empty=True):
    wb_src = openpyxl.load_workbook(input_path, data_only=True)
    ws_src = wb_src.active
    header_row = 1
    for r in range(1, min(10, ws_src.max_row + 1)):
        row_vals = [str(ws_src.cell(row=r, column=c).value or '').strip() for c in range(1, ws_src.max_column + 1)]
        if any((k in row_vals for k in ['المادة', 'تجهيز', 'اسم العميل'])):
            header_row = r
            break
    col_item, col_client, col_qty = (3, 5, 6)
    for c in range(1, ws_src.max_column + 1):
        val = str(ws_src.cell(row=header_row, column=c).value or '').strip()
        if val in ['المادة', 'اسم المادة']:
            col_item = c
        elif val in ['اسم العميل', 'العميل', 'المحل']:
            col_client = c
        elif val in ['تجهيز', 'الكمية', 'كمية']:
            col_qty = c
    wb_out = openpyxl.Workbook()
    ws_out = wb_out.active
    ws_out.title = report_title
    ws_out.views.sheetView[0].rightToLeft = True
    ws_out.append(['المادة', 'اسم العميل', 'الكمية'])
    totals_per_item = {}
    for r in range(header_row + 1, ws_src.max_row + 1):
        qty_val = ws_src.cell(row=r, column=col_qty).value
        if remove_empty and (qty_val is None or str(qty_val).strip() in ['', '0', 'None', '0.0']):
            continue
        item_val = str(ws_src.cell(row=r, column=col_item).value or '').strip()
        client_val = ws_src.cell(row=r, column=col_client).value
        if selected_items and item_val not in selected_items:
            continue
        try:
            num_qty = float(qty_val)
        except (ValueError, TypeError):
            num_qty = 0.0
        if isinstance(qty_val, float) and qty_val.is_integer():
            qty_val = int(qty_val)
        else:
            if isinstance(qty_val, (int, float)):
                pass
            else:
                try:
                    qty_val = int(num_qty) if num_qty.is_integer() else num_qty
                except:
                    pass
        totals_per_item[item_val] = totals_per_item.get(item_val, 0) + num_qty
        ws_out.append([item_val, client_val, qty_val])
    ws_out.append([])
    total_rows_start = ws_out.max_row + 1
    for item, total in totals_per_item.items():
        total_formatted = int(total) if total.is_integer() else round(total, 2)
        ws_out.append([f'مجموع {item}', 'الإجمالي', total_formatted])
    font_header = Font(name='Arial', size=12, bold=True, color='FFFFFF')
    font_body = Font(name='Arial', size=11, bold=True, color='000000')
    font_total = Font(name='Arial', size=12, bold=True, color='000000')
    fill_header = PatternFill(start_color='000000', end_color='000000', fill_type='solid')
    fill_zebra = PatternFill(start_color='F2F2F2', end_color='F2F2F2', fill_type='solid')
    fill_total = PatternFill(start_color='E2E2E2', end_color='E2E2E2', fill_type='solid')
    thin_side = Side(style='thin', color='000000')
    double_side = Side(style='double', color='000000')
    border_all = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    border_total = Border(left=thin_side, right=thin_side, top=thin_side, bottom=double_side)
    center_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws_out.row_dimensions[1].height = 28
    for r in range(1, ws_out.max_row + 1):
        if r > 1:
            ws_out.row_dimensions[r].height = 25
        is_header = r == 1
        is_total_row = r >= total_rows_start
        if not is_header and ws_out.cell(row=r, column=1).value is None:
            ws_out.row_dimensions[r].height = 10
            continue
        for c in range(1, 4):
            cell = ws_out.cell(row=r, column=c)
            cell.alignment = center_alignment
            if is_header:
                cell.font = font_header
                cell.border = border_all
                cell.fill = fill_header
            else:
                if is_total_row:
                    cell.font = font_total
                    cell.border = border_total
                    cell.fill = fill_total
                else:
                    cell.font = font_body
                    cell.border = border_all
                    if r % 2 == 0:
                        cell.fill = fill_zebra
    ws_out.column_dimensions['A'].width = 18
    ws_out.column_dimensions['B'].width = 18
    ws_out.column_dimensions['C'].width = 8
    ws_out.page_margins.left = ws_out.page_margins.right = 0.0
    ws_out.page_margins.top = ws_out.page_margins.bottom = 0.01
    ws_out.page_setup.orientation = ws_out.ORIENTATION_PORTRAIT
    ws_out.sheet_properties.pageSetUpPr.fitToPage = True
    ws_out.page_setup.fitToWidth = 1
    ws_out.page_setup.fitToHeight = 1
    wb_out.save(output_path)

def format_standard_two_column_sheet(input_path, output_path, report_title):
    wb_src = openpyxl.load_workbook(input_path, data_only=True)
    ws_src = wb_src.active
    data = []
    for row in ws_src.iter_rows(values_only=True):
        if any(row):
            valid_vals = [v for v in row if v is not None and str(v).strip() != '']
            if len(valid_vals) >= 2:
                item_name = valid_vals[0]
                qty = valid_vals[(-1)]
                data.append([item_name, qty])
    wb_out = openpyxl.Workbook()
    ws_out = wb_out.active
    ws_out.title = report_title
    ws_out.views.sheetView[0].rightToLeft = True
    ws_out.merge_cells('A1:B1')
    ws_out['A1'] = report_title
    ws_out['A2'] = 'اسم المادة'
    ws_out['B2'] = 'المطلوب'
    for r in data:
        item_name, qty = (r[0], r[1])
        if item_name in ('اسم المادة', None) or str(item_name).strip() == '':
            continue
        else:
            try:
                val_float = float(qty)
                formatted_qty = int(val_float) if val_float.is_integer() else round(val_float, 2)
            except (ValueError, TypeError):
                formatted_qty = qty
            ws_out.append([item_name, formatted_qty])
    font_title = Font(name='Arial', size=16, bold=True)
    font_header = Font(name='Arial', size=14, bold=True, color='FFFFFF')
    font_body = Font(name='Arial', size=13, bold=True)
    fill_header = PatternFill(start_color='000000', end_color='000000', fill_type='solid')
    fill_zebra = PatternFill(start_color='F2F2F2', end_color='F2F2F2', fill_type='solid')
    thin_side = Side(style='thin', color='000000')
    border_all = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws_out.row_dimensions[1].height = 32
    ws_out['A1'].font = font_title
    ws_out['A1'].alignment = center_align
    ws_out.row_dimensions[2].height = 28
    for c in range(1, 3):
        cell = ws_out.cell(row=2, column=c)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = center_align
        cell.border = border_all
    for r in range(3, ws_out.max_row + 1):
        ws_out.row_dimensions[r].height = 26
        for c in range(1, 3):
            cell = ws_out.cell(row=r, column=c)
            cell.font = font_body
            cell.alignment = center_align
            cell.border = border_all
            if r % 2 == 1:
                cell.fill = fill_zebra
    ws_out.column_dimensions['A'].width = 28
    ws_out.column_dimensions['B'].width = 15
    ws_out.page_margins.left = ws_out.page_margins.right = 0.0
    ws_out.page_margins.top = ws_out.page_margins.bottom = 0.01
    ws_out.page_setup.orientation = ws_out.ORIENTATION_PORTRAIT
    ws_out.sheet_properties.pageSetUpPr.fitToPage = True
    ws_out.page_setup.fitToWidth = 1
    ws_out.page_setup.fitToHeight = 1
    wb_out.save(output_path)

def process_single_warehouse_file(input_path, base_output_dir, direct_print=False):
    """معالجة ملف مستودعات مستقل (لفزر المستودع 1 و 3)"""
    file_name = os.path.basename(input_path)
    store_name = os.path.splitext(file_name)[0]
    
    wb_src = openpyxl.load_workbook(input_path, data_only=True)
    ws_src = wb_src.active
    rows = list(ws_src.iter_rows(values_only=True))
    if not
