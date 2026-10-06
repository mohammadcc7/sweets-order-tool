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
                qty = valid_vals[-1]
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
    file_name = os.path.basename(input_path)
    store_name = os.path.splitext(file_name)[0]
    
    wb_src = openpyxl.load_workbook(input_path, data_only=True)
    ws_src = wb_src.active
    rows = list(ws_src.iter_rows(values_only=True))
    if not rows:
        return 0
        
    header_row_idx = 0
    col_item, col_qty, col_wh = 0, 1, -1
    
    for idx, row in enumerate(rows[:5]):
        row_strs = [str(v or '').strip() for v in row]
        for c_idx, val in enumerate(row_strs):
            if val in ['المادة', 'اسم المادة']:
                col_item = c_idx
                header_row_idx = idx
            elif val in ['الكمية', 'المطلوب', 'تجهيز']:
                col_qty = c_idx
            elif 'مستودع' in val or val in ['رقم المستودع', 'المستودع']:
                col_wh = c_idx

    filtered_data_wh1 = []
    filtered_data_wh3 = []
    
    for row in rows[header_row_idx + 1:]:
        if len(row) <= max(col_item, col_qty, col_wh if col_wh != -1 else 0):
            continue
        item_val = row[col_item]
        qty_val = row[col_qty]
        wh_val = row[col_wh] if col_wh != -1 else 1
        
        if item_val is None or str(item_val).strip() == '' or str(item_val) == 'المادة':
            continue
        if qty_val is None or str(qty_val).strip() in ['', '0', 'None', '0.0']:
            continue
            
        try:
            wh_num = int(float(wh_val))
        except (ValueError, TypeError):
            wh_num = 1
            
        if wh_num == 2:
            continue
            
        try:
            f_qty = float(qty_val)
            formatted_qty = int(f_qty) if f_qty.is_integer() else round(f_qty, 2)
        except:
            formatted_qty = qty_val
            
        unit_val = row[2] if len(row) > 2 and row[2] is not None else ''
        item_record = [str(item_val), formatted_qty, str(unit_val)]
        
        if wh_num == 1:
            filtered_data_wh1.append(item_record)
        elif wh_num == 3:
            filtered_data_wh3.append(item_record)
            
    processed_files_count = 0
    
    def save_to_warehouse_folder(sub_data, wh_title_suffix, folder_name):
        if not sub_data:
            return None
        target_dir = os.path.join(base_output_dir, folder_name)
        os.makedirs(target_dir, exist_ok=True)
        
        wb_out = openpyxl.Workbook()
        ws_out = wb_out.active
        report_title = f"{store_name} - {wh_title_suffix}"
        ws_out.title = wh_title_suffix
        ws_out.views.sheetView[0].rightToLeft = True
        
        ws_out.merge_cells('A1:C1')
        ws_out['A1'] = report_title
        ws_out.append(['المادة', 'الكمية', 'الوحدة'])
        
        for r_data in sub_data:
            ws_out.append(r_data)
            
        font_title = Font(name='Arial', size=14, bold=True)
        font_header = Font(name='Arial', size=12, bold=True, color='FFFFFF')
        font_body = Font(name='Arial', size=11, bold=True)
        fill_header = PatternFill(start_color='000000', end_color='000000', fill_type='solid')
        fill_zebra = PatternFill(start_color='F2F2F2', end_color='F2F2F2', fill_type='solid')
        thin_side = Side(style='thin', color='000000')
        border_all = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
        center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
        
        ws_out.row_dimensions[1].height = 30
        ws_out['A1'].font = font_title
        ws_out['A1'].alignment = center_align
        
        ws_out.row_dimensions[2].height = 25
        for c in range(1, 4):
            cell = ws_out.cell(row=2, column=c)
            cell.font = font_header
            cell.fill = fill_header
            cell.alignment = center_align
            cell.border = border_all
            
        for r in range(3, ws_out.max_row + 1):
            ws_out.row_dimensions[r].height = 24
            for c in range(1, 4):
                cell = ws_out.cell(row=r, column=c)
                cell.font = font_body
                cell.alignment = center_align
                cell.border = border_all
                if r % 2 == 1:
                    cell.fill = fill_zebra
                    
        ws_out.column_dimensions['A'].width = 22
        ws_out.column_dimensions['B'].width = 10
        ws_out.column_dimensions['C'].width = 10
        
        ws_out.page_margins.left = ws_out.page_margins.right = 0.0
        ws_out.page_margins.top = ws_out.page_margins.bottom = 0.01
        ws_out.page_setup.orientation = ws_out.ORIENTATION_PORTRAIT
        ws_out.sheet_properties.pageSetUpPr.fitToPage = True
        ws_out.page_setup.fitToWidth = 1
        ws_out.page_setup.fitToHeight = 1
        
        out_file_path = os.path.join(target_dir, f"{store_name}_{wh_title_suffix}.xlsx")
        wb_out.save(out_file_path)
        return out_file_path

    path_w1 = save_to_warehouse_folder(filtered_data_wh1, "مستودع المواد الأولية", "مستودع المواد الأولية")
    if path_w1:
        processed_files_count += 1
        if direct_print:
            print_excel_file(path_w1)
            
    path_w3 = save_to_warehouse_folder(filtered_data_wh3, "مستودع الجاهز", "مستودع الجاهز")
    if path_w3:
        processed_files_count += 1
        if direct_print:
            print_excel_file(path_w3)
            
    return processed_files_count

def print_excel_file(file_path):
    try:
        if platform.system() == 'Windows':
            os.startfile(file_path, 'print')
            return True
        else:
            subprocess.run(['lpr', file_path], check=True)
            return True
    except Exception as e:
        return str(e)

class App:
    def __init__(self, root):
        self.root = root
        self.root.title('منسق طلبات الأمين الحراري الشامل (مع محول TXT)')
        self.root.geometry('720x800')
        self.root.resizable(False, False)
        
        self.current_files = []
        self.warehouse_files = []
        self.txt_converter_files = []

        lbl_title = tk.Label(root, text='منسق ملفات الأمين للطابعة الحرارية (8سم)', font=('Arial', 13, 'bold'))
        lbl_title.pack(pady=5)

        frame_controls = tk.LabelFrame(root, text='إعدادات ومعالجة التقارير الحرارية', font=('Arial', 10, 'bold'))
        frame_controls.pack(fill='x', padx=15, pady=3)

        frame_type = tk.Frame(frame_controls)
        frame_type.pack(pady=5, fill='x', padx=10)
        tk.Label(frame_type, text='نوع التقرير:', font=('Arial', 10, 'bold')).pack(side='right', padx=5)
        
        report_types = ['ورقة الفرن', 'هرايس بانواعها', 'مستودع الجاهز القديم', 'محلاية + خمس مواد']
        self.file_type_var = tk.StringVar(value='ورقة الفرن')
        self.combo_type = ttk.Combobox(frame_type, textvariable=self.file_type_var, values=report_types, state='readonly', font=('Arial', 10, 'bold'), width=22)
        self.combo_type.pack(side='right', padx=5)

        self.chk_var = tk.BooleanVar(value=True)
        chk = tk.Checkbutton(frame_controls, text='حذف الصفوف فارغة/صفرية الكمية تلقائياً', variable=self.chk_var, font=('Arial', 9))
        chk.pack(anchor='e', padx=15, pady=2)

        self.direct_print_var = tk.BooleanVar(value=False)
        chk_print = tk.Checkbutton(frame_controls, text='إرسال الملفات للطابعة الحرارية مباشرة بعد إنتاجها', variable=self.direct_print_var, font=('Arial', 9, 'bold'), fg='#d9534f')
        chk_print.pack(anchor='e', padx=15, pady=2)

        frame_btns = tk.Frame(root)
        frame_btns.pack(pady=8)
        
        btn_select = tk.Button(frame_btns, text='📂 اختر ملفات المحلات (القياسية)', font=('Arial', 10, 'bold'), bg='#007bff', fg='white', padx=8, pady=5, command=self.load_files_dialog)
        btn_select.pack(side='left', padx=5)

        btn_warehouse = tk.Button(frame_btns, text='🏢 فرز المستودعات (مواد أولية وجاهز)', font=('Arial', 10, 'bold'), bg='#fd7e14', fg='white', padx=8, pady=5, command=self.load_warehouse_files_dialog)
        btn_warehouse.pack(side='left', padx=5)

        frame_preview = tk.LabelFrame(root, text='معاينة الملفات القياسية المختارة', font=('Arial', 9, 'bold'))
        frame_preview.pack(fill='both', expand=True, padx=15, pady=3)

        scroll_x = ttk.Scrollbar(frame_preview, orient='horizontal')
        scroll_y = ttk.Scrollbar(frame_preview, orient='vertical')
        self.tree = ttk.Treeview(frame_preview, show='headings', height=4, xscrollcommand=scroll_x.set, yscrollcommand=scroll_y.set)
        scroll_x.config(command=self.tree.xview)
        scroll_y.config(command=self.tree.yview)
        scroll_x.pack(side='bottom', fill='x')
        scroll_y.pack(side='left', fill='y')
        self.tree.pack(fill='both', expand=True, padx=5, pady=5)

        self.btn_process = tk.Button(root, text='⚡ معالجة واستخراج ملفات المحلات القياسية', font=('Arial', 11, 'bold'), bg='#28a745', fg='white', padx=15, pady=6, state='disabled', command=self.process_files)
        self.btn_process.pack(pady=3)

        frame_txt_tool = tk.LabelFrame(root, text='أداة محول ملفات Excel إلى TXT (Unicode)', font=('Arial', 10, 'bold'), fg='#0056b3')
        frame_txt_tool.pack(fill='x', padx=15, pady=5)

        self.lbl_txt_status = tk.Label(frame_txt_tool, text='لم يتم اختيار أي ملف للتحويل', font=('Arial', 9), fg='gray')
        self.lbl_txt_status.pack(pady=2)

        frame_txt_btns = tk.Frame(frame_txt_tool)
        frame_txt_btns.pack(pady=5)

        btn_txt_select = tk.Button(frame_txt_btns, text='📂 (واحدة أو أكثر) Excel اختر ملفات', font=('Arial', 10, 'bold'), bg='#28a745', fg='white', padx=8, pady=4, command=self.load_txt_files_dialog)
        btn_txt_select.pack(side='left', padx=5)

        self.btn_txt_convert = tk.Button(frame_txt_btns, text='🔄 TXT (Unicode) تحويل الملفات إلى', font=('Arial', 10, 'bold'), bg='#007bff', fg='white', padx=8, pady=4, state='disabled', command=self.convert_to_txt_files)
        self.btn_txt_convert.pack(side='left', padx=5)

    def load_files_dialog(self):
        file_paths = filedialog.askopenfilenames(filetypes=[('Excel Files', '*.xlsx *.xls')])
        if not file_paths:
            return
        self.current_files = list(file_paths)
        self.preview_files(self.current_files)
        self.btn_process.config(state='normal')

    def load_warehouse_files_dialog(self):
        file_paths = filedialog.askopenfilenames(filetypes=[('Excel Files', '*.xlsx *.xls')])
        if not file_paths:
            return
        
        desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')
        direct_print = self.direct_print_var.get()
        processed_count = 0
        
        for file_path in file_paths:
            try:
                count = process_single_warehouse_file(file_path, desktop_path, direct_print)
                processed_count += count
            except Exception as e:
                continue
                
        msg = f'تمت معالجة وفرز ملفات المستودعات بنجاح تام ({processed_count} ملف ناتج)!\nوموجودة في مجلداتها على سطح المكتب.'
        if direct_print:
            msg += '\nوتم إرسالها للطباعة المباشرة.'
        messagebox.showinfo('نجاح فرز المستودعات', msg)

    def preview_files(self, file_paths):
        for item in self.tree.get_children():
            self.tree.delete(item)
        try:
            first_file = file_paths[0]
            wb = openpyxl.load_workbook(first_file, data_only=True)
            ws = wb.active
            rows = list(ws.iter_rows(values_only=True))
            if not rows:
                return
            header_row = rows[0]
            cols = [f'col_{i}' for i in range(len(header_row))]
            self.tree['columns'] = cols
            for i, col_name in enumerate(header_row):
                header_text = str(col_name) if col_name is not None else f"عمود {i + 1}"
                self.tree.heading(cols[i], text=header_text)
                self.tree.column(cols[i], width=100, anchor='center')
            for row in rows[1:6]:
                values = [str(cell) if cell is not None else '' for cell in row]
                self.tree.insert('', 'end', values
