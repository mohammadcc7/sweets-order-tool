import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
import subprocess
import platform

def extract_clean_store_name(filename):
    base = os.path.splitext(os.path.basename(filename))[0]
    base_lower = base.lower()
    if 'ضاحية قدسية' in base or 'مجمع ضاحية قدسية' in base:
        return 'ضاحية قدسية'
    if 'مشروع دمر' in base:
        return 'مشروع دمر'
    if 'دمر البلد' in base:
        return 'دمر البلد'
    if 'شيخ سعد' in base or 'الشيخ سعد' in base:
        return 'شيخ سعد'

    stop_words = ['طلبات', 'طلب', 'طلبية', 'مجمع']
    words = base.split()
    clean_words = [w for w in words if w not in stop_words]
    
    cleaned = ' '.join(clean_words).strip()
    if not cleaned:
        cleaned = base
    return cleaned

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
        item_name, qty = r[0], r[1]
        if item_name in ('اسم المادة', None) or str(item_name).strip() == '':
            continue
        try:
            val_float = float(qty)
            formatted_qty = int(val_float) if val_float.is_integer() else round(val_float, 2)
        except (ValueError, TypeError):
            formatted_qty = qty
        ws_out.append([item_name, formatted_qty])
    
    font_title = Font(name='Arial', size=14, bold=True)
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
    wb_out.save(output_path)

def format_warehouse_sheet_three_columns(rows_data, output_path, store_name, subtitle_type):
    wb_out = openpyxl.Workbook()
    ws_out = wb_out.active
    ws_out.title = subtitle_type
    ws_out.views.sheetView[0].rightToLeft = True
    
    ws_out.merge_cells('A1:C1')
    ws_out.merge_cells('A2:C2')
    
    ws_out['A1'] = f"محل {store_name}"
    ws_out['A2'] = subtitle_type
    
    ws_out['A3'] = 'اسم المادة'
    ws_out['B3'] = 'الوحدة'
    ws_out['C3'] = 'المطلوب'
    
    for r in rows_data:
        item_name = r[0]
        unit_name = r[1]
        qty = r[2]
        try:
            val_float = float(qty)
            formatted_qty = int(val_float) if val_float.is_integer() else round(val_float, 2)
        except (ValueError, TypeError):
            formatted_qty = qty
        ws_out.append([item_name, unit_name, formatted_qty])
        
    font_title1 = Font(name='Arial', size=15, bold=True)
    font_title2 = Font(name='Arial', size=14, bold=True)
    font_header = Font(name='Arial', size=14, bold=True, color='FFFFFF')
    font_body = Font(name='Arial', size=13, bold=True)
    fill_header = PatternFill(start_color='000000', end_color='000000', fill_type='solid')
    fill_zebra = PatternFill(start_color='F2F2F2', end_color='F2F2F2', fill_type='solid')
    thin_side = Side(style='thin', color='000000')
    border_all = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    
    ws_out.row_dimensions[1].height = 25
    ws_out['A1'].font = font_title1
    ws_out['A1'].alignment = center_align
    
    ws_out.row_dimensions[2].height = 25
    ws_out['A2'].font = font_title2
    ws_out['A2'].alignment = center_align
    
    ws_out.row_dimensions[3].height = 28
    for c in range(1, 4):
        cell = ws_out.cell(row=3, column=c)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = center_align
        cell.border = border_all
        
    for r in range(4, ws_out.max_row + 1):
        ws_out.row_dimensions[r].height = 26
        for c in range(1, 4):
            cell = ws_out.cell(row=r, column=c)
            cell.font = font_body
            cell.alignment = center_align
            cell.border = border_all
            if r % 2 == 0:
                cell.fill = fill_zebra
                
    ws_out.column_dimensions['A'].width = 28
    ws_out.column_dimensions['B'].width = 12
    ws_out.column_dimensions['C'].width = 15
    wb_out.save(output_path)

class App:
    def __init__(self, root):
        self.root = root
        self.root.title('منسق طلبات الأمين الحراري الشامل (مع التنسيق الذكي للأعمدة)')
        self.root.geometry('720x780')
        self.root.resizable(False, False)
        
        self.current_files = []
        self.txt_converter_files = []

        lbl_title = tk.Label(root, text='منسق ملفات الأمين للطابعة الحرارية (8سم)', font=('Arial', 13, 'bold'))
        lbl_title.pack(pady=5)

        frame_controls = tk.LabelFrame(root, text='إعدادات ومعالجة التقارير الحرارية', font=('Arial', 10, 'bold'))
        frame_controls.pack(fill='x', padx=15, pady=3)

        frame_type = tk.Frame(frame_controls)
        frame_type.pack(pady=5, fill='x', padx=10)
        tk.Label(frame_type, text='نوع التقرير:', font=('Arial', 10, 'bold')).pack(side='right', padx=5)
        self.file_type_var = tk.StringVar(value='تعرّف تلقائي')
        self.combo_type = ttk.Combobox(frame_type, textvariable=self.file_type_var, values=['تعرّف تلقائي', 'محلاية + خمس مواد', 'ورقة الفرن', 'هرايس بانواعها', 'مستودع الجاهز'], state='readonly', font=('Arial', 10, 'bold'), width=20)
        self.combo_type.pack(side='right', padx=5)

        self.chk_var = tk.BooleanVar(value=True)
        chk = tk.Checkbutton(frame_controls, text='حذف الصفوف فارغة/صفرية الكمية تلقائياً', variable=self.chk_var, font=('Arial', 9))
        chk.pack(anchor='e', padx=15, pady=2)

        self.direct_print_var = tk.BooleanVar(value=False)
        chk_print = tk.Checkbutton(frame_controls, text='إرسال الملفات للطابعة الحرارية مباشرة بعد إنتاجها', variable=self.direct_print_var, font=('Arial', 9, 'bold'), fg='#d9534f')
        chk_print.pack(anchor='e', padx=15, pady=2)

        frame_btns = tk.Frame(root)
        frame_btns.pack(pady=5)
        
        btn_select = tk.Button(frame_btns, text='📂 اختر ملف أو عدة ملفات Excel', font=('Arial', 10, 'bold'), bg='#007bff', fg='white', padx=10, pady=5, command=self.load_files_dialog)
        btn_select.pack(side='left', padx=5)

        btn_warehouses = tk.Button(frame_btns, text='🏢 فرز وتوزيع ملفات المستودعات (1 و 3)', font=('Arial', 10, 'bold'), bg='#6f42c1', fg='white', padx=10, pady=5, command=self.process_warehouses_files)
        btn_warehouses.pack(side='left', padx=5)

        frame_preview = tk.LabelFrame(root, text='معاينة الملفات المختارة', font=('Arial', 9, 'bold'))
        frame_preview.pack(fill='both', expand=True, padx=15, pady=3)

        scroll_x = ttk.Scrollbar(frame_preview, orient='horizontal')
        scroll_y = ttk.Scrollbar(frame_preview, orient='vertical')
        self.tree = ttk.Treeview(frame_preview, show='headings', height=5, xscrollcommand=scroll_x.set, yscrollcommand=scroll_y.set)
        scroll_x.config(command=self.tree.xview)
        scroll_y.config(command=self.tree.yview)
        scroll_x.pack(side='bottom', fill='x')
        scroll_y.pack(side='left', fill='y')
        self.tree.pack(fill='both', expand=True, padx=5, pady=5)

        self.btn_process = tk.Button(root, text='⚡ معالجة واستخراج جميع الملفات إلى مجلد سطح المكتب', font=('Arial', 11, 'bold'), bg='#28a745', fg='white', padx=15, pady=7, state='disabled', command=self.process_files)
        self.btn_process.pack(pady=5)

        frame_txt_tool = tk.LabelFrame(root, text='أداة محول ملفات Excel إلى TXT (Unicode)', font=('Arial', 10, 'bold'), fg='#0056b3')
        frame_txt_tool.pack(fill='x', padx=15, pady=5)

        self.lbl_txt_status = tk.Label(frame_txt_tool, text='لم يتم اختيار أي ملف للتحويل', font=('Arial', 9), fg='gray')
        self.lbl_txt_status.pack(pady=2)

        frame_txt_btns = tk.Frame(frame_txt_tool)
        frame_txt_btns.pack(pady=5)

        btn_txt_select = tk.Button(frame_txt_btns, text='📂 (واحدة أو أكثر) Excel اختر ملفات', font=('Arial', 10, 'bold'), bg='#28a745', fg='white', padx=10, pady=5, command=self.load_txt_files_dialog)
        btn_txt_select.pack(side='left', padx=5)

        self.btn_txt_convert = tk.Button(frame_txt_btns, text='🔄 TXT (Unicode) تحويل الملفات إلى', font=('Arial', 10, 'bold'), bg='#007bff', fg='white', padx=10, pady=5, state='disabled', command=self.convert_to_txt_files)
        self.btn_txt_convert.pack(side='left', padx=5)

    def load_files_dialog(self):
        file_paths = filedialog.askopenfilenames(filetypes=[('Excel Files', '*.xlsx *.xls')])
        if not file_paths:
            return
        self.current_files = file_paths
        self.preview_files(file_paths)
        self.btn_process.config(state='normal')

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
                header_text = str(col_name) if col_name is not None else f'عمود {i + 1}'
                self.tree.heading(f'col_{i}', text=header_text)
                self.tree.column(f'col_{i}', width=100, anchor='center')
            for row in rows[1:6]:
                self.tree.insert('', 'end', values=[str(v) if v is not None else '' for v in row])
        except Exception as e:
            pass

    def process_files(self):
        pass

    def process_warehouses_files(self):
        file_paths = filedialog.askopenfilenames(filetypes=[('Excel Files', '*.xlsx *.xls')])
        if not file_paths:
            return
        
        desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')
        main_output_dir = os.path.join(desktop_path, 'فرز_المستودعات')
        raw_dir = os.path.join(main_output_dir, 'مستودع المواد الأولية')
        ready_dir = os.path.join(main_output_dir, 'مستودع الجاهز')
        
        os.makedirs(raw_dir, exist_ok=True)
        os.makedirs(ready_dir, exist_ok=True)
        
        processed_count = 0
        try:
            for file_path in file_paths:
                wb_src = openpyxl.load_workbook(file_path, data_only=True)
                ws_src = wb_src.active
                
                clean_store = extract_clean_store_name(file_path)
                
                raw_rows = []
                ready_rows = []
                
                for row in ws_src.iter_rows(values_only=True):
                    if not any(row):
                        continue
                    
                    valid_vals = [v for v in row if v is not None]
                    if len(valid_vals) < 2:
                        continue
                        
                    item_name = None
                    qty_val = None
                    unit_val = ''
                    wh_val = None
                    
                    # الفحص الذكي للقيم في الصف حسب نمط ملفات الأمين:
                    # الاسم غالباً نص في البداية (العنصر الأول أو الثاني)
                    # الكمية هي رقم (float أو int)
                    # رقم المستودع غالباً يكون 1 أو 3 في آخر الأعمدة أو قبل الأخير
                    
                    # استخراج العناصر الأساسية بمرونة تامة:
                    item_name = row[0] # العمود الأول غالباً المادة
                    if item_name is None or str(item_name).strip() == '':
                        continue
                        
                    # البحث عن رقم المستودع (غالباً القيمة الأخيرة أو قبل الأخيرة التي تساوي 1 أو 3)
                    for v in reversed(row):
                        if v is not None:
                            try:
                                v_int = int(float(v))
                                if v_int in [1, 3]:
                                    wh_val = v_int
                                    break
                            except:
                                pass
                                
                    if wh_val is None:
                        continue # إذا لم يوجد رقم مستودع (1 أو 3) نتخطي الصف
                        
                    # البحث عن الكمية (رقم غير رقم المستودع)
                    possible_qtys = []
                    possible_units = []
                    for v in row[1:]:
                        if v is not None and v != wh_val:
                            try:
                                num = float(v)
                                possible_qtys.append(num)
                            except:
                                if isinstance(v, str) and v.strip() != '':
                                    possible_units.append(v.strip())
                                    
                    if not possible_qtys:
                        continue
                        
                    qty_val = possible_qtys[0] # أول رقم يتم اعتباره الكمية
                    if qty_val == 0:
                        continue # حذف الصفوف صفرية الكمية
                        
                    if possible_units:
                        unit_val = possible_units[0] # أخذ الوحدة الأولى المكتشفة
                        
                    item_data = [item_name, unit_val, qty_val]
                    
                    if wh_val == 1:
                        raw_rows.append(item_data)
                    elif wh_val == 3:
                        ready_rows.append(item_data)
                
                if raw_rows:
                    format_warehouse_sheet_three_columns(
                        raw_rows, 
                        os.path.join(raw_dir, f'{clean_store}_أولية.xlsx'), 
                        clean_store, 
                        'مواد أولية'
                    )
                
                if ready_rows:
                    format_warehouse_sheet_three_columns(
                        ready_rows, 
                        os.path.join(ready_dir, f'{clean_store}_جاهز.xlsx'), 
                        clean_store, 
                        'الحلو العربي'
                    )
                
                processed_count += 1
                
            messagebox.showinfo('نجاح الفرز', f'تمت معالجة وفرز {processed_count} ملف بنجاح مع إضافة عمود الوحدة وتصفية الصفرية!\nالمجلدات موجودة على سطح المكتب داخل (فرز_المستودعات).')
        except Exception as e:
            messagebox.showerror('خطأ', f'حدث خطأ أثناء معالجة المستودعات:\n{str(e)}')

    def load_txt_files_dialog(self):
        file_paths = filedialog.askopenfilenames(filetypes=[('Excel Files', '*.xlsx *.xls')])
        if not file_paths:
            return
        self.txt_converter_files = list(file_paths)
        self.lbl_txt_status.config(text=f'تم اختيار {len(self.txt_converter_files)} ملف للتحويل', fg='green')
        self.btn_txt_convert.config(state='normal')

    def convert_to_txt_files(self):
        if not self.txt_converter_files:
            return
        desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')
        output_dir = os.path.join(desktop_path, 'ملفات_TXT_الأمين')
        os.makedirs(output_dir, exist_ok=True)
        
        success_count = 0
        try:
            for file_path in self.txt_converter_files:
                wb = openpyxl.load_workbook(file_path, data_only=True)
                ws = wb.active
                
                base_name = os.path.splitext(os.path.basename(file_path))[0]
                out_txt_path = os.path.join(output_dir, f'{base_name}.txt')
                
                with open(out_txt_path, 'w', encoding='utf-16') as f:
                    for row in ws.iter_rows(values_only=True):
                        if any(row):
                            row_vals = []
                            for v in row:
                                if v is None:
                                    row_vals.append('')
                                elif isinstance(v, float) and v.is_integer():
                                    row_vals.append(str(int(v)))
                                else:
                                    row_vals.append(str(v))
                            while row_vals and row_vals[-1] == '':
                                row_vals.pop()
                            if row_vals:
                                f.write('\t'.join(row_vals) + '\n')
                success_count += 1
                
            messagebox.showinfo('نجاح التحويل', f'تم تحويل {success_count} ملف بنجاح!\nموجودة الآن في مجلد (ملفات_TXT_الأمين) على سطح المكتب.')
            self.lbl_txt_status.config(text='لم يتم اختيار أي ملف للتحويل', fg='gray')
            self.txt_converter_files = []
            self.btn_txt_convert.config(state='disabled')
        except Exception as e:
            messagebox.showerror('خطأ في التحويل', f'حدث خطأ أثناء معالجة الملفات:\n{str(e)}')

if __name__ == '__main__':
    root = tk.Tk()
    app = App(root)
    root.mainloop()
