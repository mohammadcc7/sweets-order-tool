import os
import pandas as pd
from tkinter import Tk, Label, Button, filedialog, messagebox
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def process_excel_files():
    file_paths = filedialog.askopenfilenames(
        title="اختر ملفات الاكسل المطلوبة",
        filetypes=[("Excel Files", "*.xlsx *.xls")]
    )
    if not file_paths:
        return

    output_dir = filedialog.askdirectory(title="اختر مجلد حفظ ملفات الإكسل المنسقة")
    if not output_dir:
        return

    success_count = 0
    for file_path in file_paths:
        try:
            xls = pd.ExcelFile(file_path)
            sheet_names = xls.sheet_names
            target_sheet = 'حركة المواد المطلوبة' if 'حركة المواد المطلوبة' in sheet_names else sheet_names[0]
            df_sheet = pd.read_excel(file_path, sheet_name=target_sheet)
            
            sheet_title = str(target_sheet).strip()
            
            wb = Workbook()
            ws = wb.active
            ws.title = sheet_title[:31]
            ws.sheet_view.rightToLeft = True
            
            # العنوان الرئيسي
            ws.merge_cells('A1:C1')
            ws['A1'] = sheet_title
            ws['A1'].font = Font(name='Tahoma', size=13, bold=True, color='FFFFFF')
            ws['A1'].fill = PatternFill(start_color='333333', end_color='333333', fill_type='solid')
            ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
            ws.row_dimensions[1].height = 25
            
            # العنوان الفرعي
            ws.merge_cells('A2:C2')
            ws['A2'] = "الحلويات الشرقية - طلبيات المبيع"
            ws['A2'].font = Font(name='Tahoma', size=10, bold=True, color='333333')
            ws['A2'].fill = PatternFill(start_color='EAEAEA', end_color='EAEAEA', fill_type='solid')
            ws['A2'].alignment = Alignment(horizontal='center', vertical='center')
            ws.row_dimensions[2].height = 20
            
            # رؤوس الأعمدة
            ws.append([])
            headers = ['المادة', 'اسم العميل', 'الكمية المطلوبة']
            ws.append(headers)
            
            header_fill = PatternFill(start_color='D0D0D0', end_color='D0D0D0', fill_type='solid')
            header_font = Font(name='Tahoma', size=11, bold=True)
            thin_border = Border(
                left=Side(style='thin', color='000000'),
                right=Side(style='thin', color='000000'),
                top=Side(style='thin', color='000000'),
                bottom=Side(style='thin', color='000000')
            )
            
            ws.row_dimensions[4].height = 22
            for col_num in range(1, 4):
                cell = ws.cell(row=4, column=col_num)
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = thin_border
            
            # بيانات الصفوف
            if len(df_sheet.columns) >= 3:
                df_data = df_sheet.iloc[1:].copy()
                row_idx = 5
                for _, row in df_data.iterrows():
                    mat = str(row.iloc[0]).strip()
                    cust = str(row.iloc[1]).strip()
                    qty = row.iloc[2] if len(row) > 2 else row.iloc[-1]
                    
                    if pd.notna(mat) and mat != 'nan':
                        ws.append([mat, cust, qty])
                        ws.row_dimensions[row_idx].height = 20
                        
                        for col_num in range(1, 4):
                            c = ws.cell(row=row_idx, column=col_num)
                            c.font = Font(name='Tahoma', size=10)
                            c.alignment = Alignment(horizontal='center', vertical='center')
                            c.border = thin_border
                        row_idx += 1
            
            for col in ws.columns:
                max_len = 0
                col_letter = get_column_letter(col[0].column)
                for cell in col:
                    if cell.row > 2 and cell.value:
                        max_len = max(max_len, len(str(cell.value)))
                ws.column_dimensions[col_letter].width = max(max_len + 5, 18)
                
            safe_filename = sheet_title.replace('/', '-').replace('\\', '-')
            output_excel_path = os.path.join(output_dir, f"{safe_filename}.xlsx")
            wb.save(output_excel_path)
            success_count += 1
            
        except Exception as e:
            print(f"خطأ في معالجة الملف {file_path}: {e}")
            
    messagebox.showinfo("اكتمل التجهيز", f"تم معالجة وتنسيق وتصدير {success_count} ملف إكسل بنجاح وتسميتها بأسماء الأوراق الداخلية!")

root = Tk()
root.title("معالج طلبيات الحلويات")
root.geometry("420x260")
root.config(bg="#f5f5f5")

label = Label(root, text="نظام تنسيق وتصدير ملفات الإكسل", font=("Tahoma", 12, "bold"), bg="#f5f5f5")
label.pack(pady=25)

desc_label = Label(root, text="يقبل عدة ملفات، ينسقها، ويسمي الملف الناتج باسم الورقة الداخلية", font=("Tahoma", 8), bg="#f5f5f5", fg="#555")
desc_label.pack(pady=5)

btn = Button(root, text="اختيار ملفات الإكسل والبدء", command=process_excel_files, font=("Tahoma", 11, "bold"), bg="#1b5e20", fg="white", padx=15, pady=8)
btn.pack(pady=20)

root.mainloop()
