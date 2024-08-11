import os
import git
import tkinter as tk
from tkinter import filedialog, messagebox
from openpyxl import load_workbook
from tkinter import ttk
from datetime import datetime
from collections import Counter
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

class ExcelApp:
    def __init__(self, root, repo_path):
        self.root = root
        self.root.title("برنامج الورشه تم التطوير بواسطة حازم ايمن محمود")
        
        # تعيين أبعاد الشاشة ولون الخلفية مع مؤثرات بصرية
        self.root.geometry("500x600")
        self.root.configure(bg='#e6e6fa')
        
        # إضافة تدرج لون كخلفية للشاشة
        self.canvas = tk.Canvas(self.root, width=500, height=600)
        self.canvas.pack(fill="both", expand=True)
        
        self.gradient_background(self.canvas, '#e6e6fa', '#b0c4de')  # مؤثر التدرج
        
        self.repo_path = repo_path
        self.file_path = ""
        
        # تحقق من التحديثات عند بدء التشغيل
        self.update_repo()
        
        # زر لاختيار ملف Excel
        self.select_button = tk.Button(self.canvas, text="اختر ملف Excel", command=self.select_file, bg='#b0c4de')
        self.select_button.pack(pady=10)
        
        # مربع نص لإدخال الكود
        self.code_label = tk.Label(self.canvas, text="أدخل الكود:", bg='#e6e6fa')
        self.code_label.pack()
        self.code_entry = tk.Entry(self.canvas)
        self.code_entry.pack(pady=5)
        self.code_entry.bind("<Return>", self.add_data)  # إضافة البيانات عند الضغط على Enter
        
        # قائمة منسدلة لاختيار اسم الفني
        self.technician_label = tk.Label(self.canvas, text="اختر اسم الفني:", bg='#e6e6fa')
        self.technician_label.pack(pady=1)
        self.technician_combo = ttk.Combobox(self.canvas, values=[
            "محمد عادل", "محمد الزغبى", "حسنى ضاحى", 
            "حسام محمد ابراهيم", "احمد عشرى", "محمد سمير مصطفى",
            "احمد عزازى", "كريم اشرف", "محمد حسن عامر", 
            "عبد الرحمن هلال", "عبد الرحمن"
        ])
        self.technician_combo.pack(pady=5)
        self.technician_combo.bind("<<ComboboxSelected>>", self.show_technician_details)  # عرض التفاصيل عند اختيار الفني
        
        # Checkboxes
        self.checkbox1_var = tk.BooleanVar()
        self.checkbox2_var = tk.BooleanVar()
        self.checkbox3_var = tk.BooleanVar()
        
        self.checkbox1 = tk.Checkbutton(self.canvas, text="داخل الضمان", variable=self.checkbox1_var, bg='#e6e6fa')
        self.checkbox1.pack(pady=2)
        self.checkbox2 = tk.Checkbutton(self.canvas, text="خارج الضمان ", variable=self.checkbox2_var, bg='#e6e6fa')
        self.checkbox2.pack(pady=2)
        self.checkbox3 = tk.Checkbutton(self.canvas, text="مرتجع", variable=self.checkbox3_var, bg='#e6e6fa')
        self.checkbox3.pack(pady=2)
        
        # حقل إدخال التاريخ
        self.date_label = tk.Label(self.canvas, text="أدخل التاريخ (YYYY-MM-DD):", bg='#e6e6fa')
        self.date_label.pack(pady=5)
        self.date_entry = tk.Entry(self.canvas)
        self.date_entry.pack(pady=5)
        self.date_entry.insert(0, datetime.today().strftime('%Y-%m-%d'))  # تعبئة التاريخ الحالي تلقائيًا
        
        # زر لإضافة البيانات
        self.add_button = tk.Button(self.canvas, text="إضافة البيانات", command=self.add_data, bg='#b0c4de')
        self.add_button.pack(pady=10)
        
        # زر لعرض شاشة البيانات الاحترافية
        self.view_button = tk.Button(self.canvas, text="عرض البيانات", command=self.view_data, bg='#b0c4de')
        self.view_button.pack(pady=10)
        
        # زر لعرض شاشة التقارير الاحترافية
        self.report_button = tk.Button(self.canvas, text="عرض التقارير", command=self.show_reports, bg='#b0c4de')
        self.report_button.pack(pady=10)

        # زر لعرض شاشة التعديل والحذف
        self.edit_delete_button = tk.Button(self.canvas, text="تعديل/حذف البيانات", command=self.show_edit_delete, bg='#b0c4de')
        self.edit_delete_button.pack(pady=10)
    
    def gradient_background(self, canvas, color1, color2):
        """Create a gradient background."""
        for i in range(500):
            r = int(i * (int(color2[1:3], 16) - int(color1[1:3], 16)) / 500 + int(color1[1:3], 16))
            g = int(i * (int(color2[3:5], 16) - int(color1[3:5], 16)) / 500 + int(color1[3:5], 16))
            b = int(i * (int(color2[5:7], 16) - int(color1[5:7], 16)) / 500 + int(color1[5:7], 16))
            color = f'#{r:02x}{g:02x}{b:02x}'
            canvas.create_line(0, i, 500, i, fill=color)
    
    def add_scroll(self, window):
        """Add a scrollbar to a window."""
        container = tk.Frame(window)
        canvas = tk.Canvas(container)
        scrollbar = tk.Scrollbar(container, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(
                scrollregion=canvas.bbox("all")
            )
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        container.pack(fill="both", expand=True)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        return scrollable_frame
    
    def update_repo(self):
        try:
            repo = git.Repo(self.repo_path)
            origin = repo.remotes.origin
            origin.pull()
            messagebox.showinfo("تحديث", "تم تحديث البرنامج إلى أحدث إصدار.")
        except Exception as e:
            messagebox.showerror("خطأ في التحديث", f"حدث خطأ أثناء التحديث: {e}")
    
    def select_file(self):
        self.file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if self.file_path:
            messagebox.showinfo("تم الاختيار", f"تم اختيار الملف: {self.file_path}")
    
    def add_data(self, event=None):
        if not self.file_path:
            messagebox.showerror("خطأ", "يرجى اختيار ملف Excel أولاً")
            return
        
        try:
            wb = load_workbook(self.file_path)
            sheet = wb.active
            
            # البحث عن أول صف فارغ
            row = 2
            while sheet[f"A{row}"].value is not None:
                row += 1
            
            # إضافة البيانات
            sheet[f"A{row}"] = row - 1
            sheet[f"B{row}"] = self.code_entry.get()
            sheet[f"E{row}"] = "True" if self.checkbox1_var.get() else ""
            sheet[f"F{row}"] = "True" if self.checkbox2_var.get() else ""
            sheet[f"G{row}"] = "True" if self.checkbox3_var.get() else ""
            sheet[f"H{row}"] = self.technician_combo.get()
            sheet[f"I{row}"] = self.date_entry.get()  # إضافة التاريخ
            
            wb.save(self.file_path)
            messagebox.showinfo("نجاح", "تمت إضافة البيانات بنجاح")
            
            # إعادة تعيين المدخلات
            self.code_entry.delete(0, tk.END)
            self.date_entry.delete(0, tk.END)
            self.date_entry.insert(0, datetime.today().strftime('%Y-%m-%d'))
        
        except Exception as e:
            messagebox.showerror("خطأ", f"حدث خطأ أثناء حفظ البيانات: {e}")
    
    def view_data(self):
        if not self.file_path:
            messagebox.showerror("خطأ", "يرجى اختيار ملف Excel أولاً")
            return
        
        try:
            wb = load_workbook(self.file_path)
            sheet = wb.active
            
            technician_count = Counter()
            for row in range(2, sheet.max_row + 1):
                technician = sheet[f"H{row}"].value
                if technician:
                    technician_count[technician] += 1
            
            # إنشاء نافذة جديدة لعرض البيانات مع سكرول
            view_window = tk.Toplevel(self.root)
            view_window.title("عرض البيانات")
            view_window.geometry("400x300")
            
            scrollable_frame = self.add_scroll(view_window)
            
            # عرض البيانات
            for technician, count in technician_count.items():
                label = tk.Label(scrollable_frame, text=f"{technician}: {count} قطع")
                label.pack(pady=5)
        
        except Exception as e:
            messagebox.showerror("خطأ", f"حدث خطأ أثناء عرض البيانات: {e}")
    
    def show_technician_details(self, event):
        if not self.file_path:
            messagebox.showerror("خطأ", "يرجى اختيار ملف Excel أولاً")
            return
        
        try:
            wb = load_workbook(self.file_path, data_only=True)  # عرض ناتج المعادلات
            sheet = wb.active
            
            technician_name = self.technician_combo.get()
            details_window = tk.Toplevel(self.root)
            details_window.title(f"تفاصيل {technician_name}")
            details_window.geometry("500x400")
            
            scrollable_frame = self.add_scroll(details_window)
            
            part_data = []
            for row in range(2, sheet.max_row + 1):
                technician = sheet[f"H{row}"].value
                if technician == technician_name:
                    part_code = sheet[f"B{row}"].value
                    part_name = sheet[f"C{row}"].value
                    part_price = sheet[f"D{row}"].value
                    warranty_status = "داخل الضمان" if sheet[f"E{row}"].value else "خارج الضمان"
                    returned_status = "مرتجع" if sheet[f"G{row}"].value else "غير مرتجع"
                    date = sheet[f"I{row}"].value
                    part_data.append((part_code, part_name, part_price, warranty_status, returned_status, date))
            
            for data in part_data:
                part_code, part_name, part_price, warranty_status, returned_status, date = data
                label = tk.Label(scrollable_frame, text=f"كود: {part_code}, اسم القطعة: {part_name}, السعر: {part_price}, الضمان: {warranty_status}, الحالة: {returned_status}, التاريخ: {date}")
                label.pack(pady=5)
        
        except Exception as e:
            messagebox.showerror("خطأ", f"حدث خطأ أثناء عرض التفاصيل: {e}")
    
    def show_edit_delete(self):
        if not self.file_path:
            messagebox.showerror("خطأ", "يرجى اختيار ملف Excel أولاً")
            return
        
        edit_delete_window = tk.Toplevel(self.root)
        edit_delete_window.title("تعديل/حذف البيانات")
        edit_delete_window.geometry("600x400")
        
        scrollable_frame = self.add_scroll(edit_delete_window)
        
        # إنشاء واجهة تعديل/حذف هنا
        wb = load_workbook(self.file_path, data_only=True)  # استخدام data_only لعرض ناتج المعادلات
        sheet = wb.active
        
        def delete_data(row):
            sheet.delete_rows(row)
            wb.save(self.file_path)
            messagebox.showinfo("نجاح", "تم حذف البيانات بنجاح")
            edit_delete_window.destroy()
        
        for row in range(2, sheet.max_row + 1):
            part_code = sheet[f"B{row}"].value
            part_name = sheet[f"C{row}"].value
            part_price = sheet[f"D{row}"].value
            
            label = tk.Label(scrollable_frame, text=f"كود: {part_code}, اسم القطعة: {part_name}, السعر: {part_price}")
            label.grid(row=row, column=0, padx=10, pady=5)
            
            edit_button = tk.Button(scrollable_frame, text="تعديل", command=lambda r=row: self.edit_data(r))
            edit_button.grid(row=row, column=1, padx=10, pady=5)
            
            delete_button = tk.Button(scrollable_frame, text="حذف", command=lambda r=row: delete_data(r))
            delete_button.grid(row=row, column=2, padx=10, pady=5)
    
    def edit_data(self, row):
        # يمكن إضافة واجهة تعديل هنا
        edit_window = tk.Toplevel(self.root)
        edit_window.title("تعديل البيانات")
        edit_window.geometry("400x300")
        
        wb = load_workbook(self.file_path)
        sheet = wb.active
        
        part_code = tk.Entry(edit_window)
        part_code.insert(0, sheet[f"B{row}"].value)
        part_code.pack(pady=5)
        
        part_name = tk.Entry(edit_window)
        part_name.insert(0, sheet[f"C{row}"].value)
        part_name.pack(pady=5)
        
        part_price = tk.Entry(edit_window)
        part_price.insert(0, sheet[f"D{row}"].value)
        part_price.pack(pady=5)
        
        def save_changes():
            sheet[f"B{row}"] = part_code.get()
            sheet[f"C{row}"] = part_name.get()
            sheet[f"D{row}"] = part_price.get()
            wb.save(self.file_path)
            messagebox.showinfo("نجاح", "تم تحديث البيانات بنجاح")
            edit_window.destroy()
        
        save_button = tk.Button(edit_window, text="حفظ التغييرات", command=save_changes)
        save_button.pack(pady=10)
    
    def show_reports(self):
        if not self.file_path:
            messagebox.showerror("خطأ", "يرجى اختيار ملف Excel أولاً")
            return
        
        try:
            wb = load_workbook(self.file_path)
            sheet = wb.active
            
            # حساب عدد القطع لكل فني
            technician_count = Counter()
            for row in range(2, sheet.max_row + 1):
                technician = sheet[f"H{row}"].value
                if technician:
                    technician_count[technician] += 1
            
            # إنشاء نافذة جديدة لعرض التقارير مع سكرول
            report_window = tk.Toplevel(self.root)
            report_window.title("التقارير")
            report_window.geometry("600x400")
            
            scrollable_frame = self.add_scroll(report_window)
            
            # رسم بياني لمقارنة عدد القطع لكل فني
            fig, ax = plt.subplots(figsize=(6, 4))
            technicians = list(technician_count.keys())
            counts = list(technician_count.values())
            ax.bar(technicians, counts, color='lightblue')
            ax.set_xlabel("الفني")
            ax.set_ylabel("عدد القطع")
            ax.set_title("عدد القطع لكل فني")
            
            canvas = FigureCanvasTkAgg(fig, master=scrollable_frame)
            canvas.draw()
            canvas.get_tk_widget().pack(fill="both", expand=True)
        
        except Exception as e:
            messagebox.showerror("خطأ", f"حدث خطأ أثناء عرض التقارير: {e}")

if __name__ == "__main__":
    repo_path = os.path.dirname(os.path.abspath(__file__))  # تحديد مسار المستودع
    root = tk.Tk()
    app = ExcelApp(root, repo_path)
    root.mainloop()
