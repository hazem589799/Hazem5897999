from tkinter import *
import tkinter as tk
from tkinter import ttk
import os
import sys
import shutil
import subprocess
import zipfile
import requests
from tkinter import messagebox

#=====================start to make my main window ============================

try:
    url = "https://github.com/hazem589799/Hazem5897999/archive/refs/heads/main.zip"
    local_zip_path = "update.zip"
    new_program_path = "Hazem5897999-main/Hazem.py"
    
    # تحميل الملف المضغوط
    response = requests.get(url)
    with open(local_zip_path, 'wb') as file:
        file.write(response.content)
    
    # فك ضغط الملف
    with zipfile.ZipFile(local_zip_path, 'r') as zip_ref:
        zip_ref.extractall()
    
    # تحويل الملف إلى exe
    subprocess.call([
        'pyinstaller', '--onefile', '--distpath', '.', '--workpath', '.', '--specpath', '.', new_program_path
    ])
    
    # اسم الملف التنفيذي الجديد
    exe_file = "Hazem.exe"
    
    # الحصول على مسار الملف الأصلي
    original_path = os.path.abspath(__file__)
    original_dir = os.path.dirname(original_path)
    
    # المسار الجديد للملف التنفيذي
    new_exe_path = os.path.join(original_dir, exe_file)
    
    # حذف النسخة القديمة
    if os.path.isfile(new_exe_path):
        os.remove(new_exe_path)
    
    # نقل النسخة الجديدة إلى المسار الأصلي
    shutil.move(os.path.join('.', exe_file), new_exe_path)
    
    # حذف الملف المضغوط والتحديث
    os.remove(local_zip_path)
    shutil.rmtree('Hazem5897999-main')


    # إعادة تشغيل البرنامج بالنسخة الجديدة
    subprocess.call([new_exe_path])
    self.root.destroy()


except Exception as e:
    messagebox.showerror("خطأ", f"حدث خطأ أثناء تحديث البرنامج: {e}")

def enter():
    password="30507280201237"
    passwrd = txt_username_login.get()
    if passwrd == password :
        main_wd = tk.Tk()
        main_wd.title("main page")
        main_wd.geometry("1307x500+10+10")
        main_wd.resizable(False , False)
        main_wd.configure(bg='#aed6f1')
        title_label = tk.Label(main_wd , text="بسم الله الرحمن الرحيم" , font=("Calibri" , 18 , "bold") , fg='#6c3483' , bg='#aed6f1')
        title_label.place(x=650 , y=10)
        login_wd.withdraw()
        main_wd.mainloop()

        
    else:
        messagebox.showerror( 'كلمة السر غلط' , 'كلمة السر غلط روح اعرف كلمة السر وتعال حاول تانى')
        
        
        pass
login_wd = tk.Tk()
login_wd.geometry("1307x500+10+10")
login_wd.resizable(False , False)
login_wd.configure(bg='#aed6f1')
title_label = tk.Label(login_wd , text="بسم الله الرحمن الرحيم" , font=("Calibri" , 18 , "bold") , fg='#6c3483' , bg='#aed6f1')
title_label.place(x=650 , y=10)
wd_frame = Frame(login_wd , width=750 , height=350 , bg='#d4e6f1')
wd_frame.place(x=250 , y=70)
label_title_login = tk.Label(wd_frame , text="تسجيل الدخول" ,font=("Calibri" , 18 , "bold") , fg='blue' , bg='#d4e6f1')
label_title_login.place(x=350 , y=10)
label_username_login = Label(wd_frame ,text=":من فضلك ادخل كلمة المرور" ,font=("Calibri" , 18 , "bold") , fg='#6e2c00' , bg='#d4e6f1')
label_username_login.place(x=500, y=150)
txt_username_login = Entry(wd_frame,show="*" ,width=20 , bg='#fcf3cf' , fg='#ba4a00', font=("Calibri" , 20 , "bold"))
txt_username_login.place(x=200, y=155)
btn_user_name = Button(wd_frame , text="دخول" , command=enter , width=18 , height=1 , bg="green" ,fg="white" ,  font=("Calibri" , 16 , "bold"))
btn_user_name.place(x=250 , y=220)

login_wd.mainloop()
