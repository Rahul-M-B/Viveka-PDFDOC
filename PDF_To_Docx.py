#!/usr/bin/env python
# coding: utf-8

# In[1]:


pip install pdf2docx


# In[4]:


import tkinter as tk
from tkinter import filedialog
from tkinter import Label
import tkinter.font as font
import os
from pdf2docx import Converter

def add_file():
    filetypes = [("PDF Files", "*.pdf")]
    f_path = filedialog.askopenfilename(initialdir="/", title="Select PDF File", filetypes=filetypes)
    if f_path:
        convert_file(f_path)

def convert_file(f_path):
    save_dir = r'C:\Users\user\Desktop\Converter_Demo' 
    if not os.path.exists(save_dir):
        os.makedirs(save_dir) 

    file_name = os.path.basename(f_path)  
    docx_filename = os.path.join(save_dir, os.path.splitext(file_name)[0] + '.docx')

    con = Converter(f_path)
    con.convert(docx_filename, start=0, end=None)
    con.close()

    success_message = Label(frame, text="PDF has been converted successfully!", font=('Consolas 16'),bg="#d8e0ed", fg="#00080d")
    success_message.pack()
    s=Label(frame,text="Please check the same directory for the docx file",font=('Consolas 10'),bg="#d8e0ed",fg="#00080d")
    s.pack()

base = tk.Tk()
base.title("PDF to Docx Converter")

interface = tk.Canvas(base, height=600, width=600)
interface.pack()

frame = tk.Frame(base,borderwidth=5,bg="#d8e0ed",relief="solid")
frame.place(relwidth=0.8, relheight=0.8, relx=0.1, rely=0.1)
tk.Label(frame, text=" ", font='Century ', bg="#d8e0ed").pack()

tk.Label(frame, text="PDF to Docx Converter", font='Consolas 26 underline', bg="#d8e0ed").pack()
tk.Label(frame, text=" ", font='Century', bg="#d8e0ed").pack()
tk.Label(frame, text=" ", font='Century', bg="#d8e0ed").pack()

o_button = tk.Button(frame, text="Select PDF File", font='Helvetica',padx= 15, pady=15, bg="#89afc9", fg="#00080d", command=add_file,borderwidth=3,relief="solid")
o_button.place(x=325, y=600)
o_font = font.Font(size=15)
o_button["font"] = o_font
o_button.pack()
tk.Label(frame, text=" ", font='Century', bg="#d8e0ed").pack()
tk.Label(frame, text=" ", font='Century', bg="#d8e0ed").pack()
base.mainloop()


# In[ ]:




