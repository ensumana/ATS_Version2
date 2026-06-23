# ATS_Version2
This is a repo for demo and experimentation purpose.

This code extracts the resumes from the emails in outlook desktop app and extracts the important details from the resumes extracted. These are stored in a folder of your choice.

The Package requirements for this is : 
  os
  re
  win32com.client
  pythoncom
  pandas
  pdfplumber
  docx
  datetime
  customtkinter
  tkinter
  datetime
  threading

  If you dont have these please use pip install command to install these. I have run this code on pycharm, ans it executes well. for the execution you can run gui_app2.py file.

  also to create a .exe file, run the command below in the bash
  **pyinstaller --onefile --windowed --hidden-import=win32com --hidden-import=win32com.client --hidden-import=pythoncom gui_app2.py**

  
