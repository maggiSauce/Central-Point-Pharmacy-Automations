import tkinter as tk
from tkinter.filedialog import askopenfilename
from pypdf import PdfReader

PDFEXPORTPATH = r"C:\Users\small\Central-Point-Pharmacy\StudentForms\TempExport\Tester.pdf"


def main():
    tk.Tk().withdraw() # part of the import if you are not using other tkinter functions
    path = askopenfilename()
    reader = PdfReader(path)
    fields = reader.get_fields()
    for fieldName, fieldData in fields.items():
        print(f"{fieldName}: {fieldData.get('/V')}")


main()