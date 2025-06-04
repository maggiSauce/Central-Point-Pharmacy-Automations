from pypdf import PdfReader
import subprocess
import sys
import json

JSONPATH = r"C:\Users\kroll\Documents\Central-Point-Pharmacy\TempPDFs\tempFields.json"
# JSONPATH = r"C:\Users\small\Central-Point-Pharmacy\TempPDFs\tempFields.json"

def main():
    log = open("pythonLog.txt", 'w')
    if len(sys.argv) > 1:
        filepath = sys.argv[1]
    else:
        log.write("No filepath entered")
        print("No filepath entered")
        exit(101)
    
    try:
        reader = PdfReader(filepath)
    except Exception as e:
        log.write(f"Error reading PDF: {e}")
        print(f"Error reading PDF: {e}")
        exit(102)

    if type(reader) == None:
        log.write("Not a parsable pdf file.")
        print("Not a parsable pdf file.")
        exit(103)
    
    fields = reader.get_fields()

    if not fields:
        log.write("No fields found in the PDF")
        print("No fields found in the PDF")
        exit(201)
    
    pdfData = {}    # create empty dictionary

    for fieldName, fieldData in fields.items():
        print(f"{fieldName}: {fieldData.get('/V')}")
        value = fieldData.get('/V')
        pdfData[fieldName] = value  # add key/value pair to dictionary

    jsonObject = json.dumps(pdfData, indent = 4)
    with(open(JSONPATH, 'w', encoding='utf-8') as outfile):
        outfile.write(jsonObject)

    log.write("completed with no errors")
    log.close()
    print("completed with no errors")
main()