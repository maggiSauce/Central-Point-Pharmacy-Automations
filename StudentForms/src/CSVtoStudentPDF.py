import csv
import sys
import traceback
from pypdf import PdfReader, PdfWriter
from pypdf.generic import NameObject, BooleanObject
import tkinter as tk
from tkinter.filedialog import askopenfilename

PDFTEMPLATEPATH = r"C:\Users\small\Central-Point-Pharmacy\StudentForms\Templates"
PDFEXPORTPATH = r"C:\Users\small\Central-Point-Pharmacy\StudentForms\TempExport"
# CSVPATH = r"C:\Users\small\Central-Point-Pharmacy\StudentForms\Patient listing report - Copy.csv"

# PDFTEMPLATEPATH = r"C:\Users\kroll\Desktop\School Forms\Templates"
# PDFEXPORTPATH = r"C:\Users\kroll\Desktop\School Forms\Output"

SCHOOLSLIST = ["CDI",
               "Norquest",
               "MacEwan",
               "UofA",
               "NAIT"]

def openFile(filepath:str) -> list:
    '''
    opens and formats a file
    returns a list of dicts for each row in the file
    '''

    rowsList = []
    
    with open(filepath, newline='', encoding='utf-8') as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            filledFieldsDict = {
            key: value for key, value in row.items() if value and value.strip()
            }
            rowsList.append(filledFieldsDict)
    return rowsList

def extractPhoneNumber(numberString):
    """
    Extracts the first available phone number and returns it. 
    If no number is found, returns false
    """

    phoneNumber = ''
    activeNumber = False

    for char in numberString:
        if char.isdigit() or char in '()- ':
            phoneNumber += char
            activeNumber = True
        elif activeNumber:
            break
    if phoneNumber:
        return phoneNumber
    else:
        return None
    
def isMale(genderString):
    if genderString == "M":
        return "/On"
    return '/Off'

def isFemale(genderString):
    if genderString == "F":
        return "/On"
    return '/Off'
    
def formatPLR(PLRList: list) -> list:
    '''
    Formats the Patient Listing Report dictionary
    Returns PatientInfoList which is a list of dicts that hold pdf fields as keys and corresponding values
    '''
    PatientInfoList = []

    for PLRDict in PLRList:
        commentsValue = PLRDict.pop("Comments")
        commentsList = commentsValue.split("\n")
        commentsList[0] = commentsList[0][9:]       # removes "General: " from first element
        for commentVal in commentsList:
            pair = commentVal.split(":")
            for element in pair:
                element = element.strip()
            if len(pair) < 2:
                continue
            PLRDict[pair[0]] = pair[1]

        try:
            PDFDict = {}
            PDFDict["Last Name"] = PLRDict["LastName"]
            PDFDict["First Name"] = PLRDict["FirstName"]
            PDFDict["Date of Birth"] = PLRDict["Birthday"]
            PDFDict["PHN"] = PLRDict["PHN"]
            PDFDict["Address"] = PLRDict["Address1"]
            PDFDict["City Town"] = PLRDict["City"]
            PDFDict["Province"] = PLRDict["Province"]
            PDFDict["Postal Code"] = PLRDict["Postal"]
            PDFDict["Phone"] = extractPhoneNumber(PLRDict["PhoneNumbers"])
            PDFDict["Program"] = PLRDict["Program"]
            PDFDict["Student ID"] = PLRDict["StudentNumber"]
            PDFDict["Male"] = isMale(PLRDict["Sex"])
            PDFDict["Female"] = isFemale(PLRDict["Sex"])
            PDFDict["School"] = PLRDict["School"].strip()
        except:
            continue

        if PDFDict["School"] == '':
            continue

        PatientInfoList.append(PDFDict)

    return PatientInfoList

def writeToPDF(reader, PDFDict, patientName):
    path = PDFEXPORTPATH + '\\' + patientName + '.pdf'

    writer = PdfWriter()
    writer.append(reader)

    writer.update_page_form_field_values(
        writer.pages[0], 
        PDFDict,
        auto_regenerate = True
    )

    # Ensure that checkbox fields like Male and Female are set across all pages
    for page_num, page in enumerate(writer.pages):
        for annot in page['/Annots']:
            field_name = annot.get('/T')
            if field_name and field_name[1:-1] in PDFDict:  # Field name should match, strip leading and trailing quotes
                value = PDFDict[field_name[1:-1]]  # Get the value to set for this checkbox
                if value in ["/On", "/Off"]:
                    # Set the correct value for the checkbox
                    if value == "/On":
                        annot.update({
                            NameObject("/V"): NameObject("/Yes")  # Set as checked
                        })
                    elif value == "/Off":
                        annot.update({
                            NameObject("/V"): NameObject("/Off")  # Set as unchecked
                        })

    # Copy over AcroForm and set NeedAppearances = True
    writer._root_object.update({NameObject("/AcroForm"): reader.trailer["/Root"]["/AcroForm"]}) 
    writer._root_object["/AcroForm"].update({NameObject("/NeedAppearances"): BooleanObject(True)})

    with open(path, "wb") as outputStream:     # 'wb' is for write binary mode
        writer.write(outputStream)

def main():
    log = open('CSVtoPDFLog.txt', 'w')

    tk.Tk().withdraw() # part of the import if you are not using other tkinter functions

    chosenCSVPath = askopenfilename()
    try:
        PDFInfoList = formatPLR(openFile(chosenCSVPath))
    except KeyError as e:
        log.write(f"Error Reading CSV: {e}, this key is not in the CSV file\n")
        log.write(traceback.format_exc())
        tk.messagebox.showinfo("CSV Converter Error", "There was an error converting your CSV. \nPlease read CSVtoPDFLog")
        sys.exit(101)
    except Exception as e:
        print(f"Error reading CSV: {e}")
        log.write(f"Error reading CSV: {e}\n")
        log.write(traceback.format_exc())
        tk.messagebox.showinfo("CSV Converter Error", "There was an error converting your CSV. \nPlease read CSVtoPDFLog")

        sys.exit(101)
    # print(PDFInfoList)
    
    for i in range(len(PDFInfoList)):
        if PDFInfoList[i]['School'] in SCHOOLSLIST:
            templatePath = PDFTEMPLATEPATH + '\\' + PDFInfoList[i]['School'] + '.pdf'
        else:
            log.write(f"{PDFInfoList[i]['First Name']} {PDFInfoList[i]['Last Name']} does not attend a listed school")
            print(f"{PDFInfoList[i]['First Name']} {PDFInfoList[i]['Last Name']} does not attend a listed school")
            continue

        try:
            reader = PdfReader(templatePath)
            patientName = f'{PDFInfoList[i]["First Name"]}{PDFInfoList[i]["Last Name"]}'
            writeToPDF(reader, PDFInfoList[i], patientName)
        except Exception as e:
            log.write(f"Error reading PDF template or writing to patient PDF: {e}\n")
            log.write(traceback.format_exc())
            tk.messagebox.showinfo("CSV Converter Error", "There was an error converting your CSV. \nPlease read CSVtoPDFLog")
            sys.exit(102)

    log.write("Exit code 0")
    log.close()
    tk.messagebox.showinfo("CSV Converter Completed", "Successfully created all eligible pdfs")
main()