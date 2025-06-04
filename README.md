# Central-Point-Pharmacy-Automations

This repo contains the automation suite created for Central Point Pharamacy originally created in Spring 2025. The tools were created in order to speed up pharmacy workflow by removing the repetitive, manual tasks phamacists and pharmacy techs must do on a daily basis.\
The project is broken up into two distinct parts that act independantly of each other:
  1. Automations
  2. Student Forms

# Automations
The automations were created to interact with the Kroll Pharmacy software. These automations save the bulk of a pharmacists time by shortening what would be 30 minutes of data entry per patient down to just 30 seconds.\
KrollAllHotkeys.ahk is a set of 3 different pre-defined workflows (Students, Travel, Weight-Loss) that must are triggered in anticipation of a new patient.
### Patient Travel Form
The automation that handles the patient travel forms is done through the interaction of an AutoHotkey and a Python script. The goal is to fill prescriptions based on information gathered from a patient traveller form that is filled out prior to an appointment. 

The KrollPDR.ahk file is the entry and endpoint of the program. It calls PDFscrape.py to gather the information stored on the patient traveller forms and, using the drug information stored in a pre-made JSON file, it prescribes the correct medications according to the individual patient form. Many of the functions within this file are drug specific as several drugs require unique keystrokes to prescribe.

The PDFScrape.py file handles PDF file selection and parsing before passing the results back to KrollPDR.ahk. It uses PDFReader to scrape the PDF data before storing it into a JSON file which is used as the form of communication.

# Student Forms
Using Kroll, the pharmacist can bulk export patient data to be used to fill in patient profile documents which are then automatically filled by CSVtoStudentPDF.py. Previously, the pharmacist would have to fill the documents in by hand for each new student. During the school season, this automated workflow saves over an estimated 50 hours of manual work. 

### CSVtoStudent.PDF
This program allows the pharmacist to select a CSV file filled with student information and produce a PDF version of the student profile documents for every student in the CSV. It achieves this by first extracting the information from the CSV and parsing each row. Then it creates a new individual pdf for every student and fills in the corresponding information. There are 5 templates for different universities in Edmonton, AB that are automatically chosen based on the information given in the CSV

# Outcome
The automations and programs in this repo are in use today at Central Point Pharmacy in Edmonton, AB. Each year, hundreds of cumulative hours are saved on data entry.\
Author: Benjamin Bui
