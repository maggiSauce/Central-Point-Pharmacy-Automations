#include JSON.ahk

programPath = C:\Users\kroll\Documents\Central-Point-Pharmacy\PDFScrape.py
JSONPath = C:\Users\kroll\Documents\Central-Point-Pharmacy\TempPDFs\tempFields.json
; medDataPath = C:\Users\kroll\Documents\Central-Point-Pharmacy\TempPDFs\medications_data.json
medDataPath = C:\Users\kroll\Documents\Central-Point-Pharmacy\TempPDFs\medications_data_with_DSigs.json
travelPDFPath = C:\Users\kroll\Desktop\TravelVaccinationForms\

;programPath = C:\Users\small\Central-Point-Pharmacy\PDFScrape.py
;JSONPath = C:\Users\small\Central-Point-Pharmacy\TempPDFs\tempFields.json
;medDataPath = C:\Users\small\Central-Point-Pharmacy\TempPDFs\medications_data.json

^Esc::ExitApp	; CTRL + esc
^+q::
    InputBox, filepath, Input the pdf file
    if (filepath == "") {
        MsgBox, No filepath given
        return
    }
    filepath := % travelPDFPath filepath
    ; MsgBox % filepath
    if (!runPython(programPath, filepath)) {    ; exit hotkey execution if runPython returns False
        return
    }
    if (!tempData := parseJSON(JSONPath)) {     ; load tempFields.Json
        return
    }
    if (!medData := parseJSON(medDataPath)) {    ; load medications_data.json
        return
    }
    ; MsgBox, Successful JSON reads
    fillIndividualDrug(tempData, medData)
    MsgBox, Done!!! :) 
    return

runPython(programPath, filepath) {
    ; MsgBox, python "%programPath%" "%filepath%"
    RunWait, python "%programPath%" "%filepath%"
    exitCode := ErrorLevel      ; save the python exit code

    if (exitCode == 102) {
        MsgBox, %filepath% does not exist
        return 0
    }
    if (exitCode != 0) {
        MsgBox, Python file exited with non-zero code: %exitCode%
        return false
    }
    return true
}

parseJSON(pathToJson) {
    ; MsgBox, % pathToJson
    
    if !FileExist(pathToJson) {
        MsgBox % "Error: JSON file does not exist at " pathToJson
        return 0
    }

    FileRead, jsonContent, %pathToJson%   ; load the file
    if (jsonContent = "") {
        MsgBox % "Error: JSON file is empty."
        return 0
    }

    data := JSON.Load(jsonContent)      ; parse json
    if (IsObject(data)) {
        ;MsgBox % "JSON loaded successfully!"
    } else {
        MsgBox % "Failed to load JSON."
        return 0
    }
    return data
}

fillIndividualDrug(data, medData) {
    firstTime := true

    for key, value in data {
        ; MsgBox, % "Key: " key " Value: " value
        if (value != "/Yes") {
            continue    ; do not do iteration if not checked yes
        }
        ; MsgBox, Yes
        if (!medData.HasKey(key)) {      ; check if data.key is in medData.key
            continue
        }
        ; MsgBox, HasKey
        item := medData[key]
        if (!(IsObject(item) && item.HasKey("DIN"))) {   ; check if item has a DIN
            continue
        }

        ; at this point, the item will have a been checked yes and have an associated DIN
        ; MsgBox, % "Key: " key " Value: " value " DIN: " item["DIN"]

        if (firstTime) {
            Send, {F12}     ; create new Rx
            Sleep, 1000
            firstTime := false
        }

        ; DIN and Doc
        Send, % item["DIN"]
        Send, {Tab}
        Send, 14127
        Send, {Enter}
        Sleep, 1000

        if (item["DIN"] == "02247208") {
            handleDukoral(key)
        } else if (item["DIN"] == "02466783") {    ; malerone adult
            handleVariableDQandDays("Total Malaria tabs Malarone", data)
            continue
        }   else if (item["DIN"] == "02264935") {
            handleMalaronePed(key, data)
            continue
        } else if (item["DIN"] == "00725250") {     ; Doxycyline
            handleVariableDQandDays("Total Malaria tabs Doxy", data)
            continue
        } else if (item["DIN"] == "02244366") {     ; mefloquine
            handleMefloquine(data)
            continue
        } else if (item["DIN"] == "00545015") {     ; diamox prevention
            Send, % item["sig"]
            Send, {Tab}
            handleVariableDQandDays("Diamox Tabs Prevent", data)
            Continue
        } else if (item["DIN"] == "02333279") {
            handleIxiaro()
        } else if (item["DIN"] == "02247600") {
            handleBoostrix()
        } else if (item["DIN"] == "02332396") {
            handleAzithro600(data, item)
            Continue
        }
       
        ; sig
        if (item["sig"] == "DEFAULT") {
            ; pass
        } else if (item["sig"] == "VARIABLE") {
            if (item["DIN"] == "02274388") {
                handleAzithroSusp(data)
            }
            Send, {Tab}
        } else {    ; case for non default sig
            Send, % item["sig"]
            Send, {Tab}
        }
        Sleep, 1000

        ; Disp QTY
        Send, % item["quantity"]
        Send, {Tab}
        Sleep, 500

        Send, ^r	; sent ctrl r to specify repeats
		Send, % item["refills"]
		Send, {Enter}

        ; make Rx unfilled
		Send, {Alt}
		Send, r
		Send, {Enter}
		Sleep, 1000

        ; Days
        Send, % item["days_supply"]	; DAYS
		Sleep, 500

        Send, {F12}	; final fill
		Sleep, 500

        Send, {Enter}
		Send, {Enter}
		Send, {Enter}
		Send, {Enter}
		Sleep, 500
		Send, {Enter}
		Sleep, 500
		Send, {Enter}
		Sleep, 3000
    }
    return
}

handleDukoral(key) {
     ; Dukoral exception
    if (key == "Dukoral") {     
        Send, {Down}	; select second version of Dukorol
        Send, {Enter}
        Sleep, 500
        Send, {Enter}
        Sleep, 1000
    }
    if (key == "Dukoral booster") { ; select first version of dukoral
        Send, {Enter}
        Sleep, 500
        Send, {Enter}
        Sleep, 1000
    }
    return
}

handleIxiaro() {
    Send, {Down}	; select second version of Ixario
    Send, {Enter}
    Sleep, 500
    Send, {Enter}
    Sleep, 1000
}

handleBoostrix() {
    Send, {Down}	; select second version of Boostrix
    Send, {Enter}
    Sleep, 500
    Send, {Enter}
    Sleep, 1000
}

handleVariableDQandDays(fieldName, data) {
    ; calculate the DQ and Days
    dspQty := data[fieldName]   ; get the dspqty for malarone adult
    
     ; Disp QTY
        Send, % dspQty
        Send, {Tab}
        Sleep, 500

    Send, ^r	; sent ctrl r to specify repeats
    Send, % item["refills"]
	;	Sleep, 3000	;Remove
    Send, {Enter}
	;	Sleep, 3000	; Test

        ; make Rx unfilled
		Send, {Alt}
		Send, r
		Send, {Enter}
		Sleep, 1000

        ; Days
        Send, % dspQty	; DAYS
		Sleep, 500

        Send, {F12}	; final fill
		Sleep, 500

        Send, {Enter}
		Send, {Enter}
		Send, {Enter}
		Send, {Enter}
		Sleep, 500
		Send, {Enter}
		Sleep, 500
		Send, {Enter}
		Sleep, 3000
        Return
}

handleMalaronePed(key, data) {
    dspQty := data["Total Malaria tabs Malarone"]   ; get the dspqty for malarone adult
    ; MsgBox, % dspQty
    if (data["3T QD - MP"] == "/Yes") {
        days := dspQty / 3
        sigNum := 3
    } else if (data["2T QD - MP"] == "/Yes") {
        days := dspQty / 2
        sigNum := 2
    } else if (data["1T QD - MP"] == "/Yes") {
        days := dspQty / 1
        sigNum := 1
    } else if (data["3/4T QD - MP"] == "/Yes") {
        days := dspQty / (3/4)
        sigNum := "3/4"
    } else if (data["1/2T QD - MP"] == "/Yes") {
        days := dspQty / (1/2)
        sigNum := "1/2"
    } else {
        MsgBox, "No malerone pediatric box selected"
        ExitApp, 301
    }

    sigTemplate := % "TAKE " sigNum " TABLET ONCE DAILY (WITH FOOD), START 1 DAY PRIOR TO EXPOSURE, DURING STAY IN REGION AND FOR 1 WEEK AFTER LEAVING ENDEMIC AREA"
    Send, % sigTemplate
    Send, {Tab}

    ; Disp QTY
    Send, % dspQty
    Send, {Tab}
    Sleep, 500

    Send, ^r	; sent ctrl r to specify repeats
    Send, % item["refills"]
;	Sleep, 3000	;Remove
    Send, {Enter}
;	Sleep, 3000	; Test

    ; make Rx unfilled
    Send, {Alt}
    Send, r
    Send, {Enter}
    Sleep, 1000

    ; Days
    ; MsgBox, % Ceil(days)
    Send, % Ceil(days)	; DAYS
    Sleep, 500

    Send, {F12}	; final fill
    Sleep, 500

    Send, {Enter}
    Send, {Enter}
    Send, {Enter}
    Send, {Enter}
    Sleep, 500
    Send, {Enter}
    Sleep, 500
    Send, {Enter}
    Sleep, 3000
    Return
}

handleMefloquine(data) {
    dspQty := data["Days travelling Mefloquine"]   ; get the dspqty for malarone adult
    ; MsgBox, % dspQty
    if (data["1T QD - MFLQ"] == "/Yes") {
        days := dspQty / 1
        sigNum := 1
    } else if (data["1/4T QD - MFLQ"] == "/Yes") {
        days := dspQty / (1/4)
        sigNum := "1/4"
    } else if (data["3/4T QD - MFLQ"] == "/Yes") {
        days := dspQty / (3/4)
        sigNum := "3/4"
    } else if (data["1/2T QD - MFLQ"] == "/Yes") {
        days := dspQty / (1/2)
        sigNum := "1/2"
    } else {
        MsgBox, "No mefloquine box selected"
        ExitApp, 302
    }

    sigTemplate := % "TAKE " sigNum " TABLET ONCE A WEEK, START 1 WEEK PRIOR TO EXPOSURE, DURING STAY IN REGION AND WEEKLY FOR 4 WEEKS AFTER LEAVING ENDEMIC AREA"
    Send, % sigTemplate
    Send, {Tab}

    ; Disp QTY
    Send, % dspQty
    Send, {Tab}
    Sleep, 500

    Send, ^r	; sent ctrl r to specify repeats
    Send, % item["refills"]
;	Sleep, 3000	;Remove
    Send, {Enter}
;	Sleep, 3000	; Test

    ; make Rx unfilled
    Send, {Alt}
    Send, r
    Send, {Enter}
    Sleep, 1000

    ; Days
    ; MsgBox, % Ceil(days)
    Send, % Ceil(days)	; DAYS
    Sleep, 500

    Send, {F12}	; final fill
    Sleep, 500

    Send, {Enter}
    Send, {Enter}
    Send, {Enter}
    Send, {Enter}
    Sleep, 500
    Send, {Enter}
    Sleep, 500
    Send, {Enter}
    Sleep, 3000
    Return
}

handleAzithroSusp(data) {
    azML := data["AZ_mL"]   ; get the ml amout for az

    sigTemplate := % "Give " azML " mL daily for 3 days (for travellers diarrhea)(discard remaining)"
    Send, % sigTemplate
    Return
}

handleAzithro600(data, item) {
    azML := data["AZ_mL"]   ; get the ml amout for az
    if (azML <= 5) {    ; select first (pack size 15)
        Send, {Enter}
        Sleep, 500
        Send, {Enter}
        Sleep, 1000
        dspQty := 15
    } else if (azML > 5 && azML <= 7.5) {
        Send, {Down}	; select second (pack size 22.5)
        Send, {Enter}
        Sleep, 500
        Send, {Enter}
        Sleep, 1000
        dspQty := 22.5
    } else if (azML > 7.5) {
        Send, {Down}	; select third (pack size 37.5)
        Send, {Down}
        Send, {Enter}
        Sleep, 500
        Send, {Enter}
        Sleep, 1000
        dspQty := 37.5
    } else {
        MsgBox, incorrect azithro suspension ML amount. Terminating
        ExitApp, 201
    }
    sigTemplate := % "Give " azML " mL daily for 3 days (for travellers diarrhea)(discard remaining)"
    Send, % sigTemplate
    Send, {Tab}

    ; Disp QTY
    Send, % dspQty
    Send, {Tab}
    Sleep, 500

    Send, ^r	; sent ctrl r to specify repeats
    Send, % item["refills"]
;	Sleep, 3000	;Remove
    Send, {Enter}
;	Sleep, 3000	; Test

    ; make Rx unfilled
    Send, {Alt}
    Send, r
    Send, {Enter}
    Sleep, 1000

    ; Days
    Send, % item["days_supply"]	; DAYS
    Sleep, 500

    Send, {F12}	; final fill
    Sleep, 500

    Send, {Enter}
    Send, {Enter}
    Send, {Enter}
    Send, {Enter}
    Sleep, 500
    Send, {Enter}
    Sleep, 500
    Send, {Enter}
    Sleep, 3000
    Return
}