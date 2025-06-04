; this is a script that autofills a drug on the Kroll software
; Preconditions
;	[1]	Profile has Pharmacist Privaleges
;	[2] must be a patient page and have selected a patient
;	[3]	Hotbar must not be active (If any of the top row is blue, it is active, to deactivate press alt)
;	[4]	all variable values such as DIN, AMOUNT etc are how we want them in the code
;
;

;		Ozempic    Wegovy      Rybelsus    Mounjaro    Contrave    Saxenda    Orlistat
WLDIN := ["02540258", "02528509", "02497581", "02551950", "02472945", "02437899", "02240325"]
WLDspQty := ["3", "1.5", "30", "1", "120", "15", "84"]
WLDays := ["28", "28", "30", "28", "30", "30", "28"]

;	Students = [MMR, Varivax, Engerix B, Boostrix]
StuDIN := ["466085", "2246081", "02487039", "02247600"]
StuDspQty := ["1", "1", "1", "0.5"]
StuDays := ["28", "28", "28", "28"]
StuRepeats := ["1", "1", "2", "0"]

;	Travel = [Azithro, Avaxim, Typhim]
TravDIN := ["2442434", "2237792", "2130955"]
TravDspQty := ["8", "1", "1"]
TravDays := ["3", "30", "30"]
TravRepeats := ["0", "1", "0"]

^Esc::ExitApp	; CTRL + esc

;==================================================
; WEIGHT LOSS (1)
;==================================================

^+1::	; Ctrl (^) + Shift (+) + 1
	weightLoss(WLDIN, WLDspQty, WLDays) 
return

weightLoss(DINArr, DspQtyArr, DaysArr) {
	Send, {F12}	; hit F12 to navigate to fill Rx
	Sleep, 1000
	
	; loop over all DINS
	Loop, % DINArr.Length() {
		i := A_Index	; initialise index
		
		; active field should be "drug search"
		Send, % DINArr[i]		; 	DIN
	;	Sleep, 500
		Send, {Tab}				; 	tab to move to "doc search" field
	;	Sleep, 500
		Send, 14127				; 	doc primary key
		Send, {Enter}
		Sleep, 1000

		; active field should be "Disp Qty"
		Send, % DspQtyArr[i]	; DISP qTY
		Sleep, 1000

		; make Rx unfilled
		Send, {Alt}
	;	Sleep, 500
		Send, r
	;	Sleep, 500
		Send, {Enter}
		Sleep, 1000

		Send, % DaysArr[i]	; DAYS
		Sleep, 500
		
		Send, {F12}	; final fill
		Sleep, 500
		
		Send, {Enter}
		Sleep, 500
		Send, {Enter}
	;	Send, ENDITER
		Sleep, 3000
	}
	return
}


;==================================================
; Students (2)
;==================================================

^+2::
	fillWithRepeats(StuDIN, StuDspQty, StuRepeats, StuDays)
return


;==================================================
; Travel (3)
;==================================================
^+3::	; CTRL + SHIFT + 3
	fillWithRepeats(TravDIN, TravDspQty, TravRepeats, TravDays)
return

fillWithRepeats(DINArr, DspQtyArr, RepeatsArr, DaysArr) {
	Send, {F12}	; hit F12 to navigate to fill Rx
	Sleep, 1000

	; loop over all DINS
	Loop, % DINArr.Length() {
		i := A_Index	; initialise index

		; active field should be "drug search"
		Send, % DINArr[i]		; 	DIN
	;	Sleep, 500	;Remove
		Send, {Tab}				; 	tab to move to "doc search" field
	;	Sleep, 500	;Remove
		Send, 14127				; 	doc primary key
		Send, {Enter}
		Sleep, 1000
		
		if (DINArr[i] == "02247600") {
			Send, {Down}	; select second version of boostrix
			Send, {Enter}
			Sleep, 500
			Send, {Enter}
			Sleep, 1000
		}

		; active field should be "Disp Qty"
		Send, % DspQtyArr[i]	; DISP qTY
		Send, {Tab}
		Sleep, 500
		
		Send, ^r	; sent ctrl r to specify repeats
	;	Send, {Ctrl down}r{Ctrl up}
	;	Sleep, 3000	;Remove
		Send, % RepeatsArr[i]
	;	Sleep, 3000	;Remove
		Send, {Enter}
	;	Sleep, 3000	; Test

		; make Rx unfilled
		Send, {Alt}
	;	Sleep, 500	;Remove
		Send, r
	;	Sleep, 500	;Remove
		Send, {Enter}
		Sleep, 1000

		Send, % DaysArr[i]	; DAYS
		Sleep, 500
		
		Send, {F12}	; final fill
		Sleep, 500
		
		Send, {Enter}
		;Sleep, 500
		Send, {Enter}
		;Sleep, 500
		Send, {Enter}
		;Sleep, 500
		Send, {Enter}
		Sleep, 500
		Send, {Enter}
		Sleep, 500
		Send, {Enter}
	;	Send, ENDITER
		Sleep, 3000
	}
	return
}