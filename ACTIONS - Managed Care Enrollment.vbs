'**********THIS IS A HENNEPIN SPECIFIC SCRIPT.  IF YOU REVERSE ENGINEER THIS SCRIPT, JUST BE CAREFUL.************
'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "Action - Managed Care Enrollment.vbs"
start_time = timer

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN		'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF default_directory = "C:\DHS-MAXIS-Scripts\Script Files\" OR default_directory = "" THEN			'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		ELSEIF beta_agency = "" or beta_agency = True then							'If you're a beta agency, you should probably use the beta branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/BETA/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else																		'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		End if
		SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a FuncLib_URL
		req.open "GET", FuncLib_URL, FALSE							'Attempts to open the FuncLib_URL
		req.send													'Sends request
		IF req.Status = 200 THEN									'200 means great success
			Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
			Execute req.responseText								'Executes the script code
		ELSE														'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
			MsgBox 	"Something has gone wrong. The code stored on GitHub was not able to be reached." & vbCr &_ 
					vbCr & _
					"Before contacting Veronica Cary, please check to make sure you can load the main page at www.GitHub.com." & vbCr &_
					vbCr & _
					"If you can reach GitHub.com, but this script still does not work, ask an alpha user to contact Veronica Cary and provide the following information:" & vbCr &_
					vbTab & "- The name of the script you are running." & vbCr &_
					vbTab & "- Whether or not the script is ""erroring out"" for any other users." & vbCr &_
					vbTab & "- The name and email for an employee from your IT department," & vbCr & _
					vbTab & vbTab & "responsible for network issues." & vbCr &_
					vbTab & "- The URL indicated below (a screenshot should suffice)." & vbCr &_
					vbCr & _
					"Veronica will work with your IT department to try and solve this issue, if needed." & vbCr &_ 
					vbCr &_
					"URL: " & FuncLib_URL
					script_end_procedure("Script ended due to error connecting to GitHub.")
		END IF
	ELSE
		FuncLib_URL = "C:\BZS-FuncLib\MASTER FUNCTIONS LIBRARY.vbs"
		Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
		Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
		text_from_the_other_script = fso_command.ReadAll
		fso_command.Close
		Execute text_from_the_other_script
	END IF
END IF
'END OF FUNCLIB BLOCK======================================================================================

'DIALOG----------------------------------------------------------------------------------------------------
BeginDialog Enrollment_dlg, 0, 0, 266, 280, "Enrollment Information"
  EditBox 55, 25, 60, 15, PMI_number
  EditBox 205, 25, 25, 15, enrollment_month
  EditBox 230, 25, 25, 15, enrollment_year
  CheckBox 80, 75, 25, 10, "Yes", Insurance_yes
  CheckBox 80, 90, 25, 10, "Yes", Pregnant_yes
  CheckBox 80, 105, 30, 10, "Yes", Interpreter_yes
  DropListBox 170, 100, 85, 15, "Spanish - 01"+chr(9)+"Hmong - 02"+chr(9)+"Vietnamese - 03"+chr(9)+"Khmer - 04"+chr(9)+"Laotian - 05"+chr(9)+"Russian - 06"+chr(9)+"Somali - 07"+chr(9)+"ASL - 08"+chr(9)+"Arabic - 10"+chr(9)+"Serbo-Croatian - 11"+chr(9)+"Oromo - 12"+chr(9)+"Other - 98", Interpreter_type
  CheckBox 80, 120, 25, 10, "Yes", foster_care_yes
  EditBox 80, 135, 75, 15, Medical_clinic_code
  EditBox 80, 155, 75, 15, Dental_clinic_code
  DropListBox 100, 195, 120, 15, "Select one..."+chr(9)+"Blue Plus PMAP"+chr(9)+"Medica PMAP"+chr(9)+"Medica MSC plus"+chr(9)+"Medica MSC plus (with EW)"+chr(9)+"Medica MSHO"+chr(9)+"Hennepin Health PMAP"+chr(9)+"Health Partners MSC plus"+chr(9)+"Health Partners MSC plus (with EW)"+chr(9)+"Health Partners MSHO"+chr(9)+"UCare MSC plus"+chr(9)+"UCare MSC plus (with EW)"+chr(9)+"UCare MSHO", Health_plan
  DropListBox 100, 215, 120, 15, "Select one..."+chr(9)+"First year change option"+chr(9)+"Health plan contract end"+chr(9)+"Initial enrollment"+chr(9)+"Move"+chr(9)+"Ninety Day change option"+chr(9)+"Open enrollment"+chr(9)+"PMI merge"+chr(9)+"Reenrollment", change_reason
  DropListBox 100, 235, 120, 15, "Select one..."+chr(9)+"Eligibility ended"+chr(9)+"Exclusion"+chr(9)+"First year change option"+chr(9)+"Health plan contract end"+chr(9)+"Jail - Incarceration"+chr(9)+"Move"+chr(9)+"Loss of disability"+chr(9)+"Ninety Day change option"+chr(9)+"Open Enrollment"+chr(9)+"PMI merge"+chr(9)+"Voluntary", disenrollment_reason
  ButtonGroup ButtonPressed
    OkButton 150, 260, 50, 15
    CancelButton 205, 260, 50, 15
  Text 10, 90, 40, 10, "Pregnant?"
  Text 10, 140, 70, 10, "Medical Clinic Code:"
  Text 10, 160, 65, 10, "Dental Clinic Code:"
  Text 10, 105, 45, 10, "Interpreter?"
  Text 110, 105, 55, 10, "If so what type?"
  Text 125, 30, 80, 10, "Enrollment Month/Year:"
  Text 10, 120, 50, 10, "Foster Care?"
  GroupBox 0, 55, 260, 120, "REFM questions (will enter no if nothing is selected)"
  GroupBox 0, 10, 260, 35, "*Note: You do not need enter zeros before the client's PMI number."
  Text 10, 75, 60, 10, "Other Insurance?"
  Text 35, 200, 40, 10, "Health plan:"
  Text 35, 220, 55, 10, "Change reason:"
  Text 10, 30, 45, 10, "PMI Number:"
  ButtonGroup ButtonPressed
    OkButton 155, 330, 50, 15
  Text 35, 240, 60, 10, "Disenroll reason:"
  GroupBox 0, 180, 260, 75, "RPPH information"
EndDialog

BeginDialog correct_pmi_check, 0, 0, 246, 60, "PMI check"
  ButtonGroup ButtonPressed
    OkButton 70, 40, 50, 15
    CancelButton 125, 40, 50, 15
  Text 10, 15, 225, 20, "Please verify that the PMI and client are correct then click OK. If the PMI was entered incorrectl then hit cancel, and start the script again. "
EndDialog

BeginDialog correct_REFM_check, 0, 0, 261, 65, "REFM check"
  ButtonGroup ButtonPressed
    OkButton 75, 40, 50, 15
    CancelButton 130, 40, 50, 15
  Text 15, 15, 235, 20, "Please verify that the information entered is correct then click OK. If the information was entered incorrectly hit cancel and start the script again. "
EndDialog

'Custom function----------------------------------------------------------------------------------------------------
'Sending MMIS back to the beginning screen and checking for a password prompt
Function check_for_MMIS(end_script) 'if end_script is set to true, the script will end; if set to false, script will continue once password is entered
	Do
		transmit
		row = 1
		col = 1
		EMSearch "MMIS", row, col
		IF row <> 1 then
			If end_script = True then 
				script_end_procedure("You do not appear to be in MMIS. You may be passworded out. Please check your MMIS screen and try again.")
			Else
				warning_box = MsgBox("You do not appear to be in MMIS. You may be passworded out. Please check your MMIS screen and try again, or press ""cancel"" to exit the script.", vbOKCancel)
				If warning_box = vbCancel then stopscript
			End if
		End if
	Loop until row = 1
End function

'SCRIPT----------------------------------------------------------------------------------------------------
EMConnect ""
call check_for_MMIS(False) 'Sending MMIS back to the beginning screen and checking for a password prompt
EMReadScreen MMIS_panel_check, 4, 1, 52	'checking to see if user is on the RKEY panel in MMIS. If not, then it will go to there.
IF MMIS_panel_check <> "RKEY" THEN 
	DO
		PF6
		EMReadScreen session_terminated_check, 18, 1, 7
	LOOP until session_terminated_check = "SESSION TERMINATED"
	'Getting back in to MMIS and transmitting past the warning screen (workers should already have accepted the warning screen when they logged themself into MMIS the first time!)
	EMWriteScreen "mw00", 1, 2
	transmit
	transmit
	EMWriteScreen "x", 8, 3
	transmit
END IF 

'date calculations
enrollment_month = Dateadd("m", + 1, date)
enrollment_year = DatePart("yyyy", enrollment_month)
enrollment_month = DatePart("m", enrollment_month)
enrollment_year = "" & enrollment_year - 2000
enrollment_month = CStr(enrollment_month)	'fix for changing a 1-digit month (1-9), to a 2-digit month (10-12)

'grabs the PMI number if one is listed on RKEY
EMReadscreen PMI_number, 8, 4, 19
PMI_number= trim(PMI_number)


'do the dialog here
Do
	Do
		Dialog Enrollment_dlg
		cancel_confirmation
		If pmi_number = "" then MsgBox "You must have a PMI number to continue!"
		If health_plan = "Select one..." then MsgBox " You must select a health plan."
		If change_reason = "Select one..." then MsgBox " You must select a change reason."
		IF disenrollment_reason = "Select one..." then MsgBox " You must select a disenrollment reason."
		If Interpreter_yes = 1 and Interpreter_type = "Select one..." then MsgBox "You must select an interpreter language."
	Loop until Interpreter_yes = 0 or (Interpreter_yes = 1 and Interpreter_type <> "Select one...")
Loop until (PMI_number <> "" and health_plan <> "Select one..." and change_reason <> "Select one..." and disenrollment_reason <> "Select one...")

'checking for an active MMIS session
Call check_for_MMIS(False)

'formatting variables----------------------------------------------------------------------------------------------------
If len(enrollment_month) = 1 THEN enrollment_month = "0" & enrollment_month
IF len(enrollment_year) <> 2 THEN enrollment_year = right(enrollment_year, 2)
interpreter_type = right(interpreter_type, 2)

Do	'adds zeros to PMI number until number becomes 8 digits
  If len(PMI_number) < 8 then PMI_number = "0" & PMI_number
Loop until len(PMI_number) = 8

enrollment_date = enrollment_month & "/01/" & enrollment_year

'Coding to be inputed in the MMIS----------------------------------------------------------------------------------------------------
'Blue Plus plan
If health_plan = "Blue Plus PMAP" then 
	health_plan_code = "A065813800"
	Contract_code_part_one = "MA"
	Contract_code_part_two = "12"
END IF 
'Medica plans 
If health_plan = "Medica PMAP" then 
	health_plan_code = "A405713900"
	Contract_code_part_one = "MA"
	Contract_code_part_two = "12"
END IF
If health_plan = "Medica MSC plus" then 
	health_plan_code = "A405713900"
	Contract_code_part_one = "MA"
	Contract_code_part_two = "30"
END IF
If health_plan = "Medica MSC plus (with EW)" then 
	health_plan_code = "A405713900"
	Contract_code_part_one = "MA"
	Contract_code_part_two = "35"
END IF
If health_plan = "Medica MSHO" then 
	health_plan_code = "A405713900"
	Contract_code_part_one = "MA"
	Contract_code_part_two = "20"
END IF
'Hennepin Health plan
If health_plan = "Hennepin Health PMAP" then 
	health_plan_code = "A836618200"
	Contract_code_part_one = "MA"
	Contract_code_part_two = "12"
END IF
'Health Partners plans
If health_plan = "Health Partners MSC plus" then 
	health_plan_code = "A585713900"
	Contract_code_part_one = "MA"
	Contract_code_part_two = "30"
END IF
If health_plan = "Health Partners MSC plus (with EW)" then 
	health_plan_code = "A585713900"
	Contract_code_part_one = "MA"
	Contract_code_part_two = "35"
END IF
If health_plan = "Health Partners MSHO" then 
	health_plan_code = "A585713900"
	Contract_code_part_one = "MA"
	Contract_code_part_two = "20"
END IF
'UCARE plans
If health_plan = "UCare MSC plus" then 
	health_plan_code = "A565813600"
	Contract_code_part_one = "MA"
	Contract_code_part_two = "30"
END IF
If health_plan = "UCare MSC plus (with EW)" then 
	health_plan_code = "A565813600"
	Contract_code_part_one = "MA"
	Contract_code_part_two = "35"
END IF
If health_plan = "UCare MSHO" then 
	health_plan_code = "A565813600"
	Contract_code_part_one = "MA"
	Contract_code_part_two = "20"
END IF

'change reasons
If change_reason = "First year change option" then
	change_reason = "FY"
End if
If change_reason = "Health plan contract end" then
	change_reason = "HP"
End if
If change_reason = "Initial enrollment" then
	change_reason = "IN"
End if
If change_reason = "Move" then
	change_reason = "MV"
End if
If change_reason = "Ninety Day change option" then
	change_reason = "NT"
End if
If change_reason = "Open enrollment" then
	change_reason = "OE"
End if
If change_reason = "PMI merge" then
	change_reason = "PM"
End if
If change_reason = "Reenrollment" then
	change_reason = "RE"
END IF

'Disenrollment reasons
If disenrollment_reason = "Eligibility ended" then
	disenrollment_reason = "EE"
END IF
If disenrollment_reason = "Exclusion" then
	disenrollment_reason = "EX"
END IF
If disenrollment_reason = "First year change option" then
	disenrollment_reason = "FY"
END IF
If disenrollment_reason = "Health plan contract end" then
	disenrollment_reason = "HP"
END IF
If disenrollment_reason = "Jail - Incarceration" then
	disenrollment_reason = "JL"
END IF
If disenrollment_reason = "Move" then
	disenrollment_reason = "MV"
END IF
If disenrollment_reason = "Loss of disability" then
	disenrollment_reason = "ND"
END IF
If disenrollment_reason = "Ninety Day change option" then
	disenrollment_reason = "NT"
END IF
If disenrollment_reason = "Open Enrollment" then
	disenrollment_reason = "OE"
END IF
If disenrollment_reason = "PMI merge" then
	disenrollment_reason = "PM"
END IF
If disenrollment_reason = "Voluntary" then
	disenrollment_reason = "VL"
END IF

'REFM check box values
If insurance_yes = 1 then 
	insurance_yn = "y"
   else
	insurance_yn = "n"
end if
If pregnant_yes = 1 then
	pregnant_yn = "y"
   else
	pregnant_yn = "n"
end if
If interpreter_yes = 1 then
	interpreter_yn = "y"
   else
	interpreter_yn = "n"
end if
If foster_care_yes = 1 then
	foster_care_yn = "y"
   else
	foster_care_yn = "n"
end if
'End of MMIS coding==============================================

'Now we are in RKEY, and it navigates into the case, transmits, and makes sure we've moved to the next screen.
EMWriteScreen "c", 2, 19
EMWriteScreen PMI_number, 4, 19 
transmit
EMReadscreen RKEY_check, 4, 1, 52
If RKEY_check = "RKEY" then script_end_procedure("The listed PMI number was not found. Check your PMI number and try again.")

'Now it gets to RELG for member 01 of this case.
EMWriteScreen "rcin", 1, 8
transmit
EMWriteScreen "x", 11, 2
'check Rpol to see if there is other insurance available, if so worker processes manually
EMWriteScreen "rpol", 1, 8
transmit
'making sure script got to right panel
EMReadScreen RPOL_check, 4, 1, 52
If RPOL_check <> "RPOL" then script_end_procedure("The script was unable to navigate to RPOL process manually if needed.")
EMreadscreen policy_number, 1, 7, 8
if policy_number <> " " then 
	msgbox "This case has spans on RPOL. Please evaluate manually at this time."
	pf6
	stopscript
end if
EMWriteScreen "rpph", 1, 8
transmit
'making sure script got to right panel
EMReadScreen RPPH_check, 4, 1, 52
If RPPH_check <> "RPPH" then script_end_procedure("The script was unable to navigate to RPPH process manually if needed.")

'Grabs client's name
EMreadscreen client_first_name, 13, 3, 20
client_first_name  = replace(client_first_name, " ", "")
EMreadscreen client_last_name, 18, 3, 2
client_last_name  = replace(client_last_name, " ", "")
'clears and enters info for relg
Emreadscreen managed_care_span, 1, 13, 5
'resets to bottom of the span list. 
pf11

'Checks for exclusion code only deletes if YY or blank, if any other span entered it stops script.
EMReadscreen XCL_code, 2, 6, 2
If XCL_code = "YY" or XCL_code = "* " then
	EMSetCursor 6, 2
	EMSendKey "..."
Else
	MSGbox "There is an exclusion code other than YY. Please process manually."
	stopscript
End if

'enter enrollment date
EMsetcursor 13, 5
EMSendKey Enrollment_date
'enter managed care plan code
EMsetcursor 13, 23
EMSendKey Health_plan_code
'enter contract code
EMSetcursor 13, 34
EMSendkey contract_code_part_one
EMsetcursor 13, 37
EMSendkey contract_code_part_two
'enter change reason
EMsetcursor 13, 71
EMsendkey change_reason
'enter disenrollment reason
EMsetcursor 13, 75
EMsendkey disenrollment_reason
'Asks worker to make sure the script has entered into the right case and cancels out to RKEY if worker hits cancel to no save anything. 
Dialog correct_pmi_check
IF buttonpressed = 0 then
	pf6
	stopscript
End IF

'error handling to ensure that enrollment date and exclusion dates don't conflict
EMReadScreen RPPH_error_check, 4, 24, 2
If RPPH_error_check <> "    " then 
	MsgBox "The enrollment date you are entering is conflicting with the exclusion date.  Please review."
END IF 

'REFM screen
EMWriteScreen "refm", 1, 8
transmit
'making sure script got to right panel
EMReadScreen REFM_check, 4, 1, 52
If REFM_check <> "REFM" then script_end_procedure("The script was unable to navigate to REFM process manually if needed.")
'checks for edit after hitting transmit
Emreadscreen edit_check, 1, 24, 2
If edit_check <> " " then
	msgbox "There is an edit on this action. Please review the edit and proceed manually."
	stopscript
end if

'form rec'd
EMsetcursor 10, 16
EMSendkey "y"
'other insurance y/n
EMsetcursor 11, 18
EMsendkey insurance_yn
'preg y/n
EMsetcursor 12, 19
EMsendkey pregnant_yn
'interpreter y/n
EMsetcursor 13, 29
EMsendkey interpreter_yn
'interpreter type
if interpreter_type <> "" then
	EMsetcursor 13, 52
	EMsendKey interpreter_type
end if
'medical clinic code
EMsetcursor 19, 4
EMsendkey Medical_clinic_code
'dental clinic code if applicable
EMsetcursor 19, 24
EMsendkey Dental_clinic_code
'foster care y/n
EMsetcursor 21, 15
EMsendkey foster_care_yn

'Asks worker to make sure the script has entered the correct information and cancels out to RKEY if worker hits cancel to no save anything. 
Dialog correct_REFM_check
IF buttonpressed = 0 then
	msgbox "You have identified the REFM panel is not correct.  Please check REFM, and case note manually or run the script again."
	stopscript
End IF
transmit

'checks for edit after hitting transmit
EMReadscreen edit_check, 1, 24, 2 'checks for an inhibiting edit
If edit_check <> " " then
	msgbox "There is an edit on this action. Please review the edit and proceed manually."
	stopscript
end if

'Saves & case notes
pf3
EMWriteScreen "c", 2, 19
transmit
pf4
pf11
EMSendkey "***HMO Note*** " & Client_first_name & " " & client_last_name & " enrolled into " & health_plan & " " & Enrollment_date & " " & worker_signature
pf3
pf3

script_end_procedure("")