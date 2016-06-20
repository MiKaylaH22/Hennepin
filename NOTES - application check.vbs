'OPTION EXPLICIT
name_of_script = "NOTES - APPLICATION CHECK.vbs"
start_time = timer

'DIM name_of_script, start_time, FuncLib_URL, run_locally, default_directory, beta_agency, req, fso, row, col

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			   'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else											'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		End if
		SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a FuncLib_URL
		req.open "GET", FuncLib_URL, FALSE							'Attempts to open the FuncLib_URL
		req.send													'Sends request
		IF req.Status = 200 THEN									'200 means great success
			Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
			Execute req.responseText								'Executes the script code
		ELSE														'Error message
			critical_error_msgbox = MsgBox ("Something has gone wrong. The Functions Library code stored on GitHub was not able to be reached." & vbNewLine & vbNewLine &_
                                            "FuncLib URL: " & FuncLib_URL & vbNewLine & vbNewLine &_
                                            "The script has stopped. Please check your Internet connection. Consult a scripts administrator with any questions.", _
                                            vbOKonly + vbCritical, "BlueZone Scripts Critical Error")
            StopScript
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
'END FUNCTIONS LIBRARY BLOCK================================================================================================

'Declaring variables---------------------------------------------------------------------------------------------------- 
'DIM application_check_dialog
'DIM application_check_type_dialog
'DIM mutiple_app_dates_dialog
'DIM application_check_date
'DIM case_number_dialog
'DIM outlook_calendar_check
'DIM pending_progs
'DIM app_date
'DIM ButtonPressed
'DIM case_number
'DIM MAXIS_footer_month
'DIM MAXIS_footer_year
'DIM application_type_droplist
'DIM application_status_droplist
'DIM application_check_droplist
'DIM actions_taken
'DIM other_app_notes
'DIM worker_signature
'DIM verifs_needed_check
'DIM approved_progs_check
'DIM denied_progs_check
'DIM documents_rec_check
'DIM PEND_CASH_check
'DIM PEND_EMER_check 
'DIM PEND_GRH_check
'DIM PEND_SNAP_check
'DIM PEND_HC_check
'DIM CASH_I_type_check
'DIM CASH_II_type_check
'DIM EMER_type_check
'DIM add_PEND_CASH_check
'DIM add_PEND_SNAP_check
'DIM add_PEND_HC_check 
'DIM add_PEND_EMER_check
'DIM add_PEND_GRH_check
'DIM app_month
'DIM app_day
'DIM app_year
'DIM pending_EMER_progs
'DIM pending_GRH_progs
'DIM pending_HC_progs
'DIM pending_SNAP_progs
'DIM pending_cash_progs
'DIM AREP_button
'DIM DISA_button
'DIM HCRE_button
'DIM HEST_button
'DIM JOBS_button
'DIM PROG_button
'DIM REVW_button
'DIM SHEL_button
'DIM UNEA_button
'DIM multiple_programs
'DIM MAXIS_row
'DIM MAXIS_col
'DIM pending_application_date
'DIM days_pending
'DIM not_pending_check
'DIM application_date
'DIM additional_application_check
'DIM additional_application_date
'DIM add_app_month 
'DIM add_app_day 
'DIM add_app_year
'DIM multiple_apps
'DIM additional_apps
'DIM date_found
'DIM additional_date_found

'DIALOGS----------------------------------------------------------------------------------------------------
BeginDialog case_number_dialog, 0, 0, 256, 95, "Case number dialog"
  EditBox 60, 10, 60, 15, case_number					
  DropListBox 185, 10, 60, 10, "Select one..."+chr(9)+"Day 1"+chr(9)+"Day 10"+chr(9)+"Day 20"+chr(9)+"Day 30"+chr(9)+"Day 45"+chr(9)+"Day 60"+chr(9)+"Over 60 days", application_check_droplist
  CheckBox 10, 35, 230, 10, "Check here to add application check date to your Outlook calendar.", outlook_calendar_check
  ButtonGroup ButtonPressed
    OkButton 140, 70, 50, 15
    CancelButton 195, 70, 50, 15
  Text 10, 15, 45, 10, "Case number: "
  Text 130, 15, 50, 10, "App check day:"
  Text 25, 50, 170, 10, "(Appointment time in Outlook will defult to 9:00 a.m.)"
EndDialog

BeginDialog application_check_dialog, 0, 0, 341, 165, "Application check dialog"
  DropListBox 190, 5, 145, 15, "Select one..."+chr(9)+"Apply MN"+chr(9)+"CAF"+chr(9)+"CAF addendum"+chr(9)+"HC - certain populations"+chr(9)+"HC - LTC"+chr(9)+"HC - EMA Mnsure ", application_type_droplist
  DropListBox 75, 30, 260, 15, "Select one..."+chr(9)+"Case is ready to approve or deny"+chr(9)+"No verifs rec'd yet (verification request has been sent)"+chr(9)+"Some verifs rec'd & more verification are needed"+chr(9)+"Other", application_status_droplist
  EditBox 75, 50, 260, 15, other_app_notes
  EditBox 75, 70, 260, 15, actions_taken
  EditBox 75, 145, 145, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 230, 145, 50, 15
    CancelButton 285, 145, 50, 15
    PushButton 195, 105, 30, 10, "AREP", AREP_button
    PushButton 230, 105, 30, 10, "DISA", DISA_button
    PushButton 265, 105, 30, 10, "HCRE", HCRE_button
    PushButton 300, 105, 30, 10, "JOBS", JOBS_button
    PushButton 195, 120, 30, 10, "PROG", PROG_button
    PushButton 230, 120, 30, 10, "REVW", REVW_button
    PushButton 265, 120, 30, 10, "SHEL", SHEL_button
    PushButton 300, 120, 30, 10, "UNEA", UNEA_button
  EditBox 90, 95, 90, 15, application_check_date
  EditBox 90, 115, 90, 15, pending_progs
  Text 15, 120, 70, 10, "Pending program(s):"
  Text 10, 35, 60, 10, "Application status:"
  Text 10, 55, 55, 10, "Other app notes:"
  Text 10, 150, 60, 10, "Worker signature:"
  Text 25, 100, 55, 10, "Application date:"
  Text 20, 75, 50, 10, "Actions taken:"
  Text 10, 10, 175, 10, "If Day 1 application check, select the application type:"
  GroupBox 190, 90, 145, 45, "MAXIS navigation"
EndDialog

BeginDialog mutiple_app_dates_dialog, 0, 0, 186, 80, "Multiple application dates exist"
  EditBox 60, 35, 60, 15, application_check_date
  ButtonGroup ButtonPressed
    OkButton 40, 60, 50, 15
    CancelButton 95, 60, 50, 15
  Text 10, 10, 170, 20, "Multiple application dates/pending programs exist.  Please select the application date to review."
EndDialog

'THE SCRIPT----------------------------------------------------------------------------------------------------
'Connecting to MAXIS & case number
EMConnect ""
Call MAXIS_case_number_finder(case_number)
Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)

'initial case number dialog
Do	
	dialog case_number_dialog
	if ButtonPressed = 0 Then StopScript
	If case_number = "" or IsNumeric(case_number) = False then MsgBox " You must enter a valid case number."
	If application_check_droplist = "Select one..." THEN Msgbox "You must enter the application check day."
LOOP UNTIL (IsNumeric(case_number) = TRUE AND application_check_droplist <> "Select one...")

'checking for an active MAXIS session
Call check_for_MAXIS(False)

'information gathering to auto-populate the application date
'pending programs information
back_to_self
EMWriteScreen case_number, 18, 43
Call navigate_to_MAXIS_screen("REPT", "PND2")

'grabs row and col number that the cursor is at 
EMGetCursor MAXIS_row, MAXIS_col
EMReadScreen not_pending_check, 5, 24, 2
EMReadScreen app_month, 2, MAXIS_row, 38
EMReadScreen app_day, 2, MAXIS_row, 41
EMReadScreen app_year, 2, MAXIS_row, 44
EMReadScreen days_pending, 3, MAXIS_row, 50
EMReadScreen PEND_CASH_check, 1, MAXIS_row, 54
EMReadScreen PEND_SNAP_check, 1, MAXIS_row, 62
EMReadScreen PEND_HC_check, 1, MAXIS_row, 65
EMReadScreen PEND_EMER_check, 1, MAXIS_row, 68
EMReadScreen PEND_GRH_check, 1, MAXIS_row, 72
EMReadScreen additional_application_check, 14, MAXIS_row + 1, 17
EMReadScreen add_app_month, 2, MAXIS_row + 1, 38
EMReadScreen add_app_day, 2, MAXIS_row + 1, 41
EMReadScreen add_app_year, 2, MAXIS_row + 1, 44
EMReadScreen add_PEND_CASH_check, 1, MAXIS_row + 1, 54
EMReadScreen add_PEND_SNAP_check, 1, MAXIS_row + 1, 62
EMReadScreen add_PEND_HC_check, 1, MAXIS_row  + 1, 65
EMReadScreen add_PEND_EMER_check, 1, MAXIS_row + 1, 68
EMReadScreen add_PEND_GRH_check, 1, MAXIS_row + 1, 72

'application date formatting
application_check_date = app_month & "/" & app_day & "/" & app_year

'this information auto-fills programs pending into main dialog if one app date is found
IF PEND_CASH_check = "A" or PEND_CASH_check = "P" THEN pending_progs = pending_progs & "CASH " & ", "
IF PEND_SNAP_check = "A" or PEND_SNAP_check = "P" THEN pending_progs = pending_progs & "SNAP " & ", "
IF PEND_HC_check = "A" or PEND_HC_check = "P" THEN pending_progs = pending_progs & "HC " & ", "
IF PEND_EMER_check = "A" or PEND_EMER_check = "P" THEN pending_progs = pending_progs & "EMER " & ", "
IF PEND_GRH_check = "A" or PEND_GRH_check = "P" THEN pending_progs = pending_progs & "GRH " & ", "
'trims excess spaces of pending_progs
pending_progs = trim(pending_progs)
'takes the last comma off of pending_progs when autofilled into dialog if more more than one app date is found and additional app is selected
If right(pending_progs, 1) = "," THEN pending_progs = left(pending_progs, len(pending_progs) - 1) 

'checking the case to make sure there is a pending case.  If not script will end & inform the user no pending case exists.
If not_pending_check = "CASE " THEN 
	Call navigate_to_MAXIS_screen("REPT", "PND1")
	EMReadScreen not_pending_check, 10, 24, 2
	If not_pending_check <> "NO PENDING" THEN 
		application_check_date = ""
	END IF
	If not_pending_check = "NO PENDING" THEN 
		script_end_procedure("There is not a pending program on this case." & vbNewLine & vbNewLine & "Please make sure you have the right case number, and/or check your case notes to ensure that this application has been completed.")
	END IF
END IF 

'additional application date formatting
additional_application_date = add_app_month & "/" & add_app_day & "/" & add_app_year

'checking for multiple application dates.  Creates message boxes giving the user an option of which app date to choose
If additional_application_check = "ADDITIONAL APP" THEN 
	multiple_apps = MsgBox("Do you want this application date:  " & application_date, VbYesNoCancel)
	If multiple_apps = 2 then stopscript
	If multiple_apps = 6 then 
		date_found = TRUE
	END IF 
	IF multiple_apps = 7 then 
		date_found = False
		additional_apps = Msgbox("Do you want this application date:  " & additional_application_date, VbYesNoCancel)
		If additional_apps = 2 then stopscript
		If additional_apps = 6 then 
			additional_date_found = TRUE 
		END IF
	END If
END if 

'formatting the application date for the main dialog if the additional application date is the date selected.
If date_found = TRUE THEN 
	application_check_date = application_date
END IF 

'this information auto-fills programs pending into main dialog
If additional_date_found = TRUE THEN
	application_check_date = additional_application_date
	IF add_PEND_CASH_check = "A" or add_PEND_CASH_check = "P" THEN pending_progs = pending_progs & "CASH " & ", "
	IF add_PEND_SNAP_check = "A" or add_PEND_SNAP_check = "P" THEN pending_progs = pending_progs & "SNAP " & ", "
	IF add_PEND_HC_check = "A" or add_PEND_HC_check = "P" THEN pending_progs = pending_progs & "HC " & ", "
	IF add_PEND_EMER_check = "A" or add_PEND_EMER_check = "P" THEN pending_progs = pending_progs & "EMER " & ", "
	IF add_PEND_GRH_check = "A" or add_PEND_GRH_check = "P" THEN pending_progs = pending_progs & "GRH " & ", "
	'trims excess spaces of pending_progs
	pending_progs = trim(pending_progs)
	'takes the last comma off of pending_progs when autofilled into dialog
	If right(pending_progs, 1) = "," THEN pending_progs = left(pending_progs, len(pending_progs) - 1) 
END IF 

'main dialog 
Do
	Do
		Do 	
			dialog application_check_dialog
			cancel_confirmation
			'function for navigation buttons on dialog
			MAXIS_dialog_navigation	
		Loop until ButtonPressed = -1
		If worker_signature = "" then MsgBox " You must sign your case note."
		If application_status_droplist = "Select one..." then MsgBox "You must enter the application status."
		IF actions_taken = "" then MsgBox "You must enter your case actions."	
	Loop until (worker_signature <> "" AND application_status_droplist <> "Select one..." and actions_taken <> "") 
	If application_status_droplist = "Other" AND other_app_notes = "" THEN MsgBox "You must enter more information about the 'other' application status."	
Loop until (application_status_droplist = "Other" AND other_app_notes <> "") or application_status_droplist <> "Other"	

'checking for an active MAXIS session
Call check_for_MAXIS(False)

'THE TIKL's----------------------------------------------------------------------------------------------------
'DAY 1 
If application_check_droplist = "Day 1" THEN
	If application_status_droplist = "Case is ready to approve or deny" THEN 
		Msgbox "You identified your case is ready to approve or deny.  A TIKL will NOT be made."
	ELSE 	
		Call navigate_to_MAXIS_screen("DAIL", "WRIT")
		call create_MAXIS_friendly_date(application_check_date, 10, 5, 18)
		Call write_variable_in_TIKL("Application check day 10. Please review case, and use the ""NOTES - APPLICATION CHECK"" script.")
		PF3
	END IF
END IF	

'DAY 10
If application_check_droplist = "Day 10" Then
	If application_status_droplist = "Case is ready to approve or deny" THEN 
		Msgbox "You identified your case is ready to approve or deny.  A TIKL will NOT be made."
	ELSE
		Call navigate_to_MAXIS_screen("DAIL", "WRIT")
		call create_MAXIS_friendly_date(application_check_date, 20, 5, 18)
		Call write_variable_in_TIKL("Application check day 20. Please review case, and use the ""NOTES - APPLICATION CHECK"" script.")
		PF3
	END IF 
END IF 

'DAY 20
If application_check_droplist = "Day 20" THEN
	If application_status_droplist = "Case is ready to approve or deny" THEN 
		Msgbox "You identified your case is ready to approve or deny.  A TIKL will NOT be made."
	ELSE
		Call navigate_to_MAXIS_screen("DAIL", "WRIT")
		call create_MAXIS_friendly_date(application_check_date, 30, 5, 18)
		Call write_variable_in_TIKL("Application check day 30. Please review case, and use the ""NOTES - APPLICATION CHECK"" script.")
		PF3
	END IF 
END IF 

'DAY 30
If application_check_droplist = "Day 30" THEN
	If application_status_droplist = "Case is ready to approve or deny" THEN 
		Msgbox "You identified your case is ready to approve or deny.  A TIKL will NOT be made."
	ELSE
		Call navigate_to_MAXIS_screen("DAIL", "WRIT")
		call create_MAXIS_friendly_date(application_check_date, 45, 5, 18)
		Call write_variable_in_TIKL("Application check day 45. Please review case, and use the ""NOTES - APPLICATION CHECK"" script.")
		PF3
	END IF  
END IF 

'DAY 45
If application_check_droplist = "Day 45" THEN
	If application_status_droplist = "Case is ready to approve or deny" THEN  
		Msgbox "You identified your case is ready to approve or deny.  A TIKL will NOT be made."
	ELSE
		Call navigate_to_MAXIS_screen("DAIL", "WRIT")
		call create_MAXIS_friendly_date(application_check_date, 60, 5, 18)
		Call write_variable_in_TIKL("Application check day 60. Please review case, and use the ""NOTES - APPLICATION CHECK"" script.")
		PF3
	END IF  
END IF 

'DAY 60
If application_check_droplist = "Day 60" THEN
	If application_status_droplist = "Case is ready to approve or deny" THEN 
	Msgbox "You identified your case is ready to approve or deny.  A TIKL will NOT be made."
	ELSE
		Call navigate_to_MAXIS_screen("DAIL", "WRIT")
		call create_MAXIS_friendly_date(application_check_date, 10, 5, 18)
		Call write_variable_in_TIKL("Application check day: over 60 days. Please review case, and use the ""NOTES - APPLICATION CHECK"" script.")
		PF3
	END IF  
END IF  

'OVER 60 DAYS 
If application_check_droplist = "Over 60 days" THEN
	If application_status_droplist = "Case is ready to approve or deny" THEN 
	Msgbox "You identified your case is ready to approve or deny.  A TIKL will NOT be made."
	ELSE
		Call navigate_to_MAXIS_screen("DAIL", "WRIT")
		call create_MAXIS_friendly_date(application_check_date, 10, 5, 18)
		Call write_variable_in_TIKL("Application check day: over 60 days. An additional 10 day period has been given to evaluate the case. Please review case, and use the ""NOTES - APPLICATION CHECK"" script.")
		PF3
	END IF  
END IF

'THE CASE NOTE----------------------------------------------------------------------------------------------------
start_a_blank_CASE_NOTE
Call write_variable_in_CASE_NOTE("-------------------------" & application_check_droplist & " application check")
Call write_bullet_and_variable_in_CASE_NOTE("Type of application rec'd", application_type_droplist)
Call write_bullet_and_variable_in_CASE_NOTE("Program applied for", pending_progs)
Call write_bullet_and_variable_in_CASE_NOTE("Application date", application_check_date)
Call write_variable_in_CASE_NOTE("---")
Call write_bullet_and_variable_in_CASE_NOTE("Application status", application_status_droplist)
Call write_bullet_and_variable_in_CASE_NOTE("Other application notes", other_app_notes)
Call write_bullet_and_variable_in_CASE_NOTE("Actions taken", actions_taken)
call write_variable_in_CASE_NOTE("---")
Call write_variable_in_CASE_NOTE(worker_signature)


'message boxes based on the application status chosen instructing workers which scripts to use next
If application_status_droplist = "Case is ready to approve or deny" Then 
	Msgbox "Success!  You have identified that the case is either ready to approve or deny." & vbNewLine & vbNewLine & _
	"If your case is ready to approve, please use the ""NOTES - APPROVED PROGRAMS"" script." & vbNewLine & vbNewLine & _
	"If your case is ready to be denied, please use the ""NOTES -DENIED PROGRAMS"" script."
ELSEIF application_status_droplist = "No verifs rec'd yet (verification request has been sent)" Then
	Msgbox "Success!  You have identified that no verifications have been received yet, and a verification request has been sent." & vbNewLine & vbNewLine & _
	"Please check to see that there is a verification requested case note, and if not, please use the ""NOTES - VERIFICATIONS REQUESTED"" script."
ELSEIF application_status_droplist = "Some verifs rec'd & more verification are needed" Then 
	Msgbox "Success!  You have identified that the your case has received some verifications, but others are needed." & vbNewLine & vbNewLine & _
	"Please check to see that the documents received have been case noted, as well as which verifications are still needed, and if a new verification request was sent." & vbNewLine & _
	"Please use the ""NOTES - DOCUMENTS RECEIVED"" script and/or the ""NOTES - VERIFICATIONS REQUESTED"" as needed."
END IF 

'commented out until we can figure out how to do this.
'run next script if worker selected one of the "script actions" check boxes
'IF verifs_needed_check = 1 THEN call run_from_GitHub(script_repository & "NOTES - VERIFICATIONS NEEDED.vbs")
'IF approved_progs_check = 1 THEN run_another_script(script_repository & "NOTES - APPROVED PROGRAMS.vbs")
'IF denied_progs_check = 1 THEN run_another_script(script_repository & "NOTES - DENIED PROGRAMS.vbs")
'IF documents_rec_check = 1 THEN run_another_script(script_repository & "NOTES - DOCUMENTS RECEIVED.vbs")

script_end_procedure("")