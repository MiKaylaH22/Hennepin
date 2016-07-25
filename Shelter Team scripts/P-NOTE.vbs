'LOADING GLOBAL VARIABLES
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("\\hcgg.fr.co.hennepin.mn.us\lobroot\HSPH\Team\Eligibility Support\Scripts\Script Files\SETTINGS - GLOBAL VARIABLES.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "P-NOTE.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 0         	'manual run time in seconds
STATS_denomination = "C"        'C is for each case
'END OF stats block=========================================================================================================

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

'DIALOGS-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
BeginDialog pnote_dialog, 0, 0, 321, 150, "P-NOTE"
  EditBox 60, 5, 60, 15, MAXIS_case_number
  EditBox 230, 5, 20, 15, number_nights
  EditBox 205, 20, 20, 15, number_tokens
  EditBox 235, 40, 75, 15, ACF_EA_dates
  EditBox 185, 65, 125, 15, reason_for_homelessness
  EditBox 55, 85, 255, 15, resolution
  EditBox 105, 105, 205, 15, other_notes
  EditBox 105, 130, 90, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 205, 130, 50, 15
    CancelButton 260, 130, 50, 15
  Text 10, 10, 45, 10, "Case number:"
  Text 40, 135, 60, 10, "Worker Signature:"
  Text 175, 45, 55, 10, " ACF/EA Dates:"
  Text 60, 110, 40, 10, "Other notes:"
  Text 10, 70, 175, 10, "Funds issued when client become Homeless due to:"
  Text 235, 25, 80, 10, "# bus tokens/bus cards"
  Text 260, 10, 65, 10, "# nights shelter"
  Text 10, 90, 45, 10, "Resolution"
EndDialog

'THE SCRIPT--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Connecting to BlueZone, grabbing case number
EMConnect ""
CALL MAXIS_case_number_finder(MAXIS_case_number)

'Running the initial dialog
DO
	DO
		err_msg = ""
		Dialog pnote_dialog
		cancel_confirmation
		If MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbNewLine & "* Enter a valid case number."
		If number_nights = "" then err_msg = err_msg & vbNewLine & "* Enter the nubmer of nights of shelter"
		If number_tokens = "" then err_msg = err_msg & vbNewLine & "* Enter the number of tokens or buscards"
		If ACF_EA_dates = "" then err_msg = err_msg & vbNewLine & "* Enter the ACF/EA dates."
		If reason_for_homelessness = "" then err_msg = err_msg & vbNewLine & "* Enter the reason for homelessness."
		If resolution = "" then err_msg = err_msg & vbNewLine & "* Enter the resolution."		
		If worker_signature = "" then err_msg = err_msg & vbNewLine & "* Enter your worker signature."		
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & "(enter NA in all fields that do not apply)" & vbNewLine & err_msg & vbNewLine
	LOOP until err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS						
Loop until are_we_passworded_out = false					'loops until user passwords back in					
		
'adding the case number 	
back_to_self
EMWriteScreen "________", 18, 43
EMWriteScreen MAXIS_case_number, 18, 43
EMWriteScreen CM_mo, 20, 43	'entering current footer month/year
EMWriteScreen CM_yr, 20, 46

'The case note'
start_a_blank_CASE_NOTE
Call write_variable_in_CASE_NOTE("### P-note at End of EA and ACF Shelter Stay ###")
Call write_bullet_and_variable_in_CASE_NOTE("nights shelter", number_nights)
Call write_bullet_and_variable_in_CASE_NOTE("tokens or bus cards", number_tokens)
Call write_bullet_and_variable_in_CASE_NOTE("ACF/EA Dates", ACF_EA_dates)
Call write_bullet_and_variable_in_CASE_NOTE("Funds issued when client become Homeless due to ", reason_for_homelessness)
Call write_bullet_and_variable_in_CASE_NOTE("Resolution ", resolution)
Call write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes)
Call write_variable_in_CASE_NOTE(worker_signature)
Call write_variable_in_CASE_NOTE("Hennepin County Shelter Team.")

script_end_procedure("")