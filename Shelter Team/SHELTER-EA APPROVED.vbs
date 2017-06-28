'LOADING GLOBAL VARIABLES
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("\\hcgg.fr.co.hennepin.mn.us\lobroot\HSPH\Team\Eligibility Support\Scripts\Script Files\SETTINGS - GLOBAL VARIABLES.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - EA APPROVED.vbs"
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
BeginDialog EA_approval_dialog, 0, 0, 316, 110, "EA Approved "
  EditBox 65, 10, 55, 15, MAXIS_case_number
  EditBox 200, 10, 100, 15, approval_dates
  DropListBox 65, 35, 95, 15, "Select one..."+chr(9)+"FMF "+chr(9)+"PSP"+chr(9)+"St. Anne's"+chr(9)+"The Drake", shelter_droplist
  CheckBox 170, 40, 140, 10, "Send mandatory vendor MEMO to client.", send_MEMO_checkbox
  EditBox 65, 60, 235, 15, other_notes
  EditBox 65, 85, 125, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 195, 85, 50, 15
    CancelButton 250, 85, 50, 15
  Text 135, 15, 65, 10, "EA approval dates:"
  Text 20, 65, 40, 10, "Other notes: "
  Text 15, 15, 45, 10, "Case number:"
  Text 15, 40, 45, 10, "Shelter name:"
  Text 5, 90, 60, 10, "Worker signature: "
EndDialog

'THE SCRIPT--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Connecting to BlueZone, grabbing case number
EMConnect ""
CALL MAXIS_case_number_finder(MAXIS_case_number)

'Running the initial dialog
DO
	DO
		err_msg = ""
		Dialog EA_approval_dialog
        cancel_confirmation
		IF len(case_number) > 8 or IsNumeric(case_number) = False THEN err_msg = err_msg & vbNewLine & "* Please enter a valid case number."
		IF approval_dates = "" then err_msg = err_msg & vbNewLine & "* Please enter the EA approval dates."
		If shelter_droplist = "Select one..." then err_msg = err_msg & vbNewLine & "* Please select the name of the shelter."
		IF worker_signature = "" then err_msg = err_msg & vbNewLine & "* Please enter your worker signature."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP UNTIL err_msg = ""
 Call check_for_password(are_we_passworded_out)
LOOP UNTIL check_for_password(are_we_passworded_out) = False
 
back_to_SELF
EMWriteScreen "________", 18, 43
EMWriteScreen case_number, 18, 43
EMWriteScreen CM_mo, 20, 43	'entering current footer month/year
EMWriteScreen CM_yr, 20, 46

IF send_MEMO_checkbox = 1 then 
	call navigate_to_MAXIS_screen("SPEC", "MEMO")		'Navigating to SPEC/MEMO
	'Creates a new MEMO. If it's unable the script will stop.
	PF5
	EMReadScreen memo_display_check, 12, 2, 33
	If memo_display_check = "Memo Display" then script_end_procedure("You are not able to go into update mode. Did you enter in inquiry by mistake? Please try again in production.")
	
	'Checking for an AREP. If there's an AREP it'll navigate to STAT/AREP, check to see if the forms go to the AREP. If they do, it'll write X's in those fields below.
	row = 4                             'Defining row and col for the search feature.
	col = 1
	EMSearch "ALTREP", row, col         'Row and col are variables which change from their above declarations if "ALTREP" string is found.
	IF row > 4 THEN                     'If it isn't 4, that means it was found.
		arep_row = row                                          'Logs the row it found the ALTREP string as arep_row
		call navigate_to_MAXIS_screen("STAT", "AREP")           'Navigates to STAT/AREP to check and see if forms go to the AREP
		EMReadscreen forms_to_arep, 1, 10, 45                   'Reads for the "Forms to AREP?" Y/N response on the panel.
		call navigate_to_MAXIS_screen("SPEC", "MEMO")           'Navigates back to SPEC/MEMO
		PF5                                                     'PF5s again to initiate the new memo process
	END IF
		
	'Checking for SWKR
	row = 4                             'Defining row and col for the search feature.
	col = 1
	EMSearch "SOCWKR", row, col         'Row and col are variables which change from their above declarations if "SOCWKR" string is found.
	IF row > 4 THEN                     'If it isn't 4, that means it was found.
		swkr_row = row                                          'Logs the row it found the SOCWKR string as swkr_row
		call navigate_to_MAXIS_screen("STAT", "SWKR")         'Navigates to STAT/SWKR to check and see if forms go to the SWKR
		EMReadscreen forms_to_swkr, 1, 15, 63                'Reads for the "Forms to SWKR?" Y/N response on the panel.
		call navigate_to_MAXIS_screen("SPEC", "MEMO")         'Navigates back to SPEC/MEMO
		PF5                                           'PF5s again to initiate the new memo process
	END IF
	EMWriteScreen "x", 5, 10                                        'Initiates new memo to client
	IF forms_to_arep = "Y" THEN EMWriteScreen "x", arep_row, 10     'If forms_to_arep was "Y" (see above) it puts an X on the row ALTREP was found.
	IF forms_to_swkr = "Y" THEN EMWriteScreen "x", swkr_row, 10     'If forms_to_arep was "Y" (see above) it puts an X on the row ALTREP was found.
	transmit  
	'writing the MEMO'
	Call write_variable_in_SPEC_MEMO("************************************************************")
	Call write_variable_in_SPEC_MEMO("Call 1 (888) 577-2227 to get more information or enroll. You will be placed on mandatory vendor because you have used shelter or have requested assistance for housing issues.")   
	Call write_variable_in_SPEC_MEMO(" ")
	Call write_variable_in_SPEC_MEMO("You will remain on mandatory vendor for 12 months. If you move, or your rent changes you must let your team know at least 15 days before the end of the month to make this change.")                                 
	Call write_variable_in_SPEC_MEMO(" ")
	Call write_variable_in_SPEC_MEMO("Call your Human Service Representative Team at the end of this 12 month period if you want them to stop vendoring your rent at that time. Budgeting classes are free to you and available through the Lutheran Social Services. If you have any questions call the Shelter Team at (612)-348-9410.")
	Call write_variable_in_SPEC_MEMO(" ")
	Call write_variable_in_SPEC_MEMO("Sincerely, Shelter Team")	
	Call write_variable_in_SPEC_MEMO("************************************************************")
	PF4
END IF

'The case note---------------------------------------------------------------------------------------
start_a_blank_case_note      'navigates to case/note and puts case/note into edit mode
Call write_variable_in_CASE_NOTE("### EA approved for: " & approval_dates & " for shelter stay at " & shelter_droplist & " ###")
If send_MEMO_checkbox = 1 then Call write_variable_in_CASE_NOTE("* Sent SPEC/MEMO to client re: mandatory vendoring for the next 12 months.")
Call write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes)
Call write_variable_in_CASE_NOTE ("---")
Call write_variable_in_CASE_NOTE (worker_signature)

script_end_procedure("")	