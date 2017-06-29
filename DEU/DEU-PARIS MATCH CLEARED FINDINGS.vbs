name_of_script = "DEU-PARIS MATCH CLEARED FINDINGS.vbs"
start_time = timer
STATS_counter = 1              'sets the stats counter at one
STATS_manualtime = 90          'manual run time in seconds
STATS_denomination = "C"      'C is for each case
'END OF stats block=========================================================================================================

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_loCally = FALSE or run_loCally = "" THEN	   'If the scripts are set to run loCally, it skips this and uses an FSO below.
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

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: Call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
Call changelog_update("05/17/2017", "Initial version.", "MiKayla Handley")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'DIALOGS-------------------------------------------------------------
BeginDialog Paris_match_dialog, 0, 0, 211, 120, "NOTES-PARIS MATCH CLEARED FINDINGS"
  EditBox 60, 5, 55, 15, Maxis_Case_number
  DropListBox 155, 25, 50, 15, "Select One..."+chr(9)+"YES"+chr(9)+"NO", accessing_benefits_state_dropdown
  DropListBox 80, 40, 50, 15, "Select One..."+chr(9)+"YES"+chr(9)+"NO", Contact_other_state_dropdown
  EditBox 45, 55, 160, 15, findings_field
  EditBox 70, 75, 135, 15, verif_used
  ButtonGroup ButtonPressed
    OkButton 95, 100, 50, 15
    CancelButton 150, 100, 50, 15
  Text 5, 60, 35, 10, "Findings:"
  Text 5, 40, 70, 10, "Contacted other state"
  Text 5, 80, 60, 10, "Verification used:"
  Text 5, 25, 145, 10, "Is client accessing benefits in the other state"
  Text 5, 10, 45, 10, "Case number:"
EndDialog

'--THE SCRIPT----------------------------------------------------
EMConnect ""
Call MAXIS_case_number_finder(MAXIS_case_number)

Do
	DO
		Err_msg = ""
		Dialog Paris_match_dialog
		cancel_confirmation
			If Maxis_Case_number = "" or IsNumeric(case_number) = False or len(case_number) > 8 THEN err_msg = err_msg & vbNewLine & "*Please enter a valid case number"
			If accessing_benefits_state_dropdownn = "Select One..." THEN err_msg = err_msg & vbNewLine & "*Please select if client is accessing benefits in other state"
			If Contact_other_state_dropdown = "Select One..." THEN err_msg = err_msg & vbNewLine & "*Please select if agency has contacted other state"
			IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	Loop until err_msg = ""
	Call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = false

'Writing the case note to MAXIS---
Call start_a_blank_CASE_NOTE
Call write_variable_in_case_note("---PARIS MATCH CLEARED FINDINGS---" & paris_match_header)
Call write_bullet_and_variable_in_case_note("Client accessing benefits in the other state:", accessing_benefits_state_dropdown)
Call write_bullet_and_variable_in_case_note("Agency Contact other state:", Contact_other_state_dropdown )
Call write_bullet_and_variable_in_case_note("Findings:", findings_field)
Call write_bullet_and_variable_in_case_note("Verification used to clear the PARIS match:", verif_used)
Call write_variable_in_CASE_NOTE ("----- ----- ----- ----- ----- ----- -----")
Call write_variable_in_CASE_NOTE ("DEBT ESTABLISHMENT UNIT 612-348-4290 EXT 1-1-1")

script_end_procedure("")