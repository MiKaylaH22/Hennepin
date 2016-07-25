'LOADING GLOBAL VARIABLES
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("\\hcgg.fr.co.hennepin.mn.us\lobroot\HSPH\Team\Eligibility Support\Scripts\Script Files\SETTINGS - GLOBAL VARIABLES.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - CLIENT SHELTERED BY WINDOW A.vbs"
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
BeginDialog client_sheltered_window_A, 0, 0, 301, 265, "Client Sheltered Window A"
  EditBox 75, 5, 65, 15, MAXIS_case_number
  EditBox 75, 30, 65, 15, client_housed_at
  EditBox 260, 30, 30, 15, nights_housed
  EditBox 105, 55, 35, 15, adults_vouchered
  EditBox 260, 55, 30, 15, children_vouchered
  EditBox 105, 80, 185, 15, reason_for_homelessness
  EditBox 70, 120, 220, 15, name_of_person_verifying
  EditBox 70, 140, 220, 15, relationship
  EditBox 70, 160, 220, 15, phone_number
  CheckBox 15, 190, 280, 10, "Informed client that they will need to see Rapid ReHousing Screener first, then see ", informed_client_checkbox
  EditBox 70, 215, 220, 15, other_notes
  EditBox 70, 240, 110, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 185, 240, 50, 15
    CancelButton 240, 240, 50, 15
  Text 20, 35, 55, 10, "Client housed at:"
  Text 25, 10, 45, 10, "Case number:"
  Text 35, 125, 25, 10, "Name:"
  Text 25, 145, 45, 10, "Relationship:"
  Text 15, 165, 50, 10, "Phone Number:"
  Text 25, 220, 40, 10, "Other notes:"
  Text 5, 245, 60, 10, "Worker Signature:"
  Text 5, 60, 100, 10, "Number of Aduilts vouchered:"
  Text 145, 35, 115, 10, "for the following number of nights:"
  Text 155, 60, 105, 10, "Number of Children vouchered:"
  GroupBox 10, 105, 285, 80, "Homelessness verified by contacting:"
  Text 15, 85, 90, 10, "Reason for homelessness:"
  Text 25, 200, 150, 10, "the Shelter team for interview and revoucher."
EndDialog

'THE SCRIPT--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Connecting to BlueZone, grabbing case number
EMConnect ""
CALL MAXIS_case_number_finder(MAXIS_case_number)

'Running the initial dialog
DO
	DO
		err_msg = ""
		Dialog client_sheltered_window_A
		cancel_confirmation
		If MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbNewLine & "* Enter a valid case number."
		If client_housed_at = "" then err_msg = err_msg & vbNewLine & "* Enter the name of the shelter where client(s) housed"		
		If nights_housed = "" then err_msg = err_msg & vbNewLine & "* Enter the number of nights clients housed"
		If adults_vouchered = "" then err_msg = err_msg & vbNewLine & "* Enter the number of adults vouchered"
		If children_vouchered = "" then err_msg = err_msg & vbNewLine & "* Enter the number of children vouchered"
		If reason_for_homelessness = "" then err_msg = err_msg & vbNewLine & "* Enter the reason for client's homelessness"
		If name_of_person_verifying = "" then err_msg = err_msg & vbNewLine & "* Enter the name of the person who verified client's homelessness"	
		If relationship = "" then err_msg = err_msg & vbNewLine & "* Enter the relationship to the client of the person who verified client's homelessness"
		If phone_number = "" then err_msg = err_msg & vbNewLine & "* Enter the phone number of the person who verified client's homelessness"
		If worker_signature = "" then err_msg = err_msg & vbNewLine & "* Enter your worker signature"		
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

Call write_variable_in_CASE_NOTE("### APPROVED FOR SHELTER BY " & worker_signature & " ###")
Call write_variable_in_CASE_NOTE("* Client has been housed at " & client_housed_at & " for " & nights_housed & " nights")
Call write_variable_in_CASE_NOTE("* Voucher keyed for " & adults_vouchered & " adults and " & children_vouchered & " children")
If informed_client_checkbox = 1 then Call write_variable_in_CASE_NOTE("* Client will need to see Rapid ReHousing Screener first, then see shelter team for interview and revoucher. ")
Call write_bullet_and_variable_in_CASE_NOTE("* Reason for client's homelessness", reason_for_homelessness)
Call write_variable_in_CASE_NOTE("---")
Call write_variable_in_CASE_NOTE("* Homelessness verified by contacting")
Call write_bullet_and_variable_in_CASE_NOTE("Name", name_of_person_verifying)
Call write_bullet_and_variable_in_CASE_NOTE("Relationship to the client", relationship)
Call write_bullet_and_variable_in_CASE_NOTE("Phone Number of person verifying client's homelessness", phone_number)
Call write_variable_in_CASE_NOTE("---")
Call write_variable_in_CASE_NOTE(worker_signature)

script_end_procedure("")