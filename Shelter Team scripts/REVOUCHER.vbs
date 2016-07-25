'LOADING GLOBAL VARIABLES
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("\\hcgg.fr.co.hennepin.mn.us\lobroot\HSPH\Team\Eligibility Support\Scripts\Script Files\SETTINGS - GLOBAL VARIABLES.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - REVOUCHER.vbs"
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
BeginDialog revoucher_dialog, 0, 0, 146, 75, "Select a revoucher option"
  EditBox 80, 10, 60, 15, MAXIS_case_number
  DropListBox 80, 30, 60, 10, "Select one.."+chr(9)+"Family"+chr(9)+"Single", revoucher_option
  ButtonGroup ButtonPressed
    OkButton 35, 50, 50, 15
    CancelButton 90, 50, 50, 15
  Text 15, 35, 60, 10, "Revoucher option:"
  Text 35, 15, 45, 10, "Case number:"
EndDialog

BeginDialog family_revoucher_dialog, 0, 0, 341, 330, "Family revoucher"
  DropListBox 55, 10, 60, 15, "Select one..."+chr(9)+"ACF"+chr(9)+"EA", voucher_type
  EditBox 195, 10, 55, 15, revoucher_date
  EditBox 305, 10, 25, 15, num_nights
  EditBox 55, 35, 110, 15, shelter_name
  EditBox 225, 35, 25, 15, children
  EditBox 305, 35, 25, 15, adults
  EditBox 55, 75, 275, 15, goal_one
  EditBox 55, 95, 275, 15, goal_two
  EditBox 55, 115, 275, 15, goal_three
  EditBox 55, 135, 275, 15, goal_four
  EditBox 55, 180, 275, 15, next_goal_one
  EditBox 55, 200, 275, 15, next_goal_two
  EditBox 55, 220, 275, 15, next_goal_three
  EditBox 55, 240, 275, 15, next_goal_four
  EditBox 100, 265, 230, 15, bus_issued
  EditBox 100, 285, 230, 15, other_notes
  EditBox 100, 305, 120, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 225, 305, 50, 15
    CancelButton 280, 305, 50, 15
  Text 265, 15, 40, 10, "# of nights:"
  Text 55, 290, 40, 10, "Other notes: "
  Text 180, 40, 45, 10, "# of Children:"
  GroupBox 5, 60, 330, 95, "Goals accomplished:"
  Text 15, 120, 40, 10, "Goal three:"
  Text 15, 140, 35, 10, "Goal four:"
  Text 15, 100, 35, 10, "Goal two:"
  Text 5, 15, 45, 10, "Voucher type:"
  Text 265, 40, 40, 10, "# of Adults:"
  Text 15, 80, 35, 10, "Goal one:"
  Text 15, 185, 35, 10, "Goal one:"
  Text 15, 205, 35, 10, "Goal two:"
  GroupBox 5, 165, 330, 95, "Goals for the next voucher:"
  Text 15, 225, 40, 10, "Goal three:"
  Text 15, 245, 35, 10, "Goal four:"
  Text 35, 310, 60, 10, "Worker signature: "
  Text 130, 15, 60, 10, "Date of revoucher:"
  Text 5, 40, 50, 10, "Shelter name:"
  Text 15, 270, 85, 10, "Bus tokens/cards issued:"
EndDialog

BeginDialog single_revoucher_dialog, 0, 0, 341, 345, "Single revoucher"
  DropListBox 55, 10, 60, 15, "Select one..."+chr(9)+"GA/GRH"+chr(9)+"O/C", voucher_type
  EditBox 195, 10, 55, 15, revoucher_date
  EditBox 300, 10, 30, 15, num_nights
  DropListBox 55, 35, 60, 15, "Select one..."+chr(9)+"PSP"+chr(9)+"SA-HL", shelter_type
  EditBox 210, 35, 120, 15, shelter_dates
  EditBox 55, 75, 275, 15, goal_one
  EditBox 55, 95, 275, 15, goal_two
  EditBox 55, 115, 275, 15, goal_three
  EditBox 55, 135, 275, 15, goal_four
  EditBox 55, 180, 275, 15, next_goal_one
  EditBox 55, 200, 275, 15, next_goal_two
  EditBox 55, 220, 275, 15, next_goal_three
  EditBox 55, 240, 275, 15, next_goal_four
  EditBox 100, 265, 230, 15, bus_issued
  EditBox 100, 285, 230, 15, other_notes
  EditBox 100, 305, 120, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 225, 305, 50, 15
    CancelButton 280, 305, 50, 15
  Text 55, 290, 40, 10, "Other notes: "
  Text 125, 40, 85, 10, "Dates shelter issued for:"
  GroupBox 5, 60, 330, 95, "Goals accomplished (18-21 YRS):"
  Text 15, 120, 40, 10, "Goal three:"
  Text 15, 140, 35, 10, "Goal four:"
  Text 15, 100, 35, 10, "Goal two:"
  Text 5, 15, 45, 10, "Voucher type:"
  Text 15, 80, 35, 10, "Goal one:"
  Text 15, 185, 35, 10, "Goal one:"
  Text 15, 205, 35, 10, "Goal two:"
  GroupBox 5, 165, 330, 95, "Goals for the next voucher:"
  Text 15, 225, 40, 10, "Goal three:"
  Text 15, 245, 35, 10, "Goal four:"
  Text 35, 310, 60, 10, "Worker signature: "
  Text 130, 15, 60, 10, "Date of revoucher:"
  Text 5, 40, 45, 10, "Shelter type:"
  Text 15, 270, 85, 10, "Bus tokens/cards issued:"
  Text 260, 15, 40, 10, "# of nights:"
EndDialog

'THE SCRIPT--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Connecting to BlueZone, grabbing case number
EMConnect ""
CALL MAXIS_case_number_finder(MAXIS_case_number)

'autofilling the date to the current Date
revoucher_date = date & ""

'Running the initial dialog
DO
	DO
		err_msg = ""
		Dialog revoucher_dialog
        cancel_confirmation
		IF len(case_number) > 8 or IsNumeric(case_number) = False THEN err_msg = err_msg & vbNewLine & "* Enter a valid case number."
		IF revoucher_option = "Select one..." then err_msg = err_msg & vbNewLine & "* Select a revoucher option."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP UNTIL err_msg = ""
 Call check_for_password(are_we_passworded_out)
LOOP UNTIL check_for_password(are_we_passworded_out) = False

If revoucher_option = "Family" then 
	DO
		DO
			err_msg = ""
			Dialog family_revoucher_dialog
			cancel_confirmation
			IF voucher_type = "Select one..." then err_msg = err_msg & vbNewLine & "* Select a voucher type."
			If isDate(revoucher_date) = False then err_msg = err_msg & vbNewLine & "* Enter the revoucher date."
			If IsNumeric(num_nights) = False then err_msg = err_msg & vbNewLine & "* Enter the number of nights issued."
			If shelter_name = "" then err_msg = err_msg & vbNewLine & "* Enter the shelter name."
			If IsNumeric(children) = False then err_msg = err_msg & vbNewLine & "* Enter the number of children."
			If IsNumeric(adults) = False then err_msg = err_msg & vbNewLine & "* Enter the number of adults."
			If goal_one = "" AND goal_two = "" AND goal_three = "" AND goal_four = "" then err_msg = err_msg & vbNewLine & "* Enter at least one goal accomplished."
			If next_goal_one = "" AND next_goal_two = "" AND next_goal_three = "" AND next_goal_four = "" then err_msg = err_msg & vbNewLine & "* Enter at least one goal for the next voucher." 
			If bus_issued = "" then err_msg = err_msg & vbNewLine & "* Enter information about bus cards/tokens issued." 
			If worker_signature = "" then err_msg = err_msg & vbNewLine & "* Enter your worker signature."
			IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
		LOOP UNTIL err_msg = ""
 	Call check_for_password(are_we_passworded_out)
	LOOP UNTIL check_for_password(are_we_passworded_out) = False
END IF 

If revoucher_option = "Single" then 
	DO
		DO
			err_msg = ""
			Dialog single_revoucher_dialog
			cancel_confirmation
			IF voucher_type = "Select one..." then err_msg = err_msg & vbNewLine & "* Select a voucher type."
			If isDate(revoucher_date) = False then err_msg = err_msg & vbNewLine & "* Enter the revoucher date."
			If IsNumeric(num_nights) = False then err_msg = err_msg & vbNewLine & "* Enter the number of nights issued."
			If shelter_type = "Select one..." then err_msg = err_msg & vbNewLine & "* Select the shelter type."
			If shelter_dates = "" then err_msg = err_msg & vbNewLine & "* Enter the dates of the shelter stay."
			If goal_one = "" AND goal_two = "" AND goal_three = "" AND goal_four = "" then err_msg = err_msg & vbNewLine & "* Enter at least one goal accomplished."
			If next_goal_one = "" AND next_goal_two = "" AND next_goal_three = "" AND next_goal_four = "" then err_msg = err_msg & vbNewLine & "* Enter at least one goal for the next voucher." 
			If bus_issued = "" then err_msg = err_msg & vbNewLine & "* Enter information about bus cards/tokens issued." 
			If worker_signature = "" then err_msg = err_msg & vbNewLine & "* Enter your worker signature."
			IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
		LOOP UNTIL err_msg = ""
 		Call check_for_password(are_we_passworded_out)
	LOOP UNTIL check_for_password(are_we_passworded_out) = False
END IF 

back_to_SELF
EMWriteScreen "________", 18, 43
EMWriteScreen case_number, 18, 43
EMWriteScreen CM_mo, 20, 43	'entering current footer month/year
EMWriteScreen CM_yr, 20, 46

'The case note---------------------------------------------------------------------------------------
start_a_blank_case_note      'navigates to case/note and puts case/note into edit mode
Call write_variable_in_CASE_NOTE("~~" & revoucher_option & " Revoucher~~")
Call write_bullet_and_variable_in_CASE_NOTE("Voucher type", voucher_type)
Call write_bullet_and_variable_in_CASE_NOTE("Date of revoucher", revoucher_date)
IF revoucher_option = "Family" then 
	Call write_variable_in_CASE_NOTE("* HH comp: " & adults & " adults, " & children & " children.")
	Call write_variable_in_CASE_NOTE("* Revoucher issued for " & shelter_name & " for " & num_nights & " nights.")
Else
	Call write_variable_in_CASE_NOTE("* Revoucher issued for " & voucher_type & " shelter for " & num_nights & " nights.")
END IF
IF revoucher_option = "Family" then 
	Call write_variable_in_CASE_NOTE("--Goals accomplished--")
Else 
	Call write_variable_in_CASE_NOTE("--Goals accomplished (18-21 yrs)--")
END IF
Call write_bullet_and_variable_in_CASE_NOTE("Goal one", goal_one)
Call write_bullet_and_variable_in_CASE_NOTE("Goal two", goal_two)
Call write_bullet_and_variable_in_CASE_NOTE("Goal three", goal_three) 
Call write_bullet_and_variable_in_CASE_NOTE("Goal four", goal_four)
Call write_variable_in_CASE_NOTE("---")
Call write_variable_in_CASE_NOTE("--Goals for the next voucher--")
Call write_bullet_and_variable_in_CASE_NOTE("Goal one", next_goal_one)
Call write_bullet_and_variable_in_CASE_NOTE("Goal two", next_goal_two)
Call write_bullet_and_variable_in_CASE_NOTE("Goal three", next_goal_three) 
Call write_bullet_and_variable_in_CASE_NOTE("Goal four", next_goal_four)
Call write_variable_in_CASE_NOTE ("---")
Call write_bullet_and_variable_in_CASE_NOTE("Bus tickets/tokens issued", bus_issued)
Call write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes)
Call write_variable_in_CASE_NOTE ("---")
Call write_variable_in_CASE_NOTE (worker_signature)

script_end_procedure("")