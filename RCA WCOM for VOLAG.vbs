'Functions needed since we are not connecting to the Functions LIBRARY

function back_to_SELF()
'--- This function will return back to the 'SELF' menu or the MAXIS home menu 
'===== Keywords: MAXIS, SELF, navigate
  Do
    EMSendKey "<PF3>"
    EMWaitReady 0, 0
    EMReadScreen SELF_check, 4, 2, 50
  Loop until SELF_check = "SELF"
end function

function check_for_MAXIS(end_script)
'--- This function checks to ensure the user is in a MAXIS panel
'~~~~~ end_script: If end_script = TRUE the script will end. If end_script = FALSE, the user will be given the option to cancel the script, or manually navigate to a MAXIS screen.
'===== Keywords: MAXIS, production, script_end_procedure
	Do
		transmit
		EMReadScreen MAXIS_check, 5, 1, 39
		If MAXIS_check <> "MAXIS"  and MAXIS_check <> "AXIS " then
			If end_script = True then
				script_end_procedure("You do not appear to be in MAXIS. You may be passworded out. Please check your MAXIS screen and try again.")
			Else
				warning_box = MsgBox("You do not appear to be in MAXIS. You may be passworded out. Please check your MAXIS screen and try again, or press ""cancel"" to exit the script.", vbOKCancel)
				If warning_box = vbCancel then stopscript
			End if
		End if
	Loop until MAXIS_check = "MAXIS" or MAXIS_check = "AXIS "
end function

function check_for_password(are_we_passworded_out)
'--- This function checks to make sure a user is not passworded out. If they are, it allows the user to password back in. NEEDS TO BE ADDED INTO dialog DO...lOOPS
'~~~~~ are_we_passworded_out: When adding to dialog enter "Call check_for_password(are_we_passworded_out)", then Loop until are_we_passworded_out = false. Parameter will remain true if the user still needs to input password.
'===== Keywords: MAXIS, PRISM, password
	Transmit 'transmitting to see if the password screen appears
	Emreadscreen password_check, 8, 2, 33 'checking for the word password which will indicate you are passworded out
	If password_check = "PASSWORD" then 'If the word password is found then it will tell the worker and set the parameter to be true, otherwise it will be set to false.
		Msgbox "Are you passworded out? Press OK and the dialog will reappear. Once it does, you can enter your password."
		are_we_passworded_out = true
	Else
		are_we_passworded_out = false
	End If
end function

function navigate_to_MAXIS_screen(function_to_go_to, command_to_go_to)
'--- This function is to be used to navigate to a specific MAXIS screen
'~~~~~ function_to_go_to: needs to be MAXIS function like "STAT" or "REPT"
'~~~~~ command_to_go_to: needs to be MAXIS function like "WREG" or "ACTV"
'===== Keywords: MAXIS, navigate
  EMSendKey "<enter>"
  EMWaitReady 0, 0
  EMReadScreen MAXIS_check, 5, 1, 39
  If MAXIS_check = "MAXIS" or MAXIS_check = "AXIS " then
    EMReadScreen locked_panel, 23, 2, 30
    IF locked_panel = "Program History Display" then
	PF3 'Checks to see if on Program History panel - which does not allow the Command line to be updated
    END IF
    row = 1
    col = 1
    EMSearch "function: ", row, col
    If row <> 0 then
      EMReadScreen MAXIS_function, 4, row, col + 10
      EMReadScreen STAT_note_check, 4, 2, 45
      row = 1
      col = 1
      EMSearch "Case Nbr: ", row, col
      EMReadScreen current_case_number, 8, row, col + 10
      current_case_number = replace(current_case_number, "_", "")
      current_case_number = trim(current_case_number)
    End if
    If current_case_number = MAXIS_case_number and MAXIS_function = ucase(function_to_go_to) and STAT_note_check <> "NOTE" then
      row = 1
      col = 1
      EMSearch "Command: ", row, col
      EMWriteScreen command_to_go_to, row, col + 9
      EMSendKey "<enter>"
      EMWaitReady 0, 0
    Else
      Do
        EMSendKey "<PF3>"
        EMWaitReady 0, 0
        EMReadScreen SELF_check, 4, 2, 50
      Loop until SELF_check = "SELF"
      EMWriteScreen function_to_go_to, 16, 43
      EMWriteScreen "________", 18, 43
      EMWriteScreen MAXIS_case_number, 18, 43
      EMWriteScreen MAXIS_footer_month, 20, 43
      EMWriteScreen MAXIS_footer_year, 20, 46
      EMWriteScreen command_to_go_to, 21, 70
      EMSendKey "<enter>"
      EMWaitReady 0, 0
      EMReadScreen abended_check, 7, 9, 27
      If abended_check = "abended" then
        EMSendKey "<enter>"
        EMWaitReady 0, 0
      End if
	  EMReadScreen ERRR_check, 4, 2, 52			'Checking for the ERRR screen
	  If ERRR_check = "ERRR" then transmit		'If the ERRR screen is found, it transmits
    End if
  End if
end function

function transmit()
'--- This function sends or hits the transmit key. 
 '===== Keywords: MAXIS, MMIS, PRISM, transmit
  EMSendKey "<enter>"
  EMWaitReady 0, 0
end function

function PF3()
'--- This function sends or hits the PF3 key. 
 '===== Keywords: MAXIS, MMIS, PRISM, PF3
  EMSendKey "<PF3>"
  EMWaitReady 0, 0
end function

function PF4()
'--- This function sends or hits the PF4 key. 
 '===== Keywords: MAXIS, MMIS, PRISM, PF4
  EMSendKey "<PF4>"
  EMWaitReady 0, 0
end function

function PF8()
'--- This function sends or hits the PF9 key. 
 '===== Keywords: MAXIS, MMIS, PRISM, PF9
  EMSendKey "<PF8>"
  EMWaitReady 0, 0
end function

function PF9()
'--- This function sends or hits the PF9 key. 
 '===== Keywords: MAXIS, MMIS, PRISM, PF9
  EMSendKey "<PF9>"
  EMWaitReady 0, 0
end function

function write_variable_in_SPEC_MEMO(variable)
'--- This function writes a variable in SPEC/MEMO
'~~~~~ variable: information to be entered into SPEC/MEMO 
'===== Keywords: MAXIS, SPEC, MEMO
	EMGetCursor memo_row, memo_col						'Needs to get the row and col to start. Doesn't need to get it in the array function because that uses EMWriteScreen.
	memo_col = 15										'The memo col should always be 15 at this point, because it's the beginning. But, this will be dynamically recreated each time.
	'The following figures out if we need a new page
	Do
		EMReadScreen character_test, 1, memo_row, memo_col 	'Reads a single character at the memo row/col. If there's a character there, it needs to go down a row, and look again until there's nothing. It also needs to trigger these events if it's at or above row 18 (which means we're beyond memo range).
		If character_test <> " " or memo_row >= 18 then
			memo_row = memo_row + 1

			'If we get to row 18 (which can't be written to), it will go to the next page of the memo (PF8).
			If memo_row >= 18 then
				PF8
				memo_row = 3					'Resets this variable to 3
			End if
		End if
	Loop until character_test = " "

	'Each word becomes its own member of the array called variable_array.
	variable_array = split(variable, " ")

	For each word in variable_array
		'If the length of the word would go past col 74 (you can't write to col 74), it will kick it to the next line
		If len(word) + memo_col > 74 then
			memo_row = memo_row + 1
			memo_col = 15
		End if

		'If we get to row 18 (which can't be written to), it will go to the next page of the memo (PF8).
		If memo_row >= 18 then
			PF8
			memo_row = 3					'Resets this variable to 3
		End if

		'Writes the word and a space using EMWriteScreen
		EMWriteScreen word & " ", memo_row, memo_col

		'Increases memo_col the length of the word + 1 (for the space)
		memo_col = memo_col + (len(word) + 1)
	Next

	'After the array is processed, set the cursor on the following row, in col 15, so that the user can enter in information here (just like writing by hand).
	EMSetCursor memo_row + 1, 15
end function

'Dialog--------------------------------------------
BeginDialog rca_dialog, 0, 0, 131, 70, "RCA WCOM dialog"
  EditBox 65, 5, 55, 15, MAXIS_case_number
  EditBox 75, 25, 20, 15, MAXIS_footer_month
  EditBox 100, 25, 20, 15, MAXIS_footer_year
  ButtonGroup ButtonPressed
    OkButton 15, 50, 50, 15
    CancelButton 70, 50, 50, 15
  Text 10, 10, 55, 10, "Case Number: "
  Text 10, 30, 65, 10, "Footer month/year:"
EndDialog

'The script-------------------------------------
EMConnect ""

'the dialog
Do
	Do
  		err_msg = ""
  		Dialog rca_dialog
  		If ButtonPressed = 0 then stopscript
  		If MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbNewLine & "* Enter a valid case number."
  		If IsNumeric(MAXIS_footer_month) = False or len(MAXIS_footer_month) > 2 or len(MAXIS_footer_month) < 2 then err_msg = err_msg & vbNewLine & "* Enter a valid footer month."
  		If IsNumeric(MAXIS_footer_year) = False or len(MAXIS_footer_year) > 2 or len(MAXIS_footer_year) < 2 then err_msg = err_msg & vbNewLine & "* Enter a valid footer year."
  		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP UNTIL err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS						
Loop until are_we_passworded_out = false					'loops until user passwords back in		

'Navigating to the spec wcom screen
CALL Check_for_MAXIS(false)
									
back_to_self

Emwritescreen MAXIS_case_number, 18, 43
Emwritescreen MAXIS_footer_month, 20, 43
Emwritescreen MAXIS_footer_year, 20, 46

CALL navigate_to_MAXIS_screen("SPEC", "WCOM")

'Searching for waiting SNAP notice
wcom_row = 6
Do
	wcom_row = wcom_row + 1
	Emreadscreen program_type, 2, wcom_row, 26
	Emreadscreen print_status, 7, wcom_row, 71
	If program_type = "RC" then
		If print_status = "Waiting" then
			Emwritescreen "x", wcom_row, 13
			Transmit
			PF9
			Emreadscreen rca_wcom_exists, 3, 3, 15
			If rca_wcom_exists <> "   " then 
				Msgbox("It appears you already have a WCOM added to this notice. The script will now end.")
				stopscript 
			END IF 
			
			If program_type = "RC" AND print_status = "Waiting" then
				rca_wcom_writen = true
				'This will write if the notice is for SNAP only
				CALL write_variable_in_SPEC_MEMO("******************************************************")
				CALL write_variable_in_SPEC_MEMO("")
				CALL write_variable_in_SPEC_MEMO("As of March 1, 2017 the monthly RCA standard has increased by $110.00. If you are receiving benefits from the Supplemental Nutrition Assistance Program (SNAP), this increase may affect those benefits.")
				CALL write_variable_in_SPEC_MEMO("")
				CALL write_variable_in_SPEC_MEMO("******************************************************")
				PF4
				PF3
			End if
		End If
	End If
	If rca_wcom_writen = true then Exit Do
	If wcom_row = 17 then
		PF8
		Emreadscreen spec_edit_check, 6, 24, 2
		wcom_row = 6
	end if
	If spec_edit_check = "NOTICE" THEN no_rca_waiting = true
Loop until spec_edit_check = "NOTICE"

If no_rca_waiting = true then 
	msgbox("No waiting RCA notice was found for the requested month")
	stopscript
Else 
	msgbox("WCOM has been added to the first found waiting RCA notice for the month and case selected. Please review the notice.")
	stopscript
END IF 