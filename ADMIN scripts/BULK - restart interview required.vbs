Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("C:\DHS-MAXIS-Scripts\Script Files\SETTINGS - GLOBAL VARIABLES.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "BULK - restart interview required.vbs" 'BULK script that creates a list of cases that require an interview, and the contact phone numbers'
start_time = timer

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

'Dialog----------------------------------------------------------------------------------------------------
BeginDialog excel_row_to_start_dialog, 0, 0, 156, 45, "Enter the excel row where the script should start:"
  EditBox 90, 5, 60, 15, excel_row_to_restart
  ButtonGroup ButtonPressed
    OkButton 45, 25, 50, 15
    CancelButton 100, 25, 50, 15
  Text 5, 10, 85, 10, "Excel row to restart from:"
EndDialog

'FUNCTIONS that are currently not in the FuncLib that are used in this script----------------------------------------------------------------------------------------------------
Function File_Selection_System_Dialog(file_selected)
    'Creates a Windows Script Host object
    Set wShell=CreateObject("WScript.Shell")

    'Creates an object which executes the "select a file" dialog, using a Microsoft HTML application (MSHTA.exe), and some handy-dandy HTML.
    Set oExec=wShell.Exec("mshta.exe ""about:<input type=file id=FILE><script>FILE.click();new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(1).WriteLine(FILE.value);close();resizeTo(0,0);</script>""")

    'Creates the file_selected variable from the exit
    file_selected = oExec.StdOut.ReadLine
End function

'-------THIS FUNCTION ALLOWS THE USER TO PICK AN EXCEL FILE---------
Function BrowseForFile()
    Dim shell : Set shell = CreateObject("Shell.Application")
    Dim file : Set file = shell.BrowseForFolder(0, "Choose a file:", &H4000, "Computer")
	IF file is Nothing THEN 
		script_end_procedure("The script will end.")
	ELSE
		BrowseForFile = file.self.Path
	END IF
End Function

'THE SCRIPT-------------------------------------------------------------------------------------------------------------------------
EMConnect ""		'Connects to BlueZone

'function to call up a local/network file
'Call File_Selection_System_Dialog(file_selected)

DO
	'file_location = InputBox("Please enter the file location.")
	Set objExcel = CreateObject("Excel.Application")
	Set objWorkbook = objExcel.Workbooks.Open(BrowseForFile)
	objExcel.Visible = True
	objExcel.DisplayAlerts = True
	
	confirm_file = MsgBox("Is this the correct file? Press YES to continue. Press NO to try again. Press CANCEL to stop the script.", vbYesNoCancel)
	IF confirm_file = vbCancel THEN 
		objWorkbook.Close
		objExcel.Quit
		stopscript
	ELSEIF confirm_file = vbNo THEN 
		objWorkbook.Close
		objExcel.Quit
	END IF
LOOP UNTIL confirm_file = vbYes

'DISPLAYS DIALOG
DO
	DO
		err_msg = ""
		Dialog excel_row_to_start_dialog
		If ButtonPressed = 0 then StopScript
		If excel_row_to_restart = "" or IsNumeric(excel_row_to_restart) = False then err_msg = err_msg & vbNewLine & "* Enter the excel row where the script should start."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP until err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS						
Loop until are_we_passworded_out = false					'loops until user passwords back in					
	

'resets the case number and footer month/year back to the CM (REVS for current month plus two has is going to be a problem otherwise)
back_to_self
EMwritescreen "________", 18, 43
EMWriteScreen CM_mo, 20, 43
EMWriteScreen CM_yr, 20, 46
transmit

'starts adding phone numbers at row selected
Excel_row = excel_row_to_restart
msgbox excel_row_to_restart
Do
	Do
		'Grabs the case number
		MAXIS_case_number = objExcel.cells(excel_row, 2).value
		Call navigate_to_MAXIS_screen("STAT", "ADDR")
		EMReadScreen ADDR_panel_check, 4, 2, 44
		If ADDR_panel_check <> "ADDR" then PF10 
	Loop until ADDR_panel_check = "ADDR"	
	EMReadScreen phone_number_one, 16, 17, 43	' if phone numbers are blank it doesn't add them to EXCEL
	If phone_number_one <> "( ___ ) ___ ____" then objExcel.cells(excel_row, 3).Value = phone_number_one
	EMReadScreen phone_number_two, 16, 18, 43
	If phone_number_two <> "( ___ ) ___ ____" then objExcel.cells(excel_row, 4).Value = phone_number_two
	EMReadScreen phone_number_three, 16, 19, 43
	If phone_number_three <> "( ___ ) ___ ____" then objExcel.cells(excel_row, 5).Value = phone_number_three	
	excel_row = excel_row + 1
	STATS_counter = STATS_counter + 1
LOOP UNTIL objExcel.Cells(excel_row, 2).value = ""	'looping until the list of cases to check for recert is complete

STATS_counter = STATS_counter - 1 'removes one from the count since 1 is counted at the beginning (because counting :p)
Msgbox STATS_counter
script_end_procedure("Success! The Excel file now has all of the cases that require interviews for renewals.  Please manually review the list of privileged cases (if any).")