Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("C:\DHS-MAXIS-Scripts\Script Files\SETTINGS - GLOBAL VARIABLES.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "BULK - AUTO DIALER CASE STATUS.vbs" 'BULK script that creates a list of cases that require an interview, and the contact phone numbers'
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

FUNCTION get_case_status	
	back_to_self
	EMWriteScreen "________", 18, 43
	EMWriteScreen MAXIS_case_number, 18, 43
	
	Call navigate_to_MAXIS_screen("CASE", "CURR")
	EMReadScreen CURR_panel_check, 4, 2, 55
	If CURR_panel_check <> "CURR" then msgbox MAXIS_case_number & " cannot access CASE/CURR."
	
	EMReadScreen case_status, 8, 8, 9
	case_status = trim(case_status)
	ObjExcel.Cells(excel_row, 2).Value = case_status
	MAXIS_case_number = ""
	excel_row = excel_row + 1
	STATS_counter = STATS_counter + 1
END FUNCTION

'using new variable count to calculate percentages
IF case_status = "ACTIVE" then active_status = active_status + 1
IF case_status = "APP CLOS" then app_close_status = app_close_status + 1
IF case_status = "APP OPEN" then app_open_status = app_open_status + 1
IF case_status = "INACTIVE" then inactive_status = inactive_status + 1
IF case_status = "REIN" then rein_status = rein_status + 1

BeginDialog excel_row_to_start_dialog, 0, 0, 156, 45, "Enter the excel row where the script should start:"
  EditBox 90, 5, 60, 15, excel_row_to_start
  ButtonGroup ButtonPressed
    OkButton 45, 25, 50, 15
    CancelButton 100, 25, 50, 15
  Text 5, 10, 85, 10, "Excel row to restart from:"
EndDialog

Do 	
	Do
		dialog excel_row_to_start_dialog
		If ButtonPressed = 0 then stopscript
		If isnumeric(excel_row_to_start) = False then msgbox "Enter a valid numeric row to start."
	Loop until isnumeric(excel_row_to_start) = True
	call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = false
 
'resets the case number and footer month/year back to the CM (REVS for current month plus two has is going to be a problem otherwise)
back_to_self
EMwritescreen "________", 18, 43
EMWriteScreen CM_mo, 20, 43
EMWriteScreen CM_yr, 20, 46
transmit

'Gathering case status for answered call cases
'ObjExcel.ActiveSheet.Name = "Answer"
objExcel.worksheets("Answer").Activate
'Zeroing out variables
stats_counter = 0
active_status = 0
app_close_status = 0
app_open_status = 0
inactive_status = 0 
rein_status = 0 

excel_row = excel_row_to_start
Do 
	'Grabs the case number
	MAXIS_case_number = objExcel.cells(excel_row, 1).value
	If MAXIS_case_number = "" then exit do
	get_case_status
LOOP UNTIL objExcel.Cells(excel_row, 1).value = ""	'looping until the list of cases to check for recert is complete
STATS_counter = STATS_counter - 1 'removes one from the count since 1 is counted at the beginning (because counting :p)

'ObjExcel.Cells(1,4).Value = "=COUNTA(B2:B & abs(excel_row))"	'Excel formula
''ObjExcel.Cells(row_to_use, col_to_use).Value = "=COUNTA(" & SNAP_letter_col & ":" & SNAP_letter_col & ") - 1"	'Excel formula
'
'''=COUNTA(A2:A7)
'ObjExcel.Cells(1,1).Value = "Percentage of ACTIVE cases:"		'Row header
'ObjExcel.Cells(2,1).Value = "Percentage of APP CLOSED cases:"	'Row header
'ObjExcel.Cells(3,1).Value = "Percentage of APP OPEN cases:"		'Row header
'ObjExcel.Cells(4,1).Value = "Percentage of INACTIVE cases:"		'Row header
'ObjExcel.Cells(5,1).Value = "Percentage of REIN cases:"			'Row header
'
'objExcel.Cells(1,3).Font.Bold = TRUE							'Row header should be bold
'ObjExcel.Cells(1,3).NumberFormat = "0.00%"						'Formula should be percent
''Gathering case status for unanswered call cases
'objExcel.worksheets("No Answer").Activate	
''Zeroing out variables
'stats_counter = 0
'active_status = 0
'app_close_status = 0
'app_open_status = 0
'inactive_status = 0 
'rein_status = 0 
'
'excel_row = 2
'Do 
'	'Grabs the case number
'	MAXIS_case_number = objExcel.cells(excel_row, 1).value
'	If MAXIS_case_number = "" then exit do
'	get_case_status
'LOOP UNTIL objExcel.Cells(excel_row, 1).value = ""	'looping until the list of cases to check for recert is complete
'
script_end_procedure("Success! The Excel file now has been update for all inactive SNAP cases.")