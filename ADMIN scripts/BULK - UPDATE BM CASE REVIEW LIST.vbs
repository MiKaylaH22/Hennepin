Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("C:\DHS-MAXIS-Scripts\Script Files\SETTINGS - GLOBAL VARIABLES.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "BULK - UPDATE BM CASE REVIEW LIST.vbs" 'BULK script that creates a list of cases that require an interview, and the contact phone numbers'
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

BeginDialog update_banked_month_status_dialog, 0, 0, 191, 90, "Dialog"
  DropListBox 80, 10, 105, 15, "Select one..."+chr(9)+"January"+chr(9)+"February"+chr(9)+"March"+chr(9)+"April"+chr(9)+"May"+chr(9)+"June"+chr(9)+"July"+chr(9)+"August"+chr(9)+"September"+chr(9)+"October"+chr(9)+"November"+chr(9)+"December", month_selection
  EditBox 80, 30, 30, 15, excel_row_to_start
  ButtonGroup ButtonPressed
    OkButton 80, 65, 50, 15
    CancelButton 135, 65, 50, 15
  Text 15, 35, 60, 10, "Excel row to start:"
  Text 10, 50, 160, 10, "(NOTE: the 1st row is row 2 (the header is row 1)."
  Text 5, 15, 70, 10, "Update status month:"
EndDialog


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
	
'DISPLAYS DIALOG
DO
	DO
		err_msg = ""
		Dialog update_banked_month_status_dialog
		If ButtonPressed = 0 then StopScript
		If month_selection = "Select one..." then err_msg = err_msg & vbNewLine & "* Please select the status month to update."
		If isNumeric(excel_row_to_start) = False then err_msg = err_msg & vbNewLine & "* Please select the excel row to start."
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
Select Case month_selection
Case "January"
	MAXIS_footer_month = "01"
	MAXIS_footer_year = "16"
	excel_col = 5
Case "February"
	MAXIS_footer_month = "02"
	MAXIS_footer_year = "16"
	excel_col = 6
Case "March"
	MAXIS_footer_month = "03"
	MAXIS_footer_year = "16"
	excel_col = 7
Case "April"
	MAXIS_footer_month = "04"
	MAXIS_footer_year = "16"
	excel_col = 8
Case "May"
	MAXIS_footer_month = "05"
	MAXIS_footer_year = "16"
	excel_col = 9
Case "June"
	MAXIS_footer_month = "06"
	MAXIS_footer_year = "16"
	excel_col = 10
Case "July"
	MAXIS_footer_month = "07"
	MAXIS_footer_year = "16"
	excel_col = 11
Case "August"
	MAXIS_footer_month = "08"
	MAXIS_footer_year = "16"
	excel_col = 12
Case "September"
	MAXIS_footer_month = "09"
	MAXIS_footer_year = "16"
	excel_col = 13
Case "October"
	MAXIS_footer_month = "10"
	MAXIS_footer_year = "16"
	excel_col = 14
Case "November"
	MAXIS_footer_month = "11"
	MAXIS_footer_year = "16"
	excel_col = 15
Case "December"
	MAXIS_footer_month = "12"
	MAXIS_footer_year = "16"
	excel_col = 16
End Select

excel_row = excel_row_to_start
DO  
    'Grabs the case number
	MAXIS_case_number = objExcel.cells(excel_row, 3).value
    If MAXIS_case_number = "" then exit do
	back_to_self
	EMWriteScreen "________", 18, 43
	EMWriteScreen MAXIS_case_number, 18, 43
	
    Call navigate_to_MAXIS_screen("CASE", "CURR")
    EMReadScreen CURR_panel_check, 4, 2, 55
	If CURR_panel_check <> "CURR" then msgbox MAXIS_case_number & " cannot access CASE/CURR."
    
    EMReadScreen case_status, 8, 8, 9
    case_status = trim(case_status)
    If case_status = "INACTIVE" then 
        ObjExcel.Cells(excel_row, excel_col).Value = "Inactive"
	Elseif case_status = "ACTIVE" then 
        MAXIS_row = 9
        Do 
            EMReadScreen prog_name, 4, MAXIS_row, 3
            prog_name = trim(prog_name)
            if prog_name = "" then exit do
            If prog_name = "FS" then 
                EMReadScreen case_status, 8, MAXIS_row, 9
                case_status = trim(case_status)
	            if case_status = "ACTIVE" then 
                    exit do
                ELSE 
                    MAXIS_row = MAXIS_row + 1
                END IF 
            Else
                MAXIS_row = MAXIS_row + 1
            END IF 
	    Loop until MAXIS_row = 17
        If prog_name <> "FS" then ObjExcel.Cells(excel_row, excel_col).Value = "Inactive"
    END If 

    MAXIS_case_number = ""
    excel_row = excel_row + 1
	STATS_counter = STATS_counter + 1
LOOP UNTIL objExcel.Cells(excel_row, 3).value = ""	'looping until the list of cases to check for recert is complete

STATS_counter = STATS_counter - 1 'removes one from the count since 1 is counted at the beginning (because counting :p)

script_end_procedure("Success! The Excel file now has been update for all inactive SNAP cases.")