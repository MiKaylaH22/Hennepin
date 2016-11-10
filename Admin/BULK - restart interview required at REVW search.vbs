Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("C:\DHS-MAXIS-Scripts\Script Files\SETTINGS - GLOBAL VARIABLES.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "BULK - restart IR from REVW search.vbs" 'BULK script that creates a list of cases that require an interview, and the contact phone numbers'
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
BeginDialog excel_row_to_start_dialog, 0, 0, 141, 90, "Enter the excel row where the script should start:"
  EditBox 90, 5, 45, 15, excel_row_to_restart
  EditBox 90, 25, 20, 15, REPT_month
  EditBox 115, 25, 20, 15, REPT_year
  CheckBox 20, 45, 120, 10, "Add phone numbers to Excel list?", add_phone_numbers_check
  ButtonGroup ButtonPressed
    OkButton 30, 65, 50, 15
    CancelButton 85, 65, 50, 15
  Text 5, 10, 85, 10, "Excel row to restart from:"
  Text 25, 30, 60, 10, "REPT month/year:"
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
		If ButtonPressed = 0 then Script_end_procedure("")
		If REPT_month = "" or REPT_year = "" then err_msg = err_msg & vbNewLine & "*Enter a REPT month/year."
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

DO 'Loops until there are no more cases in the Excel list
	recert_status = ""
	'Grabs the case number
	MAXIS_case_number = objExcel.cells(excel_row, 2).value
	'Goes to STAT/REVW
	Call navigate_to_MAXIS_screen("STAT", "REVW")
	
	'Checking for PRIV cases.
	EMReadScreen priv_check, 6, 24, 14 'If it can't get into the case needs to skip
	IF priv_check = "PRIVIL" THEN 'Delete priv cases from excel sheet, save to a list for later
		priv_case_list = priv_case_list & "|" & MAXIS_case_number
		SET objRange = objExcel.Cells(excel_row, 1).EntireRow
		objRange.Delete	
	ELSE		'For all of the cases that aren't privileged...
        'Looks at review details
        EMReadScreen SNAP_review_check, 8, 9, 57
        IF SNAP_review_check = "__ 01 __" then 
			SET objRange = objExcel.Cells(excel_row, 1).EntireRow
			objRange.Delete				'all other cases that are not due for a recert will be deleted
			excel_row = excel_row - 1
        ELSE 
            EMwritescreen "x", 5, 58
            Transmit
		    DO
			    EMReadScreen SNAP_popup_check, 7, 5, 43
		    LOOP until SNAP_popup_check = "Reports"

		    'The script will now read the CSR MO/YR and the Recert MO/YR
		    EMReadScreen CSR_mo, 2, 9, 26
		    EMReadScreen CSR_yr, 2, 9, 32
		    EMReadScreen recert_mo, 2, 9, 64
		    EMReadScreen recert_yr, 2, 9, 70

	        'It then compares what it read to the previously established current month plus 2 and determines if it is a recert or not. If it is a recert we need an interview
		    IF CSR_mo = left(REPT_month, 2) and CSR_yr = right(REPT_year, 2) THEN recert_status = "NO"
            If recert_mo = left(REPT_month, 2) and recert_yr <> right(REPT_year, 2) THEN recert_status = "NO"
			IF recert_mo = left(REPT_month, 2) and recert_yr = right(REPT_year, 2) THEN recert_status = "YES"

			'If it's not a recert that requires an interview, delete it from the excel list and move on with our lives
			'this section ensure that only ACTIVE SNAP and MFIP cases have a review scheduled
			Call navigate_to_MAXIS_screen("STAT", "PROG")
			IF recert_status = "YES" then
			 	SNAP_status_check = ""
				EMReadScreen SNAP_status_check, 4, 10, 74
				If SNAP_status_check <> "ACTV" then 
					SET objRange = objExcel.Cells(excel_row, 1).EntireRow
					objRange.Delete				'all other cases that are not due for a recert will be deleted
					excel_row = excel_row - 1
				END IF 
			ELSEIF recert_status = "NO" then 
				MFIP_prog_check = ""
				MFIP_status_check = ""
				EMReadScreen MFIP_prog_check, 2, 6, 67		'checking for an active MFIP case
				EMReadScreen MFIP_status_check, 4, 6, 74
				If MFIP_prog_check = "MF" THEN
					IF MFIP_status_check <> "ACTV" THEN				'if MFIP is active, then case will not be deleted.
						SET objRange = objExcel.Cells(excel_row, 1).EntireRow
						objRange.Delete				'all other cases that are not due for a recert will be deleted
						excel_row = excel_row - 1
					END IF
				ELSE 
					SET objRange = objExcel.Cells(excel_row, 1).EntireRow
					objRange.Delete				'all other cases that are not due for a recert will be deleted
					excel_row = excel_row - 1
				END If
			END IF
			'handling for cases that do not have a completed HCRE panel
			PF3		'exits PROG to prommpt HCRE if HCRE insn't complete
			Do
				EMReadscreen HCRE_panel_check, 4, 2, 50
				If HCRE_panel_check = "HCRE" then 
					PF10	'exists edit mode in cases where HCRE isn't complete for a member
					PF3
				END IF
			Loop until HCRE_panel_check <> "HCRE"
		END IF
	END IF 	
	STATS_counter = STATS_counter + 1						'adds one instance to the stats counter
	excel_row = excel_row + 1
LOOP UNTIL objExcel.Cells(excel_row, 2).value = ""	'looping until the list of cases to check for recert is complete

'Creating the list of privileged cases and adding to the spreadsheet
prived_case_array = split(priv_case_list, "|")
excel_row = 2

FOR EACH MAXIS_case_number in prived_case_array
	objExcel.cells(excel_row, 6).value = MAXIS_case_number
	excel_row = excel_row + 1
NEXT

'If user selects to add phone numbers to the Excel list
IF add_phone_numbers_check = 1 then 
	Excel_row = 2
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
	LOOP UNTIL objExcel.Cells(excel_row, 2).value = ""	'looping until the list of cases to check for recert is complete
End if

'POST MAXIS ACTIONS----------------------------------------------------------------------------------------------------
'Query date/time/runtime info
'ObjExcel.Cells(1, 7).Value = "Query date and time:"	'Goes back one, as this is on the next row
'ObjExcel.Cells(1, 8).Value = now
'ObjExcel.Cells(2, 7).Value = "Query runtime (in seconds):"	'Goes back one, as this is on the next row
'ObjExcel.Cells(2, 8).Value = timer - query_start_time

FOR i = 1 to 8		'formatting the cells'
	objExcel.Cells(1, i).Font.Bold = True		'bold font'
	objExcel.Columns(i).AutoFit()						'sizing the columns'
NEXT

STATS_counter = STATS_counter - 1 'removes one from the count since 1 is counted at the beginning (because counting :p)

script_end_procedure("Success! The Excel file now has all of the cases that require interviews for renewals.  Please manually review the list of privileged cases (if any).")
