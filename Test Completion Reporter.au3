#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_UseUpx=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#RequireAdmin
;#AutoIt3Wrapper_usex64=n
#include <File.au3>
#include <Array.au3>
#Include "Json.au3"
#include "Jira.au3"
#include "Confluence.au3"
#include "TestRail.au3"
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <GuiListView.au3>
#include <ButtonConstants.au3>
#include <ComboConstants.au3>
#include <SQLite.au3>
#include <SQLite.dll.au3>
#include <Crypt.au3>
#include <GuiComboBox.au3>

;$jira_json = FileRead("D:\dwn\fred.txt")
;ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $jira_json = ' & $jira_json & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
;$rr = StringRegExp($jira_json, '(?U)"key":"(.*)".*"summary":"(.*)".*"name":"(.*)"', 3)
;@error
;ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : @error = ' & @error & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
;Exit


;$rr = "<td>[Task for RMS-147] - Create an expandable \</td>"
;$tt = StringRegExpReplace($rr, "([^\\])\\([^""])", "$1\\\\$2")
;ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $tt = ' & $tt & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
;Exit
#cs
$tmp_json = FileRead("D:\dwn\fred2.txt")
Local $Data1 = Json_Decode($tmp_json)

;local $obj
;$uu = Json_ObjGetKeys($Data1)
;ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $uu = ' & $uu & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
;local $newobj = Json_ObjGet($Data1, 'issues')
;ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $newobj = ' & $newobj & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
;local $gg = Json_ObjGetCount($newobj)


;$yy = Json_IsObject($newobj)
;ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $yy = ' & $yy & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
;Exit

Local $issuelink = Json_Get($Data1, '.issues[0].fields.issuelinks[2].inwardIssue.key')
ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $issuelink = ' & $issuelink & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
Exit

for $i = 0 to 99999

	Local $key = Json_Get($Data1, '.issues[' & $i & '].key')

	if @error > 0 Then ExitLoop

	Local $summary = Json_Get($Data1, '.issues[' & $i & '].fields.summary')
	Local $reporter = Json_Get($Data1, '.issues[' & $i & '].fields.reporter.displayName')
	Local $fixVersions = Json_Get($Data1, '.issues[' & $i & '].fields.fixVersions[0].name')
	Local $resolution = Json_Get($Data1, '.issues[' & $i & '].fields.resolution.name')
	Local $ScrumTeam = Json_Get($Data1, '.issues[' & $i & '].fields.customfield_12000.value')
	Local $priority = Json_Get($Data1, '.issues[' & $i & '].fields.priority.name')
	Local $labels = Json_Get($Data1, '.issues[' & $i & '].fields.labels[0]')
	Local $AffectsVersions = Json_Get($Data1, '.issues[' & $i & '].fields.versions[0]')
	Local $status = Json_Get($Data1, '.issues[' & $i & '].fields.status.name')
	Local $customfield_10007 = Json_Get($Data1, '.issues[' & $i & '].fields.customfield_10007[0]')
	Local $environment = Json_Get($Data1, '.issues[' & $i & '].fields.environment')
	Local $assignee = Json_Get($Data1, '.issues[' & $i & '].fields.assignee.displayName')
	Local $Sprint = ""

	$query = "INSERT INTO Bug (Key,Summary,Reporter,Assignee,Status,Priority,AffectsVersions,FixVersions,Resolution,Labels,Environment,ScrumTeam,Sprint) VALUES ('" & $key & "','" & $summary & "','" & $reporter & "','" & $assignee & "','" & $status & "','" & $priority & "','" & $AffectsVersions & "','" & $fixVersions & "','" & $resolution & "','" & $labels & "','" & $environment & "','" & $ScrumTeam & "','" & $Sprint & "');"
	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $query = ' & $query & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
Next


Exit
#ce

;for $i = 0 to 99999

;	Local $user_id = Json_Get($tmp_json, '[' & $i & '].id')
;	Local $user_name = Json_Get($tmp_json, '[' & $i & '].name')

;	if @error > 0 Then ExitLoop

;	$user_dict.Add($user_id, $user_name)
;Next




Global $run_ids
Global $html
Global $app_name = "Test Completion Reporter"
Global $ini_filename = @ScriptDir & "\" & $app_name & ".ini"

Global $main_gui = GUICreate("TCR - " & $app_name, 860, 600)

GUICtrlCreateGroup("TestRail Setup", 10, 10, 410, 110)
GUICtrlCreateLabel("TestRail Username", 20, 30, 100, 20)
Global $testrail_username_input = GUICtrlCreateInput(IniRead($ini_filename, "main", "testrailusername", "sgriffin@janison.com"), 140, 30, 250, 20)
GUICtrlCreateLabel("TestRail Password", 20, 50, 100, 20)
Global $testrail_password_input = GUICtrlCreateInput("", 140, 50, 250, 20, $ES_PASSWORD)
GUICtrlCreateLabel("TestRail Project", 20, 70, 100, 20)
Global $testrail_project_combo = GUICtrlCreateCombo("", 140, 70, 250, 20, BitOR($CBS_DROPDOWNLIST, $WS_VSCROLL))
GUICtrlCreateLabel("TestRail Plan", 20, 90, 100, 20)
Global $testrail_plan_combo = GUICtrlCreateCombo("", 140, 90, 250, 20, BitOR($CBS_DROPDOWNLIST, $WS_VSCROLL))
GUICtrlCreateGroup("", -99, -99, 1, 1)

Local $testrail_encrypted_password = IniRead($ini_filename, "main", "testrailpassword", "")
Global $testrail_decrypted_password = ""

if stringlen($testrail_encrypted_password) > 0 Then

	$testrail_decrypted_password = _Crypt_DecryptData($testrail_encrypted_password, "applesauce", $CALG_AES_256)
	$testrail_decrypted_password = BinaryToString($testrail_decrypted_password)
	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $testrail_decrypted_password = ' & $testrail_decrypted_password & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
	GUICtrlSetData($testrail_password_input, $testrail_decrypted_password)
Else

	$testrail_decrypted_password = ""
EndIf

GUICtrlCreateGroup("Jira Setup", 440, 10, 410, 110)
GUICtrlCreateLabel("Jira Username", 450, 30, 100, 20)
Global $jira_username_input = GUICtrlCreateInput(IniRead($ini_filename, "main", "jirausername", "sgriffin@janison.com.au"), 570, 30, 250, 20)
GUICtrlCreateLabel("Jira Password", 450, 50, 100, 20)
Global $jira_password_input = GUICtrlCreateInput("", 570, 50, 250, 20, $ES_PASSWORD)
GUICtrlCreateGroup("", -99, -99, 1, 1)

Local $jira_encrypted_password = IniRead($ini_filename, "main", "jirapassword", "")
Global $jira_decrypted_password = ""

if stringlen($jira_encrypted_password) > 0 Then

	$jira_decrypted_password = _Crypt_DecryptData($jira_encrypted_password, "applesauce", $CALG_AES_256)
	$jira_decrypted_password = BinaryToString($jira_decrypted_password)
	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $jira_decrypted_password = ' & $jira_decrypted_password & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
	GUICtrlSetData($jira_password_input, $jira_decrypted_password)
Else

	$jira_decrypted_password = ""
EndIf

GUICtrlCreateLabel("Jira Project", 20, 130, 100, 20)
Global $epic_key_input = GUICtrlCreateInput(IniRead($ini_filename, "main", "epickeys", "RMS"), 140, 130, 680, 20)
Global $start_button = GUICtrlCreateButton("Start", 10, 160, 100, 20, -1, $BS_DEFPUSHBUTTON)
;GUICtrlSetState(-1, $GUI_DISABLE)
Global $display_report_button = GUICtrlCreateButton("Display Report", 120, 160, 100, 20)
;GUICtrlSetState(-1, $GUI_DISABLE)
Global $display_data_button = GUICtrlCreateButton("Display Data", 240, 160, 100, 20)
GUICtrlSetState(-1, $GUI_DISABLE)
Global $update_confluence_button = GUICtrlCreateButton("Update Confluence", 360, 160, 100, 20)
;GUICtrlSetState(-1, $GUI_DISABLE)

Global $listview = GUICtrlCreateListView("Epic Key|PID|Extract Status", 10, 200, 410, 300, $LVS_SHOWSELALWAYS)
_GUICtrlListView_SetColumnWidth(-1, 0, 200)
_GUICtrlListView_SetColumnWidth(-1, 1, 50)
_GUICtrlListView_SetColumnWidth(-1, 2, 200)
_GUICtrlListView_SetExtendedListViewStyle($listview, BitOR($LVS_EX_GRIDLINES, $LVS_EX_FULLROWSELECT, $LVS_EX_SUBITEMIMAGES))

Global $status_input = GUICtrlCreateInput("Enter the ""Epic Key"" and click ""Start""", 10, 600 - 25, 400, 20, $ES_READONLY, $WS_EX_STATICEDGE)
Global $progress = GUICtrlCreateProgress(420, 600 - 25, 400, 20)


GUISetState(@SW_SHOW, $main_gui)

; Startup SQLite

_SQLite_Startup()
ConsoleWrite("_SQLite_LibVersion=" & _SQLite_LibVersion() & @CRLF)

GUICtrlSetData($status_input, "")
GUIRegisterMsg($WM_COMMAND, "WM_COMMAND")

; Loop until the user exits.
While 1

	; GUI msg loop...
	$msg = GUIGetMsg()

	Switch $msg

		Case $GUI_EVENT_CLOSE

			IniWrite($ini_filename, "main", "testrailusername", GUICtrlRead($testrail_username_input))
			IniWrite($ini_filename, "main", "testrailproject", GUICtrlRead($testrail_project_combo))
			IniWrite($ini_filename, "main", "testrailplan", GUICtrlRead($testrail_plan_combo))
			IniWrite($ini_filename, "main", "jirausername", GUICtrlRead($jira_username_input))
			IniWrite($ini_filename, "main", "epickeys", GUICtrlRead($epic_key_input))

			$testrail_encrypted_password = _Crypt_EncryptData(GUICtrlRead($testrail_password_input), "applesauce", $CALG_AES_256)
			IniWrite($ini_filename, "main", "testrailpassword", $testrail_encrypted_password)

			$jira_encrypted_password = _Crypt_EncryptData(GUICtrlRead($jira_password_input), "applesauce", $CALG_AES_256)
			IniWrite($ini_filename, "main", "jirapassword", $jira_encrypted_password)


			ExitLoop

		Case $start_button

			FileDelete(@ScriptDir & "\" & $app_name & ".sqlite")
			_SQLite_Open(@ScriptDir & "\" & $app_name & ".sqlite")
			_SQLite_Exec(-1, "CREATE TABLE Epic (Key,Summary);") ; CREATE a Table
			_SQLite_Exec(-1, "CREATE TABLE Story (Key,Summary,EpicKey,ReqID,FixVersion,Status,TestNotes);") ; CREATE a Table
			_SQLite_Exec(-1, "CREATE TABLE Task (Key,Summary,StoryKey);") ; CREATE a Table
			_SQLite_Exec(-1, "CREATE TABLE Bug (Key,Summary,Reporter,Assignee,Status,Priority,AffectsVersions,FixVersions,Resolution,Labels,Environment,ScrumTeam,Sprint);") ; CREATE a Table
			_SQLite_Exec(-1, "CREATE TABLE TestCase (Id,Title,Reference);") ; CREATE a Table
			_SQLite_Exec(-1, "CREATE TABLE Run (Id,Name);") ; CREATE a Table
			_SQLite_Exec(-1, "CREATE TABLE Test (Id,Title,TestCaseId,RunId);") ; CREATE a Table
			_SQLite_Exec(-1, "CREATE TABLE Result (Id,TestId,StatusId,CreatedOn,Defects);") ; CREATE a Table
			_SQLite_Exec(-1, "CREATE TABLE TestStatus (Id,Label);") ; CREATE a Table

;			_SQLite_Exec(-1, "DELETE FROM Epic;") ; CREATE a Table
;			_SQLite_Exec(-1, "DELETE FROM Story;") ; CREATE a Table
;			_SQLite_Exec(-1, "DELETE FROM Task;") ; CREATE a Table

			GUICtrlSetData($progress, 0)
			GUICtrlSetState($epic_key_input, $GUI_DISABLE)
			GUICtrlSetState($start_button, $GUI_DISABLE)
			GUICtrlSetState($display_report_button, $GUI_DISABLE)
			GUICtrlSetState($display_data_button, $GUI_DISABLE)
			GUISetCursor(15, 1, $main_gui)
			_GUICtrlListView_DeleteAllItems($listview)

			$each = "blank"
			Local $pid = ShellExecute(@ScriptDir & "\data_extractor.exe", """" & GUICtrlRead($testrail_username_input) & """ """ & GUICtrlRead($testrail_password_input) & """ """ & $run_ids & """ """ & GUICtrlRead($jira_username_input) & """ """ & GUICtrlRead($jira_password_input) & """ """ & GUICtrlRead($epic_key_input) & """", "", "", @SW_HIDE)

#cs
			; populate listview with epic keys

			Local $epic_key = StringSplit(GUICtrlRead($epic_key_input), ",;|", 2)

			for $each in $epic_key

				Local $pid = ShellExecute(@ScriptDir & "\data_extractor.exe", """" & GUICtrlRead($testrail_username_input) & """ """ & GUICtrlRead($testrail_password_input) & """ """ & $run_ids & """ """ & GUICtrlRead($jira_username_input) & """ """ & GUICtrlRead($jira_password_input) & """ """ & $each & """", "", "", @SW_HIDE)
				GUICtrlCreateListViewItem($each & "|" & $pid & "|In Progress", $listview)
			Next

			While True

				Local $all_epics_done = True

				for $index = 0 to (_GUICtrlListView_GetItemCount($listview) - 1)

					Local $pid = _GUICtrlListView_GetItemText($listview, $index, 1)
					Local $status = _GUICtrlListView_GetItemText($listview, $index, 2)

					if StringCompare($status, "In Progress") = 0 Then

						$all_epics_done = False

						if ProcessExists($pid) = False Then

							_GUICtrlListView_SetItemText($listview, $index, "Done", 2)

						EndIf
					EndIf
				Next

				if $all_epics_done = True Then

					ExitLoop
				EndIf

				Sleep(1000)
			WEnd
#ce


			GUICtrlSetData($progress, 0)
			GUICtrlSetData($status_input, "")
			GUICtrlSetState($epic_key_input, $GUI_ENABLE)
			GUICtrlSetState($start_button, $GUI_ENABLE)
			GUICtrlSetState($display_report_button, $GUI_ENABLE)
			GUICtrlSetState($display_data_button, $GUI_ENABLE)
			GUISetCursor(2, 0, $main_gui)


		Case $display_report_button

			GUICtrlSetData($status_input, "Creating HTML report ...")
			Create_HTML_Report(True)
;			Create_HTML_Report(False)

			ShellExecute(@ScriptDir & "\html_report.html")

			GUICtrlSetData($status_input, "")

		case $display_data_button

			ShellExecute("notepad", "data.txt", @ScriptDir)

		case $update_confluence_button

			GUICtrlSetData($status_input, "Creating HTML report ...")
			Create_HTML_Report(True)

			GUICtrlSetData($status_input, "Reading HTML report ...")
			Local $html = FileRead(@ScriptDir & "\html_report.html")

			GUICtrlSetData($status_input, "Posting to Confluence ...")
			_ConfluenceSetup()
			_ConfluenceDomainSet("https://janisoncls.atlassian.net")
			_ConfluenceLogin(GUICtrlRead($jira_username_input), GUICtrlRead($jira_password_input))
			_ConfluenceUpdatePage("JAST", "386105345", "386105349", "RMS Traceability Report", $html)
			_ConfluenceShutdown()

			GUICtrlSetData($status_input, "")
	EndSwitch

WEnd

GUIDelete($main_gui)

; Shutdown Jira

GUICtrlSetData($status_input, "Closing Jira ... ")
_JiraShutdown()


Func WM_COMMAND($hWnd, $iMsg, $wParam, $lParam)
    #forceref $hWnd, $iMsg
    Local $hWndFrom, $iIDFrom, $iCode
    $hWndFrom = $lParam
    $iIDFrom = BitAND($wParam, 0xFFFF) ; Low Word
    $iCode = BitShift($wParam, 16) ; Hi Word

	Switch $hWndFrom

        Case GUICtrlGetHandle($testrail_project_combo)

			Switch $iCode

                Case $CBN_SELCHANGE ; Sent when the user changes the current selection in the list box of a combo box

					query_testrail_plans()
            EndSwitch

        Case GUICtrlGetHandle($testrail_plan_combo)

			Switch $iCode

                Case $CBN_SELCHANGE ; Sent when the user changes the current selection in the list box of a combo box

					query_testrail_runs()

            EndSwitch
    EndSwitch
    Return $GUI_RUNDEFMSG
EndFunc   ;==>WM_COMMAND

Func SQLite_to_HTML_table_old($query, $td_alignments, $empty_message)

	Local $td_alignment = StringSplit($td_alignments, ",", 2)

	Local $aResult, $iRows, $iColumns, $iRval

	$iRval = _SQLite_GetTable2d(-1, $query, $aResult, $iRows, $iColumns)
;	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $query = ' & $query & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
;	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $iRval = ' & $iRval & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

	If $iRval = $SQLITE_OK Then

;		_SQLite_Display2DResult($aResult)

		Local $num_rows = UBound($aResult, 1)
		ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $num_rows = ' & $num_rows & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
		Local $num_cols = UBound($aResult, 2)
		ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $num_cols = ' & $num_cols & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

		if $num_rows < 2 Then

			$html = $html &	"<p>" & $empty_message & "</p>" & @CRLF
		Else

			$html = $html &	"<table>" & @CRLF
			$html = $html & "<tr>"

			for $i = 0 to ($num_cols - 1)

				$html = $html & "<th>" & $aResult[0][$i] & "</th>" & @CRLF
			Next

			$html = $html & "</tr>" & @CRLF

			for $i = 1 to ($num_rows - 1)

				$html = $html & "<tr>"

				for $j = 0 to ($num_cols - 1)

					$html = $html & "<td align=""" & $td_alignment[$j] & """>" & $aResult[$i][$j] & "</td>" & @CRLF
				Next

				$html = $html & "</tr>" & @CRLF
			Next

			$html = $html &	"</table>" & @CRLF
		EndIf
	Else
		MsgBox($MB_SYSTEMMODAL, "SQLite Error: " & $iRval, _SQLite_ErrMsg())
	EndIf
EndFunc



Func SQLite_to_HTML_table($query, $th_classes, $td_classes, $empty_message, $run_id, $confluence_html = False)

	Local $double_quotes = """"

	if $confluence_html = true Then

		$double_quotes = "\"""
	EndIf

	Local $th_class = StringSplit($th_classes, ",", 2)
	Local $td_class = StringSplit($td_classes, ",", 2)

	Local $aResult, $iRows, $iColumns, $iRval, $run_name = ""

;	$xx = "SELECT RunName AS ""Run Name"" FROM report WHERE RunID = '" & $run_id & "';"
;	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $xx = ' & $xx & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

	if StringLen($run_id) > 0 Then

		$iRval = _SQLite_GetTable2d(-1, "SELECT RunName AS ""Run Name"" FROM report WHERE RunID = '" & $run_id & "';", $aResult, $iRows, $iColumns)

		If $iRval = $SQLITE_OK Then

;			_SQLite_Display2DResult($aResult)

			$run_name = $aResult[1][0]
		EndIf

		$html = $html &	"<h3>Test Run " & $run_id & " - " & $run_name & "</h3>" & @CRLF
	EndIf

;	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $query = ' & $query & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
	$iRval = _SQLite_GetTable2d(-1, $query, $aResult, $iRows, $iColumns)

	If $iRval = $SQLITE_OK Then

;		_SQLite_Display2DResult($aResult)

		Local $num_rows = UBound($aResult, 1)
		Local $num_cols = UBound($aResult, 2)

		if $num_rows < 2 Then

			$html = $html &	"<p>" & $empty_message & "</p>" & @CRLF
		Else

			if $confluence_html = true Then

				$html = $html &	"<font size=\""1\""><table class=\""wrapped fixed-table\"">" & @CRLF
			Else

				$html = $html &	"<table style=" & $double_quotes & "table-layout:fixed" & $double_quotes & ">" & @CRLF
			EndIf

			$html = $html & "<tr>"

			for $i = 0 to ($num_cols - 1)

				if $confluence_html = true Then

					$html = $html & "<th width=" & $double_quotes & $th_class[$i] & $double_quotes & ">" & $aResult[0][$i] & "</th>" & @CRLF
				Else

					$html = $html & "<th class=" & $double_quotes & $th_class[$i] & $double_quotes & ">" & $aResult[0][$i] & "</th>" & @CRLF
				EndIf
			Next

			$html = $html & "</tr>" & @CRLF

			for $i = 1 to ($num_rows - 1)

				$html = $html & "<tr>"

				for $j = 0 to ($num_cols - 1)

					if $j = 2 Then

	;					ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $aResult[$i][$j] = ' & $aResult[$i][$j] & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

						Switch $aResult[$i][$j]

							case "Passed"

								$td_class[$j] = "trp"

							case "Failed"

								$td_class[$j] = "trf"

							case "Untested"

								$td_class[$j] = "tru"

							case "Blocked"

								$td_class[$j] = "trb"
						EndSwitch
					EndIf


					if $confluence_html = true Then

;						$aResult[$i][$j] = StringReplace($aResult[$i][$j], " \</td>", " \\</td>")
						$aResult[$i][$j] = StringRegExpReplace($aResult[$i][$j], "([^\\])\\$", "$1\\\\")
;						$a = StringRegExpReplace($a, "([^\\])\\$", "$1\\\\")
						$aResult[$i][$j] = StringReplace($aResult[$i][$j], "<br>", "<br/>")
						$aResult[$i][$j] = StringReplace($aResult[$i][$j], "&", "&amp;")
						$aResult[$i][$j] = StringReplace($aResult[$i][$j], """", "\""")
						$aResult[$i][$j] = StringReplace($aResult[$i][$j], "\\""", "\""")
						$html = $html & "<td>" & $aResult[$i][$j] & "</td>" & @CRLF
					Else

						$html = $html & "<th class=" & $double_quotes & $th_class[$i] & $double_quotes & ">" & $aResult[0][$i] & "</th>" & @CRLF
					EndIf
				Next

				$html = $html & "</tr>" & @CRLF
			Next

			$html = $html &	"</table></font>" & @CRLF
		EndIf
	Else
		MsgBox($MB_SYSTEMMODAL, "SQLite Error: " & $iRval, _SQLite_ErrMsg())
	EndIf
EndFunc


Func query_testrail_plans()

	Local $project_id_name = GUICtrlRead($testrail_project_combo)

	IniWrite($ini_filename, "main", "testrailproject", $project_id_name)

	Local $project_part = StringSplit($project_id_name, " - ", 3)

	GUICtrlSetData($status_input, "Querying TestRail Plans ... ")

	Local $plan_id_name = _TestRailGetPlansIDAndNameArray($project_part[0])
	Local $plan_id_str = ""

	for $i = 0 to (UBound($plan_id_name) - 1)

		if StringLen($plan_id_str) > 0 Then

			$plan_id_str = $plan_id_str & "|"
		EndIf

		$plan_id_str = $plan_id_str & $plan_id_name[$i][0] & " - " & $plan_id_name[$i][1]
	Next

	GUICtrlSetData($testrail_plan_combo, $plan_id_str)
	GUICtrlSetData($status_input, "")
	GUICtrlSetState($testrail_plan_combo, $GUI_ENABLE)

EndFunc

Func query_testrail_runs()

	Local $plan_id_name = GUICtrlRead($testrail_plan_combo)
	Local $plan_part = StringSplit($plan_id_name, " - ", 3)

	GUICtrlSetData($status_input, "Querying TestRail Runs ... ")
	Local $run_id = _TestRailGetPlanRunsID($plan_part[0])
	$run_ids = _ArrayToString($run_id)
	GUICtrlSetData($status_input, "")
	GUICtrlSetState($start_button, $GUI_ENABLE)

EndFunc

Func Create_HTML_Report($confluence_html = False)

	_SQLite_Open(@ScriptDir & "\" & $app_name & ".sqlite")

	$html = 				""

	if $confluence_html = False Then

		$html = $html &		"<!DOCTYPE html>" & @CRLF & _
							"<html>" & @CRLF & _
							"<head>" & @CRLF & _
							"<style>" & @CRLF & _
							"table, th, td {" & @CRLF & _
							"    border: 1px solid black;" & @CRLF & _
							"    border-collapse: collapse;" & @CRLF & _
							"    font-size: 12px;" & @CRLF & _
							"    font-family: Arial;" & @CRLF & _
							"}" & @CRLF & _
							".ds {min-width: 400px; text-align: left;}" & @CRLF & _
							".tes {min-width: 800px; text-align: left;}" & @CRLF & _
							".mti {min-width: 110px; text-align: center;}" & @CRLF & _
							".tt {min-width: 500px; text-align: left;}" & @CRLF & _
							".ati {min-width: 150px; text-align: center;}" & @CRLF & _
							".sd {min-width: 1000px; text-align: left;}" & @CRLF & _
							".tc {min-width: 300px; text-align: left;}" & @CRLF & _
							".tr {min-width: 110px; text-align: center;}" & @CRLF & _
							".ts {min-width: 90px; text-align: center;}" & @CRLF & _
							".trp {min-width: 110px; text-align: center; background-color: yellowgreen;}" & @CRLF & _
							".trf {min-width: 110px; text-align: center; background-color: lightcoral; color:white;}" & @CRLF & _
							".tru {min-width: 110px; text-align: center; background-color: lightgray;}" & @CRLF & _
							".trb {min-width: 110px; text-align: center; background-color: darkred; color: white;}" & @CRLF & _
							".pass {background-color: yellowgreen;}" & @CRLF & _
							".fail {background-color: lightcoral; color:white;}" & @CRLF & _
							".untested {background-color: lightgray;}" & @CRLF & _
							".mp {background-color: yellow;}" & @CRLF & _
							".rh {background-color: seagreen; color: white;}" & @CRLF & _
							".rhr {background-color: seagreen; color: white; text-align:center; white-space:nowrap; transform-origin:50% 50%; transform: rotate(-90deg);}" & @CRLF & _
							".rhr:before {background-color: seagreen; color: white; content:''; padding-top:100%; display:inline-block; vertical-align:middle;}" & @CRLF & _
							".i {background-color: deepskyblue;}" & @CRLF & _
							"</style>" & @CRLF & _
							"</head>" & @CRLF & _
							"<body>" & @CRLF
	EndIf

	$html = $html &			"<h2>" & GUICtrlRead($epic_key_input) & " Requirement Traceability</h2>" & @CRLF

	if $confluence_html = True Then

;		SQLite_to_HTML_table("select '<a href=""https://janisoncls.atlassian.net/browse/' || Epic.Key || '"" target=""_blank"">' || Epic.Key || '</a>' AS ""Epic Key"", Epic.Summary AS ""Epic Summary"", Story.ReqID AS ""Req ID"", '<a href=""https://janisoncls.atlassian.net/browse/' || Story.Key || '"" target=""_blank"">' || Story.Key || '</a>' AS ""Story Key"", Story.Summary AS ""Story Summary"", Story.Status AS ""Story Status"", Story.FixVersion AS ""Story Fix Version"", Story.TestNotes AS ""Story Test Notes"", '<a href=""https://janisoncls.atlassian.net/browse/' || Task.Key || '"" target=""_blank"">' || Task.Key || '</a>' AS ""Task Key"", Task.Summary AS ""Task Summary"", TestCase.Id AS ""Case Id linked to Story"", TestCase.Title AS ""Case Title linked to Story"", Test.RunId AS ""Run Id linked to Case"", Run.Name AS ""Run Name linked to Case"", Test.Id AS ""Test Id linked to Case+Run"", Test.Title AS ""Test Title linked to Case+Run"", max(Result.CreatedOn) AS ""Result Date linked to Test"", Result.Id AS ""Result Id linked to Test"", TestStatus.Label AS ""Result Status linked to Test"", Result.Defects AS ""Defects linked to Test"" from Epic left join Story on Epic.Key = Story.EpicKey left join Task on Task.StoryKey like '%,' || Story.Key || ',%' left join TestCase on TestCase.Reference like '%, ' || Story.Key || ',%' left join Test on TestCase.Id = Test.TestCaseId left join Run on Test.RunId = Run.Id left join Result on Test.Id = Result.TestId left join TestStatus on Result.StatusId = TestStatus.Id group by Epic.Key, Story.Key, Task.Key, TestCase.Id, TestCase.Title, Test.RunId, Run.Name, Test.Id, Test.Title order by cast(replace(Epic.Key,'RMS-','') as int), cast(replace(Story.Key,'RMS-','') as int), cast(replace(Task.Key,'RMS-','') as int), TestCase.Id;", "90,300,90,90,500,90,90,300,90,500,90,500,90,170,90,500,90,90,90,200", "ts,tc,ts,ts,tt,ts,ts,tc,ts,tt,ts,tt,ts,tc,ts,tc,ts,ts,ts,tc", "", "", $confluence_html)
		SQLite_to_HTML_table("select '<a href=""https://janisoncls.atlassian.net/browse/' || Epic.Key || '"" target=""_blank"">' || Epic.Key || '</a>' AS ""Epic Key"", Epic.Summary AS ""Epic Summary"", Story.ReqID AS ""Req ID"", '<a href=""https://janisoncls.atlassian.net/browse/' || Story.Key || '"" target=""_blank"">' || Story.Key || '</a>' AS ""Story Key"", Story.Summary AS ""Story Summary"", Story.Status AS ""Story Status"", Story.FixVersion AS ""Story Fix Version"", Story.TestNotes AS ""Story Test Notes"", '<a href=""https://janisoncls.atlassian.net/browse/' || Task.Key || '"" target=""_blank"">' || Task.Key || '</a>' AS ""Task Key"", Task.Summary AS ""Task Summary"", TestCase.Id AS ""Case Id linked to Story"", TestCase.Title AS ""Case Title linked to Story"", Test.RunId AS ""Run Id linked to Case"", Run.Name AS ""Run Name linked to Case"", Test.Id AS ""Test Id linked to Case+Run"", Test.Title AS ""Test Title linked to Case+Run"", max(Result.CreatedOn) AS ""Result Date linked to Test"", Result.Id AS ""Result Id linked to Test"", TestStatus.Label AS ""Result Status linked to Test"", Result.Defects AS ""Defects linked to Test"" from Epic left join Story on Epic.Key = Story.EpicKey left join Task on Task.StoryKey like '%,' || Story.Key || ',%' left join TestCase on TestCase.Reference like '%, ' || Story.Key || ',%' left join Test on TestCase.Id = Test.TestCaseId left join Run on Test.RunId = Run.Id left join Result on Test.Id = Result.TestId left join TestStatus on Result.StatusId = TestStatus.Id group by Epic.Key, Story.Key, Task.Key, TestCase.Id, TestCase.Title, Test.RunId, Run.Name, Test.Id, Test.Title order by cast(replace(Epic.Key,'RMS-','') as int), cast(replace(Story.Key,'RMS-','') as int), cast(replace(Task.Key,'RMS-','') as int), TestCase.Id;", "60,200,60,60,330,60,60,200,60,330,60,330,60,110,60,330,60,60,60,130", "ts,tc,ts,ts,tt,ts,ts,tc,ts,tt,ts,tt,ts,tc,ts,tc,ts,ts,ts,tc", "", "", $confluence_html)
	Else

		SQLite_to_HTML_table("select '<a href=""https://janisoncls.atlassian.net/browse/' || Epic.Key || '"" target=""_blank"">' || Epic.Key || '</a>' AS ""Epic Key"", Epic.Summary AS ""Epic Summary"", Story.ReqID AS ""Req ID"", '<a href=""https://janisoncls.atlassian.net/browse/' || Story.Key || '"" target=""_blank"">' || Story.Key || '</a>' AS ""Story Key"", Story.Summary AS ""Story Summary"", Story.Status AS ""Story Status"", Story.FixVersion AS ""Story Fix Version"", Story.TestNotes AS ""Story Test Notes"", '<a href=""https://janisoncls.atlassian.net/browse/' || Task.Key || '"" target=""_blank"">' || Task.Key || '</a>' AS ""Task Key"", Task.Summary AS ""Task Summary"", TestCase.Id AS ""Case Id linked to Story"", TestCase.Title AS ""Case Title linked to Story"", Test.RunId AS ""Run Id linked to Case"", Run.Name AS ""Run Name linked to Case"", Test.Id AS ""Test Id linked to Case+Run"", Test.Title AS ""Test Title linked to Case+Run"", max(Result.CreatedOn) AS ""Result Date linked to Test"", Result.Id AS ""Result Id linked to Test"", TestStatus.Label AS ""Result Status linked to Test"", Result.Defects AS ""Defects linked to Test"" from Epic left join Story on Epic.Key = Story.EpicKey left join Task on Task.StoryKey like '%,' || Story.Key || ',%' left join TestCase on TestCase.Reference like '%, ' || Story.Key || ',%' left join Test on TestCase.Id = Test.TestCaseId left join Run on Test.RunId = Run.Id left join Result on Test.Id = Result.TestId left join TestStatus on Result.StatusId = TestStatus.Id group by Epic.Key, Story.Key, Task.Key, TestCase.Id, TestCase.Title, Test.RunId, Run.Name, Test.Id, Test.Title order by cast(replace(Epic.Key,'RMS-','') as int), cast(replace(Story.Key,'RMS-','') as int), cast(replace(Task.Key,'RMS-','') as int), TestCase.Id;", "rh,rh,rh,rh,rh,rh,rh,rh,rh,rh,rh,rh,rh,rh,rh,rh,rh,rh,rh,rh", "ts,tc,ts,ts,tt,ts,ts,tc,ts,tt,ts,tt,ts,tc,ts,tc,ts,ts,ts,tc", "", "", $confluence_html)
	EndIf

	if $confluence_html = False Then

		$html = $html &		"</body>" & @CRLF & _
							"</html>" & @CRLF
	EndIf

	FileDelete(@ScriptDir & "\html_report.html")
	FileWrite(@ScriptDir & "\html_report.html", $html)
EndFunc


