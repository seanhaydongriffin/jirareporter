#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_UseUpx=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#RequireAdmin
;#AutoIt3Wrapper_usex64=n
#include <Date.au3>
#include <File.au3>
#include <Array.au3>
#include "Jira.au3"
#include "TestRail.au3"
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <GuiListView.au3>
#include <ButtonConstants.au3>
#include <ComboConstants.au3>
#include <SQLite.au3>
#include <SQLite.dll.au3>

Global $app_name = "Data Extractor"
Global $log_filepath = @ScriptDir & "\" & $app_name & ".log"

;#cs
Global $testrail_username = $CmdLine[1]
Global $testrail_password = $CmdLine[2]
Global $testrail_run_ids = $CmdLine[3]
Global $jira_username = $CmdLine[4]
Global $jira_password = $CmdLine[5]
Global $jira_project = $CmdLine[6]
;#ce

#cs
Global $testrail_username = $CmdLine[1]
Global $testrail_password = $CmdLine[2]
Global $testrail_plan = $CmdLine[4]
Global $jira_username = $CmdLine[5]
Global $jira_password = $CmdLine[6]
Global $jira_project = $CmdLine[7]
#ce

Local $iRows, $iColumns, $aMyDate, $aMyTime
Local $num_opened_bugs, $num_resolved_bugs
Local $num_opened_blocker_bugs, $num_resolved_blocker_bugs
Local $num_opened_critical_bugs, $num_resolved_critical_bugs
Local $num_opened_major_bugs, $num_resolved_major_bugs
Local $num_opened_minor_bugs, $num_resolved_minor_bugs
Local $num_opened_trivial_bugs, $num_resolved_trivial_bugs


; Startup SQLite

_SQLite_Startup()
ConsoleWrite("_SQLite_LibVersion=" & _SQLite_LibVersion() & @CRLF)
_SQLite_Open(@ScriptDir & "\Jira Reporter.sqlite")
_SQLite_Exec(-1, "PRAGMA synchronous = OFF;")		; this should speed up DB transactions

; Startup Jira & TestRail

;GUICtrlSetData($status_input, "Starting the Jira connection ... ")
_JiraSetup()
_JiraDomainSet("https://janisoncls.atlassian.net")
_JiraLogin($jira_username, $jira_password)


;GUICtrlSetData($status_input, "Starting the TestRail connection ... ")
_TestRailDomainSet("https://janison.testrail.com")
_TestRailLogin($testrail_username, $testrail_password)


FileDelete($log_filepath)
SplashTextOn($app_name, "", 1200, 400, Default, Default, 16)


; get all versions (drops, releases, etc) for the project

$version = _JiraGetVersionNames($jira_project)
_ArrayInsert($version, 0, "all")


#cs

; get all epics for the project

ControlSetText($app_name, "", "Static1", "get all epics for the project")
;GUICtrlSetData($status_input, "Querying Project RMS ... ")
$issue = _JiraGetSearchResultKeysSummariesAndIssueTypeNames("summary,issuetype", "project = " & $jira_project & " AND issuetype = Epic")

for $i = 0 to (UBound($issue) - 1) Step 3


	$query = "INSERT INTO Epic (Key,Summary) VALUES ('" & $issue[$i] & "','" & $issue[$i + 1] & "');"
	_FileWriteLog($log_filepath, "Epic " & ($i + 1) & " of " & UBound($issue) & " = " & $query)
	_SQLite_Exec(-1, $query) ; INSERT Data

Next

; get all stories for the project

ControlSetText($app_name, "", "Static1", "get all stories for the project")
$issue = _JiraGetSearchResultKeysSummariesIssueTypeNameEpicKeyRequirements("summary,issuetype,customfield_10008,labels,fixVersions,status,customfield_15255", "project = " & $jira_project & " AND issuetype in (Improvement, Story)")

for $i = 0 to (UBound($issue) - 1) Step 8

	$query = "INSERT INTO Story (Key,Summary,EpicKey,ReqID,FixVersion,Status,TestNotes) VALUES ('" & $issue[$i] & "','" & $issue[$i + 1] & "','" & $issue[$i + 3] & "','" & $issue[$i + 4] & "','" & $issue[$i + 5] & "','" & $issue[$i + 6] & "','" & $issue[$i + 7] & "');"
	_FileWriteLog($log_filepath, "Story " & ($i + 1) & " of " & UBound($issue) & " = " & $query)
	_SQLite_Exec(-1, $query) ; INSERT Data
Next

; get all tasks for the project

ControlSetText($app_name, "", "Static1", "get all tasks for the project")
$issue = _JiraGetSearchResultKeysSummariesIssueTypeNamesIssueLinks("summary,issuetype,issuelinks", "project = " & $jira_project & " AND issuetype = Task")

for $i = 0 to (UBound($issue) - 1) Step 4

	$query = "INSERT INTO Task (Key,Summary,StoryKey) VALUES ('" & $issue[$i] & "','" & $issue[$i + 1] & "','" & $issue[$i + 3] & "');"
	_FileWriteLog($log_filepath, "Task " & ($i + 1) & " of " & UBound($issue) & " = " & $query)
	_SQLite_Exec(-1, $query) ; INSERT Data
Next



; get all bugs for the project

ControlSetText($app_name, "", "Static1", "get all bugs for the project")
;$issue = _JiraGetSearchResultBugs("summary,reporter,fixVersions,resolution,customfield_12000,priority,labels,versions,status,customfield_10007,environment,customfield_11602,assignee", "project = " & $jira_project & " AND issuetype = Bug")
$issue = _JiraGetSearchResultBugs("summary,reporter,fixVersions,resolution,customfield_12000,priority,labels,versions,status,customfield_10007,environment,assignee", "project = " & $jira_project & " AND issuetype = Bug")

for $i = 0 to (UBound($issue) - 1) Step 4

	$query = "INSERT INTO Bug (Key,Summary,Reporter,Assignee,Status,Priority,AffectsVersions,FixVersions,Resolution,Labels,Environment,ScrumTeam,Sprint) VALUES ('" & $issue[$i] & "','" & $issue[$i + 1] & "','" & $issue[$i + 3] & "');"
	_FileWriteLog($log_filepath, "Task " & ($i + 1) & " of " & UBound($issue) & " = " & $query)
	_SQLite_Exec(-1, $query) ; INSERT Data
Next


#ce


; get all bug change logs for the project

for $version_num = 0 to (UBound($version) - 1)

	ControlSetText($app_name, "", "Static1", "get all bug change logs from the jira project" & @CRLF & "for " & $version[$version_num] & " version")
	;$issue = _JiraGetSearchResultBugs("summary,reporter,fixVersions,resolution,customfield_12000,priority,labels,versions,status,customfield_10007,environment,customfield_11602,assignee", "project = " & $jira_project & " AND issuetype = Bug")

	Local $change

	if StringCompare($version[$version_num], "all") = 0 Then

		$change = _JiraGetSearchResultBugStatusHistory("status,versions,priority,created,changelog", "project = " & $jira_project & " AND issuetype = Bug")
	Else

		$change = _JiraGetSearchResultBugStatusHistory("status,versions,priority,created,changelog", "project = " & $jira_project & " AND issuetype = Bug AND affectedVersion = """ & $version[$version_num] & """")
	EndIf

	ControlSetText($app_name, "", "Static1", "update local db with all bug change logs")

	_SQLite_Exec(-1, "BEGIN TRANSACTION")

	for $i = 0 to (UBound($change) - 1) Step 4

		Local $open_date = ""
		Local $resolved_date = ""
		Local $resolved = False
		Local $previous_created = ""

		if StringCompare($version[$version_num], "all") = 0 Then

			$change[$i + 1] = "all"
		EndIf

		Local $change_item = StringSplit($change[$i + 3], "|", 3)

		for $j = 0 to (UBound($change_item) - 1)

			Local $change_item_part = StringSplit($change_item[$j], ",", 3)

			$query = "INSERT INTO BugChangeLog (Key,AffectsVersion,Priority,Created,Field,Old,New) VALUES ('" & $change[$i] & "','" & $change[$i + 1] & "','" & $change[$i + 2] & "','" & $change_item_part[0] & "','status','" & $change_item_part[1] & "','" & $change_item_part[2] & "');"
			_FileWriteLog($log_filepath, "BugChangeLog " & ($i + 1) & " of " & UBound($change) & " = " & $query)
			_SQLite_Exec(-1, $query) ; INSERT Data

			if StringInStr($change_item_part[2], "Closed") > 0 Or StringInStr($change_item_part[2], "Resolved") > 0 Or StringInStr($change_item_part[2], "Client Testing") > 0 Or StringInStr($change_item_part[2], "Waiting for UAT Deploy") > 0 Then

				$resolved = True
			EndIf

			if StringLen($resolved_date) = 0 Then

				if $resolved = True And (StringInStr($change_item_part[2], "Open") > 0 Or StringInStr($change_item_part[2], "Awaiting Clarification") > 0 Or StringInStr($change_item_part[2], "Ready for Dev") > 0 Or StringInStr($change_item_part[2], "Dev In Progress") > 0 Or StringInStr($change_item_part[2], "Code Review") > 0) Then

					$resolved_date = $previous_created
				EndIf
			EndIf

			$previous_created = $change_item_part[0]
		Next

		$open_date = $previous_created

		Local $tmp_resolved_date = $resolved_date, $unresolved_age = "", $BlockerUnresolvedAge = "", $CriticalUnresolvedAge = "", $MajorUnresolvedAge = "", $MinorUnresolvedAge = "", $TrivialUnresolvedAge = ""

		if StringLen($tmp_resolved_date) = 0 Then

			$tmp_resolved_date = _NowCalc()
		EndIf

		$unresolved_age = _DateDiff("D", $open_date, $tmp_resolved_date)

		Switch $change[$i + 2]

			Case "Blocker"

				$BlockerUnresolvedAge = $unresolved_age

			Case "Critical"

				$CriticalUnresolvedAge = $unresolved_age

			Case "Major"

				$MajorUnresolvedAge = $unresolved_age

			Case "Minor"

				$MinorUnresolvedAge = $unresolved_age

			Case "Trivial"

				$TrivialUnresolvedAge = $unresolved_age

		EndSwitch

		$query = "INSERT INTO BugStateDate (Key,AffectsVersion,Priority,OpenDate,ResolvedDate,BlockerUnresolvedAge,CriticalUnresolvedAge,MajorUnresolvedAge,MinorUnresolvedAge,TrivialUnresolvedAge) VALUES ('" & $change[$i] & "','" & $change[$i + 1] & "','" & $change[$i + 2] & "','" & $open_date & "','" & $resolved_date & "','" & $BlockerUnresolvedAge & "','" & $CriticalUnresolvedAge & "','" & $MajorUnresolvedAge & "','" & $MinorUnresolvedAge & "','" & $TrivialUnresolvedAge & "');"
		_FileWriteLog($log_filepath, "BugStateDate " & ($i + 1) & " of " & UBound($change) & " = " & $query)
		_SQLite_Exec(-1, $query) ; INSERT Data

	Next

Next

_SQLite_Exec(-1, "COMMIT TRANSACTION")






; get the total number of bugs resolved vs total number of bugs (defect removal efficiency)

ControlSetText($app_name, "", "Static1", "update local db with the total number of bugs resolved vs total number of bugs")

_SQLite_Exec(-1, "BEGIN TRANSACTION")

for $version_num = 0 to (UBound($version) - 1)

	ControlSetText($app_name, "", "Static1", "update local db with the total number of bugs resolved vs total number of bugs" & @CRLF & "Version #" & $version_num & " " & $version[$version_num])

	Local $BugTotalResolvedPerDate_first_rowid = 0
	Local $BugTotalResolvedPerDate_last_rowid = 0
	Local $BugTotalActivePerDate_first_rowid = 0
	Local $BugTotalActivePerDate_last_rowid = 0
	Local $BugTotalOpenedPerDate_first_rowid = 0
	Local $BugTotalOpenedPerDate_last_rowid = 0
	Local $now_date = _NowCalcDate() ; "2018/07/30"
;	Local $date = _DateAdd("M", -6, $now_date)
;	Local $date = _DateAdd("Y", -1, $now_date)
	Local $date = _DateAdd("Y", -2, $now_date)

	While True

		$date = _DateAdd("D", 1, $date)

		if _DateDiff("D", $now_date, $date) >= 0 Then

			ExitLoop
		EndIf

		$sqlite_date = StringReplace($date, "/", "-") & "T23:59:59"
		ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $sqlite_date = ' & $sqlite_date & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

		_SQLite_GetTable2d(-1, "select count(*) from BugStateDate where AffectsVersion = '" & $version[$version_num] & "' and OpenDate <= '" & $sqlite_date & "'", $num_opened_bugs, $iRows, $iColumns)
		_SQLite_GetTable2d(-1, "select count(*) from BugStateDate where AffectsVersion = '" & $version[$version_num] & "' and OpenDate <= '" & $sqlite_date & "' and Priority = 'Blocker'", $num_opened_blocker_bugs, $iRows, $iColumns)
		_SQLite_GetTable2d(-1, "select count(*) from BugStateDate where AffectsVersion = '" & $version[$version_num] & "' and OpenDate <= '" & $sqlite_date & "' and Priority = 'Critical'", $num_opened_critical_bugs, $iRows, $iColumns)
		_SQLite_GetTable2d(-1, "select count(*) from BugStateDate where AffectsVersion = '" & $version[$version_num] & "' and OpenDate <= '" & $sqlite_date & "' and Priority = 'Major'", $num_opened_major_bugs, $iRows, $iColumns)
		_SQLite_GetTable2d(-1, "select count(*) from BugStateDate where AffectsVersion = '" & $version[$version_num] & "' and OpenDate <= '" & $sqlite_date & "' and Priority = 'Minor'", $num_opened_minor_bugs, $iRows, $iColumns)
		_SQLite_GetTable2d(-1, "select count(*) from BugStateDate where AffectsVersion = '" & $version[$version_num] & "' and OpenDate <= '" & $sqlite_date & "' and Priority = 'Trivial'", $num_opened_trivial_bugs, $iRows, $iColumns)

		_SQLite_GetTable2d(-1, "select count(*) from BugStateDate where AffectsVersion = '" & $version[$version_num] & "' and ResolvedDate <> '' and ResolvedDate <= '" & $sqlite_date & "'", $num_resolved_bugs, $iRows, $iColumns)
		_SQLite_GetTable2d(-1, "select count(*) from BugStateDate where AffectsVersion = '" & $version[$version_num] & "' and ResolvedDate <> '' and ResolvedDate <= '" & $sqlite_date & "' and Priority = 'Blocker'", $num_resolved_blocker_bugs, $iRows, $iColumns)
		_SQLite_GetTable2d(-1, "select count(*) from BugStateDate where AffectsVersion = '" & $version[$version_num] & "' and ResolvedDate <> '' and ResolvedDate <= '" & $sqlite_date & "' and Priority = 'Critical'", $num_resolved_critical_bugs, $iRows, $iColumns)
		_SQLite_GetTable2d(-1, "select count(*) from BugStateDate where AffectsVersion = '" & $version[$version_num] & "' and ResolvedDate <> '' and ResolvedDate <= '" & $sqlite_date & "' and Priority = 'Major'", $num_resolved_major_bugs, $iRows, $iColumns)
		_SQLite_GetTable2d(-1, "select count(*) from BugStateDate where AffectsVersion = '" & $version[$version_num] & "' and ResolvedDate <> '' and ResolvedDate <= '" & $sqlite_date & "' and Priority = 'Minor'", $num_resolved_minor_bugs, $iRows, $iColumns)
		_SQLite_GetTable2d(-1, "select count(*) from BugStateDate where AffectsVersion = '" & $version[$version_num] & "' and ResolvedDate <> '' and ResolvedDate <= '" & $sqlite_date & "' and Priority = 'Trivial'", $num_resolved_trivial_bugs, $iRows, $iColumns)

		_DateTimeSplit($sqlite_date, $aMyDate, $aMyTime)
		Local $confluence_date = StringFormat("%02i/%02i/%04i", $aMyDate[3], $aMyDate[2], $aMyDate[1])

		$query = "INSERT INTO BugTotalResolvedPerDate (AffectsVersion,Date,TotalBugs,TotalBugsResolved,TotalBlockersResolved,TotalCriticalResolved,TotalMajorResolved,TotalMinorResolved,TotalTrivialResolved,TotalUnresolved) VALUES ('" & $version[$version_num] & "','" & $confluence_date & "','" & $num_opened_bugs[1][0] & "','" & $num_resolved_bugs[1][0] & "','" & $num_resolved_blocker_bugs[1][0] & "','" & $num_resolved_critical_bugs[1][0] & "','" & $num_resolved_major_bugs[1][0] & "','" & $num_resolved_minor_bugs[1][0] & "','" & $num_resolved_trivial_bugs[1][0] & "','" & ($num_opened_bugs[1][0] - $num_resolved_bugs[1][0]) & "');"
		_FileWriteLog($log_filepath, "rowid " & _SQLite_LastInsertRowID() & " = " & $query)
		_SQLite_Exec(-1, $query) ; INSERT Data

		if $BugTotalResolvedPerDate_first_rowid = 0 Then

			$BugTotalResolvedPerDate_first_rowid = _SQLite_LastInsertRowID()
		Else

			$BugTotalResolvedPerDate_last_rowid = _SQLite_LastInsertRowID()
		EndIf

		$query = "INSERT INTO BugTotalActivePerDate (AffectsVersion,Date,TotalActive,BlockersActive,CriticalActive,MajorActive,MinorActive,TrivialActive) VALUES ('" & $version[$version_num] & "','" & $confluence_date & "','" & ($num_opened_bugs[1][0] - $num_resolved_bugs[1][0]) & "','" & ($num_opened_blocker_bugs[1][0] - $num_resolved_blocker_bugs[1][0]) & "','" & ($num_opened_critical_bugs[1][0] - $num_resolved_critical_bugs[1][0]) & "','" & ($num_opened_major_bugs[1][0] - $num_resolved_major_bugs[1][0]) & "','" & ($num_opened_minor_bugs[1][0] - $num_resolved_minor_bugs[1][0]) & "','" & ($num_opened_trivial_bugs[1][0] - $num_resolved_trivial_bugs[1][0]) & "');"
		_FileWriteLog($log_filepath, "rowid " & _SQLite_LastInsertRowID() & " = " & $query)
		_SQLite_Exec(-1, $query) ; INSERT Data

		$query = "INSERT INTO BugTotalOpenedPerDate (AffectsVersion,Date,TotalOpened,BlockersOpened,CriticalOpened,MajorOpened,MinorOpened,TrivialOpened) VALUES ('" & $version[$version_num] & "','" & $confluence_date & "','" & $num_opened_bugs[1][0] & "','" & $num_opened_blocker_bugs[1][0] & "','" & $num_opened_critical_bugs[1][0] & "','" & $num_opened_major_bugs[1][0] & "','" & $num_opened_minor_bugs[1][0] & "','" & $num_opened_trivial_bugs[1][0] & "');"
		_FileWriteLog($log_filepath, "rowid " & _SQLite_LastInsertRowID() & " = " & $query)
		_SQLite_Exec(-1, $query) ; INSERT Data

	WEnd

	; locate records to delete:
	;	1. leading and trailing records that are the same
	;	2. too many records / detail to include in the chart (40 dates on the x axis maximum if possible)

	Local $tmp_result, $previous_tmp_result
	Local $dates_to_delete[0]

	for $i = ($BugTotalResolvedPerDate_first_rowid + 1) to $BugTotalResolvedPerDate_last_rowid

		_SQLite_GetTable2d(-1, "select TotalBugs, TotalBugsResolved, Date from BugTotalResolvedPerDate where rowid = " & $i, $tmp_result, $iRows, $iColumns)

		if $iRows > 0 Then

			_SQLite_GetTable2d(-1, "select TotalBugs, TotalBugsResolved, Date from BugTotalResolvedPerDate where rowid = " & ($i - 1), $previous_tmp_result, $iRows, $iColumns)

			if $iRows > 0 Then

				if StringCompare($tmp_result[1][0], $previous_tmp_result[1][0]) = 0 And StringCompare($tmp_result[1][1], $previous_tmp_result[1][1]) = 0 Then

					_ArrayAdd($dates_to_delete, $previous_tmp_result[1][2], 0, "|", @CRLF, 1)
;					$query = "delete from BugTotalResolvedPerDate where rowid = " & ($i - 1)
;					_FileWriteLog($log_filepath, $query)
;					_SQLite_Exec(-1, $query)
				Else

					ExitLoop
				EndIf
			EndIf
		EndIf
	Next

	for $i = ($BugTotalResolvedPerDate_last_rowid - 1) to $BugTotalResolvedPerDate_first_rowid Step -1

		_SQLite_GetTable2d(-1, "select TotalBugs, TotalBugsResolved, Date from BugTotalResolvedPerDate where rowid = " & $i, $tmp_result, $iRows, $iColumns)

		if $iRows > 0 Then

			_SQLite_GetTable2d(-1, "select TotalBugs, TotalBugsResolved, Date from BugTotalResolvedPerDate where rowid = " & ($i + 1), $previous_tmp_result, $iRows, $iColumns)

			if $iRows > 0 Then

				if StringCompare($tmp_result[1][0], $previous_tmp_result[1][0]) = 0 And StringCompare($tmp_result[1][1], $previous_tmp_result[1][1]) = 0 Then

					_ArrayAdd($dates_to_delete, $previous_tmp_result[1][2], 0, "|", @CRLF, 1)
				Else

					ExitLoop
				EndIf
			EndIf
		EndIf
	Next

	for $i = 0 to (UBound($dates_to_delete) - 1)

		$query = "delete from BugTotalResolvedPerDate where AffectsVersion = '" & $version[$version_num] & "' and Date = '" & $dates_to_delete[$i] & "'"
		_FileWriteLog($log_filepath, $query)
		_SQLite_Exec(-1, $query)

		$query = "delete from BugTotalActivePerDate where AffectsVersion = '" & $version[$version_num] & "' and Date = '" & $dates_to_delete[$i] & "'"
		_FileWriteLog($log_filepath, $query)
		_SQLite_Exec(-1, $query)

		$query = "delete from BugTotalOpenedPerDate where AffectsVersion = '" & $version[$version_num] & "' and Date = '" & $dates_to_delete[$i] & "'"
		_FileWriteLog($log_filepath, $query)
		_SQLite_Exec(-1, $query)
	Next

	Local $dates_to_delete[0]

	_SQLite_GetTable2d(-1, "select Date from BugTotalResolvedPerDate where AffectsVersion = '" & $version[$version_num] & "'", $tmp_result, $iRows, $iColumns)

	if $iRows > 50 Then

		Local $nth_date_to_keep = Ceiling($iRows / 50)
		Local $count_to_nth_date_to_keep = 0

		for $i = 1 to $iRows

			$count_to_nth_date_to_keep = $count_to_nth_date_to_keep + 1

			if $count_to_nth_date_to_keep = $nth_date_to_keep Then

				$count_to_nth_date_to_keep = 0
			Else

				if $i > 1 and $i < $iRows Then

					_ArrayAdd($dates_to_delete, $tmp_result[$i][0], 0, "|", @CRLF, 1)
				EndIf
			EndIf
		Next

		for $i = 0 to (UBound($dates_to_delete) - 1)

			$query = "delete from BugTotalResolvedPerDate where AffectsVersion = '" & $version[$version_num] & "' and Date = '" & $dates_to_delete[$i] & "'"
			_FileWriteLog($log_filepath, $query)
			_SQLite_Exec(-1, $query)

			$query = "delete from BugTotalActivePerDate where AffectsVersion = '" & $version[$version_num] & "' and Date = '" & $dates_to_delete[$i] & "'"
			_FileWriteLog($log_filepath, $query)
			_SQLite_Exec(-1, $query)

			$query = "delete from BugTotalOpenedPerDate where AffectsVersion = '" & $version[$version_num] & "' and Date = '" & $dates_to_delete[$i] & "'"
			_FileWriteLog($log_filepath, $query)
			_SQLite_Exec(-1, $query)
		Next
	EndIf

Next

_SQLite_Exec(-1, "COMMIT TRANSACTION")


; get the number of bugs opened per week

ControlSetText($app_name, "", "Static1", "update local db with the number of bugs opened per week")

_SQLite_Exec(-1, "BEGIN TRANSACTION")

for $version_num = 0 to (UBound($version) - 1)

	ControlSetText($app_name, "", "Static1", "update local db with the number of bugs opened per week" & @CRLF & "Version #" & $version_num & " " & $version[$version_num])

	; get the date of the last opened bug

	Local $last_opened_bug_date

	_SQLite_GetTable2d(-1, "select max(OpenDate) from BugStateDate where AffectsVersion = '" & $version[$version_num] & "'", $last_opened_bug_date, $iRows, $iColumns)

	if $iRows > 0 And StringLen($last_opened_bug_date[1][0]) > 0 Then

		; get the date 12 months prior to the date of the last opened bug

		Local $date = _DateAdd("Y", -1, $last_opened_bug_date[1][0])

		While True

			; advance days until the next Monday (start of week) is found

			while True

				_DateTimeSplit($date, $aMyDate, $aMyTime)
				Local $weekday_number = _DateToDayOfWeekISO($aMyDate[1], $aMyDate[2], $aMyDate[3])

				if $weekday_number = 1 Then

					ExitLoop
				EndIf

				$date = _DateAdd("D", 1, $date)
			WEnd

			if _DateDiff("D", $last_opened_bug_date[1][0], $date) >= 0 Then

				ExitLoop
			EndIf

			_DateTimeSplit($date, $aMyDate, $aMyTime)
			Local $sqlite_week_start_date = StringFormat("%04i-%02i-%02iT00:00:00", $aMyDate[1], $aMyDate[2], $aMyDate[3])
			$date = _DateAdd("D", 6, $date)
			_DateTimeSplit($date, $aMyDate, $aMyTime)
			Local $sqlite_week_end_date = StringFormat("%04i-%02i-%02iT23:59:59", $aMyDate[1], $aMyDate[2], $aMyDate[3])

			_SQLite_GetTable2d(-1, "select count(*) from BugStateDate where AffectsVersion = '" & $version[$version_num] & "' and OpenDate >= '" & $sqlite_week_start_date & "' and OpenDate <= '" & $sqlite_week_end_date & "'", $num_opened_bugs, $iRows, $iColumns)
			_SQLite_GetTable2d(-1, "select count(*) from BugStateDate where AffectsVersion = '" & $version[$version_num] & "' and OpenDate >= '" & $sqlite_week_start_date & "' and OpenDate <= '" & $sqlite_week_end_date & "' and Priority = 'Blocker'", $num_opened_blocker_bugs, $iRows, $iColumns)
			_SQLite_GetTable2d(-1, "select count(*) from BugStateDate where AffectsVersion = '" & $version[$version_num] & "' and OpenDate >= '" & $sqlite_week_start_date & "' and OpenDate <= '" & $sqlite_week_end_date & "' and Priority = 'Critical'", $num_opened_critical_bugs, $iRows, $iColumns)
			_SQLite_GetTable2d(-1, "select count(*) from BugStateDate where AffectsVersion = '" & $version[$version_num] & "' and OpenDate >= '" & $sqlite_week_start_date & "' and OpenDate <= '" & $sqlite_week_end_date & "' and Priority = 'Major'", $num_opened_major_bugs, $iRows, $iColumns)
			_SQLite_GetTable2d(-1, "select count(*) from BugStateDate where AffectsVersion = '" & $version[$version_num] & "' and OpenDate >= '" & $sqlite_week_start_date & "' and OpenDate <= '" & $sqlite_week_end_date & "' and Priority = 'Minor'", $num_opened_minor_bugs, $iRows, $iColumns)
			_SQLite_GetTable2d(-1, "select count(*) from BugStateDate where AffectsVersion = '" & $version[$version_num] & "' and OpenDate >= '" & $sqlite_week_start_date & "' and OpenDate <= '" & $sqlite_week_end_date & "' and Priority = 'Trivial'", $num_opened_trivial_bugs, $iRows, $iColumns)

			_DateTimeSplit($sqlite_week_start_date, $aMyDate, $aMyTime)
			Local $confluence_week_start_date = StringFormat("%02i/%02i/%04i", $aMyDate[3], $aMyDate[2], $aMyDate[1])

			$query = "INSERT INTO BugOpenedPerWeek (AffectsVersion,WeekStarting,AllOpened,BlockersOpened,CriticalOpened,MajorOpened,MinorOpened,TrivialOpened) VALUES ('" & $version[$version_num] & "','" & $confluence_week_start_date & "','" & $num_opened_bugs[1][0] & "','" & $num_opened_blocker_bugs[1][0] & "','" & $num_opened_critical_bugs[1][0] & "','" & $num_opened_major_bugs[1][0] & "','" & $num_opened_minor_bugs[1][0] & "','" & $num_opened_trivial_bugs[1][0] & "');"
			_FileWriteLog($log_filepath, "rowid " & _SQLite_LastInsertRowID() & " = " & $query)
			_SQLite_Exec(-1, $query) ; INSERT Data
		WEnd
	EndIf
Next

_SQLite_Exec(-1, "COMMIT TRANSACTION")







#cs




; get all test cases for the project

ControlSetText($app_name, "", "Static1", "get all test cases for the project")
local $suite = _TestRailGetSuitesIdName(43)

for $i = 0 to (UBound($suite) - 1) Step 2

;	$suite[$i]
	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $suite[$i] = ' & $suite[$i] & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

	local $case = _TestRailGetCasesIdTitleReferences(43, $suite[$i])
;	_ArrayDisplay($case)

	if IsArray($case) = True Then

		for $j = 0 to (UBound($case) - 1) Step 3

;			$ref = StringSplit($case[$j + 2], ",", 3)

;			for $k = 0 to (UBound($ref) - 1)

;				$ref[$k] = StringStripWS($ref[$k], 3)
;				$query = "INSERT INTO TestCase (Id,Title,Reference) VALUES ('" & $case[$j] & "','" & $case[$j + 1] & "','" & $ref[$k] & "');"
;				_FileWriteLog($log_filepath, "Suite " & ($i + 1) & " of " & UBound($suite) & " Case " & ($j + 1) & " of " & UBound($case) & " = " & $query)
;				_SQLite_Exec(-1, $query) ; INSERT Data
;			Next

			if StringLen($case[$j + 2]) > 0 Then

				$case[$j + 2] = ", " & $case[$j + 2] & ", "
			EndIf

			$query = "INSERT INTO TestCase (Id,Title,Reference) VALUES ('" & $case[$j] & "','" & $case[$j + 1] & "','" & $case[$j + 2] & "');"
			_FileWriteLog($log_filepath, "Suite " & ($i + 1) & " of " & UBound($suite) & " Case " & ($j + 1) & " of " & UBound($case) & " = " & $query)
			_SQLite_Exec(-1, $query) ; INSERT Data

	;		$case[$j]
;			ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $case[$j] = ' & $case[$j] & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
	;		$case[$j + 1]
;			ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $case[$j + 1] = ' & $case[$j + 1] & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
	;		$case[$j + 2]
;			ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $case[$j + 2] = ' & $case[$j + 2] & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
		Next

	EndIf

;	$query = "INSERT INTO Story (Key,Summary,EpicKey,ReqID) VALUES ('" & $issue[$i] & "','" & $issue[$i + 1] & "','" & $issue[$i + 3] & "','" & $issue[$i + 4] & "');"
;	_FileWriteLog($log_filepath, "Story " & ($i + 1) & " of " & UBound($issue) & " = " & $query)
;	_SQLite_Exec(-1, $query) ; INSERT Data
Next


; get all test statuses

ControlSetText($app_name, "", "Static1", "get all test statuses")
local $status = _TestRailGetStatusesIdLabel()

for $i = 0 to (UBound($status) - 1) Step 2

	$query = "INSERT INTO TestStatus (Id,Label) VALUES ('" & $status[$i] & "','" & $status[$i + 1] & "');"
	_FileWriteLog($log_filepath, "Status " & ($i + 1) & " of " & (UBound($status) / 2) & " = " & $query)
	_SQLite_Exec(-1, $query) ; INSERT Data
Next

; get all runs for the project (not in test plans)

ControlSetText($app_name, "", "Static1", "get all runs for the project (not in test plans)")
local $run = _TestRailGetRunsIdName(43)

for $i = 0 to (UBound($run) - 1) Step 2

	$query = "INSERT INTO Run (Id,Name) VALUES ('" & $run[$i] & "','" & $run[$i + 1] & "');"
	_FileWriteLog($log_filepath, "Run " & (($i / 2) + 1) & " of " & (UBound($run) / 2) & " = " & $query)
	_SQLite_Exec(-1, $query) ; INSERT Data

	ControlSetText($app_name, "", "Static1", "get all tests for run id " & $run[$i] & " - " & (($i / 2) + 1) & " of " & (UBound($run) / 2))
	local $test = _TestRailGetTestsIdTitleCaseIdRunId($run[$i])

	for $j = 0 to (UBound($test) - 1) Step 4

		$query = "INSERT INTO Test (Id,Title,TestCaseId,RunId) VALUES ('" & $test[$j] & "','" & $test[$j + 3] & "','" & $test[$j + 1] & "','" & $test[$j + 2] & "');"
		_FileWriteLog($log_filepath, "Test " & (($j / 4) + 1) & " of " & (UBound($test) / 4) & " = " & $query)
		_SQLite_Exec(-1, $query) ; INSERT Data
	Next

	ControlSetText($app_name, "", "Static1", "get all results for run id " & $run[$i] & " - " & (($i / 2) + 1) & " of " & (UBound($run) / 2))
	local $result = _TestRailGetResultsForRunIdStatusIdCreatedOnDefects($run[$i])

	for $j = 0 to (UBound($result) - 1) Step 5

		$result[$j + 4] = StringReplace($result[$j + 4], "null", "")
		$result[$j + 4] = StringReplace($result[$j + 4], """", "")

		$query = "INSERT INTO Result (Id,TestId,StatusId,CreatedOn,Defects) VALUES ('" & $result[$j] & "','" & $result[$j + 1] & "','" & $result[$j + 2] & "','" & $result[$j + 3] & "','" & $result[$j + 4] & "');"
		_FileWriteLog($log_filepath, "Result " & (($j / 5) + 1) & " of " & (UBound($test) / 5) & " = " & $query)
		_SQLite_Exec(-1, $query) ; INSERT Data
	Next



;		ControlSetText($app_name, "", "Static1", "get all results for the run id " & $run[$i] & " - " & (($i / 2) + 1) & " of " & (UBound($run) / 2) & ", test " & (($j / 4) + 1) & " of " & (UBound($test) / 4))
;		local $result = _TestRailGetResultsIdStatusIdDefects($test[$j])

;		if (($j / 4) + 1) >= 26 Then

;			_ArrayDisplay($result)
;		Exit
;		EndIf

;		if IsArray($result) = True Then
;
;			for $k = 0 to (UBound($result) - 1) Step 4

;				$query = "INSERT INTO Result (Id,CaseId,TestId,StatusId,Defects) VALUES ('" & $result[$k + 3] & "','" & $test[$j + 1] & "','" & $test[$j] & "','" & $result[$k + 2] & "','" & $result[$k] & "');"
;				_FileWriteLog($log_filepath, "Result " & (($k / 4) + 1) & " of " & (UBound($result) / 4) & " = " & $query)
;				_SQLite_Exec(-1, $query) ; INSERT Data
;			Next
;		EndIf

;	Next

Next

#ce


; Shutdown TestRail & Jira

;GUICtrlSetData($status_input, "Closing Jira ... ")
_TestRailShutdown()
_JiraShutdown()
SplashOff()

