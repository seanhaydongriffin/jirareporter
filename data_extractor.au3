#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_UseUpx=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#RequireAdmin
;#AutoIt3Wrapper_usex64=n
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


; Startup SQLite

_SQLite_Startup()
ConsoleWrite("_SQLite_LibVersion=" & _SQLite_LibVersion() & @CRLF)
_SQLite_Open(@ScriptDir & "\Test Completion Reporter.sqlite")


; Startup Jira & TestRail

;GUICtrlSetData($status_input, "Starting the Jira connection ... ")
_JiraSetup()
_JiraDomainSet("https://janisoncls.atlassian.net")
_JiraLogin($jira_username, $jira_password)


;GUICtrlSetData($status_input, "Starting the TestRail connection ... ")
_TestRailDomainSet("https://janison.testrail.com")
_TestRailLogin($testrail_username, $testrail_password)


FileDelete($log_filepath)




; get all epics for the project

SplashTextOn($app_name, "get all epics for the project")
;GUICtrlSetData($status_input, "Querying Project RMS ... ")
$issue = _JiraGetSearchResultKeysSummariesAndIssueTypeNames("summary,issuetype", "project = " & $jira_project & " AND issuetype = Epic")

for $i = 0 to (UBound($issue) - 1) Step 3


	$query = "INSERT INTO Epic (Key,Summary) VALUES ('" & $issue[$i] & "','" & $issue[$i + 1] & "');"
	_FileWriteLog($log_filepath, "Epic " & ($i + 1) & " of " & UBound($issue) & " = " & $query)
	_SQLite_Exec(-1, $query) ; INSERT Data

Next

; get all stories for the project

SplashTextOn($app_name, "get all stories for the project")
$issue = _JiraGetSearchResultKeysSummariesIssueTypeNameEpicKeyRequirements("summary,issuetype,customfield_10008,labels,fixVersions,status,customfield_15255", "project = " & $jira_project & " AND issuetype in (Improvement, Story)")

for $i = 0 to (UBound($issue) - 1) Step 8

	$query = "INSERT INTO Story (Key,Summary,EpicKey,ReqID,FixVersion,Status,TestNotes) VALUES ('" & $issue[$i] & "','" & $issue[$i + 1] & "','" & $issue[$i + 3] & "','" & $issue[$i + 4] & "','" & $issue[$i + 5] & "','" & $issue[$i + 6] & "','" & $issue[$i + 7] & "');"
	_FileWriteLog($log_filepath, "Story " & ($i + 1) & " of " & UBound($issue) & " = " & $query)
	_SQLite_Exec(-1, $query) ; INSERT Data
Next

; get all tasks for the project

SplashTextOn($app_name, "get all tasks for the project")
$issue = _JiraGetSearchResultKeysSummariesIssueTypeNamesIssueLinks("summary,issuetype,issuelinks", "project = " & $jira_project & " AND issuetype = Task")

for $i = 0 to (UBound($issue) - 1) Step 4

	$query = "INSERT INTO Task (Key,Summary,StoryKey) VALUES ('" & $issue[$i] & "','" & $issue[$i + 1] & "','" & $issue[$i + 3] & "');"
	_FileWriteLog($log_filepath, "Task " & ($i + 1) & " of " & UBound($issue) & " = " & $query)
	_SQLite_Exec(-1, $query) ; INSERT Data
Next


#cs
; get all bugs for the project

SplashTextOn($app_name, "get all bugs for the project")
;$issue = _JiraGetSearchResultBugs("summary,reporter,fixVersions,resolution,customfield_12000,priority,labels,versions,status,customfield_10007,environment,customfield_11602,assignee", "project = " & $jira_project & " AND issuetype = Bug")
$issue = _JiraGetSearchResultBugs("summary,reporter,fixVersions,resolution,customfield_12000,priority,labels,versions,status,customfield_10007,environment,assignee", "project = " & $jira_project & " AND issuetype = Bug")

for $i = 0 to (UBound($issue) - 1) Step 4

	$query = "INSERT INTO Bug (Key,Summary,Reporter,Assignee,Status,Priority,AffectsVersions,FixVersions,Resolution,Labels,Environment,ScrumTeam,Sprint) VALUES ('" & $issue[$i] & "','" & $issue[$i + 1] & "','" & $issue[$i + 3] & "');"
	_FileWriteLog($log_filepath, "Task " & ($i + 1) & " of " & UBound($issue) & " = " & $query)
	_SQLite_Exec(-1, $query) ; INSERT Data
Next
#ce

; get all test cases for the project

SplashTextOn($app_name, "get all test cases for the project")
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

SplashTextOn($app_name, "get all test statuses", 500, 400, Default, Default, 16)
local $status = _TestRailGetStatusesIdLabel()

for $i = 0 to (UBound($status) - 1) Step 2

	$query = "INSERT INTO TestStatus (Id,Label) VALUES ('" & $status[$i] & "','" & $status[$i + 1] & "');"
	_FileWriteLog($log_filepath, "Status " & ($i + 1) & " of " & (UBound($status) / 2) & " = " & $query)
	_SQLite_Exec(-1, $query) ; INSERT Data
Next

; get all runs for the project (not in test plans)

SplashTextOn($app_name, "get all runs for the project (not in test plans)", 500, 400, Default, Default, 16)
local $run = _TestRailGetRunsIdName(43)

for $i = 0 to (UBound($run) - 1) Step 2

	$query = "INSERT INTO Run (Id,Name) VALUES ('" & $run[$i] & "','" & $run[$i + 1] & "');"
	_FileWriteLog($log_filepath, "Run " & (($i / 2) + 1) & " of " & (UBound($run) / 2) & " = " & $query)
	_SQLite_Exec(-1, $query) ; INSERT Data

	SplashTextOn($app_name, "get all tests for run id " & $run[$i] & " - " & (($i / 2) + 1) & " of " & (UBound($run) / 2), 500, 400, Default, Default, 16)
	local $test = _TestRailGetTestsIdTitleCaseIdRunId($run[$i])

	for $j = 0 to (UBound($test) - 1) Step 4

		$query = "INSERT INTO Test (Id,Title,TestCaseId,RunId) VALUES ('" & $test[$j] & "','" & $test[$j + 3] & "','" & $test[$j + 1] & "','" & $test[$j + 2] & "');"
		_FileWriteLog($log_filepath, "Test " & (($j / 4) + 1) & " of " & (UBound($test) / 4) & " = " & $query)
		_SQLite_Exec(-1, $query) ; INSERT Data
	Next

	SplashTextOn($app_name, "get all results for run id " & $run[$i] & " - " & (($i / 2) + 1) & " of " & (UBound($run) / 2), 500, 400, Default, Default, 16)
	local $result = _TestRailGetResultsForRunIdStatusIdCreatedOnDefects($run[$i])

	for $j = 0 to (UBound($result) - 1) Step 5

		$result[$j + 4] = StringReplace($result[$j + 4], "null", "")
		$result[$j + 4] = StringReplace($result[$j + 4], """", "")

		$query = "INSERT INTO Result (Id,TestId,StatusId,CreatedOn,Defects) VALUES ('" & $result[$j] & "','" & $result[$j + 1] & "','" & $result[$j + 2] & "','" & $result[$j + 3] & "','" & $result[$j + 4] & "');"
		_FileWriteLog($log_filepath, "Result " & (($j / 5) + 1) & " of " & (UBound($test) / 5) & " = " & $query)
		_SQLite_Exec(-1, $query) ; INSERT Data
	Next



;		SplashTextOn($app_name, "get all results for the run id " & $run[$i] & " - " & (($i / 2) + 1) & " of " & (UBound($run) / 2) & ", test " & (($j / 4) + 1) & " of " & (UBound($test) / 4), 500, 400, Default, Default, 16)
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




; Shutdown TestRail & Jira

;GUICtrlSetData($status_input, "Closing Jira ... ")
_TestRailShutdown()
_JiraShutdown()
SplashOff()

