'******************************************************************************************************************************
' Name Of Business Component		:	Start Active Workspace
'
' Purpose							:	Invoke the application and Login to Active Workspace
'
' Input	Parameter					:	Parameter 1: str_Instance 
'
'										Parameter 2: str_UserName 
'
'										Parameter 3: str_Password 
'
' Output							:	True / False
'
' Remarks							:
'
' Author							:	Mohini Deshmukh 			  29 July 2020

'******************************************************************************************************************************
Option Explicit

Dim testPath
testPath =Environment("TestDir")
Reporter.ReportEvent micPass, "Test directory", testPath


Dim repoCount, i
repoCount = RepositoriesCollection.Count

For i = 1 To repoCount
    Reporter.ReportEvent micPass, "FOUND OBJECT REPOSITORY:", "Repository " & i & ": " & RepositoriesCollection.Item(i)
Next
'-------------------------------------------------------------------------------------------------------------------------------
'Variable Declaration
'-------------------------------------------------------------------------------------------------------------------------------
Dim objEdgeBrowser,objBrowsers,objApp,obj_AWCTeamcenterHome,Processes,objBrowsersNew,objWshShell
Dim iCount,iCounter
Dim sVersion,sGroup,sRole,sTempValue,Process,targetUrl,sPassword,sUserName,sHeaderText,sTemp,sTempUserNm
Dim myProcess1,myProcess
'--------------------------------------------------------------------------------------------------------------------------------
'Get AWC PLM window object from xml
'--------------------------------------------------------------------------------------------------------------------------------
Set obj_AWCTeamcenterHome=Eval(SearchAndLoadResourceByName("ActiveWorkspace_OR.xml").GetValue("wpage_AWCTeamcenterHome"))

Set obj_AWCTeamcenterHome=Nothing

