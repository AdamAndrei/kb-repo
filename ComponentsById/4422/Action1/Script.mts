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

Set objWshShell = CreateObject("WScript.shell")
'--------------------------------------------------------------------------------------------------------------------------------
'Terminate all sprinter processes 
'--------------------------------------------------------------------------------------------------------------------------------
myProcess="Sprinter.exe"    
myProcess1="SprinterAgent.exe" 

Set Processes = GetObject("winmgmts:").InstancesOf("Win32_Process")

For Each Process In Processes
	'--------------------------------------------------------------------------------------------------------------------------------
	'Check the sprinter process if exist then terminate it 
	'--------------------------------------------------------------------------------------------------------------------------------
	  If StrComp(Process.Name, myProcess, vbTextCompare) = 0 Then 
	      Process.Terminate()                                   
	  End If
	    '--------------------------------------------------------------------------------------------------------------------------------
	    'Check the sprinteragent process if exist then terminate it
	    '--------------------------------------------------------------------------------------------------------------------------------
	    If StrComp(Process.Name, myProcess1, vbTextCompare) = 0 Then 
	       Process.Terminate()                                   
	    End If
	 
Next
'--------------------------------------------------------------------------------------------------------------------------------
'Creating Description object
'--------------------------------------------------------------------------------------------------------------------------------
Set objEdgeBrowser = Description.Create()
objEdgeBrowser("micclass").Value = "Browser"
objEdgeBrowser("version").Value="Chromium Edge.*"

sVersion="version:=Chromium Edge.*"

'Set objBrowsers = desktop.ChildObjects(objEdgeBrowser)
'wait 1
'For iCount =0 to objBrowsers.count-1

	Set objBrowsersNew = desktop.ChildObjects(objEdgeBrowser)	
	If objBrowsersNew.count>0 Then
		If obj_AWCTeamcenterHome.exist Then
			If obj_AWCTeamcenterHome.WebButton("wbtn_Your_Profile").exist Then
	 			obj_AWCTeamcenterHome.WebButton("wbtn_Your_Profile").Click
	 			wait 2
		 		If Fn_Web_UI_WebElement_Operations("Signout from Active Workspace application","Click",obj_AWCTeamcenterHome.WebButton("wbtn_SignOut"),"","","","")=True Then
					Reporter.ReportEvent micPass, "Click on [ Sign out ] button in Active Workspace", "Successfully Clicked on [ Sign out ] button in Active Workspace"
				Else
					Reporter.ReportEvent micFail, "Click on [ Sign out ] button in Active Workspace", "Fail to click on [ Sign out ] button in Active Workspace"
					ExitComponent
				End  If
			End  IF	
			SystemUtil.CloseProcessByName "msedge.exe"
		Else
			SystemUtil.CloseProcessByName "msedge.exe"
		End  IF	
		wait 2
	End  IF	

Set Processes = GetObject("winmgmts:").InstancesOf("Win32_Process")
myProcess="msedge.exe"
For Each Process In Processes
	'--------------------------------------------------------------------------------------------------------------------------------
	'Check the sprinter process if exist then terminate it 
	'--------------------------------------------------------------------------------------------------------------------------------
	  If StrComp(Process.Name, myProcess, vbTextCompare) = 0 Then 
	      Process.Terminate()                                   
	  End If
Next

If Browser(sVersion,"CreationTime:=0").Exist(3)  Then
	Reporter.ReportEvent micFail, "Close the all instances of Edge Chromium Browser", "Fail to Close the all instances of Edge Chromium Browser"					
	ExitComponent
Else
	Reporter.ReportEvent micPass, "Close the all instances of Edge Chromium Browser", "Successfully Close the all instances of Edge Chromium Browser"					
End If
'Next
'--------------------------------------------------------------------------------------------------------------------------------
'Select Instance Name
'--------------------------------------------------------------------------------------------------------------------------------
Select Case lCase(Parameter("str_Instance"))
	'--------------------------------------------------------------------------------------------------------------------------------
	Case "project"
		targetUrl = SearchAndLoadResourceByName("ActiveWorkspace2406_Data.xml").GetValue("PROJECT")
	'--------------------------------------------------------------------------------------------------------------------------------
	Case "prod"
		targetUrl = SearchAndLoadResourceByName("ActiveWorkspace_Data.xml").GetValue("PROD")
	'--------------------------------------------------------------------------------------------------------------------------------
	Case "prod-nonsso"
		targetUrl = SearchAndLoadResourceByName("ActiveWorkspace_Data.xml").GetValue("PROD-NonSSO")
	'--------------------------------------------------------------------------------------------------------------------------------
	Case "int2406" 
		targetUrl = SearchAndLoadResourceByName("ActiveWorkspace2406_Data.xml").GetValue("INT2406")
	'--------------------------------------------------------------------------------------------------------------------------------	
	Case "int2" 
		targetUrl = SearchAndLoadResourceByName("ActiveWorkspace_Data.xml").GetValue("INT2")
	'--------------------------------------------------------------------------------------------------------------------------------	
	Case "int1" 
		targetUrl = SearchAndLoadResourceByName("ActiveWorkspace_Data.xml").GetValue("INT1")
	'--------------------------------------------------------------------------------------------------------------------------------
	Case "pre"
		targetUrl = SearchAndLoadResourceByName("ActiveWorkspace_Data.xml").GetValue("PRE")
	'--------------------------------------------------------------------------------------------------------------------------------
	Case "pre-nonsso"
		targetUrl = SearchAndLoadResourceByName("ActiveWorkspace_Data.xml").GetValue("PRE-NONSSO")
	'--------------------------------------------------------------------------------------------------------------------------------
	Case "cloud"
		targetUrl = SearchAndLoadResourceByName("ActiveWorkspace_Data.xml").GetValue("CLOUD")
	'--------------------------------------------------------------------------------------------------------------------------------
	Case "int2-nonsso"
		targetUrl = SearchAndLoadResourceByName("ActiveWorkspace_Data.xml").GetValue("INT2-NONSSO")
	'--------------------------------------------------------------------------------------------------------------------------------
	Case Else 
			ExitComponent
	'--------------------------------------------------------------------------------------------------------------------------------			
End Select
'--------------------------------------------------------------------------------------------------------------------------------
'Invoke the Active Workspace (AWC) application 
'--------------------------------------------------------------------------------------------------------------------------------
'systemutil.Run "msedge.exe",targetUrl
Reporter.ReportEvent micPass, "Launch youtube", "hope it works"
InvokeApplication "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe "& "https://www.youtube.com"
Reporter.ReportEvent micPass, "Launched youtube", "hope it works"
Wait 5
Set obj_AWCTeamcenterHome=Nothing
Set objWshShell=Nothing
