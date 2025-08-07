'******************************************************************************************************************************
' Name Of Business Component		:	Select Individual objects to submit
'
' Purpose							:	Select Individual objects to submit
'
' Input	Parameter					:	
'
' Output							:	True / False
'
' Remarks							:
'
' Author							:	Mohini  Deshmukh			  4 Dec 2024

'******************************************************************************************************************************
Option Explicit
'-------------------------------------------------------------------------------------------------------------------------------
'Variable Declaration
'--------------------------------------------------------------------------------------------------------------------------------
Dim obj_AWCTeamcenterHome
Dim sTempObject,sXpath,sColumnName,sTempColumnName,sData,sTempData,sColumnData
Dim iCount
Dim sTemp,sTempError
'--------------------------------------------------------------------------------------------------------------------------------
'Get AWC PLM window object from xml
'--------------------------------------------------------------------------------------------------------------------------------
Set obj_AWCTeamcenterHome=Eval(GetResource("ActiveWorkspace_OR.xml").GetValue("wpage_AWCTeamcenterHome"))
'--------------------------------------------------------------------------------------------------------------------------------
'Select the Item from navigation list
'--------------------------------------------------------------------------------------------------------------------------------
If Parameter("str_object_to_submit")<>"" Then
	Call Fn_AWC_Object_Navigation_Operations(obj_AWCTeamcenterHome,"select",Parameter("str_object_to_submit"))
End If
'--------------------------------------------------------------------------------------------------------------------------------
'Click on more command on Primary tab
'--------------------------------------------------------------------------------------------------------------------------------
If  Fn_Web_UI_WebElement_Operations("Select Individual objects to submit ","Click",obj_AWCTeamcenterHome,"wele_WorkAreaTool_MoreCommand","","","")  =False Then
	Reporter.ReportEvent micFail, "Click on [ More Command ] on primary toolbar", "Fail to click on [ More Command ] on primary toolbar"
	ExitComponent
Else
	Reporter.ReportEvent micPass, "Click on [ More Command ] on primary toolbar", "Successfully clicked on [ More Command ] on primary toolbar"
End If
Call  Fn_AWC_ReadyStatusSync(1)
'--------------------------------------------------------------------------------------------------------------------------------
'Click on Manage -> Submit to Workflow button of primary tool bar
'--------------------------------------------------------------------------------------------------------------------------------
Call Fn_AWC_Common_Business_Primary_Toolbar_Operations(obj_AWCTeamcenterHome,gManage,gSubmitToWorkflow)
'--------------------------------------------------------------------------------------------------------------------------------
'Select template from template list
'--------------------------------------------------------------------------------------------------------------------------------
If  Parameter("str_template") ="" Then
	Reporter.ReportEvent micFail, "Fail to select template", "Fail to select template as [ Template is empty ]"
	ExitComponent
Else
	sXpath=Replace(gEditBoxXpath,"~",gWfTempNm)
	Call Fn_Web_UI_WebObject_Operations("Select Individual objects to submit", "settoproperty", obj_AWCTeamcenterHome.WebEdit("wedit_ObjectPropertyEditBox_2"), "2", "xpath", sXpath)
	If Fn_Web_UI_WebEdit_Operations("Select Individual objects to submit","Click",obj_AWCTeamcenterHome, "wedit_ObjectPropertyEditBox_2", "" ) Then
		If WaitUntilExist(obj_AWCTeamcenterHome.WebElement("wele_ObjectListValuePanel"), 3, 5) Then
			Call Fn_Web_UI_WebObject_Operations("Select Individual objects to submit", "settoproperty", obj_AWCTeamcenterHome.WebElement("wele_ObjectListValue"), "2", "innertext", Parameter("str_template"))
			 If WaitUntilExist(obj_AWCTeamcenterHome.WebElement("wele_ObjectListValue"), 3, 5) Then
				Call Fn_Web_UI_WebElement_Operations("Select Individual objects to submit","Click",obj_AWCTeamcenterHome.WebElement("wele_ObjectListValue"),"","","","")
		 		Reporter.ReportEvent micPass, "Select [ " & Parameter("str_template") & " ] template from template list", "Successfully Selected [ " & Parameter("str_template") & " ] template from template list"
			 Else
			 	Reporter.ReportEvent micFail, "Select [ " & Parameter("str_template") & " ] template from template list", "Fail to Select [ " & Parameter("str_template") & " ] template from template list"
				ExitComponent
			 End  IF
			Call Fn_AWC_ReadyStatusSync(1)
		Else
			Reporter.ReportEvent micFail, "Select [ " & Parameter("str_template") & " ] template from template list", "Fail to Select [ " & Parameter("str_template") & " ] template from template list"
			ExitComponent
		End  If
	Else
		Reporter.ReportEvent micFail, "Click in  " & gWfTempNm & " list field", "Fail to Click in " & gWfTempNm & " list field as [  List does not exist ]"
		ExitComponent
	End  IF	
End  IF
wait 2
'--------------------------------------------------------------------------------------------------------------------------------
'Set Description
'--------------------------------------------------------------------------------------------------------------------------------
If Parameter("str_description") <> "" Then
	call Fn_Web_UI_WebObject_Operations("Select Individual objects to submit", "settoproperty", obj_AWCTeamcenterHome.WebEdit("wedit_ObjectTextArea"),"", "acc_name",glblDescription)
	If Fn_Web_UI_WebObject_Operations("Select Individual objects to submit", "waituntilexist", obj_AWCTeamcenterHome.WebEdit("wedit_ObjectTextArea"), "20000", "visible", "True") Then
		
		IF Fn_Web_UI_WebEdit_Operations("Select Individual objects to submit","Set",obj_AWCTeamcenterHome.WebEdit("wedit_ObjectTextArea"), "", Parameter("str_description")) Then
			Reporter.ReportEvent micPass, "Set the [ "&glblDescription&" ] field for Workflow Process Panel", "Successfully Set[  "& Parameter("str_description") &" ] for[ "&glblDescription&" ] field in Workflow Process Panel"
		Else
			Reporter.ReportEvent micFail, "Set the [ "&glblDescription&" ] field for Workflow Process Panel", "Fail to Set[  "& Parameter("str_description") &" ] for[ "&glblDescription&" ] field in Workflow Process Panel"
			ExitComponent
		End  IF     
		Call Fn_AWC_ReadyStatusSync(2)	
	Else
		Reporter.ReportEvent micFail, "Verify existance of ["&glblDescription&"] edit field.", "["&glblDescription&"] edit field does not exist."
		ExitComponent		
	End If	
End If
'--------------------------------------------------------------------------------------------------------------------------------
'Click on OK Button
'--------------------------------------------------------------------------------------------------------------------------------
If Fn_WEB_UI_WebButton_Operations("Select Individual objects to submit", "click", obj_AWCTeamcenterHome, "wbtn_CheckCompletenessOK","","","")Then
	Reporter.ReportEvent micPass, "Click on [OK ] button", "Successfully clicked on [  OK ] button"
Else
	Reporter.ReportEvent micFail, "Click on [  OK ] button", "Fail to Click on [  OK ] button"
	ExitComponent
End  If
'--------------------------------------------------------------------------------------------------------------------------------
'Verify the Notification message for submit workflow
'--------------------------------------------------------------------------------------------------------------------------------
If Parameter("str_expected_check_completeness_error")="" Then
	If obj_AWCTeamcenterHome.WebTable("wtbl_CheckCompletenessErrorDetails").exist Then
		Reporter.ReportEvent micFail, "Fail to submit to workflow", "Fail to submit to workflow as [ It shows an error ] "
		ExitComponent
	End  IF	
	'--------------------------------------------------------------------------------------------------------------------------------
ElseIf Parameter("str_expected_check_completeness_error") <>"" Then 
	'--------------------------------------------------------------------------------------------------------------------------------
	'Verify the Submit to workflow status
	'--------------------------------------------------------------------------------------------------------------------------------
	If WaitUntilExist(obj_AWCTeamcenterHome.WebTable("wtbl_CheckCompletenessErrorDetails"), 60,1) then
		sColumnName=Fn_Web_UI_WebObject_Operations("Select Individual objects to submit", "getroproperty", obj_AWCTeamcenterHome.WebTable("wtbl_CheckCompletenessErrorDetails"), "", "column names", "") 
		sTempColumnName=Split(sColumnName,";")
		
		sData=Fn_Web_UI_WebObject_Operations("Select Individual objects to submit", "getroproperty", obj_AWCTeamcenterHome.WebTable("wtbl_CheckCompletenessErrorDetails"), "", "innertext", "") 
		sTempData=Split(sData,sTempColumnName(0))
		sColumnData=Split(sTempData(1),sTempColumnName(1))
		
		Reporter.ReportEvent micPass, "Check Completeness with error objects", "Check Completeness with error as  [ " & sTempColumnName(0) & " column contains following data : " & sColumnData(0) & " ]"
		Parameter("str_ObjectName_out")=sColumnData(0)
		
		Reporter.ReportEvent micPass, "Check Completeness with error details", "Check Completeness with error as  [ " & sTempColumnName(1) & " column contains following data : " & sColumnData(1) & " ]"
		Parameter("str_ErrorDeatils_out")=sColumnData(1)
		'--------------------------------------------------------------------------------------------------------------------------------
		sTempError=Split(Parameter("str_expected_check_completeness_error"),"~")
		For iCount = 0 To ubound(sTempError)
			If Instr(1,Parameter("str_ErrorDeatils_out"),sTempError(iCount))>0 Then
				Reporter.ReportEvent micPass, "Verify the Check Completeness error", "Successfully verified the Check Completeness error [ " & sTempError(iCount) & " ] for selected object(s)"
			Else
				Reporter.ReportEvent micFail, "Verify the Check Completeness error", "Fail to verify the Check Completeness error [ " & sTempError(iCount) & " ] for selected object(s)"
				ExitComponent
			End If
		Next
		'--------------------------------------------------------------------------------------------------------------------------------
		'Close the check completeness error panel
		'--------------------------------------------------------------------------------------------------------------------------------
		If Fn_WEB_UI_WebButton_Operations("Select Individual objects to submit", "click", obj_AWCTeamcenterHome, "wbtn_CheckCompletenessClose","","","")Then
			Reporter.ReportEvent micPass, "Close [ Check Completeness Error ] dialog", "Successfully Closed [ Check Completeness Error ] dialog"
		Else
			Reporter.ReportEvent micFail, "Close [ Check Completeness Error ] dialog", "Fail to Close [ Check Completeness Error ] dialog"
			ExitComponent
		End  If
		Call Fn_AWC_ReadyStatusSync(1)
	End  If
	'--------------------------------------------------------------------------------------------------------------------------------
Else
	Reporter.ReportEvent micPass, "Submit to workflow", "Successfully performed submit to workflow operation"
End  If
'--------------------------------------------------------------------------------------------------------------------------------
If Parameter("str_WaitTime")<>"" Then
	wait Parameter("str_WaitTime")
End If
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
If Err.Number<> 0 Then
	Reporter.ReportEvent micFail, "Select Individual object to submit", "Fail to perform [ Select Individual object to submit ]  Operation due to [ " & Err.Description & " ]"
	ExitComponent
Else
	Reporter.ReportEvent micPass, "Select Individual object to submit", "Successfully performed [ Select Individual object to submit ] operation "
End If
'-----------------------------------------------------------------------------------------------------------------------------------------------------------	
'Set object nothing
Set obj_AWCTeamcenterHome=Nothing
'-----------------------------------------------------------------------------------------------------------------------------------------------------------

