'******************************************************************************************************************************
' Name Of Business Component		:	Validate Functionality of Context menu for items in a list - Right-click Cut
'
' Purpose							:	Validate Functionality of Context menu for items in a list - Right-click Cut
'
' Input	Parameter					:	
'
' Output							:	True / False
'
' Remarks							:
'
' Author							:	Pooja Bondarde 			  23 Sep 2021

'******************************************************************************************************************************
Option Explicit
'-------------------------------------------------------------------------------------------------------------------------------
'Variable Declaration
'--------------------------------------------------------------------------------------------------------------------------------
Dim obj_AWCTeamcenterHome,sTemp,sReg,sNotification,sNotication1
'--------------------------------------------------------------------------------------------------------------------------------
'Get CVS PLM window object from xml
'--------------------------------------------------------------------------------------------------------------------------------
Set obj_AWCTeamcenterHome = Eval(GetResource("ActiveWorkspace_OR.xml").GetValue("wpage_AWCTeamcenterHome"))
'--------------------------------------------------------------------------------------------------------------------------------
'Select Business object
'-----------------------------------------------------------------------------------------------------------------------------------------------------------	
sReg = ".*"
sTemp = Parameter("business_object")&sReg
Call Fn_Web_UI_WebObject_Operations("", "settoproperty", obj_AWCTeamcenterHome.WebElement("wele_Search_Result_SelectedCell"), "2", "innertext", sTemp)
If Fn_Web_UI_WebObject_Operations("", "exist", obj_AWCTeamcenterHome.WebElement("wele_Search_Result_SelectedCell"), "5", "", "")= False Then
	
	Call Fn_Web_UI_WebObject_Operations("", "settoproperty", obj_AWCTeamcenterHome.WebElement("wele_SearchResult_List_Listwithsummary"), "2", "innertext",Parameter("business_object"))
	If Fn_Web_UI_WebElement_Operations("","Click",obj_AWCTeamcenterHome.WebElement("wele_SearchResult_List_Listwithsummary"),"","","","") Then
		Reporter.ReportEvent micPass, "Select ["&Parameter("business_object") &"] from search result.", "Successfully Select ["&Parameter("business_object") &"] from search result."
	Else
		Reporter.ReportEvent micFail, "Select ["&Parameter("business_object") &"]from search result.", "Fail to Select ["&Parameter("business_object") &"] from search result."
		ExitComponent
	End  If
		
Else
	Reporter.ReportEvent micPass, "["&Parameter("business_object")&"] is already selected.", "Successfully verified ["&Parameter("business_object")&"] is already selected."
End If
 Call Fn_Common_ReadyStatusSync(obj_AWCTeamcenterHome,"")
 
 '--------------------------------------------------------------------------------------------------------------------------------
'Right Mouse button Click
'-----------------------------------------------------------------------------------------------------------------------------------------------------------	
Call Fn_Web_UI_WebObject_Operations("", "settoproperty", obj_AWCTeamcenterHome.WebElement("wele_SearchResult_List_Listwithsummary"), "2", "innertext",Parameter("business_object"))

Call Fn_Web_UI_WebObject_Operations("", "settoproperty", obj_AWCTeamcenterHome.WebElement("wele_RightMouseButton_Submenu"), "2", "innertext",Parameter("str_ContextMenu") )

If Fn_Web_UI_WebElement_Operations("","rightmousebutton",obj_AWCTeamcenterHome.WebElement("wele_SearchResult_List_Listwithsummary"),"","","","") Then
	Reporter.ReportEvent micPass, "Perform Right Mouse Button click", "Successfully Performed Right Mouse Button click"
Else
	Reporter.ReportEvent micFail, "Perform Right Mouse Button click", "Fail to Perform Right Mouse Button click"
	ExitComponent
End  If
'--------------------------------------------------------------------------------------------------------------------------------
'Wait until exist Context menu and click on it
'-----------------------------------------------------------------------------------------------------------------------------------------------------------	
If Fn_Web_UI_WebObject_Operations("", "waituntilexist", obj_AWCTeamcenterHome.WebElement("wele_RightMouseButton_Submenu"), "2", "visible","True") Then
	If Fn_Web_UI_WebElement_Operations("","Click",obj_AWCTeamcenterHome.WebElement("wele_RightMouseButton_Submenu"),"","","","") Then
		Reporter.ReportEvent micPass, "Click on ["& Parameter("str_ContextMenu") &"] submenu.", "Successfully Clicked on ["& Parameter("str_ContextMenu") &"] submenu."
	Else
		Reporter.ReportEvent micFail, "Click on ["& Parameter("str_ContextMenu") &"] submenu.", "Fail to Click on ["& Parameter("str_ContextMenu") &"] submenu."
		ExitComponent
	End  If
End  If

'--------------------------------------------------------------------------------------------------------------------------------
'Wait until exist Notification message
'-----------------------------------------------------------------------------------------------------------------------------------------------------------	

If Fn_Web_UI_WebObject_Operations("", "waituntilexist", obj_AWCTeamcenterHome.WebElement("wele_ObjectCreationNotificationMsg"), "2", "visible","True") Then
	sTemp = Fn_Web_UI_WebObject_Operations("", "getroproperty", obj_AWCTeamcenterHome.WebElement("wele_ObjectCreationNotificationMsg"), "5", "innertext","") 
	sNotification="was cut from"
	sNotication1="and added to clipboard"
	If instr(sTemp,Parameter("business_object"))<> 0 And instr(sTemp,sNotification)<> 0 And instr(sTemp,sNotication1)<> 0 And instr(sTemp,Parameter("str_Copied_Business_Object") )<> 0 Then
		Reporter.ReportEvent micPass, "verify Notification msg ["&sTemp&"]", "Successfully verified Notification msg ["&sTemp&"]"
	Else
		Reporter.ReportEvent micFail, "verify Notification msg ["&sTemp&"]", "Fail to verify Notification msg ["&sTemp&"]"
		ExitComponent
	End  If
Else
	Reporter.ReportEvent micFail, "Notification does not exist", "Fail to verify Notification exist"
	ExitComponent
End If


Parameter("business_object_out") = Parameter("business_object")
'-----------------------------------------------------------------------------------------------------------------------------------------------------------	
'Set object nothing
Set obj_AWCTeamcenterHome=Nothing
'-----------------------------------------------------------------------------------------------------------------------------------------------------------


 
