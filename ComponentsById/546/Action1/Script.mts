'******************************************************************************************
' Name Of Business Component		:	Right Mouse Button - Add - Add New Manufacturing Document - UFT\Right Mouse Button
'
' Purpose							:	Right Mouse Button - Add - Add New Manufacturing Document - UFT\Right Mouse Button
'
' Input	Parameter					:	
'
' Output							:	True / False
'
' Remarks							:
'
' Author							:	Mohini Deshmukh			  27 April 2020

'******************************************************************************************
Dim obj_CVSPLM,obj_ManufacturingPart,obj_QueryAndRelate
Dim sStatusMsg,sData
Dim sID,sManDoc
Dim iRowNumber,iCount,iRowCount
'----------------------------------------------------------------------------------------------------------------------
'Get CVS PLM window object from xml
'----------------------------------------------------------------------------------------------------------------------
Set obj_CVSPLM=Eval(GetResource("CreatePart_OR_XML.xml").GetValue("jwnd_CVSPLM"))
Set obj_ManufacturingPart=Eval(GetResource("CreatePart_OR_XML.xml").GetValue("jdlg_Create_Part"))
Set obj_QueryAndRelate=Eval(GetResource("CreatePart_OR_XML.xml").GetValue("jdlg_QueryAndRelate"))
'----------------------------------------------------------------------------------------------------------------------
'Verify R & D is Opened
'----------------------------------------------------------------------------------------------------------------------
Call Fn_UI_Object_Operations("RMB Add New Manufacturing Document","settoproperty", obj_CVSPLM.JavaButton("jbtn_Common"),"label", glblRND, DEFAULT_MIN_TIMEOUT) 
If Waituntilexist(obj_CVSPLM.JavaButton("jbtn_Common"), 2, 5) Then
	Reporter.ReportEvent micPass, "Existence of [ R & D ] button", "[ R & D ] button is Exist"
Else
	Reporter.ReportEvent micFail, "Existence of [ R & D ] button", "[ R & D ] button is Not Exist"
End IF
'----------------------------------------------------------------------------------------------------------------------
'Select Part Button
'----------------------------------------------------------------------------------------------------------------------
Call Fn_UI_Object_Operations("RMB Add New Manufacturing Document","settoproperty", obj_CVSPLM.JavaButton("jbtn_Common"),"label", gbtnParts, DEFAULT_MIN_TIMEOUT) 
If Fn_UI_JavaButton_Operations("RMB Add New Manufacturing Document", "Click", obj_CVSPLM.JavaButton("jbtn_Common"), "")=False then
	Reporter.ReportEvent micFail, "Click on [ " & gbtnParts & " ] button", "Fail to click on [ " & gbtnParts & " ] button"
Else
	Reporter.ReportEvent micPass, "Click on [ " & gbtnParts & " ] button", "Successfully clicked on [ " & gbtnParts & " ] button"
End If

'Select Query Parts menu
Call Fn_UI_Object_Operations("RMB Add New Manufacturing Document","settoproperty", obj_CVSPLM.JavaButton("jbtn_HoverButton"),"label", gbtnQueryParts, DEFAULT_MIN_TIMEOUT) 
Call Fn_UI_JavaButton_Operations("RMB Add New Manufacturing Document", "Click", obj_CVSPLM.JavaButton("jbtn_HoverButton"), "")
'----------------------------------------------------------------------------------------------------------------------
'Click on Clear Button
'----------------------------------------------------------------------------------------------------------------------
Call Fn_UI_Object_Operations("RMB Add New Manufacturing Document","settoproperty", obj_CVSPLM.JavaButton("jbtn_Common"),"label", gbtnClear, DEFAULT_MIN_TIMEOUT) 
If Waituntilexist(obj_CVSPLM.JavaButton("jbtn_Common"), 2, 5)  Then
	If Fn_UI_JavaButton_Operations("RMB Add New Manufacturing Document", "Click", obj_CVSPLM.JavaButton("jbtn_Common"), "")=False Then
		Reporter.ReportEvent micFail, "Click on [ " & gbtnClear & " ] button", "Fail to click on [ " & gbtnClear & " ] button"
		ExitComponent
	Else
		Reporter.ReportEvent micPass, "Click on [ " & gbtnClear & " ] button", "Successfully clicked on [ " & gbtnClear & " ] button"	
	End  IF
End  If	
'----------------------------------------------------------------------------------------------------------------------
'Set ID value
'----------------------------------------------------------------------------------------------------------------------
If Parameter("str_Part_Revision") <> "" Then
	sID=Split(Parameter("str_Part_Revision"), ",", -1, 1)
	Call Fn_UI_Object_Operations("RMB Add New Manufacturing Document","settoproperty", obj_CVSPLM.JavaEdit("jedt_Common"),"attached text", glblID, DEFAULT_MIN_TIMEOUT) 
	If Waituntilexist(obj_CVSPLM.JavaEdit("jedt_Common"), 2, 5)  Then
		If Fn_UI_JavaEdit_Operations("RMB Add New Manufacturing Document", "Set",  obj_CVSPLM, "jedt_Common", sID(0) )=False Then
			Reporter.ReportEvent micFail, "Set [ " & sID(0) & " ] value in edit box", "Fail to set [ " & glblID & " ] value in edit box"
			ExitComponent
		Else
			Reporter.ReportEvent micPass, "Set [ " & sID(0) & " ] value in edit box", "Successfully set [ " & sID(0) & " ] in [ " & glblID & " ] edit box"		
		End If
	End  If	
End If
'----------------------------------------------------------------------------------------------------------------------
'Click on OK button
'----------------------------------------------------------------------------------------------------------------------
Call Fn_UI_Object_Operations("RMB Add New Manufacturing Document","settoproperty", obj_CVSPLM.JavaButton("jbtn_Common"),"label", gbtnOK, DEFAULT_MIN_TIMEOUT) 
If Fn_UI_JavaButton_Operations("RMB Add New Manufacturing Document", "Click", obj_CVSPLM.JavaButton("jbtn_Common"), "")=False then
	Reporter.ReportEvent micFail, "Click on [ " & gbtnParts & " ] button", "Fail to click on [ " & gbtnOK & " ] button"
Else
	Reporter.ReportEvent micPass, "Click on [ " & gbtnParts & " ] button", "Successfully clicked on [ " & gbtnOK & " ] button"
End If
wait 5
'----------------------------------------------------------------------------------------------------------------------
'Check Status Message for query result
'---------------------------------------------------------------------------------------------------------------------
Call Fn_UI_Object_Operations("RMB Add New Manufacturing Document","settoproperty", obj_CVSPLM.JavaStaticText("jstxt_StatusMessage"),"label", gFoundItems, DEFAULT_MIN_TIMEOUT) 
If Fn_UI_Object_Operations("RMB Add New Manufacturing Document","exist", obj_CVSPLM.JavaStaticText("jstxt_StatusMessage"),"", "", DEFAULT_MAX_TIMEOUT) = False Then
	Reporter.ReportEvent micFail, "Verify [ " & gFoundItems & " ] value in status message", "Fail to Verify [ " & gFoundItems & " ] value in Status message"
	ExitComponent
Else
	sStatusMsg = Fn_UI_Object_Operations("RMB Add New Manufacturing Document","getroproperty", obj_CVSPLM.JavaStaticText("jstxt_StatusMessage"),"label", "", DEFAULT_MAX_TIMEOUT)
	If cint(Split(sStatusMsg," ",-1,1)(1))>0 Then
		Reporter.ReportEvent micPass, "Verify search result", "Successfully verified [ " & sStatusMsg & " ] values in searched result"
	Else
		Reporter.ReportEvent micPass, "Verify search result", "No item found in search result"
		Exitcomponent
	End If	
End If

'Select Relations tab for verifying the status message for RMB Add New Manufacturing Document (which we need to verify in further steps to verify the status messages)
Call Fn_UI_JavaTab_Operations("RMB Add New Manufacturing Document", "select",obj_CVSPLM,"jtab_FrameContainer",gtabRelations)
'----------------------------------------------------------------------------------------------------------------------
'Select the Searched Part in Result Tab and Perform RMB Add ...:Add New Manufacturing Document
'---------------------------------------------------------------------------------------------------------------------
If Fn_UI_JavaTable_Operations("RMB Add New Manufacturing Document","popupmenuselectext",obj_CVSPLM,"jtbl_Results","","Name",Parameter("str_Part_Revision"),"","Add ...:Add New Manufacturing Document")=False then
	Reporter.ReportEvent micFail, "Select Part in result Window and  Perform RMB [ Add -> Add New Manufacturing Document ]", "Fail to Select Part in result Window and Perform RMB [ Add -> Add New Manufacturing Document ]"
	ExitComponent
Else
	Reporter.ReportEvent micPass, "Select Part in result Window and  Perform RMB [ Add -> Add New Manufacturing Document ]", "Successfully Select Part in result Window and Perform RMB [ Add -> Add New Manufacturing Document ]"
End If
'----------------------------------------------------------------------------------------------------------------------
'Existence of Create Manufacturing Document dialog  Tc
'----------------------------------------------------------------------------------------------------------------------
If Waituntilexist(obj_ManufacturingPart, 2, 5)  Then
	Reporter.ReportEvent micPass, "Existence of RMB Add New Manufacturing Document Dialog", "[ RMB Add New Manufacturing Document ] Dialog Existence"
Else
	Reporter.ReportEvent micFail, "Existence of RMB Add New Manufacturing Document Dialog", "[ RMB Add New Manufacturing Document ] Dialog Does not Exist"
	ExitComponent
End If
'----------------------------------------------------------------------------------------------------------------------
'Set Language
'----------------------------------------------------------------------------------------------------------------------
If Parameter("str_language") <> "" Then
	Call Fn_UI_Object_Operations("RMB Add New Manufacturing Document","settoproperty", obj_ManufacturingPart.JavaStaticText("jstxt_Part"),"label", glblLanguage, DEFAULT_MIN_TIMEOUT) 
	If Fn_UI_JavaList_Operations("RMB Add New Manufacturing Document", "select", obj_ManufacturingPart,"jlst_Part_2",Parameter("str_language"), "", "")=False Then
		Reporter.ReportEvent micFail, "Select [ " & glblLanguage & " ] from list box", "Fail to Select [ " & glblTitleID & " ] from list box"
		ExitComponent
	Else
		Reporter.ReportEvent micPass, "Select [ " & glblLanguage & " ] from list box", "Successfully select [ " & Parameter("str_language") & " ] from [ " & glblLanguage & " ] from list box"		
	End If
End If
obj_ManufacturingPart.RefreshObject
'----------------------------------------------------------------------------------------------------------------------
'Set Document Group
'----------------------------------------------------------------------------------------------------------------------
If Parameter("str_document_group") <> "" Then
	Call Fn_UI_Object_Operations("RMB Add New Manufacturing Document","settoproperty", obj_ManufacturingPart.JavaStaticText("jstxt_Part"),"label", glblDocumentGroup, DEFAULT_MIN_TIMEOUT) 
	If Fn_UI_JavaList_Operations("RMB Add New Manufacturing Document", "select", obj_ManufacturingPart,"jlst_Part_2",Parameter("str_document_group"), "", "")=False Then
		Reporter.ReportEvent micFail, "Select [ " & glblDocumentGroup & " ] from list box", "Fail to Select [ " & glblDocumentGroup & " ] from list box"
		ExitComponent
	Else
		Reporter.ReportEvent micPass, "Select [ " & glblDocumentGroup & " ] from list box", "Successfully select [ " & Parameter("str_document_group") & " ] from [ " & glblDocumentGroup & " ] from list box"		
	End If
End If
obj_ManufacturingPart.RefreshObject
'----------------------------------------------------------------------------------------------------------------------
'Set Document Kind
'----------------------------------------------------------------------------------------------------------------------
If Parameter("str_document_kind") <> "" Then
	Call Fn_UI_Object_Operations("RMB Add New Manufacturing Document","settoproperty", obj_ManufacturingPart.JavaStaticText("jstxt_Part"),"label", glblDocumentKind, DEFAULT_MIN_TIMEOUT) 
	If Fn_UI_JavaList_Operations("RMB Add New Manufacturing Document", "select", obj_ManufacturingPart,"jlst_Part_2",Parameter("str_document_kind"), "", "")=False Then
		Reporter.ReportEvent micFail, "Select [ " & glblDocumentKind & " ] from list box", "Fail to Select [ " & glblDocumentKind & " ] from list box"
		ExitComponent
	Else
		Reporter.ReportEvent micPass, "Select [ " & glblDocumentKind & " ] from list box", "Successfully select [ " & Parameter("str_document_kind") & " ] from [ " & glblDocumentKind & " ] from list box"		
	End If
End If
'----------------------------------------------------------------------------------------------------------------------
'Set Use Status
'----------------------------------------------------------------------------------------------------------------------
If Parameter("str_use_status") <> "" Then
	Call Fn_UI_Object_Operations("RMB Add New Manufacturing Document","settoproperty", obj_ManufacturingPart.JavaStaticText("jstxt_Part"),"label", glblUseStatus, DEFAULT_MIN_TIMEOUT) 
	If Fn_UI_JavaList_Operations("RMB Add New Manufacturing Document", "type", obj_ManufacturingPart,"jlst_Part_2",Parameter("str_use_status"), "", "")=False Then
		Reporter.ReportEvent micFail, "Select [ " & glblUseStatus & " ] from list box", "Fail to Select [ " & glblUseStatus & " ] from list box"
		ExitComponent
	Else
		Reporter.ReportEvent micPass, "Select [ " & glblUseStatus & " ] from list box", "Successfully select [ " & Parameter("str_use_status") & " ] from [ " & glblUseStatus & " ] from list box"		
	End If
End If
'----------------------------------------------------------------------------------------------------------------------
'Set CVS Company Name
'----------------------------------------------------------------------------------------------------------------------
If Parameter("str_cvs_company_name") <> "" Then
	Call Fn_UI_Object_Operations("RMB Add New Manufacturing Document","settoproperty", obj_ManufacturingPart.JavaStaticText("jstxt_Part"),"label", glbCVSCompanyName, DEFAULT_MIN_TIMEOUT) 
	If Fn_UI_JavaList_Operations("RMB Add New Manufacturing Document", "type", obj_ManufacturingPart,"jlst_Part_2",Parameter("str_cvs_company_name"), "", "")=False Then
		Reporter.ReportEvent micFail, "Select [ " & glbCVSCompanyName & " ] from list box", "Fail to Select [ " & glbCVSCompanyName & " ] from list box"
		ExitComponent
	Else
		Reporter.ReportEvent micPass, "Select [ " & glbCVSCompanyName & " ] from list box", "Successfully select [ " & Parameter("str_cvs_company_name") & " ] from [ " & glbCVSCompanyName & " ] from list box"		
	End If
End If
'----------------------------------------------------------------------------------------------------------------------
'Set Details of Revision
'----------------------------------------------------------------------------------------------------------------------
If Parameter("str_detail_of_revision") <> "" Then
	Call Fn_UI_Object_Operations("RMB Add New Manufacturing Document","settoproperty", obj_ManufacturingPart.JavaEdit("jedt_Part"),"attached text", glblDetailOfRevision, DEFAULT_MIN_TIMEOUT) 
	If Fn_UI_JavaEdit_Operations("RMB Add New Manufacturing Document", "Set",  obj_ManufacturingPart, "jedt_Part", Parameter("str_detail_of_revision") )=False Then
		Reporter.ReportEvent micFail, "Set [ " & glblDetailOfRevision & " ] value in edit box", "Fail to set [ " & glblDetailOfRevision & " ] value in edit box"
		ExitComponent
	Else
		Reporter.ReportEvent micPass, "Set [ " & glblDetailOfRevision & " ] value in edit box", "Successfully set [ " & Parameter("str_detail_of_revision") & " ] in [ " & glblDetailOfRevision & " ] from edit box"		
	End If
End If
'----------------------------------------------------------------------------------------------------------------------
'Set Title ID
'----------------------------------------------------------------------------------------------------------------------
If Parameter("str_title_ID") <> "" Then
	Call Fn_UI_Object_Operations("RMB Add New Manufacturing Document","settoproperty", obj_ManufacturingPart.JavaStaticText("jstxt_Part"),"label", glblTitleID, DEFAULT_MIN_TIMEOUT) 
	If Fn_UI_JavaList_Operations("RMB Add New Manufacturing Document", "type", obj_ManufacturingPart,"jlst_Part_2",Parameter("str_title_ID"), "", "")=False Then
		Reporter.ReportEvent micFail, "Select [ " & glblTitleID & " ] from list box", "Fail to Select [ " & glblTitleID & " ] from list box"
		ExitComponent
	Else
		Reporter.ReportEvent micPass, "Select [ " & glblTitleID & " ] from list box", "Successfully select [ " & Parameter("str_title_ID") & " ] from [ " & glblTitleID & " ] from list box"		
	End If
End If
'----------------------------------------------------------------------------------------------------------------------
'Set Title 1
'----------------------------------------------------------------------------------------------------------------------
If Parameter("str_title_1") <> "" Then
	Call Fn_UI_Object_Operations("RMB Add New Manufacturing Document","settoproperty", obj_ManufacturingPart.JavaEdit("jedt_Part"),"attached text", glblTitle1, DEFAULT_MIN_TIMEOUT) 
	If Fn_UI_JavaEdit_Operations("RMB Add New Manufacturing Document", "Set",  obj_ManufacturingPart, "jedt_Part", Parameter("str_title_1") )=False Then
		Reporter.ReportEvent micFail, "Set [ " & glblTitle1 & " ] value in edit box", "Fail to set [ " & glblTitle2 & " ] value in edit box"
		ExitComponent
	Else
		Reporter.ReportEvent micPass, "Set [ " & glblTitle1 & " ] value in edit box", "Successfully set [ " & Parameter("str_title_1") & " ] in [ " & glblTitle1 & " ] from edit box"		
	End If
End If
'----------------------------------------------------------------------------------------------------------------------
'Set Title 2
'----------------------------------------------------------------------------------------------------------------------
If Parameter("str_title_2") <> "" Then
	Call Fn_UI_Object_Operations("RMB Add New Manufacturing Document","settoproperty", obj_ManufacturingPart.JavaEdit("jedt_Part"),"attached text", glblTitle2, DEFAULT_MIN_TIMEOUT) 
	If Fn_UI_JavaEdit_Operations("RMB Add New Manufacturing Document", "Set",  obj_ManufacturingPart, "jlst_Part2", Parameter("str_title_2") )=False Then
		Reporter.ReportEvent micFail, "Set [ " & glblTitle2 & " ] value in edit box", "Fail to set [ " & glblTitle2 & " ] value in edit box"
		ExitComponent
	Else
		Reporter.ReportEvent micPass, "Set [ " & glblTitle2 & " ] value in edit box", "Successfully set [ " & Parameter("str_title_2") & " ] in [ " & glblTitle2 & " ] from edit box"		
	End If
End If
'----------------------------------------------------------------------------------------------------------------------
'Set Title 3
'----------------------------------------------------------------------------------------------------------------------
If Parameter("str_title_3") <> "" Then
	Call Fn_UI_Object_Operations("RMB Add New Manufacturing Document","settoproperty", obj_ManufacturingPart.JavaEdit("jedt_Part"),"attached text", glblTitle3, DEFAULT_MIN_TIMEOUT) 
	If Fn_UI_JavaEdit_Operations("RMB Add New Manufacturing Document", "Set",  obj_ManufacturingPart, "jedt_Part", Parameter("str_title_3") )=False Then
		Reporter.ReportEvent micFail, "Set [ " & glblTitle3 & " ] value in edit box", "Fail to set [ " & glblTitle3 & " ] value in edit box"
		ExitComponent
	Else
		Reporter.ReportEvent micPass, "Set [ " & glblTitle3 & " ] value in edit box", "Successfully set [ " & Parameter("str_title_3") & " ] in [ " & glblTitle3 & " ] from edit box"		
	End If
End If
'----------------------------------------------------------------------------------------------------------------------
'Set Title 4
'----------------------------------------------------------------------------------------------------------------------
If Parameter("str_title_4") <> "" Then
	Call Fn_UI_Object_Operations("RMB Add New Manufacturing Document","settoproperty", obj_ManufacturingPart.JavaEdit("jedt_Part"),"attached text", glblTitle4, DEFAULT_MIN_TIMEOUT) 
	If Fn_UI_JavaEdit_Operations("RMB Add New Manufacturing Document", "Set",  obj_ManufacturingPart, "jedt_Part", Parameter("str_title_4") )=False Then
		Reporter.ReportEvent micFail, "Set [ " & glblTitle4 & " ] value in edit box", "Fail to set [ " & glblTitle4 & " ] value in edit box"
		ExitComponent
	Else
		Reporter.ReportEvent micPass, "Set [ " & glblTitle4 & " ] value in edit box", "Successfully set [ " & Parameter("str_title_4") & " ] in [ " & glblTitle4 & " ] from edit box"		
	End If
End If
'----------------------------------------------------------------------------------------------------------------------
'Set Branded For
'----------------------------------------------------------------------------------------------------------------------
If Parameter("str_branded_for") <> "" Then
	Call Fn_UI_Object_Operations("RMB Add New Manufacturing Document","settoproperty", obj_ManufacturingPart.JavaStaticText("jstxt_Part"),"label", glblBrandedFor, DEFAULT_MIN_TIMEOUT) 
	If Fn_UI_JavaList_Operations("RMB Add New Manufacturing Document", "type", obj_ManufacturingPart,"jlst_Part_2",Parameter("str_branded_for"), "", "")=False Then
		Reporter.ReportEvent micFail, "Select [ " & glblBrandedFor & " ] from list box", "Fail to Select [ " & glblBrandedFor & " ] from list box"
		ExitComponent
	Else
		Reporter.ReportEvent micPass, "Select [ " & glblBrandedFor & " ] from list box", "Successfully select [ " & Parameter("str_branded_for") & " ] from [ " & glblBrandedFor & " ] from list box"			
	End If
End If
'******************** Select Access tab **************************
If Fn_UI_JavaTab_Operations("RMB Add New Manufacturing Document", "select",obj_ManufacturingPart,"jtab_Part",gtabAccess)=False Then
	Reporter.ReportEvent micFail, "Select [ " & gtabAccess & " ] tab", "Fail to Select [ " & gtabAccess & " ] tab"
	ExitComponent
Else
	Reporter.ReportEvent micPass, "Select [ " & gtabAccess & " ] tab", "Successfully select [ " & gtabAccess & " ] tab"				
End If 
'----------------------------------------------------------------------------------------------------------------------
'Set Team
'----------------------------------------------------------------------------------------------------------------------
If Parameter("str_team") <> "" Then
	Call Fn_UI_Object_Operations("RMB Add New Manufacturing Document","settoproperty", obj_ManufacturingPart.JavaStaticText("jstxt_Part"),"label", glblReqTeam, DEFAULT_MIN_TIMEOUT) 
	If Fn_UI_JavaList_Operations("RMB Add New Manufacturing Document", "type", obj_ManufacturingPart,"jlst_Part_2",Parameter("str_team"), "", "")=False Then
		Reporter.ReportEvent micFail, "Select [ " & glblReqTeam & " ] from list box", "Fail to Select [ " & glblReqTeam & " ] from list box"
		ExitComponent
	Else
		Reporter.ReportEvent micPass, "Select [ " & glblReqTeam & " ] from list box", "Successfully select [ " & Parameter("str_team") & " ] from [ " & glblReqTeam & " ] from list box"			
	End If
End  If	
'----------------------------------------------------------------------------------------------------------------------
'Set Responsible
'----------------------------------------------------------------------------------------------------------------------
If Parameter("str_responsible") <> "" Then
	Call Fn_UI_Object_Operations("RMB Add New Manufacturing Document","settoproperty", obj_ManufacturingPart.JavaStaticText("jstxt_Part"),"label", glblResponsible, DEFAULT_MIN_TIMEOUT) 
	If Fn_UI_JavaList_Operations("RMB Add New Manufacturing Document", "type", obj_ManufacturingPart,"jlst_Part_2",Parameter("str_responsible"), "", "")=False Then
		Reporter.ReportEvent micFail, "Select [ " & glblResponsible & " ] from list box", "Fail to Select [ " & glblResponsible & " ] from list box"
		ExitComponent
	Else
		Reporter.ReportEvent micPass, "Select [ " & glblResponsible & " ] from list box", "Successfully select [ " & Parameter("str_responsible") & " ] from [ " & glblResponsible & " ] from list box"			
	End If
End  If	
'----------------------------------------------------------------------------------------------------------------------
'Set Organizations
'----------------------------------------------------------------------------------------------------------------------
If Parameter("str_organizations") <> "" Then
	Call Fn_UI_Object_Operations("RMB Add New Manufacturing Document","settoproperty", obj_ManufacturingPart.JavaStaticText("jstxt_Part"),"label", glblOrganizations, DEFAULT_MIN_TIMEOUT) 
	If Fn_UI_JavaList_Operations("RMB Add New Manufacturing Document", "type", obj_ManufacturingPart,"jlst_Part_2",Parameter("str_organizations"), "", "")=False Then
		Reporter.ReportEvent micFail, "Select [ " & glblOrganizations & " ] from list box", "Fail to Select [ " & glblOrganizations & " ] from list box"
		ExitComponent
	Else
		Reporter.ReportEvent micPass, "Select [ " & glblOrganizations & " ] from list box", "Successfully select [ " & Parameter("str_organizations") & " ] from [ " & glblOrganizations & " ] from list box"			
	End If
End  If	

'----Set All Location Access---

'----------------------------------------------------------------------------------------------------------------------
'Set Locations
'----------------------------------------------------------------------------------------------------------------------
If Parameter("str_add_locations") <> "" Then
	sMaterial = Split(Parameter("str_add_locations"), "~", -1, 1)
		
	'Click on Empty list button in material panel
	Call Fn_UI_Object_Operations("RMB Add New Manufacturing Document","settoproperty", obj_ManufacturingPart.JavaButton("jbtn_PartTableButton"),"label", gbtnEmptyList, DEFAULT_MIN_TIMEOUT) 
	Call Fn_UI_Object_Operations("RMB Add New Manufacturing Document","settoproperty", obj_ManufacturingPart.JavaButton("jbtn_PartTableButton"),"index", "0", DEFAULT_MIN_TIMEOUT) 
	If Fn_UI_JavaButton_Operations("RMB Add New Manufacturing Document", "Click", obj_ManufacturingPart.JavaButton("jbtn_PartTableButton"), "")=False Then
		Reporter.ReportEvent micFail, "Click on Empty List button in Material panel", "Fail to click Empty List button in Material panel"
		ExitComponent
	End  IF
	
	'Click on Add button in Material panel
	Call Fn_UI_Object_Operations("RMB Add New Manufacturing Document","settoproperty", obj_ManufacturingPart.JavaButton("jbtn_PartTableButton"),"label", gbtnAdd, DEFAULT_MIN_TIMEOUT) 
	Call Fn_UI_Object_Operations("RMB Add New Manufacturing Document","settoproperty", obj_ManufacturingPart.JavaButton("jbtn_PartTableButton"),"index", "0", DEFAULT_MIN_TIMEOUT) 
	If Fn_UI_JavaButton_Operations("RMB Add New Manufacturing Document", "Click", obj_ManufacturingPart.JavaButton("jbtn_PartTableButton"), "")=False Then
		Reporter.ReportEvent micFail, "Click on Add button in Material panel", "Fail to click Add button in Material panel"
		ExitComponent
	End  IF
		
	If Fn_UI_Object_Operations("RMB Add New Manufacturing Document","exist", obj_QueryAndRelate,"", "", DEFAULT_MAX_TIMEOUT) Then
		obj_QueryAndRelate.Activate	
		wait 2			
		Call Fn_UI_Object_Operations("RMB Add New Manufacturing Document","exist", obj_QueryAndRelate.JavaTable("jtbl_SelectItemsToRelate"),"", "", DEFAULT_MAX_TIMEOUT) 
		iRowCount=Fn_UI_JavaTable_Operations("RMB Add New Manufacturing Document","getrowcount",obj_QueryAndRelate,"jtbl_SelectItemsToRelate","","","","","")
		For iCount = 0 To Ubound(sMaterial)	
			For iCounter = 0 To iRowCount-1
				sData=Fn_UI_JavaTable_Operations("RMB Add New Manufacturing Document","getcelldata",obj_QueryAndRelate,"jtbl_SelectItemsToRelate",iCounter,"Name","","","")
				If trim(sData)=Trim(sMaterial(iCount)) Then
					Call Fn_UI_JavaTable_Operations("RMB Add New Manufacturing Document","extendrow",obj_QueryAndRelate,"jtbl_SelectItemsToRelate",iCounter,"Name",sMaterial(iCount),"","")
					Reporter.ReportEvent micPass, "Select the Material [ " & sMaterial(iCount) & " ]", "Successfully Select the Material [ " & sMaterial(iCount) & " ]"	
					Exit For
				End If
			Next
		Next	
	Else
		Reporter.ReportEvent micFail, "Select the Material [ " & sMaterial(iCount) & " ]", "Fail to select Material [ " & sMaterial(iCount) & " ] from Material table"
		ExitComponent
	End  If
	
	'Click on Add button
	If Fn_UI_JavaButton_Operations("RMB Add New Manufacturing Document", "Click", obj_QueryAndRelate.JavaButton("jbtn_Add"), "") Then
		Reporter.ReportEvent micPass, "Click on [ Add ] button on [ Material Query And Relate ] table on [RMB Add New Manufacturing Document ] dialog", "Sucessfully clicked on [ Add ] button on [ Material Query And Relate ] table on [RMB Add New Manufacturing Document ] dialog"	
	End If
End If

'----click on Scroll Button----
'----------------------------------------------------------------------------------------------------------------------
'Set Project
'----------------------------------------------------------------------------------------------------------------------
If Parameter("str_add_projects") <> "" Then
	'Click on Empty list button in Project panel
	Call Fn_UI_Object_Operations("RMB Add New Manufacturing Document","settoproperty", obj_ManufacturingPart.JavaButton("jbtn_Part"),"label", gbtnEmptyList, DEFAULT_MIN_TIMEOUT) 
	If Fn_UI_JavaButton_Operations("RMB Add New Manufacturing Document", "Click", obj_ManufacturingPart.JavaButton("jbtn_Part"), "")=False Then
		Reporter.ReportEvent micFail, "Click on Empty List button in Project panel", "Fail to click Empty List button in Project panel"
		ExitComponent
	End IF
	
	'Click on Add button in Project panel
	Call Fn_UI_Object_Operations("RMB Add New Manufacturing Document","settoproperty", obj_ManufacturingPart.JavaButton("jbtn_Part"),"label", gbtnAdd, DEFAULT_MIN_TIMEOUT) 
	If Fn_UI_JavaButton_Operations("RMB Add New Manufacturing Document", "Click", obj_ManufacturingPart.JavaButton("jbtn_Part"), "")=False Then
		Reporter.ReportEvent micFail, "Click on Empty List button in Project panel", "Fail to click Empty List button in Project panel"
		ExitComponent
	End IF
	
	If Fn_UI_Object_Operations("RMB Add New Manufacturing Document","exist", obj_QueryAndRelate,"", "", DEFAULT_MAX_TIMEOUT) Then
		
		obj_QueryAndRelate.Activate
		wait 2
		Call Fn_UI_Object_Operations("RMB Add New Manufacturing Document","exist", obj_QueryAndRelate.JavaTable("jtbl_SelectItemsToRelate"),"", "", DEFAULT_MAX_TIMEOUT) 
		iRowCount=Fn_UI_JavaTable_Operations("RMB Add New Manufacturing Document","getrowcount",obj_QueryAndRelate,"jtbl_SelectItemsToRelate","","","","","")
		For iCounter = 0 To iRowCount-1
			sData=Fn_UI_JavaTable_Operations("RMB Add New Manufacturing Document","getcelldata",obj_QueryAndRelate,"jtbl_SelectItemsToRelate",iCounter,"Name","","","")
			If trim(sData)=Trim(Parameter("str_Projects")) Then
				Call Fn_UI_JavaTable_Operations("RMB Add New Manufacturing Document","selectrow",obj_QueryAndRelate,"jtbl_SelectItemsToRelate",iCounter,"Name",Parameter("str_Projects"),"","")
				'Click on Add button
				Call Fn_UI_JavaButton_Operations("RMB Add New Manufacturing Document", "Click", obj_QueryAndRelate.JavaButton("jbtn_Add"), "")
				Exit For
			End If
		Next
	Else
		Reporter.ReportEvent micFail, "Select the Project [ " & Parameter("str_Projects") & " ]", "Fail to select Project [ " & Parameter("str_Projects") & " ] from Project table"
		ExitComponent
	End If
	'----------------------------------------------------------------------------------------------------------------------
	'Set Override secure with
	'----------------------------------------------------------------------------------------------------------------------
	Call Fn_UI_Object_Operations("RMB Add New Manufacturing Document","settoproperty", obj_ManufacturingPart.JavaStaticText("jstxt_Part"),"label", gchkbOverrideSecureWith, DEFAULT_MIN_TIMEOUT) 
	If Parameter("bln_Override_Secure_With") = True Then
		Call Fn_UI_JavaCheckBox_Operations("RMB Add New Manufacturing Document", "Set", obj_ManufacturingPart, "jchkb_Part", "ON") 
		Call Fn_UI_JavaCheckBox_Operations("RMB Add New Manufacturing Document", "Set", obj_ManufacturingPart, "jchkb_Part", "ON")
	Else
		Call Fn_UI_JavaCheckBox_Operations("RMB Add New Manufacturing Document", "Set", obj_ManufacturingPart, "jchkb_Part", "ON")
	End If
End If

'******************** Select More tab **************************
If Fn_UI_JavaTab_Operations("RMB Add New Manufacturing Document", "select",obj_ManufacturingPart,"jtab_Part",gtabMore)=False Then
	Reporter.ReportEvent micFail, "Select [ " & gtabMore & " ] tab", "Fail to Select [ " & gtabMore & " ] tab"
	ExitComponent
Else
	Reporter.ReportEvent micPass, "Select [ " & gtabMore & " ] tab", "Successfully select [ " & gtabMore & " ] tab"				
End If
'----------------------------------------------------------------------------------------------------------------------
'Set Orignal Doc
'----------------------------------------------------------------------------------------------------------------------
If Parameter("str_original_doc") <> "" Then
	Call Fn_UI_Object_Operations("RMB Add New Manufacturing Document","settoproperty", obj_CVSPLM.JavaEdit("jedt_Common"),"attached text", glblOrignalDoc, DEFAULT_MIN_TIMEOUT) 
	If Fn_UI_JavaEdit_Operations("RMB Add New Manufacturing Document", "Set",  obj_CVSPLM, "jedt_Common", Parameter("str_original_doc")  )=False Then
		Reporter.ReportEvent micFail, "Set [ " & Parameter("str_original_doc")  & " ] value in edit box", "Fail to set [ " & glblOrignalDoc & " ] value in edit box"
		ExitComponent
	Else
		Reporter.ReportEvent micPass, "Set [ " & Parameter("str_original_doc")  & " ] value in edit box", "Successfully set [ " & Parameter("str_original_doc")  & " ] in [ " & glblOrignalDoc & " ] edit box"		
	End If

End If
'----------------------------------------------------------------------------------------------------------------------
'Set Comment 
'----------------------------------------------------------------------------------------------------------------------
If Parameter("str_comment") <> "" Then
	Call Fn_UI_Object_Operations("RMB Add New Manufacturing Document","settoproperty", obj_CVSPLM.JavaEdit("jedt_Common"),"attached text", glblComment, DEFAULT_MIN_TIMEOUT) 
	If Fn_UI_JavaEdit_Operations("RMB Add New Manufacturing Document", "Set",  obj_CVSPLM, "jedt_Common", Parameter("str_comment")  )=False Then
		Reporter.ReportEvent micFail, "Set [ " & Parameter("str_comment")  & " ] value in edit box", "Fail to set [ " & glblComment & " ] value in edit box"
		ExitComponent
	Else
		Reporter.ReportEvent micPass, "Set [ " & Parameter("str_comment")  & " ] value in edit box", "Successfully set [ " & Parameter("str_comment")  & " ] in [ " & glblComment & " ] edit box"		
	End If
End If
'----------------------------------------------------------------------------------------------------------------------
'Set External Document Number 
'----------------------------------------------------------------------------------------------------------------------
If Parameter("str_external_document_number") <> "" Then
	Call Fn_UI_Object_Operations("RMB Add New Manufacturing Document","settoproperty", obj_CVSPLM.JavaEdit("jedt_Common"),"attached text", glblExternalDocumentNumber, DEFAULT_MIN_TIMEOUT) 
	If Fn_UI_JavaEdit_Operations("RMB Add New Manufacturing Document", "Set",  obj_CVSPLM, "jedt_Common", Parameter("str_external_document_number")  )=False Then
		Reporter.ReportEvent micFail, "Set [ " & Parameter("str_external_document_number")  & " ] value in edit box", "Fail to set [ " & glblExternalDocumentNumber & " ] value in edit box"
		ExitComponent
	Else
		Reporter.ReportEvent micPass, "Set [ " & Parameter("str_external_document_number")  & " ] value in edit box", "Successfully set [ " & Parameter("str_external_document_number")  & " ] in [ " & glblExternalDocumentNumber & " ] edit box"		
	End If
End If
'----------------------------------------------------------------------------------------------------------------------
'Set External Creator  
'----------------------------------------------------------------------------------------------------------------------
If Parameter("str_external_creator") <> "" Then
	Call Fn_UI_Object_Operations("RMB Add New Manufacturing Document","settoproperty", obj_CVSPLM.JavaEdit("jedt_Common"),"attached text", glblExternalCreator, DEFAULT_MIN_TIMEOUT) 
	If Fn_UI_JavaEdit_Operations("RMB Add New Manufacturing Document", "Set",  obj_CVSPLM, "jedt_Common", Parameter("str_external_creator")  )=False Then
		Reporter.ReportEvent micFail, "Set [ " & Parameter("str_external_creator")  & " ] value in edit box", "Fail to set [ " & glblExternalCreator & " ] value in edit box"
		ExitComponent
	Else
		Reporter.ReportEvent micPass, "Set [ " & Parameter("str_external_creator")  & " ] value in edit box", "Successfully set [ " & Parameter("str_external_creator")  & " ] in [ " & glblExternalCreator & " ] edit box"		
	End If
End If
'----------------------------------------------------------------------------------------------------------------------
'Click on OK button
'----------------------------------------------------------------------------------------------------------------------
Call Fn_UI_Object_Operations("RMB Add New Manufacturing Document","settoproperty", obj_ManufacturingPart.JavaButton("jbtn_Part"),"label", gbtnOK, DEFAULT_MIN_TIMEOUT) 
If Fn_UI_JavaButton_Operations("RMB Add New Manufacturing Document", "Click", obj_ManufacturingPart.JavaButton("jbtn_Part"), "")=False Then
	Reporter.ReportEvent micFail, "Click on [ " & gbtnOK & " ] button", "Fail to click on [ " & gbtnOK & " ] button"
	ExitComponent
Else
	Reporter.ReportEvent micPass, "Click on [ " & gbtnOK & " ] button", "Successfully clicked on [ " & gbtnOK & " ] button"	
End If		
'----------------------------------------------------------------------------------------------------------------------
'Check Status Message for Create Data Sheet
'---------------------------------------------------------------------------------------------------------------------
Call Fn_UI_Object_Operations("RMB Add New Manufacturing Document","settoproperty", obj_CVSPLM.JavaStaticText("jstxt_Common"),"label", gAddDataSheetLaunched, DEFAULT_MIN_TIMEOUT) 
If Fn_UI_Object_Operations("RMB Add New Manufacturing Document","exist", obj_CVSPLM.JavaStaticText("jstxt_Common"),"", "", DEFAULT_MAX_TIMEOUT) = False Then
	Reporter.ReportEvent micFail, "Verify [ " & gAddDataSheetLaunched & " ] value in Status message", "Fail to Verify [ " & gAddDataSheetLaunched & " ] value in Status message"
	ExitComponent
Else
	Reporter.ReportEvent micPass, "Verify [ " & gAddDataSheetLaunched & " ] value in Status message", "Successfully Verify [ " & gAddDataSheetLaunched & " ] value in Status message"
End If
'----------------------------------------------------------------------------------------------------------------------
'Click on Continue Without Template button on Create Data Sheet dialog
'---------------------------------------------------------------------------------------------------------------------
If Parameter("bln_ContinueWithoutTemplate")=True Then
	If Fn_UI_JavaButton_Operations("RMB Add New Manufacturing Document", "Click", obj_ManufacturingPart.JavaButton("jbtn_ContinueWithoutTemplate"), "")=False Then
		Reporter.ReportEvent micFail, "Click on [ Continue Without Template ] button", "Fail to click on [ Continue Without Template ] button"
		ExitComponent
	Else
		Reporter.ReportEvent micPass, "Click on [ Continue Without Template ] button", "Successfully clicked on [ Continue Without Template ] button"	
	End If
	'----------------------------------------------------------------------------------------------------------------------
	'Check Status Message for Continue without template
	'---------------------------------------------------------------------------------------------------------------------
	Call Fn_UI_Object_Operations("RMB Add New Manufacturing Document","settoproperty", obj_CVSPLM.JavaStaticText("jstxt_Common"),"label", gAddNewManufacturingDocCancel, DEFAULT_MIN_TIMEOUT) 
	If Waituntilexist(obj_CVSPLM.JavaStaticText("jstxt_Common"),5,5) = False Then
		Reporter.ReportEvent micFail, "Verify [ " & gAddNewManufacturingDocCancel & " ] value in Status message", "Fail to Verify [ " & gAddNewManufacturingDocCancel & " ] value in Status message"
		ExitComponent
	Else
		Reporter.ReportEvent micPass, "Verify [ " & gAddNewManufacturingDocCancel & " ] value in Status message", "Successfully Verify [ " & gAddNewManufacturingDocCancel & " ] value in Status message"
	End If
End If


'----------------------------------------------------------------------------------------------------------------------
'Get Created Manufacturing Document Name
'----------------------------------------------------------------------------------------------------------------------
sManDoc = Fn_UI_JavaTable_Operations("RMB Add New Manufacturing Document","getcelldata",obj_CVSPLM,"jtbl_Results",1,"Name","","","") 
Parameter("str_new_manufacturing_document_revision_out") = sManDoc

'Select Created Manufacturing Document
Call Fn_UI_JavaTable_Operations("RMB Add New Manufacturing Document","selectdata",obj_CVSPLM,"jtbl_Results","","Name",sManDoc,"","")
'----------------------------------------------------------------------------------------------------------------------
'Select Relation tab->Parts Tab
'----------------------------------------------------------------------------------------------------------------------
'Select Relations tab
Call Fn_UI_JavaTab_Operations("RMB Add New Manufacturing Document", "select",obj_CVSPLM,"jtab_FrameContainer",gtabRelations)
'Select Parts tab
If Fn_UI_JavaTab_Operations("RMB Add New Manufacturing Document", "select",obj_CVSPLM,"jtab_Common",gtabParts)=False Then
	Reporter.ReportEvent micFail, "Select [ " & gtabParts & " ] tab", "Fail to Select [ " & gtabParts & " ] tab"
	ExitComponent
Else
	Reporter.ReportEvent micPass, "Select [ " & gtabParts & " ] tab", "Successfully select [ " & gtabParts & " ] tab" 
	'----------------------------------------------------------------------------------------------------------------------
	'Verify the Part Revision under Relation tab->Parts Tab
	'----------------------------------------------------------------------------------------------------------------------
	If Fn_UI_JavaTable_Operations("RMB Add New Manufacturing Document","verifyexist",obj_CVSPLM,"jtbl_Relations","","Name",Parameter("str_Part_Revision"),"","") Then
		Reporter.ReportEvent micPass, "Verify Part Revision Under the [ Relations -> Parts ] Tab", "Successfully  verified Part Revision Under the [ Relations -> Parts ] Tab" 
	Else
		Reporter.ReportEvent micFail, "Verify Part Revision Under the [ Relations -> Parts ] Tab", "Failed verify Part Revision Under the [ Relations -> Parts ] Tab"
		Exitcomponent
	End If
End If
'Select Part revision
Call Fn_UI_JavaTable_Operations("RMB Add New Manufacturing Document","selectdata",obj_CVSPLM,"jtbl_Results","","Name",Parameter("str_Part_Revision"),"","")
'----------------------------------------------------------------------------------------------------------------------
'Select Relation TAB->Documents Tab
'----------------------------------------------------------------------------------------------------------------------
'Select Relations tab
Call Fn_UI_JavaTab_Operations("RMB Add New Manufacturing Document", "select",obj_CVSPLM,"jtab_FrameContainer",gtabRelations)
'Select Documents tab
If Fn_UI_JavaTab_Operations("RMB Add New Manufacturing Document", "select",obj_CVSPLM,"jtab_Common",gtabDocuments)=False Then
	Reporter.ReportEvent micFail, "Select [ " & gtabDocuments & " ] tab", "Fail to Select [ " & gtabDocuments & " ] tab"
	ExitComponent
Else
	Reporter.ReportEvent micPass, "Select [ " & gtabDocuments & " ] tab", "Successfully select [ " & gtabDocuments & " ] tab" 
	'----------------------------------------------------------------------------------------------------------------------
	'Verify the Part Revision under Relation TAB->Documents Tab
	'----------------------------------------------------------------------------------------------------------------------
	If Fn_UI_JavaTable_Operations("RMB Add New Manufacturing Document","verifyexist",obj_CVSPLM,"jtbl_Relations","","Name",sManDoc,"","") Then
		Reporter.ReportEvent micPass, "Verify Manufacturing Document  Under the [ Relations -> Documents ] Tab", "Successfully verified Manufacturing Document  Under the [ Relations -> Documents ] Tab" 
	Else
		Reporter.ReportEvent micFail, "Verify Manufacturing Document  Under the [ Relations -> Documents ] Tab", "Failed to verify Manufacturing Document  Under the [ Relations -> Documents ] Tab"
		Exitcomponent
	End If
End If
'-----------------------------------------------------------------------------------------------------------------------------------------------------------	
If Err.Number<> 0 Then
	Reporter.ReportEvent micFail, "RMB Add New Manufacturing Document", "Fail to add [ Manufacturing Document ]  nder Part Revision due to [ " & Err.Description & " ]"
	ExitComponent
Else
	Reporter.ReportEvent micPass, "RMB Add New Manufacturing Document", "Successfully added [ " &   Parameter("str_new_manufacturing_document_revision_out")  & " ] [ Manufacturing Document ]  under the Part Revision in eCenter"
End If
'-----------------------------------------------------------------------------------------------------------------------------------------------------------	
'Set object nothing
Set obj_CVSPLM=Nothing
Set obj_ManufacturingPart =Nothing
Set obj_QueryAndRelate=Nothing
'-----------------------------------------------------------------------------------------------------------------------------------------------------------	



 
