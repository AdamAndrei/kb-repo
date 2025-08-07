'******************************************************************************************************************************
' Name Of Business Component		:	Find Recipients of Workflow Task
'
' Purpose							:	Perform the Find Recipients of Workflow Task from Workflow History
'
' Input	Parameter					:	
'
' Output							:	True / False
'
' Remarks							:
'
' Author							:	Mohini Deshmukh 			  16 Dec 2024

'******************************************************************************************************************************
Option Explicit
'--------------------------------------------------------------------------------------------------------------------------------
'Variable Declaration
'--------------------------------------------------------------------------------------------------------------------------------
Dim obj_AWCTeamcenterHome
Dim sUserList,sTempRecipientUser,sTempGroupName
'--------------------------------------------------------------------------------------------------------------------------------
'Get AWC PLM window object from xml
'--------------------------------------------------------------------------------------------------------------------------------
Set obj_AWCTeamcenterHome=Eval(GetResource("ActiveWorkspace2406_OR.xml").GetValue("wpage_AWCTeamcenterHome"))
'--------------------------------------------------------------------------------------------------------------------------------
'Find recipients of task from workflow history tab 
'--------------------------------------------------------------------------------------------------------------------------------
Browser("Browser").Refresh
wait 2
sUserList= Fn_AWC_WorkFlow_History_Tab_Operations(obj_AWCTeamcenterHome,"findrecipient",gWorkflowHistory,"",Parameter("str_TaskName") )
wait 1

Parameter("str_AlternateFlag_out")=Split(sUserList,"^")(1)
sTempGroupName=Split(sUserList,"^")(2)
sUserList=Split(sUserList,"^")(0)
Parameter("str_RecipientList_out")=sUserList
Parameter("str_Recipient_out")=Split(sUserList,"~")(0)

If Parameter("str_AlternateFlag_out")="True" Then
	IF instr(lcase(Parameter("str_LoggedInUser")), lcase(Parameter("str_Recipient_out")))>0 and instr(lcase(Parameter("str_Recipient_out")), lcase(sTempGroupName))>0 Then
		Parameter("str_AlternateFlag_out")="False"
	End  If
End If


' SSO login is available so commented code from line no. 330line no. 37
''If reciept list contains the only one user and which is #T_ user then don't do alternate work 
'If instr(sUserList,"~")=0 and instr(sUserList,"#T")>0 Then
'	Parameter("str_AlternateFlag_out")="False"
'End If
'-----------------------------------------------------------------------------------------------------------------------------------------------------------	
If Err.Number<> 0 Then
	Reporter.ReportEvent micFail, "Find Recipients of Workflow Task", "Fail to perform [ Find Recipients of Workflow Task ]  Operation due to [ " & Err.Description & " ]"
	ExitComponent
Else
	Reporter.ReportEvent micPass, "Find Recipients of Workflow Task", "Successfully performed [ Find Recipients of Workflow Task ] operation "
End If
'-----------------------------------------------------------------------------------------------------------------------------------------------------------	
'Set object nothing
Set obj_AWCTeamcenterHome=Nothing
'-----------------------------------------------------------------------------------------------------------------------------------------------------------	










