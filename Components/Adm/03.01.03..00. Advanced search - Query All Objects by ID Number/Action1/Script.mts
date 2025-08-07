'******************************************************************************************************************************
' Name Of Business Component		:	 Advanced search - Query All Objects by ID Number
'
' Purpose							:	 Advanced search - Query All Objects by ID Number
'
' Input	Parameter					:	
'
' Output							:	True / False
'
' Remarks							:
'
' Author							:	Mohini Deshmukh			 15 July 2024

'******************************************************************************************************************************
Set objDic=CreateObject("Scripting.dictionary")
'--------------------------------------------------------------------------------------------------------------------------------
'Set ID field
'--------------------------------------------------------------------------------------------------------------------------------
If Parameter("str_ItemId")<>""  Then
	If  instr(Parameter("str_ItemId"),"~")>0  Then
		 Parameter("str_ItemId")=Replace( Parameter("str_ItemId"),"~",";")
	End If
	objDic.Add gItemID,Parameter("str_ItemId")
	Parameter("str_itemid_out")=Parameter("str_ItemId")
End If
'-------------------------------------------------------------------------------------------------------------------------------
'Set Current Version field
'--------------------------------------------------------------------------------------------------------------------------------
If Parameter("str_current_version") <> "" Then
	objDic.Add gCurrentVer,Parameter("str_current_version")
End If
'-------------------------------------------------------------------------------------------------------------------------------
'Store search result
'--------------------------------------------------------------------------------------------------------------------------------
Parameter("str_searchresult_out")= Fn_AWC_Advanced_Search_Operations(Parameter("str_expected_search_result"),Parameter("str_SearchType"),objDic)
objDic.RemoveAll
'-------------------------------------------------------------------------------------------------------------------------------
