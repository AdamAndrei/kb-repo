Dim testPath, resourcesParentFolder, filePath
testPath =Environment("TestDir")

resourcesParentFolder = FindParentFolderWithResources(testPath)

If resourcesParentFolder = "" Then
	Reporter.ReportEvent micFail, "Search", "Parent folder not found."
Else
	Reporter.ReportEvent micPass, "Search", "Found parent folder: " & resourcesParentFolder
End If

filePath = FindResourceFullPath(resourcesParentFolder & "\TestResources", "search.txt")

If filePath = "" Then
	Reporter.ReportEvent micFail, "Search file", "File not found."
Else
	Reporter.ReportEvent micPass, "Search file", "Found file: " & filePath
End If

Dim resource
Set resource = SearchAndLoadResourceByName("ActiveWorkspace_OR.xml")

If Not resource Is Nothing Then
	Reporter.ReportEvent micPass, "Nice", "Found: " & resource.Count() & " elements."
Else
	Reporter.ReportEvent micPass, "Bad", "Found: nothing"
End If
