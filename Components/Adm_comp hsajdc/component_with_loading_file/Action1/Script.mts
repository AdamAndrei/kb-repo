Option Explicit

Dim fileT, resource

'Set resource = GetResource(Trim("C:\res\ActiveWorkspace_OR.xml"))

'Set fileT=Eval(GetResource("C:\res\ActiveWorkspace_OR.xml").GetValue("Browser"))
Dim savePath
savePath = "Start: "

Dim testPath
testPath =Environment("TestDir")

Dim parts, i, partialPath
Dim fso, folder, subfolder

Set fso = CreateObject("Scripting.FileSystemObject")

parts = Split(testPath, "\")

partialPath = parts(0)


For i = 1 To UBound(parts)
	partialPath = partialPath & "\" & parts(i)
	Print partialPath
	Set folder = fso.GetFolder(partialPath)
	For Each subfolder In folder.SubFolders
		Print subfolder.Path
	Next
Next


Reporter.ReportEvent micPass, testPath, savePath
