Option Explicit

Dim fileT, resource

Set resource = GetResource(Trim("C:\res\ActiveWorkspace_OR.xml"))

Set fileT=Eval(GetResource("C:\res\ActiveWorkspace_OR.xml").GetValue("Browser"))


Dim testPath
testPath ="Test directory: " & Environment("TestDir")

Reporter.ReportEvent micPass, testPath, "bbbbbbbbbb"
