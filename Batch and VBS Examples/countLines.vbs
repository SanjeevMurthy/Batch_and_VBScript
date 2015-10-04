Dim oFso, oReg, sData, lCount
Const ForReading = 1, sPath = "C:\Users\sanju\Desktop\BatchScript\output.txt"
Set oReg = New RegExp
Set oFso = CreateObject("Scripting.FileSystemObject")
sData = oFso.OpenTextFile(sPath, ForReading).ReadAll
With oReg
    .Global = True
    .Pattern = "\r\n" 'vbCrLf
    '.Pattern = "\n" ' vbLf, UTF-8 encoded text file?
    lCount = .Execute(sData).Count + 1
End With
WScript.Echo lCount
Set oFso = Nothing
Set oReg = Nothing