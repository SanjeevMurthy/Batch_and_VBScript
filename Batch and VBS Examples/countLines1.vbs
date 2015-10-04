Dim objFSO, txsInput, strTemp, arrLines
Const ForReading = 1
Set objFSO = CreateObject("Scripting.FileSystemObject")

strTextFile = "C:\Users\sanju\Desktop\BatchScript\output.txt"
txsInput = objFSO.OpenTextFile(strTextFile, ForReading).ReadAll


Do While txsInput.AtEndOfStream <> True
    txsInput.SkipLine
Loop

wscript.echo txsInput.Line-1


Set objFSO = Nothing