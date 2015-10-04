Option Explicit
Dim objWMIService, objItem, colItems, counter
Dim strComputer, strList, objOS, stringArray

On Error Resume Next
strComputer = "."
strOS = "Windows7"
str7 = "7"


' WMI Connection to the object in the CIM namespace
Set objWMIService = GetObject("winmgmts:\\" _
& strComputer & "\root\cimv2")

' WMI Query to the Win32_OperatingSystem
Set colItems = objWMIService.ExecQuery _
("Select * from Win32_OperatingSystem")

' For Each... In Loop (Next at the very end)
For Each objItem in colItems
objOS=objItem.Caption & VbCr
Next

stringArray = Split(objOS)
for counter=0 to UBound(stringArray)
if (stringArray(counter)=strOS) then
Wscript.Echo "Windows XP"
Exit for
End if
if (stringArray(counter)=str7) then
Wscript.Echo "Windows 7"
Exit for
End if
next
WSCript.Quit