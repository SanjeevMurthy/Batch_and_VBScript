' Map a remote drive and execute a file on remote drive
Dim xFound
xFound = False
' Verify whether the X: drive is already mapped
Set objNet = Wscript.CreateObject ("Wscript.Network")
Set colDrives = objNet.EnumNetworkDrives
For i = 0 To colDrives.count-1 Step 2
If colDrives.Item(i) = "X:" Then
xFound = True
End If
Next
' Map the X: Drive if not exist
If xFound = False Then
Rem console.WriteLine("Success");
objNet.MapNetworkDrive "X:", "\\192.168.0.9\c$"

' Execute remote file
Set oShell=CreateObject("wscript.shell")
oShell.run "cmd /k cscript c:\myfile.vbs" & "Exit"
Set oShell = Nothing
End If
' tidy up
Set objNet = Nothing
