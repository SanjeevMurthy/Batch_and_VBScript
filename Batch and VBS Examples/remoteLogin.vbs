Const intMin = 3600
strComputer = "sanju-hp" 
strDomain = "WORKGROUP"  
Wscript.StdOut.Write "Please enter your user name:"
strUser = Wscript.StdIn.ReadLine 
strPassword = "Sanju@10"
Rem Set objPassword = CreateObject("ScriptPW.Password")
Rem Wscript.StdOut.Write "Please enter your password:"
Rem strPassword = objPassword.GetPassword()
Wscript.Echo

Set objSWbemLocator = CreateObject("WbemScripting.SWbemLocator") 
Rem Set objWMIService = objSWbemLocator.ConnectServer(strComputer, _ 
Rem    "root\CIMV2", _ 
Rem    strUser, _ 
Rem    strPassword, _ 
Rem    "MS_409", _ 
Rem    "NTLMDomain:" + strDomain) 

Set objWMIService = objSWbemLocator.ConnectServer(strComputer, "root\cimv2",strUser,strPassword)

Set colItems = objWMIService.ExecQuery( _
    "SELECT * FROM Win32_IP4RouteTable",,48) 
For Each objItem in colItems
    WScript.Echo "Age in Minutes: " _
        & int(objItem.Age/intMin) & VBNewLine _
    & "Description: " & objItem.Description & VBNewLine _
    & "Destination: " & objItem.Destination & VBNewLine _
    & "InterfaceIndex: " & objItem.InterfaceIndex & VBNewLine _
    & "Mask: " & objItem.Mask & VBNewLine _
    & "Metric1: " & objItem.Metric1 & VBNewLine _
    & "Metric2: " & objItem.Metric2 & VBNewLine _
    & "Metric3: " & objItem.Metric3 & VBNewLine _
    & "Metric4: " & objItem.Metric4 & VBNewLine _
    & "Metric5: " & objItem.Metric5 & VBNewLine _
    & "Name: " & objItem.Name & VBNewLine _
    & "NextHop: " & objItem.NextHop & VBNewLine _
    & "Protocol: " & objItem.Protocol & VBNewLine _
    & "Type: " & objItem.Type
    WScript.Echo
Next
