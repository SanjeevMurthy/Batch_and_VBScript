'#==============================================================================
'#==============================================================================
'#  SCRIPT.........:  remoteCommand.vbs
'#  AUTHOR.........:  Joe Glessner
'#  EMAIL..........:  
'#  VERSION........:  1.0
'#  DATE...........:  2004JUN18
'#  COPYRIGHT......:  2010, Joe Glessner
'#  LICENSE........:  Freeware
'#  EXAMPLE........:  cscript.exe remoteCommand.vbs COMP1 "gpupdate.exe /force"
'#                        This example will cause the remote computer "COMP1"
'#                        to refresh it's Group Policy Settings.
'#  REQUIREMENTS...:  
'#
'#  DESCRIPTION....:  Executes a command on a remote computer using a WMI 
'#                    connection.
'#
'#  NOTES..........:  The command to be executed must exist on the remote 
'#                    computer, in %PATH%.
'#                    Tested on Windows 2000 Professional, Windows XP, 
'#                    Windows Vista, Windows 7.
'# 
'#  CUSTOMIZE......:  
'#==============================================================================
'#  REVISED BY.....:  
'#  EMAIL..........:  
'#  REVISION DATE..:  
'#  REVISION NOTES.:
'#
'#==============================================================================
'#==============================================================================
'**Start Encode**

'#==============================================================================
'#  START OF SCRIPT
'#==============================================================================
Option Explicit
'On Error Resume Next

    '#--------------------------------------------------------------------------
    '#  Declare Constants
    '#--------------------------------------------------------------------------
    
    '#--------------------------------------------------------------------------
    '#  Declare Variables
    '#--------------------------------------------------------------------------
    Dim strComputer:strComputer = WScript.Arguments.Item(0)
     Dim strCmd:strCmd = WScript.Arguments.Item(1)
    Dim objWMIService, errReturn, intProcessID
    '#--------------------------------------------------------------------------
    '#  Ensure this is being executed with cscript
    '#--------------------------------------------------------------------------
    forceCScript()
    
    '#--------------------------------------------------------------------------
    '#  Initiate WMI connection to remote computer
    '#--------------------------------------------------------------------------
    Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & _
    "\root\cimv2:Win32_Process")

    '#--------------------------------------------------------------------------
    '#  Execute remote command
    '#--------------------------------------------------------------------------
    errReturn = objWMIService.Create(strCmd,null,null,intProcessID)
    if errReturn = 0 Then
        Wscript.Echo strCmd & " was started with a process ID of " _
        & intProcessID & "."
    Else
        Wscript.Echo strCmd & " could not be started. Error: " & errReturn & "."
    End If 

'#==============================================================================
'#  SUBROUTINES/FUNCTIONS/CLASSES
'#==============================================================================
    '#--------------------------------------------------------------------------
    '#  SUBROUTINE.......:  forceCScript
    '#  PURPOSE..........:  If a script is launched with Wscript, this sub 
    '#                      called at the top of the script will relaunch the 
    '#                      script using CScript as the WSH host.
    '#  ARGUMENTS........:  None
    '#  EXAMPLE..........:  forceCScript
    '#                      wscript.echo "Working!"
    '#  REQUIREMENTS.....:  
    '#  NOTES............:  The above example will cause a shell window to
    '#                      display briefly, showing the text "Working!".
    '#--------------------------------------------------------------------------
    Sub forceCScript()
        Dim oShell: Set oShell = CreateObject("Wscript.Shell")
        If Not WScript.FullName = WScript.Path & "\cscript.exe" Then
            oShell.Run WScript.Path & "\cscript.exe //NOLOGO " & Chr(34) & _
            WScript.scriptFullName & Chr(34),1,False
            WScript.Sleep 2000
            WScript.Quit 0
        End If
    End Sub

'#==============================================================================
'#  END OF FILE
'#==============================================================================