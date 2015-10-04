machine = InputBox("Enter Name of the Remote Computer:","Prompt!")
Wscript.StdOut.Write machine
cimv2_string = "WINMGMTS:" & "\\" & machine & "\root\cimv2"
query = "select * from Win32_TerminalServiceSetting"
value_to_set = 1 ' 0=off, 1=on
set cimv2 = GetObject( cimv2_string )
set items = cimv2.ExecQuery( query )
for each item in items
item.SetAllowTSConnections( value_to_set )
next
wscript.Echo "Terminal Services has been enabled."