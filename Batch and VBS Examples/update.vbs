Set updateSession = CreateObject("Microsoft.Update.Session")

Set updateSearcher = updateSession.CreateUpdateSearcher()
Set updateDownloader = updateSession.CreateUpdateDownloader()
Set updateInstaller = updateSession.CreateUpdateInstaller()

Set fso = WScript.CreateObject("Scripting.Filesystemobject")
Set f = fso.OpenTextFile("C:\Users\sanju\Desktop\BatchScript\output.txt", 2)


WScript.Echo "Searching for approved updates ..."


  Set updateSearch = updateSearcher.Search("IsInstalled=0")
  If updateSearch.ResultCode <> 2 Then

    WScript.Echo "Search failed with result code", updateSearch.ResultCode
    WScript.Quit 1

  End If

  If updateSearch.Updates.Count = 0 Then

    WScript.Echo "There are no updates to install."
    WScript.Quit 2

  End If

  Set updateList = updateSearch.Updates
  REM Set updatesAvailable = updateSearch.Updates.Count
  WScript.Echo updateSearch.Updates.Count , "Updates available"

  For I = 0 to updateSearch.Updates.Count - 1

    Set update = updateList.Item(I)

    f.WriteLine update.Title

  NEXT

REM LOOP
f.Close