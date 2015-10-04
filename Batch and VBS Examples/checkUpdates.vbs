Set updateSession = CreateObject("Microsoft.Update.Session")

Set updateSearcher = updateSession.CreateUpdateSearcher()
Set updateDownloader = updateSession.CreateUpdateDownloader()
Set updateInstaller = updateSession.CreateUpdateInstaller()


Do

   REM WScript.Echo
  WScript.Echo "Searching for approved updates ..."
  REM WScript.Echo

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

  For I = 0 to updateSearch.Updates.Count - 1

    Set update = updateList.Item(I)

    WScript.Echo "Update found:", update.Title

  NEXT

LOOP
f.Close