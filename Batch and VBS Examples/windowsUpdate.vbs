' Written in 2007 by Harry Johnston, University of Waikato, New Zealand.
' This code has been placed in the public domain.  It may be freely
' used, modified, and distributed.  However it is provided with no
' warranty, either express or implied.
'
' Exit Codes:
'   0 = scripting failure
'   1 = error obtaining or installing updates
'   2 = installation successful, no further updates to install
'   3 = reboot needed; rerun script after reboot
'
' Note that exit code 0 has to indicate failure because that is what
' is returned if a scripting error is raised.
'

Set updateSession = CreateObject("Microsoft.Update.Session")

Set updateSearcher = updateSession.CreateUpdateSearcher()
Set updateDownloader = updateSession.CreateUpdateDownloader()
Set updateInstaller = updateSession.CreateUpdateInstaller()

Do

  WScript.Echo
  WScript.Echo "Searching for approved updates ..."
  WScript.Echo

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

  Next

  WScript.Echo

  updateDownloader.Updates = updateList
  updateDownloader.Priority = 3

  Set downloadResult = updateDownloader.Download()

  If downloadResult.ResultCode <> 2 Then

    WScript.Echo "Download failed with result code", downloadResult.ResultCode
    WScript.Echo

    WScript.Quit 1

  End If

  WScript.Echo "Download complete.  Installing updates ..."
  WScript.Echo

  updateInstaller.Updates = updateList

  Set installationResult = updateInstaller.Install()

  If installationResult.ResultCode <> 2 Then

    WScript.Echo "Installation failed with result code", installationResult.ResultCode

    For I = 0 to updateList.Count - 1

      Set updateInstallationResult = installationResult.GetUpdateResult(I)
      WScript.Echo "Result for " & updateList.Item(I).Title & " is " & installationResult.GetUpdateResult(I).ResultCode

    Next

    WScript.Quit 1

  End If

  If installationResult.RebootRequired Then

    WScript.Echo "The system must be rebooted to complete installation."

    WScript.Quit 3

  End If

  WScript.Echo "Installation complete."

Loop 