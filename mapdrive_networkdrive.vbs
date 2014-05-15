' Taha Al-jumaily

On Error Resume Next
Dim Network: Set Network = WScript.CreateObject("WScript.Network")
Dim CheckDrive: Set CheckDrive = Network.EnumNetworkDrives()
Dim DriveExists: DriveExists = False
Dim ShareN: ShareN = "My Drive Share Name"
Dim DriveL: DriveL = "N:"
Dim DriveP: DriveP = "\\Server.domain.path\folder\etc"
Dim i


For i = 0 to CheckDrive.Count - 1
  If CheckDrive.Item(i) = DriveL Then
    DriveExists = True
  End If
Next

If DriveExists = False Then
   Network.MapNetworkDrive DriveL, DriveP, False
   
   If Err Then   
      MsgBox l + ShareN + " Can not be use by: " + Network.UserName
   Else
      set WshShell = WScript.CreateObject("WScript.Shell")
      strDesktop = WshShell.SpecialFolders("Desktop")
      set oShellLink = WshShell.CreateShortcut(strDesktop & "\" + ShareN + ".lnk")
      oShellLink.TargetPath = WScript.ScriptFullName
      oShellLink.WindowStyle = 1
      oShellLink.TargetPath = DriveL
      oShellLink.WorkingDirectory = DriveL
      oShellLink.Save 
   End If

'Else
'  MsgBox l + " Drive already mapped! " + Network.UserName
End If
