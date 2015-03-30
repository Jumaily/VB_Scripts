Set objNet = CreateObject("WScript.NetWork") 
Set objFSO = CreateObject("Scripting.FileSystemObject")
 
dim Apath
dim Bpath 
dim Cpath

Apath = "C:\Users\" & objNet.UserName & "\AppData\Roaming\Coda\"
Cpath = Apath & "Data.sdf"

' path to the database file to be copied if doesn't exist
Bpath = "C:\Coda\Data.sdf"
' dynamic path
'Bpath = objFSO.GetAbsolutePathName(folderName) & "\Data.sdf"


If objFSO.FileExists(Cpath) Then
   Wscript.Quit
Else
   If NOT objFSO.FolderExists(Apath) Then
      objFSO.CreateFolder Apath
   End If
   objFSO.CopyFile Bpath, Apath
End If

