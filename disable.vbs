Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")

strStartupFolder = objShell.SpecialFolders("Startup")
strFilePath = strStartupFolder & "\injector.vbs"

If objFSO.FileExists(strFilePath) Then
   objFSO.DeleteFile strFilePath
   WScript.Echo "File deleted successfully."
Else
   WScript.Echo "File not found."
End If
