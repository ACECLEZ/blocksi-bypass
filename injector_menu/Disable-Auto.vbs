'#################################
'########@PUBLIC AUTO 1.0.0#######
'#################################

'=================================
'INITALISE SCRIPT
'=================================
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")

'=================================
'GET FILEPATH
'=================================
strStartupFolder = objShell.SpecialFolders("Startup")
strFilePath = strStartupFolder & "\injector.vbs"

'=================================
'DELETE FILE IF FOUND
'=================================
If objFSO.FileExists(strFilePath) Then
   objFSO.DeleteFile strFilePath
   WScript.Echo "File deleted successfully."
Else
   WScript.Echo "File not found."
End If

'=================================
'RELEASE 
'=================================
Set objFSO = Nothing