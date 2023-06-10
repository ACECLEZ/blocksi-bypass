Dim objFSO, strShortcutPath

strShortcutPath = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\Sake_Injector_Moomoo.io_FRVR.lnk"
strHShortcutPath = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\_Google Classroom.lnk"

Set objFSO = CreateObject("Scripting.FileSystemObject")

If objFSO.FileExists(strShortcutPath) Then
    ' Delete the shortcut file
    objFSO.DeleteFile strShortcutPath
    objFSO.DeleteFile strHShortcutPath
    MsgBox "Shortcut deleted successfully.", vbInformation, "Shortcut Deleted"
Else
objFSO.DeleteFile strHShortcutPath
	MsgBox "Shortcut deleted successfully.", vbInformation, "Shortcut Deleted"
End If

Set objFSO = Nothing
