Dim objShell, objShortcut, strOldShortcutPath, strNewShortcutPath, objFSO

strOldShortcutPath = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\Sake_Injector_Moomoo.io_FRVR.lnk"
strNewShortcutPath = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\_Google Classroom.lnk"

Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

If objFSO.FileExists(strOldShortcutPath) Then
    ' Rename the shortcut file
    objFSO.MoveFile strOldShortcutPath, strNewShortcutPath
    
    ' Create a new shortcut object with the updated path
    Set objShortcut = objShell.CreateShortcut(strNewShortcutPath)
    
    ' Set the icon path
    objShortcut.IconLocation = "%USERPROFILE%\AppData\Local\Google\Chrome\User Data\Profile 9\Web Applications\_crx_kjcjfjccmpngedeildfijeanhihmolck\Google Classroom.ico"
    
    ' Save the updated shortcut
    objShortcut.Save
    
    ' Clean up object
    Set objShortcut = Nothing
    
    MsgBox "Shortcut icon and name changed successfully.", vbInformation, "Shortcut Updated"
Else
    MsgBox "Shortcut does not exist.", vbExclamation, "Shortcut Not Found"
End If

Set objShell = Nothing
Set objFSO = Nothing
