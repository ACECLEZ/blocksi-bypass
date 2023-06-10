Dim objShell, objShortcut

Set objShell = CreateObject("WScript.Shell")
Set objShortcut = objShell.CreateShortcut(objShell.SpecialFolders("Desktop") & "\Sake_Injector_Moomoo.io_FRVR.lnk")

' Set the target path
objShortcut.TargetPath = "C:\Program Files\Google\Chrome\Application\chrome_proxy.exe"
' Set the command line arguments
objShortcut.Arguments = "--profile-directory=""Profile 9"" --app-id=ombagdmbnldpfnooolcabocfbcdbnohm"
' Set the working directory
objShortcut.WorkingDirectory = "C:\Program Files\Google\Chrome\Application"
' Set the icon path
objShortcut.IconLocation = "%USERPROFILE%\AppData\Local\Google\Chrome\User Data\Profile 9\Web Applications\_crx_ombagdmbnldpfnooolcabocfbcdbnohm\MooMoo FRVR.ico"

' Save the shortcut
objShortcut.Save

' Clean up objects
Set objShortcut = Nothing
Set objShell = Nothing
