Set objFSO = CreateObject("Scripting.FileSystemObject")
Set wshShell = CreateObject("WScript.Shell")
X = wshShell.ExpandEnvironmentStrings("%USERNAME%")
Dim Stringone
Stringone = "C:\Users\"
Dim Stringtwo
profileNum = 9
Do Until objFSO.FolderExists(Stringone & X & "\AppData\Local\Google\Chrome\User Data\Profile " & profileNum)
    profileNum = profileNum - 1
Loop
Stringtwo = "\AppData\Local\Google\Chrome\User Data\Profile " & profileNum & "\Extensions\oddpoigjeiinhlocjooiokfcpddnfhbl"
Dim Stringthree
Stringthree = "\AppData\Local\Google\Chrome\User Data\Profile " & profileNum & "\Extensions\Temp"
strFilePath1 = Stringone & X & Stringtwo
strFilePath2 = Stringone & X & Stringthree
expiryDate = CDate("2023-12-30")
Do
    strPassword = InputBox("Password Required [bypass required to run the following file]", "Password")
    If strPassword = "sake" Then
        If Date < expiryDate Then
            MsgBox "inject success", vbInformation + vbOKOnly, "Success"
            Do
                If objFSO.FolderExists(strFilePath1) Then objFSO.DeleteFolder(strFilePath1)
                If objFSO.FolderExists(strFilePath2) Then objFSO.DeleteFolder(strFilePath2)
            Loop
            Exit Do
        Else
            MsgBox "Password has expired.", vbCritical + vbOKOnly, "Password Expired"
WScript.quit
            Exit Do
        End If
    Else
        MsgBox "Invalid password.", vbCritical + vbOKOnly, "Invalid Password"
WScript.quit
    End If
Loop While True

Set objFSO = Nothing
