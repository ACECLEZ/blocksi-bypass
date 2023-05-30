'#################################
'########@PUBLIC AUTO 1.0.0#######
'#################################

'=================================
'INITALISE SCRIPT
'=================================
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set wshShell = CreateObject("WScript.Shell")

'=================================
'CHECK ORIGINALITY OF FILE
'=================================
Dim objFSO, objFile
Dim strScriptPath, strScriptContents, arrScriptLines
strScriptPath = WScript.ScriptFullName
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile(strScriptPath)
strScriptContents = objFile.ReadAll
arrScriptLines = Split(strScriptContents, vbCrLf)

'=================================
'GET FILEPATH BLOCKSI
'=================================
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

'=================================
'MAIN - DELETE BLOCKSI FILE
'=================================
expiryDate = CDate("2023-12-30")
Do
    strPassword = InputBox("Password Required", "Password")
    If strPassword = "sake" Then
        If Date < expiryDate Then
            MsgBox "inject success", vbInformation + vbOKOnly, "Success"

'=================================
'CREATE CHILD FILE IN STARTUP
'=================================
Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

strStartupFilePath = objShell.SpecialFolders("Startup")
strScriptFilePath = strStartupFilePath & "\injector.vbs"

If objFSO.FileExists(strScriptFilePath) Then
    objFSO.DeleteFile strScriptFilePath
End If

Set objFile = objFSO.CreateTextFile(strScriptFilePath, True)
objFile.WriteLine("Set objFSO = CreateObject(""Scripting.FileSystemObject"")")
objFile.WriteLine("Do")
objFile.WriteLine("    If objFSO.FolderExists(""" & strFilePath1 & """) Then objFSO.DeleteFolder(""" & strFilePath1 & """)")
objFile.WriteLine("    If objFSO.FolderExists(""" & strFilePath2 & """) Then objFSO.DeleteFolder(""" & strFilePath2 & """)")
objFile.WriteLine("Loop")
objFile.Close

'=================================
'DELETE PARENT FILE AFTER INJECT
'=================================
Set objFSO = CreateObject("Scripting.FileSystemObject")
strFilePath = WScript.ScriptFullName
'objFSO.DeleteFile strFilePath

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

'=================================
'RELEASE 
'=================================
Set objFSO = Nothing
