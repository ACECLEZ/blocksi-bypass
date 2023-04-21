'==========================================================================================
'list this file as protected file and conduct checks for unauthorised modifications
'==========================================================================================
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
'========================================
' Check if file at strFilePath1 exists
'========================================
If objFSO.FileExists(strFilePath1) Then
    WScript.Echo "File at strFilePath1 exists: " & strFilePath1
Else
    WScript.Echo "File at strFilePath1 does not exist: " & strFilePath1
End If
'========================================
' Check if file at strFilePath2 exists
'========================================
If objFSO.FileExists(strFilePath2) Then
    WScript.Echo "File at strFilePath2 exists: " & strFilePath2
Else
    WScript.Echo "File at strFilePath2 does not exist: " & strFilePath2
End If

Set objFSO = Nothing
