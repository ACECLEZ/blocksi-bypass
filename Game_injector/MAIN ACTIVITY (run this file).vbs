Option Explicit

Dim continueProgram
continueProgram = True

Dim ie

Dim password
password = "sake"
Dim enteredPassword

Do While continueProgram
If enteredPassword = "" Then
enteredPassword = InputBox("Enter Password:" & vbCrLf & vbCrLf & _
                     "https://github.com/ACECLEZ/blocksi-bypass/", "Sake Injector Login - {Non Root 1.2.0} - {Game Edition}")


End If
If enteredPassword = password Then
    Dim input
    input = InputBox("Please enter a command:" & vbCrLf & vbCrLf & _
                     "1 - Current Session Bypass [ACTIVATE THIS]" & vbCrLf & _
                     "2 - Install Moomoo.io FRVR" & vbCrLf & _
                     "3 - Hide Game [Disguise as Google Classroom]" & vbCrLf & _
                     "4 - Delete Game" & vbCrLf & _
                     "5 - Exit the program", "Sake Injector - Non Root - Non Jailbreak - V1.2.0")

    Select Case input
        Case "1"
            Dim bypassManualPath
            bypassManualPath = ".\Bypass-Manual.vbs" ' Replace with the actual path to Bypass-Manual.vbs
            CreateObject("WScript.Shell").Run bypassManualPath, 0, False

        Case "2"
            Dim InstallGame
            InstallGame = ".\moomoo.io.vbs" 
            CreateObject("WScript.Shell").Run InstallGame, 0, False

        Case "3"
            Dim HideGame
            HideGame = ".\hide_games.vbs" 
            CreateObject("WScript.Shell").Run HideGame, 0, False

        Case "4"
            Dim DeleteGame
            DeleteGame = ".\delete_games.vbs"
            CreateObject("WScript.Shell").Run DeleteGame, 0, False

        Case "5"
            continueProgram = False

        Case Else
            MsgBox "Invalid command. Please try again."
    End Select
Else
    MsgBox "Incorrect password. Exiting the program."
    Exit Do
End If

If continueProgram Then
    Dim continueResponse
    continueResponse = MsgBox("Do you want to continue using the program?", vbQuestion + vbYesNo)
    If continueResponse = vbNo Then
        continueProgram = False
    End If
End If
Loop
