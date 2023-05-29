Option Explicit

Dim continueProgram
continueProgram = True

Dim password
Dim enteredPassword

Do While continueProgram
    If enteredPassword = "" Then
        enteredPassword = InputBox("Enter password:")
    End If
    
    If enteredPassword = "" Then
        MsgBox "Password cannot be empty. Exiting the program."
        Exit Do
    End If
    
    ' Retrieve password from online JSON file
    Dim passwordURL
    passwordURL = "https://pwcreatures.vercel.app/bin/injector_password.json" ' Replace with the URL of your JSON file
    
    Dim httpRequest
    Set httpRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
    httpRequest.Open "GET", passwordURL, False
    httpRequest.Send
    
    If httpRequest.Status = 200 Then ' Successful request
        Dim json
        json = httpRequest.ResponseText
        
        ' Parse the JSON response
        Dim objJSON
        Set objJSON = CreateObject("Scripting.Dictionary")
        objJSON("password") = ""
        
        On Error Resume Next
        Set objJSON = ScriptEngine.Eval("(" + json + ")")
        On Error Goto 0
        
        password = objJSON("password")
        
        If password = "" Then
            MsgBox "Failed to retrieve password. Exiting the program."
            Exit Do
        End If
    Else
        MsgBox "Internet error - Failed to retrieve password. Exiting the program."
        Exit Do
    End If
    
    If enteredPassword = password Then
        ' Rest of your code here...
        ' Handle the different command options
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
