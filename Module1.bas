Attribute VB_Name = "HackModul"
Public HPath As String
Public HPath2 As String
Public PassDriveA As Boolean
Public PassDriveD As Boolean
Public HDat As String
Public HSend As Boolean
Public Speed As Integer

Sub Pause(interval)
'you probably dont want to use this,
'its used in other procedures of this
'module
Dim Current
Current = Timer
Do While Timer - Current < Val(interval)
DoEvents
Loop
End Sub
Function small()
Pause (0.5)
Form1.Height = 8200
Pause (0.5)
Form1.Height = 8300
Pause (0.5)
Form1.Height = 8400
Pause (0.5)
Form1.Height = 8500
Pause (0.5)
Form1.Height = 8600
Pause (0.5)
Form1.Height = 8700
End Function
Function big()
Pause (0.5)
Form1.Height = 8700
Pause (0.5)
Form1.Height = 8600
Pause (0.5)
Form1.Height = 8500
Pause (0.5)
Form1.Height = 8400
Pause (0.5)
Form1.Height = 8300
Pause (0.5)
Form1.Height = 8200
Pause (0.5)
Form1.Height = 8100
End Function
Function Scan()

Form1.Text1.Text = Form1.Text1.Text & "Booting"
Pause (1)
Form1.Text1.Text = Form1.Text1.Text & "."
Pause (1)
Form1.Text1.Text = Form1.Text1.Text & "."
Pause (1)
Form1.Text1.Text = Form1.Text1.Text & "." & vbCrLf
Pause (1)
Form1.Text1.Text = Form1.Text1.Text & "Searching for OS System" & vbCrLf
Pause (1)
Form1.Text1.Text = Form1.Text1.Text & "" & vbCrLf
Form1.Text1.Text = Form1.Text1.Text & "neOS is booting" & vbCrLf
Pause (1)
Form1.Text1.Text = Form1.Text1.Text & "."
Pause (1)
Form1.Text1.Text = Form1.Text1.Text & "."
Pause (1)
Form1.Text1.Text = Form1.Text1.Text & "."
Pause (1)
Form1.Text1.Text = Form1.Text1.Text & "." & vbCrLf
Pause (1)
Form1.Text1.Text = Form1.Text1.Text & "" & vbCrLf
Pause (0.9)
Form1.Text1.Text = Form1.Text1.Text & "Welcome at Neo's Computer" & vbCrLf
Form1.Text1.Text = Form1.Text1.Text & "-------------------------" & vbCrLf
Form1.Text1.Text = Form1.Text1.Text & "" & vbCrLf
Form1.Text1.Text = Form1.Text1.Text & "c:\"

Form1.Text2.Enabled = True
Form1.Text2.SetFocus



End Function
Function send()
Form1.Text1.SelStart = (Len(Form1.Text1.Text))
Form1.Text1.Text = Form1.Text1.Text & ""
Form1.Text1.Text = Form1.Text1.Text & "sending file to 169.34.56.12"
Form1.Text1.SelStart = (Len(Form1.Text1.Text))
Pause (1)
Form1.Text1.Text = Form1.Text1.Text & " - Account GHOST" & vbCrLf
Form1.Text1.SelStart = (Len(Form1.Text1.Text))
Pause (1)
Form1.Text1.Text = Form1.Text1.Text & "sending"
Form1.Text1.SelStart = (Len(Form1.Text1.Text))
Pause (1)
Form1.Text1.Text = Form1.Text1.Text & "."
Form1.Text1.SelStart = (Len(Form1.Text1.Text))
Pause (1)
Form1.Text1.Text = Form1.Text1.Text & "."
Form1.Text1.SelStart = (Len(Form1.Text1.Text))
Pause (1)
Form1.Text1.Text = Form1.Text1.Text & "."
Form1.Text1.SelStart = (Len(Form1.Text1.Text))
Pause (1)
Form1.Text1.Text = Form1.Text1.Text & "."
Form1.Text1.SelStart = (Len(Form1.Text1.Text))
Pause (1)
Form1.Text1.Text = Form1.Text1.Text & "."
Form1.Text1.SelStart = (Len(Form1.Text1.Text))
Pause (1)
Form1.Text1.Text = Form1.Text1.Text & "." & vbCrLf
Form1.Text1.SelStart = (Len(Form1.Text1.Text))
Pause (1)
Form1.Text1.Text = Form1.Text1.Text & "File Transmitted" & vbCrLf
Form1.Text1.SelStart = (Len(Form1.Text1.Text))
Pause (1)
Form1.Text1.Text = Form1.Text1.Text & "" & vbCrLf
Form1.Text1.SelStart = (Len(Form1.Text1.Text))
Pause (3)
Form1.Text1.Text = Form1.Text1.Text & "GHOST send message" & vbCrLf
Form1.Text1.Text = Form1.Text1.Text & "Thanks for doing your job" & vbCrLf & vbCrLf
Form1.Text1.Text = Form1.Text1.Text & "" & vbCrLf
Form1.Text1.Text = Form1.Text1.Text & "" & vbCrLf
Form1.Text1.SelStart = (Len(Form1.Text1.Text))
Pause (3)
Form1.Text1.Text = Form1.Text1.Text & "That was NeOS HackingStation alpha 2.4" & vbCrLf
Form1.Text1.SelStart = (Len(Form1.Text1.Text))
Pause (1)
Form1.Text1.Text = Form1.Text1.Text & "programmed by freequenzy.wave" & vbCrLf
Form1.Text1.SelStart = (Len(Form1.Text1.Text))
Pause (1)
Form1.Text1.Text = Form1.Text1.Text & "" & vbCrLf
Form1.Text1.SelStart = (Len(Form1.Text1.Text))
Form1.Text1.Text = Form1.Text1.Text & "THx to: James Dougherty for UXC Encryption Algorithm" & vbCrLf
Form1.Text1.SelStart = (Len(Form1.Text1.Text))
Pause (1)
Form1.Text1.Text = Form1.Text1.Text & "        .:LorD NikoN:. / TechnoMusic / Belinda / Marry" & vbCrLf
Form1.Text1.SelStart = (Len(Form1.Text1.Text))
Pause (1)
Form1.Text1.Text = Form1.Text1.Text & "        mqs / shordee / and all i forgot.............." & vbCrLf
Form1.Text1.SelStart = (Len(Form1.Text1.Text))
End Function

Function FalsePass()
Form1.Text1.Text = Form1.Text1.Text & "compare passwords" & vbCrLf
Form1.Text1.SelStart = (Len(Form1.Text1.Text))
Pause (2)
Form1.Text1.Text = Form1.Text1.Text & "Password denied" & vbCrLf & vbCrLf
End Function
Function RightPass()
Form1.Text1.Text = Form1.Text1.Text & "compare passwords" & vbCrLf
Form1.Text1.SelStart = (Len(Form1.Text1.Text))
Pause (2)
Form1.Text1.Text = Form1.Text1.Text & "Password access" & vbCrLf & vbCrLf
End Function

Function Hack()

If Form1.Text2.Text = "?help" Then
Form1.Text1.Text = Form1.Text1.Text & "" & vbCrLf
Form1.Text1.Text = Form1.Text1.Text & "cls - clear screen" & vbCrLf
Form1.Text1.Text = Form1.Text1.Text & "dir - list files" & vbCrLf
Form1.Text1.Text = Form1.Text1.Text & "exit - exit" & vbCrLf
Form1.Text1.Text = Form1.Text1.Text & "neonewsoff - power of neonews" & vbCrLf
Form1.Text1.Text = Form1.Text1.Text & "neonewson - power on neonews" & vbCrLf
Form1.Text1.Text = Form1.Text1.Text & "read - open files" & vbCrLf
Exit Function
End If

If Form1.Text2.Text = "neonewsoff" Then
Form1.Label1.Visible = False
Form1.Text1.Text = Form1.Text1.Text & "" & vbCrLf
Form1.Text1.Text = Form1.Text1.Text & "NeoNews Version 2.4 is powered off" & vbCrLf
Exit Function
End If

If Form1.Text2.Text = "neonewson" Then
Form1.Label1.Visible = True
Form1.Text1.Text = Form1.Text1.Text & "" & vbCrLf
Form1.Text1.Text = Form1.Text1.Text & "NeoNews Version 2.4 is powered on" & vbCrLf
Form1.Label1.Top = 360
Exit Function
End If

If Form1.Text2.Text = "c:\" Then
HPath = "c:\"
Exit Function
End If

If Form1.Text2.Text = "d:\" Then
HPath2 = HPath
If HPath = "d:\" Then
Exit Function
End If
Form1.Text1.Text = Form1.Text1.Text & "" & vbCrLf
Call HackDriveD
Exit Function
End If

If Form1.Text2.Text = "a:\" Then
HPath2 = HPath
    If HPath = "a:\" Then
        Exit Function
    End If
Form1.Text1.Text = Form1.Text1.Text & "" & vbCrLf
Call HackDriveA
Exit Function
End If

If Form1.Text2.Text = "a:\system" Then
HPath2 = HPath
If HPath = "a:\" Then
HPath = "a:\system\"
Exit Function
End If
Form1.Text1.Text = Form1.Text1.Text & "" & vbCrLf
Form1.Text1.Text = Form1.Text1.Text & "please change to specific drive to view structure" & vbCrLf
Exit Function
End If

If Form1.Text2.Text = "d:\files" Then
HPath2 = HPath
If HPath = "d:\" Then
HPath = "d:\files\"
Exit Function
End If
Form1.Text1.Text = Form1.Text1.Text & "" & vbCrLf
Form1.Text1.Text = Form1.Text1.Text & "please change to specific drive to view structure" & vbCrLf
Exit Function
End If

If Form1.Text2.Text = "b:\" Then
Form1.Text1.Text = Form1.Text1.Text & "" & vbCrLf
Form1.Text1.Text = Form1.Text1.Text & "Drive don't exist" & vbCrLf
Exit Function
End If

If Form1.Text2.Text = "cls" Then
Form1.Text1.Text = ""
Exit Function
End If

If Form1.Text2.Text = "exit" Then
Form1.Text1.Text = ""
Call ShutDown
Exit Function
End If

If Form1.Text2.Text = "send" Then
    Form1.Text1.Text = Form1.Text1.Text & "" & vbCrLf
    Form1.Text1.Text = Form1.Text1.Text & "send (file) (ip)" & vbCrLf
    Exit Function
End If

If Form1.Text2.Text = "send security.loc 169.34.56.12" Then
    If HPath = "d:\files\" Then
        Call send
        Exit Function
    Else
        Form1.Text1.Text = Form1.Text1.Text & "" & vbCrLf
        Form1.Text1.Text = Form1.Text1.Text & "file was not found" & vbCrLf
        Exit Function
    End If
End If

If Form1.Text2.Text = "dir" Then
    If HPath = "c:\" Then
        Form1.Text1.Text = Form1.Text1.Text & "" & vbCrLf
        Form1.Text1.Text = Form1.Text1.Text & "abis.dot" & vbCrLf
        Form1.Text1.Text = Form1.Text1.Text & "boot.txt" & vbCrLf
        Form1.Text1.Text = Form1.Text1.Text & "mail.txt" & vbCrLf
        Form1.Text1.Text = Form1.Text1.Text & "mute.cfg" & vbCrLf
        Form1.Text1.Text = Form1.Text1.Text & "pass.txt" & vbCrLf
    End If
    
    '-----------a:\------------
    
    If HPath = "a:\" Then
        Form1.Text1.Text = Form1.Text1.Text & "" & vbCrLf
        Form1.Text1.Text = Form1.Text1.Text & "system" & vbCrLf
        Form1.Text1.Text = Form1.Text1.Text & "matrix.exe" & vbCrLf
        Form1.Text1.Text = Form1.Text1.Text & "res.txt" & vbCrLf
    End If
    
    If HPath = "a:\system\" Then
        Form1.Text1.Text = Form1.Text1.Text & "" & vbCrLf
        Form1.Text1.Text = Form1.Text1.Text & "ping.exe" & vbCrLf
        Form1.Text1.Text = Form1.Text1.Text & "send.exe" & vbCrLf
        Form1.Text1.Text = Form1.Text1.Text & "secur.txt" & vbCrLf
        HSend = True
        Form1.send.Visible = True
    End If
    
    '--------d:\----------------
    
    If HPath = "d:\" Then
        Form1.Text1.Text = Form1.Text1.Text & "" & vbCrLf
        Form1.Text1.Text = Form1.Text1.Text & "files" & vbCrLf
        Form1.Text1.Text = Form1.Text1.Text & "agent.txt" & vbCrLf
        Form1.Text1.Text = Form1.Text1.Text & "girl.img" & vbCrLf
        Form1.Text1.Text = Form1.Text1.Text & "matrix.txt" & vbCrLf
    End If
    
    If HPath = "d:\files\" Then
        Form1.Text1.Text = Form1.Text1.Text & "" & vbCrLf
        Form1.Text1.Text = Form1.Text1.Text & "security.loc" & vbCrLf
    End If
    
Exit Function
End If

If Left(Form1.Text2.Text, 4) = "read" Then
    If HPath = "c:\" Then
        If Form1.Text2.Text = "read abis.dot" Then
            Form1.Text1.Text = Form1.Text1.Text & "" & vbCrLf
            Form1.Text1.Text = Form1.Text1.Text & "file is not compatible with this command." & vbCrLf
        ElseIf Form1.Text2.Text = "read boot.txt" Then
            Form1.Text1.Text = Form1.Text1.Text & "" & vbCrLf
            HDat = "\DAT\1.dat"
            Call Decrypt
        ElseIf Form1.Text2.Text = "read mail.txt" Then
            Form1.Text1.Text = Form1.Text1.Text & "" & vbCrLf
            HDat = "\DAT\2.dat"
            Call Decrypt
        ElseIf Form1.Text2.Text = "read mute.cfg" Then
            Form1.Text1.Text = Form1.Text1.Text & "" & vbCrLf
            HDat = "\DAT\3.dat"
            Call Decrypt
        ElseIf Form1.Text2.Text = "read pass.txt" Then
            Form1.Text1.Text = Form1.Text1.Text & "" & vbCrLf
            HDat = "\DAT\4.dat"
            Call Decrypt
        Else
            Form1.Text1.Text = Form1.Text1.Text & "" & vbCrLf
            Form1.Text1.Text = Form1.Text1.Text & "File dont exist" & vbCrLf
        End If
    End If


    '----------A:\--------------------------------
    '---------------------------------------------
    
    If HPath = "a:\" Then
        If Form1.Text2.Text = "read matrix.exe" Then
            Form1.Text1.Text = Form1.Text1.Text & "" & vbCrLf
            Form1.Text1.Text = Form1.Text1.Text & "file is not compatible with this command." & vbCrLf
        ElseIf Form1.Text2.Text = "read res.txt" Then
            Form1.Text1.Text = Form1.Text1.Text & "" & vbCrLf
            HDat = "\DAT\11.dat"
            Call Decrypt
        Else
            Form1.Text1.Text = Form1.Text1.Text & "" & vbCrLf
            Form1.Text1.Text = Form1.Text1.Text & "File dont exist" & vbCrLf
        End If
    End If
    
    If HPath = "a:\system\" Then
        If Form1.Text2.Text = "read ping.exe" Then
            Form1.Text1.Text = Form1.Text1.Text & "" & vbCrLf
            Form1.Text1.Text = Form1.Text1.Text & "file is not compatible with this command." & vbCrLf
        ElseIf Form1.Text2.Text = "read send.exe" Then
            Form1.Text1.Text = Form1.Text1.Text & "" & vbCrLf
            Form1.Text1.Text = Form1.Text1.Text & "file is not compatible with this command." & vbCrLf
        ElseIf Form1.Text2.Text = "read secur.txt" Then
            Form1.Text1.Text = Form1.Text1.Text & "" & vbCrLf
            HDat = "\DAT\12.dat"
            Call Decrypt
        Else
            Form1.Text1.Text = Form1.Text1.Text & "" & vbCrLf
            Form1.Text1.Text = Form1.Text1.Text & "File dont exist" & vbCrLf
        End If
    End If
    
    If HPath = "d:\" Then
        If Form1.Text2.Text = "read girl.img" Then
            Form1.Text1.Text = Form1.Text1.Text & "" & vbCrLf
            Form1.Text1.Text = Form1.Text1.Text & "file is not compatible with this command." & vbCrLf
        ElseIf Form1.Text2.Text = "read agent.txt" Then
            Form1.Text1.Text = Form1.Text1.Text & "" & vbCrLf
            HDat = "\DAT\21.dat"
            Call Decrypt
        ElseIf Form1.Text2.Text = "read matrix.txt" Then
            Form1.Text1.Text = Form1.Text1.Text & "" & vbCrLf
            HDat = "\DAT\22.dat"
            Call Decrypt
        Else
            Form1.Text1.Text = Form1.Text1.Text & "" & vbCrLf
            Form1.Text1.Text = Form1.Text1.Text & "File dont exist" & vbCrLf
        End If
    End If
    
    If HPath = "d:\files\" Then
        If Form1.Text2.Text = "read security.loc" Then
            Form1.Text1.Text = Form1.Text1.Text & "" & vbCrLf
            Form1.Text1.Text = Form1.Text1.Text & "file is not compatible with this command." & vbCrLf
        Else
            Form1.Text1.Text = Form1.Text1.Text & "" & vbCrLf
            Form1.Text1.Text = Form1.Text1.Text & "File dont exist" & vbCrLf
        End If
    End If

Exit Function
End If

Form1.Text1.Text = Form1.Text1.Text & "Unknown Command" & vbCrLf

End Function
Function ShutDown()
Form1.Text2.Text = ""
Form1.Text2.Enabled = False
Form1.Text1.Text = "Closing Connections..." & vbCrLf
Pause (2)
Form1.Text1.Text = Form1.Text1.Text & "Sending Kill all Term..." & vbCrLf
Pause (3)
Form1.Text1.Text = Form1.Text1.Text & "Shutting off" & vbCrLf
Pause (3)
End
End Function
Function HackDriveA()

Form1.Text1.Text = Form1.Text1.Text & "please type the password for this Drive" & vbCrLf
HPath = ""
PassDriveA = True
End Function

Function HackDriveD()

Form1.Text1.Text = Form1.Text1.Text & "please type the password for this Drive" & vbCrLf
HPath = ""
PassDriveD = True
End Function
