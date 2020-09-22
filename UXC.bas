Attribute VB_Name = "UXC"


'Dim Encryptor As New UXC

Public Function EncryptUXC(Data As String, Key As String) As String
 On Local Error Resume Next
 Dim i As Long
 
 If Key = "" Or Len(Key) < 6 Then
  MsgBox "Please enter a key with 6 characters", vbOKOnly Or vbInformation, "UXC"
  EncryptUXC = Data
  Exit Function
 End If
 
 For i = 1 To CDbl(Len(Key))
  EncryptUXC = EncryptUXC + Chr(Asc(Mid(Key, i, 1)) + Len(WeekdayName(5, , vbSaturday)) + Len(TimeSerial(18, 21, 1)))
 Next
 
 For i = 1 To CDbl(Len(Data))
  EncryptUXC = EncryptUXC + Chr(Asc(Mid(Data, i, 1)) + ((Len(Key) + (40 - Len(Key))) / 2) + Len(TimeSerial(3, 55, 37)) + ((Len(WeekdayName(1, , vbSunday)) / 10) + 2))
 Next
 
End Function

Public Function DecryptUXC(Data As String, Key As String) As String
 On Local Error Resume Next
 Dim i As Long
 Dim tmpData As String
 
 If Key = "" Then DecryptUXC = Data: Exit Function
 
 For i = 1 To CDbl(Len(Key))
  tmpData = tmpData + Chr(Asc(Mid(Key, i, 1)) + Len(WeekdayName(5, , vbSaturday)) + Len(TimeSerial(18, 21, 1)))
 Next
 
 If Left$(tmpData, 1) = Left$(Data, 1) And _
    Left$(tmpData, 2) = Left$(Data, 2) And _
    Left$(tmpData, 3) = Left$(Data, 3) And _
    Left$(tmpData, 4) = Left$(Data, 4) And _
    Left$(tmpData, 5) = Left$(Data, 5) And _
    Left$(tmpData, 6) = Left$(Data, 6) Then
  For i = 1 To CDbl(Len(tmpData))
   DecryptUXC = DecryptUXC + Chr(Asc(Mid(Key, i, 1)) - Len(WeekdayName(5, , vbSaturday)) - Len(TimeSerial(18, 21, 1)))
  Next
  For i = Len(tmpData) + 1 To CDbl(Len(Data))
   DecryptUXC = DecryptUXC + Chr(Asc(Mid(Data, i, 1)) - ((Len(Key) + (40 - Len(Key))) / 2) - Len(TimeSerial(3, 55, 37)) - ((Len(WeekdayName(1, , vbSunday)) / 10) + 2))
  Next
  DecryptUXC = Right$(DecryptUXC, (Len(DecryptUXC) - Len(Key)))
 Else
  DecryptUXC = Data
  Exit Function
 End If
 
End Function

Public Function ProEncryptUXC(Data As String, Key As String) As String
 Dim tmpString As String
 
 tmpString = EncryptUXC(Data, Key)
 ProEncryptUXC = EncryptUXC(tmpString, Key)
 
End Function

Public Function ProDecryptUXC(Data As String, Key As String) As String
 Dim tmpString As String
 
 tmpString = DecryptUXC(Data, Key)
 ProDecryptUXC = DecryptUXC(tmpString, Key)
 
End Function

'----------------------------------------------------------------------



Public Sub Decrypt()
Dim temp2 As String
Dim i As String


    On Error Resume Next
    Open App.Path & HDat For Input As 1
    Do

    Input #1, temp2
    'temp2 = temp2 & temp2 & vbCrLf
    Loop Until EOF(1) = True
    Close 1
    
 temp2 = ProDecryptUXC(temp2, "231246")
 
    Form1.Text1.Text = Form1.Text1.Text & temp2 & vbCrLf
End Sub



