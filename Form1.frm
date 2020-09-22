VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "NeOS HackingStation © alpha 0.5 "
   ClientHeight    =   7965
   ClientLeft      =   150
   ClientTop       =   495
   ClientWidth     =   10845
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7965
   ScaleWidth      =   10845
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton Command2 
      Caption         =   "addon on"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   170
      Left            =   1215
      TabIndex        =   12
      Top             =   6785
      Width           =   960
   End
   Begin VB.CommandButton Command1 
      Caption         =   "addon off"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   170
      Left            =   135
      TabIndex        =   11
      Top             =   6785
      Width           =   960
   End
   Begin MSComctlLib.ProgressBar ProgressBar4 
      Height          =   135
      Left            =   150
      TabIndex        =   10
      Top             =   7440
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar ProgressBar3 
      Height          =   135
      Left            =   150
      TabIndex        =   9
      Top             =   7560
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar ProgressBar2 
      Height          =   135
      Left            =   150
      TabIndex        =   8
      Top             =   7680
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   10560
      TabIndex        =   7
      Top             =   7440
      Width           =   135
      _ExtentX        =   238
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   1770
      Orientation     =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   255
      Left            =   8640
      TabIndex        =   5
      Top             =   7560
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   450
      _Version        =   393216
      Min             =   1
      SelStart        =   1
      Value           =   1
      TextPosition    =   1
   End
   Begin VB.Timer Timer3 
      Interval        =   1
      Left            =   0
      Top             =   1200
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   0
      Top             =   720
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '2D
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      ScaleHeight     =   345
      ScaleWidth      =   10545
      TabIndex        =   2
      Top             =   6960
      Width           =   10575
      Begin VB.Label Label1 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   2000
         Left            =   0
         TabIndex        =   3
         Top             =   360
         Width           =   10575
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   0
      Top             =   240
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  '2D
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   6480
      Width           =   10575
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  '2D
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   6255
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   0
      Top             =   120
      Width           =   10575
   End
   Begin VB.Shape Shape4 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000C0C0&
      FillStyle       =   0  'Ausgefüllt
      Height          =   150
      Left            =   970
      Shape           =   3  'Kreis
      Top             =   7680
      Width           =   150
   End
   Begin VB.Shape Shape3 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000C0&
      FillStyle       =   0  'Ausgefüllt
      Height          =   150
      Left            =   970
      Shape           =   3  'Kreis
      Top             =   7440
      Width           =   150
   End
   Begin VB.Shape Shape2 
      Height          =   435
      Left            =   945
      Top             =   7410
      Width           =   190
   End
   Begin VB.Shape Shape1 
      Height          =   435
      Left            =   120
      Top             =   7410
      Width           =   795
   End
   Begin VB.Label Label3 
      Caption         =   "NeoNews ScrollSpeed"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   8760
      TabIndex        =   6
      Top             =   7440
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Zentriert
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "---------- NeoNews 2.4 ----------"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   3915
      TabIndex        =   4
      Top             =   6840
      Width           =   3120
   End
   Begin VB.Menu commands 
      Caption         =   "Commands"
      Begin VB.Menu cls 
         Caption         =   "cls"
      End
      Begin VB.Menu dir 
         Caption         =   "dir"
      End
      Begin VB.Menu read 
         Caption         =   "read"
      End
      Begin VB.Menu exit 
         Caption         =   "exit"
      End
      Begin VB.Menu neooff 
         Caption         =   "neonewsoff"
      End
      Begin VB.Menu neoon 
         Caption         =   "neonewson"
      End
      Begin VB.Menu send 
         Caption         =   "send"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu about 
      Caption         =   "About"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub about_Click()
MsgBox "Chilling Welcome to you," & vbCrLf & _
"this is my first try to do somthing coolness 4 hacking peoples." & vbCrLf & _
"Because I was searching for something like that, but i dont get it. So i did it myself" & vbCrLf & _
"I'm sorry 4 that bad english I'm a german... ;) I hope you've fun with da shit!!!" & vbCrLf & vbCrLf & _
"email me for questions mrisklan@hotmail.com", vbOKOnly, Form1.Caption

End Sub

Private Sub cls_Click()
Text2.Text = "cls"
End Sub

Private Sub Command1_Click()
Call big
End Sub

Private Sub Command2_Click()
Call small
End Sub

Private Sub dir_Click()
Text2.Text = "dir"
End Sub

Private Sub exit_Click()
Text2.Text = "exit"
End Sub

Private Sub Form_Load()
Dim temp As String
Slider1.Value = 3
Speed = 3
HPath = "c:\"
PassDriveA = False
    Timer1.Enabled = True
    Timer3.Enabled = True
    Text2.Enabled = False
    Timer2.Enabled = True
    Label1.Caption = "NeoNews v2.4" & vbCrLf
    On Error Resume Next
    Open App.Path & "\neonews.txt" For Input As 1
    Do
    Line Input #1, temp
    Label1.Caption = Label1.Caption & temp & vbCrLf
    Loop Until EOF(1) = True
    Close 1

End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub neooff_Click()
Text2.Text = "neonewsoff"
End Sub

Private Sub neoon_Click()
Text2.Text = "neonewson"
End Sub

Private Sub read_Click()
Text2.Text = "read"
End Sub

Private Sub send_Click()
Text2.Text = "send (file) (IP)"
End Sub

Private Sub Slider1_Scroll()
Speed = Slider1.Value
End Sub

Private Sub Text2_KeyUp(KeyCode As Integer, Shift As Integer)
    If PassDriveA = True Then
        If KeyCode = 13 Then
            If Text2.Text = "raider" Then
                Text1.Text = Text1.Text & Text2.Text & vbCrLf
                Text1.Text = Text1.Text & vbCrLf
                Text2.Text = ""
                Call RightPass
                HPath = "a:\"
                Text1.Text = Text1.Text & HPath
                PassDriveA = False
                Text1.SelStart = (Len(Text1.Text))
                Exit Sub
            End If
        Text1.Text = Text1.Text & Text2.Text & vbCrLf
        Text1.Text = Text1.Text & vbCrLf
        Text2.Text = ""
        Call FalsePass
        HPath = HPath2
        Text1.Text = Text1.Text & HPath
        PassDriveA = False
        Text1.SelStart = (Len(Text1.Text))
        Exit Sub
        End If
    End If
    
        If PassDriveD = True Then
        If KeyCode = 13 Then
            If Text2.Text = "trust" Then
                Text1.Text = Text1.Text & Text2.Text & vbCrLf
                Text1.Text = Text1.Text & vbCrLf
                Text2.Text = ""
                Call RightPass
                HPath = "d:\"
                Text1.Text = Text1.Text & HPath
                PassDriveD = False
                Text1.SelStart = (Len(Text1.Text))
                Exit Sub
            End If
        Text1.Text = Text1.Text & Text2.Text & vbCrLf
        Text1.Text = Text1.Text & vbCrLf
        Text2.Text = ""
        Call FalsePass
        HPath = HPath2
        Text1.Text = Text1.Text & HPath
        PassDriveD = False
        Text1.SelStart = (Len(Text1.Text))
        Exit Sub
        End If
    End If
        
    If KeyCode = 13 Then
        Text1.Text = Text1.Text & Text2.Text & vbCrLf
        Call Hack
        Text2.Text = ""
        Text1.Text = Text1.Text & vbCrLf
        Text1.Text = Text1.Text & HPath
    End If
    
    Text1.SelStart = (Len(Text1.Text))

End Sub

Private Sub Timer1_Timer()
Unload Form2
Timer3.Enabled = False
Call Scan
Timer1.Enabled = False
End Sub

Private Sub Timer2_Timer()
Dim i As String
Shape4.Visible = False
Shape3.Visible = True
Label1.Top = Label1.Top - Speed
ProgressBar1.Value = ProgressBar1.Value + Speed

ProgressBar2.Value = ProgressBar2.Value + 10
ProgressBar3.Value = ProgressBar3.Value + 25
ProgressBar4.Value = ProgressBar4.Value + 5

If Label1.Top < -1400 Then
    Label1.Top = 360
    ProgressBar1.Value = 0
End If

If ProgressBar2.Value = 100 Then
    ProgressBar2.Value = 0
    Shape3.Visible = False
    Shape4.Visible = True
End If
If ProgressBar3.Value = 100 Then
    ProgressBar3.Value = 0
    Shape3.Visible = False
    Shape4.Visible = True
End If
If ProgressBar4.Value = 100 Then
    ProgressBar4.Value = 0
    Shape3.Visible = False
    Shape4.Visible = True
End If



End Sub

Private Sub Timer3_Timer()
Form2.Show
End Sub
