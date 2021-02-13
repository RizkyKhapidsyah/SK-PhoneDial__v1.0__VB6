VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "PHONE DIALER"
   ClientHeight    =   5835
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3915
   Icon            =   "PhoneD.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5835
   ScaleWidth      =   3915
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   23
      Top             =   5460
      Width           =   3915
      _ExtentX        =   6906
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.OptionButton Option2 
      Caption         =   "MF TONES"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   840
      TabIndex        =   22
      Top             =   5160
      Width           =   2115
   End
   Begin VB.OptionButton Option1 
      Caption         =   "DTMF TONES"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   21
      Top             =   4800
      Width           =   2055
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Not In ServiceTones"
      Height          =   615
      Left            =   2880
      TabIndex        =   20
      Top             =   3240
      Width           =   735
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Not In Service"
      Height          =   495
      Left            =   2880
      TabIndex        =   19
      Top             =   4320
      Width           =   735
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Busy"
      Height          =   495
      Left            =   2880
      TabIndex        =   18
      Top             =   2280
      Width           =   735
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Ring"
      Height          =   495
      Left            =   2880
      TabIndex        =   17
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Dial Tone"
      Height          =   495
      Left            =   2880
      TabIndex        =   16
      Top             =   360
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Play #"
      Height          =   495
      Left            =   1920
      TabIndex        =   15
      Top             =   4080
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   720
      MaxLength       =   19
      TabIndex        =   14
      Top             =   3600
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   495
      Index           =   11
      Left            =   600
      TabIndex        =   12
      Top             =   2880
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   495
      Index           =   10
      Left            =   2040
      TabIndex        =   11
      Top             =   2880
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Clear #"
      Height          =   495
      Left            =   720
      TabIndex        =   10
      Top             =   4080
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   495
      Index           =   9
      Left            =   2040
      TabIndex        =   9
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   495
      Index           =   8
      Left            =   1320
      TabIndex        =   8
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   495
      Index           =   7
      Left            =   600
      TabIndex        =   7
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   495
      Index           =   6
      Left            =   2040
      TabIndex        =   6
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   495
      Index           =   5
      Left            =   1320
      TabIndex        =   5
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   495
      Index           =   4
      Left            =   600
      TabIndex        =   4
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   495
      Index           =   3
      Left            =   2040
      TabIndex        =   3
      Top             =   720
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   495
      Index           =   2
      Left            =   1320
      TabIndex        =   2
      Top             =   720
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   495
      Index           =   1
      Left            =   600
      TabIndex        =   1
      Top             =   720
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   495
      Index           =   0
      Left            =   1320
      TabIndex        =   0
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "PHONE DIALER"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   13
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function sndPlaySound Lib "winmm" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Sub PlaySound(strFileName As String)
sndPlaySound strFileName, 1
End Sub



Private Sub Command2_Click()
  Dim temp As String
  Dim mydir$, TempText$
  StatusBar1.Panels(1) = "PLAYING NUMBER"
  mydir$ = App.Path & "\DTMF TONES\"
    If Option1.Value = True Then mydir$ = App.Path & "\DTMF TONES\"
  If Option2.Value = True Then mydir$ = "D:\Visual Basic Forms\Phone Dialer\MF TONES\"
  TempText$ = LTrim$(RTrim$(Text1.Text))
  
  For i = 1 To Len(TempText$)
    temp = Mid$(TempText$, i, 1)
    If IsNumeric(temp) Then
        myFile$ = LTrim$(Str$(temp)) + ".wav"
    Else
        myFile$ = String((Abs(temp = "#")), "#") & String((Abs(temp = "X")), "X") + ".wav"
    End If
   
   lSound = sndSound(mydir$ & myFile$, 1)
   'PauseTime = 0.1  ' Set duration.
   Start = Timer   ' Set start time.
   Do While Timer < Start + 0.3
      DoEvents   ' Yield to other processes.
   Loop
  Next

End Sub

Private Sub Command3_Click()
StatusBar1.Panels(1) = "DIAL TONE"
PlaySound App.Path & "\SOUND EFFECTS\Dial Tone.wav"
End Sub

Private Sub Command4_Click()
'clear
StatusBar1.Panels(1) = "CLEAR NUMBER"
Text1.Text = ""
Form1.SetFocus
End Sub

Private Sub Command5_Click()

StatusBar1.Panels(1) = "RING"
PlaySound App.Path & "\SOUND EFFECTS\Ring.wav"
End Sub

Private Sub Command6_Click()

StatusBar1.Panels(1) = "BUSY"
PlaySound App.Path & "\SOUND EFFECTS\Busy.wav"
End Sub

Private Sub Command7_Click()

StatusBar1.Panels(1) = "NOT IN SERVICE"
PlaySound App.Path & "\SOUND EFFECTS\Not In Service.wav"
End Sub

Private Sub Command8_Click()

StatusBar1.Panels(1) = "NOT IN SERVICE TONES"
PlaySound App.Path & "\SOUND EFFECTS\We're Sorry Tones.wav"
End Sub



Private Sub Form_Load()

For X = 0 To 9
    Command1(X).Caption = X
Next

Command1(10).Caption = "#"
Command1(11).Caption = "*"
Option1.Value = True
End Sub
Private Sub Command1_Click(index As Integer)
  mydir$ = "\DTMF TONES\"
  StatusBar1.Panels(1) = "DIALING NUMBER"
  If Option1.Value = True Then mydir$ = App.Path & "\DTMF TONES\"
  If Option2.Value = True Then mydir$ = App.Path & "\MF TONES\"
  temp = index
  'Text1.Enabled = True
  
  If temp > 9 Then
    myFile$ = String((Abs(index = 10)), "#") & String((Abs(index = 11)), "X") + ".wav"
    Text1.Text = Text1.Text & Left$(myFile$, 1)
  Else
    myFile$ = LTrim$(Str$(index)) + ".wav"
    Text1.Text = Text1.Text & LTrim$(Str$(index))
  End If
  lSound = sndSound(mydir$ & myFile$, 1)
    
End Sub
Private Sub Command1_GotFocus(index As Integer)
temp = index
End Sub

Private Sub Command1_KeyPress(index As Integer, KeyAscii As Integer)
    doKey (KeyAscii)
 
End Sub
Sub doKey(KeyAscii)

If (KeyAscii > 47) And (KeyAscii < 58) Or (KeyAscii = 35) Or (KeyAscii = 42) Then
   If KeyAscii > 42 Then
        temp = KeyAscii - 48
   Else
        'this figures out whether its an asterisk(10) or a pound key (11)
        ' it if either of the comparisions are true (-1) then the ABS converts
        ' it to a 1 and multiplies it by the number you want assigned. If it's not
        ' then it's a zero and the two added together give only one answer
        temp = (Abs(KeyAscii = 35) * 10) + (Abs(KeyAscii = 42) * 11)
   End If
   junk = SendMessage(Command1(temp).hwnd, BM_SETSTATE, True, 0)
   Command1_Click (temp)
   PauseIt
   
   junk = SendMessage(Command1(temp).hwnd, BM_SETSTATE, False, 0)
    
End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   doKey (KeyAscii)
End Sub

Sub PauseIt()
   'PauseTime = 0.1  ' Set duration.
   Start = Timer   ' Set start time.
   Do While Timer < Start + 0.05
      DoEvents   ' Yield to other processes.
   Loop
End Sub

Private Sub Label2_Click()
StatusBar1.Panels(2) = "PHONE DIALER"
frmAbout.Show
End Sub

Private Sub Option1_Click()
StatusBar1.Panels(1) = "DTMF TONES"
End Sub

Private Sub Option2_Click()
StatusBar1.Panels(1) = "MF TONES"
End Sub

Private Sub StatusBar1_PanelClick(ByVal Panel As MSComctlLib.Panel)
StatusBar1.Panels(1) = "Hello"
End Sub

