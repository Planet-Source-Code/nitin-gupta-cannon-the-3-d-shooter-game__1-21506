VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8970
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   11970
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Form2.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   598
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   798
   Begin VB.Timer Timer2 
      Interval        =   10
      Left            =   360
      Top             =   360
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   1440
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00000000&
      Caption         =   "Sound Enabled"
      BeginProperty Font 
         Name            =   "Justov"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7560
      TabIndex        =   9
      Top             =   4920
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4800
      Top             =   2880
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3180
      Left            =   120
      Picture         =   "Form2.frx":0152
      ScaleHeight     =   3150
      ScaleWidth      =   1155
      TabIndex        =   1
      Top             =   2160
      Width           =   1185
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   5055
      Left            =   1080
      ScaleHeight     =   335
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   199
      TabIndex        =   0
      Top             =   1320
      Width           =   3015
      Begin VB.Image Image2 
         Height          =   6000
         Left            =   360
         Picture         =   "Form2.frx":BFE4
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1500
      End
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Enable Joystick"
      BeginProperty Font 
         Name            =   "Justov"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4200
      MouseIcon       =   "Form2.frx":50656
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   4920
      Width           =   1575
   End
   Begin VB.Label Label8 
      Caption         =   "0"
      Height          =   255
      Left            =   2520
      TabIndex        =   10
      Top             =   480
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Loading"
      BeginProperty Font 
         Name            =   "Justov"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5040
      TabIndex        =   8
      Top             =   5760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   180
      Left            =   5040
      Top             =   5520
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Wimp Out !"
      BeginProperty Font 
         Name            =   "Justov"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5640
      TabIndex        =   7
      Top             =   4080
      Width           =   4215
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "About the game"
      BeginProperty Font 
         Name            =   "Justov"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5640
      TabIndex        =   6
      Top             =   3600
      Width           =   4215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Save a game"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Justov"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5640
      TabIndex        =   5
      Top             =   3120
      Width           =   4215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Load a game"
      BeginProperty Font 
         Name            =   "Justov"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5640
      TabIndex        =   4
      Top             =   2640
      Width           =   4215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Play a game"
      BeginProperty Font 
         Name            =   "Justov"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5640
      TabIndex        =   3
      Top             =   2160
      Width           =   4215
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   304
      X2              =   800
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Please take your option"
      BeginProperty Font 
         Name            =   "Justov"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4560
      TabIndex        =   2
      Top             =   1320
      Width           =   4215
   End
   Begin VB.Image Image1 
      Height          =   3120
      Left            =   4080
      Picture         =   "Form2.frx":50960
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   4920
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Check1.Value = 0 Then GForm.Stopsound
If Check1.Value = 1 Then GForm.Playsound
End Sub

Private Sub Form_Load()
If Label8.Caption = "5" Then Hide
End Sub

Private Sub Label2_Click()
If Label8.Caption = "0" Then
  Timer1.Enabled = True
  Shape1.Visible = True
  Label7.Visible = True
Else
  GForm.Unl
End If
End Sub

Private Sub Label3_Click()
On Error GoTo 10
CD1.CancelError = True
CD1.ShowOpen
GForm.LoadGame CD1.FileName
GForm.DoNitin
Timer1.Enabled = True
Shape1.Visible = True
Label7.Visible = True
10
End Sub

Private Sub Label4_Click()
On Error GoTo 20
CD1.CancelError = True
CD1.ShowSave
GForm.SaveGame CD1.FileName
Hide
20
End Sub

Private Sub Label5_Click()
frmAbout.Show
Me.Enabled = False
End Sub

Private Sub Label6_Click()
ChangeScreenSettings 800, 600, 24
End
End Sub

Private Sub Label9_Click()
EnableJoy
End Sub

Private Sub Timer1_Timer()
Shape1.Width = Shape1.Width + 10
If Shape1.Width >= 200 Then
If GForm.Enabled = False Then GForm.Unl
End If
End Sub

Private Sub Timer2_Timer()
If Label8.Caption = "5" Then Hide
End Sub

Public Sub EnableJoy()
Dim rt As Long
Dim JoyTestInfo As JOYINFO
Dim JoyStickCaps As JOYCAPS

'Test to see if a Joystick is connected
rt = joyGetPos(JOYSTICK1, JoyTestInfo)

If rt <> JOYERR_NOERROR Then
    If rt = JOYERR_UNPLUGGED Then
        MsgBox "No joystick connected" & vbCrLf & "Finishing..."
    ElseIf rt = MMSYSERR_NODRIVER Then
        MsgBox "No Joystick driver on system" & vbCrLf & "Finishing..."
    Else
        MsgBox "Unknown Error" & vbCrLf & "finishing..."
    End If
    Exit Sub
End If

'Get the max and min position on the joystick
joyGetDevCaps JOYSTICK1, JoyStickCaps, Len(JoyStickCaps)
GForm.Label5.Caption = "1"
End Sub
