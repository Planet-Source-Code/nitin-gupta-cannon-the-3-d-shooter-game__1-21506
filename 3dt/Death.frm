VERSION 5.00
Begin VB.Form Death 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   480
      Top             =   2640
   End
   Begin VB.Image Image1 
      Height          =   2640
      Left            =   0
      Picture         =   "Death.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9840
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
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
      Height          =   1335
      Left            =   840
      TabIndex        =   0
      Top             =   3240
      Width           =   7695
   End
   Begin VB.Image Image2 
      Height          =   2640
      Left            =   120
      Picture         =   "Death.frx":186D4
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   9840
   End
End
Attribute VB_Name = "Death"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Label1.Caption = "DEATH...the aliens have defeated you. Your escape has cost you your life. Better luck next time :)"
End Sub

Private Sub Image1_Click()
On Error Resume Next
GForm.Show
Unload Me
End Sub

Private Sub Image2_Click()
On Error Resume Next
GForm.Show
Unload Me
End Sub

Private Sub Label1_Click()
On Error Resume Next
GForm.Show
Unload Me
End Sub
