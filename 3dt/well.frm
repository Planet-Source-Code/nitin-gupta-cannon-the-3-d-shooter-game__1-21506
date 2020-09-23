VERSION 5.00
Begin VB.Form Welldone 
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
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   480
      Top             =   2520
   End
   Begin VB.Image Image2 
      Height          =   2640
      Left            =   120
      Picture         =   "well.frx":0000
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   9840
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Justov"
         Size            =   15.75
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
      Top             =   3120
      Width           =   7695
   End
   Begin VB.Image Image1 
      Height          =   2640
      Left            =   0
      Picture         =   "well.frx":15F942
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   9840
   End
End
Attribute VB_Name = "Welldone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
'Unload GForm
'Unload Form2
Label1.Caption = "Well done, Cannon. You have succesfully won your freedom from the aliens. Now you may peacefully return back to your home - the Earth. Proud humans are waiting for you there eagerly. You are a HERO !!"
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

