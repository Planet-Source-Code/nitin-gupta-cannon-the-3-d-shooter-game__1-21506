VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "SWFLASH.OCX"
Begin VB.Form Form3 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   3120
      Top             =   8040
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2520
      Top             =   8040
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Mov1 
      Height          =   7200
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9600
      _cx             =   4211237
      _cy             =   4207004
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Movie1_GotFocus()
End Sub

Private Sub Form_Load()
Mov1.Movie = App.Path + "\nitin.swf"
Mov1.Play
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Load GForm
GForm.Show
End Sub

Private Sub Timer1_Timer()
If Mov1.IsPlaying = 0 Then
   GForm.Unl1
End If
If (GetKeyState(vbKeyEscape) And KEY_DOWN) Or (GetKeyState(vbKeySpace) And KEY_DOWN) Then
   GForm.Unl1
End If
End Sub

