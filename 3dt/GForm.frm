VERSION 5.00
Object = "{82351433-9094-11D1-A24B-00A0C932C7DF}#1.5#0"; "ANIGIF.OCX"
Object = "{C61830C1-8B47-11D4-9F3F-0000B45C4CF6}#1.0#0"; "EASYSOUND.OCX"
Begin VB.Form GForm 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "VB DOOM - by Simon Price - www.VBgames.co.uk"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   DrawWidth       =   2
   ForeColor       =   &H00FFFFFF&
   Icon            =   "GForm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2280
      Top             =   480
   End
   Begin VB.PictureBox Picture4 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3600
      Left            =   2880
      Picture         =   "GForm.frx":0442
      ScaleHeight     =   3600
      ScaleWidth      =   4800
      TabIndex        =   22
      Top             =   2040
      Visible         =   0   'False
      Width           =   4800
   End
   Begin VB.Timer Aspeak 
      Enabled         =   0   'False
      Interval        =   7000
      Left            =   1920
      Top             =   960
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1080
      Top             =   840
   End
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   8160
      Picture         =   "GForm.frx":38884
      ScaleHeight     =   615
      ScaleWidth      =   1215
      TabIndex        =   20
      Top             =   6720
      Width           =   1215
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "100"
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
         Left            =   -120
         TabIndex        =   21
         Top             =   0
         Width           =   975
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   360
      Picture         =   "GForm.frx":3BAF6
      ScaleHeight     =   615
      ScaleWidth      =   1215
      TabIndex        =   18
      Top             =   6720
      Width           =   1215
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "100"
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
         Left            =   -120
         TabIndex        =   19
         Top             =   0
         Width           =   975
      End
   End
   Begin AniGIFCtrl.AniGIF AniGIF1 
      Height          =   975
      Left            =   0
      TabIndex        =   17
      Top             =   6600
      Width           =   9600
      BackColor       =   12632256
      PLaying         =   -1  'True
      Transparent     =   -1  'True
      Speed           =   1
      Stretch         =   0
      AutoSize        =   0   'False
      SequenceString  =   ""
      Sequence        =   0
      HTTPProxy       =   ""
      HTTPUserName    =   ""
      HTTPPassword    =   ""
      MousePointer    =   0
      GIF             =   "GForm.frx":3ED68
      ExtendWidth     =   16933
      ExtendHeight    =   1720
      Loop            =   0
      AutoRewind      =   0   'False
      Synchronized    =   -1  'True
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   2280
      ScaleHeight     =   145
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   249
      TabIndex        =   16
      Top             =   3960
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.PictureBox Buf 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   7680
      ScaleHeight     =   95
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   223
      TabIndex        =   15
      Top             =   5160
      Visible         =   0   'False
      Width           =   3375
      Begin VB.Image Image2 
         Height          =   1500
         Left            =   0
         Picture         =   "GForm.frx":41FE8
         Stretch         =   -1  'True
         Top             =   0
         Visible         =   0   'False
         Width           =   2250
      End
   End
   Begin VB.PictureBox Bufm 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   7920
      ScaleHeight     =   95
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   223
      TabIndex        =   14
      Top             =   3600
      Visible         =   0   'False
      Width           =   3375
      Begin VB.Image Image1 
         Height          =   1500
         Left            =   0
         Picture         =   "GForm.frx":4D0BA
         Stretch         =   -1  'True
         Top             =   0
         Visible         =   0   'False
         Width           =   2250
      End
   End
   Begin VB.ListBox List3 
      Height          =   840
      ItemData        =   "GForm.frx":5818C
      Left            =   240
      List            =   "GForm.frx":5818E
      TabIndex        =   13
      Top             =   4440
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.ListBox List2 
      Height          =   1035
      ItemData        =   "GForm.frx":58190
      Left            =   240
      List            =   "GForm.frx":58192
      TabIndex        =   12
      Top             =   3240
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.ListBox List1 
      Height          =   1035
      ItemData        =   "GForm.frx":58194
      Left            =   240
      List            =   "GForm.frx":58196
      TabIndex        =   11
      Top             =   1920
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.PictureBox WallPic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1212
      Index           =   8
      Left            =   4080
      ScaleHeight     =   81
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   401
      TabIndex        =   10
      Top             =   4680
      Visible         =   0   'False
      Width           =   6012
   End
   Begin VB.PictureBox WallPic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1212
      Index           =   7
      Left            =   3480
      ScaleHeight     =   81
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   401
      TabIndex        =   9
      Top             =   4440
      Visible         =   0   'False
      Width           =   6012
   End
   Begin VB.PictureBox WallPic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1212
      Index           =   6
      Left            =   6960
      ScaleHeight     =   81
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   401
      TabIndex        =   8
      Top             =   2520
      Visible         =   0   'False
      Width           =   6012
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   492
      Left            =   3240
      TabIndex        =   7
      Top             =   5160
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.PictureBox WallPic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1212
      Index           =   5
      Left            =   5880
      ScaleHeight     =   81
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   401
      TabIndex        =   6
      Top             =   2160
      Visible         =   0   'False
      Width           =   6012
   End
   Begin VB.PictureBox WallPic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1212
      Index           =   4
      Left            =   6120
      ScaleHeight     =   81
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   401
      TabIndex        =   4
      Top             =   1800
      Visible         =   0   'False
      Width           =   6012
   End
   Begin VB.PictureBox WallPic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1212
      Index           =   3
      Left            =   6360
      ScaleHeight     =   81
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   401
      TabIndex        =   5
      Top             =   1440
      Visible         =   0   'False
      Width           =   6012
   End
   Begin VB.PictureBox WallPic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1212
      Index           =   2
      Left            =   6600
      ScaleHeight     =   81
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   401
      TabIndex        =   3
      Top             =   1080
      Visible         =   0   'False
      Width           =   6012
   End
   Begin VB.PictureBox WallPic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1212
      Index           =   1
      Left            =   3600
      ScaleHeight     =   81
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   401
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   6012
   End
   Begin VB.PictureBox LevelPic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   360
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   1
      Top             =   1080
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox PB 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   3600
      Left            =   120
      Picture         =   "GForm.frx":58198
      ScaleHeight     =   240
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   320
      TabIndex        =   0
      Top             =   3720
      Visible         =   0   'False
      Width           =   4800
   End
   Begin VB.Label Label5 
      Caption         =   "0"
      Height          =   255
      Left            =   360
      TabIndex        =   25
      Top             =   1440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Score :0"
      BeginProperty Font 
         Name            =   "Justov"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   24
      Top             =   720
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Level  :1"
      BeginProperty Font 
         Name            =   "Justov"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   23
      Top             =   300
      Width           =   1695
   End
   Begin EASYSOUNDLibCtl.ESound Media 
      Left            =   2520
      OleObjectBlob   =   "GForm.frx":905DA
      Top             =   1560
   End
End
Attribute VB_Name = "GForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Exi As Integer
Dim nitin As Integer
Dim Sound As Sound
Dim Level As tLevel 'the level
Dim FrameCount As Integer 'I used this to test the fps
Dim StopFlashing As Boolean 'true tells the start message to stop flashing
Dim DoorPos As tCoOrd 'exit co-ords
Dim LevelNo As Byte 'level chosen
Dim Collisions As Boolean 'if collision detection is on or not
Dim Health As Integer
Dim Ammo1 As Integer
Dim Ex As Integer
Dim Ey As Integer
Dim Pts As Integer
Dim sGun As Integer
Dim MaxX As Integer
Dim MaxY As Integer
Dim MinX As Integer
Dim MinY As Integer
Dim RelativeX As Integer
Dim RelativeY As Integer

Public Sub NowJoy()
Dim JoyStickCaps As JOYCAPS
joyGetDevCaps JOYSTICK1, JoyStickCaps, Len(JoyStickCaps)
With JoyStickCaps
    MaxX = .wXmax
    MinX = .wXmin
    MaxY = .wYmax
    MinY = .wYmin
End With
End Sub

Public Sub LoadGame(fname)
On Error GoTo 5
Open fname For Input As #1
Input #1, LevelNo
Input #1, Man.X, Man.Y, Man.Angle
Input #1, Level.Size
Input #1, Health, Ammo1, Pts
For i = 0 To Level.Size - 1
For j = 0 To Level.Size - 1
    Input #1, Level.Tile(i, j)
Next
Next
Close #1
Exit Sub
5
MsgBox "Error loading game!", vbOKOnly + vbExclamation, "Cannon"
End Sub

Public Sub SaveGame(fname)
On Error GoTo 10
Open fname For Output As #1
Print #1, LevelNo
Print #1, Man.X, Man.Y, Man.Angle
Print #1, Level.Size
Print #1, Health, Ammo1, Pts
For i = 0 To Level.Size - 1
For j = 0 To Level.Size - 1
    Print #1, Level.Tile(i, j)
Next
Next
Close #1
Exit Sub
10
MsgBox "Error saving game!", vbOKOnly + vbExclamation, "Cannon"
End Sub

Public Sub Playsound()
Media.PlayStreamingSound Sound.Birds, 1
Sound.Enabled = True
End Sub

Public Sub Stopsound()
Media.StopStreamingSound Sound.Birds
Sound.Enabled = False
End Sub

Public Sub Welcome()
If Sound.Enabled = True Then Media.PlayStaticSound Sound.Welcome, 0
End Sub

Public Sub InitializeSound()
 Sound.Enabled = True
 If Sound.Enabled = True Then
 Dim rt As Long
 Media.Window = GForm.hwnd
 rt = Media.InitializeSound()
 If rt = 1 Then
    Sound.Enabled = True
    Sound.Shoot = Media.CreateStaticSound(App.Path & "\Sounds\Cannon.wav")
    Media.SetStaticVolume Sound.Shoot, 0
    If Sound.Shoot < 0 Then Sound.Enabled = False
    Sound.Shoot1 = Media.CreateStreamingSound(App.Path & "\Sounds\Cannon.wav")
    Media.SetStreamingVolume Sound.Shoot1, 0
    If Sound.Shoot1 < 0 Then Sound.Enabled = False
    Sound.Birds = Media.CreateStreamingSound(App.Path & "\Sounds\birds.wav")
    Media.SetStreamingVolume Sound.Birds, -1200
    If Sound.Birds < 0 Then Sound.Enabled = False
    Sound.Alien = Media.CreateStaticSound(App.Path & "\Sounds\doing.wav")
    Media.SetStaticVolume Sound.Alien, 0
    If Sound.Alien < 0 Then Sound.Enabled = False
    Sound.Kill = Media.CreateStaticSound(App.Path & "\Sounds\ooh.wav")
    Media.SetStaticVolume Sound.Kill, -500
    If Sound.Kill < 0 Then Sound.Enabled = False
    Sound.Welcome = Media.CreateStaticSound(App.Path & "\Sounds\welcome.wav")
    Media.SetStaticVolume Sound.Welcome, 0
    If Sound.Welcome < 0 Then Sound.Enabled = False
    Sound.Hitme = Media.CreateStaticSound(App.Path & "\Sounds\hit.wav")
    Media.SetStaticVolume Sound.Hitme, 0
    If Sound.Hitme < 0 Then Sound.Enabled = False
  Else
    Sound.Enabled = False
  End If
  End If
  If Sound.Enabled = False Then MsgBox "Sound Error !", vbCritical, "Dungeon"
End Sub

Private Sub Aspeak_Timer()
Aspeak.Enabled = False
End Sub

Public Sub Unl()
Unload Form2
GForm.Enabled = True
GForm.Welcome
Show
End Sub

Public Sub Unl1()
Unload Form3
'GForm.Enabled = True
Form2.Show
'GForm.Welcome
'Show
End Sub

Private Sub Form_Click()
'Welldone.Show
'Exi = 1
'Form2.Visible = False
'ChangeScreenSettings 800, 600, 24
'Open "d:\nitin\a.txt" For Output As #1
'For i = 0 To Level.Size * SC
'h$ = ""
'For j = 0 To Level.Size * SC
'h$ = h$ + Str$(Level.nTile(i, j))
'Next
'Print #1, h$
'Next
'Close #1
End Sub

Public Sub DoNitin()
nitin = 1
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeySpace Then
   Ammo1 = Ammo1 - 1
   If Ammo1 >= 0 And Sound.Enabled = True Then Media.PlayStreamingSound Sound.Shoot1, 0
   If Picture1.Visible = True And Picture1.Left + Picture1.Width / 2 > 200 And Picture1.Left + Picture1.Width / 2 < 300 Then
      Level.Tile(Ex, Ey) = NOWT
      Setntile 0, Ex, Ey
      Pts = Pts + 10
      Timer1.Enabled = True
   End If
   If Ammo1 >= 0 Then nitin = 1
End If

Select Case KeyCode
  Case vbKeyS 'go to options screen
    DoOptions
  Case vbKeyEscape
    'Hide 'exit the game
'Select Case iWidth 'put screen res back to normal
'Case 800
'ChangeScreenSettings 800, 600, 16
'Case 1024
'ChangeScreenSettings 1024, 768, 16
'End Select
    Form2.Label8.Caption = "N"
    Form2.Label2.Caption = "Resume play"
    Form2.Label3.Caption = "Load a new game"
    Form2.Label4.Caption = "Save this game"
    Form2.Label4.Enabled = True
    Form2.Show
    Hide
  Case vbKeyC 'capture the screen
    PB.Picture = PB.Image
    SavePicture PB.Picture, App.Path & "\Pic.bmp"
End Select
End Sub

Public Sub DoOptions()
StopFlashing = True
'clear title screen
Picture = LoadPicture()
Cls
'show all the options
'LevelF.Visible = True
'CollisionF.Visible = True
'StyleF.Visible = True
'cmdStart.Visible = True
End Sub

Public Sub NewGame()

Pts = 0
Health = 100
Ammo1 = 100
Dim Path As String
'check wot graphics to load up
'If RealO.Value Then
  Path = App.Path & "\Real\"
'Else
'  Path = App.Path & "\Wierd\"
'End If
'load the chosen pics
For i = 1 To 5
  WallPic(i) = LoadPicture(Path & i & ".bmp", , &H2)
Next
'load level map
'make options dissapear
'LevelF.Visible = False
'CollisionF.Visible = False
'StyleF.Visible = False
'cmdStart.Visible = False
'set collsion detection
'If CollisionO.Value Then Collisions = True
'load the the level
LoadLevel
'enter main loop
MainLoop
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeySpace Then nitin = 1
End Sub

Private Sub Form_Load()
Me.Enabled = False
Me.Picture = LoadPicture()
Randomize Timer
'remember the screen res b4 messing with it
RememberScreenRes
'change to low-res, 16 bit color
ChangeScreenSettings 640, 480, 16
Show
Form2.Show
Form3.Show
If Label5.Caption = "1" Then NowJoy
RelativeX = MaxX / Me.ScaleWidth
RelativeY = MaxY / Me.ScaleHeight
'build TONS of look-up tables
RememberStuff
InitializeSound
'sort out the forms layout
Move 0, 0, RAYSby2 * Screen.TwipsPerPixelX, RAYSby1andHALF * Screen.TwipsPerPixelY
PB.Move 0, 0, RAYSby2, RAYSby1andHALF
LevelNo = 1
'keep flashing a message
'Do
'For i = 1 To 10000
'DoEvents
'Next
'Print " "
'Print "Press 'S' to begin"
'For i = 1 To 10000
'DoEvents
'Next
'Cls
'Loop Until StopFlashing
LevelNo = 1
If Sound.Enabled = True Then Media.PlayStreamingSound Sound.Birds, 1
NewGame
End Sub

Public Sub Dead()
Death.Show
Exi = 1
End Sub

Public Sub NextLevel()
LevelNo = LevelNo + 1
If LevelNo = 5 Then
Welldone.Show
Exi = 1
End If
Label3.Caption = "Level :" + LTrim$(Str$(LevelNo))
Label4.Caption = "Score :0"
Label1.Caption = "100"
Label2.Caption = "100"
Pts = 0
Health = 100
Ammo1 = 100
LoadLevel
End Sub

Public Sub MainLoop()
On Error Resume Next
Dim X As Long
Dim Y As Long
Dim Temp As POINTAPI
Dim RayAngle As Single
Dim ScrX As Integer
Dim StepX As Integer
Dim StepY As Integer
Dim Length As Integer
Dim Hit As Byte
Dim DarkX As Integer
Dim JoyInformation As JOYINFO
Exi = 0
sGun = 0
LoadLevel
nitin = 1
Do
sGun = 0
DoEvents
nk = 0
If Label5.Caption = "1" Then
   joyGetPos JOYSTICK1, JoyInformation
   ax = JoyInformation.wXpos / (RelativeX * 2)
   ay = JoyInformation.wYpos / (RelativeY * 2)
   If ax < 0 Then nk = 1
   If ay > 0 Then nk = 3
   If ax < 0 Then nk = 2
   If ax > 0 Then nk = 4
End If
FrameCount = FrameCount + 1
Caption = FrameCount

PB.Cls
'walk forward
If (GetKeyState(vbKeySpace) And KEY_DOWN) Then
   sGun = 1
End If

If (GetKeyState(vbKeyUp) And KEY_DOWN) Or nk = 1 Then
      h1 = Man.X + Cosine(Man.Angle + ADD90DEGREES) / 10
      h2 = Man.Y - Sine(Man.Angle + ADD90DEGREES) / 10
      'If Level.Tile(h1, h2) = 8 Then Exit Do
      If Level.nTile(h1 * SC, h2 * SC) = 4 Then
        Pts = Pts + 2
        Health = Health + 5
        Level.Tile(h1, h2) = NOWT
        Setntile 0, h1, h2
      End If
      If Level.nTile(h1 * SC, h2 * SC) = 3 Then
        Ammo1 = Ammo1 + 5
        Level.Tile(h1, h2) = NOWT
        Setntile 0, h1, h2
      End If
      If Level.nTile(h1 * SC, h2 * SC) = 5 And Pts >= 50 Then
         NextLevel
      End If
      If Level.nTile(h1 * SC, h2 * SC) = NOWT Or Level.nTile(h1 * SC, h2 * SC) >= 15 Then
        Man.X = h1
        Man.Y = h2
      End If
nitin = 1
End If
'walk backwards
If (GetKeyState(vbKeyDown) And KEY_DOWN) Or nk = 3 Then
      h1 = Man.X - Cosine(Man.Angle + ADD90DEGREES) / 10
      h2 = Man.Y + Sine(Man.Angle + ADD90DEGREES) / 10
      'If Level.Tile(h1, h2) = 8 Then Exit Do
      If Level.nTile(h1 * SC, h2 * SC) = 4 Then
        Pts = Pts + 2
        Health = Health + 5
        Level.Tile(h1, h2) = NOWT
        Setntile 0, h1, h2
      End If
      If Level.nTile(h1 * SC, h2 * SC) = 3 Then
        Ammo1 = Ammo1 + 5
        Level.Tile(h1, h2) = NOWT
        Setntile 0, h1, h2
      End If
      If Level.nTile(h1 * SC, h2 * SC) = 5 And Pts >= 50 Then
         Exit Do
      End If
      If Level.nTile(h1 * SC, h2 * SC) = NOWT Or Level.nTile(h1 * SC, h2 * SC) >= 15 Then
        Man.X = h1
        Man.Y = h2
      End If
nitin = 1
End If
'turn left
If (GetKeyState(vbKeyLeft) And KEY_DOWN) Or nk = 2 Then
    If Man.Angle < 0 Then
      Man.Angle = BACKHALFVIEW
    Else
      Man.Angle = Man.Angle - TURNANGLE
    End If
nitin = 1
End If
'turn right
If (GetKeyState(vbKeyRight) And KEY_DOWN) Or nk = 4 Then
    If Man.Angle > BACKHALFVIEW Then
      Man.Angle = 0
    Else
      Man.Angle = Man.Angle + TURNANGLE
    End If
nitin = 1
End If
'this set the first ray 30 degrees to the left of view
RayAngle = Man.Angle - HALFVIEWRAYS

'loop through all 320 rays, drawing a slither of screen each time
If nitin = 1 Then
If Health > 100 Then Health = 100
If Health <= 0 Then Health = 0: Dead
If Ammo1 > 100 Then Ammo1 = 100
If Ammo1 <= 0 Then Ammo1 = 0
Label1.Caption = Int(Health)
Label2.Caption = Int(Ammo1)
Label4.Caption = "Score :" + LTrim$(Str$(Pts))
ob1 = 0
h = 0
oa1 = 0
st1 = 0
st2 = 0
For ScrX = 0 To RAYS
  X = Man.X * 1200000 'multiply up so that the fixed-point maths is
  Y = Man.Y * 1200000 'accurate enough
  StepX = Sine(RayAngle) / RAYDETAIL * 1200000 'i.e. 1/10th of a unit
  StepY = Cosine(RayAngle) / RAYDETAIL * 1200000
  Length = 0 'length of ray is reset
  ol = 0
  List1.Clear
  List2.Clear
  Do
    X = X - StepX
    Y = Y - StepY 'move ray along a bit
    Length = Length + 1 'increment length
    If oa1 = 0 Then
        If Level.Tile(X \ (1200000), Y \ (1200000)) = 15 Then
           Ex = X \ (1200000)
           Ey = Y \ (1200000)
           oa1 = 1
           h = Dist2Height(Length)
           st1 = ScrX
           st2 = 15
           Health = Health - h / 100
           If Sound.Enabled = True Then Media.PlayStaticSound Sound.Hitme, 0
        End If
        If Level.Tile(X \ (1200000), Y \ (1200000)) = 16 Then
           If Aspeak.Enabled = False And Sound.Enabled = True Then Media.PlayStaticSound Sound.Alien, 0: Aspeak.Enabled = True
           Ex = X \ (1200000)
           Ey = Y \ (1200000)
           oa1 = 1
           h = Dist2Height(Length)
           st1 = ScrX
           st2 = 16
           Health = Health - h / 100
           If Sound.Enabled = True Then Media.PlayStaticSound Sound.Hitme, 0
        End If
    End If
  Hit = Level.Tile(X \ (1200000), Y \ (1200000)) 'see wot's hit
  Loop Until Hit > 0 And Hit < 15 'keep the ray going until a hit is detected
  DarkX = Dist2Dark(Length) 'see how dark the wall should be
  Length = Dist2Height(Length) 'and how tall it looks based on ray length
  If Hit = 3 Or Hit = 4 Then Length = Length / 2
  Temp.X = (X Mod 1200000) \ 12000
  Temp.Y = (Y Mod 1200000) \ 12000 'scale stuff back down again
  
  'here's the clever bit that no-ones ever done b4
  'perspective textures are put onto a wall which
  'was only represented by 1 byte of memory - now that's
  'efficient!!! The drawback is that it's only 90% accurate
  '-this technique gives incorrect results at the sides of walls
  'but for textured 3d walls in VB I think we can forget that
  'since it's hardly noticable anyway
  If Abs(50 - Temp.X) < Abs(50 - Temp.Y) Then
    StretchBlt PB.hdc, ScrX, RAYSby3div8 - Length, 1, Length + Length, WallPic(Hit).hdc, Temp.X + DarkX, 0, 1, 100, vbSrcCopy
  Else
    StretchBlt PB.hdc, ScrX, RAYSby3div8 - Length, 1, Length + Length, WallPic(Hit).hdc, Temp.Y + DarkX, 0, 1, 100, vbSrcCopy
  End If
'fire next ray 1 pixel further along
RayAngle = RayAngle + 1
Next
st1 = st1 * 2
'copy from backbuffer
Picture4.AutoRedraw = True
If Health < 50 Then BitBlt PB.hdc, 0, 0, 320, 240, Picture4.hdc, 0, 0, vbSrcAnd
StretchBlt hdc, 0, 0, RAYSby2, RAYSby1andHALF, PB.hdc, 0, 0, RAYS, RAYSby3div4, vbSrcCopy
If sGun = 0 Then
  BitBlt hdc, 390, 544 - Form1.Gun2.ScaleHeight, Form1.Gun2.ScaleWidth, Form1.Gun2.ScaleHeight, Form1.Gun2Mask.hdc, 0, 0, vbSrcAnd
  BitBlt hdc, 390, 544 - Form1.Gun2.ScaleHeight, Form1.Gun2.ScaleWidth, Form1.Gun2.ScaleHeight, Form1.Gun2.hdc, 0, 0, vbSrcPaint
End If
If sGun = 1 Then
  BitBlt hdc, 390, 544 - Form1.Gun2.ScaleHeight, Form1.Gun2.ScaleWidth, Form1.Gun2.ScaleHeight, Form1.Gun2sm.hdc, 0, 0, vbSrcAnd
  BitBlt hdc, 390, 544 - Form1.Gun2.ScaleHeight, Form1.Gun2.ScaleWidth, Form1.Gun2.ScaleHeight, Form1.Gun2s.hdc, 0, 0, vbSrcPaint
  sGun = 0
End If
DrawWidth = 2
Circle (250, 220), 22
DrawWidth = 1
Circle (250, 220), 10
Line (220, 220)-(280, 220)
Line (250, 190)-(250, 250)
Refresh
Picture1.Visible = False
Picture1.Top = 220 - h / 2
Picture1.Height = h
'MsgBox st1
'Start the first Enemy
If st2 = 15 Then
If st1 < 100 Then
W = (h * Form1.Picture3(st2 - 15).Width) / Form1.Picture3(st2 - 15).Height
Picture1.Width = W
Picture1.Left = st1
BitBlt Picture1.hdc, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, hdc, Picture1.Left, Picture1.Top, vbSrcCopy
StretchBlt Picture1.hdc, 0, 0, W, h, Form1.Picture3(st2 - 15).hdc, 0, 0, Form1.Picture3(st2 - 15).Width - 1, Form1.Picture3(st2 - 15).Height - 1, vbSrcAnd
StretchBlt Picture1.hdc, 0, 0, W, h, Form1.Picture4(st2 - 15).hdc, 0, 0, Form1.Picture3(st2 - 15).Width - 1, Form1.Picture3(st2 - 15).Height - 1, vbSrcPaint
End If
If st1 >= 100 And st1 <= 220 Then
W = (h * Form1.Picture1(st2 - 15).Width) / Form1.Picture1(st2 - 15).Height
Picture1.Width = W
Picture1.Left = st1
BitBlt Picture1.hdc, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, hdc, Picture1.Left, Picture1.Top, vbSrcCopy
StretchBlt Picture1.hdc, 0, 0, W, h, Form1.Picture2(st2 - 15).hdc, 0, 0, Form1.Picture1(st2 - 15).Width - 1, Form1.Picture1(st2 - 15).Height - 1, vbSrcAnd
StretchBlt Picture1.hdc, 0, 0, W, h, Form1.Picture1(st2 - 15).hdc, 0, 0, Form1.Picture1(st2 - 15).Width - 1, Form1.Picture1(st2 - 15).Height - 1, vbSrcPaint
End If
If st1 > 220 Then
W = (h * Form1.Picture5(st2 - 15).Width) / Form1.Picture5(st2 - 15).Height
Picture1.Width = W
Picture1.Left = st1
BitBlt Picture1.hdc, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, hdc, Picture1.Left, Picture1.Top, vbSrcCopy
StretchBlt Picture1.hdc, 0, 0, W, h, Form1.Picture5(st2 - 15).hdc, 0, 0, Form1.Picture5(st2 - 15).Width - 1, Form1.Picture5(st2 - 15).Height - 1, vbSrcAnd
StretchBlt Picture1.hdc, 0, 0, W, h, Form1.Picture6(st2 - 15).hdc, 0, 0, Form1.Picture5(st2 - 15).Width - 1, Form1.Picture5(st2 - 15).Height - 1, vbSrcPaint
End If
'Finish Enemy 1
End If

'Start the second Enemy
If st2 = 16 Then
If st1 < 100 Then
W = (h * Form1.Picture3(st2 - 15).Width) / Form1.Picture3(st2 - 15).Height
Picture1.Width = W
Picture1.Left = st1
BitBlt Picture1.hdc, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, hdc, Picture1.Left, Picture1.Top, vbSrcCopy
StretchBlt Picture1.hdc, 0, 0, W, h, Form1.Picture3(st2 - 15).hdc, 1, 1, Form1.Picture3(st2 - 15).Width - 2, Form1.Picture3(st2 - 15).Height - 2, vbSrcAnd
StretchBlt Picture1.hdc, 0, 0, W, h, Form1.Picture4(st2 - 15).hdc, 1, 1, Form1.Picture3(st2 - 15).Width - 2, Form1.Picture3(st2 - 15).Height - 2, vbSrcPaint
End If
If st1 >= 100 And st1 <= 220 Then
W = (h * Form1.Picture1(st2 - 15).Width) / Form1.Picture1(st2 - 15).Height
Picture1.Width = W
Picture1.Left = st1
BitBlt Picture1.hdc, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, hdc, Picture1.Left, Picture1.Top, vbSrcCopy
StretchBlt Picture1.hdc, 0, 0, W, h, Form1.Picture2(st2 - 15).hdc, 1, 1, Form1.Picture1(st2 - 15).Width - 2, Form1.Picture1(st2 - 15).Height - 2, vbSrcAnd
StretchBlt Picture1.hdc, 0, 0, W, h, Form1.Picture1(st2 - 15).hdc, 1, 1, Form1.Picture1(st2 - 15).Width - 2, Form1.Picture1(st2 - 15).Height - 2, vbSrcPaint
End If
If st1 > 220 Then
W = (h * Form1.Picture5(st2 - 15).Width) / Form1.Picture5(st2 - 15).Height
Picture1.Width = W
Picture1.Left = st1
BitBlt Picture1.hdc, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, hdc, Picture1.Left, Picture1.Top, vbSrcCopy
StretchBlt Picture1.hdc, 0, 0, W, h, Form1.Picture5(st2 - 15).hdc, 1, 1, Form1.Picture5(st2 - 15).Width - 2, Form1.Picture5(st2 - 15).Height - 2, vbSrcAnd
StretchBlt Picture1.hdc, 0, 0, W, h, Form1.Picture6(st2 - 15).hdc, 1, 1, Form1.Picture5(st2 - 15).Width - 2, Form1.Picture5(st2 - 15).Height - 2, vbSrcPaint
End If
'Finish Enemy 2
End If

If st2 < 15 Then
Picture1.Visible = False
BitBlt Picture1.hdc, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, hdc, Picture1.Left, Picture1.Top, vbSrcCopy
Else
Picture1.Visible = True
End If
Picture1.Refresh
'Picture1.Visible = True
End If

nitin = 0
If Exi = 1 Then Exit Do
Loop
If Exi = 1 Then Unload Me: Exit Sub
NextLevel

'level complete
'If MsgBox("Well done, you have completed level " & LevelNo & ". Do want to play again? Click 'Yes' to pick another level and play again, or click 'No' to exit the game.", vbQuestion + vbYesNo, "Level Complete") = vbYes Then
'  Cls
'  DoOptions
'Else
'  Unload Me
'End If

End Sub

Public Sub LoadLevel()
LevelPic = LoadPicture(App.Path & "\Levels\" & LevelNo & ".bmp")
Dim X As Byte
Dim Y As Byte
Man.Angle = 0
'loads a level by transferring bitmap into memory
Level.Size = LevelPic.Width - 1
ReDim Level.Tile(0 To Level.Size, 0 To Level.Size)
ReDim Level.nTile(0 To (Level.Size + 1) * SC, 0 To (Level.Size + 1) * SC)
For X = 0 To Level.Size
For Y = 0 To Level.Size
  ok = 0
  Select Case GetPixel(GForm.LevelPic.hdc, X, Y)
    Case vbBlack
      Level.Tile(X, Y) = NOWT
      Setntile NOWT, X, Y
      ok = 1
    Case vbCyan
      Level.Tile(X, Y) = 1
      Setntile 1, X, Y
      ok = 1
    Case vbYellow
      Level.Tile(X, Y) = 3
      Setntile 3, X, Y
      ok = 1
    Case vbMagenta
      Level.Tile(X, Y) = 4
      Setntile 4, X, Y
      ok = 1
    Case vbBlue
      Level.Tile(X, Y) = 2
      Setntile 2, X, Y
      ok = 1
    Case 128
      Level.Tile(X, Y) = 5
      Setntile 5, X, Y
      ok = 1
    Case vbWhite
      Level.Tile(X, Y) = 16
      Setntile 16, X, Y
      ok = 1
    Case vbGreen
      Level.Tile(X, Y) = NOWT
      Setntile NOWT, X, Y
      Man.X = X
      Man.Y = Y
      ok = 1
    Case vbRed
      Level.Tile(X, Y) = 15
      'DoorPos.X = X
      'DoorPos.Y = Y
      Setntile 15, X, Y
      ok = 1
  End Select
  If ok = 0 Then
'      MsgBox GetPixel(GForm.LevelPic.hdc, X, Y)
      Level.Tile(X, Y) = 0
      Setntile NOWT, X, Y
  End If
Next
Next

End Sub

Public Sub Setntile(W, X, Y)
For i = X * SC To (X + 1) * SC - 1
For j = Y * SC To (Y + 1) * SC - 1
    Level.nTile(i, j) = W
Next
Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
'change screen res back 2 norm
End Sub

Private Sub LevelO_Click(Index As Integer)
LevelNo = Index
End Sub

Private Sub Timer1_Timer()
'If Sound.Enabled = True Then Media.StopStaticSound Sound.Shoot
If Sound.Enabled = True Then Media.PlayStreamingSound Sound.Shoot1, 0
If Sound.Enabled = True Then Media.PlayStaticSound Sound.Kill, 0
Timer1.Enabled = False
End Sub

