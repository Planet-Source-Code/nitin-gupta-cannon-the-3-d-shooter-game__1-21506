VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6045
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10605
   LinkTopic       =   "Form1"
   ScaleHeight     =   403
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   707
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Gun2Mask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3165
      Left            =   1440
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   211
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   145
      TabIndex        =   17
      Top             =   2400
      Width           =   2175
   End
   Begin VB.PictureBox Gun2s 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3165
      Left            =   5400
      Picture         =   "Form1.frx":1679E
      ScaleHeight     =   211
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   16
      Top             =   0
      Width           =   2250
   End
   Begin VB.PictureBox Gun2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3165
      Left            =   3240
      Picture         =   "Form1.frx":2DC6C
      ScaleHeight     =   211
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   145
      TabIndex        =   15
      Top             =   600
      Width           =   2175
   End
   Begin VB.PictureBox Gun2sm 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3165
      Left            =   5520
      Picture         =   "Form1.frx":4440A
      ScaleHeight     =   211
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   14
      Top             =   2400
      Width           =   2250
   End
   Begin VB.PictureBox GunM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1710
      Left            =   2880
      Picture         =   "Form1.frx":5B8D8
      ScaleHeight     =   114
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   211
      TabIndex        =   13
      Top             =   3720
      Width           =   3165
   End
   Begin VB.PictureBox Gun 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1710
      Left            =   840
      Picture         =   "Form1.frx":6D452
      ScaleHeight     =   114
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   211
      TabIndex        =   12
      Top             =   3720
      Width           =   3165
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1185
      Index           =   1
      Left            =   5160
      Picture         =   "Form1.frx":7EFCC
      ScaleHeight     =   79
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   75
      TabIndex        =   11
      Top             =   120
      Width           =   1125
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1185
      Index           =   1
      Left            =   7560
      Picture         =   "Form1.frx":8366A
      ScaleHeight     =   79
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   75
      TabIndex        =   10
      Top             =   120
      Width           =   1125
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1185
      Index           =   1
      Left            =   7560
      Picture         =   "Form1.frx":87D08
      ScaleHeight     =   79
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   75
      TabIndex        =   9
      Top             =   1560
      Width           =   1125
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1185
      Index           =   1
      Left            =   5160
      Picture         =   "Form1.frx":8C3A6
      ScaleHeight     =   79
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   75
      TabIndex        =   8
      Top             =   1560
      Width           =   1125
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1185
      Index           =   1
      Left            =   7560
      Picture         =   "Form1.frx":90A44
      ScaleHeight     =   79
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   75
      TabIndex        =   7
      Top             =   3360
      Width           =   1125
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1185
      Index           =   1
      Left            =   5160
      Picture         =   "Form1.frx":950E2
      ScaleHeight     =   79
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   75
      TabIndex        =   6
      Top             =   2880
      Width           =   1125
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1725
      Index           =   0
      Left            =   120
      Picture         =   "Form1.frx":99780
      ScaleHeight     =   115
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   134
      TabIndex        =   5
      Top             =   3360
      Width           =   2010
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1725
      Index           =   0
      Left            =   2520
      Picture         =   "Form1.frx":A4D3E
      ScaleHeight     =   115
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   134
      TabIndex        =   4
      Top             =   3360
      Width           =   2010
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1725
      Index           =   0
      Left            =   120
      Picture         =   "Form1.frx":B02FC
      ScaleHeight     =   115
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   134
      TabIndex        =   3
      Top             =   1560
      Width           =   2010
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1725
      Index           =   0
      Left            =   2520
      Picture         =   "Form1.frx":BB8BA
      ScaleHeight     =   115
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   134
      TabIndex        =   2
      Top             =   1560
      Width           =   2010
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1290
      Index           =   0
      Left            =   2520
      Picture         =   "Form1.frx":C6E78
      ScaleHeight     =   86
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   153
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1290
      Index           =   0
      Left            =   120
      Picture         =   "Form1.frx":D0942
      ScaleHeight     =   86
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   153
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
