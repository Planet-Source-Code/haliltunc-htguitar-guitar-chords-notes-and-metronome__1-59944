VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H0040DCFF&
   Caption         =   "HtGuitar"
   ClientHeight    =   3390
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5085
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmMain.frx":058A
   ScaleHeight     =   3390
   ScaleWidth      =   5085
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblChord 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "- Chord finder"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000034D0&
      Height          =   345
      Left            =   1680
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label lblAbout 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "- About"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000034D0&
      Height          =   345
      Left            =   1680
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label lblQuit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000034D0&
      Height          =   345
      Left            =   4200
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label lblMetronome 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "- Metronome"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000034D0&
      Height          =   345
      Left            =   1680
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   1440
      Width           =   1515
   End
   Begin VB.Label lblTuner 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "- Tuner"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000034D0&
      Height          =   345
      Left            =   1680
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   2760
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label lblNoteTable 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "- Note table"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000034D0&
      Height          =   345
      Left            =   1680
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   960
      Width           =   1485
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00D8FFFF&
      BackStyle       =   1  'Opaque
      Height          =   2415
      Left            =   1320
      Top             =   240
      Width           =   3495
   End
   Begin VB.Image Image1 
      Height          =   3210
      Left            =   120
      Picture         =   "frmMain.frx":1518
      Top             =   120
      Width           =   1065
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Halil TUNC © 2004
'HtGuitar ...\Form\frmMain.frm
'19/09/2004 Ankara TURKEY
'halil_tunc@hotmail.com

Option Explicit

Private Sub Form_Load()
   Me.Icon = LoadResPicture(106, vbResIcon)
   lblChord.MouseIcon = LoadResPicture(101, vbResCursor)
   lblNoteTable.MouseIcon = LoadResPicture(101, vbResCursor)
   lblMetronome.MouseIcon = LoadResPicture(101, vbResCursor)
   lblQuit.MouseIcon = LoadResPicture(101, vbResCursor)
   lblAbout.MouseIcon = LoadResPicture(101, vbResCursor)
End Sub

Private Sub lblAbout_Click()
   frmAbout.Show vbModal, Me
End Sub

Private Sub lblChord_Click()
    frmChord.Show
End Sub

Private Sub lblMetronome_Click()
   frmMetronome.Show
End Sub

Private Sub lblNoteTable_Click()
    frmNoteTable.Show
End Sub

Private Sub lblQuit_Click()
    Unload Me
End Sub
