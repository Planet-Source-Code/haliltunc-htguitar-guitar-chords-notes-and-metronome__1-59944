VERSION 5.00
Begin VB.Form frmMetronome 
   BackColor       =   &H0048DCFF&
   Caption         =   "Metronome"
   ClientHeight    =   1875
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4485
   Icon            =   "frmMetronome.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   125
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   299
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2500
      Left            =   5160
      Top             =   240
   End
   Begin VB.PictureBox picHold 
      BackColor       =   &H0048DCFF&
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   3480
      Picture         =   "frmMetronome.frx":000C
      ScaleHeight     =   315
      ScaleWidth      =   855
      TabIndex        =   0
      Top             =   240
      Width           =   855
      Begin VB.TextBox txtValue 
         BackColor       =   &H00D8FFFF&
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   60
         TabIndex        =   1
         Text            =   "120"
         Top             =   60
         Width           =   735
      End
   End
   Begin VB.Label lblExp 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   210
      Left            =   225
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   1425
      Width           =   3975
   End
   Begin VB.Label lblClose 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000034D0&
      Height          =   315
      Left            =   3600
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   960
      Width           =   660
   End
   Begin VB.Label lblStartStop 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Start/Stop"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000034D0&
      Height          =   315
      Left            =   2040
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   960
      Width           =   1320
   End
   Begin VB.Image imgSliderCursor 
      Enabled         =   0   'False
      Height          =   480
      Left            =   555
      Picture         =   "frmMetronome.frx":0F16
      Top             =   405
      Width           =   480
   End
   Begin VB.Image imgSlider 
      Height          =   330
      Left            =   120
      MousePointer    =   99  'Custom
      Picture         =   "frmMetronome.frx":17E0
      Top             =   240
      Width           =   3300
   End
   Begin VB.Image imgFilePath 
      Height          =   255
      Left            =   120
      MousePointer    =   99  'Custom
      Picture         =   "frmMetronome.frx":50DA
      Top             =   1440
      Width           =   4215
   End
   Begin VB.Menu mnuSound 
      Caption         =   "Ses"
      Visible         =   0   'False
      Begin VB.Menu mnuSound01 
         Caption         =   "Rhythm 01"
      End
      Begin VB.Menu mnuSound02 
         Caption         =   "Rhythm 02"
      End
      Begin VB.Menu mnuSound03 
         Caption         =   "Rhythm 03"
      End
      Begin VB.Menu mnuSoundS1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSound00 
         Caption         =   "Other"
      End
   End
End
Attribute VB_Name = "frmMetronome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Halil TUNC © 2004
'HtGuitar ...\Form\frmMetronome.frm
'19/09/2004 Ankara TURKEY
'halil_tunc@hotmail.com

Option Explicit

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Declare Function waveOutGetNumDevs Lib "winmm.dll" () As Long
Const SND_ASYNC = &H1
Const SND_NODEFAULT = &H2


Private Const cInterval = 550
Dim pImgCur01 As StdPicture

Dim lValue As Long
Dim sFilePath As String

Dim bRunning As Boolean
Dim bSliderMDown As Boolean

Private Sub Form_Load()
   Me.Icon = LoadResPicture(106, vbResIcon)
   sFilePath = App.Path & "\Rhythm01.wav"
   lblExp.Caption = "Rhythm 01"
   mnuSound01.Checked = True
   pSetValue 120
   txtValue.Text = lValue
   bRunning = False
   bSliderMDown = False
   Set pImgCur01 = LoadResPicture(101, vbResCursor)
   lblStartStop.MouseIcon = pImgCur01
   lblClose.MouseIcon = pImgCur01
   lblExp.MouseIcon = pImgCur01
   imgSlider.MouseIcon = pImgCur01
End Sub

Private Sub imgSlider_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   bSliderMDown = True
   pSetValue fPosToValue(X)
   txtValue.Text = fPosToValue(X)
End Sub

Private Sub imgSlider_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If bSliderMDown Then
      pSetValue fPosToValue(X)
      txtValue.Text = fPosToValue(X)
   End If
End Sub

Private Sub imgSlider_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   bSliderMDown = False
End Sub

Private Sub lblClose_Click()
   Unload Me
End Sub

Private Sub lblExp_Click()
   Me.PopupMenu mnuSound
End Sub

Private Sub lblStartStop_Click()
   bRunning = Not bRunning
   lblStartStop.ForeColor = IIf(bRunning, &H8000&, &H34D0&)
   Timer1.Enabled = bRunning
End Sub

Private Sub mnuSound00_Click()
   fSetFilePath
End Sub

Private Sub mnuSound01_Click()
   lblExp.Caption = "Rhythm 01"
   mnuSound00.Checked = False
   mnuSound01.Checked = True
   mnuSound02.Checked = False
   mnuSound03.Checked = False
   pSetSoundFile App.Path & "\Rhythm01.wav"
End Sub

Private Sub mnuSound02_Click()
   lblExp.Caption = "Rhythm 02"
   mnuSound00.Checked = False
   mnuSound01.Checked = False
   mnuSound02.Checked = True
   mnuSound03.Checked = False
   pSetSoundFile App.Path & "\Rhythm02.wav"
End Sub

Private Sub mnuSound03_Click()
   lblExp.Caption = "Rhythm 03"
   mnuSound00.Checked = False
   mnuSound01.Checked = False
   mnuSound02.Checked = False
   mnuSound03.Checked = True
   pSetSoundFile App.Path & "\Rhythm03.wav"
End Sub

Private Sub Timer1_Timer()
   pPlaySound sFilePath
End Sub

Private Sub txtValue_Change()
   pSetValue Val(txtValue.Text)
End Sub

Private Sub pPlaySound(sFile As String)
   Dim ret&
   ret& = waveOutGetNumDevs
   If ret& > 0 Then
      If sndPlaySound(sFile, SND_ASYNC Or SND_NODEFAULT) = 0 Then
         lblExp.Caption = "Error..!"
         bRunning = False
         lblStartStop.ForeColor = &H34D0&
         Timer1.Enabled = False
      End If
   End If
End Sub

Private Sub pSetSoundFile(sText As String)
   pPlaySound sText
   sFilePath = sText
End Sub

Private Sub pSetValue(New_Value As Long)
   If New_Value > 0 And New_Value < 400 Then
      lValue = New_Value
      Timer1.Interval = cInterval / New_Value * 100
      If lValue < 40 Then
         imgSliderCursor.Left = imgSlider.Left + 10 - 9
      ElseIf lValue > 240 Then
         imgSliderCursor.Left = imgSlider.Left + 240 - 30 - 9
      Else
         imgSliderCursor.Left = imgSlider.Left + lValue - 30 - 9
      End If
   End If
End Sub

Private Function fPosToValue(lX As Single) As Long
Dim i As Long
   i = ScaleX(lX, vbTwips, vbPixels) + 30
   If i < 40 Then
      fPosToValue = 40
   ElseIf i > 240 Then
      fPosToValue = 240
   Else
      fPosToValue = i
   End If
End Function

Private Function fSetFilePath(Optional sText As String) As String
    Dim tOpenFileName As OPENFILENAME
    tOpenFileName.lStructSize = Len(tOpenFileName)
    tOpenFileName.hwndOwner = Me.hwnd
    tOpenFileName.hInstance = App.hInstance
    tOpenFileName.lpstrFilter = "Wave files (*.wav)" + Chr$(0) + "*.wav" + Chr$(0) + "All files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
    tOpenFileName.lpstrFile = Space$(254)
    tOpenFileName.nMaxFile = 255
    tOpenFileName.lpstrFileTitle = Space$(254)
    tOpenFileName.nMaxFileTitle = 255
    tOpenFileName.lpstrInitialDir = IIf(sText = "", App.Path, sText)
    tOpenFileName.lpstrTitle = "Select sound file"
    tOpenFileName.flags = 0

    If GetOpenFileName(tOpenFileName) Then
        lblExp.Caption = "Other rhythms"
         mnuSound00.Checked = True
         mnuSound01.Checked = False
         mnuSound02.Checked = False
         mnuSound03.Checked = False
        pSetSoundFile Trim$(tOpenFileName.lpstrFile)
    Else
        '
    End If
End Function
