VERSION 5.00
Begin VB.Form frmAbout 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2775
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3960
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmAbout.frx":058A
   ScaleHeight     =   2775
   ScaleWidth      =   3960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblHome 
      BackStyle       =   0  'Transparent
      Caption         =   "- http://www.haliltunc.com"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   240
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   1800
      Width           =   3135
   End
   Begin VB.Image Image1 
      Enabled         =   0   'False
      Height          =   480
      Left            =   360
      Picture         =   "frmAbout.frx":1570
      Top             =   600
      Width           =   480
   End
   Begin VB.Label lblEmail 
      BackStyle       =   0  'Transparent
      Caption         =   "- halil_tunc@hotmail.com"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   1200
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   1020
      Width           =   1935
   End
   Begin VB.Label lblClose 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Close"
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
      Left            =   2880
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   2235
      Width           =   690
   End
   Begin VB.Label lblX 
      BackStyle       =   0  'Transparent
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3675
      MousePointer    =   99  'Custom
      TabIndex        =   0
      ToolTipText     =   "Close"
      Top             =   0
      Width           =   255
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Halil TUNC © 2004
'HtGuitar ...\Form\frmAbout.frm
'19/09/2004 Ankara TURKEY
'halil_tunc@hotmail.com

Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOWNORMAL = 1
Dim iMoveX As Integer, iMoveY As Integer

Private Sub Form_Load()
Dim lRectX As Long, lRectY As Long

   lblX.MouseIcon = LoadResPicture(101, vbResCursor)
   lblClose.MouseIcon = LoadResPicture(101, vbResCursor)
   lblEmail.MouseIcon = LoadResPicture(101, vbResCursor)
   lblHome.MouseIcon = LoadResPicture(101, vbResCursor)
   
   Me.PaintPicture LoadResPicture(106, vbResIcon), 60, 60, 240, 240
   Me.CurrentX = 375
   Me.CurrentY = 60
   Me.Print "HtGutar"
   Me.CurrentX = 1200
   Me.CurrentY = 480
   Me.Print "- HtGuitar 1.0"
   Me.CurrentX = 1200
   Me.CurrentY = 750
   Me.Print "- Halil TUNÇ © 2004"
   Me.Line (240, 1395)-(3695, 1395)
   Me.CurrentX = 240
   Me.CurrentY = 1500
   Me.Print "- Visit page for Visual Basic Source Code "

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = 5
   iMoveX = X
   iMoveY = Y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If iMoveX <> 0 Then Me.Left = Me.Left + X - iMoveX
    If iMoveY <> 0 Then Me.Top = Me.Top + Y - iMoveY
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = 0
   iMoveX = 0
   iMoveY = 0
End Sub

Private Sub lblEmail_Click()
   Call ShellExecute(0&, vbNullString, "mailto:halil_tunc@hotmail.com", vbNullString, "C:\", SW_SHOWNORMAL)
End Sub

Private Sub lblHome_Click()
   Call ShellExecute(0&, vbNullString, "http://www.haliltunc.com/htguitar.asp", vbNullString, "C:\", SW_SHOWNORMAL)
End Sub

Private Sub lblClose_Click()
   Unload Me
End Sub

Private Sub lblX_Click()
   Unload Me
End Sub
