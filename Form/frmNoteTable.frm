VERSION 5.00
Begin VB.Form frmNoteTable 
   BackColor       =   &H00D0FCFF&
   Caption         =   "Note table"
   ClientHeight    =   5535
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9045
   Icon            =   "frmNoteTable.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   9045
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picC 
      BorderStyle     =   0  'None
      Height          =   2820
      Left            =   240
      Picture         =   "frmNoteTable.frx":000C
      ScaleHeight     =   2820
      ScaleWidth      =   3975
      TabIndex        =   27
      Top             =   2400
      Width           =   3975
      Begin VB.PictureBox picStave 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2250
         Left            =   120
         MouseIcon       =   "frmNoteTable.frx":0FF2
         MousePointer    =   99  'Custom
         Picture         =   "frmNoteTable.frx":1144
         ScaleHeight     =   2250
         ScaleWidth      =   3630
         TabIndex        =   28
         ToolTipText     =   "Right click for note type"
         Top             =   360
         Width           =   3630
         Begin VB.Image ImgNote 
            Enabled         =   0   'False
            Height          =   480
            Left            =   2040
            Top             =   1560
            Width           =   480
         End
         Begin VB.Image imgNoteType 
            Enabled         =   0   'False
            Height          =   375
            Left            =   1800
            Top             =   1320
            Width           =   375
         End
      End
      Begin VB.Image imgFlat 
         Height          =   375
         Left            =   3705
         MousePointer    =   99  'Custom
         ToolTipText     =   "Flat"
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgSharp 
         Height          =   315
         Left            =   3480
         MousePointer    =   99  'Custom
         ToolTipText     =   "Sharp"
         Top             =   0
         Width           =   180
      End
      Begin VB.Label lblNoteName 
         BackColor       =   &H00F0E8D8&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   29
         Top             =   0
         Width           =   3255
      End
   End
   Begin VB.Image imgRB 
      Height          =   330
      Left            =   12480
      Top             =   2400
      Width           =   300
   End
   Begin VB.Image Es1 
      Height          =   330
      Left            =   12480
      Top             =   2040
      Width           =   300
   End
   Begin VB.Image A1 
      Height          =   330
      Left            =   12480
      Top             =   1680
      Width           =   300
   End
   Begin VB.Image D1 
      Height          =   330
      Left            =   12480
      Top             =   1320
      Width           =   300
   End
   Begin VB.Image G1 
      Height          =   330
      Left            =   12480
      Top             =   960
      Width           =   300
   End
   Begin VB.Image B1 
      Height          =   330
      Left            =   12480
      Top             =   600
      Width           =   300
   End
   Begin VB.Image E1 
      Height          =   330
      Left            =   12480
      Top             =   240
      Width           =   300
   End
   Begin VB.Image imgRT 
      Height          =   330
      Left            =   12480
      Top             =   0
      Width           =   300
   End
   Begin VB.Label lblEs 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000034D0&
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label lblTop 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "O"
      ForeColor       =   &H000034D0&
      Height          =   255
      Index           =   19
      Left            =   11880
      TabIndex        =   26
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label lblTop 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "O"
      ForeColor       =   &H000034D0&
      Height          =   255
      Index           =   18
      Left            =   11280
      TabIndex        =   25
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label lblTop 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "O"
      ForeColor       =   &H000034D0&
      Height          =   255
      Index           =   17
      Left            =   10680
      TabIndex        =   24
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label lblTop 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "O"
      ForeColor       =   &H000034D0&
      Height          =   255
      Index           =   16
      Left            =   10080
      TabIndex        =   23
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label lblTop 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000034D0&
      Height          =   255
      Index           =   15
      Left            =   9480
      TabIndex        =   22
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label lblTop 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "O"
      ForeColor       =   &H000034D0&
      Height          =   255
      Index           =   14
      Left            =   8880
      TabIndex        =   21
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label lblTop 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "O"
      ForeColor       =   &H000034D0&
      Height          =   255
      Index           =   13
      Left            =   8280
      TabIndex        =   20
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label lblTop 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000034D0&
      Height          =   255
      Index           =   12
      Left            =   7680
      TabIndex        =   19
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label lblTop 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "O"
      ForeColor       =   &H000034D0&
      Height          =   255
      Index           =   11
      Left            =   7080
      TabIndex        =   18
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label lblTop 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "O"
      ForeColor       =   &H000034D0&
      Height          =   255
      Index           =   10
      Left            =   6480
      TabIndex        =   17
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label lblTop 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000034D0&
      Height          =   255
      Index           =   9
      Left            =   5880
      TabIndex        =   16
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label lblTop 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "O"
      ForeColor       =   &H000034D0&
      Height          =   255
      Index           =   8
      Left            =   5280
      TabIndex        =   15
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label lblTop 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000034D0&
      Height          =   255
      Index           =   7
      Left            =   4680
      TabIndex        =   14
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label lblTop 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "O"
      ForeColor       =   &H000034D0&
      Height          =   255
      Index           =   6
      Left            =   4080
      TabIndex        =   13
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label lblTop 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000034D0&
      Height          =   255
      Index           =   5
      Left            =   3480
      TabIndex        =   12
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label lblTop 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "O"
      ForeColor       =   &H000034D0&
      Height          =   255
      Index           =   4
      Left            =   2880
      TabIndex        =   11
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label lblTop 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "O"
      ForeColor       =   &H000034D0&
      Height          =   255
      Index           =   3
      Left            =   2280
      TabIndex        =   10
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label lblTop 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "O"
      ForeColor       =   &H000034D0&
      Height          =   255
      Index           =   2
      Left            =   1680
      TabIndex        =   9
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label lblTop 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "O"
      ForeColor       =   &H000034D0&
      Height          =   255
      Index           =   1
      Left            =   1080
      TabIndex        =   8
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label lblTop 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "O"
      ForeColor       =   &H000034D0&
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   7
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label lblA 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000034D0&
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label lblD 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000034D0&
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label lblG 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "G"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000034D0&
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   960
      Width           =   375
   End
   Begin VB.Label lblB 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000034D0&
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   375
   End
   Begin VB.Label lblE 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000034D0&
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   240
      Width           =   375
   End
   Begin VB.Image imgLB 
      Height          =   330
      Left            =   120
      Tag             =   "15, ,Mi"
      Top             =   2400
      Width           =   300
   End
   Begin VB.Image imgLT 
      Height          =   330
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Image imgBottom 
      Height          =   285
      Index           =   19
      Left            =   11880
      MousePointer    =   99  'Custom
      Top             =   2400
      Width           =   555
   End
   Begin VB.Image imgBottom 
      Height          =   285
      Index           =   18
      Left            =   11280
      MousePointer    =   99  'Custom
      Top             =   2400
      Width           =   555
   End
   Begin VB.Image imgBottom 
      Height          =   285
      Index           =   17
      Left            =   10680
      MousePointer    =   99  'Custom
      Top             =   2400
      Width           =   555
   End
   Begin VB.Image imgBottom 
      Height          =   285
      Index           =   16
      Left            =   10080
      MousePointer    =   99  'Custom
      Top             =   2400
      Width           =   555
   End
   Begin VB.Image imgBottom 
      Height          =   285
      Index           =   15
      Left            =   9480
      MousePointer    =   99  'Custom
      Top             =   2400
      Width           =   555
   End
   Begin VB.Image imgBottom 
      Height          =   285
      Index           =   14
      Left            =   8880
      MousePointer    =   99  'Custom
      Top             =   2400
      Width           =   555
   End
   Begin VB.Image imgBottom 
      Height          =   285
      Index           =   13
      Left            =   8280
      MousePointer    =   99  'Custom
      Top             =   2400
      Width           =   555
   End
   Begin VB.Image imgBottom 
      Height          =   285
      Index           =   12
      Left            =   7680
      MousePointer    =   99  'Custom
      Top             =   2400
      Width           =   555
   End
   Begin VB.Image imgBottom 
      Height          =   285
      Index           =   11
      Left            =   7080
      MousePointer    =   99  'Custom
      Top             =   2400
      Width           =   555
   End
   Begin VB.Image imgBottom 
      Height          =   285
      Index           =   10
      Left            =   6480
      MousePointer    =   99  'Custom
      Top             =   2400
      Width           =   555
   End
   Begin VB.Image imgBottom 
      Height          =   285
      Index           =   9
      Left            =   5880
      MousePointer    =   99  'Custom
      Top             =   2400
      Width           =   555
   End
   Begin VB.Image imgBottom 
      Height          =   285
      Index           =   8
      Left            =   5280
      MousePointer    =   99  'Custom
      Top             =   2400
      Width           =   555
   End
   Begin VB.Image imgBottom 
      Height          =   285
      Index           =   7
      Left            =   4680
      MousePointer    =   99  'Custom
      Top             =   2400
      Width           =   555
   End
   Begin VB.Image imgBottom 
      Height          =   285
      Index           =   6
      Left            =   4080
      MousePointer    =   99  'Custom
      Top             =   2400
      Width           =   555
   End
   Begin VB.Image imgBottom 
      Height          =   285
      Index           =   5
      Left            =   3480
      MousePointer    =   99  'Custom
      Top             =   2400
      Width           =   555
   End
   Begin VB.Image imgBottom 
      Height          =   285
      Index           =   4
      Left            =   2880
      MousePointer    =   99  'Custom
      Top             =   2400
      Width           =   555
   End
   Begin VB.Image imgBottom 
      Height          =   285
      Index           =   3
      Left            =   2280
      MousePointer    =   99  'Custom
      Top             =   2400
      Width           =   555
   End
   Begin VB.Image imgBottom 
      Height          =   285
      Index           =   2
      Left            =   1680
      MousePointer    =   99  'Custom
      Top             =   2400
      Width           =   555
   End
   Begin VB.Image imgBottom 
      Height          =   285
      Index           =   1
      Left            =   1080
      MousePointer    =   99  'Custom
      Top             =   2400
      Width           =   555
   End
   Begin VB.Image imgBottom 
      Height          =   285
      Index           =   0
      Left            =   480
      MousePointer    =   99  'Custom
      Top             =   2400
      Width           =   555
   End
   Begin VB.Image imgTop 
      Height          =   285
      Index           =   19
      Left            =   11880
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   555
   End
   Begin VB.Image imgTop 
      Height          =   285
      Index           =   18
      Left            =   11280
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   555
   End
   Begin VB.Image imgTop 
      Height          =   285
      Index           =   17
      Left            =   10680
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   555
   End
   Begin VB.Image imgTop 
      Height          =   285
      Index           =   16
      Left            =   10080
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   555
   End
   Begin VB.Image imgTop 
      Height          =   285
      Index           =   15
      Left            =   9480
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   555
   End
   Begin VB.Image imgTop 
      Height          =   285
      Index           =   14
      Left            =   8880
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   555
   End
   Begin VB.Image imgTop 
      Height          =   285
      Index           =   13
      Left            =   8280
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   555
   End
   Begin VB.Image imgTop 
      Height          =   285
      Index           =   12
      Left            =   7680
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   555
   End
   Begin VB.Image imgTop 
      Height          =   285
      Index           =   11
      Left            =   7080
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   555
   End
   Begin VB.Image imgTop 
      Height          =   285
      Index           =   10
      Left            =   6480
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   555
   End
   Begin VB.Image imgTop 
      Height          =   285
      Index           =   9
      Left            =   5880
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   555
   End
   Begin VB.Image imgTop 
      Height          =   285
      Index           =   8
      Left            =   5280
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   555
   End
   Begin VB.Image imgTop 
      Height          =   285
      Index           =   7
      Left            =   4680
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   555
   End
   Begin VB.Image imgTop 
      Height          =   285
      Index           =   6
      Left            =   4080
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   555
   End
   Begin VB.Image imgTop 
      Height          =   285
      Index           =   5
      Left            =   3480
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   555
   End
   Begin VB.Image imgTop 
      Height          =   285
      Index           =   4
      Left            =   2880
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   555
   End
   Begin VB.Image imgTop 
      Height          =   285
      Index           =   3
      Left            =   2280
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   555
   End
   Begin VB.Image imgTop 
      Height          =   285
      Index           =   2
      Left            =   1680
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   555
   End
   Begin VB.Image imgTop 
      Height          =   285
      Index           =   1
      Left            =   1080
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   585
   End
   Begin VB.Image imgTop 
      Height          =   285
      Index           =   0
      Left            =   480
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   585
   End
   Begin VB.Image B 
      Height          =   330
      Index           =   19
      Left            =   11880
      MousePointer    =   99  'Custom
      Tag             =   "23,#,F"
      Top             =   600
      Width           =   555
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
      Left            =   7800
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   4800
      Width           =   690
   End
   Begin VB.Image E 
      Height          =   330
      Index           =   1
      Left            =   1080
      MousePointer    =   99  'Custom
      Tag             =   "16, ,F"
      Top             =   240
      Width           =   555
   End
   Begin VB.Image E 
      Height          =   330
      Index           =   2
      Left            =   1680
      MousePointer    =   99  'Custom
      Tag             =   "16,#,F"
      Top             =   240
      Width           =   555
   End
   Begin VB.Image Es0 
      Height          =   330
      Left            =   120
      Top             =   2040
      Width           =   300
   End
   Begin VB.Image A0 
      Height          =   330
      Left            =   120
      Top             =   1680
      Width           =   300
   End
   Begin VB.Image D0 
      Height          =   330
      Left            =   120
      Top             =   1320
      Width           =   300
   End
   Begin VB.Image G0 
      Height          =   330
      Left            =   120
      Tag             =   "10, ,G"
      Top             =   960
      Width           =   300
   End
   Begin VB.Image B0 
      Height          =   330
      Left            =   120
      Top             =   600
      Width           =   300
   End
   Begin VB.Image E0 
      Height          =   330
      Left            =   120
      Top             =   240
      Width           =   300
   End
   Begin VB.Image E 
      Height          =   330
      Index           =   0
      Left            =   480
      MousePointer    =   99  'Custom
      Tag             =   "15, ,E"
      Top             =   240
      Width           =   555
   End
   Begin VB.Image B 
      Height          =   330
      Index           =   0
      Left            =   480
      MousePointer    =   99  'Custom
      Tag             =   "12, ,B"
      Top             =   600
      Width           =   555
   End
   Begin VB.Image G 
      Height          =   330
      Index           =   0
      Left            =   480
      MousePointer    =   99  'Custom
      Tag             =   "10, ,G"
      Top             =   960
      Width           =   555
   End
   Begin VB.Image D 
      Height          =   330
      Index           =   0
      Left            =   480
      MousePointer    =   99  'Custom
      Tag             =   "07, ,D"
      Top             =   1320
      Width           =   555
   End
   Begin VB.Image A 
      Height          =   330
      Index           =   0
      Left            =   480
      MousePointer    =   99  'Custom
      Tag             =   "04, ,A"
      Top             =   1680
      Width           =   555
   End
   Begin VB.Image Es 
      Height          =   330
      Index           =   0
      Left            =   480
      MousePointer    =   99  'Custom
      Tag             =   "01, ,E"
      Top             =   2040
      Width           =   555
   End
   Begin VB.Image Es 
      Height          =   330
      Index           =   19
      Left            =   11880
      MousePointer    =   99  'Custom
      Tag             =   "12, ,B"
      Top             =   2040
      Width           =   555
   End
   Begin VB.Image A 
      Height          =   330
      Index           =   19
      Left            =   11880
      MousePointer    =   99  'Custom
      Tag             =   "E"
      Top             =   1680
      Width           =   555
   End
   Begin VB.Image D 
      Height          =   330
      Index           =   19
      Left            =   11880
      MousePointer    =   99  'Custom
      Tag             =   "18, ,A"
      Top             =   1320
      Width           =   555
   End
   Begin VB.Image G 
      Height          =   330
      Index           =   19
      Left            =   11880
      MousePointer    =   99  'Custom
      Tag             =   "21, ,D"
      Top             =   960
      Width           =   555
   End
   Begin VB.Image E 
      Height          =   330
      Index           =   19
      Left            =   11880
      MousePointer    =   99  'Custom
      Tag             =   "26, ,B"
      Top             =   240
      Width           =   555
   End
   Begin VB.Image Es 
      Height          =   330
      Index           =   18
      Left            =   11280
      MousePointer    =   99  'Custom
      Tag             =   "11,#,A"
      Top             =   2040
      Width           =   555
   End
   Begin VB.Image A 
      Height          =   330
      Index           =   18
      Left            =   11280
      MousePointer    =   99  'Custom
      Tag             =   "14,#,D"
      Top             =   1680
      Width           =   555
   End
   Begin VB.Image D 
      Height          =   330
      Index           =   18
      Left            =   11280
      MousePointer    =   99  'Custom
      Tag             =   "17,#,G"
      Top             =   1320
      Width           =   555
   End
   Begin VB.Image G 
      Height          =   330
      Index           =   18
      Left            =   11280
      MousePointer    =   99  'Custom
      Tag             =   "20,#,C"
      Top             =   960
      Width           =   555
   End
   Begin VB.Image B 
      Height          =   330
      Index           =   18
      Left            =   11280
      MousePointer    =   99  'Custom
      Tag             =   "23, ,F"
      Top             =   600
      Width           =   555
   End
   Begin VB.Image E 
      Height          =   330
      Index           =   18
      Left            =   11280
      MousePointer    =   99  'Custom
      Tag             =   "25,#,A"
      Top             =   240
      Width           =   555
   End
   Begin VB.Image Es 
      Height          =   330
      Index           =   17
      Left            =   10680
      MousePointer    =   99  'Custom
      Tag             =   "11, ,A"
      Top             =   2040
      Width           =   555
   End
   Begin VB.Image A 
      Height          =   330
      Index           =   17
      Left            =   10680
      MousePointer    =   99  'Custom
      Tag             =   "14, ,D"
      Top             =   1680
      Width           =   555
   End
   Begin VB.Image D 
      Height          =   330
      Index           =   17
      Left            =   10680
      MousePointer    =   99  'Custom
      Tag             =   "17, ,G"
      Top             =   1320
      Width           =   555
   End
   Begin VB.Image G 
      Height          =   330
      Index           =   17
      Left            =   10680
      MousePointer    =   99  'Custom
      Tag             =   "20, ,C"
      Top             =   960
      Width           =   555
   End
   Begin VB.Image B 
      Height          =   330
      Index           =   17
      Left            =   10680
      MousePointer    =   99  'Custom
      Tag             =   "22, ,E"
      Top             =   600
      Width           =   555
   End
   Begin VB.Image E 
      Height          =   330
      Index           =   17
      Left            =   10680
      MousePointer    =   99  'Custom
      Tag             =   "25, ,A"
      Top             =   240
      Width           =   555
   End
   Begin VB.Image Es 
      Height          =   330
      Index           =   16
      Left            =   10080
      MousePointer    =   99  'Custom
      Tag             =   "10,#,G"
      Top             =   2040
      Width           =   555
   End
   Begin VB.Image A 
      Height          =   330
      Index           =   16
      Left            =   10080
      MousePointer    =   99  'Custom
      Tag             =   "13,#,C"
      Top             =   1680
      Width           =   555
   End
   Begin VB.Image D 
      Height          =   330
      Index           =   16
      Left            =   10080
      MousePointer    =   99  'Custom
      Tag             =   "16,#,F"
      Top             =   1320
      Width           =   555
   End
   Begin VB.Image G 
      Height          =   330
      Index           =   16
      Left            =   10080
      MousePointer    =   99  'Custom
      Tag             =   "19, ,B"
      Top             =   960
      Width           =   555
   End
   Begin VB.Image B 
      Height          =   330
      Index           =   16
      Left            =   10080
      MousePointer    =   99  'Custom
      Tag             =   "21,#,D"
      Top             =   600
      Width           =   555
   End
   Begin VB.Image E 
      Height          =   330
      Index           =   16
      Left            =   10080
      MousePointer    =   99  'Custom
      Tag             =   "24,#,G"
      Top             =   240
      Width           =   555
   End
   Begin VB.Image Es 
      Height          =   330
      Index           =   15
      Left            =   9480
      MousePointer    =   99  'Custom
      Tag             =   "10, ,G"
      Top             =   2040
      Width           =   555
   End
   Begin VB.Image A 
      Height          =   330
      Index           =   15
      Left            =   9480
      MousePointer    =   99  'Custom
      Tag             =   "13, ,C"
      Top             =   1680
      Width           =   555
   End
   Begin VB.Image D 
      Height          =   330
      Index           =   15
      Left            =   9480
      MousePointer    =   99  'Custom
      Tag             =   "16, ,F"
      Top             =   1320
      Width           =   555
   End
   Begin VB.Image G 
      Height          =   330
      Index           =   15
      Left            =   9480
      MousePointer    =   99  'Custom
      Tag             =   "18,#,A"
      Top             =   960
      Width           =   555
   End
   Begin VB.Image B 
      Height          =   330
      Index           =   15
      Left            =   9480
      MousePointer    =   99  'Custom
      Tag             =   "21, ,D"
      Top             =   600
      Width           =   555
   End
   Begin VB.Image E 
      Height          =   330
      Index           =   15
      Left            =   9480
      MousePointer    =   99  'Custom
      Tag             =   "24, ,G"
      Top             =   240
      Width           =   555
   End
   Begin VB.Image Es 
      Height          =   330
      Index           =   14
      Left            =   8880
      MousePointer    =   99  'Custom
      Tag             =   "09,#,F"
      Top             =   2040
      Width           =   555
   End
   Begin VB.Image A 
      Height          =   330
      Index           =   14
      Left            =   8880
      MousePointer    =   99  'Custom
      Tag             =   "12, ,B"
      Top             =   1680
      Width           =   555
   End
   Begin VB.Image D 
      Height          =   330
      Index           =   14
      Left            =   8880
      MousePointer    =   99  'Custom
      Tag             =   "15, ,E"
      Top             =   1320
      Width           =   555
   End
   Begin VB.Image G 
      Height          =   330
      Index           =   14
      Left            =   8880
      MousePointer    =   99  'Custom
      Tag             =   "18, ,A"
      Top             =   960
      Width           =   555
   End
   Begin VB.Image B 
      Height          =   330
      Index           =   14
      Left            =   8880
      MousePointer    =   99  'Custom
      Tag             =   "20,#,C"
      Top             =   600
      Width           =   555
   End
   Begin VB.Image E 
      Height          =   330
      Index           =   14
      Left            =   8880
      MousePointer    =   99  'Custom
      Tag             =   "23,#,F"
      Top             =   240
      Width           =   555
   End
   Begin VB.Image Es 
      Height          =   330
      Index           =   13
      Left            =   8280
      MousePointer    =   99  'Custom
      Tag             =   "09, ,F"
      Top             =   2040
      Width           =   555
   End
   Begin VB.Image A 
      Height          =   330
      Index           =   13
      Left            =   8280
      MousePointer    =   99  'Custom
      Tag             =   "11,#,A"
      Top             =   1680
      Width           =   555
   End
   Begin VB.Image D 
      Height          =   330
      Index           =   13
      Left            =   8280
      MousePointer    =   99  'Custom
      Tag             =   "14,#,D"
      Top             =   1320
      Width           =   555
   End
   Begin VB.Image G 
      Height          =   330
      Index           =   13
      Left            =   8280
      MousePointer    =   99  'Custom
      Tag             =   "17,#,G"
      Top             =   960
      Width           =   555
   End
   Begin VB.Image B 
      Height          =   330
      Index           =   13
      Left            =   8280
      MousePointer    =   99  'Custom
      Tag             =   "20, ,C"
      Top             =   600
      Width           =   555
   End
   Begin VB.Image E 
      Height          =   330
      Index           =   13
      Left            =   8280
      MousePointer    =   99  'Custom
      Tag             =   "23, ,F"
      Top             =   240
      Width           =   555
   End
   Begin VB.Image Es 
      Height          =   330
      Index           =   12
      Left            =   7680
      MousePointer    =   99  'Custom
      Tag             =   "08, ,E"
      Top             =   2040
      Width           =   555
   End
   Begin VB.Image A 
      Height          =   330
      Index           =   12
      Left            =   7680
      MousePointer    =   99  'Custom
      Tag             =   "11, ,A"
      Top             =   1680
      Width           =   555
   End
   Begin VB.Image D 
      Height          =   330
      Index           =   12
      Left            =   7680
      MousePointer    =   99  'Custom
      Tag             =   "14, ,D"
      Top             =   1320
      Width           =   555
   End
   Begin VB.Image G 
      Height          =   330
      Index           =   12
      Left            =   7680
      MousePointer    =   99  'Custom
      Tag             =   "17, ,G"
      Top             =   960
      Width           =   555
   End
   Begin VB.Image B 
      Height          =   330
      Index           =   12
      Left            =   7680
      MousePointer    =   99  'Custom
      Tag             =   "19, ,B"
      Top             =   600
      Width           =   555
   End
   Begin VB.Image E 
      Height          =   330
      Index           =   12
      Left            =   7680
      MousePointer    =   99  'Custom
      Tag             =   "22, ,E"
      Top             =   240
      Width           =   555
   End
   Begin VB.Image Es 
      Height          =   330
      Index           =   11
      Left            =   7080
      MousePointer    =   99  'Custom
      Tag             =   "07,#,D"
      Top             =   2040
      Width           =   555
   End
   Begin VB.Image A 
      Height          =   330
      Index           =   11
      Left            =   7080
      MousePointer    =   99  'Custom
      Tag             =   "10,#,G"
      Top             =   1680
      Width           =   555
   End
   Begin VB.Image D 
      Height          =   330
      Index           =   11
      Left            =   7080
      MousePointer    =   99  'Custom
      Tag             =   "13,#,C"
      Top             =   1320
      Width           =   555
   End
   Begin VB.Image G 
      Height          =   330
      Index           =   11
      Left            =   7080
      MousePointer    =   99  'Custom
      Tag             =   "16,#,F"
      Top             =   960
      Width           =   555
   End
   Begin VB.Image B 
      Height          =   330
      Index           =   11
      Left            =   7080
      MousePointer    =   99  'Custom
      Tag             =   "18,#,A"
      Top             =   600
      Width           =   555
   End
   Begin VB.Image E 
      Height          =   330
      Index           =   11
      Left            =   7080
      MousePointer    =   99  'Custom
      Tag             =   "21,#,D"
      Top             =   240
      Width           =   555
   End
   Begin VB.Image Es 
      Height          =   330
      Index           =   10
      Left            =   6480
      MousePointer    =   99  'Custom
      Tag             =   "07, ,D"
      Top             =   2040
      Width           =   555
   End
   Begin VB.Image A 
      Height          =   330
      Index           =   10
      Left            =   6480
      MousePointer    =   99  'Custom
      Tag             =   "10, ,G"
      Top             =   1680
      Width           =   555
   End
   Begin VB.Image D 
      Height          =   330
      Index           =   10
      Left            =   6480
      MousePointer    =   99  'Custom
      Tag             =   "13, ,C"
      Top             =   1320
      Width           =   555
   End
   Begin VB.Image G 
      Height          =   330
      Index           =   10
      Left            =   6480
      MousePointer    =   99  'Custom
      Tag             =   "16, ,F"
      Top             =   960
      Width           =   555
   End
   Begin VB.Image B 
      Height          =   330
      Index           =   10
      Left            =   6480
      MousePointer    =   99  'Custom
      Tag             =   "18, ,A"
      Top             =   600
      Width           =   555
   End
   Begin VB.Image E 
      Height          =   330
      Index           =   10
      Left            =   6480
      MousePointer    =   99  'Custom
      Tag             =   "21, ,D"
      Top             =   240
      Width           =   555
   End
   Begin VB.Image Es 
      Height          =   330
      Index           =   9
      Left            =   5880
      MousePointer    =   99  'Custom
      Tag             =   "06,#,C"
      Top             =   2040
      Width           =   555
   End
   Begin VB.Image A 
      Height          =   330
      Index           =   9
      Left            =   5880
      MousePointer    =   99  'Custom
      Tag             =   "09,#,F"
      Top             =   1680
      Width           =   555
   End
   Begin VB.Image D 
      Height          =   330
      Index           =   9
      Left            =   5880
      MousePointer    =   99  'Custom
      Tag             =   "12, ,B"
      Top             =   1320
      Width           =   555
   End
   Begin VB.Image G 
      Height          =   330
      Index           =   9
      Left            =   5880
      MousePointer    =   99  'Custom
      Tag             =   "15, ,E"
      Top             =   960
      Width           =   555
   End
   Begin VB.Image B 
      Height          =   330
      Index           =   9
      Left            =   5880
      MousePointer    =   99  'Custom
      Tag             =   "17,#,G"
      Top             =   600
      Width           =   555
   End
   Begin VB.Image E 
      Height          =   330
      Index           =   9
      Left            =   5880
      MousePointer    =   99  'Custom
      Tag             =   "20,#,C"
      Top             =   240
      Width           =   555
   End
   Begin VB.Image Es 
      Height          =   330
      Index           =   8
      Left            =   5280
      MousePointer    =   99  'Custom
      Tag             =   "06, ,C"
      Top             =   2040
      Width           =   555
   End
   Begin VB.Image A 
      Height          =   330
      Index           =   8
      Left            =   5280
      MousePointer    =   99  'Custom
      Tag             =   "09, ,F"
      Top             =   1680
      Width           =   555
   End
   Begin VB.Image D 
      Height          =   330
      Index           =   8
      Left            =   5280
      MousePointer    =   99  'Custom
      Tag             =   "11,#,A"
      Top             =   1320
      Width           =   555
   End
   Begin VB.Image G 
      Height          =   330
      Index           =   8
      Left            =   5280
      MousePointer    =   99  'Custom
      Tag             =   "14,#,D"
      Top             =   960
      Width           =   555
   End
   Begin VB.Image B 
      Height          =   330
      Index           =   8
      Left            =   5280
      MousePointer    =   99  'Custom
      Tag             =   "17, ,G"
      Top             =   600
      Width           =   555
   End
   Begin VB.Image E 
      Height          =   330
      Index           =   8
      Left            =   5280
      MousePointer    =   99  'Custom
      Tag             =   "20, ,C"
      Top             =   240
      Width           =   555
   End
   Begin VB.Image Es 
      Height          =   330
      Index           =   7
      Left            =   4680
      MousePointer    =   99  'Custom
      Tag             =   "05, ,B"
      Top             =   2040
      Width           =   555
   End
   Begin VB.Image A 
      Height          =   330
      Index           =   7
      Left            =   4680
      MousePointer    =   99  'Custom
      Tag             =   "08, ,E"
      Top             =   1680
      Width           =   555
   End
   Begin VB.Image D 
      Height          =   330
      Index           =   7
      Left            =   4680
      MousePointer    =   99  'Custom
      Tag             =   "11, ,A"
      Top             =   1320
      Width           =   555
   End
   Begin VB.Image G 
      Height          =   330
      Index           =   7
      Left            =   4680
      MousePointer    =   99  'Custom
      Tag             =   "14, ,D"
      Top             =   960
      Width           =   555
   End
   Begin VB.Image B 
      Height          =   330
      Index           =   7
      Left            =   4680
      MousePointer    =   99  'Custom
      Tag             =   "16,#,F"
      Top             =   600
      Width           =   555
   End
   Begin VB.Image E 
      Height          =   330
      Index           =   7
      Left            =   4680
      MousePointer    =   99  'Custom
      Tag             =   "19, ,B"
      Top             =   240
      Width           =   555
   End
   Begin VB.Image Es 
      Height          =   330
      Index           =   6
      Left            =   4080
      MousePointer    =   99  'Custom
      Tag             =   "04,#,A"
      Top             =   2040
      Width           =   555
   End
   Begin VB.Image A 
      Height          =   330
      Index           =   6
      Left            =   4080
      MousePointer    =   99  'Custom
      Tag             =   "07,#,D"
      Top             =   1680
      Width           =   555
   End
   Begin VB.Image D 
      Height          =   330
      Index           =   6
      Left            =   4080
      MousePointer    =   99  'Custom
      Tag             =   "10,#,G"
      Top             =   1320
      Width           =   555
   End
   Begin VB.Image G 
      Height          =   330
      Index           =   6
      Left            =   4080
      MousePointer    =   99  'Custom
      Tag             =   "13,#,C"
      Top             =   960
      Width           =   555
   End
   Begin VB.Image B 
      Height          =   330
      Index           =   6
      Left            =   4080
      MousePointer    =   99  'Custom
      Tag             =   "16, ,F"
      Top             =   600
      Width           =   555
   End
   Begin VB.Image E 
      Height          =   330
      Index           =   6
      Left            =   4080
      MousePointer    =   99  'Custom
      Tag             =   "18,#,A"
      Top             =   240
      Width           =   555
   End
   Begin VB.Image Es 
      Height          =   330
      Index           =   5
      Left            =   3480
      MousePointer    =   99  'Custom
      Tag             =   "04, ,A"
      Top             =   2040
      Width           =   555
   End
   Begin VB.Image A 
      Height          =   330
      Index           =   5
      Left            =   3480
      MousePointer    =   99  'Custom
      Tag             =   "07, ,D"
      Top             =   1680
      Width           =   555
   End
   Begin VB.Image D 
      Height          =   330
      Index           =   5
      Left            =   3480
      MousePointer    =   99  'Custom
      Tag             =   "10, ,G"
      Top             =   1320
      Width           =   555
   End
   Begin VB.Image G 
      Height          =   330
      Index           =   5
      Left            =   3480
      MousePointer    =   99  'Custom
      Tag             =   "13, ,C"
      Top             =   960
      Width           =   555
   End
   Begin VB.Image B 
      Height          =   330
      Index           =   5
      Left            =   3480
      MousePointer    =   99  'Custom
      Tag             =   "15, ,E"
      Top             =   600
      Width           =   555
   End
   Begin VB.Image E 
      Height          =   330
      Index           =   5
      Left            =   3480
      MousePointer    =   99  'Custom
      Tag             =   "18, ,A"
      Top             =   240
      Width           =   555
   End
   Begin VB.Image Es 
      Height          =   330
      Index           =   4
      Left            =   2880
      MousePointer    =   99  'Custom
      Tag             =   "03,#,G"
      Top             =   2040
      Width           =   555
   End
   Begin VB.Image A 
      Height          =   330
      Index           =   4
      Left            =   2880
      MousePointer    =   99  'Custom
      Tag             =   "06,#,C"
      Top             =   1680
      Width           =   555
   End
   Begin VB.Image D 
      Height          =   330
      Index           =   4
      Left            =   2880
      MousePointer    =   99  'Custom
      Tag             =   "09,#,F"
      Top             =   1320
      Width           =   555
   End
   Begin VB.Image G 
      Height          =   330
      Index           =   4
      Left            =   2880
      MousePointer    =   99  'Custom
      Tag             =   "12, ,B"
      Top             =   960
      Width           =   555
   End
   Begin VB.Image B 
      Height          =   330
      Index           =   4
      Left            =   2880
      MousePointer    =   99  'Custom
      Tag             =   "14,#,D"
      Top             =   600
      Width           =   555
   End
   Begin VB.Image E 
      Height          =   330
      Index           =   4
      Left            =   2880
      MousePointer    =   99  'Custom
      Tag             =   "17,#,G"
      Top             =   240
      Width           =   555
   End
   Begin VB.Image Es 
      Height          =   330
      Index           =   3
      Left            =   2280
      MousePointer    =   99  'Custom
      Tag             =   "03, ,G"
      Top             =   2040
      Width           =   555
   End
   Begin VB.Image A 
      Height          =   330
      Index           =   3
      Left            =   2280
      MousePointer    =   99  'Custom
      Tag             =   "06, ,C"
      Top             =   1680
      Width           =   555
   End
   Begin VB.Image D 
      Height          =   330
      Index           =   3
      Left            =   2280
      MousePointer    =   99  'Custom
      Tag             =   "09, ,F"
      Top             =   1320
      Width           =   555
   End
   Begin VB.Image G 
      Height          =   330
      Index           =   3
      Left            =   2280
      MousePointer    =   99  'Custom
      Tag             =   "11,#,A"
      Top             =   960
      Width           =   555
   End
   Begin VB.Image B 
      Height          =   330
      Index           =   3
      Left            =   2280
      MousePointer    =   99  'Custom
      Tag             =   "14, ,D"
      Top             =   600
      Width           =   555
   End
   Begin VB.Image E 
      Height          =   330
      Index           =   3
      Left            =   2280
      MousePointer    =   99  'Custom
      Tag             =   "17, ,G"
      Top             =   240
      Width           =   555
   End
   Begin VB.Image Es 
      Height          =   330
      Index           =   2
      Left            =   1680
      MousePointer    =   99  'Custom
      Tag             =   "02,#,F"
      Top             =   2040
      Width           =   555
   End
   Begin VB.Image A 
      Height          =   330
      Index           =   2
      Left            =   1680
      MousePointer    =   99  'Custom
      Tag             =   "05, ,B"
      Top             =   1680
      Width           =   555
   End
   Begin VB.Image D 
      Height          =   330
      Index           =   2
      Left            =   1680
      MousePointer    =   99  'Custom
      Tag             =   "08, ,E"
      Top             =   1320
      Width           =   555
   End
   Begin VB.Image G 
      Height          =   330
      Index           =   2
      Left            =   1680
      MousePointer    =   99  'Custom
      Tag             =   "11, ,A"
      Top             =   960
      Width           =   555
   End
   Begin VB.Image B 
      Height          =   330
      Index           =   2
      Left            =   1680
      MousePointer    =   99  'Custom
      Tag             =   "13,#,C"
      Top             =   600
      Width           =   555
   End
   Begin VB.Image Es 
      Height          =   330
      Index           =   1
      Left            =   1080
      MousePointer    =   99  'Custom
      Tag             =   "02, ,F"
      Top             =   2040
      Width           =   555
   End
   Begin VB.Image A 
      Height          =   330
      Index           =   1
      Left            =   1080
      MousePointer    =   99  'Custom
      Tag             =   "04,#,A"
      Top             =   1680
      Width           =   555
   End
   Begin VB.Image D 
      Height          =   330
      Index           =   1
      Left            =   1080
      MousePointer    =   99  'Custom
      Tag             =   "07,#,D"
      Top             =   1320
      Width           =   555
   End
   Begin VB.Image G 
      Height          =   330
      Index           =   1
      Left            =   1080
      MousePointer    =   99  'Custom
      Tag             =   "10,#,G"
      Top             =   960
      Width           =   555
   End
   Begin VB.Image B 
      Height          =   330
      Index           =   1
      Left            =   1080
      MousePointer    =   99  'Custom
      Tag             =   "13, ,C"
      Top             =   600
      Width           =   555
   End
   Begin VB.Image Image1 
      Height          =   2415
      Left            =   5160
      Picture         =   "frmNoteTable.frx":2B04
      Top             =   2400
      Width           =   3585
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Main"
      Visible         =   0   'False
      Begin VB.Menu mnuComp 
         Caption         =   "Normal"
      End
      Begin VB.Menu mnuSharp 
         Caption         =   "Sharp (#)"
      End
      Begin VB.Menu mnuFlat 
         Caption         =   "Flat (b)"
      End
   End
End
Attribute VB_Name = "frmNoteTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Halil TUNC © 2004
'HtGuitar ...\Form\frmNoteTable.frm
'19/09/2004 Ankara TURKEY
'halil_tunc@hotmail.com

Option Explicit

Dim lNoteLine As Long, bNoteSharp As Boolean, bNoteFlat As Boolean, sNoteName As String
Dim pImgPart01 As StdPicture, pImgPart02 As StdPicture, pImgPart03 As StdPicture
Dim pImgIco01 As StdPicture, pImgIco02 As StdPicture, pImgIco03 As StdPicture, pImgIco04 As StdPicture
Dim pImgCur01 As StdPicture
Dim iNoteType As Integer
Private Const lLineHeight = 5
Private Const lLineTop = 0

Private Sub E_Click(Index As Integer)
   pSetValues E(Index).Tag
   pShowNoteInfo
   E(Index).Picture = pImgPart02
   'Debug.Print lNoteLine & "-" & bNoteSharp & "-" & sNoteName
End Sub

Private Sub B_Click(Index As Integer)
   pSetValues B(Index).Tag
   pShowNoteInfo
   B(Index).Picture = pImgPart02
   'Debug.Print lNoteLine & "-" & bNoteSharp & "-" & sNoteName
End Sub

Private Sub G_Click(Index As Integer)
   pSetValues G(Index).Tag
   pShowNoteInfo
   G(Index).Picture = pImgPart02
   'Debug.Print lNoteLine & "-" & bNoteSharp & "-" & sNoteName
End Sub

Private Sub D_Click(Index As Integer)
   pSetValues D(Index).Tag
   pShowNoteInfo
   D(Index).Picture = pImgPart02
   'Debug.Print lNoteLine & "-" & bNoteSharp & "-" & sNoteName
End Sub

Private Sub A_Click(Index As Integer)
   pSetValues A(Index).Tag
   pShowNoteInfo
   A(Index).Picture = pImgPart02
   'Debug.Print lNoteLine & "-" & bNoteSharp & "-" & sNoteName
End Sub

Private Sub Es_Click(Index As Integer)
   pSetValues Es(Index).Tag
   pShowNoteInfo
   Es(Index).Picture = pImgPart02
   'Debug.Print lNoteLine & "-" & bNoteSharp & "-" & sNoteName
End Sub

Private Sub Form_Load()

   Set pImgCur01 = LoadResPicture(101, vbResCursor)

   Set pImgPart01 = LoadResPicture(101, vbResBitmap)
   Set pImgPart02 = LoadResPicture(102, vbResBitmap)
   Set pImgPart03 = LoadResPicture(103, vbResBitmap)

   Set pImgIco01 = LoadResPicture(101, vbResIcon)
   Set pImgIco02 = LoadResPicture(102, vbResIcon)
   Set pImgIco03 = LoadResPicture(103, vbResIcon)
   Set pImgIco04 = LoadResPicture(104, vbResIcon)

   Me.Icon = LoadResPicture(106, vbResIcon)
   lblClose.MouseIcon = pImgCur01
   imgSharp.MouseIcon = pImgCur01
   imgFlat.MouseIcon = pImgCur01
   imgSharp.Picture = pImgIco01
   imgFlat.Picture = pImgIco04

   iNoteType = 0

   pPartSetting

   ImgNote.Top = 5000
   imgNoteType.Top = 5000
   lNoteLine = 50
   ImgNote.Picture = LoadResPicture(105, vbResIcon)

End Sub

Private Sub imgFlat_Click()
   If Not bNoteFlat Then
      bNoteFlat = True
      imgFlat.Picture = pImgIco02
      imgSharp.Picture = pImgIco03
      If bNoteSharp Then
         lNoteLine = lNoteLine + 1
         iNoteType = 2
      End If
      pShowNoteInfo
   End If
End Sub

Private Sub imgSharp_Click()
   If bNoteFlat Then
      bNoteFlat = False
      imgFlat.Picture = pImgIco04
      imgSharp.Picture = pImgIco01
      If bNoteSharp Then
         lNoteLine = lNoteLine - 1
         iNoteType = 1
      End If
      pShowNoteInfo
   End If
End Sub

Private Sub lblClose_Click()
   Unload Me
End Sub

Private Sub mnuComp_Click()
If iNoteType = 0 Then
   pShowNoteInfo
ElseIf iNoteType = 1 Then
   iNoteType = 0
   bNoteSharp = False
   pShowNoteInfo
ElseIf iNoteType = 2 Then
   iNoteType = 0
   bNoteSharp = False
   sNoteName = fSharpToFlat(sNoteName)
   pShowNoteInfo
End If
End Sub

Private Sub mnuFlat_Click()
If iNoteType = 0 Then
   iNoteType = 2
   bNoteSharp = True
   bNoteFlat = True
   imgSharp.Picture = pImgIco03
   imgFlat.Picture = pImgIco02
   sNoteName = fFlatToSharp(sNoteName)
   pShowNoteInfo
ElseIf iNoteType = 1 Then
   iNoteType = 2
   bNoteSharp = True
   bNoteFlat = True
   imgSharp.Picture = pImgIco03
   imgFlat.Picture = pImgIco02
   sNoteName = fFlatToSharp(sNoteName)
   pShowNoteInfo
ElseIf iNoteType = 2 Then
   pShowNoteInfo
End If
End Sub

Private Sub mnuSharp_Click()
If iNoteType = 0 Then
   iNoteType = 1
   bNoteSharp = True
   bNoteFlat = False
   imgSharp.Picture = pImgIco01
   imgFlat.Picture = pImgIco04
   pShowNoteInfo
ElseIf iNoteType = 1 Then
   pShowNoteInfo
ElseIf iNoteType = 2 Then
   iNoteType = 1
   bNoteSharp = True
   bNoteFlat = False
   imgSharp.Picture = pImgIco01
   imgFlat.Picture = pImgIco04
   sNoteName = fSharpToFlat(sNoteName)
   pShowNoteInfo
End If
End Sub

Private Sub picStave_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = vbLeftButton Then
      fStaveClick X, Y
   ElseIf Button = vbRightButton Then
      Me.PopupMenu mnuMain
   End If
End Sub

Private Function fFlatToSharp(sValue As String) As String
   If sValue = "C" Then
      fFlatToSharp = "B"
   ElseIf sValue = "D" Then
      fFlatToSharp = "C"
   ElseIf sValue = "E" Then
      fFlatToSharp = "D"
   ElseIf sValue = "F" Then
      fFlatToSharp = "E"
   ElseIf sValue = "G" Then
      fFlatToSharp = "F"
   ElseIf sValue = "A" Then
      fFlatToSharp = "G"
   ElseIf sValue = "B" Then
      fFlatToSharp = "A"
   End If
End Function

Private Function fGetNoteName() As String
   If bNoteFlat Then
      If bNoteSharp Then
         fGetNoteName = fSharpToFlat(sNoteName) & "b"
      Else
         fGetNoteName = sNoteName
      End If
   Else
      If bNoteSharp Then
         fGetNoteName = sNoteName & "#"
      Else
         fGetNoteName = sNoteName
      End If
   End If
End Function

Private Function fGetToolTipText(sTag As String) As String
Dim sTemp As String, sName As String, bSharp As Boolean

sName = Mid(sTag, 6, 3)
bSharp = IIf(Mid(sTag, 4, 1) = " ", False, True)

   If Not bSharp Then
      sTemp = sName
   Else
      sTemp = sName & "# / "
      sTemp = sTemp & fSharpToFlat(sName)
      sTemp = sTemp & "b"
   End If

   fGetToolTipText = sTemp

End Function

Private Function fLineToNote(lLine As Long) As Boolean
lNoteLine = lLine
Select Case lLine
   Case 1
      sNoteName = "E"
   Case 2
      sNoteName = "F"
   Case 3
      sNoteName = "G"
   Case 4
      sNoteName = "A"
   Case 5
      sNoteName = "B"
   Case 6
      sNoteName = "C"
   Case 7
      sNoteName = "D"
   Case 8
      sNoteName = "E"
   Case 9
      sNoteName = "F"
   Case 10
      sNoteName = "G"
   Case 11
      sNoteName = "A"
   Case 12
      sNoteName = "B"
   Case 13
      sNoteName = "C"
   Case 14
      sNoteName = "D"
   Case 15
      sNoteName = "E"
   Case 16
      sNoteName = "F"
   Case 17
      sNoteName = "G"
   Case 18
      sNoteName = "A"
   Case 19
      sNoteName = "B"
   Case 20
      sNoteName = "C"
   Case 21
      sNoteName = "D"
   Case 22
      sNoteName = "E"
   Case 23
      sNoteName = "F"
   Case 24
      sNoteName = "G"
   Case 25
      sNoteName = "A"
   Case 26
      sNoteName = "B"
End Select
If bNoteFlat And bNoteSharp Then sNoteName = fFlatToSharp(sNoteName)
End Function

Private Function fSharpToFlat(sValue As String) As String
   If sValue = "C" Then
      fSharpToFlat = "D"
   ElseIf sValue = "D" Then
      fSharpToFlat = "E"
   ElseIf sValue = "E" Then
      fSharpToFlat = "F"
   ElseIf sValue = "F" Then
      fSharpToFlat = "G"
   ElseIf sValue = "G" Then
      fSharpToFlat = "A"
   ElseIf sValue = "A" Then
      fSharpToFlat = "B"
   ElseIf sValue = "B" Then
      fSharpToFlat = "C"
   End If
End Function

Private Function fStaveClick(lX As Single, lY As Single) As Boolean
Dim lLine As Long, lMaxLine As Long, lMinLine As Long
lX = ScaleX(lX, picStave.ScaleMode, vbPixels)
lY = ScaleY(lY, picStave.ScaleMode, vbPixels)
lMaxLine = lLineTop + lLineHeight * 27
   If lY < 14 Then
      fLineToNote 26
      pShowNoteInfo
   ElseIf lY > lMaxLine Then
      fLineToNote 1
      pShowNoteInfo
   Else
      fLineToNote 26 - lY / lLineHeight - lLineTop + 2
      pShowNoteInfo
   End If
      Debug.Print lNoteLine
End Function

Private Sub pPartSetting()
Dim i As Long, lTop As Long, lLeft As Long
   lLeft = 240
   lTop = 200

   imgLT.Top = lTop
   imgLT.Left = lLeft
   imgLT.Picture = LoadResPicture(104, vbResBitmap)
   lLeft = lLeft + imgLT.Width
   For i = 0 To 19
      imgTop(i).Left = lLeft
      imgTop(i).Top = lTop
      imgTop(i).Picture = LoadResPicture(107, vbResBitmap)
      lblTop(i).Top = lTop - lblTop(i).Height + imgTop(i).Height
      lblTop(i).Left = lLeft
      lblTop(i).Caption = i
      lblTop(i).Width = imgTop(i).Width
      lLeft = lLeft + imgTop(i).Width
   Next
   imgRT.Top = lTop
   imgRT.Left = lLeft
   imgRT.Picture = LoadResPicture(109, vbResBitmap)
   lTop = lTop + imgLT.Height
   lLeft = 240

   E0.Top = lTop
   E0.Left = lLeft
   E0.Picture = LoadResPicture(105, vbResBitmap)
   lblE.Left = lLeft
   lblE.Top = lTop + 30
   lLeft = lLeft + E0.Width
   For i = 0 To 19
      E(i).Picture = pImgPart01
      E(i).Left = lLeft
      E(i).Top = lTop
      E(i).MouseIcon = pImgCur01
      E(i).ToolTipText = fGetToolTipText(E(i).Tag)
      lLeft = lLeft + E(i).Width
   Next
   E1.Top = lTop
   E1.Left = lLeft
   E1.Picture = LoadResPicture(110, vbResBitmap)
   lTop = lTop + E0.Height
   lLeft = 240

   B0.Top = lTop
   B0.Left = lLeft
   B0.Picture = LoadResPicture(105, vbResBitmap)
   lblB.Left = lLeft
   lblB.Top = lTop + 30
   lLeft = lLeft + B0.Width
   For i = 0 To 19
      B(i).Picture = pImgPart01
      B(i).Left = lLeft
      B(i).Top = lTop
      B(i).MouseIcon = pImgCur01
      B(i).ToolTipText = fGetToolTipText(B(i).Tag)
      lLeft = lLeft + B(i).Width
   Next
   B1.Top = lTop
   B1.Left = lLeft
   B1.Picture = LoadResPicture(110, vbResBitmap)
   lTop = lTop + B0.Height
   lLeft = 240

   G0.Top = lTop
   G0.Left = lLeft
   G0.Picture = LoadResPicture(105, vbResBitmap)
   lblG.Left = lLeft
   lblG.Top = lTop + 30
   lLeft = lLeft + G0.Width
   For i = 0 To 19
      G(i).Picture = pImgPart01
      G(i).Left = lLeft
      G(i).Top = lTop
      G(i).MouseIcon = pImgCur01
      G(i).ToolTipText = fGetToolTipText(G(i).Tag)
      lLeft = lLeft + G(i).Width
   Next
   G1.Top = lTop
   G1.Left = lLeft
   G1.Picture = LoadResPicture(110, vbResBitmap)
   lTop = lTop + G0.Height
   lLeft = 240

   D0.Top = lTop
   D0.Left = lLeft
   D0.Picture = LoadResPicture(105, vbResBitmap)
   lblD.Left = lLeft
   lblD.Top = lTop + 30
   lLeft = lLeft + D0.Width
   For i = 0 To 19
      D(i).Picture = pImgPart01
      D(i).Left = lLeft
      D(i).Top = lTop
      D(i).MouseIcon = pImgCur01
      D(i).ToolTipText = fGetToolTipText(D(i).Tag)
      lLeft = lLeft + D(i).Width
   Next
   D1.Top = lTop
   D1.Left = lLeft
   D1.Picture = LoadResPicture(110, vbResBitmap)
   lTop = lTop + D0.Height
   lLeft = 240

   A0.Top = lTop
   A0.Left = lLeft
   A0.Picture = LoadResPicture(105, vbResBitmap)
   lblA.Left = lLeft
   lblA.Top = lTop + 30
   lLeft = lLeft + A0.Width
   For i = 0 To 19
      A(i).Picture = pImgPart01
      A(i).Left = lLeft
      A(i).Top = lTop
      A(i).MouseIcon = pImgCur01
      A(i).ToolTipText = fGetToolTipText(A(i).Tag)
      lLeft = lLeft + A(i).Width
   Next
   A1.Top = lTop
   A1.Left = lLeft
   A1.Picture = LoadResPicture(110, vbResBitmap)
   lTop = lTop + A0.Height
   lLeft = 240

   Es0.Top = lTop
   Es0.Left = lLeft
   Es0.Picture = LoadResPicture(105, vbResBitmap)
   lblEs.Left = lLeft
   lblEs.Top = lTop + 30
   lLeft = lLeft + Es0.Width
   For i = 0 To 19
      Es(i).Picture = pImgPart01
      Es(i).Left = lLeft
      Es(i).Top = lTop
      Es(i).MouseIcon = pImgCur01
      Es(i).ToolTipText = fGetToolTipText(Es(i).Tag)
      lLeft = lLeft + Es(i).Width
   Next
   Es1.Top = lTop
   Es1.Left = lLeft
   Es1.Picture = LoadResPicture(110, vbResBitmap)
   lTop = lTop + Es0.Height
   lLeft = 240

   imgLB.Top = lTop
   imgLB.Left = lLeft
   imgLB.Picture = LoadResPicture(106, vbResBitmap)
   lLeft = lLeft + imgLB.Width
   For i = 0 To 19
      imgBottom(i).Left = lLeft
      imgBottom(i).Top = lTop
      imgBottom(i).Picture = LoadResPicture(108, vbResBitmap)
      lLeft = lLeft + imgBottom(i).Width
   Next
   imgRB.Top = lTop
   imgRB.Left = lLeft
   imgRB.Picture = LoadResPicture(111, vbResBitmap)

End Sub

Private Sub pSetValues(sTag As String)
   bNoteSharp = IIf(Mid(sTag, 4, 1) = " ", False, True)
   lNoteLine = Val(Left(sTag, 2)) + IIf(bNoteSharp, IIf(bNoteFlat, 1, 0), 0)
   sNoteName = Mid(sTag, 6, 3)
   If Not bNoteSharp Then
      iNoteType = 0
   Else
      iNoteType = IIf(bNoteFlat, 2, 1)
   End If
End Sub

Private Sub pShowNoteInfo()
Dim i As Long
   lblNoteName.Caption = fGetNoteName

   For i = 0 To 19
      If Val(Left(E(i).Tag, 2)) = lNoteLine - IIf(bNoteFlat And bNoteSharp, 1, 0) And Mid(E(i).Tag, 4, 1) = IIf(bNoteSharp, "#", " ") Then
         E(i).Picture = pImgPart03
         E(i).Appearance = 0
      Else
         If E(i).Appearance = 0 Then E(i).Picture = pImgPart01
      End If

      If Val(Left(B(i).Tag, 2)) = lNoteLine - IIf(bNoteFlat And bNoteSharp, 1, 0) And Mid(B(i).Tag, 4, 1) = IIf(bNoteSharp, "#", " ") Then
         B(i).Picture = pImgPart03
         B(i).Appearance = 0
      Else
         If B(i).Appearance = 0 Then B(i).Picture = pImgPart01
      End If

      If Val(Left(G(i).Tag, 2)) = lNoteLine - IIf(bNoteFlat And bNoteSharp, 1, 0) And Mid(G(i).Tag, 4, 1) = IIf(bNoteSharp, "#", " ") Then
         G(i).Picture = pImgPart03
         G(i).Appearance = 0
      Else
         If G(i).Appearance = 0 Then G(i).Picture = pImgPart01
      End If

      If Val(Left(D(i).Tag, 2)) = lNoteLine - IIf(bNoteFlat And bNoteSharp, 1, 0) And Mid(D(i).Tag, 4, 1) = IIf(bNoteSharp, "#", " ") Then
         D(i).Picture = pImgPart03
         D(i).Appearance = 0
      Else
         If D(i).Appearance = 0 Then D(i).Picture = pImgPart01
      End If

      If Val(Left(A(i).Tag, 2)) = lNoteLine - IIf(bNoteFlat And bNoteSharp, 1, 0) And Mid(A(i).Tag, 4, 1) = IIf(bNoteSharp, "#", " ") Then
         A(i).Picture = pImgPart03
         A(i).Appearance = 0
      Else
         If A(i).Appearance = 0 Then A(i).Picture = pImgPart01
      End If

      If Val(Left(Es(i).Tag, 2)) = lNoteLine - IIf(bNoteFlat And bNoteSharp, 1, 0) And Mid(Es(i).Tag, 4, 1) = IIf(bNoteSharp, "#", " ") Then
         Es(i).Picture = pImgPart03
         Es(i).Appearance = 0
      Else
         If Es(i).Appearance = 0 Then Es(i).Picture = pImgPart01
      End If

   Next
    
   ImgNote.Top = ScaleY((lLineTop + lLineHeight * (26 - lNoteLine)) + lLineHeight - 1, vbPixels, picStave.ScaleMode)
   imgNoteType.Top = ImgNote.Top - 60

   If iNoteType = 0 Then
      imgNoteType.Picture = Nothing
   ElseIf iNoteType = 1 Then
      imgNoteType.Picture = LoadResPicture(101, vbResIcon)
   ElseIf iNoteType = 2 Then
      imgNoteType.Picture = LoadResPicture(102, vbResIcon)
   End If

End Sub
