VERSION 5.00
Begin VB.Form frmChord 
   BackColor       =   &H00D8FFFF&
   Caption         =   "Chord finder"
   ClientHeight    =   4440
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5955
   Icon            =   "frmChord.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   5955
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstGroup 
      Appearance      =   0  'Flat
      BackColor       =   &H0048DCFF&
      ForeColor       =   &H000034D0&
      Height          =   1395
      ItemData        =   "frmChord.frx":000C
      Left            =   330
      List            =   "frmChord.frx":000E
      TabIndex        =   21
      Top             =   2280
      Width           =   2415
   End
   Begin VB.ListBox lstChord 
      Appearance      =   0  'Flat
      BackColor       =   &H0048DCFF&
      ForeColor       =   &H000034D0&
      Height          =   1395
      Left            =   2955
      TabIndex        =   20
      Top             =   2280
      Width           =   2655
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
   Begin VB.Image G 
      Height          =   330
      Index           =   1
      Left            =   1080
      MousePointer    =   99  'Custom
      Tag             =   "10,#,G"
      Top             =   960
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
   Begin VB.Image A 
      Height          =   330
      Index           =   1
      Left            =   1080
      MousePointer    =   99  'Custom
      Tag             =   "04,#,A"
      Top             =   1680
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
   Begin VB.Image B 
      Height          =   330
      Index           =   2
      Left            =   1680
      MousePointer    =   99  'Custom
      Tag             =   "13,#,C"
      Top             =   600
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
   Begin VB.Image D 
      Height          =   330
      Index           =   2
      Left            =   1680
      MousePointer    =   99  'Custom
      Tag             =   "08, ,E"
      Top             =   1320
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
   Begin VB.Image Es 
      Height          =   330
      Index           =   2
      Left            =   1680
      MousePointer    =   99  'Custom
      Tag             =   "02,#,F"
      Top             =   2040
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
   Begin VB.Image B 
      Height          =   330
      Index           =   3
      Left            =   2280
      MousePointer    =   99  'Custom
      Tag             =   "14, ,D"
      Top             =   600
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
   Begin VB.Image D 
      Height          =   330
      Index           =   3
      Left            =   2280
      MousePointer    =   99  'Custom
      Tag             =   "09, ,F"
      Top             =   1320
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
   Begin VB.Image Es 
      Height          =   330
      Index           =   3
      Left            =   2280
      MousePointer    =   99  'Custom
      Tag             =   "03, ,G"
      Top             =   2040
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
   Begin VB.Image B 
      Height          =   330
      Index           =   4
      Left            =   2880
      MousePointer    =   99  'Custom
      Tag             =   "14,#,D"
      Top             =   600
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
   Begin VB.Image D 
      Height          =   330
      Index           =   4
      Left            =   2880
      MousePointer    =   99  'Custom
      Tag             =   "09,#,F"
      Top             =   1320
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
   Begin VB.Image Es 
      Height          =   330
      Index           =   4
      Left            =   2880
      MousePointer    =   99  'Custom
      Tag             =   "03,#,G"
      Top             =   2040
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
   Begin VB.Image B 
      Height          =   330
      Index           =   5
      Left            =   3480
      MousePointer    =   99  'Custom
      Tag             =   "15, ,E"
      Top             =   600
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
   Begin VB.Image D 
      Height          =   330
      Index           =   5
      Left            =   3480
      MousePointer    =   99  'Custom
      Tag             =   "10, ,G"
      Top             =   1320
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
   Begin VB.Image Es 
      Height          =   330
      Index           =   5
      Left            =   3480
      MousePointer    =   99  'Custom
      Tag             =   "04, ,A"
      Top             =   2040
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
   Begin VB.Image B 
      Height          =   330
      Index           =   6
      Left            =   4080
      MousePointer    =   99  'Custom
      Tag             =   "16, ,F"
      Top             =   600
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
   Begin VB.Image D 
      Height          =   330
      Index           =   6
      Left            =   4080
      MousePointer    =   99  'Custom
      Tag             =   "10,#,G"
      Top             =   1320
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
   Begin VB.Image Es 
      Height          =   330
      Index           =   6
      Left            =   4080
      MousePointer    =   99  'Custom
      Tag             =   "04,#,A"
      Top             =   2040
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
   Begin VB.Image B 
      Height          =   330
      Index           =   7
      Left            =   4680
      MousePointer    =   99  'Custom
      Tag             =   "16,#,F"
      Top             =   600
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
   Begin VB.Image D 
      Height          =   330
      Index           =   7
      Left            =   4680
      MousePointer    =   99  'Custom
      Tag             =   "11, ,A"
      Top             =   1320
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
   Begin VB.Image Es 
      Height          =   330
      Index           =   7
      Left            =   4680
      MousePointer    =   99  'Custom
      Tag             =   "05, ,B"
      Top             =   2040
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
   Begin VB.Image B 
      Height          =   330
      Index           =   8
      Left            =   5280
      MousePointer    =   99  'Custom
      Tag             =   "17, ,G"
      Top             =   600
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
   Begin VB.Image D 
      Height          =   330
      Index           =   8
      Left            =   5280
      MousePointer    =   99  'Custom
      Tag             =   "11,#,A"
      Top             =   1320
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
   Begin VB.Image Es 
      Height          =   330
      Index           =   8
      Left            =   5280
      MousePointer    =   99  'Custom
      Tag             =   "06, ,C"
      Top             =   2040
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
   Begin VB.Image B 
      Height          =   330
      Index           =   9
      Left            =   5880
      MousePointer    =   99  'Custom
      Tag             =   "17,#,G"
      Top             =   600
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
   Begin VB.Image D 
      Height          =   330
      Index           =   9
      Left            =   5880
      MousePointer    =   99  'Custom
      Tag             =   "12, ,B"
      Top             =   1320
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
   Begin VB.Image Es 
      Height          =   330
      Index           =   9
      Left            =   5880
      MousePointer    =   99  'Custom
      Tag             =   "06,#,C"
      Top             =   2040
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
   Begin VB.Image B 
      Height          =   330
      Index           =   10
      Left            =   6480
      MousePointer    =   99  'Custom
      Tag             =   "18, ,A"
      Top             =   600
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
   Begin VB.Image D 
      Height          =   330
      Index           =   10
      Left            =   6480
      MousePointer    =   99  'Custom
      Tag             =   "13, ,C"
      Top             =   1320
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
   Begin VB.Image Es 
      Height          =   330
      Index           =   10
      Left            =   6480
      MousePointer    =   99  'Custom
      Tag             =   "07, ,D"
      Top             =   2040
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
   Begin VB.Image B 
      Height          =   330
      Index           =   11
      Left            =   7080
      MousePointer    =   99  'Custom
      Tag             =   "18,#,A"
      Top             =   600
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
   Begin VB.Image D 
      Height          =   330
      Index           =   11
      Left            =   7080
      MousePointer    =   99  'Custom
      Tag             =   "13,#,C"
      Top             =   1320
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
   Begin VB.Image Es 
      Height          =   330
      Index           =   11
      Left            =   7080
      MousePointer    =   99  'Custom
      Tag             =   "07,#,D"
      Top             =   2040
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
   Begin VB.Image B 
      Height          =   330
      Index           =   12
      Left            =   7680
      MousePointer    =   99  'Custom
      Tag             =   "19, ,B"
      Top             =   600
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
   Begin VB.Image D 
      Height          =   330
      Index           =   12
      Left            =   7680
      MousePointer    =   99  'Custom
      Tag             =   "14, ,D"
      Top             =   1320
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
   Begin VB.Image Es 
      Height          =   330
      Index           =   12
      Left            =   7680
      MousePointer    =   99  'Custom
      Tag             =   "08, ,E"
      Top             =   2040
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
   Begin VB.Image E 
      Height          =   330
      Index           =   1
      Left            =   1080
      MousePointer    =   99  'Custom
      Tag             =   "16, ,F"
      Top             =   240
      Width           =   555
   End
   Begin VB.Label lblChordEditor 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Chord editor "
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
      Left            =   1080
      MousePointer    =   99  'Custom
      TabIndex        =   23
      Top             =   3840
      Visible         =   0   'False
      Width           =   1470
   End
   Begin VB.Label lblFavorites 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Favorites"
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
      Left            =   3600
      MousePointer    =   99  'Custom
      TabIndex        =   22
      Top             =   3840
      Width           =   1155
   End
   Begin VB.Label lblDesc 
      BackStyle       =   0  'Transparent
      Caption         =   "Desc"
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   3945
      Width           =   2775
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
      Left            =   5040
      MousePointer    =   99  'Custom
      TabIndex        =   18
      Top             =   3840
      Width           =   570
   End
   Begin VB.Image imgRB 
      Height          =   330
      Left            =   8280
      Tag             =   "15, ,Mi"
      Top             =   2400
      Width           =   300
   End
   Begin VB.Image Es1 
      Height          =   330
      Left            =   8280
      Tag             =   "15, ,Mi"
      Top             =   2040
      Width           =   300
   End
   Begin VB.Image A1 
      Height          =   330
      Left            =   8280
      Tag             =   "15, ,Mi"
      Top             =   1680
      Width           =   300
   End
   Begin VB.Image D1 
      Height          =   330
      Left            =   8280
      Tag             =   "15, ,Mi"
      Top             =   1320
      Width           =   300
   End
   Begin VB.Image G1 
      Height          =   330
      Left            =   8280
      Tag             =   "15, ,Mi"
      Top             =   960
      Width           =   300
   End
   Begin VB.Image B1 
      Height          =   330
      Left            =   8280
      Tag             =   "15, ,Mi"
      Top             =   600
      Width           =   300
   End
   Begin VB.Image E1 
      Height          =   330
      Left            =   8280
      Tag             =   "15, ,Mi"
      Top             =   240
      Width           =   300
   End
   Begin VB.Image imgRT 
      Height          =   330
      Left            =   8280
      Tag             =   "15, ,Mi"
      Top             =   0
      Width           =   300
   End
   Begin VB.Label lblTop 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "O"
      ForeColor       =   &H000034D0&
      Height          =   255
      Index           =   1
      Left            =   1080
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
      Index           =   2
      Left            =   1680
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
      Index           =   4
      Left            =   2880
      TabIndex        =   9
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
      Index           =   6
      Left            =   4080
      TabIndex        =   7
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
      TabIndex        =   6
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
      TabIndex        =   5
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
      TabIndex        =   4
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
      TabIndex        =   3
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
      TabIndex        =   2
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
      TabIndex        =   1
      Top             =   2640
      Width           =   495
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
      Left            =   -45
      TabIndex        =   17
      ToolTipText     =   "E"
      Top             =   240
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
      Left            =   -45
      TabIndex        =   16
      ToolTipText     =   "B"
      Top             =   600
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
      Left            =   -45
      TabIndex        =   15
      ToolTipText     =   "G"
      Top             =   960
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
      Left            =   -45
      TabIndex        =   14
      ToolTipText     =   "D"
      Top             =   1320
      Width           =   375
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
      Left            =   -45
      TabIndex        =   13
      ToolTipText     =   "A"
      Top             =   1680
      Width           =   375
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
      Left            =   -45
      TabIndex        =   0
      ToolTipText     =   "E"
      Top             =   2040
      Width           =   375
   End
   Begin VB.Image E0 
      Height          =   330
      Left            =   120
      ToolTipText     =   "E"
      Top             =   240
      Width           =   300
   End
   Begin VB.Image B0 
      Height          =   330
      Left            =   120
      Tag             =   "12, ,Si"
      ToolTipText     =   "B"
      Top             =   600
      Width           =   300
   End
   Begin VB.Image G0 
      Height          =   330
      Left            =   120
      Tag             =   "10, ,Sol"
      ToolTipText     =   "G"
      Top             =   960
      Width           =   300
   End
   Begin VB.Image D0 
      Height          =   330
      Left            =   120
      ToolTipText     =   "D"
      Top             =   1320
      Width           =   300
   End
   Begin VB.Image A0 
      Height          =   330
      Left            =   120
      ToolTipText     =   "A"
      Top             =   1680
      Width           =   300
   End
   Begin VB.Image Es0 
      Height          =   330
      Left            =   120
      ToolTipText     =   "E"
      Top             =   2040
      Width           =   300
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
      Index           =   2
      Left            =   1680
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
      Index           =   4
      Left            =   2880
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
      Index           =   6
      Left            =   4080
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
      Index           =   8
      Left            =   5280
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
      Index           =   10
      Left            =   6480
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
      Index           =   12
      Left            =   7680
      MousePointer    =   99  'Custom
      Top             =   0
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
      Index           =   2
      Left            =   1680
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
      Index           =   4
      Left            =   2880
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
      Index           =   6
      Left            =   4080
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
      Index           =   8
      Left            =   5280
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
      Index           =   10
      Left            =   6480
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
      Index           =   12
      Left            =   7680
      MousePointer    =   99  'Custom
      Top             =   2400
      Width           =   555
   End
   Begin VB.Image imgLT 
      Height          =   330
      Left            =   120
      Tag             =   "15, ,Mi"
      Top             =   0
      Width           =   300
   End
   Begin VB.Image imgLB 
      Height          =   330
      Left            =   120
      Tag             =   "15, ,Mi"
      Top             =   2400
      Width           =   300
   End
End
Attribute VB_Name = "frmChord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Halil TUNC © 2004
'HtGuitar ...\Form\frmNoteTable.frm
'19/09/2004 Ankara TURKEY
'halil_tunc@hotmail.com

Option Explicit

Dim pImgPart01 As StdPicture, pImgPartSel As StdPicture, pImgPartNP As StdPicture, pImgPartBarreF As StdPicture, pImgPartBarreM As StdPicture, pImgPartBarreL As StdPicture
Dim pImgCur01 As StdPicture

Private Type tpChords
   lChordID As Long
   sGroupName As String
   sChordName As String
   iLine1 As Integer
   iLine2 As Integer
   iLine3 As Integer
   iLine4 As Integer
   iLine5 As Integer
   iLine6 As Integer
   iFinger1 As Integer
   iFinger2 As Integer
   iFinger3 As Integer
   iFinger4 As Integer
   iFinger5 As Integer
   iFinger6 As Integer
   sDescription As String
   iBarreFirst As Integer
   iBarreLast As Integer
   iBarreFret As Integer
End Type

Private bFavorites As Boolean
Private m_Chords() As tpChords, m_ChordCount As Long

Private Sub Form_Load()
   Set pImgCur01 = LoadResPicture(101, vbResCursor)
   Set pImgPart01 = LoadResPicture(101, vbResBitmap)
   Set pImgPartSel = LoadResPicture(120, vbResBitmap)
   Set pImgPartNP = LoadResPicture(112, vbResBitmap)
   Set pImgPartBarreF = LoadResPicture(121, vbResBitmap)
   Set pImgPartBarreM = LoadResPicture(122, vbResBitmap)
   Set pImgPartBarreL = LoadResPicture(123, vbResBitmap)
   Me.Icon = LoadResPicture(106, vbResIcon)
   lblClose.MouseIcon = pImgCur01
   lblFavorites.MouseIcon = pImgCur01
   lblDesc.Caption = ""
   bFavorites = False
   pPartSetting
   pLoadChord
   pFillChordList
End Sub

Private Sub lblClose_Click()
   Unload Me
End Sub

Private Sub lblFavorites_Click()
   bFavorites = Not bFavorites
   lblFavorites.ForeColor = IIf(bFavorites, &H8000&, &H34D0&)
   pLoadChord
   pFillChordList
End Sub

Private Sub lstChord_Click()
   pShowChord
End Sub

Private Sub lstGroup_Click()
Dim i As Long, iTempIndex As Integer
   iTempIndex = lstChord.ListIndex
   lstChord.Clear
   For i = 1 To m_ChordCount
   With m_Chords(i)
      If .sGroupName = lstGroup.Text Then
         If .sDescription = "" Then
            lstChord.AddItem .sChordName
         Else
            lstChord.AddItem .sChordName
         End If
      End If
   End With
   Next i
   With lstChord
   If .ListCount > 0 Then
      If .ListCount <= iTempIndex Then
         .ListIndex = .ListCount - 1
      ElseIf iTempIndex = -1 Then
         .ListIndex = 0
      Else
         .ListIndex = iTempIndex
      End If
   End If
   End With
End Sub

Private Function fGetIndex() As Long
Dim i As Long, sGr As String, sCh As String
sGr = lstGroup.Text
sCh = lstChord.Text
   For i = 1 To m_ChordCount
      If m_Chords(i).sGroupName = sGr And m_Chords(i).sChordName = sCh Then
         fGetIndex = i
      End If
   Next
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

Private Sub pFillChordList()
Dim i As Long, sTempText As String
   lstGroup.Clear
   For i = 1 To m_ChordCount
      If m_Chords(i).sGroupName <> sTempText Then
         sTempText = m_Chords(i).sGroupName
         lstGroup.AddItem sTempText
      End If
   Next
   If lstGroup.ListCount > 0 Then lstGroup.ListIndex = 0
End Sub

Private Sub pLoadChord()
On Error GoTo Err_pLoadChord:
Dim sTextLine As String
   m_ChordCount = 0
   ReDim Preserve m_Chords(1 To 1)
   Open App.Path & "\ChordData.txt" For Input As #1
   Do While Not EOF(1)
      Line Input #1, sTextLine
      If Left(sTextLine, 2) = IIf(bFavorites, "!F", "!C") Then 'if sTextLine is ChordLine
         m_ChordCount = m_ChordCount + 1
         ReDim Preserve m_Chords(1 To m_ChordCount)
         With m_Chords(m_ChordCount)
            .lChordID = Val(Mid(sTextLine, 4, 4))
            .sGroupName = Trim(Mid(sTextLine, 9, 10))
            .sChordName = Trim(Mid(sTextLine, 20, 10))
            .iLine1 = Val(Mid(sTextLine, 31, 2))
            .iLine2 = Val(Mid(sTextLine, 34, 2))
            .iLine3 = Val(Mid(sTextLine, 37, 2))
            .iLine4 = Val(Mid(sTextLine, 40, 2))
            .iLine5 = Val(Mid(sTextLine, 43, 2))
            .iLine6 = Val(Mid(sTextLine, 46, 2))
            .sDescription = Trim(Mid(sTextLine, 49, 20))
            .iBarreFirst = Val(Mid(sTextLine, 70, 1))
            .iBarreLast = Val(Mid(sTextLine, 72, 1))
            .iBarreFret = Val(Mid(sTextLine, 74, 2))
            .iFinger1 = Val(Mid(sTextLine, 77, 2))
            .iFinger2 = Val(Mid(sTextLine, 80, 2))
            .iFinger3 = Val(Mid(sTextLine, 83, 2))
            .iFinger4 = Val(Mid(sTextLine, 86, 2))
            .iFinger5 = Val(Mid(sTextLine, 89, 2))
            .iFinger6 = Val(Mid(sTextLine, 92, 2))
         End With 'With m_Chords(m_ChordCount)
      End If 'If Left(sTextLine, 2) = "!C" Then 'if sTextLine is ChordLine
   Loop 'Do While Not EOF(1)
   Close #1
Exit Sub
Err_pLoadChord:
   MsgBox "Akor bilgilerinin bulunduðu ChordData.txt dosyasý bulunamadý.", vbCritical, ""
End Sub

Private Sub pPartSetting()
Dim i As Long, lTop As Long, lLeft As Long
   lLeft = 240
   lTop = 200

   imgLT.Top = lTop
   imgLT.Left = lLeft
   imgLT.Picture = LoadResPicture(104, vbResBitmap)
   lLeft = lLeft + imgLT.Width
   For i = 1 To 12
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
   lblE.Left = lLeft '- 300
   lblE.Top = lTop + 30
   lLeft = lLeft + E0.Width
   For i = 1 To 12
      E(i).Picture = pImgPart01
      E(i).Left = lLeft
      E(i).Top = lTop
      'E(i).MouseIcon = pImgCur01
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
   lblB.Left = lLeft '- 300
   lblB.Top = lTop + 30
   lLeft = lLeft + B0.Width
   For i = 1 To 12
      B(i).Picture = pImgPart01
      B(i).Left = lLeft
      B(i).Top = lTop
      'B(i).MouseIcon = pImgCur01
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
   lblG.Left = lLeft '- 300
   lblG.Top = lTop + 30
   lLeft = lLeft + G0.Width
   For i = 1 To 12
      G(i).Picture = pImgPart01
      G(i).Left = lLeft
      G(i).Top = lTop
      'G(i).MouseIcon = pImgCur01
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
   lblD.Left = lLeft '- 300
   lblD.Top = lTop + 30
   lLeft = lLeft + D0.Width
   For i = 1 To 12
      D(i).Picture = pImgPart01
      D(i).Left = lLeft
      D(i).Top = lTop
      'D(i).MouseIcon = pImgCur01
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
   lblA.Left = lLeft '- 300
   lblA.Top = lTop + 30
   lLeft = lLeft + A0.Width
   For i = 1 To 12
      A(i).Picture = pImgPart01
      A(i).Left = lLeft
      A(i).Top = lTop
      'A(i).MouseIcon = pImgCur01
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
   lblEs.Left = lLeft '- 300
   lblEs.Top = lTop + 30
   lLeft = lLeft + Es0.Width
   For i = 1 To 12
      Es(i).Picture = pImgPart01
      Es(i).Left = lLeft
      Es(i).Top = lTop
      'Es(i).MouseIcon = pImgCur01
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
   For i = 1 To 12
      imgBottom(i).Left = lLeft
      imgBottom(i).Top = lTop
      imgBottom(i).Picture = LoadResPicture(108, vbResBitmap)
      lLeft = lLeft + imgBottom(i).Width
   Next
   imgRB.Top = lTop
   imgRB.Left = lLeft
   imgRB.Picture = LoadResPicture(111, vbResBitmap)

End Sub

Private Sub pShowChord()
Dim i As Long
Dim bEisBarre As Boolean, bBisBarre As Boolean, bGisBarre As Boolean, bDisBarre As Boolean, bAisBarre As Boolean, bESisBarre As Boolean
If m_ChordCount < 1 Or fGetIndex < 1 Then Exit Sub
'Clear
   If E0.Appearance = 0 Then
      E0.Appearance = 1
      E0.Picture = LoadResPicture(105, vbResBitmap)
   End If
   If B0.Appearance = 0 Then
      B0.Appearance = 1
      B0.Picture = LoadResPicture(105, vbResBitmap)
   End If
   If G0.Appearance = 0 Then
      G0.Appearance = 1
      G0.Picture = LoadResPicture(105, vbResBitmap)
   End If
   If D0.Appearance = 0 Then
      D0.Appearance = 1
      D0.Picture = LoadResPicture(105, vbResBitmap)
   End If
   If A0.Appearance = 0 Then
      A0.Appearance = 1
      A0.Picture = LoadResPicture(105, vbResBitmap)
   End If
   If Es0.Appearance = 0 Then
      Es0.Appearance = 1
      Es0.Picture = LoadResPicture(105, vbResBitmap)
   End If

For i = 1 To 12
   If E(i).Appearance = 0 Then
      E(i).Appearance = 1
      E(i).Picture = pImgPart01
   End If
   If B(i).Appearance = 0 Then
      B(i).Appearance = 1
      B(i).Picture = pImgPart01
   End If
   If G(i).Appearance = 0 Then
      G(i).Appearance = 1
      G(i).Picture = pImgPart01
   End If
   If D(i).Appearance = 0 Then
      D(i).Appearance = 1
      D(i).Picture = pImgPart01
   End If
   If A(i).Appearance = 0 Then
      A(i).Appearance = 1
      A(i).Picture = pImgPart01
   End If
   If Es(i).Appearance = 0 Then
      Es(i).Appearance = 1
      Es(i).Picture = pImgPart01
   End If
Next

With m_Chords(fGetIndex)
   'Chord
   If .iLine1 > 0 Then
      E(.iLine1).Picture = pImgPartSel
      E(.iLine1).Appearance = 0
   ElseIf .iLine1 = -1 Then
      E0.Picture = pImgPartNP
      E0.Appearance = 0
   End If
   If .iLine2 > 0 Then
      B(.iLine2).Picture = pImgPartSel
      B(.iLine2).Appearance = 0
   ElseIf .iLine2 = -1 Then
      B0.Picture = pImgPartNP
      B0.Appearance = 0
   End If
   If .iLine3 > 0 Then
      G(.iLine3).Picture = pImgPartSel
      G(.iLine3).Appearance = 0
   ElseIf .iLine3 = -1 Then
      G0.Picture = pImgPartNP
      G0.Appearance = 0
   End If
   If .iLine4 > 0 Then
      D(.iLine4).Picture = pImgPartSel
      D(.iLine4).Appearance = 0
   ElseIf .iLine4 = -1 Then
      D0.Picture = pImgPartNP
      D0.Appearance = 0
   End If
   If .iLine5 > 0 Then
      A(.iLine5).Picture = pImgPartSel
      A(.iLine5).Appearance = 0
   ElseIf .iLine5 = -1 Then
      A0.Picture = pImgPartNP
      A0.Appearance = 0
   End If
   If .iLine6 > 0 Then
      Es(.iLine6).Picture = pImgPartSel
      Es(.iLine6).Appearance = 0
   ElseIf .iLine6 = -1 Then
      Es0.Picture = pImgPartNP
      Es0.Appearance = 0
   End If

   'Barre
   If .iBarreFret > 0 Then
      'E
      If .iBarreFirst = 1 Then
         E(.iBarreFret).Picture = pImgPartBarreF
         E(.iBarreFret).Appearance = 0
      ElseIf .iBarreFirst < 1 And .iBarreLast > 1 Then
      ElseIf .iBarreLast = 1 Then
      End If
      'B
      If .iBarreFirst = 2 Then
         B(.iBarreFret).Picture = pImgPartBarreF
         B(.iBarreFret).Appearance = 0
      ElseIf .iBarreFirst < 2 And .iBarreLast > 2 Then
         B(.iBarreFret).Picture = pImgPartBarreM
         B(.iBarreFret).Appearance = 0
      ElseIf .iBarreLast = 2 Then
         B(.iBarreFret).Picture = pImgPartBarreL
         B(.iBarreFret).Appearance = 0
      End If
      'G
      If .iBarreFirst = 3 Then
         G(.iBarreFret).Picture = pImgPartBarreF
         G(.iBarreFret).Appearance = 0
      ElseIf .iBarreFirst < 3 And .iBarreLast > 3 Then
         G(.iBarreFret).Picture = pImgPartBarreM
         G(.iBarreFret).Appearance = 0
      ElseIf .iBarreLast = 3 Then
         G(.iBarreFret).Picture = pImgPartBarreL
         G(.iBarreFret).Appearance = 0
      End If
      'D
      If .iBarreFirst = 4 Then
         D(.iBarreFret).Picture = pImgPartBarreF
         D(.iBarreFret).Appearance = 0
      ElseIf .iBarreFirst < 4 And .iBarreLast > 4 Then
         D(.iBarreFret).Picture = pImgPartBarreM
         D(.iBarreFret).Appearance = 0
      ElseIf .iBarreLast = 4 Then
         D(.iBarreFret).Picture = pImgPartBarreL
         D(.iBarreFret).Appearance = 0
      End If
      'A
      If .iBarreFirst = 5 Then
         A(.iBarreFret).Picture = pImgPartBarreF
         A(.iBarreFret).Appearance = 0
      ElseIf .iBarreFirst < 5 And .iBarreLast > 5 Then
         A(.iBarreFret).Picture = pImgPartBarreM
         A(.iBarreFret).Appearance = 0
      ElseIf .iBarreLast = 5 Then
         A(.iBarreFret).Picture = pImgPartBarreL
         A(.iBarreFret).Appearance = 0
      End If
      'ES
      If .iBarreFirst = 6 Then
         Es(.iBarreFret).Picture = pImgPartBarreF
         Es(.iBarreFret).Appearance = 0
      ElseIf .iBarreFirst < 6 And .iBarreLast > 6 Then
         Es(.iBarreFret).Picture = pImgPartBarreM
         Es(.iBarreFret).Appearance = 0
      ElseIf .iBarreLast = 6 Then
         Es(.iBarreFret).Picture = pImgPartBarreL
         Es(.iBarreFret).Appearance = 0
      End If
      
   End If

   lblDesc.Caption = .sDescription
End With
End Sub
