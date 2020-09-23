Attribute VB_Name = "mMain"
'Halil TUNC © 2004
'HtGuitar ...\Module\mMain.bas
'19/09/2004 Ankara TURKEY
'halil_tunc@hotmail.com

Option Explicit

Private Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type
Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean
Private Const ICC_USEREX_CLASSES = &H200

Public Sub Main()
   On Error Resume Next
   Dim iccex As tagInitCommonControlsEx
   With iccex
       .lngSize = LenB(iccex)
       .lngICC = ICC_USEREX_CLASSES
   End With
   InitCommonControlsEx iccex
   On Error GoTo 0
   frmMain.Show
End Sub




