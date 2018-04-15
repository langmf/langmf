Attribute VB_Name = "modXPControl"
Option Explicit

Private Type tagInitCommonControlsEx
    lngSize As Long
    lngICC As Long
End Type

Private Declare Function InitCommonControlsEx Lib "comctl32" (iccex As tagInitCommonControlsEx) As Boolean

Sub InitCommonControlsXP()
    Dim iccex As tagInitCommonControlsEx
    
    On Error GoTo err1
    
    With iccex
      .lngSize = Len(iccex)
      .lngICC = &H7FFF&
    End With
    
    InitCommonControlsEx iccex
err1:
End Sub
