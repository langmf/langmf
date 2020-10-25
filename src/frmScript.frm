VERSION 5.00
Begin VB.Form frmScript 
   BorderStyle     =   0  'None
   Caption         =   "frmScript"
   ClientHeight    =   612
   ClientLeft      =   -3000
   ClientTop       =   -3000
   ClientWidth     =   600
   ControlBox      =   0   'False
   Icon            =   "frmScript.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   51
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   50
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
End
Attribute VB_Name = "frmScript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private SubClass As New clsSubClass

Public CScript As New Collection


Public Function ActiveScript_Error(ByVal Obj As clsActiveScript) As Variant
    If (Not frmError.Visible) And (Not mf_IsEnd) Then
        If Not mf_NoError Then frmError.Display Obj:    EndMF
    End If
    ActiveScript_Error = True
End Function

Private Sub Form_Activate()
    Me.Visible = False
End Sub

Private Sub Form_Load()
    Me.WindowState = 0
    Me.Visible = False
    Me.Move Screen.Width / 3, Screen.Height / 4, 1, 1
    Me.Visible = False
    SubClass.Msg(WM_QUERYENDSESSION) = 1
    SubClass.HookSet Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SubClass.HookClear Me
End Sub

Public Function WindowProc(ByRef bHandled As Boolean, ByVal u_hWnd As Long, ByVal uMsg As Long, ByRef wParam As Long, ByRef lParam As Long, ByVal dwRefData As Long) As Long
    On Error Resume Next
    If uMsg = WM_QUERYENDSESSION Then WindowProc = CLng(CAS.CodeObject.QueryEndSession(lParam))
End Function


