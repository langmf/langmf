VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPowerATX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public Sub Control(ByVal Flags As Long)
    On Error Resume Next
    AdjustToken "SeShutdownPrivilege"
    ExitWindowsEx Flags, &HFFFF&
End Sub

Public Sub POWEROFF(Optional ByVal Flags As Long = EWX_FORCE)
    Call Control(EWX_POWEROFF Or Flags)
End Sub

Public Sub REBOOT(Optional ByVal Flags As Long = EWX_FORCE)
    Call Control(EWX_REBOOT Or Flags)
End Sub

Public Sub LOGOFF(Optional ByVal Flags As Long = EWX_FORCE)
    Call Control(EWX_LOGOFF Or Flags)
End Sub

Public Sub SUSPEND(Optional ByVal forceCrit As Long = 1, Optional ByVal disWakeEvents As Long = 0, Optional ByVal hibern As Long = 0)
    AdjustToken "SeShutdownPrivilege"
    SetSuspendState hibern, forceCrit, disWakeEvents
End Sub

Public Sub Monitor(Optional ByVal Flags As Long = -1)
    SendMessageW frmScript.hWnd, WM_SYSCOMMAND, SC_MONITORPOWER, ByVal Flags
End Sub


