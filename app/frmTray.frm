VERSION 5.00
Begin VB.Form frmTray 
   BorderStyle     =   0  'None
   ClientHeight    =   612
   ClientLeft      =   -996
   ClientTop       =   -996
   ClientWidth     =   588
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   612
   ScaleWidth      =   588
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   120
      Top             =   120
   End
End
Attribute VB_Name = "frmTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Private Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 128
    dwState As Long
    dwStateMask As Long
    szInfo As String * 256
    uTimeOutOrVersion As Long
    szInfoTitle As String * 64
    dwInfoFlags As Long
End Type

Private Const NIB_SHOW = (WM_USER + 2), NIB_HIDE = (WM_USER + 3), NIB_TIMEOUT = (WM_USER + 4), NIB_USERCLICK = (WM_USER + 5)
Private Const NIIF_NONE = 0&, NIIF_INFO = 1&, NIIF_WARNING = 2&, NIIF_ERROR = 3&, NIIF_NOSOUND = &H10&
Private Const NIM_ADD = 0&, NIM_MODIFY = 1&, NIM_DELETE = 2&, NIM_SETFOCUS = 3&, NIM_SETVERSION = 4&
Private Const NIF_MESSAGE = 1&, NIF_ICON = 2&, NIF_TIP = 4&, NIF_INFO = &H10&

Private m_Show As Boolean, m_Tip As String, m_Visible As Boolean, m_Balloon As Boolean
Private m_Icon As StdPicture, m_IconDef As StdPicture

Public Parent As frmForm, Anim As New Collection, Counter As Long


Public Sub Balloon(Optional ByVal strMessage As String, Optional ByVal strTitle As String = "Info", Optional ByVal typeIcon As Long = 1)
    Dim Tray As NOTIFYICONDATA

    If LenB(strMessage) > 0 And m_Balloon = True Then Exit Sub
    
    With Tray
        .uFlags = NIF_INFO
        .szInfo = strMessage + vbNullChar
        .szInfoTitle = strTitle + vbNullChar
        .uTimeOutOrVersion = 30000
        .dwInfoFlags = typeIcon
    End With

    Call Send(Tray)
End Sub

Public Sub TrayPlay()
    If Timer.Enabled Or Anim.Count = 0 Then Exit Sub
    If m_IconDef Is Nothing Then Set m_IconDef = m_Icon
    Set TrayIcon = Anim.Item(1)
    Timer.Enabled = True
End Sub

Public Sub TrayStop()
    Timer.Enabled = False
    If Not m_IconDef Is Nothing Then Set TrayIcon = m_IconDef:  Set m_IconDef = Nothing
    Counter = 1
End Sub

Public Property Set TrayIcon(vIcon As StdPicture)
    Dim Tray As NOTIFYICONDATA
    If Not (vIcon Is Nothing) Then
        If (vIcon.Type = vbPicTypeIcon) Then
            If m_Visible Then Tray.hIcon = vIcon:       Tray.uFlags = NIF_ICON:      Call Send(Tray)
            Set m_Icon = vIcon
        End If
    End If
End Property

Public Property Get TrayIcon() As StdPicture
    Set TrayIcon = m_Icon
End Property

Public Property Let TrayTip(ByVal vTip As String)
    Dim Tray As NOTIFYICONDATA
    If m_Visible Then Tray.szTip = vTip & vbNullChar:    Tray.uFlags = NIF_TIP:     Call Send(Tray)
    m_Tip = vTip
End Property

Public Property Get TrayTip() As String
    TrayTip = m_Tip
End Property

Public Property Let InTray(ByVal vShow As Boolean)
    If (vShow <> m_Show) Then
        If vShow Then
            AddIcon TrayTip, TrayIcon
            m_Visible = True
        Else
            If m_Visible Then
                DeleteIcon
                m_Visible = False
            End If
        End If
        
        m_Show = vShow
    End If
End Property

Public Property Get InTray() As Boolean
    InTray = m_Show
End Property

Private Sub AddIcon(vTip As String, vIcon As StdPicture)
    Dim Tray As NOTIFYICONDATA
    With Tray
        .uFlags = NIF_MESSAGE
        If Not (vIcon Is Nothing) Then .hIcon = vIcon:      .uFlags = .uFlags Or NIF_ICON:      Set m_Icon = vIcon
        If LenB(vTip) Then .szTip = vTip & vbNullChar:      .uFlags = .uFlags Or NIF_TIP:       m_Tip = vTip
    End With
    Call Send(Tray, NIM_ADD)
End Sub

Private Sub DeleteIcon()
    Dim Tray As NOTIFYICONDATA
    Call Send(Tray, NIM_DELETE)
End Sub

Private Function Send(Tray As NOTIFYICONDATA, Optional ByVal dwMsg As Long = NIM_MODIFY) As Long
    With Tray:      .cbSize = Len(Tray):    .hWnd = Me.hWnd:    .uCallbackMessage = WM_MOUSEMOVE:       End With
    Send = Shell_NotifyIcon(dwMsg, Tray)
End Function

Private Sub Timer_Timer()
    On Error Resume Next
    If m_IconDef Is Nothing Then Set m_IconDef = m_Icon
    Parent.Events "Tray_Anim"
    If Counter < 1 Then Counter = 1
    If Counter > Anim.Count - 1 Then Counter = 1
    Set TrayIcon = Anim(Counter)
    Counter = Counter + 1
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lX As Long
    lX = ScaleX(x, Me.ScaleMode, vbPixels)
    Select Case lX
        Case WM_MOUSEMOVE:          Parent.Events "Tray_MouseMove"
        
        Case WM_LBUTTONDOWN:        Parent.Events "Tray_MouseDown", vbLeftButton
        Case WM_LBUTTONUP:          Parent.Events "Tray_MouseUp", vbLeftButton
        Case WM_LBUTTONDBLCLK:      Parent.Events "Tray_DblClick", vbLeftButton
        
        Case WM_RBUTTONDOWN:        Parent.Events "Tray_MouseDown", vbRightButton
        Case WM_RBUTTONUP:          Parent.Events "Tray_MouseUp", vbRightButton
        Case WM_RBUTTONDBLCLK:      Parent.Events "Tray_DblClick", vbRightButton
        
        Case WM_MBUTTONDOWN:        Parent.Events "Tray_MouseDown", vbMiddleButton
        Case WM_MBUTTONUP:          Parent.Events "Tray_MouseUp", vbMiddleButton
        Case WM_MBUTTONDBLCLK:      Parent.Events "Tray_DblClick", vbMiddleButton
         
        Case NIB_SHOW:              Parent.Events "Tray_BalloonShow":     m_Balloon = True
        Case NIB_HIDE:              Parent.Events "Tray_BalloonHide":     m_Balloon = False
        Case NIB_TIMEOUT:           Parent.Events "Tray_BalloonTimeout":  m_Balloon = False
        Case NIB_USERCLICK:         Parent.Events "Tray_BalloonClick":    m_Balloon = False
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    InTray = False
End Sub
