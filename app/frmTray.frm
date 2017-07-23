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

Private Declare Function Shell_NotifyIcon Lib "Shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

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


Private Const NIM_ADD = 0&
Private Const NIM_MODIFY = 1&
Private Const NIM_DELETE = 2&
Private Const NIM_SETFOCUS = 3&
Private Const NIM_SETVERSION = 4&

Private Const NIF_MESSAGE = 1&
Private Const NIF_ICON = 2&
Private Const NIF_TIP = 4&
Private Const NIF_INFO = &H10&

Private Const WM_USER = &H400&

Private Const WM_MOUSEMOVE = &H200&
Private Const WM_LBUTTONDBLCLK = &H203&
Private Const WM_LBUTTONDOWN = &H201&
Private Const WM_LBUTTONUP = &H202&
Private Const WM_RBUTTONDBLCLK = &H206&
Private Const WM_RBUTTONDOWN = &H204&
Private Const WM_RBUTTONUP = &H205&
Private Const WM_MBUTTONDBLCLK = &H209&
Private Const WM_MBUTTONDOWN = &H207&
Private Const WM_MBUTTONUP = &H208&

Private Const NIN_BALLOONSHOW = (WM_USER + 2)
Private Const NIN_BALLOONHIDE = (WM_USER + 3)
Private Const NIN_BALLOONTIMEOUT = (WM_USER + 4)
Private Const NIN_BALLOONUSERCLICK = (WM_USER + 5)

Private Enum EBalloonIconTypes
   NIIF_NONE = 0&
   NIIF_INFO = 1&
   NIIF_WARNING = 2&
   NIIF_ERROR = 3&
   NIIF_NOSOUND = &H10&
End Enum

Private gInTray As Boolean
Private gTrayTip As String
Private gTrayIcon As StdPicture
Private gAddedToTray As Boolean
Private gBalloon As Boolean

Private def_TrayIcon As StdPicture

Public Parent As frmForm
Public Anim As New Collection
Public Counter As Long


Public Sub Balloon(Optional ByVal strMessage As String, Optional ByVal strTitle As String = "Info", Optional ByVal typeIcon As Integer = 1)
    Dim Tray As NOTIFYICONDATA

    If LenB(strMessage) > 0 And gBalloon = True Then Exit Sub
    
    With Tray
        .uFlags = NIF_INFO
        .szInfo = strMessage + vbNullChar
        .szInfoTitle = strTitle + vbNullChar
        .uTimeOutOrVersion = 30000
        .dwInfoFlags = typeIcon
    End With

    Call Send(NIM_MODIFY, Tray)
End Sub

Public Sub TrayPlay()
    If Timer.Enabled Or Anim.Count = 0 Then Exit Sub
    If def_TrayIcon Is Nothing Then Set def_TrayIcon = gTrayIcon
    Set TrayIcon = Anim.Item(1)
    Timer.Enabled = True
End Sub

Public Sub TrayStop()
    Timer.Enabled = False
    If Not def_TrayIcon Is Nothing Then Set TrayIcon = def_TrayIcon: Set def_TrayIcon = Nothing
    Counter = 1
End Sub

Public Property Set TrayIcon(Ico As StdPicture)
    Dim Tray As NOTIFYICONDATA

    If Not (Ico Is Nothing) Then
        If (Ico.Type = vbPicTypeIcon) Then
            If gAddedToTray Then
                Tray.hIcon = Ico
                Tray.uFlags = NIF_ICON
                Call Send(NIM_MODIFY, Tray)
            End If
    
            Set gTrayIcon = Ico
        End If
    End If
End Property

Public Property Get TrayIcon() As StdPicture
    Set TrayIcon = gTrayIcon
End Property

Public Property Let TrayTip(ByVal Tip As String)
    Dim Tray As NOTIFYICONDATA

    If gAddedToTray Then
        Tray.szTip = Tip & vbNullChar
        Tray.uFlags = NIF_TIP
        Call Send(NIM_MODIFY, Tray)
    End If
    
    gTrayTip = Tip
End Property

Public Property Get TrayTip() As String
    TrayTip = gTrayTip
End Property

Public Property Let InTray(ByVal Show As Boolean)
    If (Show <> gInTray) Then
        If Show Then
                AddIcon TrayTip, TrayIcon
                gAddedToTray = True
        Else
            If gAddedToTray Then
                DeleteIcon
                gAddedToTray = False
            End If
        End If
        
        gInTray = Show
    End If
End Property

Public Property Get InTray() As Boolean
    InTray = gInTray
End Property

Private Sub Timer_Timer()
    On Error Resume Next
    If def_TrayIcon Is Nothing Then Set def_TrayIcon = gTrayIcon
    Parent.Events "Tray_Anim"
    If Counter < 1 Then Counter = 1
    If Counter > Anim.Count - 1 Then Counter = 1
    Set TrayIcon = Anim(Counter)
    Counter = Counter + 1
End Sub

Private Sub AddIcon(Tip As String, Ico As StdPicture)
    Dim Tray As NOTIFYICONDATA

    With Tray
        .uFlags = NIF_MESSAGE
    
        If Not (Ico Is Nothing) Then
            .hIcon = Ico
            .uFlags = .uFlags Or NIF_ICON
            Set gTrayIcon = Ico
        End If
        
        If LenB(Tip) Then
            .szTip = Tip & vbNullChar
            .uFlags = .uFlags Or NIF_TIP
            gTrayTip = Tip
        End If
    End With

    Call Send(NIM_ADD, Tray)
End Sub

Private Sub DeleteIcon()
    Dim Tray As NOTIFYICONDATA
    Call Send(NIM_DELETE, Tray)
End Sub

Private Function Send(dwMsg As Long, Tray As NOTIFYICONDATA) As Long
    With Tray
        .hWnd = Me.hWnd
        .uCallbackMessage = WM_MOUSEMOVE
        .cbSize = Len(Tray)
    End With
    Send = Shell_NotifyIcon(dwMsg, Tray)
End Function

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
         
        Case NIN_BALLOONSHOW:       Parent.Events "Tray_BalloonShow": gBalloon = True
        Case NIN_BALLOONHIDE:       Parent.Events "Tray_BalloonHide": gBalloon = False
        Case NIN_BALLOONTIMEOUT:    Parent.Events "Tray_BalloonTimeout": gBalloon = False
        Case NIN_BALLOONUSERCLICK:  Parent.Events "Tray_BalloonClick": gBalloon = False
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    InTray = False
End Sub
