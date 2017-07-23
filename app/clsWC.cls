VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


'=========================================================================
Const PBM_SETRANGE = &H401, PBM_SETPOS = &H402, PBM_DELTAPOS = &H403, PBM_SETSTEP = &H404
Const PBM_STEPIT = &H405, PBM_SETRANGE32 = &H406, PBM_GETRANGE = &H407, PBM_GETPOS = &H408
Const PBM_SETBARCOLOR = &H409, PBM_SETBKCOLOR = 8193, PBM_SETMARQUEE = &H40A  'XP
Const PBM_GETSTEP = &H40D, PBM_GETBKCOLOR = &H40E, PBM_GETBARCOLOR = &H40F  'VISTA
Const PBM_SETSTATE = &H410, PBM_GETSTATE = &H411  'VISTA
Const PBS_SMOOTH = &H1, PBS_VERTICAL = &H4, PBS_MARQUEE = &H8, PBS_SMOOTHREVERSE = &H10  'VISTA
Const PBST_NORMAL = &H1, PBST_ERROR = &H2, PBST_PAUSED = &H3  'VISTA

Private lPos As Long, lMin As Long, lMax As Long


'=========================================================================
Private lTop As Long, lLeft As Long, lWidth As Long, lHeight As Long, hFont As Long, vColor As Long, vText As String

Public Name As String, Class As String, hWnd As Long, hDC As Long, onWindowProc As String
Public Parent As Object, Bag As Variant, SubClass As Object
Public WithEvents Font As StdFont
Attribute Font.VB_VarHelpID = -1


Public Function Create(ByVal objParent As Object, ByVal txtClass As String, Optional ByVal txtName As String, Optional ByVal txtCaption As String, Optional ByVal dwStyle As Long, Optional ByVal dwExStyle As Long) As Long
    Name = txtName
    Class = txtClass
    Set Parent = objParent

    lLeft = 0: lTop = 0: lWidth = 100: lHeight = 25
    hWnd = CreateWindowExW(dwExStyle, StrPtr(txtClass), StrPtr(txtCaption), WS_CHILD Or WS_VISIBLE Or dwStyle, lLeft, lTop, lWidth, lHeight, Parent.hWnd, 0, App.hInstance, 0)
    vColor = GetSysColor(COLOR_WINDOWTEXT)
    Set Font = SYS.Ext.CreateStdFont
    hDC = GetDC(hWnd)
    Create = hWnd
    
    Call Init
End Function

Public Sub Destroy()
    If hWnd <> 0 Then
        If hFont <> 0 Then DeleteObject hFont
        SubClass.HookClear Me
        ReleaseDC hWnd, hDC
        DestroyWindow hWnd
    End If
End Sub

Public Function Msg(ByVal wMsg As Long, Optional ByVal wParam As Long, Optional ByVal lParam As Long) As Long
    Msg = SendMessageW(hWnd, wMsg, wParam, lParam)
End Function

Public Function Win(Optional ByVal mStyle As Long, Optional ByVal value As Variant, Optional ByVal nIndex As Long = GWL_STYLE) As Long
    Win = WindowLong(hWnd, value, mStyle, nIndex)
End Function

Public Sub Move(Optional ByVal vLeft As Single = -1000, Optional ByVal vTop As Single = -1000, Optional ByVal vWidth As Single = -1000, Optional ByVal vHeight As Single = -1000)
    With Parent
        If vLeft <> -1000 Then lLeft = .ScaleX(vLeft, .ScaleMode, vbPixels)
        If vTop <> -1000 Then lTop = .ScaleY(vTop, .ScaleMode, vbPixels)
        If vWidth <> -1000 Then lWidth = .ScaleX(vWidth, .ScaleMode, vbPixels)
        If vHeight <> -1000 Then lHeight = .ScaleY(vHeight, .ScaleMode, vbPixels)
    End With
    Resize
End Sub

Public Property Get Left() As Single: Left = Parent.ScaleX(lLeft, vbPixels, Parent.ScaleMode): End Property                             'no_check
Public Property Let Left(ByVal value As Single): lLeft = Parent.ScaleX(value, Parent.ScaleMode, vbPixels): Resize: End Property
Public Property Get Top() As Single: Top = Parent.ScaleX(lTop, vbPixels, Parent.ScaleMode): End Property                                'no_check
Public Property Let Top(ByVal value As Single): lTop = Parent.ScaleY(value, Parent.ScaleMode, vbPixels): Resize: End Property
Public Property Get Width() As Single: Width = Parent.ScaleX(lWidth, vbPixels, Parent.ScaleMode): End Property                          'no_check
Public Property Let Width(ByVal value As Single): lWidth = Parent.ScaleX(value, Parent.ScaleMode, vbPixels): Resize: End Property
Public Property Get Height() As Single: Height = Parent.ScaleX(lHeight, vbPixels, Parent.ScaleMode): End Property                       'no_check
Public Property Let Height(ByVal value As Single): lHeight = Parent.ScaleY(value, Parent.ScaleMode, vbPixels): Resize: End Property
Public Property Get PText() As String: PText = vText: End Property                                                                      'no_check
Public Property Let PText(ByVal value As String): vText = value: Refresh: End Property
Public Property Get PColor() As Long: PColor = vColor: End Property                                                                     'no_check
Public Property Let PColor(ByVal value As Long): vColor = value: Refresh: End Property
Public Property Get Visible() As Boolean: Visible = WindowLong(hWnd, , WS_VISIBLE): End Property                                      'no_check
Public Property Let Visible(ByVal value As Boolean): Call WindowLong(hWnd, value, WS_VISIBLE): End Property
Public Sub Refresh(Optional ByVal mStyle As Long = SW_SHOW): ShowWindow hWnd, mStyle: End Sub                                           'no_check
Public Sub SetFocus(): Call modConst.SetFocus(hWnd): End Sub                                                                            'no_check


Public Function WindowProc(ByRef bHandled As Boolean, ByVal u_hWnd As Long, ByVal uMsg As Long, ByRef wParam As Long, ByRef lParam As Long, ByVal dwRefData As Long) As Long              'no_check
    Dim mShift As Long, mButton As Long, mx As Long, my As Long
    
    On Error Resume Next
            
    mShift = (wParam And 15) \ 4
    mButton = (wParam And 3) + (wParam And 16) \ 4
    mx = lParam And &HFFFF&
    my = lParam \ &H10000
    
    Select Case uMsg
        Case WM_PAINT
            SubClass.HookDefault u_hWnd, uMsg, wParam, lParam
            bHandled = True
            RenderText

        Case WM_MOUSEMOVE
            Parent.Events Name & "_MouseMove", mButton, mShift, mx, my
            
        Case WM_LBUTTONDOWN, WM_MBUTTONDOWN, WM_RBUTTONDOWN
            Parent.Events Name & "_MouseDown", mButton, mShift, mx, my
        
        Case WM_LBUTTONUP, WM_MBUTTONUP, WM_RBUTTONUP
            Parent.Events Name & "_MouseUp", mButton, mShift, mx, my
            
        Case WM_CHAR
            Parent.Events Name & "_KeyPress", wParam
    End Select
    
    If LenB(onWindowProc) Then WindowProc = Parent.Events(CStr(onWindowProc), Deref(VarPtr(u_hWnd) - 4), u_hWnd, uMsg, wParam, lParam, dwRefData)
End Function

Private Sub Resize()
    If hWnd <> 0 Then Call SetWindowPos(hWnd, 0, lLeft, lTop, lWidth, lHeight, SWP_NOZORDER)
End Sub

Private Sub Font_FontChanged(ByVal PropertyName As String)
    If hWnd <> 0 Then
        If hFont <> 0 Then DeleteObject hFont
        hFont = CreateFontIndirectW(StdFontToLogFont(Font))
        If (hFont <> 0) Then SendMessageW hWnd, WM_SETFONT, hFont, ByVal 1&
    End If
End Sub

Private Sub Class_Initialize()
    Set SubClass = New clsSubClass
End Sub

Private Sub Class_Terminate()
    If hWnd <> 0 Then Destroy
End Sub
'=========================================================================


Private Sub Init()
    Select Case Class
        Case "msctls_progress32":     SetRange 0, 100:  PText = vbNullChar + "0"
    End Select
    Call Font_FontChanged("")
    SubClass.List(WM_PAINT, WM_CHAR, WM_MOUSEMOVE, WM_LBUTTONDOWN, WM_MBUTTONDOWN, WM_RBUTTONDOWN, WM_LBUTTONUP, WM_MBUTTONUP, WM_RBUTTONUP) = 1
    SubClass.HookSet Me
End Sub

Private Sub RenderText()
    Dim txt As String
    If LenB(vText) = 0 Then Exit Sub
    txt = vText
    If Left$(txt, 1) = vbNullChar Then txt = Replace$(txt, vbNullChar, ""): txt = Format$((lPos / (lMax - lMin)) * 100, txt) & " %"
    PrintText hWnd, hDC, txt, vColor, Font
End Sub


'=========================================================================
Public Property Get Pos() As Long:              Pos = SendMessageW(hWnd, PBM_GETPOS, 0, 0):                  End Property                                                      'no_check
Public Property Let Pos(ByVal value As Long):   lPos = value:   SendMessageW hWnd, PBM_SETPOS, value, 0:     End Property
Public Property Get Step() As Long:             Step = SendMessageW(hWnd, PBM_GETSTEP, 0, 0):                End Property                                                      'no_check
Public Property Let Step(ByVal value As Long):  SendMessageW hWnd, PBM_SETSTEP, value, 0:                    End Property
Public Property Get State() As Long:            State = SendMessageW(hWnd, PBM_GETSTATE, 0, 0):              End Property                                                      'no_check
Public Property Let State(ByVal value As Long): SendMessageW hWnd, PBM_SETSTATE, value, 0:                   End Property
Public Sub StepIt():                            SendMessageW hWnd, PBM_STEPIT, 0, 0:                         End Sub                                                           'no_check
Public Sub SetRange(ByVal m_Min As Long, ByVal m_Max As Long):  lMin = m_Min: lMax = m_Max: SendMessageW hWnd, PBM_SETRANGE, 0, m_Max * 65536 + m_Min:   End Sub               'no_check
