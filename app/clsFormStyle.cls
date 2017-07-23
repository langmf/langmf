VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFormStyle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private ITB As TaskbarList

Private m_MaxButt As Boolean, m_MinButt As Boolean, m_CtrlBox As Boolean, m_SizeBrd As Boolean
Private m_TaskBar As Boolean, m_TitleBar As Boolean, m_ToolWin As Boolean, m_CloseButt As Boolean

Public hWnd As Long


Public Function CreateNew() As Object
    Set CreateNew = New clsFormStyle
End Function

Public Property Get CloseButton() As Boolean
    CloseButton = m_CloseButt
End Property

Public Property Let CloseButton(ByVal vData As Boolean)
    Dim hSysMenu As Long, nIndex As Long, Frm As Form
    
    m_CloseButt = vData
    
    hSysMenu = GetSystemMenu(hWnd, vData)
    nIndex = GetMenuItemCount(hSysMenu)
    
    Call RemoveMenu(hSysMenu, nIndex - 1, MF_BYPOSITION)
    Call RemoveMenu(hSysMenu, nIndex - 2, MF_BYPOSITION)
    Call SetMenu(hWnd, hSysMenu)
    
    For Each Frm In Forms
        If Frm.hWnd = hWnd Then Frm.Refresh
    Next
End Property

Public Property Get ToolWindow() As Boolean
    ToolWindow = m_ToolWin
End Property

Public Property Let ToolWindow(ByVal vData As Boolean)
    m_ToolWin = vData
    SetWindowStyle True, WS_EX_TOOLWINDOW, vData, False
End Property

Public Property Get TitleBar() As Boolean
    TitleBar = m_TitleBar
End Property

Public Property Let TitleBar(ByVal vData As Boolean)
    m_TitleBar = vData
    SetWindowStyle False, WS_CAPTION, vData, False
End Property

Public Property Get TaskBar() As Boolean
    TaskBar = m_TaskBar
End Property

Public Property Let TaskBar(ByVal vData As Boolean)
    m_TaskBar = vData
    SetWindowStyle True, WS_EX_APPWINDOW, vData, False
End Property

Public Property Get SizeBorder() As Boolean
    SizeBorder = m_SizeBrd
End Property

Public Property Let SizeBorder(ByVal vData As Boolean)
    m_SizeBrd = vData
    SetWindowStyle False, WS_THICKFRAME, vData, False
End Property

Public Property Get ControlBox() As Boolean
    ControlBox = m_CtrlBox
End Property

Public Property Let ControlBox(ByVal vData As Boolean)
    m_CtrlBox = vData
    SetWindowStyle False, WS_SYSMENU, vData, False
End Property

Public Property Get MinButton() As Boolean
    MinButton = m_MinButt
End Property

Public Property Let MinButton(ByVal vData As Boolean)
    m_MinButt = vData
    SetWindowStyle False, WS_MINIMIZEBOX, vData, False
End Property

Public Property Get MaxButton() As Boolean
    MaxButton = m_MaxButt
End Property

Public Property Let MaxButton(ByVal vData As Boolean)
    m_MaxButt = vData
    SetWindowStyle False, WS_MAXIMIZEBOX, vData, False
End Property


'Default Member
Public Sub SetWindowStyle(ByVal extended_Style As Boolean, ByVal style_Value As Long, ByVal new_Value As Boolean, ByVal bRefresh As Boolean)
Attribute SetWindowStyle.VB_UserMemId = 0
   Dim style_Type As Long, Style As Long
   
   If extended_Style Then style_Type = GWL_EXSTYLE Else style_Type = GWL_STYLE
   
   Style = GetWindowLongW(hWnd, style_Type)
   
   If new_Value Then Style = Style Or style_Value Else Style = Style And Not style_Value
   
   If bRefresh Then ShowWindow hWnd, SW_HIDE
   SetWindowLongW hWnd, style_Type, Style
   
   If bRefresh Then ShowWindow hWnd, SW_SHOW
   SetWindowPos hWnd, 0, 0, 0, 0, 0, SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER
End Sub


'---------------------------------------------------------------------------------
Public Property Get Child(Optional ByVal newLong As Long = WS_CHILD) As Long
    Child = GetWindowLongW(hWnd, GWL_STYLE) And newLong
End Property

Public Property Let Child(Optional ByVal newLong As Long = WS_CHILD, ByVal vData As Long)
    If vData <> 0 Then
        SetParent hWnd, vData
        SetWindowLongW hWnd, GWL_STYLE, GetWindowLongW(hWnd, GWL_STYLE) Or newLong
    Else
        SetParent hWnd, 0
        SetWindowLongW hWnd, GWL_STYLE, GetWindowLongW(hWnd, GWL_STYLE) And (Not newLong)
    End If
End Property

Public Property Get TopMost() As Boolean
    TopMost = (GetForegroundWindow = hWnd)
End Property

Public Property Let TopMost(ByVal vData As Boolean)
    SetTopMost hWnd, vData
    If vData Then SetForegroundWindow hWnd
End Property

Public Property Get Text() As String
    Dim sz As Long, txt As String
    sz = GetWindowTextLengthW(hWnd)
    If sz Then txt = String$(sz, 0):    GetWindowTextW hWnd, StrPtr(txt), sz + 1:      Text = txt
End Property

Public Property Let Text(ByVal vData As String)
    SetWindowTextW hWnd, StrPtr(vData)
End Property

Public Sub Icon(Optional ByVal indexIcon As Variant = 1, Optional ByVal nameIcon As String)
    SetIconWindow hWnd, indexIcon, nameIcon
End Sub


'---------------------------------------------------------------------------------
Public Function TB_State(ByVal tbpFlags As Long) As Long
    DoEvents
    If ITB Is Nothing Then Exit Function
    TB_State = ITB.SetProgressState(hWnd, tbpFlags)
End Function

Public Function TB_Value(ByVal ullCompleted As Long, ByVal ullTotal As Long) As Long
    DoEvents
    If ITB Is Nothing Then Exit Function
    TB_Value = ITB.SetProgressValue(hWnd, ullCompleted, 0, ullTotal, 0)
End Function

Public Function TB_Icon(ByVal hIcon As Long, Optional ByVal pszDescription As String) As Long
    DoEvents
    If ITB Is Nothing Then Exit Function
    TB_Icon = ITB.SetOverlayIcon(hWnd, hIcon, StrPtr(pszDescription))
End Function

Public Function TB_Clip(ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long) As Long
    Dim rc As RECT
    DoEvents
    If ITB Is Nothing Then Exit Function
    rc.Left = Left: rc.Top = Top: rc.Right = Right: rc.Bottom = Bottom
    TB_Clip = ITB.SetThumbnailClip(hWnd, rc)
End Function

Public Function TB_Tooltip(Optional ByVal pszTip As String) As Long
    DoEvents
    If ITB Is Nothing Then Exit Function
    TB_Tooltip = ITB.SetThumbnailTooltip(hWnd, StrPtr(pszTip))
End Function

Public Function TB_Buttons(ByVal buttons As Variant) As Long
    Dim i As Integer
    DoEvents
    If ITB Is Nothing Or Not IsArray(buttons) Then Exit Function
    ReDim bts(UBound(buttons)) As THUMBBUTTON
    For i = 0 To UBound(buttons)
        bts(i).dwMask = THB_ICON Or THB_TOOLTIP Or THB_FLAGS
        bts(i).IID = i
        bts(i).hIcon = buttons(i)(0)
        bts(i).dwFlags = buttons(i)(1)
        bts(i).szTip = buttons(i)(2) & vbNullChar
    Next
    TB_Buttons = ITB.ThumbBarAddButtons(hWnd, UBound(bts) + 1, VarPtr(bts(0)))
End Function


'---------------------------------------------------------------------------------
Private Sub Class_Initialize()
    On Error Resume Next
    m_MaxButt = False:  m_MinButt = True:   m_CtrlBox = True:   m_SizeBrd = False
    m_TaskBar = True:   m_TitleBar = True:  m_ToolWin = False:  m_CloseButt = True
    Set ITB = New TaskbarList
End Sub

Private Sub Class_Terminate()
    Set ITB = Nothing
End Sub
