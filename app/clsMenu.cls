VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function RegisterHotKey Lib "user32" (ByVal hWnd As Long, ByVal ID As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
Private Declare Function UnregisterHotKey Lib "user32" (ByVal hWnd As Long, ByVal ID As Long) As Long

Private m_Popup As Boolean

Public hWnd As Long, IDClick As Long, Parent As Object, SrvMenu As Object, SubMenu As Object


Public Property Get Popup() As Boolean
    Popup = m_Popup
End Property

Public Property Let Popup(ByVal value As Boolean)
    m_Popup = value
    If hWnd Then DestroyMenu hWnd
    If value Then hWnd = CreatePopupMenu Else hWnd = CreateMenu
End Property

Public Function Find(ByVal nameItem As Variant, Optional ByVal prtHWnd As Long) As Long
    Dim a As Long, ID As Long, isSubMenu As Boolean
    
    If VarType(nameItem) <> vbString Then Find = nameItem: Exit Function
    
    Find = -1
    If prtHWnd = 0 Then prtHWnd = hWnd
    
    For a = 0 To GetMenuItemCount(prtHWnd) - 1
        isSubMenu = False
        
        ID = GetMenuItemID(prtHWnd, a)
        If ID = -1 Then
            ID = GetSubMenu(prtHWnd, a)
            isSubMenu = True
        End If
        
        If Text(ID) = nameItem Then
            Find = ID
        Else
            If isSubMenu Then Find = Find(nameItem, ID)
        End If
        
        If Find > 0 Then Exit For
    Next
End Function

Public Sub Style(Optional ByVal fMask As Long = MIM_BACKGROUND Or MIM_MAXHEIGHT Or MIM_APPLYTOSUBMENUS, Optional ByVal bkgColor As Long = &H80000004, Optional ByVal maxHeight As Long = 0, Optional ByVal dwStyle As Long = 0)
    Dim mi As MENUINFO
    With mi
        .cbSize = Len(mi)
        .fMask = fMask
        .hbrBack = CreateSolidBrush(bkgColor)
        .cyMax = maxHeight
        .dwStyle = dwStyle
    End With
    SetMenuInfo hWnd, mi
    DrawMenuBar hWnd
End Sub

Public Property Get Text(ByVal ID_Menu As Variant) As String
    Dim txt As String, Size As Long
    ID_Menu = Find(ID_Menu)
    txt = String$(MAX_PATH_X2, 0)
    Size = GetMenuStringW(hWnd, ID_Menu, StrPtr(txt), Len(txt), 0)
    If Size > 0 Then Text = Left$(txt, Size)
End Property

Public Property Let Text(ByVal ID_Menu As Variant, ByVal value As String)
    Dim Flag As Long
    ID_Menu = Find(ID_Menu)
    Flag = GetMenuState(hWnd, ID_Menu, MF_BYCOMMAND)
    If IsMask(Flag, MF_OWNERDRAW) Then Flag = Flag And &HFFFFFEFF
    If IsMask(Flag, MF_BYPOSITION) Then Flag = Flag And &HFFFFFBFF
    ModifyMenuW hWnd, ID_Menu, Flag, ID_Menu, StrPtr(value)
End Property

Public Property Get State(ByVal ID_Menu As Variant, Optional ByVal Mask As Long = 65535) As Long
    ID_Menu = Find(ID_Menu)
    State = GetMenuState(hWnd, ID_Menu, MF_BYCOMMAND) And Mask
End Property

Public Property Let State(ByVal ID_Menu As Variant, Optional ByVal Mask As Long = 65535, ByVal value As Long)
    Dim tmp As Long
    ID_Menu = Find(ID_Menu)
    tmp = GetMenuState(hWnd, ID_Menu, MF_BYCOMMAND)
    If value <> 0 Then tmp = tmp Or Mask Else tmp = tmp And Not Mask
    ModifyMenuW hWnd, ID_Menu, tmp, ID_Menu, StrPtr(Text(ID_Menu))
End Property

Public Property Get NameEvent(ByVal ID_Menu As Variant) As String
    ID_Menu = Find(ID_Menu)
    NameEvent = SrvMenu("ev_" & ID_Menu)
End Property

Public Property Let NameEvent(ByVal ID_Menu As Variant, ByVal value As String)
    ID_Menu = Find(ID_Menu)
    If LenB(value) Then SrvMenu("ev_" & ID_Menu) = value Else Call SrvMenu.Remove("ev_" & ID_Menu)
End Property

Public Sub Bitmap(ByVal ID_Menu As Variant, Optional ByVal img1 As Variant, Optional ByVal img2 As Variant)
    Dim pic1 As Long, pic2 As Long
    
    On Error Resume Next
    
    ID_Menu = Find(ID_Menu)
    
    If VarType(img1) = vbString Then
        SrvMenu("p1_" & ID_Menu) = LoadPicture(img1)
        pic1 = SrvMenu("p1_" & ID_Menu).Handle
    Else
        If IsMissing(img1) Then Call SrvMenu.Remove("p1_" & ID_Menu) Else SrvMenu("p1_" & ID_Menu) = img1:   pic1 = img1
    End If
    
    If VarType(img2) = vbString Then
        SrvMenu("p2_" & ID_Menu) = LoadPicture(img2)
        pic2 = SrvMenu("p2_" & ID_Menu).Handle
    Else
        If IsMissing(img2) Then Call SrvMenu.Remove("p2_" & ID_Menu) Else SrvMenu("p2_" & ID_Menu) = img2:   pic2 = img2
    End If
    
    SetMenuItemBitmaps hWnd, ID_Menu, MF_BITMAP, pic1, pic2
End Sub

Public Sub Remove(ByVal ID_Menu As Variant)
    Dim ID As Long
    ID = Find(ID_Menu)
    If ID > -1 Then NameEvent(ID) = "":    Bitmap ID:    HotKey ID:    DeleteMenu hWnd, ID, MF_BYCOMMAND
End Sub

'Default Member
Public Function Add(ByVal ID_Menu As Long, ByVal Names As Variant, Optional ByVal nameSubMenu As String, Optional ByVal bDisable As Boolean = False, Optional ByVal bDefault As Boolean = False, Optional ByVal bChecked As Boolean = False, Optional ByVal img1 As Variant, Optional ByVal img2 As Variant, Optional ByVal bPopup As Boolean, Optional ByVal arrSubMenu As Variant) As Object
Attribute Add.VB_UserMemId = 0
    Dim a As Long, itemName As String, itemEvent As String, itemHotKey As Long
    Static offset As Long
    
    On Error Resume Next
    
    If VarType(Names) = vbString Then
        itemName = Names
    Else
        itemName = Names(0)
        If UBound(Names) > 0 Then itemEvent = Names(1)
        If UBound(Names) > 1 Then itemHotKey = Names(2)
    End If
    
    If LenB(nameSubMenu) = 0 Then

        If LenB(itemEvent) Then NameEvent(ID_Menu) = itemEvent
        If itemHotKey Then HotKey ID_Menu, itemHotKey
        
        If itemName = "-" Then
            AppendMenuW hWnd, MF_STRING Or MF_SEPARATOR, ID_Menu, 0
        Else
            AppendMenuW hWnd, MF_STRING Or -bChecked * MF_CHECKED, ID_Menu, StrPtr(itemName)
            If Not IsMissing(img1) Or Not IsMissing(img2) Then Bitmap ID_Menu, img1, img2
            If bDefault Then SetMenuDefaultItem hWnd, ID_Menu, 0
            If bDisable Then EnableMenuItem hWnd, ID_Menu, MF_BYCOMMAND Or MF_GRAYED
        End If
        
        Set Add = Me
    Else
    
        Dim newSubMenu As New clsMenu
        SubMenu.Add newSubMenu, nameSubMenu
        
        With newSubMenu
            .Initialize Parent, SrvMenu
            
            If bPopup Then .Popup = True
            If LenB(itemName) Then AppendMenuW hWnd, MF_STRING Or MF_POPUP, .hWnd, StrPtr(itemName)
    
            If IsArray(arrSubMenu) Then
                If arrSubMenu(0) <> vbEmpty Then offset = arrSubMenu(0)
                For a = 1 To UBound(arrSubMenu)
                    .Add offset, arrSubMenu(a)
                    offset = offset + 1
                Next
            End If
        End With

        Set Add = newSubMenu
        Set newSubMenu = Nothing
        
    End If
End Function

Public Sub Show(Optional ByVal x As Long = -1, Optional ByVal y As Long = -1)
    Dim pt As POINTAPI, ID_Menu As Long
    
    On Error Resume Next
    
    If Popup Then
        GetCursorPos pt
        
        If x = -1 Then x = pt.x
        If y = -1 Then y = pt.y
    
        SetForegroundWindow Parent.hWnd
        ID_Menu = TrackPopupMenuEx(hWnd, TPM_RETURNCMD Or TPM_RIGHTBUTTON, x, y, Parent.hWnd, ByVal 0&)
        PostMessageA Parent.hWnd, 0, 0, 0

        If ID_Menu <> 0 Then Click ID_Menu
    Else
        Refresh
    End If
End Sub

Public Sub Refresh(Optional ByVal bShow As Boolean = True)
    If Parent Is Nothing Then Exit Sub
    If Not Popup Then SetMenu Parent.hWnd, IIF(bShow, hWnd, 0)
End Sub

Public Sub Click(ByVal ID_Menu As Long)
    IDClick = ID_Menu
    If LenB(NameEvent(ID_Menu)) = 0 Then
        If Parent.Events("Menu" & ID_Menu & "_Click") = False Then Parent.Events "Menu_Click", ID_Menu
    Else
        Parent.Events NameEvent(ID_Menu) & "_Click"
    End If
    Parent.SetFocus
End Sub

Public Function HotKey(ByVal ID As Variant, Optional ByVal vKey As Long) As Long
    Dim v As Variant
    
    If Parent Is Nothing Or SrvMenu Is Nothing Then Exit Function
    
    With Parent
        If VarType(ID) = vbBoolean Then
            For Each v In SrvMenu.Items(, "^hk_\d+")
                If ID Then
                    RegisterHotKey .hWnd, v(0), v(1) \ &H10000, v(1) And &HFFFF&
                Else
                    UnregisterHotKey .hWnd, v(0)
                End If
            Next
            
        ElseIf VarType(ID) = vbEmpty Then
            For Each v In SrvMenu.Items(, "^hk_\d+")
                SrvMenu.Remove "hk_" & v(0)
                UnregisterHotKey .hWnd, v(0)
            Next
            
        Else
            If vKey = 0 Then
                SrvMenu.Remove "hk_" & ID
                HotKey = UnregisterHotKey(.hWnd, ID)
            Else
                If SrvMenu.Exists("hk_" & ID) Then Call UnregisterHotKey(.hWnd, ID)
                SrvMenu.Item("hk_" & ID) = Array(ID, vKey)
                HotKey = RegisterHotKey(.hWnd, ID, vKey \ &H10000, vKey And &HFFFF&)       'ALT=1 CTRL=2 SHIFT=4 WIN=8
            End If
        End If
    End With
End Function

Public Sub Initialize(ByVal objPrt As Object, ByVal objSrv As Object)
    hWnd = CreateMenu
    Set SubMenu = New clsHash
    Set Parent = objPrt
    Set SrvMenu = objSrv
    SrvMenu("ref") = SrvMenu("ref") + 1
End Sub

Private Sub Class_Terminate()
    If hWnd Then DestroyMenu hWnd
    hWnd = 0
    
    SrvMenu("ref") = SrvMenu("ref") - 1
    
    If SrvMenu("ref") = 0 Then HotKey Empty

    Set Parent = Nothing
    Set SrvMenu = Nothing
    Set SubMenu = Nothing
End Sub
