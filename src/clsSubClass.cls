VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSubClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' comctl32 versions 5.7- exported the APIs by Ordinal
Private Declare Function SetWindowSubclassOrdinal Lib "comctl32" Alias "#410" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long, ByVal dwRefData As Long) As Long
Private Declare Function DefSubclassProcOrdinal Lib "comctl32" Alias "#413" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function RemoveWindowSubclassOrdinal Lib "comctl32" Alias "#412" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long) As Long

' comctl32 versions 5.8+ exported the APIs by Name
Private Declare Function DefSubclassProc Lib "comctl32" (ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowSubclass Lib "comctl32" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long, ByVal dwRefData As Long) As Long
Private Declare Function RemoveWindowSubclass Lib "comctl32" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long) As Long

' comctl32 versions 6.0+ exported the APIs by Name
Private Declare Function GetWindowSubclass Lib "comctl32" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long, pdwRefData As Long) As Long

Private hLib As Long, hMem As Long, isFail As Boolean, bALL As Boolean, arrMsg() As Byte, udsMsg As Long, isOrdinal As Boolean

Const WM_NCDESTROY As Long = &H82, p1 As Long = 15, p3 As Long = 55, p2 As Long = 63, p4 As Long = 82, p5 As Long = 119
Const CODE_STR As String = "C808000031C08945FC8945F8EB0EE8xxxxx01x83F802740F85C07417E839000000837DF8007505E8190000008B45FCC9C21800FF751868xxxxx03xFF7508E8xxxxx02xEBE7FF7514FF7510FF750CFF7508E8xxxxx04x8945FCC38D45FC50FF751CFF75188D4514508D451050FF750CFF75088D45F850B8xxxxx05x508B00FF501CC3"


'THIS PROCEDURE SHOULD BE FIRST IN CLASS - don't move it anywhere
Public Function Proc(ByRef bHandled As Boolean, ByVal hWnd As Long, ByVal uMsg As Long, ByRef wParam As Long, ByRef lParam As Long, ByVal Obj As Object, ByVal dwRefData As Long) As Long            'no_check
    Dim vMsg As Byte
    
    If bALL Then
        Proc = Obj.WindowProc(bHandled, hWnd, uMsg, wParam, lParam, dwRefData)
    Else
        If uMsg >= 0 And uMsg <= udsMsg Then
            vMsg = arrMsg(uMsg)
            If (vMsg And 4) Then Proc = HookDefault(hWnd, uMsg, wParam, lParam)
            If (vMsg And 2) Then bHandled = True
            If (vMsg And 1) Then Proc = Obj.WindowProc(bHandled, hWnd, uMsg, wParam, lParam, dwRefData)
        End If
    End If
    If uMsg = WM_NCDESTROY Then Call HookClear(Obj, hWnd)
End Function

Public Property Get Msg(Optional ByVal Index As Long = -1) As Byte
    If Index < 0 Then Msg = bALL: Exit Property
    If Index >= &H10000 Then Exit Property
    If Index > udsMsg Then ReDim Preserve arrMsg(Index):  udsMsg = Index
    Msg = arrMsg(Index)
End Property

Public Property Let Msg(Optional ByVal Index As Long = -1, ByVal value As Byte)
    If Index < 0 Then bALL = value: Exit Property
    If Index >= &H10000 Then Exit Property
    If Index > udsMsg Then ReDim Preserve arrMsg(Index):  udsMsg = Index
    arrMsg(Index) = value
End Property

Public Property Let List(ParamArray arr() As Variant, ByVal value As Byte)            'no_check
    Dim a As Long, b As Long
    If UBound(arr) = -1 Then
        For a = 0 To UBound(arrMsg):    arrMsg(a) = value:    Next
    Else
        For a = 0 To UBound(arr)
            If IsArray(arr(a)) Then
                For b = 0 To UBound(arr(a)):      Msg(arr(a)(b)) = value:      Next
            Else
                Msg(arr(a)) = value
            End If
        Next
    End If
End Property

Public Function HookSet(ByVal Obj As Object, Optional ByVal hWnd As Long = 0, Optional ByVal dwRefData As Long) As Boolean
    If isFail Then Exit Function
    If hWnd = 0 Then hWnd = Obj.hWnd
    If isOrdinal Then
        HookSet = CBool(SetWindowSubclassOrdinal(hWnd, hMem, ObjPtr(Obj), dwRefData))
    Else
        HookSet = CBool(SetWindowSubclass(hWnd, hMem, ObjPtr(Obj), dwRefData))
    End If
End Function

Public Function HookGetData(ByVal Obj As Object, Optional ByVal hWnd As Long, Optional dwRefData As Variant) As Long
    Dim value As Long
    If isFail Then Exit Function
    If hWnd = 0 Then hWnd = Obj.hWnd
    HookGetData = GetWindowSubclass(hWnd, hMem, ObjPtr(Obj), value)
    dwRefData = value
End Function

Public Function HookClear(ByVal Obj As Object, Optional ByVal hWnd As Long = 0) As Boolean
    If isFail Then Exit Function
    If hWnd = 0 Then hWnd = Obj.hWnd
    If isOrdinal Then
        HookClear = CBool(RemoveWindowSubclassOrdinal(hWnd, hMem, ObjPtr(Obj)))
    Else
        HookClear = CBool(RemoveWindowSubclass(hWnd, hMem, ObjPtr(Obj)))
    End If
End Function

Public Function HookDefault(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If isFail Then Exit Function
    If isOrdinal Then
        HookDefault = DefSubclassProcOrdinal(hWnd, uMsg, wParam, lParam)
    Else
        HookDefault = DefSubclassProc(hWnd, uMsg, wParam, lParam)
    End If
End Function

Private Sub Class_Initialize()
    Dim i As Long, nLen As Long, sCode As String, ptrRemove As Long, ptrDefProc As Long
    
    On Error Resume Next
    
    Msg(&H400&) = 0
    
    nLen = Len(CODE_STR) / 2
    For i = 0 To nLen - 1
      sCode = sCode & ChrB$(Val("&H" & Mid$(CODE_STR, i * 2 + 1, 2)))
    Next
    
    hMem = CoTaskMemAlloc(nLen)
    Call VirtualProtect(hMem, nLen, PAGE_EXECUTE_READWRITE, 0&)
    Call CopyMemory(ByVal hMem, ByVal StrPtr(sCode), nLen)
    Call PatchVal(p3, hMem)   'addr. of subclass proc
    
    hLib = LoadLibrary(StrPtr("comctl32"))
    If hLib Then
        If GetProcAddress(hLib, "SetWindowSubclass") = 0& Then
            If GetProcAddressByOrdinal(hLib, 410&) = 0& Then isFail = True
            isOrdinal = True
            ptrRemove = GetProcAddressByOrdinal(hLib, 412)
            ptrDefProc = GetProcAddressByOrdinal(hLib, 413)
        Else
            ptrRemove = GetProcAddress(hLib, "RemoveWindowSubclass")
            ptrDefProc = GetProcAddress(hLib, "DefSubclassProc")
        End If
        
        Call PatchRel(p2, ptrRemove)
        Call PatchRel(p4, ptrDefProc)
    Else
        isFail = True
    End If
    
    If IsIDE Then
        Call CopyMemory(ByVal hMem + 12, &H9090&, 2)
        Call PatchRel(p1, GetProcAddress(GetModuleHandleW(StrPtr("vba6")), "EbMode"))
    End If
    
    Call PatchVal(p5, ObjPtr(Me))
End Sub

Private Sub Class_Terminate()
    If hMem Then CoTaskMemFree hMem
    If hLib Then Call FreeLibrary(hLib)
End Sub

Private Sub PatchRel(ByVal nOffset As Long, ByVal nTargetAddr As Long)
    Call CopyMemory(ByVal (hMem + nOffset), nTargetAddr - hMem - nOffset - 4, 4)
End Sub

Private Sub PatchVal(ByVal nOffset As Long, ByVal nValue As Long)
    Call CopyMemory(ByVal (hMem + nOffset), nValue, 4)
End Sub
