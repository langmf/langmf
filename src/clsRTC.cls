VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRTC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'============================================================================
'
' Update for use in class for VB6.   Copyright (c) 2020 by atomix
' Runtime Tiny C Compiler for VB6.   Copyright (c) 2018 by wqweto@gmail.com
' Obfuscated Tiny C Compiler.        Copyright (c) 2001-2003 Fabrice Bellard
'
'============================================================================

Option Explicit

Private Type RTC_BUFFER
    Data() As Byte
End Type

Private Type RTC_CONTEXT
    Prog        As Long
    Syms        As Long
    Mods        As Long
    Glob        As Long
    Vars        As Long
    State(31)   As Long
    Buffer(3)   As RTC_BUFFER
End Type

Private ctx As RTC_CONTEXT

Public ptrFunc As Long, ptrThunk As Long


Public Function Run(ParamArray Args() As Variant) As Variant
Attribute Run.VB_UserMemId = 0
    Dim i As Long, s As Integer, c As Long, p(31) As Long, t(31) As Integer, a(31) As Variant
    
    c = UBound(Args) + 1
    
    For i = 0 To c - 1
        s = VarType(Args(i))
        Select Case s
        Case vbString:      a(i) = StrPtr(Args(i))
        Case vbObject:      a(i) = ObjPtr(Args(i))
        Case Else:          If s And vbArray Then a(i) = VarPtr(Args(i)) Else a(i) = CLng(Args(i))
        End Select
        p(i) = VarPtr(a(i)):        t(i) = 3
    Next

    Call DispCallFunc(0, ptrFunc, 4, vbLong, c, t(0), p(0), Run)
End Function

Public Function Compile(ByVal value As String, Optional ByVal aSize As Long = 32768) As Object
    Set Compile = Me

    If ctx.Prog = 0 Then
        ctx.Prog = VirtualAlloc(0, aSize, MEM_COMMIT, PAGE_EXECUTE_READWRITE)
        
        ReDim ctx.Buffer(0).Data(aSize - 1):            ctx.Syms = VarPtr(ctx.Buffer(0).Data(0))
        ReDim ctx.Buffer(1).Data(aSize - 1):            ctx.Glob = VarPtr(ctx.Buffer(1).Data(0))
        ReDim ctx.Buffer(2).Data(aSize * 8 - 1):        ctx.Vars = VarPtr(ctx.Buffer(2).Data(0))
        ReDim ctx.Buffer(3).Data(4000 - 1):             ctx.Mods = VarPtr(ctx.Buffer(3).Data(0))
        
        Call EnumProcessModules(GetCurrentProcess(), ByVal ctx.Mods, 1000, ByVal 0&)
    End If
    
    ptrFunc = CallWindowProcA(ptrThunk, VarPtr(ctx), StrPtr(value))
End Function

Public Function Free() As Object
    Dim tmp As RTC_CONTEXT
    Set Free = Me
    If ctx.Prog <> 0 Then Call VirtualFree(ctx.Prog, 0, MEM_RELEASE):       ctx = tmp:      ptrFunc = 0
End Function

Public Function Patch(ByVal pfn As Long) As Object
    Dim bInIDE As Boolean
    Set Patch = Me
    Debug.Assert pvSetTrue(bInIDE)
    If bInIDE Then Call CopyMemory(pfn, ByVal pvAdd(pfn, &H16), 4) Else Call VirtualProtect(pfn, 8, PAGE_EXECUTE_READWRITE, 0)
    Call CopyMemory(ByVal pfn, -802975883527609.7192@, 8)
End Function

Public Function Ptr(Optional ByVal value As Variant) As Variant
    Dim p As Long
    
    If IsMissing(value) Then
    
        Ptr = VarPtr(ctx)
        
    ElseIf VarType(value) = vbString Then
    
        p = StrStrA(ctx.Syms, " " & value & " ")
        If p Then p = pvAdd(ctx.Vars, pvDiff(ctx.Syms, p) * 8 + &H100):    Call CopyMemory(p, ByVal p, 4)
        Ptr = p
        
    ElseIf IsNumeric(value) Then
    
        Set Ptr = Me
        ptrFunc = CLng(value)

    End If
End Function

Public Property Get State(ByVal i As Long) As Long
    State = ctx.State(i)
End Property

Public Property Let State(ByVal i As Long, ByVal value As Long)
    ctx.State(i) = value
End Property


Private Function pvSetTrue(value As Boolean) As Boolean
    value = True:      pvSetTrue = True
End Function

Private Function pvAdd(ByVal Ptr1 As Long, ByVal ofs As Long) As Long
    ' safely add *signed* offset to *unsigned* ptr for *unsigned* retval w/o overflow
    pvAdd = ((Ptr1 Xor &H80000000) + ofs) Xor &H80000000
End Function

Private Function pvDiff(ByVal Ptr1 As Long, ByVal Ptr2 As Long) As Long
    ' retval is *signed* offset b/n *unsigned* ptr1 and *unsigned* ptr2 w/o overflow
    If (Ptr1 Xor Ptr2) < 0 Then pvDiff = (Ptr2 - (Ptr1 Xor &H80000000)) Xor &H80000000:   Exit Function
    pvDiff = Ptr2 - Ptr1
End Function


Private Sub Class_Initialize()
    Dim b() As Byte
    b = LoadResData("rtc.bin", 50)
    ptrThunk = VirtualAlloc(0, UBound(b) + 1, MEM_COMMIT, PAGE_EXECUTE_READWRITE)
    Call CopyMemory(ByVal ptrThunk, b(0), UBound(b) + 1)
End Sub

Private Sub Class_Terminate()
    Call VirtualFree(ptrThunk, 0, MEM_RELEASE)
End Sub
