VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFuncPointer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const MAXCODE As Long = &HEC00&, MAXTPL As Long = 256
Private btASM(MAXTPL - 1) As Byte, btFnc(MAXCODE - 1) As Byte, pc As Long

Public hLib As Long, ptrIfc As Long, ptrASM As Long, ptrCode As Long, CallType As Long, Unicode As Boolean


Public Function Config(Optional ByVal lData As String, Optional ByVal sData As Variant, Optional vUnicode As Variant) As Boolean
    Dim vt As Integer

    'хэндл библиотеки
    If Left$(lData, 1) = "*" Then CallType = API_CDecl:    lData = Mid$(lData, 2) Else CallType = API_StdCall
    If Left$(lData, 1) = "@" Then Unicode = True:          lData = Mid$(lData, 2) Else Unicode = False
    If Left$(lData, 1) = "!" Then ptrCode = wglGetProcAddress(sData):    lData = Mid$(lData, 2)
    If LenB(lData) Then hLib = LoadLibrary(StrPtr(lData))

    'хэндл функции
    vt = VarType(sData)
    If vt = vbString Then
        If hLib Then
            If Left$(sData, 1) = "#" Then
                ptrCode = GetProcAddressByOrdinal(hLib, Val(Mid$(sData, 2)))
            Else
                ptrCode = GetProcAddress(hLib, sData)
            End If
        End If
    Else
        ptrCode = CLng(sData)
    End If
    
    vUnicode = Unicode
    
    Config = (ptrCode <> 0)
End Function

'Default Member
Public Function PCall(ParamArray Params() As Variant) As Long
Attribute PCall.VB_UserMemId = 0
    Dim a As Long, b As Long, uds As Long, fnc As Long, cnt As Long, cntEnd As Long, t(40) As Integer

    pc = 6
    uds = UBound(Params)
    If ptrCode <> 0 Then cntEnd = 0 Else cntEnd = 1

    For a = uds To cntEnd Step -1
        t(a) = VariantType(Params(a), True)
        If t(a) = vbArray + vbVariant Then
            For b = m_ArraySize(Params(a)) - 1 To 0 Step -1
                AddPush CLng(Params(a)(b))        ' PUSH dword
                cnt = cnt + 1
            Next
        ElseIf t(a) = vbString Then
            If Not Unicode Then Params(a) = StrConv(Params(a), vbFromUnicode)
            AddPush StrPtr(Params(a))             ' PUSH dword
            cnt = cnt + 1
        Else
            AddPush CLng(Params(a))               ' PUSH dword
            cnt = cnt + 1
        End If
    Next

    If ptrIfc <> 0 Then
        If uds > -1 Then fnc = VTable(Params(0))
        AddPush ptrIfc
    Else
        If ptrCode = 0 Then
            If uds > -1 Then fnc = CLng(Params(0))
        Else
            fnc = ptrCode
        End If
    End If

    If fnc <> 0 Then
        AddCall fnc                   ' CALL rel addr
        
        If CallType > 0 Then For a = 1 To cnt:    AddByte &H59:    Next
        
        AddByte &HC3                  ' RET
        
        PCall = CallWindowProcA(ptrASM, 0, 0, 0, 0)
    End If
    
    If Not Unicode Then
        For a = uds To cntEnd Step -1
            If t(a) = vbString Then Params(a) = StrConv(Params(a), vbUnicode)
        Next
    End If
End Function

Public Function VTable(Optional ByVal Entry As Long = 0, Optional ByVal Obj As Long = 0) As Long
    Dim ptrVTable As Long
    
    If Obj <> 0 Then ptrVTable = Obj Else ptrVTable = ptrIfc
    
    GetMem4 ptrVTable, ptrVTable
    GetMem4 ptrVTable + Entry * 4, VTable
End Function

Public Function GetBuffer() As Byte()
    GetBuffer = btFnc
End Function

Public Sub SetBuffer(Buf As Variant)
    Dim tmpBuf() As Byte
    
    If ConvToBufferByte(Buf, tmpBuf) Then
        ReDim Preserve tmpBuf(UBound(btFnc))
        CopyMemory btFnc(0), tmpBuf(0), UBound(btFnc) + 1
        ptrCode = VarPtr(btFnc(0))
        AllowExecuteCode ptrCode, UBound(btFnc) + 1
    End If
End Sub

Private Sub AddPush(ByVal lng As Long)
    AddByte &H68
    AddLong lng
End Sub

Private Sub AddCall(ByVal addr As Long)
    AddByte &HE8
    AddLong addr - ptrASM - pc - 4
End Sub

Private Sub AddLong(ByVal lng As Long)
    CopyMem4 lng, btASM(pc)
    pc = pc + 4
End Sub

Private Sub AddByte(ByVal bt As Byte)
    btASM(pc) = bt
    pc = pc + 1
End Sub

Private Sub Class_Initialize()
    btASM(0) = &H58                 ' POP EAX
    btASM(1) = &H59                 ' POP ECX
    btASM(2) = &H59                 ' POP ECX
    btASM(3) = &H59                 ' POP ECX
    btASM(4) = &H59                 ' POP ECX
    btASM(5) = &H50                 ' PUSH EAX

    ptrASM = VarPtr(btASM(0))
    AllowExecuteCode ptrASM, MAXTPL

    Unicode = True
End Sub

Private Sub Class_Terminate()
    If hLib Then FreeLibrary hLib
End Sub
