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
            For b = ArraySize(Params(a)) - 1 To 0 Step -1
                AddPush CLng(Params(a)(b))        ' PUSH dword
                cnt = cnt + 1
            Next
        ElseIf t(a) = vbString Then
            If Not Unicode Then Params(a) = Conv_W2A_Str(CStr(Params(a)))
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
        fnc = ptrCode:    If fnc = 0 And uds > -1 Then fnc = CLng(Params(0))
    End If

    If fnc <> 0 Then
        AddCall fnc                   ' CALL rel addr
        
        If CallType > 0 Then For a = 1 To cnt:    AddByte &H59:    Next
        
        AddByte &HC3                  ' RET
        
        PCall = CallWindowProcA(ptrASM, 0, 0, 0, 0)
    End If
    
    If Not Unicode Then
        For a = uds To cntEnd Step -1
            If t(a) = vbString Then Params(a) = Conv_A2W_Str(CStr(Params(a)))
        Next
    End If
End Function

Public Property Get VTable(Optional ByVal Entry As Long, Optional ByVal Obj As Variant) As Long
    Dim Ptr As Long
    If IsObject(Obj) Then Ptr = ObjPtr(Obj) Else If IsMissing(Obj) Then Ptr = ptrIfc Else Ptr = Obj
    GetMem4 Ptr, Ptr:    Ptr = Ptr + Entry * 4:    GetMem4 Ptr, VTable
End Property

Public Property Let VTable(Optional ByVal Entry As Long, Optional ByVal Obj As Variant, ByVal value As Long)
    Dim Ptr As Long
    If IsObject(Obj) Then Ptr = ObjPtr(Obj) Else If IsMissing(Obj) Then Ptr = ptrIfc Else Ptr = Obj
    GetMem4 Ptr, Ptr:    Ptr = Ptr + Entry * 4:    AllowExecuteCode Ptr, 4:    PutMem4 Ptr, value
End Property

Public Property Get Buffer() As Variant
    Buffer = btFnc
End Property

Public Property Let Buffer(Buf As Variant)
    Dim tmpBuf() As Byte
    ConvToBufferByte Buf, tmpBuf
    ReDim Preserve tmpBuf(UBound(btFnc))
    CopyMemory btFnc(0), tmpBuf(0), UBound(btFnc) + 1
    ptrCode = VarPtr(btFnc(0))
    AllowExecuteCode ptrCode, UBound(btFnc) + 1
End Property

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
