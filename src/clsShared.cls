VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsShared"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private HashDLL As clsHash


Public Function timeGetTime(Optional ByVal tMode As Long = -11) As Currency
    Static vOld As Currency, vNew As Currency, vAll As Currency
    If tMode = -13 Then timeGetTime = modConst.timeBeginPeriod():                           Exit Function
    If tMode = -12 Then timeGetTime = modConst.timeEndPeriod():                             Exit Function
    If tMode = -11 Then timeGetTime = modConst.timeGetTime():                               Exit Function
    If tMode = -3 Then timeGetTime = ((vNew - vOld) / QPF) * 1000:                          Exit Function
    If tMode = -2 Then timeGetTime = vAll:      vAll = 0:                                   Exit Function
    If tMode Then QueryPerformanceCounter vOld:     timeGetTime = (vOld / QPF) * 1000:      Exit Function
    QueryPerformanceCounter vNew:       timeGetTime = ((vNew - vOld) / QPF) * 1000:     vAll = vAll + timeGetTime
End Function

Public Function WindowLong(ByVal hWnd As Long, Optional ByVal mStyle As Long, Optional ByVal value As Variant, Optional ByVal nIndex As Long = GWL_STYLE) As Long
    WindowLong = modCommon.WindowLong(hWnd, value, mStyle, nIndex)
End Function

Public Sub GMove(Optional ByVal Obj As Object, Optional ByVal typeX As Single = -1, Optional ByVal typeY As Single = -1, Optional ByVal typeW As Single, Optional ByVal typeH As Single, Optional ByVal offsetX As Single, Optional ByVal offsetY As Single, Optional ByVal prtWidth As Single, Optional ByVal prtHeight As Single, Optional x As Variant = "Left", Optional y As Variant = "Top", Optional Width As Variant = "Width", Optional Height As Variant = "Height")
    FlexMove Obj, typeX, typeY, typeW, typeH, offsetX, offsetY, prtWidth, prtHeight, x, y, Width, Height
End Sub

Public Sub GPrint(ByVal Obj As Object, ByVal txt As String, Optional ByVal cx As Single = -0.2, Optional ByVal cy As Single = -0.2)
    If cx <> -0.2 And cx <> -1 Then If cx = -0.1 Then Obj.CurrentX = Obj.ScaleWidth / 2 - Obj.TextWidth(txt) / 2 Else Obj.CurrentX = cx
    If cy <> -0.2 And cy <> -1 Then If cy = -0.1 Then Obj.CurrentY = Obj.ScaleHeight / 2 - Obj.TextHeight(txt) / 2 Else Obj.CurrentY = cy
    If cy = -1 Or cx = -1 Then Obj.Print txt; Else Obj.Print txt
End Sub

Public Sub GLine(ByVal Obj As Object, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, Optional ByVal cColor As Long = 0, Optional ByVal cType As Integer = 0)
    Select Case cType
       Case 0:    Obj.Line (x1, y1)-(x2, y2), cColor
       Case 1:    Obj.Line (x1, y1)-(x2, y2), cColor, B
       Case 2:    Obj.Line (x1, y1)-(x2, y2), cColor, BF
    End Select
End Sub

Public Sub GCircle(ByVal Obj As Object, ByVal x1 As Single, ByVal y1 As Single, ByVal radius As Single, Optional ByVal Color As Long = 0, Optional ByVal cStart As Single = -10, Optional ByVal cEnd As Single = -10, Optional ByVal cAspect As Single = -10)
    If cStart = -10 And cEnd = -10 And cAspect = -10 Then Obj.Circle (x1, y1), radius, Color Else Obj.Circle (x1, y1), radius, Color, cStart, cEnd, cAspect
End Sub

Public Function QBColor(ByVal value As Integer) As Long
    If value >= 0 Then QBColor = VBA.QBColor(value) Else QBColor = GetSysColor(Abs(value) And 255)
End Function

Public Property Get vbPX() As Single
    vbPX = Screen.TwipsPerPixelX
End Property

Public Property Get vbPY() As Single
    vbPY = Screen.TwipsPerPixelY
End Property

Public Property Get vbSM(ByVal nIndex As Long) As Long
    vbSM = GetSystemMetrics(nIndex)
End Property

Public Property Get vbDC(ByVal nIndex As Long, Optional ByVal hDC As Long) As Long
    If hDC = 0 Then hDC = frmScript.hDC
    vbDC = GetDeviceCaps(hDC, nIndex)
End Property

Public Property Get vbMissing() As Variant
    vbMissing = DISP_E_PARAMNOTFOUND:     VariantType(vbMissing) = VT_ERROR
End Property

Public Function IsMissing(Var As Variant) As Boolean
    IsMissing = Information.IsMissing(Var)
End Function

Public Function IsError(Var As Variant) As Boolean
    IsError = Information.IsError(Var)
End Function

Public Function StrConv(ByVal txt As String, ByVal typeStr As Integer, Optional ByVal LocaleID As Long = -1) As String
    If LocaleID = -1 Then StrConv = VBA.StrConv(txt, typeStr) Else StrConv = VBA.StrConv(txt, typeStr, LocaleID)
End Function

Public Function Format(ByVal txt As String, ByVal fmt As String) As String
    Format = VBA.Format$(txt, fmt)
End Function

Public Function Val(Var As Variant) As Double
    Val = VBA.Val(Var)
End Function

Public Function StrPtr(Var As Variant, Optional ByVal bEmptyToAlloc As Boolean = False) As Long
    If bEmptyToAlloc Then If Len(Var) = 0 Then Var = String$(MAX_PATH_UNI, 0)
    StrPtr = VBA.StrPtr(Var)
End Function

Public Function ObjPtr(Var As Variant) As Long
    ObjPtr = VBA.ObjPtr(Var)
End Function

Public Function VarPtr(Var As Variant, Optional ByVal bEmptyToLong As Boolean = False) As Long
    If bEmptyToLong Then If IsEmpty(Var) Then Var = CLng(0)
    VarPtr = VBA.VarPtr(Var)
End Function

Public Property Get VType(vrtSrc As Variant, Optional ByVal isBYREF As Boolean) As Integer
    VType = VariantType(vrtSrc, isBYREF)
End Property

Public Property Let VType(vrtSrc As Variant, Optional ByVal isBYREF As Boolean, ByVal value As Integer)
    VariantType(vrtSrc, isBYREF) = value
End Property

Public Function VarCopy(srcVar As Variant, Optional dstVar As Variant, Optional ByVal sz As Long, Optional ByVal srcOffset As Long = 8, Optional ByVal dstOffset As Long = 8) As Variant
    If sz = 0 Then
        dstVar = Empty:     VariantCopy dstVar, srcVar
    ElseIf sz = -1 Then
        dstVar = Empty:     VariantCopyInd2 dstVar, srcVar
    ElseIf sz > 0 Then
        If IsMissing(dstVar) Then dstVar = CLng(0)
        CopyMemory ByVal VBA.VarPtr(dstVar) + dstOffset, ByVal VBA.VarPtr(srcVar) + srcOffset, sz
    End If
End Function

Public Function ObjFromPtr(ByVal vPtr As Long, Optional ByVal isIUnknown As Boolean = False) As Variant
    Dim VD As Object, IUnk As ATL.IUnknown
    If isIUnknown Then
        CopyMem4 vPtr, IUnk:  Set ObjFromPtr = IUnk:  CopyMem4 0&, IUnk
    Else
        CopyMem4 vPtr, VD:    Set ObjFromPtr = VD:    CopyMem4 0&, VD
    End If
End Function

Public Property Get ByRefVar(Var As Variant, ParamArray d() As Variant) As Variant
    Select Case UBound(d)
        Case -1:    ByRefVar = Var
        Case 0:     ByRefVar = Var(d(0))
        Case 1:     ByRefVar = Var(d(0), d(1))
        Case 2:     ByRefVar = Var(d(0), d(1), d(2))
    End Select
End Property

Public Property Let ByRefVar(Var As Variant, ParamArray d() As Variant, value As Variant)
    Select Case UBound(d)
        Case -1:    Var = value
        Case 0:     Var(d(0)) = value
        Case 1:     Var(d(0), d(1)) = value
        Case 2:     Var(d(0), d(1), d(2)) = value
    End Select
End Property

Public Property Set ByRefVar(Var As Variant, ParamArray d() As Variant, value As Variant)
    Select Case UBound(d)
        Case -1:    Set Var = value
        Case 0:     Set Var(d(0)) = value
        Case 1:     Set Var(d(0), d(1)) = value
        Case 2:     Set Var(d(0), d(1), d(2)) = value
    End Select
End Property

Public Property Get Is64() As Boolean
    Is64 = modCommon.Is64
End Property

Public Property Get IsAdmin() As Boolean
    IsAdmin = modCommon.IsAdmin
End Property

Public Property Get IsWine() As Boolean
    IsWine = modCommon.IsWine
End Property

Public Property Get IsEnd() As Boolean
    IsEnd = mf_IsEnd
End Property

Public Property Get IsExe() As Boolean
    IsExe = Info.IsExe
End Property

Public Sub WaitMs(Optional ByVal msec As Long)
    modCommon.WaitMs msec
End Sub

Public Sub Sleep(Optional ByVal msec As Long)
    modConst.Sleep msec
End Sub

Public Function DoEvents2() As Long
    DoEvents2 = DoEvents
End Function

Public Function DoEvents3() As Long
    DoEvents3 = API_DoEvents
End Function

Public Sub EndMF(Optional ByVal Interval As Long = 1)
    modCommon.EndMF Interval
End Sub

Public Function ExistsMember(Obj As Variant, ByVal ProcName As String) As Boolean
    ExistsMember = modCommon.ExistsMember(Obj, ProcName)
End Function

Public Function CallByName(Obj As Variant, ProcName As Variant, ByVal typeCall As Long, ParamArray Args() As Variant)
    CBN Obj, ProcName, typeCall, Args, , VBA.VarPtr(CallByName)
End Function

Public Function CallByName2(Obj As Variant, ProcName As Variant, ByVal typeCall As Long, Args() As Variant, Optional ByVal CountArgs As Long = -1)
    CBN Obj, ProcName, typeCall, Args, CountArgs, VBA.VarPtr(CallByName2)
End Function

Public Function GetFunc(Optional ByVal value As Variant) As Variant
    VariantCopy GetFunc, modCommon.GetFunc(value)
End Function

Public Property Get Struct(ByVal This As Long, Optional vsp As Variant) As Long
    Dim Ptr As Long
    Ptr = StructFunc(This, vsp)
    If vsp(3) <> 0 Then CopyMemory ByVal CLng(vsp(3)), ByVal Ptr, vsp(2) Else CopyMemory Struct, ByVal Ptr, vsp(2)
End Property

Public Property Let Struct(ByVal This As Long, Optional vsp As Variant, ByVal value As Long)
    Dim Ptr As Long
    Ptr = StructFunc(This, vsp)
    If vsp(3) <> 0 Then CopyMemory ByVal Ptr, ByVal CLng(vsp(3)), vsp(2) Else CopyMemory ByVal Ptr, value, vsp(2)
End Property

Public Property Get Object(Optional ByVal vHash As Object) As Object
    Dim Obj As IObject
    Set Obj = New IObject:      If vHash Is Nothing Then Set Obj.Hash = New clsHash Else Set Obj.Hash = vHash
    Set Object = Create_Interface(Obj)
End Property

Public Function ArraySize(Var As Variant) As Long
    ArraySize = modCommon.ArraySize(Var)
End Function

Public Function ArrayPtr(Var As Variant) As Long
    ArrayPtr = GetSafeArray(Var).pvData
End Function

Public Function ArrayByte(Var As Variant) As Byte()
    If VarType(Var) = vbArray + vbByte Then ArrayByte = Var Else ConvToBufferByte Var, ArrayByte
End Function

Public Function ArrayParam(ByVal dummy As Long, ParamArray Args() As Variant) As Variant()
    Dim p As Long
    GetMem4 VBA.VarPtr(dummy) + 4, p:   GetMem4 p, dummy:   PutMem4 VarPtrArray(ArrayParam), dummy:   PutMem4 p, 0&
End Function

Public Property Get ResToPic(ByVal Key As String) As IPictureDisp
    If Not SYS.Resource.Exists(Key) Then Exit Property
    Set ResToPic = LoadPictureFromByte(SYS.Resource(Key).Data)
End Property

Public Property Get ResToStr(ByVal Key As String) As String
    Dim vt As Integer, cd As clsDim, tmpBuf() As Byte
    If Not SYS.Resource.Exists(Key) Then Exit Property
    Set cd = SYS.Resource(Key):     vt = VarType(cd.Data)
    If vt = vbString Then ResToStr = cd.Data
    If vt = vbArray + vbByte Then tmpBuf = cd.Data: If modCommon.ArraySize(tmpBuf) Then ResToStr = Conv_A2W_Buf(tmpBuf)
End Property

Public Property Get ResToBuf(ByVal Key As String) As Byte()
    Dim vt As Integer, cd As clsDim
    If Not SYS.Resource.Exists(Key) Then Exit Property
    Set cd = SYS.Resource(Key):     vt = VarType(cd.Data)
    If vt = vbString Then ResToBuf = Conv_W2A_Buf(cd.Data)
    If vt = vbArray + vbByte Then ResToBuf = cd.Data
End Property

Public Function IIF(varIF As Variant, var1 As Variant, var2 As Variant) As Variant
    If varIF Then
        If IsObject(var1) Then Set IIF = var1 Else IIF = var1
    Else
        If IsObject(var2) Then Set IIF = var2 Else IIF = var2
    End If
End Function

Public Function VOF(ParamArray Vars() As Variant) As Variant
    Dim i As Long, t As Integer
    For i = 0 To UBound(Vars)
        VariantCopy VOF, Vars(i)
        t = VariantType(VOF)
        If t = vbString Then
            If Len(VOF) <> 0 Then Exit For
        ElseIf t = vbObject Then
            If Not VOF Is Nothing Then Exit For
        ElseIf t And vbArray Then
            Exit For
        ElseIf t = vbError Or t = vbNull Then
        ElseIf CBool(VOF) Then
            Exit For
        End If
    Next
End Function

Public Function VAF(ParamArray Vars() As Variant) As Variant
    Dim i As Long, t As Integer
    For i = 0 To UBound(Vars)
        VariantCopy VAF, Vars(i)
        t = VariantType(VAF)
        If t = vbString Then
            If Len(VAF) = 0 Then Exit For
        ElseIf t = vbObject Then
            If VAF Is Nothing Then Exit For
        ElseIf t And vbArray Then
        ElseIf t = vbError Or t = vbNull Then
            Exit For
        ElseIf Not CBool(VAF) Then
            Exit For
        End If
    Next
End Function

Public Function LenType(Obj As Variant) As Long
    On Error Resume Next
    LenType = Obj.•SIZE
    Err.Clear
End Function

Public Function InitType(Obj As Variant, Optional ByVal vOffset As Long, Optional ByVal wrap As Object) As Object
    Dim tmp As Variant
    
    On Error GoTo err1
    
    Obj.•OFS = vOffset
    
    With Obj
        If wrap Is Nothing Then Set .•WRAP = New clsBuffer: .•WRAP.Size = .•SIZE Else Set .•WRAP = wrap
        For Each tmp In .•ENUM:      InitType tmp(0), tmp(1), .•WRAP:      Next
    End With
    
err1:
    Err.Clear
    Set InitType = Obj
End Function

Public Function DllCall(ByVal lData As String, ByVal sData As Variant, ParamArray Params() As Variant) As Long
    Dim i As Long, a As Long, vt As Integer, l As LONG64, uds As Long, v() As Variant, t(40) As Long
    Dim isUnicode As Boolean, ID As String, clsFCP As clsFuncPointer

    ID = lData & " " & sData:      Set clsFCP = HashDLL(ID)

    If clsFCP Is Nothing Then Set clsFCP = New clsFuncPointer:  Call clsFCP.Config(lData, sData):  HashDLL.Add clsFCP, ID
    
    isUnicode = clsFCP.Unicode
    
    uds = UBound(Params)
    If uds > -1 Then ReDim v(uds)

    'Ïðåä-îáðàáîòêà ïàðàìåòðîâ
    For i = 0 To uds
        t(i) = VariantType(Params(i))
        
        Select Case t(i)
            Case vbArray + vbVariant
                vt = VarType(Params(i)(0))
                Select Case vt
                    Case vbSingle, vbDouble:    CopyMemory l, ByVal VBA.VarPtr(Params(i)(0)) + 8, 8:    v(a) = l.LowPart:    If vt = vbDouble Then ReDim Preserve v(UBound(v) + 1):    a = a + 1:   v(a) = l.HighPart
                    Case vbArray + vbByte:      t(i) = GetSafeArray(Params(i)(0)).pvData:     v(a) = VBA.VarPtr(t(i))
                    Case vbString:              t(i) = VBA.StrPtr(Params(i)(0)):              v(a) = VBA.VarPtr(t(i))
                    Case Else:                  v(a) = VBA.VarPtr(Params(i)(0)) + 8
                End Select
                            
            Case vbString
                If Not isUnicode Then Params(i) = Conv_W2A_Str(CStr(Params(i)))
                v(a) = VBA.StrPtr(Params(i))
                
            Case Else
                v(a) = CLng(Params(i))
        End Select
        
        a = a + 1
    Next

    DllCall = clsFCP.PCall(v)

    'Ïîñëå-îáðàáîòêà ïàðàìåòðîâ
    If Not isUnicode Then
        For i = 0 To uds
            If t(i) = vbString Then Params(i) = Conv_A2W_Str(CStr(Params(i)))
        Next
    End If
End Function

Private Sub Class_Initialize()
    Set HashDLL = New clsHash
    Set HashDLL.Default = Nothing
End Sub
