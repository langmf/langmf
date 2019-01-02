VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFileAPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type MCurr
    MVal As Currency
End Type

Const FILE_BEGIN As Long = 0, FILE_CURRENT As Long = 1, FILE_END As Long = 2

Private hFile As Long, hMap As Long, hMem As Long, MC As MCurr, ML As LONG64

Public FileName As String, AutoFlush As Boolean, EOF As Boolean


Public Property Get Handle() As Long
    Handle = hFile
End Property

Public Sub FClose()
    UnMap
    If hFile <> INVALID_HANDLE Then CloseHandle hFile
    FileName = "":    AutoFlush = False:    EOF = False:    hFile = INVALID_HANDLE
End Sub

Public Function FOpen(ByVal fName As String, Optional ByVal dwCreation As Long = OPEN_ALWAYS, Optional ByVal dwAccess As Long = GENERIC_READ Or GENERIC_WRITE, Optional ByVal dwShare As Long = FILE_SHARE_READ, Optional ByVal dwFlags As Long = FILE_ATTRIBUTE_NORMAL) As Long
    Call FClose
    hFile = CreateFile(StrPtr(LongPath(fName)), dwAccess, dwShare, 0, dwCreation, dwFlags, 0)
    If hFile = INVALID_HANDLE Or hFile = 0 Then Call FClose
    FileName = fName
    FOpen = hFile
End Function

Public Function Map(Optional ByVal dwProtect As Long = PAGE_READONLY, Optional ByVal dwAccess As Long = FILE_MAP_READ, Optional ByVal numPos As Currency = 0, Optional ByVal cntBytes As Long) As Long
    If hFile = INVALID_HANDLE Or hMap <> 0 Or hMem <> 0 Then Exit Function
    
    hMap = CreateFileMappingW(hFile, 0&, dwProtect, 0&, 0&, 0&):    If hMap = 0 Then Exit Function
    
    MC.MVal = numPos / 10000@:      LSet ML = MC:       ML.LowPart = ML.LowPart And &HFFFF0000
    hMem = MapViewOfFile(hMap, dwAccess, ML.HighPart, ML.LowPart, cntBytes)
    Map = hMem
End Function

Public Sub UnMap()
    If hMem <> 0 Then Call UnmapViewOfFile(hMem):    hMem = 0
    If hMap <> 0 Then Call CloseHandle(hMap):        hMap = 0
End Sub

Public Function FGet(Buf As Variant, Optional ByVal numPos As Currency = 0) As Long
    Dim tmpByte As Byte, tmpInteger As Integer, tmpLong As Long, tmpSingle As Single, tmpString As String
    Dim tmpBoolean As Boolean, tmpCurrency As Currency, tmpDouble As Double
    Dim SA As SafeArray, sz As Long, tmpBuf() As Byte
    
    On Error Resume Next
    
    If hFile = INVALID_HANDLE Then Exit Function
    
    Select Case VariantType(Buf, True)
        Case vbArray + vbByte:      SA = GetSafeArray(Buf):       FGet = GetMem(SA.pvData, SA.rgSABound(0).cElements, numPos)
        Case vbArray + vbVariant:   SA = GetSafeArray(Buf):       sz = SA.rgSABound(0).cElements:       ReDim tmpBuf(sz - 1):      FGet = GetMem(VarPtr(tmpBuf(0)), sz, numPos):        ConvFromBufferByte Buf, tmpBuf
        Case vbObject:      FGet = GetMem(Buf.Class___Wrapper.Ptr + Buf.Class___Offset, Buf.Class___Size, numPos)
        Case vbByte:        FGet = GetMem(VarPtr(tmpByte), 1, numPos):          Buf = tmpByte
        Case vbLong:        FGet = GetMem(VarPtr(tmpLong), 4, numPos):          Buf = tmpLong
        Case vbInteger:     FGet = GetMem(VarPtr(tmpInteger), 2, numPos):       Buf = tmpInteger
        Case vbString:      tmpString = Buf:        FGet = GetStr(tmpString, numPos):         Buf = tmpString
        Case vbDouble:      FGet = GetMem(VarPtr(tmpDouble), 8, numPos):        Buf = tmpDouble
        Case vbSingle:      FGet = GetMem(VarPtr(tmpSingle), 4, numPos):        Buf = tmpSingle
        Case vbBoolean:     FGet = GetMem(VarPtr(tmpBoolean), 2, numPos):       Buf = tmpBoolean
        Case vbCurrency:    FGet = GetMem(VarPtr(tmpCurrency), 8, numPos):      Buf = tmpCurrency
    End Select
End Function

Public Function FPut(Buf As Variant, Optional ByVal numPos As Currency = 0) As Long
    Dim tmpByte As Byte, tmpInteger As Integer, tmpLong As Long, tmpSingle As Single, tmpString As String
    Dim tmpBoolean As Boolean, tmpCurrency As Currency, tmpDouble As Double
    Dim SA As SafeArray, tmpBuf() As Byte
    
    On Error Resume Next
    
    If hFile = INVALID_HANDLE Then Exit Function
    
    Select Case VariantType(Buf, True)
        Case vbArray + vbByte:      SA = GetSafeArray(Buf):       FPut = PutMem(SA.pvData, SA.rgSABound(0).cElements, numPos)
        Case vbArray + vbVariant:   ConvToBufferByte Buf, tmpBuf:   FPut = PutMem(VarPtr(tmpBuf(0)), UBound(tmpBuf) + 1, numPos)
        Case vbObject:      FPut = PutMem(Buf.Class___Wrapper.Ptr + Buf.Class___Offset, Buf.Class___Size, numPos)
        Case vbByte:        tmpByte = Buf:          FPut = PutMem(VarPtr(tmpByte), 1, numPos)
        Case vbLong:        tmpLong = Buf:          FPut = PutMem(VarPtr(tmpLong), 4, numPos)
        Case vbInteger:     tmpInteger = Buf:       FPut = PutMem(VarPtr(tmpInteger), 2, numPos)
        Case vbString:      tmpString = Buf:        FPut = PutStr(tmpString, numPos)
        Case vbDouble:      tmpDouble = Buf:        FPut = PutMem(VarPtr(tmpDouble), 8, numPos)
        Case vbSingle:      tmpSingle = Buf:        FPut = PutMem(VarPtr(tmpSingle), 4, numPos)
        Case vbBoolean:     tmpBoolean = Buf:       FPut = PutMem(VarPtr(tmpBoolean), 2, numPos)
        Case vbCurrency:    tmpCurrency = Buf:      FPut = PutMem(VarPtr(tmpCurrency), 8, numPos)
    End Select
End Function

Public Function FInput(Optional value As Variant, Optional ByVal fndStr As String = vbCrLf) As Long
    Dim Buf As String, txt As String, rc As Long
    If hFile = INVALID_HANDLE Then Exit Function
    Buf = String$(MAX_PATH_UNI, 0)
    Do
        If ReadFile(hFile, ByVal Buf, Len(Buf), FInput, 0) = 0 Then Exit Do
        If (FInput = 0) Then Exit Do
        rc = InStr(Buf, fndStr)
        If rc Then
            txt = txt + Left$(Buf, rc - 1):   Pos = Pos - FInput + rc + Len(fndStr) - 1:   Exit Do
        Else
            txt = txt + TrimNull(Buf)
        End If
    Loop
    EOF = (FInput = 0)
    value = txt
End Function

Public Function FPrint(Optional ByVal value As String, Optional ByVal addStr As String = vbCrLf) As Long
    If hFile = INVALID_HANDLE Then Exit Function
    value = value + addStr
    If WriteFile(hFile, ByVal value, Len(value), FPrint, 0) = 0 Then Exit Function
    If AutoFlush Then Flush
End Function

Public Function GetBuf(value() As Byte, Optional ByVal numPos As Currency = 0) As Long
    If hFile = INVALID_HANDLE Then Exit Function
    If numPos > 0 Then Pos = numPos
    If ReadFile(hFile, value(0), UBound(value) + 1, GetBuf, 0) = 0 Then Exit Function
    EOF = (GetBuf = 0)
End Function

Public Function PutBuf(value() As Byte, Optional ByVal numPos As Currency = 0) As Long
    If hFile = INVALID_HANDLE Then Exit Function
    If numPos > 0 Then Pos = numPos
    If WriteFile(hFile, value(0), UBound(value) + 1, PutBuf, 0) = 0 Then Exit Function
    If AutoFlush Then Flush
End Function

Public Function GetStr(value As String, Optional ByVal numPos As Currency = 0) As Long
    If hFile = INVALID_HANDLE Then Exit Function
    If numPos > 0 Then Pos = numPos
    If ReadFile(hFile, ByVal value, Len(value), GetStr, 0) = 0 Then Exit Function
    EOF = (GetStr = 0)
End Function

Public Function PutStr(value As String, Optional ByVal numPos As Currency = 0) As Long
    If hFile = INVALID_HANDLE Then Exit Function
    If numPos > 0 Then Pos = numPos
    If WriteFile(hFile, ByVal value, Len(value), PutStr, 0) = 0 Then Exit Function
    If AutoFlush Then Flush
End Function

Public Function GetMem(ByVal value As Long, ByVal sz As Long, Optional ByVal numPos As Currency = 0) As Long
    If hFile = INVALID_HANDLE Then Exit Function
    If numPos > 0 Then Pos = numPos
    If ReadFile(hFile, ByVal value, sz, GetMem, 0) = 0 Then Exit Function
    EOF = (GetMem = 0)
End Function

Public Function PutMem(ByVal value As Long, ByVal sz As Long, Optional ByVal numPos As Currency = 0) As Long
    If hFile = INVALID_HANDLE Then Exit Function
    If numPos > 0 Then Pos = numPos
    If WriteFile(hFile, ByVal value, sz, PutMem, 0) = 0 Then Exit Function
    If AutoFlush Then Flush
End Function

Public Property Get Pos(Optional ByVal dwMethod As Long = FILE_BEGIN) As Currency
    If hFile = INVALID_HANDLE Then Exit Property
    ML.LowPart = SetFilePointer(hFile, 0, ML.HighPart, FILE_CURRENT)
    If ML.LowPart = INVALID_HANDLE Then If Err.LastDllError Then Exit Property
    LSet MC = ML
    Pos = MC.MVal * 10000@
    Pos = Pos + 1
End Property

Public Property Let Pos(Optional ByVal dwMethod As Long = FILE_BEGIN, ByVal value As Currency)
    If hFile = INVALID_HANDLE Then Exit Property
    
    If dwMethod = FILE_END Then
        ML.HighPart = 0
        ML.LowPart = 0
    ElseIf dwMethod = FILE_CURRENT Then
        ML.HighPart = 0
        ML.LowPart = value
    Else
        value = value - 1
        If value < 0 Then Exit Property
        MC.MVal = value / 10000@
        LSet ML = MC
    End If
    
    Call SetFilePointer(hFile, ML.LowPart, ML.HighPart, dwMethod)
End Property

Public Property Get LOF() As Currency
    If hFile = INVALID_HANDLE Then Exit Property
    ML.LowPart = GetFileSize(hFile, ML.HighPart)
    If ML.LowPart = INVALID_HANDLE Then If Err.LastDllError Then Exit Property
    LSet MC = ML
    LOF = MC.MVal * 10000@
End Property

Public Property Let LOF(ByVal value As Currency)
    If hFile = INVALID_HANDLE Then Exit Property
    If SetEndOfFile(hFile) <> 0 And value > 0 Then Pos = value:     WriteFile hFile, 0&, 1, 0&, 0
End Property

Public Sub Flush()
    If hFile = INVALID_HANDLE Then Exit Sub
    FlushFileBuffers hFile
End Sub


Private Sub Class_Initialize()
    hFile = INVALID_HANDLE
End Sub

Private Sub Class_Terminate()
    Call FClose
End Sub
