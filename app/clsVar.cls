VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public Property Get MemCurrency(ByVal addr As Long) As Currency
    GetMem8 addr, MemCurrency
End Property

Public Property Let MemCurrency(ByVal addr As Long, ByVal value As Currency)
    PutMem8 addr, value
End Property

Public Property Get MemLong(ByVal addr As Long) As Long
    GetMem4 addr, MemLong
End Property

Public Property Let MemLong(ByVal addr As Long, ByVal value As Long)
    PutMem4 addr, value
End Property

Public Property Get MemWord(ByVal addr As Long) As Long
    GetMem2_Wrd addr, MemWord
End Property

Public Property Let MemWord(ByVal addr As Long, ByVal value As Long)
    PutMem2_Wrd addr, value
End Property

Public Property Get MemByte(ByVal addr As Long) As Byte
    GetMem1 addr, MemByte
End Property

Public Property Let MemByte(ByVal addr As Long, ByVal value As Byte)
    PutMem1 addr, value
End Property

Public Sub MemFill(ByVal addr As Long, Optional ByVal Size As Long = 16, Optional ByVal value As Byte)
    FillMemory ByVal addr, Size, value
End Sub

Public Function VarByVal(ByVal addr As Long) As Variant
    VariantCopyInd VarPtr(VarByVal), addr
End Function

Public Sub VarZero(vrtSrc As Variant)
    ZeroMemory vrtSrc, 16
End Sub

Public Function CVar2(Var As Variant) As Variant
    CVar2 = CVar(Var)
End Function

Public Function VErr(ByVal errCode As Long) As Variant
    VErr = CVErr(errCode)
End Function

Public Function VLng(Var As Variant) As Long
    VLng = CLng(Var)
End Function

Public Function VStr(Var As Variant) As String
    VStr = CStr(Var)
End Function

Public Function Conv(Optional Var As Variant, Optional ByVal nType As Integer = 10, Optional ByVal mskAnd As Long) As Variant
    VariantCopy Conv, Var:      If nType > -1 Then VariantType(Conv) = nType
    If IsNumeric(Conv) = True And mskAnd <> 0 Then Conv = Conv And mskAnd
End Function

Public Function Map(arr As Variant, Optional cbk As Variant) As Variant()
    Dim a As Long, i As Long, j As Long, c As Long, uds As Long, t As Long, v As Variant, h As Variant, tmp() As Variant
    
    On Error Resume Next
    
    t = VariantType(cbk):       If t = VT_ARRVAR Then c = UBound(cbk)

    If TypeName(arr) = "clsHash" Then
        For Each h In arr.Items(2)
            VariantCopy v, h(0)
            If i >= uds Then uds = uds + 100:        ReDim Preserve tmp(uds - 1)
            If t = VT_ARRVAR Then For a = 0 To c:    VariantCopy v, cbk(a)(v, h(1), arr):     Next
            If t = VT_OBJECT Then VariantCopy v, cbk(v)
            If Not IsMissing(v) Then VariantCopy tmp(i), v:     i = i + 1
        Next
    Else
        For Each v In arr
            If i >= uds Then uds = uds + 100:        ReDim Preserve tmp(uds - 1)
            If t = VT_ARRVAR Then For a = 0 To c:    VariantCopy v, cbk(a)(v, j, arr):        Next
            If t = VT_OBJECT Then VariantCopy v, cbk(v)
            If Not IsMissing(v) Then VariantCopy tmp(i), v:     i = i + 1
            j = j + 1
        Next
    End If
    
    Err.Clear:          If i Then ReDim Preserve tmp(i - 1):  Map = tmp Else Map = Array()
End Function

