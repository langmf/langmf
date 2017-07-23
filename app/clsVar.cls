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

Public Function VarByVal(ByVal ptrDst As Long, ByVal ptrSrc As Long) As Long
    VarByVal = VariantCopyInd(ptrDst, ptrSrc)
End Function

Public Sub VarZero(vrtSrc As Variant)
    ZeroMemory vrtSrc, 16
End Sub

Public Function CopyVar(srcVar As Variant, Optional dstVar As Variant, Optional ByVal Size As Long = 4, Optional ByVal srcOffset As Long = 8, Optional ByVal dstOffset As Long = 8) As Variant
    If IsMissing(dstVar) Then dstVar = CLng(0)
    CopyMemory ByVal VBA.VarPtr(dstVar) + dstOffset, ByVal VBA.VarPtr(srcVar) + srcOffset, Size
    CopyVar = dstVar
End Function

Public Function Val(Var As Variant) As Double
    Val = VBA.Val(Var)
End Function

Public Function CVar2(Var As Variant) As Variant
    CVar2 = CVar(Var)
End Function

