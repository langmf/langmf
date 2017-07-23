VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private iPtr As Long, Buffer() As Byte


Public Property Get Ptr(Optional ByVal allowExec As Boolean) As Long
    If m_ArraySize(Buffer) = 0 Then Exit Property
    iPtr = VarPtr(Buffer(0))
    If allowExec Then AllowExecuteCode iPtr, Size
    Ptr = iPtr
End Property

Public Property Let Ptr(Optional ByVal allowExec As Boolean, ByVal value As Long)
    iPtr = value
End Property

Public Property Get Size() As Long
    Size = m_ArraySize(Buffer)
End Property

Public Property Let Size(ByVal value As Long)
    If value = 0 Then
        Erase Buffer
    ElseIf value < 0 Then
        ReDim Preserve Buffer(Abs(value) - 1)
    Else
        ReDim Buffer(value - 1)
    End If
    iPtr = Ptr
End Property

Public Sub SetData(Buf As Variant)
    ConvToBufferByte Buf, Buffer
    iPtr = Ptr
End Sub

Public Function GetData(Optional Buf As Variant) As Byte()
    If IsMissing(Buf) Then GetData = Buffer: Exit Function
    ConvFromBufferByte Buf, Buffer
End Function


'Default Member
Public Property Get PByte(ByVal offset As Long) As Byte
Attribute PByte.VB_UserMemId = 0
    GetMem1 iPtr + offset, PByte
End Property

Public Property Let PByte(ByVal offset As Long, ByVal value As Byte)
    PutMem1 iPtr + offset, value
End Property

Public Property Get PWord(ByVal offset As Long) As Long
    GetMem2_Wrd iPtr + offset, PWord
End Property

Public Property Let PWord(ByVal offset As Long, ByVal value As Long)
    PutMem2_Wrd iPtr + offset, value
End Property

Public Property Get PInteger(ByVal offset As Long) As Integer
    GetMem2 iPtr + offset, PInteger
End Property

Public Property Let PInteger(ByVal offset As Long, ByVal value As Integer)
    PutMem2 iPtr + offset, value
End Property

Public Property Get PBoolean(ByVal offset As Long) As Boolean
    GetMem2_Bln iPtr + offset, PBoolean
End Property

Public Property Let PBoolean(ByVal offset As Long, ByVal value As Boolean)
    PutMem2_Bln iPtr + offset, value
End Property

Public Property Get PLong(ByVal offset As Long) As Long
    GetMem4 iPtr + offset, PLong
End Property

Public Property Let PLong(ByVal offset As Long, ByVal value As Long)
    PutMem4 iPtr + offset, value
End Property

Public Property Get PSingle(ByVal offset As Long) As Single
    GetMem4_Sng iPtr + offset, PSingle
End Property

Public Property Let PSingle(ByVal offset As Long, ByVal value As Single)
    PutMem4_Sng iPtr + offset, value
End Property

Public Property Get PDouble(ByVal offset As Long) As Double
    GetMem8_Dbl iPtr + offset, PDouble
End Property

Public Property Let PDouble(ByVal offset As Long, ByVal value As Double)
    PutMem8_Dbl iPtr + offset, value
End Property

Public Property Get PCurrency(ByVal offset As Long) As Currency
    GetMem8 iPtr + offset, PCurrency
End Property

Public Property Let PCurrency(ByVal offset As Long, ByVal value As Currency)
    PutMem8 iPtr + offset, value
End Property

Public Property Get PString(ByVal offset As Long, ByVal lenStr As Long) As String
    PString = String$(Abs(lenStr), 0)
    CopyMemory ByVal PString, ByVal iPtr + offset, Abs(lenStr)
    If lenStr < 0 Then PString = TrimNull(PString)
End Property

Public Property Let PString(ByVal offset As Long, ByVal lenStr As Long, ByVal value As String)
    lenStr = Abs(lenStr)
    If Len(value) < lenStr Then value = value & String$(lenStr - Len(value), 0)
    CopyMemory ByVal iPtr + offset, ByVal value, lenStr
End Property

Public Property Get PArray(ByVal offset As Long, ByVal lenArr As Long) As Byte()
    ReDim value(lenArr - 1) As Byte
    CopyMemory value(0), Buffer(offset), lenArr
End Property

Public Property Let PArray(ByVal offset As Long, ByVal lenArr As Long, value() As Byte)
    CopyMemory Buffer(offset), value(0), lenArr
End Property
