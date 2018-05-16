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

Private m_Ptr As Long, m_Offset As Long, m_Size As Long, m_Overflow As Boolean, Buffer() As Byte


Public Sub SetData(Buf As Variant)
    ConvToBufferByte Buf, Buffer:    m_Ptr = Ptr
End Sub

Public Function GetData(Optional ByVal typeVar As Integer = vbArray + vbByte) As Variant
    ConvFromBufferByte GetData, Buffer, typeVar
End Function

Public Property Get Ptr(Optional ByVal allowExec As Boolean) As Long
    m_Size = m_ArraySize(Buffer):    If m_Size = 0 Then Exit Property
    m_Ptr = VarPtr(Buffer(0)):       If allowExec Then AllowExecuteCode m_Ptr, Size
    Ptr = m_Ptr
End Property

Public Property Let Ptr(Optional ByVal allowExec As Boolean, ByVal value As Long)
    m_Ptr = value
End Property

Public Property Get Size() As Long
    Size = m_Size
End Property

Public Property Let Size(ByVal value As Long)
    If value = 0 Then Erase Buffer
    If value < 0 Then ReDim Preserve Buffer(Abs(value) - 1)
    If value > 0 Then ReDim Buffer(value - 1)
    m_Ptr = Ptr
End Property

Public Property Get Offset() As Long
    Offset = m_Offset
End Property

Public Property Let Offset(ByVal value As Long)
    If value < 0 Then m_Offset = m_Offset + value:  If m_Offset < 0 Then m_Offset = 0
    If value >= 0 Then m_Offset = value
End Property

Public Property Get Overflow() As Boolean
    Overflow = m_Overflow
End Property

Public Property Let Overflow(ByVal value As Boolean)
    m_Overflow = value
End Property


'Default Member
Public Property Get PByte(Optional ByVal vOffset As Long = -1) As Byte
Attribute PByte.VB_UserMemId = 0
    If vOffset <> -1 Then m_Offset = vOffset
    If Overflow = False And m_Offset > (m_Size - 1) Then Exit Property
    GetMem1 m_Ptr + m_Offset, PByte
    m_Offset = m_Offset + 1
End Property

Public Property Let PByte(Optional ByVal vOffset As Long = -1, ByVal value As Byte)
    If vOffset <> -1 Then m_Offset = vOffset
    If Overflow = False And m_Offset > (m_Size - 1) Then Exit Property
    PutMem1 m_Ptr + m_Offset, value
    m_Offset = m_Offset + 1
End Property

Public Property Get PWord(Optional ByVal vOffset As Long = -1) As Long
    If vOffset <> -1 Then m_Offset = vOffset
    If Overflow = False And m_Offset > (m_Size - 2) Then Exit Property
    GetMem2_Wrd m_Ptr + m_Offset, PWord
    m_Offset = m_Offset + 2
End Property

Public Property Let PWord(Optional ByVal vOffset As Long = -1, ByVal value As Long)
    If vOffset <> -1 Then m_Offset = vOffset
    If Overflow = False And m_Offset > (m_Size - 2) Then Exit Property
    PutMem2_Wrd m_Ptr + m_Offset, value
    m_Offset = m_Offset + 2
End Property

Public Property Get PInteger(Optional ByVal vOffset As Long = -1) As Integer
    If vOffset <> -1 Then m_Offset = vOffset
    If Overflow = False And m_Offset > (m_Size - 2) Then Exit Property
    GetMem2 m_Ptr + m_Offset, PInteger
    m_Offset = m_Offset + 2
End Property

Public Property Let PInteger(Optional ByVal vOffset As Long = -1, ByVal value As Integer)
    If vOffset <> -1 Then m_Offset = vOffset
    If Overflow = False And m_Offset > (m_Size - 2) Then Exit Property
    PutMem2 m_Ptr + m_Offset, value
    m_Offset = m_Offset + 2
End Property

Public Property Get PBoolean(Optional ByVal vOffset As Long = -1) As Boolean
    If vOffset <> -1 Then m_Offset = vOffset
    If Overflow = False And m_Offset > (m_Size - 2) Then Exit Property
    GetMem2_Bln m_Ptr + m_Offset, PBoolean
    m_Offset = m_Offset + 2
End Property

Public Property Let PBoolean(Optional ByVal vOffset As Long = -1, ByVal value As Boolean)
    If vOffset <> -1 Then m_Offset = vOffset
    If Overflow = False And m_Offset > (m_Size - 2) Then Exit Property
    PutMem2_Bln m_Ptr + m_Offset, value
    m_Offset = m_Offset + 2
End Property

Public Property Get PLong(Optional ByVal vOffset As Long = -1) As Long
    If vOffset <> -1 Then m_Offset = vOffset
    If Overflow = False And m_Offset > (m_Size - 4) Then Exit Property
    GetMem4 m_Ptr + m_Offset, PLong
    m_Offset = m_Offset + 4
End Property

Public Property Let PLong(Optional ByVal vOffset As Long = -1, ByVal value As Long)
    If vOffset <> -1 Then m_Offset = vOffset
    If Overflow = False And m_Offset > (m_Size - 4) Then Exit Property
    PutMem4 m_Ptr + m_Offset, value
    m_Offset = m_Offset + 4
End Property

Public Property Get PSingle(Optional ByVal vOffset As Long = -1) As Single
    If vOffset <> -1 Then m_Offset = vOffset
    If Overflow = False And m_Offset > (m_Size - 4) Then Exit Property
    GetMem4_Sng m_Ptr + m_Offset, PSingle
    m_Offset = m_Offset + 4
End Property

Public Property Let PSingle(Optional ByVal vOffset As Long = -1, ByVal value As Single)
    If vOffset <> -1 Then m_Offset = vOffset
    If Overflow = False And m_Offset > (m_Size - 4) Then Exit Property
    PutMem4_Sng m_Ptr + m_Offset, value
    m_Offset = m_Offset + 4
End Property

Public Property Get PDouble(Optional ByVal vOffset As Long = -1) As Double
    If vOffset <> -1 Then m_Offset = vOffset
    If Overflow = False And m_Offset > (m_Size - 8) Then Exit Property
    GetMem8_Dbl m_Ptr + m_Offset, PDouble
    m_Offset = m_Offset + 8
End Property

Public Property Let PDouble(Optional ByVal vOffset As Long = -1, ByVal value As Double)
    If vOffset <> -1 Then m_Offset = vOffset
    If Overflow = False And m_Offset > (m_Size - 8) Then Exit Property
    PutMem8_Dbl m_Ptr + m_Offset, value
    m_Offset = m_Offset + 8
End Property

Public Property Get PCurrency(Optional ByVal vOffset As Long = -1) As Currency
    If vOffset <> -1 Then m_Offset = vOffset
    If Overflow = False And m_Offset > (m_Size - 8) Then Exit Property
    GetMem8 m_Ptr + m_Offset, PCurrency
    m_Offset = m_Offset + 8
End Property

Public Property Let PCurrency(Optional ByVal vOffset As Long = -1, ByVal value As Currency)
    If vOffset <> -1 Then m_Offset = vOffset
    If Overflow = False And m_Offset > (m_Size - 8) Then Exit Property
    PutMem8 m_Ptr + vOffset, value
    m_Offset = m_Offset + 8
End Property

Public Property Get PString(Optional ByVal vOffset As Long = -1, Optional ByVal lenStr As Long) As String
    If vOffset <> -1 Then m_Offset = vOffset
    If lenStr = 0 Then lenStr = m_Size - m_Offset
    If (Overflow = False And m_Offset > (m_Size - Abs(lenStr))) Or (lenStr = 0) Then Exit Property
    PString = String$(Abs(lenStr), 0)
    CopyMemory ByVal PString, ByVal m_Ptr + m_Offset, Abs(lenStr)
    If lenStr < 0 Then PString = TrimNull(PString)
    m_Offset = m_Offset + Abs(lenStr)
End Property

Public Property Let PString(Optional ByVal vOffset As Long = -1, Optional ByVal lenStr As Long, ByVal value As String)
    If vOffset <> -1 Then m_Offset = vOffset
    If lenStr = 0 Then lenStr = Len(value)
    If (Overflow = False And m_Offset > (m_Size - Abs(lenStr))) Or (lenStr = 0) Then Exit Property
    lenStr = Abs(lenStr)
    If Len(value) < lenStr Then value = value & String$(lenStr - Len(value), 0)
    CopyMemory ByVal m_Ptr + m_Offset, ByVal value, lenStr
    m_Offset = m_Offset + lenStr
End Property

Public Property Get PArray(Optional ByVal vOffset As Long = -1, Optional ByVal lenArr As Long) As Byte()
    If vOffset <> -1 Then m_Offset = vOffset
    If (Overflow = False And m_Offset > (m_Size - lenArr)) Or (lenArr <= 0) Then Exit Property
    ReDim value(lenArr - 1) As Byte
    CopyMemory value(0), Buffer(m_Offset), lenArr
    m_Offset = m_Offset + lenArr
End Property

Public Property Let PArray(Optional ByVal vOffset As Long = -1, Optional ByVal lenArr As Long, value() As Byte)
    If vOffset <> -1 Then m_Offset = vOffset
    If (Overflow = False And m_Offset > (m_Size - lenArr)) Or (lenArr <= 0) Then Exit Property
    CopyMemory Buffer(m_Offset), value(0), lenArr
    m_Offset = m_Offset + lenArr
End Property
