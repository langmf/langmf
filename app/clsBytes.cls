VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBytes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private lOn(31) As Long, lPow(31) As Long, crcTable(255) As Byte


Public Function LShift(ByVal value As Long, ByVal bits As Long) As Long
    If bits < 0 Or bits > 31 Then Exit Function
    If bits = 0 Then LShift = value: Exit Function
    If bits = 31 Then If value And 1 Then LShift = &H80000000: Exit Function Else LShift = 0: Exit Function
    If (value And lPow(31 - bits)) Then LShift = ((value And lOn(31 - (bits + 1))) * lPow(bits)) Or &H80000000: Exit Function
    LShift = ((value And lOn(31 - bits)) * lPow(bits))
End Function

Public Function RShift(ByVal value As Long, ByVal bits As Long) As Long
    If bits < 0 Or bits > 31 Then Exit Function
    If bits = 0 Then RShift = value: Exit Function
    If bits = 31 Then If value And &H80000000 Then RShift = 1: Exit Function Else RShift = 0: Exit Function
    RShift = (value And &H7FFFFFFE) \ lPow(bits)
    If (value And &H80000000) Then RShift = (RShift Or (&H40000000 \ lPow(bits - 1)))
End Function

Public Function LRotate(ByVal value As Long, ByVal bits As Long, Optional ByVal szBit As Long = 32) As Long
    If bits > szBit Then bits = bits Mod szBit
    LRotate = LShift(value, bits) Or RShift(value, (szBit - bits))
    LRotate = LRotate And lOn(szBit - 1)
End Function

Public Function SetBit(ByVal value As Long, ByVal bits As Long, Optional ByVal iValue As Long = 1) As Long
    If bits < 0 Or bits > 31 Then Exit Function
    If iValue = 0 Then value = value And (Not lPow(bits)) Else value = value Or (lPow(bits))
    SetBit = value
End Function

Public Function GetBit(ByVal value As Long, ByVal bits As Long) As Long
    If bits < 0 Or bits > 31 Then Exit Function
    If value And lPow(bits) Then GetBit = 1 Else GetBit = 0
End Function

Public Function Bin2Long(ByVal value As String) As Long
    Dim a As Long, i As Long, s As String, lenTxt As Long

    lenTxt = Len(value)
    
    For a = lenTxt To 1 Step -1
        s = Mid$(value, a, 1):      If i > 31 Then Exit For
        If s = "0" Then i = i + 1
        If s = "1" Then Bin2Long = Bin2Long Or lPow(i):   i = i + 1
    Next
End Function

Public Function Long2Bin(ByVal value As Long, Optional ByVal vShort As Long = -32) As String
    Dim a As Long
    
    Long2Bin = String$(32, "0")
    
    For a = 0 To 31
        If GetBit(value, a) Then Mid$(Long2Bin, 32 - a, 1) = "1"
    Next
    
    Long2Bin = Right$(Long2Bin, Abs(vShort))
    
    If vShort >= 0 Then Exit Function
    a = InStr(Long2Bin, "1"):    If a = 0 Then Long2Bin = "0"
    If a > 1 Then Long2Bin = Right$(Long2Bin, Len(Long2Bin) - a + 1)
End Function

Public Function Byte2Long(ByVal num1 As Byte, ByVal num2 As Byte, ByVal num3 As Byte, ByVal num4 As Byte) As Long
    Dim b(3) As Byte
    b(0) = num4:  b(1) = num3:  b(2) = num2:  b(3) = num1:    CopyMem4 b(0), Byte2Long
End Function

Public Sub Long2Byte(ByVal mLong As Long, num1 As Variant, num2 As Variant, num3 As Variant, num4 As Variant)
    Dim b(3) As Byte
    CopyMem4 mLong, b(0):    num1 = b(3):  num2 = b(2):  num3 = b(1):  num4 = b(0)
End Sub

Public Function Long2Integer(ByVal value As Long) As Integer
    GetMem2 VarPtr(value), Long2Integer
End Function

Public Function CRC(Buf As Variant, Optional ByVal addToEnd As Boolean) As Variant
    Dim a As Long, sz As Long, vCrc As Byte
    
    sz = m_ArraySize(Buf)
    If sz = 0 Then Exit Function
    
    For a = 0 To sz - 1
        vCrc = crcTable(vCrc Xor CByte(Buf(a)))
    Next
    
    CRC = vCrc
    If addToEnd Then ReDim Preserve Buf(sz):   Buf(sz) = vCrc:   CRC = Buf
End Function

Public Function Crypt(Optional value As Variant, Optional ByVal Flag As Long = -1) As Variant
    Dim a As Long, tmp As Byte, Buf() As Byte

    Static i As Long, s As Byte, p As Long, u As Long, Key() As Byte

    If Flag = 0 Or Flag = 1 Then
        If Flag = 0 Then ConvToBufferByte value, Key:    u = UBound(Key)
        For a = 0 To u:     tmp = crcTable(tmp Xor Key(a)):     Next
        i = u * 256 + tmp:    s = 0:    p = 0:    Exit Function
    End If

    ConvToBufferByte value, Buf
 
    For a = 0 To m_ArraySize(Buf) - 1
        tmp = Buf(a) Xor Key(p) Xor (i Mod 256) Xor (i \ 256) Xor s
        If Flag = 2 Then s = Buf(a)
        If Flag = 3 Then s = tmp
        Buf(a) = tmp
        p = p + 1
        If p > u Then p = 0: If i >= 65535 Then i = 0 Else i = i + 1
    Next
    
    Crypt = Buf
End Function

Public Sub ArrayRedim(Buf As Variant, vsp() As Variant, Optional ByVal isPreserve As Boolean = False, Optional ByVal value As Long = -1)
    Dim tmpBuf() As Byte
   
    If VarType(Buf) <> (vbArray + vbByte) Then Buf = tmpBuf
    
    If VerifyArrayRange(vsp, , , 1, 3) Then
        If isPreserve Then
            Select Case UBound(vsp)
                Case 0:     ReDim Preserve Buf(vsp(0))
                Case 1:     ReDim Preserve Buf(vsp(0), vsp(1))
                Case 2:     ReDim Preserve Buf(vsp(0), vsp(1), vsp(2))
            End Select
        Else
            Select Case UBound(vsp)
                Case 0:     ReDim Buf(vsp(0))
                Case 1:     ReDim Buf(vsp(0), vsp(1))
                Case 2:     ReDim Buf(vsp(0), vsp(1), vsp(2))
            End Select
        End If
    End If
    
    If VerifyLongRange(value, 0, 255) Then
        With m_GetSafeArray(Buf)
            If .pvData Then
                Select Case .cDims
                    Case 1:     FillMemory ByVal .pvData, .rgSABound(0).cElements, value
                    Case 2:     FillMemory ByVal .pvData, .rgSABound(0).cElements * .rgSABound(1).cElements, value
                    Case 3:     FillMemory ByVal .pvData, .rgSABound(0).cElements * .rgSABound(1).cElements * .rgSABound(2).cElements, value
                End Select
            End If
        End With
    End If
End Sub

Private Sub Class_Initialize()
    Dim a As Long, v As Long

    lOn(30) = &H7FFFFFFF:   lPow(30) = &H40000000:   lOn(31) = &HFFFFFFFF:   lPow(31) = &H80000000
    v = 1:   For a = 0 To 29:   lPow(a) = v:   v = v * 2:   lOn(a) = v - 1:   Next

    For a = 0 To 255
        v = 0
        If (a And &H1) Then v = v Xor &H5E
        If (a And &H2) Then v = v Xor &HBC
        If (a And &H4) Then v = v Xor &H61
        If (a And &H8) Then v = v Xor &HC2
        If (a And &H10) Then v = v Xor &H9D
        If (a And &H20) Then v = v Xor &H23
        If (a And &H40) Then v = v Xor &H46
        If (a And &H80) Then v = v Xor &H8C
        crcTable(a) = v
    Next
End Sub
