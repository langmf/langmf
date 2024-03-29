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

Private lOn(31) As Long, lPow(31) As Long


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

Public Function Long2Bin(ByVal value As Long, Optional ByVal vShort As Long = -32, Optional ByVal vStep As Long = 0, Optional ByVal vAdd As String = " ") As String
    Dim a As Long, r As String
    
    r = String$(32, "0")
    
    For a = 0 To 31
        If GetBit(value, a) Then Mid$(r, 32 - a, 1) = "1"
    Next
    
    r = Right$(r, Abs(vShort)):     If vShort < 0 Then a = InStr(r, "1"):  If a = 0 Then r = "0" Else r = Right$(r, Len(r) - a + 1)
    If vStep = 0 Then Long2Bin = r:    Exit Function
    
    For a = 1 To Len(r) Step vStep
        Long2Bin = Long2Bin & Mid$(r, a, vStep):        If (a + vStep) <= Len(r) Then Long2Bin = Long2Bin & vAdd
    Next
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

Public Function CRC(Buf As Variant, Optional ByVal addCrc As Boolean, Optional ByVal Alg As Long) As Variant
    Dim a As Long, b As Long, x As Long, v As Long, sz As Long, vCrc As Long
    
    sz = ArraySize(Buf):      If sz = 0 Then Exit Function
    
    If Alg = 0 Then                 ' Dallas
        vCrc = 0
        For a = 0 To sz - 1
            v = CByte(Buf(a))
            For b = 7 To 0 Step -1
                x = &H8C * ((vCrc Xor v) And 1):     vCrc = vCrc \ 2:     vCrc = vCrc Xor x:     v = v \ 2
            Next
        Next
        If addCrc Then ReDim Preserve Buf(sz):       Buf(sz) = vCrc
        
    ElseIf Alg = 1 Then             ' Modbus
        vCrc = &HFFFF&
        For a = 0 To sz - 1
            vCrc = vCrc Xor CByte(Buf(a))
            For b = 7 To 0 Step -1
                x = &HA001& * (vCrc And 1):          vCrc = vCrc \ 2:     vCrc = vCrc Xor x
            Next
        Next
        If addCrc Then ReDim Preserve Buf(sz + 1):   Buf(sz) = vCrc And 255:   Buf(sz + 1) = vCrc \ 256
    End If
    
    If addCrc Then CRC = Buf Else CRC = vCrc
End Function

Public Function NewArr(Buf As Variant, vsp() As Variant, Optional ByVal isPreserve As Boolean = False, Optional ByVal value As Long = -1) As Variant
    Dim tmpBuf() As Byte

    If VarType(Buf) = (vbArray + vbByte) Then tmpBuf = Buf
    
    If ArrayValid(vsp, , , 1, 3) Then
        If isPreserve Then
            Select Case UBound(vsp)
                Case 0:     ReDim Preserve tmpBuf(vsp(0))
                Case 1:     ReDim Preserve tmpBuf(vsp(0), vsp(1))
                Case 2:     ReDim Preserve tmpBuf(vsp(0), vsp(1), vsp(2))
            End Select
        Else
            Select Case UBound(vsp)
                Case 0:     ReDim tmpBuf(vsp(0))
                Case 1:     ReDim tmpBuf(vsp(0), vsp(1))
                Case 2:     ReDim tmpBuf(vsp(0), vsp(1), vsp(2))
            End Select
        End If
    End If
    
    If value >= 0 And value <= 255 Then
        With GetSafeArray(tmpBuf)
            If .pvData Then
                Select Case .cDims
                    Case 1:     FillMemory ByVal .pvData, .rgSABound(0).cElements, value
                    Case 2:     FillMemory ByVal .pvData, .rgSABound(0).cElements * .rgSABound(1).cElements, value
                    Case 3:     FillMemory ByVal .pvData, .rgSABound(0).cElements * .rgSABound(1).cElements * .rgSABound(2).cElements, value
                End Select
            End If
        End With
    End If
    
    NewArr = tmpBuf
End Function

Private Sub Class_Initialize()
    Dim a As Long, v As Long

    lOn(30) = &H7FFFFFFF:   lPow(30) = &H40000000:   lOn(31) = &HFFFFFFFF:   lPow(31) = &H80000000
    v = 1:   For a = 0 To 29:   lPow(a) = v:   v = v * 2:   lOn(a) = v - 1:   Next
End Sub
