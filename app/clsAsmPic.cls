VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAsmPic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const chRGB As Long = 0, chBlue As Long = 1, chGreen As Long = 2, chRed As Long = 3, chAlpha As Long = 4

Private bi As BITMAPINFO, BX As Long, BY As Long, BW As Long, BH As Long, B3D() As Byte

Public Obj As Variant, ASM As Object


Public Function GetDI(Optional ByVal invHeight As Boolean = True) As Long
    Dim vHDC As Long, vHBitmap As Long, vWidth As Long, vHeight As Long
    
    If IsEmpty(Obj) Then Exit Function
    
    If IsObject(Obj) Then
        If ExistsMember(Obj, "ScaleWidth") Then
            vHDC = Obj.hDC:         vHBitmap = Obj.Image:      vWidth = Obj.ScaleWidth:       vHeight = Obj.ScaleHeight
        Else
            vHDC = frmScript.hDC:   vHBitmap = Obj.Handle:     vWidth = GetHPX(Obj.Width):    vHeight = GetHPY(Obj.Height)
        End If
    Else
        vHDC = Obj(0):              vHBitmap = Obj(1):         vWidth = Obj(2):               vHeight = Obj(3)
        If vWidth < 0 Then vWidth = GetHPX(-vWidth)
        If vHeight < 0 Then vHeight = GetHPY(-vHeight)
    End If
    
    If invHeight Then vHeight = -vHeight
    
    With bi.bmiHeader
        .biSize = 40
        .biPlanes = 1
        .biBitCount = 32
        .biCompression = 0
        .biWidth = vWidth
        .biHeight = vHeight
        
        BX = 1:   BY = 1:   BW = .biWidth:   BH = Abs(.biHeight)
        
        ReDim B3D(1 To 4, 1 To BW, 1 To BH)
    End With

    GetDI = GetDIBits(vHDC, vHBitmap, 0, Abs(bi.bmiHeader.biHeight), B3D(1, 1, 1), bi, 0)
End Function

Public Function SetDI() As Long
    Dim vHDC As Long, vHBitmap As Long
    
    If IsEmpty(Obj) Then Exit Function
    
    If IsObject(Obj) Then
        If ExistsMember(Obj, "ScaleWidth") Then
            vHDC = Obj.hDC:         vHBitmap = Obj.Image
        Else
            vHDC = frmScript.hDC:   vHBitmap = Obj.Handle
        End If
    Else
        vHDC = Obj(0):              vHBitmap = Obj(1)
    End If
    
    SetDI = SetDIBits(vHDC, vHBitmap, 0, Abs(bi.bmiHeader.biHeight), B3D(1, 1, 1), bi, 0)
End Function

Public Function Ptr(Optional ByVal value As Long = 0) As Long
    Select Case value
        Case -1:    Ptr = VarPtr(bi)
        Case 0:     Ptr = VarPtr(B3D(1, 1, 1))
        Case 1:     Ptr = VarPtrArray(B3D)
    End Select
End Function

Public Function GetBuffer() As Byte()
    GetBuffer = B3D
End Function

Public Sub SetBuffer(Buf As Variant)
    B3D = Buf
End Sub

'Default Member
Public Property Get Point(ByVal x As Long, ByVal y As Long) As Long
Attribute Point.VB_UserMemId = 0
    CopyMem4 B3D(1, x, y), Point
End Property

Public Property Let Point(ByVal x As Long, ByVal y As Long, ByVal value As Long)
    CopyMem4 value, B3D(1, x, y)
End Property

Public Property Get Width() As Long
    Width = bi.bmiHeader.biWidth
End Property

Public Property Get Height() As Long
    Height = Abs(bi.bmiHeader.biHeight)
End Property

Public Function GetRect() As Variant
    GetRect = Array(BX, BY, BW, BH)
End Function

Public Sub SetRect(Optional ByVal typeX As Single = -2, Optional ByVal typeY As Single = -2, Optional ByVal offsetX As Single = 0, Optional ByVal offsetY As Single = 0, Optional ByVal typeW As Single = 1, Optional ByVal typeH As Single = 1)
    DoRect BX, BY, BW, BH, typeX, typeY, offsetX, offsetY, typeW, typeH
End Sub

Public Sub Main(ByVal numFunc As Long, Optional ByVal channel As Long = 0, Optional ByVal value As Byte = 0)
    Dim x As Long, y As Long, GT_Tmp(255) As Byte
    
    If channel <> chRGB Then
        Select Case numFunc
            Case 1      'Bits
                For x = 0 To 255
                    GT_Tmp(x) = (x \ value) * value
                Next
                For y = BY To BY + BH - 1
                    For x = BX To BX + BW - 1
                        B3D(channel, x, y) = GT_Tmp(B3D(channel, x, y))
                    Next
                Next
                
            Case 2      'And
                For y = BY To BY + BH - 1
                    For x = BX To BX + BW - 1
                        B3D(channel, x, y) = B3D(channel, x, y) And value
                    Next
                Next
                
            Case 3      'Or
                For y = BY To BY + BH - 1
                    For x = BX To BX + BW - 1
                        B3D(channel, x, y) = B3D(channel, x, y) Or value
                    Next
                Next
                
            Case 4      'Xor
                For y = BY To BY + BH - 1
                    For x = BX To BX + BW - 1
                        B3D(channel, x, y) = B3D(channel, x, y) Xor value
                    Next
                Next
        End Select
    Else
        Call Main(numFunc, chBlue, value)
        Call Main(numFunc, chGreen, value)
        Call Main(numFunc, chRed, value)
    End If
End Sub

Public Sub Gray(Optional ByVal pR As Integer = 33, Optional ByVal pg As Integer = 34, Optional ByVal pb As Integer = 33)
    Dim x As Long, y As Long, b As Integer
    For y = BY To BY + BH - 1
        For x = BX To BX + BW - 1
            b = GT_Tran(B3D(3, x, y), pR) + GT_Tran(B3D(2, x, y), pg) + GT_Tran(B3D(1, x, y), pb)
            b = GT_Mix(b)
            B3D(1, x, y) = b
            B3D(2, x, y) = b
            B3D(3, x, y) = b
        Next
    Next
End Sub

Public Sub Contrast(ByVal channel As Long, Optional ByVal value As Long = 0)
    Dim x As Long, y As Long, GT_Tmp(255) As Byte
    
    For x = 0 To 255
        y = x + (((x - 127) * value) \ 200)
        If y < 0 Then y = 0
        If y > 255 Then y = 255
        GT_Tmp(x) = y
    Next
                
    Call DoLevels(channel, GT_Tmp)
End Sub

Public Sub Curves(ByVal channel As Long, ParamArray Arg() As Variant)
    Dim a As Long, b As Integer, c As Long, m_Sub As Long, m_KF As Single, GT_Tmp(255) As Byte

    For a = 0 To 255
        GT_Tmp(a) = a
    Next
    
    For a = 0 To UBound(Arg)
        m_Sub = Arg(a)(2)
        m_KF = Arg(a)(3)
        For c = Arg(a)(0) To Arg(a)(1)
            b = m_Sub - c * m_KF
            If b < 0 Then b = 0
            If b > 255 Then b = 255
            GT_Tmp(c) = b
        Next
    Next
    
    Call DoLevels(channel, GT_Tmp)
End Sub

Public Sub Levels(ByVal channel As Long, Optional ByVal m_IL As Long = 0, Optional ByVal m_IR As Long = 255, Optional ByVal m_OL As Long = 0, Optional ByVal m_OR As Long = 255, Optional ByVal m_KF As Single = 1)
    Dim pStep As Single, oStep As Double, dGamma As Double, a As Long, b As Integer, gLevels(255) As Byte, GT_Tmp(255) As Byte
        
    For a = 0 To 255
        dGamma = CDbl(a) / 255
        dGamma = dGamma ^ (1 / m_KF)
        dGamma = dGamma * 255
        If dGamma > 255 Then dGamma = 255
        If dGamma < 0 Then dGamma = 0
        gLevels(a) = dGamma
    Next
    
    pStep = 255 / (CSng(m_IR) - CSng(m_IL))
    For a = 0 To 255
        If a < m_IL Then
            GT_Tmp(a) = 0
        ElseIf a > m_IR Then
            GT_Tmp(a) = 255
        Else
            b = (CSng(a) - CSng(m_IL)) * pStep
            If b < 0 Then b = 0
            If b > 255 Then b = 255
            GT_Tmp(a) = b
        End If
        GT_Tmp(a) = gLevels(GT_Tmp(a))
    Next
    
    oStep = (CSng(m_OR) - CSng(m_OL)) / 255
    For a = 0 To 255
        b = CSng(m_OL) + CSng(GT_Tmp(a)) * oStep
        If b < 0 Then b = 0
        If b > 255 Then b = 255
        GT_Tmp(a) = b
    Next
    
    Call DoLevels(channel, GT_Tmp)
End Sub

Public Sub Mix(Optional ByVal CR_PR As Long = 100, Optional ByVal CR_PG As Long = 0, Optional ByVal CR_PB As Long = 0, Optional ByVal CG_PR As Long = 0, Optional ByVal CG_PG As Long = 100, Optional ByVal CG_PB As Long = 0, Optional ByVal CB_PR As Long = 0, Optional ByVal CB_PG As Long = 0, Optional ByVal CB_PB As Long = 100)
    Dim x As Long, y As Long, r As Byte, g As Byte, b As Byte, o As Integer
    
    For y = BY To BY + BH - 1
        For x = BX To BX + BW - 1
            b = B3D(1, x, y)
            g = B3D(2, x, y)
            r = B3D(3, x, y)

            o = GT_Tran(r, CR_PR)
            o = o + GT_Tran(g, CR_PG)
            o = o + GT_Tran(b, CR_PB)
            B3D(3, x, y) = GT_Mix(o)

            o = GT_Tran(r, CG_PR)
            o = o + GT_Tran(g, CG_PG)
            o = o + GT_Tran(b, CG_PB)
            B3D(2, x, y) = GT_Mix(o)

            o = GT_Tran(r, CB_PR)
            o = o + GT_Tran(g, CB_PG)
            o = o + GT_Tran(b, CB_PB)
            B3D(1, x, y) = GT_Mix(o)
        Next
    Next
End Sub

Public Sub Gradient(ParamArray Arg() As Variant)
    Dim GT_Tmp(255) As RGBQUAD, c1 As RGBQUAD, c2 As RGBQUAD, x As Long, y As Long, a As Long, c As Long, id1 As Long, id2 As Long
    Dim r As Single, g As Single, b As Single, kr As Single, kg As Single, kb As Single

    For a = 0 To UBound(Arg) - 1
        id1 = Arg(a)(0)
        id2 = Arg(a + 1)(0)
        
        If UBound(Arg(a)) > 1 Then
            c1.rgbRed = Arg(a)(1)
            c1.rgbGreen = Arg(a)(2)
            c1.rgbBlue = Arg(a)(3)
        Else
            c = Arg(a)(1)
            CopyMem4 c, c1
        End If

        If UBound(Arg(a + 1)) > 1 Then
            c2.rgbRed = Arg(a + 1)(1)
            c2.rgbGreen = Arg(a + 1)(2)
            c2.rgbBlue = Arg(a + 1)(3)
        Else
            c = Arg(a + 1)(1)
            CopyMem4 c, c2
        End If

        kr = (CInt(c2.rgbRed) - CInt(c1.rgbRed)) / (id2 - id1)
        kg = (CInt(c2.rgbGreen) - CInt(c1.rgbGreen)) / (id2 - id1)
        kb = (CInt(c2.rgbBlue) - CInt(c1.rgbBlue)) / (id2 - id1)
        
        r = c1.rgbRed
        g = c1.rgbGreen
        b = c1.rgbBlue

        For c = id1 To id2
            GT_Tmp(c).rgbRed = r
            GT_Tmp(c).rgbGreen = g
            GT_Tmp(c).rgbBlue = b
            r = r + kr
            g = g + kg
            b = b + kb
        Next
    Next

    For y = BY To BY + BH - 1
        For x = BX To BX + BW - 1
            a = GT_Tran(B3D(3, x, y), 33) + GT_Tran(B3D(2, x, y), 34) + GT_Tran(B3D(1, x, y), 33)
            a = GT_Mix(a)
            
            With GT_Tmp(a)
                B3D(1, x, y) = .rgbBlue
                B3D(2, x, y) = .rgbGreen
                B3D(3, x, y) = .rgbRed
            End With
        Next
    Next
End Sub

Public Function Gistogram(ByVal channel As Long, Optional iMax As Variant, Optional iMin As Variant, Optional iBrs As Variant)
    Dim gLevels(255) As Long, Buf() As Variant, x As Long, y As Long, vMax As Long, vMin As Long, vBrs As Long, b As Byte
    
    vMax = 0: vMin = 2147483647
    
    Select Case channel
         Case chRGB
            For y = BY To BY + BH - 1
                For x = BX To BX + BW - 1
                    b = B3D(1, x, y)
                    vBrs = (vBrs + b) / 2
                    gLevels(b) = gLevels(b) + 1
                    b = B3D(2, x, y)
                    vBrs = (vBrs + b) / 2
                    gLevels(b) = gLevels(b) + 1
                    b = B3D(3, x, y)
                    vBrs = (vBrs + b) / 2
                    gLevels(b) = gLevels(b) + 1
                Next
            Next
            
        Case chBlue To chAlpha
            For y = BY To BY + BH - 1
                For x = BX To BX + BW - 1
                    b = B3D(channel, x, y)
                    vBrs = (vBrs + b) / 2
                    gLevels(b) = gLevels(b) + 1
                Next
            Next
    End Select
    
    ReDim Buf(255)
    For x = 0 To 255
        y = gLevels(x)
        Buf(x) = y
        If y > vMax Then vMax = y
        If y < vMin Then vMin = y
    Next
    
    iBrs = vBrs
    iMax = vMax
    iMin = vMin
    
    Gistogram = Buf
End Function

Public Sub Matrix(Optional ByVal value As Variant)
    Dim x As Long, y As Long, cx As Long, cy As Long, mx As Long, my As Long, r As Long, g As Long, b As Long
    Dim sum As Long, d As Long, m As Long, w As Long, h As Long, k As Long, Mat() As Long, B3D_Tmp() As Byte

    d = Sqr(ArraySize(value)) - 1
    If d < 0 Then Exit Sub
    B3D_Tmp = B3D:    m = d \ 2:    ReDim Mat(d, d):    cx = 10000:    cy = -10000:    w = BX + BW - 1:    h = BY + BH - 1
    
    For y = 0 To d
        For x = 0 To d
            k = value(x + y * (d + 1)):    Mat(x, y) = k:    sum = sum + k
            If k < cx Then cx = k
            If k > cy Then cy = k
        Next
    Next
    
    If sum <= 0 Then sum = 1
    
    For y = BY To h
        For x = BX To w
        
            r = 0:   g = 0:   b = 0
                 
            For my = 0 To d
                For mx = 0 To d
                    cx = x + mx - m:    cy = y + my - m:    k = Mat(mx, my)
                    
                    If k <> 0 Then
                        If cx < BX Then cx = BX
                        If cx > w Then cx = w
                        If cy < BY Then cy = BY
                        If cy > h Then cy = h
                        
                        r = r + B3D(3, cx, cy) * k
                        g = g + B3D(2, cx, cy) * k
                        b = b + B3D(1, cx, cy) * k
                    End If
                Next
            Next
            
            r = r / sum:    g = g / sum:    b = b / sum
            
            If r < 0 Then r = 0
            If r > 255 Then r = 255
            If g < 0 Then g = 0
            If g > 255 Then g = 255
            If b < 0 Then b = 0
            If b > 255 Then b = 255
            
            B3D_Tmp(3, x, y) = r
            B3D_Tmp(2, x, y) = g
            B3D_Tmp(1, x, y) = b
        Next
    Next
    
    B3D = B3D_Tmp
End Sub

Public Sub Replace(Optional ByVal fColor As Long = 0, Optional ByVal rColor As Long = 0, Optional ByVal RP As Long = 8, Optional ByVal GP As Long = 8, Optional ByVal BP As Long = 8)
    Dim fc_r As Byte, fc_g As Byte, fc_b As Byte, rc_r As Byte, rc_g As Byte, rc_b As Byte
    Dim x As Long, y As Long, r As Integer, g As Integer, b As Integer
    
    fc_b = fColor And 255&: fColor = fColor \ 256
    fc_g = fColor And 255&: fColor = fColor \ 256
    fc_r = fColor And 255&
    
    rc_b = rColor And 255&: rColor = rColor \ 256
    rc_g = rColor And 255&: rColor = rColor \ 256
    rc_r = rColor And 255&

    For y = BY To BY + BH - 1
        For x = BX To BX + BW - 1
            b = B3D(1, x, y)
            g = B3D(2, x, y)
            r = B3D(3, x, y)
            
            If Abs(b - fc_b) < BP Then
                If Abs(g - fc_g) < GP Then
                    If Abs(r - fc_r) < RP Then
                        B3D(1, x, y) = rc_b
                        B3D(2, x, y) = rc_g
                        B3D(3, x, y) = rc_r
                    End If
                End If
            End If
        Next
    Next
End Sub

Public Sub Table(Optional channel As Variant = chRGB)
    Static GT_Tmp(255) As Byte
    Dim a As Long
    
    If IsArray(channel) Then
        If Not VerifyArrayRange(channel, , , , 256) Then Exit Sub
    
        For a = 0 To 255
            GT_Tmp(a) = a
        Next
        
        For a = 0 To UBound(channel)
            GT_Tmp(a) = channel(a)
        Next
    Else
        Call DoLevels(channel, GT_Tmp)
    End If
End Sub

Public Sub Copy(srcBuf As Variant, Optional ByVal dstX As Long, Optional ByVal dstY As Long, Optional ByVal srcX As Long, Optional ByVal srcY As Long, Optional ByVal srcWidth As Long, Optional ByVal srcHeight As Long, Optional ByVal typeOper As Long = 0, Optional ByVal maskColor As Long = -1, Optional ByVal vAlpha As Long = 255)
    Dim x As Long, y As Long, x1 As Long, y1 As Long, x2 As Long, y2 As Long, x3 As Long, y3 As Long
    Dim r As Long, b As Long, g As Long, iMask As Long, SA As SafeArray, Buf() As Byte
    Dim sw As Long, sh As Long, dw As Long, dh As Long
    
    SA = GetSafeArray(srcBuf)
    
    With SA
        If .cDims = 3 And .cbElements = 1 And .rgSABound(2).cElements = 4 Then
            sw = .rgSABound(1).cElements
            sh = .rgSABound(0).cElements
        End If
    End With
    
    If sw = 0 Or sh = 0 Then Exit Sub
    
    PutMem4 VarPtrArray(Buf), VarPtr(SA)
    
    dw = UBound(B3D, 2)
    dh = UBound(B3D, 3)


    '-------------------- Correct coord -------------------------
    CorrRect srcX, srcY, srcWidth, srcHeight, sw, sh, True
    
    x1 = dstX:  x2 = dstX + srcWidth - 1
    y1 = dstY:  y2 = dstY + srcHeight - 1
    
    If (x2 > 0) And (y2 > 0) And (x1 <= dw) And (y1 <= dh) Then
    
        CorrRect x1, y1, 0&, 0&, dw, dh
        CorrRect x2, y2, 0&, 0&, dw, dh
        
        srcX = srcX - x1
        srcY = srcY - y1
        
        '-----------------------------------------------------
        dw = (vAlpha And 256) > 0:   vAlpha = vAlpha And 255:   sh = vAlpha:   sw = 255 - sh
        
        For y = y1 To y2
            For x = x1 To x2
                x3 = x + srcX:   y3 = y + srcY

                r = Buf(3, x3, y3):    g = Buf(2, x3, y3):    b = Buf(1, x3, y3)
                If maskColor >= 0 Then iMask = r * 65536 + g * 256 + b

                If maskColor <> iMask Then
                    If typeOper = 0 Then                    'vbSrcCopy + Alpha
                        If dw Then sh = GT_Grad(Buf(4, x3, y3), vAlpha):   sw = 255 - sh
                        
                    ElseIf typeOper = &H8800C6 Then         'vbSrcAnd
                        r = B3D(3, x, y) And r
                        g = B3D(2, x, y) And g
                        b = B3D(1, x, y) And b
                        
                    ElseIf typeOper = &HEE0086 Then         'vbSrcPaint
                        r = B3D(3, x, y) Or r
                        g = B3D(2, x, y) Or g
                        b = B3D(1, x, y) Or b
                        
                    ElseIf typeOper = &H660046 Then         'vbSrcInvert
                        r = B3D(3, x, y) Xor r
                        g = B3D(2, x, y) Xor g
                        b = B3D(1, x, y) Xor b
                    End If
                    
                    B3D(3, x, y) = GT_Grad(B3D(3, x, y), sw) + GT_Grad(r, sh)
                    B3D(2, x, y) = GT_Grad(B3D(2, x, y), sw) + GT_Grad(g, sh)
                    B3D(1, x, y) = GT_Grad(B3D(1, x, y), sw) + GT_Grad(b, sh)
                End If
            Next
        Next
    End If
    
    PutMem4 VarPtrArray(Buf), 0
End Sub

Private Sub DoLevels(ByVal channel As Long, GT_Tmp() As Byte)
    Dim x As Long, y As Long
    
    Select Case channel
         Case chRGB
            For y = BY To BY + BH - 1
                For x = BX To BX + BW - 1
                    B3D(3, x, y) = GT_Tmp(B3D(3, x, y))
                    B3D(2, x, y) = GT_Tmp(B3D(2, x, y))
                    B3D(1, x, y) = GT_Tmp(B3D(1, x, y))
                Next
            Next
            
        Case chBlue To chAlpha
            For y = BY To BY + BH - 1
                For x = BX To BX + BW - 1
                    B3D(channel, x, y) = GT_Tmp(B3D(channel, x, y))
                Next
            Next
    End Select
End Sub

Private Sub DoRect(x As Long, y As Long, w As Long, h As Long, typeX As Single, typeY As Single, offsetX As Single, offsetY As Single, typeW As Single, typeH As Single)
    Dim iWidth As Long, iHeight As Long
    
    iWidth = bi.bmiHeader.biWidth
    iHeight = Abs(bi.bmiHeader.biHeight)
    
    Select Case typeW
        Case Is > 1:    w = typeW
        Case Is > 0:    w = iWidth * typeW
        Case Is < 0:    w = iWidth + typeW
    End Select
    
    Select Case typeH
        Case Is > 1:    h = typeH
        Case Is > 0:    h = iHeight * typeH
        Case Is < 0:    h = iHeight + typeH
    End Select

    Select Case typeX
        Case -1:        x = iWidth / 2 - w / 2 + offsetX
        Case Is > 0:    x = iWidth * typeX + offsetX
        Case -1.1:      x = iWidth / 2 + offsetX
        Case -1.2:      x = iWidth / 2 - w + offsetX
        Case -2:        x = offsetX
        Case -3:        x = iWidth - w + offsetX
    End Select

    Select Case typeY
        Case -1:        y = iHeight / 2 - h / 2 + offsetY
        Case Is > 0:    y = iHeight * typeY + offsetY
        Case -1.1:      y = iHeight / 2 + offsetY
        Case -1.2:      y = iHeight / 2 - h + offsetY
        Case -2:        y = offsetY
        Case -3:        y = iHeight - h + offsetY
    End Select
    
    CorrRect x, y, w, h, iWidth, iHeight
End Sub

Private Sub CorrRect(x As Long, y As Long, w As Long, h As Long, ByVal iWidth As Long, ByVal iHeight As Long, Optional ByVal isAutoWH As Boolean)
    If isAutoWH Then
        If w = 0 Then w = iWidth
        If h = 0 Then h = iHeight
    End If
    
    If x < 1 Then x = 1
    If y < 1 Then y = 1
    If w < 1 Then w = 1
    If h < 1 Then h = 1
    If x > iWidth Then x = iWidth
    If y > iHeight Then y = iHeight
    If (x + w - 1) > iWidth Then w = iWidth - x + 1
    If (y + h - 1) > iHeight Then h = iHeight - y + 1
End Sub

Private Sub Class_Initialize()
    Set ASM = New clsFuncPointer
End Sub
