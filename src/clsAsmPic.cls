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

Const nm_Blue As Long = 1, nm_Green As Long = 2, nm_Red As Long = 3, nm_Alpha As Long = 4, nm_RGB As Long = 5
Const ch_Blue As Long = 16, ch_Green As Long = 32, ch_Red As Long = 64, ch_Alpha As Long = 128, ch_RGB As Long = 112

Private bi As BITMAPINFO, BX As Long, BY As Long, BW As Long, BH As Long, B3D() As Byte

Public Result As Variant, Obj As Variant, Bag As Object


Public Function GetDI(Optional ByVal invHeight As Boolean = True) As Object
    Dim vHDC As Long, vHBitmap As Long, vWidth As Long, vHeight As Long
    
    Set GetDI = Me:         Result = 0:         If IsEmpty(Obj) Then Exit Function
    
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

    Result = GetDIBits(vHDC, vHBitmap, 0, Abs(bi.bmiHeader.biHeight), B3D(1, 1, 1), bi, 0)
End Function

Public Function SetDI() As Object
    Dim vHDC As Long, vHBitmap As Long
    
    Set SetDI = Me:         Result = 0:         If IsEmpty(Obj) Then Exit Function
    
    If IsObject(Obj) Then
        If ExistsMember(Obj, "ScaleWidth") Then
            vHDC = Obj.hDC:         vHBitmap = Obj.Image
        Else
            vHDC = frmScript.hDC:   vHBitmap = Obj.Handle
        End If
    Else
        vHDC = Obj(0):              vHBitmap = Obj(1)
    End If
    
    Result = SetDIBits(vHDC, vHBitmap, 0, Abs(bi.bmiHeader.biHeight), B3D(1, 1, 1), bi, 0)
End Function


Public Property Get Ptr(Optional ByVal value As Long = 0) As Long
    Select Case value
        Case -1:    Ptr = VarPtr(bi)
        Case 0:     Ptr = VarPtr(B3D(1, 1, 1))
        Case 1:     Ptr = VarPtrArray(B3D)
    End Select
End Property

Public Property Get Buffer() As Variant
    Buffer = B3D
End Property

Public Property Let Buffer(Buf As Variant)
    B3D = Buf
End Property


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

Public Function GetRect() As Object
    Set GetRect = Me
    Result = Array(BX, BY, BW, BH)
End Function

Public Function SetRect(Optional ByVal typeX As Single = -2, Optional ByVal typeY As Single = -2, Optional ByVal typeW As Single = 1, Optional ByVal typeH As Single = 1, Optional ByVal offsetX As Single = 0, Optional ByVal offsetY As Single = 0) As Object
    Set SetRect = Me
    DoRect bi.bmiHeader.biWidth, Abs(bi.bmiHeader.biHeight), BX, BY, BW, BH, typeX, typeY, typeW, typeH, offsetX, offsetY
End Function


Public Function Main(ByVal numFunc As Long, Optional ByVal channel As Long = nm_RGB, Optional ByVal value As Byte = 0) As Object
    Dim x As Long, y As Long, ch As Long, GT_Tmp(255) As Byte
    
    Set Main = Me
    
    Do
        ch = NextChan(channel):      If ch = 0 Then Exit Do

        Select Case numFunc
            Case 1      'Bits
                For x = 0 To 255
                    GT_Tmp(x) = (x \ value) * value
                Next
                For y = BY To BY + BH - 1
                    For x = BX To BX + BW - 1
                        B3D(ch, x, y) = GT_Tmp(B3D(ch, x, y))
                    Next
                Next
                
            Case 2      'And
                For y = BY To BY + BH - 1
                    For x = BX To BX + BW - 1
                        B3D(ch, x, y) = B3D(ch, x, y) And value
                    Next
                Next
                
            Case 3      'Or
                For y = BY To BY + BH - 1
                    For x = BX To BX + BW - 1
                        B3D(ch, x, y) = B3D(ch, x, y) Or value
                    Next
                Next
                
            Case 4      'Xor
                For y = BY To BY + BH - 1
                    For x = BX To BX + BW - 1
                        B3D(ch, x, y) = B3D(ch, x, y) Xor value
                    Next
                Next
        End Select
    Loop
End Function

Public Function Gray(Optional ByVal pR As Integer = 33, Optional ByVal pg As Integer = 34, Optional ByVal pb As Integer = 33) As Object
    Dim x As Long, y As Long, b As Integer

    Set Gray = Me

    For y = BY To BY + BH - 1
        For x = BX To BX + BW - 1
            b = GT_Tran(B3D(3, x, y), pR) + GT_Tran(B3D(2, x, y), pg) + GT_Tran(B3D(1, x, y), pb)
            b = GT_Mix(b)
            B3D(3, x, y) = b
            B3D(2, x, y) = b
            B3D(1, x, y) = b
        Next
    Next
End Function

Public Function Contrast(ByVal channel As Long, Optional ByVal value As Long = 0) As Object
    Dim x As Long, y As Long, GT_Tmp(255) As Byte
    
    Set Contrast = Me

    For x = 0 To 255
        y = x + (((x - 127) * value) \ 200)
        If y < 0 Then y = 0
        If y > 255 Then y = 255
        GT_Tmp(x) = y
    Next
                
    Call DoLevels(channel, GT_Tmp)
End Function

Public Function Curves(ByVal channel As Long, ParamArray Arg() As Variant) As Object
    Dim a As Long, b As Long, c As Long, m_Sub As Long, m_KF As Single, GT_Tmp(255) As Byte

    Set Curves = Me

    For a = 0 To 255:       GT_Tmp(a) = a:      Next
    
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
End Function

Public Function Levels(ByVal channel As Long, Optional ByVal m_IL As Long = 0, Optional ByVal m_IR As Long = 255, Optional ByVal m_OL As Long = 0, Optional ByVal m_OR As Long = 255, Optional ByVal m_KF As Single = 1) As Object
    Dim pStep As Single, oStep As Double, dGamma As Double, a As Long, b As Long, m(255) As Byte, GT_Tmp(255) As Byte
        
    Set Levels = Me
    
    For a = 0 To 255
        dGamma = CDbl(a) / 255
        dGamma = dGamma ^ (1 / m_KF)
        dGamma = dGamma * 255
        If dGamma > 255 Then dGamma = 255
        If dGamma < 0 Then dGamma = 0
        m(a) = dGamma
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
        GT_Tmp(a) = m(GT_Tmp(a))
    Next
    
    oStep = (CSng(m_OR) - CSng(m_OL)) / 255
    For a = 0 To 255
        b = CSng(m_OL) + CSng(GT_Tmp(a)) * oStep
        If b < 0 Then b = 0
        If b > 255 Then b = 255
        GT_Tmp(a) = b
    Next
    
    Call DoLevels(channel, GT_Tmp)
End Function

Public Function Mix(Optional ByVal CR_PR As Long = 100, Optional ByVal CR_PG As Long = 0, Optional ByVal CR_PB As Long = 0, Optional ByVal CG_PR As Long = 0, Optional ByVal CG_PG As Long = 100, Optional ByVal CG_PB As Long = 0, Optional ByVal CB_PR As Long = 0, Optional ByVal CB_PG As Long = 0, Optional ByVal CB_PB As Long = 100) As Object
    Dim x As Long, y As Long, r As Byte, g As Byte, b As Byte, o As Integer
    
    Set Mix = Me
    
    For y = BY To BY + BH - 1
        For x = BX To BX + BW - 1
            r = B3D(3, x, y)
            g = B3D(2, x, y)
            b = B3D(1, x, y)

            o = GT_Tran(r, CR_PR):    o = o + GT_Tran(g, CR_PG):    o = o + GT_Tran(b, CR_PB):    B3D(3, x, y) = GT_Mix(o)
            o = GT_Tran(r, CG_PR):    o = o + GT_Tran(g, CG_PG):    o = o + GT_Tran(b, CG_PB):    B3D(2, x, y) = GT_Mix(o)
            o = GT_Tran(r, CB_PR):    o = o + GT_Tran(g, CB_PG):    o = o + GT_Tran(b, CB_PB):    B3D(1, x, y) = GT_Mix(o)
        Next
    Next
End Function

Public Function Gradient(ParamArray Arg() As Variant) As Object
    Dim GT_Tmp(255) As RGBQUAD, c1 As RGBQUAD, c2 As RGBQUAD, x As Long, y As Long, a As Long, c As Long, id1 As Long, id2 As Long
    Dim r As Single, g As Single, b As Single, kr As Single, kg As Single, kb As Single

    Set Gradient = Me
    
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
                B3D(3, x, y) = .rgbRed
                B3D(2, x, y) = .rgbGreen
                B3D(1, x, y) = .rgbBlue
            End With
        Next
    Next
End Function

Public Function Gistogram(ByVal channel As Long) As Object
    Dim m(255) As Long, Buf() As Variant, x As Long, y As Long, vMax As Long, vMin As Long, vBrs As Long, b As Byte
    
    Set Gistogram = Me:         vMax = 0:       vMin = 2147483647:      NextChan channel, 2
    
    For y = BY To BY + BH - 1
        For x = BX To BX + BW - 1
            If (channel And ch_Alpha) Then b = B3D(4, x, y):    vBrs = (vBrs + b) / 2:      m(b) = m(b) + 1
            If (channel And ch_Red) Then b = B3D(3, x, y):      vBrs = (vBrs + b) / 2:      m(b) = m(b) + 1
            If (channel And ch_Green) Then b = B3D(2, x, y):    vBrs = (vBrs + b) / 2:      m(b) = m(b) + 1
            If (channel And ch_Blue) Then b = B3D(1, x, y):     vBrs = (vBrs + b) / 2:      m(b) = m(b) + 1
        Next
    Next
    
    ReDim Buf(255)
    For x = 0 To 255
        y = m(x)
        Buf(x) = y
        If y > vMax Then vMax = y
        If y < vMin Then vMin = y
    Next
    
    Result = Array(vMin, vMax, vBrs, Buf)
End Function

Public Function Matrix(Optional ByVal value As Variant) As Object
    Dim x As Long, y As Long, cx As Long, cy As Long, mx As Long, my As Long, r As Long, g As Long, b As Long
    Dim sum As Long, d As Long, m As Long, w As Long, h As Long, k As Long, Mat() As Long, B3D_Tmp() As Byte

    Set Matrix = Me

    d = Sqr(ArraySize(value)) - 1:      If d < 0 Then Exit Function

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
        
            r = 0:      g = 0:      b = 0
                 
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
            
            r = r / sum:        g = g / sum:        b = b / sum
            
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
End Function

Public Function Replace(Optional ByVal fColor As Long = 0, Optional ByVal rColor As Long = 0, Optional ByVal RP As Long = 8, Optional ByVal GP As Long = 8, Optional ByVal BP As Long = 8) As Object
    Dim fc_r As Byte, fc_g As Byte, fc_b As Byte, rc_r As Byte, rc_g As Byte, rc_b As Byte
    Dim x As Long, y As Long, r As Integer, g As Integer, b As Integer
    
    Set Replace = Me
    
    fc_b = fColor And 255&:     fColor = fColor \ 256
    fc_g = fColor And 255&:     fColor = fColor \ 256
    fc_r = fColor And 255&
    
    rc_b = rColor And 255&:     rColor = rColor \ 256
    rc_g = rColor And 255&:     rColor = rColor \ 256
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
End Function

Public Function Table(Optional ByVal channel As Long = nm_RGB, Optional TBL As Variant) As Object
    Static GT_Tmp(255) As Byte
    Dim a As Long
    
    Set Table = Me
    
    If IsArray(TBL) Then
        If Not ArrayValid(TBL, , , , 256) Then Exit Function
        
        For a = 0 To 255:               GT_Tmp(a) = a:          Next
        For a = 0 To UBound(TBL):       GT_Tmp(a) = TBL(a):     Next
    End If

    Call DoLevels(channel, GT_Tmp)
End Function

Public Function Copy(srcBuf As Variant, Optional ByVal dstX As Long, Optional ByVal dstY As Long, Optional ByVal srcX As Long, Optional ByVal srcY As Long, Optional ByVal srcWidth As Long, Optional ByVal srcHeight As Long, Optional ByVal typeOper As Long = 0, Optional ByVal maskColor As Long = -1, Optional ByVal vAlpha As Long = 255) As Object
    Dim x As Long, y As Long, x1 As Long, y1 As Long, x2 As Long, y2 As Long, x3 As Long, y3 As Long
    Dim r As Long, b As Long, g As Long, iMask As Long, SA As SafeArray, Buf() As Byte
    Dim sw As Long, sh As Long, dw As Long, dh As Long
    
    Set Copy = Me
    
    SA = GetSafeArray(srcBuf)
    
    With SA
        If .cDims = 3 And .cbElements = 1 And .rgSABound(2).cElements = 4 Then
            sw = .rgSABound(1).cElements
            sh = .rgSABound(0).cElements
        End If
    End With
    
    If sw = 0 Or sh = 0 Then Exit Function
    
    PutMem4 VarPtrArray(Buf), VarPtr(SA):       dw = UBound(B3D, 2):        dh = UBound(B3D, 3)

    '-------------------- Correct coord -------------------------
    CorrRect sw, sh, srcX, srcY, srcWidth, srcHeight, True
    
    x1 = dstX:      x2 = dstX + srcWidth - 1
    y1 = dstY:      y2 = dstY + srcHeight - 1
    
    '------------------------------------------------------------
    If (x2 > 0) And (y2 > 0) And (x1 <= dw) And (y1 <= dh) Then
    
        CorrRect dw, dh, x1, y1, 0&, 0&:       srcX = srcX - x1
        CorrRect dw, dh, x2, y2, 0&, 0&:       srcY = srcY - y1
        
        dw = (vAlpha And 256) > 0:      vAlpha = vAlpha And 255:    sh = vAlpha:    sw = 255 - sh
        
        For y = y1 To y2
            For x = x1 To x2
                x3 = x + srcX:      y3 = y + srcY:      r = Buf(3, x3, y3):     g = Buf(2, x3, y3):     b = Buf(1, x3, y3)

                If maskColor >= 0 Then iMask = r * 65536 + g * 256 + b

                If maskColor <> iMask Then
                    If typeOper = 0 Then                                    'vbSrcCopy + Alpha
                        If dw Then sh = GT_Grad(Buf(4, x3, y3), vAlpha):    sw = 255 - sh

                    ElseIf typeOper = &H8800C6 Then                         'vbSrcAnd
                        r = r And B3D(3, x, y)
                        g = g And B3D(2, x, y)
                        b = b And B3D(1, x, y)

                    ElseIf typeOper = &HEE0086 Then                         'vbSrcPaint
                        r = r Or B3D(3, x, y)
                        g = g Or B3D(2, x, y)
                        b = b Or B3D(1, x, y)

                    ElseIf typeOper = &H660046 Then                         'vbSrcInvert
                        r = r Xor B3D(3, x, y)
                        g = g Xor B3D(2, x, y)
                        b = b Xor B3D(1, x, y)
                    End If

                    B3D(3, x, y) = GT_Grad(B3D(3, x, y), sw) + GT_Grad(r, sh)
                    B3D(2, x, y) = GT_Grad(B3D(2, x, y), sw) + GT_Grad(g, sh)
                    B3D(1, x, y) = GT_Grad(B3D(1, x, y), sw) + GT_Grad(b, sh)
                End If
            Next
        Next
    End If
    
    PutMem4 VarPtrArray(Buf), 0
End Function


Private Function NextChan(c As Long, Optional ByVal Flag As Long) As Long
    Dim v As Long

    If c <= 0 Then Exit Function

    v = c And &HF:      c = c And &HF0
    If v = nm_RGB Then c = c Or ch_RGB
    If v >= nm_Blue And v <= nm_Alpha Then c = c Or (2 ^ (v + 3))
    If Flag = 2 Then Exit Function

    v = c And ch_RGB
    If Flag = 1 And v = ch_RGB Then c = c Xor ch_RGB:   NextChan = nm_RGB:       Exit Function

    If (c And ch_Red) Then c = c Xor ch_Red:            NextChan = nm_Red:       Exit Function
    If (c And ch_Blue) Then c = c Xor ch_Blue:          NextChan = nm_Blue:      Exit Function
    If (c And ch_Green) Then c = c Xor ch_Green:        NextChan = nm_Green:     Exit Function
    If (c And ch_Alpha) Then c = c Xor ch_Alpha:        NextChan = nm_Alpha:     Exit Function
End Function

Private Sub DoLevels(ByVal channel As Long, GT_Tmp() As Byte)
    Dim x As Long, y As Long, ch As Long
    
    Do
        ch = NextChan(channel, 1):       If ch = 0 Then Exit Do

        If ch = nm_RGB Then
            For y = BY To BY + BH - 1
                For x = BX To BX + BW - 1
                    B3D(3, x, y) = GT_Tmp(B3D(3, x, y))
                    B3D(2, x, y) = GT_Tmp(B3D(2, x, y))
                    B3D(1, x, y) = GT_Tmp(B3D(1, x, y))
                Next
            Next
        Else
            For y = BY To BY + BH - 1
                For x = BX To BX + BW - 1
                    B3D(ch, x, y) = GT_Tmp(B3D(ch, x, y))
                Next
            Next
        End If
    Loop
End Sub

Private Sub DoRect(ByVal srcW As Long, ByVal srcH As Long, x As Long, y As Long, w As Long, h As Long, typeX As Single, typeY As Single, typeW As Single, typeH As Single, offsetX As Single, offsetY As Single)
    Select Case typeW
        Case Is > 1:    w = typeW
        Case Is > 0:    w = srcW * typeW
        Case Is < 0:    w = srcW + typeW
    End Select
    
    Select Case typeH
        Case Is > 1:    h = typeH
        Case Is > 0:    h = srcH * typeH
        Case Is < 0:    h = srcH + typeH
    End Select

    Select Case typeX
        Case -1:        x = srcW / 2 - w / 2 + offsetX
        Case Is > 0:    x = srcW * typeX + offsetX
        Case -1.1:      x = srcW / 2 + offsetX
        Case -1.2:      x = srcW / 2 - w + offsetX
        Case -2:        x = offsetX
        Case -3:        x = srcW - w + offsetX
    End Select

    Select Case typeY
        Case -1:        y = srcH / 2 - h / 2 + offsetY
        Case Is > 0:    y = srcH * typeY + offsetY
        Case -1.1:      y = srcH / 2 + offsetY
        Case -1.2:      y = srcH / 2 - h + offsetY
        Case -2:        y = offsetY
        Case -3:        y = srcH - h + offsetY
    End Select
    
    CorrRect srcW, srcH, x, y, w, h
End Sub

Private Sub CorrRect(ByVal srcW As Long, ByVal srcH As Long, x As Long, y As Long, w As Long, h As Long, Optional ByVal isAutoWH As Boolean)
    If isAutoWH Then
        If w = 0 Then w = srcW
        If h = 0 Then h = srcH
    End If
    If x < 1 Then x = 1
    If y < 1 Then y = 1
    If w < 1 Then w = 1
    If h < 1 Then h = 1
    If x > srcW Then x = srcW
    If y > srcH Then y = srcH
    If (x + w - 1) > srcW Then w = srcW - x + 1
    If (y + h - 1) > srcH Then h = srcH - y + 1
End Sub
