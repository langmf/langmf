VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function AlphaBlend Lib "msimg32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal widthSrc As Long, ByVal heightSrc As Long, blendFunct As BLENDFUNCTION) As Long
Private Declare Function GradientFill Lib "msimg32" (ByVal hDC As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, pMesh As Gradient_Triangle, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long
Private Declare Function func_Polygon Lib "gdi32" Alias "Polygon" (ByVal hDC As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, ByRef lpDevMode As DEVMODE) As Long
Private Declare Function ChangeDisplaySettingsEx Lib "user32" Alias "ChangeDisplaySettingsExA" (lpszDeviceName As Any, lpDevMode As Any, ByVal hWnd As Long, ByVal dwFlags As Long, lParam As Any) As Long

Private Const DM_BITSPERPEL = &H40000, DM_PELSHEIGHT = &H100000, DM_PELSWIDTH = &H80000, DM_DISPLAYFREQUENCY = &H400000

Public Clip As Object
Private m_Plus As clsGDIP

Public Property Get Plus() As Object                            'no_check
    If m_Plus Is Nothing Then Set m_Plus = New clsGDIP
    Set Plus = m_Plus
End Property


Public Sub Scale2(ByVal Obj As Object, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single)
    Obj.Scale (x1, y1)-(x2, y2)
End Sub

Public Sub Polygon(ByVal m_hDC As Long, m_Vert As Variant)
    Dim a As Long, Vert() As POINTAPI

    If Not VerifyArrayRange(m_Vert) Then Exit Sub
    
    ReDim Vert(UBound(m_Vert))
    
    For a = 0 To UBound(m_Vert)
        Vert(a).x = CLng(m_Vert(a)(0))
        Vert(a).y = CLng(m_Vert(a)(1))
    Next
    
    func_Polygon m_hDC, Vert(0), UBound(m_Vert) + 1
End Sub

Public Sub Gradient(ByVal Frm As Object, m_Vert As Variant, Optional m_gTri As Variant, Optional ByVal m_Type As Long = 1)
    Dim a As Long, Vert() As TRIVERTEX, gTri() As Gradient_Triangle
    
    If Not VerifyArrayRange(m_Vert) Then Exit Sub
    
    ReDim Vert(UBound(m_Vert))
    
    For a = 0 To UBound(m_Vert)
        m_ParamArray m_Vert(a), 0, 0, 0, 0, 0, 0
        
        Vert(a).x = CLng(m_Vert(a)(0))
        Vert(a).y = CLng(m_Vert(a)(1))
        Vert(a).red = LongToUShort(m_Vert(a)(2) * 256)
        Vert(a).green = LongToUShort(m_Vert(a)(3) * 256)
        Vert(a).blue = LongToUShort(m_Vert(a)(4) * 256)
        Vert(a).alpha = LongToUShort(m_Vert(a)(5) * 256)
    Next
    
    If IsMissing(m_gTri) Then m_ParamArray m_gTri, Array(0, 1, 0)
    
    ReDim gTri(UBound(m_gTri))
    
    For a = 0 To UBound(m_gTri)
        gTri(a).Vertex1 = CLng(m_gTri(a)(0))
        gTri(a).Vertex2 = CLng(m_gTri(a)(1))
        gTri(a).Vertex3 = CLng(m_gTri(a)(2))
    Next
    
    Frm.AutoRedraw = True
    GradientFill Frm.hDC, Vert(0), UBound(m_Vert) + 1, gTri(0), UBound(m_gTri) + 1, m_Type
    Frm.Refresh
End Sub

Public Sub AlphaPic(ByVal dstHdc As Long, ByVal dstLeft As Long, ByVal dstTop As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal srcHdc As Long, ByVal srcLeft As Long, ByVal srcTop As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal pR As Integer, Optional ByVal Flags As Long = 0)
    Dim bf As BLENDFUNCTION

    bf.SourceConstantAlpha = pR
    bf.AlphaFormat = Flags
    Call AlphaBlend(dstHdc, dstLeft, dstTop, dstWidth, dstHeight, srcHdc, srcLeft, srcTop, srcWidth, srcHeight, bf)
End Sub

Public Function IsAlpha(ByVal hWnd As Long) As Boolean
    Dim Msg As Long

    Msg = GetWindowLongW(hWnd, GWL_EXSTYLE)
    If (Msg And WS_EX_LAYERED) = WS_EX_LAYERED Then IsAlpha = True Else IsAlpha = False
End Function

Public Function SetAlpha(ByVal hWnd As Long, ByVal Perc As Integer, Optional ByVal Flags As Integer = 2) As Long
    Dim Msg As Long

    If Perc >= 0 And Perc <= 255 Then
      Msg = GetWindowLongW(hWnd, GWL_EXSTYLE)
      Msg = Msg Or WS_EX_LAYERED
      Call SetWindowLongW(hWnd, GWL_EXSTYLE, Msg)
      SetAlpha = SetLayeredWindowAttributes(hWnd, 0, Perc, Flags)
    End If
End Function

Public Function GetGrad(ByVal Level As Long, ByVal color1 As Long, ByVal color2 As Long) As Long
    Dim iLevel As Long, r As Long, g As Long, b As Long, c1(3) As Byte, c2(3) As Byte

    If Level > 255 Then Level = 255
    If Level < 0 Then Level = 0
    
    If Level = 0 Then GetGrad = color1:    Exit Function
    If Level = 255 Then GetGrad = color2:  Exit Function
    
    iLevel = 255 - Level
    
    CopyMem4 color1, c1(0)
    CopyMem4 color2, c2(0)

    r = GT_Grad(c1(0), iLevel) + GT_Grad(c2(0), Level)
    g = GT_Grad(c1(1), iLevel) + GT_Grad(c2(1), Level)
    b = GT_Grad(c1(2), iLevel) + GT_Grad(c2(2), Level)

    GetGrad = b * 65536 + g * 256 + r
End Function

Public Function LoadPicture(Optional ByVal FileName As String, Optional Size As Variant, Optional ColorDepth As Variant, Optional x As Variant, Optional y As Variant) As IPictureDisp
    Set LoadPicture = VB.LoadPicture(FileName, Size, ColorDepth, x, y)
End Function

Public Function SetDisplayMode(ByVal Width As Long, ByVal Height As Long, ByVal BPP As Long, ByVal Freq As Long, ByVal change_flags As Long) As Long
    Dim dev_mode As DEVMODE

    ' Set the desired mode.
    With dev_mode
        .dmSize = Len(dev_mode)
        .dmDriverExtra = 0
        .dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT Or DM_BITSPERPEL Or DM_DISPLAYFREQUENCY
        .dmPelsWidth = Width
        .dmPelsHeight = Height
        .dmBitsPerPel = BPP
        .dmDisplayFrequency = Freq
    End With

    ' Make the change.
    SetDisplayMode = ChangeDisplaySettingsEx(ByVal 0&, dev_mode, 0, change_flags, ByVal 0&)
End Function

Public Function EnumDisplay() As Collection
    Dim tmpCollect As New Collection, tmpCl As Collection, dev_mode As DEVMODE, mode_num As Long

    ' List the display modes.
    dev_mode.dmSize = Len(dev_mode)
    dev_mode.dmDriverExtra = 0
    
    Do While EnumDisplaySettings(0, mode_num, dev_mode) <> 0
        
        Set tmpCl = New Collection
        
        tmpCl.Add dev_mode.dmPelsWidth, "Width"
        tmpCl.Add dev_mode.dmPelsHeight, "Height"
        tmpCl.Add dev_mode.dmBitsPerPel, "Bpp"
        tmpCl.Add dev_mode.dmDisplayFrequency, "Freq"
        
        tmpCollect.Add tmpCl

        Set tmpCl = Nothing
            
        mode_num = mode_num + 1
    Loop
    
    Set EnumDisplay = tmpCollect
End Function

Public Function GetDisplayMode() As Collection
    Dim tmpCl As New Collection
    
    tmpCl.Add GetDeviceCaps(frmScript.hDC, 8), "Width"
    tmpCl.Add GetDeviceCaps(frmScript.hDC, 10), "Height"
    tmpCl.Add GetDeviceCaps(frmScript.hDC, 12), "Bpp"
    tmpCl.Add GetDeviceCaps(frmScript.hDC, 116), "Freq"
    
    Set GetDisplayMode = tmpCl
End Function

Public Function PicToIco(ByVal hSrcDC As Long, ByVal nWidth As Long, ByVal nHeight As Long, Optional ByVal lMaskColor As Long = -1) As Long
    Dim hDCMask As Long, hDCColor As Long, hScrDC As Long, hDstDC As Long, Ico As ICONINFO
    Dim hBmMask As Long, hBmColor As Long, hBmColorOld As Long, hBmMaskOld As Long

    hScrDC = GetDC(0&)
    hDstDC = CreateCompatibleDC(hSrcDC)

    hBmColor = CreateCompatibleBitmap(hSrcDC, nWidth, nHeight)                      ' Создаем цветную картинку
    hBmMask = CreateBitmap(nWidth, nHeight, 1&, 1&, ByVal 0&)                        ' Создаем черно-белую картинку

    hDCColor = CreateCompatibleDC(hSrcDC)                                           ' Создаем DC
    hBmColorOld = SelectObject(hDCColor, hBmColor)                                  ' Объединяем...

    Call SetBkColor(hDCColor, GetBkColor(hSrcDC))                                   ' Ставим нашему DC цвета исходного
    Call SetTextColor(hDCColor, GetTextColor(hSrcDC))
    Call BitBlt(hDCColor, 0&, 0&, nWidth, nHeight, hSrcDC, 0&, 0&, vbSrcCopy)       ' Копируем исходную картинку

    hDCMask = CreateCompatibleDC(hSrcDC)                                            ' Создаем маску и
    hBmMaskOld = SelectObject(hDCMask, hBmMask)                                     ' Крепим ее к DC

    If lMaskColor = -1 Then lMaskColor = GetPixel(hSrcDC, 0&, 0&)                   ' Цвет маски ...

    Call SetBkColor(hDCColor, lMaskColor)                                           ' Ставим прозрачный цвет основным
    Call SetTextColor(hDCColor, vbWhite)
    Call BitBlt(hDCMask, 0, 0, nWidth, nHeight, hDCColor, 0, 0, vbSrcCopy)          ' Создаем маску

    Call SetTextColor(hDCColor, vbBlack)                                            ' Ставим цвет и создаем белую маску (XOR)
    Call SetBkColor(hDCColor, vbWhite)
    Call BitBlt(hDCColor, 0, 0, nWidth, nHeight, hDCMask, 0, 0, &H220326)           ' Ставим бел. маску на ненужные места

    Call BitBlt(hDstDC, 0, 0, nWidth, nHeight, hDCMask, 0, 0, vbSrcAnd)             ' Накладываем маску
    Call BitBlt(hDstDC, 0, 0, nWidth, nHeight, hDCColor, 0, 0, vbSrcPaint)          ' Объединяем содержимое картинки и hdcColor

    Ico.fIcon = True
    Ico.hBmColor = SelectObject(hDCColor, hBmColorOld)
    Ico.hBmMask = SelectObject(hDCMask, hBmMaskOld)
    PicToIco = CreateIconIndirect(Ico)

    Call DeleteObject(Ico.hBmColor)                                                 ' Удаляем созданную hBmColor
    Call DeleteDC(hDCColor)                                                         ' Удаляем hDCColor
    Call DeleteObject(Ico.hBmMask)                                                  ' Удаляем созданную hBmMask
    Call DeleteDC(hDCMask)                                                          ' Удаляем hDCMask
    Call DeleteDC(hDstDC)                                                           ' Удаляем hDCMask
    Call ReleaseDC(0&, hScrDC)
End Function
    
Public Function IcoToPic(ByVal hImage As Long, Optional ByVal imgType As Long = 3) As IPicture
    Dim lpPictDesc As PictDesc, riid As UUID
    
    With lpPictDesc
        .Size = Len(lpPictDesc)
        .Type = imgType
        .hHandle = hImage
    End With

    With riid
        .Data1 = &H7BF80980:    .Data2 = &HBF32:    .Data3 = &H101A
        .Data4(0) = &H8B:   .Data4(1) = &HBB:   .Data4(2) = &H0:    .Data4(3) = &HAA
        .Data4(4) = &H0:    .Data4(5) = &H30:   .Data4(6) = &HC:    .Data4(7) = &HAB
    End With

    Call OleCreatePictureIndirect(lpPictDesc, riid, True, IcoToPic)
End Function

Public Function RegionFromBitmap(ByVal picSrc As IPictureDisp, Optional ByVal TransColor As Variant) As Long
    RegionFromBitmap = m_RegionFromBitmap(picSrc, TransColor)
End Function

Public Function Trans(ByVal Obj As Object, ByVal ipdFon As IPictureDisp, ByVal ipdMask As IPictureDisp) As Long
    Dim rg As Long
    
    With Obj
        .Picture = ipdFon
        .Move .Left, .Top, m_HPX(ipdFon.Width) * Screen.TwipsPerPixelX, m_HPY(ipdFon.Height) * Screen.TwipsPerPixelY
        rg = m_RegionFromBitmap(ipdMask)
        Trans = SetWindowRgn(.hWnd, rg, True)
        DeleteObject rg
    End With
End Function

Public Function Pic2DIB(Buf As Variant) As Variant
    Dim FAP As New clsAsmPic
    Set FAP.Obj = LoadPictureFromByte(Buf)
    FAP.GetDI
    Pic2DIB = FAP.GetBuffer
End Function

Public Sub SaveBMP(src As Variant, ByVal saveFile As String)
    If VarType(src) = vbString Then SavePicture VB.LoadPicture(src), saveFile Else SavePicture src, saveFile
End Sub

Private Function LongToUShort(ByVal ULong As Long) As Integer
    GetMem2 VarPtr(ULong), LongToUShort
End Function

Private Sub Class_Initialize()
    Set Clip = New clsClip
End Sub
