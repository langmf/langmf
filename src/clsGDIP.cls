VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGDIP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' GDI+ Functions
Private Declare Sub GdiplusShutdown Lib "GdiPlus" (ByVal Token As Long)
Private Declare Function GdiplusStartup Lib "GdiPlus" (Token As Long, inputbuf As Any, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Function GdipDisposeImage Lib "GdiPlus" (ByVal Image As Long) As Long
Private Declare Function GdipDrawImageRect Lib "GdiPlus" (ByVal graphics As Long, ByVal Image As Long, ByVal x As Single, ByVal y As Single, ByVal Width As Single, ByVal Height As Single) As Long
Private Declare Function GdipCreateFromHDC Lib "GdiPlus" (ByVal hDC As Long, GpGraphics As Long) As Long
Private Declare Function GdipGetImageWidth Lib "GdiPlus" (ByVal Image As Long, Width As Long) As Long
Private Declare Function GdipGetImageHeight Lib "GdiPlus" (ByVal Image As Long, Height As Long) As Long
Private Declare Function GdipDrawImageRectI Lib "GdiPlus" (ByVal graphics As Long, ByVal Img As Long, ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "GdiPlus" (ByVal graphics As Long) As Long
Private Declare Function GdipBitmapLockBits Lib "GdiPlus" (ByVal hBitmap As Long, rRect As Long, ByVal Flags As Long, ByVal Format As Long, lockedBitmapData As BitmapData) As Long
Private Declare Function GdipSaveImageToFile Lib "GdiPlus" (ByVal Image As Long, ByVal FileName As Long, ByRef clsidEncoder As UUID, ByRef encoderParams As Any) As Long
Private Declare Function GdipSaveImageToFile2 Lib "GdiPlus" Alias "GdipSaveImageToFile" (ByVal Image As Long, ByVal FileName As Long, clsidEncoder As UUID, encoderParams As Any) As Long
Private Declare Function GdipBitmapUnlockBits Lib "GdiPlus" (ByVal hBitmap As Long, lockedBitmapData As BitmapData) As Long
Private Declare Function GdipGetImageEncoders Lib "GdiPlus" (ByVal numEncoders As Long, ByVal Size As Long, Encoders As Any) As Long
Private Declare Function GdipSaveImageToStream Lib "GdiPlus" (ByVal Image As Long, ByVal pStream As stdole.IUnknown, clsidEncoder As Any, encoderParams As Any) As Long
Private Declare Function GdipLoadImageFromFile Lib "GdiPlus" (ByVal FileName As Long, GpImage As Long) As Long
Private Declare Function GdipDrawImageRectRectI Lib "GdiPlus" (ByVal graphics As Long, ByVal GpImage As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal srcUnit As Long, ByVal imageAttributes As Long, ByVal callbackFunc As Long, ByVal callbackData As Long) As Long
Private Declare Function GdipLoadImageFromStream Lib "GdiPlus" (ByVal pStream As stdole.IUnknown, rImage As Long) As Long
Private Declare Function GdipSetInterpolationMode Lib "GdiPlus" (ByVal graphics As Long, ByVal InterMode As Long) As Long
Private Declare Function GdipGetImageEncodersSize Lib "GdiPlus" (numEncoders As Long, Size As Long) As Long
Private Declare Function GdipCreateBitmapFromHICON Lib "GdiPlus" (ByVal hIcon As Long, GpBitmap As Long) As Long
Private Declare Function GdipCreateMetafileFromEmf Lib "GdiPlus" (ByVal hEmf As Long, ByVal deleteEmf As Long, Metafile As Long) As Long
Private Declare Function GdipCreateBitmapFromHBITMAP Lib "GdiPlus" (ByVal hBMP As Long, ByVal hPal As Long, GpBitmap As Long) As Long

Private Type GdiplusStartupInput
    GdiplusVersion           As Long
    DebugEventCallback       As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs   As Long
End Type

Private Type BitmapData
  Width As Long
  Height As Long
  Stride As Long
  PixelFormat As Long
  Scan0Ptr As Long
  ReservedPtr As Long
End Type

Private Type EncoderParameter
    GUID As UUID
    NumberOfValues As Long
    Type As Long
    value As Long
End Type

Private Type EncoderParameters
    Count As Long
    Parameter As EncoderParameter
End Type

Private Type ImageCodecInfo
    ClassID(0 To 3)   As Long
    FormatID(0 To 3)  As Long
    CodecName         As Long
    DllName           As Long
    FormatDescription As Long
    FilenameExtension As Long
    MimeType          As Long
    Flags             As Long
    Version           As Long
    SigCount          As Long
    SigSize           As Long
    SigPattern        As Long
    SigMask           As Long
End Type

Const PF_24RGB As Long = 137224, PF_32RGB As Long = 139273, PF_32ARGB As Long = 2498570, PF_32PARGB As Long = 925707
Const PICTYPE_BITMAP As Long = 1, InterpolationMode As Long = 7, UnitPixel As Long = 2, DIB_RGB_COLORS As Long = 0
Const AC_SRC_OVER As Long = 0, ULW_OPAQUE As Long = 4, ULW_COLORKEY As Long = 1, ULW_ALPHA As Long = 2
Const PATCOPY As Long = &HF00021, AC_SRC_ALPHA As Long = 1
 
Public Token As Long


Private Sub Class_Initialize()
    Init
End Sub

Private Sub Class_Terminate()
    Free
End Sub

Public Function Init() As Long
    Dim GSI As GdiplusStartupInput
    On Error Resume Next
    GSI.GdiplusVersion = 1
    Init = GdiplusStartup(Token, GSI)
End Function

Public Sub Free()
    On Error Resume Next
    If Token Then GdiplusShutdown Token
    Token = 0
End Sub

' Loads the picture (optionally resized)
Public Function LoadPicture(datFile As Variant, Optional ByVal Width As Long = -1, Optional ByVal Height As Long = -1, Optional ByVal BackColor As Long = vbWhite, Optional ByVal RetainRatio As Boolean = False) As IPictureDisp
    Dim hDC As Long, hBitmap As Long, Img As Long
    
    On Error Resume Next
    
    ' Load the image
    If VarType(datFile) = vbString Then If IsFile(CStr(datFile)) Then Call GdipLoadImageFromFile(StrPtr(LongPath(CStr(datFile))), Img)  ' Load Png
    
    If Img = 0 Then Img = GDIP_LoadPictureFromByte(datFile)
    
    ' Calculate picture's width and height if not specified
    If Width = -1 Or Height = -1 Then GdipGetImageWidth Img, Width:    GdipGetImageHeight Img, Height
    
    ' Initialise the hDC
    GDIP_InitDC hDC, hBitmap, BackColor, Width, Height

    ' Resize the picture
    GDIP_Resize Img, hDC, Width, Height, RetainRatio
    GdipDisposeImage Img
    
    ' Get the bitmap back
    GDIP_GetBitmap hDC, hBitmap

    ' Create the picture
    Set LoadPicture = GDIP_CreatePicture(hBitmap)
End Function

Public Function SavePicture(ByVal pict As StdPicture, ByVal FileName As String, Optional ByVal PicType As String, Optional ByVal Quality As Byte = 85, Optional ByVal TIFF_ColorDepth As Long = 24, Optional ByVal TIFF_Compression As Long = 5) As Long
    Dim lBitmap As Long, aEncParams() As Byte, tEncoder As UUID, tParams As EncoderParameters
    
    On Error Resume Next
    
    If LenB(PicType) = 0 Then PicType = "." & GetExtension(FileName)
    
    If LCase$(PicType) = ".bmp" Then VB.SavePicture pict, FileName:  Exit Function
    
    If GdipCreateBitmapFromHBITMAP(pict.Handle, 0, lBitmap) = 0 Then
        Select Case LCase$(PicType)
            Case ".jpg"
                CLSIDFromString StrPtr("{557CF401-1A04-11D3-9A73-0000F81EF32E}"), tEncoder
                tParams.Count = 1
                With tParams.Parameter
                    CLSIDFromString StrPtr("{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"), .GUID
                    .NumberOfValues = 1
                    .Type = 4
                    .value = VarPtr(Quality)
                End With
                ReDim aEncParams(1 To Len(tParams))
                Call CopyMemory(aEncParams(1), tParams, Len(tParams))
                
            Case ".png"
                CLSIDFromString StrPtr("{557CF406-1A04-11D3-9A73-0000F81EF32E}"), tEncoder
                ReDim aEncParams(1 To Len(tParams))
                
            Case ".gif"
                CLSIDFromString StrPtr("{557CF402-1A04-11D3-9A73-0000F81EF32E}"), tEncoder
                ReDim aEncParams(1 To Len(tParams))
                
            Case ".tiff"
                CLSIDFromString StrPtr("{557CF405-1A04-11D3-9A73-0000F81EF32E}"), tEncoder
                tParams.Count = 2
                ReDim aEncParams(1 To Len(tParams) + Len(tParams.Parameter))
                With tParams.Parameter
                    .NumberOfValues = 1
                    .Type = 4
                    CLSIDFromString StrPtr("{E09D739D-CCD4-44EE-8EBA-3FBF8BE4FC58}"), .GUID
                    .value = VarPtr(TIFF_Compression)
                End With
                Call CopyMemory(aEncParams(1), tParams, Len(tParams))
                With tParams.Parameter
                    .NumberOfValues = 1
                    .Type = 4
                    CLSIDFromString StrPtr("{66087055-AD66-4C7C-9A18-38A2310B8337}"), .GUID
                    .value = VarPtr(TIFF_ColorDepth)
                End With
                Call CopyMemory(aEncParams(Len(tParams) + 1), tParams.Parameter, Len(tParams.Parameter))
        End Select
        
        If tParams.Count = 0 Then
            SavePicture = GdipSaveImageToFile(lBitmap, StrPtr(LongPath(FileName)), tEncoder, ByVal 0&)
        Else
            SavePicture = GdipSaveImageToFile2(lBitmap, StrPtr(LongPath(FileName)), tEncoder, aEncParams(1))
        End If
        
        GdipDisposeImage lBitmap
    End If
    
    Erase aEncParams
End Function

Public Function SaveBuffer(ByVal pict As StdPicture, Optional ByVal MimeType As String = "image/png") As Byte()
    Dim lBitmap As Long, uEnc(0 To 3) As Long, istm As stdole.IUnknown

    If GDIP_GetEncoder(MimeType, uEnc) = -1 Then Exit Function
    Set istm = IStreamFromArray(0, 0):      If istm Is Nothing Then Exit Function

    If GdipCreateBitmapFromHBITMAP(pict.Handle, 0, lBitmap) = 0 Then
        If GdipSaveImageToStream(lBitmap, istm, uEnc(0), ByVal 0&) = 0 Then IStreamToArray istm, SaveBuffer
    End If

    Set istm = Nothing
End Function

Public Function Resize(ByVal Handle As Long, ByVal PicType As PictureTypeConstants, ByVal Width As Long, ByVal Height As Long, Optional ByVal BackColor As Long = vbWhite, Optional ByVal RetainRatio As Boolean = False) As IPictureDisp
    Dim Img As Long, hDC As Long, hBitmap As Long
    
    ' Determine pictyre type
    Select Case PicType
        Case vbPicTypeBitmap:       GdipCreateBitmapFromHBITMAP Handle, ByVal 0&, Img
        Case vbPicTypeEMetafile:    GdipCreateMetafileFromEmf Handle, False, Img
        Case vbPicTypeIcon:         GdipCreateBitmapFromHICON Handle, Img
    End Select
    
    ' Continue with resizing only if we have a valid image object
    If Img Then
        GDIP_InitDC hDC, hBitmap, BackColor, Width, Height
        GDIP_Resize Img, hDC, Width, Height, RetainRatio
        GdipDisposeImage Img
        GDIP_GetBitmap hDC, hBitmap
        Set Resize = GDIP_CreatePicture(hBitmap)
    End If
End Function


Public Function LoadPngSkin(ByVal Frm As Object, datFile As Variant, Optional ByVal mAlpha As Integer = 255) As Long
    Dim curWinLong As Long, graphics As Long, Img As Long, lngHeight As Long, lngWidth As Long
    Dim tempBI As BITMAPINFO
    
    If VarType(datFile) = vbString Then If IsFile(CStr(datFile)) Then Call GdipLoadImageFromFile(StrPtr(LongPath(CStr(datFile))), Img)  ' Load Png
    
    If Img = 0 Then Img = GDIP_LoadPictureFromByte(datFile)
    
    LoadPngSkin = Img
    
    Call GdipGetImageHeight(Img, lngHeight)
    Call GdipGetImageWidth(Img, lngWidth)
    
    With tempBI.bmiHeader
        .biSize = Len(tempBI.bmiHeader)
        .biBitCount = 32    ' Each pixel is 32 bit's wide
        .biHeight = Frm.ScaleHeight  ' Height of the form
        .biWidth = Frm.ScaleWidth    ' Width of the form
        .biPlanes = 1   ' Always set to 1
        .biSizeImage = .biWidth * .biHeight * (.biBitCount / 8) ' This is the number of bytes that the bitmap takes up. It is equal to the Width*Height*ByteCount (bitCount/8)
    End With
    
    Frm.gdip_mDC = CreateCompatibleDC(Frm.hDC)
    Frm.gdip_MainBitmap = CreateDIBSection(Frm.gdip_mDC, tempBI, DIB_RGB_COLORS, 0, 0, 0)
    Frm.gdip_OldBitmap = SelectObject(Frm.gdip_mDC, Frm.gdip_MainBitmap)   ' Select the new bitmap, track the old that was selected
        
    ' GDI Initializations
    Call GdipCreateFromHDC(Frm.gdip_mDC, graphics)

    Call GdipDrawImageRect(graphics, Img, 0, 0, lngWidth, lngHeight)
    
    ' Change windows extended style to be used by updatelayeredwindow
    curWinLong = GetWindowLongW(Frm.hWnd, GWL_EXSTYLE)
    ' Accidently did This line below which flipped entire form, it's neat so I left it in
    ' Comment out the line above and uncomment line below.
    'curWinLong = GetWindowLong(Frm.hWnd, GWL_STYLE)
    SetWindowLongW Frm.hWnd, GWL_EXSTYLE, curWinLong Or WS_EX_LAYERED
       
    ' Make the window a top-most window so we can always see the cool stuff
    SetWindowPos Frm.hWnd, hWnd_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
       
    ' Needed for updateLayeredWindow call
    UpdatePngSkin Frm, mAlpha
    
    Call GdipDisposeImage(Img)
    Call GdipDeleteGraphics(graphics)
End Function

Public Function UpdatePngSkin(ByVal Frm As Object, Optional ByVal mAlpha As Integer = 255) As Long
    Dim blendFunc32bpp As BLENDFUNCTION, srcPoint As POINTAPI, winSize As POINTAPI

    ' Needed for updateLayeredWindow call
    srcPoint.x = 0
    srcPoint.y = 0
    winSize.x = Frm.ScaleWidth
    winSize.y = Frm.ScaleHeight
        
    With blendFunc32bpp
        .AlphaFormat = AC_SRC_ALPHA ' 32 bit
        .BlendFlags = 0
        .BlendOp = AC_SRC_OVER
        .SourceConstantAlpha = mAlpha
    End With
    
    UpdatePngSkin = UpdateLayeredWindow(Frm.hWnd, Frm.hDC, ByVal 0&, winSize, Frm.gdip_mDC, srcPoint, 0, blendFunc32bpp, ULW_ALPHA)
End Function

Public Sub UnloadPngSkin(ByVal Frm As Object)
    SelectObject Frm.gdip_mDC, Frm.gdip_OldBitmap
    DeleteObject Frm.gdip_MainBitmap
    DeleteObject Frm.gdip_OldBitmap
    DeleteDC Frm.gdip_mDC
End Sub


Public Function Pic2DIB(Buf As Variant, Optional ByVal BackColor As Variant) As Variant
    Dim x As Long, y As Long, a1 As Byte, a2 As Byte, c(3) As Byte, Img As Long, BD As BitmapData, tmpBuf() As Byte

    Img = GDIP_LoadPictureFromByte(Buf)

    If Img Then
        GdipBitmapLockBits Img, ByVal 0&, 1, PF_32ARGB, BD
        ReDim tmpBuf(1 To 4, 1 To BD.Width, 1 To BD.Height)
        CopyMemory tmpBuf(1, 1, 1), ByVal BD.Scan0Ptr, BD.Width * BD.Height * 4
        GdipBitmapUnlockBits Img, BD
        GdipDisposeImage Img

        If Not IsMissing(BackColor) Then
            x = BackColor:    CopyMem4 x, c(0):    c(3) = 255 - c(3)
            For y = 1 To BD.Height
                For x = 1 To BD.Width
                    a1 = tmpBuf(4, x, y):   a2 = 255 - a1
                    If a1 <= c(3) Then
                        tmpBuf(3, x, y) = GT_Grad(tmpBuf(3, x, y), a1) + GT_Grad(c(0), a2)
                        tmpBuf(2, x, y) = GT_Grad(tmpBuf(2, x, y), a1) + GT_Grad(c(1), a2)
                        tmpBuf(1, x, y) = GT_Grad(tmpBuf(1, x, y), a1) + GT_Grad(c(2), a2)
                    End If
                Next
            Next
        End If

        Pic2DIB = tmpBuf
    End If
End Function


Private Function GDIP_LoadPictureFromByte(value As Variant) As Long
    Dim sz As Long, istm As stdole.IUnknown, tmpBuf() As Byte

    ConvToBufferByte value, tmpBuf:      sz = ArraySize(tmpBuf):        If sz = 0 Then Exit Function

    Set istm = IStreamFromArray(VarPtr(tmpBuf(0)), sz)
    Call GdipLoadImageFromStream(istm, GDIP_LoadPictureFromByte)
    Set istm = Nothing
End Function

Private Function GDIP_GetEncoder(strMimeType As String, ClassID() As Long) As Long
    Dim cnt As Long, sz As Long, idx As Long, sMime As String, Buf() As Byte, ICI() As ImageCodecInfo

    ' Get the encoder array size
    Call GdipGetImageEncodersSize(cnt, sz)
    If sz = 0 Or cnt = 0 Then Exit Function

    ' Allocate room for the arrays dynamically
    ReDim ICI(1 To cnt) As ImageCodecInfo
    ReDim Buf(1 To sz) As Byte

    Call GdipGetImageEncoders(cnt, sz, Buf(1))
    Call CopyMemory(ICI(1), Buf(1), (Len(ICI(1)) * cnt))

    sz = Len(strMimeType)
    sMime = String$(sz, vbNullChar)

    For idx = cnt To 1 Step -1
        With ICI(idx)
            If sz = lstrlenW(ByVal .MimeType) Then
                Call CopyMemory(ByVal StrPtr(sMime), ByVal .MimeType, sz * 2)
                If StrComp(sMime, strMimeType, vbTextCompare) = 0 Then
                    CopyMemory ClassID(0), .ClassID(0), 16
                    Exit For
                End If
            End If
        End With
    Next

    GDIP_GetEncoder = idx
End Function

Private Sub GDIP_Resize(Img As Long, hDC As Long, Width As Long, Height As Long, Optional RetainRatio As Boolean = False)
    Dim graphics As Long, OrWidth As Long, OrHeight As Long, OrRatio As Double, DesRatio As Double
    Dim DestX As Long, DestY As Long, DestWidth As Long, DestHeight As Long

    GdipCreateFromHDC hDC, graphics
    GdipSetInterpolationMode graphics, InterpolationMode

    If RetainRatio Then
        GdipGetImageWidth Img, OrWidth
        GdipGetImageHeight Img, OrHeight

        OrRatio = OrWidth / OrHeight
        DesRatio = Width / Height

        ' Calculate destination coordinates
        DestWidth = IIF(DesRatio < OrRatio, Width, Height * OrRatio)
        DestHeight = IIF(DesRatio < OrRatio, Width / OrRatio, Height)
        DestX = (Width - DestWidth) / 2
        DestY = (Height - DestHeight) / 2

        GdipDrawImageRectRectI graphics, Img, DestX, DestY, DestWidth, DestHeight, 0, 0, OrWidth, OrHeight, UnitPixel, 0, 0, 0
    Else
        GdipDrawImageRectI graphics, Img, 0, 0, Width, Height
    End If

    GdipDeleteGraphics graphics
End Sub

Private Sub GDIP_InitDC(hDC As Long, hBitmap As Long, BackColor As Long, Width As Long, Height As Long)
    Dim hBrush As Long

    hDC = CreateCompatibleDC(0)
    hBitmap = CreateBitmap(Width, Height, GetDeviceCaps(hDC, DC_PLANES), GetDeviceCaps(hDC, DC_BITSPIXEL), ByVal 0&)
    hBitmap = SelectObject(hDC, hBitmap)

    If BackColor >= 0 Then
        hBrush = CreateSolidBrush(BackColor)
        hBrush = SelectObject(hDC, hBrush)
        PatBlt hDC, 0, 0, Width, Height, PATCOPY
        DeleteObject SelectObject(hDC, hBrush)
    End If
End Sub

Private Function GDIP_CreatePicture(hBitmap As Long) As IPictureDisp
    Dim tPic As PictDesc

    tPic.Size = Len(tPic)
    tPic.Type = vbPicTypeBitmap
    tPic.hHandle = hBitmap

    Call OleCreatePictureIndirect(tPic, IID_IDispatch, True, GDIP_CreatePicture)
End Function

Private Sub GDIP_GetBitmap(hDC As Long, hBitmap As Long)
    hBitmap = SelectObject(hDC, hBitmap)
    DeleteDC hDC
End Sub
