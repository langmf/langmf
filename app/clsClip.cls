VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsClip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_Picture As IPictureDisp, m_StretchX As Long, m_StretchY As Long, m_BackColor As Long
Private m_Cols As Long, m_Rows As Long, m_ClipX As Long, m_ClipY As Long, m_ClipWidth As Long, m_ClipHeight As Long


Public Function CreateNew() As Object
    Set CreateNew = New clsClip
End Function

Public Property Get Picture() As IPictureDisp
    Set Picture = m_Picture
End Property

Public Property Let Picture(ByVal Pic As IPictureDisp)
    Set Picture = Pic
End Property

Public Property Set Picture(ByVal Pic As IPictureDisp)
    If Pic Is Nothing Then Exit Property
    Set m_Picture = Pic
End Property

Public Property Get BackColor() As Long
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal value As Long)
    m_BackColor = value
End Property

Public Property Get ClipX() As Long
    ClipX = m_ClipX
End Property

Public Property Let ClipX(ByVal value As Long)
    If value < 0 Then value = 0
    If value >= Width Then value = Width - 1
    m_ClipX = value
End Property

Public Property Get ClipY() As Long
    ClipY = m_ClipY
End Property

Public Property Let ClipY(ByVal value As Long)
    If value < 0 Then value = 0
    If value >= Height Then value = Height - 1
    m_ClipY = value
End Property

Public Property Get ClipWidth() As Long
    ClipWidth = m_ClipWidth
End Property

Public Property Let ClipWidth(ByVal value As Long)
    If value < 1 Then value = 1
    If value > Width Then value = Width
    m_ClipWidth = value
End Property

Public Property Get ClipHeight() As Long
    ClipHeight = m_ClipHeight
End Property

Public Property Let ClipHeight(ByVal value As Long)
    If value < 1 Then value = 1
    If value > Height Then value = Height
    m_ClipHeight = value
End Property

Public Property Get StretchX() As Long
    StretchX = m_StretchX
End Property

Public Property Let StretchX(ByVal value As Long)
    If value < 0 Then value = 0
    m_StretchX = value
End Property

Public Property Get StretchY() As Long
    StretchY = m_StretchY
End Property

Public Property Let StretchY(ByVal value As Long)
    If value < 0 Then value = 0
    m_StretchY = value
End Property

Public Property Get Cols() As Long
    Cols = m_Cols
End Property

Public Property Let Cols(ByVal value As Long)
    If value < 1 Then value = 1
    m_Cols = value
End Property

Public Property Get Rows() As Long
    Rows = m_Rows
End Property

Public Property Let Rows(ByVal value As Long)
    If value < 1 Then value = 1
    m_Rows = value
End Property

Public Property Get CellWidth() As Long
    CellWidth = Fix(Width / Cols)
End Property

Public Property Get CellHeight() As Long
    CellHeight = Fix(Height / Rows)
End Property

Public Property Get Width() As Long
    If IsStdPic Then Width = m_HPX(m_Picture.Width)
End Property

Public Property Get Height() As Long
    If IsStdPic Then Height = m_HPY(m_Picture.Height)
End Property

Public Property Get Clip(Optional ByVal x As Long = -1, Optional ByVal y As Long = -1, Optional ByVal cw As Long = -1, Optional ByVal ch As Long = -1) As IPictureDisp
    If x = -1 Then x = m_ClipX
    If y = -1 Then y = m_ClipY
    If cw = -1 Then cw = m_ClipWidth
    If ch = -1 Then ch = m_ClipHeight
    
    Set Clip = GetPicture(x, y, cw, ch)
End Property

Public Property Get GraphicCell(ByVal Index As Long) As IPictureDisp
    Dim x As Long, y As Long, cw As Long, ch As Long
    
    If Index < 0 Or (Index >= m_Cols * m_Rows) Then Exit Property
    
    cw = CellWidth:    ch = CellHeight
    If cw = 0 Or ch = 0 Then Exit Property
    
    y = (Index \ m_Cols) * ch
    x = (Index Mod m_Cols) * cw
    
    Set GraphicCell = GetPicture(x, y, cw, ch)
End Property

Public Function Render(ByVal dstDC As Long, Optional ByVal dstX As Long, Optional ByVal dstY As Long, Optional ByVal dstWidth As Long, Optional ByVal dstHeight As Long, Optional ByVal srcX As Long, Optional ByVal srcY As Long, Optional ByVal srcWidth As Long, Optional ByVal srcHeight As Long, Optional ByVal Zoom As Single = 1) As Boolean
    On Error Resume Next
    
    If Not IsStdPic Or Zoom <= 0! Then Exit Function

    If dstWidth Then dstWidth = dstWidth * Zoom Else dstWidth = Width * Zoom
    If dstHeight Then dstHeight = dstHeight * Zoom Else dstHeight = Height * Zoom
    If srcX Then srcX = m_PHX(srcX)
    If srcY Then srcY = m_PHY(srcY)
    If srcWidth Then srcWidth = m_PHX(srcWidth) Else srcWidth = m_Picture.Width
    If srcHeight Then srcHeight = m_PHY(srcHeight) Else srcHeight = m_Picture.Height

    m_Picture.Render dstDC Or 0&, dstX Or 0&, dstY Or 0&, dstWidth Or 0&, dstHeight Or 0&, _
                     srcX Or 0&, m_Picture.Height - srcY, srcWidth Or 0&, -srcHeight, 0&

    If Err Then Err.Clear Else Render = True
End Function

Private Function IsStdPic() As Boolean
    If m_Picture Is Nothing Then Exit Function
    If m_Picture.Handle = 0 Then Exit Function
    IsStdPic = True
End Function

Private Function CreatePicture(ByVal w As Long, ByVal h As Long) As IPictureDisp
    Dim hBMP As Long, hDC As Long, tPic As PictDesc, Pic As IPictureDisp

    hDC = GetDC(0)
    hBMP = CreateCompatibleBitmap(hDC, w, h)
    With tPic:   .Size = Len(tPic):   .Type = vbPicTypeBitmap:   .hHandle = hBMP:   End With
    Call OleCreatePictureIndirect(tPic, IID_IDispatch, 1, Pic)
    Set CreatePicture = Pic
    ReleaseDC 0, hDC
End Function

Private Function GetPicture(ByVal x As Long, ByVal y As Long, ByVal cw As Long, ByVal ch As Long) As IPictureDisp
    Dim cw2 As Long, ch2 As Long, dstDC As Long, dstOld As Long, rc As RECT, hBrush As Long, Pic As IPictureDisp
    
    If m_StretchX Then cw2 = m_StretchX Else cw2 = cw
    If m_StretchY Then ch2 = m_StretchY Else ch2 = ch
    
    Set Pic = CreatePicture(cw2, ch2)
   
    dstDC = CreateCompatibleDC(0)
    dstOld = SelectObject(dstDC, Pic.Handle)
    
    If m_BackColor > -1 Then
        rc.Bottom = ch2
        rc.Right = cw2
        hBrush = CreateSolidBrush(m_BackColor)
        FillRect dstDC, rc, hBrush
        DeleteObject hBrush
    End If

    Render dstDC, 0, 0, cw2, ch2, x, y, cw, ch

    SelectObject dstDC, dstOld
    DeleteDC dstDC
    
    Set GetPicture = Pic
End Function

Private Sub Class_Initialize()
    m_Cols = 1
    m_Rows = 1
    m_BackColor = -1
End Sub
