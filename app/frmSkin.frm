VERSION 5.00
Begin VB.Form frmSkin 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1056
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.4
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   50
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   88
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Capture 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   600
      Top             =   120
   End
   Begin VB.Timer Anim 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   120
      Top             =   120
   End
End
Attribute VB_Name = "frmSkin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_Child As Boolean, m_Hover As Integer

Public Parent As frmForm, Bag As Variant, OverOutX As Long, OverOutY As Long
Public Disabled As New clsDim, Default As New clsDim, Hover As New clsDim, Down As New clsDim
Public Action As Long, HoverSimple As Boolean, CaptureSimple As Boolean, NoMoveMouse As Boolean


Public Sub Update()
    If Parent.Events("Skin" & Me.Tag & "_Update") = False Then Parent.Events "Skin_Update", Val(Me.Tag)
    
    If Me.Enabled Then
        If Action > 2 Then Action = 0
        
        Select Case Action
            Case 0
                Paint Default, Disabled, Hover, Down
            
            Case 1
                If Hover.Slides.Count Then
                    Paint Hover, Disabled, Default, Down
                Else
                    Action = 0
                End If
            
            Case 2
                If Down.Slides.Count Then
                    Paint Down, Disabled, Default, Hover
                Else
                    Action = 1
                End If
        End Select
    Else
        If Disabled.Slides.Count Then
            Paint Disabled, Default, Hover, Down
            Action = 3
        End If
    End If
    
    If Parent.Events("Skin" & Me.Tag & "_UpdateEnd") = False Then Parent.Events "Skin_UpdateEnd", Val(Me.Tag)
End Sub

Public Function Paint(ByVal Obj As clsDim, ParamArray rstObj() As Variant) As Long
    Dim cd As clsDim, rg As Long, mRect As RECT, m_hWnd As Long, m_hDC As Long, tmp As Variant
    
    With Obj
        If .Counter < 1 Then
            .Counter = 1
        ElseIf .Counter >= .Slides.Count Then
            If .One = False Then .Counter = 1
        Else
            .Counter = .Counter + 1
        End If
    
        If .Slides.Count >= .Counter Then Set cd = .Slides(.Counter)
    End With
    
    If Not cd Is Nothing Then
        With cd
            If .Picture <> 0 Then
                    Picture = .Picture
                    
                    Move Left, Top, m_HPX(Picture.Width) * Screen.TwipsPerPixelX, m_HPY(Picture.Height) * Screen.TwipsPerPixelY
                    
                    If VarType(.TransColor) <> vbBoolean Then
                        rg = m_RegionFromBitmap(.MaskPicture, .TransColor)
                        
                        If rg = 0 Then
                            Call GetWindowRect(hWnd, mRect)
                            m_hWnd = GetDesktopWindow
                            m_hDC = GetDC(m_hWnd)
                            BitBlt hDC, 0, 0, 1, 1, m_hDC, mRect.Left - 1, mRect.Top - 1, vbSrcCopy
                            ReleaseDC m_hWnd, m_hDC
                            BackColor = GetPixel(hDC, 0, 0)
                            rg = CreateRectRgn(0, 0, 1, 1)
                        End If
                    Else
                        rg = CreateRectRgn(0, 0, ScaleX(Width, vbTwips, vbPixels), ScaleY(Height, vbTwips, vbPixels))
                    End If
                    
                    Paint = SetWindowRgn(hWnd, rg, True)
                    DeleteObject rg
                    
                    If m_hDC Then Picture = LoadPicture
                
                If Not UpdateTimer(.Interval) Then UpdateTimer Obj.Interval
            End If
        End With
    End If
    
    For Each tmp In rstObj
        tmp.Counter = 0
    Next
End Function

Public Function Fill(ByVal Obj As clsDim, ByVal Pic As IPictureDisp, Optional ByVal nCols As Long = 1, Optional ByVal nRows As Long = 1, Optional ByVal TransColor As Variant, Optional ByVal mInterval As Long)
    Dim x As Long, y As Long, img1 As IPictureDisp, cd As clsDim, Clip As New clsClip
    
    With Clip
        Set .Picture = Pic
        .Cols = nCols
        .Rows = nRows

        For y = 0 To .Rows - 1
            For x = 0 To .Cols - 1
                Set cd = New clsDim
                Set img1 = .GraphicCell(y * .Cols + x)
                With cd
                    Set .Picture = img1
                    Set .MaskPicture = img1
                    .TransColor = TransColor
                    .Interval = mInterval
                End With
                Set img1 = Nothing
                Obj.Slides.Add cd
                Set cd = Nothing
            Next
        Next
    End With
End Function

Public Property Get Child() As Boolean
    Child = m_Child
End Property

Public Property Let Child(ByVal value As Boolean)
    If value Then
        SetParent Me.hWnd, Parent.hWnd
        SetWindowLongW Me.hWnd, GWL_STYLE, GetWindowLongW(Me.hWnd, GWL_STYLE) Or WS_CHILD
    Else
        SetParent Me.hWnd, 0
        SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, &H2 Or &H1
        SetWindowLongW Me.hWnd, GWL_STYLE, GetWindowLongW(Me.hWnd, GWL_STYLE) And Not WS_CHILD
    End If
    m_Child = value
End Property

Public Sub Center()
    If m_Child Then
        Me.Move (ScaleX(Parent.ScaleWidth, Parent.ScaleMode, vbTwips) - Me.Width) / 2, (ScaleY(Parent.ScaleHeight, Parent.ScaleMode, vbTwips) - Me.Height) / 2
    Else
        Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    End If
End Sub

Public Sub Interval(ByVal value As Long, ParamArray Obj() As Variant)
    Dim tmp As Variant
    For Each tmp In Obj
        tmp.Interval = value
    Next
End Sub

Private Function UpdateTimer(ByVal value As Long) As Boolean
    If value > 0 Then
        Anim.Enabled = False
        Anim.Interval = value
        Anim.Enabled = True
        UpdateTimer = True
    End If
End Function

Private Sub Anim_Timer()
    Update
End Sub

Private Sub Capture_Timer()
    CaptureMouseMove
End Sub

Private Sub CaptureMouseMove()
    Dim mRect As RECT, Pos As POINTAPI, cx As Long, cy As Long, rc As Long
    
    With Me
        rc = .hWnd
        GetCursorPos Pos:   cx = Pos.x:   cy = Pos.y
        GetWindowRect rc, mRect
        
        If HoverSimple Then
            With mRect
                If cx < .Left Or cx > .Right Or cy < .Top Or cy > .Bottom Then rc = 0
            End With
        Else
            rc = WindowFromPoint(cx, cy)
        End If
        
        OverOutX = cx - mRect.Left
        OverOutY = cy - mRect.Top
        
        If rc <> .hWnd Then
            ReleaseCapture
            m_Hover = 0
            Action = 0
            Update
            If Not m_Child Then Capture.Enabled = False
            If Parent.Events("Skin" & Me.Tag & "_MouseOut") = False Then Parent.Events "Skin_MouseOut", Val(Me.Tag)
        Else
            If m_Hover = 0 Then
                SetCapture .hWnd
                m_Hover = 1
                Action = 1
                Update
                If Not m_Child Then Capture.Enabled = True
                If Parent.Events("Skin" & Me.Tag & "_MouseOver") = False Then Parent.Events "Skin_MouseOver", Val(Me.Tag)
            End If
        End If
    End With
End Sub

Private Sub Form_Load()
    HoverSimple = True
    NoMoveMouse = True
End Sub

Private Sub Form_Click()
    If Parent.Events("Skin" & Me.Tag & "_Click") = False Then Parent.Events "Skin_Click", Val(Me.Tag)
End Sub

Private Sub Form_DblClick()
    Parent.Events "Skin" & Me.Tag & "_DblClick"
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    Parent.Events "Skin" & Me.Tag & "_OLEDragDrop", Data, Effect, Button, Shift, x, y
End Sub

Private Sub Form_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    Parent.Events "Skin" & Me.Tag & "_OLEDragOver", Data, Effect, Button, Shift, x, y, State
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Parent.Events "Skin" & Me.Tag & "_KeyDown", KeyCode, Shift
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Parent.Events "Skin" & Me.Tag & "_KeyPress", KeyAscii
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Parent.Events "Skin" & Me.Tag & "_KeyUp", KeyCode, Shift
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Action = 2
    Update
    Parent.Events "Skin" & Me.Tag & "_MouseDown", Button, Shift, x, y
    If Button = 1 And Not NoMoveMouse Then
        ReleaseCapture
        Call SendMessageW(Me.hWnd, &HA1, 2, 0)
        Call Form_MouseUp(Button, Shift, x, y)
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 0 Or CaptureSimple Then CaptureMouseMove
    Parent.Events "Skin" & Me.Tag & "_MouseMove", Button, Shift, x, y
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    m_Hover = 0
    Action = 0
    Update
    Parent.Events "Skin" & Me.Tag & "_MouseUp", Button, Shift, x, y
End Sub
