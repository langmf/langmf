VERSION 5.00
Begin VB.Form frmForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form"
   ClientHeight    =   2736
   ClientLeft      =   5040
   ClientTop       =   6732
   ClientWidth     =   5136
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.4
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmForm.frx":0000
   LinkTopic       =   "frmForm"
   MaxButton       =   0   'False
   ScaleHeight     =   2736
   ScaleWidth      =   5136
   Begin VB.TextBox MText 
      Height          =   285
      Index           =   0
      Left            =   1920
      MultiLine       =   -1  'True
      TabIndex        =   14
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ComboBox LCombo 
      Height          =   330
      Index           =   0
      Left            =   2520
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Frame Frame 
      Caption         =   "Frame"
      ClipControls    =   0   'False
      Height          =   372
      Index           =   0
      Left            =   480
      TabIndex        =   12
      Top             =   840
      Visible         =   0   'False
      Width           =   852
   End
   Begin VB.PictureBox icoInet 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   384
      Left            =   4680
      Picture         =   "frmForm.frx":058A
      ScaleHeight     =   384
      ScaleWidth      =   384
      TabIndex        =   11
      Top             =   2280
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.TextBox TextBox 
      Height          =   285
      Index           =   0
      Left            =   2400
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Index           =   0
      Left            =   0
      Top             =   840
   End
   Begin VB.VScrollBar VScroll 
      Height          =   735
      Index           =   0
      Left            =   4200
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.HScrollBar HScroll 
      Height          =   255
      Index           =   0
      Left            =   3720
      TabIndex        =   8
      Top             =   840
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ListBox List 
      Height          =   240
      Index           =   0
      Left            =   1920
      TabIndex        =   7
      Top             =   480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.OptionButton Opt 
      Caption         =   "Option"
      Height          =   255
      Index           =   0
      Left            =   960
      TabIndex        =   6
      Top             =   480
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CheckBox Check 
      Caption         =   "Check"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   5
      Top             =   480
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ComboBox Combo 
      Height          =   330
      Index           =   0
      Left            =   3360
      TabIndex        =   4
      Text            =   "Combo"
      Top             =   480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text 
      Height          =   285
      Index           =   0
      Left            =   1440
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox Pic 
      AutoRedraw      =   -1  'True
      ClipControls    =   0   'False
      Height          =   255
      Index           =   0
      Left            =   960
      ScaleHeight     =   204
      ScaleWidth      =   324
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command 
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Img 
      Height          =   372
      Index           =   0
      Left            =   1440
      Top             =   840
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.Line CLine 
      Index           =   0
      Visible         =   0   'False
      X1              =   3168
      X2              =   3168
      Y1              =   192
      Y2              =   0
   End
   Begin VB.Shape CShape 
      Height          =   255
      Index           =   0
      Left            =   3600
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "label"
      Height          =   210
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   330
   End
End
Attribute VB_Name = "frmForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_frmWsk As frmWsk, m_MinMax As MINMAXINFO, m_CT As Long
Private m_Child As Boolean, m_CodeObject As Object, m_onEvent As Object

Public Parent As Object
Public SubClass As clsSubClass
Public Style As clsFormStyle
Public Menu As clsMenu
Public Tray As frmTray
Public Resize As clsHash
Public WC As clsHash

Public xMin As Long, yMin As Long, xMax As Long, yMax As Long
Public gdip_mDC As Long, gdip_MainBitmap As Long, gdip_OldBitmap As Long
Public NoMoveMouse As Boolean, NoOverOutPic As Boolean, NoOverOutFrame As Boolean, NoOverOutCommand As Boolean


Public Function Ctrl(ByVal typeObj As String, Optional ByVal ID As Long, Optional ByVal isAdd As Boolean = True, Optional ByVal dataArg As Variant) As Object
    Dim Obj As Object, txt() As String

    If LenB(typeObj) = 0 Or ID < 0 Then Exit Function
    
    typeObj = LCase$(typeObj)
    
    Select Case typeObj
        Case Is = "check":     If isAdd Then Load Me.Check(ID):    Set Obj = Me.Check(ID) Else:    Unload Me.Check(ID)
        Case Is = "cline":     If isAdd Then Load Me.CLine(ID):    Set Obj = Me.CLine(ID) Else:    Unload Me.CLine(ID)
        Case Is = "combo":     If isAdd Then Load Me.Combo(ID):    Set Obj = Me.Combo(ID) Else:    Unload Me.Combo(ID)
        Case Is = "command":   If isAdd Then Load Me.Command(ID):  Set Obj = Me.Command(ID) Else:  Unload Me.Command(ID)
        Case Is = "cshape":    If isAdd Then Load Me.CShape(ID):   Set Obj = Me.CShape(ID) Else:   Unload Me.CShape(ID)
        Case Is = "frame":     If isAdd Then Load Me.Frame(ID):    Set Obj = Me.Frame(ID) Else:    Unload Me.Frame(ID)
        Case Is = "hscroll":   If isAdd Then Load Me.HScroll(ID):  Set Obj = Me.HScroll(ID) Else:  Unload Me.HScroll(ID)
        Case Is = "img":       If isAdd Then Load Me.Img(ID):      Set Obj = Me.Img(ID) Else:      Unload Me.Img(ID)
        Case Is = "label":     If isAdd Then Load Me.Label(ID):    Set Obj = Me.Label(ID) Else:    Unload Me.Label(ID)
        Case Is = "lcombo":    If isAdd Then Load Me.LCombo(ID):   Set Obj = Me.LCombo(ID) Else:   Unload Me.LCombo(ID)
        Case Is = "list":      If isAdd Then Load Me.List(ID):     Set Obj = Me.List(ID) Else:     Unload Me.List(ID)
        Case Is = "mtext":     If isAdd Then Load Me.MText(ID):    Set Obj = Me.MText(ID) Else:    Unload Me.MText(ID)
        Case Is = "opt":       If isAdd Then Load Me.Opt(ID):      Set Obj = Me.Opt(ID) Else:      Unload Me.Opt(ID)
        Case Is = "pic":       If isAdd Then Load Me.Pic(ID):      Set Obj = Me.Pic(ID) Else:      Unload Me.Pic(ID)
        Case Is = "text":      If isAdd Then Load Me.Text(ID):     Set Obj = Me.Text(ID) Else:     Unload Me.Text(ID)
        Case Is = "textbox":   If isAdd Then Load Me.TextBox(ID):  Set Obj = Me.TextBox(ID) Else:  Unload Me.TextBox(ID)
        Case Is = "timer":     If isAdd Then Load Me.Timer(ID):    Set Obj = Me.Timer(ID) Else:    Unload Me.Timer(ID)
        Case Is = "vscroll":   If isAdd Then Load Me.VScroll(ID):  Set Obj = Me.VScroll(ID) Else:  Unload Me.VScroll(ID)

        Case Is = "wsk":       Set Obj = WskCtrl(ID, isAdd)
        Case Is = "skin":      Set Obj = SkinCtrl(ID, isAdd)
        Case Is = "pbar":      Set Obj = CreateWC(typeObj & ID, IIF(isAdd, "msctls_progress32", ""), &H1)
        
        Case Else
            txt = Split(typeObj, "=")
            ReDim Preserve txt(2)
            Set Obj = CreateOCX(Trim$(txt(0)), IIF(isAdd, Trim$(txt(1)), ""), Val(txt(2)))
    End Select
    
    If isAdd Then
        If Not Obj Is Nothing Then
            CBN Obj, "Visible", VbLet, Array(True)
            Call m_DoParams(Obj, dataArg)
            Set Ctrl = Obj
        End If
    Else
        Resize.Remove typeObj & ID
    End If
End Function

Public Function Add(ByVal typeObj As String, ByVal ID As Long, ParamArray dataArg() As Variant) As Object
    Set Add = Ctrl(typeObj, ID, True, dataArg)
End Function

Public Function Remove(ByVal typeObj As String, Optional ByVal ID As Long)
    Ctrl typeObj, ID, False
End Function


Public Property Get PBar(ByVal ID As Long) As clsWC
    If WC.Exists("pbar" & ID) Then Set PBar = WC("pbar" & ID)
End Property

Public Property Get Skin(ByVal ID As Long) As frmSkin
    If WC.Exists("skin" & ID) Then Set Skin = WC("skin" & ID)
End Property

Public Property Get Wsk(ByVal ID As Long) As Winsock
    Set Wsk = m_frmWsk.Wsk(ID)
End Property

Private Function SkinCtrl(ByVal ID As Long, ByVal isAdd As Boolean) As frmSkin
    Dim Obj As New frmSkin
    If isAdd Then
        With Obj:    Set .Parent = Me:    .Tag = ID:    .Child = True:    .Move 0, 0, 0, 0:     End With
        WC.Add Obj, "skin" & ID
        Set SkinCtrl = Obj
    Else
        WC.Remove "skin" & ID
    End If
End Function

Private Function WskCtrl(ByVal ID As Long, ByVal isAdd As Boolean) As Winsock
    On Error GoTo err1
    If m_frmWsk Is Nothing Then Set m_frmWsk = New frmWsk:    Set m_frmWsk.Parent = Me
    If isAdd Then Load m_frmWsk.Wsk(ID):  Set WskCtrl = m_frmWsk.Wsk(ID) Else:  Unload m_frmWsk.Wsk(ID)
err1:
End Function

Public Function CreateWC(ByVal NameControl As String, Optional ByVal NameClass As String, Optional ByVal txtCaption As String, Optional ByVal dwStyle As Long, Optional ByVal dwExStyle As Long) As clsWC
    Dim Obj As New clsWC
    If LenB(NameClass) Then
        Obj.Create Me, NameClass, NameControl, txtCaption, dwStyle, dwExStyle
        WC.Add Obj, NameControl
        Set CreateWC = Obj
    Else
        WC.Remove NameControl
    End If
End Function

Public Function CreateOCX(ByVal NameControl As String, Optional ByVal strGUID As String, Optional ByVal NotWrapEvent As Boolean) As Object
    Dim OCX As New clsOCX
    If LenB(strGUID) Then
        Set CreateOCX = OCX.Create(Me, NameControl, strGUID, NotWrapEvent)
        WC.Add OCX, NameControl
    Else
        WC.Remove NameControl
    End If
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

Public Sub Move2(ByVal Obj As Object, Optional ByVal typeX As Single = -1, Optional ByVal typeY As Single = -1, Optional ByVal offsetX As Single = 0, Optional ByVal offsetY As Single = 0, Optional ByVal typeW As Single = 0, Optional ByVal typeH As Single = 0, Optional ByVal AddItem As Variant)
    Dim x As Single, y As Single, v As Variant, txt As String

    On Error Resume Next
    
    x = Obj.Left:   y = Obj.Top
    FlexMove Obj, x, y, , , Me.ScaleWidth, Me.ScaleHeight, typeX, typeY, offsetX, offsetY, typeW, typeH
    Obj.Move x, y

    If Not IsMissing(AddItem) And Not IsEmpty(AddItem) Then
        If VarType(AddItem) = vbString Then txt = AddItem Else txt = Obj.Name & Obj.Index
        v = Array(Obj, typeX, typeY, offsetX, offsetY, typeW, typeH)
        If LenB(txt) Then Resize.Add v, txt Else Resize.Add v
    End If
    
    Err.Clear
End Sub

Public Function IsFont(ByVal nameFont As String) As Boolean
    On Error GoTo err1
    frmScript.FontName = nameFont
    IsFont = True
err1:
End Function

Public Property Get UniCaption(Optional ByVal Obj As Object) As String
    Dim lngLen As Long, lngPtr As Long

    If Obj Is Nothing Then Set Obj = Me
    If (TypeOf Obj Is CheckBox) Or (TypeOf Obj Is CommandButton) Or (TypeOf Obj Is Form) Or (TypeOf Obj Is Frame) Or (TypeOf Obj Is OptionButton) Then
        lngLen = DefWindowProcW(Obj.hWnd, WM_GETTEXTLENGTH, 0, 0)
        If lngLen Then
            lngPtr = SysAllocStringLen(0, lngLen)
            PutMem4 VarPtr(UniCaption), lngPtr
            DefWindowProcW Obj.hWnd, WM_GETTEXT, lngLen + 1, lngPtr
        End If
    Else
        On Error Resume Next
        UniCaption = Obj
    End If
End Property

Public Property Let UniCaption(Optional ByVal Obj As Object, ByVal value As String)
    If Obj Is Nothing Then Set Obj = Me
    If (TypeOf Obj Is CheckBox) Or (TypeOf Obj Is CommandButton) Or (TypeOf Obj Is Form) Or (TypeOf Obj Is Frame) Or (TypeOf Obj Is OptionButton) Then
        DefWindowProcW Obj.hWnd, WM_SETTEXT, 0, ByVal StrPtr(value)
    Else
        On Error Resume Next
        Obj = value
    End If
End Property



'========================= Timer ===========================
Private Sub Timer_Timer(Index As Integer)
    Events "Timer" & Index & "_Timer"
End Sub


'========================= Check ===========================
Private Sub Check_Click(Index As Integer)
    If Events("Check" & Index & "_Click") = False Then Events "Check_Click", Index
End Sub

Private Sub Check_DragDrop(Index As Integer, Source As Control, x As Single, y As Single)
    Events "Check" & Index & "_DragDrop", Source, x, y
End Sub

Private Sub Check_DragOver(Index As Integer, Source As Control, x As Single, y As Single, State As Integer)
    Events "Check" & Index & "_DragOver", Source, x, y, State
End Sub

Private Sub Check_KeyPress(Index As Integer, KeyAscii As Integer)
    Events "Check" & Index & "_KeyPress", KeyAscii
End Sub

Private Sub Check_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Events "Check" & Index & "_MouseDown", Button, Shift, x, y
End Sub

Private Sub Check_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Events "Check" & Index & "_MouseUp", Button, Shift, x, y
End Sub


'========================= Opt ===========================
Private Sub Opt_Click(Index As Integer)
    If Events("Opt" & Index & "_Click") = False Then Events "Opt_Click", Index
End Sub

Private Sub Opt_DragDrop(Index As Integer, Source As Control, x As Single, y As Single)
    Events "Opt" & Index & "_DragDrop", Source, x, y
End Sub

Private Sub Opt_DragOver(Index As Integer, Source As Control, x As Single, y As Single, State As Integer)
    Events "Opt" & Index & "_DragOver", Source, x, y, State
End Sub

Private Sub Opt_KeyPress(Index As Integer, KeyAscii As Integer)
    Events "Opt" & Index & "_KeyPress", KeyAscii
End Sub

Private Sub Opt_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Events "Opt" & Index & "_MouseDown", Button, Shift, x, y
End Sub

Private Sub Opt_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Events "Opt" & Index & "_MouseUp", Button, Shift, x, y
End Sub


'========================= Combo ===========================
Private Sub Combo_Change(Index As Integer)
    Events "Combo" & Index & "_Change"
End Sub

Private Sub Combo_Click(Index As Integer)
    If Events("Combo" & Index & "_Click") = False Then Events "Combo_Click", Index
End Sub

Private Sub Combo_DragDrop(Index As Integer, Source As Control, x As Single, y As Single)
    Events "Combo" & Index & "_DragDrop", Source, x, y
End Sub

Private Sub Combo_DragOver(Index As Integer, Source As Control, x As Single, y As Single, State As Integer)
    Events "Combo" & Index & "_DragOver", Source, x, y, State
End Sub

Private Sub Combo_DropDown(Index As Integer)
    Events "Combo" & Index & "_DropDown"
End Sub

Private Sub Combo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Events "Combo" & Index & "_KeyDown", KeyCode, Shift
End Sub

Private Sub Combo_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    Events "Combo" & Index & "_KeyUp", KeyCode, Shift
End Sub

Private Sub Combo_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    Events "Combo" & Index & "_OLEDragDrop", Data, Effect, Button, Shift, x, y
End Sub

Private Sub Combo_OLEDragOver(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    Events "Combo" & Index & "_OLEDragOver", Data, Effect, Button, Shift, x, y, State
End Sub

Private Sub Combo_Scroll(Index As Integer)
    Events "Combo" & Index & "_Scroll"
End Sub

Private Sub Combo_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Events("Combo" & Index & "_Validate")
End Sub


'========================= LCombo ===========================
Private Sub LCombo_Change(Index As Integer)
    Events "LCombo" & Index & "_Change"
End Sub

Private Sub LCombo_Click(Index As Integer)
    If Events("LCombo" & Index & "_Click") = False Then Events "LCombo_Click", Index
End Sub

Private Sub LCombo_DragDrop(Index As Integer, Source As Control, x As Single, y As Single)
    Events "LCombo" & Index & "_DragDrop", Source, x, y
End Sub

Private Sub LCombo_DragOver(Index As Integer, Source As Control, x As Single, y As Single, State As Integer)
    Events "LCombo" & Index & "_DragOver", Source, x, y, State
End Sub

Private Sub LCombo_DropDown(Index As Integer)
    Events "LCombo" & Index & "_DropDown"
End Sub

Private Sub LCombo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Events "LCombo" & Index & "_KeyDown", KeyCode, Shift
End Sub

Private Sub LCombo_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    Events "LCombo" & Index & "_KeyUp", KeyCode, Shift
End Sub

Private Sub LCombo_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    Events "LCombo" & Index & "_OLEDragDrop", Data, Effect, Button, Shift, x, y
End Sub

Private Sub LCombo_OLEDragOver(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    Events "LCombo" & Index & "_OLEDragOver", Data, Effect, Button, Shift, x, y, State
End Sub

Private Sub LCombo_Scroll(Index As Integer)
    Events "LCombo" & Index & "_Scroll"
End Sub

Private Sub LCombo_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Events("LCombo" & Index & "_Validate")
End Sub


'========================= List ===========================
Private Sub List_Click(Index As Integer)
    If Events("List" & Index & "_Click") = False Then Events "List_Click", Index
End Sub

Private Sub List_DblClick(Index As Integer)
    Events "List" & Index & "_DblClick"
End Sub

Private Sub List_DragDrop(Index As Integer, Source As Control, x As Single, y As Single)
    Events "List" & Index & "_DragDrop", Source, x, y
End Sub

Private Sub List_DragOver(Index As Integer, Source As Control, x As Single, y As Single, State As Integer)
    Events "List" & Index & "_DragOver", Source, x, y, State
End Sub

Private Sub List_ItemCheck(Index As Integer, Item As Integer)
    Events "List" & Index & "_ItemCheck", Item
End Sub

Private Sub List_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Events "List" & Index & "_KeyDown", KeyCode, Shift
End Sub

Private Sub List_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    Events "List" & Index & "_KeyUp", KeyCode, Shift
End Sub

Private Sub List_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    Events "List" & Index & "_OLEDragDrop", Data, Effect, Button, Shift, x, y
End Sub

Private Sub List_OLEDragOver(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    Events "List" & Index & "_OLEDragOver", Data, Effect, Button, Shift, x, y, State
End Sub

Private Sub List_Scroll(Index As Integer)
    Events "List" & Index & "_Scroll"
End Sub

Private Sub List_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Events("List" & Index & "_Validate")
End Sub


'========================= HScroll ===========================
Private Sub HScroll_Change(Index As Integer)
    Events "HScroll" & Index & "_Change"
End Sub

Private Sub HScroll_DragDrop(Index As Integer, Source As Control, x As Single, y As Single)
    Events "HScroll" & Index & "_DragDrop", Source, x, y
End Sub

Private Sub HScroll_DragOver(Index As Integer, Source As Control, x As Single, y As Single, State As Integer)
    Events "HScroll" & Index & "_DragOver", Source, x, y, State
End Sub

Private Sub HScroll_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Events "HScroll" & Index & "_KeyDown", KeyCode, Shift
End Sub

Private Sub HScroll_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    Events "HScroll" & Index & "_KeyUp", KeyCode, Shift
End Sub

Private Sub HScroll_Scroll(Index As Integer)
    Events "HScroll" & Index & "_Scroll"
End Sub


'========================= VScroll ===========================
Private Sub VScroll_Change(Index As Integer)
    Events "VScroll" & Index & "_Change"
End Sub

Private Sub VScroll_DragDrop(Index As Integer, Source As Control, x As Single, y As Single)
    Events "VScroll" & Index & "_DragDrop", Source, x, y
End Sub

Private Sub VScroll_DragOver(Index As Integer, Source As Control, x As Single, y As Single, State As Integer)
    Events "VScroll" & Index & "_DragOver", Source, x, y, State
End Sub

Private Sub VScroll_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Events "VScroll" & Index & "_KeyDown", KeyCode, Shift
End Sub

Private Sub VScroll_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    Events "VScroll" & Index & "_KeyUp", KeyCode, Shift
End Sub

Private Sub VScroll_Scroll(Index As Integer)
    Events "VScroll" & Index & "_Scroll"
End Sub


'========================= Command ===========================
Private Sub Command_Click(Index As Integer)
    If Events("Command" & Index & "_Click") = False Then Events "Command_Click", Index
End Sub

Private Sub Command_DragDrop(Index As Integer, Source As Control, x As Single, y As Single)
    Events "Command" & Index & "_DragDrop", Source, x, y
End Sub

Private Sub Command_DragOver(Index As Integer, Source As Control, x As Single, y As Single, State As Integer)
    Events "Command" & Index & "_DragOver", Source, x, y, State
End Sub

Private Sub Command_KeyPress(Index As Integer, KeyAscii As Integer)
    Events "Command" & Index & "_KeyPress", KeyAscii
End Sub

Private Sub Command_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Events "Command" & Index & "_MouseDown", Button, Shift, x, y
End Sub

Private Sub Command_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Events "Command" & Index & "_MouseMove", Button, Shift, x, y

    If Not NoOverOutCommand Then
        With Command(Index)
            If (x < 0) Or (y < 0) Or (x > ScaleX(.Width, ScaleMode, vbTwips)) Or (y > ScaleY(.Height, ScaleMode, vbTwips)) Then
                ReleaseCapture
                .Tag = "0"
                Events "Command" & Index & "_MouseOut", Button, Shift, x, y
            Else
                If Val(.Tag) = 0 Then
                   SetCapture .hWnd
                   .Tag = "1"
                    Events "Command" & Index & "_MouseOver", Button, Shift, x, y
                End If
            End If
        End With
    End If
End Sub

Private Sub Command_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Not NoOverOutCommand Then
        ReleaseCapture
        Command(Index).Tag = "0"
    End If
    Events "Command" & Index & "_MouseUp", Button, Shift, x, y
End Sub

Private Sub Command_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    Events "Command" & Index & "_OLEDragDrop", Data, Effect, Button, Shift, x, y
End Sub

Private Sub Command_OLEDragOver(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    Events "Command" & Index & "_OLEDragOver", Data, Effect, Button, Shift, x, y, State
End Sub


'========================= Frame ===========================
Private Sub Frame_Click(Index As Integer)
    If Events("Frame" & Index & "_Click") = False Then Events "Frame_Click", Index
End Sub

Private Sub Frame_DblClick(Index As Integer)
    Events "Frame" & Index & "_DblClick"
End Sub

Private Sub Frame_DragDrop(Index As Integer, Source As Control, x As Single, y As Single)
    Events "Frame" & Index & "_DragDrop", Source, x, y
End Sub

Private Sub Frame_DragOver(Index As Integer, Source As Control, x As Single, y As Single, State As Integer)
    Events "Frame" & Index & "_DragOver", Source, x, y, State
End Sub

Private Sub Frame_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Events "Frame" & Index & "_MouseDown", Button, Shift, x, y
End Sub

Private Sub Frame_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Events "Frame" & Index & "_MouseMove", Button, Shift, x, y
    
    If Not NoOverOutFrame Then
        With Frame(Index)
            If (x < 0) Or (y < 0) Or (x > ScaleX(.Width, ScaleMode, vbTwips)) Or (y > ScaleY(.Height, ScaleMode, vbTwips)) Then
                ReleaseCapture
                .Tag = "0"
                Events "Frame" & Index & "_MouseOut", Button, Shift, x, y
            Else
                If Val(.Tag) = 0 Then
                    SetCapture .hWnd
                    .Tag = "1"
                    Events "Frame" & Index & "_MouseOver", Button, Shift, x, y
                End If
            End If
        End With
    End If
End Sub

Private Sub Frame_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Not NoOverOutFrame Then
        ReleaseCapture
        Frame(Index).Tag = "0"
    End If
    Events "Frame" & Index & "_MouseUp", Button, Shift, x, y
End Sub

Private Sub Frame_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    Events "Frame" & Index & "_OLEDragDrop", Data, Effect, Button, Shift, x, y
End Sub

Private Sub Frame_OLEDragOver(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    Events "Frame" & Index & "_OLEDragOver", Data, Effect, Button, Shift, x, y, State
End Sub


'========================= Label ===========================
Private Sub Label_Click(Index As Integer)
    If Events("Label" & Index & "_Click") = False Then Events "Label_Click", Index
End Sub

Private Sub Label_DblClick(Index As Integer)
    Events "Label" & Index & "_DblClick"
End Sub

Private Sub Label_DragDrop(Index As Integer, Source As Control, x As Single, y As Single)
    Events "Label" & Index & "_DragDrop", Source, x, y
End Sub

Private Sub Label_DragOver(Index As Integer, Source As Control, x As Single, y As Single, State As Integer)
    Events "Label" & Index & "_DragOver", Source, x, y, State
End Sub

Private Sub Label_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Events "Label" & Index & "_MouseDown", Button, Shift, x, y
End Sub

Private Sub Label_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Events "Label" & Index & "_MouseMove", Button, Shift, x, y
End Sub

Private Sub Label_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Events "Label" & Index & "_MouseUp", Button, Shift, x, y
End Sub

Private Sub Label_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    Events "Label" & Index & "_OLEDragDrop", Data, Effect, Button, Shift, x, y
End Sub

Private Sub Label_OLEDragOver(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    Events "Label" & Index & "_OLEDragOver", Data, Effect, Button, Shift, x, y, State
End Sub


'========================= Img ===========================
Private Sub Img_Click(Index As Integer)
    If Events("Img" & Index & "_Click") = False Then Events "Img_Click", Index
End Sub

Private Sub Img_DblClick(Index As Integer)
    Events "Img" & Index & "_DblClick"
End Sub

Private Sub Img_DragDrop(Index As Integer, Source As Control, x As Single, y As Single)
    Events "Img" & Index & "_DragDrop", Source, x, y
End Sub

Private Sub Img_DragOver(Index As Integer, Source As Control, x As Single, y As Single, State As Integer)
    Events "Img" & Index & "_DragOver", Source, x, y, State
End Sub

Private Sub Img_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Events "Img" & Index & "_MouseDown", Button, Shift, x, y
End Sub

Private Sub Img_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Events "Img" & Index & "_MouseMove", Button, Shift, x, y
End Sub

Private Sub Img_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Events "Img" & Index & "_MouseUp", Button, Shift, x, y
End Sub

Private Sub Img_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    Events "Img" & Index & "_OLEDragDrop", Data, Effect, Button, Shift, x, y
End Sub

Private Sub Img_OLEDragOver(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    Events "Img" & Index & "_OLEDragOver", Data, Effect, Button, Shift, x, y, State
End Sub


'========================= Pic ===========================
Private Sub Pic_Click(Index As Integer)
    If Events("Pic" & Index & "_Click") = False Then Events "Pic_Click", Index
End Sub

Private Sub Pic_DblClick(Index As Integer)
    Events "Pic" & Index & "_DblClick"
End Sub

Private Sub Pic_DragDrop(Index As Integer, Source As Control, x As Single, y As Single)
    Events "Pic" & Index & "_DragDrop", Source, x, y
End Sub

Private Sub Pic_DragOver(Index As Integer, Source As Control, x As Single, y As Single, State As Integer)
    Events "Pic" & Index & "_DragOver", Source, x, y, State
End Sub

Private Sub Pic_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Events "Pic" & Index & "_KeyDown", KeyCode, Shift
End Sub

Private Sub Pic_KeyPress(Index As Integer, KeyAscii As Integer)
    Events "Pic" & Index & "_KeyPress", KeyAscii
End Sub

Private Sub Pic_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    Events "Pic" & Index & "_KeyUp", KeyCode, Shift
End Sub

Private Sub Pic_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Events "Pic" & Index & "_MouseDown", Button, Shift, x, y
End Sub

Private Sub Pic_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Events "Pic" & Index & "_MouseMove", Button, Shift, x, y
    
    If Not NoOverOutPic Then
        With Pic(Index)
            If (x < 0) Or (y < 0) Or (x > ScaleX(.Width, ScaleMode, .ScaleMode)) Or (y > ScaleY(.Height, ScaleMode, .ScaleMode)) Then
                ReleaseCapture
                .Tag = "0"
                Events "Pic" & Index & "_MouseOut", Button, Shift, x, y
            Else
                If Val(.Tag) = 0 Then
                    SetCapture .hWnd
                    .Tag = "1"
                    Events "Pic" & Index & "_MouseOver", Button, Shift, x, y
                End If
            End If
        End With
    End If
End Sub

Private Sub Pic_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Not NoOverOutPic Then
        ReleaseCapture
        Pic(Index).Tag = "0"
    End If
    Events "Pic" & Index & "_MouseUp", Button, Shift, x, y
End Sub

Private Sub Pic_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    Events "Pic" & Index & "_OLEDragDrop", Data, Effect, Button, Shift, x, y
End Sub

Private Sub Pic_OLEDragOver(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    Events "Pic" & Index & "_OLEDragOver", Data, Effect, Button, Shift, x, y, State
End Sub

Private Sub Pic_Paint(Index As Integer)
    Events "Pic" & Index & "_Paint"
End Sub

Private Sub Pic_Resize(Index As Integer)
    Events "Pic" & Index & "_Resize"
End Sub

Private Sub Pic_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Events("Pic" & Index & "_Validate")
End Sub


'========================= Text ===========================
Private Sub Text_Change(Index As Integer)
    Events "Text" & Index & "_Change"
End Sub

Private Sub Text_Click(Index As Integer)
    If Events("Text" & Index & "_Click") = False Then Events "Text_Click", Index
End Sub

Private Sub Text_DblClick(Index As Integer)
    Events "Text" & Index & "_DblClick"
End Sub

Private Sub Text_DragDrop(Index As Integer, Source As Control, x As Single, y As Single)
    Events "Text" & Index & "_DragDrop", Source, x, y
End Sub

Private Sub Text_DragOver(Index As Integer, Source As Control, x As Single, y As Single, State As Integer)
    Events "Text" & Index & "_DragOver", Source, x, y, State
End Sub

Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Events "Text" & Index & "_KeyDown", KeyCode, Shift
End Sub

Private Sub Text_KeyPress(Index As Integer, KeyAscii As Integer)
    Events "Text" & Index & "_KeyPress", KeyAscii
End Sub

Private Sub Text_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    Events "Text" & Index & "_KeyUp", KeyCode, Shift
End Sub

Private Sub Text_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Events "Text" & Index & "_MouseDown", Button, Shift, x, y
End Sub

Private Sub Text_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Events "Text" & Index & "_MouseUp", Button, Shift, x, y
End Sub

Private Sub Text_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    Events "Text" & Index & "_OLEDragDrop", Data, Effect, Button, Shift, x, y
End Sub

Private Sub Text_OLEDragOver(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    Events "Text" & Index & "_OLEDragOver", Data, Effect, Button, Shift, x, y, State
End Sub

Private Sub Text_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Events("Text" & Index & "_Validate")
End Sub


'========================= MText ===========================
Private Sub MText_Change(Index As Integer)
    Events "MText" & Index & "_Change"
End Sub

Private Sub MText_Click(Index As Integer)
    If Events("MText" & Index & "_Click") = False Then Events "MText_Click", Index
End Sub

Private Sub MText_DblClick(Index As Integer)
    Events "MText" & Index & "_DblClick"
End Sub

Private Sub MText_DragDrop(Index As Integer, Source As Control, x As Single, y As Single)
    Events "MText" & Index & "_DragDrop", Source, x, y
End Sub

Private Sub MText_DragOver(Index As Integer, Source As Control, x As Single, y As Single, State As Integer)
    Events "MText" & Index & "_DragOver", Source, x, y, State
End Sub

Private Sub MText_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Events "MText" & Index & "_KeyDown", KeyCode, Shift
End Sub

Private Sub MText_KeyPress(Index As Integer, KeyAscii As Integer)
    Events "MText" & Index & "_KeyPress", KeyAscii
End Sub

Private Sub MText_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    Events "MText" & Index & "_KeyUp", KeyCode, Shift
End Sub

Private Sub MText_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Events "MText" & Index & "_MouseDown", Button, Shift, x, y
End Sub

Private Sub MText_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Events "MText" & Index & "_MouseUp", Button, Shift, x, y
End Sub

Private Sub MText_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    Events "MText" & Index & "_OLEDragDrop", Data, Effect, Button, Shift, x, y
End Sub

Private Sub MText_OLEDragOver(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    Events "MText" & Index & "_OLEDragOver", Data, Effect, Button, Shift, x, y, State
End Sub

Private Sub MText_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Events("MText" & Index & "_Validate")
End Sub


'========================= TextBox ===========================
Private Sub TextBox_Change(Index As Integer)
    Events "TextBox" & Index & "_Change"
End Sub

Private Sub TextBox_Click(Index As Integer)
    If Events("TextBox" & Index & "_Click") = False Then Events "TextBox_Click", Index
End Sub

Private Sub TextBox_DblClick(Index As Integer)
    Events "TextBox" & Index & "_DblClick"
End Sub

Private Sub TextBox_DragDrop(Index As Integer, Source As Control, x As Single, y As Single)
    Events "TextBox" & Index & "_DragDrop", Source, x, y
End Sub

Private Sub TextBox_DragOver(Index As Integer, Source As Control, x As Single, y As Single, State As Integer)
    Events "TextBox" & Index & "_DragOver", Source, x, y, State
End Sub

Private Sub TextBox_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Events "TextBox" & Index & "_KeyDown", KeyCode, Shift
End Sub

Private Sub TextBox_KeyPress(Index As Integer, KeyAscii As Integer)
    Events "TextBox" & Index & "_KeyPress", KeyAscii
End Sub

Private Sub TextBox_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    Events "TextBox" & Index & "_KeyUp", KeyCode, Shift
End Sub

Private Sub TextBox_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Events "TextBox" & Index & "_MouseDown", Button, Shift, x, y
End Sub

Private Sub TextBox_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Events "TextBox" & Index & "_MouseUp", Button, Shift, x, y
End Sub

Private Sub TextBox_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    Events "TextBox" & Index & "_OLEDragDrop", Data, Effect, Button, Shift, x, y
End Sub

Private Sub TextBox_OLEDragOver(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    Events "TextBox" & Index & "_OLEDragOver", Data, Effect, Button, Shift, x, y, State
End Sub

Private Sub TextBox_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Events("TextBox" & Index & "_Validate")
End Sub


'========================= Form ===========================
Private Sub Form_Initialize()
    Set WC = New clsHash
    Set Resize = New clsHash
    Set SubClass = New clsSubClass
    Set Style = New clsFormStyle
    Set Tray = New frmTray
    Set Tray.Parent = Me
    Set Menu = New clsMenu
    Menu.Initialize Me, New clsHash
    NoMoveMouse = True
End Sub

Private Sub Form_Terminate()
    Set WC = Nothing
    Set Resize = Nothing
    Set SubClass = Nothing
    Set Style = Nothing
    Set Menu = Nothing
    Set Tray.Parent = Nothing
    Set Tray = Nothing
End Sub

Private Sub Form_Load()
    Style.hWnd = Me.hWnd
    Set Tray.TrayIcon = Me.Icon
    
    Events "Form_Load"
    
    If WinVer.dwMajorVersion >= 6 Then Call ChangeWindowMessageFilter(WM_COMMAND, MSGFLT_ADD)
    
    SubClass.List(WM_COMMAND, WM_HOTKEY, WM_GETMINMAXINFO, WM_ACTIVATE) = 1
    SubClass.HookSet Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = Events("Form_Unload")

    Select Case Cancel
        Case 2
            Cancel = 1
            Me.Hide
            
        Case 0
            Tray.InTray = False
            SubClass.HookClear Me
            WC.Init
            Resize.Init
            
            If Not m_frmWsk Is Nothing Then
                Set m_frmWsk.Parent = Nothing
                Unload m_frmWsk
                Set m_frmWsk = Nothing
            End If

            SetMenu hWnd, 0
            Set Menu = New clsMenu
            Menu.Initialize Me, New clsHash
    End Select
End Sub

Private Sub Form_Resize()
    Dim v As Variant
    
    On Error Resume Next
    
    For Each v In Resize.Items
        Move2 v(0), v(1), v(2), v(3), v(4), v(5), v(6)
    Next

    Events "Form_Resize"
End Sub

Private Sub Form_Paint()
    Events "Form_Paint"
End Sub

Private Sub Form_Activate()
    Events "Form_Activate"
End Sub

Private Sub Form_LostFocus()
    Events "Form_LostFocus"
End Sub

Private Sub Form_Click()
    Events "Form_Click"
End Sub

Private Sub Form_DblClick()
    Events "Form_DblClick"
End Sub

Private Sub Form_DragDrop(Source As Control, x As Single, y As Single)
    Events "Form_DragDrop", Source, x, y
End Sub

Private Sub Form_DragOver(Source As Control, x As Single, y As Single, State As Integer)
    Events "Form_DragOver", Source, x, y, State
End Sub

Private Sub Form_GotFocus()
    Events "Form_GotFocus"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Events "Form_KeyDown", KeyCode, Shift
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Events "Form_KeyPress", KeyAscii
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Events "Form_KeyUp", KeyCode, Shift
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Events "Form_MouseDown", Button, Shift, x, y

    If Button = 1 And Not NoMoveMouse Then
        ReleaseCapture
        Call SendMessageW(Me.hWnd, &HA1, 2, 0)
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Events "Form_MouseMove", Button, Shift, x, y
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Events "Form_MouseUp", Button, Shift, x, y
End Sub

Private Sub Form_OLECompleteDrag(Effect As Long)
    Events "Form_OLECompleteDrag", Effect
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    Events "Form_OLEDragDrop", Data, Effect, Button, Shift, x, y
End Sub

Private Sub Form_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    Events "Form_OLEDragOver", Data, Effect, Button, Shift, x, y, State
End Sub

Private Sub Form_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
    Events "Form_OLEGiveFeedback", Effect, DefaultCursors
End Sub

Private Sub Form_OLESetData(Data As DataObject, DataFormat As Integer)
    Events "Form_OLESetData", Data, DataFormat
End Sub

Private Sub Form_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    Events "Form_OLEStartDrag", Data, AllowedEffects
End Sub


'========================= WindowProc ===========================
Public Function WindowProc(ByRef bHandled As Boolean, ByVal u_hWnd As Long, ByVal uMsg As Long, ByRef wParam As Long, ByRef lParam As Long, ByVal dwRefData As Long) As Long
    WindowProc = Events("WindowProc", Deref(VarPtr(u_hWnd) - 4), u_hWnd, uMsg, wParam, lParam, dwRefData)

    Select Case uMsg
        Case WM_ACTIVATE
            If wParam > 0 And wParam < 3 Then Menu.HotKey (True) Else Menu.HotKey (False)
            
        Case WM_COMMAND
            If lParam = 0 Then Call Menu.Click(wParam)
            If (wParam \ &H10000) = THBN_CLICKED Then Call Events("TB_Click", wParam And &HFFFF&)
            
        Case WM_HOTKEY
            Call Menu.Click(wParam)
            
        Case WM_GETMINMAXINFO
            If xMin <> 0 Or yMin <> 0 Or xMax <> 0 Or yMax <> 0 Then
                CopyMemory m_MinMax, ByVal lParam, Len(m_MinMax)
                If xMin <> 0 Then m_MinMax.ptMinTrackSize.x = xMin
                If yMin <> 0 Then m_MinMax.ptMinTrackSize.y = yMin
                If xMax <> 0 Then m_MinMax.ptMaxTrackSize.x = xMax
                If yMax <> 0 Then m_MinMax.ptMaxTrackSize.y = yMax
                CopyMemory ByVal lParam, m_MinMax, Len(m_MinMax)
                bHandled = True
            End If
    End Select
End Function


'====================================================================================
Public Function Events(nEvent As String, ParamArray Args() As Variant) As Variant
    On Error Resume Next
    If m_CodeObject Is Nothing Then Exit Function
    If Not m_onEvent Is Nothing Then nEvent = m_onEvent(Me, nEvent)
    If ExistsMember(m_CodeObject, nEvent) = False Then Exit Function
    Events = CBN(m_CodeObject, nEvent, m_CT, Args)
End Function


'====================================================================================
Public Property Get CodeObject() As Object
    Set CodeObject = m_CodeObject
End Property

Public Property Let CodeObject(ByVal value As Object)
    m_CT = IIF(IsJS(value), -1, VbFunc)
    Set m_CodeObject = value
End Property

Public Property Set CodeObject(ByVal value As Object)
    m_CT = IIF(IsJS(value), -1, VbFunc)
    Set m_CodeObject = value
End Property

Private Function IsJS(ByVal Obj As Object) As Boolean
    Dim tmp As String
    On Error Resume Next
    tmp = Obj:    IsJS = (tmp = "[object Object]")
End Function


'====================================================================================
Public Property Get onEvent() As Object
    Set onEvent = m_onEvent
End Property

Public Property Let onEvent(ByVal value As Object)
    Set m_onEvent = value
End Property

Public Property Set onEvent(ByVal value As Object)
    Set m_onEvent = value
End Property
