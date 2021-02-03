VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " About"
   ClientHeight    =   2256
   ClientLeft      =   2340
   ClientTop       =   1956
   ClientWidth     =   5628
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
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2256
   ScaleWidth      =   5628
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   5040
      Picture         =   "frmAbout.frx":0000
      ScaleHeight     =   384
      ScaleWidth      =   384
      TabIndex        =   5
      Top             =   600
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   5160
      Top             =   120
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4680
      TabIndex        =   0
      Top             =   1883
      Width           =   900
   End
   Begin VB.Image Image2 
      Height          =   1152
      Left            =   240
      MouseIcon       =   "frmAbout.frx":0152
      MousePointer    =   99  'Custom
      Picture         =   "frmAbout.frx":02A4
      ToolTipText     =   "Internet: LangMF.ru      Email: support@LangMF.ru"
      Top             =   240
      Width           =   1152
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright ©  LangMF"
      Height          =   210
      Left            =   1680
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   1920
      Width           =   1995
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      X1              =   0
      X2              =   5640
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   "Скрипт-движок позволяет создавать мощные скрипты, в то же время обладая простым языком и графическим интерфейсом."
      Height          =   690
      Left            =   1680
      TabIndex        =   1
      Top             =   1080
      Width           =   3915
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LangMF"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C86A25&
      Height          =   630
      Left            =   1680
      TabIndex        =   2
      Top             =   0
      Width           =   2055
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   0
      X2              =   5640
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1680
      TabIndex        =   3
      Top             =   720
      Width           =   3765
   End
   Begin VB.Shape Bullet 
      BorderColor     =   &H00000000&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   70
      Index           =   0
      Left            =   4800
      Top             =   120
      Visible         =   0   'False
      Width           =   200
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdOK_Click()
    ShellExecuteW 0, 0, StrPtr(GetAppPath + "help.chm"), 0, StrPtr("C:\"), SW_SHOWNORMAL
    Unload Me
End Sub

Private Sub Form_Load()
    Set Image2.MouseIcon = icoInet.Picture
    Set Label1.MouseIcon = icoInet.Picture
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    InitBullet
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Label1.ForeColor = QBColor(9) Then Label1.ForeColor = &H80000012
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
    Script_End
End Sub

Private Sub Image2_Click()
    GoWeb
End Sub

Private Sub Label1_Click()
    GoWeb
End Sub

Private Sub GoWeb()
    Call ShellExecuteW(Me.hWnd, 0, StrPtr("http://langmf.ru"), 0, StrPtr("C:\"), SW_SHOWNORMAL)
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Label1.ForeColor = QBColor(9)
End Sub

Private Sub lblTitle_Click()
    Timer1.Enabled = True
End Sub


Private Sub Timer1_Timer()
    Static gCnt As Long, gX As Integer, gY As Integer
    Dim a As Long, One As Boolean, cx As Long, cy As Long
    
    If gCnt Mod 10 = 0 Then
        Randomize Timer
        gX = Int(Rnd * 3 - 1) * Screen.TwipsPerPixelX
        gY = Int(Rnd * 3 - 1) * Screen.TwipsPerPixelY
    End If
    cx = Image2.Left + gX
    cy = Image2.Top + gY
    If cx < 0 Then cx = 0
    If cx > lblTitle.Left - Image2.Width Then cx = lblTitle.Left - Image2.Width
    If cy < 0 Then cy = 0
    If cy > Me.ScaleHeight - Image2.Height Then cy = Me.ScaleHeight - Image2.Height
    Image2.Move cx, cy

    
    For a = 0 To Bullet.Count - 1
        With Bullet(a)
            If .Visible Then
                .Move .Left + Val(.Tag)
                If .Left > Me.Width Then .Visible = False
            Else
                If Not One Then
                    If gCnt Mod 50 = 0 Then
                        One = True
                        .Move Image2.Left + Image2.Width - .Width, Image2.Top + Image2.Height / 2 - .Height / 2
                        .FillColor = RGB(GenRnd, GenRnd, GenRnd)
                        .Tag = Int(Rnd * 3 + 1) * Screen.TwipsPerPixelX
                        .Visible = True
                        Call PlaySoundW(StrPtr(GetWindowsPath + "\Media\ir_begin.wav"), 0, 1)
                    End If
                End If
            End If
        End With
    Next
    
    gCnt = gCnt + 1
End Sub

Private Function GenRnd() As Long
    Randomize Timer
    GenRnd = ((Rnd * 200 + 55) / 16) * 16
End Function

Private Sub InitBullet()
    Dim a As Long
    For a = 1 To 5:     Load Me.Bullet(a):      Next
End Sub
