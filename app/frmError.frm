VERSION 5.00
Begin VB.Form frmError 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4308
   ClientLeft      =   48
   ClientTop       =   468
   ClientWidth     =   6228
   Icon            =   "frmError.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4308
   ScaleWidth      =   6228
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Отправить отчет"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   1
      Top             =   3840
      Width           =   1692
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   252
      Left            =   120
      TabIndex        =   6
      Top             =   3960
      Width           =   4212
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   6270
      Y1              =   3745
      Y2              =   3745
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   972
      Left            =   120
      TabIndex        =   4
      Top             =   2640
      Width           =   5892
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   852
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   5892
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   492
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   5892
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   6270
      Y1              =   900
      Y2              =   900
   End
   Begin VB.Image Image1 
      Height          =   252
      Left            =   5640
      Top             =   240
      Width           =   252
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   732
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4932
   End
   Begin VB.Label lblFon 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2856
      Left            =   0
      TabIndex        =   5
      Top             =   900
      Width           =   6252
   End
End
Attribute VB_Name = "frmError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private txtError As String

Private Sub Form_Load()
    Dim sz As Long
    
    Me.Caption = GetFileName(Info.File)
    Label5.Caption = "LangMF Engine v." & App.Major & "." & App.Minor & "." & App.Revision

    If GetSystemMetrics(SM_CXICON) <= 40 Then sz = 32 Else sz = 64
    Set Image1.Picture = SYS.GDI.IcoToPic(LoadImageAsString(App.hInstance, 1, IMAGE_ICON, sz, sz, LR_SHARED))
    Image1.Move Me.ScaleWidth - Image1.Width * 2, lblFon.Top / 2 - Image1.Height / 2
    
    If GetSystemDefaultLangID = 1049 Then
        Label1.Caption = "В программе обнаружена ошибка." + vbCrLf + "Приложение будет закрыто." + vbCrLf + "Приносим извинения за неудобства."
    Else
        Label1.Caption = "Program has encountered a" + vbCrLf + "problem and needs to close." + vbCrLf + "We are sorry for the inconvenience."
        Command1.Caption = "Report send"
    End If
End Sub

Public Sub Display(ByVal Obj As clsActiveScript)
    Dim num As Long, cnx As Long, sErr As String, sType As String, sCode As String, txt() As String
    
    Const d = "    |    "
    
    With Obj.Error
        num = .Item("Line")
        cnx = .Item("Context")
        sErr = "Error = " & Hex$(.Item("Number")) & d & "Line = " & num & d & "Pos = " & .Item("Pos") & d & Obj.Name & " => " & IIF(Len(Obj.Tag), Obj.Tag, MDL(cnx).Name)
        sType = "Type -> " & .Item("Descr")
        sCode = "Code -> " & .Item("Code")
    End With

    If Len(Obj.Error.Item("Code")) = 0 And (cnx > 0 And cnx <= UBound(MDL)) Then
        If MDL(cnx).MFC = False And Len(MDL(cnx).Code) Then
            txt = Split(MDL(cnx).Code, vbCrLf)
            If UBound(txt) >= num Then sCode = sCode & Trim$(txt(num))
        End If
    End If

    txtError = "------Begin of Report-------" + vbCrLf + sErr + vbCrLf + sType + vbCrLf + sCode + vbCrLf + "Path -> " + Info.File + vbCrLf + "Version -> " + CStr(App.Major) + "." + CStr(App.Minor) + "." + CStr(App.Revision) + vbCrLf + "Windows -> " & WinVer.dwMajorVersion & "." & WinVer.dwMinorVersion & "." & WinVer.dwBuildNumber & " [" & Left$(WinVer.szCSDVersion, InStr(WinVer.szCSDVersion, vbNullChar) - 1) & "]" & vbCrLf & "-------End of Report--------"

    Label2.Caption = sErr
    Label3.Caption = sType
    Label4.Caption = sCode
    
    Me.Show vbModal
End Sub

Private Sub Command1_Click()
    Call ShellExecuteW(Me.hWnd, StrPtr("open"), StrPtr("mailto:" + EMailDevelop + "?subject=LangMF_Error_Report&body=" + SYS.Conv.EncodeUrl(txtError)), 0, 0, SW_SHOWNORMAL)
    Unload Me
End Sub


