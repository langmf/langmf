VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4812
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   5076
   LinkTopic       =   "Form1"
   ScaleHeight     =   4812
   ScaleWidth      =   5076
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "CallBack"
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   4320
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   3975
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   120
      Width           =   4815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LMF As Object, LMF2 As Object

Private Sub Form_Load()
    Dim rc As String
    
    Set LMF = CreateObject("Atomix.LangMF")
    
    Set LMF.Script.Parent = Me
    
    LMF.Script.AddObject "form", Me
    
    rc = GetSetting("Example", "Save", "text")
    If rc <> "" Then Text1.Text = rc
    
    Set LMF2 = CreateObject("Atomix.LangMF")
    LMF2.Command App.Path + "\..\Alpha Blend\demo.mf"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    LMF.Reset
    LMF2.Reset
    Set LMF = Nothing
    Set LMF2 = Nothing
    SaveSetting "Example", "Save", "text", Text1.Text
End Sub

Private Sub Command1_Click()
    On Error Resume Next
    LMF.Command Text1.Text
    MsgBox "Result: " & LMF.Script.CodeObject.Test(5, 3)
End Sub

Public Sub MY()
    MsgBox "Call MY is complete!"
End Sub

Public Sub ActiveScript_Error(ByVal Obj As Object)
    MsgBox vbCrLf & "Error = " & Hex$(Obj.Error.Item("Number")) & "    |    Line = " & Obj.Error.Item("Line") & _
           "    |    Pos = " & Obj.Error.Item("Pos") & "    |    " & Obj.Name & " => " & Obj.Tag & vbCrLf & _
           "Type -> " & Obj.Error.Item("Descr") & vbCrLf & "Code -> " & Obj.Error.Item("Code") & vbCrLf, , "Error"
End Sub

