VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGLOB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Obj As Object

Public Function Pipe(Optional ByVal nGlob As String = "lmf_glob") As Boolean
    On Error Resume Next

    For Each Obj In VBA.CreateObject("Shell.Application").Windows
        If LCase$(nGlob) = LCase$(TrimNull(Obj.StatusText)) Then Pipe = True:    Exit Function
    Next
    
    Set Obj = VBA.GetObject("new:{C08AFD90-F2A1-11D1-8455-00A0C91F3880}")
    If Not Obj Is Nothing Then Pipe = True:   Obj.StatusText = nGlob
End Function

'Default Member
Public Property Get value(ByVal nProp As String)
    If IsObject(Obj.GetProperty(nProp)) Then Set value = Obj.GetProperty(nProp) Else value = Obj.GetProperty(nProp)
End Property

Public Property Let value(ByVal nProp As String, v As Variant)
Attribute value.VB_UserMemId = 0
    Obj.PutProperty nProp, v
End Property


