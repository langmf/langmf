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

Public GLOB As Object

Public Function Pipe(Optional ByVal nGlob As String = "lmf_glob") As Boolean
    On Error Resume Next

    For Each GLOB In CreateObject("Shell.Application").Windows
        If LCase$(nGlob) = LCase$(TrimNull(GLOB.StatusText)) Then Pipe = True:    Exit Function
    Next
    
    Set GLOB = GetObject("new:{C08AFD90-F2A1-11D1-8455-00A0C91F3880}")
    If Not GLOB Is Nothing Then Pipe = True:   GLOB.StatusText = nGlob
End Function

'Default Member
Public Property Get value(ByVal nProp As String)
    If IsObject(GLOB.GetProperty(nProp)) Then Set value = GLOB.GetProperty(nProp) Else value = GLOB.GetProperty(nProp)
End Property

Public Property Let value(ByVal nProp As String, v As Variant)
Attribute value.VB_UserMemId = 0
    GLOB.PutProperty nProp, v
End Property


