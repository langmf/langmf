VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LangMF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private oleInstance As Long, isFail As Boolean


Public Sub Command(ByVal cmdLine As String)
    Dim RX As New clsRXP
    
    If isFail Then Exit Sub
    
    With Info
        .IsCmd = Len(cmdLine)

        If .IsExe Then
            .Arg = cmdLine
            .File = GetAppPath(True)
        Else
            If RX.Execute(cmdLine, "^""([^""]+)"" *(.*)").Count Then
                .Arg = Trim$(RX.Mts(0).SubMatches(1))
                .File = Trim$(RX.Mts(0).SubMatches(0))
            Else
                .Arg = vbNullString
                .File = cmdLine
            End If
        
            If LCase$(.Arg) = "/make" Then Script_Init:   MakeMF .File:   Script_End:   Exit Sub
        End If
        
        Call Code_Run(.File)
    End With
End Sub

Public Property Get Script(Optional ByVal Index As Long = -1) As Object
    If isFail Then Exit Property
    
    If CAS Is Nothing Then Script_Init
    
    If Index = -1 Then Set Script = CAS:   Exit Property
    
    If Not VerifyLongRange(Index, 1, frmScript.CScript.Count) Then Exit Property
    
    Set Script = frmScript.CScript(CStr(Index))
End Property

Public Property Get Instance() As Object
    Dim a As Long
    
    On Error Resume Next
    
    Shell """" + GetAppPath(True) + """ " + mf_New
    
    Do
        a = a + 1
        m_Wait
        Set Instance = GetObject(, "Atomix.LangMF")
    Loop Until a > 300 Or (Not Instance Is Nothing)
    
    Instance.ROT False
End Property
    
Public Function ROT(Optional ByVal value As Boolean = True) As Long
Attribute ROT.VB_Description = "Running Object Table"
    Dim mGuid As UUID
    
    If value Then
        oleInstance = 0
        If CLSIDFromProgID(StrPtr("Atomix.LangMF"), mGuid) = 0 Then ROT = RegisterActiveObject(Me, mGuid, ACTIVEOBJECT_STRONG, oleInstance)
    Else
        If oleInstance <> 0 Then ROT = RevokeActiveObject(oleInstance, 0)
        CoDisconnectObject Me, 0
    End If
End Function

Public Property Get State() As Boolean
    If Not isFail Then State = CBool(Not CAS Is Nothing)
End Property

Public Sub Reset()
    If Not isFail Then Script_End
End Sub


Private Sub Class_Initialize()
    If curRunMF Then isFail = True:   Exit Sub
    If LMF Is Nothing Then Set LMF = Me
End Sub

Private Sub Class_Terminate()
    If Not isFail Then Script_End:    Set LMF = Nothing
End Sub
