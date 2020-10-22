VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRXP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Obj As RegExp, Mts As MatchCollection


Public Property Get CreateNew() As Object
    Set CreateNew = New clsRXP
End Property

Public Function Test(ByVal rTxt As String, Optional ByVal rPat As String = "", Optional ByVal fi As Boolean = True, Optional ByVal fg As Boolean = True) As Boolean
    With Obj
        .Global = fg:       .IgnoreCase = fi:       If LenB(rPat) Then .Pattern = rPat
        Set Mts = .Execute(rTxt)
    End With
    Test = (Mts.Count > 0)
End Function

Public Function Execute(ByVal rTxt As String, Optional ByVal rPat As String = "", Optional ByVal fi As Boolean = True, Optional ByVal fg As Boolean = True) As MatchCollection
    With Obj
        .Global = fg:       .IgnoreCase = fi:       If LenB(rPat) Then .Pattern = rPat
        Set Mts = .Execute(rTxt)
    End With
    Set Execute = Mts
End Function

Public Function Replace(ByVal rTxt As String, Optional ByVal rPat As String = "", Optional rRep As Variant = "", Optional ByVal fi As Boolean = True, Optional ByVal fg As Boolean = True) As Variant
    Dim i As Long, c As Long, t As Long, p As Long, txt As String, v As Variant, m As Match, h As New clsHash
    
    t = VariantType(rRep):      Obj.Global = fg:       Obj.IgnoreCase = fi:       If LenB(rPat) Then Obj.Pattern = rPat

    If t = vbString Then
        Replace = Obj.Replace(rTxt, rRep)
    
    ElseIf t = vbObject Then
        Set Mts = Obj.Execute(rTxt):            c = GetFunc(rRep)(1):           p = 1
        
        For i = 0 To Mts.Count - 1
            Set m = Mts(i):     v = Empty
            
            If c = 1 Then v = rRep(m)
            If c = 2 Then v = rRep(m, i)
            If c = 3 Then v = rRep(m, i, h)

            If Not IsEmpty(v) And c <> 3 Then txt = txt & Mid$(rTxt, p, m.FirstIndex - p + 1) & v:      p = m.FirstIndex + m.Length + 1
        Next
        
        If c = 3 Then Set Replace = h Else Replace = txt + Mid$(rTxt, p)
    End If
End Function

Public Function Matches(Optional ByVal rNum As Long = 0, Optional ByVal numMts As Long = 0) As Variant
    Dim cntMts As Long, cntSbm As Long, numSbm As Long
    
    If Mts Is Nothing Then Exit Function
    If Mts.Count = 0 Or numMts >= Mts.Count Then Exit Function
    If numMts < 0 Then numMts = 0:   cntMts = Mts.Count Else cntMts = 1
    
    For numMts = numMts To numMts + cntMts - 1
        With Mts.Item(numMts)
            numSbm = rNum:      If numSbm < 0 Then numSbm = 0:   cntSbm = .SubMatches.Count Else cntSbm = 1
            For numSbm = numSbm To numSbm + cntSbm - 1
                If numSbm < .SubMatches.Count Then Matches = Matches & .SubMatches(numSbm)
            Next
        End With
    Next
End Function

Public Function Eval(ByVal rTxt As String, Optional ByVal rPat As String = "", Optional ByVal rDef As String = "", Optional ByVal fi As Boolean = True, Optional ByVal fg As Boolean = True, Optional ByVal rNum As Long = 0, Optional ByVal numMts As Long = 0) As String
    Dim value As Variant
    
    Call Execute(rTxt, rPat, fi, fg):       value = Matches(rNum, numMts)
    
    If IsEmpty(value) Then Eval = rDef Else Eval = value
End Function

Private Sub Class_Initialize()
    Set Obj = New RegExp
End Sub
