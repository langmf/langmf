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

Public Function Test(ByVal txt As String, Optional ByVal pat As String, Optional ByVal fi As Boolean = True, Optional ByVal fg As Boolean = True, Optional ByVal fm As Boolean = True) As Boolean
    With Obj
        .Global = fg:     .IgnoreCase = fi:     .MultiLine = fm:     If LenB(pat) Then .Pattern = pat
        Set Mts = .Execute(txt)
    End With
    Test = (Mts.Count > 0)
End Function

Public Function Execute(ByVal txt As String, Optional ByVal pat As String, Optional ByVal fi As Boolean = True, Optional ByVal fg As Boolean = True, Optional ByVal fm As Boolean = True) As MatchCollection
    With Obj
        .Global = fg:     .IgnoreCase = fi:     .MultiLine = fm:     If LenB(pat) Then .Pattern = pat
        Set Mts = .Execute(txt)
    End With
    Set Execute = Mts
End Function

Public Function Replace(ByVal txt As String, Optional ByVal pat As String, Optional Rep As Variant = "", Optional ByVal fi As Boolean = True, Optional ByVal fg As Boolean = True, Optional ByVal fm As Boolean = True) As Variant
    Dim i As Long, c As Long, t As Long, p As Long, res As String, v As Variant, m As Match, h As New clsHash
    
    With Obj
        .Global = fg:     .IgnoreCase = fi:     .MultiLine = fm:     If LenB(pat) Then .Pattern = pat
    End With
    
    t = VariantType(Rep)

    If t = vbString Then
        Replace = Obj.Replace(txt, Rep)
    
    ElseIf t = vbObject Then
        Set Mts = Obj.Execute(txt):            c = GetFunc(Rep)(1):           p = 1
        
        For i = 0 To Mts.Count - 1
            Set m = Mts(i):     v = Empty
            
            If c = 1 Then v = Rep(m)
            If c = 2 Then v = Rep(m, i)
            If c = 3 Then v = Rep(m, i, h)

            If Not IsEmpty(v) And c <> 3 Then res = res & Mid$(txt, p, m.FirstIndex - p + 1) & v:      p = m.FirstIndex + m.Length + 1
        Next
        
        If c = 3 Then Set Replace = h Else Replace = res & Mid$(txt, p)
    End If
End Function

Public Function Matches(Optional ByVal curSBM As Long, Optional ByVal curMTS As Long) As Variant
    Dim cntMTS As Long, cntSBM As Long, numSBM As Long
    
    If Mts Is Nothing Then Exit Function
    If Mts.Count = 0 Or curMTS >= Mts.Count Then Exit Function
    If curMTS < 0 Then curMTS = 0:   cntMTS = Mts.Count Else cntMTS = 1
    
    For curMTS = curMTS To curMTS + cntMTS - 1
        With Mts.Item(curMTS)
            numSBM = curSBM:      If numSBM < 0 Then numSBM = 0:   cntSBM = .SubMatches.Count Else cntSBM = 1
            For numSBM = numSBM To numSBM + cntSBM - 1
                If numSBM < .SubMatches.Count Then Matches = Matches & .SubMatches(numSBM)
            Next
        End With
    Next
End Function

Public Function Eval(ByVal txt As String, Optional ByVal pat As String, Optional ByVal def As String, Optional ByVal curSBM As Long, Optional ByVal curMTS As Long, Optional ByVal fi As Boolean = True, Optional ByVal fg As Boolean = True, Optional ByVal fm As Boolean = True) As String
    Dim value As Variant
    
    Call Execute(txt, pat, fi, fg, fm):       value = Matches(curSBM, curMTS)
    
    If IsEmpty(value) Then Eval = def Else Eval = value
End Function

Private Sub Class_Initialize()
    Set Obj = New RegExp
End Sub
