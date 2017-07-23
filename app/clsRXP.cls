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


Public Function CreateNew() As Object
    Set CreateNew = New clsRXP
End Function

Public Function Replace(ByVal rTxt As String, Optional ByVal rPat As String = "", Optional ByVal rRep As String = "", Optional ByVal fi As Boolean = True, Optional ByVal fg As Boolean = True) As String
    With Obj
        .Global = fg:       .IgnoreCase = fi:       If LenB(rPat) Then .Pattern = rPat
        Replace = .Replace(rTxt, rRep)
    End With
End Function

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
    Call Execute(rTxt, rPat, fi, fg)
    value = Matches(rNum, numMts)
    If IsEmpty(value) Then Eval = rDef Else Eval = value
End Function

Private Sub Class_Initialize()
    Set Obj = New RegExp
End Sub
