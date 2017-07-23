VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCron"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private RXP As RegExp, st As SYSTEMTIME, hWnd As Long, oldSecond As Integer

Public Jobs As Object


Public Function Add(Optional ByVal value As String, Optional Action As Variant, Optional ByVal ID As String) As Boolean
    Add = Jobs.Add(Array(Parse(value), value, Action, ID), ID)
End Function

Public Function Remove(Optional ByVal ID As Variant) As Boolean
    If IsMissing(ID) Then
        Jobs.Init
        Remove = True
    Else
        Remove = Jobs.Remove(ID)
    End If
End Function

Public Sub Timer()
    Dim a As Long, itm As Variant, v As Variant
    
    GetTime
    
    With st
        If oldSecond <> .wSecond Then
            For Each itm In Jobs.Items
                If Check(itm(0)) Then
                    v = itm(2)
                    If IsArray(v) Then
                        For a = 0 To UBound(v)
                            DoFunc v(a)
                        Next
                    Else
                        DoFunc v
                    End If
                End If
            Next
        End If
        
        oldSecond = .wSecond
    End With
End Sub

Public Function Parse(ByVal value As String) As Variant
    Dim txt() As String, txt2() As String, out As Variant, tmp As Variant, mt As Match, Mts As MatchCollection
    Dim a As Long, b As Long, fb As Long, fe As Long, fs As Long, cnt As Long
    
    GetTime
    
    ReDim out(6)
    
    With RXP
        .Pattern = "[ \t]+"
        value = .Replace(value, " ")
                
        .Pattern = "[a-zа-я_]+"
        Set Mts = .Execute(value)
        For Each mt In Mts
            a = GetMonthOrWeekDay(mt.value)
            If a > 0 Then value = Replace$(value, mt.value, a)
        Next
                
        txt = Split(value, " ")
        cnt = UBound(txt)
        
        For a = 0 To 6
            If a > cnt Then
                out(a) = Array(Array(0, 10000, 1))
            Else
                If txt(a) = "?" Then
                    With st
                        out(a) = Array(Array(.wSecond, .wMinute, .wHour, .wDay, .wDayOfWeek, .wMonth, .wYear)(a))
                    End With
                Else
                    txt2 = Split(txt(a), ",")
                    
                    ReDim tmp(UBound(txt2))
                
                    For b = 0 To UBound(txt2)
                        .Pattern = "((\d+)\-(\d+)|\*)(\/(\d+))?"
                        Set Mts = .Execute(txt2(b))
                        
                        If Mts.Count Then
                            fs = Val(Mts(0).SubMatches(4))
                            If fs < 1 Then fs = 1
                                
                            If Mts(0).SubMatches(0) = "*" Then
                                fb = Array(0, 0, 0, 1, 1, 1, 0)(a)
                                fe = 10000
                            Else
                                fb = Val(Mts(0).SubMatches(1))
                                fe = Val(Mts(0).SubMatches(2))
                                If fe = 0 Then fe = 10000
                            End If
                            
                            tmp(b) = Array(fb, fe, fs)
                        Else
                            tmp(b) = Val(txt2(b))
                        End If
                    Next
                    
                    out(a) = tmp
                End If
            End If
        Next
    End With
    
    Parse = out
End Function

Public Sub StartTimer(Optional ByVal Interval As Long = 500)
    Call SetTimer(hWnd, 30000, Interval, AddressOf Timer_Cron)
End Sub

Public Sub StopTimer()
    Call KillTimer(hWnd, 30000)
End Sub

Private Sub GetTime()
    Call GetLocalTime(st)
    If st.wDayOfWeek = 0 Then st.wDayOfWeek = 7
End Sub

Private Sub DoFunc(ByVal value As String)
    On Error Resume Next
    
    Select Case Left$(value, 2)
        Case "R!":    Remove Mid$(value, 3)
        Case "E!":    CAS.ExecuteStatement Mid$(value, 3)
        Case Else:    CAS.Eval value
    End Select
End Sub

Private Function Check(Prm As Variant) As Boolean
    Dim a As Long, v As Variant
    
    On Error Resume Next
    
    With st
        v = Array(.wSecond, .wMinute, .wHour, .wDay, .wDayOfWeek, .wMonth, .wYear)
    End With
                
    For a = 0 To UBound(Prm)
        If Not CheckItem(Prm(a), v(a)) Then Exit Function
    Next
    
    Check = True
End Function

Private Function CheckItem(Prm As Variant, ByVal value As Long) As Boolean
    Dim a As Long, v As Variant
    
    On Error Resume Next
    
    CheckItem = True
    
    For a = 0 To UBound(Prm)
        v = Prm(a)
        
        If IsArray(v) Then
            If value >= v(0) And value <= v(1) Then If (value - v(0)) Mod v(2) = 0 Then Exit Function
        Else
            If value = v Then Exit Function
        End If
    Next
    
    CheckItem = False
End Function

Private Function GetMonthOrWeekDay(value As String) As Long
    Select Case LCase$(value)
        Case "jan", "янв", "mon", "пн":     GetMonthOrWeekDay = 1
        Case "feb", "фев", "tue", "вт":     GetMonthOrWeekDay = 2
        Case "mar", "мар", "wed", "ср":     GetMonthOrWeekDay = 3
        Case "apr", "апр", "thu", "чт":     GetMonthOrWeekDay = 4
        Case "may", "май", "fri", "пт":     GetMonthOrWeekDay = 5
        Case "jun", "июн", "sat", "сб":     GetMonthOrWeekDay = 6
        Case "jul", "июл", "sun", "вс":     GetMonthOrWeekDay = 7
        Case "aug", "авг":                  GetMonthOrWeekDay = 8
        Case "sep", "сен":                  GetMonthOrWeekDay = 9
        Case "oct", "окт":                  GetMonthOrWeekDay = 10
        Case "nov", "ноя":                  GetMonthOrWeekDay = 11
        Case "dec", "дек":                  GetMonthOrWeekDay = 12
    End Select
End Function

Private Sub Class_Initialize()
    Set Jobs = New clsHash
    Set RXP = New RegExp
    RXP.Global = True
    RXP.IgnoreCase = True
    hWnd = frmScript.hWnd
    StartTimer
End Sub

Private Sub Class_Terminate()
    StopTimer
End Sub
