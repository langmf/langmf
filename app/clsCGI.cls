VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCGI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private arrGIV() As String

Public Files As Object


Public Function IsStdOut() As Boolean
    If GetStdHandle(-11) <> 65537 Then IsStdOut = True
End Function

Public Function GetStdPipe(ByVal mPipe As Long) As Long
    GetStdPipe = GetStdHandle(mPipe)
End Function

Public Function SetStdPipe(ByVal mPipe As Long, ByVal mHandle As Long) As Long
    SetStdPipe = SetStdHandle(mPipe, mHandle)
End Function

Public Function StdIn(txt As Variant, Optional ByVal numPipe As Long) As Long
    Dim sBuf As String, sz As Long
    
    If numPipe = 0 Then numPipe = GetStdHandle(STD_INPUT_HANDLE)

    sz = GetFileSize(numPipe, ByVal 0&)
    If sz > 0 Then
        sBuf = String$(sz, " ")
        StdIn = ReadFile(numPipe, ByVal sBuf, sz, sz, 0)
    End If
    txt = sBuf
End Function

Public Function StdOut(ByVal txt As String, Optional ByVal numPipe As Long) As Long
    Dim nb As Long

    If numPipe = 0 Then numPipe = GetStdHandle(STD_OUTPUT_HANDLE)
    
    StdOut = WriteFile(numPipe, ByVal txt, Len(txt), nb, 0)
End Function

Public Function Inp(Optional ByVal ID As Variant) As Variant
    Dim tmp As Variant, Col As New clsHash, txt() As String
    
    If IsMissing(ID) Or IsEmpty(ID) Then
        For Each tmp In arrGIV
            txt = Split(tmp, "=")
            txt(1) = SYS.Conv.DecodeUrl(txt(1))
            Col.Add txt(1), txt(0)
        Next
        Set Inp = Col
    Else
        Set Col = Inp()
        Inp = Col(ID)
    End If
End Function

Private Sub GetInputVars()
    Dim cl As Long, br As Long, sBuf As String, StdIn As Long, lb As Long, txt As String

    Select Case UCase$(GEV("REQUEST_METHOD"))
        Case "POST"
            
            cl = Val(GEV("CONTENT_LENGTH"))
            lb = 8192

            If cl > 0 Then
                StdIn = GetStdHandle(STD_INPUT_HANDLE)
                Do
                    sBuf = String$(lb, 0)
                    br = lb
                    If ReadFile(StdIn, ByVal sBuf, lb, br, 0) Then
                        txt = txt & Left$(sBuf, br)
                        If Len(txt) >= cl Then Exit Do
                    End If
                Loop
            End If
            
            If InStr(LCase$(GEV("CONTENT_TYPE")), "multipart/form-data") > 0 Then Upload txt
            
            
        Case "GET"
            txt = GEV("QUERY_STRING")
    End Select
    
    arrGIV = Split(txt, "&")
End Sub

Private Sub Upload(txt As String)
    Dim hdr1 As String, hdr2 As String, hdr3 As String, Boundary As String
    Dim txtData As String, txtName As String, txtFile As String, txtType As String
    Dim a As Long, arrValue() As String, REG1 As RegExp, Mts As MatchCollection
    
    Set REG1 = New RegExp
    
    REG1.Global = True
    REG1.IgnoreCase = True
    
    Boundary = GEV("CONTENT_TYPE")
    
    REG1.Pattern = "boundary=([\-]+[a-z0-9_]+)"
    Set Mts = REG1.Execute(Boundary)
    Boundary = "--" & Mts(0).SubMatches(0)
    
    arrValue = Split(txt, Boundary, , vbTextCompare)
    
    txt = ""
    
    For a = 1 To UBound(arrValue) - 1
        REG1.Pattern = "^\r\n(.+)\r\n(.*)\r\n(.*)\r\n"
        Set Mts = REG1.Execute(arrValue(a))
        
        txtData = REG1.Replace(arrValue(a), "")
        txtData = Left$(txtData, Len(txtData) - 2)

        hdr1 = Mts(0).SubMatches(0)
        hdr2 = Mts(0).SubMatches(1)
        hdr3 = Mts(0).SubMatches(2)
        
        REG1.Pattern = "name=\""([^\""]+)\"""
        Set Mts = REG1.Execute(hdr1)
        If Mts.Count > 0 Then txtName = Mts(0).SubMatches(0) Else txtName = ""
        
        REG1.Pattern = "filename=\""([^\""]+)\"""
        Set Mts = REG1.Execute(hdr1)
        If Mts.Count > 0 Then txtFile = Mts(0).SubMatches(0) Else txtFile = ""
        
        REG1.Pattern = "Content\-Type: (.+)"
        Set Mts = REG1.Execute(hdr2)
        If Mts.Count > 0 Then txtType = Mts(0).SubMatches(0) Else txtType = ""
        
        If LenB(txtFile) Then
            With Files.Child(CStr(a))
                .Add txtName, "ID"
                .Add txtFile, "Name"
                .Add txtType, "Type"
                .Add txtData, "Data"
            End With
        Else
            txt = txt & txtName & "=" & hdr3 & "&"
        End If
    Next
    
    If LenB(txt) Then txt = Left$(txt, Len(txt) - 1)
End Sub

Private Sub Class_Initialize()
    Set Files = New clsHash
    GetInputVars
End Sub
