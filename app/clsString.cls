VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function Asc(ByVal txt As String) As Long
    Asc = VBA.Asc(txt)
End Function

Public Function AscW(ByVal txt As String) As Long
    AscW = VBA.AscW(txt)
End Function

Public Function Chr(ByVal numChr As Long) As String
    Chr = VBA.Chr$(numChr)
End Function

Public Function ChrW(ByVal numChr As Long) As String
    ChrW = VBA.ChrW$(numChr)
End Function

Public Property Get Mid(txt As Variant, ByVal mStart As Long, Optional ByVal mLen As Long = 0) As String
    Call Mid_Ext(txt, mStart, mLen)
    If mLen = 0 Then Mid = VBA.Mid$(txt, mStart) Else Mid = VBA.Mid$(txt, mStart, mLen)
End Property

Public Property Let Mid(txt As Variant, ByVal mStart As Long, Optional ByVal mLen As Long = 0, ByVal value As String)
    Call Mid_Ext(txt, mStart, mLen)
    If mLen = 0 Then Mid(txt, mStart) = value Else Mid(txt, mStart, mLen) = value
End Property

Public Function Left(ByVal txt As String, ByVal mLen As Long) As String
    Call LR_Ext(txt, mLen)
    Left = VBA.Left$(txt, mLen)
End Function

Public Function Right(ByVal txt As String, ByVal mLen As Long) As String
    Call LR_Ext(txt, mLen)
    Right = VBA.Right$(txt, mLen)
End Function

Public Function Trim(ByVal txt As String) As String
    Trim = VBA.Trim$(txt)
End Function

Public Function RTrim(ByVal txt As String) As String
    RTrim = VBA.RTrim$(txt)
End Function

Public Function LTrim(ByVal txt As String) As String
    LTrim = VBA.LTrim$(txt)
End Function

Public Function TrimNull(ByVal txt As String) As String
    TrimNull = modCommon.TrimNull(txt)
End Function

Public Function Rpt(ByVal txt As String, ByVal cnt As Long) As String
    Dim i As Long, j As Long, sz As Long, p As Long, b() As Byte, o() As Byte
    
    If Len(txt) = 0 Or cnt = 0 Then Exit Function

    b = txt:   sz = Len(txt) * 2:   If cnt > 0 Then cnt = sz * cnt - 1 Else cnt = -2 * cnt - 1
    
    ReDim o(cnt)

    For p = 0 To cnt
        o(p) = b(i):   i = i + 1:   If i >= sz Then i = 0
    Next

    Rpt = o
End Function

Public Function Pad(ByVal txt As String, ByVal cnt As Long, Optional ByVal pad_str As String = " ", Optional ByVal pad_dir As Long = 2) As String
    Dim sl As Long, sz As Long
    
    If Len(txt) = 0 Or cnt = 0 Then Pad = txt:   Exit Function
    
    If cnt < 0 Then
        cnt = -cnt
        If Len(txt) >= cnt Then Pad = VBA.Left$(txt, cnt):   Exit Function
    End If
    
    If Len(txt) >= cnt Then Pad = txt:   Exit Function
    
    sz = cnt - Len(txt)

    If pad_dir = 2 Then Pad = txt + Rpt(pad_str, -sz):   Exit Function
    If pad_dir = 1 Then Pad = Rpt(pad_str, -sz) + txt:   Exit Function
    
    sl = Fix(sz / 2):   sz = sz - sl:   Pad = Rpt(pad_str, -sl) + txt + Rpt(pad_str, -sz)
End Function

Public Function Sort(txt As Variant, Optional ByVal bAscending As Boolean = True, Optional ByVal bCaseSensitive As Boolean = False) As Variant()
    Dim tmpStr() As String, a As Long
    
    If m_ArraySize(txt) = 0 Then Exit Function
    
    ReDim tmpStr(UBound(txt))
    
    For a = 0 To UBound(txt)
        tmpStr(a) = CStr(txt(a))
    Next
    
    QuickSortStringsStart tmpStr, bAscending, bCaseSensitive
    
    For a = 0 To UBound(txt)
        txt(a) = CVar(tmpStr(a))
    Next
    
    Sort = txt
End Function

Public Function LSplit(ByVal txt As String, Optional ByVal sz As Long = 80) As Variant()
    Dim a As Long, i As Long, tmpStr() As Variant
    
    ReDim tmpStr((Len(txt) - 1) \ sz)
    
    For a = 1 To Len(txt) Step sz
        tmpStr(i) = Mid$(txt, a, sz)
        i = i + 1
    Next
    
    LSplit = tmpStr
End Function


Private Sub Mid_Ext(txt As Variant, mStart As Long, mLen As Long)
    If mStart < 0 Then mStart = Len(txt) + mStart + 1
    If mStart <= 0 Then mStart = 1
    If mLen < 0 Then mLen = Len(txt) - mStart + mLen + 1
    If mLen < 0 Then mLen = 0
End Sub

Private Sub LR_Ext(txt As String, mLen As Long)
    If mLen < 0 Then mLen = Len(txt) + mLen
    If mLen < 0 Then mLen = 0
End Sub
