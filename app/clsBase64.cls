VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBase64"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private EB64(63) As Byte, DB64(255) As Byte, m_Table As String


Private Sub Class_Initialize()
    Table = vbNullString
End Sub

Public Property Get Table() As String
    Table = m_Table
End Property

Public Property Let Table(ByVal value As String)
    Dim a As Long, b As Byte, sz As Long
    
    If LenB(value) Then m_Table = value Else m_Table = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
    sz = Len(m_Table):      If sz > 64 Then sz = 64
    
    For a = 0 To UBound(DB64)
        DB64(a) = 255              'skip char
    Next
    
    DB64(61) = 64                  'char of end '='
    
    For a = 0 To sz - 1
        b = Asc(Mid$(m_Table, a + 1, 1)):      EB64(a) = b:      DB64(b) = a
    Next
End Property

Public Function Encode(Buf() As Byte, Optional ByVal vCrLf As Long) As Byte()
    Dim a As Long, p As Long, i As Long, sz As Long, sz3 As Long, sz4 As Long, szM As Long
    Dim b1 As Byte, b2 As Byte, b3 As Byte, tmp() As Byte
    
    sz = ArraySize(Buf):      If sz = 0 Then Exit Function

    sz4 = sz \ 3:      sz3 = sz4 * 3:      sz4 = sz4 * 4:      szM = sz - sz3
    
    If vCrLf Then sz4 = sz4 + (sz4 \ (vCrLf * 4)) * 2
    If szM > 0 Then sz4 = sz4 + 4
    
    ReDim tmp(sz4 - 1)

    For a = 0 To sz3 - 1 Step 3
        b1 = Buf(a):      b2 = Buf(a + 1):      b3 = Buf(a + 2)

        tmp(p + 0) = EB64(b1 \ 4)
        tmp(p + 1) = EB64((b1 And &H3) * 16 Or (b2 \ 16))
        tmp(p + 2) = EB64((b2 And &HF) * 4 Or (b3 \ 64))
        tmp(p + 3) = EB64(b3 And &H3F)
        
        p = p + 4
        
        If vCrLf Then
            i = i + 1:    If i >= vCrLf Then tmp(p + 0) = 13:    tmp(p + 1) = 10:    p = p + 2:    i = 0
        End If
    Next

    If szM = 1 Then
        b1 = Buf(sz - 1)
        tmp(p + 0) = EB64(b1 \ 4)
        tmp(p + 1) = EB64((b1 And &H3) * 16)
        tmp(p + 2) = 61
        tmp(p + 3) = 61
        
    ElseIf szM = 2 Then
        b1 = Buf(sz - 2)
        b2 = Buf(sz - 1)
        tmp(p + 0) = EB64(b1 \ 4)
        tmp(p + 1) = EB64((b1 And &H3) * 16 Or (b2 \ 16))
        tmp(p + 2) = EB64((b2 And &HF) * 4)
        tmp(p + 3) = 61
    End If
    
    Encode = tmp
End Function

Public Function Decode(Buf() As Byte) As Byte()
    Dim a As Long, p As Long, i As Long, sz As Long, b1 As Byte, b2 As Byte, b3 As Byte, b4 As Byte, tmp() As Byte

    sz = ArraySize(Buf):      If sz Then ReDim tmp(sz - 1)
    
    For a = 0 To sz - 1
        b4 = DB64(Buf(a))
        
        If b4 < 65 Then
            i = i + 1
        
            If i = 1 Then
                b1 = b4
            ElseIf i = 2 Then
                b2 = b4
            ElseIf i = 3 Then
                b3 = b4
            Else
                tmp(p + 0) = ((b1 * 4) Or (b2 \ 16))
                tmp(p + 1) = (((b2 And &HF) * 16) Or (b3 \ 4))
                tmp(p + 2) = (((b3 And &H3) * 64) Or b4)
    
                i = 0:    p = p + 3 - (b3 \ 64) - (b4 \ 64)
            End If
        End If
    Next
    
    If p Then ReDim Preserve tmp(p - 1):   Decode = tmp
End Function
