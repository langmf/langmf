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

Const def_Table As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"

Private TB64(63) As Byte, iTB64(255) As Byte, m_Table As String


Private Sub Class_Initialize()
    Table = vbNullString
End Sub

Public Property Get Table() As String
    Table = m_Table
End Property

Public Property Let Table(ByVal value As String)
    Dim a As Long
    
    If LenB(value) = 0 Then
        m_Table = def_Table
    ElseIf Len(value) = 64 Then
        m_Table = value
    Else
        Exit Property
    End If
    
    For a = 0 To UBound(iTB64)
        iTB64(a) = 64           'skip char
    Next
    
    iTB64(61) = 255             'char of end '='
    iTB64(126) = 254            'char of extra '~'
    
    For a = 0 To Len(m_Table) - 1
        TB64(a) = Asc(Mid$(m_Table, a + 1, 1))
        iTB64(TB64(a)) = a
    Next
End Property

Public Function Encode(Buf() As Byte, Optional ByVal isCrLf As Boolean = True) As Byte()
    Dim a As Long, Pos As Long, cnt As Long, Size As Long, size3 As Long, size4 As Long, szMod As Long
    Dim Bt1 As Byte, Bt2 As Byte, Bt3 As Byte, temp() As Byte
    
    Size = m_ArraySize(Buf)
    
    If Size = 0 Then Exit Function

    size4 = Size \ 3
    size3 = size4 * 3
    size4 = size4 * 4
    szMod = Size - size3
    
    If isCrLf Then size4 = size4 + (size4 \ 76) * 2
    If szMod > 0 Then size4 = size4 + 4
    
    ReDim temp(size4 - 1)

    For a = 0 To size3 - 1 Step 3
        Bt1 = Buf(a):   Bt2 = Buf(a + 1):   Bt3 = Buf(a + 2)

        temp(Pos) = TB64(Bt1 \ 4)
        temp(Pos + 1) = TB64((Bt1 And &H3) * 16 Or (Bt2 \ 16))
        temp(Pos + 2) = TB64((Bt2 And &HF) * 4 Or (Bt3 \ 64))
        temp(Pos + 3) = TB64(Bt3 And &H3F)
        
        Pos = Pos + 4
        
        If isCrLf Then
            cnt = cnt + 1
            If cnt >= 19 Then
                temp(Pos) = 13
                temp(Pos + 1) = 10
                Pos = Pos + 2
                cnt = 0
            End If
        End If
    Next

    Select Case szMod
        Case 1
            Bt1 = Buf(Size - 1)
            
            temp(Pos) = TB64(Bt1 \ 4)
            temp(Pos + 1) = TB64((Bt1 And &H3) * 16)
            temp(Pos + 2) = 61
            temp(Pos + 3) = 61

        Case 2
            Bt1 = Buf(Size - 2)
            Bt2 = Buf(Size - 1)
    
            temp(Pos) = TB64(Bt1 \ 4)
            temp(Pos + 1) = TB64((Bt1 And &H3) * 16 Or (Bt2 \ 16))
            temp(Pos + 2) = TB64((Bt2 And &HF) * 4)
            temp(Pos + 3) = 61
    End Select
    
    Encode = temp
End Function

Public Function Decode(Buf() As Byte, Optional Extra As Variant) As Byte()
    Dim a As Long, Pos As Long, Size As Long, Bt1 As Byte, Bt2 As Byte, Bt3 As Byte, Bt4 As Byte
    Dim isExt As Boolean, posExt As Long, udsExt As Long, bufExt() As Byte, temp() As Byte

    Size = m_ArraySize(Buf)
    
    If Size < 4 Then Exit Function

    posExt = -1
    udsExt = posExt
    
    ReDim temp(Size - 1)
    
    For a = 0 To Size - 4 Step 4
        Bt1 = iTB64(Buf(a))
        
        If Bt1 < 64 And isExt = False Then
            Bt2 = iTB64(Buf(a + 1)):   Bt3 = iTB64(Buf(a + 2)):   Bt4 = iTB64(Buf(a + 3))

            temp(Pos) = ((Bt1 * 4) Or (Bt2 \ 16))
            temp(Pos + 1) = (((Bt2 And &HF) * 16) Or (Bt3 \ 4))
            temp(Pos + 2) = (((Bt3 And &H3) * 64) Or Bt4)

            If Bt3 = 255 Then
                Pos = Pos + 1
            ElseIf Bt4 = 255 Then
                Pos = Pos + 2
            Else
                Pos = Pos + 3
            End If
        Else
            If Bt1 = 254 Then
                isExt = Not isExt
            Else
                If isExt Then
                    posExt = posExt + 1
                    If posExt > udsExt Then
                        udsExt = udsExt + 10000
                        ReDim Preserve bufExt(udsExt)
                    End If
                    bufExt(posExt) = Buf(a)
                End If
            End If
            a = a - 3
        End If
    Next
    
    If posExt > -1 Then ReDim Preserve bufExt(posExt)
    Extra = bufExt
    
    ReDim Preserve temp(Pos - 1)
    
    Decode = temp
End Function
