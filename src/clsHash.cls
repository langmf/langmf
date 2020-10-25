VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const DEFAULT_SizeHash     As Long = 991   'простое число
Const DEFAULT_SizeList     As Long = 256
Const DEFAULT_SizeChunk    As Long = 1024

Const DEFAULT_HashMult     As Long = -37
Const DEFAULT_HashMask     As Long = &H7FFFFF
Const DEFAULT_HashOffset   As Long = 0

Private Enum HashItems
    Item_Value = 0&
    Item_Key = 1&
    Item_Both = 2&
    Item_Ptr = 3&
End Enum

Private Type SlotType
    nextItem As Long      ' 0 if last item
    Key As String
    value As Variant
End Type

Private hashTable() As Long, slotTable() As SlotType, RXP As New RegExp
Private hash_Mult As Long, hash_Mask As Long, hash_Offset As Long
Private m_HashSize As Long, m_ListSize As Long, m_ChunkSize As Long, m_Count As Long, FreeNdx As Long
Private m_IgnoreCase As Boolean, m_Locked As Boolean, m_LockedPasswd As String, m_NoKey As Long

Public Default As Variant


Public Function Add(value As Variant, Optional ByVal Key As String) As Boolean
    Dim ndx As Long
    
    If m_Locked Then Exit Function
    
    If LenB(Key) = 0 Then
        Do
            Add = True
            m_NoKey = m_NoKey + 1
            Key = Chr$(7) & m_NoKey
            ndx = GetSlotIndex(Key, Add)
        Loop Until Add
    Else
        Add = True
        ndx = GetSlotIndex(Key, Add)
    End If
    
    If Add Then If IsObject(value) Then Set slotTable(ndx).value = value Else slotTable(ndx).value = value
End Function

'Default Member
Public Property Get Item(Optional ByVal Key As String) As Variant
Attribute Item.VB_UserMemId = 0
    Dim ndx As Long
    
    If LenB(Key) = 0 Then Exit Property

    ndx = GetSlotIndex(Key)
    
    If ndx = 0 Then     ' not found
        If IsObject(Default) Then Set Item = Default Else Item = Default
    ElseIf IsObject(slotTable(ndx).value) Then
        Set Item = slotTable(ndx).value
    Else
        Item = slotTable(ndx).value
    End If
End Property

Public Property Let Item(ByVal Key As String, value As Variant)
    Dim ndx As Long

    If LenB(Key) = 0 Or m_Locked Then Exit Property
    
    ndx = GetSlotIndex(Key, True)
    
    If IsObject(value) Then Set slotTable(ndx).value = value Else slotTable(ndx).value = value
End Property

Public Property Set Item(ByVal Key As String, ByVal value As Object)
    Dim ndx As Long

    If LenB(Key) = 0 Or m_Locked Then Exit Property
    
    ndx = GetSlotIndex(Key, True)
    Set slotTable(ndx).value = value
End Property

Public Function Remove(ByVal Key As String) As Boolean
    Dim ndx As Long, HCode As Long, LastNdx As Long
    
    If LenB(Key) = 0 Or m_Locked Then Exit Function
    
    ndx = GetSlotIndex(Key, False, HCode, LastNdx)

    If ndx = 0 Then Exit Function
    
    If LastNdx Then
        slotTable(LastNdx).nextItem = slotTable(ndx).nextItem
    ElseIf slotTable(ndx).nextItem Then
        hashTable(HCode) = slotTable(ndx).nextItem
    Else
        hashTable(HCode) = 0
    End If
    
    With slotTable(ndx)
        .nextItem = FreeNdx
        .Key = vbNullString
        .value = Empty
    End With
    
    FreeNdx = ndx
    
    m_Count = m_Count - 1
    
    Remove = True
End Function

Public Function Exists(ByVal Key As String) As Boolean
    Exists = GetSlotIndex(Key) <> 0
End Function

Public Property Get Count() As Long
    Count = m_Count
End Property

Public Property Get Items(Optional ByVal Flags As Long, Optional ByVal txtPattern As String) As Variant
    Dim a As Long, i As Long, n As Long, isRXP As Boolean, isFind As Boolean, v() As Variant
    
    If Flags = -1 Then
        ReDim v(UBound(hashTable))
        For a = 0 To UBound(hashTable)
            n = 0:    i = hashTable(a)
            Do While i:    n = n + 1:    i = slotTable(i).nextItem:    Loop
            v(a) = n
        Next
        Items = v:    Exit Property
    End If
    
    If m_Count = 0 Then Items = Array():   Exit Property
    
    isRXP = Len(txtPattern):      RXP.Pattern = txtPattern:      ReDim v(0 To m_Count - 1)
    
    For i = 0 To UBound(slotTable)
        With slotTable(i)
            If LenB(.Key) Then
                If isRXP Then isFind = RXP.Test(.Key) Else isFind = True
                
                If isFind Then
                    Select Case Flags
                        Case 0:     If IsObject(.value) Then Set v(n) = .value Else v(n) = .value
                        Case 1:     v(n) = .Key
                        Case 2:     v(n) = Array(.value, .Key)
                        Case 3:     If Not m_Locked Then v(n) = VarPtr(.nextItem)
                    End Select
                    
                    n = n + 1
                End If
            End If
        End With
    Next
    
    If isRXP Then If n = 0 Then Items = Array(): Exit Property Else ReDim Preserve v(n - 1)

    Items = v
End Property

Public Property Get IgnoreCase() As Boolean
    IgnoreCase = m_IgnoreCase
End Property

Public Property Let IgnoreCase(ByVal newValue As Boolean)
    If m_Count = 0 Then m_IgnoreCase = newValue
End Property

Public Property Get Locked(Optional ByVal txtPasswd As String) As Boolean
    Locked = m_Locked
End Property

Public Property Let Locked(Optional ByVal txtPasswd As String, ByVal vNewValue As Boolean)
    Dim Pos As Long, pwdOld As String
    
    Pos = InStr(txtPasswd, "@")
    If Pos Then
        pwdOld = Left$(txtPasswd, Pos - 1)
        txtPasswd = Mid$(txtPasswd, Pos + 1)
        If pwdOld = m_LockedPasswd Then m_LockedPasswd = txtPasswd
    End If
    
    If txtPasswd <> m_LockedPasswd Then Exit Property
    m_Locked = vNewValue
End Property

Public Function Child(Optional ByVal Key As String) As Object           'As clsHash
    If m_Locked Then Exit Function
    
    Set Child = New clsHash
    If IsObject(Default) Then Set Child.Default = Default Else Child.Default = Default
    Child.IgnoreCase = m_IgnoreCase
    Add Child, Key
End Function

Public Function Fill(ParamArray vsp() As Variant) As Object             'As clsHash
    Dim a As Long, uds As Long

    Set Fill = Me:          uds = UBound(vsp):          If m_Locked = True Or uds < 1 Then Exit Function
    
    For a = 0 To uds Step 2
        If Not IsObject(vsp(a)) Then If IsMissing(vsp(a)) Then vsp(a) = Empty
        Add vsp(a + 1), vsp(a)
    Next
End Function

Public Function Init(Optional ByVal HashSize As Long, Optional ByVal ListSize As Long, Optional ByVal ChunkSize As Long) As Boolean
    If m_Locked Then Exit Function
    
    If HashSize < 0 Or ListSize < 0 Or ChunkSize < 0 Then
        If m_Count Then Exit Function
        hash_Mult = Abs(HashSize)
        If ListSize Then hash_Mask = ListSize Else hash_Mask = DEFAULT_HashMask
        hash_Offset = ChunkSize
    Else
        If HashSize Then m_HashSize = HashSize
        If ListSize Then m_ListSize = ListSize
        If ChunkSize Then m_ChunkSize = ChunkSize
        
        If m_HashSize Mod 2 = 0 Then m_HashSize = m_HashSize + 1
        
        m_NoKey = 0:     m_Count = 0:     FreeNdx = 0
        
        ReDim hashTable(0 To m_HashSize - 1)
        ReDim slotTable(0)
        ExpandSlotTable m_ListSize
    End If
    
    Init = True
End Function


'--------------------------------------------------------------------------
Private Sub Class_Initialize()
    Call Init(DEFAULT_SizeHash, DEFAULT_SizeList, DEFAULT_SizeChunk)
    Call Init(DEFAULT_HashMult, DEFAULT_HashMask, DEFAULT_HashOffset)
    m_IgnoreCase = True
    RXP.IgnoreCase = True
End Sub

Private Sub Class_Terminate()
    Erase slotTable
End Sub

Private Sub ExpandSlotTable(ByVal numEls As Long)
    Dim newFreeNdx As Long, i As Long
    
    newFreeNdx = UBound(slotTable) + 1
    
    ReDim Preserve slotTable(0 To UBound(slotTable) + numEls) As SlotType

    For i = newFreeNdx To UBound(slotTable)             ' create the linked list of free items
        slotTable(i).nextItem = i + 1
    Next
    
    slotTable(UBound(slotTable)).nextItem = FreeNdx     ' overwrite the last (wrong) value
    FreeNdx = newFreeNdx                                ' we now know where to pick the first free item
End Sub

Private Function HashCode(Key As String) As Long
    Dim i As Long, b() As Byte

    If LenB(Key) = 0 Then Exit Function

    b = Key

    HashCode = hash_Offset

    For i = 0 To UBound(b)
        HashCode = (HashCode * hash_Mult + b(i)) And hash_Mask
    Next

    HashCode = HashCode Mod m_HashSize
End Function

Private Function GetSlotIndex(Key As String, Optional Create As Boolean, Optional HCode As Long, Optional LastNdx As Long) As Long
    Dim ndx As Long, lowKey As String
    
    If m_IgnoreCase Then
        lowKey = LCase$(Key)
        HCode = HashCode(lowKey)
        ndx = hashTable(HCode)
    
        Do While ndx
            If LCase$(slotTable(ndx).Key) = lowKey Then Exit Do
            LastNdx = ndx
            ndx = slotTable(ndx).nextItem
        Loop
    Else
        HCode = HashCode(Key)
        ndx = hashTable(HCode)
    
        Do While ndx
            If slotTable(ndx).Key = Key Then Exit Do
            LastNdx = ndx
            ndx = slotTable(ndx).nextItem
        Loop
    End If
    
    If (ndx = 0) And (Create = True) Then
        ndx = GetFreeSlot()
        PrepareSlot ndx, Key, HCode, LastNdx
    Else
        Create = False                                  ' signal that no item has been created
    End If

    GetSlotIndex = ndx
End Function

Private Function GetFreeSlot() As Long
    If FreeNdx = 0 Then ExpandSlotTable m_ChunkSize     ' allocate new memory if necessary
    GetFreeSlot = FreeNdx                               ' use the first slot
    FreeNdx = slotTable(GetFreeSlot).nextItem           ' update the pointer to the first slot
    slotTable(GetFreeSlot).nextItem = 0                 ' signal this as the end of the linked list
    m_Count = m_Count + 1                               ' we have one more item
End Function

Private Sub PrepareSlot(ByVal mIndex As Long, Key As String, ByVal HCode As Long, ByVal LastNdx As Long)
    slotTable(mIndex).Key = Key                         ' assign the key
    If LastNdx Then
        slotTable(LastNdx).nextItem = mIndex            ' this is the successor of another slot
    Else
        hashTable(HCode) = mIndex                       ' this is the first slot for a given hash code
    End If
End Sub
