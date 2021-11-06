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

Const DEF_SizeHash     As Long = 991        ' must be prime number
Const DEF_SizeList     As Long = 64
Const DEF_SizeChunk    As Long = 256

Const DEF_HashMult     As Long = -37
Const DEF_HashMask     As Long = &H7FFFFF
Const DEF_HashOffset   As Long = 0

Private Enum HashItems
    Item_Value = 0&
    Item_Key = 1&
    Item_Both = 2&
    Item_Ptr = 3&
End Enum

Private Type SlotType
    Next As Long
    Key As String
    value As Variant
End Type

Private hashTable() As Long, slotTable() As SlotType, RXP As New RegExp
Private hash_Mult As Long, hash_Mask As Long, hash_Offset As Long
Private m_IgnoreCase As Boolean, m_NoKey As Long, FreeNdx As Long
Private m_HashSize As Long, m_ListSize As Long, m_ChunkSize As Long, m_Count As Long

Public Default As Variant


Public Function Add(value As Variant, Optional ByVal Key As String) As Boolean
    Dim ndx As Long
    If LenB(Key) Then Add = True:  ndx = GetSlotIndex(Key, Add) Else ndx = AutoKey(Key, Add)
    If Add Then SetVar slotTable(ndx).value, value
End Function

'Default Member
Public Property Get Item(Optional ByVal Key As String) As Variant
Attribute Item.VB_UserMemId = 0
    Dim ndx As Long
    If LenB(Key) Then ndx = GetSlotIndex(Key):   If ndx Then SetVar Item, slotTable(ndx).value Else SetVar Item, Default
End Property

Public Property Let Item(ByVal Key As String, value As Variant)
    Dim ndx As Long
    If LenB(Key) Then ndx = GetSlotIndex(Key, True):   SetVar slotTable(ndx).value, value
End Property

Public Property Set Item(ByVal Key As String, ByVal value As Object)
    Dim ndx As Long
    If LenB(Key) Then ndx = GetSlotIndex(Key, True):   Set slotTable(ndx).value = value
End Property

Public Function Remove(ByVal Key As String) As Boolean
    Dim ndx As Long, HCode As Long, LastNdx As Long
    
    If LenB(Key) = 0 Then Exit Function Else ndx = GetSlotIndex(Key, False, HCode, LastNdx):  If ndx = 0 Then Exit Function
    
    If LastNdx Then
        slotTable(LastNdx).Next = slotTable(ndx).Next
    ElseIf slotTable(ndx).Next Then
        hashTable(HCode) = slotTable(ndx).Next
    Else
        hashTable(HCode) = 0
    End If
    
    With slotTable(ndx):      .Next = FreeNdx:      .value = Empty:      .Key = vbNullString:      End With
    
    FreeNdx = ndx:      m_Count = m_Count - 1:      Remove = True
End Function

Public Property Get Items(Optional ByVal Flags As Long, Optional ByVal pat As String, Optional ByVal fi As Boolean = True) As Variant
    Dim i As Long, n As Long, x As Long, isRXP As Boolean, v() As Variant
    
    If Flags = -1 Then
        ReDim v(UBound(hashTable)):    For x = 0 To UBound(v):    n = 0:    i = hashTable(x)
        Do While i:    n = n + 1:    i = slotTable(i).Next:    Loop:    v(x) = n:    Next:    Items = v:    Exit Property
    
    ElseIf m_Count Then
        ReDim v(0 To m_Count - 1):      If LenB(pat) Then isRXP = True:   RXP.Pattern = pat:   RXP.IgnoreCase = fi
        
        For i = 0 To UBound(slotTable)
            Do
                If LenB(slotTable(i).Key) = 0 Then Exit Do
                If isRXP Then If RXP.Test(slotTable(i).Key) = False Then Exit Do
                x = n:       n = n + 1
                If Flags = 0 Then SetVar v(x), slotTable(i).value:                      Exit Do
                If Flags = 1 Then v(x) = slotTable(i).Key:                              Exit Do
                If Flags = 2 Then v(x) = Array(slotTable(i).value, slotTable(i).Key):   Exit Do
                If Flags = 3 Then v(x) = VarPtr(slotTable(i).Next):                     Exit Do
            Loop Until True
        Next
    End If
    
    If n Then ReDim Preserve v(n - 1): Items = v Else Items = Array()
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Dim i As Long:     Static v As Collection:      Set v = New Collection
    If m_Count Then
        For i = 0 To UBound(slotTable)
            If LenB(slotTable(i).Key) Then v.Add slotTable(i).Value
        Next
    End If
    Set NewEnum = v.[_NewEnum]
End Property

Public Function Exists(ByVal Key As String) As Boolean
    Exists = GetSlotIndex(Key) <> 0
End Function

Public Property Get Count() As Long
    Count = m_Count
End Property

Public Property Get IgnoreCase() As Boolean
    IgnoreCase = m_IgnoreCase
End Property

Public Property Let IgnoreCase(ByVal newValue As Boolean)
    If m_Count = 0 Then m_IgnoreCase = newValue
End Property

Public Function Child(Optional ByVal Key As String) As Object                           'As clsHash
    Set Child = New clsHash:            Child.IgnoreCase = m_IgnoreCase
    SetVar Child.Default, Default:      Add Child, Key
End Function

Public Function Fill(ParamArray vsp() As Variant) As Object                             'As clsHash
    Dim a As Long, uds As Long

    Set Fill = Me:      uds = UBound(vsp):      If uds < 1 Then Exit Function
    
    For a = 0 To uds Step 2
        If Not IsObject(vsp(a)) Then If IsMissing(vsp(a)) Then vsp(a) = Empty
        Add vsp(a + 1), vsp(a)
    Next
End Function

Public Function Init(Optional ByVal HashSize As Long, Optional ByVal ListSize As Long, Optional ByVal ChunkSize As Long) As Boolean
    If HashSize < 0 Or ListSize < 0 Or ChunkSize < 0 Then
        If m_Count Then Exit Function
        hash_Mult = Abs(HashSize):      hash_Offset = ChunkSize:      hash_Mask = IIF(ListSize, ListSize, DEF_HashMask)
    Else
        If HashSize Then m_HashSize = HashSize:     If m_HashSize Mod 2 = 0 Then m_HashSize = m_HashSize + 1
        If ListSize Then m_ListSize = ListSize
        If ChunkSize Then m_ChunkSize = ChunkSize
        ReDim hashTable(0 To m_HashSize - 1):     ReDim slotTable(0):     m_NoKey = 0:     m_Count = 0:     FreeNdx = 0
        ExpandSlotTable m_ListSize
    End If
    Init = True
End Function


'--------------------------------------------------------------------------
Private Sub Class_Initialize()
    Call Init(DEF_SizeHash, DEF_SizeList, DEF_SizeChunk)
    Call Init(DEF_HashMult, DEF_HashMask, DEF_HashOffset)
    m_IgnoreCase = True
End Sub

Private Sub Class_Terminate()
    Erase slotTable
End Sub


Private Function GetSlotIndex(Key As String, Optional Create As Boolean, Optional HCode As Long, Optional LastNdx As Long) As Long
    Dim ndx As Long, LKey As String
    
    If m_IgnoreCase Then
        LKey = LCase$(Key):       HCode = HashCode(LKey):       ndx = hashTable(HCode)
    
        Do While ndx
            If LCase$(slotTable(ndx).Key) = LKey Then Exit Do Else LastNdx = ndx:   ndx = slotTable(ndx).Next
        Loop
    Else
        HCode = HashCode(Key):      ndx = hashTable(HCode)
    
        Do While ndx
            If slotTable(ndx).Key = Key Then Exit Do Else LastNdx = ndx:   ndx = slotTable(ndx).Next
        Loop
    End If
    
    If (ndx = 0) And Create Then ndx = GetFreeSlot():   PrepareSlot ndx, Key, HCode, LastNdx Else Create = False
    GetSlotIndex = ndx
End Function

Private Function HashCode(Key As String) As Long
    Dim i As Long, b() As Byte
    If LenB(Key) Then b = Key:  HashCode = hash_Offset Else Exit Function
    For i = 0 To UBound(b):     HashCode = (HashCode * hash_Mult + b(i)) And hash_Mask:     Next
    HashCode = HashCode Mod m_HashSize
End Function

Private Sub ExpandSlotTable(ByVal numAdd As Long)
    Dim newFreeNdx As Long, i As Long
    newFreeNdx = UBound(slotTable) + 1
    ReDim Preserve slotTable(0 To UBound(slotTable) + numAdd) As SlotType
    For i = newFreeNdx To UBound(slotTable):    slotTable(i).Next = i + 1:    Next      ' create the linked list of free items
    slotTable(UBound(slotTable)).Next = FreeNdx                                         ' overwrite the last (wrong) Value
    FreeNdx = newFreeNdx                                                                ' we now know where to pick the first free item
End Sub

Private Function GetFreeSlot() As Long
    If FreeNdx = 0 Then ExpandSlotTable m_ChunkSize                                     ' allocate new memory if necessary
    GetFreeSlot = FreeNdx                                                               ' use the first slot
    FreeNdx = slotTable(GetFreeSlot).Next                                               ' update the pointer to the first slot
    slotTable(GetFreeSlot).Next = 0                                                     ' signal this as the end of the linked list
    m_Count = m_Count + 1                                                               ' we have one more item
End Function

Private Sub PrepareSlot(ByVal mIndex As Long, Key As String, ByVal HCode As Long, ByVal LastNdx As Long)
    slotTable(mIndex).Key = Key
    If LastNdx Then slotTable(LastNdx).Next = mIndex Else hashTable(HCode) = mIndex
End Sub

Private Function AutoKey(Key As String, b As Boolean) As Long
    Do:   b = True:   m_NoKey = m_NoKey + 1:   Key = Chr$(7) & m_NoKey:   AutoKey = GetSlotIndex(Key, b):   Loop Until b
End Function

Private Sub SetVar(vDest As Variant, vSrc As Variant)
    If IsObject(vSrc) Then Set vDest = vSrc Else vDest = vSrc
End Sub
