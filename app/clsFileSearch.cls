VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFileSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function FindFirstFileW Lib "kernel32" (ByVal lpFileName As Long, ByVal lpFindFileData As Long) As Long
Private Declare Function FindNextFileW Lib "kernel32" (ByVal hFindFile As Long, ByVal lpFindFileData As Long) As Long

Private findTxt() As Variant, Count As Long, REG1 As RegExp, useREG As Boolean
Private OnlyFile As Boolean, IsRecursion As Boolean


'Default Member
Public Function Find(ByVal MPath As String, Optional ByVal patFile As String, Optional ByVal m_OnlyFile As Boolean = True, Optional ByVal m_IsRecursion As Boolean = True) As Variant
Attribute Find.VB_UserMemId = 0
    Count = -1
    OnlyFile = m_OnlyFile
    IsRecursion = m_IsRecursion
    useREG = Len(patFile)
    REG1.Pattern = patFile
    
    If Right$(MPath, 1) <> "\" Then MPath = MPath + "\"
    
    Call m_Find(MPath)

    Find = findTxt
    Erase findTxt
End Function

Private Sub m_Find(MPath As String)
    Dim hFnd As Long, nFile As String, isFind As Boolean, WFD As WIN32_FIND_DATA

    hFnd = FindFirstFileW(StrPtr(LongPath(MPath & "*")), VarPtr(WFD))
    If hFnd = INVALID_HANDLE Then Exit Sub
    
    With REG1
        Do
            nFile = TrimNull(WFD.cFileName)

            If nFile <> "." And nFile <> ".." Then
                If (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = 0 Then
                    If OnlyFile Then
                        If useREG Then isFind = .Test(nFile) Else isFind = True
                        If isFind Then AddItem MPath & nFile
                    End If
                Else
                    If Not OnlyFile Then AddItem MPath & nFile
                    If IsRecursion Then Call m_Find(MPath & nFile & "\")
                End If
            End If
        Loop While FindNextFileW(hFnd, VarPtr(WFD))
    End With
    
    FindClose hFnd
End Sub

Private Sub AddItem(value As String)
    Count = Count + 1
    ReDim Preserve findTxt(Count)
    findTxt(Count) = value
End Sub

Private Sub Class_Initialize()
    Set REG1 = New RegExp
    REG1.IgnoreCase = True
    REG1.Global = False
End Sub
