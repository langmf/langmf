VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
 
Private Type def_Processes
    pid As Long
    Sid As Long
    pMemory As Long
    pname As String
    pPath As String
    pUser As String
End Type

Private WTS As Collection, txtUser As String, txtDomain As String, txtNamePath As String, txtDos(25) As String


Public Sub AdjustToken(Optional ByVal typePrivilege As String = "SeDebugPrivilege", Optional ByVal processID As Long = -1, Optional ByVal Flag As Long = SE_PRIVILEGE_ENABLED)
    modCommon.AdjustToken typePrivilege, processID, Flag
End Sub

Public Function Kill(ByVal processID As Variant) As Boolean
    Dim a As Long, hProcess As Long, pid As Long, cl As Collection, isStr As Boolean, isFind As Boolean
    
    isStr = (VarType(processID) = vbString)
    
    Set cl = List
    
    For a = 1 To cl.Count
        If isStr Then
            If LCase$(cl(a)("Name")) = LCase$(processID) Then isFind = True
        Else
            If cl(a)("Pid") = processID Then isFind = True
        End If
        
        If isFind Then
            pid = cl(a)("Pid")
            hProcess = OpenProcess(1, False, pid)
            Kill = TerminateProcess(hProcess, 0)
            CloseHandle hProcess
            Exit For
        End If
    Next
End Function

'Default Member
Public Function Find(ByVal processID As Variant) As Boolean
Attribute Find.VB_UserMemId = 0
    Dim a As Long, cl As Collection, isStr As Boolean, isFind As Boolean
    
    isStr = (VarType(processID) = vbString)
    
    Set cl = List
    
    For a = 1 To cl.Count
        If isStr Then
            If LCase$(cl(a)("Name")) = LCase$(processID) Then isFind = True
        Else
            If cl(a)("Pid") = processID Then isFind = True
        End If
        
        If isFind Then Find = True:   Exit For
    Next
End Function

Public Function List() As Collection
    Dim lstProcesses() As Long, cbNeeded As Long, a As Long, cl As New Collection

    AdjustToken
    ListWTS
    ListWindows
    InitNTDevice
    
    ReDim lstProcesses(1023) As Long
    
    If EnumProcesses(lstProcesses(0), 1024 * 4, cbNeeded) <> 0 Then
        For a = 0 To cbNeeded / 4 - 1
            cl.Add GetProcessInfo(lstProcesses(a)), "P" & lstProcesses(a)
        Next
    End If
    
    Set List = cl
End Function

Private Function GetProcessInfo(ByVal lProcessID As Long) As Collection
    Dim hProcess As Long, tPMC As PROCESS_MEMORY_COUNTERS, Proc As def_Processes, cl As New Collection, v As Variant
    
    Proc.pid = lProcessID
    
    Select Case lProcessID
        Case 0:       Proc.pname = "System Idle Process"
        Case 4, 8:    Proc.pname = "System"
        Case Else
            hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, lProcessID)
            If hProcess = 0 Then hProcess = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, 0, lProcessID)
    
            If hProcess <> 0 Then
                GetNamePath hProcess, Proc.pname, Proc.pPath
                If GetProcessMemoryInfo(hProcess, tPMC, Len(tPMC)) <> 0 Then Proc.pMemory = tPMC.WorkingSetSize
            End If
            
            CloseHandle hProcess
    End Select
    
    If WTS.Count Then Proc.pUser = WTS("U" & Proc.pid):     Proc.Sid = WTS("S" & Proc.pid)

    v = HashWins(Proc.pid)
    If IsEmpty(v) Then v = Array(0, "", "")
    
    With cl
        .Add Proc.pid, "Pid":       .Add Proc.pname, "Name":    .Add Proc.pMemory, "Mem":   .Add Proc.Sid, "Sid"
        .Add Proc.pUser, "User":    .Add Proc.pPath, "Path":    .Add v, "Win"
    End With
    
    Set GetProcessInfo = cl
End Function

Private Sub ListWTS()
    Dim Count As Long, i As Long, lpBuffer As Long, p As Long, udtInfo As WTS_PROCESS_INFO
    
    Set WTS = New Collection
    
    If WTSEnumerateProcesses(WTS_CURRENT_SERVER, 0, 1, lpBuffer, Count) > 0 Then
        p = lpBuffer
        
        For i = 0 To Count - 1
            CopyMemory udtInfo, ByVal p, LenB(udtInfo)
            
            WTS.Add udtInfo.SessionID, "S" & udtInfo.processID
            WTS.Add GetUserFromID(udtInfo.pUserSid), "U" & udtInfo.processID

            p = p + LenB(udtInfo)
        Next
        
        WTSFreeMemory lpBuffer
    End If
End Sub
    
Private Sub InitNTDevice()
    Dim a As Long, txt As String
    
    txt = Space$(1000)
        
    For a = 0 To 25
        If QueryDosDeviceA(Chr$(65 + a) & ":", txt, Len(txt)) <> 0 Then
            txtDos(a) = LCase$(TrimNull(txt))
        Else
            txtDos(a) = vbNullString
        End If
    Next
End Sub

Private Sub GetNamePath(ByVal hProcess As Long, txtName As String, txtPath As String)
    Dim a As Long, rc As Long
    
    If WinVer.dwMajorVersion = 5 And WinVer.dwMinorVersion = 0 Then
        rc = GetModuleFileNameExW(hProcess, 0, StrPtr(txtNamePath), Len(txtNamePath))
    Else
        rc = GetProcessImageFileNameW(hProcess, StrPtr(txtNamePath), Len(txtNamePath))
    End If

    If rc > 0 Then
        txtPath = Left$(txtNamePath, rc)
        txtPath = Replace(txtPath, "\??\", "")
        For a = 0 To 25
            If Len(txtDos(a)) Then
                If InStr(1, txtPath, txtDos(a), vbTextCompare) = 1 Then
                    txtPath = Chr$(65 + a) & ":" & Mid$(txtPath, Len(txtDos(a)) + 1)
                    Exit For
                End If
            End If
        Next
        txtName = GetFileTitle(txtPath)
    Else
        txtPath = ""
        txtName = "unknown"
    End If
End Sub

Private Function GetUserFromID(ByVal Sid As Long) As String
    Dim szUser As Long, szDomain As Long
    
    szUser = Len(txtUser):      szDomain = Len(txtDomain)
    
    If LookupAccountSidW(0, Sid, StrPtr(txtUser), szUser, StrPtr(txtDomain), szDomain, 0) <> 0 Then
        GetUserFromID = Left$(txtDomain, szDomain) & "\" & Left$(txtUser, szUser)
    End If
End Function

Private Sub Class_Initialize()
    txtUser = Space$(256):      txtDomain = Space$(256):       txtNamePath = Space$(1024)
End Sub
