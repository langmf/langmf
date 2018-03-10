VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public E As Object
Public App As Object
Public Forms As Object
Public Screen As Object
Public Printers As Object
Public Licenses As Object
Public Clipboard As Object

Public API As Object
Public CGI As Object
Public GDI As Object
Public SHD As Object
Public COM As Object
Public RXP As Object
Public Net As Object
Public Var As Object
Public Ext As Object
Public Str As Object
Public Conv As Object
Public File As Object
Public CDlg As Object
Public Cron As Object
Public Bytes As Object
Public NatRes As Object
Public Process As Object
Public Registry As Object
Public Resource As Object
Public Timers As Object


Private Sub Class_Initialize()
    Set E = Err
    Set App = VB.App
    Set Forms = VB.Forms
    Set Screen = VB.Screen
    Set Printers = VB.Printers
    Set Licenses = VB.Licenses
    Set Clipboard = VB.Clipboard
    Set API = New clsAPI
    Set CGI = New clsCGI
    Set GDI = New clsGDI
    Set SHD = New clsShared
    Set COM = New clsCOM
    Set RXP = New clsRXP
    Set Net = New clsNet
    Set Var = New clsVar
    Set Ext = New clsExt
    Set Str = New clsString
    Set Conv = New clsConv
    Set File = New clsFile
    Set CDlg = New clsCDlg
    Set Cron = New clsCron
    Set Bytes = New clsBytes
    Set NatRes = New clsNativeRes
    Set Process = New clsProcess
    Set Registry = New clsRegistry
    Set Resource = New clsHash
    Set Timers = New clsHash
End Sub

Public Function AddCodeMF(ByVal txtCode As String, ByVal nameLib As String) As String
    Dim Buf() As Byte
    Buf = ChrW$(65279) + txtCode
    AddCodeMF = Code_Parse(Buf, nameLib)
End Function

Public Function AddFileMF(ByVal nameScript As String, Optional ByVal nameLib As String = "") As String
    Dim Buf() As Byte
    nameScript = Parse_MPath(nameScript)
    If Not IsFileExt(nameScript, , Array(".mf")) Then Exit Function
    m_File2Buf Buf, nameScript
    AddFileMF = Code_Parse(Buf, IIF(LenB(nameLib), nameLib, nameScript))
End Function

Public Function Env(Optional ByVal ID As Variant) As Variant
    If IsMissing(ID) Or IsEmpty(ID) Then Set Env = GEV Else Env = GEV(ID)
End Function

Public Function GenStr(Optional ByVal lenGen As Integer = 8, Optional ByVal txtPat As String) As String
    GenStr = GenTempStr(lenGen, txtPat)
End Function

Public Function DoParams(ByVal Obj As Object, Arg As Variant) As Object
    Set DoParams = m_DoParams(Obj, Arg)
End Function

Public Function NewForm(Optional ByVal Param As Variant) As Object
    Set NewForm = m_DoParams(New frmForm, Param)
End Function

Public Function NewScript(Optional ByVal Param As Variant) As Object                                'As clsActiveScript
    Set NewScript = m_DoParams(New clsActiveScript, Param)
End Function

Public Function NewSubClass(Optional ByVal Param As Variant) As Object                              'As clsSubClass
    Set NewSubClass = m_DoParams(New clsSubClass, Param)
End Function

Public Function NewCollection(Optional ByVal Param As Variant) As Collection
    Set NewCollection = m_DoParams(New Collection, Param)
End Function

Public Function NewHash(Optional ByVal Param As Variant) As Object                                  'As clsHash
    Set NewHash = m_DoParams(New clsHash, Param)
End Function

Public Function NewDim(Optional ByVal Param As Variant) As Object                                   'As clsDim
    Set NewDim = m_DoParams(New clsDim, Param)
End Function

Public Function NewGlob(Optional ByVal Param As Variant) As Object                                  'As clsGLOB
    Set NewGlob = m_DoParams(New clsGLOB, Param)
End Function

Public Function NewBuf(Optional Buf As Variant, Optional ByVal Param As Variant) As Object          'As clsBuffer
    Set NewBuf = m_DoParams(New clsBuffer, Param)
    If Not IsMissing(Buf) Then NewBuf.SetData Buf
End Function

Public Function NewFCP(Optional value As Variant, Optional ByVal Param As Variant) As Object        'As clsFuncPointer
    Set NewFCP = New clsFuncPointer
    
    If Not IsMissing(value) Then
        Select Case VarType(value)
            Case vbObject:              NewFCP.ptrIfc = ObjPtr(value)
            Case vbLong, vbInteger:     NewFCP.ptrIfc = CLng(value)
            Case Else:                  NewFCP.SetBuffer value
        End Select
    End If
    
    m_DoParams NewFCP, Param
End Function

Public Function NewFAP(Optional ByVal Obj As Object, Optional Buf As Variant, Optional ByVal Param As Variant) As Object             'As clsAsmPic
    Set NewFAP = New clsAsmPic
    
    Set NewFAP.Obj = Obj
    If Not IsMissing(Buf) Then NewFAP.ASM.SetBuffer Buf
        
    m_DoParams NewFAP, Param
End Function

Public Function Path(Optional ByVal value As Variant = False, Optional ByVal backSlash As Boolean = True) As String
    Dim vt As Long, Buf As String
    
    vt = VarType(value)
    
    If vt = 11 Then
        If value Then Path = Info.File Else Path = CPath(GetDirectory(Info.File), backSlash)
        
    ElseIf vt = 8 Then
        Buf = String$(MAX_PATH_X2, 0):      vt = Len(Buf)
        
        Select Case LCase$(value)
            Case "script":          Path = CPath(Path(False), backSlash)
            Case "script_full":     Path = Path(True)
            Case "engine":          Path = CPath(GetAppPath(False), backSlash)
            Case "engine_full":     Path = GetAppPath(True)
            Case "system":          Path = CPath(GetSystemPath, backSlash)
            Case "windows":         Path = CPath(GetWindowsPath, backSlash)
            Case "temp":            Path = CPath(GetTmpPath, backSlash)
            Case "tmp":             Path = mf_Tmp
            Case "user":            Call GetUserNameW(StrPtr(Buf), vt):       Path = Left$(Buf, vt - 1)
            Case "computer":        Call GetComputerNameW(StrPtr(Buf), vt):   Path = Left$(Buf, vt)
        End Select
    End If
End Function

Public Function Host(Optional ByVal value As String, Optional ByVal ID As Long = 1) As Variant
    On Error Resume Next
    
    Select Case LCase$(value)
        Case "host", "":        Set Host = LMF
        Case "arg":             Host = Info.Arg
        Case "time_lmf":        Host = mf_TimeLMF
        Case "time_parse":      Host = mf_TimeParse
        Case "mdl_name":        Host = MDL(ID).Name
        Case "mdl_type":        Host = MDL(ID).Type
    End Select
End Function

Public Function Power(Optional ByVal value As String = "off", Optional ByVal Flags As Long) As Long
    On Error Resume Next
    
    AdjustToken "SeShutdownPrivilege"
    
    Select Case LCase$(value)
        Case "":                Power = ExitWindowsEx(Flags, &HFFFF&)
        Case "off":             Power = ExitWindowsEx(Flags Or EWX_POWEROFF, &HFFFF&)
        Case "reboot":          Power = ExitWindowsEx(Flags Or EWX_REBOOT, &HFFFF&)
        Case "logoff":          Power = ExitWindowsEx(Flags Or EWX_LOGOFF, &HFFFF&)
        Case "suspend":         Power = SetSuspendState((Flags And 1&), 0, (Flags And 2&))
        Case "monitor":         Power = SendMessageW(frmScript.hWnd, WM_SYSCOMMAND, SC_MONITORPOWER, ByVal Flags)
    End Select
End Function

Public Function OnTimer(ByVal Interval As Long, ByVal Param As Variant, ParamArray dataArg() As Variant) As Long
    Dim idT As Long, tmp As String

    If Interval < 0 Then idT = -1
    
    While Timers.Exists("#" & idT)
        If Interval < 0 Then idT = idT - 1 Else idT = idT + 1
    Wend
    
    If Not IsArray(Param) Then tmp = Param: Param = Empty
    m_ParamArray Param, tmp, "", dataArg, idT
    Timers("#" & idT) = Param

    Call SetTimer(frmScript.hWnd, idT, Abs(Interval), AddressOf Timer_Event)
    
    OnTimer = idT
End Function

Public Sub OffTimer(ByVal Param As Variant)
    Dim v As Variant, nMod As Variant, nFunc As String
    
    If IsArray(Param) Then m_ParamArray Param, "", 0:   nFunc = Param(0):   nMod = Param(1) Else nFunc = Param:   nMod = ""

    For Each v In Timers.Items
        If v(1) = nMod Or (IsNumeric(nMod) And nMod = 0) Then
            If v(0) = nFunc Or LenB(nFunc) = 0 Then Call KillTimer(frmScript.hWnd, v(3)):    Timers.Remove "#" & v(3)
        End If
    Next
End Sub

Public Sub Prm(Param As Variant, ParamArray vsp() As Variant)
    Dim a As Long, uds As Long

    uds = UBound(vsp)
    If uds < 0 Then Exit Sub
    
    If Not IsArray(Param) Then
        Param = Empty
        ReDim Param(uds)
    ElseIf m_ArraySize(Param) = 0 Then
        ReDim Param(uds)
    Else
        If UBound(Param) <> uds Then ReDim Preserve Param(uds)
    End If
    
    For a = 0 To uds
        If Not IsObject(Param(a)) Then
            If IsEmpty(Param(a)) Or IsMissing(Param(a)) Then
                If IsObject(vsp(a)) Then Set Param(a) = vsp(a) Else Param(a) = vsp(a)
            End If
        End If
    Next
End Sub

Public Function CallBack(Optional ByVal vMethod As Variant = 0, Optional Obj As Variant = "", Optional ByVal vCount As Long = -1, Optional ByVal isCDecl As Boolean) As Object             'As clsCallback
    Dim i As clsHash

    Set CallBack = New clsCallback
    
    If IsObject(Obj) Then
        If VerifyLongRange(vCount, 0, 10) Then CallBack.Init Obj, vMethod, vCount, isCDecl
    Else
        If vCount > -1 Then
            If ExistsMember(CAS.CodeObject(Obj), CStr(vMethod)) = False Then Exit Function
        Else
            Set i = COM.Enums(CAS.CodeObject(Obj))
            If Not i.Exists("Funcs") Then Exit Function
            If Not i("Funcs").Exists(vMethod) Then Exit Function
            If i("Funcs")(vMethod).Item("Type") = VT_VOID Then Exit Function
            vCount = i("Funcs")(vMethod).Item("Prms")
            If VerifyLongRange(vCount, 0, 10) = False Then Exit Function
        End If
        CallBack.Init CAS.Eval("GetRef(""" & vMethod & """)", Obj), , vCount, isCDecl
    End If
End Function

Public Function Content(ByVal value As String, Optional ByVal bString As Boolean = True) As Variant
    Dim Buf() As Byte, tmp As String, RX As New clsRXP
    
    If bString Then Content = "" Else Content = Buf
    
    If RX.Test(value, "^([a-z]+:\/\/)?(.+)") = False Then Exit Function
    
    tmp = RX.Mts(0).SubMatches(1)
    
    Select Case LCase$(RX.Mts(0).SubMatches(0))
        Case "http://", "https://"
            With Net.XMLHTTP(value)
                If m_ArraySize(.responseBody) = 0 Then Exit Function
                Buf = .responseBody
                If bString Then Content = Int_A2W_Buf(Buf) Else Content = Buf
            End With
    
        Case "res://"
            If bString Then Content = SHD.ResToStr(tmp) Else Content = SHD.ResToBuf(tmp)
        
        Case "file://", ""
            If m_File2Buf(Buf, tmp) Then
                If bString Then Content = Int_A2W_Buf(Buf) Else Content = Buf
            End If
    End Select
End Function
