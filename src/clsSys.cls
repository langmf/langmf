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

Public Function AddCodeMF(ByVal txtCode As String, Optional ByVal nameLib As String, Optional ByVal forceRunMF As Long) As String
    Dim Buf() As Byte
    Buf = ChrW$(65279) + txtCode
    AddCodeMF = Code_Parse(Buf, nameLib, forceRunMF)
End Function

Public Function AddFileMF(ByVal nameScript As String, Optional ByVal nameLib As String, Optional ByVal forceRunMF As Long) As String
    Dim Buf() As Byte
    nameScript = Parse_MPath(nameScript)
    If Not IsFileExt(nameScript, , Array(".mf")) Then Exit Function
    If File2Buf(Buf, nameScript) Then AddFileMF = Code_Parse(Buf, IIF(LenB(nameLib), nameLib, nameScript), forceRunMF)
End Function

Public Property Get Env(Optional ByVal ID As Variant) As Variant
    If IsMissing(ID) Or IsEmpty(ID) Then Set Env = GEV Else Env = GEV(ID)
End Property

Public Property Get GenStr(Optional ByVal value As Variant, Optional ByVal txtPat As String) As String
    GenStr = GenTempStr(value, txtPat)
End Property

Public Function DoParams(ByVal Obj As Object, Arg As Variant) As Object
    Set DoParams = modCommon.DoParams(Obj, Arg)
End Function

Public Property Get NewForm(Optional ByVal Param As Variant) As Object
    Set NewForm = modCommon.DoParams(New frmForm, Param)
End Property

Public Property Get NewScript(Optional ByVal Param As Variant) As Object                            'As clsActiveScript
    Set NewScript = modCommon.DoParams(New clsActiveScript, Param)
End Property

Public Property Get NewSubClass(Optional ByVal Param As Variant) As Object                          'As clsSubClass
    Set NewSubClass = modCommon.DoParams(New clsSubClass, Param)
End Property

Public Property Get NewCol(Optional ByVal Param As Variant) As Collection
    Set NewCol = modCommon.DoParams(New Collection, Param)
End Property

Public Property Get NewHash(Optional ByVal Param As Variant) As Object                              'As clsHash
    Set NewHash = modCommon.DoParams(New clsHash, Param)
End Property

Public Property Get NewDim(Optional ByVal Param As Variant) As Object                               'As clsDim
    Set NewDim = modCommon.DoParams(New clsDim, Param)
End Property

Public Property Get NewGlob(Optional ByVal Param As Variant) As Object                              'As clsGLOB
    Set NewGlob = modCommon.DoParams(New clsGLOB, Param)
End Property

Public Property Get NewBuf(Optional Buf As Variant, Optional ByVal Param As Variant) As Object      'As clsBuffer
    Set NewBuf = modCommon.DoParams(New clsBuffer, Param)
    If Not IsMissing(Buf) Then NewBuf.SetData Buf
End Property

Public Property Get NewRtc(Optional value As Variant, Optional ByVal Param As Variant) As Object    'As clsRTC
    Set NewRtc = modCommon.DoParams(New clsRTC, Param)
    If Not IsMissing(value) Then NewRtc.Compile value
End Property

Public Property Get NewFCP(Optional value As Variant, Optional ByVal Param As Variant) As Object    'As clsFuncPointer
    Set NewFCP = New clsFuncPointer
    
    If Not IsMissing(value) Then
        Select Case VarType(value)
            Case vbObject:              NewFCP.ptrIfc = ObjPtr(value)
            Case vbLong, vbInteger:     NewFCP.ptrIfc = CLng(value)
            Case Else:                  NewFCP.Buffer = value
        End Select
    End If
    
    modCommon.DoParams NewFCP, Param
End Property

Public Property Get NewFAP(Optional ByVal Obj As Variant, Optional Buf As Variant, Optional ByVal Param As Variant) As Object         'As clsAsmPic
    Set NewFAP = New clsAsmPic
    
    If IsObject(Obj) Then Set NewFAP.Obj = Obj Else NewFAP.Obj = Obj
    If Not IsMissing(Buf) Then NewFAP.ASM.Buffer = Buf
        
    modCommon.DoParams NewFAP, Param
End Property

Public Property Get Path(Optional ByVal value As Variant = False, Optional ByVal backSlash As Boolean = True) As String
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
End Property

Public Property Get Host(ParamArray Args() As Variant) As Variant
    Dim uds As Long, Param As String
    
    On Error Resume Next
    
    uds = UBound(Args):    If uds >= 0 Then Param = LCase$(Args(0))
    
    If Left$(Param, 3) = "arg" Then
        Select Case Mid$(Param, 4)
        Case "v":               Set Host = GetArgv(Info.Arg)
        Case "o":               Set Host = SYS.SHD.Object(GetArgv(Info.Arg))
        Case "":                Host = Info.Arg
        End Select
    Else
        Select Case Param
        Case "host", "":        Set Host = LMF
        Case "mdl_name":        Host = MDL(Args(1)).Name
        Case "mdl_type":        Host = MDL(Args(1)).Type
        Case "noerror":         If uds < 1 Then Host = mf_NoError Else mf_NoError = Args(1):   Host = mf_NoError
        End Select
    End If
End Property

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
        Case "nosleep":         Power = SetThreadExecutionState(IIF(Flags = -1, ES_CONTINUOUS Or ES_SYSTEM_REQUIRED Or ES_DISPLAY_REQUIRED, IIF(Flags = 0, ES_CONTINUOUS, Flags)))
    End Select
End Function

Public Function OnTimer(ByVal Interval As Long, ByVal Param As Variant, ParamArray Args() As Variant) As Long
    Dim idT As Long, tmp As String

    If Interval < 0 Then idT = -1
    
    While Timers.Exists("#" & idT)
        If Interval < 0 Then idT = idT - 1 Else idT = idT + 1
    Wend
    
    If Not IsArray(Param) Then tmp = Param:   Param = Empty
    ArrayDef Param, tmp, "", Args, idT
    Timers("#" & idT) = Param

    Call SetTimer(frmScript.hWnd, idT, Abs(Interval), AddressOf Timer_Event)
    
    OnTimer = idT
End Function

Public Function OffTimer(ByVal Param As Variant) As Long
    Dim v As Variant, nMod As Variant, nFunc As String
    
    If IsArray(Param) Then ArrayDef Param, "", 0:   nFunc = Param(0):   nMod = Param(1) Else nFunc = Param:   nMod = ""

    For Each v In Timers.Items
        If v(1) = nMod Or (IsNumeric(nMod) And nMod = 0) Then
            If v(0) = nFunc Or LenB(nFunc) = 0 Then OffTimer = KillTimer(frmScript.hWnd, v(3)):   Timers.Remove "#" & v(3)
        End If
    Next
End Function

Public Sub Prm(Param As Variant, ParamArray vsp() As Variant)
    Dim a As Long, uds As Long

    uds = UBound(vsp)
    If uds < 0 Then Exit Sub
    
    If Not IsArray(Param) Then
        Param = Empty
        ReDim Param(uds)
    ElseIf ArraySize(Param) = 0 Then
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
        If vCount >= 0 And vCount <= 10 Then CallBack.Init Obj, vMethod, vCount, isCDecl
    Else
        If vCount > -1 Then
            If ExistsMember(CAS.CodeObject(Obj), CStr(vMethod)) = False Then Exit Function
        Else
            Set i = COM.Enums(CAS.CodeObject(Obj))
            If Not i.Exists("Funcs") Then Exit Function
            If Not i("Funcs").Exists(vMethod) Then Exit Function
            If i("Funcs")(vMethod).Item("Type") = VT_VOID Then Exit Function
            vCount = i("Funcs")(vMethod).Item("Prms")
            If vCount < 0 Or vCount > 10 Then Exit Function
        End If
        CallBack.Init CAS.Eval("GetRef(""" & vMethod & """)", Obj), , vCount, isCDecl
    End If
End Function

Public Function Content(ByVal value As String, Optional ByVal bString As Boolean = True, Optional vStatus As Variant) As Variant
    Dim vProt As String, vName As String, Buf() As Byte, RXP As New clsRXP
    
    vStatus = Array(False, 0, "", "", 0, ""):      If bString Then Content = "" Else Content = Buf
    
    If Not RXP.Test(value, "^([a-z]+:\/\/)?(.+)") Then Exit Function
    
    vProt = LCase$(RXP.Mts(0).SubMatches(0)):      vStatus(2) = vProt
    vName = RXP.Mts(0).SubMatches(1):              vStatus(3) = vName
    
    Select Case vProt
        Case "http://", "https://"
            vStatus(1) = 3
            With Net.XMLHTTP(value)
                On Error Resume Next
                vStatus(4) = .Status:    vStatus(5) = .StatusText
                On Error GoTo 0
                If .readyState = 4 And vStatus(4) = 200 Then
                    If ArraySize(.responseBody) Then
                        Buf = .responseBody
                        If bString Then Content = Conv_A2W_Buf(Buf) Else Content = Buf
                    End If
                    vStatus(0) = True
                Else
                    If .readyState = 0 Then vStatus(5) = "Timeout abort."
                    If vStatus(4) >= 12000 Then vStatus(5) = API_Error(vStatus(4), "wininet.dll")
                End If
            End With

        Case "res://"
            vStatus(1) = 2
            If SYS.Resource.Exists(vName) Then
                If bString Then Content = SHD.ResToStr(vName) Else Content = SHD.ResToBuf(vName)
                vStatus(0) = True
            Else
                vStatus(4) = 3:      vStatus(5) = "Not found resource '" & vName & "'"
            End If

        Case "file://", ""
            vStatus(1) = 1
            If File2Buf(Buf, vName) Then
                If bString Then Content = Conv_A2W_Buf(Buf) Else Content = Buf
                vStatus(0) = True
            Else
                vStatus(4) = Err.LastDllError:      vStatus(5) = API_Error(vStatus(4))
            End If
    End Select
End Function
