Attribute VB_Name = "modMain"
Option Explicit

Type def_ParseCustom
    Param   As String
    Data    As String
End Type

Type def_HeaderCompile
    Signature   As Long
    Length      As Integer
    Packer      As Integer
    Reserved    As Long
    HExtOffset  As Long
    HExtCount   As Integer
    VerMajor    As Integer
    VerMinor    As Integer
    VerBuild    As Integer
    DataOffset  As Long
    DataSize    As Long
End Type

Type def_HeaderExt
    HeaderID    As Long
    Reserved    As Long
    DataOffset  As Long
    DataSize    As Long
End Type

Type def_MDL
    MFC   As Boolean
    Type  As type_MDL
    Name  As String
    Code  As String
    Path  As String
End Type

Type def_Info
    IsExe     As Boolean
    IsCmd     As Boolean
    StartExe  As Long
    SizeExe   As Long
    File      As String
    Arg       As String
End Type

Public Enum type_MDL
    tpModule = 0
    tpForm = 1
End Enum

Global Const mf_Sign = &H2043464D, mf_Setup = "/regsetup", mf_New = "/regnew", mf_Embed = "-Embedding"
Global Const mf_Hdr = vbNullChar + "-=~lmfhdr~=-" + vbNullChar
Global Const mf_EMailDevelop = "support@langmf.ru"

Global EMailDevelop As String

Global Info As def_Info
Global MDL(255) As def_MDL

Global LMF As LangMF
Global SYS As clsSys
Global CAS As clsActiveScript
Global REG As New RegExp
Global Types As clsHash

Global mf_TimeLMF As Long, mf_TimeParse As Long, mf_AsyncLoad As Long, mf_NoError As Boolean
Global mf_Counter As Integer, mf_IsEnd As Boolean, mf_Tmp As String


Sub Main()
    Dim stp As Boolean, cmd As String
    
    mf_TimeLMF = timeGetTime

    InitGlobal
    stp = SetupLMF

    mf_TimeLMF = timeGetTime - mf_TimeLMF
    
    cmd = Command

    If cmd <> mf_Embed And cmd <> mf_Setup Then
        If Not stp Then ShellSyncEx GetAppPath(True), mf_Setup, , , "runas"
        Set LMF = New LangMF
        If cmd = mf_New Then LMF.ROT Else Script_EXE:  LMF.Command cmd
    End If
End Sub


Function SetupLMF() As Boolean
    Dim fName As String, v As Variant, tmp As String, isOK As Boolean, Buf() As Byte
    Dim clsRG As New clsRegistry, clsNR As New clsNativeRes, Prm As clsHash
    
    On Error Resume Next
    
    SetupLMF = True

    For Each v In clsNR.EnumResource(GetAppPath(True), 10)
        If VarType(v) = vbString Then

            Set Prm = GetQV(v, fName)
            fName = GetSystemPath + "\" + LCase$(fName)
            
            If Prm.Exists("ProgID") Then
                tmp = clsRG.RegRead("HKCR\CLSID\" + clsRG.RegRead("HKCR\" + Prm("ProgID") + "\CLSID\") + "\InProcServer32\" + IIF(Prm.Exists("NET"), "CodeBase", ""))
                If Prm.Exists("NET") Then tmp = Replace$(Mid$(tmp, 9), "/", "\")
            Else
                If IsFile(fName) Then tmp = fName Else tmp = GetAppPath + GetFileName(fName)
            End If
            
            isOK = False
            If IsFile(tmp) Then If VersionDLL(tmp, Prm("Ver")) Then isOK = True
            
            If Not isOK Then
                Buf = LoadResData(v, 10)
                
                If m_Buf2File(Buf, fName) Then
                    If Prm.Exists("ProgID") Then
                        If Prm.Exists("NET") Then
                            isOK = (ShellSync(GetWindowsPath + "\Microsoft.NET\Framework\" + Prm("NET") + "\regasm.exe " & fName & " /codebase /tlb", 8000, True) = 0)
                        Else
                            isOK = RegisterDLL(fName)
                        End If
                    Else
                        If InStr(1, v, "INPOUT32.DLL") Then Call IsInpOutDriverOpen
                        isOK = True
                    End If
                End If
            End If
            
            If Not isOK Then SetupLMF = False
        End If
    Next
End Function


Sub Code_Run(Optional ByVal nameScript As String)
    Dim f As New clsFileAPI, Buf() As Byte
    
    mf_TimeParse = timeGetTime
    
    Script_Init

    If Not Info.IsExe And Not Info.IsCmd Then
        If Len(GEV("REQUEST_URI")) Then
            nameScript = Replace$(GEV("PATH_TRANSLATED"), "/", "\")
            Info.File = nameScript
            Info.IsCmd = True
        Else
            frmAbout.Show
            Exit Sub
        End If
    End If
    
    If IsFile(nameScript) Then
        If f.FOpen(nameScript, OPEN_EXISTING, GENERIC_READ) <> INVALID_HANDLE Then
            If Info.IsExe Then
                If Info.SizeExe Then
                    ReDim Buf(Info.SizeExe - 1)
                    f.GetBuf Buf, Info.StartExe
                End If
            Else
                If f.LOF Then
                    ReDim Buf(f.LOF - 1)
                    f.GetBuf Buf
                End If
            End If
            f.FClose
        End If
    Else
        If Left$(nameScript, 1) = "$" Then nameScript = "<#Module=main>" & vbCrLf & Mid$(nameScript, 2) & vbCrLf & "<#Module>"
        Buf = ChrW$(65279) + nameScript
        nameScript = ""
    End If
    
    Code_Parse Buf, nameScript
    
    mf_TimeParse = timeGetTime - mf_TimeParse
    
    If mf_AsyncLoad Then
        Call SetTimer(frmScript.hWnd, 30001, mf_AsyncLoad, AddressOf Timer_Func)
    Else
        Timer_Func
    End If
End Sub


Function Code_Parse(Buf() As Byte, ByVal nameScript As String) As String
    Dim txtMain As String, txtForm As String, txtName As String, txtCode As String, txtLib As String, txtDLL As String
    Dim txtTmp As String, tmpForm As frmForm, PCD() As def_ParseCustom, cs As Collection, v As clsActiveScript
    Dim a As Long, mainRunMF As Integer, isMFC As Boolean
    
    On Error Resume Next
    
    If m_ArraySize(Buf) = 0 Then Exit Function
    
    '------------------------- Decompress Code --------------------------
    isMFC = DeCompressMF(Buf)
    
    '---------------------------- CodePage ------------------------------
    txtCode = vbCrLf + ToUnicode(Buf) + vbCrLf

    '--------------------------------------------------------------------
    REG.Global = True:  REG.IgnoreCase = True:  REG.MultiLine = True
    
    '-------------------------- Parse Data ------------------------------
    Call Parse_Data(txtCode)

    '--------------------------- Include Parse --------------------------
    Call Parse_Include(txtCode, "(?!res:)")
    Call Parse_Include(txtCode)
    
    '--------------------------- CScript Parse --------------------------
    Set cs = Parse_CScript(txtCode)
    
    '----------------- Symbol _ Parse & Comments Delete -----------------
    Call Parse_Preprocess(txtCode)

    '--------------------------- Directiv Parse -------------------------
    Call Parse_Directiv(txtCode)
    
    '--------------------------- LMF_Parser_Raw -------------------------
    For Each v In cs
        If ExistsMember(v.CodeObject, "LMF_Parser_Raw") Then txtCode = v.CodeObject.LMF_Parser_Raw(txtCode)
    Next
    
    '----------------------------- Lib Parse ----------------------------
    txtLib = Parse_Lib(txtCode)

    '----------------------------- DLL Parse ----------------------------
    txtDLL = Parse_DLL(txtCode)

    '-------------------------- Interface Parse -------------------------
    Call Parse_Interface(txtCode)

    '---------------------------- Types Parse ---------------------------
    Call Parse_Types(txtCode)

    '---------------------- RegExp Perl Style Parse ---------------------
    Call Parse_RegExp(txtCode)
    
    '--------------------------- LMF_Parser_Code ------------------------
    For Each v In cs
        If ExistsMember(v.CodeObject, "LMF_Parser_Code") Then txtCode = v.CodeObject.LMF_Parser_Code(txtCode)
    Next


    '----------------------------- DEBUG MODE ---------------------------
    'm_Str2File txtCode, "script" & mf_Counter & ".txt"
    'm_Str2File txtDLL, "dll" & mf_Counter & ".txt"
    '--------------------------------------------------------------------
    
    
    '---------------------------- Module Parse --------------------------
    If Parse_Custom(PCD, txtCode, "<#module", ">", "<#module>") Then
        For a = 0 To UBound(PCD)
            txtMain = txtMain + PCD(a).Data + vbCrLf
        Next
    End If
    
    '------------------------------ Modules Count -----------------------
    mf_Counter = mf_Counter + 1
    mainRunMF = mf_Counter
    
    '--------------------- Add Const to Main Module ---------------------
    If mainRunMF = 1 Then AddBaseMF txtMain
    
    '------------------------------ Name Parse --------------------------
    MDL(mainRunMF).Name = GetFileTitle(nameScript)
    MDL(mainRunMF).Path = GetDirectory(nameScript)
    
    '------------------------------ Form Parse --------------------------
    txtMain = "Private Const mf_IDM = " + CStr(mainRunMF) + vbCrLf + vbCrLf + txtMain
    txtMain = "Const mf_NameMod = """ + MDL(mainRunMF).Name + """" + vbCrLf + txtMain

    '--------------------------------------------------------------------
    If Parse_Custom(PCD, txtCode, "<#form=", ">", "<#form>") Then
        For a = 0 To UBound(PCD)
            mf_Counter = mf_Counter + 1
            
            txtName = PCD(a).Param
            txtForm = PCD(a).Data
            
            '-----------------------------------
            If mainRunMF > 1 Then
                txtTmp = "Dim " + txtName + vbCrLf + "Set " + txtName + " = " + "mf_forms_" + CStr(mf_Counter) + vbCrLf
                txtMain = txtTmp + txtMain
                txtForm = txtTmp + txtForm
            End If
            
            '-----------------------------------
            txtForm = "Const mf_NameMod = """ + txtName + """" + vbCrLf + vbCrLf + txtForm
            txtForm = "Const mf_NameLib = """ + MDL(mainRunMF).Name + """" + vbCrLf + txtForm
            txtForm = "Private Const mf_IDM = " + CStr(mf_Counter) + vbCrLf + txtForm
                
            '-----------------------------------
            If mainRunMF > 1 Then txtName = "mf_forms_" + CStr(mf_Counter)
            
            '------------------------------------
            txtForm = "Private This" + vbCrLf + "Set This = " + txtName + vbCrLf + txtForm
                
            '------------------------------------
            With MDL(mf_Counter)
                .Name = txtName
                .Path = ""
                .Code = txtForm
                .MFC = isMFC
                .Type = tpForm
            End With

            '------------------------------------
            Set tmpForm = New frmForm
            CAS.AddObject txtName, tmpForm
            Set tmpForm.CodeObject = CAS.AddModule(mf_Counter, txtForm)
            Set tmpForm = Nothing
        Next
    End If
    '-------------------------------------------------------------------------

    With MDL(mainRunMF)
        .Code = txtMain
        .MFC = isMFC
        .Type = tpModule
    End With
    
    '-------------------------------------------------------------------------
    If mainRunMF = 1 Then
        CAS.Objects.Add 1, ""
        CAS.AddCode txtMain
    Else
        CAS.AddObject MDL(mainRunMF).Name, CAS.AddModule(mainRunMF, txtMain)
    End If

    If Not mf_IsEnd Then
        CAS.AddCode txtDLL                  ' Add DLL Code
        Call Parse_AddLib(txtLib)           ' Add Lib Code
    End If

    Err.Clear
    
    Code_Parse = txtCode
End Function

Sub Parse_Data(txtCode As String)
    Call Parse_Resource(txtCode)
    Call Parse_VBNET(txtCode)
End Sub

Function Parse_Data_Mode(PCD As def_ParseCustom, Mode As String, Optional cd As clsDim, Optional ByVal bString As Boolean) As Variant
    Dim isBuf As Boolean, vExtra As Variant, tmpBuf() As Byte
    Static Base64 As New clsBase64
    
    If InStr(Mode, "base64") Then
        If Len(PCD.Data) Then
            If Not isBuf Then tmpBuf = Conv_W2A_Buf(PCD.Data)
            tmpBuf = Base64.Decode(tmpBuf, vExtra)
            If Not cd Is Nothing Then cd.Extra = vExtra
        End If
        isBuf = True
    End If
    
    If InStr(Mode, "zlib") Then
        If Len(PCD.Data) Then
            If Not isBuf Then tmpBuf = Conv_W2A_Buf(PCD.Data)
            DecompressData tmpBuf
        End If
        isBuf = True
    End If
    
    If InStr(Mode, "bin") Then
        If Len(PCD.Data) > 0 And isBuf = False Then tmpBuf = Conv_W2A_Buf(PCD.Data)
        isBuf = True
    End If
    
    If InStr(Mode, "null") Then
        If isBuf Then ReDim Preserve tmpBuf(m_ArraySize(tmpBuf)) Else PCD.Data = PCD.Data + Chr$(0)
    End If
    
    If isBuf Then
        If bString Then Parse_Data_Mode = Conv_A2W_Buf(tmpBuf) Else Parse_Data_Mode = tmpBuf
    Else
        Parse_Data_Mode = PCD.Data
    End If
End Function

Sub Parse_Resource(txtCode As String)
    Dim a As Long, REG1 As RegExp, Mts As MatchCollection, cd As clsDim, PCD() As def_ParseCustom
    
    If Not Parse_Custom(PCD, txtCode, "<#res ", "#>", "<#res#>", vbBinaryCompare) Then Exit Sub

    Set REG1 = New RegExp:      REG1.Global = True:      REG1.IgnoreCase = True
    REG1.Pattern = "^id=""([^""]+)""( +mode=([^ ]+))?"

    For a = 0 To UBound(PCD)
        If Len(PCD(a).Param) Then
            Set Mts = REG1.Execute(PCD(a).Param)
            
            If Mts.Count Then
                Set cd = New clsDim
                
                With cd
                    .ID = Mts(0).SubMatches(0)
                    .Mode = Mts(0).SubMatches(2)
                    .Data = Parse_Data_Mode(PCD(a), .Mode, cd)
                    SYS.Resource.Add cd, .ID
                End With
            End If
        End If
    Next
End Sub

Sub Parse_VBNET(txtCode As String)
    Dim a As Long, REG1 As RegExp, Mts As MatchCollection, Obj As Object, tmp As Variant, PCD() As def_ParseCustom

    On Error Resume Next
    
    If Not Parse_Custom(PCD, txtCode, "<#vbnet", "#>", "<#vbnet#>") Then Exit Sub
    
    Set REG1 = New RegExp:      REG1.Global = True:      REG1.IgnoreCase = True
    REG1.Pattern = "^(=([a-z0-9_]*))?( +noerror)?( +instance=""([^""]+)"")?( +start=""([^""]+)"")?( +lang=([a-z0-9_]+))?( +mode=([^ ]+))?"

    For a = 0 To UBound(PCD)
        Set Mts = REG1.Execute(PCD(a).Param)
        
        Set Obj = CreateObject("Atomix.VBNET")
        
        With Mts(0)
            If Len(.SubMatches(1)) > 0 Then CAS.AddObject .SubMatches(1), Obj

            If Len(PCD(a).Data) Then
                tmp = Parse_Data_Mode(PCD(a), .SubMatches(10), , True)
                If Len(.SubMatches(8)) > 0 Then tmp = Obj.Build(tmp, .SubMatches(8)) Else tmp = Obj.Build(tmp)
                If Len(tmp) > 0 And Len(.SubMatches(2)) = 0 Then MsgBox tmp, , "VBNET Error!":   m_Str2File CStr(tmp), "vbnet.log"
            
                If Len(.SubMatches(5)) > 0 Then
                    For Each tmp In Obj.Find(Trim$(.SubMatches(6)))
                        Obj.CallMethod tmp.Member, LMF
                    Next
                End If
                
                If Len(.SubMatches(3)) > 0 Then
                    For Each tmp In Split(.SubMatches(4), ",")
                        tmp = Split(tmp, "->")
                        If UBound(tmp) = 0 Then
                            CAS.AddObject Trim$(tmp(0)), Obj.CreateInstance(Trim$(tmp(0)))
                        Else
                            CAS.AddObject Trim$(tmp(1)), Obj.CreateInstance(Trim$(tmp(0)))
                        End If
                    Next
                End If
            End If
        End With
    Next
End Sub

Function Parse_CScript(txtCode As String) As Collection
    Dim a As Long, Obj As clsActiveScript, REG1 As RegExp, Mts As MatchCollection, PCD() As def_ParseCustom
    Static cntCScript As Long
    
    Set Parse_CScript = New Collection
    
    If Not Parse_Custom(PCD, txtCode, "<#script", "#>", "<#script#>") Then Exit Function
    
    If frmScript.CScript.Count = 0 Then cntCScript = 0

    Set REG1 = New RegExp:      REG1.Global = True:      REG1.IgnoreCase = True
    REG1.Pattern = "^=? *([a-z0-9_]*) *,? *([a-z0-9_]*) *,? *([a-z0-9_]*)"
    
    For a = 0 To UBound(PCD)
        If Len(PCD(a).Param) Then
            Set Mts = REG1.Execute(PCD(a).Param)
            
            If Mts.Count Then
                Set Obj = New clsActiveScript
                Set Obj.Parent = frmScript
                Obj.Name = "Custom"
                
                With Mts(0)
                    If Len(.SubMatches(1)) Then Obj.Language = .SubMatches(1) Else Obj.Language = "JavaScript"
                    Obj.Tag = .SubMatches(0)
                    Obj.AddObject .SubMatches(0), CAS.CodeObject
                    Obj.AddCode "mf_IDS = " & cntCScript & vbCrLf & PCD(a).Data
                    CAS.AddObject .SubMatches(0), Obj.CodeObject, Val(.SubMatches(2))
                End With
                
                Parse_CScript.Add Obj
                
                frmScript.CScript.Add Obj, CStr(cntCScript)
                cntCScript = cntCScript + 1
            End If
        End If
    Next
End Function

Function Parse_Template(txtCode As String, Optional ByVal fnPrint As String, Optional dmStart As String, Optional dmStop As String) As String
    Dim a As Long, isEval As Boolean, isInt As Boolean, fnBuffer As String, fnExec As String, out As String, tmp() As String, txt() As String

    If Len(dmStart) = 0 Then dmStart = "<?="
    If Len(dmStop) = 0 Then dmStop = "?>"
    If Len(fnPrint) = 0 Then fnPrint = "*Print*"
    
    If Left$(fnPrint, 1) = "*" Then isEval = True:   fnPrint = Mid$(fnPrint, 2)
    If Right$(fnPrint, 1) = "*" Then isInt = True:   fnPrint = Mid$(fnPrint, 1, Len(fnPrint) - 1)
    
    fnBuffer = fnPrint + "_Buffer":    fnExec = fnPrint + "_Execute"
    
    If isInt Then
        CAS.AddCode "Dim " + fnBuffer + vbCrLf + _
        "Sub " + fnPrint + "(mf_v1) : " + fnBuffer + " = " + fnBuffer + " & mf_v1 : End Sub" + vbCrLf + _
        "Function " + fnExec + "(mf_v1) : On Error Resume Next : " + fnBuffer + " = """" : Execute mf_v1 : " + fnExec + " = " + fnBuffer + " : End Function"
    End If

    If Len(txtCode) = 0 Then Exit Function

    txt = Split(txtCode, dmStart):    If isEval Then out = txt(0) Else out = Parse_Template_Sub(txt(0), fnPrint)

    For a = 1 To UBound(txt)
        tmp = Split(txt(a), dmStop)
        If isEval Then
            On Error Resume Next
            out = out & CallByName(CAS.CodeObject, fnExec, VbMethod, tmp(0))
            On Error GoTo 0
        Else
            out = out + tmp(0) + vbCrLf
        End If
        If UBound(tmp) > 0 Then If isEval Then out = out + tmp(1) Else out = out + Parse_Template_Sub(tmp(1), fnPrint)
    Next

    Parse_Template = out
End Function

Function Parse_Template_Sub(txtCode As String, fnPrint As String) As String
    Dim a As Long, uds As Long, t1 As String, t2 As String, t3 As String, txt() As String
    If Len(txtCode) = 0 Then Exit Function
    t1 = fnPrint + "(""":    t2 = """ + vbCrLf)":    t3 = fnPrint + "(vbCrLf)"
    txt = Split(txtCode, vbCrLf):    uds = UBound(txt)
    For a = 0 To uds
        If a = uds Then t2 = """)":    t3 = ""
        If Len(txt(a)) Then txt(a) = t1 + Replace$(txt(a), """", """""") + t2 Else txt(a) = t3
    Next
    If Len(txt(uds)) Then txt(uds) = txt(uds) & vbCrLf
    Parse_Template_Sub = Join(txt, vbCrLf)
End Function

Function Parse_Types_Sub(ByVal Mts As MatchCollection, wrapProp As String, ByVal sz As Long, vOffset As Long) As String
    Dim txt As String, pArg As String, pProp As String, nameVar As String, cntBound As Long, isBound As Boolean, sz2 As Long
    
    sz2 = sz:  sz = Abs(sz)
    
    With Mts(0)
        nameVar = .SubMatches(0)
        cntBound = Val(.SubMatches(2))
        isBound = Len(.SubMatches(2))
    End With
    
                        
    If isBound Then pArg = "Class___index" & IIF(sz > 1, " * " & sz, "") & " + "
    pArg = pArg & "Class___Offset + " & vOffset & IIF(wrapProp = "PString", ", " & sz2, "")
    pProp = "Class___Wrapper." & wrapProp & "(" & pArg & ")"
    
    txt = txt & "  Property Get " & nameVar & "(" & IIF(isBound, "Class___index", "") & ") : " & nameVar & " = " & pProp & " : End Property" & vbCrLf
    txt = txt & "  Property Let " & nameVar & "(" & IIF(isBound, "Class___index, ", "") & "Class___value) : " & pProp & " = Class___value : End Property" & vbCrLf
    
    vOffset = vOffset + sz * (cntBound + 1)
    
    Parse_Types_Sub = vbCrLf & txt
End Function

Sub Parse_Types(txtCode As String)
    Dim cntBound As Long, isBound As Boolean, nameVar As String, cntString As Long, nameType As String
    Dim aDim As String, aPtr As String, oTxt As String, iTxt As String, vTxt As String, rTxt As String
    Dim a As Long, Mts As MatchCollection, mts1 As MatchCollection, txt() As String
    Dim vOffset As Long, offsetUnion As Long
    
    REG.Global = False
    
    Do
        aDim = "": aPtr = "": vTxt = "": iTxt = "": oTxt = "": vOffset = 0
        
        REG.Pattern = "\n[ \t\v]*(private |public )?type +([a-z0-9_]+)([\d\D]+?)\n[ \t\v]*end type"
        
        Set Mts = REG.Execute(txtCode)
        
        If Mts.Count Then
            With Mts(0)
            
            offsetUnion = -1
            
            txt = Split(.SubMatches(2), vbCrLf)
            
            For a = 0 To UBound(txt)
                If Len(txt(a)) > 0 Then
                
                    REG.Pattern = "^[ \t\v]*([a-z0-9_]+)(\((\d+)\))? +as +([^ \r]+)( +\* +(\-?\d+))?"
                    
                    Set mts1 = REG.Execute(txt(a))
                    
                    If mts1.Count > 0 Then
                        With mts1(0)
                            nameVar = .SubMatches(0)
                            nameType = .SubMatches(3)
                            cntBound = Val(.SubMatches(2))
                            isBound = Len(.SubMatches(2))
                            cntString = Val(.SubMatches(5))
                        End With
                        
                        Select Case LCase$(nameType)
                            Case "@"
                                Select Case LCase$(nameVar)
                                    Case "union"
                                        If vOffset > offsetUnion Then offsetUnion = vOffset
                                        vOffset = cntString
                                        If vOffset < 0 Then vOffset = 0
                                        
                                    Case "offset"
                                        If cntString < 0 Then vOffset = vOffset + cntString Else vOffset = cntString
                                        If vOffset < 0 Then vOffset = 0
                                End Select
                                
                            Case "string"
                                If cntString = 0 Then
                                    'nonfixed string
                                    If isBound Then
                                        aDim = aDim & "  Dim " & nameVar & "(" & cntBound & ")" & vbCrLf
                                        aPtr = aPtr & "    mf_t_ofs = " & vOffset & vbCrLf
                                        aPtr = aPtr & "    For mf_v1 = 0 To " & cntBound & vbCrLf
                                        aPtr = aPtr & "      Class___Wrapper.PLong(Class___Offset + mf_t_ofs) = StrPtr(" & nameVar & "(mf_v1), True)" & vbCrLf
                                        aPtr = aPtr & "      mf_t_ofs = mf_t_ofs + 4" & vbCrLf
                                        aPtr = aPtr & "    Next" & vbCrLf
                                    Else
                                        aDim = aDim & "  Dim " & nameVar & vbCrLf
                                        aPtr = aPtr & "    Class___Wrapper.PLong(Class___Offset + " & vOffset & ") = StrPtr(" & nameVar & ", True)" & vbCrLf
                                    End If
                                    vOffset = vOffset + 4 * (cntBound + 1)
                                Else
                                    'fixed string
                                    vTxt = vTxt & Parse_Types_Sub(mts1, "PString", cntString, vOffset)
                                End If
                            
                            Case "byte"
                                vTxt = vTxt & Parse_Types_Sub(mts1, "PByte", 1, vOffset)
                                
                            Case "word"
                                vTxt = vTxt & Parse_Types_Sub(mts1, "PWord", 2, vOffset)
                                
                            Case "integer"
                                vTxt = vTxt & Parse_Types_Sub(mts1, "PInteger", 2, vOffset)
                                
                            Case "boolean"
                                vTxt = vTxt & Parse_Types_Sub(mts1, "PBoolean", 2, vOffset)
                                
                            Case "long"
                                vTxt = vTxt & Parse_Types_Sub(mts1, "PLong", 4, vOffset)
                                
                            Case "single"
                                vTxt = vTxt & Parse_Types_Sub(mts1, "PSingle", 4, vOffset)
                                
                            Case "double"
                                vTxt = vTxt & Parse_Types_Sub(mts1, "PDouble", 8, vOffset)
                                
                            Case "currency"
                                vTxt = vTxt & Parse_Types_Sub(mts1, "PCurrency", 8, vOffset)
                                
                            Case Else
                                If isBound Then
                                    aDim = aDim & "  Dim " & nameVar & "(" & cntBound & ")" & vbCrLf
                                    iTxt = iTxt & "    mf_t_ofs = " & vOffset & vbCrLf
                                    iTxt = iTxt & "    For mf_v1 = 0 To " & cntBound & vbCrLf
                                    iTxt = iTxt & "      Set " & nameVar & "(mf_v1) = New " & nameType & vbCrLf
                                    iTxt = iTxt & "      Class___Enum.Add Array(" & nameVar & "(mf_v1), mf_t_ofs)" & vbCrLf
                                    iTxt = iTxt & "      mf_t_ofs = mf_t_ofs + " & Types(nameType) & vbCrLf
                                    iTxt = iTxt & "    Next" & vbCrLf
                                Else
                                    aDim = aDim & "  Dim " & nameVar & vbCrLf
                                    iTxt = iTxt & "    Set " & nameVar & " = New " & nameType & vbCrLf
                                    iTxt = iTxt & "    Class___Enum.Add Array(" & nameVar & ", " & vOffset & ")" & vbCrLf
                                End If
                                vOffset = vOffset + Types(nameType) * (cntBound + 1)
                        End Select
                    Else
                        rTxt = rTxt + txt(a) + vbCrLf
                    End If
                    
                End If
            Next
            
            If offsetUnion <> -1 Then If vOffset < offsetUnion Then vOffset = offsetUnion

            oTxt = oTxt + "Class " + .SubMatches(1) + vbCrLf
            oTxt = oTxt + "  Dim Class___Wrapper, Class___Enum, Class___Offset" + vbCrLf
            oTxt = oTxt + aDim + rTxt + vTxt + vbCrLf
            oTxt = oTxt + "  Private Sub Class_Initialize()" + vbCrLf
            oTxt = oTxt + "    Set Class___Enum = Sys.NewCollection" + vbCrLf
            oTxt = oTxt + iTxt
            oTxt = oTxt + "  End Sub" + vbCrLf + vbCrLf
            oTxt = oTxt + "  Public Property Get Class___Size() : Class___Size = " + CStr(vOffset) + " : End Property" + vbCrLf + vbCrLf
            oTxt = oTxt + "  Public Default Property Get Class___Ptr()" + vbCrLf
            oTxt = oTxt + "    Class___Ptr = Class___Wrapper.Ptr(True) + Class___Offset" + vbCrLf
            oTxt = oTxt + aPtr
            oTxt = oTxt + "  End Property" + vbCrLf
            oTxt = oTxt + "  Public Property Let Class___Ptr(Class___value)" + vbCrLf
            oTxt = oTxt + "    Class___Wrapper.Ptr = Class___value" + vbCrLf
            oTxt = oTxt + "  End Property" + vbCrLf
            oTxt = oTxt + "End Class" + vbCrLf

            txtCode = Left$(txtCode, .FirstIndex + 1) + oTxt + Right$(txtCode, Len(txtCode) - .FirstIndex - .Length - 2)
        
            Types.Add vOffset, .SubMatches(1)
            
            End With
        End If
    Loop Until Mts.Count = 0
    
    
    Do
        oTxt = ""
        
        REG.Pattern = "\n[ \t\v]*Dim +([a-z0-9_]+)(\(([a-z0-9_]+)\))? +as +(new )?([a-z0-9_]+)(\([^\r]*\))?"
        Set Mts = REG.Execute(txtCode)

        If Mts.Count Then
            With Mts(0)
            
            If Len(.SubMatches(1)) Then
                If Len(.SubMatches(3)) = 0 Then oTxt = oTxt + "ReDim " + .SubMatches(0) + .SubMatches(1) + vbCrLf
                oTxt = oTxt + "For mf_d_vc = 0 To " + CStr(.SubMatches(2)) + vbCrLf
                oTxt = oTxt + "   Set " + Trim$(CStr(.SubMatches(0))) + "(mf_d_vc) = InitType(New " + .SubMatches(4) + ")" + vbCrLf
                If Len(.SubMatches(5)) Then oTxt = oTxt + "   Call " + .SubMatches(0) + ".Class_New" + .SubMatches(5) + vbCrLf
                oTxt = oTxt + "Next"
            Else
                If Len(.SubMatches(3)) = 0 Then oTxt = "Dim " + .SubMatches(0) + vbCrLf
                oTxt = oTxt + "Set " + .SubMatches(0) + " = InitType(New " + .SubMatches(4) + ")"
                If Len(.SubMatches(5)) Then oTxt = oTxt + vbCrLf + "Call " + .SubMatches(0) + ".Class_New" + .SubMatches(5)
            End If
            
            txtCode = REG.Replace(txtCode, vbLf + oTxt)
            
            End With
        End If
    Loop Until Mts.Count = 0

    REG.Global = True
End Sub


Sub Parse_Interface(txtCode As String)
    Dim Mts As MatchCollection, mts1 As MatchCollection, a As Long, b As Long, cnt As Long, mbr As Long
    Dim txt As String, txtVars As String
    
    Do
        txt = ""
        
        REG.Global = False
        REG.Pattern = "\ninterface +([a-z0-9_]+) *\[ *(\{[a-z0-9\-]*\}) *, *(\{[a-z0-9\-]*\}) *, *(\d+) *\] *= *(.+)"
        Set Mts = REG.Execute(txtCode)
    
        If Mts.Count > 0 Then
            With Mts(0)
            
            REG.Global = True
            REG.Pattern = "(\d+:)?([a-z0-9_]+)(\((\d+)\))?"
            Set mts1 = REG.Execute(.SubMatches(4))
            
            txt = txt + "Class " + .SubMatches(0) + vbCrLf + vbCrLf
            txt = txt + "Dim ifc_FCP, ifc_Obj, ifc_Arg" + vbCrLf + vbCrLf

            For a = 0 To mts1.Count - 1
                With mts1(a)
                    cnt = Val(.SubMatches(3))
                    txt = txt + "Function " + .SubMatches(1) + "("
                    
                    txtVars = ""
                    For b = 1 To cnt
                        txtVars = txtVars + "ifc_var" + CStr(b) + ", "
                    Next
                    If cnt > 0 Then txtVars = Left$(txtVars, Len(txtVars) - 2)
                    
                    If Len(.SubMatches(0)) > 0 Then mbr = Val(.SubMatches(0))
                    
                    txt = txt + txtVars + ")" + vbCrLf
                    txt = txt + .SubMatches(1) + " = ifc_FCP(" + CStr(mbr) + IIF(Len(txtVars) > 0, ", " + txtVars, "") + ")" + vbCrLf
                    txt = txt + "End Function" + vbCrLf + vbCrLf
                    
                    mbr = mbr + 1
                End With
            Next
            
            txt = txt + "Private Sub Class_Initialize" + vbCrLf
            txtVars = """" + .SubMatches(1) + """, """ + .SubMatches(2) + """, " + CStr(.SubMatches(3))
            txt = txt + "ifc_Arg = Array(" + txtVars + ")" + vbCrLf
            If .SubMatches(3) = 0 Then
                txt = txt + "Set ifc_FCP = Sys.NewFCP" + vbCrLf
            Else
                txt = txt + "ifc_Obj = Sys.Com.CreateInstance(" + txtVars + ")" + vbCrLf
                txt = txt + "If ifc_Obj <> 0 Then Set ifc_FCP = Sys.NewFCP(ifc_Obj)" + vbCrLf
            End If
            txt = txt + "End Sub" + vbCrLf + vbCrLf
            
            txt = txt + "Private Sub Class_Terminate" + vbCrLf
            txt = txt + "If ifc_Obj <> 0 then Call ifc_FCP(2)" + vbCrLf
            txt = txt + "End Sub" + vbCrLf + vbCrLf
            
            txt = txt + "End Class " + vbCrLf + vbCrLf
            
            txtCode = Left$(txtCode, .FirstIndex + 1) + txt + Right$(txtCode, Len(txtCode) - .FirstIndex - .Length - 2)
            
            End With
        End If
    Loop Until Mts.Count = 0

    REG.Global = True
End Sub

Sub Parse_RegExp(txtCode As String)
    Dim p0 As String, p1 As String, p2 As String, p3 As String, pm As String, Mts As MatchCollection
    Dim fi As Boolean, fg As Boolean, fr As Boolean, fe As Boolean, fm As Boolean

    REG.Global = False
    
    Do
        REG.Pattern = "([a-z0-9_\.\(\)]+) *=~ *([igem]*)\/(.*?(\\\\|[^\\]))?\/(.*)\/"
        Set Mts = REG.Execute(txtCode)
    
        If Mts.Count Then
            With Mts(0)
                p0 = .SubMatches(0):    p1 = LCase$(.SubMatches(1)):    p2 = .SubMatches(2):    p3 = .SubMatches(4)
                
                fr = CBool(Len(p3))
                fi = CBool(InStr(p1, "i"))
                fg = CBool(InStr(p1, "g"))
                fe = CBool(InStr(p1, "e"))
                fm = CBool(InStr(p1, "m"))
                
                If fm Then pm = "" Else pm = """":  If p3 = vbTab Then p3 = ""
    
                If fr Then
                    txtCode = Replace$(txtCode, .value, p0 + " = sys.rxp.replace(" + p0 + ", " + pm + p2 + pm + ", " + pm + p3 + pm + ", " + CStr(fi) + ", " + CStr(fg) + ")")
                Else
                    txtCode = Replace$(txtCode, .value, "sys.rxp." + IIF(fe, "execute", "test") + "(" + p0 + ", " + pm + p2 + pm + ", " + CStr(fi) + ", " + CStr(fg) + ")")
                End If
            End With
        End If
    Loop Until Mts.Count = 0
    
    Do
        REG.Pattern = "\$\$(\d+)"
        Set Mts = REG.Execute(txtCode)
        If Mts.Count Then txtCode = REG.Replace(txtCode, "sys.rxp.matches(" + CStr(Val(Mts(0).SubMatches(0)) - 1) + ")")
    Loop Until Mts.Count = 0
    
    REG.Global = True
End Sub

Function Parse_DLL(txtCode As String) As String
    Dim a As Long, b As Long, r As Boolean, v1 As String, v2 As String, txt As String, nAlias As String
    Dim tFunc As String, nFunc As String, REG1 As RegExp, Mts As MatchCollection, mt As MatchCollection
    
    REG.Pattern = "\n(private |public )?declare +(function|sub) +([^ ]+) +(Lib +""([^""]+)"" +)?(alias +""([^""]+)"" +)?\(([^\)]*)\)( as )?([a-z0-9_]*)"
    Set Mts = REG.Execute(txtCode)

    If Mts.Count > 0 Then
        Set REG1 = New RegExp
        REG1.Global = True
        REG1.IgnoreCase = True
        REG1.Pattern = " *(byval|byref)? *([a-z0-9_]+) *(as)? *([a-z0-9_]*) *,?"
    End If

    For a = 0 To Mts.Count - 1
        With Mts(a)
            Set mt = REG1.Execute(.SubMatches(7))
            
            tFunc = .SubMatches(1)
            nFunc = .SubMatches(2)
            nAlias = .SubMatches(6)
            If LenB(nAlias) = 0 Then nAlias = nFunc
            
            txt = txt & .SubMatches(0) & tFunc & " " & nFunc & "("
            
            For b = 0 To mt.Count - 1
                With mt(b)
                    If LCase$(.SubMatches(3)) <> "string" Then txt = txt & .SubMatches(0)
                    txt = txt & " mf_v" & b & ", "
                End With
            Next
            If mt.Count > 0 Then txt = Left$(txt, Len(txt) - 2)
            
            txt = txt & ")" & vbCrLf
            
            If LCase$(tFunc) = "sub" Then txt = txt & "Call" Else txt = txt & nFunc & " ="
            txt = txt & " DllCall(""" & .SubMatches(4) & """,""" & nAlias & """"
            
            For b = 0 To mt.Count - 1
                With mt(b)
                    v1 = "":    v2 = ""
                    r = (LCase$(.SubMatches(0)) <> "byval")
                    Select Case LCase$(.SubMatches(3))
                        Case "long":        If r Then v1 = "VarPtr(":   v2 = ", True) + 8"
                        Case "single":      If Not r Then v1 = "Array(CSng(":   v2 = "))"
                        Case "double":      If Not r Then v1 = "Array(CDbl(":   v2 = "))"
                    End Select
                    txt = txt & ", " & v1 & "mf_v" & b & v2
                End With
            Next
            
            txt = txt & ")" & vbCrLf
            
            If LCase$(.SubMatches(9)) = "string" Then txt = txt & nFunc & " = sys.conv.ptr2str(" & nFunc & ")" & vbCrLf
                
            txt = txt & "End " & tFunc & vbCrLf & vbCrLf
        End With
    Next
    
    txtCode = REG.Replace(txtCode, vbLf)
    
    Parse_DLL = txt
End Function

Function Parse_Lib(txtCode As String) As String
    Dim a As Long, txt As String, Mts As MatchCollection
    
    REG.Pattern = "\n[ \t\v]*#Lib +""([^\""]+)"""
    Set Mts = REG.Execute(txtCode)
    For a = 0 To Mts.Count - 1
        txt = txt + Parse_MPath(Mts(a).SubMatches(0)) + vbCr
    Next
    txtCode = REG.Replace(txtCode, vbLf)
    
    Parse_Lib = txt
End Function

Sub Parse_Directiv(txtCode As String)
    Dim a As Long, Mts As MatchCollection
    
    With REG
        .Pattern = "\n<#--([\w\d_]+)( *= *""([^""]+)"")? *>"
        Set Mts = .Execute(txtCode)
        
        For a = 0 To Mts.Count - 1
            Select Case LCase$(Mts(a).SubMatches(0))
                Case "addrus":          Call Parse_Modify(txtCode)
                Case "develop":         EMailDevelop = Mts(0).SubMatches(2)
                Case "asyncload":       mf_AsyncLoad = CLng(Mts(0).SubMatches(2))
            End Select
        Next
    End With
End Sub

Sub Parse_Preprocess(txtCode As String)
    With REG
        .Pattern = " _[ \x09]*\r\n"
        txtCode = .Replace(txtCode, " ")
    End With
End Sub

Sub Parse_AddLib(txtLib As String)
    Dim a As Long, Mts As MatchCollection, mt As Match, isLibAdd As Boolean
    Dim txt As String, tmp As String, Buf() As Byte
    
    If LenB(txtLib) = 0 Then Exit Sub

    REG.Pattern = "[^\r]+"
    Set Mts = REG.Execute(txtLib)
    
    For Each mt In Mts
        txt = LCase$(mt.value)
    
        If IsFileExt(txt, Array(SYS.Path), Array(".mf")) Then
            tmp = GetFileTitle(txt)

            isLibAdd = True
            For a = 1 To 255
               If Len(MDL(a).Name) = 0 Then Exit For
               If tmp = LCase$(MDL(a).Name) Then isLibAdd = False:  Exit For
            Next
            
            If isLibAdd Then
                m_File2Buf Buf, txt
                Code_Parse Buf, txt
            End If
        End If
    Next
End Sub

Sub Parse_Include(txtCode As String, Optional ByVal noFind As String)
    Dim txt As String, tmpBuf() As Byte, dat As Variant, RX As New clsRXP, Mts As MatchCollection
    
    REG.Global = False
    REG.Pattern = "\n[ \t\v]*#Include +""" + noFind + "([^""]+)"""
    
    Do
        Set Mts = REG.Execute(txtCode)
        
        If Mts.Count <> 0 Then
            With Mts(0)
                txt = "":    dat = Empty
                
                If RX.Test(.SubMatches(0), "^([a-z]+:\/\/)?(.+)") Then
                    tmpBuf = SYS.Content(.SubMatches(0), False)
                    DeCompressMF tmpBuf
                    dat = ToUnicode(tmpBuf)
                    
                    If Not IsEmpty(dat) Then
                        txt = vbCrLf + dat + vbCrLf
                        Call Parse_Data(txt)
                    End If
                End If
                txtCode = Left$(txtCode, .FirstIndex + 1) + txt + Right$(txtCode, Len(txtCode) - .FirstIndex - .Length)
            End With
        End If
    Loop Until Mts.Count = 0
    
    REG.Global = True
End Sub

Function Parse_Modify(txtCode As String, Optional txtConv As Variant, Optional ByVal Flags As Long) As String
    Dim a As Long, b As Long, st As Long, Fnd As String, Rep As String, txt() As String, REG1 As RegExp
    
    Set REG1 = New RegExp:      REG1.Global = True:      REG1.IgnoreCase = True
    
    If IsMissing(txtConv) Then
        txtConv = Array("Функци(я|и)", "Function", "Процедур(а|ы)", "Sub", "Вызвать", "Call", "Константа", "Const", _
        "Переменная", "Dim", "Переопределить", "Redim", "Комментарий", "Rem", "Конец Цикла", "Loop", "Цикл", "Do", _
        "Конец Пока", "Wend", "Пока", "While", "Для", "For", "Перебора", "Each", "Дальше", "Next", "Если", "If", _
        "Тогда", "Then", "Иначе", "Else", "ИначеЕсли", "ElseIf", "Свойство", "Property", "Присвоить", "Let", _
        "Установить", "Set", "Получить", "Get", "Выбор", "Select", "Условие", "Case", "По условию", "Until", _
        "Связать", "With", "Истина", "True", "Ложь", "False", "Нуль", "Null", "Пусто", "Empty", "Стереть", "Erase", _
        "Очистить", "Nothing", "Выход", "Exit", "Есть", "Is", "До", "To", "Шаг", "Step", "Частн(ая|ое)", "Private", _
        "Публичн(ая|ое)", "Public", "Новый", "New", "Как", "As", "Тип", "Type", "Конец", "End", _
        "Сообщение", "MsgBox", "СоздатьОбъект", "CreateObject", "ПолучитьОбъект", "GetObject", _
        "а", "a", "б", "b", "в", "v", "г", "g", "д", "d", "е", "e", "ё", "jo", "ж", "zh", "з", "z", "и", "i", "й", "j", _
        "к", "k", "л", "l", "м", "m", "н", "n", "о", "o", "п", "p", "р", "r", "с", "s", "т", "t", "у", "u", "ф", "f", _
        "х", "x", "ц", "c", "ч", "ch", "ш", "sh", "щ", "sch", "ъ", "qi", "ы", "y", "ь", "qu", "э", "e", "ю", "yu", "я", "ya")
    End If
    
    If IsArray(txtConv) Then
        If Flags = -1 Then
            For a = 0 To UBound(txtConv) - 1 Step 2
               REG1.Pattern = txtConv(a)
               txtCode = REG1.Replace(txtCode, txtConv(a + 1))
            Next
        ElseIf Flags < -1 Then
            Flags = Abs(Flags) - 2
            For a = 0 To UBound(txtConv) - 1 Step 2
               txtCode = Replace$(txtCode, txtConv(a), txtConv(a + 1), , , Flags)
            Next
        Else
            st = Flags And 1
            Flags = (Flags And &HFE) / 2 - 1

            txt = Split(txtCode, """")

            If Flags = -1 Then
                For b = 0 To UBound(txtConv) - 1 Step 2
                    REG1.Pattern = txtConv(b)
                    Rep = txtConv(b + 1)
                    For a = st To UBound(txt) Step 2
                        If Len(txt(a)) Then txt(a) = REG1.Replace(txt(a), Rep)
                    Next
                Next
            Else
                For b = 0 To UBound(txtConv) - 1 Step 2
                    Fnd = txtConv(b)
                    Rep = txtConv(b + 1)
                    For a = st To UBound(txt) Step 2
                        If Len(txt(a)) Then txt(a) = Replace$(txt(a), Fnd, Rep, , , Flags)
                    Next
                Next
            End If

            txtCode = Join(txt, """")
        End If
    End If

    Parse_Modify = txtCode
End Function

Function Parse_MPath(ByVal MPath As String) As String
    Dim clsReg As New clsRegistry, REG1 As RegExp, Mts As MatchCollection, isFind As Boolean
    Dim txt As String, Arg As String, tmp As String, a As Long, idx As Long
    
    Set REG1 = New RegExp:    REG1.Global = True:    REG1.IgnoreCase = True:    REG1.Pattern = "%((\w+?_)?([^%]+))%"
    Set Mts = REG1.Execute(MPath)
    
    For a = 0 To Mts.Count - 1
        txt = Mts(a).value
        
        isFind = False
        
        Select Case LCase$(Mts(a).SubMatches(1))
        
            Case "mf_"
                Arg = LCase$(Mts(a).SubMatches(2))
                If LCase$(Left$(Arg, 3)) = "rnd" Then
                    idx = Val(Mid$(Arg, 4))
                    MPath = Replace$(MPath, txt, IIF(idx, GenTempStr(idx), GenTempStr))
                Else
                    MPath = Replace$(MPath, txt, SYS.Path(Arg, False))
                End If
                isFind = True

            Case "env_"
                Arg = Mts(a).SubMatches(2)
                idx = InStr(Arg, "*")
                If idx Then
                    tmp = Mid$(Arg, idx + 1)
                    Arg = Left$(Arg, idx - 1)
                    SetEnvironmentVariableW StrPtr(Arg), StrPtr(tmp)
                    MPath = Replace$(MPath, txt, "")
                Else
                    tmp = String$(32767, 0)
                    idx = GetEnvironmentVariableW(StrPtr(Arg), StrPtr(tmp), Len(tmp))
                    tmp = Left$(tmp, idx)
                    MPath = Replace$(MPath, txt, tmp)
                End If
                isFind = True
                
            Case "sfp_"
                idx = Val(Mts(a).SubMatches(2))
                MPath = Replace$(MPath, txt, SpecialFolderPath(idx))
                isFind = True

            Case "reg_"
                Arg = Mts(a).SubMatches(2)
                MPath = Replace$(MPath, txt, clsReg.RegRead(Arg))
                isFind = True
                
        End Select
        
        If isFind = False Then MPath = Replace$(MPath, txt, GEV(Mid$(txt, 2, Len(txt) - 2)))
    Next
    
    Parse_MPath = MPath
End Function

Function Parse_Custom(PCD() As def_ParseCustom, Buf As String, ByVal cst1 As String, ByVal cst2 As String, ByVal cst3 As String, Optional ByVal Compare As VbCompareMethod = vbTextCompare) As Boolean
    Dim num As Long, rc As Long, rc2 As Long, rc3 As Long, p As Long
    
    cst1 = vbLf + cst1:     cst2 = cst2 + vbCr:     cst3 = vbLf + cst3 + vbCr
    num = -1:    rc = 1
    
    Do
        rc = InStr(rc, Buf, cst1, Compare)
        
        If rc Then
            rc2 = InStr(rc + Len(cst1), Buf, cst2, Compare)
            If rc2 = 0 Then Exit Function
            
            num = num + 1
            ReDim Preserve PCD(num)
            
            p = rc + Len(cst1)
            PCD(num).Param = Mid$(Buf, p, rc2 - p)
            
            rc3 = InStr(rc2 + Len(cst2), Buf, cst3, Compare)
            If rc3 = 0 Then Exit Function
            
            p = rc2 + Len(cst2) + 1
            If (rc3 - p) > -1 Then PCD(num).Data = Mid$(Buf, p, rc3 - p - 1)
            Buf = Left$(Buf, rc) + Right$(Buf, Len(Buf) - rc3 - Len(cst3))
        End If
    Loop Until rc = 0
    
    Parse_Custom = (num > -1)
End Function

Sub Script_Init()
    mf_IsEnd = False
    
    frmScript.Show
    
    If CAS Is Nothing Then
        Set CAS = New clsActiveScript
        Set CAS.Parent = frmScript
        CAS.Name = "Main"
    End If

    If SYS Is Nothing Then
        Set SYS = New clsSys
        CAS.AddObject "Sys", SYS
        CAS.AddObject "Shd", SYS.SHD, True
        Set Types = New clsHash
        EMailDevelop = mf_EMailDevelop
    End If
    
    SetDllDirectoryW StrPtr(CurDir)
End Sub

Sub Script_End()
    Dim Frm As Form

    On Error Resume Next

    For Each Frm In Forms
        Unload Frm
        Set Frm = Nothing
    Next
                
    Erase MDL:    mf_Counter = 0:     mf_AsyncLoad = 0:     mf_NoError = False

    If Not CAS Is Nothing Then CAS.Reset
    Set CAS = Nothing
    Set SYS = Nothing
    If App.StartMode = vbSModeStandalone Then Set LMF = Nothing
End Sub

Sub Script_EXE()
    Dim f As New clsFileAPI, ofs As Long, sz As Long, lh As Long, txt As String

    lh = Len(mf_Hdr):      txt = String$(lh, 0):      Info.IsExe = False
      
    If f.FOpen(GetAppPath(True), OPEN_EXISTING, GENERIC_READ) = INVALID_HANDLE Then Exit Sub
    
    ofs = f.LOF:        f.GetMem VarPtr(sz), 4, ofs - 3:        ofs = ofs - sz - 3

    If (ofs > 0) And (ofs < (f.LOF - lh)) Then
        f.GetStr txt, ofs
        If txt = mf_Hdr Then Info.IsExe = True:    Info.StartExe = f.Pos:    Info.SizeExe = f.LOF - Info.StartExe - 3
    End If
    
    f.FClose
End Sub

Sub AddBaseMF(txtCode As String)
    Dim txt As String
    
    txt = "Const vbSrcCopy = 13369376 : Const vbSrcAnd = 8913094 : Const vbSrcPaint = 15597702 : Const vbSrcInvert = 6684742" + vbCrLf + _
          "Const vbUnicode = 64 : Const vbFromUnicode = 128 : Const vbLowerCase = 2 : Const vbUpperCase = 1" + vbCrLf + _
          "Const vbSHA1 = 32772 : Const vbSHA256 = 32780 : Const vbSHA512 = 32782" + vbCrLf + _
          "Const vbMethod = 1 : Const vbGet = 2 : Const vbFunc = 3 : Const vbLet = 4 : Const vbSet = 8 : Const vbModal = 1" + vbCrLf + _
          "Sub Unload(mf_v1) : Sys.Ext.VB_Unload mf_v1 : End Sub" + vbCrLf + _
          "Sub DoEvents() : DoEvents2 : End Sub" + vbCrLf + vbCrLf
        
    txtCode = txt + txtCode
End Sub

Function CompressMF(ByVal fName As String, Optional VExt As Variant, Optional ByVal Packer As Long = CMS_FORMAT_ZLIB) As Boolean
    Dim a As Long, n As Long, ofs As Long, Buf() As Byte, f As New clsFileAPI
    Dim MFHC As def_HeaderCompile, HExt() As def_HeaderExt
    
    If IsFile(fName) = False Then Exit Function
    
    If f.FOpen(fName) = INVALID_HANDLE Then Exit Function
    f.GetMem VarPtr(MFHC), Len(MFHC)
    f.FClose
    
    If MFHC.Signature = mf_Sign Then Exit Function
    If Not m_File2Buf(Buf, fName) Then Exit Function
    
    FileKill fName
    
    '-------------------------------------------------
    If CompressData(Buf(), Packer) < 0 Then Exit Function
    
    '-------------------------------------------------
    ReDim HExt(0)

    With MFHC
        .Signature = mf_Sign
        .Length = Len(MFHC)
        .Packer = Packer
        .Reserved = 0
        .VerMajor = App.Major
        .VerMinor = App.Minor
        .VerBuild = App.Revision
        .HExtCount = m_ArraySize(VExt) \ 2
        .HExtOffset = .Length
        .DataOffset = .HExtOffset + .HExtCount * Len(HExt(0))
        .DataSize = UBound(Buf) + 1
        
        If .HExtCount Then
            ReDim HExt(.HExtCount - 1)
            ofs = .DataOffset + .DataSize
            
            For a = 0 To .HExtCount - 1
                n = a * 2 + 1
                HExt(a).HeaderID = CLng(VExt(n - 1))
                HExt(a).DataOffset = ofs
                If VarType(VExt(n)) = vbString Then VExt(n) = Conv_W2A_Buf(CStr(VExt(n)))
                If IsArray(VExt(n)) Then HExt(a).DataSize = m_ArraySize(VExt(n))
                ofs = ofs + HExt(a).DataSize
            Next
        End If
    End With
    
    If f.FOpen(fName) = INVALID_HANDLE Then Exit Function
    f.PutMem VarPtr(MFHC), Len(MFHC)
    If MFHC.HExtCount Then f.PutMem VarPtr(HExt(0)), MFHC.HExtCount * Len(HExt(0))
    f.PutBuf Buf
    For a = 0 To MFHC.HExtCount - 1
        If HExt(a).DataSize Then Buf = VExt(a * 2 + 1):   f.PutBuf Buf
    Next
    f.FClose
    
    CompressMF = True
End Function

Function DeCompressMF(tmpBuf() As Byte, Optional VExt As Variant, Optional ByVal bString As Boolean) As Boolean
    Dim a As Long, MFHC As def_HeaderCompile, HExt() As def_HeaderExt, Buf() As Byte
    
    If m_ArraySize(tmpBuf) = 0 Then Exit Function
    
    CopyMemory MFHC, tmpBuf(0), Len(MFHC)
    If MFHC.Signature <> mf_Sign Then Exit Function
    
    With MFHC
        If .HExtCount Then
            ReDim VExt(.HExtCount * 2 - 1)
            ReDim HExt(.HExtCount - 1)
            CopyMemory HExt(0), tmpBuf(MFHC.HExtOffset), .HExtCount * Len(HExt(0))
            
            For a = 0 To .HExtCount - 1
                VExt(a * 2) = HExt(a).HeaderID
                If HExt(a).DataSize Then
                    ReDim Buf(HExt(a).DataSize - 1)
                    CopyMemory Buf(0), tmpBuf(HExt(a).DataOffset), HExt(a).DataSize
                    If bString Then VExt(a * 2 + 1) = Conv_A2W_Buf(Buf) Else VExt(a * 2 + 1) = Buf
                End If
            Next
        End If

        CopyMemory tmpBuf(0), tmpBuf(.DataOffset), .DataSize
        ReDim Preserve tmpBuf(.DataSize - 1)
    End With
    
    '-------------------------------------------------
    If DecompressData(tmpBuf(), MFHC.Packer) < 0 Then Script_End Else DeCompressMF = True
End Function

Sub BackupMF(ByVal nameMF As String)
    Dim f As New clsFileAPI, MFHC As def_HeaderCompile
    
    If f.FOpen(nameMF) = INVALID_HANDLE Then Exit Sub
    f.GetMem VarPtr(MFHC), Len(MFHC)
    f.FClose
    
    If MFHC.Signature <> mf_Sign Then FileCopy nameMF, nameMF + ".bak"
End Sub


Sub MakeMF(ByVal nameMF As String, Optional ByVal Packer As Long = CMS_FORMAT_ZLIB)
    Dim setupPath As String, txtINI As String, txtMode As String, txtRes As String, txtOpt As String, txtFls As String
    Dim nameIcon As String, nameExe As String, txt As String, a As Long, v As Variant, Buf() As Byte
    Dim Mts As MatchCollection, Base64 As New clsBase64, RX As New clsRXP, f As New clsFileAPI
    
    nameMF = FileLongName(nameMF)
    setupPath = RX.Eval(nameMF, "(.+\\)", GetAppPath)

    If m_File2Buf(Buf, setupPath + "make.ini") Then
        txtINI = Code_Parse(Buf, "Make")
        If ExistsMember(CAS.CodeObject, "LMF_Make_Begin") Then txtINI = CBN("", "LMF_Make_Begin", VbFunc, Array(txtINI))
        
        txtFls = RX.Eval(txtINI, "\[files\]([^\[]+)", , , , , -1)
        txtRes = RX.Eval(txtINI, "\[resource\]([^\[]+)", , , , , -1)
        txtOpt = RX.Eval(txtINI, "\[options\]([^\[]+)", , , , , -1)
        
        If RX.Test(txtOpt, "\npacker[ \t]*=[ \t]*([^\r]+)") Then
            Packer = Val(Parse_MPath(RX.Mts(0).SubMatches(0)))
        End If
        
        If RX.Test(txtOpt, "\nfile[ \t]*=[ \t]*([^\r]+)") Then
            txt = Parse_MPath(RX.Mts(0).SubMatches(0))
            FileCopy nameMF, txt
            nameMF = IIF(Mid$(txt, 2, 1) = ":", txt, setupPath + txt)
        Else
            BackupMF nameMF
        End If
        
        If f.FOpen(nameMF) = INVALID_HANDLE Then Exit Sub
            f.Pos = f.LOF + 1
            
            Set Mts = RX.Execute(txtFls, "\n""([^""]+?)""(\.([a-z0-9_\-]+))?[ \t]*=[ \t]*([^\r]+)")
            
            For a = 0 To Mts.Count - 1
                txtMode = IIF(Len(Mts(a).SubMatches(2)), " mode=" & Parse_MPath(Mts(a).SubMatches(2)), "")
                        
                f.PutStr vbCrLf & vbCrLf
                
                txt = Parse_MPath(Mts(a).SubMatches(3))
                
                f.PutStr "<#res id=""" + CStr(Parse_MPath(Mts(a).SubMatches(0))) + """" + txtMode + " #>" + vbCrLf
                
                If m_File2Buf(Buf, txt) Then
                    If InStr(txtMode, "zlib") > 0 Then CompressData Buf
                    If InStr(txtMode, "base64") > 0 Then Buf = Base64.Encode(Buf)
                    f.PutBuf Buf
                End If
                
                f.PutStr vbCrLf + "<#res#>" + vbCrLf
            Next
        f.FClose

        nameIcon = Parse_MPath(RX.Eval(txtOpt, "\nicon[ \t]*=[ \t]*([^\r]+)"))

        mf_Tmp = nameMF

        If RX.Test(txtOpt, "\ntype[ \t]*=[ \t]*([^\r]+)") Then
            Select Case LCase$(Parse_MPath(RX.Mts(0).SubMatches(0)))
                Case "exe"
                    CompressMF nameMF, , Packer
                    nameExe = Parse_MPath(RX.Eval(txtOpt, "\nexe[ \t]*=[ \t]*([^\r]+)"))
                    mf_Tmp = MakeEXE(nameExe, nameMF, nameIcon, txtRes)
                    
                Case "full"
                    CompressMF nameMF, , Packer
            End Select
        End If

        If RX.Test(txtOpt, "\nshell(\-hide)*[ \t]*=[ \t]*([^\r]+)") Then
            ShellSync Parse_MPath(RX.Mts(0).SubMatches(1)), , Len(RX.Mts(0).SubMatches(0))
        End If
        
        If ExistsMember(CAS.CodeObject, "LMF_Make_End") Then Call CBN("", "LMF_Make_End", VbFunc, Array(txtINI))
        
        If RX.Test(txtOpt, "\nend[ \t]*=[ \t]*([^\r]+)") Then MsgBox Parse_MPath(RX.Mts(0).SubMatches(0)), , "LangMF"
    Else
        BackupMF nameMF
        'v = Array(101, "my", 2000, , -5, "привет")                  'headers extension
        CompressMF nameMF, v, Packer
    End If
    
    mf_Tmp = ""
End Sub


Function MakeEXE(ByVal nameExe As String, ByVal nameDest As String, ByVal nameIcon As String, ByVal txtRes As String) As String
    Dim txtOper As String, txtTable1 As String, txtTable2 As String, txtKey As String, txtValue As String
    Dim a As Long, f As New clsFileAPI, RXP As New clsRXP, ver As clsHash, clsNR As New clsNativeRes
    Dim isVerModify As Boolean, lngType As Long, lngLang As Long
    
    If Len(nameExe) = 0 Then nameExe = SYS.Path("engine_full")
    If Not IsFile(nameExe) Then Exit Function
    FileCopy nameExe, nameDest & ".tmp":    nameExe = nameDest & ".tmp"
    If Not IsFile(nameExe) Then Exit Function

    If IsFile(nameIcon) Then clsNR.UpdateMainIcon nameExe, nameIcon

    RXP.Obj.MultiLine = True
    
    '-------------------<Type>-------<Oper>-----<Lang>----<Table1>---<Table2>----<Key>------<Value>---
    RXP.Obj.Pattern = "\n([a-z0-9]+)\.([a-z0-9]*)\.([0-9]*)\.([a-z]*)\.([a-f0-9]*)\.(.+)[ \t]*=[ \t]*([^\r]+)"

    If RXP.Execute(txtRes).Count Then
        clsNR.UpdateVersion nameExe, ver

        For a = 0 To RXP.Mts.Count - 1
            txtOper = LCase$(RXP.Matches(1, a))
            lngLang = Val(RXP.Matches(2, a))
            txtTable1 = RXP.Matches(3, a)
            txtTable2 = RXP.Matches(4, a)
            txtKey = Trim$(RXP.Matches(5, a))
            txtValue = RXP.Matches(6, a)
            
            Select Case LCase$(RXP.Matches(0, a))
                Case "version"
                    lngType = 16
                    
                    If LenB(txtTable1) Then
                        If LenB(txtTable2) Then
                            If txtOper = "remove" Then
                                ver(txtTable1).Item(txtTable2).Remove txtKey
                            Else
                                ver(txtTable1).Item(txtTable2).Item(txtKey) = txtValue
                            End If
                        Else
                            If txtOper = "remove" Then
                                ver(txtTable1).Remove txtKey
                            Else
                                ver(txtTable1).Child txtKey
                            End If
                        End If
                    Else
                        If txtOper = "remove" Then
                            ver.Remove txtKey
                        Else
                            ver.Item(txtKey) = txtValue
                        End If
                    End If
                    
                    isVerModify = True
            
                Case "string"
                    lngType = 6
                    
                    If txtOper = "remove" Then txtValue = ""
                    clsNR.UpdateString nameExe, Val(txtKey), txtValue, lngLang
                
                Case Else
                    lngType = Val(RXP.Matches(0, a))
                    
                    clsNR.PutResourceFromFile Parse_MPath(txtValue), nameExe, lngType, IIF(Val(txtKey) > 0, Val(txtKey), txtKey), lngLang
            End Select
        Next
        
        If isVerModify Then clsNR.UpdateVersion nameExe, ver
    End If
    
    If f.FOpen(nameExe) <> INVALID_HANDLE Then
        txtValue = mf_Hdr & SYS.Conv.File2Str(nameDest)
        a = (f.LOF + Len(txtValue) + 4) Mod 16
        If a <> 0 Then a = 16 - a
        f.PutStr txtValue, f.LOF + 1 + a
        f.PutMem VarPtr(CLng(Len(txtValue))), 4
        f.FClose
    End If
    
    FileKill nameDest
    
    txtKey = GetExtension(nameDest)
    If LenB(txtKey) = 0 Then
        nameDest = nameDest + ".exe"
    ElseIf LCase$(txtKey) = "mf" Then
        nameDest = Left$(nameDest, Len(nameDest) - Len(txtKey) - 1) + ".exe"
    End If
    
    FileKill nameDest
    FileMove nameExe, nameDest
    
    MakeEXE = nameDest
End Function



Sub Timer_Event(ByVal hWnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
    Dim v As Variant
    v = SYS.Timers("#" & idEvent)
    If idEvent < 0 Then Call KillTimer(hWnd, idEvent):    SYS.Timers.Remove "#" & idEvent
    If IsArray(v) Then CBN v(1), CStr(v(0)), VbMethod, v(2)
End Sub

Sub Timer_Cron(ByVal hWnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
    SYS.Cron.Timer
End Sub

Sub Timer_Func(Optional ByVal hWnd As Long, Optional ByVal uMsg As Long, Optional ByVal idEvent As Long = 30001, Optional ByVal dwTime As Long)
    If hWnd Then Call KillTimer(hWnd, idEvent)
    If idEvent = 30001 Then CBN "", "Load", VbMethod, Array(Info.Arg)
End Sub
