VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsExt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public Property Get Printer() As Object
    Set Printer = VB.Printer
End Property

Public Property Set Printer(ByVal objPrn As Object)
    Set VB.Printer = objPrn
End Property

Public Function MsgBox(ByVal arg1 As String, Optional ByVal arg2 As Long, Optional ByVal arg3 As String = vbNullChar) As Long
    MsgBox = MessageBoxW(0, StrPtr(arg1), StrPtr(arg3), arg2 Or &H2000&)
End Function

Public Function InputBox(ByVal arg1 As String, Optional ByVal arg2 As String, Optional ByVal arg3 As String) As String
    InputBox = VBA.InputBox(arg1, arg2, arg3)
End Function

Public Sub SendKeys(ByVal numKey As String, Optional ByVal cWait As Boolean = False)
    VBA.SendKeys numKey, cWait
End Sub

Public Function Play(ByVal fName As String, ByVal hMod As Long, ByVal wFlags As Long) As Long
    Play = PlaySoundW(StrPtr(fName), hMod, wFlags)
End Function

Public Function Mci(ByVal sCommand As String, Optional ByVal isVal As Boolean) As Variant
    Dim txt As String, rc As Long
    txt = String$(512, 0)
    rc = mciSendStringW(StrPtr(sCommand), StrPtr(txt), Len(txt), 0)
    If rc <> 0 Then Mci = rc Else If isVal Then Mci = Val(TrimNull(txt)) Else Mci = TrimNull(txt)
End Function

Public Function Link(ByVal nameLink As String, Optional ByVal strOper As String = vbNullChar, Optional ByVal strParam As String = vbNullChar, Optional ByVal nShowCmd As Long = SW_SHOWNORMAL) As Long
    Link = ShellExecuteW(frmScript.hWnd, StrPtr(strOper), StrPtr(nameLink), StrPtr(strParam), StrPtr("C:\"), nShowCmd)
End Function

Public Function Modify(ByVal txtCode As String, Optional txtConv As Variant, Optional ByVal Flags As Long) As String
    Modify = Parse_Modify(txtCode, txtConv, Flags)
End Function

Public Function Template(ByVal txtCode As String, Optional ByVal fnPrint As String, Optional ByVal dmStart As String, Optional ByVal dmStop As String) As String
    Template = Parse_Template(txtCode, fnPrint, dmStart, dmStop)
End Function

Public Function Font(Optional ByVal fName As String = "Arial", Optional ByVal fSize As Currency = 8, Optional ByVal fBold As Boolean, Optional ByVal fItalic As Boolean, Optional ByVal fUnderline As Boolean, Optional ByVal fStrikethrough As Boolean, Optional ByVal fCharset As Integer = 204) As StdFont
    Set Font = New StdFont
    With Font:    .Name = fName:   .Size = fSize:   .Bold = fBold:   .Italic = fItalic:   .Underline = fUnderline:   .Strikethrough = fStrikethrough:    .Charset = fCharset:    End With
End Function

Public Sub VB_Load(ByVal tmpObj As Object)
    On Error Resume Next
    Load tmpObj
End Sub

Public Sub VB_Unload(ByVal tmpObj As Object)
    On Error Resume Next
    Unload tmpObj
End Sub

Public Function VB_CreateObject(ByVal nameObj As String, Optional ByVal nameServ As String) As Object
    On Error Resume Next
    Set VB_CreateObject = CreateObject(nameObj, nameServ)
End Function

Public Function SetPriority(Optional ByVal value As Long = &H20, Optional ByVal pid As Long) As Long
    Dim hProc As Long
    If pid = 0 Then pid = GetCurrentProcessId
    hProc = OpenProcess(PROCESS_QUERY_SET, 0&, pid)
    If hProc Then SetPriority = SetPriorityClass(hProc, value):    Call CloseHandle(hProc)
End Function

Public Function SetAffinity(Optional ByVal value As Long = -1, Optional ByVal hProc As Long) As Long
    If hProc = 0 Then hProc = GetCurrentProcess
    SetAffinity = SetProcessAffinityMask(hProc, value)
End Function

Public Function OSVer(Optional ByVal typeFmt As Long = 0) As Variant
    Static verInfo As OSVERSIONINFOEX
    
    verInfo.dwOSVersionInfoSize = Len(verInfo)
    If (GetVersionExA(verInfo)) = 0 Then Exit Function
    
    Select Case typeFmt
        Case 0:     OSVer = VarPtr(verInfo.dwOSVersionInfoSize)
        Case 1:     OSVer = verInfo.dwMajorVersion & "." & verInfo.dwMinorVersion
        Case 2:     OSVer = verInfo.dwMajorVersion & "." & verInfo.dwMinorVersion & "." & verInfo.dwBuildNumber
        Case 3:     OSVer = verInfo.wServicePackMajor & "." & verInfo.wServicePackMinor
        Case 4:     OSVer = verInfo.wSuiteMask
        Case 5:     OSVer = verInfo.wProductType
    End Select
End Function

Public Function Shell(ByVal FileName As String, Optional ByVal Flags As Long = 1) As Long
    Shell = VBA.Shell(FileName, Flags)
End Function

Public Function Shell_Sync(ByVal CommandLine As String, Optional ByVal Timeout As Long = -1, Optional ByVal Hide As Boolean = False) As Long
    Shell_Sync = ShellSync(CommandLine, Timeout, Hide)
End Function

Public Function Shell_SyncEx(ByVal FileName As String, Optional ByVal CommandLine As String, Optional ByVal Timeout As Long = -1, Optional ByVal flagShow As Long = 0, Optional ByVal lpVerb As String = vbNullString) As Long
    Shell_SyncEx = ShellSyncEx(FileName, CommandLine, Timeout, flagShow, lpVerb)
End Function

Public Function Shell_SyncOut(ByVal sCommandLine As String, Optional ByVal nShowWindow As Boolean = False, Optional ByVal fOEMConvert As Boolean = True, Optional ByVal CmdOut_Event As Object) As String
    Shell_SyncOut = CmdOut(sCommandLine, nShowWindow, fOEMConvert, CmdOut_Event)
End Function

Public Function RegisterDLL(ByVal FileName As String, Optional ByVal isReg As Boolean = True) As Boolean
    RegisterDLL = modCommon.RegisterDLL(FileName, isReg)
End Function

Public Function VersionDLL(ByVal FileName As String, Optional ByVal verCmp As Variant) As Variant
    VersionDLL = modCommon.VersionDLL(FileName, verCmp)
End Function

Public Function CreateImage() As Object
    Static cntImage As Long
    On Error Resume Next
    cntImage = cntImage + 1
    Set CreateImage = frmScript.Controls.Add("VB.PictureBox", "dynPic_" & cntImage)
    With CreateImage
        .BorderStyle = 0
        .AutoRedraw = True
        .Move 0, 0, GetSystemMetrics(SM_CXSMICON), GetSystemMetrics(SM_CYSMICON)
        .BackColor = &H808080
        .Font.Name = "Arial"
        .Font.Size = 10
        .Font.Charset = 204
        .Visible = False
    End With
End Function

Public Function GetMemoryStatus() As Collection
    Dim dbl_T_Phys As Double, dbl_A_Phys As Double
    Dim dbl_T_Page As Double, dbl_A_Page As Double
    Dim dbl_T_Virt As Double, dbl_A_Virt As Double
    Dim memStat As MEMORYSTATUSEX, cl As New Collection
    
    memStat.dwLength = Len(memStat)
    
    GlobalMemoryStatusEx memStat
    
    dbl_T_Phys = BigLongToDouble(memStat.ullTotalPhys.LowPart, memStat.ullTotalPhys.HighPart)
    dbl_T_Page = BigLongToDouble(memStat.ullTotalPageFile.LowPart, memStat.ullTotalPageFile.HighPart)
    dbl_T_Virt = BigLongToDouble(memStat.ullTotalVirtual.LowPart, memStat.ullTotalVirtual.HighPart)
    
    dbl_A_Phys = BigLongToDouble(memStat.ullAvailPhys.LowPart, memStat.ullAvailPhys.HighPart)
    dbl_A_Page = BigLongToDouble(memStat.ullAvailPageFile.LowPart, memStat.ullAvailPageFile.HighPart)
    dbl_A_Virt = BigLongToDouble(memStat.ullAvailVirtual.LowPart, memStat.ullAvailVirtual.HighPart)
    
    cl.Add dbl_A_Phys, "Phys Free"
    cl.Add dbl_T_Phys, "Phys Total"
    cl.Add dbl_T_Phys - dbl_A_Phys, "Phys Used"

    cl.Add dbl_A_Page, "PageFile Free"
    cl.Add dbl_T_Page, "PageFile Total"
    cl.Add dbl_T_Page - dbl_A_Page, "PageFile Used"

    cl.Add dbl_A_Virt, "Virtual Free"
    cl.Add dbl_T_Virt, "Virtual Total"
    cl.Add dbl_T_Virt - dbl_A_Virt, "Virtual Used"

    Set GetMemoryStatus = cl
    Set cl = Nothing
End Function

Public Function GetSysTimes() As Collection
    Static old_Idle As Double, old_User As Double, old_Kernel As Double, cl_Usage As Double
    Dim cl_Idle As Double, cl_User As Double, cl_Kernel As Double, cl_Sys As Double
    Dim dbl_Idle As Double, dbl_User As Double, dbl_Kernel As Double
    Dim ft_Idle As FILETIME, ft_User As FILETIME, ft_Kernel As FILETIME
    Dim cl As New Collection
    
    GetSystemTimes ft_Idle, ft_Kernel, ft_User
    
    dbl_Idle = BigLongToDouble(ft_Idle.dwLowDateTime, ft_Idle.dwHighDateTime)
    dbl_User = BigLongToDouble(ft_User.dwLowDateTime, ft_User.dwHighDateTime)
    dbl_Kernel = BigLongToDouble(ft_Kernel.dwLowDateTime, ft_Kernel.dwHighDateTime)
    
    cl_Idle = dbl_Idle - old_Idle
    cl_User = dbl_User - old_User
    cl_Kernel = dbl_Kernel - old_Kernel
    cl_Sys = cl_User + cl_Kernel
    If cl_Sys <> 0 Then cl_Usage = (cl_Sys - cl_Idle) * 100 / cl_Sys Else cl_Usage = 0
    
    cl.Add cl_Idle, "Idle"
    cl.Add cl_User, "User"
    cl.Add cl_Kernel, "Kernel"
    cl.Add cl_Usage, "Usage"

    old_Idle = dbl_Idle: old_User = dbl_User: old_Kernel = dbl_Kernel

    Set GetSysTimes = cl
    Set cl = Nothing
End Function

Public Sub AutoIcon(ByVal hWnd As Long, Data As Variant, Optional ByVal kNormal As Single = 1, Optional ByVal kSmall As Single = 1)
    Dim a As Long, b As Long, szIconB As Long, szIconS As Long, hIconB As Long, hIconS As Long
    Dim rs As Single, rb As Single, ms As Single, mb As Single
    Dim iconDE() As ICONDIRENTRY, iconHdr As ICONDIR, Buf() As Byte
    
    ConvToBufferByte Data, Buf:      If ArraySize(Buf) < 23 Then Exit Sub
    
    szIconB = GetSystemMetrics(SM_CXICON)
    szIconS = GetSystemMetrics(SM_CXSMICON)
    hIconB = -1
    hIconS = -1
    
    CopyMemory iconHdr, Buf(0), 6
    
    ReDim iconDE(iconHdr.idCount - 1)

    For a = 0 To iconHdr.idCount - 1
        CopyMemory iconDE(a), Buf(6 + a * 16), 16
    Next
    
    For b = 32 To 8 Step -1
        ms = 1000:     mb = 1000
        For a = 0 To iconHdr.idCount - 1
            With iconDE(a)
                If b = .wBitCount Then
                    If .bWidth = 0 Then rs = 256 / szIconS Else rs = .bWidth / szIconS
                    If .bWidth = 0 Then rb = 256 / szIconB Else rb = .bWidth / szIconB
                    If rs >= kSmall And rs < ms Then hIconS = a:   ms = rs
                    If rb >= kNormal And rb < mb Then hIconB = a:  mb = rb
                End If
            End With
        Next
        If hIconS <> -1 And hIconB <> -1 Then Exit For
    Next
    
    If hIconS > -1 Then
        With iconDE(hIconS)
            hIconS = CreateIconFromResourceEx(Buf(.dwImageOffset), .dwBytesInRes, True, &H30000, .bWidth, .bHeight, 0&)
        End With
        SendMessageW hWnd, WM_SETICON, ICON_SMALL, hIconS
    End If
    
    If hIconB > -1 Then
        With iconDE(hIconB)
            hIconB = CreateIconFromResourceEx(Buf(.dwImageOffset), .dwBytesInRes, True, &H30000, .bWidth, .bHeight, 0&)
        End With
        SendMessageW hWnd, WM_SETICON, ICON_BIG, hIconB
    End If
End Sub

Public Function CreateIE(Optional ByVal txtCode As String, Optional ByVal prtHWnd As Long, Optional ByVal Param As Variant) As Object
    Dim Style As clsFormStyle
    
    On Error Resume Next
    
    Set CreateIE = CreateObject("InternetExplorer.Application")
    
    With CreateIE
        .Navigate "about:blank"
        .ToolBar = False
        .StatusBar = False
        .Resizable = False
        
        While .Busy <> 0:    WaitMs:     Wend
        
        .Width = 400
        .Height = 300
        .Left = Screen.Width / (2 * Screen.TwipsPerPixelX) - .Width / 2
        .Top = Screen.Height / (2 * Screen.TwipsPerPixelY) - .Height / 2
        
        If LenB(txtCode) Then
            .Document.Write txtCode
            .Document.Close
            WaitMs -100
            While .Busy <> 0:    WaitMs:    Wend
        End If
        
        Set Style = New clsFormStyle
        Style.hWnd = .hWnd
        Style.MaxButton = False
        Style.MinButton = False
        .PutProperty "Style", Style
    End With
    
    If prtHWnd Then Style.TitleBar = False
    
    DoParams CreateIE, Param
    
    If prtHWnd Then Style.Child = prtHWnd
End Function

Public Function WaitIE(ByVal Obj As Variant, Optional ByVal Flags As Long) As Boolean
    On Error GoTo err1
    Do
        WaitIE = True
        WaitMs
        If Len(Obj.Document.readyState) = 0 Then
err1:
            WaitIE = False
            Err.Clear
        Else
            If Flags = 2 Then Obj.Quit
        End If
        If Flags = 0 Then Exit Do
    Loop Until WaitIE = False
End Function

Public Function Setup(ByVal value As String, Optional ByVal vPrm As String, Optional ByVal sDir As String, Optional ByVal SSE As Variant) As Boolean
    Dim pathTemp As String, pathSystem As String, nameFile As String, fnShell As String, txt As String, Buf() As Byte
    Dim isFail As Boolean, isTemp As Boolean, isReg As Boolean, isClean As Boolean, vStatus As Variant
    Dim Obj As Object, sItem As Object, sFolder As Object, RX As New clsRXP, Prm As clsHash
    
    On Error Resume Next
    
    Setup = True:      mf_Tmp = "":      ArrayDef SSE, "", "", 20000, SW_SHOWNORMAL, vbNullString

    Set Prm = GetQV(vPrm, nameFile)
    isReg = CBool(Prm("Reg")):      isClean = True:     If Prm.Exists("Clean") Then isClean = CBool(Prm("Clean"))
    
    If Prm.Exists("ProgID") Then Set Obj = CreateObject(Prm("ProgID")):    If Obj Is Nothing Then isFail = True
    If Len(nameFile) And Prm.Exists("Ver") Then If VersionDLL(nameFile, Prm("Ver")) = False Then isFail = True
    If Len(nameFile) And Not IsFile(nameFile) Then isFail = True
    
    If isFail = True Or (Len(nameFile) = 0 And Prm.Exists("ProgID") = False) Then
        Buf = SYS.Content(value, False, vStatus)

        If vStatus(0) Then
            pathSystem = GetSystemPath + "\"
            If Len(sDir) = 0 Then pathTemp = GetTmpPath + "lmf_" + GenTempStr + "\":         isTemp = True
            If Len(sDir) <> 0 Then pathTemp = CPath(Parse_MPath(sDir))
            mf_Tmp = pathTemp
            CreateDir pathTemp

            nameFile = pathTemp + RX.Eval(value, "([^\\\/\?]+)$")
            Buf2File Buf, nameFile

            Select Case LCase$(GetExtension(nameFile))
                Case "cab", "zip"
                    Const regCab1 = "CLSID\\{0CD7A5C0-9F37-11CE-AE65-08002B2E1262}"
                    Const regCab2 = "{098f2470-bae0-11cd-b579-08002b30bfeb}"
    
                    If SYS.Registry.RegRead("HKCR\.cab\") <> regCab1 Then
                        SYS.Registry.RegWrite "HKCR\.cab\", regCab1
                        SYS.Registry.RegWrite "HKCR\.cab\PersistentHandler\", regCab2
                    End If

                    Set Obj = CreateObject("Shell.Application")

                    With Obj
                        Set sFolder = .NameSpace(FileLongName(pathTemp))
    
                        For Each sItem In .NameSpace(FileLongName(nameFile)).Items
                            If Not sItem Is Nothing Then
                                txt = GetFileName(sItem.Path)
    
                                sFolder.CopyHere sItem, 0

                                Select Case LCase$(GetExtension(txt))
                                    Case "dll", "ocx"
                                        If isReg Then If modCommon.RegisterDLL(pathTemp + txt) = False Then Setup = False

                                    Case "inf", "cmd"
                                        If Len(fnShell) = 0 Then fnShell = pathTemp + txt
                                End Select
                            End If
                        Next

                        Set sItem = Nothing
                        Set sFolder = Nothing
                    
                        If isClean Then FileKill nameFile
                    End With

                Case "dll", "ocx"
                    If isReg Then If modCommon.RegisterDLL(nameFile) = False Then Setup = False
                
                Case "inf", "cmd", "exe"
                    fnShell = nameFile
            End Select
            
            If VarType(SSE(0)) = vbString Then
                If Len(SSE(0)) Then
                    SSE(0) = Parse_MPath(SSE(0))
                    SSE(1) = Parse_MPath(SSE(1))
                Else
                    SSE(0) = fnShell
                End If
                If Len(SSE(0)) Then
                    If IsFile(CStr(SSE(0))) Then
                        If LCase$(GetExtension(SSE(0))) = "inf" Then
                            SSE(1) = "advpack,LaunchINFSection """ & SSE(0) & """,DefaultInstall,4"
                            SSE(0) = pathSystem & "rundll32.exe"
                            If WinVer.dwMajorVersion >= 6 Then SSE(4) = "runas"
                        End If
                        Setup = (ShellSyncEx(SSE(0), SSE(1), SSE(2), SSE(3), SSE(4)) = 0)
                    Else
                        Setup = False
                    End If
                End If
            End If

            If isClean And isTemp Then RemoveDir pathTemp
        Else
            Setup = False
        End If
    End If

    If Prm.Exists("Lic") Then DoEvents:   Licenses.Add Prm("ProgID")
End Function
