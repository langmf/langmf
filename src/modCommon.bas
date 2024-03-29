Attribute VB_Name = "modCommon"
Option Explicit

Global IID_Null As UUID, IID_IClassFactory As UUID, IID_IDispatch As UUID, IID_IUnknown As UUID, IID_IPicture As UUID

Global WinVer As OSVERSIONINFOEX, ArrNRes() As Variant, HashWins As clsHash, TypeWins As Long, QPF As Currency

Global GT_BHex(15) As Byte
Global GT_IHex(255) As Byte
Global GT_Tran(255, -200 To 200) As Integer
Global GT_Grad(255, 255) As Byte
Global GT_Mix(-2100 To 2100) As Byte


Sub InitGlobal()
    Dim a As Long, b As Long
    
    '-------------------------------------
    With IID_IPicture:       .Data1 = &H7BF80980:  .Data2 = &HBF32:      .Data3 = &H101A:       .Data4(0) = &H8B:    .Data4(1) = &HBB:    .Data4(3) = &HAA:    .Data4(5) = &H30:    .Data4(6) = &HC:    .Data4(7) = &HAB:      End With
    With IID_IClassFactory:  .Data1 = 1:           .Data4(0) = &HC0:     .Data4(7) = &H46:      End With
    With IID_IDispatch:      .Data1 = &H20400:     .Data4(0) = &HC0:     .Data4(7) = &H46:      End With
    With IID_IUnknown:       .Data4(0) = &HC0:     .Data4(7) = &H46:                            End With

    '-------------------------------------
    WinVer.dwOSVersionInfoSize = Len(WinVer)
    Call GetVersionExA(WinVer)
    
    '-------------------------------------
    InitCommonControlsXP
    
    '-------------------------------------
    QueryPerformanceFrequency QPF

    '-------------------------------------
    For a = 0 To 255:       GT_Mix(a) = a:      Next
    For a = 256 To 2100:    GT_Mix(a) = 255:    Next
    
    '-------------------------------------
    For a = 0 To 255:       For b = -200 To 200:     GT_Tran(a, b) = (a / 100) * b:         Next:   Next

    '-------------------------------------
    For a = 0 To 255:       For b = 0 To 255:        GT_Grad(a, b) = (a / 255) * b:         Next:   Next
    
    '-------------------------------------
    For a = 0 To 9:     GT_BHex(a) = 48 + a:   Next:     For a = 10 To 15:   GT_BHex(a) = 55 + a:   Next
    
    '-------------------------------------
    For a = 0 To 255:   GT_IHex(a) = 255:      Next:     For a = 48 To 57:   GT_IHex(a) = a - 48:   Next
    For a = 65 To 70:   GT_IHex(a) = a - 55:   Next:     For a = 97 To 102:  GT_IHex(a) = a - 87:   Next
End Sub

Sub AdjustToken(Optional ByVal typePrivilege As String = "SeDebugPrivilege", Optional ByVal processID As Long = -1, Optional ByVal Flag As Long = SE_PRIVILEGE_ENABLED)
    Dim tkpTokenPrivilegeTmp As TOKEN_PRIVILEGES, tkpTokenPrivilegeNew As TOKEN_PRIVILEGES, LuidTmp As LUID
    Dim lngProcessHandle As Long, lngTokenHandle As Long, lngBufferLen As Long
    
    If processID = -1 Then lngProcessHandle = GetCurrentProcess() Else lngProcessHandle = processID

    OpenProcessToken lngProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY), lngTokenHandle
    LookupPrivilegeValue "", typePrivilege, LuidTmp
    With tkpTokenPrivilegeTmp:    .PrivilegeCount = 1:    .TheLuid = LuidTmp:    .Attributes = Flag:    End With
    
    AdjustTokenPrivileges lngTokenHandle, False, tkpTokenPrivilegeTmp, Len(tkpTokenPrivilegeNew), tkpTokenPrivilegeNew, lngBufferLen
End Sub

Function IsAdmin() As Boolean
    Dim rc As Long, hToken As Long, lRet As Long, tokenElev As Long
    
    On Error GoTo err1

    If IsUserAnAdmin() Then IsAdmin = True:   Exit Function

    If WinVer.dwMajorVersion < 6 Then Exit Function
   
    If OpenProcessToken(GetCurrentProcess(), TOKEN_READ, hToken) = 1 Then
        rc = GetTokenInformation(hToken, TOKEN_ELEVATION_TYPE, tokenElev, 4, lRet)
        If rc <> 0 And tokenElev = 2 Then IsAdmin = True
        CloseHandle hToken
    End If
err1:
End Function

Function Is64() As Boolean
    Dim rc As Boolean
    If GetProcAddress(GetModuleHandleW(StrPtr("kernel32")), "IsWow64Process") > 0 Then IsWow64Process GetCurrentProcess(), rc
    Is64 = rc
End Function

Function IsWine() As Boolean
    If GetProcAddress(GetModuleHandleW(StrPtr("ntdll")), "wine_get_version") > 0 Then IsWine = True
End Function

Function IsIDE() As Boolean
    On Error GoTo err1:
    Debug.Print 1 / 0
    Exit Function
err1:  IsIDE = True
End Function

Sub EndMF(Optional ByVal Interval As Long = 1)
    mf_IsEnd = True
    If Interval = 0 Then Script_End:    End
    Call SetTimer(frmScript.hWnd, IIF(Interval < 0, 30011, 30012), Abs(Interval), AddressOf Timer_Func)
End Sub

Function GetHPX(ByVal value As Long) As Long
    GetHPX = (value * GetDeviceCaps(frmScript.hDC, DC_LOGPIXELSX)) / 2540
End Function

Function GetHPY(ByVal value As Long) As Long
    GetHPY = (value * GetDeviceCaps(frmScript.hDC, DC_LOGPIXELSY)) / 2540
End Function

Function GetPHX(ByVal value As Long) As Long
    GetPHX = (value * 2540) / GetDeviceCaps(frmScript.hDC, DC_LOGPIXELSX)
End Function

Function GetPHY(ByVal value As Long) As Long
    GetPHY = (value * 2540) / GetDeviceCaps(frmScript.hDC, DC_LOGPIXELSY)
End Function

Sub FlexMove(ByVal Obj As Object, Optional ByVal typeX As Single = -1, Optional ByVal typeY As Single = -1, Optional ByVal typeW As Single, Optional ByVal typeH As Single, Optional ByVal ofsX As Single, Optional ByVal ofsY As Single, Optional ByVal prtW As Single, Optional ByVal prtH As Single, Optional x As Variant = "Left", Optional y As Variant = "Top", Optional w As Variant = "Width", Optional h As Variant = "Height")
    Dim sx As String, sy As String, sw As String, sh As String, t As Long, p As Long

    If Not Obj Is Nothing Then
        If VarType(x) = vbString Then sx = x:   x = CBN(Obj, sx, VbGet):   t = t + 1
        If VarType(y) = vbString Then sy = y:   y = CBN(Obj, sy, VbGet):   t = t + 2
        If VarType(w) = vbString Then sw = w:   w = CBN(Obj, sw, VbGet):   t = t + 4
        If VarType(h) = vbString Then sh = h:   h = CBN(Obj, sh, VbGet):   t = t + 8
        If prtW = 0 Then If Obj.Parent Is Nothing Then prtW = Obj.ScaleWidth Else prtW = Obj.Parent.ScaleWidth
        If prtH = 0 Then If Obj.Parent Is Nothing Then prtH = Obj.ScaleHeight Else prtH = Obj.Parent.ScaleHeight
    End If

    If prtW < 0 Then prtW = Abs(prtW) * Screen.TwipsPerPixelX
    If prtH < 0 Then prtH = Abs(prtH) * Screen.TwipsPerPixelY

    p = typeX \ 256:    typeX = typeX - p * 256:    p = Abs(p):    prtW = prtW / (p \ 256 + 1):    ofsX = ofsX + (p Mod 256) * prtW
    p = typeY \ 256:    typeY = typeY - p * 256:    p = Abs(p):    prtH = prtH / (p \ 256 + 1):    ofsY = ofsY + (p Mod 256) * prtH

    If typeX < 0 Then typeX = Round(typeX, 1)
    If typeY < 0 Then typeY = Round(typeY, 1)

    Select Case typeW
        Case Is > 1:    w = typeW
        Case Is > 0:    w = prtW * typeW
        Case Is < 0:    w = prtW + typeW
    End Select

    Select Case typeH
        Case Is > 1:    h = typeH
        Case Is > 0:    h = prtH * typeH
        Case Is < 0:    h = prtH + typeH
    End Select

    Select Case typeX
        Case Is > 0:    x = prtW * typeX + ofsX
        Case -1:        x = prtW / 2 - w / 2 + ofsX
        Case -1.1:      x = prtW / 2 + ofsX
        Case -1.2:      x = prtW / 2 - w + ofsX
        Case -2:        x = ofsX
        Case -3:        x = prtW - w + ofsX
    End Select

    Select Case typeY
        Case Is > 0:    y = prtH * typeY + ofsY
        Case -1:        y = prtH / 2 - h / 2 + ofsY
        Case -1.1:      y = prtH / 2 + ofsY
        Case -1.2:      y = prtH / 2 - h + ofsY
        Case -2:        y = ofsY
        Case -3:        y = prtH - h + ofsY
    End Select

    If (t And 3) = 3 Then
        If ExistsMember(Obj, "Move") Then
            If t = 15 Then CBN Obj, "Move", VbMethod, Array(h, w, y, x), -2, , p:    If p = S_OK Then Exit Sub
            If t And 4 Then CBN Obj, "Move", VbMethod, Array(w, y, x), -2, , p:      If p = S_OK Then Exit Sub
            CBN Obj, "Move", VbMethod, Array(y, x), -2, , p:                         If p = S_OK Then Exit Sub
        End If
    End If

    If t And 1 Then CBN Obj, sx, VbLet, Array(x), -2
    If t And 2 Then CBN Obj, sy, VbLet, Array(y), -2
    If t And 4 Then CBN Obj, sw, VbLet, Array(w), -2
    If t And 8 Then CBN Obj, sh, VbLet, Array(h), -2
End Sub

Function RegionFromBitmap(ByVal picSrc As IPictureDisp, Optional ByVal TransColor As Variant) As Long
    Dim Height As Long, Width As Long, rFinal As Long, rTmp As Long, Start As Long, Row As Long, Col As Long
    Dim Buf() As Long, bi As BITMAPINFO, Color As Long, isNotClear As Boolean
    
    With bi.bmiHeader
        .biSize = 40
        .biPlanes = 1
        .biBitCount = 32
        .biCompression = 0
        .biWidth = GetHPX(picSrc.Width)
        .biHeight = -GetHPY(picSrc.Height)
        Width = .biWidth: Height = Abs(.biHeight)
    End With
    
    ReDim Buf(Width - 1, Height - 1)
    
    If GetDIBits(frmScript.hDC, picSrc.Handle, 0, Height, Buf(0, 0), bi, 0) = 0 Then Exit Function
    
    If IsMissing(TransColor) Or IsEmpty(TransColor) Then Color = Buf(0, 0) Else Color = TransColor

    rFinal = CreateRectRgn(0, 0, 0, 0)

    For Row = 0 To Height - 1
        Col = 0
        Do While Col < Width
            Do While Col < Width
                If Buf(Col, Row) <> Color Then Exit Do
                Col = Col + 1
            Loop
            If Col < Width Then
                Start = Col
                Do While Col < Width
                    If Buf(Col, Row) = Color Then Exit Do
                    Col = Col + 1
                Loop
                If Col >= Width Then Col = Width
                rTmp = CreateRectRgn(Start, Row, Col, Row + 1)
                Call CombineRgn(rFinal, rFinal, rTmp, RGN_OR)
                DeleteObject (rTmp)
                isNotClear = True
            End If
        Loop
    Next
    
    If isNotClear Then RegionFromBitmap = rFinal Else DeleteObject (rFinal)
End Function

Function CompressData(Data() As Byte, Optional ByVal cmsType As Long = CMS_FORMAT_ZLIB) As Long
    Dim sz As Long, org As Long, iL As Long, iU As Long, Buf() As Byte
    Dim WorkSpaceSize As Long, WorkSpace As Long

    If cmsType = CMS_FORMAT_NONE Then Exit Function
    
    CompressData = -1
    GetBounds Data, iL, iU
    org = iU - iL + 1
    
    If org > 0 Then
        sz = org + (org * 0.01) + 12
        ReDim Buf(0 To sz - 1)
            
        If cmsType = CMS_FORMAT_ZLIB Then
            CompressData = zlib_Compress(Buf(0), sz, Data(iL), org)
        Else
            Call RtlGetCompressionWorkSpaceSize(cmsType, WorkSpaceSize, 0)
            Call NtAllocateVirtualMemory(-1, WorkSpace, 0, WorkSpaceSize, 4096, 64)
            CompressData = RtlCompressBuffer(cmsType, Data(iL), org, Buf(0), sz, 0, sz, WorkSpace)
            Call NtFreeVirtualMemory(-1, WorkSpace, 0, 16384)
            ReDim Preserve Buf(sz - 1)
        End If
        
        If CompressData = 0 And sz > 0 Then
            ReDim Preserve Buf(sz + 3)
            CopyMemory Buf(sz), org, 4
            Data = Buf
        Else
            Erase Data
        End If
        
        Erase Buf
    End If
End Function

Function DecompressData(Data() As Byte, Optional ByVal cmsType As Long = CMS_FORMAT_ZLIB) As Long
    Dim sz As Long, org As Long, iL As Long, iU As Long, Buf() As Byte
    
    If cmsType = CMS_FORMAT_NONE Then Exit Function
    
    DecompressData = -1
    GetBounds Data, iL, iU
    org = iU - iL - 3
    
    If org > 0 Then
        CopyMemory sz, Data(iU - 3), 4:      If sz <= 0 Then Exit Function
        
        ReDim Buf(0 To sz - 1) As Byte
        
        If cmsType = CMS_FORMAT_ZLIB Then
            DecompressData = zlib_UnCompress(Buf(0), sz, Data(iL), org)
        Else
            DecompressData = RtlDecompressBuffer(cmsType, Buf(0), sz, Data(iL), org, sz)
        End If
        
        If DecompressData = 0 And sz > 0 Then Data = Buf Else Erase Data
        
        Erase Buf
    End If
End Function

Sub GetBounds(TheData() As Byte, iL As Long, iU As Long)
    With GetSafeArray(TheData)
        iL = .rgSABound(0).lLbound
        iU = .rgSABound(0).cElements + iL - 1
    End With
End Sub

Function LongPath(value As String) As String
    If Len(value) > 250 Then LongPath = "\\?\" & value Else LongPath = value
End Function

Function IsFile(value As String, Optional ByVal mask_L As Long = -17, Optional ByVal mask_H As Long = 0) As Boolean
    Dim rc As Long
    rc = GetFileAttributesW(StrPtr(LongPath(value)))
    If rc <> -1 Then If ((rc And mask_L) = rc) And ((rc And mask_H) = mask_H) Then IsFile = True
End Function

Function IsFileExt(value As String, Optional ByVal vPath As Variant, Optional ByVal VExt As Variant) As Boolean
    Dim cp As Long, ce As Long, txt As String, sPath As String, sExt As String
    
    If Not IsArray(vPath) Then vPath = Array()
    If Not IsArray(VExt) Then VExt = Array()
    
    For cp = -1 To UBound(vPath)
        If cp = -1 Then sPath = "" Else sPath = vPath(cp)
        For ce = -1 To UBound(VExt)
            If ce = -1 Then sExt = "" Else sExt = VExt(ce)
            txt = sPath + value + sExt
            If IsFile(txt) Then value = txt:  IsFileExt = True:  Exit Function
        Next
    Next
End Function

Function GenTempStr(Optional ByVal value As Variant, Optional pat As String) As String
    Dim a As Long, b As Long, sz As Long, l As Long, u As UUID, out() As Byte, p() As Byte

    Static oldTm As Long
    
    If IsMissing(value) Or IsEmpty(value) Then
        sz = 12
    ElseIf IsNumber(value) Then
        sz = value
    End If
    
    If sz Then
        If sz > 0 Then a = timeGetTime - oldTm:    If a < 0 Or a > 50 Then oldTm = timeGetTime:    Randomize oldTm
        If sz < 0 Then sz = -sz
        
        If LenB(pat) = 0 Then pat = "abcdefghijklmnopqrstuvwxyz0123456789"
        
        p = pat:        l = Len(pat):       ReDim out(sz * 2 - 1)
    
        For a = 0 To sz * 2 - 1 Step 2
            b = CLng(Rnd * (l - 1)) * 2:        out(a) = p(b):         out(a + 1) = p(b + 1)
        Next
        
        GenTempStr = out
    End If
    
    If VarType(value) = vbString Then
        value = LCase$(value)
        If Len(value) = 4 And Mid$(value, 2, 3) = "uid" Then
            CoCreateGuid u:         GenTempStr = String$(38, 0):        StringFromGUID2 VarPtr(u), StrPtr(GenTempStr)
            If Left$(value, 1) = "u" Then GenTempStr = Replace(Mid$(GenTempStr, 2, Len(GenTempStr) - 2), "-", "")
        End If
    End If
End Function

Function EncodeUTF8(value As String, Optional ByVal Cpg As Long = 65001) As String
    If LenB(value) Then EncodeUTF8 = Conv_A2W_Str(Conv_W2A_Str(value, Cpg))
End Function

Function DecodeUTF8(value As String, Optional ByVal Cpg As Long = 65001) As String
    If LenB(value) Then DecodeUTF8 = Conv_A2W_Str(Conv_W2A_Str(value, 0), Cpg)
End Function

Function Command() As String
    If IsIDE Then Command = VBA.Command$:   Exit Function
    Command = GetStringPtrW(GetCommandLineW)
    Command = Right$(Command, Len(VBA.Command$))
End Function

Function GetWindowsPath() As String
    Dim Buf As String, rc As Long
    Buf = String$(MAX_PATH_X2, 0)
    rc = GetWindowsDirectoryW(StrPtr(Buf), MAX_PATH_X2)
    GetWindowsPath = Left$(Buf, rc)
End Function

Function GetSystemPath() As String
    Dim Buf As String, rc As Long
    Buf = String$(MAX_PATH_X2, 0)
    rc = GetSystemDirectoryW(StrPtr(Buf), MAX_PATH_X2)
    GetSystemPath = Left$(Buf, rc)
End Function

Function GetTmpPath() As String
    Dim Buf As String, rc As Long
    Buf = String$(MAX_PATH_UNI, 0)
    rc = GetTempPathW(MAX_PATH_UNI, StrPtr(Buf))
    GetTmpPath = Left$(Buf, rc)
End Function

Function GetAppPath(Optional ByVal isFull As Boolean = False) As String
    Dim Buf As String, rc As Long
    Buf = String$(MAX_PATH_UNI, 0)
    rc = GetModuleFileNameW(App.hInstance, StrPtr(Buf), MAX_PATH_UNI)
    Buf = Left$(Buf, rc)
    If Not isFull Then Buf = GetDirectory(Buf)
    GetAppPath = Buf
End Function

Function GetFileName(ByVal txtFile As String, Optional ByVal dm As String = "\") As String
    Dim Pos As Long
    Pos = InStrRev(txtFile, dm)
    If Pos > 0 Then GetFileName = Right$(txtFile, Len(txtFile) - Pos) Else GetFileName = txtFile
End Function

Function GetDirectory(ByVal txtFile As String, Optional ByVal dm As String = "\") As String
    Dim Pos As Long
    Pos = InStrRev(txtFile, dm)
    If Pos > 0 Then GetDirectory = Left$(txtFile, Pos)
End Function

Function GetFileTitle(ByVal txtFile As String, Optional ByVal dm As String = "\") As String
    Dim Pos As Long
    txtFile = GetFileName(txtFile, dm)
    Pos = InStrRev(txtFile, ".")
    If Pos > 0 Then GetFileTitle = Left$(txtFile, Pos - 1) Else GetFileTitle = txtFile
End Function

Function GetExtension(ByVal txtFile As String, Optional ByVal dm As String = "\") As String
    Dim Pos As Long
    txtFile = GetFileName(txtFile, dm)
    Pos = InStrRev(txtFile, ".")
    If Pos > 0 Then GetExtension = Right$(txtFile, Len(txtFile) - Pos) Else GetExtension = ""
End Function

Public Function FileShortName(ByVal fName As String) As String
    Dim rc As Long, txt As String
    txt = String$(MAX_PATH_UNI, 0)
    rc = GetShortPathNameW(StrPtr(LongPath(fName)), StrPtr(txt), Len(txt))
    If rc > 0 Then FileShortName = Left$(txt, rc)
    If InStr(FileShortName, "\\?\") = 1 Then FileShortName = Mid$(FileShortName, 5)
End Function

Public Function FileLongName(ByVal fName As String) As String
    Dim rc As Long, txt As String
    txt = String$(MAX_PATH_UNI, 0):    rc = GetLongPathNameW(StrPtr(fName), StrPtr(txt), Len(txt))
    If rc > 0 Then FileLongName = Left$(txt, rc)
End Function

Public Function FullPathName(ByVal fName As String) As String
    Dim rc As Long, txt As String
    txt = String$(MAX_PATH_UNI, 0):    rc = GetFullPathNameW(StrPtr(fName), Len(txt), StrPtr(txt), 0)
    If rc > 0 Then FullPathName = Left$(txt, rc)
End Function

Function Str2File(Buf As String, nameFile As String) As Long
    Dim f As New clsFileAPI
    If f.FOpen(nameFile, CREATE_ALWAYS) = INVALID_HANDLE Then Exit Function
    f.PutStr Buf
    f.FClose
    Str2File = True
End Function

Function File2Str(Buf As String, nameFile As String) As Long
    Dim f As New clsFileAPI
    Buf = ""
    If f.FOpen(nameFile, OPEN_EXISTING, GENERIC_READ) = INVALID_HANDLE Then Exit Function
    File2Str = f.LOF:    If File2Str Then Buf = String$(File2Str, 0):   f.GetStr Buf
    f.FClose
End Function

Function File2Buf(Buf() As Byte, nameFile As String) As Long
    Dim f As New clsFileAPI
    Erase Buf
    If f.FOpen(nameFile, OPEN_EXISTING, GENERIC_READ) = INVALID_HANDLE Then Exit Function
    File2Buf = f.LOF:    If File2Buf Then ReDim Buf(File2Buf - 1):   f.GetBuf Buf
    f.FClose
End Function

Function Buf2File(Buf() As Byte, nameFile As String) As Long
    Dim f As New clsFileAPI
    If f.FOpen(nameFile, CREATE_ALWAYS) = INVALID_HANDLE Then Exit Function
    f.PutBuf Buf
    f.FClose
    Buf2File = True
End Function

Function ListWindows(Optional ByVal hWnd As Long, Optional ByVal VType As Long) As clsHash
    TypeWins = VType
    Set HashWins = New clsHash
    Set ListWindows = HashWins
    If hWnd Then EnumChildWindows hWnd, AddressOf EnumWinCBK, 1 Else EnumWindows AddressOf EnumWinCBK, 1
End Function

Function EnumWinCBK(ByVal hWnd As Long, ByVal lParam As Long) As Long
    Dim rc As Long, txt As String, nm As String, pid As Long
    
    rc = GetWindowTextLengthW(hWnd)
    If rc Then txt = String$(rc, 0):    GetWindowTextW hWnd, StrPtr(txt), rc + 1
    
    nm = String$(MAX_PATH_X2, 0)
    rc = GetClassNameW(hWnd, StrPtr(nm), Len(nm))
    nm = Left$(nm, rc)
    
    If TypeWins = 0 Then
        If (IsIconic(hWnd) Or IsWindowVisible(hWnd)) And (GetParent(hWnd) = 0) Then
            GetWindowThreadProcessId hWnd, pid
            HashWins.Add Array(hWnd, txt, nm), pid
        End If
    ElseIf TypeWins = 1 Then
        HashWins.Add Array(hWnd, txt, nm), hWnd
    ElseIf TypeWins = 2 Then
        If IsWindowVisible(hWnd) Then HashWins.Add Array(hWnd, txt, nm), hWnd
    End If

    EnumWinCBK = 1
End Function

Function EnumNResCBK(ByVal hModule As Long, ByVal lpszType As Long, ByVal lpszName As Long, ByVal lParam As Long) As Long
    Dim uds As Long
    uds = UBound(ArrNRes):      ReDim Preserve ArrNRes(uds + 1)
    If (lpszName > &HFFFF&) Or (lpszName < 0) Then ArrNRes(uds) = GetStringPtrA(lpszName) Else ArrNRes(uds) = lpszName
    EnumNResCBK = 1
End Function

Function Conv_A2W_Buf(Buf() As Byte, Optional ByVal Cpg As Long = -1, Optional ByVal vPos As Long = 0) As String
    Dim sz As Long, ptrSrc As Long
    
    With GetSafeArray(Buf)
        sz = .rgSABound(0).cElements - vPos
        If .cDims = 1 Then ptrSrc = .pvData + vPos
    End With
    
    If ptrSrc = 0 Then Exit Function
    If Cpg = -1 Then Cpg = GetACP()
    Conv_A2W_Buf = String$(sz + 1, vbNullChar)
    sz = MultiByteToWideChar(Cpg, 0, ptrSrc, sz, StrPtr(Conv_A2W_Buf), sz + 1)
    Conv_A2W_Buf = Left$(Conv_A2W_Buf, sz)
End Function

Function Conv_A2W_Str(txt As String, Optional ByVal Cpg As Long = -1) As String
    Dim sz As Long
    sz = Len(txt):    Conv_A2W_Str = String$(sz + 1, vbNullChar):    If Cpg = -1 Then Cpg = GetACP()
    sz = MultiByteToWideChar(Cpg, 0, StrPtr(txt), sz, StrPtr(Conv_A2W_Str), sz + 1)
    Conv_A2W_Str = Left$(Conv_A2W_Str, sz)
End Function

Function Conv_W2A_Buf(txt As String, Optional ByVal Cpg As Long = -1) As Byte()
    Dim Buf() As Byte, sz As Long
    sz = Len(txt):     ReDim Buf(sz * 2):     If Cpg = -1 Then Cpg = GetACP()
    sz = WideCharToMultiByte(Cpg, 0, StrPtr(txt), sz, VarPtr(Buf(0)), sz * 2 + 1, 0, ByVal 0&)
    If sz > 0 Then ReDim Preserve Buf(sz - 1):    Conv_W2A_Buf = Buf
End Function

Function Conv_W2A_Str(txt As String, Optional ByVal Cpg As Long = -1) As String
    Dim sz As Long
    sz = Len(txt):    Conv_W2A_Str = String$(sz * 2 + 1, vbNullChar):    If Cpg = -1 Then Cpg = GetACP()
    sz = WideCharToMultiByte(Cpg, 0, StrPtr(txt), sz, StrPtr(Conv_W2A_Str), sz * 2 + 1, 0, ByVal 0&)
    Conv_W2A_Str = Left$(Conv_W2A_Str, sz)
End Function

Function ToUnicode(Buf() As Byte) As String
    Dim sz As Long, i As Long, b As Byte

    sz = ArraySize(Buf)

    If sz >= 3 Then
        If Buf(0) = 239 And Buf(1) = 187 And Buf(2) = 191 Then               'UTF-8
            ToUnicode = Conv_A2W_Buf(Buf, 65001, 3)
            Exit Function
        End If
    End If

    If sz >= 2 Then
        If Buf(0) = 255 And Buf(1) = 254 Then                               'UTF-16 LE
            ToUnicode = Mid$(Buf, 2)
            Exit Function
        End If
        
        If Buf(0) = 254 And Buf(1) = 255 Then                               'UTF-16 BE
            For i = 0 To sz - 1 Step 2
                b = Buf(i):     Buf(i) = Buf(i + 1):     Buf(i + 1) = b
            Next
            ToUnicode = Mid$(Buf, 2)
            Exit Function
        End If
    End If

    If sz Then ToUnicode = Conv_A2W_Buf(Buf)                                'ANSI
End Function

Function IStreamToArray(istm As stdole.IUnknown, arr() As Byte) As Boolean
    Dim hMem As Long, pMem As Long, cnt As Long

    If istm Is Nothing Then Exit Function
    If GetHGlobalFromStream(istm, hMem) <> 0 Then Exit Function

    cnt = GlobalSize(hMem):         If cnt <= 0 Then Exit Function
    pMem = GlobalLock(hMem)

    If pMem <> 0 Then
        ReDim arr(0 To cnt - 1)
        CopyMemory arr(0), ByVal pMem, cnt
        GlobalUnlock hMem
        IStreamToArray = True
    End If
End Function

Function IStreamFromArray(ByVal Ptr As Long, ByVal Length As Long) As stdole.IUnknown
    Dim hMem As Long, pMem  As Long

    On Error GoTo err1

    If Ptr = 0& Then CreateStreamOnHGlobal 0, 1, IStreamFromArray:     Exit Function
    If Length = 0 Then Exit Function

    hMem = GlobalAlloc(&H2&, Length):       If hMem = 0 Then Exit Function
    pMem = GlobalLock(hMem)

    If pMem <> 0 Then
        CopyMemory ByVal pMem, ByVal Ptr, Length
        Call GlobalUnlock(hMem)
        Call CreateStreamOnHGlobal(hMem, 1, IStreamFromArray)
    End If

    If IStreamFromArray Is Nothing Then GlobalFree hMem
err1:
End Function

Function LoadPictureFromByte(value As Variant) As IPicture
    Dim sz As Long, istm As stdole.IUnknown, tmpBuf() As Byte
    
    ConvToBufferByte value, tmpBuf:      sz = ArraySize(tmpBuf):        If sz = 0 Then Exit Function
    
    Set istm = IStreamFromArray(VarPtr(tmpBuf(0)), sz)
    Call OleLoadPicture(istm, sz, 0, IID_IPicture, LoadPictureFromByte)
    Set istm = Nothing
End Function

Function BigLongToDouble(ByVal low_part As Long, ByVal high_part As Long) As Double
    Dim Result As Double

    Result = high_part:             If high_part < 0 Then Result = Result + 2 ^ 32
    Result = Result * 2 ^ 32
    Result = Result + low_part:     If low_part < 0 Then Result = Result + 2 ^ 32

    BigLongToDouble = Result
End Function

Function CmdOut(ByVal sCommandLine As String, Optional ByVal nShowWindow As Boolean = False, Optional ByVal fOEMConvert As Boolean = True, Optional ByVal CmdOut_Event As Object) As String
    Dim hPipeRead As Long, hPipeWrite1 As Long, hPipeWrite2 As Long, hCurProcess As Long, lBytesRead As Long, lpExitCode As Long
    Dim SA As SECURITY_ATTRIBUTES, si As STARTUPINFO, pi As PROCESS_INFORMATION
    Dim baOutput As String, sNewOutput As String, tp As Long
    
    Const BUFSIZE = 4096&
    
    baOutput = String$(BUFSIZE, 0):     SA.nLength = Len(SA):      SA.bInheritHandle = 1
        
    If CreatePipe(hPipeRead, hPipeWrite1, SA, BUFSIZE) = 0 Then Exit Function
    hCurProcess = GetCurrentProcess()
    Call DuplicateHandle(hCurProcess, hPipeRead, hCurProcess, hPipeRead, 0, 0, DUPLICATE_SAME_ACCESS Or DUPLICATE_CLOSE_SOURCE)
    Call DuplicateHandle(hCurProcess, hPipeWrite1, hCurProcess, hPipeWrite2, 0, 1, DUPLICATE_SAME_ACCESS)
    
    With si
        .cb = Len(si)
        .dwFlags = STARTF_USESHOWWINDOW Or STARTF_USESTDHANDLES
        .wShowWindow = IIF(nShowWindow, 1, 0)
        .hStdOutput = hPipeWrite1
        .hStdError = hPipeWrite2
    End With
    
    If CreateProcessW(0, StrPtr(sCommandLine), 0, 0, 1, 0, 0, 0, si, pi) Then
        Call CloseHandle(pi.hThread)
        Call CloseHandle(hPipeWrite1)

        hPipeWrite1 = 0:      If hPipeWrite2 Then Call CloseHandle(hPipeWrite2):    hPipeWrite2 = 0
        
        Do
            sNewOutput = "":      WaitMs:      GetExitCodeProcess pi.hProcess, lpExitCode

            If ReadFileStr(hPipeRead, baOutput, BUFSIZE, lBytesRead, 0) <> 0 Then
                If fOEMConvert Then
                    sNewOutput = String$(lBytesRead, 0)
                    Call OemToCharBuffA(baOutput, sNewOutput, lBytesRead)
                Else
                    sNewOutput = Left$(baOutput, lBytesRead)
                End If
                
                CmdOut = CmdOut & sNewOutput
                
                tp = 0
            Else
                tp = 1
                If lpExitCode <= 0 Then tp = 2
            End If
            
            If tp <> 1 Then If Not CmdOut_Event Is Nothing Then Call CmdOut_Event(Array(lpExitCode, sNewOutput, pi.hProcess, pi.dwProcessID))
        Loop Until lpExitCode <= 1
        
        Call CloseHandle(pi.hProcess)
    End If
    
    Call CloseHandle(hPipeRead)

    If hPipeWrite1 Then Call CloseHandle(hPipeWrite1)
    If hPipeWrite2 Then Call CloseHandle(hPipeWrite2)
End Function

Function GetStringPtrA(ByVal lpszA As Long) As String
    GetStringPtrA = String$(lstrlenA(ByVal lpszA), 0)
    Call lstrcpyA(ByVal GetStringPtrA, ByVal lpszA)
End Function

Function GetStringPtrW(ByVal lpszW As Long) As String
    GetStringPtrW = String$(lstrlenW(ByVal lpszW), 0)
    Call lstrcpyW(ByVal StrPtr(GetStringPtrW), ByVal lpszW)
End Function

Function GetStringPtrU(ByVal lpszU As Long) As String
    Dim sz As Long, ptrDst As Long
        
    sz = lstrlenA(ByVal lpszU)
    If sz = 0 Then Exit Function
    GetStringPtrU = String$(sz + 1, vbNullChar)
    ptrDst = StrPtr(GetStringPtrU)
    
    sz = MultiByteToWideChar(65001, 0, lpszU, sz, ptrDst, sz + 1)
    
    GetStringPtrU = Left$(GetStringPtrU, sz)
End Function

Function SpecialFolderPath(ByVal lngFolderType As Long) As String
    Dim strPath As String, IDL As ITEMIDLIST
    If SHGetSpecialFolderLocation(0&, lngFolderType, IDL) Then Exit Function
    strPath = String$(MAX_PATH_UNI, 0)
    If SHGetPathFromIDListW(ByVal IDL.mkid.cb, StrPtr(strPath)) Then SpecialFolderPath = TrimNull(strPath)
End Function

Function ConvToBufferByte(bufVar As Variant, bufByte() As Byte) As Boolean
    Dim a As Long, vt As Integer, uds As Long, SA As SafeArray, v() As Variant
    
    vt = VariantType(bufVar, True)
    
    Select Case vt
        Case vbArray + vbByte:          bufByte = bufVar
        Case vbString:                  bufByte = Conv_W2A_Buf(CStr(bufVar))
        
        Case vbArray + vbVariant
    
            SA = GetSafeArray(bufVar):         uds = SA.rgSABound(0).cElements - 1:         If uds < 0 Then Exit Function
            
            ReDim bufByte(uds):         SA.fFeatures = 128:         PutMem4 VarPtrArray(v), VarPtr(SA)
            
            For a = 0 To uds:           bufByte(a) = v(a):          Next:           PutMem4 VarPtrArray(v), 0

        Case Else
            Exit Function
    End Select
    
    ConvToBufferByte = True
End Function

Function ConvFromBufferByte(bufVar As Variant, bufByte() As Byte, Optional ByVal vt As Variant) As Boolean
    Dim a As Long, b As Long, uds As Long, SA As SafeArray, v() As Variant
    
    If VarType(bufVar) = vbEmpty Then bufVar = Array()
    If IsMissing(vt) Then vt = VariantType(bufVar, True)
    
    Select Case vt
        Case vbArray + vbByte:      bufVar = bufByte
        Case -vbString:             bufVar = ToUnicode(bufByte)
        Case vbString:              bufVar = Conv_A2W_Buf(bufByte)
        
        Case vbArray + vbVariant
            
            uds = ArraySize(bufByte) - 1:       If uds < 0 Then Exit Function
            
            GetMem2 VarPtr(bufVar), vt
            
            If vt = VT_BYREF + VT_VARIANT Then
                SA = GetSafeArray(bufVar):          SA.fFeatures = 128:       PutMem4 VarPtrArray(v), VarPtr(SA):       ReDim v(uds)
                For a = 0 To uds:       v(a) = bufByte(a):          Next:     PutMem4 VarPtrArray(v), 0&
            Else
                ReDim bufVar(uds)
                For a = 0 To uds:       bufVar(a) = bufByte(a):     Next
            End If

        Case -4 To -1
            
            b = Abs(vt) - 1:      ReDim Preserve bufByte(b):      ReDim Preserve bufByte(3):       GetMem4 VarPtr(bufByte(0)), a
            bufVar = a
            
        Case Else
            Exit Function
    End Select
    
    ConvFromBufferByte = True
End Function

Function Buf2Hex(tmpBuf() As Byte) As Byte()
    Dim i As Long, p As Long, sz As Long, n1 As Byte, n2 As Byte, tmpOut() As Byte
    
    sz = ArraySize(tmpBuf):      If sz = 0 Then Exit Function
    
    ReDim tmpOut(sz * 2 - 1)

    For i = 0 To sz - 1
        n1 = tmpBuf(i):    n2 = n1 And 15:    n1 = n1 \ 16
        tmpOut(p) = GT_BHex(n1):   p = p + 1
        tmpOut(p) = GT_BHex(n2):   p = p + 1
    Next
    
    Buf2Hex = tmpOut
End Function

Function Hex2Buf(tmpBuf() As Byte) As Byte()
    Dim i As Long, p As Long, sz As Long, n1 As Byte, n2 As Byte, tmpOut() As Byte
    
    sz = ArraySize(tmpBuf):      If sz < 2 Then Exit Function
    
    ReDim tmpOut(sz / 2 - 1):    n1 = 255

    For i = 0 To sz - 1
        n2 = GT_IHex(tmpBuf(i))
        If n2 <> 255 Then
            If n1 = 255 Then n1 = n2 Else tmpOut(p) = n1 * 16 + n2:   p = p + 1:   n1 = 255
        End If
    Next

    If p Then ReDim Preserve tmpOut(p - 1):    Hex2Buf = tmpOut
End Function

Sub SetTopMost(ByVal hWnd As Long, Optional ByVal value As Boolean = True)
    SetWindowPos hWnd, IIF(value, hWnd_TOPMOST, hWnd_NOTOPMOST), 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOSIZE
End Sub

Sub SetIconWindow(ByVal hWnd As Long, ByVal varName As Variant, Optional ByVal nameLib As String)
    Dim a As Long, cx As Long, hIconL As Long, hIconS As Long, lpName As Long, typeRes As Long, hInst As Long

    If VarType(varName) = vbString Then
        lpName = StrPtr(varName)
        typeRes = LR_LOADFROMFILE
    Else
        lpName = CLng(varName)
        typeRes = LR_SHARED
    End If
    
    If LenB(nameLib) Then hInst = LoadLibrary(StrPtr(nameLib)) Else hInst = App.hInstance
   
    cx = GetSystemMetrics(SM_CXICON)
    If cx Mod 32 <> 0 Then cx = (cx \ 32) * 64
    For a = cx To 256 Step 32
        hIconL = LoadImageAsString(hInst, lpName, IMAGE_ICON, a, a, typeRes)
        If hIconL Then Exit For
    Next
    SendMessageW hWnd, WM_SETICON, ICON_BIG, hIconL
    
    cx = GetSystemMetrics(SM_CXSMICON)
    If cx Mod 16 <> 0 Then cx = (cx \ 16) * 32
    For a = cx To 256 Step 16
        hIconS = LoadImageAsString(hInst, lpName, IMAGE_ICON, a, a, typeRes)
        If hIconS Then Exit For
    Next
    SendMessageW hWnd, WM_SETICON, ICON_SMALL, hIconS
    
    If LenB(nameLib) Then FreeLibrary hInst
End Sub

Sub InitCommonControlsXP()
    Dim iccex As tagInitCommonControlsEx
    
    On Error GoTo err1
    
    With iccex
      .lngSize = Len(iccex)
      .lngICC = &H7FFF&
    End With
    
    InitCommonControlsEx iccex
err1:
End Sub

Sub WaitMs(Optional ByVal msec As Long)
    Dim t1 As Long, t2 As Long, lowCPU As Boolean
    
    If msec = 0 Then API_DoEvents:     Sleep 1:             Exit Sub
    If msec < 0 Then lowCPU = True:    msec = Abs(msec)
    
    t1 = timeGetTime
    Do
        DoEvents:    If lowCPU Then Sleep 1
        t2 = timeGetTime
    Loop Until (t1 + msec) < t2 Or t1 > t2
End Sub

Function API_DoEvents() As Long
    Dim wMsg As WNDMsg
    While PeekMessageA(wMsg, 0, 0, 0, PM_REMOVE)
        Call TranslateMessage(wMsg)
        Call DispatchMessageA(wMsg)
        API_DoEvents = True
    Wend
End Function

Function API_Error(ByVal vLastDllError As Long, Optional ByVal nFile As String) As String
    Dim Flags As Long, hModule As Long, sz As Long
    API_Error = Space$(4096):      Flags = &H1000&:      If Len(nFile) Then Flags = &H800&:   hModule = GetModuleHandleW(StrPtr(nFile))
    sz = FormatMessageW(&H1200&, 0, vLastDllError, 0, StrPtr(API_Error), Len(API_Error))
    If sz = 0 Then sz = FormatMessageW(&HA00&, hModule, vLastDllError, 0, StrPtr(API_Error), Len(API_Error))
    If sz > 2 Then If Mid$(API_Error, sz - 1, 2) = vbCrLf Then sz = sz - 2
    API_Error = Left$(API_Error, sz)
End Function

Sub RmDir(ByVal nameDir As String)
    RemoveDirectoryW StrPtr(LongPath(nameDir))
End Sub

Sub MkDir(ByVal nameDir As String)
    CreateDirectoryW StrPtr(LongPath(nameDir)), 0
End Sub

Sub ChDir(ByVal nameDir As String)
    SetCurrentDirectoryW StrPtr(nameDir)
End Sub

Function CurDir() As String
    Dim Buf As String, rc As Long
    Buf = String$(MAX_PATH_UNI, 0)
    rc = GetCurrentDirectoryW(MAX_PATH_UNI, StrPtr(Buf))
    CurDir = Left$(Buf, rc)
End Function

Sub CreateDir(ByVal nameDir As String)
    Dim tmpDir() As String, a As Long, txt As String
    tmpDir = Split(nameDir, "\")
    For a = 0 To UBound(tmpDir) - 1
        txt = txt & tmpDir(a) & "\":    If IsFile(txt, -1, vbDirectory) = False Then MkDir txt
    Next
End Sub

Sub RemoveDir(ByVal nameDir As String)
    Dim hFnd As Long, rc As Long, nFile As String, WFD As WIN32_FIND_DATA

    hFnd = FindFirstFileW(StrPtr(LongPath(nameDir & "*")), VarPtr(WFD)):      If hFnd = INVALID_HANDLE Then Exit Sub
    
    Do
        nFile = TrimNull(WFD.cFileName)

        If nFile <> "." And nFile <> ".." Then
            rc = WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY
            If rc = 0 Then FileKill nameDir & nFile Else RemoveDir nameDir & nFile & "\"
        End If
    Loop While FindNextFileW(hFnd, VarPtr(WFD))
    
    FindClose hFnd
    RmDir nameDir
End Sub

Function CPath(fPath As String, Optional ByVal typePath As Boolean = True, Optional ByVal dm As String = "\") As String
    If typePath Then
        If Right$(fPath, 1) <> dm Then fPath = fPath + dm
    Else
        If Right$(fPath, 1) = dm Then fPath = Left$(fPath, Len(fPath) - 1)
    End If
    CPath = fPath
End Function

Function FormatBytes(ByVal value As Double, Optional ByVal arrUnit As Variant) As String
    Const KB1 As Single = 1024, MB1 As Single = KB1 * 1024, GB1 As Single = MB1 * 1024, TB1 As Single = GB1 * 1024

    If Not IsArray(arrUnit) Then arrUnit = Array(" bytes", " KB", " MB", " GB", " TB")

    If value <= 999 Then
        FormatBytes = Format$(value, "0") & arrUnit(0)
    ElseIf value <= KB1 * 999 Then
        FormatBytes = ThreeNonZeroDigits(value / KB1) & arrUnit(1)
    ElseIf value <= MB1 * 999 Then
        FormatBytes = ThreeNonZeroDigits(value / MB1) & arrUnit(2)
    ElseIf value <= GB1 * 999 Then
        FormatBytes = ThreeNonZeroDigits(value / GB1) & arrUnit(3)
    Else
        FormatBytes = ThreeNonZeroDigits(value / TB1) & arrUnit(4)
    End If
End Function

Function ThreeNonZeroDigits(ByVal value As Double) As String
    If value >= 100 Then ThreeNonZeroDigits = Format$(CInt(value)):     Exit Function
    If value >= 10 Then ThreeNonZeroDigits = Format$(value, "0.0"):     Exit Function
    ThreeNonZeroDigits = Format$(value, "0.00")
End Function

Function TrimNull(txt As String) As String
    Dim Pos As Long
    Pos = InStr(txt, vbNullChar)
    If Pos Then TrimNull = Left$(txt, Pos - 1) Else TrimNull = txt
End Function

Function WindowLong(ByVal hWnd As Long, Optional ByVal value As Variant, Optional ByVal mStyle As Long, Optional ByVal nIndex As Long = GWL_STYLE) As Long
    Dim v As Long
    
    If hWnd = 0 Then Exit Function
    
    v = GetWindowLongW(hWnd, nIndex)
    
    If IsMissing(value) Then
        WindowLong = (v And mStyle) = mStyle
    Else
        If value Then v = v Or mStyle Else v = v And Not mStyle
        WindowLong = SetWindowLongW(hWnd, nIndex, v)
    End If
End Function

Function GetArgv(ByVal value As String) As Variant
    Dim i As Long, n As Long, t As String, c As String, v As String, p As String, r As Variant, h As New clsHash, REG1 As RegExp, Mts As MatchCollection

    Set REG1 = New RegExp:          REG1.Global = True:         REG1.IgnoreCase = True:         h.IgnoreCase = False
    
    REG1.Pattern = "(^|\s+)(((\-|\-\-|\/)(\w+)\s*(=)?\s*(""[^""]*""|[^\s\-]+|(?=\s|$)))|(""[^""]*""|[^\s\-]+))"
    
    Set Mts = REG1.Execute(value):          REG1.Pattern = "^[+\-]?(\d+|\d*\.\d+|&H\d+)$"
    
    For i = 0 To Mts.Count - 1
        t = Replace$(Mts(i).SubMatches(1), """", ""):           p = Mts(i).SubMatches(3)
        v = Replace$(Mts(i).SubMatches(6), """", ""):           c = Mts(i).SubMatches(4)
        
        If LenB(c) Then
            If REG1.Test(v) Then r = Val(v) Else If LenB(v) Then r = v Else r = True
            
            Select Case p
                Case "/":       h(c) = IIF(VarType(r) = vbString, Parse_MPath(r), r)
                Case "-":       h(c) = r
                Case Else:      h(c) = v
            End Select
            
        ElseIf LenB(t) Then
            h(n) = t:           n = n + 1
        End If
    Next
    
    Set GetArgv = h
End Function

Function GEV(Optional ByVal ID As Variant) As Variant
    Dim cnt As Integer, txt() As String, Col As New clsHash

    If IsMissing(ID) Or IsEmpty(ID) Then
        While LenB(Environ$(cnt + 1))
            txt = Split(Environ$(cnt + 1), "=")
            Col.Add txt(1), txt(0)
            cnt = cnt + 1
        Wend
        Set GEV = Col
    Else
        GEV = Environ$(ID)
    End If
End Function

Function StdFontToLogFont(fnt As StdFont) As LOGFONT
    Dim s As String, i As Long, hDC As Long, b() As Byte
    If fnt Is Nothing Then Set fnt = New StdFont
    With StdFontToLogFont
        s = fnt.Name:    b = s:    For i = 0 To Len(s) * 2 - 1:   .lfFaceName(i) = b(i):      Next
        hDC = GetDC(0):     .lfHeight = -MulDiv(fnt.Size, GetDeviceCaps(hDC, DC_LOGPIXELSY), 72):      ReleaseDC 0, hDC
        'If fnt.Bold Then .lfWeight = FW_BOLD Else .lfWeight = FW_NORMAL
        .lfWeight = fnt.Weight
        .lfItalic = fnt.Italic
        .lfUnderline = fnt.Underline
        .lfStrikeOut = fnt.Strikethrough
        .lfCharSet = fnt.Charset
    End With
End Function

Sub PrintText(ByVal hWnd As Long, ByVal hDC As Long, Text As String, Optional ByVal Color As Long, Optional ByVal Font As StdFont, Optional ByVal Flag As Long = DT_SINGLELINE Or DT_CENTER Or DT_VCENTER)
    Dim hFont As Long, hTmp As Long, rc As RECT
    GetClientRect hWnd, rc
    SetTextColor hDC, Color
    hFont = CreateFontIndirectW(StdFontToLogFont(Font))
    hTmp = SelectObject(hDC, hFont)
    SetBkMode hDC, MD_TRANSPARENT
    DrawTextW hDC, StrPtr(Text), -1, rc, Flag
    DeleteObject SelectObject(hDC, hTmp)
End Sub

Function LoadResDataWNull(ByVal ID As Variant, VType As Variant) As Byte()
    Dim a As Long, Buf() As Byte
    
    Buf = LoadResData(ID, VType):    If ArraySize(Buf) = 0 Then Exit Function
    
    For a = UBound(Buf) To 0 Step -1
        If Buf(a) <> 0 Then ReDim Preserve Buf(a):   Exit For
    Next
    
    LoadResDataWNull = Buf
End Function

Sub FileMove(fScr As String, fDest As String)
    FileCopy fScr, fDest
    FileKill fScr
End Sub

Sub FileCopy(fScr As String, fDst As String, Optional ByVal bFailIfExists As Long = 0)
    On Error Resume Next
    CopyFileW StrPtr(LongPath(fScr)), StrPtr(LongPath(fDst)), bFailIfExists
End Sub

Sub FileKill(fScr As String)
    On Error Resume Next
    DeleteFileW StrPtr(LongPath(fScr))
End Sub

Function GetQV(ByVal value As String, txtFile As String) As clsHash
    Dim v As Variant, tmp() As String, txt() As String

    Set GetQV = New clsHash
    
    If LenB(value) = 0 Then Exit Function
    tmp = Split(value, "?"):    txtFile = tmp(0):    If UBound(tmp) = 0 Then Exit Function
            
    For Each v In Split(tmp(1), "&")
        txt = Split(v, "=")
        If UBound(txt) = 0 Then GetQV.Add "", txt(0) Else GetQV.Add txt(1), txt(0)
    Next
End Function

Property Get VariantType(vrtValue As Variant, Optional ByVal isBYREF As Boolean) As Integer
    Dim Ptr As Long, tmp As Integer, vt As Integer
    
    Ptr = VarPtr(vrtValue)
    
    Do
        GetMem2 Ptr, vt
        If Not isBYREF Then Exit Do
        tmp = vt And (VT_BYREF - 1)
        If tmp <> VT_VARIANT Then vt = tmp: Exit Do
        If (vt And VT_BYREF) = 0 Then Exit Do
        GetMem4 Ptr + 8, Ptr
    Loop
    
    VariantType = vt
End Property

Property Let VariantType(vrtValue As Variant, Optional ByVal isBYREF As Boolean, ByVal VrtType As Integer)
    If isBYREF Then VrtType = VrtType Or VT_BYREF
    PutMem2 VarPtr(vrtValue), VrtType
End Property

Function IsNumber(value As Variant) As Boolean
    Dim vt As Integer
    vt = VarType(value):    If vt <> vbString Then IsNumber = IsNumeric(value)
End Function

Function AllowExecuteCode(ByVal addrCode As Long, ByVal sizeCode As Long, Optional ByVal Flag As Long = PAGE_EXECUTE_READWRITE) As Long
    VirtualProtect addrCode, sizeCode, Flag, AllowExecuteCode
End Function

Function MapArray(ByVal arrPtr As Long, ByVal pvData As Long) As Long
    Dim ap As Long
    GetMem4 arrPtr, ap
    GetMem4 ap + 12, MapArray
    PutMem4 ap + 12, pvData
End Function

Function Deref(ByVal Ptr As Long) As Long
    GetMem4 Ptr, Deref
End Function

Function ObjFromIUnk(vIUnk As Variant) As ATL.IUnknown
    If IsObject(vIUnk) Then Set ObjFromIUnk = vIUnk Else Set ObjFromIUnk = ObjFromPtr(vIUnk, True)
End Function

Function ObjFromPtr(ByVal vPtr As Long, Optional isIUnknown As Boolean = False) As Variant
    Dim VD As Object, IUnk As ATL.IUnknown
    
    If isIUnknown Then
        CopyMem4 vPtr, IUnk:   Set ObjFromPtr = IUnk:   CopyMem4 0&, IUnk
    Else
        CopyMem4 vPtr, VD:     Set ObjFromPtr = VD:     CopyMem4 0&, VD
    End If
End Function

Function GetGuid(value As Variant) As UUID
    If IsMissing(value) Or IsEmpty(value) Then
        GetGuid = IID_IUnknown
    ElseIf VarType(value) = vbString Then
        If Left$(value, 1) = "{" Then CLSIDFromString StrPtr(value), GetGuid Else CLSIDFromProgID StrPtr(value), GetGuid
    ElseIf IsNumeric(value) Then
        CopyMemory GetGuid, ByVal CLng(value), 16
    End If
End Function

Function StructFunc(ByVal This As Long, Optional vsp As Variant) As Long
    Dim a As Long, ofs As Long, dsp As Long

    Select Case VarType(vsp)
        Case vbInteger, vbLong:    ofs = vsp:   ReDim vsp(4):   vsp(0) = ofs
        Case vbError:              ReDim vsp(4):   vsp(4) = 0
        Case vbSingle, vbDouble:   ofs = Fix(vsp):  dsp = Fix((vsp - ofs) * 10):  ReDim vsp(4):  vsp(0) = ofs:  vsp(4) = dsp
    End Select
    
    'vsp(0)-offset      vsp(1)-inc offset   vsp(2)-length   vsp(3)-dst Ptr  vsp(4)-deference Ptr
    ArrayDef vsp, 0, 0, 4, 0, 1
    
    StructFunc = This                                               'main Ptr
    For a = 1 To vsp(4):      CopyMem4 ByVal StructFunc, StructFunc:       Next
    StructFunc = StructFunc + CLng(vsp(0))
    vsp(0) = vsp(0) + vsp(1)
End Function

Function ShellSync(ByVal CommandLine As String, Optional ByVal Timeout As Long = -1, Optional ByVal Hide As Boolean = False) As Long
    Dim Proc As PROCESS_INFORMATION, Start As STARTUPINFO
    
    Start.cb = Len(Start):    If Hide Then Start.dwFlags = STARTF_USESHOWWINDOW:   Start.wShowWindow = SW_HIDE
    
    CreateProcessW 0, StrPtr(CommandLine), 0, 0, 1, NORMAL_PRIORITY_CLASS, 0, 0, Start, Proc
    Call WaitForSingleObject(Proc.hProcess, Timeout)
    If Proc.hProcess = 0 Then ShellSync = 2:   Exit Function
    GetExitCodeProcess Proc.hProcess, ShellSync
    CloseHandle Proc.hProcess
End Function

Function ShellSyncEx(ByVal FileName As String, Optional ByVal CommandLine As String, Optional ByVal Timeout As Long = -1, Optional ByVal flagShow As Long = 0, Optional ByVal lpVerb As String = vbNullString) As Long
    Dim SEI As SHELLEXECUTEINFO
 
    With SEI
       .cbSize = LenB(SEI)
       .fMask = SEE_MASK_NOCLOSEPROCESS Or SEE_MASK_NOASYNC
       .lpVerb = lpVerb
       .lpFile = FileName
       .lpParameters = CommandLine
       .lpDirectory = vbNullChar
       .nShow = flagShow
    End With
    
    Call ShellExecuteExW(VarPtr(SEI))
    
    ShellSyncEx = WaitForSingleObject(SEI.hProcess, Timeout)
End Function

Function RegisterDLL(ByVal FileName As String, Optional ByVal isReg As Boolean = True) As Boolean
    Dim hLib As Long, hProc As Long, f As New clsFuncPointer
    
    If LenB(FileName) = 0 Then Exit Function
    
    hLib = LoadLibrary(StrPtr(FileName)):    If hLib = 0 Then Exit Function
    
    If isReg Then hProc = GetProcAddress(hLib, "DllRegisterServer") Else hProc = GetProcAddress(hLib, "DllUnregisterServer")
    If hProc <> 0 Then If f.PCall(hProc) = 0 Then RegisterDLL = True
    If hLib Then Call FreeLibrary(hLib)
End Function

Function VersionDLL(ByVal FileName As String, Optional ByVal verCmp As Variant, Optional ByVal verProd As Boolean) As Variant
    Dim sz As Long, p As Long, txt As String, v1() As String, v2() As String, b() As Byte, ver As VS_FIXEDFILEINFO

    sz = GetFileVersionInfoSizeW(StrPtr(LongPath(FileName)), ByVal 0&):     If sz = 0 Then Exit Function

    ReDim b(sz) As Byte
    Call GetFileVersionInfoW(StrPtr(LongPath(FileName)), 0&, sz, b(0))
    Call VerQueryValueW(b(0), StrPtr("\"), p, sz)
    Call CopyMemory(ver, ByVal p, Len(ver))

    If verProd = True Then txt = ver.dwProductVersionMSh & "." & ver.dwProductVersionMSl & "." & ver.dwProductVersionLSh & "." & ver.dwProductVersionLSl
    If verProd = False Then txt = ver.dwFileVersionMSh & "." & ver.dwFileVersionMSl & "." & ver.dwFileVersionLSh & "." & ver.dwFileVersionLSl
    
    If IsMissing(verCmp) Or IsNull(verCmp) Then VersionDLL = txt:    Exit Function
    
    v1 = Split(txt, "."):    VersionDLL = True
    Select Case Left$(verCmp, 1)
        Case ">"
            v2 = Split(Mid$(verCmp, 2), ".", 4)
            For p = 0 To UBound(v2)
                sz = Val(v1(p)) - Val(v2(p))
                If sz < 0 Then VersionDLL = False:  Exit Function
                If sz > 0 Then Exit Function
            Next
            
        Case Else
            v2 = Split(verCmp, ".", 4)
            For p = 0 To UBound(v2)
                If Val(v1(p)) <> Val(v2(p)) Then VersionDLL = False:     Exit Function
            Next
    End Select
End Function

Function ExistsMember(ByVal Disp As ATL.IDispatch, ProcName As String) As Boolean
    If Disp Is Nothing Then Exit Function
    ExistsMember = (Disp.GetIDsOfNames(IID_Null, ProcName, 1, LOCALE_USER_DEFAULT, 0&) = S_OK)
End Function

Function GetFunc(Optional value As Variant) As Variant
    Dim t As Long, c As Long, fn As String, rt As String, txt As String, REG1 As RegExp, Mts As MatchCollection
    
    On Error Resume Next
    
    Err.Clear:          t = VariantType(value):         If IsMissing(value) Then Set GetFunc = Funcs:   Exit Function
    
    If t = vbObject Then fn = ObjPtr(value):    t = -1:     GetFunc = Funcs(fn)
    If t = vbString Then Set GetFunc = Funcs(value)
    If Err.Number = 0 Then Exit Function
    
    txt = value
    Set REG1 = New RegExp:          REG1.IgnoreCase = True:     REG1.Pattern = "^(\s*<([^>]+)>\s*)?(function)?\s*(\w*)\s*(\([^\)]*\))([\s\S]+)"
    Set Mts = REG1.Execute(txt):    REG1.Global = True:         REG1.Pattern = "\w+":         c = REG1.Execute(Mts(0).SubMatches(4)).Count
    
    If t = vbString Then
        If Mts.Count = 0 Then Set GetFunc = Nothing
        fn = Mts(0).SubMatches(3):      If LenB(fn) = 0 Then fn = "func_" & GenTempStr("uuid")
        rt = Mts(0).SubMatches(1):      If LenB(rt) = 0 Then rt = "result"
        txt = Replace$("Function " + fn + Mts(0).SubMatches(4) + vbCrLf + Mts(0).SubMatches(5) + vbCrLf + "End Function", rt, fn, , , vbTextCompare)
        CAS.Execute txt:                Set GetFunc = CAS.Eval("GetRef(""" + fn + """)")
        Funcs.Add GetFunc, value:       Funcs.Add Array(fn, c, value), CStr(ObjPtr(GetFunc))
    Else
        If Mts.Count = 0 Then c = 1 Else rt = Mts(0).SubMatches(3):   If LenB(txt) Then fn = rt
        GetFunc = Array(fn, c, txt)
    End If
    
    Err.Clear
End Function

Function CBN(Obj As Variant, ProcName As Variant, ByVal CallType As VbCallType, Optional ByVal Args As Variant, Optional ByVal cntArgs As Long = -1, Optional ByVal pvarResult As Long, Optional hr As Long) As Variant
    Dim Disp As ATL.IDispatch, pDispParams As ATL.DISPPARAMS, pexcepinfo As ATL.EXCEPINFO, puArgError As Long
    Dim idMember As Long, pNamed As Long, SA As SafeArray
    
    If IsObject(Obj) Then
        Set Disp = Obj
    Else
        If Not CAS Is Nothing Then Set Disp = CAS.CodeObject(Obj)
    End If
    
    If Disp Is Nothing Then Exit Function
    
    If VarType(ProcName) = vbString Then
        hr = Disp.GetIDsOfNames(IID_Null, CStr(ProcName), 1, LOCALE_USER_DEFAULT, idMember)
    Else
        idMember = CLng(ProcName)
    End If

    If hr = S_OK Then
        If cntArgs = -1 Then ArrayReverse Args

        SA = GetSafeArray(Args)
        
        If cntArgs > -1 Then pDispParams.cArgs = cntArgs Else pDispParams.cArgs = SA.rgSABound(0).cElements
        pDispParams.rgPointerToVariantArray = SA.pvData
        
        If CallType > 15 Then CallType = CallType And 15
        
        Select Case CallType
            Case VbGet, VbFunc
                If pvarResult = 0 Then pvarResult = VarPtr(CBN)
                
            Case VbLet, VbSet
                pNamed = DISPID_PROPERTYPUT
                pDispParams.cNamedArgs = 1
                pDispParams.rgPointerToLONGNamedArgs = VarPtr(pNamed)
            
            Case Is < 0
                CallType = -CallType
                If pvarResult = 0 Then pvarResult = VarPtr(CBN)
        End Select
        
        hr = Disp.Invoke(idMember, IID_Null, LOCALE_USER_DEFAULT, CallType, pDispParams, pvarResult, pexcepinfo, puArgError)
    End If
End Function

Function DoParams(ByVal Obj As Object, Arg As Variant) As Object
    Dim a As Long, b As Long, uds As Long, txt As String, nMod As String
    Static Preset(32) As String
    
    Set DoParams = Obj
    If VarType(Arg) = vbString Then Arg = Array(Arg)
    uds = ArraySize(Arg) - 1
    If uds < 0 Then Exit Function
    
    For a = 0 To uds
        Select Case VarType(Arg(a))
            Case vbString
                txt = txt + Arg(a) + vbCrLf
                
            Case vbArray + vbVariant
                Select Case ArraySize(Arg(a))
                    Case 1
                        nMod = Arg(a)(0)
                    Case Is > 1
                        For b = 0 To UBound(Arg(a)) Step 2
                            Preset(Arg(a)(b)) = Arg(a)(b + 1)
                        Next
                End Select
                
            Case vbError, vbEmpty
                Erase Preset()
        End Select
    Next
    
    Preset(0) = "ObjFromPtr(" & ObjPtr(Obj) & ")"
    
    If LenB(txt) = 0 Then Exit Function
    
    For a = 0 To UBound(Preset)
        If Len(Preset(a)) Then txt = Replace$(txt, "$" & a, Preset(a))
    Next

    CAS.Execute "With " & Preset(0) & vbCrLf & txt & vbCrLf & "End With", nMod
End Function

Sub ArrayDef(Param As Variant, ParamArray vsp() As Variant)
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

Sub ArrayReverse(arr As Variant)
    Dim SA As SafeArray, a As Long, uds As Long, NewArr() As Variant
    
    SA = GetSafeArray(arr):    uds = SA.rgSABound(0).cElements - 1:    If uds <= 0 Or SA.cDims <> 1 Then Exit Sub

    ReDim NewArr(uds)
    For a = 0 To uds
        VariantCopy NewArr(a), arr(uds - a)
    Next
        
    arr = NewArr
End Sub

Function GetSafeArray(arr As Variant, Optional vt As Integer, Optional ptrSA As Long) As SafeArray
    Dim Ptr As Long, cDims As Integer
    
    Ptr = VarPtr(arr):    GetMem2 Ptr, vt
    
    If vt = VT_BYREF + VT_VARIANT Then GetMem4 Ptr + 8, Ptr:   GetMem2 Ptr, vt

    If (vt And VT_ARRAY) Then
        GetMem4 Ptr + 8, Ptr
        If Ptr <> 0 Then If (vt And VT_BYREF) Then GetMem4 Ptr, Ptr
        If Ptr <> 0 Then ptrSA = Ptr:   GetMem2 Ptr, cDims:   CopyMemory GetSafeArray, ByVal Ptr, 16 + cDims * 8
    End If
End Function

Function ArraySize(arr As Variant) As Long
    With GetSafeArray(arr)
        If .cDims = 1 Then ArraySize = .rgSABound(0).cElements
    End With
End Function

Function ArrayValid(arr As Variant, Optional iL As Long, Optional iU As Long, Optional ByVal minCount As Long = 1, Optional ByVal maxCount As Long = 0, Optional vt As Integer, Optional ByVal ZeroiL As Boolean = True) As Boolean
    Dim sz As Long, SA As SafeArray
    
    SA = GetSafeArray(arr, vt)
    If SA.cDims <> 1 Then Exit Function
    sz = SA.rgSABound(0).cElements
    iL = SA.rgSABound(0).lLbound
    iU = sz + iL - 1
    If ZeroiL Then If iL <> 0 Then Exit Function
    If sz < minCount Then Exit Function
    If maxCount > 0 Then If sz > maxCount Then Exit Function

    ArrayValid = True
End Function



'======================== String Sort ================================
Sub InsertSortStringsStart(ListArray() As String, Optional ByVal bAscending As Boolean = True, Optional ByVal bCaseSensitive As Boolean = False)
    Dim lMin As Long, lMax As Long, lOrder As Long, lCompareType As Long
    
    lMin = LBound(ListArray):           lMax = UBound(ListArray):       If lMin = lMax Then Exit Sub
    lOrder = IIF(bAscending, -1, 1):    lCompareType = IIF(bCaseSensitive, vbBinaryCompare, vbTextCompare)
    
    InsertSortStrings ListArray, lMin, lMax, lOrder, lCompareType
End Sub


Private Sub InsertSortStrings(ListArray() As String, ByVal lMin As Long, ByVal lMax As Long, ByVal lOrder As Long, ByVal lCompareType As Long)
    Dim sValue As String, lCount1 As Long, lCount2 As Long
    
    For lCount1 = lMin + 1 To lMax
        sValue = ListArray(lCount1)
        
        For lCount2 = lCount1 - 1 To lMin Step -1
            If StrComp(ListArray(lCount2), sValue, lCompareType) <> lOrder Then Exit For
            ListArray(lCount2 + 1) = ListArray(lCount2)
        Next
        
        ListArray(lCount2 + 1) = sValue
    Next
End Sub

Sub QuickSortStringsStart(ListArray() As String, Optional ByVal bAscending As Boolean = True, Optional ByVal bCaseSensitive As Boolean = False)
    Dim lMin As Long, lMax As Long, lOrder As Long, lCompareType As Long

    lMin = LBound(ListArray):           lMax = UBound(ListArray):       If lMin = lMax Then Exit Sub
    lOrder = IIF(bAscending, 1, -1):    lCompareType = IIF(bCaseSensitive, vbBinaryCompare, vbTextCompare)
    
    QuickSortStrings ListArray, lMin, lMax, lOrder, lCompareType
End Sub

Private Sub QuickSortStrings(ListArray() As String, ByVal lLowerPoint As Long, ByVal lUpperPoint As Long, ByVal lOrder As Long, ByVal lCompareType As Long)
    Const DELEGATE_POINT As Long = 60
    Dim lMidPoint As Long
    
    If (lUpperPoint - lLowerPoint) <= DELEGATE_POINT Then
        InsertSortStrings ListArray, lLowerPoint, lUpperPoint, lOrder, lCompareType
        Exit Sub
    End If

    Do While lLowerPoint < lUpperPoint
        lMidPoint = QuickSortStringsPartition(ListArray, lLowerPoint, lUpperPoint, lOrder, lCompareType)
        
        If (lMidPoint - lLowerPoint) <= (lUpperPoint - lMidPoint) Then
            QuickSortStrings ListArray, lLowerPoint, lMidPoint - 1, lOrder, lCompareType
            lLowerPoint = lMidPoint + 1
        Else
            QuickSortStrings ListArray, lMidPoint + 1, lUpperPoint, lOrder, lCompareType
            lUpperPoint = lMidPoint - 1
        End If
    Loop
End Sub

Private Function QuickSortStringsPartition(ListArray() As String, ByVal lLow As Long, ByVal lHigh As Long, ByVal lOrder As Long, ByVal lCompareType As Long) As Long
    Dim lPivot As Long, sPivot As String, lLowCount As Long, lHighCount As Long, sTemp As String

    ' Select pivot point and exchange with first element
    lPivot = lLow + (lHigh - lLow) \ 2:     sPivot = ListArray(lPivot):     ListArray(lPivot) = ListArray(lLow)
    
    lLowCount = lLow + 1:       lHighCount = lHigh
    
    ' Continually loop moving entries smaller than pivot to One side and
    ' larger than pivot to other side
    Do
        Do While lLowCount < lHighCount
            If StrComp(sPivot, ListArray(lLowCount), lCompareType) <> lOrder Then Exit Do
            lLowCount = lLowCount + 1
        Loop
        
        Do While lHighCount >= lLowCount
            If StrComp(ListArray(lHighCount), sPivot, lCompareType) <> lOrder Then Exit Do
            lHighCount = lHighCount - 1
        Loop
        
        If lLowCount >= lHighCount Then Exit Do
        
        ' Swap the items
        sTemp = ListArray(lLowCount)
        ListArray(lLowCount) = ListArray(lHighCount)
        ListArray(lHighCount) = sTemp
        
        lHighCount = lHighCount - 1:      lLowCount = lLowCount + 1
    Loop
    
    ListArray(lLow) = ListArray(lHighCount)
    ListArray(lHighCount) = sPivot
    QuickSortStringsPartition = lHighCount
End Function
