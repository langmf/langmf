VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNativeRes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function LoadResource Lib "kernel32" (ByVal hInstance As Long, ByVal hResInfo As Long) As Long
Private Declare Function FindResource Lib "kernel32" Alias "FindResourceA" (ByVal hInstance As Long, ByVal lpName As String, ByVal lpType As Long) As Long
Private Declare Function LockResource Lib "kernel32" (ByVal hResData As Long) As Long
Private Declare Function LoadLibraryEx Lib "kernel32" Alias "LoadLibraryExW" (ByVal lpLibFileName As Long, ByVal hFile As Long, ByVal dwFlags As Long) As Long
Private Declare Function SizeofResource Lib "kernel32" (ByVal hInstance As Long, ByVal hResInfo As Long) As Long
Private Declare Function UpdateResource Lib "kernel32" Alias "UpdateResourceA" (ByVal hUpdate As Long, ByVal lpType As Long, ByVal lpName As Long, ByVal wLanguage As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function FindResourceEx Lib "kernel32" Alias "FindResourceExA" (ByVal hModule As Long, ByVal lpType As Long, ByVal lpName As String, ByVal wLanguage As Long) As Long
Private Declare Function EnumResourceNames Lib "kernel32" Alias "EnumResourceNamesA" (ByVal hModule As Long, ByVal lpType As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function EndUpdateResource Lib "kernel32" Alias "EndUpdateResourceA" (ByVal hUpdate As Long, ByVal fDiscard As Long) As Long
Private Declare Function BeginUpdateResource Lib "kernel32" Alias "BeginUpdateResourceA" (ByVal pFileName As String, ByVal bDeleteExistingResources As Long) As Long

Const LOAD_LIBRARY_AS_DATAFILE As Long = 2

Const RT_CURSOR = 1, RT_BITMAP = 2, RT_ICON = 3, RT_MENU = 4, RT_DIALOG = 5, RT_STRING = 6, RT_RCDATA = 10
Const RT_GROUP_CURSOR = 12, RT_GROUP_ICON = 14, RT_VERSION = 16, RT_HTML = 23, RT_MANIFEST = 24

Private Type NRVar
  wLength As Integer
  wValueLength As Integer
  wType As Integer
  szKey(21) As Byte
  Padding(1) As Integer
  'Value() As Long
End Type

Private Type VarFileInfo
  wLength As Integer
  wValueLength As Integer
  wType As Integer
  szKey(21) As Byte
  Padding(1) As Integer
  'Children() As NRVar
End Type

Private Type def_StringData
  szKey() As Byte
  value() As Byte
End Type

Private Type def_String
  wLength As Integer
  wValueLength As Integer
  wType As Integer
  'szKey As String
  'Padding(1) As Integer
  'Value() As Byte
End Type

Private Type StringTable
  wLength As Integer
  wValueLength As Integer
  wType As Integer
  szKey(15) As Byte
  'Padding(0) As Integer
  'Value() As def_String
End Type

Private Type StringTables
  wLength As Integer
  wValueLength As Integer
  wType As Integer
  szKey(15) As Byte
  m_SFI_Str() As def_String
  m_SFI_StrData() As def_StringData
End Type

Private Type StringFileInfo
  wLength As Integer
  wValueLength As Integer
  wType As Integer
  szKey(27) As Byte
  'Padding(0) As Integer
  'Children() As StringTable
End Type

Private Type VS_FIXEDFILEINFO
    dwSignature As Long
    dwStrucVersion As Long
    dwFileVersionMS As Long
    dwFileVersionLS As Long
    dwProductVersionMS As Long
    dwProductVersionLS As Long
    dwFileFlagsMask As Long
    dwFileFlags As Long
    dwFileOS As Long
    dwFileType As Long
    dwFileSubtype As Long
    dwFileDateMS As Long
    dwFileDateLS As Long
End Type

Private Type VS_VERSIONINFO
    wLength As Integer
    wValueLength As Integer
    wType  As Integer
    szKey(29) As Byte
    Padding1(1) As Integer
    value As VS_FIXEDFILEINFO
'    Padding2() As Integer
'    Children() As Integer
End Type

Private Type def_IconData
    Buf() As Byte
End Type

Private Type RESICONDIRENTRY
   bWidth As Byte               'Width of the image
   bHeight As Byte              'Height of the image (times 2)
   bColorCount As Byte          'Number of colors in image (0 if >=8bpp)
   bReserved As Byte            'Reserved
   wPlanes As Integer           'Color Planes
   wBitCount As Integer         'Bits per pixel
   dwBytesInRes As Long         'how many bytes in this resource?
   nID As Integer               'where in the file is this image
End Type


Public Function UpdateString(ByVal nameExeFile As String, ByVal numRes As Long, ByVal szTxt As String, Optional ByVal wLang As Long = 0) As Long
    Dim rc As Long, hLib As Long, hGlobal As Long, hRsrc As Long, lPtr As Long, lenTxt As Integer, numStr As Long
    Dim dlHw As Long, resID As Long, cntStr As Long, ofs As Long, Buf() As Byte, txt(16) As String
    
    resID = (numRes \ 16) + 1
    numStr = (numRes Mod 16) + 1
    
    '------------------------чтение ресурса------------------------------
    hLib = LoadLibraryEx(StrPtr(nameExeFile), 0, LOAD_LIBRARY_AS_DATAFILE)
    If wLang = 0 Then
        hRsrc = FindResource(hLib, "#" & resID, RT_STRING)
    Else
        hRsrc = FindResourceEx(hLib, RT_STRING, "#" & resID, wLang)
    End If
    
    If hRsrc <> 0 Then
        hGlobal = LoadResource(hLib, hRsrc)
        lPtr = LockResource(hGlobal)
        rc = SizeofResource(hLib, hRsrc)
        
        ReDim Buf(rc - 1)
        
        CopyMemory Buf(0), ByVal lPtr, UBound(Buf) + 1
        
        '----------------------Чтение таблицы----------------------------
        Do
            cntStr = cntStr + 1
            
            CopyMemory lenTxt, Buf(ofs), 2
            ofs = ofs + 2
            
            If lenTxt <> 0 Then
                txt(cntStr) = String$(lenTxt * 2, 0)
                CopyMemory ByVal txt(cntStr), Buf(ofs), lenTxt * 2
                ofs = ofs + lenTxt * 2
            End If
        Loop Until cntStr >= 16
    End If
    
    FreeLibrary hLib
    
    '----------------------Присвоение текста-------------------------
    txt(numStr) = StrConv(szTxt, vbUnicode)
    
    '----------------------Запись таблицы----------------------------
    cntStr = 0
    ofs = 0
        
    Do
        cntStr = cntStr + 1
        
        lenTxt = Len(txt(cntStr)) \ 2
        
        ReDim Preserve Buf(ofs + 2 - 1)
        
        CopyMemory Buf(ofs), lenTxt, 2
        ofs = ofs + 2
        
        If lenTxt <> 0 Then
            ReDim Preserve Buf(ofs + lenTxt * 2 - 1)
        
            CopyMemory Buf(ofs), ByVal txt(cntStr), lenTxt * 2
            ofs = ofs + lenTxt * 2
        End If
    Loop Until cntStr >= 16

    '----------------------Обновляем ресурс-----------------------------
    dlHw = BeginUpdateResource(nameExeFile, False)

    UpdateString = UpdateResource(dlHw, RT_STRING, resID, wLang, Buf(0), UBound(Buf) + 1)

    Call EndUpdateResource(dlHw, False)
End Function

Public Function UpdateMainIcon(ByVal nameExeFile As String, ByVal nameIconFile As String, Optional ByVal startID As Long = 30001, Optional ByVal resID As Long = 1, Optional ByVal wLanguage As Long = 0) As Long
    Dim a As Long, dlHw As Long, nF As Integer, Buf() As Byte, iconData() As def_IconData
    Dim iconRDE() As RESICONDIRENTRY, iconDE() As ICONDIRENTRY, iconHdr As ICONDIR

    nF = FreeFile
    Open nameIconFile For Binary As nF
        Get nF, , iconHdr
    
        ReDim iconDE(iconHdr.idCount - 1)
        ReDim iconRDE(iconHdr.idCount - 1)
        ReDim iconData(iconHdr.idCount - 1)
    
        Get nF, , iconDE
    
        For a = 0 To iconHdr.idCount - 1
            ReDim iconData(a).Buf(iconDE(a).dwBytesInRes - 1)
            Get nF, iconDE(a).dwImageOffset + 1, iconData(a).Buf
    
            CopyMemory iconRDE(a), iconDE(a), Len(iconRDE(a))
    
            iconRDE(a).nID = startID
    
            startID = startID + 1
        Next
    Close nF
    
    
    ReDim Buf(Len(iconHdr) + Len(iconRDE(0)) * iconHdr.idCount)
    
    CopyMemory Buf(0), iconHdr, Len(iconHdr)
    
    For a = 0 To iconHdr.idCount - 1
        CopyMemory Buf(Len(iconHdr) + Len(iconRDE(0)) * a), iconRDE(a), Len(iconRDE(0))
    Next
    
    '----------------------Обновляем ресурс-----------------------------
    dlHw = BeginUpdateResource(nameExeFile, False)
    
    UpdateMainIcon = UpdateResource(dlHw, RT_GROUP_ICON, resID, wLanguage, Buf(0), UBound(Buf) + 1)
    
    For a = 0 To iconHdr.idCount - 1
        Call UpdateResource(dlHw, RT_ICON, CLng(iconRDE(a).nID), wLanguage, iconData(a).Buf(0), UBound(iconData(a).Buf) + 1)
    Next
    
    Call EndUpdateResource(dlHw, False)
End Function

Public Function UpdateVersion(ByVal nameExeFile As String, dataVer As Variant, Optional ByVal resID As Long = 1) As Long
    Dim hLib As Long, hGlobal As Long, hRsrc As Long, lPtr As Long, rc As Long, dlHw As Long, isWrite As Boolean
    Dim Buf() As Byte, txtKey As String, txtValue As String, txt As String, v As Variant, v2 As Variant
    Dim sizeBuf As Long, a As Long, b As Long, numHdr As Long, numSD As Long, szHdr As Long, tmpLng As Long
    Dim ofs As Long, offSet_Start As Long, offSet_Start_KV As Long, offSet_Start_SFI As Long
    Dim offSet_End_SFI As Long, offSet_Start_SFI_Hdr As Long, korrLen As Long, korrAllLen As Long
    
    Dim b_VFI As Boolean, m_VFI As VarFileInfo, m_VFI_Hdr As NRVar, m_VFI_Value() As Long
    Dim b_SFI As Boolean, m_SFI As StringFileInfo, m_SFI_Hdr() As StringTables, m_SFI_Hdr_Len As StringTable
    
    Dim m_VS As VS_VERSIONINFO, txtVFI As String * 22, txtSFI As String * 28

    ReDim m_VFI_Value(0)
    ReDim m_SFI_Hdr(0)
    
    
    If dataVer Is Nothing Then
        Set dataVer = New clsHash
    ElseIf TypeOf dataVer Is clsHash Then
        If dataVer.Count Then isWrite = True
    Else
        Set dataVer = New clsHash
    End If
    
    '------------------------чтение ресурса------------------------------
    hLib = LoadLibraryEx(StrPtr(nameExeFile), 0, LOAD_LIBRARY_AS_DATAFILE)
    
    If CLng(dataVer("FindLng")) = 0 Then
        hRsrc = FindResource(hLib, "#" & resID, RT_VERSION)
    Else
        hRsrc = FindResourceEx(hLib, RT_VERSION, "#" & resID, CLng(dataVer("FindLng")))
    End If
    
    If hLib = 0 Or hRsrc = 0 Then
        UpdateVersion = -1
        FreeLibrary hLib
        Exit Function
    End If
    
    hGlobal = LoadResource(hLib, hRsrc)
    lPtr = LockResource(hGlobal)
    rc = SizeofResource(hLib, hRsrc)
    
    ReDim Buf(rc - 1)
    
    CopyMemory Buf(0), ByVal lPtr, UBound(Buf) + 1
    
    FreeLibrary hLib
    
    
    '--------------------------- Parse Resource Version -------------------------
    CopyMemory m_VS, Buf(0), Len(m_VS)
    ofs = Len(m_VS)
    
    With m_VS.value
        If isWrite = False Then
            dataVer("dwSignature") = .dwSignature
            dataVer("dwStrucVersion") = .dwStrucVersion
            dataVer("dwFileVersionMS") = .dwFileVersionMS
            dataVer("dwFileVersionLS") = .dwFileVersionLS
            dataVer("dwProductVersionMS") = .dwProductVersionMS
            dataVer("dwProductVersionLS") = .dwProductVersionLS
            dataVer("dwFileFlagsMask") = .dwFileFlagsMask
            dataVer("dwFileFlags") = .dwFileFlags
            dataVer("dwFileOS") = .dwFileOS
            dataVer("dwFileType") = .dwFileType
            dataVer("dwFileSubtype") = .dwFileSubtype
            dataVer("dwFileDateMS") = .dwFileDateMS
            dataVer("dwFileDateLS") = .dwFileDateLS
        Else
            .dwSignature = CLng(dataVer("dwSignature"))
            .dwStrucVersion = CLng(dataVer("dwStrucVersion"))
            .dwFileVersionMS = CLng(dataVer("dwFileVersionMS"))
            .dwFileVersionLS = CLng(dataVer("dwFileVersionLS"))
            .dwProductVersionMS = CLng(dataVer("dwProductVersionMS"))
            .dwProductVersionLS = CLng(dataVer("dwProductVersionLS"))
            .dwFileFlagsMask = CLng(dataVer("dwFileFlagsMask"))
            .dwFileFlags = CLng(dataVer("dwFileFlags"))
            .dwFileOS = CLng(dataVer("dwFileOS"))
            .dwFileType = CLng(dataVer("dwFileType"))
            .dwFileSubtype = CLng(dataVer("dwFileSubtype"))
            .dwFileDateMS = CLng(dataVer("dwFileDateMS"))
            .dwFileDateLS = CLng(dataVer("dwFileDateLS"))
        End If
    End With
    
    Do
        '----------------------- Parse VarFileInfo Struct ------------------------
        If b_VFI = False Then
            CopyMemory ByVal txtVFI, Buf(ofs + 6), Len(txtVFI)
            
            txtVFI = StrConv(txtVFI, vbFromUnicode)
            
            If InStr(txtVFI, "VarFileInfo") > 0 Then
                CopyMemory m_VFI, Buf(ofs), Len(m_VFI)
                ofs = ofs + Len(m_VFI)
                
                CopyMemory m_VFI_Hdr, Buf(ofs), Len(m_VFI_Hdr)
                ofs = ofs + Len(m_VFI_Hdr)
                
                For a = 0 To m_VFI_Hdr.wValueLength Step 5
                    CopyMemory m_VFI_Value(UBound(m_VFI_Value)), Buf(ofs), 4
                    ofs = ofs + 4
                    
                    ReDim Preserve m_VFI_Value(UBound(m_VFI_Value) + 1)
                Next
                
                ReDim Preserve m_VFI_Value(UBound(m_VFI_Value) - 1)
                
                b_VFI = True
            End If
        End If
    
        '----------------------- Parse StringFileInfo Struct ------------------------
        If b_SFI = False Then
            CopyMemory ByVal txtSFI, Buf(ofs + 6), Len(txtSFI)
            
            txtSFI = StrConv(txtSFI, vbFromUnicode)
            
            If InStr(txtSFI, "StringFileInfo") > 0 Then
                offSet_Start_SFI = ofs
                
                CopyMemory m_SFI, Buf(ofs), Len(m_SFI)
                ofs = ofs + Len(m_SFI)
                
                Padding ofs
                
                offSet_End_SFI = ofs
                
                Do
                    offSet_Start = ofs
                    
                    numHdr = UBound(m_SFI_Hdr)
                    
                    CopyMemory m_SFI_Hdr(numHdr), Buf(ofs), Len(m_SFI_Hdr_Len)
                    ofs = ofs + Len(m_SFI_Hdr_Len)
                    
                    Padding ofs
                    
                    ReDim m_SFI_Hdr(numHdr).m_SFI_Str(0)
                    ReDim m_SFI_Hdr(numHdr).m_SFI_StrData(0)
                    
                    Do
                        numSD = UBound(m_SFI_Hdr(numHdr).m_SFI_Str)
                        
                        CopyMemory m_SFI_Hdr(numHdr).m_SFI_Str(numSD), Buf(ofs), Len(m_SFI_Hdr(numHdr).m_SFI_Str(0))
                        ofs = ofs + Len(m_SFI_Hdr(numHdr).m_SFI_Str(0))

                        offSet_Start_KV = ofs
                        
                        For a = ofs To ofs + m_SFI_Hdr(numHdr).m_SFI_Str(numSD).wLength Step 2
                            If Buf(a) = 0 And Buf(a + 1) = 0 Then
                                ReDim m_SFI_Hdr(numHdr).m_SFI_StrData(numSD).szKey(a - ofs - 1)
                        
                                CopyMemory m_SFI_Hdr(numHdr).m_SFI_StrData(numSD).szKey(0), Buf(ofs), UBound(m_SFI_Hdr(numHdr).m_SFI_StrData(numSD).szKey) + 1
                                ofs = ofs + UBound(m_SFI_Hdr(numHdr).m_SFI_StrData(numSD).szKey) + 2
                                
                                Padding ofs
                                
                                Exit For
                            End If
                        Next
                        
                        For a = ofs To offSet_Start_KV + m_SFI_Hdr(numHdr).m_SFI_Str(numSD).wLength Step 2
                            If Buf(a) = 0 And Buf(a + 1) = 0 Then
                                ReDim m_SFI_Hdr(numHdr).m_SFI_StrData(numSD).value(a - ofs + 2 - 1)
                        
                                CopyMemory m_SFI_Hdr(numHdr).m_SFI_StrData(numSD).value(0), Buf(ofs), UBound(m_SFI_Hdr(numHdr).m_SFI_StrData(numSD).value) + 1
                                ofs = ofs + UBound(m_SFI_Hdr(numHdr).m_SFI_StrData(numSD).value)
                                
                                Padding ofs
                                
                                Exit For
                            End If
                        Next
                        
                        'txtKey = m_SFI_Hdr(numHdr).m_SFI_StrData(numSD).szKey
                        'txtValue = m_SFI_Hdr(numHdr).m_SFI_StrData(numSD).value
                            
                        ReDim Preserve m_SFI_Hdr(numHdr).m_SFI_Str(numSD + 1)
                        ReDim Preserve m_SFI_Hdr(numHdr).m_SFI_StrData(numSD + 1)
                        
                        korrLen = (ofs - offSet_Start) - m_SFI_Hdr(numHdr).wLength
                    Loop Until korrLen >= 0 And korrLen <= 2    '2 - корректировка у неправильных программ
                    
                    korrAllLen = korrAllLen + korrLen
                    
                    ReDim Preserve m_SFI_Hdr(numHdr).m_SFI_Str(UBound(m_SFI_Hdr(numHdr).m_SFI_Str) - 1)
                    ReDim Preserve m_SFI_Hdr(numHdr).m_SFI_StrData(UBound(m_SFI_Hdr(numHdr).m_SFI_StrData) - 1)
                    
                    szHdr = 0
                    For a = 0 To numHdr
                        szHdr = szHdr + m_SFI_Hdr(a).wLength
                    Next
                    szHdr = szHdr + korrAllLen
                    
                    ReDim Preserve m_SFI_Hdr(numHdr + 1)
                Loop Until szHdr >= m_SFI.wLength - (offSet_End_SFI - offSet_Start_SFI)
                
                ReDim Preserve m_SFI_Hdr(UBound(m_SFI_Hdr) - 1)
                
                b_SFI = True
            End If
        End If
    Loop Until ofs >= m_VS.wLength
    
    If isWrite = False Then
        dataVer.Child "Tables"
        
        For a = 0 To UBound(m_SFI_Hdr)
            txt = m_SFI_Hdr(a).szKey
            
            With dataVer("Tables").Child(txt)
                For b = 0 To UBound(m_SFI_Hdr(a).m_SFI_Str)
                    txtKey = m_SFI_Hdr(a).m_SFI_StrData(b).szKey
                    txtValue = m_SFI_Hdr(a).m_SFI_StrData(b).value
                    txtValue = Left$(txtValue, Len(txtValue) - 1)
                    .Add txtValue, txtKey
                Next
            End With
        Next

        With dataVer.Child("EnumLng")
            For a = 0 To UBound(m_VFI_Value)
                tmpLng = m_VFI_Value(a)
                tmpLng = (tmpLng Mod 65536) * 65536 + (tmpLng \ 65536)
                txt = Hex$(tmpLng)
                txt = String$(8 - Len(txt), "0") & txt
                .Add txt, txt
            Next
        End With
    Else
        'Примерное значение всех фиксированных структур
        sizeBuf = 256
        
        'Заполнение структуры m_SFI_StrData данными из коллекции dataVer
        ReDim m_SFI_Hdr(dataVer("Tables").Count - 1)
        
        a = 0
        
        For Each v In dataVer("Tables").Items
            sizeBuf = sizeBuf + Len(m_SFI_Hdr_Len)
            
            With m_SFI_Hdr(a)
                ReDim Preserve .m_SFI_Str(v.Count - 1)
                ReDim Preserve .m_SFI_StrData(v.Count - 1)
                    
                b = 0
                
                For Each v2 In v.Items(2)
                    '--- Примерное значение всех данных ---
                    sizeBuf = sizeBuf + 6
                    sizeBuf = sizeBuf + Len(v2(1)) * 2 + 4
                    sizeBuf = sizeBuf + Len(v2(0)) * 2 + 4
                    '--------------------------------------
                
                    txtKey = v2(1)
                    txtKey = StrConv(txtKey, vbUnicode)
                    txtValue = v2(0)
                    txtValue = StrConv(txtValue & vbNullChar, vbUnicode)
                    
                    .m_SFI_Str(b).wType = 1
        
                    ReDim .m_SFI_StrData(b).szKey(Len(txtKey) - 1)
                    CopyMemory .m_SFI_StrData(b).szKey(0), ByVal txtKey, Len(txtKey)
        
                    ReDim .m_SFI_StrData(b).value(Len(txtValue) - 1)
                    CopyMemory .m_SFI_StrData(b).value(0), ByVal txtValue, Len(txtValue)
        
                    b = b + 1
                Next
            End With
            a = a + 1
        Next

        'Примерно задаем размер буфера
        ReDim Buf(sizeBuf)
        
        'Заполняем перечисление языков
        a = 0
        ReDim m_VFI_Value(dataVer("EnumLng").Count - 1)
        For Each v In dataVer("EnumLng").Items
            tmpLng = Val("&H" & v)
            tmpLng = (tmpLng Mod 65536) * 65536 + (tmpLng \ 65536)
            m_VFI_Value(a) = tmpLng
            a = a + 1
        Next
        
        'Посчитываем размер всех фиксированных структур
        ofs = Len(m_VS) + Len(m_VFI) + Len(m_VFI_Hdr) + (UBound(m_VFI_Value) + 1) * 4

        'Запоминаем начало структуры StringFileInfo
        offSet_Start_SFI = ofs

        'Добавляем к общему размеру структур - структуру StringFileInfo
        ofs = ofs + Len(m_SFI)
        Padding ofs
        
        a = 0
        
        For Each v In dataVer("Tables").Items(1)
            'Запоминаем начало структуры StringTable
            offSet_Start_SFI_Hdr = ofs
            
            'Добавляем к общему размеру структур - структуру StringTable
            ofs = ofs + Len(m_SFI_Hdr(a)) - 8
            Padding ofs
    
            'Заполняем все данные структуры StringFileInfo
            For b = 0 To UBound(m_SFI_Hdr(a).m_SFI_Str)
                offSet_Start = ofs
    
                ofs = ofs + Len(m_SFI_Hdr(a).m_SFI_Str(b))

                CopyMemory Buf(ofs), m_SFI_Hdr(a).m_SFI_StrData(b).szKey(0), UBound(m_SFI_Hdr(a).m_SFI_StrData(b).szKey) + 1
                ofs = ofs + UBound(m_SFI_Hdr(a).m_SFI_StrData(b).szKey) + 1
    
                'Исправляет если ofs и так выровнен.
                ofs = ofs + 1
    
                Padding ofs
    
                CopyMemory Buf(ofs), m_SFI_Hdr(a).m_SFI_StrData(b).value(0), UBound(m_SFI_Hdr(a).m_SFI_StrData(b).value) + 1
                ofs = ofs + UBound(m_SFI_Hdr(a).m_SFI_StrData(b).value) + 1
    
                Padding ofs
    
                m_SFI_Hdr(a).m_SFI_Str(b).wLength = ofs - offSet_Start
                m_SFI_Hdr(a).m_SFI_Str(b).wValueLength = UBound(m_SFI_Hdr(a).m_SFI_StrData(b).value) + 1
                CopyMemory Buf(offSet_Start), m_SFI_Hdr(a).m_SFI_Str(b), Len(m_SFI_Hdr(a).m_SFI_Str(b))
            Next
    
            Padding ofs
            
            txt = StrConv(v, vbUnicode)
            CopyMemory m_SFI_Hdr(a).szKey(0), ByVal txt, Len(txt)
            
            m_SFI_Hdr(a).wLength = ofs - offSet_Start_SFI_Hdr
            m_SFI_Hdr(a).wType = 1
            CopyMemory m_SFI_Hdr(a).szKey(0), ByVal txt, Len(txt)
            CopyMemory Buf(offSet_Start_SFI_Hdr), m_SFI_Hdr(a), Len(m_SFI_Hdr(a)) - 8
            
            a = a + 1
        Next
        
        'Обрезаем лишнее
        ReDim Preserve Buf(ofs - 1)

        'Заполняем длины структур
        m_VFI_Hdr.wValueLength = (UBound(m_VFI_Value) + 1) * 4
        m_VFI_Hdr.wLength = Len(m_VFI_Hdr) + m_VFI_Hdr.wValueLength
        m_VFI.wLength = Len(m_VFI) + m_VFI_Hdr.wLength
        m_SFI.wLength = ofs - offSet_Start_SFI
        m_VS.wLength = ofs

        'Записываем сами структуры
        ofs = 0

        CopyMemory Buf(ofs), m_VS, Len(m_VS)
        ofs = ofs + Len(m_VS)

        CopyMemory Buf(ofs), m_VFI, Len(m_VFI)
        ofs = ofs + Len(m_VFI)

        CopyMemory Buf(ofs), m_VFI_Hdr, Len(m_VFI_Hdr)
        ofs = ofs + Len(m_VFI_Hdr)

        CopyMemory Buf(ofs), m_VFI_Value(0), (UBound(m_VFI_Value) + 1) * 4
        ofs = ofs + (UBound(m_VFI_Value) + 1) * 4

        CopyMemory Buf(ofs), m_SFI, Len(m_SFI)

        '----------------------Обновляем ресурс-----------------------------
        dlHw = BeginUpdateResource(nameExeFile, False)

        rc = UpdateResource(dlHw, RT_VERSION, resID, CLng(dataVer("SaveLng")), Buf(0), UBound(Buf) + 1)

        rc = EndUpdateResource(dlHw, False)
    End If
End Function

Public Function EnumResource(ByVal nameExeFile As String, ByVal lTypes As Long) As Variant
    Dim hLib As Long
    
    ReDim ArrNRes(0)
    
    hLib = LoadLibraryEx(StrPtr(nameExeFile), 0, LOAD_LIBRARY_AS_DATAFILE)
    If hLib = 0 Then Exit Function

    EnumResourceNames hLib, lTypes, AddressOf EnumNResCBK, 0
    FreeLibrary hLib
    
    If UBound(ArrNRes) > 0 Then
        ReDim Preserve ArrNRes(UBound(ArrNRes) - 1)
        EnumResource = ArrNRes
    Else
        EnumResource = Array()
    End If
End Function

Public Function GetResource(vrtBuf As Variant, ByVal nameExeFile As String, ByVal lTypes As Long, ByVal mNames As String, Optional ByVal wLang As Long) As Long
    Dim Buf() As Byte, hLib As Long, hGlobal As Long, hRsrc As Long, lPtr As Long, rc As Long
    
    hLib = LoadLibraryEx(StrPtr(nameExeFile), 0, LOAD_LIBRARY_AS_DATAFILE)
    If hLib = 0 Then Exit Function

    If wLang = 0 Then hRsrc = FindResource(hLib, mNames, lTypes) Else hRsrc = FindResourceEx(hLib, lTypes, mNames, wLang)

    If hRsrc <> 0 Then
        hGlobal = LoadResource(hLib, hRsrc)
        lPtr = LockResource(hGlobal)
        rc = SizeofResource(hLib, hRsrc)
        
        ReDim Buf(rc - 1)

        CopyMemory Buf(0), ByVal lPtr, rc
    End If

    FreeLibrary hLib
    vrtBuf = Buf
    
    GetResource = hRsrc
End Function

Public Function PutResourceFromFile(ByVal srcFileName As String, ByVal nameExeFile As String, ByVal mTypes As Long, ByVal mNames As Variant, Optional ByVal wLang As Long = 0, Optional ByVal bDelete As Boolean = False) As Long
    Dim Buf() As Byte
    If Not m_File2Buf(Buf, srcFileName) Then ReDim Buf(0)
    PutResourceFromFile = PutResource(Buf, nameExeFile, mTypes, mNames, wLang, bDelete)
End Function

Public Function PutResource(vrtBuf As Variant, ByVal nameExeFile As String, ByVal mTypes As Long, ByVal mNames As Variant, Optional ByVal wLang As Long = 0, Optional ByVal bDelete As Boolean = False) As Long
    Dim dlHw As Long, lNames As Long, sNames As String, bNames() As Byte, Buf() As Byte
    
    If VarType(mNames) = vbString Then
        sNames = mNames + vbNullChar
        ReDim bNames(Len(sNames) - 1)
        CopyMemory bNames(0), ByVal sNames, Len(sNames)
        lNames = VarPtr(bNames(0))
    Else
        lNames = mNames
    End If
    
    Buf = vrtBuf
    
    dlHw = BeginUpdateResource(nameExeFile, bDelete)

    PutResource = UpdateResource(dlHw, mTypes, lNames, wLang, Buf(0), UBound(Buf) + 1)

    Call EndUpdateResource(dlHw, False)
End Function


'----------------------------------------------------------------------------------------------------------------------
Public Function LoadResData(ID As Variant, VType As Variant, Optional ByVal isTrimNull As Boolean = False) As Byte()
    If isTrimNull Then LoadResData = LoadResDataWNull(ID, VType) Else LoadResData = VB.LoadResData(ID, VType)
End Function

Public Function LoadResPicture(ID As Variant, Optional resType As Variant) As IPictureDisp
    Dim rt As Integer
    Set LoadResPicture = VB.LoadResPicture(ID, rt)
    resType = rt
End Function

Public Function LoadResString(ByVal ID As Long) As String
    LoadResString = VB.LoadResString(ID)
End Function

Private Function Padding(vOffset As Long) As Boolean
    While vOffset Mod 4 <> 0
        vOffset = vOffset + 1
        Padding = True
    Wend
End Function
