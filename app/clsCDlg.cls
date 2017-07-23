VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function PrintDlgA Lib "comdlg32" (pPrintdlg As def_PRINTDLG) As Long
Private Declare Function ChooseColorW Lib "comdlg32" (pChooseColor As ChooseColor) As Long
Private Declare Function ChooseFontW Lib "comdlg32" (pChooseFont As ChooseFont) As Long
Private Declare Function GetSaveFileNameW Lib "comdlg32" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetOpenFileNameW Lib "comdlg32" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function SHBrowseForFolderW Lib "shell32" (lpBI As BROWSEINFO) As Long

Const SHGFI_DISPLAYNAME = &H200&, SHGFI_EXETYPE = &H2000&, SHGFI_SYSICONINDEX = &H4000&, SHGFI_LARGEICON = &H0&
Const SHGFI_SMALLICON = &H1&, SHGFI_SHELLICONSIZE = &H4&, SHGFI_TYPENAME = &H400&
Const BASIC_SHGFI_FLAGS = SHGFI_TYPENAME Or SHGFI_SHELLICONSIZE Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE
Const CF_ENABLETEMPLATE = &H10&, CF_ENABLEHOOK = &H8&, CF_APPLY = &H200&, CF_SCREENFONTS = &H1&
Const CF_INITTOLOGFONTSTRUCT = &H40&, CF_EFFECTS = &H100&, CF_PALETTE = 9&

Const PD_ALLPAGES = &H0&, PD_COLLATE = &H10&, PD_HIDEPRINTTOFILE = &H100000, PD_NOPAGENUMS = &H8&, PD_NOSELECTION = &H4&
Const PD_PAGENUMS = &H2&, PD_PRINTSETUP = &H40&, PD_PRINTTOFILE = &H20&, PD_SELECTION = &H1&, PD_CURRENTPAGE = &H400000

Const CCHDEVICENAME = 32&, CCHFORMNAME = 32&    'PrintDialog
Const DM_DUPLEX = &H1000&, DM_ORIENTATION = &H1&

Const GMEM_MOVEABLE = &H2&, GMEM_ZEROINIT = &H40&

Private Type def_PRINTDLG 'PrintDialog
    lStructSize As Long
    hwndOwner As Long
    hDevMode As Long
    hDevNames As Long
    hDC As Long
    Flags As Long
    nFromPage As Integer
    nToPage As Integer
    nMinPage As Integer
    nMaxPage As Integer
    nCopies As Integer
    hInstance As Long
    lCustData As Long
    lpfnPrintHook As Long
    lpfnSetupHook As Long
    lpPrintTemplateName As String
    lpSetupTemplateName As String
    hPrintTemplate As Long
    hSetupTemplate As Long
End Type

Private Type DEVMODE
    dmDeviceName As String * CCHDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCHFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
    dmICMMethod As Long
    dmICMIntent As Long
    dmMediaType As Long
    dmDitherType As Long
    dmReserved1 As Long
    dmReserved2 As Long
    dmPanningWidth As Long
    dmPanningHeight As Long
End Type

Private Type DEVNAMES 'PrintDialog
    wDriverOffset As Integer
    wDeviceOffset As Integer
    wOutputOffset As Integer
    wDefault As Integer
    Extra As String * 100
End Type

Private Type ChooseColor
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    RGBResult As Long
    lpCustColors As Long
    Flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As Long
End Type

Private Type ChooseFont
    lStructSize As Long
    hwndOwner As Long          '  caller's window handle
    hDC As Long                '  printer DC/IC or NULL
    lpLogFont As Long          '  ptr. to a LOGFONT struct
    iPointSize As Long         '  10 * size in points of selected font
    Flags As Long              '  enum. type flags
    rgbColors As Long          '  returned text color
    lCustData As Long          '  data passed to hook fn.
    lpfnHook As Long           '  ptr. to hook function
    lpTemplateName As Long     '  custom template name
    hInstance As Long          '  instance handle of.EXE that
    lpszStyle As Long          '  return the style field here
    nFontType As Integer       '  same value reported to the EnumFonts
    MISSING_ALIGNMENT As Integer
    nSizeMin As Long           '  minimum pt size allowed &
    nSizeMax As Long           '  max pt size allowed if
End Type

Private Type BROWSEINFO
     hwndOwner As Long
     pidlRoot As Long
     pszDisplayName As Long
     lpszTitle As Long
     ulFlags As Long
     lpfnCallback As Long
     lParam As Long
     iImage As Long
End Type

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As Long
    lpstrCustomFilter As Long
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As Long
    nMaxFile As Long
    lpstrFileTitle As Long
    nMaxFileTitle As Long
    lpstrInitialDir As Long
    lpstrTitle As Long
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As Long
    pvReserved As Long
    dwReserved As Long
    FlagsEx As Long
End Type

Public hwndOwner As Long, FromPage As Long, ToPage As Long, PrnFlag As Long


Public Function ShowFont(Optional ByVal defFont As StdFont, Optional nColor As Variant = 0, Optional ByVal Flags As Long = 321) As StdFont
    Dim pm As ChooseFont, lg As LOGFONT
    
    lg = StdFontToLogFont(defFont)
    
    With pm
        .lStructSize = Len(pm)
        .hInstance = App.hInstance
        .hwndOwner = hwndOwner
        .lpLogFont = VarPtr(lg)
        .iPointSize = defFont.Size * 10
        .rgbColors = CLng(nColor)
        .Flags = Flags
    End With

    If ChooseFontW(pm) Then
        With defFont
            .Name = lg.lfFaceName
            .Charset = lg.lfCharSet
            .Weight = lg.lfWeight
            .Italic = lg.lfItalic
            .Strikethrough = lg.lfStrikeOut
            .Underline = lg.lfUnderline
            .Size = pm.iPointSize / 10
        End With
        nColor = pm.rgbColors
    End If
    
    Set ShowFont = defFont
End Function

Public Function ShowColor(Optional ByVal defColor As Long, Optional vColors As Variant, Optional ByVal Flags As Long = 2) As Long
    Dim pm As ChooseColor, lColors(15) As Long, a As Long

    If IsArray(vColors) = False Then ReDim vColors(15):    For a = 0 To UBound(vColors):   vColors(a) = RGB((15 - a) * 16, (15 - a) * 16, (15 - a) * 16):    Next
    If UBound(vColors) <> 15 Then ReDim Preserve vColors(15)
    For a = 0 To UBound(lColors):     lColors(a) = CLng(vColors(a)):     Next
    
    With pm
        .lStructSize = LenB(pm)
        .hInstance = App.hInstance
        .hwndOwner = hwndOwner
        .Flags = Flags
        .lpCustColors = VarPtr(lColors(0))
    End With
    
    If ChooseColorW(pm) Then ShowColor = pm.RGBResult Else ShowColor = defColor

    For a = 0 To UBound(lColors):     vColors(a) = lColors(a):     Next
End Function

Public Function ShowFolder(Optional ByVal Title As String = "Select Folder", Optional ByVal Flags As Long = 1) As String
    Dim lRes As Long, sTemp As String, bi As BROWSEINFO
    
    Title = Title + vbNullChar
    
    With bi
        .hwndOwner = hwndOwner
        .lpszTitle = StrPtr(Title)
        .ulFlags = Flags
    End With
    
    lRes = SHBrowseForFolderW(bi)
    
    If lRes Then
        sTemp = String$(MAX_PATH_UNI, 0)
        SHGetPathFromIDListW lRes, StrPtr(sTemp)
        CoTaskMemFree lRes
        ShowFolder = TrimNull(sTemp)
    End If
End Function

Public Function ShowOpen(Optional ByVal Filter As String = "All files (*.*)|*.*", Optional ByVal Title As String, Optional ByVal nameFile As String, Optional ByVal InitDir As String, Optional ByVal defExt As String, Optional ByVal numFilter As Long, Optional ByVal Flags As Long) As String
    Dim OFN As OPENFILENAME
    
    Filter = Replace$(Filter, "|", vbNullChar) + vbNullChar
    nameFile = nameFile + String$(MAX_PATH_UNI, 0)
    
    With OFN
        .lStructSize = Len(OFN)
        .hInstance = App.hInstance
        .hwndOwner = hwndOwner
        .lpstrFile = StrPtr(nameFile)
        .nMaxFile = Len(nameFile)
        .lpstrTitle = StrPtr(Title)
        .nMaxFileTitle = Len(Title)
        .lpstrFilter = StrPtr(Filter)
        .lpstrInitialDir = StrPtr(InitDir)
        .lpstrDefExt = StrPtr(defExt)
        .nFilterIndex = numFilter
        .Flags = Flags And &H7FFFFFFF
        .FlagsEx = IIF(Flags And &H80000000, 1, 0)
    End With
    
    If GetOpenFileNameW(OFN) Then ShowOpen = TrimNull(nameFile)
End Function

Public Function ShowSave(Optional ByVal Filter As String = "All files (*.*)|*.*", Optional ByVal Title As String, Optional ByVal nameFile As String, Optional ByVal InitDir As String, Optional ByVal defExt As String, Optional ByVal numFilter As Long, Optional ByVal Flags As Long) As String
    Dim OFN As OPENFILENAME
    
    Filter = Replace$(Filter, "|", vbNullChar) + vbNullChar
    nameFile = nameFile + String$(MAX_PATH_UNI, 0)
    
    With OFN
        .lStructSize = Len(OFN)
        .hInstance = App.hInstance
        .hwndOwner = hwndOwner
        .lpstrFile = StrPtr(nameFile)
        .nMaxFile = Len(nameFile)
        .lpstrTitle = StrPtr(Title)
        .nMaxFileTitle = Len(Title)
        .lpstrFilter = StrPtr(Filter)
        .lpstrInitialDir = StrPtr(InitDir)
        .lpstrDefExt = StrPtr(defExt)
        .nFilterIndex = numFilter
        .Flags = Flags And &H7FFFFFFF
        .FlagsEx = IIF(Flags And &H80000000, 1, 0)
    End With
    
    If GetSaveFileNameW(OFN) Then ShowSave = TrimNull(nameFile)
End Function

Public Function ShowPrinter(Optional ByVal bNewStyle As Boolean = True, Optional ByVal Flags As Long) As Long
    Dim PD As def_PRINTDLG, dm As DEVMODE, DN As DEVNAMES, lpDevMode As Long, lpDevName As Long
    Dim RetValue As Long, objPrinter As Printer, NewPrinterName As String

    With PD
      .lStructSize = Len(PD)
      .hwndOwner = IIF(bNewStyle, frmScript.hWnd, hwndOwner)
      .hDC = GetDC(.hwndOwner)
      .nMinPage = 1
      .nMaxPage = 1000
      .Flags = Flags
    End With

  ' Set the current orientation and duplex setting
    On Error GoTo ExitSub

    With dm
        .dmDeviceName = Printer.DeviceName
        .dmSize = Len(dm)
        .dmFields = DM_ORIENTATION Or DM_DUPLEX
        .dmOrientation = Printer.Orientation
         On Error Resume Next
        .dmDuplex = Printer.Duplex
         On Error GoTo 0
    End With

  ' Allocate memory for the initialization hDevMode structure
  ' and copy the settings gathered above into this memory
    PD.hDevMode = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(dm))
    lpDevMode = GlobalLock(PD.hDevMode)
    If lpDevMode > 0 Then
       CopyMemory ByVal lpDevMode, dm, Len(dm)
       RetValue = GlobalUnlock(lpDevMode)
    End If

  ' Set the current driver, device, and port Name strings
    With DN
        .wDriverOffset = 8
        .wDeviceOffset = .wDriverOffset + 1 + Len(Printer.DriverName)
        .wOutputOffset = .wDeviceOffset + 1 + Len(Printer.Port)
        .wDefault = 0
    End With

    With Printer
         DN.Extra = .DriverName & vbNullChar & .DeviceName & vbNullChar & .Port & vbNullChar
    End With

  ' Allocate memory for the initial hDevName structure
  ' and copy the settings gathered above into this memory
    PD.hDevNames = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(DN))
    lpDevName = GlobalLock(PD.hDevNames)
    If lpDevName > 0 Then
       CopyMemory ByVal lpDevName, DN, Len(DN)
       RetValue = GlobalUnlock(lpDevName)
    End If

  ' Call the print dialog up and let the user make changes
    RetValue = PrintDlgA(PD)

    ShowPrinter = RetValue
    
    ReleaseDC PD.hwndOwner, PD.hDC

    If RetValue <> 0 Then
      'mhOwner = PD.hDC

      lpDevName = GlobalLock(PD.hDevNames)
      CopyMemory DN, ByVal lpDevName, 45
      RetValue = GlobalUnlock(lpDevName)
      GlobalFree PD.hDevNames

      lpDevMode = GlobalLock(PD.hDevMode)
      CopyMemory dm, ByVal lpDevMode, Len(dm)
      RetValue = GlobalUnlock(PD.hDevMode)
      GlobalFree PD.hDevMode

      NewPrinterName = UCase$(Left$(dm.dmDeviceName, InStr(dm.dmDeviceName, vbNullChar) - 1))

      If Printer.DeviceName <> NewPrinterName Then
         For Each objPrinter In Printers
             If UCase$(objPrinter.DeviceName) = NewPrinterName Then
                Set Printer = objPrinter
              End If
         Next
      End If

      On Error Resume Next

    ' Set printer object properties according to selections made by user
      FromPage = PD.nFromPage
      ToPage = PD.nToPage
      PrnFlag = PD.Flags
      
      With Printer
          .Copies = dm.dmCopies
          .Duplex = dm.dmDuplex
          .Orientation = dm.dmOrientation
          .PaperSize = dm.dmPaperSize
          .PrintQuality = dm.dmPrintQuality
      End With

      On Error GoTo 0
    End If

ExitSub:
End Function
