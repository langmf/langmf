Attribute VB_Name = "modConst"
Option Explicit


Declare Sub GetMem1 Lib "msvbvm60" (ByVal addr As Long, RetVal As Byte)
Declare Sub GetMem2 Lib "msvbvm60" (ByVal addr As Long, RetVal As Integer)
Declare Sub GetMem4 Lib "msvbvm60" (ByVal addr As Long, RetVal As Long)
Declare Sub GetMem8 Lib "msvbvm60" (ByVal addr As Long, RetVal As Currency)
Declare Sub PutMem1 Lib "msvbvm60" (ByVal addr As Long, ByVal NewVal As Byte)
Declare Sub PutMem2 Lib "msvbvm60" (ByVal addr As Long, ByVal NewVal As Integer)
Declare Sub PutMem4 Lib "msvbvm60" (ByVal addr As Long, ByVal NewVal As Long)
Declare Sub PutMem8 Lib "msvbvm60" (ByVal addr As Long, ByVal NewVal As Currency)

Declare Sub GetMem2_Wrd Lib "msvbvm60" Alias "GetMem2" (ByVal addr As Long, RetVal As Long)
Declare Sub PutMem2_Wrd Lib "msvbvm60" Alias "PutMem2" (ByVal addr As Long, ByVal NewVal As Long)
Declare Sub GetMem2_Bln Lib "msvbvm60" Alias "GetMem2" (ByVal addr As Long, RetVal As Boolean)
Declare Sub PutMem2_Bln Lib "msvbvm60" Alias "PutMem2" (ByVal addr As Long, ByVal NewVal As Boolean)
Declare Sub GetMem4_Sng Lib "msvbvm60" Alias "GetMem4" (ByVal addr As Long, RetVal As Single)
Declare Sub PutMem4_Sng Lib "msvbvm60" Alias "PutMem4" (ByVal addr As Long, ByVal NewVal As Single)
Declare Sub GetMem8_Dbl Lib "msvbvm60" Alias "GetMem8" (ByVal addr As Long, RetVal As Double)
Declare Sub PutMem8_Dbl Lib "msvbvm60" Alias "PutMem8" (ByVal addr As Long, ByVal NewVal As Double)

Declare Function VarPtrArray Lib "msvbvm60" Alias "VarPtr" (ByRef Ptr() As Any) As Long


Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Declare Function SetMenu Lib "user32" (ByVal hWnd As Long, ByVal hMenu As Long) As Long
Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Declare Function SetFocus Lib "user32" (ByVal hWnd As Long) As Long
Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Declare Function IsIconic Lib "user32" (ByVal hWnd As Long) As Long
Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Declare Function DrawTextW Lib "user32" (ByVal hDC As Long, ByVal lpStr As Long, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function CreateMenu Lib "user32" () As Long
Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal uFlags As Long) As Long
Declare Function AppendMenuW Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Long) As Long
Declare Function ModifyMenuW Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As Long) As Long
Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
Declare Function SetMenuInfo Lib "user32" (ByVal hMenu As Long, mi As MENUINFO) As Long
Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Declare Function MessageBoxW Lib "user32" (ByVal hWnd As Long, ByVal lpText As Long, ByVal lpCaption As Long, ByVal wType As Long) As Long
Declare Function GetMenuState Lib "user32" (ByVal hMenu As Long, ByVal wID As Long, ByVal wFlags As Long) As Long
Declare Function PeekMessageA Lib "user32" (lpMsg As WNDMsg, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
Declare Function PostMessageA Lib "user32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function SendMessageW Lib "user32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Declare Function GetClassNameW Lib "user32" (ByVal hWnd As Long, ByVal ClassName As Long, ByVal classlength As Long) As Long
Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Declare Function ExitWindowsEx Lib "user32" (ByVal dwOptions As Long, ByVal dwReserved As Long) As Long
Declare Function GetWindowLongW Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Declare Function SetWindowLongW Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function GetMenuStringW Lib "user32" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As Long, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Declare Function SetWindowTextW Lib "user32" (ByVal hWnd As Long, ByVal lpString As Long) As Long
Declare Function GetWindowTextW Lib "user32" (ByVal hWnd As Long, ByVal lpString As Long, ByVal cch As Long) As Long
Declare Function OemToCharBuffA Lib "user32" (ByVal lpszSrc As String, ByVal lpszDst As String, ByVal cchDstLength As Long) As Long
Declare Function DefWindowProcW Lib "user32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function EnableMenuItem Lib "user32" (ByVal hMenu As Long, ByVal wIDEnableItem As Long, ByVal wEnable As Long) As Long
Declare Function ReleaseCapture Lib "user32" () As Long
Declare Function CreatePopupMenu Lib "user32" () As Long
Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Declare Function CreateWindowExW Lib "user32" (ByVal dwExStyle As Long, ByVal lpClassName As Long, ByVal lpWindowName As Long, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByVal lpParam As Long) As Long
Declare Function CallWindowProcA Lib "user32" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function DispatchMessageA Lib "user32" (lpMsg As WNDMsg) As Long
Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Declare Function TranslateMessage Lib "user32" (lpMsg As WNDMsg) As Long
Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Declare Function GetDesktopWindow Lib "user32" () As Long
Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Declare Function TrackPopupMenuEx Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal x As Long, ByVal y As Long, ByVal hWnd As Long, ByVal lptpm As Any) As Long
Declare Function LoadImageAsString Lib "user32" Alias "LoadImageW" (ByVal hInst As Long, ByVal lpsz As Long, ByVal uType As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal fuLoad As Long) As Long
Declare Function CreateIconIndirect Lib "user32" (piconinfo As ICONINFO) As Long
Declare Function SetMenuDefaultItem Lib "user32" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPos As Long) As Long
Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long
Declare Function GetForegroundWindow Lib "user32" () As Long
Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Declare Function UpdateLayeredWindow Lib "user32" (ByVal hWnd As Long, ByVal hdcDst As Long, pptDst As Any, psize As Any, ByVal hdcSrc As Long, pptSrc As Any, ByVal crKey As Long, ByRef pblend As Any, ByVal dwFlags As Long) As Long
Declare Function GetWindowTextLengthW Lib "user32" (ByVal hWnd As Long) As Long
Declare Function CreateIconFromResourceEx Lib "user32" (presbits As Any, ByVal dwResSize As Long, ByVal fIcon As Long, ByVal dwVer As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal Flags As Long) As Long
Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Declare Function ChangeWindowMessageFilter Lib "user32" (ByVal Message As Long, ByVal dwFlag As Integer) As Boolean


Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (Destination As Any, ByVal numBytes As Long)
Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (Destination As Any, ByVal Length As Long, ByVal Fill As Byte)
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Declare Sub GetLocalTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
Declare Sub GlobalMemoryStatusEx Lib "kernel32" (lpBuffer As MEMORYSTATUSEX)
Declare Function GetACP Lib "kernel32" () As Long
Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Declare Function lstrcpyA Lib "kernel32" (ByVal RetVal As String, ByVal Ptr As Long) As Long
Declare Function lstrcpyW Lib "kernel32" (ByVal RetVal As Long, ByVal Ptr As Long) As Long
Declare Function lstrlenA Lib "kernel32" (lpString As Any) As Long
Declare Function lstrlenW Lib "kernel32" (lpString As Any) As Long
Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
Declare Function CopyFileW Lib "kernel32" (ByVal lpExistingFileName As Long, ByVal lpNewFileName As Long, ByVal bFailIfExists As Long) As Long
Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function CreateFile Lib "kernel32" Alias "CreateFileW" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Declare Function CreatePipe Lib "kernel32" (phReadPipe As Long, phWritePipe As Long, lpPipeAttributes As Any, ByVal nSize As Long) As Long
Declare Function ReadFileStr Lib "kernel32" Alias "ReadFile" (ByVal hFile As Long, ByVal lpBuffer As String, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
Declare Function GetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryW" (ByVal lpLibFileName As Long) As Long
Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessID As Long) As Long
Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Declare Function DeleteFileW Lib "kernel32" (ByVal lpFileName As Long) As Long
Declare Function GetTempPathW Lib "kernel32" (ByVal nBufferLength As Long, ByVal lpBuffer As Long) As Long
Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As Long
Declare Function SetStdHandle Lib "kernel32" (ByVal nStdHandle As Long, ByVal nHandle As Long) As Long
Declare Function SetEndOfFile Lib "kernel32" (ByVal hFile As Long) As Long
Declare Function GetVersionExA Lib "kernel32" (lpVersionInformation As Any) As Long
Declare Function GetDriveTypeA Lib "kernel32" (ByVal nDrive As String) As Long
Declare Function MapViewOfFile Lib "kernel32" (ByVal hFileMappingObject As Long, ByVal dwDesiredAccess As Long, ByVal dwFileOffsetHigh As Long, ByVal dwFileOffsetLow As Long, ByVal dwNumberOfBytesToMap As Long) As Long
Declare Function IsWow64Process Lib "kernel32" (ByVal hProc As Long, bWow64Process As Boolean) As Long
Declare Function GetSystemTimes Lib "kernel32" (lpIdleTime As FILETIME, lpKernelTime As FILETIME, lpUserTime As FILETIME) As Long
Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Declare Function VirtualProtect Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long
Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Declare Function CreateProcessW Lib "kernel32" (ByVal lpApplicationName As Long, ByVal lpCommandLine As Long, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Declare Function QueryDosDeviceA Lib "kernel32" (ByVal lpDeviceName As String, ByVal lpTargetPath As String, ByVal ucchMax As Long) As Long
Declare Function DuplicateHandle Lib "kernel32" (ByVal hSourceProcessHandle As Long, ByVal hSourceHandle As Long, ByVal hTargetProcessHandle As Long, lpTargetHandle As Long, ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwOptions As Long) As Long
Declare Function GetCommandLineW Lib "kernel32" () As Long
Declare Function UnmapViewOfFile Lib "kernel32" (ByVal lpBaseAddress As Long) As Long
Declare Function SetDllDirectoryW Lib "kernel32" (ByVal lpPathName As Long) As Long
Declare Function GetModuleHandleW Lib "kernel32" (ByVal lpModuleName As Long) As Long
Declare Function FlushFileBuffers Lib "kernel32" (ByVal hFile As Long) As Long
Declare Function RemoveDirectoryW Lib "kernel32" (ByVal lpPathName As Long) As Long
Declare Function CreateDirectoryW Lib "kernel32" (ByVal lpPathName As Long, ByVal lpSecurityAttributes As Long) As Long
Declare Function SetPriorityClass Lib "kernel32" (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Long
Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Declare Function GetLongPathNameW Lib "kernel32" (ByVal lpszShortPath As Long, ByVal lpszLongPath As Long, ByVal cchBuffer As Long) As Long
Declare Function GetComputerNameW Lib "kernel32" (ByVal lpBuffer As Long, nSize As Long) As Long
Declare Function GetShortPathNameW Lib "kernel32" (ByVal lpszLongPath As Long, ByVal lpszShortPath As Long, ByVal cchBuffer As Long) As Long
Declare Function GetCurrentProcess Lib "kernel32" () As Long
Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Declare Function GetModuleFileNameW Lib "kernel32" (ByVal hModule As Long, ByVal lpFileName As Long, ByVal nSize As Long) As Long
Declare Function SetFileAttributesW Lib "kernel32" (ByVal lpFileName As Long, ByVal dwFileAttributes As Long) As Long
Declare Function GetFileAttributesW Lib "kernel32" (ByVal lpFileName As Long) As Long
Declare Function CreateFileMappingW Lib "kernel32" (ByVal hFile As Long, ByVal lpFileMappingAttributes As Long, ByVal flProtect As Long, ByVal dwMaximumSizeHigh As Long, ByVal dwMaximumSizeLow As Long, ByVal lpName As Long) As Long
Declare Function GetDiskFreeSpaceExA Lib "kernel32" (ByVal lpRootPathName As String, lpFreeBytesAvailableToCaller As LONG64, lpTotalNumberOfBytes As LONG64, lpTotalNumberOfFreeBytes As LONG64) As Long
Declare Function GetSystemDirectoryW Lib "kernel32" (ByVal lpBuffer As Long, ByVal nSize As Long) As Long
Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Declare Function WaitForSingleObject Lib "kernel32" (ByVal hProcess As Long, ByVal dwMilliseconds As Long) As Long
Declare Function MultiByteToWideChar Lib "kernel32" (ByVal codepage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Declare Function WideCharToMultiByte Lib "kernel32" (ByVal codepage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, lpUsedDefaultChar As Long) As Long
Declare Function GetCurrentDirectoryW Lib "kernel32" (ByVal nBufferLength As Long, ByVal lpBuffer As Long) As Long
Declare Function SetCurrentDirectoryW Lib "kernel32" (ByVal lpPathName As Long) As Long
Declare Function GetWindowsDirectoryW Lib "kernel32" (ByVal lpBuffer As Long, ByVal nSize As Long) As Long
Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Declare Function GetVolumeInformationA Lib "kernel32" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Declare Function GetSystemDefaultLangID Lib "kernel32" () As Integer
Declare Function SetProcessAffinityMask Lib "kernel32" (ByVal hProcess As Long, ByVal lMask As Long) As Long
Declare Function GetLogicalDriveStringsA Lib "kernel32" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Declare Function LocalFileTimeToFileTime Lib "kernel32" (lpLocalFileTime As FILETIME, lpFileTime As FILETIME) As Long
Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Declare Function GetProcAddressByOrdinal Lib "kernel32" Alias "GetProcAddress" (ByVal hModule As Long, ByVal lpProcName As Long) As Long
Declare Function SetEnvironmentVariableW Lib "kernel32" (ByVal lpName As Long, ByVal lpValue As Long) As Long
Declare Function GetEnvironmentVariableW Lib "kernel32" (ByVal lpName As Long, ByVal lpBuffer As Long, ByVal nSize As Long) As Long


Declare Function GetUserNameW Lib "advapi32" (ByVal lpBuffer As Long, nSize As Long) As Long
Declare Function CryptHashData Lib "advapi32" (ByVal hHash As Long, pbData As Any, ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long
Declare Function CryptCreateHash Lib "advapi32" (ByVal hProv As Long, ByVal Algid As Long, ByVal hKey As Long, ByVal dwFlags As Long, ByRef phHash As Long) As Long
Declare Function CryptDestroyHash Lib "advapi32" (ByVal hHash As Long) As Long
Declare Function OpenProcessToken Lib "advapi32" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Declare Function LookupAccountSidW Lib "advapi32" (ByVal lpSystemName As Long, ByVal Sid As Long, ByVal Name As Long, cbName As Long, ByVal ReferencedDomainName As Long, cbReferencedDomainName As Long, peUse As Integer) As Long
Declare Function CryptGetHashParam Lib "advapi32" (ByVal hHash As Long, ByVal dwParam As Long, pbData As Any, pdwDataLen As Long, ByVal dwFlags As Long) As Long
Declare Function CryptAcquireContext Lib "advapi32" Alias "CryptAcquireContextA" (ByRef phProv As Long, ByVal pszContainer As String, ByVal pszProvider As String, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
Declare Function CryptReleaseContext Lib "advapi32" (ByVal hProv As Long, ByVal dwFlags As Long) As Long
Declare Function GetTokenInformation Lib "advapi32" (ByVal TokenHandle As Long, ByVal TokenInformationClass As Long, TokenInformation As Any, ByVal TokenInformationLength As Long, ReturnLength As Long) As Long
Declare Function LookupPrivilegeValue Lib "advapi32" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long
Declare Function AdjustTokenPrivileges Lib "advapi32" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long


Declare Function PatBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Declare Function SetDIBits Lib "gdi32" (ByVal hDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Declare Function GetDIBits Lib "gdi32" (ByVal hDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Declare Function GetBkColor Lib "gdi32" (ByVal hDC As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Declare Function GetTextColor Lib "gdi32" (ByVal hDC As Long) As Long
Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, ByVal lplpVoid As Long, ByVal Handle As Long, ByVal dw As Long) As Long
Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Declare Function CreateFontIndirectW Lib "gdi32" (lpLogFont As LOGFONT) As Long
Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long


Declare Function timeGetTime Lib "winmm" () As Long
Declare Function PlaySoundW Lib "winmm" (ByVal lpszName As Long, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Declare Function mciSendStringW Lib "winmm" (ByVal lpstrCommand As Long, ByVal lpstrReturnString As Long, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

Declare Sub WTSFreeMemory Lib "wtsapi32" (ByVal pMemory As Long)
Declare Function WTSEnumerateProcesses Lib "wtsapi32" Alias "WTSEnumerateProcessesA" (ByVal hServer As Long, ByVal Reserved As Long, ByVal Version As Long, ByRef ppProcessInfo As Long, ByRef pCount As Long) As Long

Declare Function IsUserAnAdmin Lib "shell32" Alias "#680" () As Integer
Declare Function ShellExecuteW Lib "shell32" (ByVal hWnd As Long, ByVal lpOperation As Long, ByVal lpFile As Long, ByVal lpParameters As Long, ByVal lpDirectory As Long, ByVal nShowCmd As Long) As Long
Declare Function ShellExecuteExW Lib "shell32" (ByVal SEI As Long) As Long
Declare Function SHGetPathFromIDListW Lib "shell32" (ByVal pidl As Long, ByVal pszPath As Long) As Long
Declare Function SHGetSpecialFolderLocation Lib "shell32" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long

Declare Function VerQueryValueW Lib "Version" (pBlock As Any, ByVal lpSubBlock As Long, lplpBuffer As Any, puLen As Long) As Long
Declare Function GetFileVersionInfoW Lib "Version" (ByVal lptstrFilename As Long, ByVal dwhandle As Long, ByVal dwlen As Long, lpData As Any) As Long
Declare Function GetFileVersionInfoSizeW Lib "Version" (ByVal lptstrFilename As Long, lpdwHandle As Long) As Long

Declare Function RtlGetCompressionWorkSpaceSize Lib "ntdll" (ByVal CompressionFormat As Integer, CompressBufferWorkSpaceSize As Long, CompressFragmentWorkSpaceSize As Long) As Long
Declare Function NtAllocateVirtualMemory Lib "ntdll" (ByVal ProcHandle As Long, BaseAddress As Long, ByVal NumBits As Long, regionsize As Long, ByVal Flags As Long, ByVal ProtectMode As Long) As Long
Declare Function NtFreeVirtualMemory Lib "ntdll" (ByVal ProcHandle As Long, BaseAddress As Long, regionsize As Long, ByVal Flags As Long) As Long
Declare Function RtlDecompressBuffer Lib "ntdll" (ByVal CompressionFormat As Integer, UncompressedBuffer As Any, ByVal UncompressedBufferSize As Long, CompressedBuffer As Any, ByVal CompressedBufferSize As Long, FinalUncompressedSize As Long) As Long
Declare Function RtlCompressBuffer Lib "ntdll" (ByVal CompressionFormat As Integer, UncompressedBuffer As Any, ByVal UncompressedBufferSize As Long, CompressedBuffer As Any, ByVal CompressedBufferSize As Long, ByVal UncompressedChunkSize As Long, FinalCompressedSize As Long, ByVal WorkSpace As Long) As Long

Declare Function EnumProcesses Lib "psapi" (lpidProcess As Long, ByVal cb As Long, cbNeeded As Long) As Long
Declare Function GetModuleFileNameExW Lib "psapi" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As Long, ByVal nSize As Long) As Long
Declare Function GetProcessMemoryInfo Lib "psapi" (ByVal hProcess As Long, ppsmemCounters As PROCESS_MEMORY_COUNTERS, ByVal cb As Long) As Long
Declare Function GetProcessImageFileNameW Lib "psapi" (ByVal hProcess As Long, ByVal lpImageFileName As Long, ByVal nSize As Long) As Long

Declare Sub CoTaskMemFree Lib "ole32" (ByVal hMem As Long)
Declare Function CoTaskMemAlloc Lib "ole32" (ByVal cb As Long) As Long
Declare Function CLSIDFromString Lib "ole32" (ByVal lpsz As Any, pClsid As UUID) As Long
Declare Function CLSIDFromProgID Lib "ole32" (ByVal lpsz As Any, pClsid As UUID) As Long
Declare Function StringFromGUID2 Lib "ole32" (ByVal ptrIID As Long, ByVal strIID As Long, Optional ByVal cbMax As Long = 39) As Long
Declare Function CoCreateInstance Lib "ole32" (rclsid As UUID, ByVal pUnkOuter As Long, ByVal dwClsContext As Long, riid As UUID, ppv As Long) As Long
Declare Function CoDisconnectObject Lib "ole32" (ByVal pUnk As IUnknown, pvReserved As Long) As Long
Declare Function CreateStreamOnHGlobal Lib "ole32" (hGlobal As Any, ByVal fDeleteOnRelease As Long, ppstm As Any) As Long

Declare Function OleLoadPicture Lib "olepro32" (pStream As Any, ByVal lSize As Long, ByVal fRunmode As Long, riid As UUID, IPic As IPictureDisp) As Long
Declare Function OleCreatePictureIndirect Lib "olepro32" (PicDesc As PictDesc, RefIID As UUID, ByVal fPictureOwnsHandle As Long, IPic As IPictureDisp) As Long

Declare Function LoadTypeLib Lib "oleaut32" (ByVal szFile As Long, pptlib As ITypeLib) As Long
Declare Function VariantCopy Lib "oleaut32" (varDest As Variant, varSrc As Variant) As Long
Declare Function VariantCopyInd Lib "oleaut32" (ByVal pvargDest As Long, ByVal pvargSrc As Long) As Long
Declare Function SysAllocStringLen Lib "oleaut32" (ByVal OleStr As Long, ByVal bLen As Long) As Long
Declare Function RevokeActiveObject Lib "oleaut32" (ByVal dwRegister As Long, ByVal pvReserved As Long) As Long
Declare Function RegisterActiveObject Lib "oleaut32" (ByVal pUnk As IUnknown, rclsid As UUID, ByVal dwFlags As Long, pdwRegister As Long) As Long

Declare Function wglGetProcAddress Lib "opengl32" (ByVal prcname As String) As Long
Declare Function SetSuspendState Lib "powrprof" (ByVal hibernate As Long, ByVal ForceCritical As Long, ByVal DisableWakeEvent As Long) As Long

Declare Function zlib_Compress Lib "zlib" Alias "compress" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long
Declare Function zlib_UnCompress Lib "zlib" Alias "uncompress" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long

Declare Sub Out32 Lib "inpout32" (ByVal PortAddress As Integer, ByVal value As Integer)
Declare Function Inp32 Lib "inpout32" (ByVal PortAddress As Integer) As Integer
Declare Function IsInpOutDriverOpen Lib "inpout32" () As Long


'--------------------------------------------
Global Const PAGE_NOACCESS          As Long = 1
Global Const PAGE_READONLY          As Long = 2
Global Const PAGE_READWRITE         As Long = 4
Global Const PAGE_WRITECOPY         As Long = 8
Global Const PAGE_EXECUTE           As Long = &H10
Global Const PAGE_EXECUTE_READ      As Long = &H20
Global Const PAGE_EXECUTE_READWRITE As Long = &H40
Global Const PAGE_EXECUTE_WRITECOPY As Long = &H80
Global Const PAGE_GUARD             As Long = &H100&
Global Const PAGE_NOCACHE           As Long = &H200&
Global Const MEM_COMMIT             As Long = &H1000&
Global Const MEM_RESERVE            As Long = &H2000&
Global Const MEM_DECOMMIT           As Long = &H4000&
Global Const MEM_RELEASE            As Long = &H8000&
Global Const MEM_FREE               As Long = &H10000
Global Const MEM_PRIVATE            As Long = &H20000
Global Const MEM_MAPPED             As Long = &H40000
Global Const MEM_RESET              As Long = &H80000
Global Const MEM_TOP_DOWN           As Long = &H100000

'--------------------------------------------
Global Const WM_ACTIVATE            As Long = 6
Global Const WM_PAINT               As Long = &HF
Global Const WM_USER                As Long = &H400&
Global Const WM_HOTKEY              As Long = &H312&
Global Const WM_COMMAND             As Long = &H111&
Global Const WM_SYSCOMMAND          As Long = &H112&
Global Const WM_GETMINMAXINFO       As Long = &H24
Global Const WM_QUERYENDSESSION     As Long = &H11
Global Const WM_ERASEBKGND          As Long = &H14
Global Const WM_SETICON             As Long = &H80
Global Const WM_MOUSEMOVE           As Long = &H200&
Global Const WM_LBUTTONDOWN         As Long = &H201&
Global Const WM_LBUTTONUP           As Long = &H202&
Global Const WM_LBUTTONDBLCLK       As Long = &H203&
Global Const WM_RBUTTONDOWN         As Long = &H204&
Global Const WM_RBUTTONUP           As Long = &H205&
Global Const WM_RBUTTONDBLCLK       As Long = &H206&
Global Const WM_MBUTTONDOWN         As Long = &H207&
Global Const WM_MBUTTONUP           As Long = &H208&
Global Const WM_MBUTTONDBLCLK       As Long = &H209&
Global Const WM_SETFONT             As Long = &H30
Global Const WM_GETTEXT             As Long = &HD
Global Const WM_GETTEXTLENGTH       As Long = &HE
Global Const WM_SETTEXT             As Long = &HC
Global Const WM_CHAR                As Long = &H102&

'--------------------------------------------
Global Const GWL_WNDPROC            As Long = (-4)
Global Const GWL_STYLE              As Long = (-16)
Global Const GWL_EXSTYLE            As Long = (-20)
Global Const GW_CHILD               As Long = 5
Global Const GW_HWNDNEXT            As Long = 2
Global Const WS_CHILD               As Long = &H40000000
Global Const WS_CAPTION             As Long = &HC00000
Global Const WS_EX_LAYERED          As Long = &H80000
Global Const WS_EX_APPWINDOW        As Long = &H40000
Global Const WS_EX_TOOLWINDOW       As Long = &H80
Global Const WS_SYSMENU             As Long = &H80000
Global Const WS_MINIMIZEBOX         As Long = &H20000
Global Const WS_MAXIMIZEBOX         As Long = &H10000
Global Const WS_THICKFRAME          As Long = &H40000
Global Const WS_VISIBLE             As Long = &H10000000
Global Const WS_BORDER              As Long = &H800000
Global Const SWP_NOACTIVATE         As Long = &H10
Global Const SWP_SHOWWINDOW         As Long = &H40
Global Const SWP_FRAMECHANGED       As Long = &H20
Global Const SWP_NOMOVE             As Long = 2
Global Const SWP_NOSIZE             As Long = 1
Global Const SWP_NOZORDER           As Long = 4
Global Const SW_HIDE                As Long = 0
Global Const SW_SHOWNORMAL          As Long = 1
Global Const SW_SHOW                As Long = 5
Global Const hWnd_TOPMOST           As Long = -1
Global Const hWnd_NOTOPMOST         As Long = -2
Global Const PM_NOREMOVE            As Long = 0
Global Const PM_REMOVE              As Long = 1
Global Const MSGFLT_ADD             As Long = 1
Global Const MSGFLT_REMOVE          As Long = 2

'--------------------------------------------
Global Const IMAGE_ICON             As Long = 1
Global Const ICON_SMALL             As Long = 0
Global Const ICON_BIG               As Long = 1
Global Const SM_CXICON              As Long = 11
Global Const SM_CYICON              As Long = 12
Global Const SM_CXSMICON            As Long = 49
Global Const SM_CYSMICON            As Long = 50
Global Const LR_SHARED              As Long = &H8000&
Global Const LR_LOADFROMFILE        As Long = &H10
Global Const WTS_CURRENT_SERVER     As Long = 0
Global Const TPM_RETURNCMD          As Long = &H100&
Global Const TPM_RIGHTBUTTON        As Long = &H2
Global Const STD_INPUT_HANDLE       As Long = -10
Global Const STD_OUTPUT_HANDLE      As Long = -11
Global Const STD_ERROR_HANDLE       As Long = -12
Global Const COLOR_WINDOWTEXT       As Long = 8
Global Const THBN_CLICKED           As Long = &H1800&
Global Const MD_TRANSPARENT         As Long = 1
Global Const MD_OPAQUE              As Long = 2
Global Const RGN_OR                 As Long = 2
Global Const FW_NORMAL              As Long = 400
Global Const FW_BOLD                As Long = 700
Global Const DT_WORDBREAK           As Long = &H10
Global Const DT_SINGLELINE          As Long = &H20
Global Const DT_VCENTER             As Long = 4
Global Const DT_BOTTOM              As Long = 8
Global Const DT_CENTER              As Long = 1
Global Const DT_RIGHT               As Long = 2
Global Const DT_LEFT                As Long = 0
Global Const DT_TOP                 As Long = 0
Global Const MAX_PATH               As Long = 260
Global Const MAX_PATH_X2            As Long = 520
Global Const MAX_PATH_UNI           As Long = 4096
Global Const API_StdCall            As Long = 0
Global Const API_CDecl              As Long = 1
Global Const VbFunc                 As Long = VbMethod Or VbGet

'--------------------------------------------
Global Const DC_HORZRES             As Long = 8
Global Const DC_VERTRES             As Long = 10
Global Const DC_BITSPIXEL           As Long = 12
Global Const DC_PLANES              As Long = 14
Global Const DC_LOGPIXELSX          As Long = 88
Global Const DC_LOGPIXELSY          As Long = 90
Global Const DC_VREFRESH            As Long = 116
Global Const DC_DESKTOPVERTRES      As Long = 117
Global Const DC_DESKTOPHORZRES      As Long = 118

'--------------------------------------------
Global Const DD_ATTACH_TO_DESKTOP   As Long = 1
Global Const DD_MULTI_DRIVER        As Long = 2
Global Const DD_PRIMARY_DEVICE      As Long = 4
Global Const DD_MIRRORING_DRIVER    As Long = 8
Global Const DD_VGA_COMPATIBLE      As Long = &H10
Global Const DD_REMOVABLE           As Long = &H20
Global Const DD_DISCONNECT          As Long = &H2000000
Global Const DD_REMOTE              As Long = &H4000000
Global Const DD_MODESPRUNED         As Long = &H8000000

'--------------------------------------------
Global Const MF_BYCOMMAND           As Long = 0
Global Const MF_BYPOSITION          As Long = &H400&
Global Const MF_STRING              As Long = 0
Global Const MF_GRAYED              As Long = 1
Global Const MF_DISABLED            As Long = 2
Global Const MF_BITMAP              As Long = 4
Global Const MF_CHECKED             As Long = 8
Global Const MF_POPUP               As Long = &H10
Global Const MF_MENUBARBREAK        As Long = &H20
Global Const MF_MENUBREAK           As Long = &H40
Global Const MF_HILITE              As Long = &H80
Global Const MF_OWNERDRAW           As Long = &H100&
Global Const MF_SEPARATOR           As Long = &H800&
Global Const MF_DEFAULT             As Long = &H1000&
Global Const MF_SYSMENU             As Long = &H2000&
Global Const MF_HELP                As Long = &H4000&
Global Const MF_RIGHTJUSTIFY        As Long = &H4000&
Global Const MF_MOUSESELECT         As Long = &H8000&

'--------------------------------------------
Global Const MIM_MAXHEIGHT          As Long = 1
Global Const MIM_BACKGROUND         As Long = 2
Global Const MIM_APPLYTOSUBMENUS    As Long = &H80000000

'--------------------------------------------
Global Const CRYPT_VERIFYCONTEXT    As Long = &HF0000000
Global Const CRYPT_NEWKEYSET        As Long = 8
Global Const PROV_RSA_FULL          As Long = 1
Global Const PROV_RSA_AES           As Long = 24
Global Const HP_HASHVAL             As Long = 2
Global Const HP_HASHSIZE            As Long = 4
Global Const ALG_CLASS_HASH         As Long = &H8000&
Global Const SHA1                   As Long = ALG_CLASS_HASH Or 4
Global Const SHA256                 As Long = ALG_CLASS_HASH Or 12
Global Const SHA512                 As Long = ALG_CLASS_HASH Or 14

'--------------------------------------------
Global Const LOCALE_USER_DEFAULT        As Long = &H400
Global Const SE_PRIVILEGE_ENABLED       As Long = 2
Global Const TOKEN_ELEVATION_TYPE       As Long = 18
Global Const TOKEN_ADJUST_PRIVILEGES    As Long = &H20
Global Const TOKEN_QUERY                As Long = 8
Global Const TOKEN_READ                 As Long = &H20008
Global Const EWX_LOGOFF                 As Long = 0
Global Const EWX_SHUTDOWN               As Long = 1
Global Const EWX_REBOOT                 As Long = 2
Global Const EWX_FORCE                  As Long = 4
Global Const EWX_POWEROFF               As Long = 8
Global Const SC_MONITORPOWER            As Long = &HF170&

'--------------------------------------------
Global Const CREATE_ALWAYS              As Long = 2
Global Const OPEN_EXISTING              As Long = 3
Global Const OPEN_ALWAYS                As Long = 4
Global Const GENERIC_WRITE              As Long = &H40000000
Global Const GENERIC_READ               As Long = &H80000000
Global Const INVALID_HANDLE             As Long = -1
Global Const FILE_SHARE_READ            As Long = 1
Global Const FILE_SHARE_WRITE           As Long = 2
Global Const FILE_ATTRIBUTE_NORMAL      As Long = &H80
Global Const FILE_ATTRIBUTE_DIRECTORY   As Long = &H10
Global Const FILE_MAP_READ              As Long = &H4

'--------------------------------------------
Global Const STARTF_USESHOWWINDOW       As Long = 1
Global Const STARTF_USESTDHANDLES       As Long = &H100&
Global Const DUPLICATE_CLOSE_SOURCE     As Long = 1
Global Const DUPLICATE_SAME_ACCESS      As Long = 2
Global Const ACTIVEOBJECT_STRONG        As Long = 0
Global Const ACTIVEOBJECT_WEAK          As Long = 1

'--------------------------------------------
Global Const SEE_MASK_NOCLOSEPROCESS    As Long = &H40
Global Const SEE_MASK_NOASYNC           As Long = &H100&

'--------------------------------------------
Global Const NORMAL_PRIORITY_CLASS          As Long = &H20
Global Const IDLE_PRIORITY_CLASS            As Long = &H40
Global Const HIGH_PRIORITY_CLASS            As Long = &H80
Global Const REALTIME_PRIORITY_CLASS        As Long = &H100&
Global Const BELOW_NORMAL_PRIORITY_CLASS    As Long = &H4000&
Global Const ABOVE_NORMAL_PRIORITY_CLASS    As Long = &H8000&

'--------------------------------------------
Global Const PROCESS_QUERY_LIMITED_INFORMATION  As Long = &H1000&
Global Const PROCESS_QUERY_INFORMATION          As Long = &H400&
Global Const PROCESS_SET_INFORMATION            As Long = &H200&
Global Const PROCESS_VM_READ                    As Long = &H10
Global Const PROCESS_QUERY_SET                  As Long = PROCESS_QUERY_INFORMATION Or PROCESS_SET_INFORMATION

'---------------------------------------------
Global Const CMS_STATUS_SUCCESS                 As Long = &H0
Global Const CMS_STATUS_BUFFER_ALL_ZEROS        As Long = &H117&
Global Const CMS_STATUS_INVALID_PARAMETER       As Long = &HC000000D
Global Const CMS_STATUS_BUFFER_TOO_SMALL        As Long = &HC0000023
Global Const CMS_STATUS_NOT_SUPPORTED           As Long = &HC00000BB
Global Const CMS_STATUS_BAD_COMPRESS_BUFFER     As Long = &HC0000242
Global Const CMS_STATUS_UNSUPPORTED_COMPRESS    As Long = &HC000025F
Global Const CMS_FORMAT_NONE                    As Long = 0
Global Const CMS_FORMAT_ZLIB                    As Long = 1             'custom mod
Global Const CMS_FORMAT_LZNT1                   As Long = 2
Global Const CMS_FORMAT_XPRESS                  As Long = 3             'added in Windows 8
Global Const CMS_FORMAT_XPRESS_HUFF             As Long = 4             'added in Windows 8
Global Const CMS_ENGINE_STANDARD                As Long = 0
Global Const CMS_ENGINE_MAXIMUM                 As Long = &H100&


'---------------------------------------------
Type PictDesc
    Size As Long
    Type As Long
    hHandle As Long
    hPal As Long
End Type

Type SafeArrayBound
    cElements As Long
    lLbound As Long
End Type

Type SafeArray
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    rgSABound(0 To 2) As SafeArrayBound
End Type

Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Type WNDMsg
    hWnd As Long
    Message As Long
    wParam As Long
    lParam As Long
    Time As Long
    pt As POINTAPI
End Type

Type PROCESS_MEMORY_COUNTERS
   cb As Long
   PageFaultCount As Long
   PeakWorkingSetSize As Long
   WorkingSetSize As Long
   QuotaPeakPagedPoolUsage As Long
   QuotaPagedPoolUsage As Long
   QuotaPeakNonPagedPoolUsage As Long
   QuotaNonPagedPoolUsage As Long
   PagefileUsage As Long
   PeakPagefileUsage As Long
End Type

Type WTS_PROCESS_INFO
    SessionID As Long
    processID As Long
    pProcessName As Long
    pUserSid As Long
End Type

Type SHELLITEMID
    cb As Long
    abID As Byte
End Type

Type ITEMIDLIST
    mkid As SHELLITEMID
End Type

Type LUID
    UsedPart As Long
    IgnoredForNowHigh32BitPart As Long
End Type

Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    TheLuid As LUID
    Attributes As Long
End Type

Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH_X2
    cAlternate As String * 28
End Type

Type MENUINFO
    cbSize As Long
    fMask As Long
    dwStyle As Long
    cyMax As Long
    hbrBack As Long
    dwContextHelpID As Long
    dwMenuData As Long
End Type

Type ICONINFO
    fIcon      As Long
    xHotspot   As Long
    yHotspot   As Long
    hBmMask    As Long
    hBmColor   As Long
End Type

Type MINMAXINFO
    ptReserved As POINTAPI
    ptMaxSize As POINTAPI
    ptMaxPosition As POINTAPI
    ptMinTrackSize As POINTAPI
    ptMaxTrackSize As POINTAPI
End Type

Type THUMBBUTTON
    dwMask As THUMBBUTTONMASK
    IID As Long
    iBitmap As Long
    hIcon As Long
    szTip As String * MAX_PATH
    dwFlags As THUMBBUTTONFLAGS
End Type

Type BLENDFUNCTION
    BlendOp As Byte
    BlendFlags As Byte
    SourceConstantAlpha As Byte
    AlphaFormat As Byte
End Type

Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbAlpha As Byte
End Type

Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type

Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(63) As Byte
End Type

Type LONG64
    LowPart As Long
    HighPart As Long
End Type

Type MEMORYSTATUSEX
    dwLength As Long
    dwMemoryLoad As Long
    ullTotalPhys As LONG64
    ullAvailPhys As LONG64
    ullTotalPageFile As LONG64
    ullAvailPageFile As LONG64
    ullTotalVirtual As LONG64
    ullAvailVirtual As LONG64
    ullAvailExtendedVirtual As LONG64
End Type

Type VS_FIXEDFILEINFO
    dwSignature As Long
    dwStrucVersionl As Integer
    dwStrucVersionh As Integer
    dwFileVersionMSl As Integer
    dwFileVersionMSh As Integer
    dwFileVersionLSl As Integer
    dwFileVersionLSh As Integer
    dwProductVersionMSl As Integer
    dwProductVersionMSh As Integer
    dwProductVersionLSl As Integer
    dwProductVersionLSh As Integer
    dwFileFlagsMask As Long
    dwFileFlags As Long
    dwFileOS As Long
    dwFileType As Long
    dwFileSubtype As Long
    dwFileDateMS As Long
    dwFileDateLS As Long
End Type

Type OSVERSIONINFOEX
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformID As Long
    szCSDVersion As String * 128
    wServicePackMajor As Integer
    wServicePackMinor As Integer
    wSuiteMask As Integer
    wProductType As Byte
    wReserved As Byte
End Type

Type PROCESS_INFORMATION
    hProcess    As Long
    hThread     As Long
    dwProcessID As Long
    dwThreadID  As Long
End Type

Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Type STARTUPINFO
    cb              As Long
    lpReserved      As Long
    lpDesktop       As Long
    lpTitle         As Long
    dwX             As Long
    dwY             As Long
    dwXSize         As Long
    dwYSize         As Long
    dwXCountChars   As Long
    dwYCountChars   As Long
    dwFillAttribute As Long
    dwFlags         As Long
    wShowWindow     As Integer
    cbReserved2     As Integer
    lpReserved2     As Long
    hStdInput       As Long
    hStdOutput      As Long
    hStdError       As Long
End Type

Type SHELLEXECUTEINFO
    cbSize        As Long
    fMask         As Long
    hWnd          As Long
    lpVerb        As String
    lpFile        As String
    lpParameters  As String
    lpDirectory   As String
    nShow         As Long
    hInstApp      As Long
    lpIDList      As Long
    lpClass       As String
    hkeyClass     As Long
    dwHotKey      As Long
    hIcon         As Long
    hProcess      As Long
End Type

Type DEVMODE
    dmDeviceName As String * 32 'Name of graphics card
    dmSpecVersion As Integer
    dmDriverVersion As Integer  'graphics card driver version
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
    dmFormName As String * 32   'Name of form
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer     'Color Quality (can be 8, 16, 24, 32 or even 4)
    dmPelsWidth As Long         'Display Width in pixels
    dmPelsHeight As Long        'Display height in pixels
    dmDisplayFlags As Long
    dmDisplayFrequency As Long  'Display frequency
    dmICMMethod As Long
    dmICMIntent As Long
    dmMediaType As Long
    dmDitherType As Long
    dmReserved1 As Long
    dmReserved2 As Long
    dmPanningWidth As Long
    dmPanningHeight As Long
End Type

Type DEVDISPLAY
   cbSize As Long
   DeviceName As String * 32
   DeviceString As String * 128
   StateFlags As Long
   DeviceID As String * 128
   DeviceKey As String * 128
End Type

Type TRIVERTEX
    x As Long
    y As Long
    red As Integer
    green As Integer
    blue As Integer
    alpha As Integer
End Type

Type Gradient_Triangle
    Vertex1 As Long
    Vertex2 As Long
    Vertex3 As Long
End Type

Type ICONDIRENTRY
   bWidth As Byte               'Width of the image
   bHeight As Byte              'Height of the image (times 2)
   bColorCount As Byte          'Number of colors in image (0 if >=8bpp)
   bReserved As Byte            'Reserved
   wPlanes As Integer           'Color Planes
   wBitCount As Integer         'Bits per pixel
   dwBytesInRes As Long         'how many bytes in this resource?
   dwImageOffset As Long        'where in the file is this image
End Type

Type ICONDIR
   idReserved As Integer
   idType As Integer
   idCount As Integer
End Type
