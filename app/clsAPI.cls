VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit


Public Function LoadLibrary(ByVal lpLibFileName As String) As Long
    LoadLibrary = modConst.LoadLibrary(StrPtr(lpLibFileName))
End Function

Public Function FreeLibrary(ByVal hLibModule As Long) As Long
    FreeLibrary = modConst.FreeLibrary(hLibModule)
End Function

Public Function GetProcAddress(ByVal hModule As Long, ByVal lpProcName As String) As Long
    GetProcAddress = modConst.GetProcAddress(hModule, lpProcName)
End Function

Public Function BitBlt(ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
    BitBlt = modConst.BitBlt(hDestDC, x, y, nWidth, nHeight, hSrcDC, XSrc, YSrc, dwRop)
End Function

Public Function StretchBlt(ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal dwRop As Long) As Long
    StretchBlt = modConst.StretchBlt(hDestDC, x, y, nWidth, nHeight, hSrcDC, XSrc, YSrc, srcWidth, srcHeight, dwRop)
End Function

Public Function GetPixel(ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
    GetPixel = modConst.GetPixel(hDC, x, y)
End Function

Public Function SetPixel(ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal cColor As Long) As Long
    SetPixel = modConst.SetPixel(hDC, x, y, cColor)
End Function

Public Sub CopyMemory(ByVal hpvDest As Long, ByVal hpvSource As Long, ByVal cbCopy As Long)
    modConst.CopyMemory ByVal hpvDest, ByVal hpvSource, cbCopy
End Sub

Public Function SendMessage(ByVal my_hWnd As Long, ByVal my_wMsg As Long, ByVal my_wParam As Long, ByVal my_lParam As Long) As Long
    SendMessage = modConst.SendMessageW(my_hWnd, my_wMsg, my_wParam, my_lParam)
End Function

Public Function SetWindowPos(ByVal my_hWnd As Long, ByVal my_hWndInsertafter As Long, ByVal my_x As Long, ByVal my_y As Long, ByVal my_cx As Long, ByVal my_cy As Long, ByVal my_wFlags As Long) As Long
    SetWindowPos = modConst.SetWindowPos(my_hWnd, my_hWndInsertafter, my_x, my_y, my_cx, my_cy, my_wFlags)
End Function

Public Function CallWindowProc(ByVal lpPrevWndFunc As Long, Optional ByVal hWnd As Long, Optional ByVal uMsg As Long, Optional ByVal wParam As Long, Optional ByVal lParam As Long) As Long
    CallWindowProc = modConst.CallWindowProcA(lpPrevWndFunc, hWnd, uMsg, wParam, lParam)
End Function

Public Function CreateWindowEx(ByVal m_dwExStyle As Long, ByVal m_lpClassName As String, ByVal m_lpWindowName As String, ByVal m_dwStyle As Long, ByVal m_x As Long, ByVal m_y As Long, ByVal m_nWidth As Long, ByVal m_nHeight As Long, ByVal m_hWndParent As Long, ByVal m_hMenu As Long, ByVal m_hInstance As Long, ByVal m_lpParam As Long) As Long
    On Error Resume Next
    CreateWindowEx = modConst.CreateWindowExW(m_dwExStyle, StrPtr(m_lpClassName), StrPtr(m_lpWindowName), m_dwStyle, m_x, m_y, m_nWidth, m_nHeight, m_hWndParent, m_hMenu, m_hInstance, m_lpParam)
End Function
