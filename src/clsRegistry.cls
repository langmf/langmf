VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Private Declare Function RegEnumKeyA Lib "advapi32" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Private Declare Function RegOpenKeyExA Lib "advapi32" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegEnumValueW Lib "advapi32" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As Long, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Private Declare Function RegQueryValueA Lib "advapi32" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpValue As String, lpcbValue As Long) As Long

Private Const KEY_QUERY_VALUE = &H1, KEY_ENUMERATE_SUB_KEYS = &H8, KEY_ACCESS = KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS
Private Const REG_DWORD = 4&, REG_SZ = 1&, REG_MULTI_SZ = 7&, REG_EXPAND_SZ = 2&, ERROR_NO_MORE_ITEMS = 259
Private Const HKCR = &H80000000, HKCC = &H80000005, HKCU = &H80000001, HKLM = &H80000002, HKUS = &H80000003

Private WSH As Object


Private Sub Class_Initialize()
    On Error Resume Next
    Set WSH = VBA.CreateObject("WScript.Shell")
End Sub


'----------------- Работа с реестром (ветвь VB and VBA) --------------------
Public Sub SaveSetting(ByVal regAppName As String, ByVal regSect As String, ByVal regKey As String, ByVal regSetting As String)
    VBA.SaveSetting regAppName, regSect, regKey, regSetting
End Sub

Public Function GetSetting(ByVal regAppName As String, ByVal regSect As String, ByVal regKey As String, Optional ByVal regDef As String) As String
    GetSetting = VBA.GetSetting(regAppName, regSect, regKey, regDef)
End Function

Public Function GetAllSettings(ByVal regAppName As String, Optional ByVal regSect As String = vbNullChar) As Variant
    GetAllSettings = VBA.GetAllSettings(regAppName, regSect)
End Function

Public Sub DeleteSetting(ByVal regAppName As String, ByVal regSect As String, ByVal regKey As String)
    VBA.DeleteSetting regAppName, regSect, regKey
End Sub


'----------------- Работа с реестром (WScript Object) --------------------
Public Sub RegWrite(ByVal strName As String, ByVal anyValue As Variant, Optional ByVal strType As String = "REG_SZ")
    On Error Resume Next
    WSH.RegWrite strName, anyValue, strType
End Sub

Public Function RegRead(ByVal strName As String) As Variant
    On Error Resume Next
    RegRead = WSH.RegRead(strName)
End Function

Public Sub RegDelete(ByVal strName As String)
    On Error Resume Next
    WSH.RegDelete strName
End Sub


'----------------- Работа с реестром (WinApi) --------------------
Public Function RegEnum(ByVal Key As String, Optional ByVal Section As Long = HKCU, Optional ByVal samDesired As Long = KEY_ACCESS) As Collection
    Dim i As Long, r As Long, hKey As Long, sz As Long, vLen As Long, VType As Long
    Dim txt As String, vName As String, Buf(1 To 16384) As Byte, b() As Byte

    Set RegEnum = New Collection:      RegSection Key, Section
    
    If RegOpenKeyExA(Section, Key, 0&, samDesired, hKey) <> 0 Then Exit Function

    '------------------- Enumerate the key values ---------------------
    i = 0
    Do
        vName = Space$(255):    vLen = Len(vName):    sz = UBound(Buf)
        r = RegEnumValueW(hKey, i, StrPtr(vName), vLen, 0, VType, Buf(1), sz)
        
        If r = ERROR_NO_MORE_ITEMS Then Exit Do
        
        If r = 0 Then
            vName = Left$(vName, vLen)
            
            Select Case VType
                Case REG_DWORD
                    CopyMemory r, Buf(1), 4
                    RegEnum.Add Array(VType, vName, r)
                Case REG_SZ, REG_MULTI_SZ, REG_EXPAND_SZ
                    txt = Left$(Buf, sz / 2)
                    If VType = REG_MULTI_SZ Then txt = Replace$(txt, Chr$(0), vbCrLf) Else txt = Replace$(txt, Chr$(0), "")
                    RegEnum.Add Array(VType, vName, txt)
                Case Else
                    If sz Then
                        ReDim b(1 To sz)
                        CopyMemory b(1), Buf(1), sz
                        RegEnum.Add Array(VType, vName, b)
                    Else
                        RegEnum.Add Array(VType, vName, Array())
                    End If
            End Select
        End If

        i = i + 1
    Loop

    '------------------- Enumerate the subkeys ------------------------
    i = 0
    Do
        vName = Space$(255):    vLen = Len(vName)
        If RegEnumKeyA(hKey, i, vName, vLen) <> 0 Then Exit Do
        vName = Left$(vName, InStr(vName, Chr$(0)) - 1)
    
        txt = Space$(255):      vLen = Len(txt)
        If RegQueryValueA(hKey, vName, txt, vLen) = 0 Then txt = Left$(txt, vLen - 1) Else txt = ""
        
        RegEnum.Add Array(-1, vName, txt)
        i = i + 1
    Loop
    
    Call RegCloseKey(hKey)
End Function


Private Sub RegSection(Key As String, Section As Long)
    Select Case Left$(UCase$(Key), 5)
        Case "HKCU\":    Key = Right$(Key, Len(Key) - 5):    Section = HKCU
        Case "HKLM\":    Key = Right$(Key, Len(Key) - 5):    Section = HKLM
        Case "HKCR\":    Key = Right$(Key, Len(Key) - 5):    Section = HKCR
        Case "HKCC\":    Key = Right$(Key, Len(Key) - 5):    Section = HKCC
        Case "HKUS\":    Key = Right$(Key, Len(Key) - 5):    Section = HKUS
    End Select
End Sub

