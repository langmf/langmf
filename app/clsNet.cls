VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function WSACleanup Lib "ws2_32" () As Long
Private Declare Function WSAStartup Lib "ws2_32" (ByVal wVersionRequired As Long, lpWSADATA As WSADATA) As Long
Private Declare Function inet_addr Lib "ws2_32" (ByVal s As String) As Long
Private Declare Function inet_ntoa Lib "ws2_32" (ByVal addr As Long) As Long
Private Declare Function gethostbyname Lib "ws2_32" (ByVal Hostname As String) As Long
Private Declare Function IcmpSendEcho Lib "icmp" (ByVal IcmpHandle As Long, ByVal DestinationAddress As Long, ByVal RequestData As String, ByVal RequestSize As Long, ByVal RequestOptions As Long, ReplyBuffer As ICMP_ECHO_REPLY, ByVal ReplySize As Long, ByVal Timeout As Long) As Long
Private Declare Function IcmpCreateFile Lib "icmp" () As Long
Private Declare Function IcmpCloseHandle Lib "icmp" (ByVal IcmpHandle As Long) As Long

Const INADDR_NONE As Long = &HFFFFFFFF, WS_VERSION_REQD As Long = &H101, IP_SUCCESS As Long = 0, PING_TIMEOUT As Long = 500

Private Type ICMP_OPTIONS
   Ttl             As Byte
   Tos             As Byte
   Flags           As Byte
   OptionsSize     As Byte
   OptionsData     As Long
End Type

Private Type ICMP_ECHO_REPLY
   Address         As Long
   Status          As Long
   RoundTripTime   As Long
   DataSize        As Long
   DataPointer     As Long
   Options         As ICMP_OPTIONS
   Data            As String * 250
End Type

Private Type WSADATA
   wVersion             As Integer
   wHighVersion         As Integer
   szDescription(256)   As Byte
   szSystemStatus(128)  As Byte
   wMaxSockets          As Long
   wMaxUDPDG            As Long
   dwVendorInfo         As Long
End Type


Public Function SocketInit() As Boolean
    Dim WSAD As WSADATA
    On Error Resume Next
    SocketInit = (WSAStartup(WS_VERSION_REQD, WSAD) = IP_SUCCESS)
End Function

Public Function SocketClean() As Boolean
    On Error Resume Next
    SocketClean = (WSACleanup = 0)
End Function

Public Function Ping(ByVal sAddress As String, Optional vTime As Variant, Optional vTTL As Variant, Optional ByVal sDataToSend As String = "echo request") As Long
    Dim lngPort As Long, lngAddress As Long, Echo As ICMP_ECHO_REPLY
    
    On Error Resume Next
    
    If SocketInit Then
        lngAddress = inet_addr(sAddress)
                        
        If lngAddress <> INADDR_NONE Then
            lngPort = IcmpCreateFile()
            If lngPort Then
                IcmpSendEcho lngPort, lngAddress, sDataToSend, Len(sDataToSend), 0, Echo, Len(Echo), PING_TIMEOUT
                Ping = Echo.Status
                IcmpCloseHandle lngPort
                
                SocketClean
                
                vTime = Echo.RoundTripTime
                vTTL = Echo.Options.Ttl
            End If
        Else
            Ping = INADDR_NONE
        End If
    End If
End Function

Public Function Name2IP(ByVal sHostname As String) As String
    Dim ptrHosent As Long, ptrName As Long, ptrAddr As Long, ptrIPAddr As Long, ptrIPAddr2 As Long
    
    On Error Resume Next
    
    If SocketInit Then
        ptrHosent = gethostbyname(sHostname & vbNullChar)
            
        If ptrHosent <> 0 Then
            ptrName = ptrHosent
            ptrAddr = ptrHosent + 12
            CopyMemory ptrName, ByVal ptrName, 4
            CopyMemory ptrAddr, ByVal ptrAddr, 4
            CopyMemory ptrIPAddr, ByVal ptrAddr, 4
            CopyMemory ptrIPAddr2, ByVal ptrIPAddr, 4
            Name2IP = GetStringPtrA(inet_ntoa(ptrIPAddr2))
            
            SocketClean
        End If
    End If
End Function

Public Sub WOL(ByVal txtIP As String, ByVal txtMAC As String, Optional ByVal txtPasswd As String)
    Dim txtHost() As String, a As Long, b As Long, arrMAC(5) As Byte, magicData() As Byte, Obj As Object

    On Error Resume Next
    
    ReDim magicData(101)
    
    txtHost = Split(txtIP, ":")
    ReDim Preserve txtHost(1)
  
    For a = 1 To 16 Step 3
      arrMAC((a - 1) / 3) = Val("&H" & Mid$(txtMAC, a, 2))
    Next
    
    For a = 0 To 5:     magicData(a) = 255:     Next
    For a = 0 To 15:    For b = 0 To 5:    magicData((a * 6) + b + 6) = arrMAC(b):     Next:    Next
    
    If LenB(txtPasswd) Then
        ReDim Preserve magicData(107)
        For a = 1 To 6
            magicData(a + 101) = Asc(Mid$(txtPasswd, a, 1))
        Next
    End If
    
    Licenses.Add "MSWinsock.Winsock"
    With frmScript.Controls.Add("MSWinsock.Winsock", "Wsk").Object
        .RemoteHost = txtHost(0)
        .RemotePort = Val(txtHost(1))
        .Protocol = 1                       ' sckUDPProtocol
        .SendData magicData
        .Close
    End With
    frmScript.Controls.Remove "Wsk"
End Sub

Public Function XMLHTTP(Optional ByVal Host As String = "https://api.ipify.org", Optional ByVal Param As Variant) As Object
    Dim t1 As Long, t2 As Long, tmp As Variant

    On Error Resume Next

    m_ParamArray Param, 10000, Empty, Empty, Empty, Empty

    Set XMLHTTP = CreateObject("MSXML2.XMLHTTP")
    If LenB(Host) = 0 Then Exit Function
    
    XMLHTTP.Open IIF(VarType(Param(1)) = vbString, "POST", "GET"), Host, CBool(Param(0)), Param(2), Param(3)
    XMLHTTP.SetRequestHeader "If-Modified-Since", "Wed, 01 Jan 1997 00:00:00 GMT"

    If IsArray(Param(4)) Then For Each tmp In Param(4):     Call XMLHTTP.SetRequestHeader(tmp(0), tmp(1)):     Next
    
    XMLHTTP.Send Param(1)

    t1 = timeGetTime
    Do
        m_Wait
        t2 = timeGetTime
    Loop Until XMLHTTP.readyState = 4 Or (t1 + Param(0) < t2) Or (t1 > t2) Or mf_IsEnd

    If XMLHTTP.readyState <> 4 Then XMLHTTP.Abort
End Function

Public Function WINHTTP(Optional ByVal Host As String = "https://api.ipify.org", Optional ByVal Param As Variant) As Object
    Dim tmp As Variant

    On Error Resume Next

    m_ParamArray Param, 10, Empty, Empty, Empty, Empty

    Set WINHTTP = CreateObject("WinHttp.WinHttpRequest.5.1"):   If WINHTTP Is Nothing Then Set WINHTTP = CreateObject("WinHttp.WinHttpRequest")
    If LenB(Host) = 0 Then Exit Function
    
    WINHTTP.Open IIF(VarType(Param(1)) = vbString, "POST", "GET"), Host, CBool(Param(0)), Param(2), Param(3)
    WINHTTP.SetRequestHeader "If-Modified-Since", "Wed, 01 Jan 1997 00:00:00 GMT"

    If IsArray(Param(4)) Then For Each tmp In Param(4):     Call WINHTTP.SetRequestHeader(tmp(0), tmp(1)):     Next
    
    WINHTTP.Send Param(1)
    
    WINHTTP.WaitForResponse Param(0)
End Function
