VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private numLenRandom(512) As Long

Public Search As Object


Public Function API(Optional ByVal Param As Variant) As Object                 'As clsFileAPI
    Set API = m_DoParams(New clsFileAPI, Param)
End Function

Public Function Open2(ByVal nameFile As String, Optional ByVal NumFile As Integer = 0, Optional ByVal typeOpen As String = "binary", Optional ByVal fLen As Integer = 1) As Long
    On Error GoTo err1
    
    If NumFile <= 0 Then NumFile = VBA.FreeFile
    
    numLenRandom(NumFile) = 1
    
    Select Case LCase$(typeOpen)
        Case "input":       Open nameFile For Input As #NumFile
        Case "output":      Open nameFile For Output As #NumFile
        Case "append":      Open nameFile For Append As #NumFile
        Case "binary":      Open nameFile For Binary As #NumFile
        Case "random":      Open nameFile For Binary As #NumFile:       numLenRandom(NumFile) = fLen
    End Select
    
    Open2 = NumFile
    Exit Function
err1:
    Open2 = -1
End Function

Public Function Seek2(ByVal NumFile As Integer, Optional ByVal numPos As Long = -1) As Long
    On Error Resume Next
    If numPos < 1 Then Seek2 = Seek(NumFile) Else Seek NumFile, numPos
End Function

Public Sub Lock2(ByVal NumFile As Integer)
    On Error Resume Next
    Lock NumFile
End Sub

Public Sub UnLock2(ByVal NumFile As Integer)
    On Error Resume Next
    Unlock NumFile
End Sub

Public Function EOF(ByVal NumFile As Integer) As Boolean
    EOF = VBA.EOF(NumFile)
End Function

Public Function LOF(ByVal NumFile As Integer) As Long
    LOF = VBA.LOF(NumFile)
End Function

Public Function LOC(ByVal NumFile As Integer) As Long
    LOC = VBA.LOC(NumFile)
End Function

Public Function LineInput2(ByVal NumFile As Integer, Optional txt As Variant) As String
    Dim tmp As String
    On Error Resume Next
    Line Input #NumFile, tmp
    txt = tmp
    LineInput2 = tmp
End Function

Public Function Input2(ByVal NumFile As Integer, Optional txt As Variant) As String
    Dim tmp As String
    On Error Resume Next
    Input #NumFile, tmp
    txt = tmp
    Input2 = tmp
End Function

Public Sub Print2(ByVal NumFile As Integer, Optional ByVal txt As String = "")
    On Error Resume Next
    Print #NumFile, txt
End Sub

Public Sub Get2(ByVal NumFile As Integer, Optional numPosFile As Variant, Optional Buf As Variant)
    Dim tmpBuf() As Byte, tmpByte As Byte, tmpInteger As Integer, tmpLong As Long
    Dim tmpBoolean As Boolean, tmpCurrency As Currency, tmpDouble As Double
    Dim tmpSingle As Single, tmpString As String, fs As Long, numPos As Long
    
    On Error Resume Next
    
    If Not IsMissing(numPosFile) Then numPos = numPosFile
    If numLenRandom(NumFile) > 1 Then numPos = (numPos - 1) * numLenRandom(NumFile) + 1
    
    Select Case VariantType(Buf, True)
        Case vbArray + vbByte
            fs = m_ArraySize(Buf) - 1
            If fs < 0 Then Exit Sub
            ReDim tmpBuf(fs)
            If numPos <= 0 Then Get #NumFile, , tmpBuf Else Get #NumFile, numPos, tmpBuf
            Buf = tmpBuf
            
        Case vbArray + vbVariant
            fs = m_ArraySize(Buf) - 1
            If fs < 0 Then Exit Sub
            ReDim tmpBuf(fs)
            If numPos <= 0 Then Get #NumFile, , tmpBuf Else Get #NumFile, numPos, tmpBuf
            ConvFromBufferByte Buf, tmpBuf
            
        Case vbObject
            fs = Buf.Class___Size
            ReDim tmpBuf(fs - 1)
            If numPos <= 0 Then Get #NumFile, , tmpBuf Else Get #NumFile, numPos, tmpBuf
            tmpLong = Buf.Class___Wrapper.Ptr + Buf.Class___Offset
            CopyMemory ByVal tmpLong, tmpBuf(0), fs
            
        Case vbByte
            If numPos <= 0 Then Get #NumFile, , tmpByte Else Get #NumFile, numPos, tmpByte
            Buf = tmpByte
            
        Case vbLong
            If numPos <= 0 Then Get #NumFile, , tmpLong Else Get #NumFile, numPos, tmpLong
            Buf = tmpLong
            
        Case vbInteger
            If numPos <= 0 Then Get #NumFile, , tmpInteger Else Get #NumFile, numPos, tmpInteger
            Buf = tmpInteger
            
        Case vbString
            tmpString = Buf
            If numPos <= 0 Then Get #NumFile, , tmpString Else Get #NumFile, numPos, tmpString
            Buf = tmpString
            
        Case vbDouble
            If numPos <= 0 Then Get #NumFile, , tmpDouble Else Get #NumFile, numPos, tmpDouble
            Buf = tmpDouble
            
        Case vbSingle
            If numPos <= 0 Then Get #NumFile, , tmpSingle Else Get #NumFile, numPos, tmpSingle
            Buf = tmpSingle
            
        Case vbBoolean
            If numPos <= 0 Then Get #NumFile, , tmpBoolean Else Get #NumFile, numPos, tmpBoolean
            Buf = tmpBoolean
            
        Case vbCurrency
            If numPos <= 0 Then Get #NumFile, , tmpCurrency Else Get #NumFile, numPos, tmpCurrency
            Buf = tmpCurrency
    End Select
End Sub

Public Sub Put2(ByVal NumFile As Integer, Optional numPosFile As Variant, Optional Buf As Variant)
    Dim tmpBuf() As Byte, tmpByte As Byte, tmpInteger As Integer, tmpLong As Long
    Dim tmpBoolean As Boolean, tmpCurrency As Currency, tmpDouble As Double
    Dim tmpSingle As Single, tmpString As String, fs As Long, numPos As Long
    
    On Error Resume Next
    
    If Not IsMissing(numPosFile) Then numPos = numPosFile
    If numLenRandom(NumFile) > 1 Then numPos = (numPos - 1) * numLenRandom(NumFile) + 1

    Select Case VariantType(Buf, True)
        Case vbArray + vbByte
            tmpBuf = Buf
            If numPos <= 0 Then Put #NumFile, , tmpBuf Else Put #NumFile, numPos, tmpBuf
            
        Case vbArray + vbVariant
            ConvToBufferByte Buf, tmpBuf
            If numPos <= 0 Then Put #NumFile, , tmpBuf Else Put #NumFile, numPos, tmpBuf
            
        Case vbObject
            fs = Buf.Class___Size
            ReDim tmpBuf(fs - 1)
            tmpLong = Buf.Class___Wrapper.Ptr + Buf.Class___Offset
            CopyMemory tmpBuf(0), ByVal tmpLong, fs
            If numPos <= 0 Then Put #NumFile, , tmpBuf Else Put #NumFile, numPos, tmpBuf
            
        Case vbByte
            tmpByte = Buf
            If numPos <= 0 Then Put #NumFile, , tmpByte Else Put #NumFile, numPos, tmpByte
            
        Case vbLong
            tmpLong = Buf
            If numPos <= 0 Then Put #NumFile, , tmpLong Else Put #NumFile, numPos, tmpLong

        Case vbInteger
            tmpInteger = Buf
            If numPos <= 0 Then Put #NumFile, , tmpInteger Else Put #NumFile, numPos, tmpInteger
            
        Case vbString
            tmpString = Buf
            If numPos <= 0 Then Put #NumFile, , tmpString Else Put #NumFile, numPos, tmpString
            
        Case vbDouble
            tmpDouble = Buf
            If numPos <= 0 Then Put #NumFile, , tmpDouble Else Put #NumFile, numPos, tmpDouble
            
        Case vbSingle
            tmpSingle = Buf
            If numPos <= 0 Then Put #NumFile, , tmpSingle Else Put #NumFile, numPos, tmpSingle
            
        Case vbBoolean
            tmpBoolean = Buf
            If numPos <= 0 Then Put #NumFile, , tmpBoolean Else Put #NumFile, numPos, tmpBoolean
            
        Case vbCurrency
            tmpCurrency = Buf
            If numPos <= 0 Then Put #NumFile, , tmpCurrency Else Put #NumFile, numPos, tmpCurrency
    End Select
End Sub

Public Sub Close2(ByVal NumFile As Integer)
    On Error Resume Next
    Close #NumFile
End Sub

Public Sub Reset()
    On Error Resume Next
    VBA.FileSystem.Reset
End Sub

Public Function FreeFile() As Long
    FreeFile = VBA.FreeFile
End Function

Public Function ShortName(ByVal fName As String) As String
    ShortName = FileShortName(fName)
End Function

Public Function LongName(ByVal fName As String) As String
    LongName = FileLongName(fName)
End Function

Public Function FileDateTime(ByVal fName As String, Optional ByVal numDT As Long = 3) As String
    Dim hFile As Long, ft1 As FILETIME, ft2 As FILETIME, ft3 As FILETIME, st As SYSTEMTIME
    
    hFile = CreateFile(StrPtr(LongPath(fName)), GENERIC_READ, 0, 0, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0)
    If hFile = INVALID_HANDLE Then Exit Function
    GetFileTime hFile, ft1, ft2, ft3
    CloseHandle hFile
    
    If numDT = 1 Then ft3 = ft1
    If numDT = 2 Then ft3 = ft2
    
    FileTimeToLocalFileTime ft3, ft3
    FileTimeToSystemTime ft3, st
    FileDateTime = DateSerial(st.wYear, st.wMonth, st.wDay) & " " & Format$(TimeSerial(st.wHour, st.wMinute, st.wSecond), "hh:mm:ss")
End Function

Public Property Get Attr(ByVal fName As String) As Long
    Attr = GetFileAttributesW(StrPtr(LongPath(fName)))
End Property

Public Property Let Attr(ByVal fName As String, ByVal value As Long)
    SetFileAttributesW StrPtr(LongPath(fName)), value
End Property

Public Function GetDir(Optional ByVal MPath As String = "-1", Optional ByVal mAttr As Integer = 16) As String
    On Error Resume Next
    If MPath = "-1" Then GetDir = Dir Else GetDir = Dir$(MPath, mAttr)
End Function

Public Sub FileCopy(ByVal fScr As String, ByVal fDest As String)
    modCommon.FileCopy fScr, fDest
End Sub

Public Sub FileMove(ByVal fScr As String, ByVal fDest As String)
    modCommon.FileMove fScr, fDest
End Sub

Public Sub FileKill(ByVal fScr As String)
    modCommon.FileKill fScr
End Sub

Public Sub FolderCopy(ByVal fScr As String, ByVal fDest As String)
    Dim fso As Object
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CopyFolder fScr, fDest, True
    Set fso = Nothing
End Sub

Public Sub FolderMove(ByVal fScr As String, ByVal fDest As String)
    Dim fso As Object
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.MoveFolder fScr, fDest
    Set fso = Nothing
End Sub

Public Sub ChDrive(ByVal nameDrive As String)
    On Error Resume Next
    VBA.ChDrive nameDrive
End Sub

Public Sub RmDir(ByVal nameDir As String)
    modCommon.RmDir nameDir
End Sub

Public Sub ChDir(ByVal nameDir As String)
    modCommon.ChDir nameDir
End Sub

Public Function CurDir() As String
    CurDir = modCommon.CurDir
End Function

Public Sub CreateDir(ByVal nameDir As String)
    modCommon.CreateDir nameDir
End Sub

Public Sub MkDir(ByVal nameDir As String)
    modCommon.MkDir nameDir
End Sub

Public Function Exists(ByVal fScr As String, Optional ByVal mask_L As Long = -17, Optional ByVal mask_H As Long = 0) As Boolean
    Exists = IsFile(fScr, mask_L, mask_H)
End Function

Public Function FormatBytes(ByVal numBytes As Double, Optional ByVal arrUnit As Variant) As String
    FormatBytes = modCommon.FormatBytes(numBytes, arrUnit)
End Function

Public Function GetFileName(ByVal txtFile As String, Optional ByVal dm As String = "\") As String
    GetFileName = modCommon.GetFileName(txtFile, dm)
End Function

Public Function GetDirectory(ByVal txtFile As String, Optional ByVal dm As String = "\") As String
    GetDirectory = modCommon.GetDirectory(txtFile, dm)
End Function

Public Function GetFileTitle(ByVal txtFile As String, Optional ByVal dm As String = "\") As String
    GetFileTitle = modCommon.GetFileTitle(txtFile, dm)
End Function

Public Function GetExtension(ByVal txtFile As String, Optional ByVal dm As String = "\") As String
    GetExtension = modCommon.GetExtension(txtFile, dm)
End Function

Public Function CPath(ByVal fPath As String, Optional ByVal typePath As Boolean = True, Optional ByVal delim As String = "\") As String
    CPath = modCommon.CPath(fPath, typePath, delim)
End Function

Public Function LPath(ByVal fPath As String, Optional ByVal numLevel As Long = 0, Optional ByVal delim As String = "\") As String
    Dim a As Long, cnt As Long, txt() As String
    
    If LenB(fPath) Then
        If Right$(fPath, 1) = delim Then fPath = Left$(fPath, Len(fPath) - 1)
        
        txt = Split(fPath, delim)
        
        cnt = UBound(txt)
        
        If numLevel > 0 Then
            For a = numLevel To cnt
                LPath = LPath + txt(a) + delim
            Next
        Else
            If cnt + numLevel < 0 Then numLevel = -cnt
            
            For a = 0 To cnt + numLevel
                LPath = LPath + txt(a) + delim
            Next
        End If
    End If
End Function

Public Function GetDiskInfo(ByVal sDrive As String) As Collection
    Dim bytes_Avail As LONG64, bytes_Total As LONG64, bytes_Free As LONG64
    Dim dbl_Total As Double, dbl_Free As Double, cl As New Collection, vSerial As Long
    Dim vLabel As String, vName As String, oSerial As String, iDriveType() As Variant
    
    On Error Resume Next
    
    GetDiskFreeSpaceExA sDrive, bytes_Avail, bytes_Total, bytes_Free
    
    dbl_Total = BigLongToDouble(bytes_Total.LowPart, bytes_Total.HighPart)
    dbl_Free = BigLongToDouble(bytes_Free.LowPart, bytes_Free.HighPart)
    
    cl.Add dbl_Free, "Free"
    cl.Add dbl_Total, "Total"
    cl.Add dbl_Total - dbl_Free, "Used"
    
    cl.Add GetDriveTypeA(sDrive), "Type"
    
    iDriveType = Array("Missing", "Unknown", "Removable", "Fixed", "Remote", "CDRom", "RAMDisk")
    cl.Add iDriveType(GetDriveTypeA(sDrive)), "TypeStr"
    
    vLabel = String$(255, 0)
    vName = String$(255, 0)
    GetVolumeInformationA sDrive, vLabel, 255, vSerial, 255, 0, vName, 255
    vLabel = Trim$(Left$(vLabel, InStr(vLabel, vbNullChar) - 1))
    vName = Trim$(Left$(vName, InStr(vName, vbNullChar) - 1))
    
    If vSerial <> 0 Then oSerial = Hex$(vSerial):    oSerial = Left$(oSerial, 4) + "-" + Right$(oSerial, 4)

    cl.Add vName, "Name"
    cl.Add oSerial, "Serial"
    cl.Add vSerial, "SerialLng"
    cl.Add vLabel, "Label"
    
    Set GetDiskInfo = cl
    Set cl = Nothing
End Function

Public Function GetLogicalDrives() As Collection
    Dim txt As String * 250, lenDrv As Integer, a As Integer
    Dim tmpTxt As String, tmpStr() As String, cl As New Collection
    
    On Error Resume Next
    
    lenDrv = GetLogicalDriveStringsA(250, txt)
    If lenDrv <> 0 Then
       tmpTxt = Left$(txt, lenDrv)
       tmpStr = Split(tmpTxt, vbNullChar)
       For a = 0 To UBound(tmpStr) - 1
         cl.Add tmpStr(a), "Drive_" & a
       Next
    End If
    
    Set GetLogicalDrives = cl
    Set cl = Nothing
End Function

Public Function GetDiskType(ByVal sDrive As String) As Long
    On Error Resume Next
    GetDiskType = GetDriveTypeA(sDrive)
End Function


Private Sub Class_Initialize()
    Set Search = New clsFileSearch
End Sub
