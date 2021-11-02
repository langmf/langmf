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
    Set API = DoParams(New clsFileAPI, Param)
End Function

Public Function Open2(ByVal nameFile As String, Optional ByVal numFile As Integer = 0, Optional ByVal typeOpen As String = "binary", Optional ByVal fLen As Integer = 1) As Long
    On Error GoTo err1
    
    If numFile <= 0 Then numFile = VBA.FreeFile
    
    numLenRandom(numFile) = 1
    
    Select Case LCase$(typeOpen)
        Case "input":       Open nameFile For Input As #numFile
        Case "output":      Open nameFile For Output As #numFile
        Case "append":      Open nameFile For Append As #numFile
        Case "binary":      Open nameFile For Binary As #numFile
        Case "random":      Open nameFile For Binary As #numFile:       numLenRandom(numFile) = fLen
    End Select
    
    Open2 = numFile
    Exit Function
err1:
    Open2 = -1
End Function

Public Function Seek2(ByVal numFile As Integer, Optional ByVal numPos As Long = -1) As Long
    On Error Resume Next
    If numPos < 1 Then Seek2 = Seek(numFile) Else Seek numFile, numPos
End Function

Public Sub Lock2(ByVal numFile As Integer)
    On Error Resume Next
    Lock numFile
End Sub

Public Sub UnLock2(ByVal numFile As Integer)
    On Error Resume Next
    Unlock numFile
End Sub

Public Function EOF(ByVal numFile As Integer) As Boolean
    EOF = VBA.EOF(numFile)
End Function

Public Function LOF(ByVal numFile As Integer) As Long
    LOF = VBA.LOF(numFile)
End Function

Public Function LOC(ByVal numFile As Integer) As Long
    LOC = VBA.LOC(numFile)
End Function

Public Function LineInput2(ByVal numFile As Integer, Optional value As Variant) As String
    Dim tmp As String
    On Error Resume Next
    Line Input #numFile, tmp:       value = tmp:        LineInput2 = tmp
End Function

Public Function Input2(ByVal numFile As Integer, Optional value As Variant) As String
    Dim tmp As String
    On Error Resume Next
    Input #numFile, tmp:            value = tmp:        Input2 = tmp
End Function

Public Sub Print2(ByVal numFile As Integer, Optional ByVal value As String = "")
    On Error Resume Next
    Print #numFile, value
End Sub

Public Sub Get2(ByVal numFile As Integer, Optional vPos As Variant, Optional Buf As Variant)
    Dim tmpBuf() As Byte, tmpByte As Byte, tmpInteger As Integer, tmpLong As Long
    Dim tmpBoolean As Boolean, tmpCurrency As Currency, tmpDouble As Double
    Dim tmpSingle As Single, tmpString As String, fs As Long, numPos As Long
    
    On Error Resume Next
    
    If Not IsMissing(vPos) Then numPos = vPos
    If numLenRandom(numFile) > 1 Then numPos = (numPos - 1) * numLenRandom(numFile) + 1
    
    Select Case VariantType(Buf, True)
        Case vbArray + vbByte
            fs = ArraySize(Buf) - 1
            If fs < 0 Then Exit Sub
            ReDim tmpBuf(fs)
            If numPos <= 0 Then Get #numFile, , tmpBuf Else Get #numFile, numPos, tmpBuf
            Buf = tmpBuf
            
        Case vbArray + vbVariant
            fs = ArraySize(Buf) - 1
            If fs < 0 Then Exit Sub
            ReDim tmpBuf(fs)
            If numPos <= 0 Then Get #numFile, , tmpBuf Else Get #numFile, numPos, tmpBuf
            ConvFromBufferByte Buf, tmpBuf
            
        Case vbObject
            fs = Buf.•SIZE
            ReDim tmpBuf(fs - 1)
            If numPos <= 0 Then Get #numFile, , tmpBuf Else Get #numFile, numPos, tmpBuf
            tmpLong = Buf.•WRAP.Ptr + Buf.•OFS
            CopyMemory ByVal tmpLong, tmpBuf(0), fs
            
        Case vbByte
            If numPos <= 0 Then Get #numFile, , tmpByte Else Get #numFile, numPos, tmpByte
            Buf = tmpByte
            
        Case vbLong
            If numPos <= 0 Then Get #numFile, , tmpLong Else Get #numFile, numPos, tmpLong
            Buf = tmpLong
            
        Case vbInteger
            If numPos <= 0 Then Get #numFile, , tmpInteger Else Get #numFile, numPos, tmpInteger
            Buf = tmpInteger
            
        Case vbString
            tmpString = Buf
            If numPos <= 0 Then Get #numFile, , tmpString Else Get #numFile, numPos, tmpString
            Buf = tmpString
            
        Case vbDouble
            If numPos <= 0 Then Get #numFile, , tmpDouble Else Get #numFile, numPos, tmpDouble
            Buf = tmpDouble
            
        Case vbSingle
            If numPos <= 0 Then Get #numFile, , tmpSingle Else Get #numFile, numPos, tmpSingle
            Buf = tmpSingle
            
        Case vbBoolean
            If numPos <= 0 Then Get #numFile, , tmpBoolean Else Get #numFile, numPos, tmpBoolean
            Buf = tmpBoolean
            
        Case vbCurrency
            If numPos <= 0 Then Get #numFile, , tmpCurrency Else Get #numFile, numPos, tmpCurrency
            Buf = tmpCurrency
    End Select
End Sub

Public Sub Put2(ByVal numFile As Integer, Optional vPos As Variant, Optional Buf As Variant)
    Dim tmpBuf() As Byte, tmpByte As Byte, tmpInteger As Integer, tmpLong As Long
    Dim tmpBoolean As Boolean, tmpCurrency As Currency, tmpDouble As Double
    Dim tmpSingle As Single, tmpString As String, fs As Long, numPos As Long
    
    On Error Resume Next
    
    If Not IsMissing(vPos) Then numPos = vPos
    If numLenRandom(numFile) > 1 Then numPos = (numPos - 1) * numLenRandom(numFile) + 1

    Select Case VariantType(Buf, True)
        Case vbArray + vbByte
            tmpBuf = Buf
            If numPos <= 0 Then Put #numFile, , tmpBuf Else Put #numFile, numPos, tmpBuf
            
        Case vbArray + vbVariant
            ConvToBufferByte Buf, tmpBuf
            If numPos <= 0 Then Put #numFile, , tmpBuf Else Put #numFile, numPos, tmpBuf
            
        Case vbObject
            fs = Buf.•SIZE
            ReDim tmpBuf(fs - 1)
            tmpLong = Buf.•WRAP.Ptr + Buf.•OFS
            CopyMemory tmpBuf(0), ByVal tmpLong, fs
            If numPos <= 0 Then Put #numFile, , tmpBuf Else Put #numFile, numPos, tmpBuf
            
        Case vbByte
            tmpByte = Buf
            If numPos <= 0 Then Put #numFile, , tmpByte Else Put #numFile, numPos, tmpByte
            
        Case vbLong
            tmpLong = Buf
            If numPos <= 0 Then Put #numFile, , tmpLong Else Put #numFile, numPos, tmpLong

        Case vbInteger
            tmpInteger = Buf
            If numPos <= 0 Then Put #numFile, , tmpInteger Else Put #numFile, numPos, tmpInteger
            
        Case vbString
            tmpString = Buf
            If numPos <= 0 Then Put #numFile, , tmpString Else Put #numFile, numPos, tmpString
            
        Case vbDouble
            tmpDouble = Buf
            If numPos <= 0 Then Put #numFile, , tmpDouble Else Put #numFile, numPos, tmpDouble
            
        Case vbSingle
            tmpSingle = Buf
            If numPos <= 0 Then Put #numFile, , tmpSingle Else Put #numFile, numPos, tmpSingle
            
        Case vbBoolean
            tmpBoolean = Buf
            If numPos <= 0 Then Put #numFile, , tmpBoolean Else Put #numFile, numPos, tmpBoolean
            
        Case vbCurrency
            tmpCurrency = Buf
            If numPos <= 0 Then Put #numFile, , tmpCurrency Else Put #numFile, numPos, tmpCurrency
    End Select
End Sub

Public Sub Close2(ByVal numFile As Integer)
    On Error Resume Next
    Close #numFile
End Sub

Public Sub Reset()
    On Error Resume Next
    VBA.FileSystem.Reset
End Sub

Public Property Get FreeFile() As Long
    FreeFile = VBA.FreeFile
End Property

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

Public Property Get GetDir(Optional ByVal MPath As String = "-1", Optional ByVal mAttr As Integer = 16) As String
    On Error Resume Next
    If MPath = "-1" Then GetDir = Dir Else GetDir = Dir$(MPath, mAttr)
End Property

Public Sub FileCopy(ByVal fName As String, ByVal fNew As String)
    modCommon.FileCopy fName, fNew
End Sub

Public Sub FileMove(ByVal fName As String, ByVal fNew As String)
    modCommon.FileMove fName, fNew
End Sub

Public Sub FileKill(ByVal fName As String)
    modCommon.FileKill fName
End Sub

Public Sub FolderCopy(ByVal fName As String, ByVal fNew As String)
    Dim fso As Object
    On Error Resume Next
    Set fso = VBA.CreateObject("Scripting.FileSystemObject")
    fso.CopyFolder fName, fNew, True
    Set fso = Nothing
End Sub

Public Sub FolderMove(ByVal fName As String, ByVal fNew As String)
    Dim fso As Object
    On Error Resume Next
    Set fso = VBA.CreateObject("Scripting.FileSystemObject")
    fso.MoveFolder fName, fNew
    Set fso = Nothing
End Sub

Public Sub ChDrive(ByVal value As String)
    On Error Resume Next
    VBA.ChDrive value
End Sub

Public Sub ChDir(ByVal value As String)
    modCommon.ChDir value
End Sub

Public Property Get CurDir() As String
    CurDir = modCommon.CurDir
End Property

Public Sub RemoveDir(ByVal value As String)
    modCommon.RemoveDir value
End Sub

Public Sub RmDir(ByVal value As String)
    modCommon.RmDir value
End Sub

Public Sub CreateDir(ByVal value As String)
    modCommon.CreateDir value
End Sub

Public Sub MkDir(ByVal value As String)
    modCommon.MkDir value
End Sub

Public Function Exists(ByVal fName As String, Optional ByVal mask_L As Long = -17, Optional ByVal mask_H As Long = 0) As Boolean
    Exists = IsFile(fName, mask_L, mask_H)
End Function

Public Function FormatBytes(ByVal value As Double, Optional ByVal arrUnit As Variant) As String
    FormatBytes = modCommon.FormatBytes(value, arrUnit)
End Function

Public Function GetFileName(ByVal fName As String, Optional ByVal dm As String = "\") As String
    GetFileName = modCommon.GetFileName(fName, dm)
End Function

Public Function GetDirectory(ByVal fName As String, Optional ByVal dm As String = "\") As String
    GetDirectory = modCommon.GetDirectory(fName, dm)
End Function

Public Function GetFileTitle(ByVal fName As String, Optional ByVal dm As String = "\") As String
    GetFileTitle = modCommon.GetFileTitle(fName, dm)
End Function

Public Function GetExtension(ByVal fName As String, Optional ByVal dm As String = "\") As String
    GetExtension = modCommon.GetExtension(fName, dm)
End Function

Public Function CPath(ByVal fPath As String, Optional ByVal typePath As Boolean = True, Optional ByVal dm As String = "\") As String
    CPath = modCommon.CPath(fPath, typePath, dm)
End Function

Public Function LPath(ByVal fPath As String, Optional ByVal numLevel As Long = 0, Optional ByVal dm As String = "\") As String
    Dim a As Long, cnt As Long, txt() As String
    
    If LenB(fPath) Then
        If Right$(fPath, 1) = dm Then fPath = Left$(fPath, Len(fPath) - 1)
        
        txt = Split(fPath, dm)
        
        cnt = UBound(txt)
        
        If numLevel > 0 Then
            For a = numLevel To cnt
                LPath = LPath + txt(a) + dm
            Next
        Else
            If cnt + numLevel < 0 Then numLevel = -cnt
            
            For a = 0 To cnt + numLevel
                LPath = LPath + txt(a) + dm
            Next
        End If
    End If
End Function

Public Property Get GetDiskInfo(Optional ByVal sDrive As String = "C:\") As Collection
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
End Property

Public Property Get GetLogicalDrives() As Collection
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
End Property

Public Property Get GetDiskType(Optional ByVal sDrive As String = "C:\") As Long
    On Error Resume Next
    GetDiskType = GetDriveTypeA(sDrive)
End Property


Private Sub Class_Initialize()
    Set Search = New clsFileSearch
End Sub
