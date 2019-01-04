VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Base64 As New clsBase64


Public Function EncodeUTF8(ByVal value As String, Optional ByVal Cpg As Long = 65001) As String
    If LenB(value) Then EncodeUTF8 = modCommon.EncodeUTF8(value, Cpg)
End Function

Public Function DecodeUTF8(ByVal value As String, Optional ByVal Cpg As Long = 65001) As String
    If LenB(value) Then DecodeUTF8 = modCommon.DecodeUTF8(value, Cpg)
End Function

Public Function A2W(ByVal value As String, Optional ByVal Cpg As Long = -1) As String
    A2W = Conv_A2W_Str(value, Cpg)
End Function

Public Function W2A(ByVal value As String, Optional ByVal Cpg As Long = -1) As String
    W2A = Conv_W2A_Str(value, Cpg)
End Function

Public Function EncodeBase64(value As Variant) As String
    Dim tmpBuf() As Byte
    If ConvToBufferByte(value, tmpBuf) Then EncodeBase64 = Conv_A2W_Buf(Base64.Encode(tmpBuf))
End Function

Public Function DecodeBase64(value As Variant) As String
    Dim tmpBuf() As Byte
    If ConvToBufferByte(value, tmpBuf) Then DecodeBase64 = Conv_A2W_Buf(Base64.Decode(tmpBuf))
End Function

Public Function Compress(value As Variant, Optional ByVal cmsType As Long = CMS_FORMAT_ZLIB) As Long
    Dim tmpBuf() As Byte
    If Not ConvToBufferByte(value, tmpBuf) Then Compress = -1:    Exit Function
    Compress = CompressData(tmpBuf(), cmsType)
    ConvFromBufferByte value, tmpBuf
End Function

Public Function Decompress(value As Variant, Optional ByVal cmsType As Long = CMS_FORMAT_ZLIB) As Long
    Dim tmpBuf() As Byte
    If Not ConvToBufferByte(value, tmpBuf) Then Decompress = -1:    Exit Function
    Decompress = DecompressData(tmpBuf(), cmsType)
    ConvFromBufferByte value, tmpBuf
End Function

Public Function Buf(value As Variant, Optional ByVal typeVar As Integer = vbArray + vbVariant) As Variant
    Dim tmpBuf() As Byte
    If ConvToBufferByte(value, tmpBuf) Then ConvFromBufferByte Buf, tmpBuf, typeVar
End Function

Public Function EncodeUrl(ByVal value As String, Optional ByVal Cpg As Long = -1) As String
    Dim i As Long, b As Byte, tmpBuf() As Byte
    If Len(value) = 0 Then Exit Function
    tmpBuf = Conv_W2A_Buf(value, Cpg):   value = ""
    For i = 0 To UBound(tmpBuf)
        b = tmpBuf(i)
        If (b >= 48 And b <= 57) Or (b >= 65 And b <= 90) Or (b >= 97 And b <= 122) Then
            EncodeUrl = EncodeUrl + Chr$(b)
        ElseIf b < 16 Then
            EncodeUrl = EncodeUrl + "%0" + Hex$(b)
        Else
            EncodeUrl = EncodeUrl + "%" + Hex$(b)
        End If
    Next
End Function

Public Function DecodeUrl(ByVal value As String, Optional ByVal Cpg As Long = -1) As String
    Dim i As Long, c As Long, b As Byte, srcBuf() As Byte, dstBuf() As Byte
    If Len(value) = 0 Then Exit Function
    srcBuf = Conv_W2A_Buf(value):   value = "":   ReDim dstBuf(UBound(srcBuf))
    For i = 0 To UBound(srcBuf)
        b = srcBuf(i)
        If b = 37 Then
            b = GT_IHex(srcBuf(i + 1)) * 16 + GT_IHex(srcBuf(i + 2)):   i = i + 2
        ElseIf b = 43 Then
            b = 32
        End If
        dstBuf(c) = b:   c = c + 1
    Next
    Erase srcBuf:   ReDim Preserve dstBuf(c - 1):   DecodeUrl = Conv_A2W_Buf(dstBuf, Cpg)
End Function

Public Function ArrayJS(ByVal Obj As ATL.IDispatchEx) As Variant()
    Dim a As Long, uds As Long, sz As Long, out() As Variant, pAE As Long, pDP As ATL.DISPPARAMS, pEI As ATL.EXCEPINFO

    On Error Resume Next:    sz = CallByName(Obj, "length", VbGet):    On Error GoTo 0
    If sz = 0 Then ArrayJS = out:   Exit Function

    ReDim out(sz - 1)

    For a = 0 To sz - 1
        Call Obj.GetDispID(a, fdexNameCaseSensitive, uds)
        Call Obj.InvokeEx(uds, LOCALE_USER_DEFAULT, VbGet, pDP, out(a), pEI, pAE)
    Next

    ArrayJS = out
End Function

Public Function Image2Str(ByVal Obj As IPictureDisp) As String
    Dim fName As String
    fName = GetTmpPath + "temp_" + GenTempStr + ".pic"
    FileKill fName:    SavePicture Obj, fName:     modCommon.File2Str Image2Str, fName:    FileKill fName
End Function

Public Function Str2Image(value As Variant) As IPictureDisp
    Set Str2Image = LoadPictureFromByte(value)
End Function

Public Function File2Buf(ByVal nameFile As String) As Byte()
    modCommon.File2Buf File2Buf, nameFile
End Function

Public Function Buf2File(value As Variant, ByVal nameFile As String) As Boolean
    Dim tmpBuf() As Byte
    If ConvToBufferByte(value, tmpBuf) Then Buf2File = modCommon.Buf2File(tmpBuf, nameFile)
End Function

Public Function File2Str(ByVal nameFile As String) As String
    modCommon.File2Str File2Str, nameFile
End Function

Public Function Str2File(ByVal value As String, ByVal nameFile As String) As Boolean
    Str2File = modCommon.Str2File(value, nameFile)
End Function

Public Function Hex2Buf(value As Variant) As Byte()
    Dim tmpBuf() As Byte
    If ConvToBufferByte(value, tmpBuf) Then Hex2Buf = modCommon.Hex2Buf(tmpBuf)
End Function

Public Function Buf2Hex(value As Variant) As String
    Dim tmpBuf() As Byte
    If ConvToBufferByte(value, tmpBuf) Then Buf2Hex = Conv_A2W_Buf(modCommon.Buf2Hex(tmpBuf))
End Function

Public Function MPath(ByVal value As String) As String
    MPath = Parse_MPath(value)
End Function

Public Function Ptr2Str(ByVal pStr As Long, Optional ByVal ptrType As Long = 0, Optional ByVal Cols As Long = 0, Optional ByVal Rows As Long = 0) As Variant
    Dim c As Long, r As Long, arr() As Variant, p() As Long, old As Long
    
    If Cols = 0 Then
        If ptrType = 0 Then Ptr2Str = GetStringPtrA(pStr)
        If ptrType = 1 Then Ptr2Str = GetStringPtrW(pStr)
        If ptrType = 2 Then Ptr2Str = GetStringPtrU(pStr)
    Else
        If Rows = 0 Then
            ReDim arr(Cols - 1)
            ReDim p(Cols - 1)
            
            old = MapArray(VarPtrArray(p), pStr)
            
            If ptrType = 0 Then For c = 0 To Cols - 1:    arr(c) = GetStringPtrA(p(c)):    Next
            If ptrType = 1 Then For c = 0 To Cols - 1:    arr(c) = GetStringPtrW(p(c)):    Next
            If ptrType = 2 Then For c = 0 To Cols - 1:    arr(c) = GetStringPtrU(p(c)):    Next
        Else
            ReDim arr(Cols - 1, Rows - 1)
            ReDim p(Cols - 1, Rows - 1)
            
            old = MapArray(VarPtrArray(p), pStr)
            
            If ptrType = 0 Then For r = 0 To Rows - 1:    For c = 0 To Cols - 1:    arr(c, r) = GetStringPtrA(p(c, r)):    Next:    Next
            If ptrType = 1 Then For r = 0 To Rows - 1:    For c = 0 To Cols - 1:    arr(c, r) = GetStringPtrW(p(c, r)):    Next:    Next
            If ptrType = 2 Then For r = 0 To Rows - 1:    For c = 0 To Cols - 1:    arr(c, r) = GetStringPtrU(p(c, r)):    Next:    Next
        End If
        
        Call MapArray(VarPtrArray(p), old)
        
        Ptr2Str = arr
    End If
End Function

Public Function XLT(srcTxt As Variant, curTBL As Variant, Optional ByVal bInvert As Boolean = False) As Variant
    Dim a As Long, uds As Long, tmpBuf() As Byte, TBL(255) As Byte
    
    For a = 0 To 255:    TBL(a) = a:    Next
    
    ConvToBufferByte curTBL, tmpBuf:    uds = ArraySize(tmpBuf) - 1
    
    If bInvert Then
        For a = 0 To uds Step 2:    TBL(tmpBuf(a + 1)) = tmpBuf(a):    Next
    Else
        For a = 0 To uds Step 2:    TBL(tmpBuf(a)) = tmpBuf(a + 1):    Next
    End If
    
    ConvToBufferByte srcTxt, tmpBuf:    uds = ArraySize(tmpBuf) - 1
    
    For a = 0 To uds:    tmpBuf(a) = TBL(tmpBuf(a)):    Next
    
    XLT = Buf(tmpBuf, VarType(srcTxt))
End Function

Public Function HashFile(ByVal FileName As String, Optional ByVal Algorithm As Long = SHA1) As String
    Dim hCtx As Long, hHash As Long, lRes As Long, lLen As Long, lIdx As Long, f As New clsFileAPI, abHash() As Byte
    
    Const BLOCK_SIZE As Long = 64 * 1024&
    
    If f.FOpen(FileName, OPEN_EXISTING, GENERIC_READ) = INVALID_HANDLE Then Exit Function

    lRes = CryptAcquireContext(hCtx, vbNullString, vbNullString, IIF(Algorithm >= SHA256, PROV_RSA_AES, PROV_RSA_FULL), CRYPT_VERIFYCONTEXT)
    
    If lRes <> 0 Then
        lRes = CryptCreateHash(hCtx, Algorithm, 0, 0, hHash)
        
        If lRes <> 0 Then
            ReDim abBlock(1 To BLOCK_SIZE) As Byte
            Dim lCount As Long, lBlocks As Long, lLastBlock As Long

            lBlocks = f.LOF \ BLOCK_SIZE
            lLastBlock = f.LOF - lBlocks * BLOCK_SIZE
            
            For lCount = 1 To lBlocks
               f.GetMem VarPtr(abBlock(1)), BLOCK_SIZE
               lRes = CryptHashData(hHash, abBlock(1), BLOCK_SIZE, 0)
               If lRes = 0 Then Exit For
            Next
            
            If lLastBlock > 0 And lRes <> 0 Then
               ReDim abBlock(1 To lLastBlock) As Byte
               f.GetMem VarPtr(abBlock(1)), lLastBlock
               lRes = CryptHashData(hHash, abBlock(1), lLastBlock, 0)
            End If
          
            If lRes <> 0 Then
                lRes = CryptGetHashParam(hHash, HP_HASHSIZE, lLen, 4, 0)
                
                If lRes <> 0 Then
                    ReDim abHash(0 To lLen - 1)
                
                    lRes = CryptGetHashParam(hHash, HP_HASHVAL, abHash(0), lLen, 0)
                    
                    If lRes <> 0 Then
                        For lIdx = 0 To UBound(abHash)
                            HashFile = HashFile & Right$("0" & Hex$(abHash(lIdx)), 2)
                        Next
                    End If
                End If
            End If
            
            CryptDestroyHash hHash
       End If
    End If
    
    f.FClose
    
    CryptReleaseContext hCtx, 0
End Function

Public Function HashString(ByVal Str As String, Optional ByVal Algorithm As Long = SHA1) As String
    Dim hCtx As Long, hHash As Long, lRes As Long, lLen As Long, lIdx As Long, abData() As Byte

    lRes = CryptAcquireContext(hCtx, vbNullString, vbNullString, IIF(Algorithm >= SHA256, PROV_RSA_AES, PROV_RSA_FULL), CRYPT_VERIFYCONTEXT)
    
    If lRes <> 0 Then
       lRes = CryptCreateHash(hCtx, Algorithm, 0, 0, hHash)
    
       If lRes <> 0 Then
          lRes = CryptHashData(hHash, ByVal Str, Len(Str), 0)
          
          If lRes <> 0 Then
             lRes = CryptGetHashParam(hHash, HP_HASHSIZE, lLen, 4, 0)
    
             If lRes <> 0 Then
                 ReDim abData(0 To lLen - 1)
    
                 lRes = CryptGetHashParam(hHash, HP_HASHVAL, abData(0), lLen, 0)
    
                 If lRes <> 0 Then
                     For lIdx = 0 To UBound(abData)
                         HashString = HashString & Right$("0" & Hex$(abData(lIdx)), 2)
                     Next
                 End If
             End If
          End If
    
          CryptDestroyHash hHash
    
       End If
    End If
    
    CryptReleaseContext hCtx, 0
End Function
