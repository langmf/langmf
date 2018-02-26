Attribute VB_Name = "modInterface"
Option Explicit

Global Const COM_Mask_QueryInterface As Long = 1
Global Const COM_Mask_GetIDsOfNames As Long = 2
Global Const COM_Mask_Terminate As Long = &H100&

Type COMTable
    pVTable As Long
    cRefs As Long
    Mask As Long
    Wrapper As Object
    iName As Variant
    IID_User As UUID
    VTable(19) As Long
End Type


Private Function COM_QueryInterface(This As COMTable, riid As UUID, pvObj As Long) As Long
    Dim isOK As Boolean
    
    With This
        If (.Mask And COM_Mask_QueryInterface) Then
            COM_QueryInterface = .Wrapper.COM_QueryInterface(VarPtr(This), VarPtr(riid), VarPtr(pvObj))
        Else
            Select Case riid.Data1
                Case IID_IDispatch.Data1:       isOK = IsEqualGUID(riid, IID_IDispatch)
                Case IID_IUnknown.Data1:        isOK = IsEqualGUID(riid, IID_IUnknown)
                Case .IID_User.Data1:           isOK = IsEqualGUID(riid, .IID_User)
            End Select
            
            If isOK Then
                pvObj = VarPtr(.pVTable)
                .cRefs = .cRefs + 1
                COM_QueryInterface = S_OK
            Else
                COM_QueryInterface = E_NOINTERFACE
            End If
        End If
    End With
End Function

Private Function COM_AddRef(This As COMTable) As Long
    With This
        .cRefs = .cRefs + 1
        COM_AddRef = .cRefs
    End With
End Function

Private Function COM_Release(This As COMTable) As Long
    With This
        .cRefs = .cRefs - 1
        COM_Release = .cRefs
        If .cRefs = 0 Then
            If (.Mask And COM_Mask_Terminate) Then Call .Wrapper.COM_Terminate(VarPtr(This))
            .iName = Empty
            Set .Wrapper = Nothing
            CoTaskMemFree VarPtr(.pVTable)
        End If
    End With
End Function

Private Function COM_GetTypeInfoCount(This As COMTable, pctinfo As Long) As Long
    COM_GetTypeInfoCount = E_NOTIMPL
End Function

Private Function COM_GetTypeInfo(This As COMTable, ByVal iTInfo As Long, ByVal LCID As Long, ppTInfo As Long) As Long
    COM_GetTypeInfo = E_NOTIMPL
End Function

Private Function COM_GetIDsOfNames(This As COMTable, riid As UUID, rgszNames As Long, ByVal cNames As Long, ByVal LCID As Long, rgDispId As Long) As Long
    Dim sName As String, sz As Long
    
    With This
        sz = lstrlenW(ByVal rgszNames)
        If sz Then
            sName = String$(sz, 0)
            CopyMemory ByVal StrPtr(sName), ByVal rgszNames, sz * 2
        End If
        .iName = sName
        
        If (.Mask And COM_Mask_GetIDsOfNames) Then
            COM_GetIDsOfNames = .Wrapper.COM_GetIDsOfNames(VarPtr(This), VarPtr(riid), rgszNames, cNames, LCID, VarPtr(rgDispId))
        Else
            rgDispId = DISPID_UNKNOWN
            COM_GetIDsOfNames = S_OK
        End If
    End With
End Function

Private Function COM_Invoke(This As COMTable, ByVal idMember As Long, riid As UUID, ByVal LCID As Long, ByVal wFlags As InvokeFlags, pDispParams As ATL.DISPPARAMS, ByVal pvarResult As Long, pexcepinfo As ATL.EXCEPINFO, puArgErr As Long) As Long
    Dim SA As SafeArray, Arg() As Variant, Result As Variant
    
    On Error Resume Next
    
    With This
        With SA
           .cDims = 1
           .cbElements = 16
           .fFeatures = 128
           .pvData = pDispParams.rgPointerToVariantArray
           .rgSABound(0).cElements = pDispParams.cArgs
        End With
        
        wFlags = wFlags And 15
    
        PutMem4 VarPtrArray(Arg), VarPtr(SA)
        COM_Invoke = .Wrapper.COM_Invoke(VarPtr(This), idMember, .iName, wFlags, Arg, Result)
        If pvarResult <> 0 Then VariantCopyInd pvarResult, VarPtr(Result)
        PutMem4 VarPtrArray(Arg), 0&
    End With
End Function


Function Create_Interface(Wrapper As Object) As stdole.IUnknown
    Dim Ptr As Long, Cts As COMTable
    
    If Wrapper Is Nothing Then Exit Function
    
    Ptr = CoTaskMemAlloc(LenB(Cts))
    If Ptr = 0 Then Exit Function
        
    With Cts
        Set .Wrapper = Wrapper
        
        .VTable(0) = AddrOf(AddressOf COM_QueryInterface)
        .VTable(1) = AddrOf(AddressOf COM_AddRef)
        .VTable(2) = AddrOf(AddressOf COM_Release)
        .VTable(3) = AddrOf(AddressOf COM_GetTypeInfoCount)
        .VTable(4) = AddrOf(AddressOf COM_GetTypeInfo)
        .VTable(5) = AddrOf(AddressOf COM_GetIDsOfNames)
        .VTable(6) = AddrOf(AddressOf COM_Invoke)
        
        .pVTable = (Ptr Xor &H80000000) + 48 Xor &H80000000
        .cRefs = 1
        
        COM_Custom Cts
                
        CopyMemory ByVal Ptr, Cts, LenB(Cts)
        CopyMemory Create_Interface, Ptr, 4
        
        ZeroMemory Cts, LenB(Cts)
    End With
End Function


Private Sub COM_Custom(This As COMTable)
    Dim a As Long, uds As Long, mbr() As Variant, Param As Variant
    
    On Error GoTo err1
    
    With This
        Param = .Wrapper.COM_Custom(VarPtr(This))
        
        uds = -1:      uds = UBound(Param)
        
        If uds >= 2 Then
            mbr = Param(2)
            If VerifyArrayRange(mbr, , , , 17) Then
                For a = 0 To UBound(mbr)
                    If IsNumeric(mbr(a)) Then .VTable(a + 3) = CLng(mbr(a))
                Next
            End If
        End If
            
        If uds >= 1 Then
            If VarType(Param(1)) = vbString Then
                If Len(Param(1)) Then CLSIDFromString StrPtr(Param(1)), .IID_User
            Else
                CopyMemory .IID_User, ByVal CLng(Param(1)), 16
            End If
        End If
            
        If uds >= 0 Then .Mask = Param(0)
    End With
err1:
End Sub

Private Function IsEqualGUID(i1 As UUID, i2 As UUID) As Boolean
    Dim Tmp1 As Currency, Tmp2 As Currency
    If i1.Data1 <> i2.Data1 Then Exit Function
    If i1.Data2 <> i2.Data2 Then Exit Function
    If i1.Data3 <> i2.Data3 Then Exit Function
    CopyMemory Tmp1, i1.Data4(0), 8
    CopyMemory Tmp2, i2.Data4(0), 8
    If Tmp1 = Tmp2 Then IsEqualGUID = True
End Function

Private Function AddrOf(ByVal value As Long) As Long
    AddrOf = value
End Function
