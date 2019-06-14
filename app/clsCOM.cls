VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCOM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public Function Create(Optional ByVal Wrapper As Object, Optional Args As Variant) As Object
    Set Create = Create_Interface(Wrapper, Args)
End Function

Public Function QI(ByVal pIUnk As Variant, Optional ByVal vRIID As Variant, Optional pObj As Variant) As Long
    Dim nIfc As Long
    QI = ObjFromIUnk(pIUnk).QueryInterface(GetGuid(vRIID), nIfc)
    pObj = nIfc
End Function

Public Function AddRef(ByVal pIUnk As Variant) As Long
    AddRef = ObjFromIUnk(pIUnk).AddRef
End Function

Public Function Release(ByVal pIUnk As Variant) As Long
    Release = ObjFromIUnk(pIUnk).Release
End Function

Public Function FromGUID(ByVal ptrGUID As Long) As String
    FromGUID = String$(38, 0)
    StringFromGUID2 ptrGUID, StrPtr(FromGUID)
End Function

Public Function ToGUID(Optional ByVal vGUID As Variant) As Long
    Static v As UUID
    v = GetGuid(vGUID):      ToGUID = VarPtr(v)
End Function

Public Function Handler(Optional ByVal Object As Object, Optional ByVal Prefix As String, Optional ByVal CallType As Long, Optional ByVal objHandler As Object) As Object               'As clsAddHandler
    Set Handler = New clsAddHandler
    Handler.Init Object, Prefix, CallType, objHandler
End Function

Public Function Instance(ByVal txtCLSID As String, ByVal txtIID As String, Optional ByVal dwClsContext As Long = 1) As Long
    Dim CLSID As UUID, IID As UUID
    CLSIDFromString StrPtr(txtCLSID), CLSID
    CLSIDFromString StrPtr(txtIID), IID
    Call CoCreateInstance(CLSID, 0, dwClsContext, IID, Instance)
End Function

Public Function Enums(ByVal value As Variant, Optional ByVal Flags As Long = -1) As Object               'As clsHash
    Dim Disp As ATL.IDispatch, Name As String, Descr As String, CLSID As String, Names() As String, Params() As ELEMDESC
    Dim TI As ITypeInfo, TA As TYPEATTR, TL As ITypeLib, FD As FUNCDESC, VD As VARDESC, pdx As PARAMDESCEX
    Dim i As Long, j As Long, hr As Long, pTA As Long, pfd As Long, pVD As Long
    Dim Col As New clsHash, Elems As clsHash, Elem As clsHash, Prms As clsHash

    On Error GoTo err1

    If IsObject(value) Then
        Set Disp = value
        If Disp.GetTypeInfoCount = 0 Then Exit Function
        Set TI = Disp.GetTypeInfo

    ElseIf VarType(value) = vbString Then
        Set Enums = New clsHash
        If LoadTypeLib(StrPtr(CStr(value)), TL) = 0 Then
            For i = 0 To TL.GetTypeInfoCount - 1
                Set TI = TL.GetTypeInfo(i)
                Enums.Add Enums(ObjPtr(TI), Flags)
            Next
        End If
        Exit Function

    Else
        PutMem4 VarPtr(TI), value
        TI.AddRef
    End If

    TI.GetDocumentation DISPID_UNKNOWN, Name, Descr, 0, vbNullString

    pTA = TI.GetTypeAttr
    CopyMemory TA, ByVal pTA, LenB(TA)
    TI.ReleaseTypeAttr pTA
    
    CLSID = FromGUID(VarPtr(TA.IID))
    
    '-------------------------- Interface -------------------------
    With Col.Child("Interface")
        .Add Name, "Name"
        .Add Descr, "Descr"
        .Add CLSID, "Clsid"
        .Add TA.cbSizeInstance, "szInst"
        .Add TA.cbSizeVft, "szVft"
        .Add TA.cFuncs, "Funcs"
        .Add TA.cImplTypes, "Impls"
        .Add TA.cVars, "Vars"
        .Add TA.TYPEKIND, "Kind"
        .Add TA.wTypeFlags, "Flags"
        .Add TA.wMajorVerNum & "." & TA.wMinorVerNum, "Ver"
    End With
    
    '-------------------------- Functions -------------------------
    If (Flags And 1) Then
        Set Elems = Col.Child("Funcs")
        
        For i = 0 To TA.cFuncs - 1
            pfd = TI.GetFuncDesc(i)
            CopyMemory FD, ByVal pfd, LenB(FD)
            
            TI.GetDocumentation FD.memid, Name, Descr, 0, vbNullString
            
            ReDim Names(1 To FD.cParams + 1)
            
            hr = TI.GetNames(FD.memid, Names(1), UBound(Names), j)
            If j <> UBound(Names) Then Names(UBound(Names)) = "RHS"
            
            If Not Elems.Exists(Name) Then Set Elem = Elems.Child(Name) Else Set Elem = Elems.Child(Name & "@" & FD.invkind)
            
            With Elem
                .Add Name, "Name"
                .Add Descr, "Descr"
                .Add FD.cParams, "Prms"
                .Add FD.cParamsOpt, "Opts"
                .Add FD.invkind, "Invk"
                .Add FD.memid, "id"
                .Add FD.oVft, "oVft"
                .Add FD.FUNCKIND, "Kind"
                .Add FD.wFuncFlags, "Flags"
                .Add FD.elemdescFunc.tdesc.vt, "Type"
            End With
            
            If (Flags And 2) Then
                Set Prms = Elem.Child("Params")
                
                If FD.cParams > 0 Then
                    ReDim Params(1 To FD.cParams)
                    CopyMemory Params(1), ByVal FD.lprgELEMDESCParam, FD.cParams * LenB(Params(1))
        
                    For j = 1 To UBound(Params)
                        With Prms.Child(Names(j + 1))
                            .Add Names(j + 1), "Name"
                            .Add Params(j).PARAMDESC.wParamFlags, "Flags"
                            .Add Params(j).tdesc.vt, "Type"
                
                            If (Params(j).PARAMDESC.pPARAMDESCEX <> 0) And ((Flags And 8) > 0) Then
                                CopyMemory pdx, ByVal Params(j).PARAMDESC.pPARAMDESCEX, LenB(pdx)
                                .Add pdx.varDefaultValue, "Default"
                            End If
                        End With
                    Next
                End If
            End If
            
            TI.ReleaseFuncDesc pfd
        Next
    End If
    
    '----------------------------- Vars ----------------------------
    If (Flags And 16) Then
        Set Elems = Col.Child("Vars")
        
        For i = 0 To TA.cVars - 1
            pVD = TI.GetVarDesc(i)
            CopyMemory VD, ByVal pVD, LenB(VD)
            
            TI.GetDocumentation VD.memid, Name, Descr, 0, vbNullString
    
            With Elems.Child(Name)
                .Add Name, "Name"
                .Add Descr, "Descr"
                .Add VD.memid, "id"
                .Add VD.oInst_varValue, "oInst"
                .Add VD.VARKIND, "Kind"
                .Add VD.wVarFlags, "Flags"
                .Add VD.elemdescVar.tdesc.vt, "Type"
                
                If (VD.elemdescVar.PARAMDESC.pPARAMDESCEX <> 0) And ((Flags And 64) > 0) Then
                    CopyMemory pdx, ByVal VD.elemdescVar.PARAMDESC.pPARAMDESCEX, LenB(pdx)
                    .Add pdx.varDefaultValue, "Default"
                End If
            End With
            
            TI.ReleaseVarDesc pVD
        Next
    End If
    
err1:
    Set Enums = Col
End Function

Public Function Enums2(ByVal DispEx As ATL.IDispatchEx) As Object               'As clsHash
    Dim Col As New clsHash, txt As String, ids As Long, hr As Long
    
    On Error GoTo err1
    
    hr = DispEx.GetNextDispID(fdexEnumAll, DISPID_STARTENUM, ids)
       
    While hr = 0
        hr = DispEx.GetMemberName(ids, txt)
        If hr = 0 Then Col.Add ids, txt
        hr = DispEx.GetNextDispID(fdexEnumAll, ids, ids)
    Wend
    
err1:
    Set Enums2 = Col
End Function

Public Function DLL(ByVal objPath As String, Optional ByVal objClass As String) As Object
    Dim v As Variant, hr As Long, CLSID As UUID, INull As IUnknown, IUnk As IUnknown
    Dim f As New clsFuncPointer, ICF As IClassFactory
    
    If LenB(objClass) Then
        If f.Config(objPath, "DllGetClassObject") Then
            If Left$(objClass, 1) = "{" Then
                CLSIDFromString StrPtr(objClass), CLSID
            Else
                For Each v In Enums(objPath, 0).Items
                    If v("Interface")("Name") = objClass Then
                        CLSIDFromString StrPtr(v("Interface")("CLSID")), CLSID
                        Exit For
                    End If
                Next
            End If
            
            If f.PCall(VarPtr(CLSID), VarPtr(IID_IClassFactory), VarPtr(ICF)) = 0 Then
                If ICF.CreateInstance(INull, IID_IUnknown, IUnk) = 0 Then Set DLL = IUnk:    f.hLib = 0
                Set ICF = Nothing
                Set IUnk = Nothing
            End If
        End If
    Else
        hr = GetModuleHandleW(StrPtr(LongPath(objPath)))
        If hr Then
            f.hLib = hr
            If f.Config(, "DllCanUnloadNow") Then If f.PCall = 0 Then Exit Function
            f.hLib = 0
        End If
    End If
End Function
