VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAddHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private IID_Event As UUID, oCP As IConnectionPoint, Sink As Object, dwCookie As Long, bAttach As Boolean

Public CallType As Long, Prefix As String, Object As ATL.IDispatch, Handler As ATL.IDispatch


Public Sub Init(Optional ByVal vObject As Object, Optional ByVal vPrefix As String, Optional ByVal vCallType As Long, Optional ByVal vHandler As Object)
    Prefix = vPrefix
    CallType = vCallType
    If Not vHandler Is Nothing Then Set Handler = vHandler
    If Not vObject Is Nothing Then Set Object = vObject:    Call Attach
End Sub

Public Function Attach() As Boolean
    Dim hr As Long, oCPC As IConnectionPointContainer, oECP As IEnumConnectionPoints

    If bAttach Then Detach

    Set oCPC = Object
    If oCPC Is Nothing Then Exit Function

    hr = oCPC.EnumConnectionPoints(oECP)
    If hr Then Exit Function

    hr = CallInterface(ObjPtr(oECP), 3, 3, 1, VarPtr(oCP))
    If hr Then Exit Function

    hr = oCP.GetConnectionInterface(IID_Event)
    If hr Then Exit Function
    
    Set Sink = Create_Interface(Me, Array(, VarPtr(IID_Event)))

    hr = oCP.Advise(Sink, dwCookie)
    If hr Then Exit Function

    Set oECP = Nothing
    Set oCPC = Nothing
    bAttach = True
    Attach = True
End Function

Public Sub Detach()
    If Not bAttach Then Exit Sub
    oCP.Unadvise dwCookie
    Set oCP = Nothing
    Set Sink = Nothing
    bAttach = False
End Sub

Public Function COM_Invoke(ByVal pThis As Long, ByVal idMember As Long, ByVal IDName As String, ByVal wFlags As Long, Params() As Variant, Result As Variant) As Long              'no_check
    IDName = GetMemberName(idMember)
    Select Case CallType
        Case 0
            Result = CBN(Handler, Prefix + IDName, wFlags, Params, -2)
        Case 1
            Result = CBN(Handler, Prefix + IDName, wFlags, Array(Params), -2)
        Case -1
            Result = CBN(Handler, Prefix, wFlags, Array(Params, IDName), -2)
    End Select
End Function

Private Function GetMemberName(ByVal dispid As Long) As String
    Dim strName As String, TL As ITypeLib, TI As ITypeInfo, TI2 As ITypeInfo

    If Object.GetTypeInfoCount Then
        Set TI = Object.GetTypeInfo
    
        If TI.GetContainingTypeLib(TL) Then
            Set TI2 = TL.GetTypeInfoOfIID(IID_Event)
            If Not TI2 Is Nothing Then If TI2.GetNames(dispid, strName, 1, 0&) = 0 Then GetMemberName = strName
        End If
        
        If LenB(strName) = 0 Then If TI.GetNames(dispid, strName, 1, 0&) = 0 Then GetMemberName = strName
    End If
    
    If LenB(strName) = 0 Then GetMemberName = dispid
End Function

Private Sub Class_Initialize()
    Set Handler = CAS.CodeObject
End Sub

Private Sub Class_Terminate()
    Call Detach
    Set Object = Nothing
    Set Handler = Nothing
End Sub
