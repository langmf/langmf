VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsActiveScript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IActiveScriptSite
Implements IActiveScriptSiteWindow

Private IScript As IActiveScript
Private IParse As IActiveScriptParse
Private Exep As ATL.EXCEPINFO

Public Parent As Object, Objects As New Collection, Error As New Collection
Public Name As String, Tag As String, Language As String, SitehWnd As Long


Private Function Init() As Long
    Dim oUnk As IUnknown, CLSID As UUID, IID As UUID

    If Not IScript Is Nothing Then Exit Function

    CLSIDFromString StrPtr(IID_IActiveScript), IID
    
    If Len(Language) = 0 Then Language = "VBScript"
    
    If Left$(Language, 1) = "{" Then
        CLSIDFromString StrPtr(Language), CLSID
    ElseIf LCase$(Language) = "chakra" Then
        CLSIDFromString StrPtr("{16D51579-A30B-4C8B-A276-0FF4DC41E755}"), CLSID
    Else
        CLSIDFromProgID StrPtr(Language), CLSID
    End If

    Init = CoCreateInstance(CLSID, 0&, CLSCTX_INPROC_SERVER, IID, ByVal VarPtr(oUnk))

    If Init = S_OK Then
        Set IScript = oUnk
        IScript.SetScriptSite Me
        
        Set IParse = IScript
        IParse.InitNew
        
        IScript.SetScriptState SCRIPTSTATE_CONNECTED
    End If
End Function

Private Sub Class_Terminate()
    Call Reset
End Sub

Public Property Get Ptr() As Long
    Ptr = ObjPtr(IScript)
End Property

Public Function Reset() As Object
    Dim v As Variant
    
    Set Reset = Me
    
    For Each v In Objects
        If IsObject(v) Then Set v = Nothing
    Next
    Set Objects = New Collection

    If Not IScript Is Nothing Then
        IScript.SetScriptState SCRIPTSTATE_DISCONNECTED
        IScript.Close
    End If
    
    Set IParse = Nothing
    Set IScript = Nothing
End Function

Public Function Parse(value As String, Optional ByVal nameMod As String, Optional ByVal Flags As Long, Optional Result As Variant) As Object
    Dim cnx As Long
    Set Parse = Me:           Call Init
    On Error Resume Next
    cnx = Objects(nameMod)
    On Error GoTo 0
    IParse.ParseScriptText value, nameMod, Nothing, vbNullString, cnx, 0, Flags, Result, Exep
End Function

Public Function AddCode(ByVal value As String, Optional ByVal nameMod As String) As Object
    Set AddCode = Parse(value, nameMod, SCRIPTTEXT_ISVISIBLE)
End Function

Public Function Eval(ByVal value As String, Optional ByVal nameMod As String) As Variant
    Parse value, nameMod, SCRIPTTEXT_ISEXPRESSION, Eval
End Function

Public Function Execute(ByVal value As String, Optional ByVal nameMod As String) As Object
    Set Execute = Parse(value, nameMod)
End Function

Public Property Get CodeObject(Optional ByVal nameMod As String) As Object
    On Error Resume Next
    Init
    IScript.GetScriptDispatch nameMod, CodeObject
End Property

Public Function AddObject(ByVal vName As String, ByVal Obj As Object, Optional ByVal AddMembers As Boolean) As Object
    On Error GoTo err1
    Set AddObject = Me:       Call Init
    If Not Obj Is Nothing Then
        Objects.Add Obj, vName
        AddName vName, SCRIPTITEM_ISVISIBLE Or IIF(AddMembers, SCRIPTITEM_GLOBALMEMBERS, 0)
    End If
err1:
End Function

Public Function AddName(ByVal value As String, Optional ByVal Flags As Long = SCRIPTITEM_ISVISIBLE Or SCRIPTITEM_CODEONLY) As Object
    Set AddName = Me:         Call Init
    IScript.AddNamedItem value, Flags
End Function

Public Function AddModule(ByVal nameMod As Variant, Optional ByVal txtCode As String) As Object
    On Error GoTo err1
    Init
    If IsNumeric(nameMod) Then nameMod = CStr(nameMod):    Objects.Add Val(nameMod), nameMod
    AddName nameMod
    AddCode txtCode, nameMod
    Set AddModule = CodeObject(nameMod)
err1:
End Function

Public Property Get State() As Long
    Init
    IScript.GetScriptState State
End Property

Public Property Let State(ByVal value As Long)
    Init
    IScript.SetScriptState value
End Property

Public Function Run(ByVal vName As String, ParamArray a() As Variant) As Variant
    Select Case UBound(a)
        Case -1:  Run = CallByName(CodeObject, vName, VbMethod)
        Case 0:   Run = CallByName(CodeObject, vName, VbMethod, a(0))
        Case 1:   Run = CallByName(CodeObject, vName, VbMethod, a(0), a(1))
        Case 2:   Run = CallByName(CodeObject, vName, VbMethod, a(0), a(1), a(2))
        Case 3:   Run = CallByName(CodeObject, vName, VbMethod, a(0), a(1), a(2), a(3))
        Case 4:   Run = CallByName(CodeObject, vName, VbMethod, a(0), a(1), a(2), a(3), a(4))
        Case 5:   Run = CallByName(CodeObject, vName, VbMethod, a(0), a(1), a(2), a(3), a(4), a(5))
        Case 6:   Run = CallByName(CodeObject, vName, VbMethod, a(0), a(1), a(2), a(3), a(4), a(5), a(6))
        Case 7:   Run = CallByName(CodeObject, vName, VbMethod, a(0), a(1), a(2), a(3), a(4), a(5), a(6), a(7))
        Case 8:   Run = CallByName(CodeObject, vName, VbMethod, a(0), a(1), a(2), a(3), a(4), a(5), a(6), a(7), a(8))
        Case 9:   Run = CallByName(CodeObject, vName, VbMethod, a(0), a(1), a(2), a(3), a(4), a(5), a(6), a(7), a(8), a(9))
    End Select
End Function



'---------------------------------------------------------------------
' IActiveScript Implementation
'---------------------------------------------------------------------

Private Sub IActiveScriptSite_GetItemInfo(ByVal pstrName As String, ByVal dwReturnMask As SCRIPTINFO, ppiunkItem As Long, ppti As Long)
    Dim Obj As ATL.IUnknown
    
    On Error Resume Next
    
    If dwReturnMask = SCRIPTINFO_IUNKNOWN Then
        Set Obj = Objects(pstrName)
        If Obj Is Nothing Then Set Obj = CodeObject(pstrName)
        ppiunkItem = ObjPtr(Obj)
        Obj.AddRef
    End If
End Sub

Private Sub IActiveScriptSite_OnScriptError(ByVal pscripterror As IActiveScriptError)
    Dim tmp As String, cnx As Long, lineNo As Long, charPos As Long

    On Error Resume Next

    With pscripterror
        .GetExceptionInfo Exep:    .GetSourceLineText tmp:   .GetSourcePosition cnx, lineNo, charPos
    End With

    Set Error = New Collection
    With Error
        .Add Exep.Source, "Source":     .Add Exep.sCode, "Number":     .Add Exep.Description, "Descr"
        .Add lineNo, "Line":    .Add charPos, "Pos":    .Add tmp, "Code":    .Add cnx, "Context"
    End With

    Parent_Event "ActiveScript_Error"
End Sub

Private Sub IActiveScriptSite_OnEnterScript()
    Parent_Event "ActiveScript_Enter"
End Sub

Private Sub IActiveScriptSite_OnLeaveScript()
    Parent_Event "ActiveScript_Leave"
End Sub

Private Sub IActiveScriptSite_OnStateChange(ByVal ssScriptState As SCRIPTSTATE)
    Parent_Event "ActiveScript_State"
End Sub

Private Function IActiveScriptSite_GetLCID() As Long
    IActiveScriptSite_GetLCID = GetSystemDefaultLangID
End Function

Private Function IActiveScriptSite_GetDocVersionString() As String
    IActiveScriptSite_GetDocVersionString = "LangMF Host"
End Function

Private Sub IActiveScriptSite_OnScriptTerminate(pvarResult As Variant, pexcepinfo As ATL.EXCEPINFO)
'
End Sub

Private Sub IActiveScriptSiteWindow_EnableModeless(ByVal fEnable As Long)
'
End Sub

Private Function IActiveScriptSiteWindow_GetWindow() As Long
    IActiveScriptSiteWindow_GetWindow = SitehWnd
End Function


'---------------------------------------------------------------------
Private Function Parent_Event(ProcName As String) As Variant
    Dim v As Variant
    On Error Resume Next
    If ExistsMember(Parent, ProcName) Then Set v = Me:   Parent_Event = CallByName(Parent, ProcName, VbMethod, v)
End Function
