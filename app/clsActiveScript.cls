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

Public Sub Reset()
    Dim v As Variant
    
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
End Sub

Public Sub Parse(value As String, Optional ByVal nMod As String, Optional ByVal Flags As Long, Optional Result As Variant)
    Dim cnx As Long
    On Error Resume Next
    Init
    cnx = Objects(nMod)
    IParse.ParseScriptText value, nMod, Nothing, vbNullString, cnx, 0, Flags, Result, Exep
End Sub

Public Sub AddCode(ByVal value As String, Optional ByVal nMod As String)
    Parse value, nMod, SCRIPTTEXT_ISVISIBLE
End Sub

Public Function Eval(ByVal value As String, Optional ByVal nMod As String) As Variant
    Parse value, nMod, SCRIPTTEXT_ISEXPRESSION, Eval
End Function

Public Sub ExecuteStatement(ByVal value As String, Optional ByVal nMod As String)
    Parse value, nMod
End Sub

Public Property Get CodeObject(Optional ByVal nMod As String) As Object
    On Error Resume Next
    Init
    IScript.GetScriptDispatch nMod, CodeObject
End Property

Public Sub AddObject(ByVal sName As String, ByVal Obj As Object, Optional ByVal AddMembers As Boolean)
    On Error GoTo err1
    Init
    If Not Obj Is Nothing Then
        Objects.Add Obj, sName
        AddName sName, SCRIPTITEM_ISVISIBLE Or IIF(AddMembers, SCRIPTITEM_GLOBALMEMBERS, 0)
    End If
err1:
End Sub

Public Sub AddName(ByVal value As String, Optional ByVal Flags As Long = SCRIPTITEM_ISVISIBLE Or SCRIPTITEM_CODEONLY)
    Init
    IScript.AddNamedItem value, Flags
End Sub

Public Function AddModule(ByVal nMod As Variant, ByVal sCode As String) As Object
    On Error GoTo err1
    Init
    If IsNumeric(nMod) Then nMod = CStr(nMod):    Objects.Add Val(nMod), nMod
    AddName nMod
    AddCode sCode, nMod
    Set AddModule = CodeObject(nMod)
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
    IActiveScriptSite_GetDocVersionString = "LangMF Script Host"
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
