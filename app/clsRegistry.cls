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

Private WSH As Object


Private Sub Class_Initialize()
    On Error Resume Next
    Set WSH = CreateObject("WScript.Shell")
End Sub

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

'Работа с регистром (ветвь VB and VBA)
Public Sub SaveSetting(ByVal regAppName As String, ByVal regSection As String, ByVal regKey As String, ByVal regSetting As String)
    VBA.SaveSetting regAppName, regSection, regKey, regSetting
End Sub

Public Function GetSetting(ByVal regAppName As String, ByVal regSection As String, ByVal regKey As String, Optional ByVal regDef As String) As String
    GetSetting = VBA.GetSetting(regAppName, regSection, regKey, regDef)
End Function

Public Sub DeleteSetting(ByVal regAppName As String, ByVal regSection As String, ByVal regKey As String)
    VBA.DeleteSetting regAppName, regSection, regKey
End Sub

