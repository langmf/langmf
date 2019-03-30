VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOCX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents Object As VBControlExtender
Attribute Object.VB_VarHelpID = -1

Public Name As String, TypeEvent As Long


Public Function Create(Optional ByVal Frm As Object, Optional ByVal strName As String, Optional ByVal strGUID As String, Optional ByVal strEvent As String) As Object
    Name = strName:     TypeEvent = 0:      If LenB(strEvent) Then Name = strEvent:     TypeEvent = 1
    If LenB(strGUID) Then Set Object = Frm.Controls.Add(strGUID, strName)
    Set Create = Object
End Function

Private Sub Object_DragDrop(Source As Control, x As Single, y As Single)
    On Error Resume Next
    Object.Parent.Events Name & "_DragDrop", Source, x, y
End Sub

Private Sub Object_DragOver(Source As Control, x As Single, y As Single, State As Integer)
    On Error Resume Next
    Object.Parent.Events Name & "_DragOver", Source, x, y
End Sub

Private Sub Object_GotFocus()
    On Error Resume Next
    Object.Parent.Events Name & "_GotFocus"
End Sub

Private Sub Object_LostFocus()
    On Error Resume Next
    Object.Parent.Events Name & "_LostFocus"
End Sub

Private Sub Object_Validate(Cancel As Boolean)
    On Error Resume Next
    Cancel = Object.Parent.Events(Name & "_Validate")
End Sub

Private Sub Object_ObjectEvent(Info As EventInfo)
    On Error Resume Next
    
    If TypeEvent Then
        Object.Parent.Events Name & "_Event", Me, Info
    Else
        With Info.EventParameters
            Select Case .Count
                Case 0:  Object.Parent.Events Name & "_" & Info.Name
                Case 1:  Object.Parent.Events Name & "_" & Info.Name, .Item(0).value
                Case 2:  Object.Parent.Events Name & "_" & Info.Name, .Item(0).value, .Item(1).value
                Case 3:  Object.Parent.Events Name & "_" & Info.Name, .Item(0).value, .Item(1).value, .Item(2).value
                Case 4:  Object.Parent.Events Name & "_" & Info.Name, .Item(0).value, .Item(1).value, .Item(2).value, .Item(3).value
                Case 5:  Object.Parent.Events Name & "_" & Info.Name, .Item(0).value, .Item(1).value, .Item(2).value, .Item(3).value, .Item(4).value
                Case 6:  Object.Parent.Events Name & "_" & Info.Name, .Item(0).value, .Item(1).value, .Item(2).value, .Item(3).value, .Item(4).value, .Item(5).value
                Case 7:  Object.Parent.Events Name & "_" & Info.Name, .Item(0).value, .Item(1).value, .Item(2).value, .Item(3).value, .Item(4).value, .Item(5).value, .Item(6).value
                Case 8:  Object.Parent.Events Name & "_" & Info.Name, .Item(0).value, .Item(1).value, .Item(2).value, .Item(3).value, .Item(4).value, .Item(5).value, .Item(6).value, .Item(7).value
                Case 9:  Object.Parent.Events Name & "_" & Info.Name, .Item(0).value, .Item(1).value, .Item(2).value, .Item(3).value, .Item(4).value, .Item(5).value, .Item(6).value, .Item(7).value, .Item(8).value
            End Select
        End With
    End If
    Err.Clear
End Sub
