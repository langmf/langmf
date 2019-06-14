VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public This As Long, Hash As New clsHash


Public Property Get Parent() As Object
    Set Parent = ObjFromPtr(This)
End Property


Public Function COM_Invoke(ByVal pThis As Long, ByVal idMember As Long, ByVal IDName As String, ByVal wFlags As Long, Arg() As Variant, Result As Variant) As Long              'no_check
    Dim uds As Long, v As Variant
    
    This = pThis
    
    uds = UBound(Arg)
    
    If idMember = 0 Then
        If uds = -1 Then
            Set Result = ObjFromPtr(pThis)
            Exit Function
        Else
            IDName = Arg(uds)
            uds = uds - 1
        End If
    ElseIf idMember = -1 And LenB(IDName) = 0 Then
        Set Result = Me
        Exit Function
    End If

    Select Case wFlags
        Case VbMethod To VbFunc
            If uds < 0 Then
                If IsObject(Hash(IDName)) Then Set Result = Hash(IDName) Else Result = Hash(IDName)
            Else
                If IsObject(Hash(IDName)) Then
                    CBN Hash(IDName), 0, VbFunc, Arg, -2, VarPtr(Result)
                Else
                    v = Hash(IDName)
                    If uds = 0 Then
                        If IsObject(v(Arg(0))) Then Set Result = v(Arg(0)) Else Result = v(Arg(0))
                    ElseIf uds = 1 Then
                        If IsObject(v(Arg(1), Arg(0))) Then Set Result = v(Arg(1), Arg(0)) Else Result = v(Arg(1), Arg(0))
                    Else
                        If IsObject(v(Arg(2), Arg(1), Arg(0))) Then Set Result = v(Arg(2), Arg(1), Arg(0)) Else Result = v(Arg(2), Arg(1), Arg(0))
                    End If
                End If
            End If
            
        Case VbLet
            Hash(IDName) = Arg(0)
    
        Case VbSet, VbLet + VbSet
            Set Hash(IDName) = Arg(0)
    End Select
End Function


Public Function Add(ParamArray vsp() As Variant) As IObject
    Dim a As Long, uds As Long

    Set Add = Me

    uds = UBound(vsp):      If uds < 1 Then Exit Function
    
    For a = 0 To uds Step 2
        If Not IsObject(vsp(a)) Then If IsMissing(vsp(a)) Then vsp(a) = Empty
        Hash.Add vsp(a + 1), vsp(a)
    Next
End Function
