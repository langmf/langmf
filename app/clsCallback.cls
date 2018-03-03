VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCallback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const CODE_STR As String = "5589E583EC08E8070000008B45FCC9C2_xx_8D45FC50FF752CFF7528FF7524FF7520FF751CFF7518FF7514FF7510FF750CFF7508B8_xxxxxx_508B00FF501CC3"

Private hMem As Long, Obj As Object, Method As Variant, Count As Long


'THIS PROCEDURE SHOULD BE FIRST IN CLASS - don't move it anywhere
Public Function Proc(ByVal p1 As Long, ByVal p2 As Long, ByVal p3 As Long, ByVal p4 As Long, ByVal p5 As Long, ByVal p6 As Long, ByVal p7 As Long, ByVal p8 As Long, ByVal p9 As Long, ByVal p10 As Long) As Long
    Dim Args() As Variant
    
    Select Case Count
        Case 1:   Args = Array(p1)
        Case 2:   Args = Array(p2, p1)
        Case 3:   Args = Array(p3, p2, p1)
        Case 4:   Args = Array(p4, p3, p2, p1)
        Case 5:   Args = Array(p5, p4, p3, p2, p1)
        Case 6:   Args = Array(p6, p5, p4, p3, p2, p1)
        Case 7:   Args = Array(p7, p6, p5, p4, p3, p2, p1)
        Case 8:   Args = Array(p8, p7, p6, p5, p4, p3, p2, p1)
        Case 9:   Args = Array(p9, p8, p7, p6, p5, p4, p3, p2, p1)
        Case 10:  Args = Array(p10, p9, p8, p7, p6, p5, p4, p3, p2, p1)
    End Select
    
    Proc = CBN(Obj, Method, VbFunc, Args, Count)
End Function

'Default Member
Public Function Ptr() As Long
Attribute Ptr.VB_UserMemId = 0
    Ptr = hMem
End Function

Public Sub Init(vObj As Variant, Optional ByVal vMethod As Variant = 0, Optional ByVal vCount As Long = 0, Optional ByVal isCDecl As Boolean)
    Dim i As Long, nLen As Long, sCode As String
    
    Call Class_Terminate
    
    nLen = Len(CODE_STR) / 2
    For i = 0 To nLen - 1
      sCode = sCode & ChrB$(Val("&H" & Mid$(CODE_STR, i * 2 + 1, 2)))
    Next
    
    hMem = CoTaskMemAlloc(nLen)
    Call VirtualProtect(hMem, nLen, PAGE_EXECUTE_READWRITE, 0&)
    Call CopyMemory(ByVal hMem, ByVal StrPtr(sCode), nLen)
    PutMem4 hMem + 53, ObjPtr(Me)
    If Not isCDecl Then PutMem2 hMem + 16, vCount * 4
    
    Set Obj = vObj
    Method = vMethod
    Count = vCount
End Sub


Private Sub Class_Terminate()
    If hMem Then CoTaskMemFree hMem
End Sub

