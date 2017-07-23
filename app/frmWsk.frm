VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmWsk 
   BorderStyle     =   0  'None
   Caption         =   "frmWsk"
   ClientHeight    =   600
   ClientLeft      =   0
   ClientTop       =   -2004
   ClientWidth     =   636
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   HasDC           =   0   'False
   ScaleHeight     =   600
   ScaleWidth      =   636
   ShowInTaskbar   =   0   'False
   Begin MSWinsockLib.Winsock Wsk 
      Index           =   0
      Left            =   120
      Top             =   120
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   393216
   End
End
Attribute VB_Name = "frmWsk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Parent As frmForm


Private Sub Wsk_Close(Index As Integer)
    If Parent.Events("Wsk" & Index & "_Close") = False Then Parent.Events "Wsk_Close", Index
End Sub

Private Sub Wsk_Connect(Index As Integer)
    If Parent.Events("Wsk" & Index & "_Connect") = False Then Parent.Events "Wsk_Connect", Index
End Sub

Private Sub Wsk_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    If Parent.Events("Wsk" & Index & "_ConnectionRequest", requestID) = False Then Parent.Events "Wsk_ConnectionRequest", Index, requestID
End Sub

Private Sub Wsk_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    If Parent.Events("Wsk" & Index & "_DataArrival", bytesTotal) = False Then Parent.Events "Wsk_DataArrival", Index, bytesTotal
End Sub

Private Sub Wsk_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    If Parent.Events("Wsk" & Index & "_Error", Number, Description, sCode) = False Then Parent.Events "Wsk_Error", Index, Number, Description, sCode
End Sub

Private Sub Wsk_SendComplete(Index As Integer)
    If Parent.Events("Wsk" & Index & "_SendComplete") = False Then Parent.Events "Wsk_SendComplete", Index
End Sub

Private Sub Wsk_SendProgress(Index As Integer, ByVal bytesSent As Long, ByVal bytesRemaining As Long)
    If Parent.Events("Wsk" & Index & "_SendProgress", bytesSent, bytesRemaining) = False Then Parent.Events "Wsk_SendProgress", Index, bytesSent, bytesRemaining
End Sub
