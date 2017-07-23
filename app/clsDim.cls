VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


'-------------------- Skin --------------------
Public Slides As New Collection
Public Counter As Long
Public One As Boolean
Public Picture As StdPicture
Public MaskPicture As StdPicture
Public TransColor As Variant
Public Interval As Long


'------------------ Resource ------------------
Public Extra As Variant
Public Data As Variant
Public Mode As String
Public ID As String


