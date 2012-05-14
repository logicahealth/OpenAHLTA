VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "POC_DEF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim sNCID As String
Dim sName As String


Public Property Get NCID() As String
    NCID = sNCID
End Property

Public Property Let NCID(ByVal RHV As String)
    sNCID = RHV
End Property

Public Property Get Name() As String
    Name = sName
End Property

Public Property Let Name(ByVal RHV As String)
    sName = RHV
End Property
