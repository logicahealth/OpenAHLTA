VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPickListItemFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private sNCID As String
Private sRepresentation As String
Private sRSFORM_ID As String
Public Property Get NCID() As String
    NCID = sNCID
End Property
Public Property Let NCID(ByVal sNCIDValue As String)
    sNCID = sNCIDValue
End Property
Public Property Get Representation() As String
    Representation = sRepresentation
End Property
Public Property Let Representation(ByVal sRepresentationValue As String)
    sRepresentation = sRepresentationValue
End Property
Public Property Get RelsFormNumId() As String
    RelsFormNumId = sRSFORM_ID
End Property
Public Property Let RelsFormNumId(ByVal sRSFORMValue As String)
    sRSFORM_ID = sRSFORMValue
End Property

Public Property Get Concept() As String
    Concept = sNCID
End Property

