VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "iMapperRads"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'This should only be called from the DataAccess library
Public Sub Init(ByRef Support As iParameters, ByRef Cache As CacheMapper)
End Sub

'Gets Rads info
Public Function RetrieveRads(ByRef Params As iParameters) As DTOs.RadClinicalNotes
End Function

'Gets value/name pairs of RadTypes for display in dropdown lists.
Public Function RetrieveRadTypes(ByRef Params As iParameters) As DTOs.RadTypes
End Function

