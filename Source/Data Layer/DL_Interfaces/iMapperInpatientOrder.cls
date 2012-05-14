VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "iMapperInpatientOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'This should only be called from the DataAccess library
Public Sub Init(ByRef Support As iParameters, ByRef Cache As CacheMapper)
End Sub

Public Function Retrieve(ByRef Params As iParameters) As DTOs.InPatientOrders
End Function

Public Function Save(ByRef Orders As DTOs.InPatientOrders) As Boolean
End Function

Public Function RetrieveInpatientID(ByRef Params As iParameters) As String
End Function

'--- SCR 88340  Sherry Wang  2/21/2006
Public Function NewOrders(ByRef Params As iParameters) As DTOs.InPatientOrders
End Function

