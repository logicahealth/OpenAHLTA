VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "iMapperOrderEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'This should only be called from the DataAccess library
Public Sub Init(ByRef Support As iParameters, ByRef Cache As CacheMapper)
End Sub

Public Function RetrieveOtherOrderTypes(ByRef Params As iParameters) As DTOs.OEOtherOrderTypes
End Function

Public Function RetrieveOtherOrderTypeCategories(ByVal FacilityNCID As String) As Collection
    'return a collection of strings containing the Categories found in the OtherOrderTypes table for
    '   a given facility
End Function

Public Function SaveOtherOrderTypes(ByRef OrderTypes As OEOtherOrderTypes) As Boolean
End Function
