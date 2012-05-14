VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MapperICD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements DL_Interfaces.iMapperOrderEntry

Private Const MODULE_NAME As String = "DM_OrderEntry.MapperICD"

Public Sub iMapperOrderEntry_Init(ByRef Support As iParameters, ByRef Cache As CacheMapper)
    On Error GoTo ErrHandler
    
    Set mCacheMapper = Cache
    Set mConn = Support.Item("DBConn").Value
    
    Exit Sub

ErrHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function iMapperOrderEntry_RetrieveOtherOrderTypes(ByRef params As iParameters) As OEOtherOrderTypes
    On Error GoTo ErrHandler
    
ErrHandler:
    Set iMapperOrderEntry_RetrieveOtherOrderTypes = Nothing
    Err.Raise ERR_METHOD_NOT_SUPPORTED, MODULE_NAME & ".RetrieveOtherOrderTypeCategories", ERR_METHOD_NOT_SUPPORTED_STR
End Function

Public Function iMapperOrderEntry_RetrieveOtherOrderTypeCategories(ByVal FacilityNCID As String) As Collection
    'return a collection of strings containing the Categories found in the OtherOrderTypes table for
    '   a given facility
    On Error GoTo ErrHandler
    
ErrHandler:
    Set iMapperOrderEntry_RetrieveOtherOrderTypeCategories = Nothing
    Err.Raise ERR_METHOD_NOT_SUPPORTED, MODULE_NAME & ".RetrieveOtherOrderTypeCategories", ERR_METHOD_NOT_SUPPORTED_STR
End Function

Public Function iMapperOrderEntry_SaveOtherOrderTypes(ByRef OrderTypes As OEOtherOrderTypes) As Boolean
    On Error GoTo ErrHandler
    
ErrHandler:
    iMapperOrderEntry_SaveOtherOrderTypes = False
    Err.Raise ERR_METHOD_NOT_SUPPORTED, MODULE_NAME & ".RetrieveOtherOrderTypeCategories", ERR_METHOD_NOT_SUPPORTED_STR
End Function


