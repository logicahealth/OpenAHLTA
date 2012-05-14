VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MapperPGUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements DL_Interfaces.iMapperDocuments

Private Const MODULE_NAME As String = "DM_Documents.MapperPGUI"

Private Sub iMapperDocuments_Init(Support As DL_Support.iParameters, Cache As DL_Support.CacheMapper)
    On Error GoTo ErrHandler
    
    Set mCacheMapper = Cache
    Set mConn = Support.Item("DBConn").Value
    
    Exit Sub

ErrHandler:
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Private Function iMapperDocuments_RetrieveInpatientAdmissions(PatientUnitNumber As String, FacilityNCID As String, Optional ApptID As String = "") As DTOs.DocAdmissions
    On Error GoTo ErrHandler
    
ErrHandler:
    Set iMapperDocuments_RetrieveInpatientAdmissions = Nothing
    Err.Raise ERR_METHOD_NOT_SUPPORTED, MODULE_NAME & ".RetrieveInpatientAdmissions", ERR_METHOD_NOT_SUPPORTED_STR

End Function
