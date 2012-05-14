VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MapperITT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements DL_Interfaces.iMapperInpatientOrder

Private mConn As CHCSII_CONN.Conn
Private SQLOps As ICHCSII_DAS
Private iSQLOps As ICHCSII_SQL.ISqlOpsEx


Private Sub iMapperInpatientOrder_Init(Support As DL_Support.iParameters, Cache As DL_Support.CacheMapper)

End Sub

Private Function iMapperInpatientOrder_NewOrders(Params As DL_Support.iParameters) As DTOs.InPatientOrders

End Function

Private Function iMapperInpatientOrder_Retrieve(Params As DL_Support.iParameters) As DTOs.InPatientOrders

End Function

Private Function iMapperInpatientOrder_RetrieveInpatientID(Params As DL_Support.iParameters) As String

End Function

Private Function iMapperInpatientOrder_Save(Ordes As DTOs.InPatientOrders) As Boolean

End Function
