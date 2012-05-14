VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MapperTheater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements DL_Interfaces.iMapperRads

Private Connection As CHCSII_CONN.Conn
Private SQLOps As ICHCSII_DAS


Private Sub iMapperRads_Init(Support As DL_Support.iParameters, ByRef Cache As CacheMapper)

  Set Connection = New CHCSII_CONN.Conn
  Set SQLOps = Connection.CHCSII_DAS(Auto)

End Sub

Private Function iMapperRads_RetrieveRads(NPParams As DL_Support.iParameters) As DTOs.RadClinicalNotes
  
  Dim MySQLGeneric As MapperSQLGeneric
  Set MySQLGeneric = New MapperSQLGeneric

  Set iMapperRads_RetrieveRads = MySQLGeneric.SQLGeneric(SQLOps, NPParams)

  Set MySQLGeneric = Nothing

End Function


Private Function iMapperRads_RetrieveRadTypes(NPParams As DL_Support.iParameters) As DTOs.RadTypes
  
  Err.Raise 0, "DM_Radiology.MapperTheater", "ImapperRads_RetrieveRadTypes is not valid in Theater mode.", Err.HelpFile, Err.HelpContext

End Function

