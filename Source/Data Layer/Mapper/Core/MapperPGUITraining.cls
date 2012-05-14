VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MapperPGUITraining"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Provides all Data Access for Core modules against the CDR.

Option Explicit

Implements iMapperCore


Private mConn As CHCSII_CONN.Conn
Private mCache As CacheMapper

Private Sub iMapperCore_Init(Support As DL_Support.iParameters, Cache As DL_Support.CacheMapper)
  Set mConn = Support.Item("DBConn").Value
  Set mCache = Cache
End Sub

'Gets the abbr rank description for a given Rank NCID.
Private Function iMapperCore_RetrieveAbbrRank(ByVal RankNCID As String) As String
Dim SQLOps As ICHCSII_SQL.ISqlOps
Dim SQL As String

Const RANK_SQL As String = "SELECT abbrev_rank from rank_abbrev WHERE ncid = '%NCID'"
Const FLD_AB_RANK As String = "abbrev_rank"

  SQL = Replace(RANK_SQL, "%NCID", RankNCID)
  
  Set SQLOps = LoadSQLOps(SQL)
  
  If Not SQLOps.EOF Then
    iMapperCore_RetrieveAbbrRank = SQLOps.Value(FLD_AB_RANK)
  End If
  
End Function

Private Function LoadSQLOps(ByVal SelectStmt As String) As ISqlOps
Dim SQLOps As ICHCSII_SQL.ISqlOps

  Set SQLOps = mConn.CHCSII_SQLOPS(CDR)
  SQLOps.Execute SelectStmt

  Set LoadSQLOps = SQLOps
  
End Function

