VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SQLOracle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements iSQLFix

Private Function iSQLFix_GetOuterJoin(ByVal sInnerColumnName As String, ByVal sOuterColumnName As String) As String
  iSQLFix_GetOuterJoin = " " & sInnerColumnName & " = " & sOuterColumnName & "(+) "
End Function


Private Function iSQLFix_SQLDate(ByVal dDate As Date) As String
End Function
