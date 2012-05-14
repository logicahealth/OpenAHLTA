VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SQLSQLServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements iSQLFix

Private Function iSQLFix_GetOuterJoin(ByVal sInnerColumnName As String, ByVal sOuterColumnName As String) As String
    If IsNumeric(sInnerColumnName) Or IsNumeric(sOuterColumnName) Then
        iSQLFix_GetOuterJoin = " " & sInnerColumnName & " = " & sOuterColumnName & " "
    Else
        iSQLFix_GetOuterJoin = " " & sInnerColumnName & " *= " & sOuterColumnName & " "
    End If
End Function


Private Function iSQLFix_SQLDate(ByVal dDate As Date) As String
    If CDbl(dDate) = 0 Then
        iSQLFix_SQLDate = "NULL"
    Else
        '---SQL server style 13 is dd mon yyyy hh:mm:ss:mmm(24h)
        iSQLFix_SQLDate = "convert(datetime,'" & Format$(dDate, "dd/mmm/yyyy hh:nn:ss") & "', 13)"
    End If
End Function


