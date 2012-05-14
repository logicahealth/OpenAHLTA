VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "iSQLFix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function GetOuterJoin(ByVal sInnerColumnName As String, ByVal sOuterColumnName As String) As String
End Function

Public Function GetSubStr(ByVal sColumn As String, ByVal nStart As Long, ByVal nLength As Long) As String
End Function

Public Function GetSubStrOuterJoin(ByVal sInnerColumnName As String _
                                  , ByVal sOuterColumnName As String _
                                  , ByVal nStart As Long _
                                  , ByVal nLength As Long) As String
End Function
