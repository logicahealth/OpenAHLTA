'***************************************************************************************
'
'  Copyright (c) 2007-2012 Northrop Grumman Corporation
'
'  Licensed by Tricare Management Activity under license from the Copyright owner.
'
'  This text file must be included in all Derivative Works of the licensed Source Code.
'
'***************************************************************************************

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

Private Function iSQLFix_GetSubStr(ByVal sColumn As String, ByVal nStart As Long, ByVal nLength As Long) As String
  iSQLFix_GetSubStr = " SubStr(" & sColumn & ", " & nStart & ", " & nLength & ")"
End Function

Private Function iSQLFix_GetSubStrOuterJoin(ByVal sInnerColumnName As String _
                                  , ByVal sOuterColumnName As String _
                                  , ByVal nStart As Long _
                                  , ByVal nLength As Long) As String
                                  
  iSQLFix_GetSubStrOuterJoin = " " & sInnerColumnName & " = " & iSQLFix_GetSubStr(sOuterColumnName & " (+)", nStart, nLength)
End Function


