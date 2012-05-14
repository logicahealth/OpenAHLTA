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
Attribute VB_Name = "SQLText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
' Simple class to assist in SQL String construction
' Ben Rendleman 12/01
Private Const msSELECT  As String = "SELECT "
Private Const msFROM    As String = "FROM "
Private Const msWHERE   As String = "WHERE "
Private Const msORDER   As String = "ORDER BY "
Private sSelect  As String
Private sFrom    As String
Private sWhere   As String
Private sOrder   As String
Public Sub AddSQLSelect(sText As String)
   sText = Trim(sText)
   If sText = "" Then Exit Sub
   If sSelect = "" Then
      sSelect = msSELECT & sText
   Else
      sSelect = sSelect & ", " & sText
   End If
End Sub
Public Sub AddSQLFrom(sText As String)
   sText = Trim(sText)
   If sText = "" Then Exit Sub
   If sFrom = "" Then
      sFrom = msFROM & sText
   Else
      sFrom = sFrom & ", " & sText
   End If
End Sub
Public Sub AddSQLWhere(sText As String)
   sText = Trim(sText)
   If sText = "" Then Exit Sub
   If sWhere = "" Then
      sWhere = msWHERE & sText
   Else
      sWhere = sWhere & " AND " & sText
   End If
End Sub
Public Sub AddSQLOrder(sText As String)
   sText = Trim(sText)
   If sText = "" Then Exit Sub
   If sOrder = "" Then
      sOrder = msORDER & sText
   Else
      sOrder = sOrder & ", " & sText
   End If
End Sub
Private Sub Class_Initialize()
   ClearSQLStatement
End Sub
Public Function GetSQLStatement() As String
   GetSQLStatement = sSelect & " " & _
      sFrom & " " & _
      sWhere & " " & _
      sOrder & " "
End Function
Public Sub ClearSQLStatement()
   sSelect = ""
   sFrom = ""
   sWhere = ""
   sOrder = ""
End Sub
