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
Attribute VB_Name = "RSObject_Ops"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IDMBE.I_RSObjects

Public TableName As String
Public DBRECORDSET As ADODB.Recordset

Private Property Set I_RSObjects_DBRecordSet(rhs As ADODB.Recordset)
   Set DBRECORDSET = rhs
End Property

Private Property Get I_RSObjects_DBRecordSet() As ADODB.Recordset
   Set I_RSObjects_DBRecordSet = DBRECORDSET
End Property

Private Property Get I_RSObjects_TableName() As String
   I_RSObjects_TableName = TableName
End Property

Private Property Let I_RSObjects_TableName(rhs As String)
   TableName = rhs
End Property

