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
Attribute VB_Name = "iBackend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get ColumnCount() As Integer
End Property

Public Property Get ColumnIndex(ByRef Name As String) As Integer
End Property

Public Property Get ColumnName(ByVal index As Integer) As String
End Property

Public Property Get BOF() As Boolean
End Property

Public Property Get EOF() As Boolean
End Property

Public Property Get InfoType(ByRef Column As Variant) As String
End Property

Public Property Get RecordCount() As Long
End Property

Public Property Get Source() As String
End Property

Public Property Get Value(ByRef Column As Variant) As Variant
End Property

Public Function CreateNewCDRClinicNCID(ByVal strClinicName As String, ByVal strFacilityNCID As String, ByVal strEnterpriseNCID As String) As String
End Function

Public Sub Execute(ByRef SqlStatement As String)
End Sub

Public Sub ExecuteMulti(ByRef SqlStatements As Collection)
End Sub

Public Sub MoveFirst()
End Sub

Public Sub MoveLast()
End Sub

Public Sub MoveNext()
End Sub

Public Sub MovePrevious()
End Sub

Public Function ValidateClinicAgainstCDR(ByRef ClinicNCID As String, ByVal FacilityNCID As String, FacilityIsValid As Boolean, NCIDIsRepresentation As Boolean, ByVal EnterpriseNCID As String, ByVal CreateIfNotFound As Boolean) As Boolean
End Function

Public Function ValiadeFacilityAgainstCDR(ByRef FacilityNCID As String, ByVal EnterpriseNCID As String, NCIDIsRepresentation As Boolean, ByVal CreateIfNotFound As Boolean) As Boolean
End Function


