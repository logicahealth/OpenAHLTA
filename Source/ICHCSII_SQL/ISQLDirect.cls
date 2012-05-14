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
Attribute VB_Name = "ISQLDirect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Enum edSeq
    edMAINSEQ = 0
    edAPPT_ID
    edAPPT_TRANS_ID
    edENC_TEMPLATE_ID
    edENCOUNTER_ID
    edIEN
    edPCM_CODE
    edPI_DATA_ID
    edUNIT_NUMBER
    edUSER_NUMBER
    edJDOE_NUMBER
    edTEMP_SSN
End Enum

Public Function Execute(SqlStatement As String) As Variant:     End Function
Public Sub ExecuteMulti(SqlStatements As Collection):           End Sub
Public Function ExecuteCmd(oAdoCmd As ADODB.Command):           End Function
Public Function GetNextID(Optional ByVal eSeqName As edSeq = edMAINSEQ): End Function
Public Function ValidateFacility(ByVal FacilityNcid As Long, _
                                 ByVal CreateIfNotFound As Boolean, _
                        Optional ByVal ClinicNCID As Long = 0) As Boolean: End Function
Public Function GetNextNcid():                                  End Function
Public Sub MovePrevious():                                      End Sub
Public Sub MoveNext():                                          End Sub
Public Sub MoveLast():                                          End Sub
Public Sub MoveFirst():                                         End Sub
Public Property Get Source() As String:                         End Property
Public Property Get RecordCount() As Long:                      End Property
Public Property Get EOF() As Boolean:                           End Property
Public Property Get BOF() As Boolean:                           End Property
Public Property Get Value(Column As Variant) As Variant:        End Property
Attribute Value.VB_UserMemId = 0
Public Property Get InfoType(Column As Variant) As String:      End Property
Public Property Get ColumnIndex(Name As String) As Integer:     End Property
Public Property Get ColumnName(index As Integer) As String:     End Property
Public Property Get ColumnCount() As Integer:                   End Property




