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
Attribute VB_Name = "ISqlOpsEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Enum exSeq
    exMAINSEQ = 0
    exAPPT_ID
    exAPPT_TRANS_ID
    exENC_TEMPLATE_ID
    exENCOUNTER_ID
    exIEN
    exPCM_CODE
    exPI_DATA_ID
    exUNIT_NUMBER
    exUSER_NUMBER
    exJDOE_NUMBER
    exTEMP_SSN
End Enum

' SCR #439;   Developer:  06/05/2002 04:17 PM
' Note: Function Changed:  ValidateFacility2  ValidateProvider2  SQLDate  SQLCase  SQLQuote    Impact of Change:



Public Sub BeginTrans():                                        End Sub
Public Sub CommitTrans():                                       End Sub
Public Sub RollBackTrans():                                     End Sub

Public Function Execute(SqlStatement As String) As Variant:     End Function
Public Sub ExecuteMulti(SqlStatements As Collection):           End Sub
Public Function ExecuteCmd(oAdoCmd As ADODB.Command):           End Function
Public Function GetNextID(Optional ByVal eSeqName As exSeq = exMAINSEQ): End Function

Public Function ValidateFacility(ByVal FacilityNcid As Long, _
                                 ByVal CreateIfNotFound As Boolean, _
                        Optional ByVal ClinicNCID As Long = 0) As Boolean: End Function
                        
Public Function ValidateProvider(ByVal UserNCID As Long, _
                                 ByVal CreateIfNotFound As Boolean, _
                        Optional ByVal FacilityNcid As Long = 0, _
                        Optional ByVal ClinicNCID As Long = 0) As Boolean: End Function


Public Function ValidateFacility2(ByRef FacilityNcid, _
                                 ByRef FacilityName, _
                                 ByVal CreateIfNotFound As Boolean, _
                        Optional ByRef ClinicNCID, _
                        Optional ByRef ClinicName) As Boolean: End Function
                                 
Public Function ValidateProvider2(ByRef ProviderNCID As String, _
                                 ByVal ProviderSSN As String, _
                                 ByVal ProviderFirstName As String, _
                                 ByVal ProviderMiddleName As String, _
                                 ByVal ProviderLastName As String, _
                                 ByVal CreateIfNotFound As Boolean) As Boolean: End Function

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

Public Property Let Value(Column As Variant, ByVal ColumnData As Variant):       End Property
Public Function AddNew() As Boolean:                            End Function
Public Function Update() As Boolean:                            End Function
Public Function Cancel() As Boolean:                            End Function

Public Function AppendAsChunk(ByVal vColumn As Variant, _
                                    vText As Variant) As Boolean: End Function
Public Function GetAsChunk(ByVal vColumn As Variant, _
                                 vText As Variant) As Boolean: End Function
                                    
Public Function SQLDate(ByVal dDate As Date) As String:         End Function
Public Function SQLQuote(ByVal sText As String) As String:     End Function


