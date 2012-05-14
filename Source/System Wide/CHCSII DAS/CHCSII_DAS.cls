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
Attribute VB_Name = "CHCSII_DAS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ICHCSII_DAS

Private mBackend As iBackend    'Use either the CDR or LGS as the data source

Private Sub Class_Initialize()
Dim CoreShared As CWShared

  Set CoreShared = New CWShared
  
  If CoreShared.AppMode <> modeCHCSI_GUI Then
    Set mBackend = New CDRBackend
  Else
    Set mBackend = New LGSBackend
  End If

End Sub

Private Sub Class_Terminate()
    Set mBackend = Nothing
End Sub

Private Property Get ICHCSII_DAS_ActiveConnection() As ADODB.Connection
    Set ICHCSII_DAS_ActiveConnection = Nothing
End Property

'- GEMS
Private Function ICHCSII_DAS_CreateRecordsetFromFile(sFileName As String, sDelimiter As String) As ADODB.Recordset
    Set ICHCSII_DAS_CreateRecordsetFromFile = mBackend.CreateRecordsetFromFile(sFileName, sDelimiter)
End Function

'- GEMS
Private Function ICHCSII_DAS_ExecuteSQL(SQL As String) As Variant
    mBackend.ExecuteSQL SQL
End Function

Private Function ICHCSII_DAS_GetNextID(Optional ByVal eSeqName As ICHCSII_DataAccess.eSeq = 0&) As Variant

End Function

Private Function ICHCSII_DAS_GetNextNcid() As Variant

End Function

'- GEMS
Private Function ICHCSII_DAS_OpenRecordset(SQL As String, _
                            Optional ByVal CursorType As ADODB.CursorTypeEnum = 0&, _
                            Optional ByVal LockType As ADODB.LockTypeEnum = 1&) As ADODB.Recordset
    Set ICHCSII_DAS_OpenRecordset = mBackend.OpenRecordset(SQL)
End Function

Private Function ICHCSII_DAS_OpenRecordsetMSP(Parms As Collection) As Collection

End Function

'- GEMS
Private Property Let ICHCSII_DAS_UserNCID(RHS As String): End Property
Public Function ICHCSII_DAS_BeginTrans() As Boolean:      Err.Raise 1001, "CHCSII_DAS", "Transaction control is not supported in CHCSII Data Access Service.": End Function
Public Function ICHCSII_DAS_CommitTrans() As Boolean:     Err.Raise 1001, "CHCSII_DAS", "Transaction control is not supported in CHCSII Data Access Service.": End Function
Public Function ICHCSII_DAS_RollbackTrans() As Boolean:   Err.Raise 1001, "CHCSII_DAS", "Transaction control is not supported in CHCSII Data Access Service.": End Function

Private Function ICHCSII_DAS_OpenRecordsetSP(Parms As CHCSII_Command.CHCSII_Cmd) As ADODB.Recordset
   
End Function

Private Function ICHCSII_DAS_ExecuteSP(Parms As CHCSII_Command.CHCSII_Cmd) As Variant
 
End Function

Private Function ICHCSII_DAS_ExecuteMSP(Parms As Collection) As Collection
 
End Function

Public Property Let ICHCSII_DAS_FacilityWorkaround(ByVal Enabled As Boolean)
    mBackend.SkipFacilityWorkAround = Not Enabled
End Property

Public Function OpenRecordset(SQL As String) As ADODB.Recordset

  Set OpenRecordset = mBackend.OpenRecordset(SQL)

End Function

Public Sub ExecuteSQL(SQL As String)

  mBackend.ExecuteSQL SQL
  
End Sub

Function CreateRecordsetFromFile(sFileName As String, sDelimiter As String) As ADODB.Recordset

  Set CreateRecordsetFromFile = mBackend.CreateRecordsetFromFile(sFileName, sDelimiter)

End Function

