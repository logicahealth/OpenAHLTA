VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LGSBackend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This class is used to connect to the Legacy Gateway Server (LGS) via the
'CHCSConnection interface.  Data is pulled from the LGS for MedcinLite
'(CWShared.AppMode = CHCS_GUI) instead of from the CDR.

Option Explicit

Implements iBackend

Private mCHCSConnection As Object   'Client-side interface to LGS (shared object in CWShared)

Private mSkipFacilityWorkAround As Boolean

Private Function iBackend_CreateRecordsetFromFile(ByVal sFileName As String, ByVal sDelimiter As String) As ADODB.Recordset
  'Not supported
End Function

Private Sub iBackend_ExecuteSQL(ByVal SQL As String)

  On Error GoTo ErrHandler

  InitCHCSConnection
  
  mCHCSConnection.Execute SQL
  
  Exit Sub
  
ErrHandler:
  Err.Raise Err.Number, Err.Source, Err.Description
  
End Sub

Private Function iBackend_OpenRecordset(ByVal SQL As String) As ADODB.Recordset

  On Error GoTo ErrHandler
  
  InitCHCSConnection
    
  Set iBackend_OpenRecordset = mCHCSConnection.OpenRecordset(SQL)
  
  Exit Function
  
ErrHandler:
  Err.Raise Err.Number, Err.Source, Err.Description
  
End Function

Private Property Let iBackend_SkipFacilityWorkAround(ByVal RHS As Boolean)
  mSkipFacilityWorkAround = RHS
End Property

Private Property Get iBackend_SkipFacilityWorkAround() As Boolean
  iBackend_SkipFacilityWorkAround = mSkipFacilityWorkAround
End Property

Private Sub InitCHCSConnection()
Dim CoreShared As CWShared

  If mCHCSConnection Is Nothing Then
    Set CoreShared = New CWShared
    
    Set mCHCSConnection = CoreShared.CHCSConnection
  End If
  
  If mCHCSConnection Is Nothing Then
    Err.Raise 62000, "CHCSII_DataAccess:iBackend_OpenRecordset", "Unable to Connect to CHCS."
  End If
  
End Sub
