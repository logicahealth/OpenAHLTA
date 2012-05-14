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
Attribute VB_Name = "LGS_DAS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Implements ICHCSII_DAS
'This class is used to connect to the Legacy Gateway Server (LGS) via the
'CHCSConnection interface.  Data is pulled from the LGS for MedcinLite
'(CWShared.AppMode = CHCS_GUI) instead of from the CDR.

Private mCHCSConnection As Object   'Client-side interface to LGS (shared object in CWShared)

Private mSkipFacilityWorkAround As Boolean


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

Private Property Get ICHCSII_DAS_ActiveConnection() As ADODB.Connection

    Set ICHCSII_DAS_ActiveConnection = Nothing

End Property

Private Function ICHCSII_DAS_BeginTrans() As Boolean

    Err.Raise 1001, "LGS_DAS", "BeginTrans is not supported in LGS DAS."

End Function


Private Function ICHCSII_DAS_CommitTrans() As Boolean

    Err.Raise 1001, "LGS_DAS", "CommitTrans is not supported in LGS DAS."

End Function


Private Function ICHCSII_DAS_CreateRecordsetFromFile(sFileName As String, sDelimiter As String) As ADODB.Recordset

    Err.Raise 1001, "LGS_DAS", "CreateRecordsetFromFile is not supported in LGS DAS."

End Function


Private Function ICHCSII_DAS_ExecuteSP(Parms As CHCSII_Cmd) As Variant

    Err.Raise 1001, "LGS_DAS", "ExecuteSP is not supported in LGS DAS."

End Function


Private Function ICHCSII_DAS_ExecuteSQL(SQL As String) As Variant

  On Error GoTo ErrHandler

  InitCHCSConnection
  
  'Adjust SQL for access to encrypted tables
  SQL = FixParans(SQL)
  Call RedirectEncrypted(SQL, Me)
    
  mCHCSConnection.Execute SQL
  
  Exit Function
  
ErrHandler:
  Err.Raise Err.Number, Err.Source, Err.Description

End Function


Private Property Let ICHCSII_DAS_FacilityWorkaround(ByVal RHS As Boolean)
    mSkipFacilityWorkAround = Not RHS
End Property


Private Function ICHCSII_DAS_GetNextID(Optional ByVal eSeqName As ICHCSII_DataAccess.eSeq = 0&) As Variant
Dim oRs As New ADODB.Recordset
Dim sSeqName As String
Dim lSeqNum As Long
Dim CoreShared As CWShared
    
    On Error GoTo ErrHandler
    
    If mCHCSConnection Is Nothing Then
      Set CoreShared = New CWShared
      
      Set mCHCSConnection = CoreShared.CHCSConnection
    End If
    
    sSeqName = SeqEnumToString(eSeqName)
    Set oRs = mCHCSConnection.OpenRecordset("Select " & sSeqName & ".nextval from dual")

    If Not oRs.EOF Then
        ICHCSII_DAS_GetNextID = oRs(0)
    End If
    
ErrHandler:
    If oRs.State > 0 Then oRs.Close
    Set oRs = Nothing

End Function


Private Function ICHCSII_DAS_GetNextNcid() As Variant

    Err.Raise 1001, "LGS_DAS", "GetNextNcid is not supported in LGS DAS."

End Function


Private Function ICHCSII_DAS_OpenRecordset(SQL As String, Optional ByVal CursorType As ADODB.CursorTypeEnum = 0&, Optional ByVal LockType As ADODB.LockTypeEnum = 1&) As ADODB.Recordset

  On Error GoTo ErrHandler
  
  InitCHCSConnection
    
  Set ICHCSII_DAS_OpenRecordset = mCHCSConnection.OpenRecordset(SQL)
  
  Exit Function
  
ErrHandler:
  Err.Raise Err.Number, Err.Source, Err.Description

End Function


Private Function ICHCSII_DAS_OpenRecordsetMSP(Parms As Collection) As Collection

    Err.Raise 1001, "LGS_DAS", "OpenRecordsetMSP is not supported in LGS DAS."

End Function


Private Function ICHCSII_DAS_OpenRecordsetSP(Parms As CHCSII_Cmd) As ADODB.Recordset

    Err.Raise 1001, "LGS_DAS", "OpenRecordsetSP is not supported in LGS DAS."

End Function


Private Function ICHCSII_DAS_RollbackTrans() As Boolean

    Err.Raise 1001, "LGS_DAS", "RollbackTrans is not supported in LGS DAS."

End Function


Private Property Let ICHCSII_DAS_UserNCID(RHS As String)

    Err.Raise 1001, "LGS_DAS", "UserNCID is not supported in LGS DAS."

End Property

Private Function SeqEnumToString(ByVal eiSeq As eSeq) As String
    Select Case eiSeq
    Case esMAINSEQ:         SeqEnumToString = "MainSeq"
    Case esAPPT_ID:         SeqEnumToString = "APPT_ID"
    Case esAPPT_TRANS_ID:   SeqEnumToString = "APPT_TRANS_ID"
    Case esENC_TEMPLATE_ID: SeqEnumToString = "ENC_TEMPLATE_ID"
    Case esENCOUNTER_ID:    SeqEnumToString = "ENCOUNTER_ID"
    Case esIEN:             SeqEnumToString = "IEN"
    Case esPCM_CODE:        SeqEnumToString = "PCM_CODE"
    Case esPI_DATA_ID:      SeqEnumToString = "PI_DATA_ID"
    Case esUNIT_NUMBER:     SeqEnumToString = "UNIT_NUMBER"
    Case esUSER_NUMBER:     SeqEnumToString = "USER_NUMBER"
    Case esJDOE_NUMBER:     SeqEnumToString = "JDOE_NUMBER"
    Case esTEMP_SSN:        SeqEnumToString = "TEMP_SSN"
    End Select

End Function

