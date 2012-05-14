VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LGS_SQLOPS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Implements ICHCSII_SQL.ISqlOps

Private mCHCSConnection As Object   'Client-side interface to LGS (shared object in CWShared)

Private mRecordSet As ADODB.Recordset

Private Property Get ISqlOps_BOF() As Boolean
  ISqlOps_BOF = mRecordSet.BOF
End Property

Private Property Get ISqlOps_ColumnCount() As Integer
  ISqlOps_ColumnCount = mRecordSet.Fields.Count
End Property

Private Property Get ISqlOps_ColumnIndex(Name As String) As Integer
Dim Flds As Fields
Dim i As Integer
Dim sText As String

  ISqlOps_ColumnIndex = -1
  
  Set Flds = mRecordSet.Fields
  
  For i = 0 To Flds.Count - 1
    If UCase$(Flds(i).Name) = UCase$(Name) Then
      ISqlOps_ColumnIndex = i
      
      Exit For
    End If
  Next
  
  If ISqlOps_ColumnIndex = -1 Then
      sText = "There is no column named '" & Name & "'.  Please select from:"
    For i = 0 To Flds.Count - 1
        sText = sText & Chr(10) & Flds(i).Name
    Next
  
    Err.Raise vbObjectError + 1, "CHCSII_SQL:SqlOps.ColumnIndex", sText
  
  End If
  
End Property


Private Property Get ISqlOps_ColumnName(index As Integer) As String

    If Not (mRecordSet.BOF Or mRecordSet.EOF) Then
        If index > mRecordSet.Fields.Count Then
            Err.Raise vbObjectError + 1, "CHCSII_SQL:SqlOps.ColumnName", "Index (" & index & ") out of range.  There are " & mRecordSet.Fields.Count & " colummns."
        End If
        
        ISqlOps_ColumnName = mRecordSet.Fields(index - 1).Name
    End If

End Property


Private Function ISqlOps_CreateNewCDRClinicNCID(ByVal strClinicName As String, ByVal strFacilityNCID As String, ByVal strEnterpriseNCID As String) As String
  'Not Supported
End Function

Private Property Get ISqlOps_EOF() As Boolean
  ISqlOps_EOF = mRecordSet.EOF
End Property


Private Function ISqlOps_Execute(SqlStatement As String) As Variant

  InitCHCSConnection
  
  If Left$(UCase$(SqlStatement), 6) = "SELECT" Then
    Set mRecordSet = mCHCSConnection.OpenRecordset(SqlStatement)
    
    If mRecordSet Is Nothing Then
      Set mRecordSet = New ADODB.Recordset
    End If
  Else
    mCHCSConnection.Execute SqlStatement
  End If

End Function


Private Function ISqlOps_ExecuteCmd(oAdoCmd As ADODB.Command) As Variant

    Err.Raise 1001, "LGS_SQLOPS", "ExecuteCmd is not supported in LGS_SQLOPS."
    
End Function

Private Sub ISqlOps_ExecuteMulti(SqlStatements As Collection)

Dim i As Long

  For i = 1 To SqlStatements.Count
    ISqlOps_Execute SqlStatements.Item(i)
  Next
  
End Sub

Private Function ISqlOps_ExecuteSP(Parms As CHCSII_Command.CHCSII_Cmd) As Variant

    Err.Raise 1001, "LGS_SQLOPS", "ExecuteSP is not supported in LGS_SQLOPS."

End Function


Private Function ISqlOps_GetNextID(Optional ByVal eSeqName As ICHCSII_SQL.eSeq = 0&) As Variant
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
        ISqlOps_GetNextID = oRs(0)
    End If
    
ErrHandler:
    If oRs.State > 0 Then oRs.Close
    Set oRs = Nothing
    
End Function


Private Function ISqlOps_GetNextNcid() As Variant

    Err.Raise 1001, "LGS_SQLOPS", "GetNextNcid is not supported in LGS_SQLOPS."

End Function


Private Property Get ISqlOps_InfoType(Column As Variant) As String
  ISqlOps_InfoType = mRecordSet.Fields(Column).Type
End Property

Private Sub ISqlOps_MoveFirst()
  mRecordSet.MoveFirst
End Sub

Private Sub ISqlOps_MoveLast()
  mRecordSet.MoveLast
End Sub

Private Sub ISqlOps_MoveNext()
  mRecordSet.MoveNext
End Sub

Private Sub ISqlOps_MovePrevious()
  mRecordSet.MovePrevious
End Sub

Private Property Get ISqlOps_RecordCount() As Long
  ISqlOps_RecordCount = mRecordSet.RecordCount
End Property

Private Property Get ISqlOps_Source() As String
  ISqlOps_Source = mRecordSet.Source
End Property

Private Function ISqlOps_ValiadeFacilityAgainstCDR(FacilityNcid As String, ByVal EnterpriseNCID As String, NCIDIsRepresentation As Boolean, ByVal CreateIfNotFound As Boolean) As Boolean
  'Not Supported
End Function

Private Function ISqlOps_ValidateClinicAgainstCDR(ClinicNCID As String, ByVal FacilityNcid As String, FacilityIsValid As Boolean, NCIDIsRepresentation As Boolean, ByVal EnterpriseNCID As String, ByVal CreateIfNotFound As Boolean) As Boolean
  'Not Supported
End Function

Private Property Get ISqlOps_Value(Column As Variant) As Variant
  ISqlOps_Value = mRecordSet.Fields(Column).Value
End Property

Private Sub InitCHCSConnection()
Dim CoreShared As CWShared

  If mCHCSConnection Is Nothing Then
    Set CoreShared = New CWShared
    
    Set mCHCSConnection = CoreShared.CHCSConnection
  End If
  
  If mCHCSConnection Is Nothing Then
    Err.Raise 62000, "CHCSII_DataAccess:ISqlOps_OpenRecordset", "Unable to Connect to CHCS."
  End If
  
End Sub

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

