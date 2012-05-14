VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MapperTheater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements DL_Interfaces.iMapperPatient

Private mobjMapperGeneric As DL_Interfaces.iMapperPatient


Private Sub Class_Initialize()
Dim objMapper As MapperGeneric
Dim oSQLFix As iSQLFix
Dim oPicklist As iPickList
Dim oShared As New CWShared

On Error GoTo ErrHandler

  Set mobjMapperGeneric = New MapperGeneric
  Set oSQLFix = New SQLOracle
  Set oPicklist = New PickListGEMS
  
  Set objMapper = mobjMapperGeneric
  
  If oShared.UseMySQLDB Then
      objMapper.BackEndDB = MYSQL
  Else
      objMapper.BackEndDB = GEMS
  End If
  Set objMapper.SQLFix = oSQLFix
  Set objMapper.PickList = oPicklist
  Exit Sub
  
ErrHandler:
  Err.Raise Err.Number, "DM_Patient.MapperTheater.Initialize", Err.Description, Err.HelpFile, Err.HelpContext


End Sub
'--------------------------------------------------------------
'--- 52797
'--- Rank Code is one of the data from CAC
'--- Need to select the correct Rank from the drop down list
'--------------------------------------------------------------
Private Function iMapperPatient_GetRankNCID(ByVal sRankCode As String, ByVal sBranch As String) As String
    Dim sSQL        As String
    Dim oConn       As CHCSII_CONN.Conn
    Dim oRS         As ADODB.Recordset

    On Error GoTo ErrHandler
    
    iMapperPatient_GetRankNCID = vbNullString
    Set oConn = New CHCSII_CONN.Conn
    sSQL = "SELECT NCID FROM RANK_ABBREV " & vbNewLine
    '--- SCR 53556
    sSQL = sSQL & "WHERE REPLACE(UPPER(ABBREV_RANK), '-', '') = '" & Replace(UCase(sRankCode), "-", vbNullString) & "'"
    sSQL = sSQL & " AND BRANCH = '" & sBranch & "'"
    Set oRS = oConn.CHCSII_DAS(Auto).OpenRecordset(sSQL)
    If oRS.RecordCount > 0 Then
        iMapperPatient_GetRankNCID = oRS("NCID")
    End If
    
    Set oConn = Nothing
    Set oRS = Nothing
    Exit Function
    
ErrHandler:
   Err.Raise Err.Number, "DM_Patient.MapperTheater.iMapperPatient_GetRankNCID", Err.Description, Err.HelpFile, Err.HelpContext
End Function

Private Sub iMapperPatient_Init(Support As DL_Support.iParameters, Cache As DL_Support.CacheMapper)
    mobjMapperGeneric.Init Support, Cache
End Sub

Private Function iMapperPatient_Retrieve(Params As DL_Support.iParameters) As DTOs.Patients

On Error GoTo ErrHandler
    
    Set iMapperPatient_Retrieve = mobjMapperGeneric.Retrieve(Params)
    
    GoTo CleanUpAndExit
ErrHandler:
    Err.Raise Err.Number, "MapperTheater.iMapperPatient_Retrieve", Err.Description, Err.HelpFile, Err.HelpContext
CleanUpAndExit:
End Function

Private Function iMapperPatient_RetrieveSelectStrings(Params As DL_Support.iParameters) As DTOs.PatientSelectStrings
On Error GoTo ErrHandler
    
    Set iMapperPatient_RetrieveSelectStrings = mobjMapperGeneric.RetrieveSelectStrings(Params)
    
    GoTo CleanUpAndExit
ErrHandler:
    Err.Raise Err.Number, "MapperTheater.iMapperPatient_RetrieveSelectStrings", Err.Description, Err.HelpFile, Err.HelpContext
CleanUpAndExit:

End Function

Private Function iMapperPatient_RetrieveRepresentation(ByRef Params As DL_Support.iParameters) As String
    On Error GoTo ErrHandler
    
    iMapperPatient_RetrieveRepresentation = mobjMapperGeneric.RetrieveRepresentation(Params)
    
    GoTo CleanUpAndExit
ErrHandler:
    Err.Raise Err.Number, "MapperTheater.iMapperPatient_RetrieveRepresentation", Err.Description, Err.HelpFile, Err.HelpContext
CleanUpAndExit:
End Function

Private Function iMapperPatient_Save(ByRef Patient As DTOs.Patient) As Boolean
On Error GoTo ErrHandler
    
    iMapperPatient_Save = mobjMapperGeneric.Save(Patient)
    
    GoTo CleanUpAndExit
ErrHandler:
    Err.Raise Err.Number, "MapperTheater.iMapperPatient_RetrieveSelectStrings", Err.Description, Err.HelpFile, Err.HelpContext
CleanUpAndExit:

End Function


Private Function iMapperPatient_TransferCHCSPatient(ByRef Params As DL_Support.iParameters) As DTOs.PatientTransferResult

End Function

