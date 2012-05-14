VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDR_SQLOPS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Implements ICHCSII_SQL.ISqlOps

Private moSqlConnect As SqlServConnect
Private msServiceName As String

Private meTimeZoneConversion As eTimeZoneConversion

Private mvarAbsolutePosition As Long
'Private mvarDSQLConnect As DSQLATLLib.DSqlConnect


Public Sub Init(TimeZoneConversion As eTimeZoneConversion, ServiceName As eServiceNameMode)

    'SF 1-22-04 Indicates if dates should be converted for the local time zone.
    meTimeZoneConversion = TimeZoneConversion
    
    'SCR 46007
    'Developer: SF 12-6-03
    'Created to expose the service name without modifying ICHCSII_DAS
    
    msServiceName = TranslateServiceName(ServiceName)

End Sub

Private Sub Class_Initialize()

    Set moSqlConnect = New SQLSERVATLLib.SqlServConnect
    'Set mvarDSQLConnect = New DSQLATLLib.DSqlConnect
    
    GetTimeZoneInformation gTZInfo
    mvarAbsolutePosition = 0
    
End Sub

Private Sub Class_Terminate()
    
    Set moSqlConnect = Nothing
    'Set mvarDSQLConnect = Nothing
    
End Sub

Private Property Get ISqlOps_BOF() As Boolean
    
    If mvarAbsolutePosition = 0 Then
        ISqlOps_BOF = True
    Else
        ISqlOps_BOF = False
    End If

End Property

Private Property Get ISqlOps_ColumnCount() As Integer
    
    ISqlOps_ColumnCount = moSqlConnect.Transaction(1).ColumnNames.Count
    'ISqlOps_ColumnCount = mvarDSQLConnect.Table.Rows(mvarAbsolutePosition).Columns.Count
    
End Property

Private Property Get ISqlOps_ColumnIndex(Name As String) As Integer

Dim i As Integer
Dim sText As String
    
    For i = 1 To ISqlOps_ColumnCount
        If UCase(ISqlOps_ColumnName(i)) = UCase(Name) Then
            ISqlOps_ColumnIndex = i
            Exit Property
        End If
    Next i
    sText = "There is no column named '" & Name & "'.  Please select from:"
    For i = 1 To ISqlOps_ColumnCount
        sText = sText & Chr(10) & moSqlConnect.Transaction(1).ColumnNames.Item(i).Text
        'sText = sText & Chr(10) & mvarDSQLConnect.Table(mvarAbsolutePosition, i).Name
    Next
    
    Err.Raise vbObjectError + 1, "CHCSII_SQL:SqlOps.ColumnIndex", sText

End Property


Private Property Get ISqlOps_ColumnName(index As Integer) As String

    ISqlOps_ColumnName = ""
    
    If Not (ISqlOps_BOF Or ISqlOps_EOF) Then
        If index > ISqlOps_ColumnCount Then
            Err.Raise vbObjectError + 1, "CHCSII_SQL:SqlOps.ColumnName", "Index (" & index & ") out of range.  There are " & ISqlOps_ColumnCount & " colummns."
            Exit Property
        End If
        ISqlOps_ColumnName = moSqlConnect.Transaction(1).ColumnNames(index).Text
        'ISqlOps_ColumnName = mvarDSQLConnect.Table(mvarAbsolutePosition, index).Name
    End If

End Property

'This function will create a new Clinic NCID in CDR.
'strClinicName will become the 'representation' of the newly created Facility NCID
'The function will return the NCID of the newly created Facility NCID
Private Function ISqlOps_CreateNewCDRClinicNCID(ByVal strClinicName As String, ByVal strFacilityNCID As String, ByVal strEnterpriseNCID As String) As String
Dim oHDD As HDDMAINTLib.HDDMaintenance
Dim strReturnedNCID As String
    
    Set oHDD = New HDDMAINTLib.HDDMaintenance
    With oHDD
        .ClearAll
        .WholeConcepts.Add
        With .WholeConcepts(1)
            .ClearAll
            .Concept.Definition = "A DoD field clinic"
            .Concept.Comments = "Created " & Format(Now, "mm/dd/yyyy hh:nn")
            '---create the Profile Domain relationships
            .Relationships.Add
            With .Relationships(1)
                .ClearAll
                .ConceptRelationNcid = NCID_CLINIC
                .RelationshipNcid = NCID_HAS_CHILD
                .EnterpriseNCID = CLng(strEnterpriseNCID)
            End With
            .Relationships.Add
            With .Relationships(2)
                .ClearAll
                .ConceptRelationNcid = NCID_CLINIC
                .RelationshipNcid = NCID_HAS_MEMBER
                .EnterpriseNCID = CLng(strEnterpriseNCID)
            End With
            .Relationships.Add
            With .Relationships(3)
                .ClearAll
                .ConceptRelationNcid = NCID_CLINIC
                .RelationshipNcid = NCID_HAS_PART
                .EnterpriseNCID = CLng(strEnterpriseNCID)
            End With
             .Relationships.Add
            With .Relationships(4)
                .ClearAll
                .ConceptRelationNcid = CLng(strFacilityNCID)
                .RelationshipNcid = NCID_HAS_CHILD
                .EnterpriseNCID = CLng(strEnterpriseNCID)
            End With
            .Relationships.Add
            With .Relationships(5)
                .ClearAll
                .ConceptRelationNcid = CLng(strFacilityNCID)
                .RelationshipNcid = NCID_HAS_MEMBER
                .EnterpriseNCID = CLng(strEnterpriseNCID)
            End With
            .Relationships.Add
            With .Relationships(6)
                .ClearAll
                .ConceptRelationNcid = CLng(strFacilityNCID)
                .RelationshipNcid = NCID_HAS_PART
                .EnterpriseNCID = CLng(strEnterpriseNCID)
            End With

            '---create a context 2000 surface form whose representation is just the name of the Profile
            With .RSForms
                .ClearAll
                .Add
                With .Item(1)
                    .ClearAll
                    .Representation = LEADSTRING_CLINIC_REPRESENTATION & strClinicName
                    .Contexts.Add
                    With .Contexts(1)
                        .ContextNcid = 2000 'NCID_3M_DEFAULT_CONTEXT
                        .EnterpriseNCID = CLng(strEnterpriseNCID)
                        .PreferredScore = 0
                    End With
                    .Contexts.Add
                    With .Contexts(2)
                        .ContextNcid = 367  'NCID_DEFAULT_CONTEXT
                        .EnterpriseNCID = CLng(strEnterpriseNCID)
                        .PreferredScore = 0
                    End With
                    .Contexts.Add
                    With .Contexts(3)
                        .ContextNcid = 21915 'DODShortName
                        .EnterpriseNCID = CLng(strEnterpriseNCID)
                        .PreferredScore = 0
                    End With
                    .Contexts.Add
                    With .Contexts(4)
                        .ContextNcid = 21916 'NCID_DOD_LONG_NAME_CONTEXT
                        .EnterpriseNCID = CLng(strEnterpriseNCID)
                        .PreferredScore = 0
                    End With
                End With
                .Add
                With .Item(2)
                    .ClearAll
                    .Representation = LEADSTRING_CLINIC_REPRESENTATION & strClinicName 'presumably the ncid of the clinical machine
                    .Contexts.Add
                    With .Contexts(1)
                        .ContextNcid = 2000 'NCID_3M_DEFAULT_CONTEXT
                        .EnterpriseNCID = CLng(strEnterpriseNCID)
                        .PreferredScore = 0
                    End With
                    .Contexts.Add
                    With .Contexts(2)
                        .ContextNcid = 367  'NCID_DEFAULT_CONTEXT
                        .EnterpriseNCID = CLng(strEnterpriseNCID)
                        .PreferredScore = 0
                    End With
                    .Contexts.Add
                    With .Contexts(3)
                        .ContextNcid = 21915
                        .EnterpriseNCID = CLng(strEnterpriseNCID)
                        .PreferredScore = 0
                    End With
                    .Contexts.Add
                    With .Contexts(4)
                        .ContextNcid = 21916 'NCID_DOD_LONG_NAME_CONTEXT
                        .EnterpriseNCID = CLng(strEnterpriseNCID)
                        .PreferredScore = 0
                    End With
                End With
            End With
        End With
        .Create
        strReturnedNCID = .WholeConcepts(1).Concept.Ncid
    End With
    
    ISqlOps_CreateNewCDRClinicNCID = strReturnedNCID

Exit Function

ErrHandler:
    Err.Raise Err.Number, "SqlOps.CreateNewCDRFacilityNCID", Err.Description
    Err.Clear

End Function

Private Property Get ISqlOps_EOF() As Boolean
    If ISqlOps_RecordCount = 0 Then
        ISqlOps_EOF = True
    ElseIf mvarAbsolutePosition = ISqlOps_RecordCount + 1 Then
        ISqlOps_EOF = True
    Else
        ISqlOps_EOF = False
    End If
End Property


Private Function ISqlOps_Execute(SqlStatement As String) As Variant

Dim oStatement As SqlServStatement

'For now, we raise the error to the client code.
'On Error GoTo ErrHnd
    
    moSqlConnect.ClearAll
    
    moSqlConnect.ServiceName = msServiceName
    
    Set oStatement = moSqlConnect.Transaction.Add

    oStatement.SqlId.ValueConcept.UniqueId = gSQLQueryNCID

    oStatement.SqlStatement.Text = SqlStatement
    
    ExecuteSQLServ moSqlConnect
    'moSqlConnect.ExecuteTransaction
    
    ISqlOps_MoveFirst
    
    ISqlOps_Execute = True

Exit Function

ErrHnd:

    ISqlOps_Execute = False

'Dim sOraError As String
'
'    On Error GoTo CheckErr
'
'    With mvarDSQLConnect
'        .ClearAll
'        .Table.sqlId.ValueConcept.UniqueId = gSQLQueryNCID
'        .Table.arguments.Add
'        .Table.arguments(1).Text = SqlStatement
'        .Query
'CheckErr:
'        If Len(.Table.Error.ErrorText) <> 0 Then
'            On Error Resume Next
'            sOraError = Left(.Table.Error.ErrorText, 9)
'            Select Case sOraError
'            Case OraErrorNoDataFound
'                ' do nothing
'            Case Else
'                Debug.Print "CHCSII_SQL:SqlOps.Execute(" & SqlStatement & ")", "Error Text = " & .Table.Error.ErrorText
'                Err.Raise vbObjectError + 1, "CHCSII_SQL:SqlOps.Execute(" & SqlStatement & ")", "Error Text = " & .Table.Error.ErrorText
'                Exit Function
'            End Select
'        ElseIf Err Then
'            Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
'        End If
'    End With
'
'    ISqlOps_MoveFirst

End Function

Private Function ISqlOps_ExecuteCmd(oAdoCmd As ADODB.Command) As Variant

    Err.Raise 1001, "CDR_SQLOPS", "ExecuteCmd is not supported in CDR_SQLOPS."

End Function

Private Sub ISqlOps_ExecuteMulti(SqlStatements As Collection)

Dim oStatement As SqlServStatement
Dim iCounter As Integer

'For now, we raise the error to the client code.
'On Error GoTo ErrHnd
    
    moSqlConnect.ClearAll
    
    moSqlConnect.ServiceName = msServiceName

    For iCounter = 1 To SqlStatements.Count
        Set oStatement = moSqlConnect.Transaction.Add
    
        oStatement.SqlId.ValueConcept.UniqueId = gSQLQueryNCID
    
        oStatement.SqlStatement.Text = SqlStatements(iCounter)
    Next
    
    ExecuteSQLServ moSqlConnect
    'moSqlConnect.ExecuteTransaction

'Dim i As Integer
'Dim sOraError As String
'
'    With mvarDSQLConnect
'        .ClearAll
'        .Table.sqlId.ValueConcept.UniqueId = gSQLQueryNCID
'        For i = 1 To SqlStatements.Count
'            .Table.arguments.Add
'            .Table.arguments(i).Text = SqlStatements(i)
'        Next i
'        .Query
'        If Len(.Table.Error.ErrorText) <> 0 Then
'            sOraError = Left(.Table.Error.ErrorText, 9)
'            Select Case sOraError
'            Case OraErrorNoDataFound
'                ' do nothing
'            Case Else
'                Debug.Print "CHCSII_SQL:SqlOps.Execute(" & .Table.SqlCString & ")", "Error Text = " & .Table.Error.ErrorText
'                Err.Raise vbObjectError + 1, "CHCSII_SQL:SqlOps.Execute(" & .Table.SqlCString & ")", "Error Text = " & .Table.Error.ErrorText
'                Exit Sub
'            End Select
'        End If
'    End With
    
    ISqlOps_MoveFirst

End Sub


Private Function ISqlOps_ExecuteSP(Parms As CHCSII_Command.CHCSII_Cmd) As Variant

    Err.Raise 1001, "CDR_SQLOPS", "ExecuteSP is not supported in CDR_SQLOPS."

End Function

Private Function ISqlOps_GetNextID(Optional ByVal eSeqName As ICHCSII_SQL.eSeq = 0&) As Variant

    Err.Raise 1001, "CDR_SQLOPS", "GetNextID is not supported in CDR_SQLOPS."

End Function

Private Function ISqlOps_GetNextNcid() As Variant

    Err.Raise 1001, "CDR_SQLOPS", "GetNextNcid is not supported in CDR_SQLOPS."

End Function

Private Property Get ISqlOps_InfoType(Column As Variant) As String
Dim index As Integer
    
    If Not (ISqlOps_BOF Or ISqlOps_EOF) Then
        If IsNumeric(Column) Then
            index = Column
        Else
            index = ISqlOps_ColumnIndex(CStr(Column))
        End If
        ISqlOps_InfoType = moSqlConnect.Transaction(1).Rows(mvarAbsolutePosition)(index).InfoType
        'ISqlOps_InfoType = mvarDSQLConnect.Table(mvarAbsolutePosition, index).Value.InfoType
    End If

End Property

Private Sub ISqlOps_MoveFirst()
    If ISqlOps_RecordCount > 0 Then
        mvarAbsolutePosition = 1
    Else
        mvarAbsolutePosition = 0
    End If
End Sub

Private Sub ISqlOps_MoveLast()
    mvarAbsolutePosition = ISqlOps_RecordCount
End Sub

Private Sub ISqlOps_MoveNext()
    If mvarAbsolutePosition < ISqlOps_RecordCount + 1 Then
        mvarAbsolutePosition = mvarAbsolutePosition + 1
    End If
End Sub

Private Sub ISqlOps_MovePrevious()
    If mvarAbsolutePosition > 0 Then
        mvarAbsolutePosition = mvarAbsolutePosition - 1
    End If
End Sub

Private Property Get ISqlOps_RecordCount() As Long
    
    ISqlOps_RecordCount = moSqlConnect.Transaction(1).Rows.Count
    
    'ISqlOps_RecordCount = mvarDSQLConnect.Table.Rows.Count
    
End Property

Private Property Get ISqlOps_Source() As String
    
    ISqlOps_Source = moSqlConnect.Transaction(1).SqlStatement
    'ISqlOps_Source = mvarDSQLConnect.Table.SqlCString
    
End Property

Private Function ISqlOps_ValiadeFacilityAgainstCDR(FacilityNcid As String, ByVal EnterpriseNCID As String, NCIDIsRepresentation As Boolean, ByVal CreateIfNotFound As Boolean) As Boolean
    On Error GoTo ErrHandler

    If NCIDIsCDRFacility(FacilityNcid) Then
        ISqlOps_ValiadeFacilityAgainstCDR = True
        Exit Function
    Else
        ISqlOps_ValiadeFacilityAgainstCDR = False
        If NCIDIsCDRFacilityRepresentation(FacilityNcid) Then
            'FacilityNCID will be changed to the one it is representing
            NCIDIsRepresentation = True
            Exit Function
        Else
            NCIDIsRepresentation = False
        End If
    End If

    ISqlOps_ValiadeFacilityAgainstCDR = False
    
    If CreateIfNotFound = True Then
        FacilityNcid = CreateNewCDRFacilityNCID(FacilityNcid, EnterpriseNCID)
    End If

Exit Function

ErrHandler:
    ISqlOps_ValiadeFacilityAgainstCDR = False
    Err.Raise Err.Number, "SqlOps.ValiadeFacilityAgainstCDR", Err.Description
    Err.Clear

End Function

Private Function ISqlOps_ValidateClinicAgainstCDR(ClinicNCID As String, ByVal FacilityNcid As String, FacilityIsValid As Boolean, NCIDIsRepresentation As Boolean, ByVal EnterpriseNCID As String, ByVal CreateIfNotFound As Boolean) As Boolean
    
Dim oStatement As SqlServStatement

Dim strSQL As String
Dim strClinicSQL As String
Dim strOriginalNCID As String
    
    'first validate the facility sent in
    If NCIDIsCDRFacility(FacilityNcid) Then
        FacilityIsValid = True
    Else
        FacilityIsValid = False
        ISqlOps_ValidateClinicAgainstCDR = False
        Exit Function
    End If
    
    strOriginalNCID = ClinicNCID
    
    strSQL = "select distinct r.ncid as DomainNCID, "
    strSQL = strSQL & "cr.RELATIONSHIP_NCID, substr(r.representation,1,65) as DomainName "
    strSQL = strSQL & "from rsform_context rc,rsform r, "
    strSQL = strSQL & "concept_relation cr "
    strSQL = strSQL & "where rc.PREFERRED_SCORE = 0 and "
    strSQL = strSQL & "rc.context_ncid = 2000 and "
    strSQL = strSQL & "rc.rsform_id = r.rsform_id and "
    strSQL = strSQL & "r.ncid = cr.concept_relation_ncid and "
    strSQL = strSQL & "r.ncid = " & FacilityNcid & " and "
    
    
    If NCIDIsCDRClinic(ClinicNCID) = True Then
        'Still need to check whether there is a "matching" FacilityNCID
        
        strClinicSQL = "cr.concept_ncid = " & ClinicNCID & " "
        
        moSqlConnect.ClearAll
        moSqlConnect.ServiceName = msServiceName
        
        Set oStatement = moSqlConnect.Transaction.Add
        oStatement.SqlId.ValueConcept.UniqueId = gSQLQueryNCID
        oStatement.SqlStatement.Text = strSQL & strClinicSQL
        'oStatement.SqlStatement.Text = "SELECT event_id, facility_ncid, event_start_gmtime from event where encounter_number = :X1"
    
        'Bind variables
    '    Set oVars = oStatement.BindVariables
    '    Set oVar = oVars.Add
    '    oVar.Value.InfoType = "textString"
    '    oVar.Value.TextString = "118013"
        
        ExecuteSQLServ moSqlConnect
        'moSqlConnect.ExecuteTransaction
        
            If moSqlConnect.Transaction(1).Rows.Count = 0 Then

                ISqlOps_ValidateClinicAgainstCDR = False

                If NCIDIsCDRClinicRepresentation(ClinicNCID) Then
                    'If it is a representation then need to check again

                    NCIDIsRepresentation = True
                    strClinicSQL = "cr.concept_ncid = " & ClinicNCID & " "
                    
                    moSqlConnect.ClearAll
                    moSqlConnect.ServiceName = msServiceName
                    
                    Set oStatement = moSqlConnect.Transaction.Add
                    oStatement.SqlId.ValueConcept.UniqueId = gSQLQueryNCID
                    oStatement.SqlStatement.Text = strSQL & strClinicSQL
                    ExecuteSQLServ moSqlConnect
                    'moSqlConnect.ExecuteTransaction
                    
                    If moSqlConnect.Transaction(1).Rows.Count = 0 Then
                        'A clinicNCID but not a representation and no matching facility
                        If CreateIfNotFound = True Then
                            'Use original NCID because the representation returned might be for some other facility
                            ClinicNCID = ISqlOps_CreateNewCDRClinicNCID(strOriginalNCID, FacilityNcid, EnterpriseNCID)
                        End If
                    Else 'A Clinic NCID a representation and the facility match up
                        'In this case, do nothing
                        Exit Function
                    End If
                Else 'A Clinic NCID and the facility ncid does not match up and Not a representation
                    NCIDIsRepresentation = False
                    If CreateIfNotFound = True Then
                        ISqlOps_ValidateClinicAgainstCDR = False
                        ClinicNCID = ISqlOps_CreateNewCDRClinicNCID(ClinicNCID, FacilityNcid, EnterpriseNCID)
                        Exit Function
                    End If
                End If
            Else ' A Clinic NCID and facility matches up
                ISqlOps_ValidateClinicAgainstCDR = True
            End If
        
        
'********************* The old DSQL way
'        With mvarDSQLConnect
'            .ClearAll
'            .Table.sqlId.ValueConcept.UniqueId = gSQLQueryNCID
'            .Table.arguments.Add
'            .Table.arguments(1).Text = strSQL & strClinicSQL
'            .Query
'            If .Table.Rows.Count = 0 Then
'
'                ISqlOps_ValidateClinicAgainstCDR = False
'
'                If NCIDIsCDRClinicRepresentation(ClinicNCID) Then
'                    'If it is a representation then need to check again
'
'                    NCIDIsRepresentation = True
'                    strClinicSQL = "cr.concept_ncid = " & ClinicNCID & " "
'                    .ClearAll
'                    .Table.sqlId.ValueConcept.UniqueId = gSQLQueryNCID
'                    .Table.arguments.Add
'                    .Table.arguments(1).Text = strSQL & strClinicSQL
'                    .Query
'                    If .Table.Rows.Count = 0 Then
'                        'A clinicNCID but not a representation and no matching facility
'                        If CreateIfNotFound = True Then
'                            'Use original NCID because the representation returned might be for some other facility
'                            ClinicNCID = ISqlOps_CreateNewCDRClinicNCID(strOriginalNCID, FacilityNcid, EnterpriseNCID)
'                        End If
'                    Else 'A Clinic NCID a representation and the facility match up
'                        'In this case, do nothing
'                        Exit Function
'                    End If
'                Else 'A Clinic NCID and the facility ncid does not match up and Not a representation
'                    NCIDIsRepresentation = False
'                    If CreateIfNotFound = True Then
'                        ISqlOps_ValidateClinicAgainstCDR = False
'                        ClinicNCID = ISqlOps_CreateNewCDRClinicNCID(ClinicNCID, FacilityNcid, EnterpriseNCID)
'                        Exit Function
'                    End If
'                End If
'            Else ' A Clinic NCID and facility matches up
'                ISqlOps_ValidateClinicAgainstCDR = True
'            End If
'        End With
    
    Else 'Not a clinic NCID
    
        ISqlOps_ValidateClinicAgainstCDR = False
        
        'Additional check whether it is a representation
        'If it is then NCID will be changed
        If NCIDIsCDRClinicRepresentation(ClinicNCID) Then
            NCIDIsRepresentation = True
        Else: NCIDIsRepresentation = False
        End If
    
        strSQL = strSQL & "cr.concept_ncid = " & ClinicNCID & " "
        strSQL = strSQL & "AND r.ncid = " & FacilityNcid & ""
        
        
        moSqlConnect.ClearAll
        moSqlConnect.ServiceName = msServiceName
        
        Set oStatement = moSqlConnect.Transaction.Add
        oStatement.SqlId.ValueConcept.UniqueId = gSQLQueryNCID
        oStatement.SqlStatement.Text = strSQL
        ExecuteSQLServ moSqlConnect
        'moSqlConnect.ExecuteTransaction
        
        If moSqlConnect.Transaction(1).Rows.Count = 0 Then
            ISqlOps_ValidateClinicAgainstCDR = False
            If CreateIfNotFound = True Then
                ClinicNCID = ISqlOps_CreateNewCDRClinicNCID(ClinicNCID, FacilityNcid, EnterpriseNCID)
                Exit Function
            End If
        End If
        
'**************** The old DSQL way
'        With mvarDSQLConnect
'            .ClearAll
'            .Table.sqlId.ValueConcept.UniqueId = gSQLQueryNCID
'            .Table.arguments.Add
'            .Table.arguments(1).Text = strSQL
'            .Query
'            If .Table.Rows.Count = 0 Then
'                ISqlOps_ValidateClinicAgainstCDR = False
'                If CreateIfNotFound = True Then
'                    ClinicNCID = ISqlOps_CreateNewCDRClinicNCID(ClinicNCID, FacilityNcid, EnterpriseNCID)
'                    Exit Function
'                End If
'            End If
'        End With
    End If

Exit Function

ErrHandler:
    ISqlOps_ValidateClinicAgainstCDR = False
    Err.Raise Err.Number, "SqlOps.ValidateClinicAgainstCDR", Err.Description
    Err.Clear

End Function

Private Property Get ISqlOps_Value(Column As Variant) As Variant
Dim index As Integer
    
    On Error GoTo ErrHandler
    
    If Not (ISqlOps_BOF Or ISqlOps_EOF) Then
        If IsNumeric(Column) Then
            index = Column
        Else
            index = ISqlOps_ColumnIndex(CStr(Column))
        End If
        
        ISqlOps_Value = SServGetValue(moSqlConnect.Transaction(1).Rows(mvarAbsolutePosition).Item(index), False, meTimeZoneConversion)

'****************** Old DSQL way
'        With mvarDSQLConnect.Table(mvarAbsolutePosition, index).Value
'            Select Case .InfoType
'                Case "textString"
'                    ISqlOps_Value = .TextString
'                Case "binaryString"
'                    ISqlOps_Value = .BinaryString
'                Case "date"
'                    If .Date <> #12/31/9999# Then
'                        ' The DSQL atl component is spuriously adjusting for local time,
'                        ' so we "undo" that adjustment here
'                        ISqlOps_Value = AdjustDate(.Date)
'                    End If
'                Case Else
'            End Select
'        End With
    End If
ErrHandler:

End Property

Private Function CreateNewCDRFacilityNCID(ByVal strFacilityName As String, ByVal EnterpriseNCID As String) As String

    'This function will create a new Facility NCID in CDR.
    'strFacilityName will become the 'representation' of the newly created Facility NCID
    'The function will return the NCID of the newly created Facility NCID

    Dim oHDD As HDDMAINTLib.HDDMaintenance
    Dim strReturnedNCID As String
    
    Set oHDD = New HDDMAINTLib.HDDMaintenance
    With oHDD
        .ClearAll
        .WholeConcepts.Add
        With .WholeConcepts(1)
            .ClearAll
            .Concept.Definition = "A DoD field facility"
            .Concept.Comments = "Created " & Format(Now, "mm/dd/yyyy hh:nn")
            '---create the Profile Domain relationships
            .Relationships.Add
            With .Relationships(1)
                .ClearAll
                .ConceptRelationNcid = NCID_FACILITY
                .RelationshipNcid = NCID_HAS_CHILD
                .EnterpriseNCID = EnterpriseNCID 'objLogon.EnterpriseNCID
            End With
            .Relationships.Add
            With .Relationships(2)
                .ClearAll
                .ConceptRelationNcid = NCID_FACILITY
                .RelationshipNcid = NCID_HAS_MEMBER
                .EnterpriseNCID = EnterpriseNCID 'objLogon.EnterpriseNCID
            End With
            .Relationships.Add
            With .Relationships(3)
                .ClearAll
                .ConceptRelationNcid = NCID_FACILITY
                .RelationshipNcid = NCID_HAS_PART
                .EnterpriseNCID = EnterpriseNCID 'objLogon.EnterpriseNCID
            End With
            '---create a context 2000 surface form whose representation is just the name of the Profile
            With .RSForms
                .ClearAll
                .Add
                With .Item(1)
                    .ClearAll
                    .Representation = LEADSTRING_FACILITY_REPRESENTATION & strFacilityName
                    .Contexts.Add
                    With .Contexts(1)
                        .ContextNcid = 2000
                        .EnterpriseNCID = EnterpriseNCID 'objLogon.EnterpriseNCID
                        .PreferredScore = 0
                    End With
                    .Contexts.Add
                    With .Contexts(2)
                        .ContextNcid = 367
                        .EnterpriseNCID = EnterpriseNCID 'objLogon.EnterpriseNCID
                        .PreferredScore = 0
                    End With
                    .Contexts.Add
                    With .Contexts(3)
                        .ContextNcid = 21915
                        .EnterpriseNCID = EnterpriseNCID 'objLogon.EnterpriseNCID
                        .PreferredScore = 0
                    End With
                    .Contexts.Add
                    With .Contexts(4)
                        .ContextNcid = 21916
                        .EnterpriseNCID = EnterpriseNCID 'objLogon.EnterpriseNCID
                        .PreferredScore = 0
                    End With
                End With
                .Add
                With .Item(2)
                    .ClearAll
                    .Representation = LEADSTRING_FACILITY_REPRESENTATION & strFacilityName 'presumably the ncid of the clinical machine
                    .Contexts.Add
                    With .Contexts(1)
                        .ContextNcid = 2000
                        .EnterpriseNCID = EnterpriseNCID 'objLogon.EnterpriseNCID
                        .PreferredScore = 0
                    End With
                    .Contexts.Add
                    With .Contexts(2)
                        .ContextNcid = 367
                        .EnterpriseNCID = EnterpriseNCID 'objLogon.EnterpriseNCID
                        .PreferredScore = 0
                    End With
                    .Contexts.Add
                    With .Contexts(3)
                        .ContextNcid = 21915
                        .EnterpriseNCID = EnterpriseNCID 'objLogon.EnterpriseNCID
                        .PreferredScore = 0
                    End With
                    .Contexts.Add
                    With .Contexts(4)
                        .ContextNcid = 21916
                        .EnterpriseNCID = EnterpriseNCID 'objLogon.EnterpriseNCID
                        .PreferredScore = 0
                    End With
                End With
            End With
        End With
        .Create
        strReturnedNCID = .WholeConcepts(1).Concept.Ncid
    End With
    
    CreateNewCDRFacilityNCID = strReturnedNCID

End Function

Private Function NCIDIsCDRClinic(ByVal strClinicNCID As String) As Boolean
    
    'This function will check whether the passed-in NCID is a valid
    'CDR Faciltity NCID.  If yes, the function returns 'True'
    
    Dim oConcept As CONCEPTATLLib.ConceptCtrl
    
    Set oConcept = New CONCEPTATLLib.ConceptCtrl
    
    oConcept.ClearAll
    oConcept.UniqueId = CLng(strClinicNCID)
    
    If oConcept.IsA(NCID_CLINIC) = True Then
        NCIDIsCDRClinic = True
    Else
        NCIDIsCDRClinic = False
    End If
        
End Function

Private Function NCIDIsCDRClinicRepresentation(ByRef strInOutClinicNCID As String) As Boolean

    'This function will check whether the passed-in NCID is a representation of
    'an already existed CDR faciltity.  If yes, the function returns true and change
    'value of the passed-in NCID to the NCID of that CDR Facility representation.

    Dim oRToNCID As CONCEPTATLLib.RepToNcid
    
    Set oRToNCID = New CONCEPTATLLib.RepToNcid
    
    With oRToNCID
        .Representation = LEADSTRING_CLINIC_REPRESENTATION & strInOutClinicNCID
        .Domain = NCID_CLINIC
        .Load
        If .Count > 0 Then
            strInOutClinicNCID = .Item(1)
            NCIDIsCDRClinicRepresentation = True
        Else
            NCIDIsCDRClinicRepresentation = False
        End If
    End With
    
    Set oRToNCID = Nothing
    
End Function

Private Function NCIDIsCDRFacility(ByVal strFacilityNCID As String) As Boolean
    
    'This function will check whether the passed-in NCID is a valid
    'CDR Faciltity NCID.  If yes, the function returns 'True'
 
    Dim oConcept As CONCEPTATLLib.ConceptCtrl
    
    Set oConcept = New CONCEPTATLLib.ConceptCtrl
    
    oConcept.ClearAll
    oConcept.UniqueId = CLng(strFacilityNCID)
    
    If oConcept.IsA(NCID_FACILITY) = True Then
        NCIDIsCDRFacility = True
    Else
        NCIDIsCDRFacility = False
    End If
    
'   The above utilization of the ConceptControl is equivalent to the
'   following query
'    strSQL = "SELECT R.NCID, SUBSTR(R.REPRESENTATION,1,75) "
'    strSQL = strSQL & "FROM RSFORM_CONTEXT RC, RSFORM R "
'    strSQL = strSQL & "WHERE RC.CONTEXT_NCID = 2000 AND "
'    strSQL = strSQL & "RC.RSFORM_ID = R.RSFORM_ID AND "
'    strSQL = strSQL & "R.NCID IN "
'        strSQL = strSQL & "(SELECT DISTINCT CONCEPT_NCID "
'        strSQL = strSQL & "FROM CONCEPT_RELATION CR "
'        strSQL = strSQL & "WHERE CR.RELATIONSHIP_NCID = 363 AND "
'        strSQL = strSQL & "CR.CONCEPT_RELATION_NCID = 119) "
'    strSQL = strSQL & "AND R.NCID = " & strFacilityNCID
    
End Function

Private Function NCIDIsCDRFacilityRepresentation(ByRef strInOutFacilityNCID As String) As Boolean

    'This function will check whether the passed-in NCID is a representation of
    'an already existed CDR faciltity.  If yes, the function returns true and change
    'value of the passed-in NCID to the NCID of that CDR Facility representation.

    Dim oRToNCID As CONCEPTATLLib.RepToNcid
    
    Set oRToNCID = New CONCEPTATLLib.RepToNcid
    
    With oRToNCID
        .Representation = LEADSTRING_FACILITY_REPRESENTATION & strInOutFacilityNCID
        .Domain = NCID_FACILITY
        .Load
        If .Count > 0 Then
            strInOutFacilityNCID = .Item(1)
            NCIDIsCDRFacilityRepresentation = True
        Else
            NCIDIsCDRFacilityRepresentation = False
        End If
    End With
    
    Set oRToNCID = Nothing
    

End Function



