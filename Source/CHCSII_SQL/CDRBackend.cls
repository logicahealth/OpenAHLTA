VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDRBackend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements iBackend

Private Const SQLQueryNCID As String = "1450582"
Private Const OraErrorNoDataFound As String = "ORA-00100"
Private Const OraErrorFetchOutOfSequence As String = "ORA-01002"

Private mvarAbsolutePosition As Long
Private mvarDSQLConnect As DSQLATLLib.DSqlConnect

'These NCIDs were copied from ncidcon.bas on 3/6/2002
Private Const NCID_FACILITY As String = "119"
Private Const NCID_CLINIC As String = "1003"
Private Const NCID_HAS_MEMBER As String = "363"
Private Const NCID_HAS_CHILD As String = "364"
Private Const NCID_HAS_PART As String = "30737"
Private Const LEADSTRING_FACILITY_REPRESENTATION As String = "THEATER FACILITY: "
Private Const LEADSTRING_CLINIC_REPRESENTATION As String = "THEATER CLINIC: "

Private TZInfo As TIME_ZONE_INFORMATION

Private Declare Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long

Private Sub Class_Initialize()
    GetTimeZoneInformation TZInfo
    Set mvarDSQLConnect = New DSQLATLLib.DSqlConnect
    mvarAbsolutePosition = 0
End Sub

Private Sub Class_Terminate()
    Set mvarDSQLConnect = Nothing
End Sub

Private Property Get iBackend_BOF() As Boolean
    If mvarAbsolutePosition = 0 Then
        iBackend_BOF = True
    Else
        iBackend_BOF = False
    End If
End Property

Private Property Get iBackend_ColumnCount() As Integer
    iBackend_ColumnCount = mvarDSQLConnect.Table.Rows(mvarAbsolutePosition).Columns.Count
End Property

Private Property Get iBackend_ColumnIndex(Name As String) As Integer
Dim i As Integer
Dim sText As String
    
    For i = 1 To iBackend_ColumnCount
        If UCase(iBackend_ColumnName(i)) = UCase(Name) Then
            iBackend_ColumnIndex = i
            Exit Property
        End If
    Next i
    sText = "There is no column named '" & Name & "'.  Please select from:"
    For i = 1 To iBackend_ColumnCount
        sText = sText & Chr(10) & mvarDSQLConnect.Table(mvarAbsolutePosition, i).Name
    Next
    
    Err.Raise vbObjectError + 1, "CHCSII_SQL:SqlOps.ColumnIndex", sText

End Property

Private Property Get iBackend_ColumnName(ByVal index As Integer) As String
    
    iBackend_ColumnName = ""
    
    If Not (iBackend_BOF Or iBackend_EOF) Then
        If index > iBackend_ColumnCount Then
            Err.Raise vbObjectError + 1, "CHCSII_SQL:SqlOps.ColumnName", "Index (" & index & ") out of range.  There are " & iBackend_ColumnCount & " colummns."
            Exit Property
        End If
        iBackend_ColumnName = mvarDSQLConnect.Table(mvarAbsolutePosition, index).Name
    End If

End Property

'This function will create a new Clinic NCID in CDR.
'strClinicName will become the 'representation' of the newly created Facility NCID
'The function will return the NCID of the newly created Facility NCID
Private Function iBackend_CreateNewCDRClinicNCID(ByVal strClinicName As String, ByVal strFacilityNCID As String, ByVal strEnterpriseNCID As String) As String
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
    
    iBackend_CreateNewCDRClinicNCID = strReturnedNCID

Exit Function

ErrHandler:
    Err.Raise Err.Number, "SqlOps.CreateNewCDRFacilityNCID", Err.Description
    Err.Clear

End Function

Private Property Get iBackend_EOF() As Boolean
    If iBackend_RecordCount = 0 Then
        iBackend_EOF = True
    ElseIf mvarAbsolutePosition = iBackend_RecordCount + 1 Then
        iBackend_EOF = True
    Else
        iBackend_EOF = False
    End If
End Property

Private Sub iBackend_Execute(SqlStatement As String)
Dim sOraError As String
    
    On Error GoTo CheckErr
    
    With mvarDSQLConnect
        .ClearAll
        .Table.sqlId.ValueConcept.UniqueId = SQLQueryNCID
        .Table.arguments.Add
        .Table.arguments(1).Text = SqlStatement
        .Query
CheckErr:
        If Len(.Table.Error.ErrorText) <> 0 Then
            On Error Resume Next
            sOraError = Left(.Table.Error.ErrorText, 9)
            Select Case sOraError
            Case OraErrorNoDataFound
                ' do nothing
            Case Else
                Debug.Print "CHCSII_SQL:SqlOps.Execute(" & SqlStatement & ")", "Error Text = " & .Table.Error.ErrorText
                Err.Raise vbObjectError + 1, "CHCSII_SQL:SqlOps.Execute(" & SqlStatement & ")", "Error Text = " & .Table.Error.ErrorText
                Exit Sub
            End Select
        ElseIf Err Then
            Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
        End If
    End With
    
    iBackend_MoveFirst

End Sub

Private Sub iBackend_ExecuteMulti(SqlStatements As Collection)
Dim i As Integer
Dim sOraError As String
    
    With mvarDSQLConnect
        .ClearAll
        .Table.sqlId.ValueConcept.UniqueId = SQLQueryNCID
        For i = 1 To SqlStatements.Count
            .Table.arguments.Add
            .Table.arguments(i).Text = SqlStatements(i)
        Next i
        .Query
        If Len(.Table.Error.ErrorText) <> 0 Then
            sOraError = Left(.Table.Error.ErrorText, 9)
            Select Case sOraError
            Case OraErrorNoDataFound
                ' do nothing
            Case Else
                Debug.Print "CHCSII_SQL:SqlOps.Execute(" & .Table.SqlCString & ")", "Error Text = " & .Table.Error.ErrorText
                Err.Raise vbObjectError + 1, "CHCSII_SQL:SqlOps.Execute(" & .Table.SqlCString & ")", "Error Text = " & .Table.Error.ErrorText
                Exit Sub
            End Select
        End If
    End With
    
    iBackend_MoveFirst

End Sub

Private Property Get iBackend_InfoType(Column As Variant) As String
Dim index As Integer
    
    If Not (iBackend_BOF Or iBackend_EOF) Then
        If IsNumeric(Column) Then
            index = Column
        Else
            index = iBackend_ColumnIndex(CStr(Column))
        End If
        iBackend_InfoType = mvarDSQLConnect.Table(mvarAbsolutePosition, index).Value.InfoType
    End If

End Property

Private Sub iBackend_MoveFirst()
    If iBackend_RecordCount > 0 Then
        mvarAbsolutePosition = 1
    Else
        mvarAbsolutePosition = 0
    End If
End Sub

Private Sub iBackend_MoveLast()
    mvarAbsolutePosition = iBackend_RecordCount
End Sub

Private Sub iBackend_MoveNext()
    If mvarAbsolutePosition < iBackend_RecordCount + 1 Then
        mvarAbsolutePosition = mvarAbsolutePosition + 1
    End If
End Sub

Private Sub iBackend_MovePrevious()
    If mvarAbsolutePosition > 0 Then
        mvarAbsolutePosition = mvarAbsolutePosition - 1
    End If
End Sub

Private Property Get iBackend_RecordCount() As Long
    iBackend_RecordCount = mvarDSQLConnect.Table.Rows.Count
End Property

Private Property Get iBackend_Source() As String
    iBackend_Source = mvarDSQLConnect.Table.SqlCString
End Property

Private Function iBackend_ValiadeFacilityAgainstCDR(FacilityNCID As String, ByVal EnterpriseNCID As String, NCIDIsRepresentation As Boolean, ByVal CreateIfNotFound As Boolean) As Boolean
    On Error GoTo ErrHandler

    If NCIDIsCDRFacility(FacilityNCID) Then
        iBackend_ValiadeFacilityAgainstCDR = True
        Exit Function
    Else
        iBackend_ValiadeFacilityAgainstCDR = False
        If NCIDIsCDRFacilityRepresentation(FacilityNCID) Then
            'FacilityNCID will be changed to the one it is representing
            NCIDIsRepresentation = True
            Exit Function
        Else
            NCIDIsRepresentation = False
        End If
    End If

    iBackend_ValiadeFacilityAgainstCDR = False
    
    If CreateIfNotFound = True Then
        FacilityNCID = CreateNewCDRFacilityNCID(FacilityNCID, EnterpriseNCID)
    End If

Exit Function

ErrHandler:
    iBackend_ValiadeFacilityAgainstCDR = False
    Err.Raise Err.Number, "SqlOps.ValiadeFacilityAgainstCDR", Err.Description
    Err.Clear

End Function

Private Function iBackend_ValidateClinicAgainstCDR(ClinicNCID As String, ByVal FacilityNCID As String, FacilityIsValid As Boolean, NCIDIsRepresentation As Boolean, ByVal EnterpriseNCID As String, ByVal CreateIfNotFound As Boolean) As Boolean
    Dim strSQL As String
    Dim strClinicSQL As String
    Dim strOriginalNCID As String
    
    'first validate the facility sent in
    If NCIDIsCDRFacility(FacilityNCID) Then
        FacilityIsValid = True
    Else
        FacilityIsValid = False
        iBackend_ValidateClinicAgainstCDR = False
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
    strSQL = strSQL & "r.ncid = " & FacilityNCID & " and "
    
    
    If NCIDIsCDRClinic(ClinicNCID) = True Then
        'Still need to check whether there is a "matching" FacilityNCID
        
        strClinicSQL = "cr.concept_ncid = " & ClinicNCID & " "
        
        
        With mvarDSQLConnect
            .ClearAll
            .Table.sqlId.ValueConcept.UniqueId = SQLQueryNCID
            .Table.arguments.Add
            .Table.arguments(1).Text = strSQL & strClinicSQL
            .Query
            If .Table.Rows.Count = 0 Then
                
                iBackend_ValidateClinicAgainstCDR = False
                
                If NCIDIsCDRClinicRepresentation(ClinicNCID) Then
                    'If it is a representation then need to check again
                    
                    NCIDIsRepresentation = True
                    strClinicSQL = "cr.concept_ncid = " & ClinicNCID & " "
                    .ClearAll
                    .Table.sqlId.ValueConcept.UniqueId = SQLQueryNCID
                    .Table.arguments.Add
                    .Table.arguments(1).Text = strSQL & strClinicSQL
                    .Query
                    If .Table.Rows.Count = 0 Then
                        'A clinicNCID but not a representation and no matching facility
                        If CreateIfNotFound = True Then
                            'Use original NCID because the representation returned might be for some other facility
                            ClinicNCID = iBackend_CreateNewCDRClinicNCID(strOriginalNCID, FacilityNCID, EnterpriseNCID)
                        End If
                    Else 'A Clinic NCID a representation and the facility match up
                        'In this case, do nothing
                        Exit Function
                    End If
                Else 'A Clinic NCID and the facility ncid does not match up and Not a representation
                    NCIDIsRepresentation = False
                    If CreateIfNotFound = True Then
                        iBackend_ValidateClinicAgainstCDR = False
                        ClinicNCID = iBackend_CreateNewCDRClinicNCID(ClinicNCID, FacilityNCID, EnterpriseNCID)
                        Exit Function
                    End If
                End If
            Else ' A Clinic NCID and facility matches up
                iBackend_ValidateClinicAgainstCDR = True
            End If
        End With
    
    Else 'Not a clinic NCID
    
        iBackend_ValidateClinicAgainstCDR = False
        
        'Additional check whether it is a representation
        'If it is then NCID will be changed
        If NCIDIsCDRClinicRepresentation(ClinicNCID) Then
            NCIDIsRepresentation = True
        Else: NCIDIsRepresentation = False
        End If
    
        strSQL = strSQL & "cr.concept_ncid = " & ClinicNCID & " "
        strSQL = strSQL & "AND r.ncid = " & FacilityNCID & ""
        
        With mvarDSQLConnect
            .ClearAll
            .Table.sqlId.ValueConcept.UniqueId = SQLQueryNCID
            .Table.arguments.Add
            .Table.arguments(1).Text = strSQL
            .Query
            If .Table.Rows.Count = 0 Then
                iBackend_ValidateClinicAgainstCDR = False
                If CreateIfNotFound = True Then
                    ClinicNCID = iBackend_CreateNewCDRClinicNCID(ClinicNCID, FacilityNCID, EnterpriseNCID)
                    Exit Function
                End If
            End If
        End With
    End If

Exit Function

ErrHandler:
    iBackend_ValidateClinicAgainstCDR = False
    Err.Raise Err.Number, "SqlOps.ValiadeClinicAgainstCDR", Err.Description
    Err.Clear

End Function

Private Property Get iBackend_Value(Column As Variant) As Variant
Dim index As Integer
    
    On Error GoTo ErrHandler
    
    If Not (iBackend_BOF Or iBackend_EOF) Then
        If IsNumeric(Column) Then
            index = Column
        Else
            index = iBackend_ColumnIndex(CStr(Column))
        End If
        With mvarDSQLConnect.Table(mvarAbsolutePosition, index).Value
            Select Case .InfoType
                Case "textString"
                    iBackend_Value = .TextString
                Case "binaryString"
                    iBackend_Value = .BinaryString
                Case "date"
                    If .Date <> #12/31/9999# Then
                        ' The DSQL atl component is spuriously adjusting for local time,
                        ' so we "undo" that adjustment here
                        iBackend_Value = AdjustDate(.Date)
                    End If
                Case Else
            End Select
        End With
    End If
ErrHandler:

End Property

Private Function AdjustDate(ByVal dDate As Date) As Date
    Static lYear As Integer
    Static dStandard As Date
    Static dDaylight As Date
    Dim lDay As Integer
    Dim lTimeBias As Long
    If Year(dDate) <> lYear Then
        lYear = Year(dDate)
        dStandard = GetTZInfoDate(TZInfo.StandardDate, lYear)
        dDaylight = GetTZInfoDate(TZInfo.DaylightDate, lYear)
    End If
    If dStandard > dDaylight Then
        If dDate >= dDaylight And dDate < dStandard Then
            lTimeBias = TZInfo.Bias + TZInfo.DaylightBias
        Else
            lTimeBias = TZInfo.Bias + TZInfo.StandardBias
        End If
    Else
        If dDate >= dStandard And dDate < dDaylight Then
            lTimeBias = TZInfo.Bias + TZInfo.StandardBias
        Else
            lTimeBias = TZInfo.Bias + TZInfo.DaylightBias
        End If
    End If
    AdjustDate = DateAdd("n", lTimeBias, dDate)
End Function

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

Private Function GetTZInfoDate(TZDate As SYSTEMTIME, TZYear As Integer) As Date
    With TZDate
        If .wDay = 5 Then
            ' means last .wDayOfWeek in month
            GetTZInfoDate = DateSerial(TZYear, .wMonth + 1, 0) + TimeSerial(.wHour, .wMinute, .wSecond)
            Do While Weekday(GetTZInfoDate) <> .wDayOfWeek + 1
                GetTZInfoDate = DateAdd("d", -1, GetTZInfoDate)
            Loop
        Else
            ' means .wday'TH .wdayofweek in month
            GetTZInfoDate = DateSerial(TZYear, .wMonth, 1) + TimeSerial(.wHour, .wMinute, .wSecond)
            Do While Weekday(GetTZInfoDate) <> .wDayOfWeek + 1
                GetTZInfoDate = DateAdd("d", 1, GetTZInfoDate)
            Loop
            GetTZInfoDate = DateAdd("d", (.wDay - 1) * 7, GetTZInfoDate)
        End If
    End With
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
