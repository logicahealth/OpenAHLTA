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
Attribute VB_Name = "CDR_SQLOPS_SP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Implements ICHCSII_SQL.ISqlOps

Private mvarAbsolutePosition As Long
Private meTimeZoneConversion As eTimeZoneConversion

Private moSqlConnect As SQLSERVATLLib.SqlServConnect
Private moStatement As SQLSERVATLLib.SqlServStatement
Private msServiceName As String

Public Sub Init(TimeZoneConversion As eTimeZoneConversion, ServiceName As eServiceNameMode)

    'SF 1-22-04 Indicates if dates should be converted for the local time zone.
    meTimeZoneConversion = TimeZoneConversion
    
    'SCR 46007
    'Developer: Scott Foster 12-6-03
    'Created to expose the service name without modifying ICHCSII_DAS
    
    msServiceName = TranslateServiceName(ServiceName)

End Sub

Public Property Get Value(Column As Variant) As Variant

    Dim index As Integer
    On Error GoTo ErrHandler
    If Not ISqlOps_EOF Then
        If IsNumeric(Column) Then
            index = Column
        Else
            index = ISqlOps_ColumnIndex(CStr(Column))
        End If
        Value = SServGetValue(moStatement.Rows(mvarAbsolutePosition)(index), True, meTimeZoneConversion)
        
'SF 1-29-04 replaced by SServGetValue
'        With moStatement.Rows(mvarAbsolutePosition)(index)
'            Select Case .InfoType
'                Case "textString"
'                    Value = .TextString
'                Case "binaryString"
'                    Value = .BinaryString
'                Case "date"
'                    If .Date <> #12/31/9999# Then
'                        Value = .DateNoTZConversion
'                    End If
'                Case Else
'            End Select
'        End With

    End If
ErrHandler:
End Property



Public Property Get Source() As String
    Source = moStatement.SqlStatement
End Property

Public Property Get RecordCount() As Long
    RecordCount = moStatement.Rows.Count
End Property

Private Sub Class_Initialize()

    'SF 1-29-04 Commented out becuase it is not used
    'GetTimeZoneInformation gTZInfo
    
    Set moSqlConnect = New SQLSERVATLLib.SqlServConnect
    mvarAbsolutePosition = 0
    
End Sub

Private Sub Class_Terminate()
    Set moSqlConnect = Nothing
    Set moStatement = Nothing
End Sub


Private Property Get ISqlOps_BOF() As Boolean
    
    If mvarAbsolutePosition = 0 Then
        ISqlOps_BOF = True
    Else
        ISqlOps_BOF = False
    End If

End Property
Private Property Get ISqlOps_ColumnCount() As Integer
    ISqlOps_ColumnCount = moStatement.ColumnCount
End Property
Private Property Get ISqlOps_ColumnIndex(Name As String) As Integer
    Dim i As Integer
    Dim sText As String
    For i = 1 To moStatement.ColumnNames.Count
        If UCase(moStatement.ColumnNames.Item(i)) = UCase(Name) Then
            ISqlOps_ColumnIndex = i
            Exit Property
        End If
    Next i
    sText = "There is no column named '" & Name & "'.  Please select from:"
    For i = 1 To ISqlOps_ColumnCount
        sText = sText & Chr(10) & moStatement.ColumnNames(i)
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
        ISqlOps_ColumnName = moStatement.ColumnName(index)
    End If
End Property
Private Property Get ISqlOps_EOF() As Boolean
    If RecordCount = 0 Then
        ISqlOps_EOF = True
    ElseIf mvarAbsolutePosition = ISqlOps_RecordCount + 1 Then
        ISqlOps_EOF = True
    Else
        ISqlOps_EOF = False
    End If
End Property
Private Function ISqlOps_Execute(SqlStatement As String) As Variant
    Err.Raise 1001, "CDR_SQLOPSSP", "Execute is not supported in CDR_SQLOPSSP."
End Function
Private Function ISqlOps_ExecuteCmd(oAdoCmd As ADODB.Command) As Variant
    ISqlOps_ExecuteCmd = False
End Function
Private Sub ISqlOps_ExecuteMulti(SqlStatements As Collection)

    Err.Raise 1001, "CDR_SQLOPSSP", "ExecuteMulti is not supported in CDR_SQLOPSSP."
    
End Sub

Private Function ISqlOps_ExecuteSP(Parms As CHCSII_Command.CHCSII_Cmd) As Variant

'</SCR 22426>
    
    Dim vParam  As CHCSII_SPparam
    Dim oBindVar As SQLSERVATLLib.SqlServBindVar
    Dim i As Integer
    
    moSqlConnect.ClearAll
    
    moSqlConnect.ServiceName = msServiceName
    
    Set moStatement = moSqlConnect.Transaction.Add
    
    With moStatement
        .ProcedureName = Parms.SPName
        For i = 1 To Parms.Count
            Set vParam = Parms.Item(i)
            Set oBindVar = .BindVariables.Add
            If Not vParam.InfoType = "" Then
                oBindVar.Value.InfoType = vParam.InfoType
                If vParam.InfoType = "date" Then
                    oBindVar.Value.DateNoTZConversion = vParam.Param
                ElseIf vParam.InfoType = "textString" Then
                    oBindVar.Value.TextString = vParam.Param
                End If
            End If
            oBindVar.IsCursor = vParam.IsCursor
            oBindVar.IsOutput = vParam.IsOutput
        Next
    End With
    
    moSqlConnect.ExecuteTransaction
    
    If moSqlConnect.Transaction.Count > 0 Then
        Set moStatement = moSqlConnect.Transaction(1)
    End If
    ISqlOps_MoveFirst 'reset the mvarAbsolutePosition

End Function

Private Function ISqlOps_GetNextID(Optional ByVal sSeqName As ICHCSII_SQL.eSeq = 0&) As Variant
    ISqlOps_GetNextID = 0&
End Function
Private Function ISqlOps_GetNextNcid() As Variant
    ISqlOps_GetNextNcid = 0
End Function
Private Property Get ISqlOps_InfoType(Column As Variant) As String
    Dim index As Integer
    If Not (ISqlOps_BOF Or ISqlOps_EOF) Then
        If IsNumeric(Column) Then
            index = Column
        Else
            index = ISqlOps_ColumnIndex(CStr(Column))
        End If
        ISqlOps_InfoType = moStatement.Rows(mvarAbsolutePosition)(index).InfoType
    End If
End Property
Private Sub ISqlOps_MoveFirst()
    If RecordCount > 0 Then
        mvarAbsolutePosition = 1
    Else
        mvarAbsolutePosition = 0
    End If
End Sub
Private Sub ISqlOps_MoveLast()
    mvarAbsolutePosition = RecordCount
End Sub
Private Sub ISqlOps_MoveNext()
    If mvarAbsolutePosition < RecordCount + 1 Then
        mvarAbsolutePosition = mvarAbsolutePosition + 1
    End If
End Sub
Private Sub ISqlOps_MovePrevious()
    If mvarAbsolutePosition > 0 Then
        mvarAbsolutePosition = mvarAbsolutePosition - 1
    End If
    If mvarAbsolutePosition > 0 Then
        mvarAbsolutePosition = mvarAbsolutePosition - 1
    End If
End Sub


Private Property Get ISqlOps_RecordCount() As Long
    ISqlOps_RecordCount = RecordCount
End Property
Private Property Get ISqlOps_Source() As String
    ISqlOps_Source = Source
End Property

Private Function ISqlOps_ValidateFacility(ByVal FacilityNcid As Long, _
                                          ByVal CreateIfNotFound As Boolean, _
                                          Optional ByVal ClinicNCID As Long = 0&) As Boolean
    ISqlOps_ValidateFacility = False

End Function
Private Property Get ISqlOps_Value(Column As Variant) As Variant
    ISqlOps_Value = Value(Column)
End Property

'<SCR 22426>

'PIC - add a function in CHCSII_SQL to add a new clinic into CDR

Private Function ISqlOps_ValiadeFacilityAgainstCDR(FacilityNcid As String, ByVal EnterpriseNCID As String, NCIDIsRepresentation As Boolean, ByVal CreateIfNotFound As Boolean) As Boolean
    
    ISqlOps_ValiadeFacilityAgainstCDR = ValiadeFacilityAgainstCDR(FacilityNcid, EnterpriseNCID, NCIDIsRepresentation, CreateIfNotFound)
    
End Function

Private Function ISqlOps_ValidateClinicAgainstCDR(ClinicNCID As String, ByVal FacilityNcid As String, FacilityIsValid As Boolean, NCIDIsRepresentation As Boolean, ByVal EnterpriseNCID As String, ByVal CreateIfNotFound As Boolean) As Boolean
    
    ISqlOps_ValidateClinicAgainstCDR = ValidateClinicAgainstCDR(ClinicNCID, FacilityNcid, FacilityIsValid, NCIDIsRepresentation, EnterpriseNCID, CreateIfNotFound)
    
End Function
Private Function ValiadeFacilityAgainstCDR(FacilityNcid As String, ByVal EnterpriseNCID As String, NCIDIsRepresentation As Boolean, ByVal CreateIfNotFound As Boolean) As Boolean

    On Error GoTo ErrHandler

    If NCIDIsCDRFacility(FacilityNcid) Then
        ValiadeFacilityAgainstCDR = True
        Exit Function
    Else
        ValiadeFacilityAgainstCDR = False
        If NCIDIsCDRFacilityRepresentation(FacilityNcid) Then
            'FacilityNCID will be changed to the one it is representing
            NCIDIsRepresentation = True
            Exit Function
        Else
            NCIDIsRepresentation = False
        End If
    End If

    ValiadeFacilityAgainstCDR = False
    
    If CreateIfNotFound = True Then
        FacilityNcid = CreateNewCDRFacilityNCID(FacilityNcid, EnterpriseNCID)
    End If

Exit Function

ErrHandler:
    ValiadeFacilityAgainstCDR = False
    Err.Raise Err.Number, "SqlOps.ValiadeFacilityAgainstCDR", Err.Description
    Err.Clear
End Function

'<SCR 22426>

Private Function ValidateClinicAgainstCDR(ClinicNCID As String, ByVal FacilityNcid As String, FacilityIsValid As Boolean, NCIDIsRepresentation As Boolean, ByVal EnterpriseNCID As String, ByVal CreateIfNotFound As Boolean) As Boolean

    Err.Raise 1001, "CDR_SQLOPSSP", "ValidateClinicAgainstCDR is not supported in CDR_SQLOPSSP."

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

Public Function CreateNewCDRClinicNCID(ByVal strClinicName As String, ByVal strFacilityNCID As String, ByVal strEnterpriseNCID As String) As String

    'This function will create a new Clinic NCID in CDR.
    'strClinicName will become the 'representation' of the newly created Facility NCID
    'The function will return the NCID of the newly created Facility NCID

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
    
    CreateNewCDRClinicNCID = strReturnedNCID

Exit Function

ErrHandler:
    Err.Raise Err.Number, "SqlOps.CreateNewCDRFacilityNCID", Err.Description
    Err.Clear
End Function
'</SCR 22426>

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
'   following ExecuteTransaction
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

