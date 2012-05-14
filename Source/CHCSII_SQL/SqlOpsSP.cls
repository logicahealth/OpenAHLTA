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
Attribute VB_Name = "SqlOpsSP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ICHCSII_SQL.ISqlOps


Private Const SQLExecuteTransactionNCID As String = "1450582"
Private Const OraErrorNoDataFound As String = "ORA-00100"
Private Const OraErrorFetchOutOfSequence As String = "ORA-01002"

Private mvarAbsolutePosition As Long

Private oSqlConnect As SQLSERVATLLib.SqlServConnect
Private oStatement As SQLSERVATLLib.SqlServStatement


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

Public Property Get Value(Column As Variant) As Variant
    Dim index As Integer
    On Error GoTo ErrHandler
    If Not EOF Then
        If IsNumeric(Column) Then
            index = Column
        Else
            index = ColumnIndex(CStr(Column))
        End If
        With oStatement.Rows(mvarAbsolutePosition)(index)
            Select Case .InfoType
                Case "textString"
                    Value = .TextString
                Case "binaryString"
                    Value = .BinaryString
                Case "date"
                    If .Date <> #12/31/9999# Then
                        Value = .DateNoTZConversion
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

Public Property Get InfoType(Column As Variant) As String
    Dim index As Integer
    If Not (BOF Or EOF) Then
        If IsNumeric(Column) Then
            index = Column
        Else
            index = ColumnIndex(CStr(Column))
        End If
        InfoType = oStatement.Rows(mvarAbsolutePosition)(index).InfoType
    End If
End Property


Public Property Get ColumnIndex(Name As String) As Integer
    Dim i As Integer
    Dim sText As String
    For i = 1 To oStatement.ColumnNames.Count
        If UCase(oStatement.ColumnNames.Item(i)) = UCase(Name) Then
            ColumnIndex = i
            Exit Property
        End If
    Next i
    sText = "There is no column named '" & Name & "'.  Please select from:"
    For i = 1 To ColumnCount
        sText = sText & Chr(10) & oStatement.ColumnNames(i)
    Next
    Err.Raise vbObjectError + 1, "CHCSII_SQL:SqlOps.ColumnIndex", sText
End Property

Public Property Get ColumnName(index As Integer) As String
    ColumnName = ""
    If Not (BOF Or EOF) Then
        If index > ColumnCount Then
            Err.Raise vbObjectError + 1, "CHCSII_SQL:SqlOps.ColumnName", "Index (" & index & ") out of range.  There are " & ColumnCount & " colummns."
            Exit Property
        End If
        ColumnName = oStatement.ColumnName(index)
    End If
End Property

Public Property Get ColumnCount() As Integer
    ColumnCount = oStatement.ColumnCount
End Property

Public Sub ExecuteMulti(SqlStatements As Collection)
    
End Sub
Public Sub ExecuteSP(Params As CHCSII_Command.CHCSII_Cmd)

    Dim vParam  As CHCSII_Command.CHCSII_SPparam
    Dim oBindVar As SQLSERVATLLib.SqlServBindVar
    Dim i As Integer
    
    oSqlConnect.ClearAll
    
    Set oStatement = oSqlConnect.Transaction.Add
    
    With oStatement
        .ProcedureName = Params.SPName
        For i = 1 To Params.Count
            Set vParam = Params.Item(i)
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
    
    oSqlConnect.ExecuteTransaction
    
    If oSqlConnect.Transaction.Count > 0 Then
        Set oStatement = oSqlConnect.Transaction(1)
    End If
    MoveFirst 'reset the mvarAbsolutePosition
End Sub
Public Sub Execute(SPName As String)

    
End Sub
Public Sub MovePrevious()
    If mvarAbsolutePosition > 0 Then
        mvarAbsolutePosition = mvarAbsolutePosition - 1
    End If
    If mvarAbsolutePosition > 0 Then
        mvarAbsolutePosition = mvarAbsolutePosition - 1
    End If
End Sub

Public Sub MoveNext()
    If mvarAbsolutePosition < RecordCount + 1 Then
        mvarAbsolutePosition = mvarAbsolutePosition + 1
    End If
End Sub

Public Sub MoveLast()
    mvarAbsolutePosition = RecordCount
End Sub

Public Sub MoveFirst()
    If RecordCount > 0 Then
        mvarAbsolutePosition = 1
    Else
        mvarAbsolutePosition = 0
    End If
End Sub

Public Property Get Source() As String
    Source = oStatement.SqlStatement
End Property

Public Property Get RecordCount() As Long
    RecordCount = oStatement.Rows.Count
End Property

Public Property Get EOF() As Boolean
    If RecordCount = 0 Then
        EOF = True
    ElseIf mvarAbsolutePosition = RecordCount + 1 Then
        EOF = True
    Else
        EOF = False
    End If
End Property

Public Property Get BOF() As Boolean
    If mvarAbsolutePosition = 0 Then
        BOF = True
    Else
        BOF = False
    End If
End Property

Private Sub Class_Initialize()
    GetTimeZoneInformation TZInfo
    Set oSqlConnect = New SQLSERVATLLib.SqlServConnect
    mvarAbsolutePosition = 0
End Sub

Private Sub Class_Terminate()
    Set oSqlConnect = Nothing
    Set oStatement = Nothing
End Sub


Private Property Get ISqlOps_BOF() As Boolean:                          ISqlOps_BOF = BOF: End Property
Private Property Get ISqlOps_ColumnCount() As Integer:                  ISqlOps_ColumnCount = ColumnCount: End Property
Private Property Get ISqlOps_ColumnIndex(Name As String) As Integer:    ISqlOps_ColumnIndex = ColumnIndex(Name): End Property
Private Property Get ISqlOps_ColumnName(index As Integer) As String:    ISqlOps_ColumnName = ColumnName(index): End Property



Private Property Get ISqlOps_EOF() As Boolean:                          ISqlOps_EOF = EOF: End Property
Private Function ISqlOps_Execute(SqlStatement As String) As Variant:    Call Execute(SqlStatement): End Function
Private Function ISqlOps_ExecuteCmd(oAdoCmd As ADODB.Command) As Variant: ISqlOps_ExecuteCmd = False: End Function
Private Sub ISqlOps_ExecuteMulti(SqlStatements As Collection):          Call ExecuteMulti(SqlStatements): End Sub
Private Function ISqlOps_GetNextID(Optional ByVal sSeqName As ICHCSII_SQL.eSeq = 0&) As Variant: ISqlOps_GetNextID = 0&: End Function
Private Function ISqlOps_GetNextNcid() As Variant:                      ISqlOps_GetNextNcid = 0: End Function
Private Property Get ISqlOps_InfoType(Column As Variant) As String:     ISqlOps_InfoType = InfoType(Column): End Property
Private Sub ISqlOps_MoveFirst():                                        MoveFirst: End Sub
Private Sub ISqlOps_MoveLast():                                         MoveLast: End Sub
Private Sub ISqlOps_MoveNext():                                         MoveNext: End Sub
Private Sub ISqlOps_MovePrevious():                                     MovePrevious: End Sub


Private Property Get ISqlOps_RecordCount() As Long:                     ISqlOps_RecordCount = RecordCount: End Property
Private Property Get ISqlOps_Source() As String:                        ISqlOps_Source = Source: End Property

Private Function ISqlOps_ValidateFacility(ByVal FacilityNCID As Long, _
                                          ByVal CreateIfNotFound As Boolean, _
                                          Optional ByVal ClinicNCID As Long = 0&) As Boolean: ISqlOps_ValidateFacility = False: End Function
Private Property Get ISqlOps_Value(Column As Variant) As Variant:       ISqlOps_Value = Value(Column): End Property

'<SCR 22426>

'PIC - add a function in CHCSII_SQL to add a new clinic into CDR

Private Function ISqlOps_ValiadeFacilityAgainstCDR(FacilityNCID As String, ByVal EnterpriseNCID As String, NCIDIsRepresentation As Boolean, ByVal CreateIfNotFound As Boolean) As Boolean
    ISqlOps_ValiadeFacilityAgainstCDR = ValiadeFacilityAgainstCDR(FacilityNCID, EnterpriseNCID, NCIDIsRepresentation, CreateIfNotFound)
End Function

Private Function ISqlOps_ValidateClinicAgainstCDR(ClinicNCID As String, ByVal FacilityNCID As String, FacilityIsValid As Boolean, NCIDIsRepresentation As Boolean, ByVal EnterpriseNCID As String, ByVal CreateIfNotFound As Boolean) As Boolean
    ISqlOps_ValidateClinicAgainstCDR = ValidateClinicAgainstCDR(ClinicNCID, FacilityNCID, FacilityIsValid, NCIDIsRepresentation, EnterpriseNCID, CreateIfNotFound)
End Function
'</SCR 22426>
Private Function ISqlOps_ExecuteSP(Parms As CHCSII_Command.CHCSII_Cmd) As Variant
    Call ExecuteSP(Parms)
End Function



Private Function ValiadeFacilityAgainstCDR(FacilityNCID As String, ByVal EnterpriseNCID As String, NCIDIsRepresentation As Boolean, ByVal CreateIfNotFound As Boolean) As Boolean

    On Error GoTo ErrHandler

    If NCIDIsCDRFacility(FacilityNCID) Then
        ValiadeFacilityAgainstCDR = True
        Exit Function
    Else
        ValiadeFacilityAgainstCDR = False
        If NCIDIsCDRFacilityRepresentation(FacilityNCID) Then
            'FacilityNCID will be changed to the one it is representing
            NCIDIsRepresentation = True
            Exit Function
        Else
            NCIDIsRepresentation = False
        End If
    End If

    ValiadeFacilityAgainstCDR = False
    
    If CreateIfNotFound = True Then
        FacilityNCID = CreateNewCDRFacilityNCID(FacilityNCID, EnterpriseNCID)
    End If

Exit Function

ErrHandler:
    ValiadeFacilityAgainstCDR = False
    Err.Raise Err.Number, "SqlOps.ValiadeFacilityAgainstCDR", Err.Description
    Err.Clear
End Function

'<SCR 22426>

Private Function ValidateClinicAgainstCDR(ClinicNCID As String, ByVal FacilityNCID As String, FacilityIsValid As Boolean, NCIDIsRepresentation As Boolean, ByVal EnterpriseNCID As String, ByVal CreateIfNotFound As Boolean) As Boolean


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

Private Function ISQLOpsSP_ExecuteSP(Params As CHCSII_Command.CHCSII_Cmd) As Variant
    Call ExecuteSP(Params)
End Function
