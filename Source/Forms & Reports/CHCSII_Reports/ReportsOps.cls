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
Attribute VB_Name = "ReportOps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*****************************************************************************************************************
'  Class Name: CustomReportOps
'
'  Description: This class manages the creation of all the Custom Reports for CHCS II.
'               Selection Criteria, etc is passed in and appropriate queries are called
'               to fill in the data.
'
'  Properties:
'
'  Methods:
'
'*****************************************************************************************************************
Option Explicit

Private moReport As CRAXDRT.Report
Private moColRecSet As Collection
Private moRS As ADODB.Recordset


Private moHeaderColl As Collection

Private moTableOps As New TableOps
Private moReportItem As ReportItem
Private moReportInfo As ReportInformation

Private Const cDBName = "CHCSII.MDB" '-- Report database Name (Access)
Private Const cDBPath = "\Reports\"          '-- Subdirectory location where database is located
Private Const cRepPath = "\Reports\"         '-- Subdirectory location where reports are located

Private sReportsPath As String

Private msDatabase As String         '-- Full path and filename of database


'-- Identifieds the Report type
Public Enum ReportCategory
    Customized = 0
    Preventive = 1
    Standard = 2
    Forms = 3
    Dental = 4
    CPG = 5
    SRTS = 6
    POP_HEALTH = 7
    DOEHRSHC = 8
End Enum

Public Enum ReportType
    NoReport = 0

    '-- Custom
    DiagnosisRep = 1
    PrimaryDiagnosisRep = 2
    ProcedureRep = 3
    LabRep = 4
    RadRep = 5
    MedRep = 6
    DispositionRep = 7
    AppointmentRep = 8
    EandMRep = 9

    '-- Preventive
    AlcoholQueryforPotentialHeavyUse = 23
    
    DiscontinuedServicesRisk = 12
    
    CholesterolScreening = 11
    HighCholesterolHDL_LDLFollowUp = 16
    HighCholesterolFollowUpCounseling = 17
    HighCholesterolFollowUpRepeat = 18


    SafeSexCounselingQueried_AdvisedofSTDs = 24

    TobaccoUseScreening = 25
    TobaccoUsersAdvisingToQuit = 10

    MammographyQueryFor = 21
    MammographyHaveNot = 27

    PAPSmearQueryFor = 22
    PAPSmearHaveNot = 28

    PreventionReportCard = 29
    AlphabeticPatientList = 30

    DTPImmunizations2yearOlds = 13
    HebBHibImmunizations2YearOlds = 15

    ImmunizationsActiveDutyA2Dose = 20
    Immunizations13yearOlds = 19
    ImmunizationsActiveDutyA1Dose = 14

    Varicellafor2yrolds = 26
    
    PatientRiskReport = 32
    PCMRiskReport = 33
    
    ADSDiagnosis = 31
    ADSClinicOutpatientWorkload = 35
    ADSEncountersAppointmentStatus = 37
    ADSInsuranceIndicator = 39
    ADSPatientCategoriesDisposition = 40
    ADSPatientCategoriesProvider = 41
    ADSInsuranceIndicatorNotMarked = 42
    ADSProcedures = 43

    Consults = 44
    CPSSummary = 45
    CPSServicesDue = 46
    
    ThirdPartyCollections = 47
    DD2766 = 48
    
    Readiness = 50
    
    ADSPatientEncounter = 51
    ADSClinicInpatientWorkload = 52
    
    '-- Forms
    DD2813 = 1000
    DD689 = 1001
    SF522 = 1002
    DD741 = 1003
    DD771 = 1004
    
    '-- Dental 'Wilson2
    SF603 = 2000
    DD2322 = 2001
    DailyProvider = 2002
    MonthlyProvider = 2003
    UnsignedEncounters = 2004
    PerioPatientSummary = 2005 '12/13/01
    
    '-- CPG
    CPG_RegistryCandidate = 3000
    CPG_FollowUpCareLetter = 3001
    CPG_DelinquentQuestionnaire = 3002
    CPG_QuestionnaireInitial = 3003
    CPG_QuestionnaireEach = 3004
    CPG_QuestionnaireAnnual = 3005
    CPG_Goals = 3006
    CPG_ClinicVisitTypeCount = 3007
    CPG_Complications = 3008
    CPG_Compliance = 3009
    CPG_POC = 3010
    CPG_QuestionnaireData = 3011

    '-- SRTS
    SRTS_OrderedLenses = 4000
    SRTS_Rejection = 4001
    SRTS_OrderedFrames = 4002
    SRTS_Overdue10days = 4003
    SRTS_Overdue20days = 4004
    SRTS_Overdue30days = 4005
    SRTS_Overdue = 4006
    SRTS_PendingOrders = 4007
    
    
    '-- Audit
    AUDIT_AllergyVerified = 5000
    
     'Pop_Health
    ProviderRepeatVisitReport = 7010
    ProviderRepeatVisitList = 7011
    Top20DrugsByCostGraph = 7062
    Top20DrugsByProviderGraph = 7060
    ProviderDrugComparison = 7066
    Top20DrugsByVolumeTabular = 7061
    Top20DrugsByCostAverageGraph = 7081
    Top20DrugsByCostAverageGraphAndTabular = 7063
End Enum

Public Enum CriteriaType
    NoCriteria = 0
    MTF = 1
    Clinic = 2
    Provider = 3
    ProviderMulti = 13
    Diagnosis = 4
    Procedure = 5
    Med = 6
    Lab = 7
    Rad = 8
    PCM = 9
    PCMMulti = 14
    Patient = 10
    AppointmentStatus = 11
    Disposition = 12
    Questionnaire = 15
    QuestionnaireAnswer = 16
    Catchment = 17
    ScreeningServices = 18
    Top20Drug = 19
    Wellness = 20
    WellnessMulti = 21
    MTFDMIS = 22
    CPGid = 23
    ClinicSingleSelect = 24
    OrderNumber = 25
End Enum

Private Function DataGetADSMultiPatientEncounter(HeaderColl As Collection) As Boolean
    
    If Not DebugMode Then On Error GoTo ErrHandler
    DataGetADSMultiPatientEncounter = False

    Dim oSQL As ICHCSII_SQL.ISqlOps
    
    Set oSQL = goDBConn.CHCSII_SQLOPS(Auto)
    
    Dim sSql As String
    Dim sStmt As String
    
    Dim i As Integer
    Dim rs As ADODB.Recordset
    
    sStmt = "select distinct n.first_name, n.last_name, n.middle_name, i.ssn, a.patient_unit_number " & vbLf & _
               "from mmi_name n, mmi_id i, appointment a" & vbLf & _
               "where n.unit_number = i.unit_number and" & vbLf & _
               "i.unit_number = a.patient_unit_number and" & vbLf & _
               "(a.appt_date_time >= " & GetOracleStartDate(moReportItem.DateFrom) & vbLf & _
               "and a.appt_date_time <= " & GetOracleEndDate(moReportItem.DateTo) & ") " & vbLf & _
               "and a.clinic_ncid = " & moReportItem.CriteriaSel1.Item(1).NCID & " " & vbLf & _
               "and a.clinician_ncid = " & moReportItem.CriteriaSel2.Item(1).NCID & " "
    goDBSqlOps.Execute sStmt
    If Not (goDBSqlOps.BOF And goDBSqlOps.EOF) Then
       If Not moTableOps.OpenTable("Patient", rs) Then GoTo CleanUp
       Do While Not goDBSqlOps.EOF
           rs.AddNew
           rs("PatientID") = goDBSqlOps.Value("patient_unit_number")
           rs("FirstName") = goDBSqlOps.Value("first_name")
           rs("MiddleName") = goDBSqlOps.Value("middle_name")
           rs("LastName") = goDBSqlOps.Value("last_name")
           
           sSql = "SELECT ID_TYPE_NCID, ID_VALUE FROM mmi_generic_id " & _
                     "WHERE (ID_TYPE_NCID = 204708 OR ID_TYPE_NCID = 186334) AND " & _
                     "UNIT_NUMBER = " & goDBSqlOps.Value("patient_unit_number")
           oSQL.Execute sSql
           If Not (oSQL.BOF And oSQL.EOF) Then
              Do Until oSQL.EOF
                  If oSQL("ID_TYPE_NCID") = 204708 Then
                     rs("FMP") = oSQL.Value("ID_VALUE")
                  ElseIf oSQL("ID_TYPE_NCID") = 186334 Then
                     rs("SponsorSSN") = oSQL.Value("ID_VALUE")
                  End If
                  oSQL.MoveNext
              Loop
           End If
           rs.Update
           goDBSqlOps.MoveNext
       Loop
       If Not moTableOps.CloseTable(rs) Then GoTo CleanUp
    End If
    
    '-- Get Diagnosis ----------------------------------------------------------------------------------------------
    sStmt = "select /*+ RULE no_index(e event_x4) */ a.appt_date_time as ApptDate, a.patient_unit_number, ov.value_ncid DiagnosisNCID " & vbLf & _
            "from observation_value ov, clinical_observation co, event e, appointment a " & vbLf & _
            "where ov.observation_seq = co.observation_seq and " & vbLf & _
            "ov.event_id = co.event_id and " & vbLf & _
            "co.observation_ncid = 658 and " & vbLf & _
            "co.event_id = e.event_id and " & vbLf & _
            "e.event_type_ncid = 90803 and " & vbLf & _
            "e.Encounter_Number = a.Encounter_Number " & vbLf & _
            " and (a.appt_date_time >= " & GetOracleStartDate(moReportItem.DateFrom) & vbLf & _
            " and a.appt_date_time <= " & GetOracleEndDate(moReportItem.DateTo) & ") " & vbLf & _
            " and a.clinic_ncid = " & moReportItem.CriteriaSel1.Item(1).NCID & " " & vbLf & _
            " and a.clinician_ncid = " & moReportItem.CriteriaSel2.Item(1).NCID & " " & vbLf & _
            " and ov.value_ncid <> 14516241 "
        'R. Foster SCR #36263;   Developer: Robert Foster 05/23/2003 03:48 PM
        'R. Foster Note: Function Changed: DataGetADSMultiPatientEncounter
        'R. Foster Note:Impact of Change: Filter out erroneous Raw Medcin ID data.
        '(ov.value_ncid <> 14516241)
        
    goDBSqlOps.Execute sStmt
    If Not (goDBSqlOps.BOF And goDBSqlOps.EOF) Then
       If Not moTableOps.OpenTable("ADSPatientEncounterSubDiagnosis", rs) Then GoTo CleanUp
       Do While Not goDBSqlOps.EOF
           rs.AddNew
           rs("PatientID") = goDBSqlOps.Value("patient_unit_number")
           rs("DiagnosisNCID") = goDBSqlOps.Value("DiagnosisNCID")
           rs("ApptDate") = Format(goDBSqlOps.Value("ApptDate"), gsDefaultFormat)
           rs.Update
           goDBSqlOps.MoveNext
       Loop
       If Not moTableOps.GetDiagnosisTableData("ADSPatientEncounterSubDiagnosis", "DiagnosisNCID") Then GoTo CleanUp
    End If
    
    If Not moTableOps.CloseTable(rs) Then GoTo CleanUp

    '-- Get Procedures ----------------------------------------------------------------------------------------------
    sStmt = "select /*+ RULE no_index(e event_x4) */ a.appt_date_time as ApptDate, a.patient_unit_number, ov.value_ncid ProcedureNCID " & vbLf & _
            "from observation_value ov, clinical_observation co, event e, appointment a " & vbLf & _
            "where ov.observation_seq = co.observation_seq and " & vbLf & _
            "ov.event_id = co.event_id and " & vbLf & _
            "co.observation_ncid = 41 and " & vbLf & _
            "co.event_id = e.event_id and " & vbLf & _
            "e.event_type_ncid = 90803 and " & vbLf & _
            "e.Encounter_Number = a.Encounter_Number " & vbLf & _
            " and (a.appt_date_time >= " & GetOracleStartDate(moReportItem.DateFrom) & vbLf & _
            " and a.appt_date_time <= " & GetOracleEndDate(moReportItem.DateTo) & ") " & vbLf & _
            " and a.clinic_ncid = " & moReportItem.CriteriaSel1.Item(1).NCID & " " & vbLf & _
            " and a.clinician_ncid = " & moReportItem.CriteriaSel2.Item(1).NCID & " "
        
    goDBSqlOps.Execute sStmt
    If Not (goDBSqlOps.BOF And goDBSqlOps.EOF) Then
       If Not moTableOps.OpenTable("ADSPatientEncounterSubProcedures", rs) Then GoTo CleanUp
       Do While Not goDBSqlOps.EOF
           rs.AddNew
           rs("PatientID") = goDBSqlOps.Value("patient_unit_number")
           rs("ProcedureNCID") = goDBSqlOps.Value("ProcedureNCID")
           rs("ApptDate") = Format(goDBSqlOps.Value("ApptDate"), gsDefaultFormat)
           rs.Update
           goDBSqlOps.MoveNext
       Loop
       If Not moTableOps.GetProcedureTableData("ADSPatientEncounterSubProcedures", "ProcedureNCID") Then GoTo CleanUp
    End If
    
    If Not moTableOps.CloseTable(rs) Then GoTo CleanUp

    '-- Get E&M ----------------------------------------------------------------------------------------------
    sStmt = "select /*+ RULE no_index(e event_x4) */ a.appt_date_time as ApptDate, a.patient_unit_number, ov.value_ncid EMNCID " & vbLf & _
            "from observation_value ov, clinical_observation co, event e, appointment a " & vbLf & _
            "where ov.observation_seq = co.observation_seq and " & vbLf & _
            "ov.event_id = co.event_id and " & vbLf & _
            "co.observation_ncid = 41 and " & vbLf & _
            "co.event_id = e.event_id and " & vbLf & _
            "e.event_subtype_ncid = 184917 and " & vbLf & _
            "e.event_type_ncid = 209614 and " & vbLf & _
            "e.Encounter_Number = a.Encounter_Number " & vbLf & _
            " and (a.appt_date_time >= " & GetOracleStartDate(moReportItem.DateFrom) & vbLf & _
            " and a.appt_date_time <= " & GetOracleEndDate(moReportItem.DateTo) & ") " & vbLf & _
            " and a.clinic_ncid = " & moReportItem.CriteriaSel1.Item(1).NCID & " " & vbLf & _
            " and a.clinician_ncid = " & moReportItem.CriteriaSel2.Item(1).NCID & " "
            
    goDBSqlOps.Execute sStmt
    If Not (goDBSqlOps.BOF And goDBSqlOps.EOF) Then
       If Not moTableOps.OpenTable("ADSPatientEncounterSubEM", rs) Then GoTo CleanUp
       Do While Not goDBSqlOps.EOF
           rs.AddNew
           rs("PatientID") = goDBSqlOps.Value("patient_unit_number")
           rs("EMNCID") = goDBSqlOps.Value("EMNCID")
           rs("ApptDate") = Format(goDBSqlOps.Value("ApptDate"), gsDefaultFormat)
           rs.Update
           goDBSqlOps.MoveNext
       Loop
       If Not moTableOps.GetEandMTableData("ADSPatientEncounterSubEM", "EMNCID") Then GoTo CleanUp
    End If
    
    If Not moTableOps.CloseTable(rs) Then GoTo CleanUp
    
    DataGetADSMultiPatientEncounter = True
    GoTo CleanUp

ErrHandler:
    If Err = -2147220874 Then
        MsgBxARMd "The selected report has timed out, please re-run this report again.", vbOKOnly + vbInformation
        Err.Clear
    Else
        Call gobjshared.ShowVBError(Err.Number, Err.Description, "ReportOps.DataGetADSPatientEncounter", App.Title, vbExclamation)
    End If
CleanUp:
    Set rs = Nothing
    Set oSQL = Nothing

End Function

Private Function InitializeCrystal() As Boolean
    
    '-- This routine is executed for the exclusive purpose of
    '-- preventing Crystal 8 errors.
    '-- It appears to initialize something
    
    On Error GoTo ErrHandler
    
    InitializeCrystal = False
    
    Dim oCrystal As New Crystal
    Dim oForm As frmCrystalView
    
    Set moReport = Nothing
    
    If Not oCrystal.ReportOpen("Initialize.rpt", moReport) Then GoTo CleanUp
    Set oForm = New frmCrystalView
    
    On Error Resume Next
    oForm.CRViewer1.Refresh
    On Error GoTo ErrHandler
    
    oForm.CRViewer1.ReportSource = moReport
    oForm.CRViewer1.DisplayGroupTree = False
    oForm.CRViewer1.EnableGroupTree = False
    oForm.CRViewer1.DisplayTabs = False
    oForm.CRViewer1.EnableDrillDown = False
    oForm.CRViewer1.EnableRefreshButton = False
    oForm.CRViewer1.EnableExportButton = True 'False 'True
    
    DoEvents
    oForm.CRViewer1.ViewReport
    DoEvents
    
    oForm.CRViewer1.Zoom 1
    DoEvents
    'Unload oForm
    
    InitializeCrystal = True
    GoTo CleanUp
    
ErrHandler:
        gobjshared.ShowVBError Err.Number, Err.Description, "ReportOps.InitializeCrystal", App.Title, vbExclamation
CleanUp:
        Set oForm = Nothing
        Set oCrystal = Nothing

End Function
Private Function DataGet(sReportName As String) As Boolean

On Error GoTo ErrHandler
DataGet = False

Dim oReportClass As Object


If moReportInfo.ClassName <> "" Then
    Set oReportClass = CreateObject("CHCSII_Reports." & moReportInfo.ClassName)
    If Not oReportClass.BuildReport(sReportName, moReportItem, moReport, moHeaderColl) Then GoTo CleanUp
    DataGet = True
    GoTo CleanUp
End If

'-- Get Data for the Report
'---------------------------------------------------------------------------------------
Select Case moReportItem.Report
    Case CholesterolScreening
        If Not DataGetCholesterolScreening(moHeaderColl) Then GoTo CleanUp
    Case HighCholesterolFollowUpRepeat
        If Not DataGetHighCholesterolFollowUpRepeat(moHeaderColl) Then GoTo CleanUp
    Case HighCholesterolHDL_LDLFollowUp
        If Not DataGetHighCholesterolHDL_LDLFollowUp(moHeaderColl) Then GoTo CleanUp
    Case HighCholesterolFollowUpCounseling
        If Not DataGetHighCholesterolFollowUpCounseling(moHeaderColl) Then GoTo CleanUp
    Case MammographyQueryFor
        If Not DataGetMammographyQueryFor(moHeaderColl) Then GoTo CleanUp
    Case PAPSmearQueryFor
        If Not DataGetPAPSmearQueryFor(moHeaderColl) Then GoTo CleanUp
    Case PAPSmearHaveNot
        If Not DataGetPAPSmearHaveNot(moHeaderColl) Then GoTo CleanUp
    Case AlcoholQueryforPotentialHeavyUse
        If Not DataGetAlcoholQueryforPotentialHeavyUse(moHeaderColl) Then GoTo CleanUp
    Case SafeSexCounselingQueried_AdvisedofSTDs
        If Not DataGetSafeSexCounselingQueried_AdvisedofSTDs(moHeaderColl) Then GoTo CleanUp
    Case TobaccoUsersAdvisingToQuit
        If Not DataGetTobaccoUsersAdvisingToQuit(moHeaderColl) Then GoTo CleanUp
    Case TobaccoUseScreening
        If Not DataGetTobaccoUseScreening(moHeaderColl) Then GoTo CleanUp
    Case PreventionReportCard
        If Not DataGetPreventionReportCard(moHeaderColl) Then GoTo CleanUp
    Case ImmunizationsActiveDutyA2Dose
        If Not DataGetImmunizationsActiveDutyA2Dose(moHeaderColl) Then GoTo CleanUp
    Case ImmunizationsActiveDutyA1Dose
        If Not DataGetImmunizationsActiveDutyA1Dose(moHeaderColl) Then GoTo CleanUp
    Case HebBHibImmunizations2YearOlds
        If Not DataGetHebBHibImmunizations2YearOlds(moHeaderColl) Then GoTo CleanUp
    Case DTPImmunizations2yearOlds
        If Not DataGetDTPImmunizations2yearOlds(moHeaderColl) Then GoTo CleanUp
    Case Immunizations13yearOlds
        If Not DataGetImmunizations13yearOlds(moHeaderColl) Then GoTo CleanUp
    Case CPSSummary
        If Not DataGetCPSSummary(moHeaderColl) Then GoTo CleanUp
    Case AlphabeticPatientList
        If Not DataGetAlphabeticPatientList(moHeaderColl) Then GoTo CleanUp
    Case ADSPatientCategoriesDisposition
        If Not DataGetADSDisposition(moHeaderColl) Then GoTo CleanUp
    Case ADSProcedures
        If Not DataGetADSProcedure(moHeaderColl) Then GoTo CleanUp
    Case Varicellafor2yrolds
        If Not DataGetVaricellaImmunizations2yearOlds(moHeaderColl) Then GoTo CleanUp
    Case ADSInsuranceIndicatorNotMarked
        If Not DataGetADSInsuranceIndicatorNotMarked(moHeaderColl) Then GoTo CleanUp
    Case ADSInsuranceIndicator
        If Not DataGetADSInsuranceIndicator(moHeaderColl) Then GoTo CleanUp
    Case ADSPatientEncounter
        If Not DataGetADSPatientEncounter(moHeaderColl) Then GoTo CleanUp

        
    '-- Forms
    '-------------------------------------------------------------------
    Case DD2766
        If Not DataGetDD2766() Then GoTo CleanUp
    Case DD2813
        If Not DataGetDD2813() Then GoTo CleanUp
    Case DD689
        If Not DataGetDD689() Then GoTo CleanUp
    Case SF522
        If Not DataGetSF522() Then GoTo CleanUp
    Case DD741
        If Not DataGetDD741() Then GoTo CleanUp
    Case DD771
        If Not DataGetDD771() Then GoTo CleanUp
    '-- CPG Forms
    Case CPG_POC
        If Not DataGetCPGPOC() Then GoTo CleanUp
    Case Else
        MsgBxARMd "This report is not yet available.", vbExclamation
        GoTo CleanUp
End Select

DataGet = True
GoTo CleanUp

ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "ReportOps.DataGet", App.Title, vbExclamation)
CleanUp:
    Set oReportClass = Nothing
    Exit Function
End Function







Private Function GetMTForPCM_SQL(sFromSQL, sWhereSQL) As Boolean

If Not DebugMode Then On Error GoTo ErrHandler
    
GetMTForPCM_SQL = False

    If moReportItem.Criteria1 = MTF Then
        sFromSQL = ", mmi_site ms " & vbNewLine
        
        sWhereSQL = _
                        " id.unit_number = ms.unit_number and " & vbNewLine & _
                        " (ms.facility_cid  = " & moReportItem.CriteriaSel1.Item(1).NCID & ") " & vbNewLine
                        
    ElseIf moReportItem.Criteria1 = PCM Then
        Dim sPCMCode As String
        
        sFromSQL = ", mmi_generic_id mg " & vbNewLine
    
        If moReportItem.CriteriaSel1.Item(1) Is Nothing Then
            sPCMCode = "0"
        Else
            On Error Resume Next
            sPCMCode = moReportItem.CriteriaSel1.Item(1).PCMCode
            If Err <> 0 Then sPCMCode = "0"
        End If

        sWhereSQL = _
                        " id.unit_number = mg.unit_number and " & vbNewLine & _
                        " mg.id_value = '" & sPCMCode & "' and " & vbNewLine & _
                        " mg.id_type_ncid = 14501724 " & vbNewLine
    ElseIf moReportItem.Criteria1 = Patient Then
        sFromSQL = "  "
    
        sWhereSQL = " e.unit_number = " & moReportItem.CriteriaSel1.Item(1).UnitNumber & vbNewLine
    End If

GetMTForPCM_SQL = True
GoTo CleanUp
ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "ReportOps.GetMTForPCM_SQL", "ReportOps", vbExclamation)
CleanUp:

End Function

Private Function GetUnitNumbersForPCM(sPCMCode As String, sPatientUnitNumbers As String) As Boolean

    If Not DebugMode Then On Error GoTo ErrHandler
    GetUnitNumbersForPCM = False
    
    Dim sStmt As String
    Dim icount As Integer
    Dim iPatientCount As Integer
    
    sStmt = "select mg.unit_number " & _
            " from mmi_generic_id mg " & _
            " where " & _
            " mg.id_type_ncid = 14501724 and " & _
            " mg.id_value = '" & sPCMCode & "'"
    
    goDBSqlOps.Execute sStmt
    
    If goDBSqlOps.BOF And goDBSqlOps.EOF Then
        GetUnitNumbersForPCM = True
        sPatientUnitNumbers = "Null"
        GoTo CleanUp
    End If
    
    icount = 0
    Do While Not goDBSqlOps.EOF
        If icount = 0 Then
            sPatientUnitNumbers = goDBSqlOps.Value("unit_number")
        Else
            sPatientUnitNumbers = sPatientUnitNumbers & "," & goDBSqlOps.Value("unit_number")
        End If
        icount = icount + 1
        
        iPatientCount = iPatientCount + 1
        If iPatientCount Mod 10 = 0 Then
           iPatientCount = 0
           sPatientUnitNumbers = sPatientUnitNumbers & vbCrLf
        End If
        
        goDBSqlOps.MoveNext
    Loop
    
    
    GetUnitNumbersForPCM = True
    GoTo CleanUp

ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "ReportOps.GetUnitNumbersForPCM", App.Title, vbExclamation)
CleanUp:
    Exit Function
End Function

'******************************************************************************************************
'  Sub Name: RunReport
'
'  Description: Manages the data extraction and printing of a report based on
'               the report criteria object parameter values
'
'  Parameters:
'       RepItem - Object that describes report and criteria for the report
'
'  Return Value:
'       False - if an error occured
'       True  - successful
'******************************************************************************************************
Friend Function RunReport(RepItem As ReportItem, bDisplayInSeparateWindow As Boolean) As Boolean

If Not DebugMode Then On Error GoTo ErrHandler

RunReport = False


Dim sReportName As String '-- Name of report to run (determined by criteria)
Dim HeaderColl As Collection '-- Header information to be displayed in report


'-- Initialize Variables
'---------------------------------------------------------------------------------------
Set moColRecSet = Nothing
Set moRS = Nothing
Set moReport = Nothing
Set moReportItem = RepItem '-- Report criteria to public for this module
Set moReportInfo = poCollRepInfo(CStr(moReportItem.Report))

'-- Empty all the tables in Access
'---------------------------------------------------------------------------------------
If Not DeleteFromAllTables() Then GoTo CleanUp

'-- Get Name of report file
'---------------------------------------------------------------------------------------
If Not GetReportName(sReportName) Then GoTo CleanUp


'-- build header information for report and other information passed in by formula
'---------------------------------------------------------------------------------------
If Not GetReportHeader(moHeaderColl) Then GoTo CleanUp

'-- Get Report
'---------------------------------------------------------------------------------------
If Not DataGet(sReportName) Then GoTo CleanUp

'-- Run Report
'-----------------------------------------------------------------------------------------
If moReport Is Nothing Then 'Delay seems to prevent error for Access DB
    Delay
End If

DoEvents

If Not DisplayReport(sReportsPath & sReportName & ".rpt", moHeaderColl, bDisplayInSeparateWindow) Then GoTo CleanUp


RunReport = True
GoTo CleanUp


ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "ReportOps.RunReport", "ReportOps", vbExclamation)
CleanUp:
    Set moReport = Nothing
    Set moHeaderColl = Nothing
    Exit Function
End Function

Public Function GetCriteria() As Boolean

If Not DebugMode Then On Error GoTo ErrHandler

GetCriteria = True

Dim rsInput As ADODB.Recordset
Dim sStmt As String
Dim oCriteriaItem As CriteriaItem


Set rsInput = New ADODB.Recordset
rsInput.CursorLocation = adUseClient
rsInput.CursorType = adOpenStatic
rsInput.ActiveConnection = poConnection

sStmt = "select * " & _
        "from Criteria "

rsInput.Open sStmt, , , adLockBatchOptimistic, adCmdText

Set pCollCriteria = Nothing
Set pCollCriteria = New Collection

Do While Not rsInput.EOF
    
    Set oCriteriaItem = New CriteriaItem
    
    oCriteriaItem.Id = rsInput.Fields("Id")

    oCriteriaItem.Name = rsInput.Fields("Name")
        
    pCollCriteria.Add oCriteriaItem, CStr(oCriteriaItem.Id)
        
    rsInput.MoveNext
    
    Set oCriteriaItem = Nothing
    
Loop

rsInput.Close

GetCriteria = True

GoTo CleanUp

ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "ReportOps.GetCriteria", "ReportOps", vbExclamation)
CleanUp:
    Set oCriteriaItem = Nothing
    Set rsInput = Nothing
    Exit Function
End Function


Private Sub SQLComments()
'-- Lab Queries
'94129

'-- Get Domain NCID for LAB Tests
'Select concept_relation_ncid, relationship_ncid
'From concept_relation
'Where
'concept_ncid = 94129;
'-- concept_relation_ncid = 600 Relationship_ncid = 363 or 364

'Select concept_relation_ncid, relationship_ncid
'From concept_relation
'Where
'concept_ncid = 66560

'-- Get a specific Lab Test
'Select r.ncid, substr(r.representation,1,75)
'from rsform_context rc, rsform r
'Where
'rc.context_ncid = 2000 and
'rc.rsform_id = r.rsform_id and
'Exists
'(select 'x'
'From concept_relation
'Where
'concept_relation_ncid = 600 and
'concept_ncid = r.ncid) and
'UPPER(representation) like '%%';
'


'-------------------------------------
'-- Radiology
'select r.ncid, substr(r.representation,1,75)
'from rsform_context rc, rsform r
'Where
'rc.context_ncid = 2000 and
'rc.rsform_id = r.rsform_id and
'r.ncid in
'(select distinct concept_relation_ncid
'from concept_relation cr
'Where
'cr.concept_ncid = 125349);










End Sub

Private Sub Class_Initialize()

    Dim sConnectionString As String
    On Error GoTo ErrHandler
    
    '--- SCR-ARM-16148 MSolano 9/10/01; COMMON TO MULTIPLE ARM-SCRs
    '--- Init routine should be called from first top level creatable coclass
    '--- It can be called multiple times, but will actually run only once.
    '--- Instrument all creatable classes in case dll's multiple entry
    '--- points are actually used
     Call ARM_InitAllTransactionsForCHCSIIReports
 
    sReportsPath = gobjshared.ReportsPath
    If Mid$(sReportsPath, Len(sReportsPath), 1) <> "\" Then
        sReportsPath = sReportsPath & "\"
    End If
    gsFormPath = sReportsPath

    '-- Open connection to the Access database
    msDatabase = sReportsPath & cDBName
    Set poConnection = New ADODB.Connection
    sConnectionString = msDatabase '"Driver={Microsoft Access Driver (*.mdb)};Dbq=" & cDBName & ";DefaultDir=" & sReportsPath & ";"
    'poConnection.Open sConnectionString
    
    With poConnection
        .Provider = "Microsoft.Jet.OLEDB.4.0"
        .ConnectionString = sConnectionString
        .Open
    End With

    '-- Routine Below is called due to errors encountered with Crystal 8
    If Not InitializeCrystal() Then GoTo CleanUp

    GoTo CleanUp
    
ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "ReportOps.Class_Initialize", App.Title, vbExclamation)
CleanUp:
    Set moReport = Nothing
    Exit Sub
End Sub

Private Sub Class_Terminate()

On Error GoTo ErrHandler
poConnection.Close


ErrHandler:
CleanUp:
Set poConnection = Nothing
Set moColRecSet = Nothing
Set moRS = Nothing
Set moTableOps = Nothing
Set moReportItem = Nothing
Set moReportInfo = Nothing

End Sub
Private Function DisplayReport(sReportName As String, HeaderFields As Collection, bDisplayInSeparateWindow As Boolean) As Boolean

    If Not DebugMode Then On Error GoTo ErrHandler
    DisplayReport = False

    Dim External_Report As CRAXDRT.Report
    Dim CrystalApp As New CRAXDRT.Application
    Dim oSubReport As CRAXDRT.SubreportObject
    Dim i As Integer
    Dim i2 As Integer
    Dim i3 As Integer
    Static ofrmInternalForm As frmCrystalViewEmb

    On Error Resume Next
    
   
    If moReport Is Nothing Then
       Set External_Report = CrystalApp.OpenReport(sReportName)
       External_Report.DiscardSavedData
       DoEvents
    Else
       Set External_Report = moReport
       GoTo DisplayIt
    End If
    
    If Err <> 0 Then
       MsgBxARMd "This report is not yet available."
       MsgBxARMd "Error: " & Error(Err) & "  FileName: " & sReportName
       DisplayReport = True
       Exit Function
    End If
    If Not DebugMode Then On Error GoTo ErrHandler
    
    If Not moColRecSet Is Nothing Then
        Set moRS = Nothing
        Set moRS = moColRecSet("MAIN")
        External_Report.Database.SetDataSource moRS
    Else
        For i = 1 To External_Report.Database.Tables.Count
            External_Report.Database.Tables.Item(i).Location = msDatabase
        Next
    End If
    
    DoEvents
    
    
        '-- Set Formula Fields
        '-----------------------------------------------------------------------
        If Not HeaderFields Is Nothing Then
            On Error Resume Next
            For i = 1 To External_Report.FormulaFields.Count
                On Error Resume Next
                External_Report.FormulaFields.Item(i).Text = _
                    HeaderFields.Item(External_Report.FormulaFields.Item(i).FormulaFieldName)
                If Err <> 5 And Err <> 0 Then
                    On Error GoTo ErrHandler
                    Err.Raise Err
                End If
            Next
            On Error GoTo ErrHandler
        End If
    
    DoEvents
    
    
    '-- For each table in a subreport update .MDB file location
    '---------------------------------------------------------------------------------------
    For i = 1 To External_Report.Sections.Count
        For i2 = 1 To External_Report.Sections.Item(i).ReportObjects.Count
            If TypeOf External_Report.Sections.Item(i).ReportObjects.Item(i2) Is SubreportObject Then
                Set oSubReport = External_Report.Sections.Item(i).ReportObjects.Item(i2)
                For i3 = 1 To External_Report.OpenSubreport(oSubReport.SubreportName).Database.Tables.Count
                    External_Report.OpenSubreport(oSubReport.SubreportName).Database.Tables.Item(i3).Location = msDatabase
                Next
            End If
        Next
    Next
    
DisplayIt:
    DoEvents
    
      
    If bDisplayInSeparateWindow Then
        Dim moForm As frmCrystalView
        Set moForm = New frmCrystalView
        moForm.CRViewer1.ReportSource = External_Report
        moForm.CRViewer1.DisplayGroupTree = False
        moForm.CRViewer1.EnableGroupTree = False
        moForm.CRViewer1.DisplayTabs = False
        moForm.CRViewer1.EnableDrillDown = False
        moForm.CRViewer1.EnableRefreshButton = False
        moForm.CRViewer1.EnableExportButton = False
        
        
        '-- Temporary Fix to meet requirement (SCR# 9179)
        If Mid$(sReportName, (Len(sReportName) - InStr(1, StrReverse(sReportName), "\") + 2)) = "CPGQuestionnaireData.rpt" Then
            moForm.cmdExport.Visible = True
        Else
            moForm.cmdExport.Visible = False
        End If
    
        DoEvents
        moForm.CRViewer1.ViewReport
        DoEvents
        moForm.CRViewer1.Zoom 1
        DoEvents
        moForm.Show
    Else
        If Not ofrmInternalForm Is Nothing Then
            Unload ofrmInternalForm
            Set ofrmInternalForm = Nothing
        End If
        Set ofrmInternalForm = New frmCrystalViewEmb
        On Error Resume Next
        ofrmInternalForm.CRViewer1.Refresh
        On Error GoTo ErrHandler
        ofrmInternalForm.CRViewer1.ReportSource = External_Report
        ofrmInternalForm.CRViewer1.DisplayGroupTree = False
        ofrmInternalForm.CRViewer1.EnableGroupTree = False
        ofrmInternalForm.CRViewer1.DisplayTabs = False
        ofrmInternalForm.CRViewer1.EnableDrillDown = False
        ofrmInternalForm.CRViewer1.EnableRefreshButton = False
        ofrmInternalForm.CRViewer1.EnableExportButton = False
        
        '-- Temporary Fix to meet requirement (SCR 9179)
        If Mid$(sReportName, (Len(sReportName) - InStr(1, StrReverse(sReportName), "\") + 2)) = "CPGQuestionnaireData.rpt" Then
            ofrmInternalForm.cmdExport.Visible = True
        Else
            ofrmInternalForm.cmdExport.Visible = False
        End If

        
        SetParent ofrmInternalForm.hwnd, frmReports.picView.hwnd
        ofrmInternalForm.CRViewer1.ViewReport
        DoEvents
        ofrmInternalForm.CRViewer1.Zoom 1
        ofrmInternalForm.BorderStyle = 0
        ofrmInternalForm.ClipControls = False
        ofrmInternalForm.WindowState = 2
        ofrmInternalForm.Show
    End If
    
    DisplayReport = True
    GoTo CleanUp
    
ErrHandler:
        Call gobjshared.ShowVBError(Err.Number, Err.Description, "ReportsOps.DisplayReport", "ReportsOps", vbExclamation)
        'Resume Next
CleanUp:
        Set moForm = Nothing
        Set External_Report = Nothing
        Set CrystalApp = Nothing
        Set oSubReport = Nothing
        Set moColRecSet = Nothing
        Set moRS = Nothing
        Exit Function
End Function

Private Function GetReportHeader(HeaderColl As Collection) As Boolean

If Not DebugMode Then On Error GoTo ErrHandler
GetReportHeader = False

Dim sH1 As String
Dim sH2 As String
Dim sH3 As String
Dim sH4 As String
Dim sH5 As String
Dim oLab As LabItem
Dim oMed As MedItem
Dim oDiagnosis As DiagnosisItem
Dim oMTF As MTFItem
Dim oProcedure As ProcedureItem
Dim oProvider As ProviderItem
Dim oClinic As ClinicItem
Dim i As Integer


    Set HeaderColl = Nothing
    Set HeaderColl = New Collection

    If moReportInfo.Header Then
    Set HeaderColl = Nothing
    Set HeaderColl = New Collection
    
    '-- Build Report Header 1
    '------------------------------------------------------------------------------------------
    If Not Len(moReportInfo.AlternativeHeader) = 0 Then
       sH1 = moReportInfo.AlternativeHeader
    Else
       sH1 = moReportInfo.LongName
    End If
    
    If moReportInfo.AppendCriteriaToHeader Then
        If moReportItem.Criteria1 <> NoCriteria Then
            sH1 = sH1 & " for " & pCollCriteria(CStr(moReportItem.Criteria1)).Name
        End If
        
        If moReportItem.Criteria2 <> NoCriteria Then
            sH1 = sH1 & ", " & pCollCriteria(CStr(moReportItem.Criteria2)).Name
        End If
        
        If moReportItem.Criteria3 <> NoCriteria Then
            sH1 = sH1 & ", " & pCollCriteria(CStr(moReportItem.Criteria3)).Name
        End If
    
        If moReportItem.GroupBy1 <> NoCriteria Then
            sH1 = sH1 & " group by " & pCollCriteria(CStr(moReportItem.GroupBy1)).Name
            If moReportItem.GroupBy2 <> NoCriteria Then
                sH1 = sH1 & "," & pCollCriteria(CStr(moReportItem.GroupBy2)).Name
            End If
        End If
    End If

    HeaderColl.Add CrystalQuote(sH1), "Header1"
    
    '-- Build Report Header 2
    '------------------------------------------------------------------------------------------
    
    If moReportItem.Criteria1 <> NoCriteria Then
        sH2 = pCollCriteria(CStr(moReportItem.Criteria1)).Name & ": "
        If moReportItem.CriteriaSel1.Count <> 0 Then
            For i = 1 To moReportItem.CriteriaSel1.Count
                If TypeOf moReportItem.CriteriaSel1.Item(i) Is CHCSII_Search.ClinicItem Then
                    Set oClinic = moReportItem.CriteriaSel1.Item(i)
                    If sH2 = "" Then
                        sH2 = " - "
                    End If
                    If moReportInfo.LongName = "Provider Satisfaction Report" Then
                        sH2 = "Clinic: " & oClinic.Name
                    Else
                        sH2 = sH2 & oClinic.Name & " (MEPRS: " & oClinic.MeprsCode & " , " & " DMIS: " & oClinic.DMISIdCode & ")"
                    End If
                    If i <> moReportItem.CriteriaSel1.Count Then sH2 = sH2 & ", "
                    Set oClinic = Nothing
                                        
                ElseIf TypeOf moReportItem.CriteriaSel1.Item(i) Is CHCSII_Search.DiagnosisItem Then
                    Set oDiagnosis = moReportItem.CriteriaSel1.Item(i)
                    sH2 = sH2 & oDiagnosis.Name
                    Set oDiagnosis = Nothing
                ElseIf TypeOf moReportItem.CriteriaSel1.Item(i) Is CHCSII_Search.LabItem Then
                    Set oLab = moReportItem.CriteriaSel1.Item(i)
                    sH2 = sH2 & oLab.Name
                    Set oLab = Nothing
                ElseIf TypeOf moReportItem.CriteriaSel1.Item(i) Is CHCSII_Search.MedItem Then
                    Set oMed = moReportItem.CriteriaSel1.Item(i)
                    sH2 = sH2 & oMed.Name
                    Set oMed = Nothing
                ElseIf TypeOf moReportItem.CriteriaSel1.Item(i) Is CHCSII_Search.MTFItem Then
                    Set oMTF = moReportItem.CriteriaSel1.Item(i)
                    sH2 = sH2 & oMTF.Name
                    Set oMTF = Nothing
                ElseIf TypeOf moReportItem.CriteriaSel1.Item(i) Is CHCSII_Search.ProcedureItem Then
                    Set oProcedure = moReportItem.CriteriaSel1.Item(i)
                    sH2 = sH2 & oProcedure.Name
                    Set oProcedure = Nothing
                ElseIf TypeOf moReportItem.CriteriaSel1.Item(i) Is CHCSII_Search.ProviderItem Then
                    Set oProvider = moReportItem.CriteriaSel1.Item(i)
                    sH2 = sH2 & oProvider.Name
                    Set oProvider = Nothing
                ElseIf TypeOf moReportItem.CriteriaSel1.Item(i) Is CHCSII_Search.PCMItem Then
                    sH2 = sH2 & moReportItem.CriteriaSel1.Item(i).Name
                    Set oProvider = Nothing
                ElseIf TypeOf moReportItem.CriteriaSel1.Item(i) Is CHCSII_Search.PatientItem Then
                    sH2 = sH2 & Trim(moReportItem.CriteriaSel1.Item(i).Last & ", " & moReportItem.CriteriaSel1.Item(i).First & " " & moReportItem.CriteriaSel1.Item(i).Middle)
                End If
            Next
        End If
        

        If Len(sH2) > 254 Then
            sH2 = Mid$(sH2, 1, 254)
        End If
        HeaderColl.Add CrystalQuote(sH2), "Header2"
    
    End If
    
    
    '-- Build Report Header 3
    '------------------------------------------------------------------------------------------
    If moReportItem.Criteria2 <> NoCriteria Then
        sH3 = pCollCriteria(CStr(moReportItem.Criteria2)).Name & ": " '(moReportItem.Criteria1, LongName) & ": "
        If moReportItem.CriteriaSel2.Count <> 0 Then
            For i = 1 To moReportItem.CriteriaSel2.Count
                If TypeOf moReportItem.CriteriaSel2.Item(i) Is CHCSII_Search.ClinicItem Then
                    Set oClinic = moReportItem.CriteriaSel2.Item(i)
                    sH3 = sH3 & oClinic.Name
                    Set oClinic = Nothing
                ElseIf TypeOf moReportItem.CriteriaSel2.Item(i) Is CHCSII_Search.DiagnosisItem Then
                    Set oDiagnosis = moReportItem.CriteriaSel2.Item(i)
                    sH3 = sH3 & oDiagnosis.Name
                    Set oDiagnosis = Nothing
                ElseIf TypeOf moReportItem.CriteriaSel2.Item(i) Is CHCSII_Search.LabItem Then
                    Set oLab = moReportItem.CriteriaSel2.Item(i)
                    sH3 = sH3 & oLab.Name
                    Set oLab = Nothing
                ElseIf TypeOf moReportItem.CriteriaSel2.Item(i) Is CHCSII_Search.MedItem Then
                    Set oMed = moReportItem.CriteriaSel2.Item(i)
                    sH3 = sH3 & oMed.Name
                    Set oMed = Nothing
                ElseIf TypeOf moReportItem.CriteriaSel2.Item(i) Is CHCSII_Search.MTFItem Then
                    Set oMTF = moReportItem.CriteriaSel2.Item(i)
                    sH3 = sH3 & oMTF.Name
                    Set oMTF = Nothing
                ElseIf TypeOf moReportItem.CriteriaSel2.Item(i) Is CHCSII_Search.ProcedureItem Then
                    Set oProcedure = moReportItem.CriteriaSel2.Item(i)
                    sH3 = sH3 & oProcedure.Name
                    Set oProcedure = Nothing
                ElseIf TypeOf moReportItem.CriteriaSel2.Item(i) Is CHCSII_Search.ProviderItem Then
                    Set oProvider = moReportItem.CriteriaSel2.Item(i)
                    sH3 = sH3 & oProvider.Name
                    Set oProvider = Nothing
                ElseIf TypeOf moReportItem.CriteriaSel2.Item(i) Is CHCSII_Search.ScreeningServiceItem Then
                    sH3 = sH3 & moReportItem.CriteriaSel2.Item(i).Name
                ElseIf TypeOf moReportItem.CriteriaSel2.Item(i) Is CHCSII_Search.PatientItem Then
                    'Set oPatient = moReportItem.CriteriaSel2.Item(i)
                    sH3 = sH3 & moReportItem.CriteriaSel2.Item(i).Last & ", " & moReportItem.CriteriaSel2.Item(i).First
                    'Set oProvider = Nothing
                ElseIf TypeOf moReportItem.CriteriaSel2.Item(i) Is CHCSII_Search.RadItem Then
                    'Set oProvider = moReportItem.CriteriaSel2.Item(i)
                    sH3 = sH3 & moReportItem.CriteriaSel2.Item(i).Name
                    'Set oProvider = Nothing
                ElseIf TypeOf moReportItem.CriteriaSel2.Item(i) Is CHCSII_Search.QuestionnaireAnswerItem Then
                    sH3 = "Questionnaire: " & moReportItem.CriteriaSel2.Item(i).QDefID & "   Question: " & moReportItem.CriteriaSel2.Item(i).Seq
                    sH5 = moReportItem.CriteriaSel2.Item(i).Text
                    HeaderColl.Add CrystalQuote(sH5), "Header5"
                End If
            Next
        End If
        
        HeaderColl.Add CrystalQuote(sH3), "Header3"

    
    End If
    
    '-- Build Report Header 4
    '------------------------------------------------------------------------------------------
    If moReportItem.Criteria3 <> NoCriteria Then
        sH4 = pCollCriteria(CStr(moReportItem.Criteria3)).Name & ": "
        If moReportItem.CriteriaSel3.Count <> 0 Then
            For i = 1 To moReportItem.CriteriaSel3.Count
                If TypeOf moReportItem.CriteriaSel3.Item(i) Is CHCSII_Search.ClinicItem Then
                    Set oClinic = moReportItem.CriteriaSel3.Item(i)
                    sH4 = sH4 & oClinic.Name
                    Set oClinic = Nothing
                ElseIf TypeOf moReportItem.CriteriaSel3.Item(i) Is CHCSII_Search.DiagnosisItem Then
                    Set oDiagnosis = moReportItem.CriteriaSel3.Item(i)
                    sH4 = sH4 & oDiagnosis.Name
                    Set oDiagnosis = Nothing
                ElseIf TypeOf moReportItem.CriteriaSel3.Item(i) Is CHCSII_Search.LabItem Then
                    Set oLab = moReportItem.CriteriaSel3.Item(i)
                    sH4 = sH4 & oLab.Name
                    Set oLab = Nothing
                ElseIf TypeOf moReportItem.CriteriaSel3.Item(i) Is CHCSII_Search.MedItem Then
                    Set oMed = moReportItem.CriteriaSel3.Item(i)
                    sH4 = sH4 & oMed.Name
                    Set oMed = Nothing
                ElseIf TypeOf moReportItem.CriteriaSel3.Item(i) Is CHCSII_Search.MTFItem Then
                    Set oMTF = moReportItem.CriteriaSel3.Item(i)
                    sH4 = sH4 & oMTF.Name
                    Set oMTF = Nothing
                ElseIf TypeOf moReportItem.CriteriaSel3.Item(i) Is CHCSII_Search.ProcedureItem Then
                    Set oProcedure = moReportItem.CriteriaSel3.Item(i)
                    sH4 = sH4 & oProcedure.Name
                    Set oProcedure = Nothing
                ElseIf TypeOf moReportItem.CriteriaSel3.Item(i) Is CHCSII_Search.ProviderItem Then
                    Set oProvider = moReportItem.CriteriaSel3.Item(i)
                    sH4 = sH4 & oProvider.Name
                    Set oProvider = Nothing
                ElseIf TypeOf moReportItem.CriteriaSel3.Item(i) Is CHCSII_Search.PatientItem Then
                    sH4 = sH4 & moReportItem.CriteriaSel3.Item(i).Last & ", " & moReportItem.CriteriaSel3.Item(i).First
                ElseIf TypeOf moReportItem.CriteriaSel3.Item(i) Is CHCSII_Search.ScreeningServiceItem Then
                    sH4 = sH4 & moReportItem.CriteriaSel3.Item(i).Name
                ElseIf TypeOf moReportItem.CriteriaSel3.Item(i) Is CHCSII_Search.RadItem Then
                    sH4 = sH4 & moReportItem.CriteriaSel3.Item(i).Name
                ElseIf TypeOf moReportItem.CriteriaSel3.Item(i) Is CHCSII_Search.QuestionnaireAnswerItem Then
                    sH4 = "Questionnaire: " & moReportItem.CriteriaSel3.Item(i).QDefID & "   Question: " & moReportItem.CriteriaSel3.Item(i).Seq
                    sH5 = moReportItem.CriteriaSel3.Item(i).Text
                    HeaderColl.Add CrystalQuote(sH5), "Header5"
                End If
            Next
        End If
        

        HeaderColl.Add CrystalQuote(sH4), "Header4"
    
    End If



End If

If moReportInfo.DateFrom Then
    HeaderColl.Add CrystalQuote(Format(moReportItem.DateFrom, gsDefaultFormat)), "DateFrom"
End If
    
If moReportInfo.DateTo Then
    HeaderColl.Add CrystalQuote(Format(moReportItem.DateTo, gsDefaultFormat)), "DateTo"
End If

HeaderColl.Add CrystalQuote(Format(Now, gsDefaultFormat)), "TodaysDate"

HeaderColl.Add CrystalQuote(Format(Now, gsDefaultFormat & "  HH:MM")), "TodaysDateTime"

HeaderColl.Add CrystalQuote(Trim(cwobjLogonEnviron.NursingDivisionName)), "ClinicName"


GetReportHeader = True
GoTo CleanUp

ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "ReportsOps.GetReportHeader", "ReportsOps", vbExclamation)
CleanUp:
    Exit Function
End Function

Private Function GetReportName(sReportName As String) As Boolean

If Not DebugMode Then On Error GoTo ErrHandler
GetReportName = False


'-- Build Report Name (Report name varies based on grouping)
'------------------------------------------------------------------------------------------
sReportName = moReportInfo.ReportFileName

If moReportInfo.Header = True Then
    If moReportInfo.IsGrouping = True Then
        If Not moReportItem.GroupBy1 = NoCriteria Then
            sReportName = sReportName & pCollCriteria(CStr(moReportItem.GroupBy1)).Name
        End If
        
        If InStr(sReportName, "/Lab") Then
            sReportName = Mid(sReportName, 1, InStr(1, sReportName, "/") - 1)
        End If
        
        If Not moReportItem.GroupBy2 = NoCriteria Then
            sReportName = sReportName & pCollCriteria(CStr(moReportItem.GroupBy2)).Name
        End If
        
        If InStr(sReportName, "/Lab") Then
            sReportName = Mid(sReportName, 1, InStr(1, sReportName, "/") - 1)
        End If
        
    End If
End If



GetReportName = True
GoTo CleanUp

ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "ReportsOps.GetReportName", "ReportsOps", vbExclamation)
CleanUp:
    Exit Function
End Function








Private Function BuildWhereClauseforLocation(sStmt As String) As Boolean
'-- Assumes alias for event table is 'e' and Action_Info is 'ai'

If Not DebugMode Then On Error GoTo ErrHandler

BuildWhereClauseforLocation = False

Dim i As Integer

    If moReportItem.Criteria1 = MTF Then
       sStmt = sStmt & " e.facility_ncid  = " & moReportItem.CriteriaSel1.Item(1).NCID & " "
    ElseIf moReportItem.Criteria1 = Clinic Then
        sStmt = sStmt & " ai.poc_location_ncid in ("
    
        For i = 1 To moReportItem.CriteriaSel1.Count
            If i <> moReportItem.CriteriaSel1.Count Then
                sStmt = sStmt & moReportItem.CriteriaSel1.Item(i).NCID & ","
            Else
                sStmt = sStmt & moReportItem.CriteriaSel1.Item(i).NCID & ") "
            End If
        Next
    ElseIf moReportItem.Criteria1 = Provider Then
        sStmt = sStmt & " ai.clinician_id_ncid = " & moReportItem.CriteriaSel1.Item(1).NCID & " "
    End If

    sStmt = sStmt & _
                " and ((ai.action_type_ncid = 1520 " & _
                " and ai.poc_location_type_ncid = 1003) " & _
                " or (ai.action_type_ncid is null)) " & _
                " and ai.event_id (+) = e.event_id "

BuildWhereClauseforLocation = True
GoTo CleanUp

ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "ReportOps.BuildWhereClauseforLocation", "frmReports", vbExclamation)
CleanUp:
    Exit Function
End Function




Private Function BuildWhereClauseforEncounterEvents2(sStmt As String, _
                                                     Optional EventSubTypeNCID As String) As Boolean
'-- Assumes alias for event table is 'e' and Action_Info is 'ai'

If Not DebugMode Then On Error GoTo ErrHandler

Dim i As Integer

BuildWhereClauseforEncounterEvents2 = False

    '-- Clinical_Event_Header
    '-------------------------------------------------------------------------------------------
    sStmt = sStmt & " and c.status_ncid = 205657 and c.event_id = e.event_id "

    '-- Location
    '--------------------------------------------------------------------------------------------
    If moReportItem.Criteria1 = Clinic Then
        sStmt = sStmt & " and ai.poc_location_ncid in ("
    
        For i = 1 To moReportItem.CriteriaSel1.Count
            If i <> moReportItem.CriteriaSel1.Count Then
                sStmt = sStmt & moReportItem.CriteriaSel1.Item(i).NCID & ","
            Else
                sStmt = sStmt & moReportItem.CriteriaSel1.Item(i).NCID & ") "
            End If
        Next
        
    ElseIf moReportItem.Criteria1 = Provider Then
        sStmt = sStmt & " and ai.clinician_id_ncid = " & moReportItem.CriteriaSel1.Item(1).NCID & " "
    End If

    sStmt = sStmt & _
                " and ((ai.action_type_ncid = 1520 " & _
                " and ai.poc_location_type_ncid = 1003) " & _
                " or (ai.action_type_ncid is null)) " & _
                " and ai.event_id (+) = e.event_id "

    If moReportItem.Criteria1 = MTF Then
       sStmt = sStmt & " and e.facility_ncid  = " & moReportItem.CriteriaSel1.Item(1).NCID & " "
    ElseIf moReportItem.Criteria1 = PCM Then
        Dim sPCMCode As String
        
        If moReportItem.CriteriaSel1.Item(1) Is Nothing Then
            sPCMCode = "0"
        Else
            On Error Resume Next
            sPCMCode = moReportItem.CriteriaSel1.Item(1).PCMCode
            If Err <> 0 Then sPCMCode = "0"
        End If
        sStmt = sStmt & " and e.unit_number in " & _
                        "(select mg.unit_number " & _
                        " from mmi_generic_id mg " & _
                        " where " & _
                        " mg.id_type_ncid = 14501724 and " & _
                        " mg.id_value = '" & sPCMCode & "') "
    End If
    
    

    '-- Event
    '-------------------------------------------------------------------------------------------
    sStmt = sStmt & " and (e.event_start_gmtime >= " & GetOracleStartDate(moReportItem.DateFrom) & _
                    " and e.event_start_gmtime <= " & GetOracleEndDate(moReportItem.DateTo) & ")"

    If EventSubTypeNCID <> "" Then
        sStmt = sStmt & " and e.event_subtype_ncid = " & EventSubTypeNCID
    End If
    
    sStmt = sStmt & " and e.event_type_ncid = 209614 "

BuildWhereClauseforEncounterEvents2 = True
GoTo CleanUp

ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "ReportOps.BuildWhereClauseforEncounterEvents", "frmReports", vbExclamation)
CleanUp:
    Exit Function
End Function

Private Function BuildWhereClauseforEncounterEvents3(sStmt As String, _
                                                     Optional EventSubTypeNCID As String) As Boolean
'-- Assumes alias for event table is 'e' and Action_Info is 'ai'

If Not DebugMode Then On Error GoTo ErrHandler

Dim i As Integer

BuildWhereClauseforEncounterEvents3 = False

'-------------------------------------------------------------------------------------------
'-- STATUS IS COMPLETED (Clinical_Event_Header)
'-------------------------------------------------------------------------------------------
    sStmt = sStmt & " and c.status_ncid = 205657 and c.event_id = e.event_id "

'-------------------------------------------------------------------------------------------
'-- Event
'-------------------------------------------------------------------------------------------
    If EventSubTypeNCID <> "" Then
        sStmt = sStmt & " and e.event_subtype_ncid = " & EventSubTypeNCID
    End If
    sStmt = sStmt & " and e.event_type_ncid = 209614 and e.encounter_number = a.Encounter_Number "
    
'-------------------------------
'- MTF
'-------------------------------
    If moReportItem.Criteria1 = MTF Then
        sStmt = sStmt & " and a.facility_ncid  = " & moReportItem.CriteriaSel1.Item(1).NCID & " "
    End If
        
'-------------------------------
'- DATE RANGE
'-------------------------------
    sStmt = sStmt & " and (a.APPT_DATE_TIME >= " & GetOracleStartDate(moReportItem.DateFrom) & _
                    " and a.APPT_DATE_TIME <= " & GetOracleEndDate(moReportItem.DateTo) & ")"


'-------------------------------
'- CLINIC
'-------------------------------
    If moReportItem.Criteria1 = Clinic Then
    
        
        sStmt = sStmt & " and a.clinic_ncid in ("
    
        For i = 1 To moReportItem.CriteriaSel1.Count
            If i <> moReportItem.CriteriaSel1.Count Then
                sStmt = sStmt & moReportItem.CriteriaSel1.Item(i).NCID & ","
            Else
                sStmt = sStmt & moReportItem.CriteriaSel1.Item(i).NCID & ") "
            End If
        Next

'-------------------------------
'- PROVIDER
'-------------------------------
    ElseIf moReportItem.Criteria1 = Provider Then
        sStmt = sStmt & " and a.clinician_ncid = " & moReportItem.CriteriaSel1.Item(1).NCID & " "

'-------------------------------
'- PCM
'-------------------------------
    ElseIf moReportItem.Criteria1 = PCM Then
        Dim sPCMCode As String
        Dim sUnitNumbers As String

        If moReportItem.CriteriaSel1.Item(1) Is Nothing Then
            sPCMCode = "0"
        Else
            On Error Resume Next
            sPCMCode = moReportItem.CriteriaSel1.Item(1).PCMCode
            If Err <> 0 Then sPCMCode = "0"
        End If
        
        If Not GetUnitNumbersForPCM(sPCMCode, sUnitNumbers) Then GoTo CleanUp

        sStmt = sStmt & " and a.patient_unit_number in " & _
                        "(" & sUnitNumbers & ")"
    End If


BuildWhereClauseforEncounterEvents3 = True
GoTo CleanUp

ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "ReportOps.BuildWhereClauseforEncounterEvents3", "frmReports", vbExclamation)
CleanUp:
    Exit Function
End Function

Private Function BuildWhereClauseforEncounterEvents4(sStmt As String, _
                                                     Optional EventSubTypeNCID As String) As Boolean
'-- Assumes alias for event table is 'e' and Action_Info is 'ai'

If Not DebugMode Then On Error GoTo ErrHandler

Dim i As Integer

BuildWhereClauseforEncounterEvents4 = False

'-------------------------------------------------------------------------------------------
'-- STATUS IS COMPLETED (Clinical_Event_Header)
'-------------------------------------------------------------------------------------------
    sStmt = sStmt & " and c.status_ncid in (205657, 14510306, 14510309) and " & vbNewLine & _
                    "c.event_id = e.event_id " & vbNewLine

'-------------------------------------------------------------------------------------------
'-- Event
'-------------------------------------------------------------------------------------------
    If EventSubTypeNCID <> "" Then
        sStmt = sStmt & " and e.event_subtype_ncid = " & EventSubTypeNCID & vbNewLine
    End If
    sStmt = sStmt & " and e.event_type_ncid = 209614 and " & vbNewLine & _
                    "e.encounter_number = a.Encounter_Number and " & vbNewLine & _
                    "a.encounter_number <> '0' " & vbNewLine
    
'-------------------------------
'- MTF
'-------------------------------
    If moReportItem.Criteria1 = MTF Then
        sStmt = sStmt & " and a.facility_ncid  = " & moReportItem.CriteriaSel1.Item(1).NCID & " " & vbNewLine
    End If
        
'-------------------------------
'- CLINIC
'-------------------------------
    If moReportItem.Criteria1 = Clinic Then
    
        
        sStmt = sStmt & " and a.clinic_ncid in ("
    
        For i = 1 To moReportItem.CriteriaSel1.Count
            If i <> moReportItem.CriteriaSel1.Count Then
                sStmt = sStmt & moReportItem.CriteriaSel1.Item(i).NCID & ","
            Else
                sStmt = sStmt & moReportItem.CriteriaSel1.Item(i).NCID & ") " & vbNewLine
            End If
        Next

'-------------------------------
'- PROVIDER
'-------------------------------
    ElseIf moReportItem.Criteria1 = Provider Then
        sStmt = sStmt & " and a.clinician_ncid = " & moReportItem.CriteriaSel1.Item(1).NCID & " " & vbNewLine

'-------------------------------
'- PCM
'-------------------------------
    ElseIf moReportItem.Criteria1 = PCM Then
        Dim sPCMCode As String
        Dim sUnitNumbers As String

        If moReportItem.CriteriaSel1.Item(1) Is Nothing Then
            sPCMCode = "0"
        Else
            On Error Resume Next
            sPCMCode = moReportItem.CriteriaSel1.Item(1).PCMCode
            If Err <> 0 Then sPCMCode = "0"
        End If
                
        sStmt = sStmt & " and a.patient_unit_number in " & _
                        "(SELECT mg.unit_number " & _
                        " from mmi_generic_id mg " & _
                        " where mg.id_type_ncid = 14501724 and mg.id_value = " & SQLQuote(sPCMCode) & ") " & vbNewLine
    End If

'-------------------------------
'- DATE RANGE
'-------------------------------
    sStmt = sStmt & " and (a.APPT_DATE_TIME >= " & GetOracleStartDate(moReportItem.DateFrom) & vbNewLine & _
                    " and a.APPT_DATE_TIME <= " & GetOracleEndDate(moReportItem.DateTo) & ")"

BuildWhereClauseforEncounterEvents4 = True
GoTo CleanUp

ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "ReportOps.BuildWhereClauseforEncounterEvents4", "frmReports", vbExclamation)
CleanUp:
    Exit Function
End Function


Private Function BuildWhereClauseforEncounterAppointments(sStmt As String, _
                                                     Optional EventSubTypeNCID As String) As Boolean
'-- Assumes alias for event table is 'e' and Action_Info is 'ai'

If Not DebugMode Then On Error GoTo ErrHandler

Dim i As Integer

BuildWhereClauseforEncounterAppointments = False

    '-- Clinical_Event_Header (Make sure status is 'Complete')
    '-------------------------------------------------------------------------------------------
    sStmt = sStmt & " c.status_ncid = 205657 and c.event_id = e.event_id "


    '-- SubType Event
    '--------------------------------------------------------------------------------------------
    If EventSubTypeNCID <> "" Then
        sStmt = sStmt & " and e.event_subtype_ncid = " & EventSubTypeNCID
    End If

    '-- Encounter and EventType
    '--------------------------------------------------------------------------------------------
    sStmt = sStmt & " and e.event_type_ncid = 209614 " & _
                    " and e.encounter_number = a.encounter_number "
    
    
'-- Filter by either MTF, Clinic, Provider, or PCM
'-------------------------------------------------------------------------------------------
    
    '-- PCM
    '-------------------------------------------------------------------------------------------
    If moReportItem.Criteria1 = PCM Then
        Dim sPCMCode As String
        
        If moReportItem.CriteriaSel1.Item(1) Is Nothing Then
            sPCMCode = "0"
        Else
            On Error Resume Next
            sPCMCode = moReportItem.CriteriaSel1.Item(1).PCMCode
            If Err <> 0 Then sPCMCode = "0"
        End If
        
        sStmt = sStmt & " and exists " & _
                        "(select 'x' " & _
                        " from mmi_generic_id mg " & _
                        " where " & _
                        " mg.id_value = '" & sPCMCode & "' and " & _
                        " mg.id_type_ncid = 14501724 and " & _
                        " mg.unit_number = a.patient_unit_number ) "
    End If
    
    '-- Provider
    '-------------------------------------------------------------------------------------------
    If moReportItem.Criteria1 = Provider Then
        sStmt = sStmt & " and a.clinician_ncid = " & moReportItem.CriteriaSel1.Item(1).NCID & " "
    End If

    '-- Clinic
    '--------------------------------------------------------------------------------------------
    If moReportItem.Criteria1 = Clinic Then
        sStmt = sStmt & " and a.clinic_ncid in ("
    
        For i = 1 To moReportItem.CriteriaSel1.Count
            If i <> moReportItem.CriteriaSel1.Count Then
                sStmt = sStmt & moReportItem.CriteriaSel1.Item(i).NCID & ","
            Else
                sStmt = sStmt & moReportItem.CriteriaSel1.Item(i).NCID & ") "
            End If
        Next
    End If

    '-- MTF
    '-------------------------------------------------------------------------------------------
    If moReportItem.Criteria1 = MTF Then
       sStmt = sStmt & " and a.facility_ncid  = " & moReportItem.CriteriaSel1.Item(1).NCID & " "
    End If

    '-- Event
    '-------------------------------------------------------------------------------------------
    sStmt = sStmt & " and (a.APPT_DATE_TIME >= " & GetOracleStartDate(moReportItem.DateFrom) & _
                    " and a.APPT_DATE_TIME <= " & GetOracleEndDate(moReportItem.DateTo) & ")"

BuildWhereClauseforEncounterAppointments = True
GoTo CleanUp

ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "ReportOps.BuildWhereClauseforEncounterAppointments", "frmReports", vbExclamation)
CleanUp:
    Exit Function
End Function

Private Function DeleteFromAllTables() As Boolean

If Not DebugMode Then On Error GoTo ErrHandler

DeleteFromAllTables = False

If Not moTableOps.DeleteFromTable("MTF") Then GoTo CleanUp
If Not moTableOps.DeleteFromTable("Clinic") Then GoTo CleanUp
If Not moTableOps.DeleteFromTable("Provider") Then GoTo CleanUp
If Not moTableOps.DeleteFromTable("Encounter") Then GoTo CleanUp
If Not moTableOps.DeleteFromTable("EncounterDiagnosis") Then GoTo CleanUp
If Not moTableOps.DeleteFromTable("Appointment") Then GoTo CleanUp
If Not moTableOps.DeleteFromTable("Category") Then GoTo CleanUp
If Not moTableOps.DeleteFromTable("EncounterProcedure") Then GoTo CleanUp
If Not moTableOps.DeleteFromTable("Diagnosis") Then GoTo CleanUp
If Not moTableOps.DeleteFromTable("EandM") Then GoTo CleanUp
If Not moTableOps.DeleteFromTable("Disposition") Then GoTo CleanUp
If Not moTableOps.DeleteFromTable("Consult_Order") Then GoTo CleanUp
If Not moTableOps.DeleteFromTable("Patient") Then GoTo CleanUp
If Not moTableOps.DeleteFromTable("Sponsor") Then GoTo CleanUp
If Not moTableOps.DeleteFromTable("Preventive") Then GoTo CleanUp
If Not moTableOps.DeleteFromTable("LabReport") Then GoTo CleanUp
If Not moTableOps.DeleteFromTable("LabReport") Then GoTo CleanUp
If Not moTableOps.DeleteFromTable("RadReport") Then GoTo CleanUp
If Not moTableOps.DeleteFromTable("DiagnosisReport") Then GoTo CleanUp
If Not moTableOps.DeleteFromTable("ProcedureReport") Then GoTo CleanUp
If Not moTableOps.DeleteFromTable("PrimaryDiagnosisReport") Then GoTo CleanUp
If Not moTableOps.DeleteFromTable("EandMReport") Then GoTo CleanUp
If Not moTableOps.DeleteFromTable("DispositionReport") Then GoTo CleanUp
If Not moTableOps.DeleteFromTable("MedReport") Then GoTo CleanUp
If Not moTableOps.DeleteFromTable("ConsultReport") Then GoTo CleanUp
If Not moTableOps.DeleteFromTable("AlphabeticPatientReport") Then GoTo CleanUp
If Not moTableOps.DeleteFromTable("ADSDiagnosisReport") Then GoTo CleanUp
If Not moTableOps.DeleteFromTable("TPC") Then GoTo CleanUp
If Not moTableOps.DeleteFromTable("ProcedureReport") Then GoTo CleanUp
If Not moTableOps.DeleteFromTable("Procedures") Then GoTo CleanUp
If Not moTableOps.DeleteFromTable("ADSEandMSubReport") Then GoTo CleanUp
If Not moTableOps.DeleteFromTable("ADSProcedureSubReport") Then GoTo CleanUp
If Not moTableOps.DeleteFromTable("ADSOutpatientWorkload") Then GoTo CleanUp
If Not moTableOps.DeleteFromTable("ADSCategoryReport") Then GoTo CleanUp
If Not moTableOps.DeleteFromTable("ADSDispositionReport") Then GoTo CleanUp
If Not moTableOps.DeleteFromTable("DiscontinuedServicesReport") Then GoTo CleanUp
If Not moTableOps.DeleteFromTable("ReadinessReport") Then GoTo CleanUp
If Not moTableOps.DeleteFromTable("MammographyHaveNotReport") Then GoTo CleanUp
If Not moTableOps.DeleteFromTable("PapSmearHaveNotReport") Then GoTo CleanUp
If Not moTableOps.DeleteFromTable("ADSPatientEncounterSubDiagnosis") Then GoTo CleanUp
If Not moTableOps.DeleteFromTable("ADSPatientEncounterSubProcedures") Then GoTo CleanUp
If Not moTableOps.DeleteFromTable("ADSPatientEncounterSubEM") Then GoTo CleanUp
If Not moTableOps.DeleteFromTable("WorkloadReport") Then GoTo CleanUp
If Not moTableOps.DeleteFromTable("ReadinessImmunReport") Then GoTo CleanUp
DeleteFromAllTables = True
GoTo CleanUp

ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "ReportOps.RemoveFromAllTables", "ReportOps", vbExclamation)
CleanUp:
    Exit Function
End Function

Private Function DataGetTobaccoUsersAdvisingToQuit(HeaderColl As Collection) As Boolean

If Not DebugMode Then On Error GoTo ErrHandler
DataGetTobaccoUsersAdvisingToQuit = False

Dim LNumerator As Long
Dim LDenominator As Long


Dim sStmt As String
Dim sWhereCriteria As String
Dim sFromCriteria As String



    If Not GetMTForPCM_SQL(sFromCriteria, sWhereCriteria) Then GoTo CleanUp
    

'-- Get Denominator
'---------------------------------------------------------------------------

    sStmt = "select count(*) Denominator " & _
            "from mmi_id Id " & sFromCriteria & _
            "Where "
    '-- Check for Coupler Event where user responded 'Yes' to smoking 100 cigarettes
    sStmt = sStmt & _
            "(exists " & _
                "(select e.event_id " & _
                " from observation_value o, clinical_observation c, event e " & _
                " where " & _
                        "o.value_ncid = 34599 and " & _
                        "o.observation_seq = c.observation_seq and " & _
                        "o.event_id = c.event_id and " & _
                        "c.observation_ncid = 204771 and " & _
                        "c.event_id = e.event_id and " & _
                        GetOracleEndDate(moReportItem.DateTo) & " >= EVENT_START_GMTIME and " & _
                        "e.event_type_ncid = 204764 and " & _
                        "e.unit_number = id.unit_number) or "
    '-- Check for indicator that patient smokes in Vitals
    sStmt = sStmt & _
            "exists " & _
                "(select e.event_id " & _
                " from clinical_observation c, clinical_event_header ce, event e " & _
                " where " & _
                        "c.observation_ncid = 1604139 and " & _
                        "c.event_id = e.event_id and " & _
                        GetOracleEndDate(moReportItem.DateTo) & " >= EVENT_START_GMTIME and " & _
                        "ce.status_ncid = 1024 and " & _
                        "ce.event_id = e.event_id and " & _
                        "e.event_type_ncid = 90782 and " & _
                        "e.unit_number = id.unit_number)) and "
    '-- Check if patient was entered into the CHCS system befor the 'TO' date
    sStmt = sStmt & _
            "(exists " & _
                "(select ma.unit_number " & _
                " from mmi_audit ma " & _
                " where " & _
                        GetOracleEndDate(moReportItem.DateTo) & " >= TX_TIME and " & _
                        "ma.unit_number = id.unit_number) and "
    '-- Check if patient was >= 17 years before the 'TO' date
    sStmt = sStmt & _
            "((((" & GetOracleEndDate(moReportItem.DateTo) & " - id.birth_date)/365) >= 17.0) "
    
    '-- Check if patient has records at specified MTF or is patieint of specified PCM
    sStmt = sStmt & _
            " and " & sWhereCriteria & "))"

    goDBSqlOps.Execute sStmt
    
    LDenominator = goDBSqlOps.Value("Denominator")


'-- Get Numerator
'---------------------------------------------------------------------------
    sStmt = "select count(*) Numerator " & _
            "from mmi_id Id " & sFromCriteria & _
            "Where "
    '-- Check if patient indicated Anti-Tobacco Counseling in HEAR
    sStmt = sStmt & _
            "(exists " & _
                "(select e.event_id " & _
                " from observation_value o, clinical_observation c, event e " & _
                " where " & _
                        "o.value_ncid = 34634 and " & _
                        "o.observation_seq = c.observation_seq and " & _
                        "o.event_id = c.event_id and " & _
                        "c.observation_ncid = 204771 and " & _
                        "c.event_id = e.event_id and " & _
                        GetOracleEndDate(moReportItem.DateTo) & " >= EVENT_START_GMTIME and " & _
                        "e.event_type_ncid = 204764 and " & _
                        "e.unit_number = id.unit_number) or "
    '-- Check if patient received Anit-Tobacco Counseling
    sStmt = sStmt & _
            "exists " & _
                "(select e.event_id " & _
                " from observation_value o, clinical_observation c, event e " & _
                " where " & _
                        "o.value_ncid = 207675 and " & _
                        "o.observation_seq = c.observation_seq and " & _
                        "o.event_id = c.event_id and " & _
                        "c.observation_ncid = 162045 and " & _
                        "c.event_id = e.event_id and " & _
                        GetOracleEndDate(moReportItem.DateTo) & " >= EVENT_START_GMTIME and " & _
                        "e.event_type_ncid = 162156 and " & _
                        "e.unit_number = id.unit_number)) and "
    '-- Check for Coupler Event where user responded 'Yes' to smoking 100 cigarettes
    sStmt = sStmt & _
            "((exists " & _
                "(select e.event_id " & _
                " from observation_value o, clinical_observation c, event e " & _
                " where " & _
                        "o.value_ncid = 34599 and " & _
                        "o.observation_seq = c.observation_seq and " & _
                        "o.event_id = c.event_id and " & _
                        "c.observation_ncid = 204771 and " & _
                        "c.event_id = e.event_id and " & _
                        GetOracleEndDate(moReportItem.DateTo) & " >= EVENT_START_GMTIME and " & _
                        "e.event_type_ncid = 204764 and " & _
                        "e.unit_number = id.unit_number) or "
    '-- Check for indicator that patient smokes in Vitals
    sStmt = sStmt & _
            "exists " & _
                "(select e.event_id " & _
                " from clinical_observation c, event e " & _
                " where " & _
                        "c.observation_ncid = 1604139 and " & _
                        "c.event_id = e.event_id and " & _
                        GetOracleEndDate(moReportItem.DateTo) & " >= EVENT_START_GMTIME and " & _
                        "e.event_type_ncid = 90782 and " & _
                        "e.unit_number = id.unit_number)) and "

    '-- Check if patient was entered into the CHCS system befor the 'TO' date
    sStmt = sStmt & _
            "(exists " & _
                "(select ma.unit_number " & _
                " from mmi_audit ma " & _
                " where " & _
                        GetOracleEndDate(moReportItem.DateTo) & " >= TX_TIME and " & _
                        "ma.unit_number = id.unit_number) and "

    '-- Check if patient was >= 17 years before the 'TO' date
    sStmt = sStmt & _
            "((((" & GetOracleEndDate(moReportItem.DateTo) & " - id.birth_date)/365) >= 17.0) "

    '-- Check if patient has records at specified MTF or is patieint of specified PCM
    sStmt = sStmt & _
            " and " & sWhereCriteria & ")))"

    goDBSqlOps.Execute sStmt
    
    LNumerator = goDBSqlOps.Value("Numerator")

'-- Set Report Field (Formula Field) Information
'---------------------------------------------------------------------------
    
    HeaderColl.Add CrystalQuote(CStr(LNumerator)), "Numerator"
    HeaderColl.Add CrystalQuote(CStr(LDenominator)), "Denominator"
    HeaderColl.Add CrystalQuote(moReportItem.CriteriaSel1.Item(1).Name), "Facility"

DataGetTobaccoUsersAdvisingToQuit = True
GoTo CleanUp
ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "ReportOps.DataGetTobaccoUsersAdvisingToQuit", App.Title, vbExclamation)
CleanUp:
    Exit Function
End Function


Private Function DataGetCholesterolScreening(HeaderColl As Collection) As Boolean

'94129 - Lipid Panel Complete, Serum
'11158 - Plasma Total Cholesterol
'67507 - Serum HDL
If Not DebugMode Then On Error GoTo ErrHandler
DataGetCholesterolScreening = False

Dim LNumerator As Long
Dim LDenominator As Long
Dim sStmt As String

Dim sWhereSQL As String
Dim sFromSQL As String

    If Not GetMTForPCM_SQL(sFromSQL, sWhereSQL) Then GoTo CleanUp


'-- Get Denominator
'---------------------------------------------------------------------------
    sStmt = "select count(*) Denominator " & vbNewLine & _
            "from mmi_id id " & sFromSQL & _
            "Where " & vbNewLine

    '-- Check if patient entered in system before 'TO' date
    sStmt = sStmt & _
            "exists " & vbNewLine & _
                "(select ma.unit_number " & vbNewLine & _
                " from mmi_audit ma " & vbNewLine & _
                " where " & vbNewLine & _
                        GetOracleEndDate(moReportItem.DateTo) & " >= TX_TIME and " & vbNewLine & _
                        "ma.unit_number = id.unit_number) and " & vbNewLine
    '-- Check if patient is a Male >= 35 or Female >= 45
    sStmt = sStmt & _
            "((((" & GetOracleEndDate(moReportItem.DateTo) & " - id.birth_date)/365) >= 35.0 and " & vbNewLine & _
            "id.sex = 'M') or " & vbNewLine & _
            "(((" & GetOracleEndDate(moReportItem.DateTo) & " - id.birth_date)/365) >= 45.0 and " & vbNewLine & _
            "id.sex = 'F')) " & vbNewLine
    
'    '-- Check if patient has records at specified MTF or is patieint of specified PCM
    sStmt = sStmt & _
            " and " & sWhereSQL
    
    
    goDBSqlOps.Execute sStmt
    
    LDenominator = goDBSqlOps.Value("Denominator")

'-- Get Numerator
'---------------------------------------------------------------------------
    sStmt = "select count(*) Numerator " & vbNewLine & _
            "from mmi_id id " & sFromSQL & _
            "Where " & vbNewLine
    '-- Check for Coupler Event where user has had cholesterol checked
    '-- 34799 Less than a year ago
    '-- 34800 1-2 years ago
    '-- 34801 3-4 years ago
    '-- 34802 5 or more years ago
    sStmt = sStmt & _
            "(exists " & vbNewLine & _
                "(select e.event_id " & vbNewLine & _
                " from clinical_observation c, observation_value o, event e " & vbNewLine & _
                " where " & vbNewLine & _
                        "o.value_ncid in (34799,34800,34801,34802) and " & vbNewLine & _
                        "o.observation_seq = c.observation_seq and " & vbNewLine & _
                        "o.event_id = c.event_id and " & vbNewLine & _
                        "c.observation_ncid = 204771 and " & vbNewLine & _
                        "c.event_id = e.event_id and " & vbNewLine & _
                        GetOracleEndDate(moReportItem.DateTo) & " >= EVENT_START_GMTIME and " & vbNewLine & _
                        "e.event_type_ncid = 204764 and " & vbNewLine & _
                        "e.unit_number = id.unit_number) or " & vbNewLine
    '-- Check for Cholesterol Lab Test
    sStmt = sStmt & _
            "exists " & vbNewLine & _
                "(select le.event_id " & vbNewLine & _
                " from lab_event_type le" & vbNewLine & _
                " where " & vbNewLine & _
                    "((" & GetOracleEndDate(moReportItem.DateTo) & " - le.EVENT_START_GMTIME)/365) <= 5.0 and " & vbNewLine & _
                    GetOracleEndDate(moReportItem.DateTo) & " >= le.EVENT_START_GMTIME and " & vbNewLine & _
                "   le.test_ncid in (94129,11158,67507) and " & vbNewLine & _
                "   le.unit_number = id.unit_number)) and " & vbNewLine

    '-- Check if patient was entered before the 'TO' date
    sStmt = sStmt & _
            "(exists " & vbNewLine & _
                "(select ma.unit_number " & vbNewLine & _
                " from mmi_audit ma " & vbNewLine & _
                " where " & vbNewLine & _
                        GetOracleEndDate(moReportItem.DateTo) & " >= TX_TIME and " & vbNewLine & _
                        "ma.unit_number = id.unit_number) and " & vbNewLine

    '-- Check if patient is a Male >= 35 or Female >= 45
    sStmt = sStmt & _
            "((((" & GetOracleEndDate(moReportItem.DateTo) & " - id.birth_date)/365) >= 35.0 and " & vbNewLine & _
            "id.sex = 'M') or " & vbNewLine & _
            "(((" & GetOracleEndDate(moReportItem.DateTo) & " - id.birth_date)/365) >= 45.0 and " & vbNewLine & _
            "id.sex = 'F')) " & vbNewLine

    '-- Check if patient has records at specified MTF or is patieint of specified PCM
    sStmt = sStmt & _
            " and " & sWhereSQL & ")"
    goDBSqlOps.Execute sStmt
    
    LNumerator = goDBSqlOps.Value("Numerator")

'-- Set Report Field (Formual Field) Information
'---------------------------------------------------------------------------

    HeaderColl.Add CrystalQuote(CStr(LNumerator)), "Numerator"
    HeaderColl.Add CrystalQuote(CStr(LDenominator)), "Denominator"
    HeaderColl.Add CrystalQuote(moReportItem.CriteriaSel1.Item(1).Name), "Facility"

DataGetCholesterolScreening = True
GoTo CleanUp
ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "ReportOps.DataGetCholesterolScreening", App.Title, vbExclamation)
CleanUp:
    Exit Function
End Function















Private Function DataGetADSPatientEncounter(HeaderColl As Collection) As Boolean

    If Not DebugMode Then On Error GoTo ErrHandler
    DataGetADSPatientEncounter = False

    Dim sStmt As String
    Dim i As Integer
    Dim rs As ADODB.Recordset
    

    If Not moReportItem.CriteriaSel3 Is Nothing Then
       If Not moTableOps.OpenTable("Patient", rs) Then GoTo CleanUp
       If Not moTableOps.SaveRec(moReportItem.CriteriaSel3.Item(1), rs) Then GoTo CleanUp
       If Not moTableOps.CloseTable(rs) Then GoTo CleanUp
    Else
       DataGetADSPatientEncounter = DataGetADSMultiPatientEncounter(HeaderColl)
       GoTo CleanUp
    End If
    
    'SCR-46324 RBELL 12/4/3 Added filter of ov.value_ncid to prevent Raw Medcin ID from showing up
    '-- Get Diagnosis ----------------------------------------------------------------------------------------------
    sStmt = "select a.appt_date_time as ApptDate, a.patient_unit_number, ov.value_ncid DiagnosisNCID " & vbLf & _
            "from observation_value ov, clinical_observation co, event e, appointment a " & vbLf & _
            "where ov.observation_seq = co.observation_seq and " & vbLf & _
            "ov.value_ncid <> 14516241 and " & vbLf & _
            "ov.event_id = co.event_id and " & vbLf & _
            "co.observation_ncid = 658 and " & vbLf & _
            "co.event_id = e.event_id and " & vbLf & _
            "e.event_type_ncid = 90803 and " & vbLf & _
            "e.Encounter_Number = a.Encounter_Number " & vbLf & _
            " and (a.appt_date_time >= " & GetOracleStartDate(moReportItem.DateFrom) & vbLf & _
            " and a.appt_date_time <= " & GetOracleEndDate(moReportItem.DateTo) & ") " & vbLf & _
            " and a.clinic_ncid = " & moReportItem.CriteriaSel1.Item(1).NCID & " " & vbLf & _
            " and a.clinician_ncid = " & moReportItem.CriteriaSel2.Item(1).NCID & vbLf & _
            " and a.patient_unit_number = " & moReportItem.CriteriaSel3.Item(1).UnitNumber
    goDBSqlOps.Execute sStmt
    If Not (goDBSqlOps.BOF And goDBSqlOps.EOF) Then
       If Not moTableOps.OpenTable("ADSPatientEncounterSubDiagnosis", rs) Then GoTo CleanUp
       Do While Not goDBSqlOps.EOF
           rs.AddNew
           rs("PatientID") = goDBSqlOps.Value("patient_unit_number")
           rs("DiagnosisNCID") = goDBSqlOps.Value("DiagnosisNCID")
           rs("ApptDate") = Format(goDBSqlOps.Value("ApptDate"), gsDefaultFormat)
           rs.Update
           goDBSqlOps.MoveNext
       Loop
       If Not moTableOps.GetDiagnosisTableData("ADSPatientEncounterSubDiagnosis", "DiagnosisNCID") Then GoTo CleanUp
    End If
    
    If Not moTableOps.CloseTable(rs) Then GoTo CleanUp

    '-- Get Procedures ----------------------------------------------------------------------------------------------
    sStmt = "select a.appt_date_time as ApptDate, a.patient_unit_number, ov.value_ncid ProcedureNCID " & vbLf & _
            "from observation_value ov, clinical_observation co, event e, appointment a " & vbLf & _
            "where ov.observation_seq = co.observation_seq and " & vbLf & _
            "ov.event_id = co.event_id and " & vbLf & _
            "co.observation_ncid = 41 and " & vbLf & _
            "co.event_id = e.event_id and " & vbLf & _
            "e.event_type_ncid = 90803 and " & vbLf & _
            "e.Encounter_Number = a.Encounter_Number " & vbLf & _
            " and (a.appt_date_time >= " & GetOracleStartDate(moReportItem.DateFrom) & vbLf & _
            " and a.appt_date_time <= " & GetOracleEndDate(moReportItem.DateTo) & ") " & vbLf & _
            " and a.clinic_ncid = " & moReportItem.CriteriaSel1.Item(1).NCID & " " & vbLf & _
            " and a.clinician_ncid = " & moReportItem.CriteriaSel2.Item(1).NCID & vbLf & _
            " and a.patient_unit_number = " & moReportItem.CriteriaSel3.Item(1).UnitNumber
        
    goDBSqlOps.Execute sStmt
    If Not (goDBSqlOps.BOF And goDBSqlOps.EOF) Then
       If Not moTableOps.OpenTable("ADSPatientEncounterSubProcedures", rs) Then GoTo CleanUp
       Do While Not goDBSqlOps.EOF
           rs.AddNew
           rs("PatientID") = goDBSqlOps.Value("patient_unit_number")
           rs("ProcedureNCID") = goDBSqlOps.Value("ProcedureNCID")
           rs("ApptDate") = Format(goDBSqlOps.Value("ApptDate"), gsDefaultFormat)
           rs.Update
           goDBSqlOps.MoveNext
       Loop
       If Not moTableOps.GetProcedureTableData("ADSPatientEncounterSubProcedures", "ProcedureNCID") Then GoTo CleanUp
    End If
    
    If Not moTableOps.CloseTable(rs) Then GoTo CleanUp

    '-- Get E&M ----------------------------------------------------------------------------------------------
    sStmt = "select a.appt_date_time as ApptDate, a.patient_unit_number, ov.value_ncid EMNCID " & vbLf & _
            "from observation_value ov, clinical_observation co, event e, appointment a " & vbLf & _
            "where ov.observation_seq = co.observation_seq and " & vbLf & _
            "ov.event_id = co.event_id and " & vbLf & _
            "co.observation_ncid = 41 and " & vbLf & _
            "co.event_id = e.event_id and " & vbLf & _
            "e.event_subtype_ncid = 184917 and " & vbLf & _
            "e.event_type_ncid = 209614 and " & vbLf & _
            "e.Encounter_Number = a.Encounter_Number " & vbLf & _
            " and (a.appt_date_time >= " & GetOracleStartDate(moReportItem.DateFrom) & vbLf & _
            " and a.appt_date_time <= " & GetOracleEndDate(moReportItem.DateTo) & ") " & vbLf & _
            " and a.clinic_ncid = " & moReportItem.CriteriaSel1.Item(1).NCID & " " & vbLf & _
            " and a.clinician_ncid = " & moReportItem.CriteriaSel2.Item(1).NCID & vbLf & _
            " and a.patient_unit_number = " & moReportItem.CriteriaSel3.Item(1).UnitNumber
            
    goDBSqlOps.Execute sStmt
    If Not (goDBSqlOps.BOF And goDBSqlOps.EOF) Then
       If Not moTableOps.OpenTable("ADSPatientEncounterSubEM", rs) Then GoTo CleanUp
       Do While Not goDBSqlOps.EOF
           rs.AddNew
           rs("PatientID") = goDBSqlOps.Value("patient_unit_number")
           rs("EMNCID") = goDBSqlOps.Value("EMNCID")
           rs("ApptDate") = Format(goDBSqlOps.Value("ApptDate"), gsDefaultFormat)
           rs.Update
           goDBSqlOps.MoveNext
       Loop
       If Not moTableOps.GetEandMTableData("ADSPatientEncounterSubEM", "EMNCID") Then GoTo CleanUp
    End If
    If Not moTableOps.CloseTable(rs) Then GoTo CleanUp
    
    DataGetADSPatientEncounter = True
    GoTo CleanUp

ErrHandler:
    If Err = -2147220874 Then
        MsgBxARMd "The selected report has timed out, please re-run this report again.", vbOKOnly + vbInformation
        Err.Clear
    Else
        Call gobjshared.ShowVBError(Err.Number, Err.Description, "ReportOps.DataGetADSPatientEncounter", App.Title, vbExclamation)
    End If
CleanUp:
    Set rs = Nothing

End Function

Private Function GetReceivingEntityName(sRecvEntityNCID As String, sRequestTo As String) As String

If Not DebugMode Then On Error GoTo ErrHandler

Dim sStmt As String

sStmt = "select trim(representation) description " & _
        "from rsform_context rc, rsform r " & _
        "Where " & _
            "rc.context_ncid = 2000 and " & _
            "rc.rsform_id = r.rsform_id and " & _
            "r.NCID = " & sRecvEntityNCID

goDBSqlOps.Execute sStmt
If goDBSqlOps.BOF And goDBSqlOps.EOF Then
    GetReceivingEntityName = "Name not found."
Else
    GetReceivingEntityName = goDBSqlOps.Value("Description")
End If

GoTo CleanUp

ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "ReportOps.GetReceivingEntityName", App.Title, vbExclamation)
CleanUp:
End Function


Private Function DataGetADSProcedure(HeaderColl As Collection) As Boolean
If Not DebugMode Then On Error GoTo ErrHandler
DataGetADSProcedure = False

Dim sStmt As String
Dim i As Integer
Dim rs As ADODB.Recordset
    
'-- Get Procedures
'------------------------------------------------------------------------------
sStmt = "SELECT  /*+ index (ai action_info_x4) */   " & _
                "ov.value_ncid ProcedureNCID, " & _
                "e.event_id, " & _
                "e.unit_number patient_unit_number, " & _
                "e.encounter_number, " & _
                "e.facility_ncid mtf_ncid, " & _
                "ai.poc_location_ncid clinic_ncid, " & _
                "ai.clinician_id_ncid provider_ncid " & _
                "FROM observation_value ov, event e2, action_info ai, clinical_event_header c, clinical_observation co, event e "
sStmt = sStmt & _
        "Where " & _
            "ov.observation_seq = co.observation_seq and " & _
            "ov.event_id = co.event_id and " & _
            "41 = co.observation_ncid and "

    '-- In this case e2 represents a procedure event
    sStmt = sStmt & _
            "co.event_id = e2.event_id and " & _
            "90803 = e2.event_type_ncid and "

            
sStmt = sStmt & "e2.Encounter_Number = e.Encounter_Number and " & _
                "" 'c.event_id = e.event_id and

    sStmt = sStmt & " ai.clinician_id_ncid = " & moReportItem.CriteriaSel2.Item(1).NCID & " "

    'If Not BuildWhereClauseforLocation(sStmt) Then GoTo CleanUp
    'If Not BuildWhereClauseforEncounterEvents(sStmt) Then GoTo CleanUp
    If Not BuildWhereClauseforEncounterEvents2(sStmt, "209683") Then GoTo CleanUp

    If moReportItem.Criteria3 <> NoCriteria Then
        sStmt = sStmt & "and e.unit_number = " & moReportItem.CriteriaSel3.Item(1).UnitNumber & " "
    End If

    goDBSqlOps.Execute sStmt
    
    If Not moTableOps.OpenTable("ADSProcedureSubReport", rs) Then GoTo CleanUp

    Do While Not goDBSqlOps.EOF
        rs.AddNew
        rs("MTF_NCID") = goDBSqlOps.Value("MTF_NCID")
        rs("ClinicNCID") = goDBSqlOps.Value("Clinic_NCID")
        rs("ProviderNCID") = goDBSqlOps.Value("Provider_NCID")
        rs("ProcedureNCID") = goDBSqlOps.Value("ProcedureNCID")
        rs.Update
        goDBSqlOps.MoveNext
    Loop

    If Not moTableOps.CloseTable(rs) Then GoTo CleanUp
    
    If Not moTableOps.GetProcedureTableData("ADSProcedureSubReport", "ProcedureNCID") Then GoTo CleanUp
    
    DoEvents
    
'-- Get E&M Codes
'------------------------------------------------------------------------------=
sStmt = "SELECT /*+ index (ai action_info_x4) */  ov.value_ncid EandMNCID, " & _
                "e.event_id, " & _
                "e.unit_number patient_unit_number, " & _
                "e.encounter_number, " & _
                "e.facility_ncid mtf_ncid, " & _
                "ai.poc_location_ncid clinic_ncid, " & _
                "ai.clinician_id_ncid provider_ncid, " & _
                "e.event_start_gmtime encounter_date " & _
        "FROM observation_value ov,  clinical_observation co, " & _
              " action_info ai, clinical_event_header c, event e "
              '"FROM observation_value ov, action_info ai, clinical_event_header c, , event e "
sStmt = sStmt & " Where "

        If moReportItem.Criteria2 = Diagnosis Then
            '-- In this case e2 represents a Diagnosis event that is linked to a _
                procedure via semantic links
            sStmt = sStmt & _
                    " exists ( " & _
                    "select e2.event_id " & _
                    "from " & _
                            " clinical_observation co2, " & _
                            " observation_value ov2, " & _
                            " event e2 " & _
                    "where " & _
                        moReportItem.CriteriaSel2.Item(1).NCID & " =  ov2.value_ncid and " & _
                        "ov2.observation_seq = co2.observation_seq and " & _
                        "ov2.event_id = co2.event_id and " & _
                        "658 = co2.observation_ncid and co2.event_id = e2.event_id and " & _
                        "90803 = e2.event_type_ncid and " & _
                        "e2.Encounter_Number = e.Encounter_Number) and "

        End If

sStmt = sStmt & _
            "ov.observation_seq = co.observation_seq and " & _
            "ov.event_id = co.event_id and " & _
            "co.observation_ncid = 41 and " & _
            "co.event_id = e.event_id and " & _
            "c.event_id = e.event_id and "

    sStmt = sStmt & " ai.clinician_id_ncid = " & moReportItem.CriteriaSel2.Item(1).NCID & " "
    If moReportItem.Criteria3 <> NoCriteria Then
        sStmt = sStmt & " and e.unit_number = " & moReportItem.CriteriaSel3.Item(1).UnitNumber & " "
    End If

    'If Not BuildWhereClauseforLocation(sStmt) Then GoTo CleanUp
    'If Not BuildWhereClauseforEncounterEvents(sStmt) Then GoTo CleanUp
    If Not BuildWhereClauseforEncounterEvents2(sStmt, "184917") Then GoTo CleanUp
    
    goDBSqlOps.Execute sStmt
    If goDBSqlOps.BOF And goDBSqlOps.EOF Then
        DataGetADSProcedure = True
        GoTo CleanUp
    End If
    
    If Not moTableOps.OpenTable("ADSEandMSubReport", rs) Then GoTo CleanUp

    Do While Not goDBSqlOps.EOF
        rs.AddNew
        rs("MTF_NCID") = goDBSqlOps.Value("MTF_NCID")
        rs("ClinicNCID") = goDBSqlOps.Value("Clinic_NCID")
        rs("ProviderNCID") = goDBSqlOps.Value("Provider_NCID")
        rs("EandMNCID") = goDBSqlOps.Value("EandMNCID")
        rs.Update
        goDBSqlOps.MoveNext
    Loop
    If Not moTableOps.CloseTable(rs) Then GoTo CleanUp
    
    If Not moTableOps.GetEandMTableData("ADSEandMSubReport", "EandMNCID") Then GoTo CleanUp
  
    DoEvents

DataGetADSProcedure = True
GoTo CleanUp
ErrHandler:
    If Err = -2147220874 Then
        MsgBxARMd "The selected report has timed out, please re-run this report again.", vbOKOnly + vbInformation
        Err.Clear
    Else
        Call gobjshared.ShowVBError(Err.Number, Err.Description, "ReportOps.DataGetADSProcedure", App.Title, vbExclamation)
    End If
CleanUp:
    Set rs = Nothing
    Exit Function
End Function








Private Function DataGetADSDisposition(HeaderColl As Collection) As Boolean


If Not DebugMode Then On Error GoTo ErrHandler
DataGetADSDisposition = False

Dim sStmt As String
Dim i As Integer
Dim rs As ADODB.Recordset
    
    
sStmt = "SELECT  /*+ ORDERED */ ov.value_ncid DispositionNCID, " & vbNewLine & _
                "e.event_id, " & vbNewLine & _
                "a.patient_unit_number, " & vbNewLine & _
                "a.encounter_number, " & vbNewLine & _
                "a.facility_ncid mtf_ncid, " & vbNewLine & _
                "a.clinic_ncid, " & vbNewLine & _
                "a.clinician_ncid provider_ncid, " & vbNewLine & _
                "a.appt_date_time encounter_date " & vbNewLine & _
        "FROM " & vbNewLine
              

        
sStmt = sStmt & "appointment a, event e, clinical_event_header c, event e2, " & vbNewLine & _
                "clinical_observation co, observation_value ov   " & vbNewLine

        If moReportItem.Criteria2 = Diagnosis Then
            sStmt = sStmt & ", semantic_link sl, semantic_link_element sle, " & vbNewLine & _
                            "clinical_observation co2, observation_value ov2 " & vbNewLine
        End If

sStmt = sStmt & " Where " & vbNewLine

        If moReportItem.Criteria2 = Diagnosis Then
            sStmt = sStmt & _
            "ov2.value_ncid = " & moReportItem.CriteriaSel2.Item(1).NCID & vbNewLine & _
            " and ov2.observation_seq = co2.observation_seq and " & vbNewLine & _
            "ov2.event_id = co2.event_id and " & vbNewLine & _
            "co2.observation_ncid = 105328 and " & vbNewLine & _
            "co2.event_id = sle.link_event_id and " & vbNewLine & _
            "sle.semantic_link_ord = sl.semantic_link_ord and " & vbNewLine & _
            "sle.event_id = sl.event_id and " & vbNewLine & _
            "sl.relationship_ncid = 205655 and " & vbNewLine & _
            "sl.event_id = e.event_id and " & vbNewLine
        End If

sStmt = sStmt & _
            "ov.observation_seq = co.observation_seq and " & vbNewLine & _
            "ov.event_id = co.event_id and " & vbNewLine & _
            "co.observation_ncid = 373 and " & vbNewLine & _
            "co.event_id = e2.event_id and " & vbNewLine & _
            "e2.event_subtype_ncid = 184917 and " & vbNewLine & _
            "e2.encounter_number = a.encounter_number " & vbNewLine
    
    If Not BuildWhereClauseforEncounterEvents4(sStmt, "209683") Then GoTo CleanUp

    goDBSqlOps.Execute sStmt
    If goDBSqlOps.BOF And goDBSqlOps.EOF Then
        DataGetADSDisposition = True
        GoTo CleanUp
    End If
    
    If Not moTableOps.OpenTable("ADSDispositionReport", rs) Then GoTo CleanUp
    Do While Not goDBSqlOps.EOF
        rs.AddNew
        rs("MTF_NCID") = goDBSqlOps.Value("MTF_NCID")
        rs("ClinicNCID") = goDBSqlOps.Value("Clinic_NCID")
        rs("ProviderNCID") = goDBSqlOps.Value("Provider_NCID")
        rs("DispositionNCID") = goDBSqlOps.Value("DispositionNCID")
        rs("UnitNumber") = goDBSqlOps.Value("patient_unit_number")
        rs.Update
        goDBSqlOps.MoveNext
    Loop

    If Not moTableOps.CloseTable(rs) Then GoTo CleanUp

    If Not moTableOps.GetDispositionTableData("ADSDispositionReport", "DispositionNCID") Then GoTo CleanUp
    
    If Not moTableOps.GetClinicTableData("ADSDispositionReport", "ClinicNCID") Then GoTo CleanUp
    
    If Not moTableOps.GetPatientTableData("ADSDispositionReport", "UnitNumber") Then GoTo CleanUp
    
    DoEvents

DataGetADSDisposition = True
GoTo CleanUp
ErrHandler:
    If Err = -2147220874 Then
        MsgBxARMd "The selected report has timed out, please re-run this report again.", vbOKOnly + vbInformation
        Err.Clear
    Else
        Call gobjshared.ShowVBError(Err.Number, Err.Description, "ReportOps.DataGetDisposition", App.Title, vbExclamation)
    End If
CleanUp:
    Set rs = Nothing
    Exit Function
End Function
Private Function DataGetAlphabeticPatientList(HeaderColl As Collection) As Boolean

If Not DebugMode Then On Error GoTo ErrHandler
DataGetAlphabeticPatientList = False

Dim sStmt As String
Dim i As Integer
Dim rs As ADODB.Recordset
    
    
sStmt = "SELECT  " & vbNewLine & _
                "e.event_id, " & _
                "a.patient_unit_number, " & vbNewLine & _
                "a.encounter_number, " & vbNewLine & _
                "a.facility_ncid mtf_ncid, " & vbNewLine & _
                "a.clinic_ncid, " & vbNewLine & _
                "a.clinician_ncid provider_ncid, " & vbNewLine & _
                "a.appt_date_time encounter_date " & vbNewLine

              
sStmt = sStmt & "FROM  " & vbNewLine & _
                    "clinical_event_header c, event e, appointment a " & vbNewLine

sStmt = sStmt & "Where " & vbNewLine & " 1=1 " & vbNewLine


    If Not BuildWhereClauseforEncounterEvents4(sStmt, "209683") Then GoTo CleanUp
    'If Not BuildWhereClauseforEncounterEvents2(sStmt, "209683") Then GoTo CleanUp
    goDBSqlOps.Execute sStmt
    If goDBSqlOps.BOF And goDBSqlOps.EOF Then
        DataGetAlphabeticPatientList = True
        GoTo CleanUp
    End If
    
    If Not moTableOps.OpenTable("AlphabeticPatientReport", rs) Then GoTo CleanUp

    Do While Not goDBSqlOps.EOF
        rs.AddNew
        rs("MTF_NCID") = goDBSqlOps.Value("MTF_NCID")
        rs("ClinicNCID") = goDBSqlOps.Value("Clinic_NCID")
        rs("ProviderNCID") = goDBSqlOps.Value("Provider_NCID")
        rs("PatientUnitNumber") = goDBSqlOps.Value("patient_unit_number")
        rs.Update
        goDBSqlOps.MoveNext
    Loop
    If Not moTableOps.CloseTable(rs) Then GoTo CleanUp
    
    If Not moTableOps.GetPatientTableData("AlphabeticPatientReport", "PatientUnitNumber") Then GoTo CleanUp
    
    If moReportItem.GroupBy1 = Clinic Then
        If Not moTableOps.GetClinicTableData("AlphabeticPatientReport", "ClinicNCID") Then GoTo CleanUp
    End If
    If moReportItem.GroupBy1 = Provider Or moReportItem.GroupBy2 = Provider Then
        If Not moTableOps.GetProviderTableData("AlphabeticPatientReport", "ProviderNCID") Then GoTo CleanUp
    End If
    
    DoEvents
    
    
DataGetAlphabeticPatientList = True
GoTo CleanUp
ErrHandler:
    If Err = -2147220874 Then
        MsgBxARMd "The selected report has timed out, please re-run this report again.", vbOKOnly + vbInformation
        Err.Clear
    Else
        Call gobjshared.ShowVBError(Err.Number, Err.Description, "ReportOps.DataGetAlphabeticPatientList", App.Title, vbExclamation)
    End If
    'Resume Next
CleanUp:
    Set rs = Nothing

End Function
Private Function DataGetADSInsuranceIndicatorNotMarked(HeaderColl As Collection) As Boolean

If Not DebugMode Then On Error GoTo ErrHandler
DataGetADSInsuranceIndicatorNotMarked = False

Dim sStmt As String
Dim i As Integer
Dim rs As ADODB.Recordset

sStmt = "SELECT  " & vbNewLine & _
                "distinct e.unit_number patient_unit_number " & vbNewLine & _
        "FROM clinical_event_header c, event e, appointment a " & vbNewLine

sStmt = sStmt & " Where " & vbNewLine

sStmt = sStmt & "not exists " & vbNewLine & _
                "(select unit_number " & vbNewLine & _
                "from mmi_genericdata mg " & vbNewLine & _
                "where " & vbNewLine & _
                        "type_ncid = 90635 and " & vbNewLine & _
                        "unit_number = e.Unit_Number) " & vbNewLine

If Not BuildWhereClauseforEncounterEvents4(sStmt, "209683") Then GoTo CleanUp

    goDBSqlOps.Execute sStmt
    If goDBSqlOps.BOF And goDBSqlOps.EOF Then
        DataGetADSInsuranceIndicatorNotMarked = True
        GoTo CleanUp
    End If

    If Not moTableOps.OpenTable("AlphabeticPatientReport", rs) Then GoTo CleanUp

    Do While Not goDBSqlOps.EOF
        rs.AddNew
        rs("MTF_NCID") = goDBSqlOps.Value("MTF_NCID")
        rs("ClinicNCID") = goDBSqlOps.Value("Clinic_NCID")
        rs("ProviderNCID") = goDBSqlOps.Value("Provider_NCID")
        rs("PatientUnitNumber") = goDBSqlOps.Value("patient_unit_number")
        rs.Update
        goDBSqlOps.MoveNext
    Loop
    If Not moTableOps.CloseTable(rs) Then GoTo CleanUp
    
    If Not moTableOps.GetPatientTableData("AlphabeticPatientReport", "PatientUnitNumber") Then GoTo CleanUp
    
    DoEvents

DataGetADSInsuranceIndicatorNotMarked = True
GoTo CleanUp
ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "ReportOps.DataGetADSInsuranceIndicatorNotMarked", App.Title, vbExclamation)
CleanUp:
    Set rs = Nothing
    Exit Function
End Function







Private Function DataGetADSInsuranceIndicator(HeaderColl As Collection) As Boolean

If Not DebugMode Then On Error GoTo ErrHandler
DataGetADSInsuranceIndicator = False

Dim sStmt As String
Dim i As Integer
Dim LMarkedNo As Long
Dim LMarkedYes As Long



'-- Get Marked No
'---------------------------------------------------------------------------
sStmt = "SELECT /*+ ordered */ COUNT (DISTINCT a.patient_unit_number) AS total_count " & vbNewLine & _
        "FROM appointment a " & vbNewLine

sStmt = sStmt & " Where " & vbNewLine

sStmt = sStmt & "not exists " & vbNewLine & _
                "(select 'x' " & vbNewLine & _
                "from mmi_genericdata mg " & vbNewLine & _
                "where " & vbNewLine & _
                        "type_ncid = 90635 and " & vbNewLine & _
                        "unit_number = a.Patient_Unit_Number) And " & vbNewLine & _
                "a.encounter_status IN (4,5,6,7) And " & vbNewLine & _
                "a.encounter_number <> '0' " & vbNewLine

'-------------------------------
'- MTF
'-------------------------------
    If moReportItem.Criteria1 = MTF Then
        sStmt = sStmt & " and a.facility_ncid  = " & moReportItem.CriteriaSel1.Item(1).NCID & " " & vbNewLine
    End If
        
'-------------------------------
'- CLINIC
'-------------------------------
    If moReportItem.Criteria1 = Clinic Then
    
        
        sStmt = sStmt & " and a.clinic_ncid in ("
    
        For i = 1 To moReportItem.CriteriaSel1.Count
            If i <> moReportItem.CriteriaSel1.Count Then
                sStmt = sStmt & moReportItem.CriteriaSel1.Item(i).NCID & ","
            Else
                sStmt = sStmt & moReportItem.CriteriaSel1.Item(i).NCID & ") " & vbNewLine
            End If
        Next

    End If

'-------------------------------
'- DATE RANGE
'-------------------------------
    sStmt = sStmt & " and (a.APPT_DATE_TIME >= " & GetOracleStartDate(moReportItem.DateFrom) & vbNewLine & _
                    " and a.APPT_DATE_TIME <= " & GetOracleEndDate(moReportItem.DateTo) & ")"
                    
goDBSqlOps.Execute sStmt


LMarkedNo = goDBSqlOps.Value("total_count")
DoEvents

'-- Get Marked Yes
'---------------------------------------------------------------------------
sStmt = Replace(sStmt, "not exists", "exists")

goDBSqlOps.Execute sStmt

LMarkedYes = goDBSqlOps.Value("total_count")
DoEvents

'-- Set Report Field (Formual Field) Information
'---------------------------------------------------------------------------
 HeaderColl.Add CrystalQuote(CStr(LMarkedNo)), "MarkedNo"
 HeaderColl.Add CrystalQuote(CStr(LMarkedYes)), "MarkedYes"
 HeaderColl.Add CrystalQuote(moReportItem.CriteriaSel1.Item(1).Name), "Facility"


DataGetADSInsuranceIndicator = True
GoTo CleanUp
ErrHandler:
    If Err = -2147220874 Then
        MsgBxARMd "The selected report has timed out, please re-run this report again.", vbOKOnly + vbInformation
        Err.Clear
    Else
        Call gobjshared.ShowVBError(Err.Number, Err.Description, "ReportOps.DataGetADSInsuranceIndicator", App.Title, vbExclamation)
    End If
CleanUp:
    Exit Function
End Function

Private Function DataGetDD2766() As Boolean
On Error GoTo ErrHandler
    DataGetDD2766 = False
    Dim oForms As CHCSII_Forms.Forms
    Set oForms = New CHCSII_Forms.Forms
    Set oForms.Comm = cwobjCWComm

    'oForms.printDD2766 PrintPreview, moReportItem.CriteriaSel1.Item(1).UnitNumber
    If Not oForms.printDD2766(ReturnReportObject, moReportItem.CriteriaSel1.Item(1).UnitNumber) Then GoTo CleanUp
    Set moReport = oForms.ReportObject
    
    DataGetDD2766 = True
    GoTo CleanUp
ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "ReportOps.DataGetDD2766", App.Title, vbExclamation)
CleanUp:
    Set oForms = Nothing
End Function

Private Function DataGetCPGPOC() As Boolean
    Dim oForms As CHCSII_Forms.Forms
On Error GoTo ErrHandler
    DataGetCPGPOC = False
    Set oForms = New CHCSII_Forms.Forms
    Set oForms.Comm = cwobjCWComm
    If Not oForms.printCPG_POC(ReturnReportObject, moReportItem.CriteriaSel1.Item(1).UnitNumber, cwobjLogonEnviron.FacilityNCID) Then GoTo CleanUp
    Set moReport = oForms.ReportObject

    DataGetCPGPOC = True
    GoTo CleanUp

ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "ReportOps.DataGetCPGPOC", App.Title, vbExclamation)
CleanUp:
    Set oForms = Nothing
End Function

Private Function DataGetDD2813() As Boolean

On Error GoTo ErrHandler
    DataGetDD2813 = False

    Dim oForms As CHCSII_Forms.Forms
    Set oForms = New CHCSII_Forms.Forms
    Set oForms.Comm = cwobjCWComm

    'oForms.printDD2766 PrintPreview, moReportItem.CriteriaSel1.Item(1).UnitNumber
    If Not oForms.printDD2813(ReturnReportObject, moReportItem.CriteriaSel1.Item(1).UnitNumber) Then GoTo CleanUp
    Set moReport = oForms.ReportObject

    DataGetDD2813 = True
    GoTo CleanUp
ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "ReportOps.DataGetDD2813", App.Title, vbExclamation)
CleanUp:
    Set oForms = Nothing
    
End Function

Private Function DataGetDD689() As Boolean

On Error GoTo ErrHandler
    DataGetDD689 = False

    Dim oForms As CHCSII_Forms.Forms
    Set oForms = New CHCSII_Forms.Forms
    Set oForms.Comm = cwobjCWComm

    'oForms.printDD2766 PrintPreview, moReportItem.CriteriaSel1.Item(1).UnitNumber
    If Not oForms.printDD689(ReturnReportObject, moReportItem.CriteriaSel1.Item(1).UnitNumber) Then GoTo CleanUp
    Set moReport = oForms.ReportObject

    DataGetDD689 = True
    GoTo CleanUp
ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "ReportOps.DataGetDD689", App.Title, vbExclamation)
CleanUp:
    Set oForms = Nothing
End Function

Private Function DataGetSF522() As Boolean

On Error GoTo ErrHandler
    DataGetSF522 = False

    Dim oForms As CHCSII_Forms.Forms
    Set oForms = New CHCSII_Forms.Forms
    Set oForms.Comm = cwobjCWComm

    'oForms.printDD2766 PrintPreview, moReportItem.CriteriaSel1.Item(1).UnitNumber
    If Not oForms.printSF522(ReturnReportObject, moReportItem.CriteriaSel1.Item(1).UnitNumber) Then GoTo CleanUp
    Set moReport = oForms.ReportObject

    DataGetSF522 = True
    GoTo CleanUp
ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "ReportOps.DataGetSF522", App.Title, vbExclamation)
CleanUp:
    Set oForms = Nothing
    
End Function

Private Function DataGetDD741() As Boolean
On Error GoTo ErrHandler
    DataGetDD741 = False

    Dim oForms As CHCSII_Forms.Forms
    Set oForms = New CHCSII_Forms.Forms
    Set oForms.Comm = cwobjCWComm

    'oForms.printDD2766 PrintPreview, moReportItem.CriteriaSel1.Item(1).UnitNumber
    If Not oForms.printDD741(ReturnReportObject, moReportItem.CriteriaSel1.Item(1).UnitNumber) Then GoTo CleanUp
    Set moReport = oForms.ReportObject

    DataGetDD741 = True
    GoTo CleanUp
ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "ReportOps.DataGetDD741", App.Title, vbExclamation)
CleanUp:
    Set oForms = Nothing
End Function

Private Function DataGetDD771() As Boolean

On Error GoTo ErrHandler
    DataGetDD771 = False

    Dim oForms As CHCSII_Forms.Forms
    Set oForms = New CHCSII_Forms.Forms
    Set oForms.Comm = cwobjCWComm
    
    Dim x As New Collection

    x.Add "009999-2001-0000021"
    'oForms.printDD2766 PrintPreview, moReportItem.CriteriaSel1.Item(1).UnitNumber
    'If Not oForms.printDD771(ReturnReportObject, moReportItem.CriteriaSel1.Item(1).UnitNumber) Then GoTo CleanUp
    If Not oForms.printDD771(PrintPreview, x) Then GoTo CleanUp
    If Not oForms.printDD771Double(PrintPreview, x) Then GoTo CleanUp
    'Set moReport = oForms.ReportObject
    
    DataGetDD771 = True
    GoTo CleanUp
ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "ReportOps.DataGetDD771", App.Title, vbExclamation)
CleanUp:
    Set oForms = Nothing
End Function

Private Function DataGetPAPSmearHaveNot(HeaderColl As Collection) As Boolean

If Not DebugMode Then On Error GoTo ErrHandler
DataGetPAPSmearHaveNot = False

Dim sStmt As String
Dim rs As ADODB.Recordset
Dim sWhereCriteria As String
Dim sFromCriteria As String

HeaderColl.Add CrystalQuote(moReportItem.CriteriaSel1.Item(1).Name), "Header1"

If Not GetMTForPCM_SQL(sFromCriteria, sWhereCriteria) Then GoTo CleanUp

    sStmt = "select distinct id.unit_number " & _
            "from mmi_id id " & sFromCriteria & _
            "Where "
            
    '-- Check for PAP SMEAR Result
    sStmt = sStmt & _
            "(Not exists " & _
                "(select e.event_id " & _
                " from event e " & _
                " where " & _
                        GetOracleEndDate(moReportItem.DateTo) & " >= e.EVENT_START_GMTIME and " & _
                        "e.event_subtype_ncid = 90269 and e.event_type_ncid = 485 and " & _
                        "e.unit_number = id.unit_number) and "
    '-- Check if patient is a Female >= 21 or Female >= 64
    sStmt = sStmt & _
            "(((((" & GetOracleEndDate(moReportItem.DateTo) & " - id.birth_date)/365) >= 21.0) and " & _
            "(((" & GetOracleEndDate(moReportItem.DateTo) & " - id.birth_date)/365) <= 64.0) and " & _
            "id.sex = 'F') "
            
    '-- Check if patient has records at specified MTF or is patieint of specified PCM
    sStmt = sStmt & _
            " and " & sWhereCriteria & "))"
            
    '-- Check if patient belongs to the PCM
    'sStmt = sStmt & _
            "id.unit_number = mi.unit_number and " & _
            "mi.id_value = '" & moReportItem.CriteriaSel1.Item(1).NCID & "' and " & _
            "mi.ID_TYPE_NCID = 14501724"

    goDBSqlOps.Execute sStmt
    If goDBSqlOps.BOF And goDBSqlOps.EOF Then
        DataGetPAPSmearHaveNot = True
        GoTo CleanUp
    End If

    If Not moTableOps.OpenTable("PapSmearHaveNotReport", rs) Then GoTo CleanUp

    Do While Not goDBSqlOps.EOF
        rs.AddNew
        rs("PatientId") = goDBSqlOps.Value("unit_number")
        rs.Update
        goDBSqlOps.MoveNext
    Loop

    If Not moTableOps.CloseTable(rs) Then GoTo CleanUp
 
    If Not moTableOps.GetPatientTableData("PapSmearHaveNotReport", "PatientId") Then GoTo CleanUp
 
DataGetPAPSmearHaveNot = True
GoTo CleanUp
ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "ReportOps.DataGetPAPSmearHaveNot", App.Title, vbExclamation)
CleanUp:
    Set rs = Nothing
    Exit Function
End Function





Private Function DataGetHighCholesterolFollowUpRepeat(HeaderColl As Collection) As Boolean
If Not DebugMode Then On Error GoTo ErrHandler
DataGetHighCholesterolFollowUpRepeat = False

Dim LNumerator As Long
Dim LDenominator As Long
Dim sStmt As String
Dim sFromSQL As String
Dim sWhereSQL As String

'94129 - Lipid Panel Complete, Serum
'11158 - Plasma Total Cholesterol

'1036 Higher Than Critical
'1037 Higher Than Normal
'1039 Lower Than Critical
'1040 Lower Than Normal

    If Not GetMTForPCM_SQL(sFromSQL, sWhereSQL) Then GoTo CleanUp

'-- Get Denominator
'---------------------------------------------------------------------------
    sStmt = "select count(*) Denominator " & _
            "from mmi_id id " & sFromSQL & _
            "Where "
    '-- Check for Abnormal Cholesterol Lab Test
    sStmt = sStmt & _
            "exists " & _
                "(select 'x' " & _
                " from observation o, lab_test lt, event e " & _
                " where " & _
                "   o.abnormal_flag_ncid in (1036,1037,1039,1040) and " & _
                "   o.event_id = lt.event_id and " & _
                "   lt.test_ncid = 11158 and " & _
                "   lt.event_id = e.event_id and " & _
                GetOracleEndDate(moReportItem.DateTo) & " >= e.EVENT_START_GMTIME and " & _
                GetOracleEndDate(moReportItem.DateFrom) & " <= e.EVENT_START_GMTIME and " & _
                "   e.event_type_ncid = 485 and " & _
                "   e.unit_number = id.unit_number) and "

    '-- Check if patient was entered before the 'TO' date
    sStmt = sStmt & _
            "(exists " & _
                "(select ma.unit_number " & _
                " from mmi_audit ma " & _
                " where " & _
                        GetOracleEndDate(moReportItem.DateTo) & " >= TX_TIME and " & _
                        "ma.unit_number = id.unit_number) and "

    '-- Check if patient was >= 17 years before 'TO' date
    sStmt = sStmt & _
            "((((" & GetOracleEndDate(moReportItem.DateTo) & " - id.birth_date)/365) >= 17.0) "
    
'    '-- Check if patient has records at specified MTF or is patieint of specified PCM
    sStmt = sStmt & _
            " and " & sWhereSQL & "))"

    goDBSqlOps.Execute sStmt
    
    LDenominator = goDBSqlOps.Value("Denominator")
    
'-- Get Numerator
'---------------------------------------------------------------------------
    sStmt = "select count(*) Numerator " & _
            "from mmi_id id " & sFromSQL & _
            "Where "
    '-- Check for Abnormal Cholesterol Lab Test and then HDL/LDL test within 2 months of Abnormal
    sStmt = sStmt & _
            "exists " & _
                "(select 'x' " & _
                " from observation o, lab_test lt, event e " & _
                " where " & _
                    "exists " & _
                        "(select 'x' " & _
                        " from lab_test lt2, event e2 " & _
                        " where " & _
                        "(e2.EVENT_END_GMTIME - e.EVENT_START_GMTIME) <= 62.0 and " & _
                        "   lt2.test_ncid = 94129 and " & _
                        "   lt2.event_id = e2.event_id and " & _
                        "   e2.event_type_ncid = 485 and " & _
                        "   e2.unit_number = e.unit_number) and " & _
                "   o.abnormal_flag_ncid in (1036,1037,1039,1040) and " & _
                "   o.event_id = lt.event_id and " & _
                "   lt.test_ncid = 11158 and " & _
                "   lt.event_id = e.event_id and " & _
                GetOracleEndDate(moReportItem.DateTo) & " >= e.EVENT_START_GMTIME and " & _
                GetOracleEndDate(moReportItem.DateFrom) & " <= e.EVENT_START_GMTIME and " & _
                "   e.event_type_ncid = 485 and " & _
                "   e.unit_number = id.unit_number) and "
    '-- Check if patient was entered before the 'TO' date
    sStmt = sStmt & _
            "(exists " & _
                "(select ma.unit_number " & _
                " from mmi_audit ma " & _
                " where " & _
                        GetOracleEndDate(moReportItem.DateTo) & " >= TX_TIME and " & _
                        "ma.unit_number = id.unit_number) and "
    '-- Check if patient was >= 17 years before 'TO' date
    sStmt = sStmt & _
            "((((" & GetOracleEndDate(moReportItem.DateTo) & " - id.birth_date)/365) >= 17.0) "
'    '-- Check if patient has records at specified MTF or is patieint of specified PCM
    sStmt = sStmt & _
            " and " & sWhereSQL & "))"
    
    goDBSqlOps.Execute sStmt
    
    LNumerator = goDBSqlOps.Value("Numerator")
    
'-- Set Report Field (Formula Field) Information
'---------------------------------------------------------------------------

 HeaderColl.Add CrystalQuote(CStr(LNumerator)), "Numerator"
 HeaderColl.Add CrystalQuote(CStr(LDenominator)), "Denominator"
 HeaderColl.Add CrystalQuote(moReportItem.CriteriaSel1.Item(1).Name), "Facility"

DataGetHighCholesterolFollowUpRepeat = True
GoTo CleanUp
ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "ReportOps. DataGetHighCholesterolFollowUpRepeat", App.Title, vbExclamation)
CleanUp:
    Exit Function
End Function

Private Function DataGetImmunizationsActiveDutyA2Dose(HeaderColl As Collection) As Boolean

If Not DebugMode Then On Error GoTo ErrHandler
DataGetImmunizationsActiveDutyA2Dose = False

Dim LNumerator As Long
Dim LDenominator As Long
Dim sStmt As String
Const HEPA = 162038 '-- Measles, Mumps, Rubella
Dim sWhereCriteria As String
Dim sFromCriteria As String


If Not GetMTForPCM_SQL(sFromCriteria, sWhereCriteria) Then GoTo CleanUp


'-- Get Denominator
'---------------------------------------------------------------------------
sStmt = "select count(*) Denominator " & _
        "from mmi_genericdata Id " & sFromCriteria & _
        "Where "
'    '-- Check for valid facility
'    sStmt = sStmt & _
'            "exists " & _
'                "(select ms.unit_number " & _
'                " from mmi_site ms " & _
'                " where " & _
'                        moReportItem.CriteriaSel1.Item(1).NCID & " = facility_cid and " & _
'                        "ms.unit_number = id.unit_number) and "
    '-- Check if patient entered in system before 'TO' date
    sStmt = sStmt & _
            "(exists " & _
                "(select ma.unit_number " & _
                " from mmi_audit ma " & _
                " where " & _
                        GetOracleEndDate(moReportItem.DateTo) & " >= TX_TIME and " & _
                        "ma.unit_number = id.unit_number) and "
    '-- Check if patient is Active Duty
    '-- Not sure how to check for active duty so checking for Unit Identification Code
    sStmt = sStmt & _
                "(Id.value_ncid in " & _
                "(119033,119034,119035,119068,119069,119070, " & _
                 "119122,119123,119124,119148,119149,119150) and " & _
                "(Id.type_ncid = 14501721  "

    '-- Check if patient has records at specified MTF or is patieint of specified PCM
    sStmt = sStmt & _
            " and (" & sWhereCriteria & "))))"

goDBSqlOps.Execute sStmt

LDenominator = goDBSqlOps.Value("Denominator")

'-- Get Numerator
'---------------------------------------------------------------------------
    sStmt = "select count(*) Numerator " & _
        "from mmi_genericdata id " & sFromCriteria & _
        "Where "
        
    '-- Check if patient has had second HEPA vaccine
    sStmt = sStmt & _
                "(exists " & _
                    "(select e.unit_number " & _
                    "from observation_value o, clinical_observation c, " & _
                        "event e " & _
                    "Where " & _
                            "(o.observation_seq = 2 and " & _
                            "(o.value_ncid = " & CStr(HEPA) & " and " & _
                            "(o.observation_seq = c.observation_seq and " & _
                            "(o.event_id = c.event_id and " & _
                            "(c.observation_ncid = 154784 and " & _
                            "(c.event_id = e.event_id and " & _
                            "(e.event_subtype_ncid = 154789 and " & _
                            "(e.event_type_ncid = 90782 and " & _
                            "(e.unit_number = id.unit_number)))))))))) and "

'    '-- Check for valid facility
'    sStmt = sStmt & _
'            "(exists " & _
'                "(select ms.unit_number " & _
'                " from mmi_site ms " & _
'                " where " & _
'                        moReportItem.CriteriaSel1.Item(1).NCID & " = facility_cid and " & _
'                        "ms.unit_number = id.unit_number) and "
    '-- Check if patient entered in system before 'TO' date
    sStmt = sStmt & _
            "(exists " & _
                "(select ma.unit_number " & _
                " from mmi_audit ma " & _
                " where " & _
                        GetOracleEndDate(moReportItem.DateTo) & " >= TX_TIME and " & _
                        "ma.unit_number = id.unit_number) and "
    '-- Check if patient is Active Duty
    '-- Not sure how to check for active duty so checking for Unit Identification Code
    sStmt = sStmt & _
                "(Id.value_ncid in " & _
                "(119033,119034,119035,119068,119069,119070, " & _
                 "119122,119123,119124,119148,119149,119150) and " & _
                "(Id.type_ncid = 14501721  "

    '-- Check if patient has records at specified MTF or is patieint of specified PCM
    sStmt = sStmt & _
            " and (" & sWhereCriteria & ")))))"
    
goDBSqlOps.Execute sStmt

LNumerator = goDBSqlOps.Value("Numerator")


'-- Set Report Field (Formula Field) Information
'---------------------------------------------------------------------------

 HeaderColl.Add CrystalQuote(CStr(LNumerator)), "Numerator"
 HeaderColl.Add CrystalQuote(CStr(LDenominator)), "Denominator"
 HeaderColl.Add CrystalQuote(moReportItem.CriteriaSel1.Item(1).Name), "Facility"

DataGetImmunizationsActiveDutyA2Dose = True
GoTo CleanUp
ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "ReportOps.DataGetImmunizationsActiveDutyA2Dose", App.Title, vbExclamation)
CleanUp:
    Exit Function
End Function
Private Function DataGetImmunizationsActiveDutyA1Dose(HeaderColl As Collection) As Boolean

If Not DebugMode Then On Error GoTo ErrHandler
DataGetImmunizationsActiveDutyA1Dose = False

Dim LNumerator As Long
Dim LDenominator As Long
Dim sStmt As String
Const HEPA = 162038 '-- Measles, Mumps, Rubella
Dim sWhereCriteria As String
Dim sFromCriteria As String


If Not GetMTForPCM_SQL(sFromCriteria, sWhereCriteria) Then GoTo CleanUp

'-- Get Denominator
'---------------------------------------------------------------------------
sStmt = "select count(*) Denominator " & vbNewLine & _
        "from mmi_genericdata Id " & sFromCriteria & _
        "Where "
    '-- Check if patient entered in system before 'TO' date
    sStmt = sStmt & _
            "(exists " & vbNewLine & _
                "(select ma.unit_number " & vbNewLine & _
                " from mmi_audit ma " & vbNewLine & _
                " where " & vbNewLine & _
                        GetOracleEndDate(moReportItem.DateTo) & " >= TX_TIME and " & vbNewLine & _
                        "ma.unit_number = id.unit_number) and " & vbNewLine
'    '-- Check for valid facility
'    sStmt = sStmt & _
'            "exists " & _
'                "(select ms.unit_number " & _
'                " from mmi_site ms " & _
'                " where " & _
'                        moReportItem.CriteriaSel1.Item(1).NCID & " = facility_cid and " & _
'                        "ms.unit_number = id.unit_number) and "

    '-- Check if patient is Active Duty
    sStmt = sStmt & _
                "(Id.value_ncid in " & vbNewLine & _
                "(119033,119034,119035,119068,119069,119070, " & vbNewLine & _
                 "119122,119123,119124,119148,119149,119150) and " & vbNewLine & _
                "(Id.type_ncid = 14501721 " & vbNewLine

    '-- Check if patient has records at specified MTF or is patieint of specified PCM
    sStmt = sStmt & _
            " and (" & sWhereCriteria & "))))"
    
goDBSqlOps.Execute sStmt

LDenominator = goDBSqlOps.Value("Denominator")

'-- Get Numerator
'---------------------------------------------------------------------------
    sStmt = "select count(*) Numerator " & vbNewLine & _
        "from mmi_genericdata id " & sFromCriteria & _
        "Where " & vbNewLine
        
    '-- Check if patient has had first HEPA vaccine
    sStmt = sStmt & vbNewLine & _
                "(exists " & vbNewLine & _
                    "(select e.unit_number " & vbNewLine & _
                    "from event e, clinical_observation c, " & vbNewLine & _
                        "observation_value o " & vbNewLine & _
                    "Where " & vbNewLine & _
                            "(o.observation_seq = 1 and " & vbNewLine & _
                            "(o.value_ncid = " & CStr(HEPA) & " and " & vbNewLine & _
                            "(o.observation_seq = c.observation_seq and " & vbNewLine & _
                            "(o.event_id = c.event_id and " & vbNewLine & _
                            "(c.observation_ncid = 154784 and " & vbNewLine & _
                            "(c.event_id = e.event_id and " & vbNewLine & _
                            "(e.event_subtype_ncid = 154789 and " & vbNewLine & _
                            "(e.event_type_ncid = 90782 and " & vbNewLine & _
                            "(e.unit_number = id.unit_number)))))))))) and " & vbNewLine

    '-- Check if patient entered in system before 'TO' date
    sStmt = sStmt & _
            "(exists " & vbNewLine & _
                "(select ma.unit_number " & vbNewLine & _
                " from mmi_audit ma " & vbNewLine & _
                " where " & vbNewLine & _
                        GetOracleEndDate(moReportItem.DateTo) & " >= TX_TIME and " & vbNewLine & _
                        "ma.unit_number = id.unit_number) and " & vbNewLine
'    '-- Check for facility
'    sStmt = sStmt & _
'            "exists " & _
'                "(select ms.unit_number " & _
'                " from mmi_site ms " & _
'                " where " & _
'                        moReportItem.CriteriaSel1.Item(1).NCID & " = facility_cid and " & _
'                        "ms.unit_number = id.unit_number) and "
    '-- Check if patient is Active Duty
    '-- Not sure how to check for active duty so checking for Unit Identification Code
    sStmt = sStmt & _
                "(Id.value_ncid in " & vbNewLine & _
                "(119033,119034,119035,119068,119069,119070, " & vbNewLine & _
                 "119122,119123,119124,119148,119149,119150) and " & vbNewLine & _
                "(Id.type_ncid = 14501721 " & vbNewLine

    '-- Check if patient has records at specified MTF or is patieint of specified PCM
    sStmt = sStmt & _
            " and (" & sWhereCriteria & ")))))"

goDBSqlOps.Execute sStmt

LNumerator = goDBSqlOps.Value("Numerator")

'-- Set Report Field (Formula Field) Information
'---------------------------------------------------------------------------

 HeaderColl.Add CrystalQuote(CStr(LNumerator)), "Numerator"
 HeaderColl.Add CrystalQuote(CStr(LDenominator)), "Denominator"
 HeaderColl.Add CrystalQuote(moReportItem.CriteriaSel1.Item(1).Name), "Facility"

DataGetImmunizationsActiveDutyA1Dose = True
GoTo CleanUp

ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "ReportOps.DataGetImmunizationsActiveDutyA1Dose", App.Title, vbExclamation)
CleanUp:
    Exit Function
End Function
Private Function DataGetHebBHibImmunizations2YearOlds(HeaderColl As Collection) As Boolean

If Not DebugMode Then On Error GoTo ErrHandler
DataGetHebBHibImmunizations2YearOlds = False

Dim LNumerator As Long
Dim LDenominator As Long
Dim sStmt As String
Const HepatitsB = 162039
Const HaemophilusInfuenzaeB = 162041
Dim sWhereCriteria As String
Dim sFromCriteria As String


If Not GetMTForPCM_SQL(sFromCriteria, sWhereCriteria) Then GoTo CleanUp

'-- Get Denominator
'---------------------------------------------------------------------------
sStmt = "select count(*) Denominator " & _
        "from mmi_id Id " & sFromCriteria & _
        "Where "

    '-- Check if patient entered in system before 'TO' date
    sStmt = sStmt & _
            "(exists " & _
                "(select ma.unit_number " & _
                " from mmi_audit ma " & _
                " where " & _
                        GetOracleEndDate(moReportItem.DateTo) & " >= TX_TIME and " & _
                        "ma.unit_number = id.unit_number) and "
'    '-- Check for facility
'    sStmt = sStmt & _
'            "exists " & _
'                "(select ms.unit_number " & _
'                " from mmi_site ms " & _
'                " where " & _
'                        moReportItem.CriteriaSel1.Item(1).NCID & " = facility_cid and " & _
'                        "ms.unit_number = id.unit_number) and "
    '-- Check if patient is >= 2 years of age and patient < 3 years of age
    sStmt = sStmt & _
            "((((" & GetOracleEndDate(moReportItem.DateTo) & " - id.birth_date)/365) >= 2.0) and " & _
            "((((" & GetOracleEndDate(moReportItem.DateTo) & " - id.birth_date)/365) < 3.0) "
            
    '-- Check if patient has records at specified MTF or is patieint of specified PCM
    sStmt = sStmt & _
            " and (" & sWhereCriteria & "))))"

goDBSqlOps.Execute sStmt

LDenominator = goDBSqlOps.Value("Denominator")

'-- Get Numerator
'---------------------------------------------------------------------------
sStmt = "select count(*) Numerator " & _
        "from mmi_id id " & sFromCriteria & _
        "Where "
    '-- Check if patient has had third Haemophilus Infuenzae B vaccine before age of 2
    sStmt = sStmt & _
                "(exists " & _
                    "(select e.unit_number " & _
                    "from observation_value o, clinical_observation c, " & _
                        "event e " & _
                    "Where " & _
                            "o.observation_seq = 3 and " & _
                            "o.value_ncid = 162041 and " & _
                            "o.observation_seq = c.observation_seq and " & _
                            "o.event_id = c.event_id and " & _
                            "c.observation_ncid = 154784 and " & _
                            "c.event_id = e.event_id and " & _
                            "e.event_start_gmtime <= (id.birth_date + (365*2)) and " & _
                            "e.event_subtype_ncid = 154789 and " & _
                            "e.event_type_ncid = 90782 and " & _
                            "e.unit_number = id.unit_number) and "
    '-- Check if patient has had third Hepatits B vaccine before age of 2
    sStmt = sStmt & _
                "(exists " & _
                    "(select e.unit_number " & _
                    "from observation_value o, clinical_observation c, " & _
                        "event e " & _
                    "Where " & _
                            "o.observation_seq = 3 and " & _
                            "o.value_ncid = 162039 and " & _
                            "o.observation_seq = c.observation_seq and " & _
                            "o.event_id = c.event_id and " & _
                            "c.observation_ncid = 154784 and " & _
                            "c.event_id = e.event_id and " & _
                            "e.event_start_gmtime <= (id.birth_date + (365*2)) and " & _
                            "e.event_subtype_ncid = 154789 and " & _
                            "e.event_type_ncid = 90782 and " & _
                            "e.unit_number = id.unit_number) and "
    '-- Check if patient entered in system before 'TO' date
    sStmt = sStmt & _
            "(exists " & _
                "(select ma.unit_number " & _
                " from mmi_audit ma " & _
                " where " & _
                        GetOracleEndDate(moReportItem.DateTo) & " >= TX_TIME and " & _
                        "ma.unit_number = id.unit_number) and "
'    '-- Check for facility
'    sStmt = sStmt & _
'            "exists " & _
'                "(select ms.unit_number " & _
'                " from mmi_site ms " & _
'                " where " & _
'                        moReportItem.CriteriaSel1.Item(1).NCID & " = facility_cid and " & _
'                        "ms.unit_number = id.unit_number) and "
    '-- Check if patient is >= 2 years of age and patient < 3 years of age
    sStmt = sStmt & _
            "((((" & GetOracleEndDate(moReportItem.DateTo) & " - id.birth_date)/365) >= 2.0) and " & _
            "((((" & GetOracleEndDate(moReportItem.DateTo) & " - id.birth_date)/365) < 3.0) "

            
    '-- Check if patient has records at specified MTF or is patieint of specified PCM
    sStmt = sStmt & _
            " and (" & sWhereCriteria & "))))))"

goDBSqlOps.Execute sStmt

LNumerator = goDBSqlOps.Value("Numerator")

'-- Set Report Field (Formula Field) Information
'---------------------------------------------------------------------------

 HeaderColl.Add CrystalQuote(CStr(LNumerator)), "Numerator"
 HeaderColl.Add CrystalQuote(CStr(LDenominator)), "Denominator"
 HeaderColl.Add CrystalQuote(moReportItem.CriteriaSel1.Item(1).Name), "Facility"

DataGetHebBHibImmunizations2YearOlds = True
GoTo CleanUp
ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "ReportOps.DataGetHebBHibImmunizations2YearOlds", App.Title, vbExclamation)
CleanUp:
    Exit Function
End Function


Private Function DataGetDTPImmunizations2yearOlds(HeaderColl As Collection) As Boolean

If Not DebugMode Then On Error GoTo ErrHandler
DataGetDTPImmunizations2yearOlds = False

Dim LNumerator As Long
Dim LDenominator As Long

Dim sStmt As String
Dim sWhereCriteria As String
Dim sFromCriteria As String


If Not GetMTForPCM_SQL(sFromCriteria, sWhereCriteria) Then GoTo CleanUp

'-- Get Denominator
'---------------------------------------------------------------------------
sStmt = "select count(*) Denominator " & _
        "from mmi_id Id " & sFromCriteria & _
        "Where "
    
    '-- Check if patient reached age 2 between the given date range
    sStmt = sStmt & _
            "(((" & GetOracleEndDate(moReportItem.DateFrom) & " - id.birth_date) < 1095) and " & _
            "((" & GetOracleEndDate(moReportItem.DateTo) & " - id.birth_date) > 1095)) "
    '-- Check if patient has records at specified MTF or is patieint of specified PCM
    sStmt = sStmt & _
            " and (" & sWhereCriteria & ")"

goDBSqlOps.Execute sStmt


LDenominator = goDBSqlOps.Value("Denominator")

'-- Get Numerator
'---------------------------------------------------------------------------
sStmt = "select /*+ ORDERED */ count(*) Numerator " & _
        "from mmi_generic_id mg, mmi_id id, mmi_generic_id ssn, mmi_generic_id fmp " & _
        "Where "

    '-- Check if patient has been given a MMR immunization before the age of 2
    sStmt = sStmt & _
                "(exists " & _
                  "(select /*+ ORDERED */ 'x'" & vbNewLine & _
                   "from imm_people_vaccine ipv, tblvaccines tv" & vbNewLine & _
                   "where" & vbNewLine & _
                   "ipv.vacc_series = 1 and ipv.vacc_date <= (id.birth_date + (365*2)) and " & _
                   "tv.genericname = 'MMR' and" & vbNewLine & _
                   "tv.vaccineid = ipv.vaccineid and" & vbNewLine & _
                   "ipv.spon_ssan_fmp = ssn.id_value || fmp.id_value) and" & vbNewLine
                
    '-- Check if patient has been given a OPV immunization before the age of 2
    sStmt = sStmt & _
                "(exists " & _
                  "(select /*+ ORDERED */ 'x'" & vbNewLine & _
                   "from imm_people_vaccine ipv, tblvaccines tv" & vbNewLine & _
                   "where" & vbNewLine & _
                   "ipv.vacc_series = 3 and ipv.vacc_date <= (id.birth_date + (365*2)) and " & _
                   "tv.genericname = 'Polio' and" & vbNewLine & _
                   "tv.vaccineid = ipv.vaccineid and" & vbNewLine & _
                   "ipv.spon_ssan_fmp = ssn.id_value || fmp.id_value) and" & vbNewLine
    
    '-- Check if patient has been given a DTP immunization before the Age of 2
    sStmt = sStmt & _
                "(exists " & _
                  "(select /*+ ORDERED */ 'x'" & vbNewLine & _
                   "from imm_people_vaccine ipv, tblvaccines tv" & vbNewLine & _
                   "where" & vbNewLine & _
                   "ipv.vacc_series = 4 and ipv.vacc_date <= (id.birth_date + (365*2)) and " & _
                   "tv.genericname = 'DTaP' and" & vbNewLine & _
                   "tv.vaccineid = ipv.vaccineid and" & vbNewLine & _
                   "ipv.spon_ssan_fmp = ssn.id_value || fmp.id_value) and" & vbNewLine
    


    '-- Get Sponsor SSN and Patient FMP
    sStmt = sStmt & _
            "fmp.id_type_ncid = 204708 and" & vbNewLine & _
            "fmp.unit_number = id.unit_number and" & vbNewLine & _
            "ssn.id_type_ncid = 186334 and" & vbNewLine & _
            "ssn.unit_number = id.unit_number and" & vbNewLine
    
    '-- Check if patient is >= 2 years of age and patient < 3 years of age
    sStmt = sStmt & _
            "(((" & GetOracleEndDate(moReportItem.DateFrom) & " - id.birth_date) < 1095) and " & _
            "((" & GetOracleEndDate(moReportItem.DateTo) & " - id.birth_date) > 1095)) "

    '-- Check if patient has records at specified MTF or is patieint of specified PCM
    sStmt = sStmt & _
            " and (" & sWhereCriteria & "))))"

goDBSqlOps.Execute sStmt

LNumerator = goDBSqlOps.Value("Numerator")

'-- Set Report Field (Formula Field) Information
'---------------------------------------------------------------------------

HeaderColl.Add CrystalQuote(CStr(LNumerator)), "Numerator"
HeaderColl.Add CrystalQuote(CStr(LDenominator)), "Denominator"
HeaderColl.Add CrystalQuote(moReportItem.CriteriaSel1.Item(1).Name), "Facility"

DataGetDTPImmunizations2yearOlds = True
GoTo CleanUp
ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "ReportOps.DataGetDTPImmunizations2yearOlds", App.Title, vbExclamation)
CleanUp:
    Exit Function
End Function

Private Function DataGetVaricellaImmunizations2yearOlds(HeaderColl As Collection) As Boolean

If Not DebugMode Then On Error GoTo ErrHandler
DataGetVaricellaImmunizations2yearOlds = False

Dim LNumerator As Long
Dim LDenominator As Long
Dim sStmt As String
Const Varicella = 162043
Dim sWhereCriteria As String
Dim sFromCriteria As String


If Not GetMTForPCM_SQL(sFromCriteria, sWhereCriteria) Then GoTo CleanUp


'-- Get Denominator
'---------------------------------------------------------------------------
sStmt = "select count(*) Denominator " & _
        "from mmi_id Id  " & sFromCriteria & _
        "Where "

    '-- Check if patient entered in system before 'TO' date
    sStmt = sStmt & _
            "(exists " & _
                "(select ma.unit_number " & _
                " from mmi_audit ma " & _
                " where " & _
                        GetOracleEndDate(moReportItem.DateTo) & " >= TX_TIME and " & _
                        "ma.unit_number = id.unit_number) and "
'    '-- Check for facility
'    sStmt = sStmt & _
'            "exists " & _
'                "(select ms.unit_number " & _
'                " from mmi_site ms " & _
'                " where " & _
'                        moReportItem.CriteriaSel1.Item(1).NCID & " = facility_cid and " & _
'                        "ms.unit_number = id.unit_number) and "
    '-- Check if patient is >= 2 years of age and patient < 3 years of age
    sStmt = sStmt & _
            "((((" & GetOracleEndDate(moReportItem.DateTo) & " - id.birth_date)/365) >= 2.0) and " & _
            "((((" & GetOracleEndDate(moReportItem.DateTo) & " - id.birth_date)/365) < 3.0) "

    '-- Check if patient has records at specified MTF or is patieint of specified PCM
    sStmt = sStmt & _
            " and (" & sWhereCriteria & "))))"

goDBSqlOps.Execute sStmt

LDenominator = goDBSqlOps.Value("Denominator")

'-- Get Numerator
'---------------------------------------------------------------------------
sStmt = "select count(*) Numerator " & _
        "from mmi_id id " & sFromCriteria & _
        "Where "
    '-- Check if patient has been given a Varicella immunization before the age of 2
    sStmt = sStmt & _
                "(exists " & _
                    "(select e.unit_number " & _
                    "from observation_value o, clinical_observation c, event e " & _
                    "Where " & _
                            "(o.observation_seq = 1 and " & _
                            "(o.value_ncid = " & CStr(Varicella) & " and " & _
                            "(o.observation_seq = c.observation_seq and " & _
                            "(o.event_id = c.event_id and " & _
                            "(c.observation_ncid = 154784 and " & _
                            "(c.event_id = e.event_id and " & _
                            "(e.event_start_gmtime <= (id.birth_date + (365*2)) and " & _
                            "(e.event_subtype_ncid = 154789 and " & _
                            "(e.event_type_ncid = 90782 and " & _
                            "(e.unit_number = id.unit_number))))))))))) and "
    '-- Check if patient entered in system before 'TO' date
    sStmt = sStmt & _
            "(exists " & _
                "(select ma.unit_number " & _
                " from mmi_audit ma " & _
                " where " & _
                        GetOracleEndDate(moReportItem.DateTo) & " >= TX_TIME and " & _
                        "ma.unit_number = id.unit_number) and "
'    '-- Check for facility
'    sStmt = sStmt & _
'            "(exists " & _
'                "(select ms.unit_number " & _
'                " from mmi_site ms " & _
'                " where " & _
'                        moReportItem.CriteriaSel1.Item(1).NCID & " = facility_cid and " & _
'                        "ms.unit_number = id.unit_number) and "
    '-- Check if patient is >= 2 years of age and patient < 3 years of age
    sStmt = sStmt & _
            "((((" & GetOracleEndDate(moReportItem.DateTo) & " - id.birth_date)/365) >= 2.0) and " & _
            "((((" & GetOracleEndDate(moReportItem.DateTo) & " - id.birth_date)/365) < 3.0) "

    '-- Check if patient has records at specified MTF or is patieint of specified PCM
    sStmt = sStmt & _
            " and (" & sWhereCriteria & ")))))"

goDBSqlOps.Execute sStmt

LNumerator = goDBSqlOps.Value("Numerator")


'-- Set Report Field (Formula Field) Information
'---------------------------------------------------------------------------

HeaderColl.Add CrystalQuote(CStr(LNumerator)), "Numerator"
HeaderColl.Add CrystalQuote(CStr(LDenominator)), "Denominator"
HeaderColl.Add CrystalQuote(moReportItem.CriteriaSel1.Item(1).Name), "Facility"

DataGetVaricellaImmunizations2yearOlds = True
GoTo CleanUp

ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "ReportOps.DataGetVaricellaImmunizations2yearOlds", App.Title, vbExclamation)
CleanUp:
    Exit Function
End Function


Private Function DataGetImmunizations13yearOlds(HeaderColl As Collection) As Boolean

If Not DebugMode Then On Error GoTo ErrHandler
DataGetImmunizations13yearOlds = False

Dim LNumerator As Long
Dim LDenominator As Long
Dim sStmt As String
Const Varicella = 162044 '--
Const MMR = 162036 '-- Measles, Mumps, Rubella
Const HBV = 162039 '-- Don't know this value yet
Const TD = 162035 '-- Right now assuming TD is the Same as these DtaP, DTP, Td, DT, TT
Dim sWhereCriteria As String
Dim sFromCriteria As String


If Not GetMTForPCM_SQL(sFromCriteria, sWhereCriteria) Then GoTo CleanUp

'-- Get Denominator
'---------------------------------------------------------------------------
sStmt = "select count(*) Denominator " & _
        "from mmi_id Id " & sFromCriteria & _
        "Where "

    '-- Check if patient entered in system before 'TO' date
    sStmt = sStmt & _
            "(exists " & _
                "(select ma.unit_number " & _
                " from mmi_audit ma " & _
                " where " & _
                        GetOracleEndDate(moReportItem.DateTo) & " >= TX_TIME and " & _
                        "ma.unit_number = id.unit_number) and "
'    '-- Check for facility
'    sStmt = sStmt & _
'            "exists " & _
'                "(select ms.unit_number " & _
'                " from mmi_site ms " & _
'                " where " & _
'                        moReportItem.CriteriaSel1.Item(1).NCID & " = facility_cid and " & _
'                        "ms.unit_number = id.unit_number) and "
    '-- Check if patient is >= 13 years of age and patient < 14 years of age
    sStmt = sStmt & _
            "((((" & GetOracleEndDate(moReportItem.DateFrom) & " - id.birth_date)/365) >= 13.0) and " & _
            "((((" & GetOracleEndDate(moReportItem.DateTo) & " - id.birth_date)/365) < 14.0) "

    '-- Check if patient has records at specified MTF or is patieint of specified PCM
    sStmt = sStmt & _
            " and (" & sWhereCriteria & "))))"

goDBSqlOps.Execute sStmt

LDenominator = goDBSqlOps.Value("Denominator")

'-- Get Numerator
'---------------------------------------------------------------------------
sStmt = "select count(*) Numerator " & _
        "from mmi_id id " & sFromCriteria & _
        "Where "
        
    '-- Check if patient has had second MMR vaccine before age of 14
    sStmt = sStmt & _
                "(exists " & _
                    "(select e.unit_number " & _
                    "from observation_value o, clinical_observation c, " & _
                        "event e " & _
                    "Where " & _
                            "o.observation_seq = 2 and " & _
                            "o.value_ncid = " & CStr(MMR) & " and " & _
                            "o.observation_seq = c.observation_seq and " & _
                            "o.event_id = c.event_id and " & _
                            "c.observation_ncid = 154784 and " & _
                            "c.event_id = e.event_id and " & _
                            "e.event_start_gmtime <= (id.birth_date + (365*14)) and " & _
                            "e.event_subtype_ncid = 154789 and " & _
                            "e.event_type_ncid = 90782 and " & _
                            "e.unit_number = id.unit_number) and "
    '-- Check if patient has had second TD vaccine before age of 14
    sStmt = sStmt & _
                "(exists " & _
                    "(select e.unit_number " & _
                    "from observation_value o, clinical_observation c, " & _
                        "event e " & _
                    "Where " & _
                            "o.observation_seq = 1 and " & _
                            "o.value_ncid = " & CStr(TD) & " and " & _
                            "o.observation_seq = c.observation_seq and " & _
                            "o.event_id = c.event_id and " & _
                            "c.observation_ncid = 154784 and " & _
                            "c.event_id = e.event_id and " & _
                            "e.event_start_gmtime <= (id.birth_date + (365*14)) and " & _
                            "e.event_subtype_ncid = 154789 and " & _
                            "e.event_type_ncid = 90782 and " & _
                            "e.unit_number = id.unit_number) and "
    '-- Check if patient has had third Varicella vaccine before age of 14
    sStmt = sStmt & _
                "(exists " & _
                    "(select e.unit_number " & _
                    "from observation_value o, clinical_observation c, " & _
                        "event e " & _
                    "Where " & _
                            "o.observation_seq = 3 and " & _
                            "o.value_ncid = " & CStr(Varicella) & " and " & _
                            "o.observation_seq = c.observation_seq and " & _
                            "o.event_id = c.event_id and " & _
                            "c.observation_ncid = 154784 and " & _
                            "c.event_id = e.event_id and " & _
                            "e.event_start_gmtime <= (id.birth_date + (365*14)) and " & _
                            "e.event_subtype_ncid = 154789 and " & _
                            "e.event_type_ncid = 90782 and " & _
                            "e.unit_number = id.unit_number) and "
    '-- Check if patient has had third HBV vaccine before age of 14
    sStmt = sStmt & _
                "(exists " & _
                    "(select e.unit_number " & _
                    "from observation_value o, clinical_observation c, " & _
                        "event e " & _
                    "Where " & _
                            "o.observation_seq = 3 and " & _
                            "o.value_ncid = " & CStr(HBV) & " and " & _
                            "o.observation_seq = c.observation_seq and " & _
                            "o.event_id = c.event_id and " & _
                            "c.observation_ncid = 154784 and " & _
                            "c.event_id = e.event_id and " & _
                            "e.event_start_gmtime <= (id.birth_date + (365*14)) and " & _
                            "e.event_subtype_ncid = 154789 and " & _
                            "e.event_type_ncid = 90782 and " & _
                            "e.unit_number = id.unit_number) and "
                            
    '-- Check if patient entered in system before 'TO' date
    sStmt = sStmt & _
            "(exists " & _
                "(select ma.unit_number " & _
                " from mmi_audit ma " & _
                " where " & _
                        GetOracleEndDate(moReportItem.DateTo) & " >= TX_TIME and " & _
                        "ma.unit_number = id.unit_number) and "
'    '-- Check for facility
'    sStmt = sStmt & _
'            "exists " & _
'                "(select ms.unit_number " & _
'                " from mmi_site ms " & _
'                " where " & _
'                        moReportItem.CriteriaSel1.Item(1).NCID & " = facility_cid and " & _
'                        "ms.unit_number = id.unit_number) and "
    '-- Check if patient is >= 13 years of age and patient < 14 years of age
    sStmt = sStmt & _
            "((((" & GetOracleEndDate(moReportItem.DateFrom) & " - id.birth_date)/365) >= 13.0) and " & _
            "((((" & GetOracleEndDate(moReportItem.DateTo) & " - id.birth_date)/365) < 14.0) "
            
    '-- Check if patient has records at specified MTF or is patieint of specified PCM
    sStmt = sStmt & _
            " and (" & sWhereCriteria & "))))))))"

goDBSqlOps.Execute sStmt

LNumerator = goDBSqlOps.Value("Numerator")

'-- Set Report Field (Formula Field) Information
'---------------------------------------------------------------------------

 HeaderColl.Add CrystalQuote(CStr(LNumerator)), "Numerator"
 HeaderColl.Add CrystalQuote(CStr(LDenominator)), "Denominator"
 HeaderColl.Add CrystalQuote(moReportItem.CriteriaSel1.Item(1).Name), "Facility"

DataGetImmunizations13yearOlds = True
GoTo CleanUp
ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "ReportOps.DataGetImmunizations13yearOlds", App.Title, vbExclamation)
CleanUp:
    Exit Function
End Function


Private Function DataGetHighCholesterolHDL_LDLFollowUp(HeaderColl As Collection) As Boolean

If Not DebugMode Then On Error GoTo ErrHandler
DataGetHighCholesterolHDL_LDLFollowUp = False

Dim LNumerator As Long
Dim LDenominator As Long
Dim sStmt As String
Dim sFromSQL As String
Dim sWhereSQL As String

'94129 - Lipid Panel Complete, Serum
'11158 - Plasma Total Cholesterol

'1036 Higher Than Critical
'1037 Higher Than Normal
'1039 Lower Than Critical
'1040 Lower Than Normal

    If Not GetMTForPCM_SQL(sFromSQL, sWhereSQL) Then GoTo CleanUp

'-- Get Denominator
'---------------------------------------------------------------------------
    sStmt = "select count(*) Denominator " & _
            "from mmi_id id " & sFromSQL & _
            "Where "
    '-- Check for Abnormal Cholesterol Lab Test
    sStmt = sStmt & _
            "exists " & _
                "(select 'x' " & _
                " from observation o, lab_test lt, event e " & _
                " where " & _
                "   o.abnormal_flag_ncid in (1036,1037,1039,1040) and " & _
                "   o.event_id = lt.event_id and " & _
                "   lt.test_ncid = 11158 and " & _
                "   lt.event_id = e.event_id and " & _
                GetOracleEndDate(moReportItem.DateTo) & " >= e.EVENT_START_GMTIME and " & _
                GetOracleEndDate(moReportItem.DateFrom) & " <= e.EVENT_START_GMTIME and " & _
                "   e.event_type_ncid = 485 and " & _
                "   e.unit_number = id.unit_number) and "

    '-- Check if patient was entered before the 'TO' date
    sStmt = sStmt & _
            "(exists " & _
                "(select ma.unit_number " & _
                " from mmi_audit ma " & _
                " where " & _
                        GetOracleEndDate(moReportItem.DateTo) & " >= TX_TIME and " & _
                        "ma.unit_number = id.unit_number) and "

    '-- Check if patient was >= 17 years before 'TO' date
    sStmt = sStmt & _
            "((((" & GetOracleEndDate(moReportItem.DateTo) & " - id.birth_date)/365) >= 17.0) "
    
'    '-- Check if patient has records at specified MTF or is patieint of specified PCM
    sStmt = sStmt & _
            " and " & sWhereSQL & "))"

    goDBSqlOps.Execute sStmt
    
    LDenominator = goDBSqlOps.Value("Denominator")
    
'-- Get Numerator
'---------------------------------------------------------------------------
    sStmt = "select count(*) Numerator " & _
            "from mmi_id id " & sFromSQL & _
            "Where "
    '-- Check for Abnormal Cholesterol Lab Test and then HDL/LDL test within 2 months of Abnormal
    sStmt = sStmt & _
            "exists " & _
                "(select 'x' " & _
                " from observation o, lab_test lt, event e " & _
                " where " & _
                    "exists " & _
                        "(select 'x' " & _
                        " from lab_test lt2, event e2 " & _
                        " where " & _
                        "(e2.EVENT_END_GMTIME - e.EVENT_START_GMTIME) <= 62.0 and " & _
                        "   lt2.test_ncid = 94129 and " & _
                        "   lt2.event_id = e2.event_id and " & _
                        "   e2.event_type_ncid = 485 and " & _
                        "   e2.unit_number = e.unit_number) and " & _
                "   o.abnormal_flag_ncid in (1036,1037,1039,1040) and " & _
                "   o.event_id = lt.event_id and " & _
                "   lt.test_ncid = 11158 and " & _
                "   lt.event_id = e.event_id and " & _
                GetOracleEndDate(moReportItem.DateTo) & " >= e.EVENT_START_GMTIME and " & _
                GetOracleEndDate(moReportItem.DateFrom) & " <= e.EVENT_START_GMTIME and " & _
                "   e.event_type_ncid = 485 and " & _
                "   e.unit_number = id.unit_number) and "
    '-- Check if patient was entered before the 'TO' date
    sStmt = sStmt & _
            "(exists " & _
                "(select ma.unit_number " & _
                " from mmi_audit ma " & _
                " where " & _
                        GetOracleEndDate(moReportItem.DateTo) & " >= TX_TIME and " & _
                        "ma.unit_number = id.unit_number) and "
    '-- Check if patient was >= 17 years before 'TO' date
    sStmt = sStmt & _
            "((((" & GetOracleEndDate(moReportItem.DateTo) & " - id.birth_date)/365) >= 17.0) "
'    '-- Check if patient has records at specified MTF or is patieint of specified PCM
    sStmt = sStmt & _
            " and " & sWhereSQL & "))"
    
    goDBSqlOps.Execute sStmt
    
    LNumerator = goDBSqlOps.Value("Numerator")
    
'-- Set Report Field (Formula Field) Information
'---------------------------------------------------------------------------

 HeaderColl.Add CrystalQuote(CStr(LNumerator)), "Numerator"
 HeaderColl.Add CrystalQuote(CStr(LDenominator)), "Denominator"
 HeaderColl.Add CrystalQuote(moReportItem.CriteriaSel1.Item(1).Name), "Facility"

 DataGetHighCholesterolHDL_LDLFollowUp = True
GoTo CleanUp
ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "ReportOps. DataGetHighCholesterolHDL_LDLFollowUp", App.Title, vbExclamation)
CleanUp:
    Exit Function


    '-- Check for Abnormal Cholesterol Lab Test
'    sStmt = sStmt & _
'            "exists " & _
'                "(select e.event_id " & _
'                " from observation o, event e " & _
'                " where " & _
'                        "o.Abnormal_Flag_NCID = 1037 and " & _
'                        "o.observation_ncid in (6094, 6096, 6097, 6090) and " & _
'                        "o.event_id = e.event_id and " & _
'                        "((" & GetOracleEndDate(moReportItem.DateTo) & " - e.EVENT_START_GMTIME)/365) <= 5.0 and " & _
'                        GetOracleEndDate(moReportItem.DateTo) & " >= e.EVENT_START_GMTIME and " & _
'                        "e.event_type_ncid = 485 and " & _
'                        "e.unit_number = id.unit_number) and "
'-------------------------
'    sStmt = sStmt & _
'            "exists " & _
'                "(select e.event_id " & _
'                " from observation o, event e " & _
'                " where " & _
'                        "exists " & _
'                            "(select e.event_id " & _
'                            " from observation o2, event e2 " & _
'                            "where " & _
'                                "o2.observation_ncid = 51141 and " & _
'                                "o2.event_id = e2.event_id and " & _
'                                "(e2.EVENT_END_GMTIME - e.EVENT_START_GMTIME) <= 62.0 and " & _
'                                "e2.event_type_ncid = 485 and " & _
'                                "e2.unit_number = e.unit_number) and " & _
'                        "o.Abnormal_Flag_NCID = 1037 and " & _
'                        "o.observation_ncid in (6094, 6096, 6097, 6090) and " & _
'                        "o.event_id = e.event_id and " & _
'                        "((" & GetOracleEndDate(moReportItem.DateTo) & " - e.EVENT_START_GMTIME)/365) <= 5.0 and " & _
'                        GetOracleEndDate(moReportItem.DateTo) & " >= e.EVENT_START_GMTIME and " & _
'                        "e.event_type_ncid = 485 and " & _
'                        "e.unit_number = id.unit_number) and "
'------------------------------
    '-- Check for facility
'    sStmt = sStmt & _
'            "exists " & _
'                "(select ms.unit_number " & _
'                " from mmi_site ms " & _
'                " where " & _
'                        moReportItem.CriteriaSel1.Item(1).NCID & " = facility_cid and " & _
'                        "ms.unit_number = id.unit_number) and "
End Function

Private Function DataGetHighCholesterolFollowUpCounseling(HeaderColl As Collection) As Boolean

If Not DebugMode Then On Error GoTo ErrHandler
DataGetHighCholesterolFollowUpCounseling = False

Dim LNumerator As Long
Dim LDenominator As Long
Dim sStmt As String
Dim sFromSQL As String
Dim sWhereSQL As String
'94129 - Lipid Panel Complete, Serum
'11158 - Plasma Total Cholesterol

'1036 Higher Than Critical
'1037 Higher Than Normal
'1039 Lower Than Critical
'1040 Lower Than Normal

    If Not GetMTForPCM_SQL(sFromSQL, sWhereSQL) Then GoTo CleanUp

'-- Get Denominator
'---------------------------------------------------------------------------
    sStmt = "select count(*) Denominator " & _
            "from mmi_id id " & sFromSQL & _
            "Where "
    '-- Check for Abnormal Cholesterol Lab Test
    sStmt = sStmt & _
            "exists " & _
                "(select 'x' " & _
                " from observation o, lab_test lt, event e " & _
                " where " & _
                "   o.abnormal_flag_ncid in (1036,1037) and " & _
                "   o.event_id = lt.event_id and " & _
                "   lt.test_ncid in (11158,94129) and " & _
                "   lt.event_id = e.event_id and " & _
                GetOracleEndDate(moReportItem.DateTo) & " >= e.EVENT_START_GMTIME and " & _
                GetOracleEndDate(moReportItem.DateFrom) & " <= e.EVENT_START_GMTIME and " & _
                "   e.event_type_ncid = 485 and " & _
                "   e.unit_number = id.unit_number) and "

    '-- Check if patient was entered before the 'TO' date
    sStmt = sStmt & _
            "(exists " & _
                "(select ma.unit_number " & _
                " from mmi_audit ma " & _
                " where " & _
                        GetOracleEndDate(moReportItem.DateTo) & " >= TX_TIME and " & _
                        "ma.unit_number = id.unit_number) and "

    '-- Check if patient was >= 17 years before 'TO' date
    sStmt = sStmt & _
            "((((" & GetOracleEndDate(moReportItem.DateTo) & " - id.birth_date)/365) >= 17.0) "
    
'    '-- Check if patient has records at specified MTF or is patieint of specified PCM
    sStmt = sStmt & _
            " and " & sWhereSQL & "))"

    goDBSqlOps.Execute sStmt
    
    LDenominator = goDBSqlOps.Value("Denominator")
    
'-- Get Numerator
'---------------------------------------------------------------------------
    sStmt = "select count(*) Numerator " & _
            "from mmi_id id " & sFromSQL & _
            "Where "
    '-- Check for Abnormal Cholesterol Lab Test
    sStmt = sStmt & _
            "exists " & _
                "(select 'x' " & _
                " from observation o, lab_test lt, event e " & _
                " where " & _
                        "exists " & _
                            "(select 'x' " & _
                            " from clinical_event_header ce2, event e2 " & _
                            "where " & _
                            "ce2.status_ncid = 162068 and " & _
                            "ce2.event_id = e2.event_id and " & _
                            "(e2.EVENT_END_GMTIME - e.EVENT_START_GMTIME) <= 62.0 and " & _
                            "e2.event_subtype_ncid = 207693 and " & _
                            "e2.event_type_ncid = 162156 and " & _
                            "e2.unit_number = e.unit_number) and " & _
                "   o.abnormal_flag_ncid in (1036,1037) and " & _
                "   o.event_id = lt.event_id and " & _
                "   lt.test_ncid in (11158,94129) and " & _
                "   lt.event_id = e.event_id and " & _
                GetOracleEndDate(moReportItem.DateTo) & " >= e.EVENT_START_GMTIME and " & _
                GetOracleEndDate(moReportItem.DateFrom) & " <= e.EVENT_START_GMTIME and " & _
                "   e.event_type_ncid = 485 and " & _
                "   e.unit_number = id.unit_number) and "

    '-- Check if patient was entered before the 'TO' date
    sStmt = sStmt & _
            "(exists " & _
                "(select ma.unit_number " & _
                " from mmi_audit ma " & _
                " where " & _
                        GetOracleEndDate(moReportItem.DateTo) & " >= TX_TIME and " & _
                        "ma.unit_number = id.unit_number) and "

    '-- Check if patient was >= 17 years before 'TO' date
    sStmt = sStmt & _
            "((((" & GetOracleEndDate(moReportItem.DateTo) & " - id.birth_date)/365) >= 17.0) "
    
'    '-- Check if patient has records at specified MTF or is patieint of specified PCM
    sStmt = sStmt & _
            " and " & sWhereSQL & "))"
    
    goDBSqlOps.Execute sStmt
    
    LNumerator = goDBSqlOps.Value("Numerator")
    
'-- Set Report Field (Formula Field) Information
'---------------------------------------------------------------------------

 HeaderColl.Add CrystalQuote(CStr(LNumerator)), "Numerator"
 HeaderColl.Add CrystalQuote(CStr(LDenominator)), "Denominator"
 HeaderColl.Add CrystalQuote(moReportItem.CriteriaSel1.Item(1).Name), "Facility"

DataGetHighCholesterolFollowUpCounseling = True
GoTo CleanUp

ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "ReportOps. DataGetHighCholesterolFollowUpCounseling", App.Title, vbExclamation)
CleanUp:
    Exit Function
End Function


Private Function DataGetMammographyQueryFor(HeaderColl As Collection) As Boolean

If Not DebugMode Then On Error GoTo ErrHandler
DataGetMammographyQueryFor = False

Dim LNumerator As Long
Dim LDenominator As Long
Dim sStmt As String
Dim sFromSQL As String
Dim sWhereSQL As String


If Not GetMTForPCM_SQL(sFromSQL, sWhereSQL) Then GoTo CleanUp

'-- Get Denominator
'---------------------------------------------------------------------------
    sStmt = "select count(*) Denominator " & _
            "from mmi_id Id " & sFromSQL & _
            "Where "
    '-- Check if patient entered in system before 'TO' date
    sStmt = sStmt & _
            "exists " & _
                "(select ma.unit_number " & _
                " from mmi_audit ma " & _
                " where " & _
                        GetOracleEndDate(moReportItem.DateTo) & " >= TX_TIME and " & _
                        "ma.unit_number = id.unit_number) and "

    '-- Check if patient is a Female >= 52 and <= 69
    sStmt = sStmt & _
            "((((" & GetOracleEndDate(moReportItem.DateFrom) & " - id.birth_date)/365) >= 52.0) and " & _
            "(((" & GetOracleEndDate(moReportItem.DateTo) & " - id.birth_date)/365) <= 69.0) and " & _
            "id.sex = 'F')"
            
'    '-- Check if patient has records at specified MTF or is patieint of specified PCM
    sStmt = sStmt & _
            " and " & sWhereSQL

    goDBSqlOps.Execute sStmt
    
    LDenominator = goDBSqlOps.Value("Denominator")

'-- Get Numerator
'---------------------------------------------------------------------------
    sStmt = "select count(*) Numerator " & _
            "from mmi_id id " & sFromSQL & _
            "Where "
    '-- Check for Mammography Result
    sStmt = sStmt & _
            "exists " & _
                "(select 'x' " & _
                " from document d, event e " & _
                " where " & _
                        "(((" & GetOracleEndDate(moReportItem.DateFrom) & " <= e.EVENT_START_GMTIME) and " & _
                        "(" & GetOracleEndDate(moReportItem.DateTo) & " >= e.EVENT_START_GMTIME)) or  " & _
                        "(((e.EVENT_START_GMTIME - " & GetOracleEndDate(moReportItem.DateFrom) & ")/365) <= 2.0)) and " & _
                        "d.report_status_ncid = 205657 and " & _
                        "d.report_type_ncid = 1255 and " & _
                        "d.event_id = e.event_id and " & _
"event_subtype_ncid in (" & _
"124957,125243,125244,125335,125338,125339,125340,125341,125342,125343,125344,125345,125346," & _
"125347,125349,125350,125351,125352,125353,125354,125355,125356,125357,125358,125360,125361," & _
"125362,125363,125364,125365,125368,125369,126171,126172,127000,149300,162240,162241,162242," & _
"162243,182767,182919,182920,213049,213457,213458,213459,213460,213461,14501664,14511016," & _
"14511037,14583723,14583724,14583725,14603448,14603729,14603754) and " & _
                        "e.Event_Type_NCID = 507 and " & _
                        "e.unit_number = id.unit_number) and "
    '-- Check if patient entered in system before 'TO' date
    sStmt = sStmt & _
            "exists " & _
                "(select ma.unit_number " & _
                " from mmi_audit ma " & _
                " where " & _
                        GetOracleEndDate(moReportItem.DateTo) & " >= TX_TIME and " & _
                        "ma.unit_number = id.unit_number) and "
'    '-- Check for valid facility
'    sStmt = sStmt & _
'            "exists " & _
'                "(select ms.unit_number " & _
'                " from mmi_site ms " & _
'                " where " & _
'                        moReportItem.CriteriaSel1.Item(1).NCID & " = facility_cid and " & _
'                        "ms.unit_number = id.unit_number) and "
    '-- Check if patient is a Female >= 52 and  <= 69
    sStmt = sStmt & _
            "((((" & GetOracleEndDate(moReportItem.DateFrom) & " - id.birth_date)/365) >= 52.0) and " & _
            "(((" & GetOracleEndDate(moReportItem.DateTo) & " - id.birth_date)/365) <= 69.0) and " & _
            "id.sex = 'F')"
    
'    '-- Check if patient has records at specified MTF or is patieint of specified PCM
    sStmt = sStmt & _
            " and " & sWhereSQL

    
    goDBSqlOps.Execute sStmt
    
    LNumerator = goDBSqlOps.Value("Numerator")

'-- Set Report Field (Formual Field) Information
'---------------------------------------------------------------------------

 HeaderColl.Add CrystalQuote(CStr(LNumerator)), "Numerator"
 HeaderColl.Add CrystalQuote(CStr(LDenominator)), "Denominator"
 HeaderColl.Add CrystalQuote(moReportItem.CriteriaSel1.Item(1).Name), "Facility"

DataGetMammographyQueryFor = True
GoTo CleanUp
ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "ReportOps.DataGetMammographyQueryFor", App.Title, vbExclamation)
CleanUp:
    Exit Function
'    '-- Check for valid facility
'    sStmt = sStmt & _
'            "exists " & _
'                "(select ms.unit_number " & _
'                " from mmi_site ms " & _
'                " where " & _
'                        moReportItem.CriteriaSel1.Item(1).NCID & " = facility_cid and " & _
'                        "ms.unit_number = id.unit_number) and "
End Function

Private Function DataGetPAPSmearQueryFor(HeaderColl As Collection) As Boolean

If Not DebugMode Then On Error GoTo ErrHandler
DataGetPAPSmearQueryFor = False

Dim LNumerator As Long
Dim LDenominator As Long
Dim sStmt As String
Dim sWhereCriteria As String
Dim sFromCriteria As String


If Not GetMTForPCM_SQL(sFromCriteria, sWhereCriteria) Then GoTo CleanUp

'-- Get Denominator
'---------------------------------------------------------------------------
    sStmt = "select count(*) Denominator " & _
            "from mmi_id Id " & sFromCriteria & _
            "Where "
    '-- Check if patient entered in system before 'TO' date
    sStmt = sStmt & _
            "(exists " & _
                "(select ma.unit_number " & _
                " from mmi_audit ma " & _
                " where " & _
                        GetOracleEndDate(moReportItem.DateTo) & " >= TX_TIME and " & _
                        "ma.unit_number = id.unit_number) and "
'    '-- Check for valid facility
'    sStmt = sStmt & _
'            "(exists " & _
'                "(select ms.unit_number " & _
'                " from mmi_site ms " & _
'                " where " & _
'                        moReportItem.CriteriaSel1.Item(1).NCID & " = facility_cid and " & _
'                        "ms.unit_number = id.unit_number) and "
    '-- Check if patient is a Female >= 21 or Female >= 64
    sStmt = sStmt & _
            "(((((" & GetOracleEndDate(moReportItem.DateTo) & " - id.birth_date)/365) >= 21.0) and " & _
            "(((" & GetOracleEndDate(moReportItem.DateTo) & " - id.birth_date)/365) <= 64.0) and " & _
            "id.sex = 'F')"
    
    '-- Check if patient has records at specified MTF or is patieint of specified PCM
    sStmt = sStmt & _
            " and " & sWhereCriteria & "))"
            
    goDBSqlOps.Execute sStmt
    
    LDenominator = goDBSqlOps.Value("Denominator")

'-- Get Numerator
'---------------------------------------------------------------------------
    sStmt = "select count(*) Numerator " & _
            "from mmi_id id " & sFromCriteria & _
            "Where "
    '-- Check for Pap Smear Lab Test Result
    sStmt = sStmt & _
            "((exists " & _
                "(select e.event_id " & _
                " from event e " & _
                " where " & _
                        "((((" & GetOracleEndDate(moReportItem.DateFrom) & " <= e.EVENT_START_GMTIME) and " & _
                        "(" & GetOracleEndDate(moReportItem.DateTo) & " >= e.EVENT_START_GMTIME)) or  " & _
                        "(((e.EVENT_START_GMTIME - " & GetOracleEndDate(moReportItem.DateFrom) & ")/365) <= 3.0)) and " & _
                        "(e.event_subtype_ncid = 90269 and (e.event_type_ncid = 485 and " & _
                        "(e.unit_number = id.unit_number)))))) and "
'    sStmt = sStmt & _
'            "(exists " & _
'                "(select e.event_id " & _
'                " from observation o, event e " & _
'                " where " & _
'                        "o.observation_ncid = 66560 and " & _
'                        "o.event_id = e.event_id and " & _
'                        "((" & GetOracleEndDate(moReportItem.DateTo) & " - e.EVENT_START_GMTIME)/365) <= 3.0 and " & _
'                        GetOracleEndDate(moReportItem.DateTo) & " >= e.EVENT_START_GMTIME and " & _
'                        "e.event_subtype_ncid = 90269 and e.event_type_ncid = 485 and " & _
'                        "e.unit_number = id.unit_number) and "
                        

    '-- Check if patient entered in system before 'TO' date
    sStmt = sStmt & _
            "(exists " & _
                "(select ma.unit_number " & _
                " from mmi_audit ma " & _
                " where " & _
                        GetOracleEndDate(moReportItem.DateTo) & " >= TX_TIME and " & _
                        "ma.unit_number = id.unit_number) and "
                        
'    '-- Check for valid facility
'    sStmt = sStmt & _
'            "exists " & _
'                "(select ms.unit_number " & _
'                " from mmi_site ms " & _
'                " where " & _
'                        moReportItem.CriteriaSel1.Item(1).NCID & " = facility_cid and " & _
'                        "ms.unit_number = id.unit_number) and "

    '-- Check if patient is a Female >= 21 or Female >= 64
    sStmt = sStmt & _
            "((((" & GetOracleEndDate(moReportItem.DateTo) & " - id.birth_date)/365) <= 64.0) and " & _
            "((((" & GetOracleEndDate(moReportItem.DateTo) & " - id.birth_date)/365) >= 21.0) and " & _
            "(id.sex = 'F'"

    '-- Check if patient has records at specified MTF or is patieint of specified PCM
    sStmt = sStmt & _
            " and (" & sWhereCriteria & "))))))"

    goDBSqlOps.Execute sStmt
    
    LNumerator = goDBSqlOps.Value("Numerator")

'-- Set Report Field (Formula Field) Information
'---------------------------------------------------------------------------

HeaderColl.Add CrystalQuote(CStr(LNumerator)), "Numerator"
HeaderColl.Add CrystalQuote(CStr(LDenominator)), "Denominator"
HeaderColl.Add CrystalQuote(moReportItem.CriteriaSel1.Item(1).Name), "Facility"

DataGetPAPSmearQueryFor = True
GoTo CleanUp
ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "ReportOps.DataGetPAPSmearQueryFor", App.Title, vbExclamation)
CleanUp:
    Exit Function
End Function
Private Function DataGetCPSSummary(HeaderColl As Collection) As Boolean

If Not DebugMode Then On Error GoTo ErrHandler
DataGetCPSSummary = False

Dim oPatient As PatientItem
Dim oPatientSearch As New PatientSearch
Dim i As Integer
Dim rs As ADODB.Recordset
Dim collReminders As Collection
Dim sFromSQL As String
Dim sWhereSQL As String
Dim sStmt As String

    If Not moTableOps.DeleteFromTable("Patient") Then GoTo CleanUp
    If Not moTableOps.DeleteFromTable("Preventive") Then GoTo CleanUp
    
    If Not GetMTForPCM_SQL(sFromSQL, sWhereSQL) Then GoTo CleanUp

'hwc0
'Convert to use MMM_DEFAULT_TEXT_RSFORM_MV to improve performance

sStmt = "select /*+ AND_EQUAL(MMM_DEFAULT_TEXT_RSFORM_MV_X1 MMM_DEFAULT_TEXT_RSFORM_MV_X2) " & vbCrLf & _
            "INDEX (MMM_DEFAULT_TEXT_RSFORM_MV MMM_DEFAULT_TEXT_RSFORM_MV_X1) " & vbCrLf & _
            "INDEX (MMM_DEFAULT_TEXT_RSFORM_MV MMM_DEFAULT_TEXT_RSFORM_MV_X2) " & vbCrLf & _
            "INDEX (MMI_GENERIC_ID MMI_GENERIC_ID_X2) " & vbCrLf & _
            "INDEX (MMI_ID MMI_ID_PK) " & vbCrLf & _
            "INDEX (EVENT EVENT_X1) " & vbCrLf & _
            "INDEX (clinical_event_header CLINICAL_EVENT_HEADER_PK) */ " & vbCrLf & _
        "e.unit_number as PatientId, " & vbCrLf & _
        "e.event_type_ncid, " & vbCrLf & _
        "e.event_subtype_ncid, " & vbCrLf & _
        "co.observation_ncid, " & vbCrLf & _
        "ov.value_datetime as EarliestRecommendedDate, " & vbCrLf & _
        "r.Representation as status, " & vbCrLf & _
        "r2.Representation As Description "

'No need massaging sFromSQL
'If Len(Trim(sFromSQL)) <> 0 Then
'    sFromSQL = Trim(sFromSQL)      'Trim leading and trailing spaces
'    If Left(sFromSQL, 1) = "," Then sFromSQL = Mid(sFromSQL, 2)  'Trim leading ","
'    If Not InStr(1, sFromSQL, vbCrLf, vbTextCompare) = 0 Then    'Trim CRLF
'       sFromSQL = Left(sFromSQL, InStr(1, sFromSQL, vbCrLf, vbTextCompare) - 1)
'    End If
'    sFromSQL = sFromSQL & ", " & vbCrLf     'Add ", "
'End If

'rearrange From Clause
'sStmt = sStmt & vbCrLf & _
'        "From " & vbCrLf & _
'        sFromSQL & _
'        "mmi_id id, event e, clinical_event_header c, " & _
'        "clinical_observation co, " & vbCrLf & _
'        "observation_value ov, " & vbCrLf & _
'        "rsform r, " & vbCrLf & _
'        "rsform_context rc, " & vbCrLf & _
'        "rsform r2, " & vbCrLf & _
'        "rsform_context rc2 " & vbCrLf

sStmt = sStmt & vbCrLf & _
        "From " & vbCrLf & _
        "event e, clinical_event_header c, " & _
        "clinical_observation co, " & vbCrLf & _
        "observation_value ov, " & vbCrLf & _
        "MMM_DEFAULT_TEXT_RSFORM_MV r, " & vbCrLf & _
        "MMM_DEFAULT_TEXT_RSFORM_MV r2"

If Len(Trim(sFromSQL)) <> 0 Then
    If Not InStr(1, sFromSQL, vbCrLf, vbTextCompare) = 0 Then    'Trim CRLF
       sFromSQL = Left(sFromSQL, InStr(1, sFromSQL, vbCrLf, vbTextCompare) - 1)
    End If
    sStmt = sStmt & sFromSQL
End If
sStmt = sStmt & ", mmi_id id " & vbCrLf        'use mmi_id as driving table
        
'sStmt = sStmt & _
'        "Where " & vbCrLf & _
'        "rc2.context_ncid = 2000 and " & vbCrLf & _
'        "rc2.rsform_id = r2.rsform_id and " & vbCrLf & _
'        "r2.ncid = e.event_subtype_ncid and " & vbCrLf & _
'        "rc.context_ncid = 2000 and " & vbCrLf & _
'        "rc.rsform_id = r.rsform_id and " & vbCrLf & _
'        "rc2.preferred_score = 0 and rc.preferred_score = 0 and " & vbCrLf & _
'        "r.ncid = c.status_ncid and " & vbCrLf & _
'        "ov.OBSERVATION_SEQ (+) = co.OBSERVATION_SEQ and " & vbCrLf & _
'        "ov.event_id (+) = co.event_id and " & vbCrLf & _
'        "co.observation_ncid = 162168 and " & vbCrLf & _
'        "co.event_id (+) = e.event_id and " & vbCrLf & _
'        "((c.status_ncid <> 1024 and " & vbCrLf & _
'        "e.event_type_ncid = 162156) and " & vbCrLf & _
'        "(c.event_id = e.event_id and " & vbCrLf & _
'        "(e.event_type_ncid = 162156 and (e.unit_number = id.unit_number "
sStmt = sStmt & _
        "Where " & vbCrLf & _
        "r2.ncid = e.event_subtype_ncid and " & vbCrLf & _
        "r.ncid = c.status_ncid and " & vbCrLf & _
        "ov.OBSERVATION_SEQ (+) = co.OBSERVATION_SEQ and " & vbCrLf & _
        "ov.event_id (+) = co.event_id and " & vbCrLf & _
        "co.observation_ncid = 162168 and " & vbCrLf & _
        "co.event_id (+) = e.event_id and " & vbCrLf & _
        "((c.status_ncid <> 1024 and " & vbCrLf & _
        "e.event_type_ncid = 162156) and " & vbCrLf & _
        "(c.event_id = e.event_id and " & vbCrLf & _
        "(e.event_type_ncid = 162156 and (e.unit_number = id.unit_number "
        

'    '-- Check if patient has records at specified MTF or is patieint of specified PCM
    sStmt = sStmt & _
            " and (" & sWhereSQL & ")))))"

    goDBSqlOps.Execute sStmt

    If Not moTableOps.OpenTable("Preventive", rs) Then GoTo CleanUp

    Do While Not goDBSqlOps.EOF
        rs.AddNew
        rs("PatientId") = goDBSqlOps.Value("PatientId")
        rs("desc") = goDBSqlOps.Value("Description")
        rs("status") = goDBSqlOps.Value("Status") 'GetReminderStatus(goDBSqlOps.Value("EarliestRecommendedDate"))
        rs("DateDue") = Format(goDBSqlOps.Value("EarliestRecommendedDate"), "dd MMM yyyy")
        rs.Update
        goDBSqlOps.MoveNext
    Loop

    If Not moTableOps.GetPatientTableData("Preventive", "PatientId") Then GoTo CleanUp

    If Not moTableOps.GetSponsorTableData("Patient", "SponsorSSN") Then GoTo CleanUp


DataGetCPSSummary = True
GoTo CleanUp
ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "ReportOps.DataGetCPSServicesDue", App.Title, vbExclamation)
CleanUp:
    Set oPatientSearch = Nothing
    Set oPatient = Nothing
    Set rs = Nothing
    Exit Function
End Function
Private Function GetReminderStatus(vDate As Variant) As String

On Error GoTo ErrHandler

If vDate <> gdNULL_DATE And vDate <= Now Then
    GetReminderStatus = "Due Today"
ElseIf vDate <> gdNULL_DATE And vDate > Now Then
    GetReminderStatus = "Next Exam Due"
Else    '--unknown dates to compare against--
    GetReminderStatus = "Next Exam Due"
End If

Exit Function
ErrHandler:
    GetReminderStatus = "Next Exam Due"
End Function

Private Function DataGetPreventionReportCard(HeaderColl As Collection) As Boolean

If Not DebugMode Then On Error GoTo ErrHandler
DataGetPreventionReportCard = False

Dim sTotalHear As String
Dim sNumeratorDrink As String
Dim sNumeratorExercise As String
Dim sNumeratorHypertension As String
Dim sNumeratorSatisfaction As String
Dim sNumeratorSmoking As String
Dim sNumeratorHealth As String
Dim sNumeratorBMI As String
Dim sStmt As String
Dim sWhereCriteria As String
Dim sFromCriteria As String



If Not GetMTForPCM_SQL(sFromCriteria, sWhereCriteria) Then GoTo CleanUp


'-- Get TotalHear (Denominator)
'---------------------------------------------------------------------------
sStmt = "select count(Id.unit_number) TotalHear " & _
        "from mmi_genericdata Id " & sFromCriteria & _
        "Where "

    '-- Check for Hears questionare
    sStmt = sStmt & _
            "(exists " & _
                "(select e.event_id " & _
                " from event e " & _
                " where " & _
                        GetOracleStartDate(moReportItem.DateFrom) & " <= EVENT_START_GMTIME and " & _
                        GetOracleEndDate(moReportItem.DateTo) & " >= EVENT_START_GMTIME and " & _
                        "(e.event_type_ncid = 204764 and " & _
                        "(e.unit_number = id.unit_number))) and "


    '-- Check for correct facility
'    sStmt = sStmt & _
'            "exists " & _
'                "(select ms.unit_number " & _
'                " from mmi_site ms " & _
'                " where " & _
'                        moReportItem.CriteriaSel1.Item(1).NCID & " = facility_cid and " & _
'                        "ms.unit_number = id.unit_number) and "

    '-- Check if patient is Active Duty
    '-- Not sure how to check for active duty so checking for Unit Identification Code
    sStmt = sStmt & _
                "(Id.value_ncid in " & _
                "(119033,119034,119035,119068,119069,119070, " & _
                 "119122,119123,119124,119148,119149,119150) and " & _
                "Id.type_ncid = 14501721  "

    '-- Check if patient has records at specified MTF or is patieint of specified PCM
    sStmt = sStmt & _
            " and (" & sWhereCriteria & ")))"

    goDBSqlOps.Execute sStmt
    sTotalHear = CStr(goDBSqlOps.Value("TotalHear"))

'-- Get Active Duty at risk of alcohol abuse
'---------------------------------------------------------------------------
    sStmt = "select count(id.unit_number) NumeratorDrink " & _
            "from mmi_genericdata id " & sFromCriteria & _
            "Where "
    '-- Check for Coupler Event where user responded 'Yes' or 'No' Finding 691,239
    sStmt = sStmt & _
            "(exists " & _
                "(select e.event_id " & _
                " from observation_value o, clinical_observation c, event e " & _
                " where " & _
                        "(o.value_ncid in " & _
                        "(34688,34689,34690,34693,34694,34695,34698,24824, " & _
                        "34699,34703,34704,34705,34708,34709,34710,34713,34714,34715, " & _
                        "34718,34719,34720,34721,1751434064) and " & _
                        "(o.observation_seq = c.observation_seq and " & _
                        "(o.event_id = c.event_id and " & _
                        "(c.observation_ncid = 204771 and " & _
                        "(c.event_id = e.event_id and (" & _
                        GetOracleStartDate(moReportItem.DateFrom) & " <= EVENT_START_GMTIME and (" & _
                        GetOracleEndDate(moReportItem.DateTo) & " >= EVENT_START_GMTIME and " & _
                        "(e.event_subtype_ncid = 33971 and (e.event_type_ncid = 204764 and " & _
                        "(e.unit_number = id.unit_number))))))))))) and "
    '-- Added for effiency
    'sStmt = sStmt & _
            "exists " & _
                "(" & GetOracleStartDate(moReportItem.DateFrom) & " <= EVENT_START_GMTIME and " & _
                GetOracleEndDate(moReportItem.DateTo) & " >= EVENT_START_GMTIME and " & _
                "e.event_type_ncid = 204764 and " & _
                "e.unit_number = id.unit_number) and "
    '-- Check for valid facility
'    sStmt = sStmt & _
'            "exists " & _
'                "(select ms.unit_number " & _
'                " from mmi_site ms " & _
'                " where " & _
'                        moReportItem.CriteriaSel1.Item(1).NCID & " = facility_cid and " & _
'                        "ms.unit_number = id.unit_number) and "
    '-- Check if patient is Active Duty
    '-- Not sure how to check for active duty so checking for Unit Identification Code
    sStmt = sStmt & _
                "(Id.value_ncid in " & _
                "(119033,119034,119035,119068,119069,119070, " & _
                 "119122,119123,119124,119148,119149,119150) and " & _
                "(Id.type_ncid = 14501721 "

      '-- Check if patient has records at specified MTF or is patieint of specified PCM
    sStmt = sStmt & _
            " and (" & sWhereCriteria & "))))"

    goDBSqlOps.Execute sStmt
    
    sNumeratorDrink = CStr(goDBSqlOps.Value("NumeratorDrink"))

'-- Get Active Duty who aren't Active
'---------------------------------------------------------------------------
    sStmt = "select count(id.unit_number) NumeratorExercise " & _
            "from mmi_genericdata id " & sFromCriteria & _
            "Where "
    '-- Check for Coupler Event where user responded 'Yes' or 'No' Finding 691,239
    sStmt = sStmt & _
            "(exists " & _
                "(select e.event_id " & _
                " from clinical_observation c, observation_value o, event e " & _
                " where " & _
                        "(o.value_ncid in " & _
                        "(34537) and " & _
                        "(o.observation_seq = c.observation_seq and " & _
                        "(o.event_id = c.event_id and " & _
                        "(c.observation_ncid = 204771 and " & _
                        "(c.event_id = e.event_id and (" & _
                        GetOracleStartDate(moReportItem.DateFrom) & " <= EVENT_START_GMTIME and (" & _
                        GetOracleEndDate(moReportItem.DateTo) & " >= EVENT_START_GMTIME and " & _
                        "(e.event_subtype_ncid = 33971 and (e.event_type_ncid = 204764 and " & _
                        "(e.unit_number = id.unit_number))))))))))) and "
    '-- Check for valid facility
'    sStmt = sStmt & _
'            "exists " & _
'                "(select ms.unit_number " & _
'                " from mmi_site ms " & _
'                " where " & _
'                        moReportItem.CriteriaSel1.Item(1).NCID & " = facility_cid and " & _
'                        "ms.unit_number = id.unit_number) and  "
    '-- Check if patient is Active Duty
    '-- Not sure how to check for active duty so checking for Unit Identification Code
    sStmt = sStmt & _
                "(Id.value_ncid in " & _
                "(119033,119034,119035,119068,119069,119070, " & _
                 "119122,119123,119124,119148,119149,119150) and " & _
                "(Id.type_ncid = 14501721 "
    
    '-- Check if patient has records at specified MTF or is patieint of specified PCM
    sStmt = sStmt & _
            " and (" & sWhereCriteria & "))))"
    
    goDBSqlOps.Execute sStmt
    
    sNumeratorExercise = CStr(goDBSqlOps.Value("NumeratorExercise"))

'-- Get Active Duty who aren't consistent with Hypertension Medication
'---------------------------------------------------------------------------
    sStmt = "select count(id.unit_number) NumeratorHypertension " & _
            "from mmi_genericdata id " & sFromCriteria & _
            "Where "
    '-- Check for Coupler Event where user responded 'Yes' or 'No' Finding 691,239
    sStmt = sStmt & _
            "(exists " & _
                "(select e.event_id " & _
                " from clinical_observation c, observation_value o, event e " & _
                " where " & _
                        "(o.value_ncid in " & _
                        "(34795,34796,34797) and " & _
                        "(o.observation_seq = c.observation_seq and " & _
                        "(o.event_id = c.event_id and " & _
                        "(c.observation_ncid = 204771 and " & _
                        "(c.event_id = e.event_id and (" & _
                        GetOracleStartDate(moReportItem.DateFrom) & " <= EVENT_START_GMTIME and (" & _
                        GetOracleEndDate(moReportItem.DateTo) & " >= EVENT_START_GMTIME and " & _
                        "(e.event_subtype_ncid = 33971 and (e.event_type_ncid = 204764 and " & _
                        "(e.unit_number = id.unit_number))))))))))) and "
    '-- Check for valid facility
'    sStmt = sStmt & _
'            "exists " & _
'                "(select ms.unit_number " & _
'                " from mmi_site ms " & _
'                " where " & _
'                        moReportItem.CriteriaSel1.Item(1).NCID & " = facility_cid and " & _
'                        "ms.unit_number = id.unit_number) and "
    '-- Check if patient is Active Duty
    '-- Not sure how to check for active duty so checking for Unit Identification Code
    sStmt = sStmt & _
                "(Id.value_ncid in " & _
                "(119033,119034,119035,119068,119069,119070, " & _
                 "119122,119123,119124,119148,119149,119150) and " & _
                "(Id.type_ncid = 14501721 "

    '-- Check if patient has records at specified MTF or is patieint of specified PCM
    sStmt = sStmt & _
            " and (" & sWhereCriteria & "))))"

    goDBSqlOps.Execute sStmt
    
    sNumeratorHypertension = CStr(goDBSqlOps.Value("NumeratorHypertension"))

'-- Get Active Duty who are not satisfied with their lives
'---------------------------------------------------------------------------
    sStmt = "select count(id.unit_number) NumeratorSatisfaction " & _
            "from mmi_genericdata id " & sFromCriteria & _
            "Where "
    '-- Check for Coupler Event where user responded 'Yes' or 'No' Finding 691,239
    sStmt = sStmt & _
            "(exists " & _
                "(select e.event_id " & _
                " from clinical_observation c, observation_value o, event e " & _
                " where " & _
                        "(o.value_ncid in " & _
                        "(34969) and " & _
                        "(o.observation_seq = c.observation_seq and " & _
                        "(o.event_id = c.event_id and " & _
                        "(c.observation_ncid = 204771 and " & _
                        "(c.event_id = e.event_id and (" & _
                        GetOracleStartDate(moReportItem.DateFrom) & " <= EVENT_START_GMTIME and (" & _
                        GetOracleEndDate(moReportItem.DateTo) & " >= EVENT_START_GMTIME and " & _
                        "(e.event_subtype_ncid = 33971 and (e.event_type_ncid = 204764 and " & _
                        "(e.unit_number = id.unit_number))))))))))) and "
    '-- Check for valid facility
'    sStmt = sStmt & _
'            "exists " & _
'                "(select ms.unit_number " & _
'                " from mmi_site ms " & _
'                " where " & _
'                        moReportItem.CriteriaSel1.Item(1).NCID & " = facility_cid and " & _
'                        "ms.unit_number = id.unit_number) and "
    '-- Check if patient is Active Duty
    '-- Not sure how to check for active duty so checking for Unit Identification Code
    sStmt = sStmt & _
                "(Id.value_ncid in " & _
                "(119033,119034,119035,119068,119069,119070, " & _
                 "119122,119123,119124,119148,119149,119150) and " & _
                "(Id.type_ncid = 14501721 "
           
    '-- Check if patient has records at specified MTF or is patieint of specified PCM
    sStmt = sStmt & _
            " and (" & sWhereCriteria & "))))"

    goDBSqlOps.Execute sStmt
    
    sNumeratorSatisfaction = CStr(goDBSqlOps.Value("NumeratorSatisfaction"))

'-- Get Active Duty who smoke everyday or some days
'---------------------------------------------------------------------------
    sStmt = "select count(id.unit_number) NumeratorSmoking " & _
            "from mmi_genericdata id " & sFromCriteria & _
            "Where "
    '-- Check for Coupler Event where user responded 'Yes' or 'No' Finding 691,239
    sStmt = sStmt & _
            "(exists " & _
                "(select e.event_id " & _
                " from clinical_observation c, observation_value o, event e " & _
                " where " & _
                        "(o.value_ncid in " & _
                        "(34577,34591) and " & _
                        "(o.observation_seq = c.observation_seq and " & _
                        "(o.event_id = c.event_id and " & _
                        "(c.observation_ncid = 204771 and " & _
                        "(c.event_id = e.event_id and (" & _
                        GetOracleStartDate(moReportItem.DateFrom) & " <= EVENT_START_GMTIME and (" & _
                        GetOracleEndDate(moReportItem.DateTo) & " >= EVENT_START_GMTIME and " & _
                        "(e.event_subtype_ncid = 33971 and (e.event_type_ncid = 204764 and " & _
                        "(e.unit_number = id.unit_number))))))))))) and "
    '-- Check for valid facility
'    sStmt = sStmt & _
'            "exists " & _
'                "(select ms.unit_number " & _
'                " from mmi_site ms " & _
'                " where " & _
'                        moReportItem.CriteriaSel1.Item(1).NCID & " = facility_cid and " & _
'                        "ms.unit_number = id.unit_number) and "
    '-- Check if patient is Active Duty
    '-- Not sure how to check for active duty so checking for Unit Identification Code
    sStmt = sStmt & _
                "(Id.value_ncid in " & _
                "(119033,119034,119035,119068,119069,119070, " & _
                 "119122,119123,119124,119148,119149,119150) and " & _
                "(Id.type_ncid = 14501721 "

    '-- Check if patient has records at specified MTF or is patieint of specified PCM
    sStmt = sStmt & _
            " and (" & sWhereCriteria & "))))"

    goDBSqlOps.Execute sStmt
    
    sNumeratorSmoking = CStr(goDBSqlOps.Value("NumeratorSmoking"))

'-- Get Active Duty who answer fair or poor to health question
'---------------------------------------------------------------------------
    sStmt = "select count(id.unit_number) NumeratorHealth " & _
            "from mmi_genericdata id " & sFromCriteria & _
            "Where "
    '-- Check for Coupler Event where user responded 'Yes' or 'No' Finding 691,239
    sStmt = sStmt & _
            "(exists " & _
                "(select e.event_id " & _
                " from clinical_observation c, observation_value o, event e " & _
                " where " & _
                        "(o.value_ncid in " & _
                        "(34516,34515) and " & _
                        "(o.observation_seq = c.observation_seq and " & _
                        "(o.event_id = c.event_id and " & _
                        "(c.observation_ncid = 204771 and " & _
                        "(c.event_id = e.event_id and (" & _
                        GetOracleStartDate(moReportItem.DateFrom) & " <= EVENT_START_GMTIME and (" & _
                        GetOracleEndDate(moReportItem.DateTo) & " >= EVENT_START_GMTIME and " & _
                        "(e.event_subtype_ncid = 33971 and (e.event_type_ncid = 204764 and " & _
                        "(e.unit_number = id.unit_number))))))))))) and "
    '-- Check for valid facility
'    sStmt = sStmt & _
'            "exists " & _
'                "(select ms.unit_number " & _
'                " from mmi_site ms " & _
'                " where " & _
'                        moReportItem.CriteriaSel1.Item(1).NCID & " = facility_cid and " & _
'                        "ms.unit_number = id.unit_number) and "
        '-- Check if patient is Active Duty
    '-- Not sure how to check for active duty so checking for Unit Identification Code
    sStmt = sStmt & _
                "(Id.value_ncid in " & _
                "(119033,119034,119035,119068,119069,119070, " & _
                 "119122,119123,119124,119148,119149,119150) and " & _
                "(Id.type_ncid = 14501721 "

    '-- Check if patient has records at specified MTF or is patieint of specified PCM
    sStmt = sStmt & _
            " and (" & sWhereCriteria & "))))"

    goDBSqlOps.Execute sStmt
    
    sNumeratorHealth = CStr(goDBSqlOps.Value("NumeratorHealth"))

'-- Get Active Duty who are over their body mass index (BMI)
'---------------------------------------------------------------------------
    sStmt = "select count(id.unit_number) NumeratorBMI " & _
            "from mmi_genericdata id " & sFromCriteria & _
            "Where "

    '-- Check for Coupler Event where user responded 'Yes' or 'No' Finding 691,239
    sStmt = sStmt & _
            "(exists " & _
                "(select e.event_id " & _
                " from " & _
                        "observation_value o2, observation_value o, clinical_observation c2," & _
                        "clinical_observation c,  event e " & _
                " where " & _
                        "(((o2.value_ncid = 34350 and " & _
                        "o.value_ncid = 34001) or " & _
                        "(o2.value_ncid = 34351 and " & _
                        "o.value_ncid = 34047)) and " & _
                        "(o2.observation_seq = c2.observation_seq and " & _
                        "(o.observation_seq = c.observation_seq and " & _
                        "(o2.event_id = c2.event_id and " & _
                        "(o.event_id = c.event_id and " & _
                        "(c2.observation_ncid = 204771 and " & _
                        "(c.observation_ncid = 204771 and " & _
                        "(c2.event_id = e.event_id and " & _
                        "(c.event_id = e.event_id and (" & _
                        GetOracleStartDate(moReportItem.DateFrom) & " <= EVENT_START_GMTIME and (" & _
                        GetOracleEndDate(moReportItem.DateTo) & " >= EVENT_START_GMTIME and " & _
                        "(e.event_subtype_ncid in (33971, 33972) and (e.event_type_ncid = 204764 and " & _
                        "(e.unit_number = id.unit_number))))))))))))))) and "
    '-- Check for valid facility
'    sStmt = sStmt & _
'            "exists " & _
'                "(select ms.unit_number " & sFromCriteria & _
'                " from mmi_site ms " & _
'                " where " & _
'                        moReportItem.CriteriaSel1.Item(1).NCID & " = facility_cid and " & _
'                        "ms.unit_number = id.unit_number) and "
    '-- Check if patient is Active Duty
    '-- Not sure how to check for active duty so checking for Unit Identification Code
    sStmt = sStmt & _
                "(Id.value_ncid in " & _
                "(119033,119034,119035,119068,119069,119070, " & _
                 "119122,119123,119124,119148,119149,119150) and " & _
                "(Id.type_ncid = 14501721 "

    '-- Check if patient has records at specified MTF or is patieint of specified PCM
    sStmt = sStmt & _
            " and (" & sWhereCriteria & "))))"


    goDBSqlOps.Execute sStmt
    
    sNumeratorBMI = CStr(goDBSqlOps.Value("NumeratorBMI"))

'-- Set Report Field (Formula Field) Information
'---------------------------------------------------------------------------
HeaderColl.Add CrystalQuote(sTotalHear), "TotalHEAR"
HeaderColl.Add CrystalQuote(sNumeratorDrink), "NumeratorDrinking"
HeaderColl.Add CrystalQuote(sNumeratorExercise), "NumeratorExercise"
HeaderColl.Add CrystalQuote(sNumeratorHypertension), "NumeratorHypertension"
HeaderColl.Add CrystalQuote(sNumeratorSatisfaction), "NumeratorSatisfaction"
HeaderColl.Add CrystalQuote(sNumeratorSmoking), "NumeratorSmoking"
HeaderColl.Add CrystalQuote(sNumeratorHealth), "NumeratorHealth"
HeaderColl.Add CrystalQuote(sNumeratorBMI), "NumeratorBMI"
HeaderColl.Add CrystalQuote(moReportItem.CriteriaSel1.Item(1).Name), "Facility"

DataGetPreventionReportCard = True
GoTo CleanUp
ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "ReportOps.DataGetPreventionReportCard", App.Title, vbExclamation)
CleanUp:
    Exit Function
End Function

Private Function DataGetAlcoholQueryforPotentialHeavyUse(HeaderColl As Collection) As Boolean

If Not DebugMode Then On Error GoTo ErrHandler
DataGetAlcoholQueryforPotentialHeavyUse = False

Dim LNumerator As Long
Dim LDenominator As Long
Dim sStmt As String
Dim sWhereCriteria As String
Dim sFromCriteria As String


If Not GetMTForPCM_SQL(sFromCriteria, sWhereCriteria) Then GoTo CleanUp

'-- Get Denominator
'---------------------------------------------------------------------------
    sStmt = "select count(*) Denominator " & _
            "from mmi_id Id " & sFromCriteria & _
            "Where "
    '-- Check if patient entered in system before 'TO' date
    sStmt = sStmt & _
            "(exists " & _
                "(select ma.unit_number " & _
                " from mmi_audit ma " & _
                " where (" & _
                        GetOracleEndDate(moReportItem.DateTo) & " >= TX_TIME and " & _
                        "(ma.unit_number = id.unit_number))) and "
'    '-- Check for valid facility
'    sStmt = sStmt & _
'            "exists " & _
'                "(select ms.unit_number " & _
'                " from mmi_site ms " & _
'                " where " & _
'                        moReportItem.CriteriaSel1.Item(1).NCID & " = facility_cid and " & _
'                        "ms.unit_number = id.unit_number) and "
    '-- Check if patient >= 17 During reporting period
    sStmt = sStmt & _
            "(((" & GetOracleEndDate(moReportItem.DateTo) & " - id.birth_date)/365) >= 17.0 "

    '-- Check if patient has records at specified MTF or is patieint of specified PCM
    sStmt = sStmt & _
            " and (" & sWhereCriteria & ")))"

    goDBSqlOps.Execute sStmt
    
    LDenominator = goDBSqlOps.Value("Denominator")
    

'-- Get Numerator
'---------------------------------------------------------------------------
    sStmt = "select count(*) Numerator " & _
            "from mmi_id id " & sFromCriteria & _
            "Where "
    '-- Check for Coupler Event where user responded 'Yes' or 'No' Finding 691,239
    sStmt = sStmt & _
            "(exists " & _
                "(select e.event_id " & _
                " from observation_value o, clinical_observation c, event e " & _
                " where (" & _
                        GetOracleEndDate(moReportItem.DateTo) & " >= EVENT_START_GMTIME and " & _
                        "(o.value_ncid in " & _
                        "(34688,34689,34690,34693,34694,34695,34698,24824, " & _
                        "34699,34703,34704,34705,34708,34709,34710,34713,34714,34715, " & _
                        "34718,34719,34720,34721,1751434064) and " & _
                        "(o.observation_seq = c.observation_seq and " & _
                        "(o.event_id = c.event_id and " & _
                        "(c.observation_ncid = 204771 and " & _
                        "(c.event_id = e.event_id and " & _
                        "(e.event_subtype_ncid = 33971 and " & _
                        "(e.event_type_ncid = 204764 and " & _
                        "(e.unit_number = id.unit_number)))))))))) and "
    '-- Check if patient entered in system before 'TO' date
    sStmt = sStmt & _
            "(exists " & _
                "(select ma.unit_number " & _
                " from mmi_audit ma " & _
                " where " & _
                        GetOracleEndDate(moReportItem.DateTo) & " >= TX_TIME and " & _
                        "ma.unit_number = id.unit_number) and "
'    '-- Check for valid facility
'    sStmt = sStmt & _
'            "exists " & _
'                "(select ms.unit_number " & _
'                " from mmi_site ms " & _
'                " where " & _
'                        moReportItem.CriteriaSel1.Item(1).NCID & " = facility_cid and " & _
'                        "ms.unit_number = id.unit_number) and "
    '-- Check if patient >= 17 During reporting period
    sStmt = sStmt & _
            "(((" & GetOracleEndDate(moReportItem.DateTo) & " - id.birth_date)/365) >= 17.0 "

    '-- Check if patient has records at specified MTF or is patieint of specified PCM
    sStmt = sStmt & _
            " and (" & sWhereCriteria & "))))"

    goDBSqlOps.Execute sStmt
    
    LNumerator = goDBSqlOps.Value("Numerator")

'-- Set Report Field (Formula Field) Information
'---------------------------------------------------------------------------
 HeaderColl.Add CrystalQuote(CStr(LNumerator)), "Numerator"
 HeaderColl.Add CrystalQuote(CStr(LNumerator)), "Denominator"
 HeaderColl.Add CrystalQuote(CStr(LNumerator)), "Facility"

DataGetAlcoholQueryforPotentialHeavyUse = True
GoTo CleanUp
ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "ReportOps.DataGetAlcoholQueryforPotentialHeavyUse", App.Title, vbExclamation)
CleanUp:
    Exit Function
End Function
Private Function DataGetTobaccoUseScreening(HeaderColl As Collection) As Boolean

If Not DebugMode Then On Error GoTo ErrHandler
DataGetTobaccoUseScreening = False

Dim LNumerator As Long
Dim LDenominator As Long
Dim sStmt As String
Dim sWhereCriteria As String
Dim sFromCriteria As String

If Not GetMTForPCM_SQL(sFromCriteria, sWhereCriteria) Then GoTo CleanUp

'-- Get Denominator
'---------------------------------------------------------------------------
    sStmt = "select count(*) Denominator " & _
            "from mmi_id Id " & sFromCriteria & _
            "Where "
    '-- Check if patient entered in system before 'TO' date
    sStmt = sStmt & _
            "(exists " & _
                "(select ma.unit_number " & _
                " from mmi_audit ma " & _
                " where " & _
                        GetOracleEndDate(moReportItem.DateTo) & " >= TX_TIME and " & _
                        "ma.unit_number = id.unit_number) and "
    
    '-- Check if patient >= 17 During reporting period
    sStmt = sStmt & _
            "((((" & GetOracleEndDate(moReportItem.DateTo) & " - id.birth_date)/365) >= 17.0) "

    '-- Check if patient has records at specified MTF or is patieint of specified PCM
    sStmt = sStmt & _
            " and " & sWhereCriteria & "))"

    
    goDBSqlOps.Execute sStmt
    
    LDenominator = goDBSqlOps.Value("Denominator")

'-- Get Numerator
'---------------------------------------------------------------------------
    sStmt = "select count(*) Numerator " & _
            "from mmi_id id " & sFromCriteria & _
            "Where "
    '-- Check for Coupler Event where user responded 'Yes' or 'No' Finding 691,239
    sStmt = sStmt & _
            "((exists " & _
                "(select e.event_id " & _
                " from observation_value o, clinical_observation c, event e " & _
                " where " & _
                        "o.value_ncid in (34599, 34599) and " & _
                        "o.observation_seq = c.observation_seq and " & _
                        "o.event_id = c.event_id and " & _
                        "c.observation_ncid = 204771 and " & _
                        "c.event_id = e.event_id and " & _
                        GetOracleEndDate(moReportItem.DateTo) & " >= EVENT_START_GMTIME and " & _
                        "e.event_subtype_ncid = 33971 and e.event_type_ncid = 204764 and " & _
                        "e.unit_number = id.unit_number) or "
    '-- Check for Vitals indicating patient smokes
    sStmt = sStmt & _
            "exists " & _
                "(select e.event_id " & _
                " from observation_value o, clinical_observation c, event e " & _
                " where " & _
                        "o.value_ncid = 207675 and " & _
                        "o.observation_seq = c.observation_seq and " & _
                        "o.event_id = c.event_id and " & _
                        "c.observation_ncid = 162045 and " & _
                        "c.event_id = e.event_id and " & _
                        GetOracleEndDate(moReportItem.DateTo) & " >= EVENT_START_GMTIME and " & _
                        "e.event_subtype_ncid = 207675 and e.event_type_ncid = 162156 and " & _
                        "e.unit_number = id.unit_number)) and "
    '-- Check if patient entered in system before 'TO' date
    sStmt = sStmt & _
            "(exists " & _
                "(select ma.unit_number " & _
                " from mmi_audit ma " & _
                " where " & _
                        GetOracleEndDate(moReportItem.DateTo) & " >= TX_TIME and " & _
                        "ma.unit_number = id.unit_number) and "

    '-- Check if patient >= 17 During reporting period
    sStmt = sStmt & _
            "((((" & GetOracleEndDate(moReportItem.DateTo) & " - id.birth_date)/365) >= 17.0) "
    
    '-- Check if patient has records at specified MTF or is patieint of specified PCM
    sStmt = sStmt & _
            " and " & sWhereCriteria & ")))"
    
    
    goDBSqlOps.Execute sStmt
    
    LNumerator = goDBSqlOps.Value("Numerator")

'-- Set Report Field (Formual Field) Information
'---------------------------------------------------------------------------
     HeaderColl.Add CrystalQuote(CStr(LNumerator)), "Numerator"
     HeaderColl.Add CrystalQuote(CStr(LDenominator)), "Denominator"
     HeaderColl.Add CrystalQuote(moReportItem.CriteriaSel1.Item(1).Name), "Facility"

DataGetTobaccoUseScreening = True
GoTo CleanUp
ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "ReportOps.DataGetTobaccoUseScreening", App.Title, vbExclamation)
CleanUp:
    Exit Function
End Function

Private Function DataGetSafeSexCounselingQueried_AdvisedofSTDs(HeaderColl As Collection) As Boolean

If Not DebugMode Then On Error GoTo ErrHandler
DataGetSafeSexCounselingQueried_AdvisedofSTDs = False

Dim LNumerator As Long
Dim LDenominator As Long
Dim sStmt As String
Dim sWhereCriteria As String
Dim sFromCriteria As String


If Not GetMTForPCM_SQL(sFromCriteria, sWhereCriteria) Then GoTo CleanUp

'-- Get Denominator
'---------------------------------------------------------------------------
    sStmt = "select count(*) Denominator " & _
            "from mmi_id Id " & sFromCriteria & _
            "Where "
    '-- Check if patient entered in system before 'TO' date
    sStmt = sStmt & _
            "(exists " & _
                "(select ma.unit_number " & _
                " from mmi_audit ma " & _
                " where " & _
                        GetOracleEndDate(moReportItem.DateTo) & " >= TX_TIME and " & _
                        "ma.unit_number = id.unit_number) and "
'    '-- Check for valid facility
'    sStmt = sStmt & _
'            "exists " & _
'                "(select ms.unit_number " & _
'                " from mmi_site ms " & _
'                " where " & _
'                        moReportItem.CriteriaSel1.Item(1).NCID & " = facility_cid and " & _
'                        "ms.unit_number = id.unit_number) and "
    '-- Check if patient >= 17 During reporting period
    sStmt = sStmt & _
            "(((" & GetOracleEndDate(moReportItem.DateTo) & " - id.birth_date)/365) >= 17.0 "

    '-- Check if patient has records at specified MTF or is patieint of specified PCM
    sStmt = sStmt & _
            " and (" & sWhereCriteria & ")))"

    goDBSqlOps.Execute sStmt
    
    LDenominator = goDBSqlOps.Value("Denominator")

'-- Get Numerator
'---------------------------------------------------------------------------
    sStmt = "select count(*) Numerator " & _
            "from mmi_id id " & sFromCriteria & _
            "Where "
    '-- Check for Coupler Event where user responded 'Yes' Finding 518
    sStmt = sStmt & _
            "((exists " & _
                "(select e.event_id " & _
                " from clinical_observation c, observation_value o, event e " & _
                " where " & _
                        "(o.value_ncid in (34883) and " & _
                        "(o.observation_seq = c.observation_seq and " & _
                        "(o.event_id = c.event_id and " & _
                        "(c.observation_ncid = 204771 and " & _
                        "(c.event_id = e.event_id and (" & _
                        GetOracleEndDate(moReportItem.DateTo) & " >= EVENT_START_GMTIME and " & _
                        "(e.event_subtype_ncid = 33971 and " & _
                        "(e.event_type_ncid = 204764 and " & _
                        "(e.unit_number = id.unit_number)))))))))) or "
    '-- Check for Wellness Event - STD Prevention
    
'                            "o.value_ncid = 207694 and " & _
'                        "o.observation_seq = c.observation_seq and " & _
'                        "o.event_id = c.event_id and " & _
'                        "c.observation_ncid = 162045 and " &
    sStmt = sStmt & _
            "exists " & _
                "(select e.event_id " & _
                " from clinical_event_header c, event e " & _
                " where " & _
                        "(c.status_ncid = 162068 and " & _
                        "(c.event_id = e.event_id and (" & _
                        GetOracleEndDate(moReportItem.DateTo) & " >= e.EVENT_START_GMTIME and " & _
                        "(e.event_subtype_ncid = 207694 and " & _
                        "(e.event_type_ncid = 162156 and " & _
                        "(e.unit_number = id.unit_number)))))))) and "

'    '-- Check for valid facility
'    sStmt = sStmt & _
'            "exists " & _
'                "(select ms.unit_number " & _
'                " from mmi_site ms " & _
'                " where " & _
'                        moReportItem.CriteriaSel1.Item(1).NCID & " = facility_cid and " & _
'                        "ms.unit_number = id.unit_number) and "
    '-- Check if patient was entered before the 'TO' date
    sStmt = sStmt & _
            "(exists " & _
                "(select ma.unit_number " & _
                " from mmi_audit ma " & _
                " where " & _
                        GetOracleEndDate(moReportItem.DateTo) & " >= TX_TIME and " & _
                        "ma.unit_number = id.unit_number) and "
    '-- Check if patient was >= 17 years before 'TO' date
    sStmt = sStmt & _
            "(((" & GetOracleEndDate(moReportItem.DateTo) & " - id.birth_date)/365) >= 17.0 "

    '-- Check if patient has records at specified MTF or is patieint of specified PCM
    sStmt = sStmt & _
            " and (" & sWhereCriteria & "))))"

    goDBSqlOps.Execute sStmt
    
    LNumerator = goDBSqlOps.Value("Numerator")

'-- Set Report Field (Formual Field) Information
'---------------------------------------------------------------------------

 HeaderColl.Add CrystalQuote(CStr(LNumerator)), "Numerator"
 HeaderColl.Add CrystalQuote(CStr(LDenominator)), "Denominator"
 HeaderColl.Add CrystalQuote(moReportItem.CriteriaSel1.Item(1).Name), "Facility"

DataGetSafeSexCounselingQueried_AdvisedofSTDs = True
GoTo CleanUp
ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "ReportOps.DataGetSafeSexCounselingQueried_AdvisedofSTDs", App.Title, vbExclamation)
CleanUp:
    Exit Function
End Function




Friend Function GetReportInformation() As Boolean

    If Not DebugMode Then On Error GoTo ErrHandler
    
    GetReportInformation = True
    
    Dim rsInput As ADODB.Recordset
    Dim rsAdditionalInformation As ADODB.Recordset
    Dim oReport_Sort_Item As Report_Sort_Item
    Dim sStmt As String
    Dim icount As Integer
    Dim sName As String
    Dim oReportInfo As ReportInformation
   
    Set rsInput = New ADODB.Recordset
    rsInput.CursorLocation = adUseClient
    rsInput.CursorType = adOpenStatic
    rsInput.ActiveConnection = poConnection
        
    If Not gobjshared.UseSQLServer Then
        sStmt = "select * from Reports order by longname "
    Else
        sStmt = "select * from Reports where ITTReports = True order by longname "
    End If
    
    rsInput.Open sStmt, , , adLockBatchOptimistic, adCmdText
    
    Set poCollRepInfo = Nothing
    Set poCollRepInfo = New Collection
    
    Do While Not rsInput.EOF
    
        If rsInput.Fields("IsDisplayed") = True Then
            Set oReportInfo = New ReportInformation
            
            Set oReportInfo.Criteria1 = New Collection
            Set oReportInfo.Criteria2 = New Collection
            Set oReportInfo.Criteria3 = New Collection
            Set oReportInfo.Groupby = New Collection
            Set oReportInfo.SortItems = New Collection
            
            oReportInfo.ADate = rsInput.Fields("ADate")
        
            oReportInfo.ReportFileName = rsInput.Fields("ReportFileName")
            
            oReportInfo.LongName = rsInput.Fields("LongName")
            
            oReportInfo.Header = rsInput.Fields("Header")
            
            On Error Resume Next
            oReportInfo.AppendCriteriaToHeader = rsInput.Fields("AppendCriteriaToHeader").Value
            oReportInfo.CustomDateFormat = rsInput.Fields("CustomDateFormat").Value
            oReportInfo.AlternativeHeader = rsInput.Fields("AlternativeHeader").Value
            oReportInfo.DefaultSortColumn = rsInput.Fields("DefaultSortColumn").Value
            If Not DebugMode Then On Error GoTo ErrHandler
            
            oReportInfo.DateFrom = rsInput.Fields("DateFrom")
            
            oReportInfo.DateTo = rsInput.Fields("DateTo")
            
            oReportInfo.ReportCategory = rsInput.Fields("ReportCategory")
            
            oReportInfo.IsGrouping = rsInput.Fields("IsGroup")
            oReportInfo.IsCriteria1 = rsInput.Fields("IsCriteria1")
            oReportInfo.IsCriteria2 = rsInput.Fields("IsCriteria2")
            oReportInfo.IsCriteria3 = rsInput.Fields("IsCriteria3")
            
            oReportInfo.IsCriteria1Required = rsInput.Fields("Criteria1Required")
            oReportInfo.IsCriteria2Required = rsInput.Fields("Criteria2Required")
            oReportInfo.IsCriteria3Required = rsInput.Fields("Criteria3Required")
            
            If oReportInfo.IsCriteria1 Then
                If rsInput.Fields("Criteria1_1") <> 0 Then
                    oReportInfo.Criteria1.Add CInt(rsInput.Fields("Criteria1_1"))
                End If
                
                If rsInput.Fields("Criteria1_2") <> 0 Then
                    oReportInfo.Criteria1.Add CInt(rsInput.Fields("Criteria1_2"))
                End If
                
                If rsInput.Fields("Criteria1_3") <> 0 Then
                    oReportInfo.Criteria1.Add CInt(rsInput.Fields("Criteria1_3"))
                End If
                    
                If rsInput.Fields("Criteria1_4") <> 0 Then
                    oReportInfo.Criteria1.Add CInt(rsInput.Fields("Criteria1_4"))
                End If
            End If
            
            If oReportInfo.IsCriteria2 Then
                If rsInput.Fields("Criteria2_1") <> 0 Then
                    oReportInfo.Criteria2.Add CInt(rsInput.Fields("Criteria2_1"))
                End If
                
                If rsInput.Fields("Criteria2_2") <> 0 Then
                    oReportInfo.Criteria2.Add CInt(rsInput.Fields("Criteria2_2"))
                End If
                
                If rsInput.Fields("Criteria2_3") <> 0 Then
                    oReportInfo.Criteria2.Add CInt(rsInput.Fields("Criteria2_3"))
                End If
                    
                If rsInput.Fields("Criteria2_4") <> 0 Then
                    oReportInfo.Criteria2.Add CInt(rsInput.Fields("Criteria2_4"))
                End If
                
                If rsInput.Fields("Criteria2_5") <> 0 Then
                    oReportInfo.Criteria2.Add CInt(rsInput.Fields("Criteria2_5"))
                End If
                
                If rsInput.Fields("Criteria2_6") <> 0 Then
                    oReportInfo.Criteria2.Add CInt(rsInput.Fields("Criteria2_6"))
                End If
                    
                If rsInput.Fields("Criteria2_7") <> 0 Then
                    oReportInfo.Criteria2.Add CInt(rsInput.Fields("Criteria2_7"))
                End If
            End If
            
            If oReportInfo.IsCriteria3 Then
                If rsInput.Fields("Criteria3_1") <> 0 Then
                    oReportInfo.Criteria3.Add CInt(rsInput.Fields("Criteria3_1"))
                End If
                    
                If rsInput.Fields("Criteria3_2") <> 0 Then
                    oReportInfo.Criteria3.Add CInt(rsInput.Fields("Criteria3_2"))
                End If
            End If
            
            If rsInput.Fields("GroupBy1") <> 0 Then
                oReportInfo.Groupby.Add CInt(rsInput.Fields("GroupBy1"))
            End If
                
            If rsInput.Fields("GroupBy2") <> 0 Then
                oReportInfo.Groupby.Add CInt(rsInput.Fields("GroupBy2"))
            End If
        
            If rsInput.Fields("GroupBy3") <> 0 Then
                oReportInfo.Groupby.Add CInt(rsInput.Fields("GroupBy3"))
            End If
            
            oReportInfo.Description = rsInput.Fields("Description")
            oReportInfo.ClassName = rsInput.Fields("ClassName") & ""
            oReportInfo.Switch = rsInput.Fields("Switch") & ""
            
            oReportInfo.ReportTypeId = rsInput.Fields("ReportId")
            oReportInfo.SortReport = rsInput.Fields("SortResults")
            If oReportInfo.SortReport Then
              Set rsAdditionalInformation = New ADODB.Recordset
              rsAdditionalInformation.CursorLocation = adUseClient
              rsAdditionalInformation.CursorType = adOpenStatic
              rsAdditionalInformation.ActiveConnection = poConnection
        
              sStmt = "select * from Report_Sort_Items where ReportId = " & CStr(rsInput("ReportId"))
              rsAdditionalInformation.Open sStmt, , , adLockBatchOptimistic, adCmdText
              If Not (rsAdditionalInformation.BOF And rsAdditionalInformation.EOF) Then
                Do Until rsAdditionalInformation.EOF
                    Set oReport_Sort_Item = New Report_Sort_Item
                    oReport_Sort_Item.DisplayValue = rsAdditionalInformation("DisplayValue")
                    oReport_Sort_Item.ReportId = rsAdditionalInformation("ReportId")
                    
                    oReport_Sort_Item.SortField = rsAdditionalInformation("SortField")
                    oReportInfo.SortItems.Add oReport_Sort_Item
                    rsAdditionalInformation.MoveNext
                Loop
              End If
              'hwcc 1/5/2004 - plug a leak here
              Set rsAdditionalInformation = Nothing
            End If
            
            'hwcc 1/5/2004 - Rounds Report requirements  DiagnosisRep
            oReportInfo.EnableSecondaryFunction = False
            'hwcc 1/12/2003 - Add customized->Diagnosis report to this list
            If (oReportInfo.ReportTypeId = PrimaryDiagnosisRep) Or _
               (oReportInfo.ReportTypeId = DiagnosisRep) Or _
               (oReportInfo.ReportTypeId = ProcedureRep) Or _
               (oReportInfo.ReportTypeId = AppointmentRep) Or _
               (oReportInfo.ReportTypeId = ADSDiagnosis) Or _
               (oReportInfo.ReportTypeId = ADSEncountersAppointmentStatus) Then
                If Not cwobjCWComm.CmdLineSwitch(RNDSAPPTSW) = "" Then     'chk RNDS switch
                    oReportInfo.EnableSecondaryFunction = rsInput.Fields("EnableSecondaryFunction")
                End If
            Else
                oReportInfo.EnableSecondaryFunction = rsInput.Fields("EnableSecondaryFunction")
            End If
            
            oReportInfo.SecondaryFunctionTtitle = rsInput.Fields("SecondaryFunctionTitlte") & ""
            If Not IsNull(rsInput.Fields("MaximumDuration")) Then
                oReportInfo.MaximumDuration = rsInput.Fields("MaximumDuration")
            End If
            
            If oReportInfo.Switch = "" Or (Not cwobjCWComm.CmdLineSwitch(oReportInfo.Switch) = "") Then
               poCollRepInfo.Add oReportInfo, CStr(oReportInfo.ReportTypeId)
            End If
            
            'hwcc 1/5/2004 - Save this line for debug
            ' Call oReportInfo.DebugPrintReportInformation
            
            Set oReportInfo = Nothing
        End If
        rsInput.MoveNext
    Loop

    rsInput.Close

    GetReportInformation = True

    GoTo CleanUp

ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "ReportOps.GetReportInformation", "frmReports", vbExclamation)
    'Resume Next
CleanUp:
    Set oReport_Sort_Item = Nothing
    Set oReportInfo = Nothing
    Set rsAdditionalInformation = Nothing
    Set rsInput = Nothing
    Exit Function
End Function

Friend Function GetOracleStartDate(vDate As Variant) As String

If Not DebugMode Then On Error GoTo ErrHandler

Dim sDate As String

    sDate = " TO_DATE('" & Format(vDate, "yyyy/mm/dd") & " 00:00" & "','YYYY/MM/DD HH24:MI')"
    GetOracleStartDate = sDate

GoTo CleanUp

ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "ReportOps.GetOracleStartDate", "ReportOps", vbExclamation)
CleanUp:
    Exit Function
End Function
Friend Function GetOracleEndDate(vDate As Variant) As String

If Not DebugMode Then On Error GoTo ErrHandler

Dim sDate As String

sDate = " TO_DATE('" & Format(vDate, "yyyy/mm/dd") & " 23:59" & "','YYYY/MM/DD HH24:MI')"
GetOracleEndDate = sDate
GoTo CleanUp

ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "ReportOps.GetOracleEndDate", "ReportOps", vbExclamation)
CleanUp:
    Exit Function
End Function

Private Sub Delay()

If Not DebugMode Then On Error GoTo ErrHandler

Dim PauseTime As Variant
Dim Start As Variant
Dim Finish  As Variant
Dim TotalTime  As Variant

   PauseTime = 1   ' Set duration
   Start = Timer   ' Set start time.
   Do While Timer < Start + PauseTime
      DoEvents   ' Yield to other processes.
   Loop

GoTo CleanUp
ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "ReportsOps.DisplayReport", "ReportsOps", vbExclamation)
CleanUp:
    Exit Sub
End Sub

