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
Attribute VB_Name = "RepForms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private moCollRecSet As Collection
Private moCollFormulas As Collection

Public Function BuildReport(sReportName As String, _
                            oReportItem As ReportItem, _
                            oReport As CRAXDRT.Report, _
                            oCollFormulas As Collection) As Boolean

On Error GoTo ErrHandler
 
BuildReport = False

Dim oCrystal As Crystal
Dim oRec As ADODB.Recordset
Dim sSponsorUnitNumber As String
Dim sUnitNumber As String
Dim sPatientUnitNumber As String
DoEvents

If oReportItem.Criteria1 = Patient Then
    sPatientUnitNumber = oReportItem.CriteriaSel1.Item(1).UnitNumber
End If

'-- Get Data for the Report
'---------------------------------------------------------------------------------------
Select Case oReportItem.Report
    Case DD2766
        Dim oDD2766 As FormDD2766
        Set oDD2766 = New FormDD2766
        If Not oDD2766.BuildDD2766(sPatientUnitNumber, oReport) Then GoTo CleanUp
        'If Not DataGetDD2766() Then GoTo CleanUp
    Case DD2813
        Dim oDD2813 As Formdd2813
        Set oDD2813 = New Formdd2813
        If Not oDD2813.BuildDD2813(sPatientUnitNumber, oReport) Then GoTo CleanUp
        'If Not DataGetDD2813() Then GoTo CleanUp
    Case DD689
        Dim oDD689 As FormDD689
        Set oDD689 = New FormDD689
        If Not oDD689.BuildDD689(sPatientUnitNumber, oReport) Then GoTo CleanUp
        'If Not DataGetDD689() Then GoTo CleanUp
    Case SF522
        Dim oSF522 As FormSF522
        Set oSF522 = New FormSF522
        If Not oSF522.BuildSF522(sPatientUnitNumber, oReport) Then GoTo CleanUp
        'If Not DataGetSF522() Then GoTo CleanUp
    Case DD741
        Dim oDD741 As FormDD741
        Set oDD741 = New FormDD741
        If Not oDD741.BuildDD741(sPatientUnitNumber, oReport) Then GoTo CleanUp
        'If Not DataGetDD741() Then GoTo CleanUp
''    Case DD771
''        If Not DataGetDD771() Then GoTo CleanUp
''    '-- CPG Forms
''    Case CPG_POC
''        If Not DataGetCPGPOC() Then GoTo CleanUp
    Case Else
End Select


'--------------------------------

'-- Get Data
'----------------------------------------------------------------------------
''    Set moCollRecSet = Nothing
''    Set moCollRecSet = New Collection
''
''    '-- Get Candidates
''    '----------------------------------------------------------
''    If Not GetCandidates(oReportItem, oRec) Then GoTo CleanUp
''    moCollRecSet.Add oRec, "candidates"
''
'''-- Create Report Object
'''-----------------------------------------------------------------------------
''    If Not CreateReport(sReportName & ".rpt", oReport, oCollFormulas, moCollRecSet) Then GoTo CleanUp

BuildReport = True

GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepRegistryCandidates.BuildReport", "CHCSII_Reports", vbExclamation
CleanUp:
    Set oCrystal = Nothing
    Set oRec = Nothing
    
End Function

Private Function GetCandidates(oReportItem As ReportItem, _
                               oRec As ADODB.Recordset) As Boolean
On Error GoTo ErrHandler

GetCandidates = False

Dim sStmt As String
Dim sFromSQL As String
Dim sWhereSQL As String

    'If Not GetMTForPCM_SQL(oReportItem, sFromSQL, sWhereSQL) Then GoTo CleanUp

'-- Select
    sStmt = "Select distinct " & vbNewLine & _
                "Trim(mn.last_name ||', ' || mn.first_name || ' ' || " & vbNewLine & _
                "mn.middle_name) as Full_Name," & vbNewLine & _
                "fs.id_value as FMP_SPONSOR_SSN, " & vbNewLine & _
                "To_Char(id.birth_date,'DD MON YYYY') as DOB_FORMATTED, " & vbNewLine & _
                "p.NAME as PCM_Name " & vbNewLine

'-- From
    sStmt = sStmt & _
            "From " & vbNewLine & _
                "mmi_generic_id fs, " & vbNewLine & _
                "mmi_generic_id pm, " & vbNewLine & _
                "provider p, " & vbNewLine & _
                "mmi_name mn, " & vbNewLine & _
                "mmi_site ms, " & vbNewLine & _
                "mmi_id id, " & vbNewLine & _
                "event e, " & vbNewLine & _
                "clinical_observation co, " & vbNewLine & _
                "observation_value ov " & vbNewLine

'-- Where
    sStmt = sStmt & _
        "Where " & vbNewLine & _
            "fs.id_type_ncid = 14511309 and " & vbNewLine & _
            "fs.unit_number = ms.unit_number and " & vbNewLine & _
            "p.pcm_code = pm.id_value and " & vbNewLine & _
            "pm.id_type_ncid = 14501724 and " & vbNewLine & _
            "pm.unit_number = ms.unit_number and " & vbNewLine & _
            "mn.name_type_cid = 187 and " & vbNewLine & _
            "mn.unit_number = ms.unit_number and " & vbNewLine & _
            "not exists " & vbNewLine & _
            "(select 'x' from cpgregistry c where c.unit_number = ms.unit_number) and " & vbNewLine & _
            "ms.facility_cid  = " & oReportItem.CriteriaSel1.Item(1).NCID & " and " & vbNewLine & _
            "ms.unit_number = id.unit_number and " & vbNewLine & _
            "id.unit_number = e.unit_number and " & vbNewLine & _
            "e.event_id = co.event_id and " & vbNewLine & _
            "co.observation_seq = ov.observation_seq and " & vbNewLine & _
            "co.event_id = ov.event_id and " & vbNewLine & _
            "ov.value_ncid in (1002435,1002437)"
    
    '-- Check if patient has records at specified MTF or is patient of specified PCM
'    sStmt = sStmt & _
'            " and " & sWhereSQL

Set oRec = goDBDAS.OpenRecordset(sStmt)

GetCandidates = True
GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepRegistryCandidates.GetCandidates", "CHCSII_Reports", vbExclamation
CleanUp:
    Exit Function
End Function

Private Sub Class_Initialize()
    '--- SCR-ARM-16148 MSolano 9/10/01; COMMON TO MULTIPLE ARM-SCRs
    '--- Init routine should be called from first top level creatable coclass
    '--- It can be called multiple times, but will actually run only once.
    '--- Instrument all creatable classes in case dll's multiple entry
    '--- points are actually used
     Call ARM_InitAllTransactionsForCHCSIIReports
End Sub

