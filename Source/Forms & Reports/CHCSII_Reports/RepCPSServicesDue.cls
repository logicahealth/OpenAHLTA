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
Attribute VB_Name = "RepCPSServicesDue"
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
DoEvents

'-- Get Data
'----------------------------------------------------------------------------
    Set moCollRecSet = Nothing
    Set moCollRecSet = New Collection
    
    '-- Patient Demographics
    '----------------------------------------------------------
    If Not GetPatientDemographics(oReportItem, oRec) Then GoTo CleanUp
    moCollRecSet.Add oRec, "demographics"

    '-- Get Services Due
    '----------------------------------------------------------
    If Not GetServicesDue(oReportItem, oRec) Then GoTo CleanUp
    FilterActive oReportItem, oRec
    moCollRecSet.Add oRec, "services"

'-- Set Formulas
'-----------------------------------------------------------------------------
'    Set moCollFormulas = Nothing
'    Set moCollFormulas = New Collection
'    moCollFormulas.Add "'" & Format(Now, gsDefaultFormat & "  HHMM") & "'", "TodaysDateTime"
'

'-- Create Report Object
'-----------------------------------------------------------------------------
    Set oCrystal = New Crystal
    Set oReport = Nothing
    If Not oCrystal.ReportOpen(sReportName & ".rpt", oReport) Then GoTo CleanUp
    If Not oCrystal.AssignRecordSets(moCollRecSet, oReport) Then GoTo CleanUp
    If Not oCrystal.AssignFormulas(oCollFormulas, oReport) Then GoTo CleanUp

BuildReport = True

GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "CPSServicesDue.BuildReport", "CHCSII_Reports", vbExclamation
    
CleanUp:
    Set oCrystal = Nothing
    Set oRec = Nothing
    
End Function

Private Function FilterActive(oReportItem As ReportItem, pDueItems As ADODB.Recordset) As Boolean

    Dim bActiveDuty As Boolean
    Dim sStmt As String
    Dim oRec2 As ADODB.Recordset
    
    On Error GoTo ErrHandler
    FilterActive = False
        
    If Not oReportItem.Criteria1 = Patient Then GoTo CleanUp
    
    sStmt = "SELECT * from mmi_genericdata " & _
            "WHERE UNIT_NUMBER = " & oReportItem.CriteriaSel1.Item(1).UnitNumber & " AND " & _
            "TYPE_NCID = 14501721 " & _
                    " AND VALUE_NCID IN (119033, 119034, 119035, 119068, 119069, 119070, 119122, 119123, 119124, 119148, 119149, 119150)"
    Set oRec2 = Nothing
    Set oRec2 = goDBDAS.OpenRecordset(sStmt)
    
    bActiveDuty = False
    If Not (oRec2.BOF And oRec2.EOF) Then
       bActiveDuty = True
    End If
    
    If Not bActiveDuty And Not (pDueItems.BOF And pDueItems.EOF) Then
       pDueItems.MoveFirst
       Do Until pDueItems.EOF
           If Left(pDueItems("DESCRIPTION"), 3) = "(A)" Then
              pDueItems.Delete
           End If
           pDueItems.MoveNext
       Loop
    End If
    
    FilterActive = True
    GoTo CleanUp
    
ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "FormCheckOut.FilterActive", "CHCSII_Forms", vbExclamation
CleanUp:
    Set oRec2 = Nothing
    
End Function


Private Function GetServicesDue(oReportItem As ReportItem, _
                                oRec As ADODB.Recordset) As Boolean
On Error GoTo ErrHandler

GetServicesDue = False

Dim sStmt As String
Dim sFromSQL As String
Dim sWhereSQL As String

If Not GetMTForPCM_SQL(oReportItem, sFromSQL, sWhereSQL) Then GoTo CleanUp

'-- Get Services Due
'----------------------------------------------------------------------------------------
sStmt = "select " & _
        "e.unit_number, " & _
        "e.event_type_ncid, " & _
        "e.event_subtype_ncid, " & _
        "co.observation_ncid, " & _
        "to_char(ov.value_datetime,'DD MON YYYY') as EarliestRecDate, " & _
        "r.representation as status, " & _
        "r2.Representation As Description "

sStmt = sStmt & _
        "From " & _
        "rsform_context rc2, " & _
        "rsform r2, " & _
        "rsform_context rc, " & _
        "rsform r, " & _
        "observation_value ov, " & _
        "clinical_observation co, " & _
        "clinical_event_header c, event e, mmi_id id " & sFromSQL

sStmt = sStmt & _
        "Where " & _
        "rc2.context_ncid = 2000 and " & _
        "rc2.rsform_id = r2.rsform_id and " & _
        "r2.ncid = e.event_subtype_ncid and " & _
        "rc.context_ncid = 2000 and " & _
        "rc.rsform_id = r.rsform_id and " & _
        "r.ncid = c.status_ncid and " & _
        "ov.OBSERVATION_SEQ (+) = co.OBSERVATION_SEQ and " & _
        "ov.event_id (+) = co.event_id and " & _
        "co.observation_ncid = 162168 and " & _
        "co.event_id (+) = e.event_id and " & _
        "(((c.status_ncid = 30475 and " & _
        "e.event_type_ncid = 162158) or " & _
        "(c.status_ncid = 1024 and " & _
        "e.event_type_ncid = 162156)) and " & _
        "(c.event_id = e.event_id and " & _
        "(e.event_type_ncid = 162156 and (e.unit_number = id.unit_number "

'    '-- Check if patient has records at specified MTF or is patieint of specified PCM
    sStmt = sStmt & _
            " and (" & sWhereSQL & ")))))"

Set oRec = goDBDAS.OpenRecordset(sStmt)

GetServicesDue = True
GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "FormCheckOut.GetServicesDue", "CHCSII_Forms", vbExclamation
CleanUp:
    Exit Function
End Function

Private Function GetPatientDemographics(oReportItem As ReportItem, _
                                        oRec As ADODB.Recordset) As Boolean

On Error GoTo ErrHandler

GetPatientDemographics = False

Dim sStmt As String
Dim sFromSQL As String
Dim sWhereSQL As String
Dim oRec2 As Recordset


If Not GetMTForPCM_SQL(oReportItem, sFromSQL, sWhereSQL) Then GoTo CleanUp

'-- Get Patients
'----------------------------------------------------------------------------------------
    sStmt = _
        "Select " & _
            "id.unit_number, id.ssn as SSN, " & _
            "tUIC.value_text as UIC, " & _
            "To_Char(id.birth_date,'DD MON YYYY') as DOB_FORMATTED, " & _
            "Upper(id.sex) as Sex, " & _
            "Trim(mn.last_name ||', ' || mn.first_name || ' ' || " & _
            "mn.middle_name) as Full_Name," & _
            "Trim(tFMP.id_value) as FMP, " & _
            "Trim(tSSN.id_value) as SPONSOR_SSN, " & _
            "Decode(Upper(substr(tPC.value_text,1,1)), " & _
                    "'A','Army', " & _
                    "'B','NOAA (National Oceanic and Atmospheric Administration)', " & _
                    "'C','Coast Guard', " & _
                    "'F','Air Force','K','Civilian', " & _
                    "'M','Marines','N','Navy', " & _
                    "'p','Public Health Service', " & _
                    "'R','NATO Reciprocal Agreement', " & _
                    "'A','Army','Unknown') as Service, " & _
            "Decode(tPC.value_ncid,119033,'Y',119034,'Y',119035,'Y',119068,'Y',119069,'Y'," & _
                    "119070,'Y',119122,'Y',119123,'Y',119124,'Y',119148,'Y',119149,'Y',119150,'Y','N') as Active_Duty, " & _
            "Trim(sMN.last_name ||', ' || sMN.first_name || ' ' || " & _
            "sMN.middle_name) as Sponsor_Full_Name," & _
            "tRank.value_text as Rank "

        sStmt = sStmt & _
            "from " & _
                "mmi_genericdata tUIC, " & _
                "mmi_name sMN, " & _
                "mmi_id sId, " & _
                "mmi_generic_id tFMP, " & _
                "mmi_generic_id tSSN, " & _
                "mmi_genericdata tPC, " & _
                "mmi_genericdata tRank, " & _
                "mmi_name MN, " & _
                "mmi_id id " & _
                sFromSQL

        sStmt = sStmt & "where " & _
                "tUIC.type_ncid(+)=14501722 and " & _
                "tUIC.unit_number(+)=id.unit_number and " & _
                "smn.name_type_cid = 187 and sMN.unit_number=sId.unit_number and " & _
                "sId.ssn(+)=tSSN.id_value and " & _
                "tFMP.id_type_ncid(+)=204708 and " & _
                "tFMP.unit_number(+)=id.unit_number and " & _
                "tSSN.id_type_ncid(+)=186334 and " & _
                "tSSN.unit_number(+)=id.unit_number and " & _
                "tPC.type_ncid(+)=14501721 and " & _
                "tPC.unit_number(+)=id.unit_number and " & _
                "tRank.type_ncid(+)=1015 and " & _
                "tRank.unit_number(+)=id.unit_number and " & _
                "mn.name_type_cid = 187 and " & _
                "MN.unit_number(+)= id.unit_number and " & _
                sWhereSQL

Set oRec = Nothing

Set oRec = goDBDAS.OpenRecordset(sStmt)

GetPatientDemographics = True
GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "FormCheckOut.GetMeds", "CHCSII_Forms", vbExclamation
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


