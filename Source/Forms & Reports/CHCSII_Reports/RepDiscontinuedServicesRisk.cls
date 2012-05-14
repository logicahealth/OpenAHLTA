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
Attribute VB_Name = "RepDisconServicesRisk"
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
    If Not GetDiscontinuedServices(oReportItem, oRec) Then GoTo CleanUp
    moCollRecSet.Add oRec, "services"

'-- Set Formulas
'-----------------------------------------------------------------------------
    'Set oCollFormulas = Nothing
    'Set oCollFormulas = New Collection
    'oCollFormulas.Add "'" & Format(Now, gsDefaultFormat & "  HHMM") & "'", "TodaysDateTime"
    

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

Private Function GetDiscontinuedServices(oReportItem As ReportItem, _
                                         oRec As ADODB.Recordset) As Boolean
On Error GoTo ErrHandler

GetDiscontinuedServices = False

Dim sStmt As String
Dim sFromSQL As String
Dim sWhereSQL As String

If Not GetMTForPCM_SQL(oReportItem, sFromSQL, sWhereSQL) Then GoTo CleanUp

'-- Get Services Due
'----------------------------------------------------------------------------------------
sStmt = "select " & vbNewLine & _
        "e.unit_number," & vbNewLine & _
        "Trim(Representation) ServiceDesc, " & vbNewLine & _
        "To_Char(event_start_gmtime,'YYYY MON DD') as DateDiscontinued " & vbNewLine

sStmt = sStmt & _
        "From " & vbNewLine & _
            "rsform_context rc, rsform r, " & vbNewLine & _
            "clinical_event_header c, event e, mmi_id id " & vbNewLine & _
            sFromSQL

sStmt = sStmt & _
        "where " & vbNewLine & _
            "rc.context_ncid = 2000 and " & vbNewLine & _
            "rc.rsform_id = r.rsform_id and " & vbNewLine & _
            "r.ncid = e.event_subtype_ncid and " & vbNewLine & _
            "c.status_ncid = 162073 and " & vbNewLine & _
            "c.event_id = e.event_id and " & vbNewLine & _
            "e.event_type_ncid = 162158 and " & vbNewLine & _
            "e.unit_number  = id.unit_number " & vbNewLine

'    '-- Check if patient has records at specified MTF or is patieint of specified PCM
    sStmt = sStmt & _
            " and (" & sWhereSQL & ")"

Set oRec = goDBDAS.OpenRecordset(sStmt)

GetDiscontinuedServices = True
GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "FormCheckOut.GetDiscontinuedServices", "CHCSII_Forms", vbExclamation
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

        sStmt = sStmt & _
            "where " & _
                "tUIC.type_ncid(+)=14501722 and " & _
                "tUIC.unit_number(+)=id.unit_number and " & _
                "sMN.unit_number(+)=sId.unit_number and " & _
                "sId.ssn(+)=tSSN.id_value and " & _
                "tFMP.id_type_ncid(+)=204708 and " & _
                "tFMP.unit_number(+)=id.unit_number and " & _
                "tSSN.id_type_ncid(+)=186334 and " & _
                "tSSN.unit_number(+)=id.unit_number and " & _
                "tPC.type_ncid(+)=14501721 and " & _
                "tPC.unit_number(+)=id.unit_number and " & _
                "tRank.type_ncid(+)=1015 and " & _
                "tRank.unit_number(+)=id.unit_number and " & _
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

