VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RepTPC"
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
DoEvents

'-- Get Data
'----------------------------------------------------------------------------
    Set moCollRecSet = Nothing
    Set moCollRecSet = New Collection

    '-- Get Insurance Information
    '----------------------------------------------------------
    If Not GetInsuranceInformation(oReportItem, oRec) Then GoTo CleanUp
    moCollRecSet.Add oRec, "Insurance"

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

Private Function GetInsuranceInformation(oReportItem As ReportItem, _
                                         oRec As ADODB.Recordset) As Boolean
On Error GoTo ErrHandler

GetInsuranceInformation = False

Dim sStmt As String


'-- Get Insurance
'----------------------------------------------------------------------------------------
sStmt = "select " & vbNewLine & _
        "Trim(mnp.last_name ||', ' || mnp.first_name || ' ' || " & vbNewLine & _
        "mnp.middle_name) as Provider_Name," & vbNewLine & _
        "ma.street_1 || ' ' || ma.street_2 || ' ' || ma.city || ', '||  ma.state || ' ' || ma.post_code as Pat_address, " & vbNewLine & _
        "Trim(mn.last_name ||', ' || mn.first_name || ' ' || " & vbNewLine & _
        "mn.middle_name) as Full_Name," & vbNewLine & _
        "Trim(tFMP.id_value) as FMP, " & vbNewLine & _
        "Trim(tSSN.id_value) as SPONSOR_SSN, " & vbNewLine & _
        "id.ssn as SSN, " & vbNewLine & _
        "To_Char(id.birth_date,'DD MON YYYY') as DOB_FORMATTED, " & vbNewLine & _
        "DECODE(hp.area_code,null,null, '(' || hp.area_code ||') ') || DECODE(hp.local_number,null, null, " & vbNewLine & _
        "SUBSTR(hp.local_number, 1, 3) || '-' || SUBSTR(hp.local_number,4,4)) as Home_phone," & vbNewLine & _
        "SUBSTR(DECODE(wp.area_code, null, null, '(' || wp.area_code ||') ') || " & vbNewLine & _
        "DECODE(wp.local_number,null, null, SUBSTR(wp.local_number, 1, 3) || '-' || " & vbNewLine & _
        "SUBSTR(wp.local_number,4,4)) || DECODE(wp.internal_extension, null, null," & vbNewLine & _
        "' Ext.' || wp.internal_extension), 1, 22) as Work_phone," & vbNewLine & _
        "cl.Name as Clinic_Name, " & vbNewLine & _
        "cl.meprs_code, " & vbNewLine & _
        "cl.dmis_id_code, " & vbNewLine

sStmt = sStmt & _
        "TO_CHAR(APPT_DATE_TIME,'DD MON YYYY') as Appt_Date, " & vbNewLine & _
        "INSCOMPANYNAME, " & vbNewLine & _
        "INSCOMPANYPHONE, " & vbNewLine & _
        "Trim(ADDRESS1) || ' ' || Trim(ADDRESS2) || ' ' || Trim(ADDRESS3) AS Ins_Address, " & vbNewLine & _
        "INSIDNUMBER, " & vbNewLine & _
        "GroupName, " & vbNewLine & _
        "GroupNumber, " & vbNewLine & _
        "SUBSCRIBERNAME, " & vbNewLine & _
        "TO_CHAR(STARTDATE,'YYYY MON DD') as StartDate, " & vbNewLine & _
        "TO_CHAR(STOPDATE,'YYYY MON DD') as StopDate, " & vbNewLine & _
        "ENTRYPERSON, " & vbNewLine & _
        "Source, " & vbNewLine & _
        "Comments, pi.RELATIONTOSPONSOR " & vbNewLine

'RELATIONTOSPONSOR

sStmt = sStmt & _
        "From " & vbNewLine & _
            "clinic cl, " & vbNewLine & _
            "mmi_name mnp, hems_users hu, " & vbNewLine & _
            "mmi_address ma, " & vbNewLine & _
            "mmi_phone hp, " & vbNewLine & _
            "mmi_phone wp, " & vbNewLine & _
            "mmi_id sId, " & vbNewLine & _
            "mmi_generic_id tFMP, " & vbNewLine & _
            "mmi_generic_id tSSN, " & vbNewLine & _
            "mmi_name MN, " & vbNewLine & _
            "mmi_id id, " & vbNewLine & _
            "patient_insurance pi, " & vbNewLine & _
            "appointment a " & vbNewLine

sStmt = sStmt & _
        "Where " & vbNewLine & _
            "cl.ncid = a.clinic_ncid and " & vbNewLine & _
            "ma.unit_number (+) = id.unit_number and " & vbNewLine & _
            "mnp.unit_number (+) = hu.unit_number and " & vbNewLine & _
            "hu.user_ncid (+) = a.clinician_ncid and " & vbNewLine & _
            "wp.phone_type_cid (+) = 194 and " & vbNewLine & _
            "wp.unit_number (+) = id.unit_number and " & vbNewLine & _
            "hp.phone_type_cid (+) = 193 and " & vbNewLine & _
            "hp.unit_number (+) = id.unit_number and " & vbNewLine & _
            "sId.ssn(+)=tSSN.id_value and " & vbNewLine & _
            "tFMP.id_type_ncid(+)=204708 and " & vbNewLine & _
            "tFMP.unit_number(+)=id.unit_number and " & vbNewLine & _
            "tSSN.id_type_ncid(+)=186334 and " & vbNewLine & _
            "tSSN.unit_number(+)=id.unit_number and " & vbNewLine & _
            "mn.unit_number = id.unit_number and " & vbNewLine & _
            "id.unit_number = a.patient_unit_number and " & vbNewLine & _
            "pi.encounter_number = a.encounter_number " & vbNewLine

    If Not BuildWhereClauseforAppointments(oReportItem, sStmt) Then GoTo CleanUp

Set oRec = goDBDAS.OpenRecordset(sStmt)

GetInsuranceInformation = True
GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepTPC.GetInsuranceInformation", "CHCSII_Forms", vbExclamation
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


