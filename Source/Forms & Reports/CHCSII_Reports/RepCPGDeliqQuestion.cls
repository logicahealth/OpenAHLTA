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
Attribute VB_Name = "RepCPGDeliqQuestion"
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

    '-- Get Candidates
    '----------------------------------------------------------
    If Not GetDelinquent(oReportItem, oRec, oCollFormulas) Then GoTo CleanUp
    moCollRecSet.Add oRec, "delinquent"

'-- Create Report Object
'-----------------------------------------------------------------------------
    If Not CreateReport(sReportName & ".rpt", oReport, oCollFormulas, moCollRecSet) Then GoTo CleanUp

BuildReport = True

GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepCPGDeliqQuestion.BuildReport", "CHCSII_Reports", vbExclamation
CleanUp:
    Set oCrystal = Nothing
    Set oRec = Nothing
    
End Function

Private Function GetDelinquent(oReportItem As ReportItem, _
                               oRec As ADODB.Recordset, _
                               oCollFormulas As Collection) As Boolean

    Dim sQDEF_ID As String
    Dim sVersion As String
    Dim sName As String

    Dim sStmt As String
    
    On Error GoTo ErrHandler
    GetDelinquent = False
        
    sQDEF_ID = oReportItem.CriteriaSel1.Item(1).QDefID
    sVersion = oReportItem.CriteriaSel1.Item(1).Version
    sName = "Questionnaire: " & oReportItem.CriteriaSel1(1).Name & " Version: " & CStr(sVersion) & " "
    
    '-- This select gets all patients who have never taken the annual questionnaire
    '--------------------------------------------------------------------------------
    '-- Select
        sStmt = "Select distinct " & vbNewLine & _
                    "mn.unit_number, " & vbNewLine & _
                    "ma.street_1, " & vbNewLine & _
                    "ma.street_2, " & vbNewLine & _
                    "ma.city, " & vbNewLine & _
                    "ma.state, " & vbNewLine & _
                    "ma.post_code, " & vbNewLine & _
                    "mn.last_name ||', ' || mn.first_name || ' ' || " & vbNewLine & _
                    "mn.middle_name as Patient_name, " & vbNewLine & _
                    "DECODE(hp.area_code,null,null, '(' || hp.area_code ||') ') || DECODE(hp.local_number,null, null, " & vbNewLine & _
                    "SUBSTR(hp.local_number, 1, 3) || '-' || SUBSTR(hp.local_number,4,4)) as Home_phone, " & vbNewLine & _
                    "To_date(Null) As start_date_time, " & vbNewLine & _
                    "To_Char(To_date(Null),'DD MON YYYY') as Date_Time_Format " & vbNewLine
    
    '-- From
        sStmt = sStmt & _
                "From " & vbNewLine & _
                    "mmi_address ma, " & vbNewLine & _
                    "mmi_name mn, " & vbNewLine & _
                    "mmi_phone hp, " & vbNewLine & _
                    "cpgregistry cr " & vbNewLine
    
    '"mmi_site ms, " & vbNewLine &
    '"ms.facility_cid = " & oReportItem.CriteriaSel1.Item(1).NCID & " and " & vbNewLine & _
    '"ms.unit_number = mn.unit_number and " & vbNewLine &
    
    '-- Where
        sStmt = sStmt & _
                "Where " & vbNewLine & _
                    "hp.phone_type_cid (+) = 193 and " & vbNewLine & _
                    "hp.unit_number (+) = cr.unit_number and " & vbNewLine & _
                    "ma.unit_number (+) = cr.unit_number and " & vbNewLine & _
                    "mn.name_type_cid  = 187 and " & vbNewLine & _
                    "mn.unit_number = cr.unit_number and " & vbNewLine & _
                    "cr.facility_ncid = '" & cwobjLogonEnviron.FacilityNCID & "' and " & vbNewLine
                    
                    '-- Check that patient has never taken the DIABETES ANNUAL questionnaire
        sStmt = sStmt & _
                    "not exists " & vbNewLine & _
                    "(select 'x' " & vbNewLine & _
                    "from qresp qr " & vbNewLine & _
                    "Where qr.QDEF_ID = " & sQDEF_ID & " and " & vbNewLine & _
                    "qr.Version = " & sVersion & " and " & vbNewLine & _
                    "qr.unit_number = cr.unit_number) " & vbNewLine
                    
                    
                    '-- Believe this code was not necessary
                    'and
                    '"not exists " & vbNewLine & _
                    "(select 'x' " & vbNewLine & _
                    "from qdef qd, qresp qr " & vbNewLine & _
                    "Where " & vbNewLine & _
                    "UPPER(qd.NAME) = '" & CPG_DIABETES_ANNUAL_Q_NAME & "' and " & vbNewLine & _
                    "qd.qdef_id = qr.qdef_id and " & vbNewLine & _
                    "qr.start_date_time between (sysdate - 365) and sysdate and " & vbNewLine & _
                    "qr.unit_number = cr.unit_number) " & vbNewLine
    
    '-- UNION
        sStmt = sStmt & "Union " & vbNewLine
    
    '-- This select gets all patients who HAVE taken the annual questionnaire
    '-- but are deliquent
    '--------------------------------------------------------------------------------
    '-- Select
        sStmt = sStmt & _
                "Select distinct " & vbNewLine & _
                "mn.unit_number, " & vbNewLine & _
                "ma.street_1, " & vbNewLine & _
                "ma.street_2, " & vbNewLine & _
                "ma.city, " & vbNewLine & _
                "ma.state, " & vbNewLine & _
                "ma.post_code, " & vbNewLine & _
                "mn.last_name ||', ' || mn.first_name || ' ' || " & vbNewLine & _
                "mn.middle_name as Patient_name, " & vbNewLine & _
                "DECODE(hp.area_code,null,null, '(' || hp.area_code ||') ') || DECODE(hp.local_number,null, null, " & vbNewLine & _
                "SUBSTR(hp.local_number, 1, 3) || '-' || SUBSTR(hp.local_number,4,4)) as Home_phone, " & vbNewLine & _
                "q.start_date_time, " & vbNewLine & _
                "To_Char(q.start_date_time,'DD MON YYYY') as Date_Time_Format " & vbNewLine
    
    '-- From
        sStmt = sStmt & _
                "From " & vbNewLine & _
                "mmi_address ma, " & vbNewLine & _
                "mmi_name mn, " & vbNewLine & _
                "mmi_phone hp, " & vbNewLine & _
                "qresp q, " & vbNewLine & _
                "qdef qd, " & vbNewLine & _
                "cpgregistry cr " & vbNewLine
                
    '"mmi_site ms, " & vbNewLine & _
    "ms.facility_cid = " & oReportItem.CriteriaSel1.Item(1).NCID & " and " & vbNewLine & _
    "ms.unit_number = mn.unit_number and " & vbNewLine & _

    '-- Where
        sStmt = sStmt & _
                "Where  cr.facility_ncid = '" & cwobjLogonEnviron.FacilityNCID & "' and " & _
                "hp.phone_type_cid = 193 and " & vbNewLine & _
                "hp.unit_number = cr.unit_number and " & vbNewLine & _
                "ma.unit_number = cr.unit_number and " & vbNewLine & _
                "mn.name_type_cid = 187 and " & vbNewLine & _
                "mn.unit_number = cr.unit_number and " & vbNewLine
            
    '-- This select determines if the latest DIABETES ANNUAL Questionnaire is deliquent
    '--------------------------------------------------------------------------------
        sStmt = sStmt & _
                "Not Exists " & vbNewLine & _
                "(select 'x' " & vbNewLine & _
                "from qresp qr " & vbNewLine & _
                "Where " & vbNewLine & _
                "qr.QDEF_ID = " & sQDEF_ID & " and " & vbNewLine & _
                "qr.VERSION  =" & sVersion & " and " & vbNewLine & _
                "qr.start_date_time between (sysdate - 365) and (sysdate + 1) and " & vbNewLine & _
                "qr.qresp_id = q.qresp_id and" & vbNewLine & _
                "qr.unit_number = cr.unit_number) and" & vbNewLine
                
    '-- This select determines the latest DIABETES ANNUAL Questionnaire
    '-- for the purpose of displaying the last date the questionnaire was taken
    '--------------------------------------------------------------------------------
        sStmt = sStmt & _
                "q.qresp_id in " & vbNewLine & _
                "(select max(qresp_id) " & vbNewLine & _
                "from qresp qr " & vbNewLine & _
                "Where qr.QDEF_ID = " & sQDEF_ID & " and " & vbNewLine & _
                "qr.VERSION = " & sVersion & " and " & vbNewLine & _
                "qr.unit_number = cr.unit_number)and " & vbNewLine & _
                "q.unit_number = cr.unit_number" & vbNewLine
    
    '-- Order by
        sStmt = sStmt & "Order by start_date_time, Patient_name "
    
    Set oRec = goDBDAS.OpenRecordset(sStmt)
    GetDelinquent = True
    
    On Error Resume Next
    oCollFormulas.Remove "Header1"
    oCollFormulas.Add CrystalQuote("Delinquent Annual Questionnaire"), "Header1"
    oCollFormulas.Remove "Header2"
    oCollFormulas.Add CrystalQuote(sName), "Header2"
    
    GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepRegistryCandidates.GetDelinquent", "CHCSII_Reports", vbExclamation
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


