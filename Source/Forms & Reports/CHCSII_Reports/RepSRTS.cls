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
Attribute VB_Name = "RepSRTS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private moCollRecSet As Collection

Public Function BuildReport(sReportName As String, _
                            oReportItem As ReportItem, _
                            oReport As CRAXDRT.Report, _
                            oCollFormulas As Collection) As Boolean
 
On Error GoTo ErrHandler
 
BuildReport = False

Select Case oReportItem.Report
    Case SRTS_OrderedLenses
        If Not BuildOrderedLensesReport(sReportName, oReportItem, oReport, oCollFormulas) Then GoTo CleanUp
    Case SRTS_Rejection
        If Not BuildRejectionReport(sReportName, oReportItem, oReport, oCollFormulas) Then GoTo CleanUp
    Case SRTS_OrderedFrames
        If Not BuildOrderedFramesReport(sReportName, oReportItem, oReport, oCollFormulas) Then GoTo CleanUp
    Case SRTS_Overdue10days
        If Not BuildOverdueReport(sReportName, oReportItem, oReport, oCollFormulas, 10) Then GoTo CleanUp
    Case SRTS_Overdue20days
        If Not BuildOverdueReport(sReportName, oReportItem, oReport, oCollFormulas, 20) Then GoTo CleanUp
    Case SRTS_Overdue30days
        If Not BuildOverdueReport(sReportName, oReportItem, oReport, oCollFormulas, 30) Then GoTo CleanUp
    Case SRTS_Overdue
        If Not BuildOverdueRangeReport(sReportName, oReportItem, oReport, oCollFormulas) Then GoTo CleanUp
End Select

   
BuildReport = True

GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "BuildReport", App.EXEName, vbExclamation
CleanUp:
    
End Function

Public Function BuildOrderedLensesReport(sReportName As String, _
                oReportItem As ReportItem, _
                oReport As CRAXDRT.Report, _
                oCollFormulas As Collection) As Boolean
    '-- Get Services Due
    '----------------------------------------------------------
    'If Not GetOrderedLenses(oReportItem, oRec) Then GoTo CleanUp

On Error GoTo ErrHandler
 
BuildOrderedLensesReport = False

Dim oRec As ADODB.Recordset
Dim sStmt As String
Dim sSQLWhere As String
Dim sDateFilter As String

DoEvents


'-- Get Clinic NCIDs
'----------------------------------------------------------------------------
If Not GetSQLWhere(oReportItem, sSQLWhere) Then GoTo CleanUp
If Not GetSQLDateRangeSRTS(oReportItem, "sot", sDateFilter) Then GoTo CleanUp

'-- Get Data
'----------------------------------------------------------------------------
    Set moCollRecSet = Nothing
    Set moCollRecSet = New Collection




'R. Foster SCR #40701; 40702; 40731; 40732; 40733; 40734; 40735; 40736; 40737; 40746; 40753; 40754; 40759; 40767; 40768; 40770; 40773; 40774;   Developer: Robert Foster 10/27/2003 12:31 PM
'R. Foster Note: Function Changed:CHCSII_FORMS.dll - GetDasOBJ, GetSQLOBJ, (Declarations), GetDD771Data2, GetDD771Data3, GetDD771Data1, Initialize    Reports.dll - BuildOrderedLensesReport, BuildOrderedFramesReport, BuildRejectionReport, BuildOverdueReport, BuildOverdueRangeReport, Initialize, GetDasOBJ, GetSQLOBJ,       Impact of Change: Allow the CHCSII SRTS II Reports and Forms moodules to operate in ITT mode.
'ITT 832 UPDATES BLOCK II
    If gobjshared.IsAppMode(modeITT) Then
        sSQLWhere = Replace(sSQLWhere, "c.", vbNullString)
        sDateFilter = Replace(sDateFilter, "sot.", vbNullString)
        sDateFilter = Replace(sDateFilter, "TO_DATE(", "CONVERT(DATETIME,")
        sDateFilter = Replace(sDateFilter, ",'YYYY/MM/DD HH24:MI')", ")")
        sStmt = "SELECT * FROM vwOrderedFramesLensesReport WHERE "
        sStmt = sStmt & "("
        sStmt = sStmt & sDateFilter & " AND " & sSQLWhere
        sStmt = sStmt & ")"
    Else
        '-- Select
        '------------
        sStmt = "select" & vbNewLine & _
                    "sco.clinical_order_id," & vbNewLine & _
                    "substr(sco.clinical_order_id,-6) as order_number," & vbNewLine & _
                    "sco.clinic_facility_id as account_number," & vbNewLine & _
                    "lab.facility_name as lab_name," & vbNewLine & _
                    "clinic.facility_name as clinic_name," & vbNewLine & _
                    "Trim(sco.last_name || ', ' || sco.first_name || ' ' || sco.middle_name) as patient_name," & vbNewLine & _
                    "sco.ssn," & vbNewLine & _
                    "sco.frame_type," & vbNewLine & _
                    "sco.eye," & vbNewLine & _
                    "sco.bridge," & vbNewLine & _
                    "sco.temple," & vbNewLine & _
                    "sco.frame_color," & vbNewLine & _
                    "sco.distpd," & vbNewLine & _
                    "sco.nearpd," & vbNewLine & _
                    "sco.lens_type," & vbNewLine & _
                    "sco.tint," & vbNewLine & _
                    "sco.lens_material," & vbNewLine & _
                    "sco.pair," & vbNewLine & _
                    "sco.cases," & vbNewLine & _
                    "sco.right_sphere," & vbNewLine & _
                    "sco.right_cylinder," & vbNewLine & _
                    "sco.right_axis," & vbNewLine & _
                    "sco.right_hprism," & vbNewLine & _
                    "sco.right_hbase," & vbNewLine
        
        sStmt = sStmt & _
                    "sco.right_vprism," & vbNewLine & _
                    "sco.right_vbase," & vbNewLine & _
                    "sco.left_sphere," & vbNewLine & _
                    "sco.left_cylinder," & vbNewLine & _
                    "sco.left_axis," & vbNewLine & _
                    "sco.left_hprism," & vbNewLine & _
                    "sco.left_hbase," & vbNewLine & _
                    "sco.left_vprism," & vbNewLine & _
                    "sco.left_vbase," & vbNewLine & _
                    "sco.right_add," & vbNewLine & _
                    "sco.right_seght," & vbNewLine & _
                    "sco.left_add," & vbNewLine & _
                    "sco.left_seght," & vbNewLine & _
                    "sco.priority," & vbNewLine & _
                    "sco.user_initials," & vbNewLine & _
                    "sp.description as priority_description," & vbNewLine & _
                    "sco.fund_cite_id," & vbNewLine & _
                    "sco.clinical_order_comment," & vbNewLine & _
                    "sco.clinician_name," & vbNewLine & _
                    "To_Char(sot.dt_transaction,'DD MON YYYY') as Trans_Date," & vbNewLine & _
                    "sfe.b_val" & vbNewLine
        
        '-- From
        '------------
        sStmt = sStmt & _
                "From" & vbNewLine & _
                    "srts_frame_eye_xrefs sfe," & vbNewLine & _
                    "srts_order_transaction sot," & vbNewLine & _
                    "srts_priority sp," & vbNewLine & _
                    "srts_facility lab," & vbNewLine & _
                    "srts_facility clinic," & vbNewLine & _
                    "srts_clinical_order sco," & vbNewLine & _
                    "srts_facility sf " & vbNewLine
        
        '-- Where
        '------------
        sStmt = sStmt & _
                "Where" & vbNewLine & _
                    "sfe.eye (+) = sco.eye and" & vbNewLine & _
                    "sfe.frame_type (+) = sco.frame_type and" & vbNewLine & _
                    "sot.dt_transaction in" & vbNewLine & _
                        "(select max(dt_transaction)" & vbNewLine & _
                        "from srts_order_transaction sot2" & vbNewLine & _
                        "Where" & vbNewLine & _
                            "sot2.order_status = 3 and" & vbNewLine & _
                            "sot2.clinical_order_id (+) = sot.clinical_order_id) and" & vbNewLine & _
                    "sot.clinical_order_id = sco.clinical_order_id and" & vbNewLine & _
                    "sp.priority (+) = sco.priority and" & vbNewLine & _
                    "lab.facility_id (+) = sco.lab_facility_id and" & vbNewLine & _
                    "clinic.facility_id (+) = sco.clinic_facility_id and " & vbNewLine & _
                    "(sco.lab_facility_id = sf.facility_id or " & vbNewLine & _
                    "sco.clinic_facility_id = sf.facility_id) and " & vbNewLine & _
                    "sf.facility_ncid in " & vbNewLine & _
                    "(select c.ncid " & vbNewLine & _
                    "from clinic c " & vbNewLine & _
                    "where " & vbNewLine & _
                    sSQLWhere & ") and" & vbNewLine & _
                    sDateFilter
    End If
          
    Set oRec = goDBDAS.OpenRecordset(sStmt)

'-- Add Recordset to list
'----------------------------------------------------------------------------
    moCollRecSet.Add oRec, "OrderedLenses"

'-- Get Clinic Header
'----------------------------------------------------------------------------
    If Not GetClinicHeader(oReportItem, oCollFormulas) Then GoTo CleanUp

'-- Create Report Object
'-----------------------------------------------------------------------------
    If Not CreateReport(sReportName & ".rpt", oReport, oCollFormulas, moCollRecSet) Then GoTo CleanUp

BuildOrderedLensesReport = True

GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "BuildOrderedLensesReport", App.EXEName, vbExclamation
CleanUp:
    Set oRec = Nothing
End Function
Public Function BuildOrderedFramesReport(sReportName As String, _
                oReportItem As ReportItem, _
                oReport As CRAXDRT.Report, _
                oCollFormulas As Collection) As Boolean
    '-- Get Services Due
    '----------------------------------------------------------
    'If Not GetOrderedLenses(oReportItem, oRec) Then GoTo CleanUp

On Error GoTo ErrHandler
 
BuildOrderedFramesReport = False

Dim oRec As ADODB.Recordset
Dim sStmt As String
Dim sSQLWhere As String
Dim sDateFilter As String
Dim sOrderNumberFilter As String
Dim sFromNum As String
Dim sToNum As String




DoEvents


'-- Get Clinic NCIDs
'----------------------------------------------------------------------------
If Not GetSQLDateRangeSRTS(oReportItem, "sot", sDateFilter) Then GoTo CleanUp
If Not GetSQLWhere(oReportItem, sSQLWhere) Then GoTo CleanUp

If Not oReportItem.CriteriaSel2 Is Nothing Then

    sFromNum = oReportItem.CriteriaSel2.Item(1).OrderNumberFrom
    
    sOrderNumberFilter = _
        "(sco.clinical_order_id Like '" & Mid$(sFromNum, 1, InStr(8, sFromNum, "-") - 1) & "%' and " & vbNewLine

    sFromNum = Mid$(sFromNum, InStr(8, oReportItem.CriteriaSel2.Item(1).OrderNumberFrom, "-") + 1)
    sToNum = oReportItem.CriteriaSel2.Item(1).OrderNumberTo
    sToNum = Mid$(sToNum, InStr(8, oReportItem.CriteriaSel2.Item(1).OrderNumberTo, "-") + 1)
    
    sOrderNumberFilter = sOrderNumberFilter & _
        "TO_NUMBER(SUBSTR(sco.clinical_order_id,13)) >= " & sFromNum & " and" & vbNewLine & _
        "TO_NUMBER(SUBSTR(sco.clinical_order_id,13)) <= " & sToNum & ") and" & vbNewLine

    oCollFormulas.Remove ("Header3")
    oCollFormulas.Add CrystalQuote("Order Number: " & oReportItem.CriteriaSel2.Item(1).OrderNumberFrom & "  to  " & _
                oReportItem.CriteriaSel2.Item(1).OrderNumberTo), "Header3"

Else
    sOrderNumberFilter = vbNullString
End If


'-- Get Data
'----------------------------------------------------------------------------
    Set moCollRecSet = Nothing
    Set moCollRecSet = New Collection

'-- Select
'------------
'R. Foster SCR #40701; 40702; 40731; 40732; 40733; 40734; 40735; 40736; 40737; 40746; 40753; 40754; 40759; 40767; 40768; 40770; 40773; 40774;   Developer: Robert Foster 10/27/2003 12:31 PM
'R. Foster Note: Function Changed:CHCSII_FORMS.dll - GetDasOBJ, GetSQLOBJ, (Declarations), GetDD771Data2, GetDD771Data3, GetDD771Data1, Initialize    Reports.dll - BuildOrderedLensesReport, BuildOrderedFramesReport, BuildRejectionReport, BuildOverdueReport, BuildOverdueRangeReport, Initialize, GetDasOBJ, GetSQLOBJ,       Impact of Change: Allow the CHCSII SRTS II Reports and Forms moodules to operate in ITT mode.
'ITT 832 UPDATES BLOCK II
    If gobjshared.IsAppMode(modeITT) Then
        sSQLWhere = Replace(sSQLWhere, "c.", vbNullString)
        sDateFilter = Replace(sDateFilter, "sot.", vbNullString)
        sDateFilter = Replace(sDateFilter, "TO_DATE(", "CONVERT(DATETIME,")
        sDateFilter = Replace(sDateFilter, ",'YYYY/MM/DD HH24:MI')", ")")
        sOrderNumberFilter = Replace(sOrderNumberFilter, "sco.", vbNullString)
        sOrderNumberFilter = Replace(sOrderNumberFilter, "TO_NUMBER(SUBSTR(", "CONVERT(NUMERIC,SUBSTRING(")
        sOrderNumberFilter = Replace(sOrderNumberFilter, ",13))", ",13,7))")
        sStmt = "SELECT * FROM vwOrderedFramesLensesReport WHERE "
        sStmt = sStmt & "("
        sStmt = sStmt & sOrderNumberFilter & " " & sDateFilter & " AND " & sSQLWhere
        sStmt = sStmt & ")"
    Else
        sStmt = "select" & vbNewLine & _
                    "sco.clinical_order_id," & vbNewLine & _
                    "substr(sco.clinical_order_id,-6) as order_number," & vbNewLine & _
                    "sco.clinic_facility_id as account_number," & vbNewLine & _
                    "lab.facility_name as lab_name," & vbNewLine & _
                    "clinic.facility_name as clinic_name," & vbNewLine & _
                    "Trim(sco.last_name || ', ' || sco.first_name || ' ' || sco.middle_name) as patient_name," & vbNewLine & _
                    "sco.ssn," & vbNewLine & _
                    "sco.frame_type," & vbNewLine & _
                    "sco.eye," & vbNewLine & _
                    "sco.bridge," & vbNewLine & _
                    "sco.temple," & vbNewLine & _
                    "sco.frame_color," & vbNewLine & _
                    "sco.distpd," & vbNewLine & _
                    "sco.nearpd," & vbNewLine & _
                    "sco.lens_type," & vbNewLine & _
                    "sco.tint," & vbNewLine & _
                    "sco.lens_material," & vbNewLine & _
                    "sco.pair," & vbNewLine & _
                    "sco.cases," & vbNewLine & _
                    "sco.right_sphere," & vbNewLine & _
                    "sco.right_cylinder," & vbNewLine & _
                    "sco.right_axis," & vbNewLine & _
                    "sco.right_hprism," & vbNewLine & _
                    "sco.right_hbase," & vbNewLine
        
        sStmt = sStmt & _
                    "sco.right_vprism," & vbNewLine & _
                    "sco.right_vbase," & vbNewLine & _
                    "sco.left_sphere," & vbNewLine & _
                    "sco.left_cylinder," & vbNewLine & _
                    "sco.left_axis," & vbNewLine & _
                    "sco.left_hprism," & vbNewLine & _
                    "sco.left_hbase," & vbNewLine & _
                    "sco.left_vprism," & vbNewLine & _
                    "sco.left_vbase," & vbNewLine & _
                    "sco.right_add," & vbNewLine & _
                    "sco.right_seght," & vbNewLine & _
                    "sco.left_add," & vbNewLine & _
                    "sco.left_seght," & vbNewLine & _
                    "sco.priority," & vbNewLine & _
                    "sco.user_initials," & vbNewLine & _
                    "sco.patient_status," & vbNewLine & _
                    "TO_CHAR(sp.priority) || ' - ' || sp.description as priority_description," & vbNewLine & _
                    "sco.fund_cite_id," & vbNewLine & _
                    "sco.clinical_order_comment," & vbNewLine & _
                    "sco.clinician_name," & vbNewLine & _
                    "To_Char(sot.dt_transaction,'DD MON YYYY') as Trans_Date," & vbNewLine & _
                    "sfe.b_val" & vbNewLine
        
        '-- From
        '------------
        sStmt = sStmt & _
                "From" & vbNewLine & _
                    "srts_frame_eye_xrefs sfe," & vbNewLine & _
                    "srts_priority sp," & vbNewLine & _
                    "srts_facility lab," & vbNewLine & _
                    "srts_facility clinic," & vbNewLine & _
                    "srts_clinical_order sco," & vbNewLine & _
                    "srts_order_transaction sot," & vbNewLine & _
                    "srts_facility sf " & vbNewLine
        
        '-- Where
        '------------
        sStmt = sStmt & _
                "Where" & vbNewLine
        
        sStmt = sStmt & sOrderNumberFilter
        
        sStmt = sStmt & _
                    "sfe.eye (+) = sco.eye and" & vbNewLine & _
                    "sfe.frame_type (+) = sco.frame_type and" & vbNewLine & _
                    "sot.dt_transaction in" & vbNewLine & _
                        "(select max(dt_transaction)" & vbNewLine & _
                        "from srts_order_transaction sot2" & vbNewLine & _
                        "Where" & vbNewLine & _
                            "sot2.order_status = 3 and" & vbNewLine & _
                            "sot2.clinical_order_id (+) = sot.clinical_order_id) and" & vbNewLine & _
                    "sot.clinical_order_id (+) = sco.clinical_order_id and" & vbNewLine & _
                    "sp.priority = sco.priority and" & vbNewLine & _
                    "lab.facility_id (+) = sco.lab_facility_id and" & vbNewLine & _
                    "clinic.facility_id (+) = sco.clinic_facility_id and " & vbNewLine & _
                    "(sco.lab_facility_id = sf.facility_id or " & vbNewLine & _
                    "sco.clinic_facility_id = sf.facility_id) and " & vbNewLine & _
                    "sf.facility_ncid in " & vbNewLine & _
                    "(select c.ncid " & vbNewLine & _
                    "from clinic c " & vbNewLine & _
                    "where " & vbNewLine & _
                    sSQLWhere & ") and" & vbNewLine & _
                    sDateFilter
    End If

    Set oRec = goDBDAS.OpenRecordset(sStmt)

'-- Add Recordset to list
'----------------------------------------------------------------------------
    moCollRecSet.Add oRec, "OrderedLenses"

'-- Get Clinic Header
'----------------------------------------------------------------------------
    If Not GetClinicHeader(oReportItem, oCollFormulas) Then GoTo CleanUp


'-- Create Report Object
'-----------------------------------------------------------------------------
    If Not CreateReport(sReportName & ".rpt", oReport, oCollFormulas, moCollRecSet) Then GoTo CleanUp

BuildOrderedFramesReport = True

GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "BuildOrderedFramesReport", App.EXEName, vbExclamation
CleanUp:
    Set oRec = Nothing
End Function

Public Function BuildRejectionReport(sReportName As String, _
                oReportItem As ReportItem, _
                oReport As CRAXDRT.Report, _
                oCollFormulas As Collection) As Boolean

On Error GoTo ErrHandler

BuildRejectionReport = False

Dim oRec As ADODB.Recordset
Dim sStmt As String
Dim sSQLWhere As String
Dim sDateFilter As String

DoEvents


'-- Get Clinic NCIDs
'----------------------------------------------------------------------------
If Not GetSQLDateRangeSRTS(oReportItem, "sot2", sDateFilter) Then GoTo CleanUp
If Not GetSQLWhere(oReportItem, sSQLWhere) Then GoTo CleanUp


'-- Get Data
'----------------------------------------------------------------------------
    Set moCollRecSet = Nothing
    Set moCollRecSet = New Collection

'-- Select
'------------
'R. Foster SCR #40701; 40702; 40731; 40732; 40733; 40734; 40735; 40736; 40737; 40746; 40753; 40754; 40759; 40767; 40768; 40770; 40773; 40774;   Developer: Robert Foster 10/27/2003 12:31 PM
'R. Foster Note: Function Changed:CHCSII_FORMS.dll - GetDasOBJ, GetSQLOBJ, (Declarations), GetDD771Data2, GetDD771Data3, GetDD771Data1, Initialize    Reports.dll - BuildOrderedLensesReport, BuildOrderedFramesReport, BuildRejectionReport, BuildOverdueReport, BuildOverdueRangeReport, Initialize, GetDasOBJ, GetSQLOBJ,       Impact of Change: Allow the CHCSII SRTS II Reports and Forms moodules to operate in ITT mode.
'ITT 832 UPDATES BLOCK II
    If gobjshared.IsAppMode(modeITT) Then
        sSQLWhere = Replace(sSQLWhere, "c.", vbNullString)
        sDateFilter = Replace(sDateFilter, "sot2.", vbNullString)
        sDateFilter = Replace(sDateFilter, "TO_DATE(", "CONVERT(DATETIME,")
        sDateFilter = Replace(sDateFilter, ",'YYYY/MM/DD HH24:MI')", ")")
        sDateFilter = Replace(sDateFilter, "order_status=3", "ORDER_STATUS = 13")
        sStmt = "SELECT * FROM vwOrderedFramesLensesReport WHERE "
        sStmt = sStmt & "("
        sStmt = sStmt & sDateFilter & " AND " & sSQLWhere & " "
        sStmt = sStmt & ")"
    Else
        sStmt = "select" & vbNewLine & _
                    "sco.clinic_facility_id as account_number," & vbNewLine & _
                    "lab.facility_name as lab_name," & vbNewLine & _
                    "clinic.facility_name as clinic_name," & vbNewLine & _
                    "substr(sot.order_comment,1,1) as Clinic_Type," & vbNewLine & _
                    "Trim(substr(sot.order_comment, instr(sot.order_comment,': ') + 1)) as TheComment" & vbNewLine
        
        '-- From
        '------------
        sStmt = sStmt & _
                "From" & vbNewLine & _
                    "srts_order_transaction sot," & vbNewLine & _
                    "srts_priority sp," & vbNewLine & _
                    "srts_facility lab," & vbNewLine & _
                    "srts_facility clinic," & vbNewLine & _
                    "srts_facility sf," & vbNewLine & _
                    "srts_clinical_order sco," & vbNewLine & _
                    "srts_order_transaction sot2" & vbNewLine
        
        '-- Where
        '------------
        sStmt = sStmt & _
                "Where" & vbNewLine & _
                    "sot.order_comment is not null and sot.order_status=13 and " & vbNewLine & _
                    "sot.clinical_order_id = sco.clinical_order_id and" & vbNewLine & _
                    "sp.priority (+) = sco.priority and" & vbNewLine & _
                    "lab.facility_id (+) = sco.lab_facility_id and" & vbNewLine & _
                    "clinic.facility_id (+) = sco.clinic_facility_id and " & vbNewLine & _
                    "(sco.lab_facility_id = sf.facility_id or " & vbNewLine & _
                    "sco.clinic_facility_id = sf.facility_id) and " & vbNewLine & _
                    "sf.facility_ncid in " & vbNewLine & _
                    "(select c.ncid " & vbNewLine & _
                    "from clinic c " & vbNewLine & _
                    "where " & vbNewLine & _
                    sSQLWhere & ") and" & vbNewLine & _
                    "sco.clinical_order_id=sot2.clinical_order_id and" & vbNewLine & _
                    sDateFilter
    End If
            
    Set oRec = goDBDAS.OpenRecordset(sStmt)

'-- Add Recordset to list
'----------------------------------------------------------------------------
    moCollRecSet.Add oRec, "rejection"

'-- Get Clinic Header
'----------------------------------------------------------------------------
    If Not GetClinicHeader(oReportItem, oCollFormulas) Then GoTo CleanUp

'-- Create Report Object
'-----------------------------------------------------------------------------
    If Not CreateReport(sReportName & ".rpt", oReport, oCollFormulas, moCollRecSet) Then GoTo CleanUp

BuildRejectionReport = True

GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "BuildRejectionReport", App.EXEName, vbExclamation
CleanUp:
    Set oRec = Nothing
End Function


Public Function BuildOverdueReport(sReportName As String, _
                oReportItem As ReportItem, _
                oReport As CRAXDRT.Report, _
                oCollFormulas As Collection, _
                iDaysOverdue As Integer) As Boolean

On Error GoTo ErrHandler

BuildOverdueReport = False

Dim oRec As ADODB.Recordset
Dim sStmt As String
Dim sSQLWhere As String

DoEvents

'-- Get Clinic NCIDs
'----------------------------------------------------------------------------
If Not GetSQLWhere(oReportItem, sSQLWhere) Then GoTo CleanUp

'-- Get Data
'----------------------------------------------------------------------------
    Set moCollRecSet = Nothing
    Set moCollRecSet = New Collection

'-- Select
'------------

'R. Foster SCR #40701; 40702; 40731; 40732; 40733; 40734; 40735; 40736; 40737; 40746; 40753; 40754; 40759; 40767; 40768; 40770; 40773; 40774;   Developer: Robert Foster 10/27/2003 12:31 PM
'R. Foster Note: Function Changed:CHCSII_FORMS.dll - GetDasOBJ, GetSQLOBJ, (Declarations), GetDD771Data2, GetDD771Data3, GetDD771Data1, Initialize    Reports.dll - BuildOrderedLensesReport, BuildOrderedFramesReport, BuildRejectionReport, BuildOverdueReport, BuildOverdueRangeReport, Initialize, GetDasOBJ, GetSQLOBJ,       Impact of Change: Allow the CHCSII SRTS II Reports and Forms moodules to operate in ITT mode.
'ITT 832 UPDATES BLOCK II
If gobjshared.IsAppMode(modeITT) Then
    sSQLWhere = Replace(sSQLWhere, "c.", vbNullString)
    sStmt = "SELECT  * FROM cwBuildOverdueRangeReport WHERE "
    sStmt = sStmt & "(((GETDATE() - DT_TRANSACTION) >= " & CStr(iDaysOverdue) & ") AND "
    sStmt = sStmt & sSQLWhere & ")"
Else
        sStmt = "select" & vbNewLine & _
                    "sco.clinical_order_id," & vbNewLine & _
                    "substr(sco.clinical_order_id,-6) as order_number," & vbNewLine & _
                    "sco.clinic_facility_id as account_number," & vbNewLine & _
                    "lab.facility_name as lab_name," & vbNewLine & _
                    "clinic.facility_name as clinic_name," & vbNewLine & _
                    "Trim(sco.last_name || ', ' || sco.first_name || ' ' || sco.middle_name) as patient_name," & vbNewLine & _
                    "sco.ssn," & vbNewLine & _
                    "sco.frame_type," & vbNewLine & _
                    "sco.eye," & vbNewLine & _
                    "sco.bridge," & vbNewLine & _
                    "sco.temple," & vbNewLine & _
                    "sco.frame_color," & vbNewLine & _
                    "sco.distpd," & vbNewLine & _
                    "sco.nearpd," & vbNewLine & _
                    "sco.lens_type," & vbNewLine & _
                    "sco.tint," & vbNewLine & _
                    "sco.lens_material," & vbNewLine & _
                    "sco.pair," & vbNewLine & _
                    "sco.cases," & vbNewLine & _
                    "sco.right_sphere," & vbNewLine & _
                    "sco.right_cylinder," & vbNewLine & _
                    "sco.right_axis," & vbNewLine & _
                    "sco.right_hprism," & vbNewLine & _
                    "sco.right_hbase," & vbNewLine
        
        sStmt = sStmt & _
                    "sco.right_vprism," & vbNewLine & _
                    "sco.right_vbase," & vbNewLine & _
                    "sco.left_sphere," & vbNewLine & _
                    "sco.left_cylinder," & vbNewLine & _
                    "sco.left_axis," & vbNewLine & _
                    "sco.left_hprism," & vbNewLine & _
                    "sco.left_hbase," & vbNewLine & _
                    "sco.left_vprism," & vbNewLine & _
                    "sco.left_vbase," & vbNewLine & _
                    "sco.right_add," & vbNewLine & _
                    "sco.right_seght," & vbNewLine & _
                    "sco.left_add," & vbNewLine & _
                    "sco.left_seght," & vbNewLine & _
                    "sco.priority," & vbNewLine & _
                    "sco.user_initials," & vbNewLine & _
                    "sp.description as priority_description," & vbNewLine & _
                    "sco.fund_cite_id," & vbNewLine & _
                    "sco.clinical_order_comment," & vbNewLine & _
                    "sco.clinician_name," & vbNewLine & _
                    "To_Char(sot.dt_transaction,'DD MON YYYY') as Trans_Date," & vbNewLine & _
                    "sot.dt_transaction," & vbNewLine & _
                    "sfe.b_val" & vbNewLine
        
        '-- From
        '------------
        sStmt = sStmt & _
                "From" & vbNewLine & _
                    "srts_frame_eye_xrefs sfe," & vbNewLine & _
                    "srts_priority sp," & vbNewLine & _
                    "srts_facility lab," & vbNewLine & _
                    "srts_facility clinic," & vbNewLine & _
                    "srts_clinical_order sco," & vbNewLine & _
                    "srts_order_transaction sot," & vbNewLine & _
                    "srts_facility sf " & vbNewLine
        
        '-- Where
        '------------
        sStmt = sStmt & _
                "Where" & vbNewLine & _
                    "sfe.eye (+) = sco.eye and" & vbNewLine & _
                    "sfe.frame_type (+) = sco.frame_type and" & vbNewLine & _
                    "(not exists" & vbNewLine & _
                        "(select 'x'" & vbNewLine & _
                        "from srts_order_transaction sot2" & vbNewLine & _
                        "Where" & vbNewLine & _
                            "sot2.order_status in (10,11,12,13,14) and" & vbNewLine & _
                            "sot2.clinical_order_id = sot.clinical_order_id)) and" & vbNewLine
        
        sStmt = sStmt & _
                    "((((sysdate - sot.dt_transaction) >= " & CStr(iDaysOverdue) & ")" & vbNewLine
        
        sStmt = sStmt & ") And" & vbNewLine
        
        sStmt = sStmt & _
                    "sot.dt_transaction in" & vbNewLine & _
                        "(select max(dt_transaction)" & vbNewLine & _
                        "from srts_order_transaction sot2" & vbNewLine & _
                        "Where" & vbNewLine & _
                            "sot2.order_status = 3 and" & vbNewLine & _
                            "sot2.clinical_order_id (+) = sot.clinical_order_id)) and" & vbNewLine & _
                    "sot.clinical_order_id (+) = sco.clinical_order_id and" & vbNewLine & _
                    "sp.priority = sco.priority and" & vbNewLine & _
                    "lab.facility_id (+) = sco.lab_facility_id and" & vbNewLine & _
                    "clinic.facility_id (+) = sco.clinic_facility_id and " & vbNewLine & _
                    "(sco.lab_facility_id = sf.facility_id or " & vbNewLine & _
                    "sco.clinic_facility_id = sf.facility_id) and " & vbNewLine & _
                    "sf.facility_ncid in " & vbNewLine & _
                    "(select c.ncid " & vbNewLine & _
                    "from clinic c " & vbNewLine & _
                    "where " & vbNewLine & _
                    sSQLWhere & ")"
    End If

    Set oRec = goDBDAS.OpenRecordset(sStmt)

'-- Add Recordset to list
'----------------------------------------------------------------------------
    moCollRecSet.Add oRec, "overdue"
    
'-- Get Clinic Header
'----------------------------------------------------------------------------
    If Not GetClinicHeader(oReportItem, oCollFormulas) Then GoTo CleanUp

'-- Create Report Object
'-----------------------------------------------------------------------------
    If Not CreateReport(sReportName & ".rpt", oReport, oCollFormulas, moCollRecSet) Then GoTo CleanUp

BuildOverdueReport = True

GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "BuildOverdueReport", App.EXEName, vbExclamation
CleanUp:
    Set oRec = Nothing
End Function
Public Function BuildOverdueRangeReport(sReportName As String, _
                oReportItem As ReportItem, _
                oReport As CRAXDRT.Report, _
                oCollFormulas As Collection) As Boolean

On Error GoTo ErrHandler

BuildOverdueRangeReport = False

Dim oRec As ADODB.Recordset
Dim sStmt As String
Dim sSQLWhere As String
Dim sDateFilter As String

DoEvents

'-- Get Clinic NCIDs
'----------------------------------------------------------------------------
If Not GetSQLDateRangeSRTS2(oReportItem, "sot", sDateFilter) Then GoTo CleanUp
If Not GetSQLWhere(oReportItem, sSQLWhere) Then GoTo CleanUp


'-- Get Data
'----------------------------------------------------------------------------
    Set moCollRecSet = Nothing
    Set moCollRecSet = New Collection

'-- Select
'------------

'R. Foster SCR #40701; 40702; 40731; 40732; 40733; 40734; 40735; 40736; 40737; 40746; 40753; 40754; 40759; 40767; 40768; 40770; 40773; 40774;   Developer: Robert Foster 10/27/2003 12:31 PM
'R. Foster Note: Function Changed:CHCSII_FORMS.dll - GetDasOBJ, GetSQLOBJ, (Declarations), GetDD771Data2, GetDD771Data3, GetDD771Data1, Initialize    Reports.dll - BuildOrderedLensesReport, BuildOrderedFramesReport, BuildRejectionReport, BuildOverdueReport, BuildOverdueRangeReport, Initialize, GetDasOBJ, GetSQLOBJ,       Impact of Change: Allow the CHCSII SRTS II Reports and Forms moodules to operate in ITT mode.
'ITT 832 UPDATES BLOCK II
If gobjshared.IsAppMode(modeITT) Then
    sSQLWhere = Replace(sSQLWhere, "c.", vbNullString)
    sDateFilter = Replace(sDateFilter, "sot.", vbNullString)
    sDateFilter = Replace(sDateFilter, "TO_DATE(", "CONVERT(DATETIME,")
    sDateFilter = Replace(sDateFilter, ",'YYYY/MM/DD HH24:MI')", ")")
    sStmt = "SELECT * FROM cwBuildOverdueRangeReport WHERE (" & sDateFilter & " AND " & sSQLWhere
    sStmt = sStmt & " AND ORDER_STATUS NOT IN (10,11,12,13,14))"
    
Else
    sStmt = "select" & vbNewLine & _
                " sco.clinical_order_id, " & vbNewLine & _
                "substr(sco.clinical_order_id,-6) as order_number," & vbNewLine & _
                "sco.clinic_facility_id as account_number," & vbNewLine & _
                "lab.facility_name as lab_name," & vbNewLine & _
                "clinic.facility_name as clinic_name," & vbNewLine & _
                "Trim(sco.last_name || ', ' || sco.first_name || ' ' || sco.middle_name) as patient_name," & vbNewLine & _
                "sco.ssn," & vbNewLine & _
                "sco.frame_type," & vbNewLine & _
                "sco.eye," & vbNewLine & _
                "sco.bridge," & vbNewLine & _
                "sco.temple," & vbNewLine & _
                "sco.frame_color," & vbNewLine & _
                "sco.distpd," & vbNewLine & _
                "sco.nearpd," & vbNewLine & _
                "sco.lens_type," & vbNewLine & _
                "sco.tint," & vbNewLine & _
                "sco.lens_material," & vbNewLine & _
                "sco.pair," & vbNewLine & _
                "sco.cases," & vbNewLine & _
                "sco.right_sphere," & vbNewLine & _
                "sco.right_cylinder," & vbNewLine & _
                "sco.right_axis," & vbNewLine & _
                "sco.right_hprism," & vbNewLine & _
                "sco.right_hbase," & vbNewLine
    
    sStmt = sStmt & _
                "sco.right_vprism," & vbNewLine & _
                "sco.right_vbase," & vbNewLine & _
                "sco.left_sphere," & vbNewLine & _
                "sco.left_cylinder," & vbNewLine & _
                "sco.left_axis," & vbNewLine & _
                "sco.left_hprism," & vbNewLine & _
                "sco.left_hbase," & vbNewLine & _
                "sco.left_vprism," & vbNewLine & _
                "sco.left_vbase," & vbNewLine & _
                "sco.right_add," & vbNewLine & _
                "sco.right_seght," & vbNewLine & _
                "sco.left_add," & vbNewLine & _
                "sco.left_seght," & vbNewLine & _
                "sco.priority," & vbNewLine & _
                "sco.user_initials," & vbNewLine & _
                "' ' as priority_description," & vbNewLine & _
                "sco.fund_cite_id," & vbNewLine & _
                "sco.clinical_order_comment," & vbNewLine & _
                "sco.clinician_name," & vbNewLine & _
                "To_Char(sot.dt_transaction,'DD MON YYYY') as Trans_Date," & vbNewLine & _
                "sot.dt_transaction," & vbNewLine & _
                "' ' as b_val," & vbNewLine & _
                "sos.description order_status" & vbNewLine & _
                " "
    
    '-- From
    '------------
    sStmt = sStmt & _
            "From" & vbNewLine & _
                "srts_facility lab," & vbNewLine & _
                "srts_facility clinic," & vbNewLine & _
                "srts_clinical_order sco," & vbNewLine & _
                "srts_order_transaction sot," & vbNewLine & _
                "srts_facility sf, " & vbNewLine & _
                "srts_order_status sos" & vbNewLine
    
    '-- Where
    '------------
    
    sStmt = sStmt & _
            "Where" & vbNewLine & _
                "sos.order_status = sot.order_status and" & vbNewLine & _
                "(not exists" & vbNewLine & _
                    "(select 'x'" & vbNewLine & _
                    "from srts_order_transaction sot2" & vbNewLine & _
                    "Where" & vbNewLine & _
                        "sot2.order_status in (10,11,12,13,14) and" & vbNewLine & _
                        "sot2.clinical_order_id = sot.clinical_order_id)) and" & vbNewLine
    
    sStmt = sStmt & "((" & sDateFilter & vbNewLine
    
    sStmt = sStmt & ") And" & vbNewLine
    
    sStmt = sStmt & _
                "sot.dt_transaction in" & vbNewLine & _
                    "(select max(dt_transaction)" & vbNewLine & _
                    "from srts_order_transaction sot2" & vbNewLine & _
                    "Where" & vbNewLine & _
                        "sot2.clinical_order_id (+) = sot.clinical_order_id)) and" & vbNewLine & _
                "sot.clinical_order_id (+) = sco.clinical_order_id and" & vbNewLine & _
                "lab.facility_id (+) = sco.lab_facility_id and" & vbNewLine & _
                "clinic.facility_id (+) = sco.clinic_facility_id and " & vbNewLine & _
                "sco.lab_facility_id = sf.facility_id and " & vbNewLine & _
                "sf.facility_ncid in " & vbNewLine & _
                "(select c.ncid " & vbNewLine & _
                "from clinic c " & vbNewLine & _
                "where " & vbNewLine & _
                sSQLWhere & ")"
End If

Set oRec = goDBDAS.OpenRecordset(sStmt)

'Call DebugPrintRecordSet(oRec)

'-- Add Recordset to list
'----------------------------------------------------------------------------
    moCollRecSet.Add oRec, "overdue"
    
'-- Get Clinic Header
'----------------------------------------------------------------------------
    If Not GetClinicHeader(oReportItem, oCollFormulas) Then GoTo CleanUp

'-- Create Report Object
'-----------------------------------------------------------------------------
    If Not CreateReport(sReportName & ".rpt", oReport, oCollFormulas, moCollRecSet) Then GoTo CleanUp

BuildOverdueRangeReport = True

GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "BuildOverdueRangeReport", App.EXEName, vbExclamation
CleanUp:
    Set oRec = Nothing
End Function

Private Function GetClinicHeader(oReportItem As ReportItem, _
                                 oCollFormulas As Collection) As Boolean

GetClinicHeader = False
On Error GoTo ErrHandler

Dim sStmt As String
Dim i As Integer
Dim sHeader As String
Dim sClinicNCIDs As String


If oReportItem.Criteria1 = Clinic Then
    For i = 1 To oReportItem.CriteriaSel1.Count
        sClinicNCIDs = sClinicNCIDs & oReportItem.CriteriaSel1.Item(i).NCID
        If i <> oReportItem.CriteriaSel1.Count Then
            sClinicNCIDs = sClinicNCIDs & ","
        End If
    Next
    
    sStmt = "select" & vbNewLine & _
                "sf.facility_name as ClinicName," & vbNewLine & _
                "c.meprs_code," & vbNewLine & _
                "c.dmis_id_code" & vbNewLine & _
            "from" & vbNewLine & _
                "clinic c," & vbNewLine & _
                "srts_facility sf" & vbNewLine & _
            "where" & vbNewLine & _
                "c.ncid=sf.facility_ncid and" & vbNewLine & _
                "sf.facility_ncid in (" & sClinicNCIDs & ")"
    
    goDBSqlOps.Execute sStmt
    
    sHeader = "Clinic/Lab: "
    Do While Not goDBSqlOps.EOF
        sHeader = sHeader & goDBSqlOps.Value("ClinicName") & " (Meprs:" & goDBSqlOps.Value("meprs_code") & " DMIS:" & goDBSqlOps.Value("DMIS_ID_CODE") & ")"
        goDBSqlOps.MoveNext
        If Not goDBSqlOps.EOF Then
            sHeader = sHeader & "; "
        End If
    Loop
    
    If Len(sHeader) > 254 Then
        sHeader = Mid$(sHeader, 1, 254)
    End If
    
    oCollFormulas.Remove ("Header2")
    oCollFormulas.Add CrystalQuote(sHeader), "Header2"
End If

GetClinicHeader = True

GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "GetClinicHeader", App.EXEName, vbExclamation
CleanUp:
End Function


Public Function GetSQLWhere(oReportItem As ReportItem, _
                            sSQLWhere As String) As Boolean

If Not DebugMode Then On Error GoTo ErrHandler
GetSQLWhere = False

Dim i As Integer
Dim sClinicNCIDs As String



    If oReportItem.Criteria1 = MTF Then
        sSQLWhere = "c.facility_ncid = " & oReportItem.CriteriaSel1.Item(1).NCID
    ElseIf oReportItem.Criteria1 = Clinic Then
        For i = 1 To oReportItem.CriteriaSel1.Count
            sClinicNCIDs = sClinicNCIDs & oReportItem.CriteriaSel1.Item(i).NCID
            If i <> oReportItem.CriteriaSel1.Count Then
                sClinicNCIDs = sClinicNCIDs & ","
            End If
        Next
        sSQLWhere = "c.ncid in (" & sClinicNCIDs & ")"
    End If

GetSQLWhere = True
GoTo CleanUp
ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "GetSQLWhere", "ReportOps", vbExclamation)
CleanUp:

End Function



Public Function GetSQLDateRangeSRTS(oReportItem As ReportItem, _
                                    sTableAlias As String, _
                                    sDateFilter As String) As Boolean
                    
On Error GoTo ErrHandler
GetSQLDateRangeSRTS = False

'-------------------------------
'- DATE RANGE
'-------------------------------
sDateFilter = sDateFilter & "(" & sTableAlias & ".order_status=3 and" & vbNewLine & _
                            "(" & sTableAlias & ".DT_TRANSACTION between " & GetOracleStartDate(oReportItem.DateFrom) & vbNewLine & _
                            "and " & GetOracleEndDate(oReportItem.DateTo) & "))"

GetSQLDateRangeSRTS = True
GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "GetSQLDateRangeSRTS", "CHCSII_Forms", vbExclamation
CleanUp:
    Exit Function
End Function

Public Function GetSQLDateRangeSRTS2(oReportItem As ReportItem, _
                                    sTableAlias As String, _
                                    sDateFilter As String) As Boolean
                    
On Error GoTo ErrHandler
GetSQLDateRangeSRTS2 = False

'-------------------------------
'- DATE RANGE
'-------------------------------
sDateFilter = "(" & sTableAlias & ".DT_TRANSACTION between " & GetOracleStartDate(oReportItem.DateFrom) & vbNewLine & _
                "and " & GetOracleEndDate(oReportItem.DateTo) & ")"

GetSQLDateRangeSRTS2 = True
GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "GetSQLDateRangeSRTS2", "CHCSII_Forms", vbExclamation
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

'Use these code for testing
Public Sub DebugPrintRecordSet(rs As ADODB.Recordset)
    If Not (rs.RecordCount > 0) Then Exit Sub   'Go Bye Bye
    rs.MoveFirst     'Go to the first record
    Dim fld As ADODB.Field
    Do While Not rs.EOF
        Debug.Print "Record number: " & rs.AbsolutePosition
        For Each fld In rs.Fields
            Debug.Print fld.Name & " = " & fld.Value & ";";
        Next
        Debug.Print
        rs.MoveNext
    Loop
    rs.MoveFirst
End Sub

