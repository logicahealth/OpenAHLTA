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
Attribute VB_Name = "RepMammographyHaveNot"
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

Dim oRec As ADODB.Recordset
Dim sSponsorUnitNumber As String
Dim sUnitNumber As String
DoEvents

'-- Get Data
'----------------------------------------------------------------------------
    Set moCollRecSet = Nothing
    Set moCollRecSet = New Collection

    '-- Get Services Due
    '----------------------------------------------------------
    If Not GetMammographyHaveNot(oReportItem, oRec) Then GoTo CleanUp
    moCollRecSet.Add oRec, "Mammography"

'-- Set Formula Fields
'----------------------------------------------------------------------------
    oCollFormulas.Add CrystalQuote(oReportItem.CriteriaSel1.Item(1).Name), "Header1"


'-- Create Report Object
'-----------------------------------------------------------------------------
    If Not CreateReport(sReportName & ".rpt", oReport, oCollFormulas, moCollRecSet) Then GoTo CleanUp

BuildReport = True

GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepMammographyHaveNot.BuildReport", "CHCSII_Reports", vbExclamation
CleanUp:
    Set oRec = Nothing
    
End Function

Private Function GetMammographyHaveNot(oReportItem As ReportItem, _
                                       oRec As ADODB.Recordset) As Boolean

If Not DebugMode Then On Error GoTo ErrHandler
GetMammographyHaveNot = False

Dim sStmt As String
Dim sFromSQL As String
Dim sWhereSQL As String


    If Not GetMTForPCM_SQL(oReportItem, sFromSQL, sWhereSQL) Then GoTo CleanUp

'-- SELECT
    sStmt = "Select " & vbNewLine & _
        "distinct mn.last_name ||', ' || mn.first_name || ' ' || " & vbNewLine & _
        "mn.middle_name as Patient_name," & vbNewLine & _
        "substr(mg2.id_value,1,2) || '/' || SUBSTR(mg2.id_value,4) as SSN," & vbNewLine & _
        "SUBSTR(Trunc((sysdate -id.birth_date)/365),1,3) as Age," & vbNewLine & _
        "DECODE(hp.area_code,null,null, '(' || hp.area_code ||') ') || DECODE(hp.local_number,null, null, " & vbNewLine & _
        "SUBSTR(hp.local_number, 1, 3) || '-' || SUBSTR(hp.local_number,4,4)) as Home_phone," & vbNewLine & _
        "SUBSTR(DECODE(wp.area_code, null, null, '(' || wp.area_code ||') ') || " & vbNewLine & _
        "DECODE(wp.local_number,null, null, SUBSTR(wp.local_number, 1, 3) || '-' || " & vbNewLine & _
        "SUBSTR(wp.local_number,4,4)) || DECODE(wp.internal_extension, null, null," & vbNewLine & _
        "' Ext.' || wp.internal_extension), 1, 22) as Work_phone," & vbNewLine & _
        "ma.street_1 || ' ' || ma.street_2 || ' ' || ma.city || ', '||  ma.state || ' ' || ma.post_code as Address "

'-- FROM
    sStmt = sStmt & _
        "From mmi_generic_id mg2, " & vbNewLine & _
            "mmi_address ma, " & vbNewLine & _
            "mmi_phone hp, " & vbNewLine & _
            "mmi_phone wp, " & vbNewLine & _
            "mmi_name mn, " & vbNewLine & _
            "mmi_id id " & vbNewLine & _
            sFromSQL
            
'-- WHERE
    sStmt = sStmt & _
        "Where " & vbNewLine & _
        "mg2.id_type_ncid = 14511309 " & vbNewLine & _
        "and mg2.Unit_Number (+) = Id.Unit_Number " & vbNewLine & _
        "and mn.Unit_Number (+) = Id.Unit_Number " & vbNewLine & _
        "and ma.unit_number (+) = id.unit_number" & vbNewLine & _
        "and wp.phone_type_cid (+) = 194" & vbNewLine & _
        "and wp.unit_number (+) = id.unit_number" & vbNewLine & _
        "and hp.phone_type_cid (+) = 193" & vbNewLine & _
        "and hp.unit_number (+) = id.unit_number" & vbNewLine & _
        "and not exists" & vbNewLine
    
    sStmt = sStmt & _
        "(select * from event e" & vbNewLine & _
        "where event_subtype_ncid in" & vbNewLine & _
        "(124957,125243,125244,125245,125335,125338,125339,125340,125341,125342," & vbNewLine & _
        "125343,125344,125345,125346,125347,125349,125350,125351,125352,125353," & vbNewLine & _
        "125354,125355,125356,125357,125358,125360,125361,125362,125363,125364," & vbNewLine & _
        "125365,125366,125367,125368,125369,125371,126171,126172,127000,14501664," & vbNewLine & _
        "14511016,14511037,14583723,14583724,14583725,14603448,14603729,14603754," & vbNewLine & _
        "14613401,14613403,14613404,14613405,14613406,14613407,14613408,14613411," & vbNewLine & _
        "14613415,14613466,14613468,14613860,14614716,14614873,14614874,14614923," & vbNewLine & _
        "14614924,14614925,14614930,14614931,14614932,14614933,14614934,14614935," & vbNewLine & _
        "14614936,14615009,14615010,14615011,149300,162240,162241,162242,162243,182767," & vbNewLine & _
        "182919,182920,184898,213049,213334,213406,213457,213458,213459,213460,213461," & vbNewLine & _
        "213808,84550,84551) and " & _
                        "e.Event_Type_NCID = 507 and " & _
                        "e.unit_number = id.unit_number) and "

    '-- Check if patient is a Female >= 21 or Female >= 64
    sStmt = sStmt & _
            "((((" & GetOracleEndDate(oReportItem.DateTo) & " - id.birth_date)/365) >= 50.0) and " & _
            "id.sex = 'F')"

'    '-- Check if patient has records at specified MTF or is patient of specified PCM
    sStmt = sStmt & _
            " and " & sWhereSQL

    Set oRec = goDBDAS.OpenRecordset(sStmt)


GetMammographyHaveNot = True
GoTo CleanUp
ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "RepMammographyHaveNot.GetMammographyHaveNot", "CHCSII_Reports", vbExclamation)
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

