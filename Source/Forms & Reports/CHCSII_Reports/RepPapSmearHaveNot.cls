VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RepPapSmearHaveNot"
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
    If Not GetPapSmearHaveNot(oReportItem, oRec) Then GoTo CleanUp
    moCollRecSet.Add oRec, "papsmear"

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

Private Function GetPapSmearHaveNot(oReportItem As ReportItem, _
                                    oRec As ADODB.Recordset) As Boolean

If Not DebugMode Then On Error GoTo ErrHandler
GetPapSmearHaveNot = False

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
        "and " & vbNewLine
    
    '-- Check for PAP SMEAR Result
    sStmt = sStmt & _
            "(Not exists " & vbNewLine & _
                "(select e.event_id " & vbNewLine & _
                " from event e " & vbNewLine & _
                " where " & vbNewLine & _
                        GetOracleEndDate(oReportItem.DateTo) & " >= e.EVENT_START_GMTIME and " & vbNewLine & _
                        "e.event_subtype_ncid = 90269 and e.event_type_ncid = 485 and " & vbNewLine & _
                        "e.unit_number = id.unit_number)) and " & vbNewLine

    '-- Check if patient is a Female >= 21 or Female >= 64
    sStmt = sStmt & _
            "((((((" & GetOracleEndDate(oReportItem.DateTo) & " - id.birth_date)/365) >= 21.0) and " & vbNewLine & _
            "((" & GetOracleEndDate(oReportItem.DateTo) & " - id.birth_date)/365) <= 64.0)) and " & vbNewLine & _
            "id.sex = 'F') " & vbNewLine

'    '-- Check if patient has records at specified MTF or is patient of specified PCM
    sStmt = sStmt & _
            " and " & sWhereSQL

    Set oRec = goDBDAS.OpenRecordset(sStmt)


GetPapSmearHaveNot = True
GoTo CleanUp
ErrHandler:
    If Err = cTuxedoTimeOut Then
        MsgBxARMd "The selected report has timed out, re-run this report again.", vbOKOnly + vbInformation
        Err.Clear
    Else
        Call gobjshared.ShowVBError(Err.Number, Err.Description, "RepPapSmearHaveNot.GetPapSmearHaveNot", "CHCSII_Reports", vbExclamation)
    End If
CleanUp:
End Function
Private Sub Class_Initialize()
    '--- SCR-ARM-16148 MSolano 9/10/01; COMMON TO MULTIPLE ARM-SCRs
    '--- Init routine should be called from first top level creatable coclass
    '--- It can be called multiple times, but will actually run only once.
    '--- Instrument all creatable classes in case dll's multiple entry
    '--- points are actually used
     Call ARM_InitAllTransactionsForCHCSIIReports
End Sub


