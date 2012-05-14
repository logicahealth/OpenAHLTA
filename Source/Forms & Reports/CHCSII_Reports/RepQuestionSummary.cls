VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RepQuestionSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private moCollRecSet As Collection
Private moCollFormulas As Collection

Private moSQL As CHCSII_SQL.SqlOps
Private moDB As CHCSII_DataAccess.CHCSII_DAS


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
    If Not GetQuestionnaireSummary(oReportItem, oRec) Then GoTo CleanUp
    moCollRecSet.Add oRec, "answers"

'-- Create Report Object
'-----------------------------------------------------------------------------
    If Not CreateReport(sReportName & ".rpt", oReport, oCollFormulas, moCollRecSet) Then GoTo CleanUp

BuildReport = True

GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepRegistryCandidates.BuildReport", "CHCSII_Reports", vbExclamation
CleanUp:
    Set oCrystal = Nothing
    Set oRec = Nothing
    
End Function

Private Function GetQuestionnaireSummary(oReportItem As ReportItem, _
                           oRec As ADODB.Recordset) As Boolean

    Dim sStmt As String
    
    On Error GoTo ErrHandler
    GetQuestionnaireSummary = False

    '-- Select
    sStmt = "Select " & vbNewLine & _
                "ar.seq as QuestionNumber, " & vbNewLine & _
                "ar.answer, " & vbNewLine & _
                "Trim(mn.last_name ||', ' || mn.first_name || ' ' || " & vbNewLine & _
                "mn.middle_name) as Full_Name," & vbNewLine & _
                "fs.id_value as FMP_SPONSOR_SSN, " & vbNewLine & _
                "cr.research_id, " & vbNewLine & _
                "p.NAME as PCM_Name, " & vbNewLine & _
                "ad.adef_type, " & vbNewLine & _
                "ad.Params, " & _
                "'  ' as Final_Answer, " & _
                "ad.QDEF_ID, " & vbNewLine & _
                "ad.Version " & vbNewLine

    '-- From
    sStmt = sStmt & _
            "From " & vbNewLine & _
                "adef ad, " & vbNewLine & _
                "CPGDEF cd, " & vbNewLine & _
                "aresp ar, " & vbNewLine & _
                "qresp qr, " & vbNewLine & _
                "mmi_generic_id fs, " & vbNewLine & _
                "mmi_generic_id pm, " & vbNewLine & _
                "provider p, " & vbNewLine & _
                "mmi_name mn, " & vbNewLine & _
                "cpgregistry cr, " & vbNewLine & _
                "mmi_id id " & vbNewLine

    '-- Where
    sStmt = sStmt & _
        "Where " & vbNewLine & _
            "ad.seq = ar.seq and " & vbNewLine & _
            "ad.version =  qr.version and " & vbNewLine & _
            "ad.qdef_id = qr.qdef_id and " & vbNewLine & _
            "cd.name = '" & CPG_DIABETES_DEF_NAME & "' and " & vbNewLine & _
            "cd.cpg_id = cr.cpg_id and cr.unit_number = id.unit_number and " & vbNewLine & _
            "ar.qresp_id = qr.qresp_id and " & vbNewLine & _
            "qr.qresp_id in " & vbNewLine & _
            "(select max(qr2.qresp_id) " & vbNewLine & _
            "from qdef qd2, qresp qr2 " & vbNewLine & _
            "where " & vbNewLine & _
            "UPPER(qd2.NAME) = '" & CPG_DIABETES_ANNUAL_Q_NAME & "' and " & vbNewLine & _
            "qd2.qdef_id = qr2.qdef_id and " & vbNewLine & _
            "qr2.unit_number = id.unit_number) and " & vbNewLine & _
            "qr.unit_number = id.unit_number and " & vbNewLine & _
            "fs.id_type_ncid = 14511309 and " & vbNewLine & _
            "fs.unit_number = id.unit_number and " & vbNewLine & _
            "p.pcm_code = pm.id_value and " & vbNewLine & _
            "pm.id_type_ncid = 14501724 and " & vbNewLine & _
            "pm.unit_number = id.unit_number and " & vbNewLine & _
            "mn.name_type_cid = 187 and " & vbNewLine & _
            "mn.unit_number = id.unit_number and " & vbNewLine

    sStmt = sStmt & _
            "id.unit_number in " & _
            "(select mgi.unit_number " & vbNewLine & _
            "from  mmi_generic_id mgi " & vbNewLine & _
            "where mgi.id_type_ncid = 14501724 and " & vbNewLine & _
            "mgi.id_value = '" & oReportItem.CriteriaSel1.Item(1).PCMCode & "') " & vbNewLine

    '-- Order By
    sStmt = sStmt & "Order by p.Name, QuestionNumber, Full_Name, research_id, QuestionNumber "

    Set oRec = moDB.OpenRecordset(sStmt)

    '-- Get Text for Answer (shouldn't have to do this)
    '-------------------------------------------------------------
    If Not UpdateQuestionnaireFields(oRec) Then GoTo CleanUp
    
    GetQuestionnaireSummary = True
    GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepCPGQuestionAnnual.GetQuestionnaireSummary", "CHCSII_Reports", vbExclamation
CleanUp:
    Exit Function
End Function


Private Sub Class_Initialize()
    Set moSQL = New CHCSII_SQL.SqlOps
    Set moDB = New CHCSII_DAS
End Sub


Private Sub Class_Terminate()
    Set moSQL = Nothing
    Set moDB = Nothing
End Sub


