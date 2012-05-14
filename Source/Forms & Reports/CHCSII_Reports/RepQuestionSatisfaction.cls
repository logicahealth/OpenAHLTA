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
Attribute VB_Name = "RepQuestionSatisfaction"
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
    Dim oPeerRec As ADODB.Recordset
    Dim sSponsorUnitNumber As String
    Dim sUnitNumber As String
    DoEvents
    
    '-- Get Data    '----------------------------------------------------------------------------
    Set moCollRecSet = Nothing
    Set moCollRecSet = New Collection

    '-- Get Candidates
    '----------------------------------------------------------
    If Not GetData(oReportItem, oRec, False) Then GoTo CleanUp         ' data for provider
    If oReportItem.Criteria1 = MTF Then
        moCollRecSet.Add oRec, "answersMTF"
    ElseIf oReportItem.Criteria1 = MTFDMIS Then
        moCollRecSet.Add oRec, "answersMTF"
    ElseIf oReportItem.Criteria1 = Clinic Or oReportItem.Criteria1 = ClinicSingleSelect Then
        If oReportItem.Criteria2 = Provider Then
            moCollRecSet.Add oRec, "answers"
            If Not GetData(oReportItem, oPeerRec, True) Then GoTo CleanUp      ' data for peer
            moCollRecSet.Add oPeerRec, "Peer"
        Else
            moCollRecSet.Add oRec, "answersClinic"
        End If
    End If

'-- Create Report Object
'-----------------------------------------------------------------------------
    If Not CreateReport(sReportName & ".rpt", oReport, oCollFormulas, moCollRecSet) Then GoTo CleanUp

    BuildReport = True

GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepQuestionSatisfaction.BuildReport", "CHCSII_Reports", vbExclamation
CleanUp:
    Set oCrystal = Nothing
    Set oRec = Nothing
    Set oPeerRec = Nothing
    
End Function

Private Function GetData(oReportItem As ReportItem, _
                           oRec As ADODB.Recordset, Peer As Boolean) As Boolean

    Dim sStmt As String
    Dim sTableA As String
    Dim sTableB As String
    Dim sTableC As String
    Dim iReportParam As Integer
    
    On Error GoTo ErrHandler
    GetData = False

    Dim sDateFilter As String
    Dim sFilter As String
    Dim sQuestionnaireFilter As String
    
    If oReportItem.Criteria2 = Provider Then
      If Not oReportItem.Criteria3 = QuestionnaireAnswer Then GoTo CleanUp
      If Not (oReportItem.CriteriaSel3(1).ADefType = 2 Or _
        oReportItem.CriteriaSel3(1).ADefType = 5 Or _
        oReportItem.CriteriaSel3(1).ADefType = 6) Then
        MsgBxARMd "The satisfaction report can only be run for multiple choice, or multiple selection answers.", vbOKOnly + vbInformation
        GoTo CleanUp
      End If
    Else
      If Not oReportItem.Criteria2 = QuestionnaireAnswer Then GoTo CleanUp
      If Not (oReportItem.CriteriaSel2(1).ADefType = 2 Or _
        oReportItem.CriteriaSel2(1).ADefType = 5 Or _
        oReportItem.CriteriaSel2(1).ADefType = 6) Then
        MsgBxARMd "The satisfaction report can only be run for multiple choice, or multiple selection answers.", vbOKOnly + vbInformation
        GoTo CleanUp
      End If
    End If
    '-- Get Filter Information for SQL
    '----------------------------------------------------------------------------
    If Not Peer Then '--- SQL not for Peer Provider Mean
        sDateFilter = sDateFilter & " (qr.Start_Date_Time >= " & GetOracleStartDate(oReportItem.DateFrom) & vbCrLf & _
                    " AND qr.End_Date_Time <= " & GetOracleEndDate(oReportItem.DateTo) & ") "
    
        If oReportItem.Criteria1 = MTF Then
            If Not GetSQLCriteria(oReportItem, 1, sFilter) Then GoTo CleanUp
            sFilter = Replace(sFilter, "a.facility_ncid", "ms.FACILITY_CID")
        ElseIf oReportItem.Criteria1 = Clinic Or oReportItem.Criteria1 = ClinicSingleSelect Then
            If oReportItem.Criteria2 = Provider Then
                sFilter = "qr.PCM = " & oReportItem.CriteriaSel2(1).NCID
            Else
                sFilter = "qr.clinic_ncid in ("
                For iReportParam = 1 To oReportItem.CriteriaSel1.Count
                      If iReportParam > 1 Then sFilter = sFilter & ", "
                      sFilter = sFilter & oReportItem.CriteriaSel1(iReportParam).NCID
                Next
                sFilter = sFilter & ")"
            End If
        ElseIf oReportItem.Criteria1 = MTFDMIS Then
            sFilter = GetClinicNCIDsFromDMIS(oReportItem.CriteriaSel1)
            If sFilter = "" Then Exit Function
            sFilter = "qr.clinic_ncid in (" & sFilter & ")"
        End If
        
        If oReportItem.Criteria2 = Provider Then
            sQuestionnaireFilter = "ar.seq = " & oReportItem.CriteriaSel3(1).Seq & " AND " & vbCrLf & _
                    "qr.QDef_ID = " & oReportItem.CriteriaSel3(1).QDefID & " AND "
        Else
            If Not oReportItem.Criteria2 = QuestionnaireAnswer Then GoTo CleanUp
            sQuestionnaireFilter = "ar.seq = " & oReportItem.CriteriaSel2(1).Seq & " AND " & vbCrLf & _
                        "qr.QDef_ID = " & oReportItem.CriteriaSel2(1).QDefID & " AND "
        End If

        
        '-- Select
        If oReportItem.Criteria1 = Clinic Or oReportItem.Criteria1 = ClinicSingleSelect Then
            If oReportItem.Criteria2 = Provider Then
                sStmt = "SELECT qd.Question_Text, qd.choice, qr.PCM" & vbCrLf & _
                    "FROM QDEF_CHOICES qd, ARESP ar, QRESP qr "
            Else
                sStmt = "SELECT qd.Question_Text, qd.choice, qr.PCM, " _
                    & "clc.ncid as clinic_ncid, clc.name as clinic_name" & vbCrLf & _
                    "FROM QDEF_CHOICES qd, ARESP ar, QRESP qr, CLINIC clc"
            End If
        ElseIf oReportItem.Criteria1 = MTFDMIS Then
            sStmt = "SELECT qd.Question_Text, qd.choice, qr.PCM, " _
                & "clc.ncid as clinic_ncid, clc.name as clinic_name" & vbCrLf & _
                "FROM QDEF_CHOICES qd, ARESP ar, QRESP qr, CLINIC clc"
        ElseIf oReportItem.Criteria1 = MTF Then
            sStmt = "SELECT qd.Question_Text, qd.choice, qr.PCM" & vbCrLf & _
                "FROM QDEF_CHOICES qd, ARESP ar, QRESP qr, MMI_SITE ms "
        End If
        
        sStmt = sStmt & vbCrLf
        
        sStmt = sStmt & "WHERE ar.QResp_ID = qr.QResp_ID AND " & vbCrLf & _
            "qd.QDef_ID = qr.QDef_ID AND " & vbCrLf & _
            "qd.Version = qr.Version AND " & vbCrLf & _
            "qd.Seq = ar.Seq AND " & vbCrLf & _
            "TO_CHAR(qd.Choice) = ar.answer AND " & vbCrLf & _
            sQuestionnaireFilter & vbCrLf & sDateFilter & " AND " & vbCrLf
        
        If oReportItem.Criteria1 = MTF Then
           sStmt = sStmt & "qr.UNIT_NUMBER = ms.UNIT_NUMBER AND " & vbCrLf
        ElseIf ((oReportItem.Criteria1 = Clinic Or oReportItem.Criteria1 = ClinicSingleSelect) And oReportItem.Criteria2 <> Provider) Or oReportItem.Criteria1 = MTFDMIS Then
            sStmt = sStmt & "qr.clinic_ncid = clc.ncid AND " & vbCrLf
        End If
        
        sStmt = sStmt & sFilter & vbCrLf

    Else '---- for Peer Provider Mean

        sDateFilter = sDateFilter & " (qr.Start_Date_Time >= " & GetOracleStartDate(oReportItem.DateFrom) & vbCrLf & _
                    " AND qr.End_Date_Time <= " & GetOracleEndDate(oReportItem.DateTo) & ") "

        sFilter = "qr.PCM <> " & oReportItem.CriteriaSel2(1).NCID

        sFilter = sFilter & " AND " & vbCrLf & _
                  "qr.CLINIC_NCID  = " & oReportItem.CriteriaSel1(1).NCID

        If Not oReportItem.Criteria3 = QuestionnaireAnswer Then GoTo CleanUp
        sQuestionnaireFilter = "ar.seq = " & oReportItem.CriteriaSel3(1).Seq & " AND " & vbCrLf & _
                    "qr.QDef_ID = " & oReportItem.CriteriaSel3(1).QDefID & " AND "

        '---------------------------------------------------------
        ' SQL for Table A
        ' Get the answer count by choice and provider
        '---------------------------------------------------------
        sTableA = "SELECT qd.Question_Text, qd.choice, qr.PCM, COUNT(ALL qr.PCM) AS ANSWER_COUNT" & vbCrLf & _
                  "FROM QDEF_CHOICES qd, ARESP ar, QRESP qr " & vbCrLf

        sTableA = sTableA & "WHERE ar.QResp_ID = qr.QResp_ID AND " & vbCrLf & _
            "qd.QDef_ID = qr.QDef_ID AND " & vbCrLf & _
            "qd.Version = qr.Version AND " & vbCrLf & _
            "qd.Seq = ar.Seq AND " & vbCrLf & _
            "TO_CHAR(qd.Choice) = ar.answer AND " & vbCrLf & _
            sQuestionnaireFilter & vbCrLf & _
            sDateFilter & " AND " & vbCrLf & _
            sFilter & vbCrLf

        sTableA = sTableA & " group by qd.choice, qd.Question_Text, qr.PCM" & vbCrLf
        
        '------------------------------------------------------------------------------------
        ' SQL for Table B
        ' Get the total answer count by provider
        '------------------------------------------------------------------------------------
        sTableB = "SELECT QR1.PCM, COUNT(ALL QR1.PCM) AS TOTAL_BY_PROV" & vbCrLf & _
                  "FROM QDEF_CHOICES qd1, ARESP ar1, QRESP qr1 " & vbCrLf

        sTableB = sTableB & "WHERE ar1.QResp_ID = qr1.QResp_ID AND " & vbCrLf & _
            "qd1.QDef_ID = qr1.QDef_ID AND " & vbCrLf & _
            "qd1.Version = qr1.Version AND " & vbCrLf & _
            "qd1.Seq = ar1.Seq AND " & vbCrLf & _
            "qd1.Choice = ar1.answer AND " & vbCrLf & _
            Replace(Replace(sQuestionnaireFilter, "qr.", "qr1."), "ar.", "ar1.") & vbCrLf & _
            Replace(sDateFilter, "qr.", "qr1.") & " AND " & vbCrLf & _
            Replace(Replace(sFilter, "qr.", "qr1."), "CLC", "CLC1") & vbCrLf

        sTableB = sTableB & " group by qr1.PCM" & vbCrLf
        
        '------------------------------------------------------------------------------------
        ' SQL for Table C
        ' Join table A and table B to get the Percentage for each provider and each choice
        '-----------------------------------------------------------------------------------
        sTableC = "SELECT A.Question_Text, A.choice,  A.ANSWER_COUNT/B.TOTAL_BY_PROV AS PERCENT, " & _
                        "B.TOTAL_BY_PROV" & vbCrLf & _
                  "FROM (" & sTableA & ") A, (" & sTableB & ") B" & vbCrLf & _
                  "WHERE A.PCM = B.PCM"

        '------------------------------------------------------------------------------------
        ' Finally calculate the standard deviations
        '-----------------------------------------------------------------------------------
        sStmt = "SELECT C.Question_Text, C.choice, TO_CHAR(ROUND(AVG(C.PERCENT), 4)) AS MEAN, " & vbCrLf & _
                    "TO_CHAR(ROUND(STDDEV(C.PERCENT), 4)) AS STD_DEV2,  " & _
                    "TO_CHAR(SUM(C.TOTAL_BY_PROV)) AS TOTAL" & vbCrLf & _
                "FROM (" & sTableC & ") C" & vbCrLf & _
                "GROUP BY  C.Question_Text, C.choice"

    End If
                
    Set oRec = goDBDAS.OpenRecordset(sStmt)
    GetData = True
    GoTo CleanUp

ErrHandler:
    'Resume Next
    gobjshared.ShowVBError Err.Number, Err.Description, "RepQuestionSatisfaction.GetData", "CHCSII_Reports", vbExclamation
CleanUp:
    Exit Function
End Function

