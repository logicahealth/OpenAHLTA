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
Attribute VB_Name = "RepERTopUtilizerPRV"
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
    
    '-- Get Data
    '----------------------------------------------------------------------------
    Set moCollRecSet = Nothing
    Set moCollRecSet = New Collection

    '-- Get Data for the MTF
    '----------------------------------------------------------
    If Not GetData(oReportItem, oRec, oCollFormulas) Then GoTo CleanUp
    moCollRecSet.Add oRec, "ERTopUtilizerPRV"

'-- Create Report Object
'-----------------------------------------------------------------------------
    If Not CreateReport(sReportName & ".rpt", oReport, oCollFormulas, moCollRecSet) Then GoTo CleanUp

    BuildReport = True

GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepERTopUtilizersPRV.BuildReport", "CHCSII_Reports", vbExclamation
CleanUp:
    Set oCrystal = Nothing
    Set oRec = Nothing
    Set oPeerRec = Nothing
    
End Function

Private Function GetData(oReportItem As ReportItem, _
                           oRec As ADODB.Recordset, _
                           ByRef colFormulas As Collection) As Boolean

    Dim sStmt As String
    
    On Error GoTo ErrHandler
    GetData = False
    Dim sDateFilter As String
    Dim sFilter As String
    Dim sTableA As String
    Dim sPCMCode As String
    
    Dim sStartDate As String
    Dim sEndDate As String
    Dim strDateInfo                    As String
    
    
    '--- Get PCM Code
    sPCMCode = oReportItem.CriteriaSel1.Item(1).PCMCode
    
    If sPCMCode = "" Then
        MsgBxARMd ("The provider you selected is not a PCM!")
        GoTo CleanUp
    End If
    
    '-- Get Filter Information for SQL
    '----------------------------------------------------------------------------
    sStartDate = GetOracleStartDate(GetFirstDayOfMonth(oReportItem.DateFrom))
    sEndDate = GetOracleEndDate(GetLastDayOfMonth(oReportItem.DateTo))
    
    sDateFilter = sDateFilter & " (APPT_DATE_TIME >= " & sStartDate & vbCrLf & _
                                " AND APPT_DATE_TIME <= " & sEndDate & ") "
    
    'INDEX (B MMI_ID_PK)
    
    sTableA = "SELECT" & vbCrLf & _
              "/*+ ORDERED USE_NL(A) USE_NL(B) USE_NL(C ) USE_NL(D) USE_NL(E) USE_NL(F) " & vbCrLf & _
              "INDEX (C MMI_NAME_PK) INDEX (D MMI_ADDRESS_PK) INDEX (E MMI_PHONE_PK) INDEX (F MMI_GENERICDATA_X3) */" & vbCrLf & _
              "PATIENT_UNIT_NUMBER, COUNT(*) AS TOT_VISITS " & vbCrLf & _
              "FROM APPOINTMENT_HISTORY" & vbCrLf & _
              "WHERE " & sDateFilter & vbCrLf & _
                    "AND (MEPRS_CODE like 'BIA%' or MEPRS_CODE like 'BHI%')" & vbCrLf & _
                    "AND (APPT_STATUS = 'KEPT' OR  APPT_STATUS = 'LWOBS' OR  (APPT_STATUS = 'CANCEL' AND APPT_CANCEL_REASON = 'LWOBS'))" & vbCrLf & _
              "GROUP BY PATIENT_UNIT_NUMBER" & vbCrLf & _
              "HAVING COUNT(*) > 1" & vbCrLf
    sTableA = "SELECT /*+ USE_NL(A) USE_NL(B) */ A.PATIENT_UNIT_NUMBER, A.TOT_VISITS" & vbCrLf & _
              "FROM (" & sTableA & ") A, (" & vbCrLf & _
                    "SELECT UNIT_NUMBER" & vbCrLf & _
                    "FROM PATIENT_GENERAL_CODES " & vbCrLf & _
                    "WHERE PCM_CODE = '" & sPCMCode & "'" & vbCrLf & _
                    ") B" & vbCrLf & _
              "WHERE A.PATIENT_UNIT_NUMBER = B.UNIT_NUMBER"
    
    sStmt = "SELECT A.PATIENT_UNIT_NUMBER, SUBSTR(B.SSN, 1, 3) || '-' || SUBSTR(B.SSN, 4, 2) || '-' || SUBSTR(B.SSN, 6) AS SSN, " & vbCrLf & _
                "C.LAST_NAME || ', ' || C.FIRST_NAME || ' ' || C.MIDDLE_NAME AS NAME," & vbCrLf & _
                "D.STREET_1 AS STREET, D.STREET_2 AS APT, D.CITY, D.STATE, D.POST_CODE AS ZIP," & vbCrLf & _
                "E.PHONE_TYPE_CID, E.AREA_CODE || '-' || SUBSTR(E.LOCAL_NUMBER, 1, 3) || '-' || SUBSTR(E.LOCAL_NUMBER, 4) AS PHONE, " & vbCrLf & _
                "E.INTERNAL_EXTENSION, F.VALUE_TEXT AS PATCAT, to_char(A.TOT_VISITS) as TOTAL_VISTIS" & vbCrLf & _
            "FROM (" & sTableA & ") A,   MMI_ID B, MMI_NAME C, MMI_ADDRESS D, MMI_PHONE E, MMI_GENERICDATA F" & vbCrLf & _
            "WHERE  A.PATIENT_UNIT_NUMBER = B.UNIT_NUMBER" & vbCrLf & _
                "AND B.UNIT_NUMBER = C.UNIT_NUMBER AND C.NAME_INDEX = 1" & vbCrLf & _
                "AND B.UNIT_NUMBER = D.UNIT_NUMBER AND D.ADDRESS_TYPE_CID = 199" & vbCrLf & _
                "AND B.UNIT_NUMBER = E.UNIT_NUMBER AND E.PHONE_TYPE_CID = 193" & vbCrLf & _
                "AND B.UNIT_NUMBER = F.UNIT_NUMBER AND F.TYPE_NCID = 14501721" & vbCrLf & _
            "ORDER BY TOT_VISITS DESC, NAME ASC" & vbCrLf


   Set oRec = goDBDAS.OpenRecordset(sStmt)
    
    On Error Resume Next
    '--- From/To date string in the report formular
    If IsDate(oReportItem.DateFrom) Then strDateInfo = "From  " & Format(oReportItem.DateFrom, "mmm yyyy") & "   "
    If IsDate(oReportItem.DateTo) Then strDateInfo = strDateInfo & "Through  " & Format(oReportItem.DateTo, "mmm yyyy")
    colFormulas.Add CrystalQuote(strDateInfo), "DateInfo"
    
    GetData = True
    GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepGenericClass.GetData", "CHCSII_Reports", vbExclamation
CleanUp:
    Exit Function
End Function

