VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RepERTopUtilizerMTF"
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
    moCollRecSet.Add oRec, "ERTopUtilizerMTF"

'-- Create Report Object
'-----------------------------------------------------------------------------
    If Not CreateReport(sReportName & ".rpt", oReport, oCollFormulas, moCollRecSet) Then GoTo CleanUp

    BuildReport = True

GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepERTopUtilizersMTF.BuildReport", "CHCSII_Reports", vbExclamation
CleanUp:
    Set oCrystal = Nothing
    Set oRec = Nothing
    Set oPeerRec = Nothing
    
End Function

Private Function GetData(oReportItem As ReportItem, _
                           oRec As ADODB.Recordset, _
                           ByRef oCollFormulas As Collection) As Boolean

    Dim sStmt As String
    
    On Error GoTo ErrHandler
    GetData = False
    Dim sDateFilter As String
    Dim sFilter As String
    Dim sTableA As String
    Dim sQuestionnaireFilter As String
    
    Dim sHeader As String
    Dim iResponse As Integer
    
    Dim sStartDate As String
    Dim sEndDate As String
    Dim strDateInfo As String
    
    Dim sHint As String
    Dim sUnitNumber As String
    
    '-- Added tabbing to SQL to help debugging
    '--------------------------------------------
    Dim svbNewTab As String
    Dim svbNewTab2 As String
    Dim svbNewTab3 As String
    Dim svbNewTab4 As String
    
    svbNewTab = vbNewLine & vbTab
    svbNewTab2 = vbNewLine & vbTab & vbTab
    svbNewTab3 = vbNewLine & vbTab & vbTab & vbTab
    svbNewTab4 = vbNewLine & vbTab & vbTab & vbTab & vbTab
    
    '-- Get Filter Information for SQL
    '----------------------------------------------------------------------------
    sStartDate = GetOracleStartDate(GetFirstDayOfMonth(oReportItem.DateFrom))
    sEndDate = GetOracleEndDate(GetLastDayOfMonth(oReportItem.DateTo))
    
    sDateFilter = " (APPT_DATE_TIME >= " & sStartDate & " " & " AND APPT_DATE_TIME <= " & sEndDate & ") "
    
    If oReportItem.Criteria1 = MTF Then
        sFilter = "Facility_NCID = " & oReportItem.CriteriaSel1.Item(1).NCID
        sHint = " /*+ INDEX (APPOINTMENT_HISTORY APPOINTMENT_HISTORY_X2) */"
    ElseIf oReportItem.Criteria1 = MTFDMIS Then
        sFilter = "CLINIC_NCID IN (" & GetClinicNCIDsFromDMIS(oReportItem.CriteriaSel1) & ")"
        sHint = " /*+ INDEX (APPOINTMENT_HISTORY APPOINTMENT_HISTORY_X3) */"
    End If

    sStmt = "SELECT /*+ RULE */ *" & vbNewLine & _
            "FROM (" & svbNewTab & _
                "SELECT *" & svbNewTab & _
                "From" & svbNewTab2 & _
                    "(SELECT A.PATIENT_UNIT_NUMBER," & svbNewTab3 & _
                        "C.LAST_NAME || ', ' || C.FIRST_NAME || ' ' || C.MIDDLE_NAME AS NAME," & svbNewTab3 & _
                        "D.STREET_1 AS STREET, D.STREET_2 AS APT, D.CITY," & svbNewTab3 & _
                        "D.STATE, D.POST_CODE AS ZIP," & svbNewTab3 & _
                        "E.PHONE_TYPE_CID, E.AREA_CODE || '-' || SUBSTR(E.LOCAL_NUMBER, 1, 3) ||" & svbNewTab3 & _
                        "'-' || SUBSTR(E.LOCAL_NUMBER, 4) AS PHONE," & svbNewTab3 & _
                        "E.INTERNAL_EXTENSION, F.VALUE_TEXT AS PATCAT," & svbNewTab3 & _
                        "TO_CHAR(A.TOT_VISITS) AS TOTAL_VISTIS" & svbNewTab2 & _
                    "From /*+ ORDERED */ " & svbNewTab3

        sStmt = sStmt & _
                        "(SELECT " & sHint & svbNewTab4 & _
                            "PATIENT_UNIT_NUMBER, COUNT(*) AS TOT_VISITS" & svbNewTab3 & _
                        "From" & svbNewTab4 & _
                            "APPOINTMENT_HISTORY" & svbNewTab3 & _
                        "Where" & svbNewTab4 & _
                            sFilter & " And" & svbNewTab4 & _
                            sDateFilter & vbCrLf & _
                            "AND (APPT_STATUS = 'KEPT' OR  APPT_STATUS = 'LWOBS' OR  (APPT_STATUS = 'CANCEL' AND APPT_CANCEL_REASON = 'LWOBS'))" & vbCrLf & _
                            "AND (MEPRS_CODE LIKE 'BIA%' OR MEPRS_CODE LIKE 'BHI%')" & svbNewTab3 & _
                        "GROUP BY PATIENT_UNIT_NUMBER" & svbNewTab3 & _
                        "HAVING COUNT(*) > 1) A," & svbNewTab3 & _
                        "MMI_NAME C, MMI_ADDRESS D, MMI_PHONE E, MMI_GENERICDATA F" & svbNewTab2 & _
                    "Where" & svbNewTab3 & _
                        "A.Patient_Unit_Number = C.Unit_Number And C.NAME_INDEX = 1 And" & svbNewTab3 & _
                        "A.PATIENT_UNIT_NUMBER = D.UNIT_NUMBER AND D.ADDRESS_TYPE_CID = 199 AND" & svbNewTab3 & _
                        "A.PATIENT_UNIT_NUMBER = E.UNIT_NUMBER AND E.PHONE_TYPE_CID = 193 AND" & svbNewTab3 & _
                        "A.PATIENT_UNIT_NUMBER = F.UNIT_NUMBER AND F.TYPE_NCID = 14501721)) A" & vbNewLine & _
            "ORDER BY TOTAL_VISTIS, A.PATIENT_UNIT_NUMBER, A.PHONE_TYPE_CID DESC"


    Set oRec = goDBDAS.OpenRecordset(sStmt)
    
    '-- Remove Duplicates
    '--------------------------------------------------
    If oRec.RecordCount > 0 Then
        oRec.MoveFirst
        sUnitNumber = ""
        Dim i As Integer
        For i = 1 To oRec.RecordCount
            If sUnitNumber = oRec("PATIENT_UNIT_NUMBER").Value Then
                oRec.Delete
            Else
                sUnitNumber = oRec("PATIENT_UNIT_NUMBER").Value
            End If
            oRec.MoveNext
        Next
    End If
    
    If IsDate(oReportItem.DateFrom) Then strDateInfo = "From  " & Format(oReportItem.DateFrom, "mmm yyyy")
    If IsDate(oReportItem.DateTo) Then strDateInfo = strDateInfo & "  Through " & Format(oReportItem.DateTo, "mmm yyyy")
       
    On Error Resume Next
    oCollFormulas.Remove "DateInfo"
    oCollFormulas.Add CrystalQuote(strDateInfo), "DateInfo"
    
    If oReportItem.Criteria1 = MTFDMIS Then
        On Error Resume Next
        oCollFormulas.Remove ("Header2")
        For iResponse = 1 To oReportItem.CriteriaSel1.Count
            If Not sHeader = "" Then sHeader = sHeader & ", "
            sHeader = sHeader & oReportItem.CriteriaSel1(iResponse).MTF_NAME
        Next
        sHeader = "DMIS IDs selected: " & sHeader
        oCollFormulas.Add CrystalQuote(sHeader), "Header2"
    End If
    
    GetData = True
    GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepGenericClass.GetData", "CHCSII_Reports", vbExclamation
    'Resume Next
CleanUp:
    Exit Function
End Function

