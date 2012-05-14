VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FormInpatientList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private moCollRecSet As Collection
Private moCollFormulas As Collection

Const sMODULE_NAME As String = "FormInpatientList"

'Save this for reference
'Const sRegNum_Hdr As String = "Register Number"
'Const sSpecialty_Hdr As String = "Specialty"
'Const sCapacity_Hdr As String = "% Capacity"
'Const sBed_Hdr As String = "Bed"
'Const sRank_Hdr As String = "Rank"
'Const sPatient_Hdr As String = "Patient"
'Const sFMPSSN_Hdr As String = "FMP/Sponsor SSN"
'Const sUnit_Hdr As String = "Unit"
'Const sPatService_Hdr As String = "Inpatient Service"
'Const sAdmittingDiagnosis_Hdr As String = "Admitting Diagnosis"
'Const sAdmissionDate_Hdr As String = "Admission Date"
'Const sAdmissionComments_Hdr As String = "Admission Comments"
'Const sDay_Hdr As String = "Day#"
'Const sLocation_Hdr As String = "Location"
'Const sStatus_Hdr As String = "Status"
'Const sDischargeDate_Hdr As String = "Discharge Date"
'Const sDisposition_Hdr As String = "Disposition"
'Const sDischargeDiagnosis_Hdr As String = " Discharge Diagnosis"
'Const sAdmittingProvider_Hdr As String = "Admitting Provider"
'Const sAttendingProvider_Hdr As String = "Provider"
'Const sTransferDate_Hdr As String = "Transfer Date"
'Const sService_Hdr As String = "Service"

Const sRegNum_ColNm As String = "RegNum"
Const sSpecialty_ColNm As String = "Specialty"
Const sCapacity_ColNm As String = "Capacity"
Const sBed_ColNm As String = "Bed"
Const sRank_ColNm As String = "Rank"
Const sPatient_ColNm As String = "Patient"
Const sFMPSSN_ColNm As String = "FMPSSN"
Const sUnit_ColNm As String = "Unit"
Const sPatService_ColNm As String = "PatService"
Const sAdmittingDiagnosis_ColNm As String = "AdmittingDiagnosis"
Const sAdmissionDate_ColNm As String = "AdmissionDate"
Const sAdmissionComments_ColNm As String = "AdmissionComments"
Const sDay_ColNm As String = "Day"
Const sLocation_ColNm As String = "Location"
Const sStatus_ColNm As String = "Status"
Const sDischargeDate_ColNm As String = "DischargeDate"
Const sDisposition_ColNm As String = "Disposition"
Const sDischargeDiagnosis_ColNm As String = " DischargeDiagnosis"
Const sAdmittingProvider_ColNm As String = "AdmittingProvider"
Const sAttendingProvider_ColNm As String = "AttendingProvider"
Const sTransferDate_ColNm As String = "TransferDate"
Const sService_ColNm As String = "Service"

Const sFieldNmBase As String = "GENERIC_"       'Generic Column in report
Const sRptColHdrBase As String = "COLHDR_"      'Column Header in report
'

'Public Sub DebugPrintRecordSet(rs As ADODB.Recordset)
'    If Not (rs.RecordCount > 0) Then Exit Sub   'Go Bye Bye
'    rs.MoveFirst     'Go to the first record
'    Dim fld As ADODB.Field
'    Do While Not rs.EOF
'        Debug.Print "Record number: " & rs.AbsolutePosition
'        For Each fld In rs.Fields
'            Debug.Print fld.Name & " = " & fld.Value & ";";
'        Next
'        Debug.Print
'        rs.MoveNext
'    Loop
'    rs.MoveFirst
'End Sub

Public Function BuildInpatientListReport(oReport As CRAXDRT.Report, _
                                        sTitle As String, _
                                        colInpatient As Collection, _
                                        colColumns As Collection) As Boolean
On Error GoTo ErrHandler
BuildInpatientListReport = False

Dim oCrystal As Crystal
Dim oRec1 As ADODB.Recordset
Dim lLineProcessing As Long
Dim sLineProcessing As String
Dim vWords As Variant
Dim vWord As Variant

DoEvents

'-- Get Data
'----------------------------------------------------------------------------
    Set moCollRecSet = Nothing
    Set moCollRecSet = New Collection

    '-- Get Appointment Data
    '------------------------------------------------------------------------
   
   Set oRec1 = New ADODB.Recordset
   Call AppendsResultSetFields(oRec1)           'Set up and record set's fields
   Call BuildRSFrom2Collections(oRec1, colInpatient, colColumns)
   
   moCollRecSet.Add oRec1, "Generic_Table"

'-- Set Formulas
'-----------------------------------------------------------------------------
    Set moCollFormulas = Nothing
    Set moCollFormulas = New Collection
    moCollFormulas.Add "'" & Format(Now, gsDefaultFormat & "  HHMM") & "'", "TodaysDateTime"

'hwcc 1/10/2004 -
'This whole thing is use for processing a Title larger than 254 characters
' quite a piece of work but do we really needs it?\
'  I do so research on the reality of a Title large than 254 char and go from there
'   Another item in the Laundry list
    'moCollFormulas.Add "'" & sTitle, "Title"
    If Len(sTitle) <= 254 Then
       moCollFormulas.Add "'" & Replace(sTitle, "'", "''") & "'", "Title"
    Else
        vWords = Split(sTitle, " ")
        lLineProcessing = 1
        For Each vWord In vWords
            If Len(sLineProcessing) + Len(vWord) + 1 > 254 Then 'Add one for space
               Select Case lLineProcessing
                    Case 1
                        moCollFormulas.Add "'" & Replace(sLineProcessing, "'", "''") & "'", "Title"
                    Case Else
                        moCollFormulas.Add "'" & Replace(sLineProcessing, "'", "''") & "'", "Title" & CStr(lLineProcessing)
                End Select
                sLineProcessing = vWord
                lLineProcessing = lLineProcessing + 1
            Else
                sLineProcessing = sLineProcessing & " " & vWord
            End If
        Next
        If Len(sLineProcessing) > 0 Then
            'Append remaining text
            Select Case lLineProcessing
                 Case 1
                     moCollFormulas.Add "'" & Replace(sLineProcessing, "'", "''") & "'", "Title"
                 Case Else
                     moCollFormulas.Add "'" & Replace(sLineProcessing, "'", "''") & "'", "Title" & CStr(lLineProcessing)
             End Select
        End If
    End If
    
    Call BuildDynamicColumnHdr(moCollFormulas, colColumns)
    
'-- Create Report Object
'-----------------------------------------------------------------------------
    Set oCrystal = New Crystal
    Set oReport = Nothing
    If Not oCrystal.ReportOpen("InpatientList.rpt", oReport) Then GoTo CleanUp
    If Not oCrystal.AssignRecordSets(moCollRecSet, oReport) Then GoTo CleanUp
    If Not oCrystal.AssignFormulas(moCollFormulas, oReport) Then GoTo CleanUp

BuildInpatientListReport = True
GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "Form.BuildHospitalSetupReport", "CHCSII_Forms", vbExclamation
    'Resume Next
CleanUp:
    Set oRec1 = Nothing
    Set oCrystal = Nothing
    Exit Function
End Function

Private Sub BuildDynamicColumnHdr(cFormulas As Collection, cOrdinal As Collection)
    Dim oColObj As Object
    Dim sColHdrName As String
    
    For Each oColObj In cOrdinal
        sColHdrName = NULLStr
        If oColObj.columnOrdinal <= 25 Then   'Our Generic Report can only take 25 columns
            sColHdrName = sRptColHdrBase & oColObj.columnOrdinal  'Assume that there are less than 25 columns
            cFormulas.Add "'" & Replace(oColObj.columnHeader, "'", "''") & "'", sColHdrName
        End If
    Next
End Sub

Private Sub BuildRSFrom2Collections(rsReturn As ADODB.Recordset, cInput As Collection, cOrdinal As Collection)
    'Check if cOrdinal is empty
    
    Dim lItems As Long
    Dim oColObj As Object
    Dim sFieldName As String
    
    On Error Resume Next
    Set rsReturn.ActiveConnection = Nothing
    rsReturn.CursorLocation = adUseClient
    rsReturn.Open
    
    If cInput Is Nothing And rsReturn.RecordCount = 0 Then
        rsReturn.AddNew
        Call InitCurrentRecordSet(rsReturn)
    ElseIf cInput.Count = 0 And rsReturn.RecordCount = 0 Then
        rsReturn.AddNew
        Call InitCurrentRecordSet(rsReturn)
    Else
        For lItems = 1 To cInput.Count
            rsReturn.AddNew
            Call InitCurrentRecordSet(rsReturn)
            For Each oColObj In cOrdinal
                sFieldName = NULLStr
                'oColObj.ColumnHeader
                'oColObj.colName
                sFieldName = sFieldNmBase & oColObj.columnOrdinal
                If UCase(oColObj.colName) = UCase(sRegNum_ColNm) Then
                   rsReturn(sFieldName) = cInput(lItems).RegNum
                ElseIf UCase(oColObj.colName) = UCase(sSpecialty_ColNm) Then
                   rsReturn(sFieldName) = cInput(lItems).Specialty
                ElseIf UCase(oColObj.colName) = UCase(sCapacity_ColNm) Then
                   rsReturn(sFieldName) = cInput(lItems).Capacity
                ElseIf UCase(oColObj.colName) = UCase(sBed_ColNm) Then
                   rsReturn(sFieldName) = cInput(lItems).Bed
                ElseIf UCase(oColObj.colName) = UCase(sRank_ColNm) Then
                   rsReturn(sFieldName) = cInput(lItems).Rank
                ElseIf UCase(oColObj.colName) = UCase(sPatient_ColNm) Then
                   rsReturn(sFieldName) = cInput(lItems).Patient
                ElseIf UCase(oColObj.colName) = UCase(sFMPSSN_ColNm) Then
                   rsReturn(sFieldName) = cInput(lItems).FMPSSN
                ElseIf UCase(oColObj.colName) = UCase(sUnit_ColNm) Then
                   rsReturn(sFieldName) = cInput(lItems).Unit
                ElseIf UCase(oColObj.colName) = UCase(sPatService_ColNm) Then
                   rsReturn(sFieldName) = cInput(lItems).PatService
                ElseIf UCase(oColObj.colName) = UCase(sAdmittingDiagnosis_ColNm) Then
                   rsReturn(sFieldName) = cInput(lItems).AdmittingDiagnosis
                ElseIf UCase(oColObj.colName) = UCase(sAdmissionDate_ColNm) Then
                   rsReturn(sFieldName) = cInput(lItems).AdmissionDate
                ElseIf UCase(oColObj.colName) = UCase(sAdmissionComments_ColNm) Then
                   rsReturn(sFieldName) = cInput(lItems).AdmissionComments
                ElseIf UCase(oColObj.colName) = UCase(sDay_ColNm) Then
                   rsReturn(sFieldName) = cInput(lItems).Day
                ElseIf UCase(oColObj.colName) = UCase(sLocation_ColNm) Then
                   rsReturn(sFieldName) = cInput(lItems).Location
                ElseIf UCase(oColObj.colName) = UCase(sStatus_ColNm) Then
                   rsReturn(sFieldName) = cInput(lItems).Status
                ElseIf UCase(oColObj.colName) = UCase(sDischargeDate_ColNm) Then
                   rsReturn(sFieldName) = cInput(lItems).DischargeDate
                ElseIf UCase(oColObj.colName) = UCase(sDisposition_ColNm) Then
                   rsReturn(sFieldName) = cInput(lItems).Disposition
                ElseIf UCase(oColObj.colName) = UCase(sDischargeDiagnosis_ColNm) Then
                   rsReturn(sFieldName) = cInput(lItems).DischargeDiagnosis
                ElseIf UCase(oColObj.colName) = UCase(sAdmittingProvider_ColNm) Then
                   rsReturn(sFieldName) = cInput(lItems).AdmittingProvider
                ElseIf UCase(oColObj.colName) = UCase(sAttendingProvider_ColNm) Then
                   rsReturn(sFieldName) = cInput(lItems).AttendingProvider
                ElseIf UCase(oColObj.colName) = UCase(sTransferDate_ColNm) Then
                   rsReturn(sFieldName) = cInput(lItems).TransferDate
                ElseIf UCase(oColObj.colName) = UCase(sService_ColNm) Then
                   rsReturn(sFieldName) = cInput(lItems).Service
                End If
            Next
        Next
    End If
End Sub

Private Sub InitCurrentRecordSet(rsReturn As ADODB.Recordset)
    rsReturn("GENERIC_1") = NULLStr
    rsReturn("GENERIC_2") = NULLStr
    rsReturn("GENERIC_3") = NULLStr
    rsReturn("GENERIC_4") = NULLStr
    rsReturn("GENERIC_5") = NULLStr
    rsReturn("GENERIC_6") = NULLStr
    rsReturn("GENERIC_7") = NULLStr
    rsReturn("GENERIC_8") = NULLStr
    rsReturn("GENERIC_9") = NULLStr
    rsReturn("GENERIC_10") = NULLStr
    rsReturn("GENERIC_11") = NULLStr
    rsReturn("GENERIC_12") = NULLStr
    rsReturn("GENERIC_13") = NULLStr
    rsReturn("GENERIC_14") = NULLStr
    rsReturn("GENERIC_15") = NULLStr
    rsReturn("GENERIC_16") = NULLStr
    rsReturn("GENERIC_17") = NULLStr
    rsReturn("GENERIC_18") = NULLStr
    rsReturn("GENERIC_19") = NULLStr
    rsReturn("GENERIC_20") = NULLStr
    rsReturn("GENERIC_21") = NULLStr
    rsReturn("GENERIC_22") = NULLStr
    rsReturn("GENERIC_23") = NULLStr
    rsReturn("GENERIC_24") = NULLStr
    rsReturn("GENERIC_25") = NULLStr
End Sub

Private Sub AppendsResultSetFields(rsReturn As ADODB.Recordset)
    rsReturn.CursorLocation = adUseClient
    With rsReturn.Fields
            .Append "GENERIC_1", adBSTR, 255, 132
            .Append "GENERIC_2", adBSTR, 255, 132
            .Append "GENERIC_3", adBSTR, 255, 132
            .Append "GENERIC_4", adBSTR, 255, 132
            .Append "GENERIC_5", adBSTR, 255, 132
            .Append "GENERIC_6", adBSTR, 255, 132
            .Append "GENERIC_7", adBSTR, 255, 132
            .Append "GENERIC_8", adBSTR, 255, 132
            .Append "GENERIC_9", adBSTR, 255, 132
            .Append "GENERIC_10", adBSTR, 255, 132
            .Append "GENERIC_11", adBSTR, 255, 132
            .Append "GENERIC_12", adBSTR, 255, 132
            .Append "GENERIC_13", adBSTR, 255, 132
            .Append "GENERIC_14", adBSTR, 255, 132
            .Append "GENERIC_15", adBSTR, 255, 132
            .Append "GENERIC_16", adBSTR, 255, 132
            .Append "GENERIC_17", adBSTR, 255, 132
            .Append "GENERIC_18", adBSTR, 255, 132
            .Append "GENERIC_19", adBSTR, 255, 132
            .Append "GENERIC_20", adBSTR, 255, 132
            .Append "GENERIC_21", adBSTR, 255, 132
            .Append "GENERIC_22", adBSTR, 255, 132
            .Append "GENERIC_23", adBSTR, 255, 132
            .Append "GENERIC_24", adBSTR, 255, 132
            .Append "GENERIC_25", adBSTR, 255, 132
    End With
End Sub

Private Sub Class_Terminate()
   Set moCollRecSet = Nothing
   Set moCollFormulas = Nothing
End Sub




