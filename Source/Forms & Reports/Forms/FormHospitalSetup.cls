VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FormHospitalSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private moCollRecSet As Collection
Private moCollFormulas As Collection

Const sMODULE_NAME As String = "FormHospitalSetup"

'Use these code for testing
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

Public Function BuildHospitalSetupReport(oReport As CRAXDRT.Report, _
                                        sTitle As String, _
                                        colHospitalSetup As Collection) As Boolean
On Error GoTo ErrHandler
BuildHospitalSetupReport = False

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
   Call BuildRSFromCollection(oRec1, colHospitalSetup)
   
   moCollRecSet.Add oRec1, "HospitalSetup"
   
'   Call DebugPrintRecordSet(oRec1)

'-- Set Formulas
'-----------------------------------------------------------------------------
    Set moCollFormulas = Nothing
    Set moCollFormulas = New Collection
    moCollFormulas.Add "'" & Format(Now, gsDefaultFormat & "  HHMM") & "'", "TodaysDateTime"
    'moCollFormulas.Add "'" & sTitle, "Title"
    
'hwcc 1/10/2004 -
'This whole thing is use for processing a Title larger than 254 characters
' quite a piece of work but do we really needs it?\
'  I do so research on the reality of a Title large than 254 char and go from there
'   Another item in the Laundry list
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
    
'-- Create Report Object
'-----------------------------------------------------------------------------

    Set oCrystal = New Crystal
    Set oReport = Nothing
    If Not oCrystal.ReportOpen("HospitalSetup.rpt", oReport) Then GoTo CleanUp
    If Not oCrystal.AssignRecordSets(moCollRecSet, oReport) Then GoTo CleanUp
    If Not oCrystal.AssignFormulas(moCollFormulas, oReport) Then GoTo CleanUp

BuildHospitalSetupReport = True
GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "Form.BuildHospitalSetupReport", "CHCSII_Forms", vbExclamation
    'Resume Next
CleanUp:
    Set oRec1 = Nothing
    Set oCrystal = Nothing
End Function

Private Sub BuildRSFromCollection(rsReturn As ADODB.Recordset, Optional pInput As Collection = Nothing)
    Dim lItems As Long
    
    On Error Resume Next
    Set rsReturn.ActiveConnection = Nothing
    rsReturn.CursorLocation = adUseClient
    
    rsReturn.Open
    If pInput Is Nothing And rsReturn.RecordCount = 0 Then
        rsReturn.AddNew
        rsReturn("SPECIALTY") = ""
    ElseIf pInput.Count = 0 And rsReturn.RecordCount = 0 Then
        rsReturn.AddNew
        rsReturn("SPECIALTY") = ""
    Else
        For lItems = 1 To pInput.Count
            rsReturn.AddNew
            rsReturn("SPECIALTY") = pInput(lItems).Specialty
            rsReturn("TOTALBEDS") = pInput(lItems).TotalBeds
            rsReturn("BEDSOPEN") = pInput(lItems).BedsOpen
            rsReturn("BEDSOCCUPIED") = pInput(lItems).BedsOccupied
            rsReturn("BEDSAVAILABLE") = pInput(lItems).BedsAvailable
            rsReturn("LOCATION") = pInput(lItems).Location
            rsReturn("STATUSOPEN") = pInput(lItems).StatusOpen
            rsReturn("STATUSACTIVE") = pInput(lItems).StatusActive
        Next
    End If
End Sub

Private Sub AppendsResultSetFields(rsReturn As ADODB.Recordset)
    rsReturn.CursorLocation = adUseClient
    With rsReturn.Fields
            .Append "SPECIALTY", adBSTR, 255, 132
            .Append "TOTALBEDS", adBSTR, 255, 132
            .Append "BEDSOPEN", adBSTR, 255, 132
            .Append "BEDSOCCUPIED", adBSTR, 255, 132
            .Append "BEDSAVAILABLE", adBSTR, 255, 132
            .Append "LOCATION", adBSTR, 255, 132
            .Append "STATUSOPEN", adBSTR, 255, 132
            .Append "STATUSACTIVE", adBSTR, 255, 132
    End With
End Sub

Private Sub Class_Terminate()
   Set moCollRecSet = Nothing
   Set moCollFormulas = Nothing
End Sub

