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
Attribute VB_Name = "MapperSQLGeneric"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public Function SQLGeneric(ByVal SQLOps As ICHCSII_DAS, ByVal NPParams As DL_Support.iParameters) As DTOs.RadClinicalNotes

On Error GoTo ErrHandler

Dim objRadParams As Retrieve_Parameters
Dim MyRads As DTOs.RadClinicalNotes

Dim MyRad As DTOs.RadClinicalNote

Dim oRads As adodb.Recordset
Dim oDetails As adodb.Recordset
Dim sSQL As String
Dim sSQLDetails As String

Dim dtRadDate As Date

Dim dtFilterStartDate As Date
Dim dtFilterEndDate As Date
Dim nFilterLastN As Integer

Dim sRadDetailsString As String

Dim nRecordCount As Integer

Set MyRads = New DTOs.RadClinicalNotes

  Set objRadParams = NPParams

    dtFilterStartDate = CDate(0)
    dtFilterEndDate = CDate(0)
    nFilterLastN = 0

    If Not (objRadParams.TimeObject Is Nothing) Then
      Select Case objRadParams.TimeSearchType
        Case cwiALL_TIMES
        Case cwiLASTN_TIMES
            nFilterLastN = objRadParams.TimeObject.LastN
        Case cwiSLIDING_TIMES
            dtFilterEndDate = CDate(Format$(Now, "DD-MMM-YYYY"))
            dtFilterStartDate = CDate(Format$(objRadParams.TimeObject.CalcSlidingTime(objRadParams.TimeObject.SlidingTime), "DD-MMM-YYYY"))
        Case cwiFIXED_TIMES
            dtFilterEndDate = CDate(Format$(objRadParams.TimeObject.EndTime, "DD-MMM-YYYY"))
            dtFilterStartDate = CDate(Format$(objRadParams.TimeObject.BeginTime, "DD-MMM-YYYY"))
        Case Else   '--do nothing--
      End Select
    End If
'**************************************

nRecordCount = 0

sSQL = "SELECT event_id, event_type, unit_number, test_ncid, test_name, " & _
              "date_collected, Placer_Number, Filler_Number, " & _
              "status_ncid, status_name, ordering_facility_ncid, ordering_facility_name, " & _
              "ordering_provider_ncid, ordering_provider_name, date_ordered, date_resulted, " & _
              "priority, specimen, sample, sensitive, comment_text, encounter_number, " & _
              "rad_result_code, rad_result_text " & _
        "FROM Result_Lab_Test " & _
        "WHERE Unit_Number = " & objRadParams.Unit_Number & " and Event_Type = 3 " & _
        "ORDER BY date_collected desc"

Set oRads = SQLOps.OpenRecordset(sSQL)


If Not (oRads.EOF And oRads.BOF) Then
  oRads.MoveFirst
  Do Until oRads.EOF

    '******************************************************
    'Filter Out the ones we don't want.
    Select Case objRadParams.TimeSearchType
        Case cwiALL_TIMES
          'Do Nothing
        Case cwiLASTN_TIMES

          If nRecordCount = nFilterLastN Then
            GoTo GetNextRecord
          End If

        Case cwiSLIDING_TIMES, cwiFIXED_TIMES

          dtRadDate = oRads.Fields("date_collected").Value
          dtRadDate = CDate(Format$(dtRadDate, "DD-MMM-YYYY"))

          If dtRadDate <= dtFilterEndDate And dtRadDate >= dtFilterStartDate Then
          Else
            GoTo GetNextRecord
          End If

        Case Else
          'Do Nothing
    End Select
    '******************************************************
      Set MyRad = New DTOs.RadClinicalNote

      MyRad.ApplicationNCID = 1505
      MyRad.DataTypeNCID = 507
      MyRad.DataType = "1"
      MyRad.EventID = oRads.Fields("event_id").Value

      MyRad.ReportNCID = oRads.Fields("test_ncid").Value
      MyRad.ReportNameOne = oRads.Fields("test_name").Value
      MyRad.ReportNameTwo = oRads.Fields("test_name").Value

      MyRad.EventDate = oRads.Fields("date_collected").Value
      MyRad.CreateDate = oRads.Fields("date_ordered").Value
      MyRad.ModifyDate = CDate(0)
      MyRad.ExamDate = CDate(0)

      MyRad.Status = oRads.Fields("status_name").Value
      MyRad.EncounterFacility = oRads.Fields("ordering_facility_name").Value
      MyRad.POCFacility = oRads.Fields("ordering_facility_name").Value
      MyRad.ProviderName = oRads.Fields("ordering_provider_name").Value
      MyRad.EncounterNumber = oRads.Fields("encounter_number").Value
      MyRad.ResultCode = oRads.Fields("rad_result_code").Value
      MyRad.ResultName = oRads.Fields("rad_result_text").Value
      MyRad.CreatedBy = "USER, INTERFACE"

      sRadDetailsString = vbNullString
      sSQLDetails = "SELECT Comment_Text from Result_MicroBio_Comments " & _
                    "WHERE Event_ID = " & oRads.Fields("event_id").Value & " "
      Set oDetails = SQLOps.OpenRecordset(sSQLDetails)

      MyRad.HTMLText = vbNullString
      If Not (oDetails.EOF And oDetails.BOF) Then
        oDetails.MoveFirst
        Do Until oDetails.EOF
          sRadDetailsString = sRadDetailsString & vbCrLf & oDetails.Fields("comment_text").Value
          oDetails.MoveNext
        Loop
        MyRad.HTMLText = sRadDetailsString
      End If

      MyRads.Add MyRad

      Set MyRad = Nothing



GetNextRecord:

    oRads.MoveNext

  Loop

End If

Set SQLGeneric = MyRads

GoTo CleanUp

Exit Function

ErrHandler:
  '---show the VB error---
  Err.Raise Err.Number, "DM_Radiology.MapperSQLGeneric.SQLGeneric", Err.Description, Err.HelpFile, Err.HelpContext

CleanUp:

  Set oRads = Nothing
  Set SQLOps = Nothing
  Set objRadParams = Nothing
  Set MyRads = Nothing
  Set MyRad = Nothing
  Set oRads = Nothing
  Set oDetails = Nothing

End Function


