VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MapperPGUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements DL_Interfaces.iMapperRads

Private nNTECount As Integer
Private nZRRCount As Integer
Private sHL7 As String
Private sMSHArray() As String
Private sORCArray() As String
Private sOBRArray() As String
Private sNTEArray() As String
Private sRadResult() As String
Private sDelimiter As String
Private objShared As CWShared


Private Sub iMapperRads_Init(Support As DL_Support.iParameters, ByRef Cache As CacheMapper)

Set objShared = New CWShared

End Sub

Private Function CHCS_Rads(ByVal RadParams As DL_Support.iParameters, ByRef dtBeginDate As Date, ByRef dtEndDate As Date) As adodb.Recordset

On Error GoTo ErrHandler

Dim dtInterfaceBeginDate As Date
Dim dtInterfaceEndDate As Date
Dim objRadParams As Retrieve_Parameters

Dim LegacyInterface As CHCS_Interface.ICHCSPatient

  Set objRadParams = RadParams
  Select Case objRadParams.TimeSearchType
    Case cwiALL_TIMES
    Case cwiLASTN_TIMES
    Case cwiSLIDING_TIMES
      dtBeginDate = objRadParams.TimeObject.CalcSlidingTime(objRadParams.TimeObject.SlidingTime)
      dtEndDate = CDate(Format$(Now, "DD-MMM-YYYY"))
      dtInterfaceEndDate = dtEndDate
      If objRadParams.TimeObject.SlidingTime <= 4320 Then
        dtInterfaceBeginDate = objRadParams.TimeObject.CalcSlidingTime(4320)
      Else
        dtInterfaceBeginDate = objRadParams.TimeObject.CalcSlidingTime(objRadParams.TimeObject.SlidingTime)
      End If
    Case cwiFIXED_TIMES
      dtBeginDate = objRadParams.TimeObject.BeginTime
      dtEndDate = objRadParams.TimeObject.EndTime
      dtInterfaceBeginDate = dtBeginDate
      dtInterfaceEndDate = dtEndDate
  End Select

Set LegacyInterface = objShared.CHCSConnection
Set CHCS_Rads = LegacyInterface.RetrievePatientRecordsResults(CStr(objRadParams.IEN_Number), eDataType_RADRESULT, dtInterfaceBeginDate, dtInterfaceEndDate, objRadParams.ForceRefresh)
Set LegacyInterface = Nothing
Set objRadParams = Nothing

'Set CHCS_XML_RecordSet = LegacyInterface.RetrievePatientRecords_Results(CStr(mobjPatient.IEN_number), eDataType_RADRESULT, dtBeginDate, dtEndDate, gbForceRefresh)
'Set CHCS_XML_RecordSet = LegacyInterface.RetrievePatientRecords2(CStr(mobjPatient.IEN_number), eDataType_RADRESULT, dtInterfaceBeginDate, dtInterfaceEndDate, gbForceRefresh, eResultsReturnMode_FULL, eResultsReturnFormat$_Recordset)

'''*******************************************************
'''This is COOL
''Set CHCS_XML_RecordSet = New ADODB.Recordset
''CHCS_XML_RecordSet.Open ("C:\rad_results.xml")
'''CHCS_XML_RecordSet.RecordCount
'''*******************************************************

Exit Function

ErrHandler:
  
  Err.Raise Err.Number, "DM_Radiology.MapperPGUI.CHCS_Rads", Err.Description, Err.HelpFile, Err.HelpContext

End Function


Private Function iMapperRads_RetrieveRads(NPParams As DL_Support.iParameters) As DTOs.RadClinicalNotes
  
On Error GoTo ErrHandler

Dim sEventID As String
Dim strValue As String
Dim dtDate As Date
Dim sDate As String


Dim MyRads As DTOs.RadClinicalNotes
Dim objRadParams As Retrieve_Parameters

Dim MyRad As DTOs.RadClinicalNote

Set objRadParams = NPParams

If Trim$(objRadParams.EventID) = "0" Then
  objRadParams.EventID = vbNullString
End If
If objRadParams.SummaryDataPull = True Then
  Set iMapperRads_RetrieveRads = PGUI_SummaryRetrieveData(objRadParams)
  Exit Function
End If

Dim CHCS_XML_RecordSet As adodb.Recordset

Dim dtCollectionDate As Date



Dim dtBeginDate As Date
Dim dtEndDate As Date



Set CHCS_XML_RecordSet = CHCS_Rads(objRadParams, dtBeginDate, dtEndDate)
 
Set MyRads = New DTOs.RadClinicalNotes

If Not (CHCS_XML_RecordSet.EOF And CHCS_XML_RecordSet.BOF) Then
  CHCS_XML_RecordSet.MoveFirst
  Do Until CHCS_XML_RecordSet.EOF
  
    sHL7 = CHCS_XML_RecordSet("PATIENT_DATA")
  
    'This will Make sure it's Radiology Data
    If Process_HL7 Then
    
      'Filter out the EventID if that is set to something
      If objRadParams.EventID <> vbNullString Then
        sEventID = sOBRArray(2)
        If UCase$(Trim$(objRadParams.EventID)) = UCase$(Trim$(sEventID)) Then
          Set MyRad = New DTOs.RadClinicalNote
          PopulateRadClass objRadParams, MyRad
          MyRads.Add MyRad
          Exit Do
        End If
      Else
        'Filter out Unwanted Dates
        strValue = sOBRArray(6)
        GetDate sDate, dtDate, strValue
        dtCollectionDate = CDate(Format$(dtDate, "DD-MMM-YYYY"))
        If Not (dtBeginDate = CDate(0) Or dtEndDate = CDate(0)) Then
          dtBeginDate = CDate(Format$(dtBeginDate, "DD-MMM-YYYY"))
          dtEndDate = CDate(Format$(dtEndDate, "DD-MMM-YYYY"))
          If (dtCollectionDate >= dtBeginDate And dtCollectionDate <= dtEndDate) Then
            Set MyRad = New DTOs.RadClinicalNote
            PopulateRadClass objRadParams, MyRad
            MyRads.Add MyRad
          End If
        End If
      End If
      
      
    End If
    
    Set MyRad = Nothing

    CHCS_XML_RecordSet.MoveNext

  Loop

End If

Set iMapperRads_RetrieveRads = MyRads

GoTo CleanUp

Exit Function

ErrHandler:
  Err.Raise Err.Number, "DM_Radiology.MapperPGUI.RetreieveRads", Err.Description, Err.HelpFile, Err.HelpContext

CleanUp:

  Set MyRad = Nothing
  Set MyRads = Nothing
  Set CHCS_XML_RecordSet = Nothing
  Set objRadParams = Nothing
  Erase sMSHArray
  Erase sORCArray
  Erase sOBRArray
  Erase sNTEArray
  Erase sRadResult

    
End Function

Private Sub PopulateRadClass(ByVal RadParams As DL_Support.iParameters, ByRef MyRad As DTOs.RadClinicalNote)

On Error GoTo ErrHandler

Dim nSlashPos As Integer
Dim sRadTestName As String
Dim sRadTestIEN As String
Dim sExamNumber As String
Dim strValue As String
Dim dtDate As Date
Dim sDate As String
Dim sExamDate As String
Dim sTransDate As String
Dim sDisplayStatus As String
Dim sArray() As String
Dim sFacilityArray() As String
Dim sFacilityName As String
Dim sProviderIEN As String
Dim sProviderName As String
Dim sOrderComment As String
Dim sReasonForOrder As String
Dim sRequestingLocation As String
Dim sOrderingProviderName As String
Dim sBigReportText As String
Dim objRadParams As Retrieve_Parameters

      Set objRadParams = RadParams
      
      MyRad.EventID = Trim$(sOBRArray(2))
      MyRad.ReportNameOne = Trim$(sOBRArray(4))
      nSlashPos = InStr(MyRad.ReportNameOne, "\")
      sRadTestIEN = Left$(MyRad.ReportNameOne, nSlashPos - 1)
      MyRad.ReportNameOne = Right$(MyRad.ReportNameOne, Len(MyRad.ReportNameOne) - nSlashPos)
      MyRad.ReportNameTwo = MyRad.ReportNameOne
      
      MyRad.ApplicationNCID = 1505
      MyRad.DataTypeNCID = 507
      MyRad.DataType = "1"
  
      MyRad.ReportNCID = sRadTestIEN
      
      sRadTestName = "Procedure:" & MyRad.ReportNameOne
  
      'Exam Number
      sExamNumber = sOBRArray(3)
      sExamNumber = "Exam #:" & sExamNumber
  
      strValue = sOBRArray(6)
      GetDate sDate, dtDate, strValue
      MyRad.EventDate = dtDate
  
      'Exam Date
      strValue = sOBRArray(14)
      sDate = vbNullString
      dtDate = CDate(0)
      GetDate sDate, dtDate, strValue
      sExamDate = sDate
      sExamDate = "Exam Date/Time: " & sExamDate
      MyRad.ExamDate = dtDate
  
      'Transcription Date
      strValue = sOBRArray(22)
      sDate = vbNullString
      dtDate = CDate(0)
      GetDate sDate, dtDate, strValue
      sTransDate = sDate
      sTransDate = "Transcription Date/Time: " & sTransDate
  
      'Status
      If nZRRCount > 1 Then
        MyRad.Status = sOBRArray(25) & " (Amended)"
      Else
        MyRad.Status = sOBRArray(25)
      End If
      sDisplayStatus = "Status:" & MyRad.Status
  
      'Facility / Requesting Location
      sFacilityArray = Split(sORCArray(13), "\")
      sFacilityName = sFacilityArray(UBound(sFacilityArray))
      MyRad.EncounterFacility = sFacilityName
      MyRad.POCFacility = sFacilityName
      sRequestingLocation = sFacilityArray(1) & " @ " & sFacilityName
      sRequestingLocation = "Requesting Location:" & sRequestingLocation
  
  
      '**********************************
      'Provider IEN and Name
      '16 Provider IEN
      GetHL7Provider objRadParams, sOBRArray(16), sProviderIEN, sProviderName
      MyRad.ProviderName = sProviderName
      MyRad.ProviderIEN = sProviderIEN
      sOrderingProviderName = "Provider:" & sProviderName
      '**********************************
  
      'Encounter Number
      MyRad.EncounterNumber = "unknown"
  
      'Order Comment and Reason for Order (NTE Segment)
      sOrderComment = vbNullString
      sReasonForOrder = vbNullString
      For nNTECount = 1 To UBound(sNTEArray)
        sArray = Split(sNTEArray(nNTECount), sDelimiter)
        If InStr(1, UCase$(sArray(3)), "ORDER COMMENT:") Then
          sOrderComment = sArray(3)
        End If
        If InStr(1, UCase$(sArray(3)), "REASON FOR ORDER:") Then
          sReasonForOrder = sReasonForOrder & sArray(3)
        Else
          If sReasonForOrder <> vbNullString Then
            sReasonForOrder = sReasonForOrder & sArray(3)
          End If
        End If
      Next nNTECount
  
  
      MyRad.ModifyDate = CDate(0)
      MyRad.CreatedBy = "Simpson, Homer"
  
      sBigReportText = sRadTestName & vbCrLf & _
                    sOrderComment & vbCrLf & _
                    sReasonForOrder & vbCrLf & _
                    sExamNumber & vbCrLf & _
                    sExamDate & vbCrLf & _
                    sTransDate & vbCrLf & _
                    sOrderingProviderName & vbCrLf & _
                    sRequestingLocation & vbCrLf & _
                    sDisplayStatus
                    
      ProcessZRRSegment objRadParams, sBigReportText, MyRad
  
      MyRad.HTMLText = sBigReportText

Erase sArray
Erase sFacilityArray
Set objRadParams = Nothing

Exit Sub

ErrHandler:
  Err.Raise Err.Number, "DM_Radiology.MapperPGUI.PopulateRadClass", Err.Description, Err.HelpFile, Err.HelpContext

End Sub

Private Function Process_HL7() As Boolean

On Error GoTo ErrHandler

Dim sHL7Array() As String
Dim nArrayCount As Integer
Dim nResultCount As Integer
Dim nMSHLocation As Integer
    
    sHL7Array = Split(sHL7, vbCr)
    nZRRCount = 0

    For nArrayCount = 0 To UBound(sHL7Array)

      If nArrayCount = 0 Then
        sDelimiter = vbNullString
        ReDim sMSHArray(0)
        ReDim sORCArray(0)
        ReDim sOBRArray(0)
        ReDim sNTEArray(0)
        nNTECount = 0
        nResultCount = 0
      End If

      'Find out what the Delimiter is.
      If nArrayCount = 0 And InStr(sHL7Array(0), "MSH") Then
        nMSHLocation = InStr(sHL7Array(0), "MSH")
        sDelimiter = Mid$(sHL7Array(0), nMSHLocation + 3, 1)
        sMSHArray = Split(sHL7Array(0), sDelimiter)
        If sMSHArray(2) = "CHCS\RAD" Then
          Process_HL7 = True
        Else
          Process_HL7 = False
        End If
      Else

        If nArrayCount > 0 And sDelimiter <> vbNullString Then
          Select Case Left$(sHL7Array(nArrayCount), 4)
            Case "PID^"
            Case "ORC^"
              sORCArray = Split(sHL7Array(nArrayCount), sDelimiter)
            Case "OBR^"
              sOBRArray = Split(sHL7Array(nArrayCount), sDelimiter)

            'We need an Array of NTE and OBX segments
            Case "NTE^"
              nNTECount = nNTECount + 1
              If nNTECount = 1 Then
                ReDim sNTEArray(1 To 1)
              Else
                ReDim Preserve sNTEArray(1 To nNTECount)
              End If
              sNTEArray(nNTECount) = sHL7Array(nArrayCount)

            'Results and Amendments
            Case "OBX^", "ZRR^"
              nResultCount = nResultCount + 1
              If nResultCount = 1 Then
                ReDim sRadResult(1 To 1)
              Else
                ReDim Preserve sRadResult(1 To nResultCount)
              End If
              sRadResult(nResultCount) = sHL7Array(nArrayCount)

              If Left$(sHL7Array(nArrayCount), 4) = "ZRR^" Then
                nZRRCount = nZRRCount + 1
              End If


          End Select
        End If

      End If

    Next nArrayCount
    
Erase sHL7Array

Exit Function

ErrHandler:
  Err.Raise Err.Number, "DM_Radiology.MapperPGUI.Process_HL7", Err.Description, Err.HelpFile, Err.HelpContext

End Function


Private Sub ProcessZRRSegment(ByVal RadParams As DL_Support.iParameters, ByRef sBigReportText As String, ByRef MyRad As DTOs.RadClinicalNote)

  On Error GoTo ErrHandler

  Dim nZRRCount As Integer
  Dim sReportText As String
  Dim nResultCount As Integer
  Dim sResultSegment() As String
  Dim sResultCodeArray() As String
  Dim sResultCode As String
  Dim strValue As String
  Dim sDate As String
  Dim dtDate As Date
  Dim sApprovedDate As String
  Dim sProviderIEN As String
  Dim sProviderName As String
  Dim sApprovedBy As String
  Dim sInterpretedBy As String
  Dim objRadParams As Retrieve_Parameters
  
      Set objRadParams = RadParams
      
      nZRRCount = 0
      sReportText = "Report Text:"
      For nResultCount = 1 To UBound(sRadResult)
        sResultSegment = Split(sRadResult(nResultCount), sDelimiter)
        Select Case sResultSegment(0)
          Case "OBX"

            If InStr(1, UCase$(sResultSegment(3)), "RESCODE") Then
              sResultCodeArray = Split(sResultSegment(5), "\")
              MyRad.ResultCode = sResultCodeArray(0)
              MyRad.ResultName = sResultCodeArray(1)
              If nZRRCount > 0 Then
                sResultCode = "Amended Result Code:" & MyRad.ResultName
              Else
                sResultCode = "Result Code:" & MyRad.ResultName
              End If
            End If
            If InStr(1, UCase$(sResultSegment(3)), "REPORT") Then
              If sReportText = "Report Text:" Then
                sReportText = sReportText & sResultSegment(5)
              Else
                sReportText = sReportText & vbCrLf & sResultSegment(5)
              End If
            End If

          Case "ZRR"
            nZRRCount = nZRRCount + 1
            strValue = sResultSegment(2)
            sDate = vbNullString
            dtDate = CDate(0)
            GetDate sDate, dtDate, strValue
            sApprovedDate = sDate
            sApprovedDate = "Approved Date/Time: " & sApprovedDate

            GetHL7Provider objRadParams, sResultSegment(5), sProviderIEN, sProviderName
            sApprovedBy = "Approved By:" & sProviderName
            GetHL7Provider objRadParams, sResultSegment(9), sProviderIEN, sProviderName
            sInterpretedBy = "Interpreted By:" & sProviderName

            sBigReportText = sBigReportText & _
                    vbCrLf & sResultCode & vbCrLf & _
                    sInterpretedBy & vbCrLf & _
                    sApprovedBy & vbCrLf & _
                    sApprovedDate & vbCrLf & _
                    sReportText

            sReportText = "Report Text:"

        End Select

      Next nResultCount


Set objRadParams = Nothing
Erase sResultSegment
Erase sResultCodeArray

Exit Sub

ErrHandler:
  
  Err.Raise Err.Number, "DM_Radiology.MapperPGUI.ProcessZRRSegment", Err.Description, Err.HelpFile, Err.HelpContext
  
End Sub


Private Function iMapperRads_RetrieveRadTypes(NPParams As DL_Support.iParameters) As DTOs.RadTypes
  
  Err.Raise 0, "DM_Radiology.MapperPGUI", "ImapperRads_RetrieveRadTypes is not valid in PGUI mode.", Err.HelpFile, Err.HelpContext
    
End Function


Private Function PGUI_SummaryRetrieveData(ByVal RadParams As DL_Support.iParameters) As DTOs.RadClinicalNotes

  On Error GoTo ErrHandler
  
  Dim MyRads As DTOs.RadClinicalNotes

  Dim MyRad As DTOs.RadClinicalNote

  Dim LegacyInterface As CHCS_Interface.ICHCSPatient
  Dim oRadsXML As MSXML2.IXMLDOMElement
  Dim oRadXML As MSXML2.IXMLDOMElement
  
  Dim dtBeginDate As Date
  Dim dtEndDate As Date
    
  Dim sArray() As String
  
  '''''''''''''''''''''''''''''''''''''''''''''''''
  Dim sTemp As String
  Dim nSlashPos As Integer
  Dim sName As String
  Dim sIEN As String
  '''''''''''''''''''''''''''''''''''''''''''''''''
      
  Dim objRadParams As Retrieve_Parameters
  
  Set objRadParams = RadParams
  
  Select Case objRadParams.TimeSearchType
    Case cwiSLIDING_TIMES
      dtBeginDate = objRadParams.TimeObject.CalcSlidingTime(objRadParams.TimeObject.SlidingTime)
      dtEndDate = CDate(Format$(Now, "DD-MMM-YYYY"))
    Case cwiFIXED_TIMES
      dtBeginDate = objRadParams.TimeObject.BeginTime
      dtEndDate = objRadParams.TimeObject.EndTime
  End Select

  Set LegacyInterface = objShared.CHCSConnection
  
  Set oRadsXML = LegacyInterface.RetrievePatientRecords2XML(CStr(objRadParams.IEN_Number), eDataType.eDataType_RADRESULT, dtBeginDate, dtEndDate, False, eResultsReturnMode_SUMMARY)

  Set MyRads = New DTOs.RadClinicalNotes
  
  If Not oRadsXML.childNodes Is Nothing Then
    For Each oRadXML In oRadsXML.childNodes
    
        Set MyRad = New DTOs.RadClinicalNote
        
        MyRad.EventID = oRadXML.getAttribute("EVENT_ID") & vbNullString
        sTemp = oRadXML.getAttribute("PROCEDURE") & vbNullString
        
        If Trim$(sTemp) = vbNullString Then
          GoTo GetNextRecord
        End If
        
        nSlashPos = InStr(sTemp, "\")
        sName = Right$(sTemp, Len(sTemp) - nSlashPos)
        sIEN = Left$(sTemp, nSlashPos - 1)
        
        MyRad.ReportNCID = sIEN
        MyRad.ReportNameOne = sName
        MyRad.ReportNameTwo = sName
        MyRad.ApplicationNCID = 1505
        MyRad.DataTypeNCID = 507
        MyRad.DataType = "1"
        
        MyRad.EventDate = ConvertStringToDate(oRadXML.getAttribute("DATE") & vbNullString)
        MyRad.ExamDate = ConvertStringToDate(oRadXML.getAttribute("DATE") & vbNullString)
        
        MyRad.Status = oRadXML.getAttribute("STATUS") & vbNullString
        
        'Facility / Requesting Location
        sTemp = oRadXML.getAttribute("MTF") & vbNullString
        sArray = Split(sTemp, "\")
        sName = sArray(UBound(sArray))
        
        MyRad.EncounterFacility = sName
        MyRad.POCFacility = sName
        MyRad.EncounterFacilityNCID = 0
         
        sName = vbNullString
        sIEN = vbNullString
        GetHL7Provider objRadParams, oRadXML.getAttribute("ORDERING_PROVIDER") & vbNullString, sIEN, sName
        
        MyRad.ProviderName = sName
        MyRad.ProviderIEN = sIEN
        MyRad.ProviderNCID = sIEN
        
        MyRad.ResultCode = oRadXML.getAttribute("RESULT_CODE") & vbNullString
        
        MyRad.ResultName = oRadXML.getAttribute("RESULT_TEXT") & vbNullString
        
        MyRads.Add MyRad
        
        Set MyRad = Nothing

        'DATE="MAY 27, 2003@08:54"
        'STATUS="COMPLETE Amended"
        'EVENT_ID="030527-00002"
        'MTF="46\0"
        'ORDERING_PROVIDER="9574\COOKE\RICHARD T"
        'RESULT_CODE="4"
        'RESULT_TEXT="ABNORMALITY, ATTN. NEEDED"
        'PROCEDURE="3\CHEST,AP"
        'PROCEDURE_ROWCOUNT="1"
        'ID="03001596"/>
        
GetNextRecord:
    Next oRadXML
    
  End If
  
  Set PGUI_SummaryRetrieveData = MyRads
  
  GoTo CleanUp
  
Exit Function

ErrHandler:

  '---show the VB error---
  Err.Raise Err.Number, "DM_Radiology.MapperPGUI.PGUI_SummaryRetrieveData", Err.Description, Err.HelpFile, Err.HelpContext

CleanUp:

  Set objRadParams = Nothing
  Set LegacyInterface = Nothing
  Set oRadsXML = Nothing
  Set MyRads = Nothing
  Set MyRad = Nothing
  Erase sArray

End Function

Private Sub GetDate(ByRef sDate As String, ByRef dtDate As Date, ByVal strValue As String)

On Error GoTo ErrHandler

Dim sYear As String
Dim sMonth As String
Dim sDay As String
Dim sTime As String
Dim sHour As String
Dim sMinute As String

Dim nPos As Integer

  'strValue should be a 14 digit string

  strValue = Trim$(strValue)
  nPos = InStr(strValue, ".")
  If nPos > 0 Then
    strValue = Left$(strValue, nPos - 1)
  End If
  
  strValue = Trim$(strValue)
  If Len(strValue) > 14 Then
    strValue = Left$(strValue, 14)
  End If
  
  If IsNumeric(strValue) Then
  Else
    sDate = "Not Recorded"
    GoTo CleanUp
  End If
  
  sDate = vbNullString
  dtDate = CDate(0)
   
  If strValue = vbNullString Or strValue = "0" Then
     sDate = "Not Recorded"
     dtDate = CDate(0)
  Else
     sYear = Mid$(strValue, 1, 4)
     sMonth = GetMonth(CInt(Mid$(strValue, 5, 2)))
     sDay = Mid$(strValue, 7, 2)
     sTime = Mid$(strValue, 9, 4)
     sHour = Left$(sTime, 2)
     sMinute = Right$(sTime, 2)
     
     If sTime = "    " Then
        sTime = "0000"
     End If
     If sHour = "  " Then
        sHour = "00"
     End If
     If sMinute = "  " Then
        sMinute = "00"
     End If
     
     If Trim$(sTime) = "0000" Then
        sDate = sDay & "-" & sMonth & "-" & sYear
     Else
        sDate = sDay & "-" & sMonth & "-" & sYear & " " & sHour & ":" & sMinute
     End If
     'Date used for Sorting
     dtDate = CDate(sDay & "-" & sMonth & "-" & sYear & " " & sHour & ":" & sMinute)
  End If
  
  GoTo CleanUp
  
  Exit Sub
  
ErrHandler:

  sDate = "Not Recorded"
  
CleanUp:
  
        
  Exit Sub
 
End Sub

Private Sub GetHL7Provider(ByVal RadParams As DL_Support.iParameters, ByVal sSegment As String, ByRef sProviderIEN As String, ByRef sProviderName As String)

On Error GoTo ErrHandler
  
  Dim sProviderArray() As String
  Dim objRadParams As Retrieve_Parameters
  
  Set objRadParams = RadParams
  
  sProviderIEN = vbNullString
  sProviderName = vbNullString

      If InStr(sSegment, "\") Then
        sProviderArray = Split(sSegment, "\")
        Select Case UBound(sProviderArray)
        Case 1
          sProviderName = sProviderArray(1)
        Case 2
          sProviderName = sProviderArray(2) & " " & sProviderArray(1)
        Case 3
          sProviderName = sProviderArray(2) & " " & sProviderArray(3) & " " & sProviderArray(1)
        Case Else
          sProviderName = vbNullString
        End Select
      
        sProviderIEN = sProviderArray(0)
      Else
        If sSegment <> vbNullString Then
          If IsNumeric(sSegment) Then
            sProviderIEN = sSegment
              sProviderName = vbNullString
              PGUI_GetProviderNameFromSQL objRadParams, sProviderIEN, sProviderName
          
          End If
        End If
        
      End If
      

  Set objRadParams = Nothing
  Erase sProviderArray
      
Exit Sub

ErrHandler:

  '---show the VB error---
  Err.Raise Err.Number, "DM_Radiology.MapperPGUI.GetHL7Provider", Err.Description, Err.HelpFile, Err.HelpContext

End Sub

'*****************************
'SCR 50824
'This procedure will be used to retrieve the Provider name from the IEN.
'This will be used if the HL7 Message has the IEN without the name.
Private Sub PGUI_GetProviderNameFromSQL(ByVal RadParams As DL_Support.iParameters, ByVal sProviderIEN As String, ByRef sProviderName As String)

On Error GoTo ErrHandler

  Dim sSQL As String
  Dim LegacySQL As CHCS_Interface.ICHCSSQL
  Dim CHCS_XML_RecordSet As adodb.Recordset
  Dim sNameArray() As String
  Dim sLastName As String
  Dim sFirstName As String
  Dim objRadParams As DM_Radiology.Retrieve_Parameters
  
  Set objRadParams = RadParams
  
  sProviderName = vbNullString
  
  sSQL = "SELECT Name, IEN " & _
          "FROM Provider " & _
          "WHERE IEN = '" & sProviderIEN & "' " & _
          "AND Facility_NCID = '" & objRadParams.LogonFacilityNCID & "'"

  Set LegacySQL = objShared.CHCSConnection
  Set CHCS_XML_RecordSet = LegacySQL.ExecuteSQL(sSQL)
  
  If CHCS_XML_RecordSet.EOF And CHCS_XML_RecordSet.BOF Then
  
  Else
    
    CHCS_XML_RecordSet.MoveFirst
    sProviderName = CHCS_XML_RecordSet("NAME")
    
    If InStr(sProviderName, ",") Then
      
      sNameArray = Split(sProviderName, ",")
      sLastName = sNameArray(0)
      sFirstName = sNameArray(1)
      
      sProviderName = Trim$(sFirstName) & " " & Trim$(sLastName)
      
    End If
  
  End If
  
  GoTo CleanUp
  
  Exit Sub

ErrHandler:

  '---show the VB error---
  Err.Raise Err.Number, "DM_Radiology.MapperPGUI.PGUI_GetProviderNameFromSQL", Err.Description, Err.HelpFile, Err.HelpContext

CleanUp:
  
  Set LegacySQL = Nothing
  Set CHCS_XML_RecordSet = Nothing
  Set objRadParams = Nothing
  Erase sNameArray
  
End Sub

Private Function ConvertStringToDate(ByVal sDate As String) As Date

On Error GoTo ErrHandler
  
  Dim nLoc As Integer
  Dim strDate As String
  Dim strTime As String
  
  nLoc = InStr(1, sDate, "@")
  
  If nLoc > 0 Then
    strDate = Left$(sDate, nLoc - 1)
    strTime = Right$(sDate, Len(sDate) - nLoc)
    If IsDate(strDate & " " & strTime) Then
      ConvertStringToDate = CDate(strDate & " " & strTime)
    End If
  Else
    If IsDate(sDate) Then
      ConvertStringToDate = CDate(sDate)
    Else
      ConvertStringToDate = CDate(0)
    End If
    
  End If
  
Exit Function

ErrHandler:

  '---show the VB error---
  Err.Raise Err.Number, "DM_Radiology.MapperPGUI.ConvertStringToDate", Err.Description, Err.HelpFile, Err.HelpContext

End Function
