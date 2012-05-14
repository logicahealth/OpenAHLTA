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
Attribute VB_Name = "MapperCDR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements DL_Interfaces.iMapperRads


Private Const NCID_RADIOLOGY_EVENT            As String = "217260"
Private Const NCID_EKG_EVENT                  As String = "162446"
Private Const NCID_PHONE_LOG                  As String = "950"

Private Const NCID_3M_DEFAULT_CONTEXT         As String = "2000"
Private Const NCID_USER_FULL_NAME_CONTEXT     As String = "187"

Private Const cwiENCOUNTER_NUMBER_TIMES       As Integer = 4



'****************************************************
'Temporary Variable for what could be the Ordering Provider.
Private sRequestingHCP As String
'****************************************************

Private Sub Class_Initialize()

  sRequestingHCP = vbNullString

End Sub

Private Sub iMapperRads_Init(Support As DL_Support.iParameters, ByRef Cache As CacheMapper)

End Sub

Private Function iMapperRads_RetrieveRads(NPParams As DL_Support.iParameters) As DTOs.RadClinicalNotes

On Error GoTo ErrHandler

Dim nCount As Integer

Dim objRadParams As Retrieve_Parameters
Dim MyRads As DTOs.RadClinicalNotes

Dim MyRad As DTOs.RadClinicalNote

Dim objDAS As DasCtrl
Dim objDataString As DataString
Dim objClinicalText As ClinicalText

'Dim nCount As Integer
Dim bItemFound As Boolean

  
  Set objRadParams = NPParams
  Set objDAS = GetRadsDasObject(objRadParams)

    '-----------------------------------------------------------------------------------------------------
    '  Loop through all datastrings in "objDas".  For each string that matches the filter criteria,
    '    extract inFormat$ion and store it in the global array, "matClinicalNotes()".
    '-----------------------------------------------------------------------------------------------------

    Set MyRads = New DTOs.RadClinicalNotes


    While objDAS.EOF = False

      Set MyRad = New DTOs.RadClinicalNote

      Set objDataString = objDAS.Events

      MyRad.DataTypeNCID = objDataString.DataTypeNCID

      If MyRad.DataTypeNCID = NCID_EKG_EVENT Then
        GoTo GetNextRecord
      End If

      Set objClinicalText = objDataString.PatientData

      MyRad.EventID = objDataString.DataId
      MyRad.ApplicationNCID = objDataString.Application.ValueConcept.UniqueId
      MyRad.DataTypeNCID = objDataString.DataTypeNCID
      MyRad.DataType = objDataString.DataType
      MyRad.ReportNCID = objClinicalText.ReportType.ValueConcept.UniqueId


      If Not (objRadParams.FilterCollection Is Nothing) Then
        If objRadParams.FilterCollection.Count > 0 Then
          bItemFound = False
          For nCount = 1 To objRadParams.FilterCollection.Count
            If MyRad.ReportNCID = objRadParams.FilterCollection(nCount) Then
              bItemFound = True
              Exit For
            End If
          Next nCount
          If bItemFound = False Then
            GoTo GetNextRecord
          End If
        End If
      End If

      If MyRad.ReportNCID = NCID_PHONE_LOG Then
        MsgBox "Phone Log Found."
        GoTo GetNextRecord
      End If

      MyRad.PlacerNumber = objClinicalText.DocumentId.RefNum

      MyRad.ReportNameOne = objClinicalText.ReportType.Value(NCID_3M_DEFAULT_CONTEXT)
      MyRad.EventDate = objClinicalText.EventDate
      MyRad.CreateDate = objDataString.CreateTime
      MyRad.ModifyDate = objDataString.ModifyTime

      MyRad.StatusNCID = objClinicalText.ReportStatus.ValueConcept.UniqueId
      MyRad.Status = objClinicalText.ReportStatus.Value(NCID_3M_DEFAULT_CONTEXT)

      MyRad.ResultName = vbNullString
      MyRad.ResultCode = 0
      MyRad.ReportNameTwo = vbNullString
      MyRad.ProviderName = vbNullString
      MyRad.ProviderIEN = 0
      MyRad.Amended = False
      MyRad.ExamDate = CDate(0)

      MyRad.HTMLText = objClinicalText.ReportText.TextObject.Text.Text

      MyRad.ProviderNCID = 0
      MyRad.EncounterNumber = "#" & objDataString.EncounterNumber
      MyRad.EncounterFacility = objDataString.EncounterFacility.Value(NCID_3M_DEFAULT_CONTEXT)
      MyRad.EncounterFacilityNCID = objDataString.EncounterFacility.ValueConcept.UniqueId

      MyRad.CreatedBy = objDataString.CreateUser.Value(NCID_USER_FULL_NAME_CONTEXT)

      If objDataString.ModifyUser Is Nothing Then
        MyRad.ModifiedBy = vbNullString
      Else
        MyRad.ModifiedBy = objDataString.ModifyUser.Value(NCID_USER_FULL_NAME_CONTEXT)
      End If

      If Not objClinicalText.ContextInfo.DictatingClinician.IsEmpty Then
        MyRad.Clinician = objClinicalText.ContextInfo.DictatingClinician.Value(NCID_USER_FULL_NAME_CONTEXT)
      Else
        MyRad.Clinician = vbNullString
      End If

      If objClinicalText.PointOfCare.PhysicalLocation.InfoType <> vbNullString Then
        MyRad.POCFacility = objClinicalText.PointOfCare.PhysicalLocation.Facility.Value(NCID_3M_DEFAULT_CONTEXT)
      Else
        MyRad.POCFacility = MyRad.EncounterFacility
      End If

      If MyRad.ProviderName = vbNullString Then
        If MyRad.Clinician <> vbNullString Then
          MyRad.ProviderName = MyRad.Clinician
        Else
          If sRequestingHCP <> vbNullString Then
            MyRad.ProviderName = sRequestingHCP
          End If
        End If
      End If

      If Trim$(MyRad.ReportNameTwo) = vbNullString Then
        MyRad.ReportNameTwo = MyRad.ReportNameOne
      End If

      'This will extract some Data from the Text Report
      CDR_GetHTMLData MyRad

      MyRads.Add MyRad

      Set MyRad = Nothing

GetNextRecord:

      objDAS.GoToNext

    Wend

  Set iMapperRads_RetrieveRads = MyRads

  GoTo CleanUp

  Exit Function

ErrHandler:

  '---show the VB error---
  Err.Raise Err.Number, "DM_Radiology.MapperCDR.RetreieveRads", Err.Description, Err.HelpFile, Err.HelpContext

CleanUp:

  Set MyRads = Nothing
  Set MyRad = Nothing
  Set objDataString = Nothing
  Set objClinicalText = Nothing
  Set objDAS = Nothing
  Set objRadParams = Nothing

End Function

Private Function GetRadsDasObject(NPParams As DL_Support.iParameters) As DasCtrl

  Dim DASObject As DasCtrl
  Set DASObject = New DasCtrl
  Dim objRadParams As Retrieve_Parameters
  Dim sWhereClause As String

  Set objRadParams = NPParams

  DASObject.ClearAll
  DASObject.GeneralCriteria.ClearAll

  DASObject.GeneralCriteria.UnitNumber = objRadParams.Unit_Number

  DASObject.GeneralCriteria.ReturnMultiMedia = False

  If objRadParams.EventID <> "0" And Trim$(objRadParams.EventID) <> vbNullString Then
    sWhereClause = "(EVENT_ID = '" & objRadParams.EventID & "')"
  Else
    sWhereClause = " EXISTS (Select 'x' from Concept_Relation WHERE Relationship_ncid = '363' and Concept_Relation_Ncid = '" & NCID_RADIOLOGY_EVENT & "' and concept_ncid = Event_SubType_Ncid) "

    'Look At the Time Filter
    If Not (objRadParams.TimeObject Is Nothing) Then
      Select Case objRadParams.TimeSearchType
        Case cwiALL_TIMES
        Case cwiLASTN_TIMES
            DASObject.GeneralCriteria.QueryCount = objRadParams.TimeObject.LastN
        Case cwiSLIDING_TIMES
            DASObject.GeneralCriteria.EventEndGMTime = Format$(Now, "General Date")
            DASObject.GeneralCriteria.EventStartGMTime = Format$(objRadParams.TimeObject.CalcSlidingTime(objRadParams.TimeObject.SlidingTime), "General Date")
        Case cwiFIXED_TIMES
            DASObject.GeneralCriteria.EventStartGMTime = Format$(objRadParams.TimeObject.BeginTime, "General Date")
            DASObject.GeneralCriteria.EventEndGMTime = Format$(objRadParams.TimeObject.EndTime + 0.99999, "General Date")
        Case cwiENCOUNTER_NUMBER_TIMES
            DASObject.GeneralCriteria.EncounterNumber = objRadParams.TimeObject.EncounterNumber
            DASObject.GeneralCriteria.Facility.ValueConcept.UniqueId = objRadParams.TimeObject.EncounterFacilityNCID
        Case Else   '--do nothing--
      End Select
    End If

  End If

  DASObject.GeneralCriteria.QueryQualifier = sWhereClause

    DASObject.DoSearch

  Set GetRadsDasObject = DASObject

End Function

Private Function iMapperRads_RetrieveRadTypes(NPParams As DL_Support.iParameters) As DTOs.RadTypes
    
On Error GoTo ErrHandler

  Dim MyRadTypes As DTOs.RadTypes
  Dim MyRadType As DTOs.RadType
  Dim objTypeParams As RetrieveTypes_Parameters
  Dim objPickListFil As PickListFilter
  Dim lListTotal As Long
  Dim lListCount As Long
  
  Dim sDomain As String
  Dim sRelationship As String
  Dim sContext As String
  Dim bPrefOnlyFlag As Boolean

  Set MyRadTypes = New DTOs.RadTypes
  
  Set objTypeParams = NPParams
  sDomain = objTypeParams.Domain
  sRelationship = objTypeParams.Relationship
  sContext = objTypeParams.Context
  bPrefOnlyFlag = objTypeParams.PrefOnlyFlag
    
  '************************************
  Set objPickListFil = New PickListFilter
  With objPickListFil
    .Domain = sDomain
    .Relationship = sRelationship
    .Context = sContext
    .PrefOnlyFlag = bPrefOnlyFlag
    .Load
  End With
  
  lListTotal = objPickListFil.ListItemCnt
  
  If lListTotal > 0 Then
      
    For lListCount = 1 To lListTotal
      Set MyRadType = New RadType
      
      MyRadType.NCID = objPickListFil.ListItem(lListCount).Concept
      MyRadType.Description = objPickListFil.ListItem(lListCount).Representation
      
      MyRadTypes.Add MyRadType
      
      Set MyRadType = Nothing
      
    Next lListCount
    
  
  End If
      
  Set iMapperRads_RetrieveRadTypes = MyRadTypes
  
  GoTo CleanUp
    
Exit Function

ErrHandler:

  '---show the VB error---
  Err.Raise Err.Number, "DM_Radiology.MapperCDR.RetreieveRads", Err.Description, Err.HelpFile, Err.HelpContext
    
CleanUp:

  Set MyRadTypes = Nothing
  Set MyRadType = Nothing
  Set objTypeParams = Nothing
  Set objPickListFil = Nothing

End Function

'This Procedure is Pulling some Data Elements from the Text Report
Private Sub CDR_GetHTMLData(ByRef MyRad As DTOs.RadClinicalNote)

On Error GoTo ErrHandler

Dim sHTMLText As String

Dim sAbnormalText As String
Dim sProcedure As String
Dim sProvider As String
Dim SubstrAbnormal() As String
Dim SubstrProcedure() As String
Dim SubstrProvider() As String
Dim ATXT() As String

Dim iCount As Integer
Dim iCount1 As Integer
Dim iCount2 As Integer

Dim sDay As String
Dim sHour As String
Dim sMinute As String
Dim sSecond As String
Dim sYear As String
Dim sMonth As String
Dim sDate As String

Dim iPos1 As Integer
Dim iPos2 As Integer

sRequestingHCP = vbNullString
sHTMLText = MyRad.HTMLText

  iCount = 0
  iCount = StrToArray(SubstrAbnormal(), sHTMLText, "Result Code:")
  If iCount > 1 Then
      sAbnormalText = SubstrAbnormal(iCount - 1)
        iCount1 = InStr(1, sAbnormalText, Chr(10))
      sAbnormalText = Mid$(sAbnormalText, 1, iCount1 - 1)
      iCount1 = 0
      iCount1 = InStr(2, sAbnormalText, " ")
      MyRad.ResultName = Mid$(sAbnormalText, iCount1 + 1)
      MyRad.ResultCode = Mid$(sAbnormalText, 2, iCount1 - 2)
   End If
   'get procedure
   iCount = 0
   If InStr(1, sHTMLText, "Procedure") Then
   
      iCount = StrToArray(SubstrProcedure(), sHTMLText, "Procedure:")
   
      If iCount > 0 Then
          sProcedure = SubstrProcedure(iCount - 1)
          iCount1 = 0
          iCount1 = InStr(1, sProcedure, Chr(10))
          If iCount1 < 1 Then
              MyRad.ReportNameTwo = sProcedure
          Else
              MyRad.ReportNameTwo = Mid$(sProcedure, 1, iCount1 - 1)
          End If
      End If
   End If
      'get provider
  iCount = 0
  If InStr(1, sHTMLText, "Provider:") Then
    iCount = StrToArray(SubstrProvider(), sHTMLText, "Provider:")
    If iCount > 0 Then
      sProvider = SubstrProvider(iCount - 1)
      iCount1 = 0
      iCount1 = InStr(1, sProvider, Chr(10))
      sProvider = Mid$(sProvider, 1, iCount1 - 1)
      iCount1 = 0
      iCount1 = InStr(1, sProvider, " ")
      MyRad.ProviderName = Mid$(sProvider, iCount1 + 1)
      MyRad.ProviderIEN = Left$(sProvider, iCount1 - 1)
    End If
  End If
  'Requesting HCP
  iCount = 0
  If InStr(1, sHTMLText, "Requesting HCP:") Then
    iCount = StrToArray(SubstrProvider(), sHTMLText, "Requesting HCP:")
    If iCount > 0 Then
      sProvider = SubstrProvider(iCount - 1)
      iCount1 = 0
      iCount1 = InStr(1, sProvider, Chr(10))
      sProvider = Mid$(sProvider, 1, iCount1 - 1)
      sRequestingHCP = Trim$(sProvider)
    End If
  End If
   
  ATXT = Split(sHTMLText, Chr(10))
  
  MyRad.Amended = False
  For iCount2 = 0 To UBound(ATXT)
      Select Case True
                    
          Case Left$(ATXT(iCount2), 15) = "Exam Date/Time:"
              sDate = Mid$(ATXT(iCount2), 16)
              sYear = Mid$(sDate, 1, 4)
              sMonth = Mid$(sDate, 5, 2)
              sMonth = GetMonth(CInt(sMonth))
              sDay = Mid$(sDate, 7, 2)
              sHour = Mid$(sDate, 9, 2)
              sMinute = Mid$(sDate, 11, 2)
              sSecond = Mid$(sDate, 13, 2)
              sDate = sDay & "-" & sMonth & "-" & sYear & " " & sHour & ":" & sMinute & ":" & sSecond
              MyRad.ExamDate = CDate(sDate)
          
          
          Case Left$(ATXT(iCount2), 14) = "Amended Result"
              MyRad.Amended = True
              iPos1 = InStr(21, ATXT(iCount2), " ")
              MyRad.ResultName = vbNullString
              MyRad.ResultName = Mid$(ATXT(iCount2), iPos1 + 1)
              MyRad.ResultCode = 0
              iPos2 = InStr(1, ATXT(iCount2), ":")
              MyRad.ResultCode = Trim$(Mid$(ATXT(iCount2), iPos2 + 2, iPos1 - iPos2 - 1))
          
          Case Left$(ATXT(iCount2), 11) = "Report Text"
              If MyRad.Amended = True Then
                  iPos1 = InStr(1, ATXT(iCount2), ":")
              End If
          '************************************************
       End Select
  Next iCount2
  
  If MyRad.Amended Then
    MyRad.Status = MyRad.Status & " (Amended)"
  End If
  
GoTo CleanUp

Exit Sub

ErrHandler:
   Err.Raise Err.Number, "DM_Radiology.MapperCDR.CDR_GetHTMLData", Err.Description, Err.HelpFile, Err.HelpContext

CleanUp:

  Erase ATXT
  Erase SubstrAbnormal
  Erase SubstrProcedure
  Erase SubstrProvider

End Sub

Private Function StrToArray(ByRef substrs() As String, ByVal SrcStr As String, ByVal Delimiter As String) As Integer
    ' Dimension variables:
     ReDim substrs(0) As String
     Dim CurPos As Long
     Dim NextPos As Long
     Dim DelLen As Integer
     Dim nCount As Integer
     Dim TStr As String
On Error GoTo ErrHandler

     'Since string has a delimiter at the end, add a delimiter to start  of
      'string to make loop simpler:
     SrcStr = Delimiter & SrcStr & Delimiter
     ' Calculate the delimiter length only once:
     DelLen = Len(Delimiter)
     ' Initialize the count and position:
     nCount = 0
     CurPos = 1
     NextPos = InStr(CurPos + DelLen, SrcStr, Delimiter)

     ' Loop searching for delimiters:
     Do Until NextPos = 0
        ' Extract a sub-string:
        TStr = Mid$(SrcStr, CurPos + DelLen, NextPos - CurPos - DelLen)
       ' Add room for the new sub-string in the array:
'       If TStr <> "SYSTEM" And TStr <> "SYSCOMM" Then
        ReDim Preserve substrs(nCount) As String
        ' Put the sub-string in the array:
        substrs(nCount) = TStr
        nCount = nCount + 1
'         End If
        ' Increment the sub string counter:
        '

        ' Position to the last found delimiter:
        CurPos = NextPos
        ' Find the next delimiter:
        NextPos = InStr(CurPos + DelLen, SrcStr, Delimiter)
     Loop

     ' Return the number of sub-strings found:
     StrToArray = nCount
    
    Exit Function
ErrHandler:

  Err.Raise Err.Number, "DM_Radiology.MapperCDR.StrToArray", Err.Description, Err.HelpFile, Err.HelpContext

End Function

