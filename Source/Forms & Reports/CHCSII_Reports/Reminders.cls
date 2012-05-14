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
Attribute VB_Name = "Reminders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit



Public Enum TypeOfReminder
    ReminderHistory = 1
    RemindersDue = 2
End Enum


Private Type gtHeaderCommentInfo
    sSourceOfCommentNCID As String
    sCodedComment() As String
    sTextComment() As String
End Type

'---object for translating NCIDs---
Private gobjConTrans As New ConTrans
'---context object---
Private gobjContext As New ContextCtrl
'****************************************************************************************
'  Function Name: ActiveDuty
'
'  Description: This function returns True if the patient is active duty or false if not
'
'  Paramters:
'    None
'
'  Return Value:
'    True - Active Duty     False - Not Active Duty
'****************************************************************************************
Private Function ActiveDuty() As Boolean
    
    
    Dim gobjMember As Member
    Dim gobjMMI As MmiDoc
    
    '---assign the MMI number of the patient to retrieve---
    gobjMMI.Member.UnitNumber = cwobjPatient.UnitNumber  '---or any other way you plan to obtain the current patient

    '----------------------------------------------------------------------------------------
    '  "Get" this record as opposed to "searching" for it.  "Getting" is MUCH FASTER than
    '    searching for a single record.
    '----------------------------------------------------------------------------------------
    gobjMMI.Get (False)      '--if "false", then "read-only", else writeable --
    
    '---point to the member object---
    Set gobjMember = gobjMMI.Member

    '--- Generic MMI Objects
    ActiveDuty = False
    If Not gobjMember.GenericData Is Nothing Then
        gobjMember.GenericData.GotoFirst
        If Err Then
            Err.Clear
        Else
            Do While Not gobjMember.GenericData.EOF
                If gobjMember.GenericData.Type = NCID_MILITARY_STATUS Then
                    If gobjMember.GenericData.Value = NCID_ACTIVE_STATUS Then
                        ActiveDuty = True
                        Exit Do
                    Else
                        ActiveDuty = False
                    End If
                End If
                gobjMember.GenericId.GoToNext
            Loop
        End If
    End If
    
    '--- Clean up
    Set gobjMMI = Nothing
    Set gobjMember = Nothing
End Function

'****************************************************************************************************************************
'  Function Name: FindObs
'
'  Description: This function loops through a collection of clinical observations and finds
'               a specific "type" of observation.
'
'  Parameters:
'    vsSearchTypeNCID - Contains the NCID of the "type" of clinical observation to search for.
'    robjSearchObs - Contains a collection of clinical observations to search.
'    robjReturnObs - Upon return will be set to the clinical observation that matches the
'                    passed in "search type NCID", otherwise it will be set to "Nothing".
'    rvarObsIndex - OPTIONAL: Upon return, if the observation was found, will contain the "index"
'                             of the observation.
'
'  Return Value:
'    Returns TRUE if the specific clinical observation was found.
'    Returns FALSE if an error occurred or the specific clinical observation was not found.
'****************************************************************************************************************************
Private Function FindObs(ByVal vsSearchTypeNCID As String, _
                        robjSearchObs As ClinicalObservations, _
                        robjReturnObs As ClinicalObservation, _
                        Optional rvarObsIndex As Variant) As Boolean
    Dim i As Integer
    
    On Error GoTo ErrHandler

    '---initialize variables---
    FindObs = False
    Set robjReturnObs = Nothing

    '--------------------------------------------------------------------------------------
    '  Loop through the passed in clinical observations and find the clinical observation
    '    that matches the passed in "search type" NCID.
    '--------------------------------------------------------------------------------------
    For i = 1 To robjSearchObs.Count
        If robjSearchObs.Item(i).ObsId.Value.CodedWOSform.ValueConcept.UniqueId = vsSearchTypeNCID Then
            '---found a match---
            Set robjReturnObs = robjSearchObs.Item(i)
            
            '--------------------------------------------------------------------------
            '  If the user wants the "index" of this observation, set the parameter.
            '--------------------------------------------------------------------------
            If IsMissing(rvarObsIndex) = False Then
                rvarObsIndex = i
            End If
            
            '---set return value---
            FindObs = True
            
            Exit For
        End If
    Next i

    Exit Function
    
ErrHandler:
    '---show the VB error---
    gobjshared.ShowVBError Err.Number, Err.Description, "FindObs", "Wellness", vbExclamation

    Exit Function
End Function

'****************************************************************************************
'  Sub Name: GetHeaderComments
'
'  Description: This subroutine extracts the comments from the passed in clinical
'               observation header.
'
'  Paramters:
'    robjClinObsHeader - Contains the clinical observations header for which to extract
'                        the comments.
'    ratCommentInfo() - Upon return will contain ALL of the comments contained in the
'                       passed in clinical observation header.
'****************************************************************************************
Private Sub GetHeaderComments(robjClinObsHeader As ClinObsHeader, _
                             ratCommentInfo() As gtHeaderCommentInfo)
    Dim i As Integer
    Dim j As Integer
    Dim iNumCoded As Integer
    Dim iNumText As Integer
    Dim objCommentGroup As CommentGroup
    Dim objCommentItem As CommentItem
    
    On Error GoTo ErrHandler
    
    '---initialize variables---
    ReDim ratCommentInfo(0)
    '------------------------------------------------------------------------
    '  Extract the comments contained in the clinical observations "header".
    '------------------------------------------------------------------------
    '---see if any comments exist---
    If Not robjClinObsHeader.Comments.IsEmpty Then
        '---loop through all comment groups---
        For i = 1 To robjClinObsHeader.Comments.Count
            '---point to the comment group---
            Set objCommentGroup = robjClinObsHeader.Comments.Item(i)
            
            '---make room to store the comments for this "group"---
            If i > 1 Then
                ReDim Preserve ratCommentInfo(1 To i)
            Else
                ReDim ratCommentInfo(1 To i)
            End If
            
            '---get the "source of comment"---
            If Not objCommentGroup.SourceOfComment.IsEmpty Then
                ratCommentInfo(i).sSourceOfCommentNCID = objCommentGroup.SourceOfComment.Value.CodedWOSform.ValueConcept.UniqueId
            Else
                ratCommentInfo(i).sSourceOfCommentNCID = ""
            End If
            
            '---initialize variables---
            ReDim ratCommentInfo(i).sCodedComment(0)
            ReDim ratCommentInfo(i).sTextComment(0)
            iNumCoded = 0
            iNumText = 0
            
            '---loop through all comment items in this comment group---
            For j = 1 To objCommentGroup.Comments.Count
                '---point to the comment item---
                Set objCommentItem = objCommentGroup.Comments.Item(j)
                
                '---comments can be TEXT AND CODED---
                If objCommentItem.InfoType = "text" Then
                    iNumText = iNumText + 1
                    
                    '---make room to store the comment---
                    If iNumText > 1 Then
                        ReDim Preserve ratCommentInfo(i).sTextComment(1 To iNumText)
                    Else
                        ReDim ratCommentInfo(i).sTextComment(1 To iNumText)
                    End If
                
                    '---store the comment---
                    ratCommentInfo(i).sTextComment(iNumText) = objCommentItem.Text.Text
                ElseIf objCommentItem.InfoType = "code" Then
                    iNumCoded = iNumCoded + 1
                    
                    '---make room to store the comment---
                    If iNumCoded > 1 Then
                        ReDim Preserve ratCommentInfo(i).sCodedComment(1 To iNumCoded)
                    Else
                        ReDim ratCommentInfo(i).sCodedComment(1 To iNumCoded)
                    End If
                
                    '---store the comment---
                    ratCommentInfo(i).sCodedComment(iNumCoded) = objCommentItem.Code.Value.CodedWOSform.ValueConcept.UniqueId
                End If
                
                Set objCommentItem = Nothing
            Next j
            
            Set objCommentGroup = Nothing
        Next i
    End If
    
    Exit Sub

ErrHandler:
    '---show the VB error---
    gobjshared.ShowVBError Err.Number, Err.Description, "GetHeaderComments", "Wellness", vbExclamation
    
    Exit Sub
End Sub

'*******************************************************************************************
'  Function Name: BuildReminders
'
'  Description: This function builds "reminder data" according to the passed in parameters.
'
'  Paramters:
'    ratReminderInfo() - Upon return will contain the reminders data that was built.
'    viTabIndex - Contains the "index" into the tabs on the main form.
'    viTimeSearchType - Contains the "type" of time search being performed.
'    robjTime - Contains the time object being used to store the time period values.
'    rlvwControl - Contains the listview control that will be "cleared".
'
'  Return Value:
'    Returns TRUE if the data was built successfully.
'    Returns FALSE if an error occurred in building the data.
'*******************************************************************************************
Public Function DataGetReminders(ReminderType As TypeOfReminder, sPatientUnitNumber As String, collReminder As Collection) As Boolean
'ByVal viTabIndex As Integer, _
'                                ByVal viTimeSearchType As Integer, _
'                                robjTime As CWTime) As Boolean
    Dim sWhereClause As String
    Dim sOldContext As String
    Dim oDatastring As DataString
    Dim oClinicalEvent As ClinicalEvent
    Dim oClinObsHeader As ClinObsHeader
    Dim oClinicalObservation As ClinicalObservation
    Dim oClinicalObservationMod As ClinicalObservation
    Dim atHeaderComments() As gtHeaderCommentInfo
    Dim i As Integer
    Dim bSkipThis As Boolean
    
    Dim oDas As New DasCtrl
    Dim collReminders As Collection
    Dim oWellnessReminder As WellnessReminder

    
    
    #If debugon Then
        DL.DebugLog Exe, Module, "BuildReminders", "P-Begin BuildReminders", DR
    #End If

    On Error GoTo ErrHandler
    
    '---set the return value---
    DataGetReminders = False
    
    '---initialize variables---
    'ReDim ratReminderInfo(0)
   Set collReminder = New Collection
    
    
'-- Builder Where Clause
'-----------------------------------------------------------------------------

    '---clear out previous "data event" strings---
    oDas.ClearAll
    
    '---clear previous search criteria---
    oDas.GeneralCriteria.ClearAll

    '---set search criteria---
Rem    Select Case viTimeSearchType
Rem        Case cwiLASTN_TIMES
Rem            gobjDas.GeneralCriteria.QueryCount = robjTime.LastN
Rem        Case cwiSLIDING_TIMES
Rem            gobjDas.GeneralCriteria.EventEndGMTime = Format$(Now, "General Date")
Rem            gobjDas.GeneralCriteria.EventStartGMTime = Format$(robjTime.CalcSlidingTime(robjTime.SlidingTime), "General Date")
Rem        Case cwiFIXED_TIMES
Rem            gobjDas.GeneralCriteria.EventStartGMTime = Format$(robjTime.BeginTime, "General Date")
Rem            gobjDas.GeneralCriteria.EventEndGMTime = Format$(robjTime.EndTime, "General Date")
Rem        Case cwiENCOUNTER_NUMBER_TIMES
Rem            gobjDas.GeneralCriteria.EncounterNumber = robjTime.EncounterNumber
Rem            gobjDas.GeneralCriteria.Facility.ValueConcept.UniqueId = robjTime.EncounterFacilityNCID
Rem        Case Else   '--do nothing--
Rem    End Select
    
    oDas.GeneralCriteria.UnitNumber = sPatientUnitNumber ' cwobjPatient.UnitNumber
    oDas.GeneralCriteria.QueryName.ValueConcept.UniqueId = NCID_CLIN_OBS_EVENT_VIEW
    
    '---Set Search Criteria Based on the Type of Reminder (History, Reminders Due)
    If ReminderType = ReminderHistory Then
        oDas.GeneralCriteria.EventType.ValueConcept.UniqueId = NCID_REMINDER_EVENT
        sWhereClause = "STATUS_NCID <> " & NCID_ACTIVE
    ElseIf ReminderType = RemindersDue Then
        sWhereClause = "(EVENT_TYPE_NCID = " & NCID_REMINDER_EVENT & " AND STATUS_NCID = " & NCID_ACTIVE
    End If
    
    '---only retrieve reminders for the "production statuses" the user has access to---
Rem    For i = 1 To UBound(gasProductionStatuses)
Rem        If i = 1 Then
Rem            sWhereClause = sWhereClause & " AND PRODUCTION_STATUS_NCID IN (" & gasProductionStatuses(i)
Rem        Else
Rem            sWhereClause = sWhereClause & " , " & gasProductionStatuses(i)
Rem        End If
        
        '---check for last item---
Rem        If i = UBound(gasProductionStatuses) Then
Rem            sWhereClause = sWhereClause & ")"
Rem        End If
Rem    Next i
    
    If ReminderType = RemindersDue Then  '????? 'viTabIndex = miDUE_REMINDERS_TAB Then
        sWhereClause = sWhereClause & ") OR (EVENT_TYPE_NCID = " & NCID_REMINDER_STATUS_EVENT & " AND STATUS_NCID = " & NCID_COLLECT_HISTORY & ")"
    End If
    
    oDas.GeneralCriteria.QueryQualifier = sWhereClause

    '---set up the context used to do most of the translations---
    sOldContext = gobjContext.SetDefault(NCID_3M_DEFAULT_CONTEXT)


'-- Run Query
'-----------------------------------------------------------------------------

    oDas.DoSearch
    
    i = 0

    Set collReminder = New Collection

'--  Loop through all of the datastrings and extract the information
'---------------------------------------------------------------------------
    
    While oDas.EOF = False
    
        '---set the skip indicator (reminder that
        '---should not apply to patient to false
        bSkipThis = False
         
        '---set up object pointers to this datastring---
        Set oDatastring = oDas.Events
        Set oClinicalEvent = oDatastring.PatientData
        Set oClinObsHeader = oClinicalEvent.Header
       
        If Left$(oClinObsHeader.ObsBatId.Value.CodedWOSform.Value(NCID_3M_DEFAULT_CONTEXT), 3) = "(A)" Then
            '---Reminder is for active duty
            '---Check for Active Duty Status
            If Not ActiveDuty() Then
                '---skip this record
                bSkipThis = True
            End If
        End If
                
        If Not bSkipThis Then
            
            i = i + 1   '--bump up string counter--
            
            '---make room to store the information---
            Set oWellnessReminder = New WellnessReminder
            'If i > 1 Then
            '    ReDim Preserve ratReminderInfo(1 To i)
            'Else
            '    ReDim ratReminderInfo(1 To i)
            'End If
            
            '---extract information about this datastring---
            oWellnessReminder.MarkedAsDeleted = False  'ratReminderInfo(i).bMarkedAsDeleted = False
            oWellnessReminder.EventId = oDatastring.DataId  'ratReminderInfo(i).sEventId = oDataString.DataId
            oWellnessReminder.ReminderDocNCID = oDatastring.DataTypeNCID 'ratReminderInfo(i).sASN1TypeNCID = oDataString.DataTypeNCID
            oWellnessReminder.TypeNCID = oClinObsHeader.ObsBatId.Value.CodedWOSform.ValueConcept.UniqueId 'ratReminderInfo(i).sTypeNCID = oClinObsHeader.ObsBatId.Value.CodedWOSform.ValueConcept.UniqueId
            
            If Left$(oClinObsHeader.ObsBatId.Value.CodedWOSform.Value(NCID_3M_DEFAULT_CONTEXT), 3) = "(A)" Then
                oWellnessReminder.ReminderText = Replace(oClinObsHeader.ObsBatId.Value.CodedWOSform.Value(NCID_3M_DEFAULT_CONTEXT), "(A)", "", 1, 3) 'ratReminderInfo(i).sTypeName = Replace(oClinObsHeader.ObsBatId.Value.CodedWOSform.Value(NCID_3M_DEFAULT_CONTEXT), "(A)", "", 1, 3)
            Else
                oWellnessReminder.ReminderText = oClinObsHeader.ObsBatId.Value.CodedWOSform.Value(NCID_3M_DEFAULT_CONTEXT) 'ratReminderInfo(i).sTypeName = oClinObsHeader.ObsBatId.Value.CodedWOSform.Value(NCID_3M_DEFAULT_CONTEXT)
            End If
            If oWellnessReminder.ReminderDocNCID = NCID_IMMUNIZATION_REMINDER_EVENT Then 'If ratReminderInfo(i).sASN1TypeNCID = NCID_IMMUNIZATION_REMINDER_EVENT Then

                '--  Immunization reminders have an optional "sequence number" that needs to be
                '--    concatenated onto the immunization type, i.e. MMR #1, Tetanus Booster, etc.
                '-----------------------------------------------------------------------------------
                '---find the "immunization reminder obs"---
                If FindObs(NCID_IMMUNIZATION_REMINDER_ID, oClinicalEvent.ClinObs, oClinicalObservation) Then
                    '---find the "med. admin sequence obs"---
                    If FindObs(NCID_MED_ADMIN_SEQ_ID, oClinicalObservation.ObsMods, oClinicalObservationMod) Then
                        '---save the sequence number NCID---
                        oWellnessReminder.SequenceNumberNCID = oClinicalObservationMod.ObsValue.Value.CodedWOSform.ValueConcept.UniqueId 'ratReminderInfo(i).sSeqNumberNCID = oClinicalObservationMod.ObsValue.Value.CodedWOSform.ValueConcept.UniqueId
                        '---concatenate the sequence number onto the vaccine "type" name---
                        oWellnessReminder.ReminderText = oWellnessReminder.ReminderText & " " & oClinicalObservationMod.ObsValue.Value.CodedWOSform.Value(NCID_3M_DEFAULT_CONTEXT) 'ratReminderInfo(i).sTypeName = ratReminderInfo(i).sTypeName & " " & oClinicalObservationMod.ObsValue.Value.CodedWOSform.Value(NCID_3M_DEFAULT_CONTEXT)
                    End If
                End If
                
                Set oClinicalObservationMod = Nothing
                Set oClinicalObservation = Nothing
            End If
            oWellnessReminder.EarliestDate = gdNULL_DATE 'ratReminderInfo(i).dEarlyDate = gdNULL_DATE
            oWellnessReminder.RecommendedDate = gdNULL_DATE 'ratReminderInfo(i).dRecDate = gdNULL_DATE
            '-----------------------------------------------------------------------------
            '  "Due Reminders" consists of three "types" of datastrings:
            '    1. Immunization Reminder Events
            '    2. Preventive Reminder Events
            '    3. Reminder Status Events
            '  Reminder Status Events do NOT have a "early date" and "recommended date".
            '    Extract information based upon the datastring string "type".
            '-----------------------------------------------------------------------------
            
            If oWellnessReminder.ReminderDocNCID <> NCID_REMINDER_STATUS_EVENT Then 'If ratReminderInfo(i).sASN1TypeNCID <> NCID_REMINDER_STATUS_EVENT Then
                '---find the "earliest allowed" obs--
                If FindObs(NCID_EARLIEST_ALLOWED_DATE_ID, oClinicalEvent.ClinObs, oClinicalObservation) Then
                    oWellnessReminder.EarliestDate = oClinicalObservation.ObsValue.Value.DateTime.DateTime 'ratReminderInfo(i).dEarlyDate = oClinicalObservation.ObsValue.Value.DateTime.DateTime
                End If
                
                '---find the "earliest recommended" obs--
                If FindObs(NCID_EARLIEST_REC_DATE_ID, oClinicalEvent.ClinObs, oClinicalObservation) Then
                    oWellnessReminder.RecommendedDate = oClinicalObservation.ObsValue.Value.DateTime.DateTime 'ratReminderInfo(i).dRecDate = oClinicalObservation.ObsValue.Value.DateTime.DateTime
                End If
                    
                '---extract the "last done date"---
                If FindObs(NCID_LAST_DONE_DATE_ID, oClinicalEvent.ClinObs, oClinicalObservation) Then
                    oWellnessReminder.LastServiceDate = oClinicalObservation.ObsValue.Value.DateTime.DateTime 'ratReminderInfo(i).dLastDoneDate = oClinicalObservation.ObsValue.Value.DateTime.DateTime
                End If
                
                Set oClinicalObservation = Nothing
            End If
            
            oWellnessReminder.StatusNCID = oClinObsHeader.TestStatus.Value.CodedWOSform.ValueConcept.UniqueId 'ratReminderInfo(i).sStatusNCID = oClinObsHeader.TestStatus.Value.CodedWOSform.ValueConcept.UniqueId
            If oWellnessReminder.StatusNCID = NCID_ACTIVE Then 'If ratReminderInfo(i).sStatusNCID = NCID_ACTIVE Then
                '-----------------------------------------------------------------------------------------
                '  If reminders are "active" you want to help the user distinguish which ones are due
                '    "today", as opposed to which ones are due "in the future".  This is dependent upon
                '    if its a "immunization" or "wellness" reminder.
                '-----------------------------------------------------------------------------------------
                If oWellnessReminder.ReminderDocNCID = NCID_IMMUNIZATION_REMINDER_EVENT Then 'If ratReminderInfo(i).sASN1TypeNCID = NCID_IMMUNIZATION_REMINDER_EVENT Then
                    If oWellnessReminder.EarliestDate <> gdNULL_DATE And oWellnessReminder.EarliestDate <> Now Then 'If ratReminderInfo(i).dEarlyDate <> gdNULL_DATE And ratReminderInfo(i).dEarlyDate <= Now Then
                        oWellnessReminder.Status = "Due Today" 'ratReminderInfo(i).sStatusName = "Due Today"
                    ElseIf oWellnessReminder.EarliestDate <> gdNULL_DATE And oWellnessReminder.EarliestDate <> Now Then 'ElseIf ratReminderInfo(i).dEarlyDate <> gdNULL_DATE And ratReminderInfo(i).dEarlyDate > Now Then
                        oWellnessReminder.Status = "Next Dose Due" 'ratReminderInfo(i).sStatusName = "Next Dose Due"
                    ElseIf oWellnessReminder.RecommendedDate <> gdNULL_DATE And oWellnessReminder.RecommendedDate <= Now Then 'ElseIf ratReminderInfo(i).dRecDate <> gdNULL_DATE And ratReminderInfo(i).dRecDate <= Now Then
                        oWellnessReminder.Status = "Due Today" 'ratReminderInfo(i).sStatusName = "Due Today"
                    ElseIf oWellnessReminder.RecommendedDate <> gdNULL_DATE And oWellnessReminder.RecommendedDate > Now Then 'ElseIf ratReminderInfo(i).dRecDate <> gdNULL_DATE And ratReminderInfo(i).dRecDate > Now Then
                        oWellnessReminder.Status = "Next Dose Due" 'ratReminderInfo(i).sStatusName = "Next Dose Due"
                    Else    '--unknown dates to compare against--
                        oWellnessReminder.Status = "Next Dose Due" 'ratReminderInfo(i).sStatusName = "Next Dose Due"
                    End If
                ElseIf oWellnessReminder.ReminderDocNCID = NCID_PREVENTIVE_REMINDER_EVENT Then 'ElseIf ratReminderInfo(i).sASN1TypeNCID = NCID_PREVENTIVE_REMINDER_EVENT Then
                    If oWellnessReminder.RecommendedDate <> gdNULL_DATE And oWellnessReminder.RecommendedDate <= Now Then 'If ratReminderInfo(i).dRecDate <> gdNULL_DATE And ratReminderInfo(i).dRecDate <= Now Then
                        oWellnessReminder.Status = "Due Today" 'ratReminderInfo(i).sStatusName = "Due Today"
                    ElseIf oWellnessReminder.RecommendedDate <> gdNULL_DATE And oWellnessReminder.RecommendedDate > Now Then 'ElseIf ratReminderInfo(i).dRecDate <> gdNULL_DATE And ratReminderInfo(i).dRecDate > Now Then
                        oWellnessReminder.Status = "Next Exam Due" 'ratReminderInfo(i).sStatusName = "Next Exam Due"
                    Else    '--unknown dates to compare against--
                        oWellnessReminder.Status = "Next Exam Due" 'ratReminderInfo(i).sStatusName = "Next Exam Due"
                    End If
                Else
                    '--if unsupported "type", just display "Due"--
                    oWellnessReminder.Status = "Due" 'ratReminderInfo(i).sStatusName = "Due"   '--if unsupported "type", just display "Due"--
                End If
            Else
                oWellnessReminder.Status = oClinObsHeader.TestStatus.Value.CodedWOSform.Value(NCID_3M_DEFAULT_CONTEXT) 'ratReminderInfo(i).sStatusName = oClinObsHeader.TestStatus.Value.CodedWOSform.Value(NCID_3M_DEFAULT_CONTEXT)
            End If
            
            '---extract the comments---
            Call GetHeaderComments(oClinObsHeader, atHeaderComments)
            oWellnessReminder.Comment = FormatHeaderComments(atHeaderComments) 'ratReminderInfo(i).sComment = FormatHeaderComments(atHeaderComments)
            collReminder.Add oWellnessReminder
        End If
            
        '---release reference to objects---
        Set oClinObsHeader = Nothing
        Set oClinicalEvent = Nothing
        Set oDatastring = Nothing
    
    
        '---move to the next datastring---
        oDas.GoToNext
    Wend
    
    '---restore the previous context---
    If Len(sOldContext) > 0 Then
        sOldContext = gobjContext.SetDefault(sOldContext)
    End If
    
    '---set the return value---
    DataGetReminders = True

    #If debugon Then
        DL.DebugLog Exe, Module, "BuildReminders", "P-End BuildReminders", DR
    #End If

    Exit Function
    
ErrHandler:
    '---show the VB error---
    gobjshared.ShowVBError Err.Number, Err.Description, "DataGetReminders.BuildReminders", "CHCSII_Reports", vbExclamation
    
    Exit Function
End Function






'****************************************************************************************
'  Function Name: FormatHeaderComments
'
'  Description: This function formats the passed in header comments into a single
'               text buffer.
'
'  Paramters:
'    ratCommentInfo() - Contains the header comments to format.
'****************************************************************************************
Private Function FormatHeaderComments(ratCommentInfo() As gtHeaderCommentInfo) As String
    Dim sBuf As String
    Dim i As Integer
    Dim j As Integer
    Dim iNumItems As Integer
    
    #If debugon Then
        DL.DebugLog Exe, Module, "FormatHeaderComments", "P-Begin FormatHeaderComments", DR
    #End If
    
    On Error GoTo ErrHandler
    
    '---set return value---
    FormatHeaderComments = ""
    
    '---initialize variables---
    iNumItems = UBound(ratCommentInfo)
    sBuf = ""
    
    '---loop through all comment items and format into a single buffer---
    For i = 1 To iNumItems
        '---format all "coded" items first, NOTE: This first loop prepares to "batch" translate the coded NCIDs---
        For j = 1 To UBound(ratCommentInfo(i).sCodedComment)
            '---prepare to translate the coded NCID---
            gobjConTrans.StoreNcid ratCommentInfo(i).sCodedComment(j), NCID_3M_DEFAULT_CONTEXT
        Next j
        
        '---get the representation for each coded comment item---
        For j = 1 To UBound(ratCommentInfo(i).sCodedComment)
            If Len(sBuf) > 0 Then
                sBuf = sBuf & ", " & gobjConTrans.Representation(ratCommentInfo(i).sCodedComment(j), NCID_3M_DEFAULT_CONTEXT)
            Else    '--first comment item--
                sBuf = gobjConTrans.Representation(ratCommentInfo(i).sCodedComment(j), NCID_3M_DEFAULT_CONTEXT)
            End If
        Next j
        
        '---format all "text" items next---
        For j = 1 To UBound(ratCommentInfo(i).sTextComment)
            If Len(sBuf) > 0 Then
                sBuf = sBuf & ", " & ratCommentInfo(i).sTextComment(j)
            Else    '--first comment item--
                sBuf = ratCommentInfo(i).sTextComment(j)
            End If
        Next j
    Next i
    
    '---set return value---
    FormatHeaderComments = sBuf
    
    #If debugon Then
        DL.DebugLog Exe, Module, "FormatHeaderComments", "P-End FormatHeaderComments", DR
    #End If
    
    Exit Function

ErrHandler:
    '---show the VB error---
    gobjshared.ShowVBError Err.Number, Err.Description, "FormatHeaderComments", "Wellness", vbExclamation
    
    Exit Function
End Function



