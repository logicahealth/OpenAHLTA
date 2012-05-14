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
Attribute VB_Name = "SO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const m_strListToolRecordColumns As String = "DataId, EncounterNumber, FacilityNCID, RTF, UpdateFlag, " & _
                                                     "CreatedBy, CreatedOn, UpdatedBy, UpdatedOn, SOURCENCID, ORIGINALSIZE "

'added SOURCENCID to SQL list WR 2/14/2002
Private Const m_strListToolRecordsColumns As String = "DataId, UpdateFlag, SnoID, Duration, Modifier, Onset, " & _
                                                      "Prefix, QLink, QLinkseq, RangeNormalHigh, RangeNormalLow, " & _
                                                      "RangeScale, ReferID, Result, Status, Unit, Value, Note, ChartFlag "

Private mobjEncounter As Encounter
Private mobjParent As SectionParent
Private mobjModule As Object
Private mobjIChild As IChild
Private msProgID As String
Private mbRegularViewRunning As Boolean
Private mNoteTool As CHCSII_NoteTool.Note

Private mobjICoreAPI3 As CHCSII_CoreAPI3.ICHCSII_CoreAPI3

Private mobjCurrNote As SONote
Private mobjCopyForwardNote As SONote
Private mcolNotes As Collection

Private Const Exe = "CHCSIIEncounterOps"
Private Const Module = "SO"
Private Const ExeModule = Exe & "." & Module
#If debugon Then
    Private DR As DebugRec
    Private DL As cDebugLog
#End If

Private mblnGreenSwitch As Boolean



Public Property Get CanDeleteWithoutHistory(ByVal Index As Integer) As Boolean
'per P. Wilkinson should always save history SCR 53837

    CanDeleteWithoutHistory = False
'    If Index <> mcolNotes.Count Then
'        Exit Sub
'    End If
'    If mcolNotes(Index).meType <> ListNoteType Then
'        Exit Sub
'    End If
'    If mcolNotes(Index).msOwnerNCID <> mobjEncounter.Logon.UserNCID Then
'        Exit Sub
'    End If
'    If mobjEncounter.Status = NeedsCoSignature _
'    Or mobjEncounter.Status = complete _
'    Or mobjEncounter.Status = Updated _
'    Or mobjEncounter.Status = Updating Then
'        Exit Sub
'    End If
'    CanDeleteWithoutHistory = True
End Property

Public Property Get CurrentNote() As SONote
    Set CurrentNote = mobjCurrNote
End Property

Public Property Get Count() As Integer
    Count = mcolNotes.Count
End Property

Public Function GetNotesList() As ADODB.Recordset

    Dim rsWorking As New ADODB.Recordset
    Dim nRow As Long
    Dim objNote As SONote
    Dim dTempDate As Date
    Dim sTemp As String
    Dim strMsg As String 'Add add'l debugging SCR-30609
    
    On Error GoTo ErrHandler
    
    strMsg = "Begin Function"
    
    Set GetNotesList = New ADODB.Recordset
    
    With rsWorking.Fields
        .Append "Title", adVarChar, 250
        .Append "Date", adDate
        .Append "Author", adVarChar, 65
        .Append "Key", adVarChar, 255
        .Append "Type", adVarChar, 10
        .Append "RTF", adVariant
        .Append "ReadOnly", adBoolean
        .Append "NewCouplerData", adBoolean  'SO Note that gets passed back to PKC in RunCouplers
        .Append "CouplerData", adBoolean
        
    End With
    rsWorking.Open
          
    strMsg = "Check Copy Forward"
    
    If Not CopyForwardNote Is Nothing Then
       rsWorking.AddNew
       rsWorking("Type") = "Structured"
       rsWorking("Title") = CopyForwardNote.msTitle
       rsWorking("Date") = CopyForwardNote.mdDTS
       rsWorking("Author") = CopyForwardNote.msOwnerName
       rsWorking("RTF") = CopyForwardNote.RTF
       rsWorking("Key") = CStr(CopyForwardNote.mnDataID)
       rsWorking.Update
    End If
    
    strMsg = "Check Collection"
    If Not mcolNotes Is Nothing Then
        For Each objNote In mcolNotes
             rsWorking.AddNew
             rsWorking("Type") = "Structured"
             rsWorking("Title") = objNote.msTitle
             rsWorking("Date") = objNote.mdDTS
             rsWorking("Author") = objNote.msOwnerName
             rsWorking("RTF") = objNote.RTF
             rsWorking("Key") = CStr(objNote.mnDataID)
             rsWorking("ReadOnly") = objNote.mbReadOnly
             rsWorking("NewCouplerData") = objNote.mbNewCouplerData
             rsWorking("CouplerData") = objNote.mbCouplerData
             rsWorking.Update
        Next
    End If
    
    'If Not mNoteTool Is Nothing Then
    '   With mNoteTool.NoteList
    '        For nRow = 1 To .Count
    '             rsWorking.AddNew
    '             rsWorking("Type") = "Text"
    '             rsWorking("Title") = .Item(nRow).Note_Title
    '             rsWorking("Date") = .Item(nRow).Date_Created
    '             rsWorking("Author") = .Item(nRow).Entered_By
    '             rsWorking("Key") = .Item(nRow).Note_Collection_Key
    '             rsWorking("RTF") = .Item(nRow).Note_Text
    '             rsWorking.Update
    '        Next nRow
    '   End With
    'End If
    
    rsWorking.Sort = "Date Asc"
    Set GetNotesList = rsWorking
    
    strMsg = "Exit Function"
    
    GoTo CleanUp
    
ErrHandler:
    Call mobjEncounter.CWShared.ShowVBError(Err.Number, Err.Description & " - " & strMsg, "SO.GetNotesList", "EncounterOps", vbCritical)
    'Resume Next
CleanUp:
    Set rsWorking = Nothing
    Set objNote = Nothing
    
End Function

Public Property Get SONote(ByVal Index As Integer) As SONote
    Set SONote = mcolNotes(Index)
End Property

Public Property Get NoteRTF(ByVal Index As Integer) As String
    Dim objNote As SONote
    On Error Resume Next
    If mcolNotes.Count < Index Then
        Exit Property
    End If
    Set objNote = mcolNotes(Index)
    NoteRTF = objNote.RTF
    Set objNote = Nothing
End Property

Public Property Get RTF() As String
    
    Dim rsRTF As ADODB.Recordset
    Dim sRTF As String
    Dim sLineHeader As String
    Dim objNote As SONote
    Dim uNote As Variant
    Dim sRTFHeader As String
    Dim strEndRtf As String
    Dim oShared As CWShared
    
    On Error Resume Next
    
'<< shaw. need to retrieve the rtf
    Set oShared = New CWShared

    sRTFHeader = "{\rtf1\ansi\ucl\deff0\deflang1033\deflangfe1033" _
                & "{\fonttbl{\f0\fswiss\fcharset0\fprq2 Arial;}}" _
                & "{\colortbl;\red0\green0\blue0;}"
    Set rsRTF = GetNotesList
    
    sRTF = ""
    
    'Add historical SO Data
    If oShared.IsAppMode(modeCHCSI_GUI) Then 'SCR 36493
        If mobjEncounter.TextNotes.Count > 0 Then 'SCR#48052
            sRTF = sRTF & mobjEncounter.TextNotes.RTF(, eSO) & "\par\par "  'SCR-38142
        End If
    End If

    If Not (rsRTF.BOF And rsRTF.EOF) Then
       rsRTF.MoveFirst
       Do Until rsRTF.EOF
           'sRTF = sRTF & " \par \pard\plain\f0\fs16\par "
           'Removed top spaces from S/O RTF SCR-18514
           sRTF = sRTF & "\pard\plain\f0\fs16"
           If rsRTF("READONLY") = "T" Then
              'Coupler Read Only Note
              sRTF = sRTF & "\pard\plain\f0\fs16\b\ul SO Note \b0 Cited by " _
                     & rsRTF("Author") & " @ " & FormatDTG(rsRTF("Date")) & "\par\ul0\b "
           ElseIf rsRTF("Type") = "Text" Then
              sRTF = sRTF & "\pard\plain\f0\fs16\b\ul SO Note \b0 Written by " _
                     & rsRTF("Author") & " @ " & FormatDTG(rsRTF("Date")) & "\par\ul0\b "
           End If

           If rsRTF("Title") <> "" Then sRTF = sRTF & rsRTF("Title") & " \par "
           'SF600 form is putting underscore characters in place of space-crlf
           'Doesn't happen in encounter doc, MS Word, TXText test app or elsewhere
           sRTF = sRTF & "\pard\plain \widctlpar\adjustright \cgrid" & Replace(rsRTF("RTF"), " " & vbCrLf, vbCrLf)
           strEndRtf = "\par "
           sRTF = sRTF & strEndRtf

           rsRTF.MoveNext
       Loop
       'Remove the last set of paragraph marks  added, let encounter handle spacing between sections
       If GreenSwitch Then
        If Right(sRTF, Len(strEndRtf)) = strEndRtf Then
         sRTF = Left(sRTF, Len(sRTF) - Len(strEndRtf))
        End If
       End If
       
    End If
                
    If sRTF <> "" Then
       RTF = sRTFHeader & sRTF & "}"
'
'
'            RTF = "{\rtf1\ansi\ucl\deff0\deflang1033\deflangfe1033" _
'                & "{\fonttbl{\f0\fswiss\fcharset0\fprq2 Arial;}}" _
'                & "{\colortbl;\red0\green0\blue0;}" _
'                & "\pard\plain\f0\fs16\b " _
'                & " You do not have the access priviledges to read this section." & "\b0\par } "
'            RTF = ""
    End If
End Property

'================================================================================
'
' C Atencio 9/26/01 SCR-17008 Pass in DataID to insure we deleting corrent SO Note
Public Function DeleteNote(ByVal Index As Integer, ByVal DataID As Long) As Boolean
    Dim objNote As SONote
    Dim objData As DataString
    Dim bWasLocked As Boolean
    Dim i As Integer
    Dim strMsg As String 'Add add'l debugging SCR-30609
    Dim AddHistory As Boolean 'True if history should be written
    
    On Error GoTo ErrHandler
    'Set objNote = mcolNotes(index)
    
    strMsg = "Begin Function"
    
    If Not CopyForwardNote Is Nothing And DataID = 0 Then
       Set mobjCopyForwardNote = Nothing
       DeleteNote = True
       Exit Function
    Else
        strMsg = "Look for SO Data ID"
        For i = 1 To mcolNotes.Count
            If mcolNotes.Item(i).mnDataID = DataID Then
                Set objNote = mcolNotes(i)
                'If index is not know, set index value
                If Index = 0 Then
                    Index = i
                End If
                Exit For
            End If
        Next
    End If
    
    strMsg = "Lock Section"
    bWasLocked = mobjEncounter.LockingObject.SectionIsLocked(eSO, objNote.mnDataID)
    If Not bWasLocked Then
        If Not mobjEncounter.LockingObject.LockSection(eSO, objNote.mnDataID, SectionLock, "Deleting Note") Then
            GoTo CleanUp
        End If
    End If
    strMsg = "Start Section Update"
    If Not mobjEncounter.LockingObject.StartSectionUpdate(eSO, objNote.mnDataID) Then
        GoTo CleanUp
    End If
    
    'CanDeleteWithoutHistory logic references the notes collection which will change after the delete
    'so go ahead and determine it's value SCR 53837 - Ed Clarke
    AddHistory = Not Me.CanDeleteWithoutHistory(Index)
    
'<< GEMS
    If Not mobjEncounter.CWShared.IsAppMode(modeCDR) Then 'SCR#48052
        If GEMS_Delete(objNote) Then
            DeleteNote = True
            mcolNotes.Remove i
        End If
    Else
        strMsg = "Initialize Data String"
        Set objData = InitializeDataString(objNote)

        'Note may have been deleted previously SCR-26756
        strMsg = "Delete Note"
        If Not objData Is Nothing Then
            objData.DeleteFlag = 1
            If objData.Save Then
                DeleteNote = True
                mcolNotes.Remove i
            End If
        Else
            MsgBxARMd "Unable to delete S/O note. S/O note was deleted by another user.", vbExclamation, "Encounter S/O Note"
            GoTo CleanUp
        End If
    End If
    
    'Add history here since S/O note may have already been deleted SCR-26756
    strMsg = "Add Change History"
    If AddHistory Then
        If Not mobjEncounter.AddHistoryItem(Deletion, "SO Note", objNote.RTF) Then
            MsgBxARMd "Could not add to encounter change history section. Did not delete note.", vbCritical, "Encounter Text Note"
            GoTo CleanUp
        End If
    End If
    
    GoTo CleanUp
ErrHandler:
    Call mobjEncounter.CWShared.ShowVBError(Err.Number, Err.Description & " - " & strMsg, "SO.DeleteNote ", "EncounterOps", vbCritical)
CleanUp:
    On Error Resume Next
    Call mobjEncounter.LockingObject.EndSectionUpdate(eSO, objNote.mnDataID)
    If Not bWasLocked Then
        Call mobjEncounter.LockingObject.UnlockSection(eSO, objNote.mnDataID)
    End If
    Set objData = Nothing
    If DeleteNote Then
        mobjEncounter.Comm.SetStatusText "Encounter " & mobjEncounter.EncounterID & " SO Note was deleted."
    Else
        mobjEncounter.Comm.SetStatusText ""
    End If
End Function

Public Sub EditNewNote(ByVal NoteType As SONoteTypeEnum)
    Set mobjCurrNote = New SONote
    mobjCurrNote.meType = NoteType
    mobjCurrNote.msOwnerName = mobjEncounter.Logon.UserName
    mobjCurrNote.msOwnerNCID = mobjEncounter.Logon.UserNCID
    mobjCurrNote.mdDTS = Now
    Select Case NoteType
    Case FormNoteType
        Call StartChartNotes
    Case ListNoteType
        Set mobjCurrNote.mcolListRecords = New Collection
        Call StartMedcin(True)
    End Select
End Sub

Public Sub EditNote(ByRef Note As SONote)
    Dim objNewNote As SONote
    Dim bNeedHistory As Boolean
    On Error Resume Next
    'Check to make sure note exists in CDR (has not been deleted) SCR-26756
    If Not CheckIfNoteExists(Note) Then
        MsgBxARMd "Unable to edit S/O Note. S/O Note has deleted by another user.", vbExclamation, "Encounter S/O Note"
        Exit Sub
    End If
    
    If Note.msOwnerNCID = mobjEncounter.Logon.UserNCID Then
        If Not mobjEncounter.LockingObject.LockSection(eSO, Note.mnDataID, SectionLock) Then
            Exit Sub
        End If
        GoTo DoEdit
    End If
    If Note.meType = FormNoteType Then
        MsgBxARMd "This Form Tool Note belongs to and can only be edited by " & Note.msOwnerName & "."
        Exit Sub
    End If
    Select Case MsgBxARMd("This SO Note belongs to " & Note.msOwnerName & "." & vbCrLf & vbCrLf _
                    & "Do you want to create your own note starting with the contents of this one?", _
                    vbYesNoCancel, "Encounter Note Ownership")
    Case vbYes
        Set objNewNote = Note.Clone()
        With objNewNote
            .msOwnerNCID = mobjEncounter.Logon.UserNCID
            .msOwnerName = mobjEncounter.Logon.UserName
            .mdDTS = Now
        End With
        Set Note = objNewNote
        Set objNewNote = Nothing
        GoTo DoEdit
    Case vbNo
        Select Case MsgBxARMd("Do you want to copy the existing note to Change History and take over ownership of it?", _
                        vbYesNo, "Encounter Note Ownership")
        Case vbNo
            Exit Sub
        Case vbYes
            bNeedHistory = True
        End Select
    Case vbCancel
        Exit Sub
    End Select
DoEdit:
    If mobjEncounter.Status = NeedsCoSignature _
    Or mobjEncounter.Status = complete _
    Or mobjEncounter.Status = Updated _
    Or mobjEncounter.Status = Updating _
    Or bNeedHistory Then
        If Not mobjEncounter.LockingObject.LockSection(eSO, Note.mnDataID, SectionLock) Then
            Exit Sub
        End If
        If Not mobjEncounter.AddHistoryItem(Overwritten, "SO Note", Note.RTF) Then
            MsgBxARMd "Unable to copy section to change history.", vbCritical, "Encounter"
            Call mobjEncounter.LockingObject.UnlockSection(eSO, Note.mnDataID)
            Exit Sub
        End If
        Note.msOwnerNCID = mobjEncounter.Logon.UserNCID
        Note.msOwnerName = mobjEncounter.Logon.UserName
        Note.mdDTS = Now
        If Not SaveNote(Note) Then
            Call mobjEncounter.LockingObject.UnlockSection(eSO, Note.mnDataID)
            Exit Sub
        End If
    End If
    Set mobjCurrNote = Note
    
    'Set edited date SCR#50966
    mobjCurrNote.mdDTS = Now
    
    Select Case mobjCurrNote.meType
    Case FormNoteType
        Call StartChartNotes
    Case ListNoteType
        Call StartMedcin
    End Select
End Sub

Private Sub StartChartNotes()
    On Error Resume Next
    Set mobjModule = mobjEncounter.Comm.InitializeOLEServer(NCID_BROWSER_APPLICATION)
    Set mobjIChild = mobjModule
    If mobjIChild Is Nothing Then
        Set mobjModule = Nothing
        MsgBxARMd "The ChartNotes application module does not support the EncounterOps.IChild interface.", vbInformation, "Encounter"
        Set mobjModule = Nothing
        Exit Sub
    End If
    Set mobjICoreAPI3 = mobjModule
    Call mobjIChild.SetParent(mobjParent)
    Call mobjEncounter.Comm.Message(cwiGET_APPLICATION_PROGID, msProgID, NCID_CW_APPLICATION, NCID_BROWSER_APPLICATION)
    mbRegularViewRunning = mobjEncounter.EncounterParent.OpenChildView(mobjModule)
    Call mobjModule.UpdateEncounterInfo(mobjCurrNote.mdDTS, _
                            mobjCurrNote.msOwnerNCID, _
                            mobjEncounter.FacilityNCID, _
                            "", _
                            mobjEncounter.EncounterID, _
                            mobjEncounter.FacilityNCID, _
                            "")
    mobjModule.Message cwiGENERIC, "AUTOSAVE|" & mobjEncounter.EncounterParent.EncounterOps.EncProperties.AutoSaveInt, NCID_ENCOUNTER_APPLICATION
    If mobjCurrNote.mnDataID <> 0 Or mobjCurrNote.msProfileNCID <> "" Then
        mobjModule.Message cwiEDIT_STRING, mobjCurrNote.mnDataID & "|" & mobjCurrNote.msProfileNCID, NCID_ENCOUNTER_APPLICATION
    End If
    Set mobjEncounter.CurrSection = mobjParent
End Sub

Private Sub StartMedcin(Optional ByVal LoadTemplate As Boolean = True)
    On Error Resume Next
    Set mobjModule = mobjEncounter.Comm.InitializeOLEServer(NCID_MEDCIN_APPLICATION)
    Set mobjIChild = mobjModule
    If mobjIChild Is Nothing Then
        Set mobjModule = Nothing
        MsgBxARMd "The Medcin application module does not support the EncounterOps.IChild interface.", vbInformation, "Encounter"
        Set mobjModule = Nothing
        Exit Sub
    End If
    Set mobjICoreAPI3 = mobjModule
    Call mobjIChild.SetParent(mobjParent)
    Call mobjEncounter.Comm.Message(cwiGET_APPLICATION_PROGID, msProgID, NCID_CW_APPLICATION, NCID_MEDCIN_APPLICATION)
    mbRegularViewRunning = mobjEncounter.EncounterParent.OpenChildView(mobjModule)
    mobjModule.Message cwiGENERIC, "AUTOSAVE|" & mobjEncounter.EncounterParent.EncounterOps.EncProperties.AutoSaveInt, NCID_ENCOUNTER_APPLICATION
    
    Dim lngHistCount As Long
    lngHistCount = mobjModule.LoadTemplateHistory
    
    'If lngHistCount <= 1 Or Not GreenSwitch Then
      If LoadTemplate Then
        If mobjCurrNote.mbCouplerData Then
            Dim oTemplate As Object
            Dim oListTemplate As Object
            Set oListTemplate = CreateObject("CHCSII_TemplateClient.ListTemplate")
            If oListTemplate.LoadListToolRecords(mobjCurrNote.mcolListRecords) Then
                mobjModule.CouplersListTemplate = oListTemplate 'oListTemplate.ListNoteItems
            End If
        End If
        'Only perform if non-GREEN, otherwise, let Template navigation control handle which templates to load
        If Not GreenSwitch Then
           Call mobjModule.LoadEncounterTemplate
        End If
      End If
    'End If
    
End Sub

Public Sub Load(ByRef objParent As SectionParent, ByRef colDataStrings As Collection)
    
    Dim objData As DATASTRATLLib.DataString
    Dim objClinicalText As CLINTEXTATLLib.ClinicalText
    Dim objEvent As CLINTYPEATLLib.ClinicalEvent
    Dim objCommentGroup As CLINTYPEATLLib.CommentGroup
    Dim objCommentItem As CLINTYPEATLLib.CommentItem
    Dim objClinObs As CLINTYPEATLLib.ClinicalObservation
    Dim objLTR As ListToolRecord
    Dim objNote As SONote
    Dim sName As String
    Dim sValue As String
    Dim i As Long
    Dim j As Long
    Dim k As Long
    
    Set mcolNotes = New Collection
    
    'Using the For...Each method with 3M objects causes memory leak
    For i = 1 To colDataStrings.Count
    'For Each objData In colDataStrings
        Set objData = colDataStrings(i)
        Select Case objData.DataTypeNCID
        Case NCID_DAS_CLINICAL_TEXT
            Set objNote = New SONote
            objNote.meType = FormNoteType
            objNote.mnDataID = objData.DataID
            If Not objData.ModifyClinician Is Nothing Then
                objNote.msOwnerNCID = objData.ModifyClinician.ValueConcept.UniqueId
                objNote.msOwnerName = objData.ModifyClinician.ValueConcept.PrefRep("2000").Representation
            ElseIf Not objData.CreateClinician Is Nothing Then
                objNote.msOwnerNCID = objData.CreateClinician.ValueConcept.UniqueId
                objNote.msOwnerName = objData.CreateClinician.ValueConcept.PrefRep("2000").Representation
            End If
            objNote.msProfileNCID = objData.DataEntryModel.ValueConcept.UniqueId
            Set objClinicalText = objData.PatientData
            objNote.msRTF = ConvertHTMLtoRTF(objClinicalText.ReportText.TextObject.Text.Text)
'            objNote.msOwnerNCID = objClinicalText.ContextInfo.DictatingClinician.ValueConcept.UniqueId
'            objNote.msOwnerName = objClinicalText.ContextInfo.DictatingClinician.ValueConcept.PrefRep("2000").Representation
            objNote.mdDTS = objClinicalText.EventDate
            'Ignore entries that have a report type of 69587 as they are generic text notes
            If objData.PatientData.ReportType.ValueConcept.UniqueId = 69587 Then Set objNote = Nothing
        Case NCID_ENCEVENT
            Set objEvent = objData.PatientData
            If objEvent.Header.ObsBatId.value.Coded.ValueConcept.UniqueId = SectionEnumToNCID(eSO) Then
                Set objNote = New SONote
                objNote.meType = ListNoteType
                Set objNote.mcolListRecords = New Collection
                objNote.mnDataID = objData.DataID
                If Not objData.ModifyClinician Is Nothing Then
                    objNote.msOwnerNCID = objData.ModifyClinician.ValueConcept.UniqueId
                    objNote.msOwnerName = objData.ModifyClinician.ValueConcept.PrefRep("2000").Representation
                ElseIf Not objData.CreateClinician Is Nothing Then
                    objNote.msOwnerNCID = objData.CreateClinician.ValueConcept.UniqueId
                    objNote.msOwnerName = objData.CreateClinician.ValueConcept.PrefRep("2000").Representation
                End If
                objNote.mdDTS = objEvent.Header.ObsDateTime.StartTime.value.DateTime
                For j = 1 To objEvent.Header.Comments.Count
                'For Each objCommentGroup In objEvent.Header.Comments
                    Set objCommentGroup = objEvent.Header.Comments(j)
                    For k = 1 To objCommentGroup.Comments.Count
                    'For Each objCommentItem In objCommentGroup.Comments
                        Set objCommentItem = objCommentGroup.Comments(k)
                        Call ParseCommentItem(objCommentItem, sName, sValue)
                        Select Case sName
                        Case "TITLE"
                            objNote.msTitle = sValue
                        Case "READONLY"
                            If sValue = "T" Then
                                objNote.mbReadOnly = True
                            End If
                        Case "NEWCOUPLERDATA"
                             If sValue = "T" Then
                                objNote.mbNewCouplerData = True
                            End If
                        Case "COUPLERDATA"
                             If sValue = "T" Then
                                objNote.mbCouplerData = True
                            End If
                        End Select
                    Next 'objCommentItem
                Next 'objCommentGroup
                For j = 1 To objEvent.ClinObs.Count
                    Set objClinObs = objEvent.ClinObs.Item(j)
                    Select Case objClinObs.ObsId.value.CodedWOSform.ValueConcept.UniqueId
                    Case NCID_ENCOUNTERRTF
                        objNote.msRTF = ExtractRTF(objClinObs)
                    Case NCID_MEDCIN_OBS
                        Set objLTR = New ListToolRecord
                        Call objLTR.ReadFromClinObs(objClinObs)
                        objNote.mcolListRecords.Add objLTR
                        Set objLTR = Nothing
                    End Select
                Next j
            End If
        End Select
        If Not objNote Is Nothing Then
            ' the events are given to use in descending order, so reverse them
            If mcolNotes.Count > 0 Then
                mcolNotes.Add objNote, CStr(objNote.mnDataID), 1
            Else
                mcolNotes.Add objNote, CStr(objNote.mnDataID)
            End If
            Set objNote = Nothing
        End If
    Next 'objData
    
    Set objData = Nothing
    Set objClinicalText = Nothing
    Set objEvent = Nothing
    Set objCommentGroup = Nothing
    Set objCommentItem = Nothing
    Set objClinObs = Nothing
    Set objLTR = Nothing
    Set objNote = Nothing

End Sub

Public Sub MenuItemSelected(ByVal MenuID As Long)
    On Error Resume Next
    Call mobjModule.MenuItemSelected(MenuID)
End Sub

Public Sub OpenEncounter(ByRef objParent As SectionParent, ByRef objEncounter As Encounter)
    Set mobjParent = objParent
    Set mobjEncounter = objEncounter

    'With mobjEncounter
    '    If mNoteTool Is Nothing Then
    '       Set mNoteTool = New CHCSII_NoteTool.Note
    '       mNoteTool.Initialize .Comm, .Logon, .Config, .Patient, NCID_MEDCIN_APPLICATION
    '    End If
    '    mNoteTool.RetrieveNoteList .EncounterID
    'End With
    
    If objEncounter.GetCommObj.CmdLineSwitch("GREEN") <> "" Then
      GreenSwitch = True
    Else
      GreenSwitch = False
    End If
End Sub

Public Sub CloseEncounter()
    Set mobjParent = Nothing
    Set mobjEncounter = Nothing
    Set mobjICoreAPI3 = Nothing
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' SCR #6890 Refresh issues.                              '
' Purpose:                                               '
' Refresh the windows of a section already active.       '
' Parameters:                                            '
' None.                                                  '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Public Function RefreshSection() As Boolean
'
'    On Error Resume Next
'    If Not mbRegularViewRunning Then
'        RefreshSection = False '999
'        Exit Function
'    End If
'
'    If Not mobjEncounter.LockingObject.LockSection(eSO, 0, SectionLock) Then
'        Exit Function
'    End If
'
'    mobjIChild.SetParent mobjParent
'
'    mbRegularViewRunning = mobjEncounter.EncounterParent.RefreshChildView(mobjModule)
'
'    If mbRegularViewRunning Then
'        RefreshSection = True
'    Else
'        mobjIChild.SetParent Nothing
'        mobjEncounter.LockingObject.UnlockSection eSO, 0
'    End If
'
'End Function
Public Function OpenSection() As Boolean
    On Error Resume Next
'<<SCR 10673,10675(dup) Encounter data refresh before children load
'   Dropping this line to stop the refresh
'   mobjEncounter.RefreshDisplay
'>>

    If Not mobjEncounter.SectionIsEditable(eSO) Then 'SCR#44686
        MsgBxARMd "This encounter has been saved by another user and you can no longer edit the SO section.", vbInformation, "Encounter"
        Exit Function
    End If
        
    'Now open SO module in PGUI mode SCR#48052
    'If Not mobjEncounter.CWShared.IsAppMode(modeCHCS I_GUI) Then '<SCR 36493

    Set mobjCurrNote = Nothing
    
    If mcolNotes.Count = 0 And mobjCopyForwardNote Is Nothing Then
        Call EditNewNote(ListNoteType)
    ElseIf mcolNotes.Count = 1 And mobjEncounter.BypassSODialog Then
        If mcolNotes(1).msOwnerNCID = mobjEncounter.Logon.UserNCID Then
           EditNote mcolNotes(1)
        Else
           Call frmSelectSO.Initialize(mobjEncounter)
           mobjEncounter.Comm.Message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_ENCOUNTER_APPLICATION
           '--- SCR 22510, 22261 Stop all Tx when modal dlg MSolano 4/24/02
           goARM.StopAllCurrTx enmARMTxStatusFailed, enmCHCSIIDlgDetect_CustDlg, "EncounterOps::frmSelectSO"
           frmSelectSO.Show vbModal
           mobjEncounter.Comm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_ENCOUNTER_APPLICATION
        End If
    Else
        Call frmSelectSO.Initialize(mobjEncounter)
        mobjEncounter.Comm.Message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_ENCOUNTER_APPLICATION
        '--- SCR 22510, 22261 Stop all Tx when modal dlg MSolano 4/24/02
        goARM.StopAllCurrTx enmARMTxStatusFailed, enmCHCSIIDlgDetect_CustDlg, "EncounterOps::frmSelectSO"
        frmSelectSO.Show vbModal
        mobjEncounter.Comm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_ENCOUNTER_APPLICATION
    End If
    
    mobjEncounter.BypassSODialog = False
    
    If Not mobjCurrNote Is Nothing Then
        Call mobjEncounter.LockingObject.LockSection(eSOOpen, mobjEncounter.Logon.UserNCID, SectionLock, , True)
        OpenSection = True
    End If
    
'    Else
''        If mobjEncounter.TextNotes.Count Is Nothing Then
''            Set mcolNotes = New Collection
''        End If
'
''<< shaw - edit only the not-yet-saved text note.  If we read it in, it will be marked as readonly.
''           and the user can only ever edit the last textnote.
''        Reenable S/O for telcon in PGUI 2.0 SCR-40959
''        If mobjEncounter.EncounterType = TeleConsult Then
''            'Disable SO for Telcons SCR-38789
''            MsgBxARMd "The S/O Module is not available when accessed from a Telcon Encounter.", vbInformation, "Encounter"
''            Exit Function
''        End If
'        With mobjEncounter.TextNotes
'            If .Count = 0 Then
'                mobjEncounter.EditNewTextNote eSO
'            Else
'                If .TextNote(.Count).ReadOnly Then
'                    mobjEncounter.EditNewTextNote eSO
'                Else
'                    Call mobjEncounter.EditTextNote(.TextNote(.Count))
'                End If
'            End If
'        End With
'        OpenSection = False
'
'    End If
End Function

Public Property Get CopyForwardNote() As SONote
    Set CopyForwardNote = mobjCopyForwardNote
End Property

Public Property Let CopyForwardNote(ByVal RHV As SONote)
    Set mobjCopyForwardNote = RHV
End Property


Public Function CloseSection() As Boolean

'''''''''''''''''''''''''''''''''
' Testing Refresh.              '
' SCR 11399  Unwanted Refresh.  '
'''''''''''''''''''''''''''''''''
    On Error Resume Next
    
    #If debugon Then
        Dim DR As DebugRec
        If Not DL Is Nothing Then DL.DebugLog Exe, Module, "CloseSection", "Begin", DR
    #End If
    
    If mbRegularViewRunning Then
    
        If mobjICoreAPI3 Is Nothing Then
            Call mobjModule.CloseView(cwiREGULAR_VIEW)
            CloseSection = (Not mbRegularViewRunning)
        Else
            CloseSection = mobjICoreAPI3.CloseView2(cwiREGULAR_VIEW)
        End If
        
    #If debugon Then
        If Not DL Is Nothing Then DL.DebugLog Exe, Module, "CloseSection", "End of CloseView", DR
    #End If
        
        If CloseSection Then
            If Not mobjCurrNote Is Nothing Then
                Call mobjEncounter.LockingObject.UnlockSection(eSO, mobjCurrNote.mnDataID)
            End If
            
            Set mobjModule = Nothing
            Set mobjICoreAPI3 = Nothing
            Set mobjIChild = Nothing
            Set mobjCurrNote = Nothing
        End If
    
    End If
    #If debugon Then
        If Not DL Is Nothing Then DL.DebugLog Exe, Module, "CloseSection", "End", DR
    #End If

End Function

Public Sub ClosedView() ' called by child through SectionParent
    On Error Resume Next
    If mbRegularViewRunning Then
        mbRegularViewRunning = False
        Call mobjIChild.SetParent(Nothing)
        Call mobjEncounter.EncounterParent.hidechild
        If Not mobjCurrNote Is Nothing Then
            Call mobjEncounter.LockingObject.UnlockSection(eSO, mobjCurrNote.mnDataID)
        End If
        Call mobjEncounter.LockingObject.UnlockSection(eSOOpen, mobjEncounter.Logon.UserNCID)
        
        'Set mobjModule = Nothing
        'Set mobjIChild = Nothing
        Set mobjCurrNote = Nothing
    End If
End Sub

Public Function SaveNote(ByRef objNote As SONote, Optional ByVal blnIgnoreStatus As Boolean = False) As Boolean
' Assume section NOT already locked for update before getting here
' Only ListNotes are saved here... chartnotes saves its own data, and we read it from the CDR
    Dim objData As DataString
    Dim bWasNew As Boolean
    
    On Error GoTo ErrHandler

    'Check to make sure note exists in CDR (has not been deleted) SCR-26756
    If Not CheckIfNoteExists(objNote) And objNote.DataID > 0 Then
        MsgBxARMd "Unable to save S/O Note. S/O Note has deleted by another user.", vbExclamation, "Encounter S/O Note"
        Exit Function
    End If

    'Refresh Encounter Status
    mobjEncounter.Refresh True
    
    'If status in NeedCoSignature and user is PP or CS then save note SCR-19388
    If (mobjEncounter.Status = complete) Or _
        ((mobjEncounter.Status = NeedsCoSignature) And Not (mobjEncounter.UserIsPrimaryProvider Or mobjEncounter.UserIsCosigner)) Or _
        (mobjEncounter.Status = Updated) Or _
        ((mobjEncounter.Status = Updating) And Not (mobjEncounter.UserIsPrimaryProvider Or mobjEncounter.UserIsCosigner)) Then
        
        If Not blnIgnoreStatus Then 'SCR 17773 This was added for the data transfer.  You want to ingnore the status if
          mobjEncounter.Refresh                'this method is being called during a data transfer
         
          MsgBxARMd "Your lock on the S/O section of the current encounter has been taken away by another user." & vbCrLf & _
                  "This encounter has been completed and signed by " & mobjEncounter.WhoSigned.FullName & " on " & Format$(mobjEncounter.WhenSigned, gsDefaultDateTime) & ".  This S/O note will not be saved to the encounter.", vbExclamation + vbOKOnly, "Encounter"
                  
          Exit Function
        End If
        
    End If
    
    bWasNew = objNote.mnDataID = 0
    SaveNote = GEMS_SaveNote(objNote)
    GoTo CleanUp

ErrHandler:
    If Err = -2147220874 Or Err = -2147417848 Or Err = -2147220877 Then
       MsgBxARMd "Unable to save the current note. The note that you are attempting to save is too large. Please save some of the notes contents to a new note", vbOKOnly + vbInformation, "EncounterOps"
       Err.Clear
    Else
       Call mobjEncounter.CWShared.ShowVBError(Err.Number, Err.Description, "SONote.Save ", "EncounterOps", vbCritical)
    End If
CleanUp:
    On Error Resume Next
    If bWasNew Then
        mcolNotes.Add objNote, CStr(objNote.mnDataID)
        Call mobjEncounter.LockingObject.LockSection(eSO, objNote.mnDataID, SectionLock)
    Else
        'Only for theater, collection has old note, so remove and then add again SCR-24107
        mcolNotes.Remove CStr(objNote.mnDataID)
        mcolNotes.Add objNote, CStr(objNote.mnDataID)
        Call mobjEncounter.LockingObject.EndSectionUpdate(eSO, objNote.mnDataID)
    End If
    Set objData = Nothing
    If SaveNote Then
        mobjEncounter.Comm.SetStatusText "Encounter " & mobjEncounter.EncounterID & " SO Note was saved."
    Else
        mobjEncounter.Comm.SetStatusText ""
    End If
End Function

Private Function InitializeDataString(ByRef Note As SONote) As DataString
    Dim objData As DataString
    Dim objLTR As ListToolRecord
    Dim objEvent As ClinicalEvent
    Dim sTemp As String
    
    Set objData = mobjEncounter.InitializeDataString(eSO, Note.mnDataID, Note.mdDTS, Note.mdDTS)
    If Not objData Is Nothing Then 'Note may already be deleted SCR-26756
        Set objEvent = objData.PatientData
        Call AddRTF(objEvent, Note.msRTF)
        
        If Len(mobjEncounter.Comm.CmdLineSwitch("PKC")) Then 'SCR#44241
            If Note.mbReadOnly Then
                sTemp = "T"
            End If
            Call AddCommentValuePair(objEvent.Header, "READONLY", sTemp)
            sTemp = vbNullString
            If Note.mbNewCouplerData Then
                sTemp = "T"
            End If
            Call AddCommentValuePair(objEvent.Header, "NEWCOUPLERDATA", sTemp)
            sTemp = vbNullString
            If Note.mbCouplerData Then
                sTemp = "T"
            End If
            Call AddCommentValuePair(objEvent.Header, "COUPLERDATA", sTemp)
        End If
        
        If Not Note.mcolListRecords Is Nothing Then
            For Each objLTR In Note.mcolListRecords
                Call objLTR.WriteToClinObs(objData.PatientData)
                'Check if any of the notes has family history
                'If objLTR.prefix = "F" Then
                '    mobjEncounter.HasFamilyHx = True
                'End If
            Next objLTR
        End If
    End If
    
    Set InitializeDataString = objData
    Set objData = Nothing
End Function


Private Function GEMS_SaveNote(ByRef Note As SONote) As Boolean
   On Error GoTo Proc_Error
'- save the note (medcin data) to the database
   Dim objSQLEx           As ICHCSII_SQL.ISqlOpsEx
   Dim rsNote           As ADODB.Recordset
   Dim strSQL           As String
   Dim objLTR           As ListToolRecord

   'Check if encounter owner is Dr. Unassigned SCR-24349
   Call mobjEncounter.CheckEncounterOwnership
        
    If mobjEncounter.CWShared.IsAppMode(modeCHCSI_GUI) Then
        GEMS_SaveNote = PGUI_SaveNote(Note)
        Exit Function
    End If
    
    Set objSQLEx = gobjCHCSIIConn.CHCSII_SQLOPS_EX(Auto)  '''<SCR 36493
    
    strSQL = "SELECT DataID, FacilityNCID, EncounterNumber, Enc_SectionsIndex, Doc, OwnerName, OwnerNCID, OriginalSize " & _
             "FROM Enc_Sections " & _
             "WHERE EncounterNumber = " & Str(mobjEncounter.EncounterID) & " " & _
             "AND Enc_SectionsIndex = " & SectionEnumToNCID(eSO)
              
    
    Call objSQLEx.Execute(strSQL)
    
    With objSQLEx
       If (.BOF And .EOF) Then
          .AddNew
          !DataID = objSQLEx.GetNextID(exMAINSEQ)  ' GEMS_Seq_NextVal("MainSEQ")
          !FacilityNCID = mobjEncounter.FacilityNCID
          !EncounterNumber = CStr(mobjEncounter.EncounterID)
          !enc_sectionsindex = SectionEnumToNCID(eSO)
          !OwnerNCID = Note.msOwnerNCID 'SCR-22999
          !OwnerName = Note.msOwnerName 'SCR-22999
           End If
       'RTF saved below in ListToolRecord
       '!Doc = Note.msRTF
       .Update
    End With
    
   If Note.mnDataID < 1 Then
      Note.mnDataID = objSQLEx.GetNextID(exMAINSEQ) 'GEMS_Seq_NextVal("MainSEQ")
   End If
   
   strSQL = "SELECT " & m_strListToolRecordColumns & " FROM ListToolRecord WHERE DataId = " & CStr(Note.mnDataID)
            
   Call objSQLEx.Execute(strSQL)
   
   With objSQLEx
      If (.BOF And .EOF) Then
         Note.mnDataID = objSQLEx.GetNextID(exMAINSEQ) 'GEMS_Seq_NextVal("MainSEQ")
         .AddNew
         !DataID = Note.mnDataID
         !UpDateFlag = "N"
         !CreatedBy = Note.msOwnerNCID
         !CreatedOn = Now
         !SourceNCID = SectionEnumToNCID(eSO)
      Else
         !UpDateFlag = "U"
      End If

    !UpdatedBy = Note.msOwnerNCID
    !UpdatedOn = Now
    
    !EncounterNumber = mobjEncounter.EncounterID
    !FacilityNCID = mobjEncounter.FacilityNCID
    'Compress RTF SCR-24683
    !OriginalSize = Len(Note.msRTF)
    Dim objZlib As EncZLib
    Dim aBytes() As Byte
    Set objZlib = New EncZLib
    
    aBytes = StrConv(Replace(Note.msRTF, Chr(0), ""), vbFromUnicode)
    objZlib.CompressData aBytes
    !RTF = aBytes
    '!RTF = CompressRTF(Note.msRTF)
    .Update
   End With
   
    '-- Rebuild the details
    '   Mark all current as deleted
    '   add any new ones (ltr's without dataids)
    '   update any 'old' ones
    '   that leaves all the deleted ones marked as deleted
    
    objSQLEx.Execute "DELETE FROM ListToolRecords where DataID = " & Note.mnDataID
    
    If Not Note.mcolListRecords Is Nothing Then 'May have blank SO Note SCR#44787
        For Each objLTR In Note.mcolListRecords
    
          Call objSQLEx.Execute("SELECT " & m_strListToolRecordsColumns & " FROM ListToolRecords WHERE DataID = -1")
          
          With objSQLEx
              If (.BOF And .EOF) Then
                 .AddNew
                 !DataID = Note.mnDataID
                 !UpDateFlag = "N"
                 !CreatedBy = Note.msOwnerNCID
                 !CreatedOn = Now
              Else
                 !UpDateFlag = "Y"
              End If
                
              Call objLTR.GEMS_WriteToClinObs(objSQLEx)
              
        
              !UpdatedBy = Note.msOwnerNCID
              !UpdatedOn = Now
              
              .Update
          End With
        Next 'objLTR
    End If

    GEMS_SaveNote = True

    
Proc_Exit:
   Set objLTR = Nothing
   Set rsNote = Nothing
   Set objSQLEx = Nothing
   Exit Function
    
Proc_Error:
   GEMS_SaveNote = False
   Err.Raise Err.Number, Err.Source, Err.Description
   GoTo Proc_Exit
End Function

'Save PGUI Note SCR#48052
Private Function PGUI_SaveNote(ByRef Note As SONote) As Boolean
    Dim oSQLGen         As SQLGen
    Dim objSQL          As ICHCSII_SQL.ISqlOps
    Dim objLTR          As ListToolRecord
    Dim strSQL          As String
    
    On Error GoTo ErrHandler
    
    Set oSQLGen = New SQLGen
    'ADD SECTION DATA
    AddEncSectionData Note.mnDataID, mobjEncounter.FacilityNCID, CStr(mobjEncounter.EncounterID), SectionEnumToNCID(eSO), Note.msOwnerNCID, _
                        Note.msOwnerName, Note.mdDTS, vbNullString, vbNullString, vbNullString, 0, False
                        
    Set objSQL = gobjCHCSIIConn.CHCSII_SQLOPS(Auto)
    
    If Note.mnDataID = 0 Then
        Note.mnDataID = objSQL.GetNextID(exMAINSEQ)
    End If
    
    'ADD LISTTOOLRECORD DATA
   strSQL = "SELECT " & m_strListToolRecordColumns & " FROM ListToolRecord WHERE DataId = " & CStr(Note.mnDataID)
            
   Call objSQL.Execute(strSQL)
   
   With objSQL
      If (.BOF And .EOF) Then
         oSQLGen.SetTransType = eSQLInsert
         oSQLGen.AddToSQL "DataId", Note.DataID, eSQLNumber
         oSQLGen.AddToSQL "UpdateFlag", "N", eSQLString
         oSQLGen.AddToSQL "CreatedBy", Note.msOwnerNCID, eSQLString
         oSQLGen.AddToSQL "CreatedOn", Now, eSQLDate
         oSQLGen.AddToSQL "SourceNCID", SectionEnumToNCID(eSO), eSQLNumber
         oSQLGen.AddToSQL "FacilityNCID", mobjEncounter.FacilityNCID, eSQLNumber
      Else
        oSQLGen.SetTransType = eSQLUpdate
        oSQLGen.AddToSQL "UpdateFlag", "Y", eSQLString
      End If

        oSQLGen.AddToSQL "UpdatedBy", Note.msOwnerNCID, eSQLNumber
        oSQLGen.AddToSQL "UpdatedOn", Now, eSQLDate
        oSQLGen.AddToSQL "EncounterNumber", mobjEncounter.EncounterID, eSQLNumber
        oSQLGen.AddToSQL "OriginalSize", Len(Note.msRTF), eSQLNumber
        
        .Execute oSQLGen.GenerateSQL("LISTTOOLRECORD", "DATAID = " & Note.DataID)
        
        oSQLGen.Clear
        'ADD RTF DATA
        mobjEncounter.CWShared.CHCSConnection.AppendChunk "LISTTOOLRECORD", "DATAID = " & Note.DataID, "RTF", Note.msRTF
        
        'ADD TEXT NOTE DATA'SCR-48052
        mobjEncounter.CWShared.CHCSConnection.AppendChunk "LISTTOOLRECORD", "DATAID = " & Note.DataID, "TEXT_NOTE_RTF", Note.TextNote.msRTF
   End With
           
   'ADD LISTTOOLRECORDS
    '-- Rebuild the details
    '   Mark all current as deleted
    '   add any new ones (ltr's without dataids)
    '   update any 'old' ones
    '   that leaves all the deleted ones marked as deleted
    
    objSQL.Execute "DELETE FROM ListToolRecords where DataID = " & Note.mnDataID
    
    If Not Note.mcolListRecords Is Nothing Then 'May have blank SO Note SCR#44787
        For Each objLTR In Note.mcolListRecords

            oSQLGen.SetTransType = eSQLInsert
            oSQLGen.AddToSQL "DataId", Note.mnDataID, eSQLNumber
            oSQLGen.AddToSQL "UpdateFlag", "N", eSQLString
            oSQLGen.AddToSQL "CreatedBy", Note.msOwnerNCID, eSQLNumber
            oSQLGen.AddToSQL "CreatedOn", Now, eSQLDate
            oSQLGen.AddToSQL "UpdatedBy", Note.msOwnerNCID, eSQLNumber
            oSQLGen.AddToSQL "UpdatedOn", Now, eSQLDate
            
            With objLTR
                oSQLGen.AddToSQL "SnoID", .SnoID, eSQLString
                oSQLGen.AddToSQL "Duration", .Duration, eSQLString
                oSQLGen.AddToSQL "Modifier", .Modifier, eSQLString
                oSQLGen.AddToSQL "Onset", .Onset, eSQLString
                If Len(.prefix) = 0 Then .prefix = " "
                oSQLGen.AddToSQL "prefix", .prefix, eSQLString
                oSQLGen.AddToSQL "qlink", .QualifierLink, eSQLString
                oSQLGen.AddToSQL "qlinkseq", .QualifierLinkSequence, eSQLNumber
                oSQLGen.AddToSQL "RangeNormalHigh", .RangeNormalHigh, eSQLNumber
                oSQLGen.AddToSQL "RangeNormalLow", .RangeNormalLow, eSQLNumber
                oSQLGen.AddToSQL "RangeScale", .RangeScale, eSQLNumber
                oSQLGen.AddToSQL "Referid", .Referid, eSQLNumber
                oSQLGen.AddToSQL "result", .result, eSQLString
                oSQLGen.AddToSQL "Status", .Status, eSQLString
                oSQLGen.AddToSQL "Unit", .Unit, eSQLString
                oSQLGen.AddToSQL "value", .value, eSQLString
                oSQLGen.AddToSQL "ChartFlag", .Flag, eSQLNumber
                oSQLGen.AddToSQL "Note", .Note, eSQLString
             End With
            objSQL.Execute oSQLGen.GenerateSQL("LISTTOOLRECORDS", "")
            oSQLGen.Clear
            
        Next 'objLTR
    End If
    
    Set oSQLGen = Nothing
    Set objSQL = Nothing
    Set objLTR = Nothing
    
    Exit Function
    
ErrHandler:
    Call mobjEncounter.CWShared.ShowVBError(Err.Number, Err.Description, "SO.PGUI_SaveNote", "EncounterOps", vbCritical)
End Function

Public Sub GEMS_Load(ByRef objParent As SectionParent, objRS As ADODB.Recordset)
' rs is loaded with enc_section for 'so' for this enc
' load the information from the enc_section first, then go after the listtoolrecord and listtoolrecords data
    Dim objLTR          As ListToolRecord
    Dim objNote         As SONote
    Dim oConcept        As New GEMS_ConceptCtrl
    Dim objISQLRecord   As ICHCSII_SQL.ISqlOps
    Dim objISQLNote     As ICHCSII_SQL.ISqlOps
    Dim strSQL          As String
    Dim oShared         As CWShared
    
    Set oShared = New CWShared
    
    Set mcolNotes = New Collection
    Set objISQLRecord = gobjCHCSIIConn.CHCSII_SQLOPS(Auto)  '''<SCR 36493
    Set objISQLNote = gobjCHCSIIConn.CHCSII_SQLOPS(Auto)  '''<SCR 36493
    
   
    If Not objRS.EOF Then
        If Not oShared.IsAppMode(modeCHCSI_GUI) Then
            strSQL = "SELECT " & m_strListToolRecordColumns & " FROM ListToolRecord " & _
                     "WHERE EncounterNumber = " & objRS("EncounterNumber").value & " " & _
                     "AND facilityncid = " & objRS("FacilityNCID") & " " & _
                     "AND SOURCENCID = " & SectionEnumToNCID(eSO)
        Else 'SCR-48052
            strSQL = "SELECT " & m_strListToolRecordColumns & ", TEXT_NOTE_RTF FROM ListToolRecord " & _
                     "WHERE EncounterNumber = " & objRS("EncounterNumber").value & " " & _
                     "AND facilityncid = " & objRS("FacilityNCID") & " " & _
                     "AND SOURCENCID = " & SectionEnumToNCID(eSO)
        End If
 
     
    objISQLNote.Execute strSQL
    
      Do While Not objISQLNote.EOF
    
        Set objNote = New SONote
        objNote.meType = ListNoteType
        objNote.mnDataID = objISQLNote("DataID")
        
        If Not IsNull(objISQLNote("UpdatedBy")) Then
            objNote.msOwnerNCID = objISQLNote("UpdatedBy")
            oConcept.UniqueId = objNote.msOwnerNCID
            objNote.msOwnerName = oConcept.PrefRep("2000").Representation
        ElseIf Not IsNull(objISQLNote("CreatedBy")) Then
            objNote.msOwnerNCID = objISQLNote("CreatedBy")
            oConcept.UniqueId = objNote.msOwnerNCID
            objNote.msOwnerName = oConcept.PrefRep("2000").Representation
        End If
        
        'If Not IsNull(objISQLNote("UpdatedOn")) Then
        'WRogers 9/21/2001 SCR 16498
        If Len(objISQLNote("UpdatedOn")) Then
            objNote.mdDTS = objISQLNote("UpdatedOn")
        End If
        'objNote.msTitle = "" & objRS("Title")
        
        If Not oShared.IsAppMode(modeCHCSI_GUI) Then 'SCR-48052
            'Decompress if compressed SCR-24683
            If CVar(objISQLNote("OriginalSize")) > 0 Then
                Dim objZlib As EncZLib
                Set objZlib = New EncZLib
                Dim aBytes() As Byte
                
                aBytes = objISQLNote("RTF")
                Call objZlib.DecompressData(aBytes, objISQLNote("OriginalSize"))
                objNote.msRTF = StrConv(aBytes, vbUnicode)
                'objNote.msRTF = DecompressRTF(objISQLNote("RTF"), objISQLNote("OriginalSize"))
            Else
                'objNote.msRTF = objISQLNote("RTF")
                 
            End If
        
        Else
            'PGUI Mode SCR#50699
            objNote.msRTF = ReadChunkCHCSConn("LISTTOOLRECORD", "ENCOUNTERNUMBER = " & objRS("ENCOUNTERNUMBER") & " AND facilityncid = " & objRS("FacilityNCID") & " AND DATAID = " & objISQLNote("DataID"), "RTF", False)
            objNote.TextNote.msRTF = ReadChunkCHCSConn("LISTTOOLRECORD", "ENCOUNTERNUMBER = " & objRS("ENCOUNTERNUMBER") & " AND facilityncid = " & objRS("FacilityNCID") & " AND DATAID = " & objISQLNote("DataID"), "TEXT_NOTE_RTF", False)
        End If
        
         strSQL = "SELECT " & m_strListToolRecordsColumns & " FROM ListToolRecords " & _
                  "WHERE dataid = " & objISQLNote("DataID")
                  
         objISQLRecord.Execute strSQL
         Set objNote.mcolListRecords = New Collection
         
         
         
         Do While Not objISQLRecord.EOF
             Set objLTR = New ListToolRecord
             Call objLTR.GEMS_ReadFromClinObs(objISQLRecord)
             objNote.mcolListRecords.Add objLTR
             objISQLRecord.MoveNext
         Loop
         
         mcolNotes.Add objNote, CStr(objNote.mnDataID)
         
         objISQLNote.MoveNext
      Loop
    End If
Proc_Exit:
    Set objLTR = Nothing
    Set objNote = Nothing
    Set oConcept = Nothing
    Set objISQLRecord = Nothing
    Set objISQLNote = Nothing
    Exit Sub
    
Proc_Error:
   Err.Raise Err.Number, Err.Source, Err.Description
   GoTo Proc_Exit
End Sub


Private Function GEMS_Delete(objNote As SONote) As Boolean
'-- delete the sonote record along with the matching listtoolrecord and listtoolrecords records
    On Error GoTo Proc_Exit
    
    Dim objGemsDas As ICHCSII_DAS
    
    
    Set objGemsDas = gobjCHCSIIConn.CHCSII_DAS(Auto) '''<SCR 36493
    objGemsDas.ExecuteSQL "DELETE FROM ListToolRecords WHERE dataid = " & objNote.mnDataID
    objGemsDas.ExecuteSQL "DELETE FROM ListToolRecord WHERE dataid = " & objNote.mnDataID

    GEMS_Delete = True
Proc_Exit:
   Set objGemsDas = Nothing
   
   If Err.Number <> 0 Then
      GEMS_Delete = False
      Err.Raise Err.Number, Err.Source, Err.Description
   End If
End Function


Public Property Get Text_NoteTool() As CHCSII_NoteTool.Note
    Set Text_NoteTool = mNoteTool
End Property

Public Sub NotifyComplete()
    Dim bCreatedObject As Boolean
    
    'S/O's Completed Encounter only adds Family Hx to Problems which isn't supported
    'in PGUI   SCR 51081
    If mobjEncounter.CWShared.IsAppMode(modeCHCSI_GUI) Then
        Exit Sub
    End If
    
    On Error Resume Next
    If mobjModule Is Nothing Then
        Set mobjModule = mobjEncounter.Comm.InitializeOLEServer(SectionEnumToNCID(eSO))
        If mobjModule Is Nothing Then
            Exit Sub
        End If
        bCreatedObject = True
    End If
    
    Call mobjModule.CompletedEncounter(mobjEncounter)
    
    If bCreatedObject Then
        Set mobjModule = Nothing
    End If
    Set mobjICoreAPI3 = mobjModule
End Sub

Friend Property Get Notes() As Collection
    Set Notes = mcolNotes
End Property
' Created to check is not has not been deleted before editing or saving note SCR-26756
Public Function CheckIfNoteExists(objNote As SONote) As Boolean

    Dim objData         As DataString
    Dim objISQLNote     As ICHCSII_SQL.ISqlOps
    Dim strSQL          As String
       
    CheckIfNoteExists = False
    
    Set objISQLNote = gobjCHCSIIConn.CHCSII_SQLOPS(Auto)  '''<SCR 36493
    
    strSQL = "SELECT " & m_strListToolRecordColumns & " FROM ListToolRecord " & _
             "WHERE DATAID = " & objNote.DataID
    
    objISQLNote.Execute strSQL
    
    If Not objISQLNote.EOF Then
        CheckIfNoteExists = True
    End If
    
    Set objISQLNote = Nothing
    
End Function

Private Property Get GreenSwitch() As Boolean
   GreenSwitch = mblnGreenSwitch
End Property
Private Property Let GreenSwitch(blnGreen As Boolean)
   mblnGreenSwitch = blnGreen
End Property
