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
Attribute VB_Name = "SectionDataCDR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements iSectionDataMapper

Private mConn As CHCSII_CONN.Conn
Private mCache As CacheMapper

Private mMode As CHCSII_CONN.eBackend
Private mUseCache As Boolean

Private mcolDataStrings As Collection
Public Function InitRetrieve(ByRef mcolDataStrings As Collection)
    Set mcolDataStrings = mcolDataStrings
End Function
Private Sub iSectionDataMapper_Init(Conn As CHCSII_CONN.Conn, Cache As DL_Support.CacheMapper, ByVal DataAccessMode As CHCSII_CONN.eBackend, Optional ByVal CacheConfigItems As Boolean = False)
  Set mConn = Conn
  Set mCache = Cache
  
  mMode = DataAccessMode
  mUseCache = CacheConfigItems
End Sub

Private Function iSectionDataMapper_RetrieveAP(Encounter As DTOs.Encounter, Params As DL_Support.iParameters) As Boolean
    Dim objData As DATASTRATLLib.DataString
    Dim objEvent As CLINTYPEATLLib.ClinicalEvent
    Dim objClinObs As CLINTYPEATLLib.ClinicalObservation

    Dim objCommentGroup As CLINTYPEATLLib.CommentGroup

    Dim objCommentItem As CLINTYPEATLLib.CommentItem
    Dim oMedcinRecord As MedcinRecord
    Dim sName As String
    Dim sValue As String
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim l As Long
    Dim oShared As New CWShared
    
    Dim AnP As AnP
    Dim AnPSect As iSection
    
    Set AnP = Encounter.Sections("AnP")
    
    Set AnPSect = AnP
    
    'Set mcolRelatedDiagnoses = New Collection
    'Set mcolSnoIDs = New Collection
    
    AnP.Status = Incomplete

    AnPSect.DataID = 0
    'Using the For...Each method with 3M objects causes memory leak
    For j = 1 To mcolDataStrings.Count
    'For Each objData In mcolDataStrings
        Set objData = mcolDataStrings(j)
        If objData.DataTypeNCID = NCID_ENCEVENT Then
            Set objEvent = objData.PatientData
            If objEvent.Header.ObsBatId.Value.Coded.ValueConcept.UniqueId = SectionEnumToNCID(eap) Then
                AnPSect.DataID = objData.DataID
                If Not objData.ModifyClinician Is Nothing Then
                    AnPSect.OwnerNCID = objData.ModifyClinician.ValueConcept.UniqueId
                    'msOwnerName = objData.ModifyClinician.ValueConcept.PrefRep("2000").Representation
                ElseIf Not objData.CreateClinician Is Nothing Then
                    AnPSect.OwnerNCID = objData.CreateClinician.ValueConcept.UniqueId
                    'msOwnerName = objData.CreateClinician.ValueConcept.PrefRep("2000").Representation
                End If
                AnPSect.Document.LastModifiedDate = objEvent.Header.ObsDateTime.StartTime.Value.DateTime
                AnP.Status = StatusNCIDToEnum(objEvent.Header.TestStatus.Value.CodedWOSform.ValueConcept)
'                For Each objCommentGroup In objEvent.Header.Comments
'                    For Each objCommentItem In objCommentGroup.Comments
'                        Call ParseCommentItem(objCommentItem, sName, sValue)
'                        Select Case sName
'                        End Select
'                    Next objCommentItem
'                Next objCommentGroup
                For i = 1 To objEvent.ClinObs.Count
                    Set objClinObs = objEvent.ClinObs.Item(i)
                    Select Case objClinObs.ObsId.Value.CodedWOSform.ValueConcept.UniqueId
                    Case NCID_ENCOUNTERRTF
                        AnPSect.Document.body = ExtractRTF(objClinObs)
                    Case NCID_MEDCIN_OBS
                        Set oMedcinRecord = New MedcinRecord
                        ReadMedcinRecordFromClinObs oMedcinRecord, objClinObs
                        'mcolSnoIDs.Add oMedcinRecord
                        Set oMedcinRecord = Nothing
                    End Select
                Next i
'                For i = 1 To objData.SemanticLinks.Count
'                    mcolRelatedDiagnoses.Add objData.SemanticLinks.Item(i).ObjectPointer.StringKey.StringId
'                Next i
                
                For i = 1 To objEvent.Header.Comments.Count
                    Set objCommentGroup = objEvent.Header.Comments(i)
                    For k = 1 To objCommentGroup.Comments.Count
                        Set objCommentItem = objCommentGroup.Comments(k)
                        Call ParseCommentItem(objCommentItem, sName, sValue)
                        Select Case sName
                        Case "UPDATED"
                            AnPSect = CBool(sValue)
                        End Select
                    Next 'k objCommentItem
                Next 'j objCommentGroup
                
                
                Exit For
            End If
        End If
    Next 'j (objData)
End Function

Private Function iSectionDataMapper_RetrieveAutocites(Encounter As DTOs.Encounter, Params As DL_Support.iParameters) As Boolean
    Dim objData As DATASTRATLLib.DataString
    Dim objEvent As CLINTYPEATLLib.ClinicalEvent
    Dim i As Long

    Dim Autocites As Autocites
    Dim Section As iSection
    
    On Error GoTo ErrHandler
    
    Set Autocites = Encounter.Sections("Autocites")
    
    Set Section = Autocites

    Section.DataID = 0
    'Using the For...Each method with 3M objects causes memory leak
    For i = 1 To mcolDataStrings.Count
    'For Each objData In colDataStrings
        Set objData = mcolDataStrings(i)
        If objData.DataTypeNCID = NCID_ENCEVENT Then
            Set objEvent = objData.PatientData
            If objEvent.Header.ObsBatId.Value.Coded.ValueConcept.UniqueId = SectionEnumToNCID(eAutoCite) Then
            
                With Section
                    .DataID = objData.DataID
                    If Not objData.ModifyClinician Is Nothing Then
                        .Document.OwnerNCID = objData.ModifyClinician.ValueConcept.UniqueId
                        .Document.OwnerName = objData.ModifyClinician.ValueConcept.PrefRep("2000").Representation
                    ElseIf Not objData.CreateClinician Is Nothing Then
                        .Document.OwnerNCID = objData.CreateClinician.ValueConcept.UniqueId
                        .Document.OwnerName = objData.CreateClinician.ValueConcept.PrefRep("2000").Representation
                    End If
                    .Document.LastModifiedDate = objEvent.Header.ObsDateTime.StartTime.Value.DateTime
                    If objEvent.ClinObs.Count > 0 Then
                        .Document.body = ExtractRTF(objEvent.ClinObs.Item(1))
                        If objEvent.ClinObs.Count > 1 Then
                            Autocites.LabRTF = ExtractRTF(objEvent.ClinObs.Item(2))
                            If objEvent.ClinObs.Count > 2 Then
                                Autocites.RadRTF = ExtractRTF(objEvent.ClinObs.Item(3))
                            End If
                            If objEvent.ClinObs.Count > 3 Then
                                Autocites.QuestionnaireRTF = ExtractRTF(objEvent.ClinObs.Item(4))
                            End If
                        End If
                    End If
                End With
                Exit For
            End If
        End If
    Next 'objData
    
    Exit Function
    
ErrHandler:

End Function

Private Function iSectionDataMapper_RetrieveDental(Encounter As DTOs.Encounter, Params As DL_Support.iParameters) As Boolean
    Dim Dental As Dental
    Dim Section As iSection
    
    Set Dental = Encounter.Sections("Dental")
    Set Section = Dental
    
    LoadSectionData Section, eDental
    
End Function

Private Function iSectionDataMapper_RetrieveDispo(Encounter As DTOs.Encounter, Params As DL_Support.iParameters) As Boolean
    Dim objData As DATASTRATLLib.DataString
    Dim objEvent As CLINTYPEATLLib.ClinicalEvent
    Dim objClinObs As CLINTYPEATLLib.ClinicalObservation
    Dim objClinObsMod As CLINTYPEATLLib.ClinicalObservation
    Dim objCommentGroup As CLINTYPEATLLib.CommentGroup
    Dim objCommentItem As CLINTYPEATLLib.CommentItem
    Dim objActionInfo As CLINTYPEATLLib.EventActionInfo
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim sName As String
    Dim sValue As String
    
    Dim Dispo As Disposition
    Dim DispoSection As iSection
    
    Set Dispo = New Disposition
    
    Set DispoSection = Dispo

    With Dispo
        'Using the For...Each method with 3M objects causes memory leak
        For i = 1 To mcolDataStrings.Count
        'For Each objData In colDataStrings
            Set objData = mcolDataStrings(i)
            If objData.DataTypeNCID = NCID_ENCEVENT Then
                Set objEvent = objData.PatientData
                If objEvent.Header.ObsBatId.Value.Coded.ValueConcept.UniqueId = SectionEnumToNCID(eDisp) Then
                    DispoSection.DataID = objData.DataID
                    If Not objData.ModifyClinician Is Nothing Then
                        DispoSection.Document.OwnerNCID = objData.ModifyClinician.ValueConcept.UniqueId
                        DispoSection.Document.OwnerName = objData.ModifyClinician.ValueConcept.PrefRep("2000").Representation
                    ElseIf Not objData.CreateClinician Is Nothing Then
                        DispoSection.Document.OwnerNCID = objData.CreateClinician.ValueConcept.UniqueId
                        DispoSection.Document.OwnerName = objData.CreateClinician.ValueConcept.PrefRep("2000").Representation
                    End If
                    DispoSection.Document.LastModifiedDate = objEvent.Header.ObsDateTime.StartTime.Value.DateTime
                    .Status = StatusNCIDToEnum(objEvent.Header.TestStatus.Value.CodedWOSform.ValueConcept)
                    
                    For j = 1 To objEvent.Header.Comments.Count
                    'For Each objCommentGroup In objEvent.Header.Comments
                        Set objCommentGroup = objEvent.Header.Comments(j)
                        For k = 1 To objCommentGroup.Comments.Count
                        'For Each objCommentItem In objCommentGroup.Comments
                            Set objCommentItem = objCommentGroup.Comments(k)
                            Call ParseCommentItem(objCommentItem, sName, sValue)
                            Select Case sName
                            Case "WSDisp"
                                .WSDisp = sValue
                            Case "EANDMCALC"
                                .EandMCalc = sValue
                            Case "EMCALCOVERRIDENOTE"
                                .EMCalcOverrideNote = sValue
                            Case "FOLLOWUPTIMEFRAME"
                                .FollowupTimeFrame = sValue
                            Case "FOLLOWUPCOMMENTS"
                                .FollowupComments = sValue
                            Case "DISCUSSEDCOMMENTS"
                                .DiscussedComments = sValue
                            Case "ITEMSDISCUSSED"
                                .ItemsDiscussed = sValue
                            Case "SUPPLDISPDATA"
                                .Text = sValue
                            Case "DENTALDISPSTATUS"
                                .DentalDispStatus = sValue
                            Case "DENTALDISPSTATUSDESC"
                                .DentalDispStatusDesc = sValue
                            Case "DENTALDISPCOMMENTS"
                                .DentalDispComments = sValue
                            Case "DENTALDISPOPTIONS"
                                .DentalDispOptions = sValue
                            Case "UPDATED"
                                .Updated = CBool(sValue)
                            End Select
                        Next 'k objCommentItem
                    Next 'j objCommentGroup
                    
                    Dim szNCID As String
                    
                    For j = 1 To objEvent.ClinObs.Count
                        Set objClinObs = objEvent.ClinObs.Item(j)
                        .NCID = objClinObs.ObsId.Value.CodedWOSform.ValueConcept.UniqueId
                        Select Case objClinObs.ObsId.Value.CodedWOSform.ValueConcept.UniqueId
                        Case NCID_METADATA
                            .DispMetaData = StrConv(objClinObs.ObsValue.Value.MultiMedia.Data, vbUnicode)
                            
                        Case NCID_ENCOUNTERRTF
                            DispoSection.Document.body = ExtractRTF(objClinObs)
                        Case NCID_EANDMCODE
                            .EandMNCID = objClinObs.ObsValue.Value.CodedWOSform.ValueConcept.UniqueId
                            'DEM - added to support E & M Coding Review 05/04/00
    '                        If Not objClinObs.ObsMods.IsEmpty Then
    '                            For k = 1 To objClinObs.ObsMods.Count
    '                            'For Each objClinObsMod In objClinObs.ObsMods
    '                                Set objClinObsMod = objClinObs.ObsMods(k)
    '                                Select Case objClinObsMod.ObsId.Value.CodedWOSform.ValueConcept
    '                                    Case NCID_REVIEWED_BY_CODER
    '                                        mbEandMReviewed = objClinObsMod.ObsValue.Value.CodedWOSform.ValueConcept
    '                                End Select
    '                                If (Not objClinObsMod.ActionsInfo.IsEmpty) Then
    '                                    For Each objActionInfo In objClinObsMod.ActionsInfo
    '                                        With objActionInfo
    '                                            If .ActionId.ValueConcept.UniqueId = NCID_VERIFIED Then
    '                                               msEandMReviewedBy = .Clinician.ValueConcept.UniqueId   '-- added 1.1.00 WAS:
    '                                               mdtEandMReviewDate = .EffectiveTime
    '                                            End If
    '                                        End With
    '                                    Next
    '                                    Set objActionInfo = Nothing
    '                                End If
    '                            Next 'k
    '                            Set objClinObsMod = Nothing
    '                        End If
                            'DEM - end add
                        Case NCID_DISPOSITION
                            .NCID = objClinObs.ObsValue.Value.CodedWOSform.ValueConcept.UniqueId
                            For k = 1 To objClinObs.ObsMods.Count
                            'For Each objClinObsMod In objClinObs.ObsMods
                                Set objClinObsMod = objClinObs.ObsMods(k)
                                Select Case objClinObsMod.ObsId.Value.CodedWOSform.ValueConcept
                                Case NCID_CONFINED_TO_QUARTERS
                                    .DispMetaData = objClinObsMod.ObsValue.Value.ULong
                                Case NCID_DISPOSITION
                                    If .AdminOptions = "" Then
                                        .AdminOptions = objClinObsMod.ObsValue.Value.CodedWOSform.ValueConcept.UniqueId
                                    Else
                                        .AdminOptions = .AdminOptions & "/" & objClinObsMod.ObsValue.Value.CodedWOSform.ValueConcept.UniqueId
                                    End If
                                End Select
                            Next 'objClinObsMod
                        End Select
                    Next 'j
                    Set objEvent = Nothing
                    Set objData = Nothing
                    Exit For ' there is only one disposition per encounter
                End If
            End If
        Next 'i (objData)
    End With
End Function

Private Function iSectionDataMapper_RetrieveEducation(Encounter As DTOs.Encounter, Params As DL_Support.iParameters) As Boolean
    Dim Education As Education
    Dim Section As iSection
    
    Set Education = Encounter.Sections("Education")
    Set Section = Education
    
    LoadSectionData Section, eEducation
End Function

Private Function iSectionDataMapper_RetrieveEncNotes(Encounter As DTOs.Encounter, Params As DL_Support.iParameters) As Boolean
    Dim objData As DATASTRATLLib.DataString
    Dim objEvent As CLINTYPEATLLib.ClinicalEvent
    Dim objNote As Note
    Dim i As Long
    Dim EncNotes As EncNotes
    Dim Section As iSection
    
    Set EncNotes = Encounter.Sections("EncNotes")
    
    'Using the For...Each method with 3M objects causes memory leak
    For i = 1 To mcolDataStrings.Count
    'For Each objData In colDataStrings
        Set objData = mcolDataStrings(i)
        If objData.DataTypeNCID = NCID_ENCEVENT Then
            Set objEvent = objData.PatientData
            If objEvent.Header.ObsBatId.Value.Coded.ValueConcept.UniqueId = SectionEnumToNCID(eNotes) Then
                Set objNote = New Note
                ReadNoteFromDataString objNote, objData
                ' the events are given to use in descending order, so reverse them
                If EncNotes.Count > 0 Then
                    EncNotes.Add objNote, CStr(Section.DataID), 1
                Else
                    EncNotes.Add objNote, CStr(Section.DataID)
                End If
                Set objNote = Nothing
            End If
        End If
    Next 'objData
End Function

Private Function iSectionDataMapper_RetrieveHistory(Encounter As DTOs.Encounter, Params As DL_Support.iParameters) As Boolean
    Dim objData As DATASTRATLLib.DataString
    Dim objEvent As CLINTYPEATLLib.ClinicalEvent
    Dim EncNote As Note
    Dim EncNotes As History
    Dim Section As iSection
    
    
    Set EncNotes = Encounter.Sections("History")
    
    For Each objData In mcolDataStrings
        Set Section = EncNote
        If objData.DataTypeNCID = NCID_ENCEVENT Then
            Set objEvent = objData.PatientData
            If objEvent.Header.ObsBatId.Value.Coded.ValueConcept.UniqueId = SectionEnumToNCID(eHistory) Then
                Set EncNote = New Note
                ReadNoteFromDataString EncNote, objData
                If EncNotes.Count > 0 Then
                    EncNotes.Add EncNote, CStr(Section.DataID), 1
                Else
                    EncNotes.Add EncNote, CStr(Section.DataID)
                End If
                Set EncNote = Nothing
            End If
        End If
    Next objData
End Function
Private Function WriteNoteToDataString(ByRef Note As Note, objData As DATASTRATLLib.DataString)
    Dim sComplete As String
    Dim sSensitive As String
    Dim sReadOnly As String
    
    Dim Section As iSection
    
    Set Section = Note
    
    'Call AddRTF(objData.PatientData, msRTF)
    Call AddRTF2(objData.PatientData, Section.Document.body, 1) 'SCR#42412
    If Note.IsComplete Then
        sComplete = "Y"
    Else
        sComplete = "N"
    End If
    
    '''''''''''''''''''''''''''''''''''''''''''''''
    ' SCR #10622 Provide means of setting a       '
    ' Note as sensitve.                           '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'sSensitive = CStr(Sensitive)                                              '
    'AddCommentValuePair objData.PatientData.Header, "SENSITIVE", sSensitive   '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    sReadOnly = CStr(Note.ReadOnly)    'SCR 22188                                             '
    AddCommentValuePair objData.PatientData.Header, "READONLY", sReadOnly
    
    Call AddCommentValuePair(objData.PatientData.Header, "COMPLETE", sComplete)
    Call AddCommentValuePair(objData.PatientData.Header, "CATEGORY", Section.Document.Category)
    Call AddCommentValuePair(objData.PatientData.Header, "TITLE", Section.Document.Title)
End Function
    
Private Function ReadNoteFromDataString(ByRef Note As Note, objData As DATASTRATLLib.DataString)
    Dim objEvent As ClinicalEvent
    Dim objCommentGroup As CLINTYPEATLLib.CommentGroup
    Dim objCommentItem As CLINTYPEATLLib.CommentItem
    Dim sName As String
    Dim sValue As String
    Dim i As Long
    Dim j As Long
    
    Dim Section As iSection
    
    Set Section = Note
    
    On Error GoTo ErrHandler
    
    With Section
    
        .DataID = objData.DataID
        If Not objData.ModifyClinician Is Nothing Then
            .Document.OwnerNCID = objData.ModifyClinician.ValueConcept.UniqueId
            .Document.OwnerName = objData.ModifyClinician.ValueConcept.PrefRep("2000").Representation
        ElseIf Not objData.CreateClinician Is Nothing Then
            .Document.OwnerNCID = objData.CreateClinician.ValueConcept.UniqueId
            .Document.OwnerNCID = objData.CreateClinician.ValueConcept.PrefRep("2000").Representation
        End If
        Set objEvent = objData.PatientData
        .Document.LastModifiedDate = objEvent.Header.ObsDateTime.StartTime.Value.DateTime
        If objEvent.ClinObs.Count > 0 Then
            If objEvent.ClinObs.Item(1).ObsMods.Count > 2 Then
                'RTF may be saved in sections SCR#42412
                'Checks to make sure RTF is stored in this ClinObs SCR#42412
                If StrComp(objEvent.ClinObs.Item(1).ObsMods(3).ObsValue.Value.InfoType, "basicString", vbTextCompare) = 0 Then
                    If StrComp(objEvent.ClinObs.Item(1).ObsMods(3).ObsValue.Value.BasicString, "RTF", vbTextCompare) = 0 Then
                        .Document.body = ExtractRTF2(objEvent, 1)
                    End If
                End If
            Else
                .Document.body = ExtractRTF(objEvent.ClinObs.Item(1))
            End If
        End If
    
    End With
    
    With Note
        'Using the For...Each method with 3M objects causes memory leak
        For i = 1 To objEvent.Header.Comments.Count
        'For Each objCommentGroup In objEvent.Header.Comments
            Set objCommentGroup = objEvent.Header.Comments(i)
            For j = 1 To objCommentGroup.Comments.Count
            'For Each objCommentItem In objCommentGroup.Comments
                Set objCommentItem = objCommentGroup.Comments(j)
                Call ParseCommentItem(objCommentItem, sName, sValue)
                Select Case sName
                Case "COMPLETE"
                    If sValue = "Y" Then
                        .IsComplete = True
                    End If
                Case "CATEGORY"
                    Section.Document.Category = sValue
                Case "TITLE"
                    Section.Document.Title = sValue
                '''''''''''''''''''''''''''''''''''''''''''''''
                ' SCR #10622 Provide means of setting a       '
                ' Note as sensitve.                           '
                '''''''''''''''''''''''''''''''''''''''''''''''
                Case "SENSITIVE"                              '
                    'Note. = CBool(sValue)                 '
                '''''''''''''''''''''''''''''''''''''''''''''''
                Case "READONLY"
                    .ReadOnly = CBool(sValue)
                End Select
            Next 'objCommentItem
        Next 'objCommentGroup
    End With
    Set objEvent = Nothing
    
    Exit Function

ErrHandler:

End Function

Private Function iSectionDataMapper_RetrieveRFV(Encounter As DTOs.Encounter, Params As DL_Support.iParameters) As Boolean
    Dim RFV As RFV
    Dim Section As iSection
    Dim objEvent As CLINTYPEATLLib.ClinicalEvent
    Dim objCommentGroup As CLINTYPEATLLib.CommentGroup
    Dim objCommentItem As CLINTYPEATLLib.CommentItem
    Dim sName As String
    Dim sValue As String
    Dim k As Long
    Dim j As Long
    
    Set RFV = Encounter.Sections("RFV")
    Set Section = RFV
    
    Set objEvent = LoadSectionData(Section, eRFV)

    If objEvent.ClinObs.Count > 0 Then
        For j = 1 To objEvent.Header.Comments.Count
        'For Each objCommentGroup In objEvent.Header.Comments
            Set objCommentGroup = objEvent.Header.Comments(j)
            For k = 1 To objCommentGroup.Comments.Count
            'For Each objCommentItem In objCommentGroup.Comments
                Set objCommentItem = objCommentGroup.Comments(k)
                Call ParseCommentItem(objCommentItem, sName, sValue)
                Select Case sName
                Case "WSRFV"
                    RFV.WSRFV = sValue
                Case "UPDATED"
                    RFV.Updated = CBool(sValue)
                End Select
            Next 'objCommentItem
        Next 'objCommentGroup
    End If
                
End Function


Private Function iSectionDataMapper_RetrieveSO(Encounter As DTOs.Encounter, Params As DL_Support.iParameters) As Boolean
    Dim objData As DATASTRATLLib.DataString
    Dim objClinicalText As CLINTEXTATLLib.ClinicalText
    Dim objEvent As CLINTYPEATLLib.ClinicalEvent
    Dim objCommentGroup As CLINTYPEATLLib.CommentGroup
    Dim objCommentItem As CLINTYPEATLLib.CommentItem
    Dim objClinObs As CLINTYPEATLLib.ClinicalObservation
    Dim oRecord As MedcinRecord
    Dim objNote As SONote
    Dim objNoteSect As iSection
    Dim SO As SO
    Dim sName As String
    Dim sValue As String
    Dim i As Long
    Dim j As Long
    Dim k As Long
    
    'Set mcolNotes = New Collection
    Set SO = Encounter.Sections("SO")
    
    'Using the For...Each method with 3M objects causes memory leak
    For i = 1 To mcolDataStrings.Count
    'For Each objData In colDataStrings
        Set objData = mcolDataStrings(i)
        Select Case objData.DataTypeNCID
        Case NCID_DAS_CLINICAL_TEXT
            Set objNote = New SONote
            Set objNoteSect = objNote
            'objNote.Type = FormNoteType
            
            With objNoteSect
                .DataID = objData.DataID
                If Not objData.ModifyClinician Is Nothing Then
                    .Document.OwnerNCID = objData.ModifyClinician.ValueConcept.UniqueId
                    .Document.OwnerName = objData.ModifyClinician.ValueConcept.PrefRep("2000").Representation
                ElseIf Not objData.CreateClinician Is Nothing Then
                    .Document.OwnerNCID = objData.CreateClinician.ValueConcept.UniqueId
                    .Document.OwnerName = objData.CreateClinician.ValueConcept.PrefRep("2000").Representation
                End If
                'objNote.msProfileNCID = objData.DataEntryModel.ValueConcept.UniqueId
                Set objClinicalText = objData.PatientData
                .Document.body = ConvertHTMLtoRTF(objClinicalText.ReportText.TextObject.Text.Text)
    '            objNote.msOwnerNCID = objClinicalText.ContextInfo.DictatingClinician.ValueConcept.UniqueId
    '            objNote.msOwnerName = objClinicalText.ContextInfo.DictatingClinician.ValueConcept.PrefRep("2000").Representation
                .Document.LastModifiedDate = objClinicalText.EventDate
                'Ignore entries that have a report type of 69587 as they are generic text notes
                If objData.PatientData.ReportType.ValueConcept.UniqueId = 69587 Then Set objNote = Nothing
            End With
        Case NCID_ENCEVENT
            Set objEvent = objData.PatientData
            If objEvent.Header.ObsBatId.Value.Coded.ValueConcept.UniqueId = SectionEnumToNCID(eSO) Then
                Set objNote = New SONote
                Set objNoteSect = objNote
                'objNote.meType = ListNoteType
                Set objNote.MedcinRecords = New Collection
                
                With objNoteSect
                    .DataID = objData.DataID
                    If Not objData.ModifyClinician Is Nothing Then
                        .Document.OwnerNCID = objData.ModifyClinician.ValueConcept.UniqueId
                        .Document.OwnerName = objData.ModifyClinician.ValueConcept.PrefRep("2000").Representation
                    ElseIf Not objData.CreateClinician Is Nothing Then
                        .Document.OwnerNCID = objData.CreateClinician.ValueConcept.UniqueId
                        .Document.OwnerName = objData.CreateClinician.ValueConcept.PrefRep("2000").Representation
                    End If
                    .Document.LastModifiedDate = objEvent.Header.ObsDateTime.StartTime.Value.DateTime
                End With
                
                With objNote
                    For j = 1 To objEvent.Header.Comments.Count
                    'For Each objCommentGroup In objEvent.Header.Comments
                        Set objCommentGroup = objEvent.Header.Comments(j)
                        For k = 1 To objCommentGroup.Comments.Count
                        'For Each objCommentItem In objCommentGroup.Comments
                            Set objCommentItem = objCommentGroup.Comments(k)
                            Call ParseCommentItem(objCommentItem, sName, sValue)
                            Select Case sName
                            Case "TITLE"
                                '.msTitle = sValue
                            Case "READONLY"
                                If sValue = "T" Then
                                    .ReadOnly = True
                                End If
                            Case "NEWCOUPLERDATA"
                                 If sValue = "T" Then
                                    .NewCouplerData = True
                                End If
                            Case "COUPLERDATA"
                                 If sValue = "T" Then
                                    .CouplerData = True
                                End If
                            End Select
                        Next 'objCommentItem
                    Next 'objCommentGroup
                    For j = 1 To objEvent.ClinObs.Count
                        Set objClinObs = objEvent.ClinObs.Item(j)
                        Select Case objClinObs.ObsId.Value.CodedWOSform.ValueConcept.UniqueId
                        Case NCID_ENCOUNTERRTF
                            objNoteSect.Document.body = ExtractRTF(objClinObs)
                        Case NCID_MEDCIN_OBS
                            Set oRecord = New MedcinRecord
                            ReadMedcinRecordFromClinObs oRecord, objClinObs
                            .MedcinRecords.Add oRecord
                            Set oRecord = Nothing
                        End Select
                    Next j
                End With
            End If
        End Select
        If Not objNote Is Nothing Then
            ' the events are given to use in descending order, so reverse them
            If SO.Count > 0 Then
                SO.Add objNote, CStr(objNoteSect.DataID), 1
            Else
                SO.Add objNote, CStr(objNoteSect.DataID)
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
    Set oRecord = Nothing
    Set objNote = Nothing

End Function

Private Function iSectionDataMapper_RetrieveVitals(Encounter As DTOs.Encounter, Params As DL_Support.iParameters) As Boolean
    Dim objData As DATASTRATLLib.DataString
    Dim objEvent As CLINTYPEATLLib.ClinicalEvent
    Dim objClinObs As CLINTYPEATLLib.ClinicalObservation
    Dim oMedcinRecord As MedcinRecord
    Dim i As Long
    Dim j As Long
    
    Dim Vitals As Vitals
    Dim VitalsSect As iSection
'    Set mcolSnoIDs = Nothing
'    Set mcolSnoIDs = New Collection
'
'    msRTF = ""
'    mbDirty = False
    VitalsSect.DataID = 0
    
    'Using the For...Each method with 3M objects causes memory leak
    For i = 1 To mcolDataStrings.Count
    'For Each objData In colDataStrings
        Set objData = mcolDataStrings(i)
        If objData.DataTypeNCID = NCID_ENCEVENT Then
            Set objEvent = objData.PatientData
            If objEvent.Header.ObsBatId.Value.Coded.ValueConcept.UniqueId = SectionEnumToNCID(eVitals) Then
            
                VitalsSect.DataID = objData.DataID
                
                ' Medcin Obs are now stored in Vitals SCR-11863
                For j = 1 To objEvent.ClinObs.Count
                    Set objClinObs = objEvent.ClinObs.Item(j)
                    Select Case objClinObs.ObsId.Value.CodedWOSform.ValueConcept.UniqueId
                    Case NCID_ENCOUNTERRTF
                        VitalsSect.Document.body = ExtractRTF(objClinObs)
                    Case NCID_MEDCIN_OBS
                        Set oMedcinRecord = New MedcinRecord
                        Call ReadMedcinRecordFromClinObs(oMedcinRecord, objClinObs)
                        Vitals.MedcinRecords.Add oMedcinRecord
                        Set oMedcinRecord = Nothing
                    End Select
                Next j
                
'                If objEvent.ClinObs.Count > 0 Then
'                    msRTF = ExtractRTF(objEvent.ClinObs.Item(1))
'                End If
'                Exit For
            End If
        End If
    Next 'objData
End Function

Private Function iSectionDataMapper_SaveAP(Encounter As DTOs.Encounter) As Boolean

' Assume section already locked for update before getting here
    Dim objData As DataString
    Dim objEvent As ClinicalEvent
    Dim objHeader As ClinObsHeader
    Dim objLink As SemanticLink
    Dim objLTR As MedcinRecord
    Dim sDiagnosis As Variant
    Dim AP As AnP
    Dim APSect As iSection
    
    On Error GoTo ErrHandler
    
    Set AP = Encounter.Sections("AnP")
    Set APSect = AP
    With AP
        APSect.Document.OwnerNCID = gobjShared.CurrentUserNCID
        APSect.Document.OwnerName = gobjShared.CurrentUserName
        
        APSect.Document.LastModifiedDate = Now
        Set objData = InitializeDataString(AP, APSect.DataID, APSect.Document.LastModifiedDate, APSect.Document.LastModifiedDate, Encounter)
        Set objEvent = objData.PatientData

        Call AddRTF(objEvent, APSect.Document.body)
        Call AddCommentValuePair(objEvent.Header, "UPDATED", CStr(.Updated))
        
'        For Each objLTR In mcolSnoIDs
'            Call objLTR.WriteToClinObs(objEvent)
'        Next objLTR
        Set objHeader = objEvent.Header
        objHeader.TestStatus.InfoType = "att"
        objHeader.TestStatus.Value.InfoType = "codedWOSform"
        objHeader.TestStatus.Value.CodedWOSform.ValueConcept.UniqueId = StatusEnumToNCID(.Status)
'        For Each sDiagnosis In mcolRelatedDiagnoses
'            Set objLink = objData.SemanticLinks.Add
'            objLink.Relationship.InfoType = "att"
'            objLink.Relationship.Value.InfoType = "codedWOSform"
'            objLink.Relationship.Value.CodedWOSform.ValueConcept.UniqueId = NCID_RELATED_PROBLEMS_DIAGNOSES
'            objLink.ObjectPointer.InfoType = "intObjPtr"
'            objLink.ObjectPointer.StringKey.StringId = sDiagnosis
'            objLink.ObjectPointer.StringKey.EditVersion = 1
'            Set objLink = Nothing
'        Next sDiagnosis
        
        If objData.Save Then
            APSect.DataID = objData.DataID
        Else
            'RAISE ERROR
            'ZZZ
            'mobjEncounter.DisplayDetailedError "AP.Save", ERROR_MESSAGE_SAVE_FAILURE
            'Err.RAISE 100000, "Encounters DL",  ERROR_MESSAGE_SAVE_FAILURE & vbCrLf & "APSave(" & mobjEncounter.PatientID & " - " & mobjEncounter.FacilityNCID & " - " & mobjEncounter.EncounterID & ")", vbCritical 'SCR-39485, 38953
        End If
    End With
    
    Exit Function
    
ErrHandler:

End Function

Private Function iSectionDataMapper_SaveAutocites(Encounter As DTOs.Encounter) As Boolean

End Function

Private Function iSectionDataMapper_SaveDental(Encounter As DTOs.Encounter) As Boolean
' Assume section already locked for update before getting here
    Dim objData As DataString
    Dim Dental As Dental
    Dim DentalSect As iSection
    
    On Error GoTo ErrHandler
    
    Set Dental = New Dental
    
    Set DentalSect = Dental
    
    With DentalSect.Document
        .OwnerNCID = gobjShared.CurrentUserNCID
        .OwnerName = gobjShared.CurrentUserNCID
        .LastModifiedDate = Now
    
        Set objData = InitializeDataString(Dental, DentalSect.DataID, .LastModifiedDate, .LastModifiedDate, Encounter)
        Call AddRTF(objData.PatientData, .body)
        If objData.Save Then
          DentalSect.DataID = objData.DataID
'          Save = True
'          mbDirty = False
        Else
            'RAISE ERROR
            'mobjEncounter.DisplayDetailedError "Dental.Save", ERROR_MESSAGE_SAVE_FAILURE
            'Err.RAISE 100000, "Encounters DL",  ERROR_MESSAGE_SAVE_FAILURE & vbCrLf & "DentalSave(" & mobjEncounter.PatientID & " - " & mobjEncounter.FacilityNCID & " - " & mobjEncounter.EncounterID & ")", vbCritical 'SCR-39485, 38953
        End If
    End With
    GoTo Cleanup


ErrHandler:
    'RAISE ERROR
   'Call mobjEncounter.CWShared.ShowVBError(Err.Number, Err.Description, "Enctr.Save ", "Enc", vbCritical)
Cleanup:
   Set objData = Nothing
'   If Save Then
'      mobjEncounter.Comm.SetStatusText "Encounter " & mobjEncounter.EncounterID & " " & SectionEnumToName(mobjParent.meSection) & " section was saved."
'   Else
'      mobjEncounter.Comm.SetStatusText ""
'   End If
   
End Function

Private Function iSectionDataMapper_SaveDispo(Encounter As DTOs.Encounter) As Boolean
' Assume section already locked for update before getting here
    Dim objData As DataString
    Dim objEvent As ClinicalEvent
    Dim objClinObs As ClinicalObservation
    Dim objObsMod As ClinicalObservation
    Dim objActionInfo As EventActionInfo
    Dim DispAdmin() As String
    Dim i As Integer
    
    Dim Dispo As Disposition
    Dim DispoSect As iSection
    
    On Error GoTo ErrHandler
    
    Set Dispo = Encounter.Sections("Dispostion")
    
    Set DispoSect = Dispo
    
    With DispoSect.Document
        .OwnerNCID = gobjShared.CurrentUserNCID
        .OwnerName = gobjShared.CurrentUserName
        .LastModifiedDate = Now
    
        Set objData = InitializeDataString(eDisp, DispoSect.DataID, .LastModifiedDate, .LastModifiedDate, Encounter)
        Set objEvent = objData.PatientData
        
        Call AddRTF(objEvent, .body)
    End With
    
    With Dispo
        Call AddClinObs(objEvent, NCID_METADATA, .DispMetaData, , "multiMedia")
        Call AddCommentValuePair(objEvent.Header, "UPDATED", CStr(.Updated))
        
        Call AddCommentValuePair(objEvent.Header, "WSDisp", .WSDisp)
        Call AddCommentValuePair(objEvent.Header, "EANDMCALC", .EandMCalc)
        Call AddCommentValuePair(objEvent.Header, "EMCALCOVERRIDENOTE", .EMCalcOverrideNote)
        Call AddCommentValuePair(objEvent.Header, "FOLLOWUPTIMEFRAME", .FollowupTimeFrame)
        Call AddCommentValuePair(objEvent.Header, "FOLLOWUPCOMMENTS", .FollowupComments)
        Call AddCommentValuePair(objEvent.Header, "DISCUSSEDCOMMENTS", .DiscussedComments)
        Call AddCommentValuePair(objEvent.Header, "ITEMSDISCUSSED", .ItemsDiscussed)
        Call AddCommentValuePair(objEvent.Header, "SUPPLDISPDATA", .Text)
        Call AddCommentValuePair(objEvent.Header, "DENTALDISPSTATUS", .DentalDispStatus)
        Call AddCommentValuePair(objEvent.Header, "DENTALDISPSTATUSDESC", .DentalDispStatusDesc)
        Call AddCommentValuePair(objEvent.Header, "DENTALDISPCOMMENTS", .DentalDispComments)
        Call AddCommentValuePair(objEvent.Header, "DENTALDISPOPTIONS", .DentalDispOptions)
        objEvent.Header.TestStatus.InfoType = "att"
        objEvent.Header.TestStatus.Value.InfoType = "codedWOSform"
        objEvent.Header.TestStatus.Value.CodedWOSform.ValueConcept.UniqueId = StatusEnumToNCID(.Status)
        If Val(.EandMNCID) > 0 Then
            Set objClinObs = AddClinObs(objEvent, NCID_EANDMCODE, .EandMNCID)
        End If
'        'DEM - added this for E & M Coding Review 05/04/00
'        If Not objClinObs Is Nothing And mbEandMReviewed And mdtEandMReviewDate <> NULL_DATE And Val(.EandMReviewedBy) > 0 Then
'            Set objObsMod = AddClinObsMod(objClinObs, NCID_REVIEWED_BY_CODER, NCID_YES)
'            With objObsMod
'                ' ActionsInfo
'                '------------
'                Set objActionInfo = .ActionsInfo.Add
'                With objActionInfo
'                    .ActionId.ValueConcept.UniqueId = NCID_VERIFIED
'                    .Clinician.ValueConcept.UniqueId = .EandMReviewedBy
'                    .EffectiveTime = mdtEandMReviewDate
'                End With
'                Set objActionInfo = Nothing
'            End With
'        End If
        Set objObsMod = Nothing
        Set objClinObs = Nothing
        'DEM - end add
        
        
        If Val(.NCID) > 0 Then
            
            Set objClinObs = AddClinObs(objEvent, NCID_DISPOSITION, .NCID)
            'MAG 05/14/01 SCR#12449 Broke the next line up since val(msDispositionText) could cause an error
            'If msDispositionNCID = NCID_CONFINED_TO_QUARTERS And Val(msDispositionText) > 0 Then
            If .NCID = NCID_CONFINED_TO_QUARTERS Then
                If Val(.Text) > 0 Then
                    Set objObsMod = AddClinObsMod(objClinObs, NCID_CONFINED_TO_QUARTERS, .Text, , "uLong")
                    objObsMod.ObsValue.Units.ValueConcept.UniqueId = NCID_UNITS_HOURS
                End If
            End If
            
            'Save The Admin Options Data
            DispAdmin = Split(.AdminOptions, "/")
            For i = 0 To UBound(DispAdmin)
               Call AddClinObsMod(objClinObs, NCID_DISPOSITION, DispAdmin(i))
            Next i
        Else
            'For Telcons
            'Save Telcons with Disp NCID = 1 is causing ADM writeback errors SCR-39950
            'Set objClinObs = AddClinObs(objEvent, NCID_DISPOSITION, 1)
            'Cant Save The Admin Options Data since no disposition exists SCR#49377
    '        DispAdmin = Split(msDispAdminOption, "/")
    '        For i = 0 To UBound(DispAdmin)
    '           Call AddClinObsMod(objClinObs, NCID_DISPOSITION, DispAdmin(i))
    '        Next i
        
        End If
    End With
    
    If objData.Save Then
      DispoSect.DataID = objData.DataID
'
'      With mobjEncounter
'         Save = True       'SCR 50470 'Indicate success so far in save operation
'         If Len(.CWShared.CmdLineSwitches("HIPAA837")) > 0 And Not .CWShared.IsAppMode(modeITT) Then
'            'Process the save of HIPAA data
'            If .AccidentData.IsDirty Then Save = Save And .AccidentData.CDRSave(.Patient.UnitNumber, .FacilityNCID, .EncounterID)
'            Save = Save And mobjEncounter.PregnancyData.CDRPlusSave(.FacilityNCID, .EncounterID, .Status)
'            .Comm.Message cwiRESET_VISUAL_INDICATORS, "", NCID_CW_APPLICATION, NCID_ENCOUNTER_APPLICATION
'         End If
'      End With
      
'      mbDirty = False
    Else
        'RAISE ERROR
        'mobjEncounter.DisplayDetailedError "Disposition.Save", ERROR_MESSAGE_SAVE_FAILURE
        'Err.RAISE 100000, "Encounters DL",  ERROR_MESSAGE_SAVE_FAILURE & vbCrLf & "DispSave(" & mobjEncounter.PatientID & " - " & mobjEncounter.FacilityNCID & " - " & mobjEncounter.EncounterID & ")", vbCritical 'SCR-39485, 38953
    End If
    GoTo Cleanup
ErrHandler:
    'RAISE ERROR
    'Call mobjEncounter.CWShared.ShowVBError(Err.Number, Err.Description, "Enctr.Save ", "Enc", vbCritical)
Cleanup:
    'Resume Next
    Set objData = Nothing
    Set objEvent = Nothing
    Set objClinObs = Nothing
    Set objObsMod = Nothing
'    If Save Then
'        mobjEncounter.Comm.SetStatusText "Encounter " & mobjEncounter.EncounterID & " " & SectionEnumToName(mobjParent.meSection) & " section was saved."
'    Else
'        mobjEncounter.Comm.SetStatusText ""
'    End If
End Function

Private Function iSectionDataMapper_SaveEducation(Encounter As DTOs.Encounter) As Boolean
' Assume section already locked for update before getting here
    Dim objData As DataString
    On Error GoTo ErrHandler
    
    Dim Education As Education
    Dim Section As iSection
    
    Set Education = Encounter.Sections("Education")
    
    Set Section = Education
    
    Set objData = InitializeDataString(eEducation, Section.Id, Now, Now, Encounter)
    Call AddRTF(objData.PatientData, Section.Document.body)
    If objData.Save Then
       Section.DataID = objData.DataID
'       Save = True
'       mbDirty = False
    End If
    GoTo Cleanup
    
ErrHandler:
    'RAISE ERROR
    'Call mobjEncounter.CWShared.ShowVBError(Err.Number, Err.Description, "Education.Save ", "Enc", vbCritical)
Cleanup:
    Set objData = Nothing
'    If Save Then
'        mobjEncounter.Comm.SetStatusText "Encounter " & mobjEncounter.EncounterID & " " & SectionEnumToName(mobjParent.meSection) & " section was saved."
'    Else
'        mobjEncounter.Comm.SetStatusText ""
'    End If
End Function

Private Function iSectionDataMapper_SaveEncNotes(Encounter As DTOs.Encounter) As Boolean
' Assume section NOT already locked for update before getting here
    Dim objData As DataString
    Dim bWasNew As Boolean
    Dim bBrokenLock As Boolean
    Dim WarningMessage  As String
    Dim sNoteType As String
    Dim EncNotes As EncNotes
    Dim Section As iSection
    Dim oNote As Note
    
    On Error GoTo ErrHandler
    
    Set EncNotes = Encounter.Sections("EncNotes")

    For Each oNote In EncNotes
    
        Set Section = oNote
        
        If Section.DataID = 0 Then
            bWasNew = True
        End If
        
        If Not StartSectionUpdate(eNotes, oNote.LockID) Then 'SCR# 47028
            'Refresh to see if encounter was signed  SCR=35244
            'mobjEncounter.Refresh
            
            If mMode = ITT Then
                sNoteType = "S/O Note"
            Else
                sNoteType = "Text Note"
            End If
            
            WarningMessage = "Your lock for this note has been broken and the " & sNoteType & "  will be added as free text to the end of the encounter as an addendum."
            
            If Not (Encounter.WhoSigned Is Nothing) Then
              WarningMessage = WarningMessage & "  This encounter has been completed and signed by " & _
                                Encounter.WhoSigned.FullName & " on " & _
                                Format$(Encounter.WhenSignedDate, gsDefaultDateTime) & ".  "
            End If
            
            Err.Raise 10000, "Encounters DL", WarningMessage
        
            'Add SO note in Add Note section SCR#44686
'            mobjEncounter.AddTextNote "Additional " & sNoteType & " Information:", Section.Document.body
'
'            If Not (mobjEncounter.WhoSigned Is Nothing) Then
'              mobjEncounter.ReplaceEncRTF
'            End If
        
            bBrokenLock = True
            GoTo Cleanup
        End If

        'CDR Mode
        Set objData = InitializeDataString(eNotes, Section.DataID, Section.Document.LastModifiedDate, Section.Document.LastModifiedDate, Encounter)
        WriteNoteToDataString oNote, objData
        If objData.Save Then
            Section.DataID = objData.DataID
            iSectionDataMapper_SaveEncNotes = True
        Else
            'mobjEncounter.DisplayDetailedError "TextNote.Save", ERROR_MESSAGE_SAVE_FAILURE
            'MsgBxARMd ERROR_MESSAGE_SAVE_FAILURE & vbCrLf & "TextNoteSave(" & mobjEncounter.PatientID & " - " & mobjEncounter.FacilityNCID & " - " & mobjEncounter.EncounterID & ")", vbCritical 'SCR-39485, 38953
        End If

        If bWasNew And iSectionDataMapper_SaveEncNotes Then
            If EncNotes Is Nothing Then
                Set EncNotes = New Collection
            End If

            EncNotes.Add oNote, CStr(Section.DataID)
        End If
    
        If bBrokenLock Then
            iSectionDataMapper_SaveEncNotes = True
        End If
        Call EndSectionUpdate(eNotes, Section.DataID)
        Call EndSectionUpdate(eNotes, 0)
    Next
'
'    If Save Then
'        mobjEncounter.Comm.SetStatusText "Encounter " & mobjEncounter.EncounterID & " Note was saved."
'    Else
'        mobjEncounter.Comm.SetStatusText ""
'    End If

    GoTo Cleanup
ErrHandler:
    'Call mobjEncounter.CWShared.ShowVBError(Err.Number, Err.Description, "TextNote.Save ", "EncounterOps", vbCritical)
Cleanup:

    Set objData = Nothing
End Function

Private Function iSectionDataMapper_SaveHistory(Encounter As DTOs.Encounter) As Boolean
' Assume section NOT already locked for update before getting here
    Dim objData As DataString
    Dim bWasNew As Boolean
    Dim History As History
    Dim Section As iSection
    Dim Note As Note
    Dim State As iState
    
    On Error GoTo ErrHandler

    Set History = Encounter.Sections("History")
    
    Set Section = History
    
    With Section
        For Each Note In History
        
            Set State = Note
            
            If State.State = osNew Then
                Set objData = InitializeDataString(eHistory, .DataID, .Document.LastModifiedDate, .Document.LastModifiedDate, Encounter)
                WriteNoteToDataString Note, objData
                If objData.Save Then
                    .DataID = objData.DataID
                    'Save = True
                Else
                    'RAISE ERROR
                    'mobjEncounter.DisplayDetailedError "History.Save", ERROR_MESSAGE_SAVE_FAILURE
                    'MsgBxARMd ERROR_MESSAGE_SAVE_FAILURE & vbCrLf & "HistorySave(" & mobjEncounter.PatientID & " - " & mobjEncounter.FacilityNCID & " - " & mobjEncounter.EncounterID & ")", vbCritical 'SCR-39485, 38953
                End If
            End If
            
            History.Add Note, CStr(Section.DataID)
        Next
        
    End With
    GoTo Cleanup
ErrHandler:
    'RAISE ERROR
    'Call mobjEncounter.CWShared.ShowVBError(Err.Number, Err.Description, "TextNote.Save ", "EncounterOps", vbCritical)
Cleanup:
    On Error Resume Next
'    If bWasNew Then
'        mcolItems.Add objItem, CStr(objItem.mnDataID)
'    Else
'        Call mobjEncounter.LockingObject.EndSectionUpdate(eHistory, objItem.mnDataID)
'    End If
    Set objData = Nothing
'    If Save Then
'        mobjEncounter.Comm.SetStatusText "Encounter " & mobjEncounter.EncounterID & " change history was saved."
'    Else
'        mobjEncounter.Comm.SetStatusText ""
'    End If
End Function

Private Function iSectionDataMapper_SaveRFV(Encounter As DTOs.Encounter) As Boolean
' Assume section already locked for update before getting here
    Dim objData As DataString
    Dim RFV As RFV
    Dim Section As iSection
    
    On Error GoTo ErrHandler
    
    Set RFV = Encounter.Sections("RFV")
    
    Set Section = RFV
    
    
'    msOwnerNCID = mobjEncounter.Logon.UserNCID
'    msOwnerName = mobjEncounter.Logon.UserName
'    mdDTS = Now
    With Section
        Set objData = InitializeDataString(eRFV, .DataID, .Document.LastModifiedDate, .Document.LastModifiedDate, Encounter)
        Call AddRTF(objData.PatientData, .Document.body)
        Call AddCommentValuePair(objData.PatientData.Header, "WSRFV", RFV.WSRFV)
        Call AddCommentValuePair(objData.PatientData.Header, "UPDATED", CStr(RFV.Updated))
        If objData.Save Then
            Section.DataID = objData.DataID
            iSectionDataMapper_SaveRFV = True
            
            'Save HIPAA Data
    '        With mobjEncounter
    '            If Len(.CWShared.CmdLineSwitches("HIPAA837")) > 0 And Not .CWShared.IsAppMode(modeITT) Then
    '               'Process the save of HIPAA data
    '               If .AccidentData.IsDirty Then Save = Save And .AccidentData.CDRSave(.Patient.UnitNumber, .FacilityNCID, .EncounterID)
    '               If .PregnancyData.IsDirty Then
    '                  Save = Save And mobjEncounter.PregnancyData.CDRPlusSave(.FacilityNCID, .EncounterID, .Status)
    '                  .Comm.Message cwiRESET_VISUAL_INDICATORS, "", NCID_CW_APPLICATION, NCID_ENCOUNTER_APPLICATION
    '               End If
    '            End If
    '
    '         End With
                    
            'mbDirty = False
        Else
            'mobjEncounter.DisplayDetailedError "RFV.Save", ERROR_MESSAGE_SAVE_FAILURE
            'MsgBxARMd ERROR_MESSAGE_SAVE_FAILURE & vbCrLf & "RFVSave(" & mobjEncounter.PatientID & " - " & mobjEncounter.FacilityNCID & " - " & mobjEncounter.EncounterID & ")", vbCritical 'SCR-39485, 38953
        End If
    End With
    GoTo Cleanup
ErrHandler:
    'RAISE ERROR
    'Call mobjEncounter.CWShared.ShowVBError(Err.Number, Err.Description, "Enctr.Save ", "Enc", vbCritical)
Cleanup:
    Set objData = Nothing
'    If Save Then
'        mobjEncounter.Comm.SetStatusText "Encounter " & mobjEncounter.EncounterID & " " & SectionEnumToName(mobjParent.meSection) & " section was saved."
'    Else
'        mobjEncounter.Comm.SetStatusText ""
'    End If
End Function

Private Function iSectionDataMapper_SaveSO(Encounter As DTOs.Encounter) As Boolean

    Dim objData As DataString
    Dim objNote As SONote
    Dim SO As SO
    Dim Section As iSection
    
    Dim State As iState
    
    Set SO = Encounter.Sections("SO")
    'Loop through SO Notes and see which ones need to be saved
    For Each objNote In SO
    
        Set State = objNote
        
        If State.State = osModified Or State.State = osNew Then
            'Save note
        
            Set Section = objNote
            
            If State.State = osModified Then
                'Make sure note still exists
                'Check to make sure note exists in CDR (has not been deleted) SCR-26756
                If Section.DataID > 0 Then
                    Set objData = InitializeDataString(eSO, Section.DataID, Section.Document.LastModifiedDate, Section.Document.LastModifiedDate, Encounter)
                    If Not objData Is Nothing Then
                        Err.Raise 10000, "Save Note DL", "Unable to save S/O Note. S/O Note has deleted by another user."
                        Exit Function
                    End If
                    
                    Set objData = Nothing
                End If
            End If
            
            'Refresh Encounter Status
            'mobjEncounter.Refresh True
            
            'If status in NeedCoSignature and user is PP or CS then save note SCR-19388
            If (Encounter.Status = complete) Or _
                ((Encounter.Status = NeedsCoSignature) And Not (UserIsPrimaryProviderorCosigner(Encounter))) Or _
                (Encounter.Status = Updated) Or _
                ((Encounter.Status = Updating) And Not (UserIsPrimaryProviderorCosigner(Encounter))) Then
                
                'If Not blnIgnoreStatus Then 'SCR 17773 This was added for the data transfer.  You want to ingnore the status if
                  'Encounter.Refresh                'this method is being called during a data transfer
                 
                  Err.Raise 100000, "Encounter.SaveSONote DL", "Your lock on the S/O section of the current encounter has been taken away by another user." & vbCrLf & _
                          "This encounter has been completed and signed by " & Encounter.WhoSigned.FullName & " on " & Format$(Encounter.WhenSignedDate, gsDefaultDateTime) & ".  This S/O note will not be saved to the encounter."
                          
                  Exit Function
                'End If
                
            End If
            
            SaveSONote objNote, Encounter
            
            If State.State = osNew Then
                SO.Add objNote, CStr(Section.DataID)
                'Call mobjEncounter.LockingObject.LockSection(eSO, objNote.mnDataID, SectionLock)
            Else
                On Error Resume Next
                'Only for theater, collection has old note, so remove and then add again SCR-24107
                SO.Remove CStr(Section.DataID)
                SO.Add objNote, CStr(Section.DataID)
                
                On Error GoTo ErrHandler
                Call EndSectionUpdate(eSO, Section.DataID)
                
            End If
            
            State.State = osSaved
            
        ElseIf State.State = osDelete Then
            'delete note
            
        End If
    Next
    
    Exit Function
    
ErrHandler:
    'Call SaveSONote
End Function

Public Function UserIsPrimaryProviderorCosigner(ByRef Encounter As Encounter) As Boolean

    If (Encounter.Providers(1).NCID = gobjShared.CurrentUserNCID) Or (Encounter.Cosigner.NCID = gobjShared.CurrentUserNCID) Then
        UserIsPrimaryProviderorCosigner = True
    End If
End Function

Public Function SaveSONote(ByRef Note As SONote, ByRef Encounter As Encounter, Optional ByVal blnIgnoreStatus As Boolean = False) As Boolean
' Assume section NOT already locked for update before getting here
' Only ListNotes are saved here... chartnotes saves its own data, and we read it from the CDR
    Dim objData As DataString
    Dim bWasNew As Boolean
    Dim Section As iSection
    Dim State As iState
    Dim oRecord As MedcinRecord
    Dim objEvent As ClinicalEvent
    Dim sTemp As String

    On Error GoTo ErrHandler

    Set Section = Note
    
    If Section.DataID = 0 Then
        bWasNew = True
    Else
        If StartSectionUpdate(eSO, Section.DataID) Then
            Err.Raise 100000, "Encounters DL", "Another user has taken away your lock on this note."
            GoTo Cleanup
        End If
    End If
    
    Set objData = InitializeDataString(eSO, Section.DataID, Section.Document.LastModifiedDate, Section.Document.LastModifiedDate, Encounter)
    If Not objData Is Nothing Then 'Note may already be deleted SCR-26756
        Set objEvent = objData.PatientData
        Call AddRTF(objEvent, Section.Document.body)
        
        If Len(gobjShared.CmdLineSwitches("PKC")) Then 'SCR#44241
            If Note.ReadOnly Then
                sTemp = "T"
            End If
            Call AddCommentValuePair(objEvent.Header, "READONLY", sTemp)
            sTemp = vbNullString
            If Note.NewCouplerData Then
                sTemp = "T"
            End If
            Call AddCommentValuePair(objEvent.Header, "NEWCOUPLERDATA", sTemp)
            sTemp = vbNullString
            If Note.CouplerData Then
                sTemp = "T"
            End If
            Call AddCommentValuePair(objEvent.Header, "COUPLERDATA", sTemp)
        End If
        
        If Not Note.MedcinRecords Is Nothing Then
            For Each oRecord In Note.MedcinRecords
                WriteMedcinRecordToClinObs oRecord, objData.PatientData
                'Check if any of the notes has family history
                'If objLTR.prefix = "F" Then
                '    mobjEncounter.HasFamilyHx = True
                'End If
            Next
        End If
    End If
    
    If objData.Save Then
        Section.DataID = objData.DataID
        'SaveNote = True
        If Encounter.Status = CheckedIn Or Encounter.Status = Waiting Then
            Encounter.Status = InProgress
            
            Set State = Encounter
            State.State = osModified
        End If
    Else
        'mobjEncounter.DisplayDetailedError "SO.Save", ERROR_MESSAGE_SAVE_FAILURE
        'Err.Raise 100000, "Encounters DL", ERROR_MESSAGE_SAVE_FAILURE & vbCrLf & "SOSave(" & Encounter.PatientUnitNum & " - " & Encounter.FacilityNCID & " - " & mobjEncounter.EncounterID & ")", vbCritical  'SCR-39485, 38953
    End If
    GoTo Cleanup
ErrHandler:
    If Err = -2147220874 Or Err = -2147417848 Or Err = -2147220877 Then
       Err.Raise 100000, "Encounters DL", "Unable to save the current note. The note that you are attempting to save is too large. Please save some of the notes contents to a new note", vbOKOnly + vbInformation, "EncounterOps"
       Err.Clear
    Else
       'Call mobjEncounter.CWShared.ShowVBError(Err.Number, Err.Description, "SONote.Save ", "EncounterOps", vbCritical)
    End If
Cleanup:
    On Error Resume Next

    Set objData = Nothing
'    If SaveNote Then
'        mobjEncounter.Comm.SetStatusText "Encounter " & mobjEncounter.EncounterID & " SO Note was saved."
'    Else
'        mobjEncounter.Comm.SetStatusText ""
'    End If
End Function

Private Function iSectionDataMapper_SaveVitals(Encounter As DTOs.Encounter) As Boolean
' Assume section already locked for update before getting here
    Dim objData As DataString
    Dim objEvent As ClinicalEvent
    Dim oRecord As MedcinRecord
    
    Dim Vitals As Vitals
    
    Dim VitalsSect As iSection
    
    On Error GoTo ErrHandler
    
    Set Vitals = Encounter.Sections("Vitals")
    
    Set VitalsSect = Vitals

    Set objData = InitializeDataString(eVitals, VitalsSect.DataID, Now, Now, Encounter)
    
    Set objEvent = objData.PatientData
    
    Call AddRTF(objData.PatientData, VitalsSect.Document.body)

    ' SnoIDs from Vitals are now stored for E&M Calculation SCR-11863
    For Each oRecord In Vitals.MedcinRecords
        'only store SnoIDs SCR-16307
        WriteMedcinRecordToClinObs oRecord, objEvent, True
    Next
    
    If objData.Save Then
        VitalsSect.DataID = objData.DataID
'        Save = True
'        mbDirty = False
    Else
        'RAISE ERROR
        'mobjEncounter.DisplayDetailedError "Vitals.Save", ERROR_MESSAGE_SAVE_FAILURE
        'Err.RAISE 100000, "Encounters DL",  ERROR_MESSAGE_SAVE_FAILURE & vbCrLf & "VitalsSave(" & mobjEncounter.PatientID & " - " & mobjEncounter.FacilityNCID & " - " & mobjEncounter.EncounterID & ")", vbCritical       'SCR-39485, 38953
    End If
    GoTo Cleanup
ErrHandler:
    'RAISE ERROR
    'Call mobjEncounter.CWShared.ShowVBError(Err.Number, Err.Description, "Enctr.Save ", "Enc", vbCritical)
Cleanup:
    Set objData = Nothing
    Set objEvent = Nothing
    Set oRecord = Nothing
    
'    If Save Then
'        mobjEncounter.Comm.SetStatusText "Encounter " & mobjEncounter.EncounterID & " " & SectionEnumToName(mobjParent.meSection) & " section was saved."
'    Else
'        mobjEncounter.Comm.SetStatusText ""
'    End If
End Function

Private Function LoadSectionData(ByRef Section As iSection, ByVal SectionEnum As EncSectionEnum)
    Dim objData As DATASTRATLLib.DataString
    Dim objEvent As CLINTYPEATLLib.ClinicalEvent
    Dim i As Long


    With Section

        .DataID = 0

        'Using the For...Each method with 3M objects causes memory leak
        For i = 1 To mcolDataStrings.Count
        'For Each objData In colDataStrings
            Set objData = mcolDataStrings(i)
            If objData.DataTypeNCID = NCID_ENCEVENT Then
                Set objEvent = objData.PatientData
                If objEvent.Header.ObsBatId.Value.Coded.ValueConcept.UniqueId = SectionEnumToNCID(SectionEnum) Then
                    .DataID = objData.DataID
                    If Not objData.ModifyClinician Is Nothing Then
                        .Document.OwnerNCID = objData.ModifyClinician.ValueConcept.UniqueId
                        .Document.OwnerName = objData.ModifyClinician.ValueConcept.PrefRep("2000").Representation
                    ElseIf Not objData.CreateClinician Is Nothing Then
                        .Document.OwnerNCID = objData.CreateClinician.ValueConcept.UniqueId
                        .Document.OwnerName = objData.CreateClinician.ValueConcept.PrefRep("2000").Representation
                    End If
                    .Document.LastModifiedDate = objEvent.Header.ObsDateTime.StartTime.Value.DateTime
                    If objEvent.ClinObs.Count > 0 Then
                        .Document.body = ExtractRTF(objEvent.ClinObs.Item(1))
                    End If
                    Exit For
                End If
            End If
        Next 'objData
    End With

End Function
