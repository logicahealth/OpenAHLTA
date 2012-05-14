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
Attribute VB_Name = "TextNotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum NotesDividedEnum
    NotesAll = 0
    NotesBeforeSigning = 1
    NotesAfterSigning = 2
End Enum

Private mobjEncounter As Encounter
Private mobjParent As SectionParent

Private mcolNotes As Collection

Public Property Get Count() As Integer

    'SCR-32476
    If Not mcolNotes Is Nothing Then
        Count = mcolNotes.Count
    Else
        Count = 0
    End If
    
End Property

Public Property Get TextNote(ByVal Index As Integer) As TextNote
    Set TextNote = mcolNotes(Index)
End Property

Public Property Get NoteRTF(ByVal Index As Integer) As String
    Dim objNote As TextNote
    On Error Resume Next
    If mcolNotes.Count < Index Then
        Exit Property
    End If
    Set objNote = mcolNotes(Index)
    NoteRTF = objNote.RTF
    Set objNote = Nothing
End Property

Public Property Get RTF(Optional bGreen As Boolean = False, Optional RelatedSection As EncounterSectionEnum = eNotes) As String
    RTF = RTFDivided(NotesAll, bGreen, RelatedSection)
End Property

Public Property Get RTFDivided(ByVal eDivided As NotesDividedEnum, Optional bGreen As Boolean = False, Optional RelatedSection As EncounterSectionEnum = eNotes) As String
    Dim objNote As TextNote
    Dim sDraft As String
    Dim bContinue As Boolean
    On Error Resume Next
    If mcolNotes.Count < 1 Then
        Exit Property
    End If
    For Each objNote In mcolNotes
        If mobjEncounter.CWShared.IsAppMode(modeCHCSI_GUI) Then
            'Use related section to differentiate between SO and Text Notes SCR-38142
            If RelatedSection <> objNote.RelatedSection Then
                GoTo NextNote
            End If
        End If
            
        If eDivided = NotesBeforeSigning Then
            If mobjEncounter.WhenSigned <> 0 _
            And mobjEncounter.WhenSigned < objNote.mdDTS Then
                GoTo NextNote
            End If
        ElseIf eDivided = NotesAfterSigning Then
            If mobjEncounter.WhenSigned = 0 _
            Or mobjEncounter.WhenSigned > objNote.mdDTS Then
                GoTo NextNote
            End If
        End If
        
        If RTFDivided = "" Then
            'Set default font Swiss to f999, textcontrol creates f0, f1, ... when
            'changing fonts in text note SCR-17791
            RTFDivided = "{\rtf1\ansi\ucl\deff0\deflang1033\deflangfe1033" _
                & "{\fonttbl{\f999\fswiss\fcharset0\fprq2 Arial;}}" _
                & "{\colortbl;\red0\green0\blue0;}"
        Else
            If bGreen Then
                RTFDivided = RTFDivided & " \pard\plain\f999\fs16\par "
            Else
                RTFDivided = RTFDivided & " \pard\plain\f999\fs16\par "
            End If
        End If
        If Not objNote.mbIsComplete Then
            sDraft = " (Draft)"
        Else
            sDraft = ""
        End If
'<< shaw
        If Not mobjEncounter.CWShared.IsAppMode(modeCHCSI_GUI) Then 'SCR 36493
            RTFDivided = RTFDivided & "\pard\plain\f999\fs16\b\ul Note" & sDraft & " \b0 Written by " _
                & objNote.msOwnerName & " @ " & FormatDTG(objNote.mdDTS) & "\par\ul0\b "
        Else
            ' Commented out SCR-39840
            'If Not objNote.ReadOnly Then
            RTFDivided = RTFDivided & "\pard\plain\f999\fs16\b\ul Note" & sDraft & " \b0 Written by " _
                & objNote.msOwnerName & " @ " & FormatDTG(objNote.mdDTS) & "\ul0\b \par "
            'End If
        End If
        If eDivided = NotesAfterSigning Then
            RTFDivided = RTFDivided & "(Added after encounter was signed.) \par "
        End If
        If objNote.msCategory <> "" Then
            RTFDivided = RTFDivided & objNote.msCategory & " \par "
        End If
        If objNote.msTitle <> "" Then
            RTFDivided = RTFDivided & objNote.msTitle & " \par "
        End If
        RTFDivided = RTFDivided & "\pard\plain \widctlpar\adjustright \cgrid" & objNote.msRTF
NextNote:
    Next objNote
    If RTFDivided <> "" Then
        RTFDivided = RTFDivided & "}"
    End If
End Property

Public Function DeleteTextNote(ByVal Index As Integer) As Boolean
    Dim objNote As TextNote
    Dim objData As DataString
    Dim bWasLocked As Boolean
    On Error GoTo ErrHandler
    Set objNote = mcolNotes(Index)
    bWasLocked = mobjEncounter.LockingObject.SectionIsLocked(eNotes, objNote.mnDataID)
    If Not bWasLocked Then
        If Not mobjEncounter.LockingObject.LockSection(eNotes, objNote.mnDataID, SectionLock, "Deleting Note") Then
            GoTo CleanUp
        End If
    End If
    If Not mobjEncounter.LockingObject.StartSectionUpdate(eNotes, objNote.mnDataID) Then
        GoTo CleanUp
    End If
    If Not mobjEncounter.AddHistoryItem(Deletion, "Text Note", NoteRTF(Index)) Then
        MsgBxARMd "Could not add to encounter change history section. Did not delete note.", vbCritical, "Encounter Text Note"
        GoTo CleanUp
    End If
    
    DeleteTextNote = GEMS_DeleteTextNote(Index, objNote.mnDataID)
    GoTo CleanUp

ErrHandler:
    Call mobjEncounter.CWShared.ShowVBError(Err.Number, Err.Description, "TextNotes.DeleteTextNote ", "EncounterOps", vbCritical)
CleanUp:
    On Error Resume Next
    Call mobjEncounter.LockingObject.EndSectionUpdate(eNotes, objNote.mnDataID)
    If Not bWasLocked Then
        Call mobjEncounter.LockingObject.UnlockSection(eNotes, objNote.mnDataID)
    End If
    Set objData = Nothing
    If DeleteTextNote Then
        mobjEncounter.Comm.SetStatusText "Encounter " & mobjEncounter.EncounterID & " Note was deleted."
    Else
        mobjEncounter.Comm.SetStatusText ""
    End If
End Function
Public Function GEMS_DeleteTextNote(ByVal Index As Integer, ByVal DataID As Long) As Boolean

  Dim objSQL   As ICHCSII_SQL.ISqlOpsEx
  Set objSQL = gobjCHCSIIConn.CHCSII_SQLOPS(Auto)  '''<SCR 36493

  objSQL.Execute "DELETE FROM enc_rtfs where ENC_RTFSINDEX = " & DataID
    
  GEMS_DeleteTextNote = True
  mcolNotes.Remove Index
    
  Exit Function
  
ErrHandler:
    Call mobjEncounter.CWShared.ShowVBError(Err.Number, Err.Description, "TextNotes.GEMS_DeleteTextNote ", "EncounterOps", vbCritical)
End Function


Public Sub Load(ByRef objParent As SectionParent, ByRef colDataStrings As Collection)
    Dim objData As DATASTRATLLib.DataString
    Dim objEvent As CLINTYPEATLLib.ClinicalEvent
    Dim objNote As TextNote
    Dim i As Long
    Set mcolNotes = New Collection
    
    'Using the For...Each method with 3M objects causes memory leak
    For i = 1 To colDataStrings.Count
    'For Each objData In colDataStrings
        Set objData = colDataStrings(i)
        If objData.DataTypeNCID = NCID_ENCEVENT Then
            Set objEvent = objData.PatientData
            If objEvent.Header.ObsBatId.value.Coded.ValueConcept.UniqueId = SectionEnumToNCID(eNotes) Then
                Set objNote = New TextNote
                Call objNote.ReadFromDataString(objData)
                ' the events are given to use in descending order, so reverse them
                If mcolNotes.Count > 0 Then
                    mcolNotes.Add objNote, CStr(objNote.mnDataID), 1
                Else
                    mcolNotes.Add objNote, CStr(objNote.mnDataID)
                End If
                Set objNote = Nothing
            End If
        End If
    Next 'objData
End Sub

Public Sub OpenEncounter(ByRef objParent As SectionParent, ByRef objEncounter As Encounter)
    Set mobjParent = objParent
    Set mobjEncounter = objEncounter
End Sub

Public Sub CloseEncounter()
    Set mobjParent = Nothing
    Set mobjEncounter = Nothing
End Sub

Public Function OpenSection() As Boolean
    On Error Resume Next
'<<SCR 10673,10675(dup) Encounter data refresh before children load
'   Dropping this line to stop the refresh
'   mobjEncounter.RefreshDisplay
'>>
    If Not mobjEncounter.SectionIsEditable(eNotes) Then
        MsgBxARMd "This encounter has been saved by another user and you can no longer edit the Text Note section.", vbInformation, "Encounter"
        Exit Function
    End If
    Call frmSelectNote.Initialize(mobjEncounter)
    mobjEncounter.Comm.Message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_ENCOUNTER_APPLICATION
    frmSelectNote.Show vbModal
    mobjEncounter.Comm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_ENCOUNTER_APPLICATION
    'RefreshDisplay will be called by Encounters hideChild SCR-27066
    'Call mobjEncounter.EncounterParent.RefreshDisplay(True)
End Function

Public Function CloseSection() As Boolean
        CloseSection = True
End Function

Public Function Save(ByRef objNote As TextNote) As Boolean
' Assume section NOT already locked for update before getting here
    Dim objData As DataString
    Dim bWasNew As Boolean
    Dim bBrokenLock As Boolean
    Dim WarningMessage  As String
    Dim sNoteType As String
    
    On Error GoTo ErrHandler

    If objNote.mnDataID = 0 Then
        bWasNew = True
    End If
    
    If Not mobjEncounter.LockingObject.StartSectionUpdate(eNotes, objNote.LockID) Then 'SCR# 47028
        'Refresh to see if encounter was signed  SCR=35244
        mobjEncounter.Refresh
        
        If mobjEncounter.CWShared.IsAppMode(modeCHCSI_GUI) Then
            sNoteType = "S/O Note"
        Else
            sNoteType = "Text Note"
        End If
        
        WarningMessage = "Your lock for this note has been broken and the " & sNoteType & "  will be added as free text to the end of the encounter as an addendum."
        
        If Not (mobjEncounter.WhoSigned Is Nothing) Then
          WarningMessage = WarningMessage & "  This encounter has been completed and signed by " & _
                            mobjEncounter.WhoSigned.FullName & " on " & _
                            Format$(mobjEncounter.WhenSigned, gsDefaultDateTime) & ".  "
        End If
        
        MsgBxARMd WarningMessage, vbExclamation + vbOKOnly, "Encounter"
    
        'Add SO note in Add Note section SCR#44686
        mobjEncounter.AddTextNote "Additional " & sNoteType & " Information:", objNote.msRTF

        If Not (mobjEncounter.WhoSigned Is Nothing) Then
          mobjEncounter.ReplaceEncRTF
        End If
    
        bBrokenLock = True
        GoTo CleanUp
    End If

    Save = GEMS_Save(objNote)
    GoTo CleanUp

ErrHandler:
    Call mobjEncounter.CWShared.ShowVBError(Err.Number, Err.Description, "TextNote.Save ", "EncounterOps", vbCritical)
CleanUp:
    On Error Resume Next
    If bWasNew And Save Then
        If mcolNotes Is Nothing Then
            Set mcolNotes = New Collection
        End If
        mcolNotes.Add objNote, CStr(objNote.mnDataID)
    End If

    If bBrokenLock Then
        Save = True
    End If
    Call mobjEncounter.LockingObject.EndSectionUpdate(eNotes, objNote.mnDataID)
    Call mobjEncounter.LockingObject.EndSectionUpdate(eNotes, 0)

    Set objData = Nothing
End Function

Public Sub GEMS_Load(ByRef objParent As SectionParent, objRS As ADODB.Recordset)
    
    Dim objNote As TextNote
    Set mcolNotes = New Collection
    Do While Not objRS.EOF
        Set objNote = New TextNote
        Call objNote.GEMS_ReadFromRecordSet(objRS)
        ' the events are given to use in descending order, so reverse them
        If mcolNotes.Count > 0 Then
            mcolNotes.Add objNote, CStr(objNote.mnDataID), 1
        Else
            mcolNotes.Add objNote, CStr(objNote.mnDataID)
        End If
        Set objNote = Nothing
        objRS.MoveNext
    Loop
    
End Sub

Public Function GEMS_Save(ByRef objNote As TextNote) As Boolean
' Assume section NOT already locked for update before getting here
    
    Dim bWasNew As Boolean
    On Error GoTo ErrHandler
    If objNote.mnDataID = 0 Then
        bWasNew = True
    Else
        If Not mobjEncounter.LockingObject.StartSectionUpdate(eNotes, objNote.LockID) Then
            GoTo CleanUp
        End If
    End If
    
    'Check if encounter owner is Dr. Unassigned SCR-24349
    Call mobjEncounter.CheckEncounterOwnership
    
    Dim sql As String
    Dim oSQLOps As ICHCSII_SQL.ISqlOps
    Set oSQLOps = gobjCHCSIIConn.CHCSII_SQLOPS(Auto)  '''<SCR 36493
    Dim rtf_index As Long
        
    If objNote.mnDataID = 0 Then
        'SCR 163399
        oSQLOps.Execute "Select enc_rtfsindex from enc_rtfs where encounternumber = " & mobjEncounter.EncounterID & " and facilityncid = " & mobjEncounter.FacilityNCID & " and DTS = " & gobjCHCSIIConn.SQLDate(objNote.mdDTS) & " and rtftype = 'NOTE'"
        If Not oSQLOps.EOF Then
            rtf_index = oSQLOps("enc_rtfsindex")
        Else
            rtf_index = oSQLOps.GetNextID()
            sql = "insert into enc_rtfs (ENCOUNTERNUMBER,FACILITYNCID,DTS, rtftype, ENC_RTFSINDEX ) "
            sql = sql & " values (" & mobjEncounter.EncounterID & "," & mobjEncounter.FacilityNCID & ","
            sql = sql & gobjCHCSIIConn.SQLDate(objNote.mdDTS) & ", 'NOTE', " & rtf_index & ")"
    
            oSQLOps.Execute sql
        End If
        objNote.mnDataID = rtf_index
    End If
    Call objNote.GEMS_WriteToDataString
    
    GEMS_Save = True

    GoTo CleanUp
ErrHandler:
    If Err.Number = -2147217873 Then
        Resume Next
    Else
        Call mobjEncounter.CWShared.ShowVBError(Err.Number, Err.Description, "TextNote.GEMS_Save ", "EncounterOps", vbCritical)
        Exit Function
        Resume
    End If
CleanUp:
    Set oSQLOps = Nothing
End Function

