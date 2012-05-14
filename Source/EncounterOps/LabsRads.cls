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
Attribute VB_Name = "LabsRads"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mobjEncounter As Encounter
Private mobjParent As SectionParent

Private msWhat As String
Private meSection As EncounterSectionEnum
Private mcolCitations As Collection

Public Property Get RTF() As String
    Dim objNote As TextNote
    If meSection = eLabs Then
        RTF = mobjEncounter.SectionParent(eAutoCite).mobjSection.LabRTF
    ElseIf meSection = eRads Then
        RTF = mobjEncounter.SectionParent(eAutoCite).mobjSection.RadRTF
    Else
        RTF = mobjEncounter.SectionParent(eAutoCite).mobjSection.QuestionnaireRTF
    End If
    For Each objNote In mcolCitations
        RTF = RTF & FormatRtfDisplay(objNote.msRTF, msWhat, "Cited", objNote.msOwnerName, objNote.mdDTS)
    Next objNote
End Property

Public Sub Load(ByRef objParent As SectionParent, ByRef colDataStrings As Collection)
    Dim objData As DATASTRATLLib.DataString
    Dim objEvent As CLINTYPEATLLib.ClinicalEvent
    Dim objNote As TextNote
    Dim i As Long
    meSection = objParent.meSection
    If meSection = eLabs Then
        msWhat = "Lab Result"
    ElseIf meSection = eRads Then
        msWhat = "Rad Result"
    Else
        msWhat = "Questionnaire Result"
    End If
    Set mcolCitations = New Collection
    
    'Using the For...Each method with 3M objects causes memory leak
    For i = 1 To colDataStrings.Count
    'For Each objData In colDataStrings
        Set objData = colDataStrings(i)
        If objData.DataTypeNCID = NCID_ENCEVENT Then
            Set objEvent = objData.PatientData
            If objEvent.Header.ObsBatId.value.Coded.ValueConcept.UniqueId = SectionEnumToNCID(meSection) Then
                Set objNote = New TextNote
                Call objNote.ReadFromDataString(objData)
                If mcolCitations.Count = 0 Then
                    mcolCitations.Add objNote
                Else
                    mcolCitations.Add objNote, , 1
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

Friend Function AddRTF(ByRef RTF As String) As Boolean
    Dim objData As DataString
    Dim objNote As TextNote
    On Error GoTo ErrHandler
    Set objNote = New TextNote
    With objNote
        .mbIsComplete = True
        .mdDTS = Now
        .msOwnerNCID = mobjEncounter.Logon.UserNCID
        .msOwnerName = mobjEncounter.Logon.UserName
        .msRTF = RTF

    End With
    
    If GEMS_Save(objNote) Then
        AddRTF = True
        mcolCitations.Add objNote
    End If
    
    GoTo CleanUp
ErrHandler:
    Call mobjEncounter.CWShared.ShowVBError(Err.Number, Err.Description, "LabsRads.Save", "EncounterOps", vbCritical)
CleanUp:
    Set objData = Nothing
    Set objNote = Nothing
    If AddRTF Then
        mobjEncounter.Comm.SetStatusText "Encounter " & mobjEncounter.EncounterID & " " & msWhat & " Citation was saved."
    Else
        mobjEncounter.Comm.SetStatusText ""
    End If
End Function

Private Sub Class_Initialize()
    Set mcolCitations = New Collection
End Sub

Private Sub Class_Terminate()
    Set mcolCitations = Nothing
End Sub
Public Sub GEMS_Load(ByRef objParent As SectionParent, objRS As ADODB.Recordset)
    Dim objNote As TextNote
    meSection = objParent.meSection
    If meSection = eLabs Then
        msWhat = "Lab/Rad Result"
    Else
        msWhat = "Rad Result"
    End If
    Set mcolCitations = New Collection
    Do While Not objRS.EOF
        Set objNote = New TextNote
        Call objNote.GEMS_ReadFromRecordSet(objRS)
        If mcolCitations.Count = 0 Then
            mcolCitations.Add objNote
        Else
            mcolCitations.Add objNote, , 1
        End If
        Set objNote = Nothing
        objRS.MoveNext
    Loop
    
End Sub

Public Function GEMS_Save(ByRef objNote As TextNote) As Boolean
' Assume section NOT already locked for update before getting here
    
    Dim bWasNew As Boolean
    Dim sNoteType As String
    Dim sql As String
    Dim oSQLOps As ICHCSII_SQL.ISqlOps
    Dim rtf_index As Long
    
    Set oSQLOps = gobjCHCSIIConn.CHCSII_SQLOPS(Auto)  '''<SCR 36493
    
    On Error GoTo ErrHandler
    
    If objNote.mnDataID = 0 Then
        bWasNew = True
    Else
        If Not mobjEncounter.LockingObject.StartSectionUpdate(meSection, objNote.mnDataID) Then
            GoTo CleanUp
        End If
    End If
    
    If meSection = eLabs Then
        sNoteType = "LABNOTE"
    ElseIf meSection = eRads Then
        sNoteType = "RADNOTE"
    Else
        sNoteType = "QNRNOTE"
    End If
        
    If objNote.mnDataID = 0 Then
    
        rtf_index = oSQLOps.GetNextID()
        sql = "insert into enc_rtfs (ENCOUNTERNUMBER,FACILITYNCID,DTS, rtftype, ENC_RTFSINDEX ) "
        sql = sql & " values (" & mobjEncounter.EncounterID & "," & mobjEncounter.FacilityNCID & ","
        sql = sql & gobjCHCSIIConn.SQLDate(objNote.mdDTS) & ", '" & sNoteType & "', " & rtf_index & ")"

        oSQLOps.Execute sql
        objNote.mnDataID = rtf_index
    End If
    Call objNote.GEMS_WriteToDataString
    
    GEMS_Save = True

    GoTo CleanUp
ErrHandler:
    Call mobjEncounter.CWShared.ShowVBError(Err.Number, Err.Description, "TextNote.GEMS_Save ", "EncounterOps", vbCritical)
CleanUp:
    Set oSQLOps = Nothing
End Function
