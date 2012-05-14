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
Attribute VB_Name = "Education"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mobjEncounter As Encounter
Private mobjParent As SectionParent
Private msRTF As String
Private mbDirty As Boolean
Private mnDataID As Long
Private mobjModule As Object
Private mobjIChild As IChild
Private mobjICoreAPI3 As CHCSII_CoreAPI3.ICHCSII_CoreAPI3
Private msProgID As String
Private mbRegularViewRunning As Boolean
Private msOwnerNCID As String
Private msOwnerName As String
Private mdDTS As Date

Public Property Get RTF() As String
    RTF = FormatRtfDisplay(msRTF, "Education", "Written", msOwnerName, mdDTS)
End Property

Public Sub SetRTF(ByRef RTF As String)
    If RTF <> msRTF Then
       msRTF = RTF
       mbDirty = True
    End If
End Sub

Public Sub Load(ByRef objParent As SectionParent, ByRef colDataStrings As Collection)
    Dim objData As DATASTRATLLib.DataString
    Dim objEvent As CLINTYPEATLLib.ClinicalEvent
    msRTF = ""
    mbDirty = False
    mnDataID = 0
    For Each objData In colDataStrings
        If objData.DataTypeNCID = NCID_ENCEVENT Then
            Set objEvent = objData.PatientData
            '///pjs 7.2.2002 scr 25333
            mdDTS = objEvent.Header.ObsDateTime.StartTime.value.DateTime
            '\\\
            If objEvent.Header.ObsBatId.value.Coded.ValueConcept.UniqueId = SectionEnumToNCID(eEducation) Then
               mnDataID = objData.DataID
               If Not objData.ModifyClinician Is Nothing Then
                  msOwnerNCID = objData.ModifyClinician.ValueConcept.UniqueId
                  msOwnerName = objData.ModifyClinician.ValueConcept.PrefRep("2000").Representation
               ElseIf Not objData.CreateClinician Is Nothing Then
                  msOwnerNCID = objData.CreateClinician.ValueConcept.UniqueId
                  msOwnerName = objData.CreateClinician.ValueConcept.PrefRep("2000").Representation
               End If
               If objEvent.ClinObs.Count > 0 Then
                  msRTF = ExtractRTF(objEvent.ClinObs.Item(1))
               End If
               Exit For
            End If
        End If
    Next objData
End Sub

Public Sub MenuItemSelected(ByVal MenuID As Long)
    On Error Resume Next
    Call mobjModule.MenuItemSelected(MenuID)
End Sub

Public Sub OpenEncounter(ByRef objParent As SectionParent, ByRef objEncounter As Encounter)
    Set mobjParent = objParent
    Set mobjEncounter = objEncounter
End Sub

Public Sub CloseEncounter()
    Set mobjParent = Nothing
    Set mobjEncounter = Nothing
    Set mobjModule = Nothing
    Set mobjIChild = Nothing
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
'    If Not mobjEncounter.LockingObject.LockSection(eEducation, 0, SectionLock) Then
'        Exit Function
'    End If
'
'    mobjIChild.SetParent mobjParent
'    mbRegularViewRunning = mobjEncounter.EncounterParent.RefreshChildView(mobjModule)
'
'    If mbRegularViewRunning Then
'        RefreshSection = True
'    Else
'        mobjIChild.SetParent Nothing
'        mobjEncounter.LockingObject.UnlockSection eEducation, 0
'    End If
'
'End Function
Public Function OpenSection() As Boolean
    On Error Resume Next
    If mbRegularViewRunning Then
        Exit Function
    End If
    If Not mobjEncounter.LockingObject.LockSection(eEducation, 0, SectionLock) Then
        Exit Function
    End If
    
    If Not mobjEncounter.SectionIsEditable(eEducation) Then
        If mobjEncounter.EncounterType = TeleConsult Then 'SCR# 46973
            MsgBxARMd "The Education Module is not available when accessed from a Telcon Encounter.", vbInformation, "Encounter"
        Else
            MsgBxARMd "You do not have the required privileges to open the Education Module.", vbInformation, "Encounter"
            'MsgBxARMd "This encounter has been saved by another user and you can no longer edit the Education section.", vbInformation, "Encounter"
            GoTo Cleanup
        End If
    End If
    If msOwnerNCID <> "" _
    And (mobjEncounter.Status = Updating Or msOwnerNCID <> mobjEncounter.Logon.UserNCID) Then
        If Not mobjEncounter.Status = Updating Then
            ' a user other than the owner is opening this section if they really meant to do this, then save a copy to the change history and
            ' save the current data as belonging to the new owner
            If MsgBxARMd("The Education section of this encounter is currently owned by " & msOwnerName & "." & vbCrLf & vbCrLf _
            & "Do you want to save a copy in change history and take over the section?", vbYesNo, "Encounter Section Ownership") = vbNo Then
                GoTo Cleanup
            End If
        End If
        If Not mobjEncounter.AddHistoryItem(Overwritten, "Education Note", RTF) Then
            MsgBxARMd "Unable to copy section to change history.", vbCritical, "Encounter"
            GoTo Cleanup
        End If
        mbDirty = True
        If Not Save Then
            GoTo Cleanup
        End If
    End If
    If mobjModule Is Nothing Then
        Set mobjModule = mobjEncounter.Comm.InitializeOLEServer(SectionEnumToNCID(eEducation))
        If mobjModule Is Nothing Then
            GoTo Cleanup
        End If
        Set mobjIChild = mobjModule
        If mobjIChild Is Nothing Then
            Set mobjModule = Nothing
            MsgBxARMd "The Educaiton application module does not support the EncounterOps.IChild interface.", vbInformation, "Encounter"
            GoTo Cleanup
        End If
        On Error Resume Next
        Set mobjICoreAPI3 = mobjModule
        Call mobjEncounter.Comm.Message(cwiGET_APPLICATION_PROGID, msProgID, NCID_CW_APPLICATION, SectionEnumToNCID(eEducation))
    End If
    Call mobjIChild.SetParent(mobjParent)
    mbRegularViewRunning = mobjEncounter.EncounterParent.OpenChildView(mobjModule)
Cleanup:
    If mbRegularViewRunning Then
        OpenSection = True
    Else
        Call mobjIChild.SetParent(Nothing)
        Call mobjEncounter.LockingObject.UnlockSection(eEducation, 0)
    End If
End Function

Public Function CloseSection() As Boolean
    On Error Resume Next
    If mbRegularViewRunning Then
        If mobjICoreAPI3 Is Nothing Then
            Call mobjModule.CloseView(cwiREGULAR_VIEW)
            CloseSection = True
        Else
            CloseSection = mobjICoreAPI3.CloseView2(cwiREGULAR_VIEW)
        End If
        If CloseSection Then
            Call mobjIChild.SetParent(Nothing)
            Call mobjEncounter.LockingObject.UnlockSection(eEducation, 0)
            mbRegularViewRunning = False
        End If
        
        'CloseSection = Not CBool(Err)
        'Err.Clear
    End If
End Function

Public Sub ClosedView() ' called by child through SectionParent
    mbRegularViewRunning = False
    Call mobjIChild.SetParent(Nothing)
    Call mobjEncounter.EncounterParent.hidechild
    Call mobjEncounter.LockingObject.UnlockSection(eEducation, 0)
End Sub

Public Function Save() As Boolean
' Assume section already locked for update before getting here
    Dim objData As DataString
    On Error GoTo ErrHandler
    
    If Not mbDirty Then
       Save = True
       Exit Function
    End If
    
    Set objData = mobjEncounter.InitializeDataString(mobjParent.meSection, mnDataID, Now, Now)
    Call AddRTF(objData.PatientData, msRTF)
    If objData.Save Then
       mnDataID = objData.DataID
       Save = True
       mbDirty = False
    End If
    GoTo Cleanup
    
ErrHandler:
    Call mobjEncounter.CWShared.ShowVBError(Err.Number, Err.Description, "Education.Save ", "Enc", vbCritical)
Cleanup:
    Set objData = Nothing
    If Save Then
        mobjEncounter.Comm.SetStatusText "Encounter " & mobjEncounter.EncounterID & " " & SectionEnumToName(mobjParent.meSection) & " section was saved."
    Else
        mobjEncounter.Comm.SetStatusText ""
    End If
End Function

Public Sub GEMS_Load(ByRef objParent As SectionParent, objRS As ADODB.Recordset)
    'Leave function shell to maintain interface
End Sub


Public Function GEMS_Save() As Boolean
    
    GEMS_Save = True
    GoTo Cleanup
    
ErrHandler:
    Call mobjEncounter.CWShared.ShowVBError(Err.Number, Err.Description, "Education.gems_Save ", "Enc", vbCritical)
Cleanup:
    If Save Then
        mobjEncounter.Comm.SetStatusText "Encounter " & mobjEncounter.EncounterID & " " & SectionEnumToName(mobjParent.meSection) & " section was saved."
    Else
        mobjEncounter.Comm.SetStatusText ""
    End If
    
End Function

