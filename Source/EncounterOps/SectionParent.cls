VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SectionParent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IParent

Public meSection As EncounterSectionEnum
Public mobjSection As Object
Public mobjLocking As Locking
Public msAppNCID As String
'<< SCR  10768
Public msOKToCloseView As String
Public msOKToChangePatient As String
'>>
Private mobjEncounter As Encounter
Private mcolMenuIDs As Collection
Private mlEncounterProgID As Long

Public Property Get RTF() As String
    If Not mobjSection Is Nothing Then
        RTF = mobjSection.RTF
    End If
End Property

Friend Function OpenEncounter(ByRef objEncounter As Encounter) As Boolean
    On Error Resume Next
    Set mobjEncounter = objEncounter
    mlEncounterProgID = mobjEncounter.EncounterParent.ProgId
    Set mcolMenuIDs = New Collection
    Call mobjSection.OpenEncounter(Me, mobjEncounter)
End Function
Friend Function OpenSection() As Boolean
    OpenSection = mobjSection.OpenSection()
    
    'Only update status is section open SCR-769 ITT
    If OpenSection Then
        'Update status when module is opened SCR-21414
        If mobjEncounter.Status = CheckedIn _
        Or mobjEncounter.Status = Waiting Then
            If meSection = eSO _
            Or meSection = eap _
            Or meSection = eDental _
            Or meSection = eDisp Then
                mobjEncounter.Status = InProgress
            End If
        End If
    End If
End Function

Friend Function CloseSection() As Boolean
    On Error Resume Next
    CloseSection = mobjSection.CloseSection()
    Err.Clear
End Function

Friend Function CloseEncounter()
    Call mobjSection.CloseEncounter
    Set mobjEncounter = Nothing
End Function

Public Function Load(ByRef objEncounter As Encounter, ByVal eSection As EncounterSectionEnum, ByRef colDataStrings) As Boolean
    ' we can get here multiple times what with loads and refreshes...
    ' if se already have a section object, all we do is reload the datastrings
    Dim oShared As CWShared
    Set oShared = New CWShared
    If mobjSection Is Nothing Then
        Set mobjLocking = objEncounter.LockingObject
        meSection = eSection
        msAppNCID = SectionEnumToNCID(meSection)
        Select Case meSection
        Case eAutoCite
            Set mobjSection = New AutoCites
        Case eRFV
            Set mobjSection = New RFV
        Case eVitals
            Set mobjSection = New Vitals
        Case eSO
            Set mobjSection = New SO
        Case eLabs, eRads
            Set mobjSection = New LabsRads
            'MAG 05/02/01 Fixes SCR#12641
            mobjSection.OpenEncounter Me, objEncounter
        Case eQuestionnaire
            Set mobjSection = New LabsRads
            'MAG 05/02/01 Fixes SCR#12641
            mobjSection.OpenEncounter Me, objEncounter
        Case eEducation
            Set mobjSection = New Education
        Case eap
            Set mobjSection = New AP
        Case eDental
            Set mobjSection = New Dental
        Case eDisp
            Set mobjSection = New Disposition
        Case eNotes
            Set mobjSection = New TextNotes
        Case eHistory
            Set mobjSection = New History
        Case Else
            MsgBxARMd "Internal error: unexpected value (" & meSection & ") for eSection in SectionParent.Init().", vbCritical, "Encounter"
            Exit Function
        End Select
    End If
    '<SCR 36493
    If (Not oShared.IsAppMode(modeCHCSI_GUI)) Or _
       (oShared.IsAppMode(modeCHCSI_GUI) And (eSection = eap Or eSection = eDisp)) _
    Then
    '>
        Load = mobjSection.Load(Me, colDataStrings)
    End If
End Function
Public Function GEMS_Load(ByRef objEncounter As Encounter, _
                          ByVal eSection As EncounterSectionEnum, _
                                objRS As ADODB.Recordset) As Boolean
                                
    ' we can get here multiple times what with loads and refreshes...
    ' if se already have a section object, all we do is reload the datastrings
    If mobjSection Is Nothing Then
        Set mobjLocking = objEncounter.LockingObject
        meSection = eSection
        msAppNCID = SectionEnumToNCID(meSection)
        Select Case meSection
        Case eAutoCite
            Set mobjSection = New AutoCites
            'mobjSection.OpenEncounter Me, objEncounter
        Case eRFV
            Set mobjSection = New RFV
            mobjSection.OpenEncounter Me, objEncounter
        Case eVitals
            Set mobjSection = New Vitals
            mobjSection.OpenEncounter Me, objEncounter
        Case eSO
            Set mobjSection = New SO
            mobjSection.OpenEncounter Me, objEncounter
        Case eLabs, eRads, eQuestionnaire
            Set mobjSection = New LabsRads
            mobjSection.OpenEncounter Me, objEncounter
        Case eap
            Set mobjSection = New AP
            mobjSection.OpenEncounter Me, objEncounter
        Case eDental
            Set mobjSection = New Dental
            mobjSection.OpenEncounter Me, objEncounter
        Case eDisp
            Set mobjSection = New Disposition
            mobjSection.OpenEncounter Me, objEncounter
        Case eNotes
            Set mobjSection = New TextNotes
            mobjSection.OpenEncounter Me, objEncounter
        Case eHistory
            Set mobjSection = New History
            mobjSection.OpenEncounter Me, objEncounter
        Case Else
            MsgBxARMd "Internal error: unexpected value (" & meSection & ") for eSection in SectionParent.Init().", vbCritical, "Encounter"
            Exit Function
        End Select
    Else
        'SCR 42939
        If meSection = eap Then
            mobjSection.OpenEncounter Me, objEncounter
        End If
    End If
    
    GEMS_Load = mobjSection.GEMS_Load(Me, objRS)
End Function

Friend Sub MenuItemSelected(ByVal MenuID As Long)
    On Error Resume Next
    Call mobjSection.MenuItemSelected(MenuID)
End Sub

Private Sub Class_Terminate()
    Set mobjLocking = Nothing
    Set mobjSection = Nothing
    Set mcolMenuIDs = Nothing
End Sub

Private Function IParent_AddMenuItem(ByVal iMenu As MenuNames, ByVal sMenuText As String, ByVal sMenuBMPFileName As String, ByVal sTip As String) As Long
    Select Case iMenu
        Case mnAction
            IParent_AddMenuItem = mobjEncounter.Comm.AddActionMenuItem(sMenuText, sMenuBMPFileName, mlEncounterProgID, sTip)
        Case mnView
            IParent_AddMenuItem = mobjEncounter.Comm.AddViewMenuItem(sMenuText, mlEncounterProgID, sTip)
        Case mnPrint  '--- SCR 84207  Sherry Wang 1/9/2006
            IParent_AddMenuItem = mobjEncounter.Comm.AddPrintMenuItem(sMenuText, mlEncounterProgID, sTip)
        Case mnPrintPreview '--- SCR 84207  Sherry Wang 1/9/2006
            IParent_AddMenuItem = mobjEncounter.Comm.AddPrintPreviewMenuItem(sMenuText, mlEncounterProgID, sTip)
    End Select
    
    If IParent_AddMenuItem > 0 Then
        mcolMenuIDs.Add IParent_AddMenuItem
    End If
End Function

Private Sub IParent_ClosedView()
    Dim lMenuID As Variant
    On Error Resume Next
    Call mobjEncounter.SectionClosed
    For Each lMenuID In mcolMenuIDs
        Call mobjEncounter.Comm.DeleteMenuItem(lMenuID)
    Next lMenuID
    Call mobjSection.ClosedView
End Sub

Private Property Get IParent_CurrentEncounter() As IEncDoc
    Set IParent_CurrentEncounter = mobjEncounter
End Property

Private Function IParent_EndSave(Optional ByVal nEventId As Long) As Boolean

    If meSection <> eSO Then
        ' SO saves are done separately, but Chart Notes does call this method
        ' in order to handle locking stuff
        mobjSection.Save
    End If
    
    If mobjEncounter.CWShared.IsAppMode(modeCHCSI_GUI) Then 'SCR 36493
        Exit Function
    End If
    
    IParent_EndSave = mobjLocking.EndSectionUpdate(meSection, nEventId)
    If meSection = eap Then
      'SCR 51743 - Release A/P header update lock of encounter provider data in addition to the section lock
      IParent_EndSave = IParent_EndSave And mobjLocking.EndSectionUpdate(eHeader, nEventId)
    End If
    
    'temp fix for theater, since gems_loadEncounter set all the checkedIn to InProgress.
    If mobjEncounter.Status = CheckedIn _
    Or mobjEncounter.Status = Waiting Then
        If meSection = eSO _
        Or meSection = eap _
        Or meSection = eDental _
        Or meSection = eDisp Then
            mobjEncounter.Status = InProgress
        End If
    End If
    If mobjEncounter.Status = CheckedIn Then
        If meSection = eVitals Then
            mobjEncounter.Status = Waiting
        End If
    End If
End Function

Private Property Get IParent_Hwnd() As Long
    IParent_Hwnd = mobjEncounter.EncounterParent.hwnd
End Property

Private Function IParent_LockEventID(ByVal nEventId As Long) As Boolean
    IParent_LockEventID = mobjLocking.LockSection(meSection, nEventId, SectionLock)
End Function

Private Function IParent_RemoveMenuItem(ByVal lMenuID As Long) As Boolean
    mobjEncounter.Comm.DeleteMenuItem lMenuID
End Function

Private Property Get IParent_SectionObject() As Object
    Set IParent_SectionObject = mobjSection
End Property

Private Sub IParent_SetRTF(RTF As String, Optional Updated As Boolean)
    On Error Resume Next
    
    Select Case meSection
        Case eDisp, eap
            Call mobjSection.SetRTF(RTF, Updated)
        Case Else
            Call mobjSection.SetRTF(RTF)
    End Select
End Sub

Private Sub IParent_SetStatus(ByVal eStatus As EncStatus)
    On Error Resume Next
    Call mobjSection.SetStatus(eStatus)
End Sub

Private Function IParent_StartSave(Optional ByVal nEventId As Long) As Boolean

   IParent_StartSave = mobjLocking.StartSectionUpdate(meSection, nEventId)
   
   If meSection = eap And IParent_StartSave Then
      'SCR 51743 - A/P will now obtain an update lock on encounter header so it can save encounter provider data - make sure this lock has not been broken
      IParent_StartSave = mobjLocking.StartSectionUpdate(eHeader, nEventId)
   End If
    
End Function

Private Function IParent_UnlockEventID(ByVal nEventId As Long) As Boolean
    IParent_UnlockEventID = mobjLocking.UnlockSection(meSection, nEventId)
End Function
