VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TeleCon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IParent

Private mobjEncounter As Encounter
Attribute mobjEncounter.VB_VarHelpID = -1
Private mobjDxModule As Object
Private mobjAnPModule As Object
Private mobjIChild As IChild
Private mobjAPSection As AP
Private mobjDispSection As Disposition
Private mobjNote As TextNote

Public ProviderNote As String
Public mbProviderNote_Changed As Boolean
Public EandMCode As String

Private Const EMPTY_CODE As String = "0"


Public Function OpenTelecon(ByRef objEncounter As Encounter, ByVal lParentHwnd As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Boolean
    Dim objAnPChild As ICHCSII_AnP
    Dim lHwnd As Long
    Dim lMouse As Long
    On Error Resume Next
    lMouse = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    Set mobjEncounter = objEncounter
    Set mobjAPSection = mobjEncounter.SectionParent(eap).mobjSection
    Set mobjDispSection = mobjEncounter.SectionParent(eDisp).mobjSection
    Set mobjAnPModule = mobjEncounter.Comm.InitializeOLEServer(NCID_ASSESMENT_AND_PLAN_APPLICATION)
    If mobjAnPModule Is Nothing Then
        MsgBxARMd "Could not initialize AnP Module."
        GoTo CleanUp
    End If
    Screen.MousePointer = vbHourglass
    Set mobjIChild = mobjAnPModule
    If mobjIChild Is Nothing Then
        MsgBxARMd "The A/P application module does not support the EncounterOps.IChild interface.", vbInformation, "Encounter"
        GoTo CleanUp
    End If
    Screen.MousePointer = vbHourglass
    Call mobjIChild.SetParent(Me)
    mobjAnPModule.InitializeAnp  'need to manually initialize AnP
    Set mobjDxModule = mobjEncounter.Comm.InitializeOLEServer(NCID_ANP_DIAGNOSIS_APPLICATION)
    Set objAnPChild = mobjDxModule
    If objAnPChild Is Nothing Then
        MsgBxARMd "The Diagnosis module does not support the ICHCSII_AnP interface."
        GoTo CleanUp
    End If
    Screen.MousePointer = vbHourglass
    Call objAnPChild.SetAnPParent(mobjAnPModule.AnPChildInterface)    'SCR 23329  Phil Crowder 4/24/02
    Set objAnPChild = Nothing
    Call mobjDxModule.OpenView(cwiENCOUNTER_SUMMARY_VIEW, lParentHwnd, lHwnd)
    Screen.MousePointer = vbHourglass
    Call mobjDxModule.ResizeView(cwiENCOUNTER_SUMMARY_VIEW, 0, 0, nWidth, nHeight)
    OpenTelecon = True
CleanUp:
    Screen.MousePointer = lMouse
End Function

Public Function Save() As Boolean
    Dim lNoteID As Long
    Dim i As Integer
    Dim objDisp As Object
    Dim objChild As IChild
    Dim lMouse As Long
    Dim prevEncStatus As EncStatus
    Dim bProviderNote_Updated As Boolean
    Dim nPos As Long
    Dim sPriorEdits As String
    
    On Error Resume Next
    
    'MAG 05/17/01 SCR#12792
    lMouse = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    DoEvents
    ' assume lock for update has already been done
    Save = True
    If Not mobjAnPModule.teleconEnd Then
        Save = False
    End If
    If Not mobjAPSection.Save Then
        Save = False
    End If
    Set objDisp = mobjEncounter.Comm.InitializeOLEServer(SectionEnumToNCID(eDisp))
    If objDisp Is Nothing Then
        MsgBxARMd "Could not initialize Disposition module."
        Save = False
    Else
        Set objChild = objDisp
        If objChild Is Nothing Then
            MsgBxARMd "The Disposition module does not implement the IChild interface."
            Save = False
        Else
            mobjDispSection.EandMNCID = EandMCode
            'If E&M code was not chosen do not mark as complete SCR-13074
            If Len(Trim$(EandMCode)) <> 0 Then
                If StrComp(EandMCode, EMPTY_CODE) Then
                    Call mobjDispSection.SetStatus(complete)
                End If
            End If
            
            Call objChild.SetParent(Me)

            Call mobjDispSection.SetRTF(objDisp.BuildRTF)
            If Not mobjDispSection.Save() Then
                Save = False
            End If
        End If
    End If
    Set objDisp = Nothing
    Set objChild = Nothing
    
    ' Moved to before AddTextNote which causes refresh and resets DXSnoID SCR-24030
    If Not mobjEncounter.Save(True) Then
        Save = False
    End If
    
    
    '>>> SCR-18260 S.McAvoy 7/23/02; If the Provider Note from the Telcon Quick
    'Entry form has been edited, then update the encounter object's TextNote.
    If mbProviderNote_Changed = True And ProviderNote <> "" Then
        If mobjEncounter Is Nothing Then
            Save = False
            GoTo ExitFunction
        End If
            
        If mobjEncounter.TextNotes.Count > 0 Then
            For i = 1 To mobjEncounter.TextNotes.Count
                If mobjEncounter.TextNotes.TextNote(i).msTitle = "Provider Telecon Note" Then

                    mobjEncounter.TextNotes.TextNote(i).msRTF = ProviderNote
                    
                    
                    Set mobjNote = mobjEncounter.TextNotes.TextNote(i)
                    mobjEncounter.Status = Updating
                    mobjEncounter.EditTextNote mobjNote
                    mobjEncounter.Status = prevEncStatus
                    bProviderNote_Updated = True
                    Exit For
                End If
            Next
            If bProviderNote_Updated = False Then
                If Not mobjEncounter.AddTextNote("Provider Telecon Note", ProviderNote, lNoteID) Then
                    Save = False
                End If
            End If
        Else
            If Not mobjEncounter.AddTextNote("Provider Telecon Note", ProviderNote, lNoteID) Then
                Save = False
            End If
        End If
        mbProviderNote_Changed = False
    End If
    
    
    
'    Moved Save to above
'    If Not mobjEncounter.Save(True) Then
'        Save = False
'    End If
    'MAG 05/17/01 SCR#12792
ExitFunction:
    Screen.MousePointer = lMouse

    
End Function

Public Sub CloseTelecon()
    If Not mobjDxModule Is Nothing Then
        Call mobjDxModule.CloseView(cwiENCOUNTER_SUMMARY_VIEW)
        Set mobjDxModule = Nothing
    End If
    If Not mobjIChild Is Nothing Then
        Call mobjIChild.SetParent(Nothing)
        Set mobjIChild = Nothing
    End If
    Set mobjAnPModule = Nothing
    Set mobjEncounter = Nothing
    Set mobjAPSection = Nothing
    Set mobjDispSection = Nothing
    Set mobjNote = Nothing
End Sub

Private Function IParent_AddMenuItem(ByVal iMenu As MenuNames, ByVal sMenuText As String, ByVal sMenuBMPFileName As String, ByVal sTip As String) As Long
End Function

Private Sub IParent_ClosedView()
End Sub

Private Property Get IParent_CurrentEncounter() As IEncDoc
    Set IParent_CurrentEncounter = mobjEncounter
End Property

Private Function IParent_EndSave(Optional ByVal nEventId As Long) As Boolean
    IParent_EndSave = mobjEncounter.LockingObject.EndSectionUpdate(eap, nEventId)
End Function

Private Property Get IParent_Hwnd() As Long
End Property

Private Function IParent_LockEventID(ByVal nEventId As Long) As Boolean
    IParent_LockEventID = mobjEncounter.LockingObject.LockSection(eap, nEventId, SectionLock)
End Function

Private Function IParent_RemoveMenuItem(ByVal lMenuID As Long) As Boolean
End Function

Private Property Get IParent_SectionObject() As Object
    Set IParent_SectionObject = mobjAPSection
End Property

Private Sub IParent_SetRTF(RTF As String, Optional Updated As Boolean)
    Call mobjAPSection.SetRTF(RTF)
End Sub

Private Sub IParent_SetStatus(ByVal eStatus As EncStatus)
    mobjAPSection.APStatus = eStatus
End Sub

Private Function IParent_StartSave(Optional ByVal nEventId As Long) As Boolean
    IParent_StartSave = mobjEncounter.LockingObject.StartSectionUpdate(eap, nEventId)
End Function

Private Function IParent_UnlockEventID(ByVal nEventId As Long) As Boolean
    IParent_UnlockEventID = mobjEncounter.LockingObject.UnlockSection(eap, nEventId)
End Function
