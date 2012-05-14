VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EncSignOps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Controller object for signature operations"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"EncProvider"
Attribute VB_Ext_KEY = "Member1" ,"SigBlock"
Option Explicit

Private Const Exe = "CHCSIIEncounterOps"

Public mobjEncounter As Encounter
Public mbAutoPrint As Boolean
Public mbAutoPOC As Boolean
Public msRTF As String

Private Sub Class_Initialize()
    mbAutoPrint = True
'-- COI Testing
#If coiDebug Then
     gCOI_DL.debugStart
#End If
End Sub

Public Function AppendNote(ByRef objEncounter As Encounter) As Boolean
    Dim objNote As TextNote
    Set mobjEncounter = objEncounter
    mobjEncounter.Refresh
    If Not (mobjEncounter.Status = complete _
         Or mobjEncounter.Status = NeedsCoSignature _
         Or mobjEncounter.Status = Updated) _
    Then
        If Not mobjEncounter.CWShared.IsAppMode(modeCHCSI_GUI) Then
            MsgBxARMd "You cannot append to an encounter whose status is " & EncStatusName(mobjEncounter.Status) & " .", vbInformation, "Encounter"
        Else
            MsgBxARMd "You cannot append to an encounter that was not documented in Provider GUI.", vbInformation, "Encounter" 'SCR-41215
        End If
        Exit Function
    End If
    Set objNote = New TextNote
    objNote.msOwnerNCID = mobjEncounter.Logon.UserNCID
    objNote.msOwnerName = mobjEncounter.Logon.UserName
    objNote.mdDTS = Now
    objNote.RelatedSection = eNotes  'SCR-40458
    Call frmTextNote.Initialize(mobjEncounter, objNote, True, Me)
    mobjEncounter.Comm.Message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_ENCOUNTER_APPLICATION
    frmTextNote.Show vbModal
    mobjEncounter.Comm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_ENCOUNTER_APPLICATION
    If objNote.mnDataID <> 0 Then
        AppendNote = True
        mobjEncounter.Status = Updated 'SCR-13083 Sends Alert
    End If
    Set objNote = Nothing
    Set mobjEncounter = Nothing
End Function

Public Function SignAndSaveNote(ByRef Note As TextNote) As Boolean
    
    Dim objSecParent As SectionParent
    Dim objNotes As TextNotes
    Dim objEncRTF As EncRTF
    Dim sMsg As String
    Dim dm As Object
    
    On Error GoTo ErrHandler
            
    If Not mobjEncounter.LockingObject.StartSectionUpdate(eEntireEncounter, 0) Then
        MsgBxARMd "Another user has broken your lock on this encounter.", vbInformation, "Encounter"
        Exit Function
    End If
    sMsg = "Get Section Parent"
    Set objSecParent = mobjEncounter.SectionParent(eNotes)
    sMsg = "Get Notes Section"
    Set objNotes = objSecParent.mobjSection
    sMsg = "Save Note"
    If objNotes.Save(Note) Then
        If mobjEncounter.Status = complete _
        Or mobjEncounter.Status = Updated Then
           mobjEncounter.Status = Updated
        End If
        Set objEncRTF = New EncRTF
        msRTF = BuildRTF(mobjEncounter)
        If msRTF = "" Then GoTo CleanUp
        sMsg = "Set Enc RTF data"
        objEncRTF.mdDTS = Now
        objEncRTF.msRTF = msRTF
        objEncRTF.msUserNCID = mobjEncounter.Logon.UserNCID 'SCR 128313 mobjEncounter.EncRTFs(1).msUserNCID
        sMsg = "Add Enc RTF"
        Call mobjEncounter.AddEncRTF(objEncRTF)
        Set objEncRTF = Nothing
       'Moved here so print out will occur quicker
       If mbAutoPrint Then
          sMsg = "Print Encounter"
          Call mobjEncounter.PrintEncounter(False, mbAutoPOC, , False) 'SCR#49496
       End If

'       If mbAutoPOC Then
'          sMsg = "Print POC"
'          'mobjEncounter.PrintPlanOfCare False
'       End If

        sMsg = "Update Appt Encounter Status"
        mobjEncounter.UpdateApptEncounterStatus 'SCR-39893

        sMsg = "Save Encounter"
        SignAndSaveNote = mobjEncounter.Save
        
        'SF 8-18-06 104010
        If SignAndSaveNote Then
            'Begin:SCR #39268; Developer Mag 8/4/2003 11:00 AM
            If mobjEncounter.CWShared.IsAppMode(modeTheater) Then
               'Commented the line below out - 61779
               'If mobjEncounter.EncounterType <> InpatientNote Then 'SCR#47895
                   Set dm = mobjEncounter.Comm.InitializeOLEServer("1999988")  'NCID_DATA_MANAGER
                   Call dm.Send2TMIP(mobjEncounter.Status)
                   Screen.MousePointer = vbHourglass
               'End If
            End If
            'End: SCR #39268
        End If
                
        On Error Resume Next
'       If mbAutoPrint Then
'          Call mobjEncounter.PrintEncounter(False)
'       End If
'
'       If mbAutoPOC Then
'          mobjEncounter.PrintPlanOfCare False
'       End If
    
    End If
    GoTo CleanUp
ErrHandler:
    mobjEncounter.CWShared.ShowVBError Err.Number, sMsg & " - " & Err.Description, "EncSignOps.SignAndSaveNote", "EncounterOps", vbCritical
CleanUp:
    Set objEncRTF = Nothing
    Set objSecParent = Nothing
    Set objNotes = Nothing
    Exit Function
    Resume
End Function

Public Function SignEncounter(ByRef objEncounter As Encounter) As Boolean
    
    On Error GoTo ErrHandler
    Dim lMousePointer As Long
    
    'SF 4-12-05
    lMousePointer = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    
    Set mobjEncounter = objEncounter  'bmstop
    mobjEncounter.Refresh
    
    'Now called in Encounter.SignEncounter function SCR-32981
    'If Not mobjEncounter.LockingObject.IsSOOpen Then
    
    'HWCC 1/30/02 -
    'SCRs 17424(also reference SCR 17422) and SCR 17133 requires
    ' Dental Parent to launch 2 pop up modal forms before provider
    '  sign the encounter. Per Dental SA Erin O'Donohue, those pop
    '   up modal forms are to be display before the
    '    "Encounter is not Complete" form. The code below call Dental Parent
    '     so Dental Parent can pop up those forms required by the SCRs
    If mobjEncounter.EncounterType = Dental Then
        On Error Resume Next
        Call mobjEncounter.SectionParent(eDental).mobjSection.PreEncounterSign
        If Err Then
            goARM.StopFarTx gsARM_Tx_SignEncounter, enmARMTxStatusFailed   'SCR 16126 MSolano 1/31/2002
            MsgBxARMd "Error notifying Dental module of encounter about to be signed (" & Err.Description & ").", vbExclamation, "Encounter"
            Err.Clear
        End If
        On Error GoTo ErrHandler
    End If
      
    If OKtoSign() Then
         
        '--- SCR 22247 flag ARM to time sign from Enc only (not children)  MSolano 4/24/02
        If mobjEncounter.SignByEncParent Then
            goARM.StartFarTx gsARM_Tx_SignEncounter '--- SCR 16126 MSolano 8/23/01
            mobjEncounter.SignByEncParent = False
            '--- SCR-ARM-25468 RWalker 3/07/03
            goARM.StartCmplxTx gsARM_Tx_CorrUT_SignEncounter, , , , , , , , eAnyFamily
        End If

        ' temporarily set the signing time to now, so appended notes will behave correctly
        If mobjEncounter.WhenSigned = 0 Then
            mobjEncounter.WhenSigned = Now
        End If
        msRTF = BuildRTF(mobjEncounter)
        ' Use cached current RTF 'SCR-24498
        'msRTF = mobjEncounter.CurrentRTF
        If msRTF = "" Then
            '--- SCR-ARM-25468 RWalker 3/07/03
            goARM.StopCmplxTx gsARM_Tx_CorrUT_SignEncounter, enmARMTxStatusFailed, enmARMStop_MeOnly
            GoTo CleanUp
        End If
        
        SignEncounter = frmSign.Display(Me)
        
        '--- SCR-ARM-25468 RWalker 3/07/03
        If SignEncounter Then
            goARM.StopCmplxTx gsARM_Tx_CorrUT_SignEncounter, enmARMTxStatusOk, enmARMStop_MeOnly
        Else
            goARM.StopCmplxTx gsARM_Tx_CorrUT_SignEncounter, enmARMTxStatusFailed, enmARMStop_MeOnly
        End If
    Else
        '--- SCR-ARM-32970/42483 RWalker 10/29/03
        goARM.StopCmplxTx gsARM_Tx_CorrUT_SignEncounter, enmARMTxStatusFailed, enmARMStop_MeOnly
    End If
    'End If
    GoTo CleanUp

ErrHandler:
    Call mobjEncounter.CWShared.ShowVBError(Err.Number, Err.Description, "EncSignOps.SignEncounter ", "EncounterOps", vbCritical)

CleanUp:
    Set mobjEncounter = Nothing
    
    'SF 4-12-05
    Screen.MousePointer = lMousePointer
    
End Function

Public Function SignAndSave(Optional ByRef objEncounter As Encounter) As Boolean
    Dim objEncRTF As EncRTF
    Dim dDTS As Date
    Dim objAMS As AlertsClient
    Dim dm As Object
    On Error GoTo ErrHandler

    If Not objEncounter Is Nothing Then
        If objEncounter.CWShared.IsAppMode(modeCHCSI_GUI) Then 'SCR 36493 Then
            Set mobjEncounter = objEncounter
            mbAutoPrint = False
        End If
    End If
'-- COI Testing
#If coiDebug Then
     Dim DR As DLog.DebugRec
     gCOI_DL.DebugLog "COI43", "", "", CStr(mobjEncounter.Logon.UserNCID), DR
#End If

    dDTS = Now 'SCR-26190

    With mobjEncounter
        If .Status = NeedsCoSignature Then
            Set .Cosigner = New EncProvider
            .Cosigner.NCID = .Logon.UserNCID
            .Cosigner.FullName = .Logon.UserName
            .WhenCoSigned = dDTS
            Set .WhoCoSigned = New EncProvider
            .WhoCoSigned.NCID = .Logon.UserNCID
            .WhoCoSigned.FullName = .Logon.UserName
            .WhoCoSigned.SigLine2 = .EncProperties.SigBlockLine2
            .WhoCoSigned.SigLine3 = .EncProperties.SigBlockLine3
            
            If Len(.Comm.CmdLineSwitch("SENSITIVE")) Then 'SCR-35365, 47577
                .Sensitivity = frmSign.cboDefaultRecordType.ItemData(frmSign.cboDefaultRecordType.ListIndex)
            Else
                If frmSign.chkSensitive.value Then
                    .Sensitivity = esrStandard
                End If
            End If
        Else
            Set .WhoCoSigned = Nothing
            .WhenCoSigned = 0
            .WhenSigned = dDTS
            Set .WhoSigned = New EncProvider
            .WhoSigned.NCID = .Logon.UserNCID
            .WhoSigned.FullName = .Logon.UserName
            .WhoSigned.SigLine2 = .EncProperties.SigBlockLine2
            .WhoSigned.SigLine3 = .EncProperties.SigBlockLine3
            '.Sensitivity = CBool(frmSign.chkSensitive.value)
            If Len(.Comm.CmdLineSwitch("SENSITIVE")) Then 'SCR-35365
                .Sensitivity = frmSign.cboDefaultRecordType.ItemData(frmSign.cboDefaultRecordType.ListIndex)
            Else
                If frmSign.chkSensitive.value Then
                    .Sensitivity = esrStandard
                Else
                    .Sensitivity = esrNone
                End If
            End If
                
        End If

        'Moved here b/c need RTF when setting status to complete
        'need to build rtf before setting whensigned affect appended notes
        msRTF = BuildRTF(mobjEncounter, mbAutoPOC)
        If msRTF = "" Then GoTo CleanUp
        Set objEncRTF = New EncRTF
        objEncRTF.mdDTS = dDTS
        objEncRTF.msRTF = msRTF
        objEncRTF.msUserNCID = mobjEncounter.Logon.UserNCID
        Call mobjEncounter.AddEncRTF(objEncRTF)
        Set objEncRTF = Nothing

        'Set new encounter status after RTF is built SCR-26526
        If .Status = NeedsCoSignature Then
            If .WasUpdated Then
                .Status = Updated
            Else
                .Status = complete
            End If
            
            'Delete Need CoSign Alert SCR#48758
            If Not .Cosigner Is Nothing Then
                Set objAMS = New AlertsClient
                objAMS.DeleteAlert .Cosigner.NCID, "{7EFF358F-8A26-11d3-8304-0050DA0C5BB9}"
            End If
        Else
            If .Cosigner Is Nothing Then
                If .Status = Updating Then
                    .Status = Updated
                Else
                    .Status = complete
                End If
            Else
                .Status = NeedsCoSignature
            End If
        End If
        ' Moved here so print out will occur quicker
        If mbAutoPrint Then
            'SCR-21246 hide sign encounter window so dental encounter can print properly
            If .EncounterType = Dental Then
                frmSign.Hide
            End If
            'Do not refresh data since it will reset data form DB SCR#48882
            Call .PrintEncounter(False, mbAutoPOC, , False)
        End If
'        If mbAutoPOC Then
'           '.PrintPlanOfCare False
'        End If

        'If an appointment status change is made before the encounter is signed,
        'there is no impact to SADR.  However, if an appointment status change
        'is made after the encounter is signed, the change will not be written
        'back to CHCS, and will  not be included in the SADR reports. per Sherry Wang
        'Update Appt Encounter Status before saving encounter
        .UpdateApptEncounterStatus 'SCR-39893

        SignAndSave = .Save
     
On Error Resume Next
        If .Status = NeedsCoSignature And Not .Cosigner Is Nothing Then
            Set objAMS = New AlertsClient
            Call objAMS.Send(.Cosigner.NCID, "{7EFF358F-8A26-11d3-8304-0050DA0C5BB9}", _
                "You have encounters needing co-signature.")
                '"The encounter for " & .Patient.Name & " requires your co-signature.")
                
            If Err Then
                'SCR 45291 RWalker 11/25/03
                goARM.StartCmplxTx gsARM_Tx_CorrUT_NoAlertToCosigner, , , , , , , , eAnyFamily
                MsgBxARMd "Could not send alert to Cosigner (" & Err.Description & ").", vbExclamation, "Encounter"
                Err.Clear
                '--- SCR-ARM-45291 RWalker 11/25/2003
                goARM.StopCmplxTx gsARM_Tx_CorrUT_NoAlertToCosigner, enmARMTxStatusOk, enmARMStop_MeOnly
            End If
        End If
        If .Status = complete Or .Status = Updated Or .Status = NeedsCoSignature Then
            .Comm.SetStatusText "Updating Problem List ..."
            'PERF - Check if encounters has any Diagnoses to convert to Problems SCR-24498
            'SCR 50793  Phil Crowder 3/12/04  Don't do for PGUI it is done by the web service
            If mobjEncounter.PrimaryDxSnoID <> "" Then
                Call .SectionParent(eap).mobjSection.NotifyComplete
                If Err Then
                    'SCR 45291 RWalker 11/25/03
                    goARM.StartCmplxTx gsARM_Tx_CorrUT_NoAPNotification, , , , , , , , eAnyFamily
                    MsgBxARMd "Error notifying AP module of encounter completion (" & Err.Description & ").", vbExclamation, "Encounter"
                    Err.Clear
                    '--- SCR-ARM-45291 RWalker 11/25/2003
                    goARM.StopCmplxTx gsARM_Tx_CorrUT_NoAPNotification, enmARMTxStatusOk, enmARMStop_MeOnly
                End If
            End If
'OA
'            'Only call if encounter has family history
'            Call .SectionParent(eSO).mobjSection.NotifyComplete
'            If Err Then
'                'SCR 45291 RWalker 11/25/03
'                goARM.StartCmplxTx gsARM_Tx_CorrUT_NoSONotification, , , , , , , , eAnyFamily
'                MsgBxARMd "Error notifying SO module of encounter completion (" & Err.Description & ").", vbExclamation, "Encounter"
'                Err.Clear
'                '--- SCR-ARM-45291 RWalker 11/25/2003
'                goARM.StopCmplxTx gsARM_Tx_CorrUT_NoSONotification, enmARMTxStatusOk, enmARMStop_MeOnly
'            End If
            'Also call when updated since user may delete profile SCR-36337
            If .DispositionNCID = NCID_ReleaseWDutyLimitations Or .Status = Updated Then
                'Need to save profile data
                .Comm.SetStatusText "Updating Profile Data ..."
                Call .SectionParent(eDisp).mobjSection.NotifyComplete
                If Err Then
                    'SCR 45291 RWalker 11/25/03
                    goARM.StartCmplxTx gsARM_Tx_CorrUT_NoDispNotification, , , , , , , , eAnyFamily
                    MsgBxARMd "Error notifying Disposition module of encounter completion (" & Err.Description & ").", vbExclamation, "Encounter"
                    Err.Clear
                    '--- SCR-ARM-45291 RWalker 11/25/2003
                    goARM.StopCmplxTx gsARM_Tx_CorrUT_NoDispNotification, enmARMTxStatusOk, enmARMStop_MeOnly
                End If
            End If
            .Comm.SetStatusText "Updating Dental ..."
            If .EncounterType = Dental Then
                Call .SectionParent(eDental).mobjSection.NotifyComplete
                If Err Then
                    'SCR 45291 RWalker 11/25/03
                    goARM.StartCmplxTx gsARM_Tx_CorrUT_NoDentNotification, , , , , , , , eAnyFamily
                    MsgBxARMd "Error notifying Dental module of encounter completion (" & Err.Description & ").", vbExclamation, "Encounter"
                    Err.Clear
                    '--- SCR-ARM-45291 RWalker 11/25/2003
                    goARM.StopCmplxTx gsARM_Tx_CorrUT_NoDentNotification, enmARMTxStatusOk, enmARMStop_MeOnly
                End If
            End If
        End If
    
        If .Status = Updated Then
            ' tell Previous Encounters to cwiREFRESH_DATA SCR-24835
            If mobjEncounter.Comm.Message(cwiOLE_SERVER_RUNNING, NCID_PREVIOUS_ENCOUNTERS_APPLICATION, NCID_CW_APPLICATION, NCID_ENCOUNTER_APPLICATION) Then
                '- tell it to update the single appt
                mobjEncounter.Comm.Message cwiREFRESH_DATA, "", NCID_PREVIOUS_ENCOUNTERS_APPLICATION, NCID_ENCOUNTER_APPLICATION
            End If
        End If
'        Moved to before save so print out occur quicker
'        If mbAutoPrint Then
'           Call .PrintEncounter(False)
'        End If
'        If mbAutoPOC Then
'           .PrintPlanOfCare False
'        End If

         'Begin:SCR #39268; Developer Mag 8/4/2003 11:00 AM
         If mobjEncounter.CWShared.IsAppMode(modeTheater) Then
            'Commented the line below out - 61779
            'If mobjEncounter.EncounterType <> InpatientNote Then 'SCR#47895
                Set dm = .Comm.InitializeOLEServer("1999988")  'NCID_DATA_MANAGER
                Call dm.Send2TMIP(.Status)
                Screen.MousePointer = vbHourglass
            'End If
         End If
         'End: SCR #39268
         
    End With
    GoTo CleanUp
ErrHandler:
    'Resume Next
    mobjEncounter.CWShared.ShowVBError Err.Number, Err.Description, "EncSignOps.SignAndSave", "EncounterOps", vbCritical
CleanUp:
    If SignAndSave Then
        mobjEncounter.Comm.SetStatusText "Encounter signed and saved."
    Else
        mobjEncounter.Comm.SetStatusText ""
    End If
    Set objEncRTF = Nothing
    Set objAMS = Nothing
End Function

'----------------------------------------------------------------------
'--- PROCEDURE: OKtoSign()
'---
'--- CHANGE HISTORY:
'--- * SCR 16126 MSolano 1/31/2002
'---   Call stack:
'---   CHCSIIEncounterOps.Encounter.SignEncounter()
'---   ..CHCSIIEncounterOps.EncounterSignOps.SignEncounter()
'---   ..CHCSIIEncounterOps.EncounterSignOps.OkToSign()
'---     (may prompt user w/ MsgBxARMd or frmIncomplete)
'---   ......CHCSIIEncounterOps.frmSign.Display()
'---   ........CHCSIIEncounterOps.frmSign.Form_Load()
'---   ..........CHCSIIEncounterOps.frmSign.Form_Resize()
'---   The first call starts a far transaction [ARM.StartFarTx()].
'---   The last call resolves the far transaction [ARM.StartFarTx()] w/ ok status.
'---   Any intervening call can resolve w/ transaction status as failed (business rules) or aborted (run time error).
'----------------------------------------------------------------------
Private Function OKtoSign() As Boolean
    Dim sIncMsg As String
    Dim objNotes As TextNotes
    Dim i As Integer
    Dim objNote As TextNote
    Dim objshared As New CWShared
        
    '--- SCR-ARM-25468/27364 RWalker 3/07/03
    Dim eARMStatus As enmARMTxStatus
    
    On Error GoTo ErrHandler
    
    Set objshared = New CWShared
    
    '--- SCR-ARM-27364 RWalker 3/07/03
    goARM.StartCmplxTx gsARM_Tx_CorrUT_Incomplete, , , , , , , , eAnyFamily
    
    With mobjEncounter
        If .Status = PatientLeftWithoutBeingSeen Then
            MsgBxARMd "You cannot sign an encounter whose status is LWOBS.", vbInformation, "Encounter"
            '--- SCR-ARM-25468/27364 RWalker 3/07/03
            eARMStatus = enmARMTxStatusFailed
            GoTo CleanUp
        End If
        If .Status = complete Or .Status = Updated Then
            MsgBxARMd "This encounter is complete and does not require signing.", vbInformation, "Encounter"
            '--- SCR-ARM-25468/27364 RWalker 3/07/03
            eARMStatus = enmARMTxStatusFailed
            GoTo CleanUp
        End If
        
        'Inpatient Note Rqmt#42960
        'No sections need to be completed for Inpatient Notes except for the Discharge Summary note (disposition section).
        'Allow anyone with access to Inpatient Notes to sign and Inpatient Note Encounter SCR#42998
        '105696 Added the "And mobjEncounter.InptNoteTypeNCID <> NCID_Discharge_Note" part per 105696 and 42982
        If mobjEncounter.EncounterType = InpatientNote And mobjEncounter.InptNoteTypeNCID <> NCID_Discharge_Note Then
        'If mobjEncounter.EncounterType = InpatientNote Then
            OKtoSign = True
            '--- SCR-ARM-25468/27364 RWalker 3/07/03
            eARMStatus = enmARMTxStatusOk
            GoTo CleanUp
        End If
        
        If .Status = NeedsCoSignature Then
        
            If Not OKtoSign_NeedsCosignature Then
                 '--- SCR-ARM-25468/27364 RWalker 3/07/03
                 eARMStatus = enmARMTxStatusFailed
                GoTo CleanUp
            End If
            
            'Since cosigner may make changes that would make section incomplete SCR#44712
            GoTo CheckSectionStatus
        End If
        
        'Made into separate if block SCR-23893
        If Not .UserIsPrimaryProvider Then
        
            'Give user option to transfer ownership if user has sign privilege, SCR-23893
            'Only enalbe for Theater and ITT SCR-26431
            If .User.HasPrivilegeEx(Priv_Current_Encounter, Priv_Sign) And mobjEncounter.EncounterType <> InpatientNote Then
            
                If Not TransferEncounterOwnership Then
                    '--- SCR-ARM-25468/27364 RWalker 3/07/03
                    eARMStatus = enmARMTxStatusFailed
                    GoTo CleanUp
                End If
                
            Else
                MsgBxARMd "You do not have the authorization to sign this encounter.", vbInformation, "Encounter"
                '--- SCR-ARM-25468/27364 RWalker 3/07/03
                eARMStatus = enmARMTxStatusFailed
                GoTo CleanUp
            End If
            
        End If
        
CheckSectionStatus:
    
        'Check to see if AP info was changed after Disposition data was entered SCR#52381
        If Not OKtoSign_APAfterDispCheck Then
            eARMStatus = enmARMTxStatusFailed
            GoTo CleanUp
        End If

        'For Dental AP not required SCR-37821
        'For an Inpatient Discharge Summary note, AP is not required. 105696
        If .EncounterType <> Dental And .APStatus <> complete And mobjEncounter.InptNoteTypeNCID <> NCID_Discharge_Note Then
            sIncMsg = sIncMsg & " - A/P Section" & vbCrLf
        End If
        If .DispStatus <> complete Then
            sIncMsg = sIncMsg & " - Disposition Section" & vbCrLf
        End If
        If Len(.CWShared.CmdLineSwitches("HIPAA837")) > 0 Then
            sIncMsg = sIncMsg & .PregnancyData.GetSigningWarningText()
        End If
        Set objNotes = .SectionParent(eNotes).mobjSection
        For i = 1 To objNotes.Count
            Set objNote = objNotes.TextNote(i)
            If Not objNote.mbIsComplete Then
                sIncMsg = sIncMsg & " - Note written by " & objNote.msOwnerName _
                    & " @ " & FormatDTG(objNote.mdDTS) & vbCrLf
            End If
            Set objNote = Nothing
        Next i
        
        If sIncMsg = "" Then
            OKtoSign = True
            '--- SCR-ARM-25468/27364 RWalker 3/07/03
            eARMStatus = enmARMTxStatusOk
            GoTo CleanUp
        End If
        
        '<<< SCR 16126 MSolano 1/31/2002
        '--- Failed preliminary (automatic) business rules for signing an encounter
        '--- Need to prompt user ...  user response time shouldn't be measured by ARM.
        goARM.StopFarTx gsARM_Tx_SignEncounter, enmARMTxStatusFailed
        '>>> SCR 16126
        
        'If only AP is complete check if Draft Constults is reason SCR-17942
        'Give user option of deleting draft consults
        'NOTE: For now this only works in CDR mode for the Dec 2002 Maintenance Release
        'Does not apply to Dental Encounters
        If .EncounterType <> Dental And .APStatus = Incomplete And .DispStatus = complete Then
            'check if AP only has Draft Consults as reason for incomplete
            Dim sTemp As String
            Dim sbuf() As String
            
            sTemp = mobjEncounter.APStatusComment
                        
            sbuf = Split(sTemp, ";")
            
            'If only one check if it is DRAFT_CONSULTS
            If UBound(sbuf) = 1 Then
                If sbuf(0) = "DRAFT_CONSULTS" Then
                    If MsgBxARMd("An unsubmitted consult order has been found.  If you choose to continue the order will be deleted.", vbOKCancel + vbExclamation, "Encounter") = vbOK Then
                        'Delete consult
                        Call DeleteDraftConsults
                        'Reset AP Statuses
                        mobjEncounter.APStatusComment = vbNullString
                        mobjEncounter.APStatus = complete
                        'Save AP Status
                        mobjEncounter.SectionParent(eap).mobjSection.Save
                        'Save AP Status Comment
                        mobjEncounter.Save
                        OKtoSign = True
                        '--- SCR-ARM-25468/27364 RWalker 3/07/03
                        eARMStatus = enmARMTxStatusOk
                        GoTo CleanUp
                    End If
                End If
            End If
        End If
        
        Load frmIncomplete
        
        frmIncomplete.txtMsg.Text = "This encounter has the following incomplete items:" & vbCrLf & vbCrLf _
            & sIncMsg & vbCrLf
            '& " Please select one of the following options:" 'SCR-42602
         '- if this is not Dental, and A&P is incomplete, only nocount and close
        If .EncounterType <> Dental _
        And (.APStatus <> complete Or .DispStatus <> complete) Then
            frmIncomplete.cmdComplete.Enabled = False
            frmIncomplete.cmdNoCount.Enabled = False 'SCR-27876 Force users to create complete encounter
        End If
        
        If .EncounterType = Dental Then
            'Show and enable no count button for Dental SCR-36904
            frmIncomplete.Height = 3165
            frmIncomplete.cmdNoCount.Visible = True
            frmIncomplete.cmdNoCount.Enabled = True
            frmIncomplete.lblNoCount.Visible = True
        Else
            frmIncomplete.Height = 2535
            frmIncomplete.cmdNoCount.Visible = False
            frmIncomplete.cmdNoCount.Enabled = False
            frmIncomplete.lblNoCount.Visible = False
        End If
        
        .Comm.Message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_ENCOUNTER_APPLICATION
        
        frmIncomplete.Show vbModal
        .Comm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_ENCOUNTER_APPLICATION
        
        If frmIncomplete.mbNoCount Then
            .Class = ec_NoCount
            .WorkLoad = "N"
            OKtoSign = True
            '--- SCR-ARM-27364 RWalker 3/07/03
            eARMStatus = enmARMTxStatusOk
        ElseIf frmIncomplete.mbComplete Then
            OKtoSign = True
            '--- SCR-ARM-27364 RWalker 3/07/03
            eARMStatus = enmARMTxStatusOk
        Else
            '--- SCR-ARM-27364 RWalker 3/07/03
            eARMStatus = enmARMTxStatusFailed
       End If
        
    End With
    
    GoTo CleanUp
    
ErrHandler:
    mobjEncounter.CWShared.ShowVBError Err.Number, Err.Description, "EncSignOps.OKtoSign", "EncounterOps", vbCritical
    '--- SCR-ARM-25468 RWalker 3/07/03
    eARMStatus = enmARMTxStatusAborted
    
CleanUp:
    '--- SCR-ARM-25468 RWalker 3/07/03
    goARM.StopCmplxTx gsARM_Tx_CorrUT_Incomplete, eARMStatus, enmARMStop_MeOnly
    
    Set objNote = Nothing
    Set objshared = Nothing
End Function

Private Function OKtoSign_NeedsCosignature() As Boolean

    On Error GoTo ErrHandler

    OKtoSign_NeedsCosignature = True
    
    With mobjEncounter
        If Not .UserIsCosigner Then 'SCR-13412
            If .UserIsPrimaryProvider Then
                MsgBxARMd "As the primary provider, you are not authorized to co-sign this encounter.", vbExclamation, "Encounter"
                OKtoSign_NeedsCosignature = False
            ElseIf .User.HasPrivilegeEx(Priv_Current_Encounter, Priv_Sign) Then
                '<<< SCR 16126 MSolano 1/31/2002
                goARM.StopFarTx gsARM_Tx_SignEncounter, enmARMTxStatusFailed
                '>>> SCR 16126
                If MsgBxARMd("Do you want to co-sign this encounter, even though it has been assigned to " _
                    & .Cosigner.FullName & "?", vbQuestion + vbYesNo, "Cosign Instead of Another Provider") = vbNo Then
                    OKtoSign_NeedsCosignature = False
                End If
            Else
                MsgBxARMd "You do not have the authorization to sign encounters.", vbInformation, "Encounter"
                OKtoSign_NeedsCosignature = False
            End If
        End If
    End With
            
    Exit Function
    
ErrHandler:
    mobjEncounter.CWShared.ShowVBError Err.Number, Err.Description, "EncSignOps.OKtoSign_NeedsCosignature", "EncounterOps", vbCritical
End Function
Private Sub Class_Terminate()
'-- COI Testin
#If coiDebug Then
     gCOI_DL.debugStop
     Set gCOI_DL = Nothing
#End If
End Sub
' Delete Draft Consults for current Encounter SCR-17942
' Need to use ConsultOrderOps to delete draft consult order
Private Sub DeleteDraftConsults()

    Dim objConsultOps As Object
    Dim objSQL As ICHCSII_SQL.ISqlOps
    
    Set objConsultOps = CreateObject("CHCSII_Consult.ConsultOrderOps")
    Set objSQL = gobjCHCSIIConn.CHCSII_SQLOPS(Auto)  '''<SCR 36493
    
    objSQL.Execute _
        ("SELECT CONSULT_ID FROM CONSULT_ORDER WHERE ENC_NUM = " & mobjEncounter.EncounterID & " AND ORDER_STATUS = 0")
    
    While Not objSQL.EOF
        Call objConsultOps.DeleteConsultOrder(objSQL!CONSULT_ID)
        objSQL.MoveNext
    Wend

    Set objSQL = Nothing
    Set objConsultOps = Nothing
    
End Sub

'Check to see if AP info was changed after Disposition data was entered SCR-52152
Private Function OKtoSign_APAfterDispCheck() As Boolean

    On Error GoTo ErrHandler
                                                    'SCR 105696 Need to do this check for Discharge Summary notes.
    If mobjEncounter.EncounterType <> InpatientNote Or mobjEncounter.InptNoteTypeNCID = NCID_Discharge_Note Then
        If mobjEncounter.APAfterDisp Then
            'SCR 45291 RWalker 11/25/03
            goARM.StartCmplxTx gsARM_Tx_CorrUT_ReturnToDisp, , , , , , , , eAnyFamily
            'Display warning SCR-35067
            
            If Not mobjEncounter.CWShared.IsAppMode(modeCHCSI_GUI) Then
                MsgBxARMd "RETURN TO DISPOSITION MODULE" & vbCrLf & vbCrLf & "Your A/P information has been updated.  If you changed diagnoses or reordered them, you will need to return to the Disposition module so that this information will be updated with the new values.", vbOKOnly + vbCritical, "EncounterOps"
                'SCR 104743 - LCI - 09/01/2006: Do not open sign form, instead take user to Disposition module
                goARM.StopCmplxTx gsARM_Tx_CorrUT_ReturnToDisp, enmARMTxStatusFailed, enmARMStop_MeOnly
                goARM.StopCmplxTx gsARM_Tx_CorrUT_SignFamily, enmARMTxStatusFailed, enmARMStop_MeOnly
                mobjEncounter.OpenSection eDisp
                Screen.MousePointer = vbDefault
                Exit Function
            Else
                If MsgBxARMd("RETURN TO DISPOSITION MODULE" & vbCrLf & vbCrLf & "You have modified A/P documentation after selecting an E&M code. If you added, deleted, or reordered diagnoses, you should return to the Disposition screen to be sure that the E&M Code/Associated Diagnoses are correct. Do you wish to return to Disposition before signing?", vbYesNo + vbCritical, "Encounter") = vbYes Then
                  '--- SCR-ARM-45291 RWalker 11/25/2003
                  goARM.StopCmplxTx gsARM_Tx_CorrUT_ReturnToDisp, enmARMTxStatusFailed, enmARMStop_MeOnly
                  '--- SCR-ARM-32970/42483 RWalker 10/29/03 Must manually stop a correlated tx which was done in CleanUp
                  goARM.StopCmplxTx gsARM_Tx_CorrUT_SignFamily, enmARMTxStatusFailed, enmARMStop_MeOnly
                  mobjEncounter.OpenSection eDisp
                  Screen.MousePointer = vbDefault
                  Exit Function
                  ' Do exit function instead of clean up b/c clean up call unlock encounters which
                  ' undos locksection done in OpenSection SCR-40960
                  'GoTo CleanUp
                End If
            End If
            '--- SCR-ARM-45291 RWalker 11/25/2003
            goARM.StopCmplxTx gsARM_Tx_CorrUT_ReturnToDisp, enmARMTxStatusFailed, enmARMStop_MeOnly
        End If
    End If
    
    OKtoSign_APAfterDispCheck = True
    
    Exit Function
    
ErrHandler:
    
    mobjEncounter.CWShared.ShowVBError Err.Number, Err.Description, "Encounter.OKtoSign_APAfterDispCheck", Exe, vbCritical
    
End Function

Private Function TransferEncounterOwnership() As Boolean

    Dim AppointmentOps  As Object 'cannot reference appt
    Dim oEncProvider    As EncProvider
    
    On Error GoTo ErrHandler
    
    With mobjEncounter
    
        If MsgBxARMd("Encounter currently owned by " & .PrimaryProvider.FullName & ".  Do you want to take ownership (credit) and sign?", vbYesNo + vbDefaultButton1 + vbExclamation, "Encounter") = vbYes Then
            
            'Need to remove current provider if on Provider list as Addl Provider
            On Error Resume Next
            .Providers.Remove (.Logon.UserNCID)
            On Error GoTo ErrHandler
            
            'Change Primary Provider in Encounter
            Set oEncProvider = .Providers(1)
            oEncProvider.FullName = .Logon.UserName
            oEncProvider.NCID = .Logon.UserNCID
            
            If Not .Save(True) Then
                MsgBxARMd "Could not update encounter database." & vbCrLf & Err.Description, vbCritical, "EncounterOps"
                GoTo CleanUp
            End If
            
            'Change Primary Provider in Appointments
            Set AppointmentOps = CreateObject("CHCSII_AppointmentClient.AppointmentOps")
            AppointmentOps.Initialize NCID_ENCOUNTER_APPLICATION, .Comm, .Logon, .Config, .Patient, .User
            If Not AppointmentOps.UpdateAppointmentProvider(.AppointmentID, .Logon.UserNCID) Then
                MsgBxARMd "Could not update appointment provider." & vbCrLf & Err.Description, vbCritical, "EncounterOps"
                GoTo CleanUp
            End If
            
            'Refresh Appointment data SCR-30773, add appt id SCR#49319
            .Comm.Message cwiREFRESH_DATA, "APPTID|" & .AppointmentID, NCID_CLINIC_SCHEDULE_APPLICATION, NCID_ENCOUNTER_APPLICATION
            
            'Update provider name in encouter document
            .RefreshDisplay
            
         Else
            'vbNo
            GoTo CleanUp
        End If
                
    End With
    
    TransferEncounterOwnership = True
    
CleanUp:

    Set AppointmentOps = Nothing
    Set oEncProvider = Nothing
    
    Exit Function
    
ErrHandler:
    
    mobjEncounter.CWShared.ShowVBError Err.Number, Err.Description, "Encounter.TransferEncounterOwnership", Exe, vbCritical
    
End Function


Private Function CheckOKtoSignMsg(ByVal sIncMsg As String) As Boolean

    On Error GoTo ErrHandler
    
    If InStr(1, sIncMsg, "does not conform to ICD_DIAGNOSIS_PRIORITYs", vbTextCompare) Then
        MsgBxARMd "RETURN TO DISPOSITION MODULE" & vbCrLf & vbCrLf & "Changes that were made in A/P have altered the Diagnosis/E&M Code associations in Disposition. You must return to Disposition to validate the current associations.", vbOKOnly + vbCritical, "Encounter"
        '--- SCR-ARM-45291 RWalker 11/25/2003
        goARM.StopCmplxTx gsARM_Tx_CorrUT_ReturnToDisp, enmARMTxStatusFailed, enmARMStop_MeOnly
        '--- SCR-ARM-32970/42483 RWalker 10/29/03 Must manually stop a correlated tx which was done in CleanUp
        goARM.StopCmplxTx gsARM_Tx_CorrUT_SignFamily, enmARMTxStatusFailed, enmARMStop_MeOnly
        mobjEncounter.OpenSection eDisp
    Else
        CheckOKtoSignMsg = True
    End If
    
    Exit Function
    
ErrHandler:
    
    mobjEncounter.CWShared.ShowVBError Err.Number, Err.Description, "Encounter.TransferEncounterOwnership", Exe, vbCritical


End Function

