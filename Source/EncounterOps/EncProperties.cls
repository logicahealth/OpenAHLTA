VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EncProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"EncAutoCite"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mcolAutoCites As Collection   '- of AutoCiteProps as provided by enctr
Private mcolTemplateAutoCites As Collection   '- of AutoCiteProps as provided by enctr
Private msSigBlockLine2 As String
Private msSigBlockLine3 As String
Private msDefaultCosignerNCID As String
Private mbDidSave As Boolean

Private uEncounter As Encounter_Key_Information
Private mbTemplatePropertiesExist As Boolean

Private mobjUser As CHCSII_ClientSecurity.ICHCSII_User
Private mobjLogon As MMMHISLogon.Logon
Private mobjComm As Object
Private mobjShared As MMMHISShared.CWShared
Private mobjConfig As MMMHISConfiguration.Config
Private mbRetrievedProperties As Boolean
Private msAutoSaveInt As Long
Private mbInpatientMedsSelected As Boolean 'SCR#44005

Private miShowActiveMeds As Integer
Private miShowActiveLabs As Integer
Private miShowActiveRads As Integer

Private miDefaultSigning As Integer
Private miAutosaveSO As Integer

Private mlDefaultDiscussedWith As Long

Private mlDefaultRecordType As Long

Private mlIncludeCodes As Long  'SCR-38957

Private mobjEncPropDic As Dictionary

Private msAutociteProperties As String
Private msProvPrefProperties As String

Private msMsg As String
Private mbRetrieveFromOSH As Boolean

Private Const msEscChr As String = "\"
Private Const gsMainDelim As String = "|"
Private Const gsListDelim As String = "^"
Private Const gsRecordDelim As String = ","
Friend Property Get User() As CHCSII_ClientSecurity.ICHCSII_User: Set User = mobjUser:                  End Property
Friend Property Get Logon() As MMMHISLogon.Logon: Set Logon = mobjLogon:                                End Property
Friend Property Get Comm() As MMMHISComm.Comm: Set Comm = mobjComm:                                     End Property
Friend Property Let DidSave(value As Boolean):             mbDidSave = value:                           End Property 'SCR-36977
Public Property Let AutoSaveInt(RHV As Long):              msAutoSaveInt = RHV:                          End Property
Public Property Get AutoSaveInt() As Long:                 AutoSaveInt = msAutoSaveInt:                  End Property
Public Property Let DefaultCosignerNCID(RHV As String):    msDefaultCosignerNCID = RHV:                  End Property
Public Property Get DefaultCosignerNCID() As String:       DefaultCosignerNCID = msDefaultCosignerNCID:  End Property
Public Property Let SigBlockLine2(RHV As String):    msSigBlockLine2 = RHV: End Property
Public Property Let SigBlockLine3(RHV As String):    msSigBlockLine3 = RHV: End Property
Public Property Set AutoCites(ByVal vData As Collection):   Set mcolAutoCites = vData:                      End Property
Public Property Get AutoCites() As Collection
    If Not mbRetrievedProperties Or mcolTemplateAutoCites.Count > 0 Then
        If mobjShared Is Nothing Then Set mobjShared = New CWShared
        If mobjShared.IsAppMode(modeCHCSI_GUI) Then 'SCR 36493
            Set mcolTemplateAutoCites = New Collection
        Else
            'If template autocite exists need to refresh autocite properties everytime just case in Template was removed SCR-35162
            Call RetrieveProperties
            RetrieveTemplateProperties mcolTemplateAutoCites
        End If
        mbRetrievedProperties = True
    End If
    If mcolTemplateAutoCites.Count > 0 Then
        Set AutoCites = mcolTemplateAutoCites
    Else
        Set AutoCites = mcolAutoCites
    End If
End Property

Public Property Get SigBlockLine2() As String
    If msSigBlockLine3 = "" Then
        SigBlockLine2 = mobjLogon.RoleName
    Else
        SigBlockLine2 = msSigBlockLine2
    End If
End Property

Public Property Get SigBlockLine3() As String
    If msSigBlockLine3 = "" Then
        SigBlockLine3 = mobjLogon.Facilityname
    Else
        SigBlockLine3 = msSigBlockLine3
    End If
End Property

Public Property Get ShowActiveMeds() As Integer
    ShowActiveMeds = miShowActiveMeds
End Property

Public Property Let ShowActiveMeds(value As Integer)
    miShowActiveMeds = value
End Property

Public Property Get ShowActiveLabs() As Integer
    ShowActiveLabs = miShowActiveLabs
End Property
Public Property Let ShowActiveLabs(value As Integer)
    miShowActiveLabs = value
End Property

Public Property Get ShowActiveRads() As Integer
    ShowActiveRads = miShowActiveRads
End Property
Public Property Let ShowActiveRads(value As Integer)
    miShowActiveRads = value
End Property

Public Property Get DefaultDiscussedWith() As Long
    DefaultDiscussedWith = mlDefaultDiscussedWith
End Property
Public Property Let DefaultDiscussedWith(value As Long)
    mlDefaultDiscussedWith = value
End Property
Public Property Get DefaultRecordType() As Long
    DefaultRecordType = mlDefaultRecordType
End Property
Public Property Let DefaultRecordType(value As Long)
    mlDefaultRecordType = value
End Property
Public Property Get InpatientMedsSelected() As Boolean
    InpatientMedsSelected = mbInpatientMedsSelected
End Property
Public Property Let InpatientMedsSelected(value As Boolean)
    mbInpatientMedsSelected = value
End Property

Public Sub Initialize(robjUser As CHCSII_ClientSecurity.ICHCSII_User, robjLogonEnvironment As MMMHISLogon.Logon, robjComm As Object, robjShared As Object, robjConfig As Object)
    Set mobjUser = robjUser
    Set mobjLogon = robjLogonEnvironment
    Set mobjComm = robjComm
    Set mobjShared = robjShared
    Set mobjConfig = robjConfig
End Sub

Private Sub Class_Initialize()
   Set mcolAutoCites = New Collection
End Sub

Private Sub Class_Terminate()
    Set mobjUser = Nothing
    Set mobjComm = Nothing
    Set mobjLogon = Nothing
    Set mobjShared = Nothing
    Set mobjConfig = Nothing
End Sub

Public Function Display(pEncounterInfo As Encounter_Key_Information, Optional pTemplateData As String) As Boolean
    
    mbDidSave = False
    On Error GoTo ErrHandler
    
    Screen.MousePointer = vbHourglass 'SCR#50255
    
    uEncounter = pEncounterInfo
    
    If Len(mobjComm.CmdLineSwitch("GREEN")) Then
        frmProperties.chkExpiredMeds.Visible = True
    End If
    
    If Len(mobjComm.CmdLineSwitch("RNDS")) Then 'SCR#44005
          'Do not display expired meds SCR 44005
          'Bill Fote 2-Jan-2004 - This If is coming back False
        If uEncounter.Category Like APPT_TYPE_RNDS Then ''SCR#45422
            frmProperties.chkInpatientMeds.Visible = True
            frmProperties.chkExpiredMeds.Visible = False
            mbInpatientMedsSelected = True
        End If
    End If

    RetrieveProperties

    Set frmProperties.Properties = Me
    frmProperties.Display pTemplateData
    
    'Hide the questionnaires checkbox for theater mode
    'SCR 53595
    If mobjShared.AppMode = modeTheater Then
        frmProperties.chkQuestionnaires.Visible = False
    Else
        frmProperties.chkQuestionnaires.Visible = True
    End If
    
    If Len(mobjComm.CmdLineSwitch("CPG")) Then
       frmProperties.chkOverdueGoals.Visible = True
       frmProperties.chkDeviceReadings.Visible = True
       frmProperties.chkHandouts.Visible = True
    End If
    
    If Len(mobjComm.CmdLineSwitch("SENSITIVE")) Then
        frmProperties.cboDefaultRecordType.Visible = True 'SCR-35365
        frmProperties.lblDefaultRecordType.Visible = True
    End If
  
    If Len(mobjComm.CmdLineSwitch("OPTION")) Then
        frmProperties.chkIncludeCodes.Visible = True
    End If
    
    frmProperties.Top = (Screen.Height - frmProperties.Height) / 3
    frmProperties.Left = (Screen.Width - frmProperties.Width) / 2
  
    Screen.MousePointer = vbDefault
  
    mobjComm.Message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_ENCOUNTER_APPLICATION
    frmProperties.Show vbModal
    mobjComm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_ENCOUNTER_APPLICATION
    Display = mbDidSave
    Exit Function

ErrHandler:
    Call mobjShared.ShowVBError(Err.Number, Err.Description, "EncProperties.Display", "EncProperties", vbCritical)
    
End Function

Friend Sub SaveProperties()
    Dim iCnt As Integer
    Dim oAC As AutoCiteProp
    Dim sModule As String
    
    On Error GoTo ErrHandler
    
    ClearEncPropertyStrings
    
    iCnt = mcolAutoCites.Count
    For Each oAC In mcolAutoCites
        Select Case oAC.AppNCID
        Case NCID_PROBLEMS_APPLICATION:           sModule = "AutoCite_Problems"
        Case NCID_FAMILY_HISTORY_APPLICATION:           sModule = "AutoCite_Family"
        Case NCID_MEDICATIONS_APPLICATION:        sModule = "AutoCite_Meds"
        Case NCID_ALLERGIES_APPLICATION:          sModule = "AutoCite_Allergies"
        Case NCID_DEVICEREADING_APPLICATION:          sModule = "AutoCite_DeviceReadings"
        Case NCID_LAB_APPLICATION:                sModule = "AutoCite_Labs"
        Case NCID_RADIOLOGY_RESULTS_APPLICATION:  sModule = "AutoCite_Rads"
        Case NCID_VITAL_SIGNS_APPLICATION:        sModule = "AutoCite_Vitals"
        Case NCID_PATIENT_PLANOFCARE:  sModule = "AutoCite_Goals"
        Case NCID_PATIENT_QUESTIONNAIRES:  sModule = "AutoCite_Questionnaires"
        Case NCID_EDUCATION: sModule = "AutoCite_Handouts"
        Case NCID_EXPIRED_MEDICATIONS: sModule = "AutoCite_ExpiredMeds" 'SCR-20884
        Case Else
           'For items we do not save ie Inpatient Meds SCR#50672
           sModule = vbNullString
        End Select
    
        If Len(sModule) Then
            Call setValue(NCID_ENCOUNTER_APPLICATION, _
                                        mobjLogon.UserNCID, _
                                        sModule, _
                                        "Selected", _
                                        oAC.Selected)
            If oAC.AppNCID = NCID_LAB_APPLICATION _
                Or oAC.AppNCID = NCID_RADIOLOGY_RESULTS_APPLICATION _
                Or oAC.AppNCID = NCID_VITAL_SIGNS_APPLICATION Then
                
                Call setValue(NCID_ENCOUNTER_APPLICATION, _
                                            mobjLogon.UserNCID, _
                                            sModule, _
                                            "NumberOfPeriods", _
                                            oAC.NumberPeriods)
                Call setValue(NCID_ENCOUNTER_APPLICATION, _
                                            mobjLogon.UserNCID, _
                                            sModule, _
                                            "Period", _
                                            oAC.Period)
            End If
        End If
    Next oAC
    
    Call SaveShowActiveOrdersProperties
    
    Call SaveDefaultOptions
     

                                    
    '---save all of the changes made with "SetValue"---
    mbDidSave = SaveChanges(NCID_ENCOUNTER_APPLICATION, mobjLogon.UserNCID)
    
    Exit Sub
ErrHandler:
    Call mobjShared.ShowVBError(Err.Number, Err.Description, "EncProperties.SaveProperties ", "EncounterOps", vbCritical)
End Sub

Friend Sub SaveTemplateProperties()
    
    GEMS_SaveTemplateProperties
    GoTo CleanUpAndExit
    
ErrHandler:
    Call mobjShared.ShowVBError(Err.Number, Err.Description, "EncProperties.SaveProperties ", "EncounterOps", vbCritical)

CleanUpAndExit:

End Sub

Friend Sub GEMS_SaveTemplateProperties()
    

    Dim sSql As String
    'Dim oSQL As ICHCSII_DAS
    Dim oSQL As ICHCSII_SQL.ISqlOps

    Set oSQL = gobjCHCSIIConn.CHCSII_SQLOPS(Auto)  '''<SCR 36493
    
    If Not mbTemplatePropertiesExist Then
       sSql = "INSERT INTO TEMPLATE_AUTOCITE_PROPERTIES (ENCOUNTER_NUMBER, " & vbLf & _
            "FACILITY_NCID, UNIT_NUMBER, CITE_PROBLEMS, CITE_MEDS, CITE_ALLERGIES, CITE_QUESTIONNAIRES, CITE_FAMILY, "
       
       If mobjComm.CmdLineSwitch("CPG") <> "" Then
          sSql = sSql & "CITE_GOALS, CITE_DEVICES, CITE_HANDOUTS, " & vbLf
       End If
       
       sSql = sSql & "CITE_VITALS, VITALS_DURATION, VITALS_PERIOD, CITE_LABS, LABS_DURATION, " & vbLf & _
            "LABS_PERIOD, CITE_RADS, RADS_DURATION, RADS_PERIOD, SO_AUTOSAVE_DURATION) " & vbLf & _
            " VALUES ("
      
       sSql = sSql & SQLQuote(uEncounter.ID) & ", " & SQLQuote(uEncounter.Facility) & ", " & CStr(uEncounter.Unit_NUmber) & ", "
      
       sSql = sSql & CStr(Val(mcolTemplateAutoCites(CStr(NCID_PROBLEMS_APPLICATION)).Selected)) & ", "
       sSql = sSql & CStr(Val(mcolTemplateAutoCites(CStr(NCID_MEDICATIONS_APPLICATION)).Selected)) & ", "
       sSql = sSql & CStr(Val(mcolTemplateAutoCites(CStr(NCID_ALLERGIES_APPLICATION)).Selected)) & ", "
       sSql = sSql & CStr(Val(mcolTemplateAutoCites(CStr(NCID_PATIENT_QUESTIONNAIRES)).Selected)) & ", "
       sSql = sSql & CStr(Val(mcolTemplateAutoCites(CStr(NCID_FAMILY_HISTORY_APPLICATION)).Selected)) & ", "
       
       If mobjComm.CmdLineSwitch("CPG") <> "" Then
          sSql = sSql & CStr(Val(mcolTemplateAutoCites(CStr(NCID_PATIENT_PLANOFCARE)).Selected)) & ", "
          sSql = sSql & CStr(Val(mcolTemplateAutoCites(CStr(NCID_DEVICEREADING_APPLICATION)).Selected)) & ", "
          sSql = sSql & CStr(Val(mcolTemplateAutoCites(CStr(NCID_EDUCATION)).Selected)) & ", "
       End If
       
       sSql = sSql & CStr(Val(mcolTemplateAutoCites(CStr(NCID_VITAL_SIGNS_APPLICATION)).Selected)) & ", "
       sSql = sSql & SQLQuote(CStr(Val(mcolTemplateAutoCites(CStr(NCID_VITAL_SIGNS_APPLICATION)).NumberPeriods))) & ", "
       sSql = sSql & CStr(Val(mcolTemplateAutoCites(CStr(NCID_VITAL_SIGNS_APPLICATION)).Period)) & ", "
       sSql = sSql & CStr(Val(mcolTemplateAutoCites(CStr(NCID_LAB_APPLICATION)).Selected)) & ", "
       sSql = sSql & SQLQuote(CStr(Val(mcolTemplateAutoCites(CStr(NCID_LAB_APPLICATION)).NumberPeriods))) & ", "
       sSql = sSql & CStr(Val(mcolTemplateAutoCites(CStr(NCID_LAB_APPLICATION)).Period)) & ", "
       sSql = sSql & CStr(Val(mcolTemplateAutoCites(CStr(NCID_RADIOLOGY_RESULTS_APPLICATION)).Selected)) & ", "
       sSql = sSql & CStr(Val(mcolTemplateAutoCites(CStr(NCID_RADIOLOGY_RESULTS_APPLICATION)).NumberPeriods)) & ", "
       sSql = sSql & SQLQuote(CStr(Val(mcolTemplateAutoCites(CStr(NCID_RADIOLOGY_RESULTS_APPLICATION)).Period))) & ", "
       sSql = sSql & msAutoSaveInt & ") "
        
    Else
       sSql = "UPDATE TEMPLATE_AUTOCITE_PROPERTIES " & vbLf
       sSql = sSql & "SET CITE_PROBLEMS = " & CStr(Val(mcolTemplateAutoCites(CStr(NCID_PROBLEMS_APPLICATION)).Selected)) & ", " & vbLf
       sSql = sSql & "CITE_MEDS = " & CStr(Val(mcolTemplateAutoCites(CStr(NCID_MEDICATIONS_APPLICATION)).Selected)) & ", " & vbLf
       sSql = sSql & "CITE_ALLERGIES = " & CStr(Val(mcolTemplateAutoCites(CStr(NCID_ALLERGIES_APPLICATION)).Selected)) & ", " & vbLf
       sSql = sSql & "CITE_QUESTIONNAIRES = " & CStr(Val(mcolTemplateAutoCites(CStr(NCID_PATIENT_QUESTIONNAIRES)).Selected)) & ", " & vbLf
       sSql = sSql & "CITE_FAMILY = " & CStr(Val(mcolTemplateAutoCites(CStr(NCID_FAMILY_HISTORY_APPLICATION)).Selected)) & ", " & vbLf
       
       If mobjComm.CmdLineSwitch("CPG") <> "" Then
          sSql = sSql & "CITE_GOALS = " & CStr(Val(mcolTemplateAutoCites(CStr(NCID_PATIENT_PLANOFCARE)).Selected)) & ", " & vbLf
          sSql = sSql & "CITE_DEVICES = " & CStr(Val(mcolTemplateAutoCites(CStr(NCID_DEVICEREADING_APPLICATION)).Selected)) & ", " & vbLf
          sSql = sSql & "CITE_HANDOUTS = " & CStr(Val(mcolTemplateAutoCites(CStr(NCID_EDUCATION)).Selected)) & ", " & vbLf
       End If
       
       sSql = sSql & "CITE_VITALS = " & CStr(Val(mcolTemplateAutoCites(CStr(NCID_VITAL_SIGNS_APPLICATION)).Selected)) & ", " & vbLf
       sSql = sSql & "VITALS_DURATION = " & CStr(Val(mcolTemplateAutoCites(CStr(NCID_VITAL_SIGNS_APPLICATION)).NumberPeriods)) & ", " & vbLf
       sSql = sSql & "VITALS_PERIOD = " & SQLQuote(CStr(Val(mcolTemplateAutoCites(CStr(NCID_VITAL_SIGNS_APPLICATION)).Period))) & ", " & vbLf
       sSql = sSql & "CITE_LABS = " & CStr(Val(mcolTemplateAutoCites(CStr(NCID_LAB_APPLICATION)).Selected)) & ", " & vbLf
       sSql = sSql & "LABS_DURATION = " & CStr(Val(mcolTemplateAutoCites(CStr(NCID_LAB_APPLICATION)).NumberPeriods)) & ", " & vbLf
       sSql = sSql & "LABS_PERIOD = " & SQLQuote(CStr(Val(mcolTemplateAutoCites(CStr(NCID_LAB_APPLICATION)).Period))) & ", " & vbLf
       sSql = sSql & "CITE_RADS = " & CStr(Val(mcolTemplateAutoCites(CStr(NCID_RADIOLOGY_RESULTS_APPLICATION)).Selected)) & ", " & vbLf
       sSql = sSql & "RADS_DURATION = " & CStr(Val(mcolTemplateAutoCites(CStr(NCID_RADIOLOGY_RESULTS_APPLICATION)).NumberPeriods)) & ", " & vbLf
       sSql = sSql & "RADS_PERIOD = " & SQLQuote(CStr(Val(mcolTemplateAutoCites(CStr(NCID_RADIOLOGY_RESULTS_APPLICATION)).Period))) & ", " & vbLf
       sSql = sSql & "SO_AUTOSAVE_DURATION = " & msAutoSaveInt & vbLf
       sSql = sSql & " WHERE ENCOUNTER_NUMBER = " & SQLQuote(uEncounter.ID) & " AND " & vbLf
       sSql = sSql & "FACILITY_NCID = " & SQLQuote(uEncounter.Facility) & " AND " & vbLf
       sSql = sSql & "UNIT_NUMBER = " & SQLQuote(CStr(uEncounter.Unit_NUmber))
    End If
    
    '-------------------------
    'Bill Fote
    '2-Jan-2004
    'oSQL.Execute sSQL, True
    oSQL.Execute sSql
    '-------------------------
    
    mbDidSave = True
    
    'No Active Order in Theater mode
    'Call SaveShowActiveOrdersProperties
    
    Call SaveDefaultOptions  'SCR#42346
                                    
    '---save all of the changes made with "SetValue"---
    mbDidSave = SaveChanges(NCID_ENCOUNTER_APPLICATION, mobjLogon.UserNCID)
    
    GoTo CleanUpAndExit
    
ErrHandler:
    Call mobjShared.ShowVBError(Err.Number, Err.Description, "EncProperties.SaveProperties ", "EncounterOps", vbCritical)
CleanUpAndExit:
    Set oSQL = Nothing

End Sub

Public Sub SaveShowActiveOrdersProperties()

On Error GoTo ErrHandler
    Call setValue(NCID_ENCOUNTER_APPLICATION, _
                                    mobjLogon.UserNCID, _
                                    "A/P", _
                                    "ShowActiveMeds", _
                                    miShowActiveMeds, False)
                                    
    Call setValue(NCID_ENCOUNTER_APPLICATION, _
                                    mobjLogon.UserNCID, _
                                    "A/P", _
                                    "ShowActiveLabs", _
                                    miShowActiveLabs, False)
                                    
    Call setValue(NCID_ENCOUNTER_APPLICATION, _
                                    mobjLogon.UserNCID, _
                                    "A/P", _
                                    "ShowActiveRads", _
                                    miShowActiveRads, False)
                                    
    Exit Sub
    
ErrHandler:
    Call mobjShared.ShowVBError(Err.Number, Err.Description, "EncProperties.SaveShowActiveOrdersProperties ", "EncounterOps", vbCritical)
                                    
End Sub

' In Medcin Lite check if user has read priv for each module SCR-32805
' Set user preference if user has read priv SCR-33268
Public Sub RetrieveProperties()
'-- fills the autocite collection with the current stuff in the OSH
  Dim sbuf As String
  Dim oAC As AutoCiteProp
  Dim bSuccess As Boolean
    
  Set mcolAutoCites = New Collection
  Set mcolTemplateAutoCites = New Collection
  
  'SCR-49534 PERF PGUI Retrieve/Save Properties
  GetEncounterProperties
  
  mbRetrieveFromOSH = True
  '------------------------------------------------------------
  ' PROBLEMS
  '------------------------------------------------------------
'-- get each element of each autocite and add it to the collection
  Set oAC = New AutoCiteProp
  oAC.AppNCID = NCID_PROBLEMS_APPLICATION
  oAC.AppName = "Problems"
  Call getValue(NCID_ENCOUNTER_APPLICATION, "AutoCite_Problems", "Selected", sbuf)
  If mobjShared.IsAppMode(modeCHCSI_GUI) Then 'SCR 36493
    If Not mobjUser.HasPrivilegeEx("Medical_Problem_List", Priv_Read) Then
        oAC.Selected = vbUnchecked
    Else
        oAC.Selected = Val(sbuf)
    End If
  ElseIf Len(Trim$(sbuf)) = 0 Then 'Set default for new user SCR-23937
        oAC.Selected = vbChecked
        mbRetrieveFromOSH = False
  Else
    oAC.Selected = Val(sbuf)
  End If
  mcolAutoCites.Add oAC, CStr(oAC.AppNCID)

  Set oAC = New AutoCiteProp
  oAC.AppNCID = NCID_FAMILY_HISTORY_APPLICATION
  oAC.AppName = "Active Family History"  'SCR-42617
  Call getValue(NCID_ENCOUNTER_APPLICATION, "AutoCite_Family", "Selected", sbuf)
  If mobjShared.IsAppMode(modeCHCSI_GUI) Then 'SCR 36493
    oAC.Selected = vbUnchecked
  Else
    oAC.Selected = Val(sbuf)
  End If
  mcolAutoCites.Add oAC, CStr(oAC.AppNCID)

  '------------------------------------------------------------
  ' M E D S
  '------------------------------------------------------------
  Set oAC = New AutoCiteProp
  oAC.AppNCID = NCID_MEDICATIONS_APPLICATION
  oAC.AppName = "Active Medications" ' SCR-22642"Medications" 'Remove Dispensed SCR-38885
  Call getValue(NCID_ENCOUNTER_APPLICATION, "AutoCite_Meds", "Selected", sbuf)
  If mobjShared.IsAppMode(modeCHCSI_GUI) Then 'SCR 36493
    If Not mobjUser.HasPrivilegeEx("Medication_List", Priv_Read) Then
        oAC.Selected = vbUnchecked
    Else
        oAC.Selected = Val(sbuf)
    End If
  ElseIf Len(Trim$(sbuf)) = 0 Then 'Set default for new user SCR-23937
    oAC.Selected = vbChecked
  Else
    oAC.Selected = Val(sbuf)
  End If
  mcolAutoCites.Add oAC, CStr(oAC.AppNCID)
  
    Set oAC = New AutoCiteProp
    oAC.AppNCID = NCID_EXPIRED_MEDICATIONS
    oAC.AppName = "Expired Medications" 'SCR-20884
    Call getValue(NCID_ENCOUNTER_APPLICATION, "AutoCite_ExpiredMeds", "Selected", sbuf)
    If mobjShared.IsAppMode(modeCHCSI_GUI) Then 'SCR 36493
        oAC.Selected = Val(sbuf) 'SCR-36316
    ElseIf Len(Trim$(sbuf)) = 0 Then
        oAC.Selected = vbUnchecked
    Else
        oAC.Selected = Val(sbuf)
    End If
    mcolAutoCites.Add oAC, CStr(oAC.AppNCID)

  
  If Len(mobjComm.CmdLineSwitch("RNDS")) Then 'SCR#44005
          
          'SCR 44005
          'Bill Fote 2-Jan-2004 - This If is coming back False
    'Does not work since enc.category is not yet set
    'If StrComp(uEncounter.Category, "RNDS", vbTextCompare) = 0 Then
        Set oAC = New AutoCiteProp
        oAC.AppNCID = NCID_INPATIENT_MEDICATIONS
        oAC.AppName = "Inpatient Medications"
        'Always default to selected
        oAC.Selected = vbChecked
        mcolAutoCites.Add oAC, CStr(oAC.AppNCID)
    'End If
  End If
  
  
  '------------------------------------------------------------
  ' ALLERGIES
  '------------------------------------------------------------
  Set oAC = New AutoCiteProp
  oAC.AppNCID = NCID_ALLERGIES_APPLICATION
  oAC.AppName = "Allergies"
  Call getValue(NCID_ENCOUNTER_APPLICATION, "AutoCite_Allergies", "Selected", sbuf)
  If mobjShared.IsAppMode(modeCHCSI_GUI) Then 'SCR 36493
    If Not mobjUser.HasPrivilegeEx("Allergies", Priv_Read) Then
        oAC.Selected = vbUnchecked
    Else
        oAC.Selected = Val(sbuf)
    End If
  ElseIf Len(Trim$(sbuf)) = 0 Then 'Set default for new user SCR-23937
    oAC.Selected = vbChecked
  Else
    oAC.Selected = Val(sbuf)
  End If
  mcolAutoCites.Add oAC, CStr(oAC.AppNCID)
  
  '------------------------------------------------------------
  ' C P G S
  '------------------------------------------------------------
'<< 02.12.01 wrap this with cmd line switch check since CPGs may not have been loaded
If mobjComm.CmdLineSwitch("CPG") <> "" Then
   Set oAC = New AutoCiteProp
   oAC.AppNCID = NCID_PATIENT_PLANOFCARE
   oAC.AppName = "Goal Summary"
   Call getValue(NCID_ENCOUNTER_APPLICATION, "AutoCite_Goals", "Selected", sbuf)
   oAC.Selected = Val(sbuf)
   mcolAutoCites.Add oAC, CStr(oAC.AppNCID)

   Set oAC = New AutoCiteProp
   oAC.AppNCID = NCID_DEVICEREADING_APPLICATION
   oAC.AppName = "External Glucometer Device Results"
   Call getValue(NCID_ENCOUNTER_APPLICATION, "AutoCite_DeviceReadings", "Selected", sbuf)
   oAC.Selected = Val(sbuf)
   mcolAutoCites.Add oAC, CStr(oAC.AppNCID)

   Set oAC = New AutoCiteProp
   oAC.AppNCID = NCID_EDUCATION
   oAC.AppName = "Education Handouts"
   Call getValue(NCID_ENCOUNTER_APPLICATION, "AutoCite_Handouts", "Selected", sbuf)
   oAC.Selected = Val(sbuf)
   mcolAutoCites.Add oAC, CStr(oAC.AppNCID)
End If
'>>
  '------------------------------------------------------------
  ' L A B S
  '------------------------------------------------------------
  Set oAC = New AutoCiteProp
  oAC.AppNCID = NCID_LAB_APPLICATION
  oAC.AppName = "Labs"
  Call getValue(NCID_ENCOUNTER_APPLICATION, "AutoCite_Labs", "Selected", sbuf)
  If mobjShared.IsAppMode(modeCHCSI_GUI) Then 'SCR 36493
    If Not mobjUser.HasPrivilegeEx("Orders/Laboratory_Orders", Priv_Read) Then
        oAC.Selected = vbUnchecked
    Else
        oAC.Selected = Val(sbuf)
    End If
  Else
    oAC.Selected = Val(sbuf)
  End If
  
  Call getValue(NCID_ENCOUNTER_APPLICATION, "AutoCite_labs", "NumberOfPeriods", sbuf)
  oAC.NumberPeriods = Val(sbuf)
  'Set to 1 if NumberPeriods = 0 SCR-33825
  If oAC.NumberPeriods = 0 Then oAC.NumberPeriods = 1
  Call getValue(NCID_ENCOUNTER_APPLICATION, "AutoCite_labs", "Period", sbuf)
  oAC.Period = Val(sbuf)
  mcolAutoCites.Add oAC, CStr(oAC.AppNCID)
  
  '------------------------------------------------------------
  ' R A D S AutoCite_Rads
  '------------------------------------------------------------
  Set oAC = New AutoCiteProp
  oAC.AppNCID = NCID_RADIOLOGY_RESULTS_APPLICATION
  oAC.AppName = "Rads"
  Call getValue(NCID_ENCOUNTER_APPLICATION, "AutoCite_Rads", "Selected", sbuf)
  If mobjShared.IsAppMode(modeCHCSI_GUI) Then 'SCR 36493
    If Not mobjUser.HasPrivilegeEx("Orders/Radiology_Orders", Priv_Read) Then
        oAC.Selected = vbUnchecked
    Else
        oAC.Selected = Val(sbuf)
    End If
  Else
    oAC.Selected = Val(sbuf)
  End If
  Call getValue(NCID_ENCOUNTER_APPLICATION, "AutoCite_Rads", "NumberOfPeriods", sbuf)
  oAC.NumberPeriods = Val(sbuf)
  'Set to 1 if NumberPeriods = 0 SCR-33825
  If oAC.NumberPeriods = 0 Then oAC.NumberPeriods = 1
  Call getValue(NCID_ENCOUNTER_APPLICATION, "AutoCite_Rads", "Period", sbuf)
  oAC.Period = Val(sbuf)
  mcolAutoCites.Add oAC, CStr(oAC.AppNCID)
     
  '------------------------------------------------------------
  ' QUESTIONNAIRES
  '------------------------------------------------------------
  Set oAC = New AutoCiteProp
  oAC.AppNCID = NCID_PATIENT_QUESTIONNAIRES
  oAC.AppName = "Questionnaires"
  Call getValue(NCID_ENCOUNTER_APPLICATION, "AutoCite_Questionnaires", "Selected", sbuf)
  oAC.Selected = Val(sbuf)
  mcolAutoCites.Add oAC, CStr(oAC.AppNCID)

  '------------------------------------------------------------
  ' VITALS
  '------------------------------------------------------------
  'Do not load Vitals in PG mode SCR-37523
  If Not mobjShared.IsAppMode(modeCHCSI_GUI) Then
    Set oAC = New AutoCiteProp
    oAC.AppNCID = NCID_VITAL_SIGNS_APPLICATION
    oAC.AppName = "Vitals"
    Call getValue(NCID_ENCOUNTER_APPLICATION, "AutoCite_Vitals", "Selected", sbuf)
    If mobjShared.IsAppMode(modeCHCSI_GUI) Then 'SCR 36493
      oAC.Selected = vbUnchecked
    Else
       oAC.Selected = Val(sbuf)
    End If
    Call getValue(NCID_ENCOUNTER_APPLICATION, "AutoCite_Vitals", "NumberOfPeriods", sbuf)
    oAC.NumberPeriods = Val(sbuf)
    'Set to 1 if NumberPeriods = 0 SCR-33825
    If oAC.NumberPeriods = 0 Then oAC.NumberPeriods = 1
    Call getValue(NCID_ENCOUNTER_APPLICATION, "AutoCite_Vitals", "Period", sbuf)
    oAC.Period = Val(sbuf)
    mcolAutoCites.Add oAC, CStr(oAC.AppNCID)
  End If
  
  Set oAC = Nothing
  
  '*****************************************************************
  ' PROVIDER PREFERENCES PROPERTIES
  '*****************************************************************
  
  'Since Templates may have been used reset flag
  mbRetrieveFromOSH = True
  
  Call getValue(NCID_ENCOUNTER_APPLICATION, "SigBlock", "Line2", sbuf)
  SigBlockLine2 = sbuf
  
  Call getValue(NCID_ENCOUNTER_APPLICATION, "SigBlock", "Line3", sbuf)
  SigBlockLine3 = sbuf
  
  Call getValue(NCID_ENCOUNTER_APPLICATION, "SigBlock", "DefaultSigning", sbuf)
  If Len(Trim$(sbuf)) = 0 Then
    DefaultSigning = vbUnchecked 'Set default for new user SCR-23937
    mbRetrieveFromOSH = False  'SCR#50607
  Else
    DefaultSigning = Val(sbuf)
  End If
  
  Call getValue(NCID_ENCOUNTER_APPLICATION, "DefaultCosigner", "NCID", sbuf)
  msDefaultCosignerNCID = sbuf
'-Mod 2b (chartnotes)
  Call getValue(NCID_ENCOUNTER_APPLICATION, "AutoSaveInt", "Minutes", sbuf)
  msAutoSaveInt = Val(sbuf)
  
  Call getValue(NCID_ENCOUNTER_APPLICATION, "AutosaveSO", "Option", sbuf)
  If Len(Trim$(sbuf)) = 0 Then
    AutosaveSO = vbUnchecked 'Set default for new user SCR-23937
  Else
    AutosaveSO = Val(sbuf)
  End If
  
  Call getValue(NCID_ENCOUNTER_APPLICATION, "A/P", "ShowActiveMeds", sbuf)
  miShowActiveMeds = Val(sbuf)
  
  Call getValue(NCID_ENCOUNTER_APPLICATION, "A/P", "ShowActiveLabs", sbuf)
  miShowActiveLabs = Val(sbuf)
  
  Call getValue(NCID_ENCOUNTER_APPLICATION, "A/P", "ShowActiveRads", sbuf)
  miShowActiveRads = Val(sbuf)
  
  Call getValue(NCID_ENCOUNTER_APPLICATION, "Disposition", "DiscussedWith", sbuf)
  If Len(Trim$(sbuf)) = 0 Then 'Set default for new user SCR-23937
    mlDefaultDiscussedWith = 691     'NCID_INFO_SOURCE_PATIENT
  Else
    mlDefaultDiscussedWith = Val(sbuf)
  End If

  'SCR-35365
  Call getValue(NCID_ENCOUNTER_APPLICATION, "Encounter", "DefaultRecordType", sbuf)
  If Len(Trim$(sbuf)) = 0 Then 'Set default for new user
    mlDefaultRecordType = esrNone
  Else
    mlDefaultRecordType = Val(sbuf)
  End If
  
  Call getValue(NCID_ENCOUNTER_APPLICATION, "IncludeCodes", "Option", sbuf)
  mlIncludeCodes = Val(sbuf)
  
  If mobjEncPropDic.Count = 0 Then
    'PERF Save values in consolidated strings in OSH
    SaveProperties
  End If
  
  mbRetrieveFromOSH = True
  mbRetrievedProperties = True
  
End Sub


Public Function ParseTemplateData(pAutoCiteString As String) As Collection
  
      Dim sbuf As String
      Dim oAC As AutoCiteProp
      Dim bSuccess As Boolean
        
      Dim colList As New Collection
      Dim oTemplateData As Variant
      Dim oTemplateItem As Variant
      
      Set ParseTemplateData = New Collection
      Set mcolTemplateAutoCites = New Collection
      Set colList = DeserializeString(pAutoCiteString, "|")
      If colList.Count = 0 Then Exit Function
      oTemplateData = Split(colList(5), "^")
    
      '-- get each element of each autocite and add it to the collection
      Set oAC = New AutoCiteProp
      oAC.AppNCID = NCID_PROBLEMS_APPLICATION
      oAC.AppName = "Problems"
      oAC.Selected = 0
      For Each oTemplateItem In oTemplateData
           If oTemplateItem = oAC.AppNCID Then
              oAC.Selected = 1
              Exit For
           End If
      Next
      mcolTemplateAutoCites.Add oAC, CStr(oAC.AppNCID)
                      
      Set oAC = New AutoCiteProp
      oAC.AppNCID = NCID_FAMILY_HISTORY_APPLICATION
      oAC.AppName = "Active Family History"
      oAC.Selected = 0
      For Each oTemplateItem In oTemplateData
           If oTemplateItem = oAC.AppNCID Then
              oAC.Selected = 1
              Exit For
           End If
      Next
      mcolTemplateAutoCites.Add oAC, CStr(oAC.AppNCID)
    
      Set oAC = New AutoCiteProp
      oAC.AppNCID = NCID_MEDICATIONS_APPLICATION
      oAC.AppName = "Active Medications"
      oAC.Selected = 0
      For Each oTemplateItem In oTemplateData
           If oTemplateItem = oAC.AppNCID Then
              oAC.Selected = 1
              Exit For
           End If
      Next
      mcolTemplateAutoCites.Add oAC, CStr(oAC.AppNCID)
      
      'Add Expired Meds SCR-36977
      Set oAC = New AutoCiteProp
      oAC.AppNCID = NCID_EXPIRED_MEDICATIONS
      oAC.AppName = "Expired Medications"
      oAC.Selected = 0
      For Each oTemplateItem In oTemplateData
           If oTemplateItem = oAC.AppNCID Then
              oAC.Selected = 1
              Exit For
           End If
      Next
      mcolTemplateAutoCites.Add oAC, CStr(oAC.AppNCID)
              
    If Len(mobjComm.CmdLineSwitch("RNDS")) Then 'SCR#44005
        'SCR 44005, 45422
        Set oAC = New AutoCiteProp
        oAC.AppNCID = NCID_INPATIENT_MEDICATIONS
        oAC.AppName = "Inpatient Medications"
        'Always default to selected
        oAC.Selected = vbChecked
        mcolTemplateAutoCites.Add oAC, CStr(oAC.AppNCID)
    End If
              
              
      Set oAC = New AutoCiteProp
      oAC.AppNCID = NCID_ALLERGIES_APPLICATION
      oAC.AppName = "Allergies"
      oAC.Selected = 0
      For Each oTemplateItem In oTemplateData
           If oTemplateItem = oAC.AppNCID Then
              oAC.Selected = 1
              Exit For
           End If
      Next
      mcolTemplateAutoCites.Add oAC, CStr(oAC.AppNCID)
            
      '<< 02.12.01 wrap this with cmd line switch check since CPGs may not have been loaded
      If mobjComm.CmdLineSwitch("CPG") <> "" Then
         Set oAC = New AutoCiteProp
         oAC.AppNCID = NCID_PATIENT_PLANOFCARE
         oAC.AppName = "Goal Summary"
         oAC.Selected = 0
         For Each oTemplateItem In oTemplateData
              If oTemplateItem = oAC.AppNCID Then
                 oAC.Selected = 1
                 Exit For
              End If
         Next
         mcolTemplateAutoCites.Add oAC, CStr(oAC.AppNCID)
      
         Set oAC = New AutoCiteProp
         oAC.AppNCID = NCID_DEVICEREADING_APPLICATION
         oAC.AppName = "AutoCite_DeviceReadings"
         oAC.Selected = 0
         For Each oTemplateItem In oTemplateData
              If oTemplateItem = oAC.AppNCID Then
                 oAC.Selected = 1
                 Exit For
              End If
         Next
         mcolTemplateAutoCites.Add oAC, CStr(oAC.AppNCID)
      
         Set oAC = New AutoCiteProp
         oAC.AppNCID = NCID_EDUCATION
         oAC.AppName = "AutoCite_Handouts"
         oAC.Selected = 0
         For Each oTemplateItem In oTemplateData
              If oTemplateItem = oAC.AppNCID Then
                 oAC.Selected = 1
                 Exit For
              End If
         Next
         mcolTemplateAutoCites.Add oAC, CStr(oAC.AppNCID)
      End If
      '>>
              
      Set oAC = New AutoCiteProp
      oAC.AppNCID = NCID_LAB_APPLICATION
      oAC.AppName = "Labs"
      oAC.Selected = 0
      For Each oTemplateItem In oTemplateData
           If oTemplateItem = oAC.AppNCID Then
              oAC.Selected = 1
              Exit For
           End If
      Next
      Call getValue(NCID_ENCOUNTER_APPLICATION, "AutoCite_labs", "NumberOfPeriods", sbuf)
      oAC.NumberPeriods = Val(sbuf)
      'Set to 1 if NumberPeriods = 0 SCR-33825
      If oAC.NumberPeriods = 0 Then oAC.NumberPeriods = 1
      Call getValue(NCID_ENCOUNTER_APPLICATION, "AutoCite_labs", "Period", sbuf)
      oAC.Period = Val(sbuf)
      mcolTemplateAutoCites.Add oAC, CStr(oAC.AppNCID)
              
      Set oAC = New AutoCiteProp
      oAC.AppNCID = NCID_RADIOLOGY_RESULTS_APPLICATION
      oAC.AppName = "Rads"
      oAC.Selected = 0
      For Each oTemplateItem In oTemplateData
           If oTemplateItem = oAC.AppNCID Then
              oAC.Selected = 1
              Exit For
           End If
      Next
      Call getValue(NCID_ENCOUNTER_APPLICATION, "AutoCite_rads", "NumberOfPeriods", sbuf)
      oAC.NumberPeriods = Val(sbuf)
      'Set to 1 if NumberPeriods = 0 SCR-33825
      If oAC.NumberPeriods = 0 Then oAC.NumberPeriods = 1
      Call getValue(NCID_ENCOUNTER_APPLICATION, "AutoCite_rads", "Period", sbuf)
      oAC.Period = Val(sbuf)
      mcolTemplateAutoCites.Add oAC, CStr(oAC.AppNCID)
                 
      Set oAC = New AutoCiteProp
      oAC.AppNCID = NCID_PATIENT_QUESTIONNAIRES
      oAC.AppName = "Questionnaires"
      oAC.Selected = 0
      For Each oTemplateItem In oTemplateData
           If oTemplateItem = oAC.AppNCID Then
              oAC.Selected = 1
              Exit For
           End If
      Next
      mcolTemplateAutoCites.Add oAC, CStr(oAC.AppNCID)
            
      'Do not load Vitals in PG mode SCR-37523
      If Not mobjShared.IsAppMode(modeCHCSI_GUI) Then
        Set oAC = New AutoCiteProp
        oAC.AppNCID = NCID_VITAL_SIGNS_APPLICATION
        oAC.AppName = "Vitals"
        oAC.Selected = 0
        For Each oTemplateItem In oTemplateData
             If oTemplateItem = oAC.AppNCID Then
                oAC.Selected = 1
                Exit For
             End If
        Next
      
        Call getValue(NCID_ENCOUNTER_APPLICATION, "AutoCite_Vitals", "NumberOfPeriods", sbuf)
        oAC.NumberPeriods = Val(sbuf)
        'Set to 1 if NumberPeriods = 0 SCR-33825
        If oAC.NumberPeriods = 0 Then oAC.NumberPeriods = 1
        Call getValue(NCID_ENCOUNTER_APPLICATION, "AutoCite_Vitals", "Period", sbuf)
        oAC.Period = Val(sbuf)
        mcolTemplateAutoCites.Add oAC, CStr(oAC.AppNCID)
      End If
                  
      Set oAC = Nothing
      Set colList = Nothing
      Set ParseTemplateData = mcolTemplateAutoCites
      
End Function
Private Function DeserializeString(rStr As String, rsDelim As String) As Collection
    'Use to parse template data
    Dim cResult As Collection
    Dim X As Long
    Dim sField As String
    Dim sChr As String
    Set cResult = New Collection
    For X = 1 To Len(rStr)
        sChr = Mid(rStr, X, 1)
        Select Case sChr
        Case msEscChr
            sField = sField & Mid(rStr, X + 1, 1)
            X = X + 1
        Case rsDelim
            cResult.Add sField
            sField = ""
        Case Else
            sField = sField & sChr
        End Select
    Next X
    Set DeserializeString = cResult
End Function

Public Function RetrieveTemplateProperties(pProperties As Collection, Optional pEncounterInformation As Variant) As Boolean
    
    On Error GoTo ErrHandler
    RetrieveTemplateProperties = GEMS_RetrieveTemplateProperties(pProperties, pEncounterInformation)
    GoTo CleanUpAndExit

ErrHandler:
    Call mobjShared.ShowVBError(Err.Number, Err.Description, "EncProperties.RetrieveTemplateProperties", "EncProperties", vbCritical)

CleanUpAndExit:

End Function

Public Function GEMS_RetrieveTemplateProperties(pProperties As Collection, Optional pEncounterInformation As Variant) As Boolean
    
    Dim sSql As String
    Dim oSQL As ICHCSII_SQL.ISqlOps
    Dim oAC As AutoCiteProp
    
    GEMS_RetrieveTemplateProperties = False
    mbTemplatePropertiesExist = False
    
    On Error GoTo ErrHandler
    Set oSQL = gobjCHCSIIConn.CHCSII_SQLOPS(Auto)  '''<SCR 36493
    Set pProperties = New Collection

    
    If Not IsMissing(pEncounterInformation) Then uEncounter = pEncounterInformation
    
    sSql = "SELECT * FROM TEMPLATE_AUTOCITE_PROPERTIES" & vbLf
    With uEncounter
        sSql = sSql & " WHERE ENCOUNTER_NUMBER = " & SQLQuote(.ID) & vbLf
        sSql = sSql & " AND FACILITY_NCID = " & SQLQuote(.Facility) & vbLf
        sSql = sSql & " AND UNIT_NUMBER = " & CStr(.Unit_NUmber)
    End With
    
    oSQL.Execute sSql
    If Not (oSQL.BOF And oSQL.EOF) Then
      '-- get each element of each autocite and add it to the collection
      Set oAC = New AutoCiteProp
      oAC.AppNCID = NCID_PROBLEMS_APPLICATION
      oAC.AppName = "Problems"
      oAC.Selected = oSQL("CITE_PROBLEMS")
      pProperties.Add oAC, CStr(oAC.AppNCID)
            
      Set oAC = New AutoCiteProp
      oAC.AppNCID = NCID_FAMILY_HISTORY_APPLICATION
      oAC.AppName = "Active Family History"
      oAC.Selected = oSQL("CITE_FAMILY")
      pProperties.Add oAC, CStr(oAC.AppNCID)
            
      Set oAC = New AutoCiteProp
      oAC.AppNCID = NCID_MEDICATIONS_APPLICATION
      oAC.AppName = "Medications"
      oAC.Selected = oSQL("CITE_MEDS")
      pProperties.Add oAC, CStr(oAC.AppNCID)
              
      Set oAC = New AutoCiteProp
      oAC.AppNCID = NCID_ALLERGIES_APPLICATION
      oAC.AppName = "Allergies"
      oAC.Selected = oSQL("CITE_ALLERGIES")
      pProperties.Add oAC, CStr(oAC.AppNCID)
            
      '<< 02.12.01 wrap this with cmd line switch check since CPGs may not have been loaded
      If mobjComm.CmdLineSwitch("CPG") <> "" Then
         Set oAC = New AutoCiteProp
         oAC.AppNCID = NCID_PATIENT_PLANOFCARE
         oAC.AppName = "Goal Summary"
         oAC.Selected = oSQL("CITE_GOALS")
         pProperties.Add oAC, CStr(oAC.AppNCID)
      
         Set oAC = New AutoCiteProp
         oAC.AppNCID = NCID_DEVICEREADING_APPLICATION
         oAC.AppName = "AutoCite_DeviceReadings"
         oAC.Selected = oSQL("CITE_DEVICES")
         pProperties.Add oAC, CStr(oAC.AppNCID)
      
         Set oAC = New AutoCiteProp
         oAC.AppNCID = NCID_EDUCATION
         oAC.AppName = "AutoCite_Handouts"
         oAC.Selected = oSQL("CITE_HANDOUTS")
         pProperties.Add oAC, CStr(oAC.AppNCID)
      End If
      '>>
              
      Set oAC = New AutoCiteProp
      oAC.AppNCID = NCID_LAB_APPLICATION
      oAC.AppName = "Labs"
      oAC.Selected = oSQL("CITE_LABS")
      oAC.NumberPeriods = Val(oSQL("LABS_DURATION"))
      oAC.Period = Val(oSQL("LABS_PERIOD"))
      pProperties.Add oAC, CStr(oAC.AppNCID)
              
      Set oAC = New AutoCiteProp
      oAC.AppNCID = NCID_RADIOLOGY_RESULTS_APPLICATION
      oAC.AppName = "Rads"
      oAC.Selected = oSQL("CITE_RADS")
      oAC.NumberPeriods = Val(oSQL("RADS_DURATION"))
      oAC.Period = Val(oSQL("RADS_PERIOD"))
      pProperties.Add oAC, CStr(oAC.AppNCID)
                 
      Set oAC = New AutoCiteProp
      oAC.AppNCID = NCID_PATIENT_QUESTIONNAIRES
      oAC.AppName = "Questionnaires"
      oAC.Selected = oSQL("CITE_QUESTIONNAIRES")
      pProperties.Add oAC, CStr(oAC.AppNCID)
            
      Set oAC = New AutoCiteProp
      oAC.AppNCID = NCID_VITAL_SIGNS_APPLICATION
      oAC.AppName = "Vitals"
      oAC.Selected = oSQL("CITE_VITALS")
      oAC.NumberPeriods = Val(oSQL("VITALS_DURATION"))
      oAC.Period = Val(oSQL("VITALS_PERIOD"))
      pProperties.Add oAC, CStr(oAC.AppNCID)
                          
      AutoSaveInt = oSQL("SO_AUTOSAVE_DURATION")
      mbTemplatePropertiesExist = True
    End If
    
    Set mcolTemplateAutoCites = pProperties
    GEMS_RetrieveTemplateProperties = True
    mbRetrievedProperties = False
    GoTo CleanUpAndExit
    
ErrHandler:
    Call mobjShared.ShowVBError(Err.Number, Err.Description, "EncProperties.RetrieveTemplateProperties", "EncProperties", vbCritical)

CleanUpAndExit:
    Set oAC = Nothing

End Function

Public Property Get Encounter_Information() As Encounter_Key_Information
    Encounter_Information = uEncounter
End Property


Public Property Get TemplatePropertiesExist() As Boolean
    TemplatePropertiesExist = mbTemplatePropertiesExist
End Property


Public Property Get TemplateAutoCites() As Collection
    If mcolTemplateAutoCites Is Nothing Then
       Set mcolTemplateAutoCites = New Collection
    End If
    Set TemplateAutoCites = mcolTemplateAutoCites
End Property


Public Property Get DefaultSigning() As Integer
    DefaultSigning = miDefaultSigning
End Property

Public Property Let DefaultSigning(ByVal RHV As Integer)
    miDefaultSigning = RHV
End Property

Public Property Get IncludeCodes() As Long  'SCR-38957
    If Len(mobjComm.CmdLineSwitch("OPTION")) Then
        IncludeCodes = mlIncludeCodes
    Else
        IncludeCodes = 0
    End If
End Property

Public Property Let IncludeCodes(ByVal value As Long)
    mlIncludeCodes = value
End Property

Public Sub SaveDefaultOptions(Optional ByVal bSaveChanges As Boolean = False)

    Call setValue(NCID_ENCOUNTER_APPLICATION, _
                                   mobjLogon.UserNCID, _
                                   "SigBlock", _
                                   "Line2", _
                                   msSigBlockLine2, False)
                                   
    Call setValue(NCID_ENCOUNTER_APPLICATION, _
                                   mobjLogon.UserNCID, _
                                   "SigBlock", _
                                   "Line3", _
                                   msSigBlockLine3, False)
    
    Call setValue(NCID_ENCOUNTER_APPLICATION, _
                                   mobjLogon.UserNCID, _
                                   "SigBlock", _
                                   "DefaultSigning", _
                                   CStr(miDefaultSigning), False)

    '-- cosigning provider if needed
    Call setValue(NCID_ENCOUNTER_APPLICATION, _
                                   mobjLogon.UserNCID, _
                                   "DefaultCosigner", _
                                   "NCID", _
                                   msDefaultCosignerNCID, False)
    '-Autosave int
    Call setValue(NCID_ENCOUNTER_APPLICATION, _
                                   mobjLogon.UserNCID, _
                                   "AutoSaveInt", _
                                   "Minutes", _
                                   msAutoSaveInt, False)
     
    Call setValue(NCID_ENCOUNTER_APPLICATION, _
                                   mobjLogon.UserNCID, _
                                   "AutosaveSO", _
                                   "Option", _
                                   CStr(miAutosaveSO), False)

    Call setValue(NCID_ENCOUNTER_APPLICATION, _
                                   mobjLogon.UserNCID, _
                                   "Disposition", _
                                   "DiscussedWith", _
                                   mlDefaultDiscussedWith, False)
                                   
    'SCR-35365
    Call setValue(NCID_ENCOUNTER_APPLICATION, _
                                   mobjLogon.UserNCID, _
                                   "Encounter", _
                                   "DefaultRecordType", _
                                   mlDefaultRecordType, False)
                                   
    'SCR-38957
    Call setValue(NCID_ENCOUNTER_APPLICATION, _
                                   mobjLogon.UserNCID, _
                                   "IncludeCodes", _
                                   "Option", _
                                   CStr(mlIncludeCodes), False)
                                   
    If bSaveChanges Then
        SaveChanges NCID_ENCOUNTER_APPLICATION, mobjLogon.UserNCID
    End If
                                   
End Sub
Public Sub SaveProviderPreferenceOptions()

    On Error GoTo ErrHandler
    
    SaveDefaultOptions

    SaveShowActiveOrdersProperties
    
    SaveChanges NCID_ENCOUNTER_APPLICATION, mobjLogon.UserNCID
    
    Exit Sub
    
ErrHandler:
    Call mobjShared.ShowVBError(Err.Number, Err.Description, "EncProperties.SaveProviderPreferenceOptions", "EncProperties", vbCritical)

End Sub

Public Property Get AutosaveSO() As Integer
    AutosaveSO = miAutosaveSO
End Property

Public Property Let AutosaveSO(ByVal RHV As Integer)
    miAutosaveSO = RHV
End Property

Public Property Get RetrievedProperties() As Boolean
    RetrievedProperties = mbRetrievedProperties
End Property

Public Property Let RetrievedProperties(ByVal RHV As Boolean)
    mbRetrievedProperties = RHV
End Property

'SCR-49534 PERF PGUI Retrieve/Save Properties
Private Function getValue(ByVal sApplicationNCID As String, ByVal sSection As String, _
                         ByVal sKey As String, ByRef sValue As String) As Boolean
                         
    If mobjEncPropDic Is Nothing Then
        Set mobjEncPropDic = New Dictionary
    End If
    
    msMsg = "Get Value - " & sSection & " : " & sKey
    
    'See if it exists in the dictionary
    If mobjEncPropDic.Exists(sSection & ":" & sKey) Then
        msMsg = "Get from Dictionary - " & sSection & " : " & sKey
        sValue = mobjEncPropDic.Item(sSection & ":" & sKey)
    Else
        'retrieve from OSH
        If mbRetrieveFromOSH Then
            msMsg = "Get from OSH - " & sSection & " : " & sKey
            Call mobjConfig.getValue(sApplicationNCID, sSection, sKey, sValue)
        Else
            'OSH Settings do not exist so use default value
            sValue = vbNullString
        End If
    End If
    
    Exit Function
    
ErrHandler:
    Call mobjShared.ShowVBError(Err.Number, Err.Description & " - " & msMsg, "EncProperties.getValue", "EncProperties", vbCritical)
    

End Function
'SCR-49534 PERF PGUI Retrieve/Save Properties
Private Function setValue(ByVal sApplicationNCID As String, ByVal sOwnerNCID As String, _
                         ByVal sSection As String, ByVal sKey As String, _
                         ByVal sValue As String, _
                         Optional ByVal bAutociteRelated As Boolean = True) As Boolean
                         
    On Error GoTo ErrHandler
    
    If mobjShared.IsAppMode(modeCHCSI_GUI) Then
        If bAutociteRelated Then
            msAutociteProperties = msAutociteProperties & sSection & ":" & sKey & gsListDelim & sValue & gsMainDelim
        Else
            msProvPrefProperties = msProvPrefProperties & sSection & ":" & sKey & gsListDelim & sValue & gsMainDelim
        End If
    Else
        mobjConfig.setValue sApplicationNCID, sOwnerNCID, sSection, sKey, sValue
    End If
    
    
    Exit Function
    
ErrHandler:
    Call mobjShared.ShowVBError(Err.Number, Err.Description, "EncProperties.setValue", "EncProperties", vbCritical)
        
End Function
'SCR-49534 PERF PGUI Retrieve/Save Properties
Private Sub GetEncounterProperties()

    Dim sbuf1() As String
    Dim sbuf2() As String
    Dim i As Long
    
    If Not mobjShared.IsAppMode(modeCHCSI_GUI) Then Exit Sub
    
    Set mobjEncPropDic = Nothing
    Set mobjEncPropDic = New Dictionary
    
    ClearEncPropertyStrings
    
    msMsg = "Get Autocite Prop"
    Call mobjConfig.getValue(NCID_ENCOUNTER_APPLICATION, "EncProperties", "AutociteValues", msAutociteProperties)
    
    msMsg = "Get Prov Pref Prop"
    Call mobjConfig.getValue(NCID_ENCOUNTER_APPLICATION, "EncProperties", "ProvPrefValues", msProvPrefProperties)
    
    msMsg = "Parse Autocite Prop - " & msAutociteProperties
    If Len(msAutociteProperties) Then
        sbuf1 = Split(msAutociteProperties, "|")
        
        For i = 0 To UBound(sbuf1)
            sbuf2 = Split(sbuf1(i), "^")

            mobjEncPropDic.Add sbuf2(0), sbuf2(1)
        Next
    End If
    
    msMsg = "Parse Prov Pref Prop - " & msProvPrefProperties
    If Len(msProvPrefProperties) Then
        sbuf1 = Split(msProvPrefProperties, "|")
        
        For i = 0 To UBound(sbuf1)
            sbuf2 = Split(sbuf1(i), "^")

            mobjEncPropDic.Add sbuf2(0), sbuf2(1)
        Next
    End If

    ClearEncPropertyStrings
    
    Exit Sub
    
ErrHandler:
    Call mobjShared.ShowVBError(Err.Number, Err.Description & " - " & msMsg, "EncProperties.GetEncProperties", "EncProperties", vbCritical)
    
End Sub
'SCR-49534 PERF PGUI Retrieve/Save Properties
Private Function SaveChanges(ByVal sApplicationNCID As String, ByVal sOwnerNCID As String) As Boolean

    If mobjShared.IsAppMode(modeCHCSI_GUI) Then
        If Len(msAutociteProperties) Then
            'Remove last delimiter
            msAutociteProperties = Left(msAutociteProperties, Len(msAutociteProperties) - 1)
            msMsg = "Set AutoCite Properties"
            mobjConfig.setValue sApplicationNCID, sOwnerNCID, "EncProperties", "AutociteValues", msAutociteProperties
        End If
        
        If Len(msProvPrefProperties) Then
            msProvPrefProperties = Left(msProvPrefProperties, Len(msProvPrefProperties) - 1)
            msMsg = "Set Prov Pref Properties"
            mobjConfig.setValue sApplicationNCID, sOwnerNCID, "EncProperties", "ProvPrefValues", msProvPrefProperties
        End If
    End If
    
    msMsg = "Save to OSH"
    SaveChanges = mobjConfig.SaveChanges(sApplicationNCID, sOwnerNCID)
    
    ClearEncPropertyStrings
    
    Exit Function
    
ErrHandler:
    Call mobjShared.ShowVBError(Err.Number, Err.Description & " - " & msMsg, "EncProperties.SaveChanges", "EncProperties", vbCritical)
End Function

Private Sub ClearEncPropertyStrings()
    msAutociteProperties = vbNullString
    msProvPrefProperties = vbNullString
End Sub
