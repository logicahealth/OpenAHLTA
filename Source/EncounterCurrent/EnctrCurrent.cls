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
Attribute VB_Name = "Enctr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "The main controing class for Encounter Summary."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"EncProperties"
Attribute VB_Ext_KEY = "Member1" ,"EncRTF"
Attribute VB_Ext_KEY = "Member2" ,"EncProperties"
'-----------------------------------------------------
'Purpose: Controler class for Encounters and Encounter GUI
'
'Comments:  controls all aspects of the encounter module.  Main enctry point.
'           contains references to all parent objects and a collection of encounters.
'----------------------------------------------------
Option Explicit

Implements ICHCSII_CoreAPI2                  '-- implements the security obj

Private mobjUser       As ICHCSII_User       '-- The Security User Object
Private mcolEncounters As Collection         '-- previous (and current) encounters
Private mcolEncChild   As Collection         '-- encounter child modules
Private WithEvents mobjCurrEnc As Encounter  '-- current active encounter
Attribute mobjCurrEnc.VB_VarHelpID = -1
Private mcolMenus      As Collection         '-- menus I've put on the screen
Private mobjPatient    As Object             '-- patient object from core
Private mobjLogonEnvironment As Logon        '-- logon env obj from core
Private mobjComm       As Object             '-- comm obj from core
Private mlProgID       As Long               '-- my prog id from core
Private mvarCurrPatientUnitNumber As String  '- mobjPatient.UnitNumber
Private mobjWorkFlow   As EncWorkFlow
Public UnloadFailed    As Boolean            '-- used to flag a failed query-unload event to calling object
Private sOriginals(8)  As sOriginals_        '-- the original set of rtf strings.
Private mobjProperties As EncProperties      '-- the autocite an other stuff from the OSH
Private mbSelectionCanceled As Boolean       '- flag for select dialog box to set if cancel is selected

Private WithEvents mEncChild As EncChild
Attribute mEncChild.VB_VarHelpID = -1
Public Event ChangedEncounter(EncNum As Long, FacilityNCID As String)
Public Event EnctrErr(ErrNum As Integer, ErrStr As String)
Public Event CurrEncStatusChanged(EStatus As EncStatus)
Public Event CreatedEncounter(oEnc As Encounter)
Public Event OpenEncounter(oEnc As Encounter)
Public Event CloseEncounter(oEnc As Encounter)

Public Enum MenuNames
  mnAction = 0
  mnView = 1
End Enum

'-- implements the dlog.dll timing module
#If debugon Then
  Private DR As DebugRec
  Private DL As cDebugLog
  Private Const Exe = "Encounter"
  Private Const Module = "Enctrs.cls"
#End If

Private mobjUtil                    As EncUtil
Private sChildMenuIDs               As String     '-- a delimited list of child menu IDs
Private bFirstRegularResize         As Boolean
Private bFirstPatientSummaryResize  As Boolean
Private bMainViewRunning            As Boolean
Private bPatientSummaryViewRunning  As Boolean

Public Property Get MainViewRunning() As Boolean:           MainViewRunning = bMainViewRunning:          End Property
Public Property Let MainViewRunning(RHV As Boolean):        bMainViewRunning = RHV:                      End Property
Public Property Get Comm() As Object:                       Set Comm = mobjComm:                         End Property
Public Property Get Config() As Object:                     Set Config = mobjConfig:                     End Property
Public Property Get CurrentEncounter() As Encounter:        Set CurrentEncounter = mobjCurrEnc:          End Property
Public Property Get EncChildren() As Collection:            Set EncChildren = mcolEncChild:              End Property
Public Property Get Encounters() As Collection:             Set Encounters = mcolEncounters:             End Property
Public Property Get EncProperties() As EncProperties:       Set EncProperties = mobjProperties:          End Property
Public Property Get hwnd() As Long:                         hwnd = frmEncParent.hwnd:                    End Property
Public Property Get LogonEnvironment() As Object:           Set LogonEnvironment = mobjLogonEnvironment: End Property
Public Property Get Menus() As Collection:                  Set Menus = mcolMenus:                       End Property
Public Property Get Patient() As Object:                    Set Patient = mobjPatient:                   End Property
Public Property Get ProgID() As Long:                       ProgID = mlProgID:                           End Property
Public Property Get User() As ICHCSII_User:                 Set User = mobjUser:                         End Property
Public Property Get WorkFlow() As EncWorkFlow:              Set WorkFlow = mobjWorkFlow:                 End Property
Public Property Let Comm(ByVal RHV As Object):              Set mobjComm = RHV:                          End Property
Public Property Let Config(ByVal RHV As Object):            Set mobjConfig = RHV:                        End Property
Public Property Let EncProperties(RHV As EncProperties):    Set mobjProperties = RHV:                    End Property
Public Property Let LogonEnvironment(ByVal RHV As Object):  Set mobjLogonEnvironment = RHV:              End Property
Public Property Let Patient(ByVal RHV As Object):           Set mobjPatient = RHV:                       End Property
Public Property Let ProgID(ByVal RHV As Long):              mlProgID = RHV:                              End Property
Public Property Let User(vData As ICHCSII_User):            Set mobjUser = vData:                        End Property
Public Property Set CurrentEncounter(RHV As Encounter):     Set mobjCurrEnc = RHV:                       End Property
Public Property Set WorkFlow(RHV As EncWorkFlow):           mobjWorkFlow = RHV:                          End Property
Public Property Let SelectionCanceled(RHV As Boolean):      mbSelectionCanceled = RHV:                   End Property

'=========== End of standard property exposure code ===========================

'====================================================
' Purpose: retrieve a specific encounter
'
' Input:
'    patientUnitNumber - the patient unit number
'    encounternumber - the desired encounter number
'    facilityNcid  - the facility for the enc
' Returns: an encounter
'Called By:
'- OpenView
'- PatientSelected
'- Initialize
'- AddProvider
'- LeftWithoutBeingSeen
'Calls:
'- Init                                                      Encounter
'- Patient [Property Get]
'- Patient [Property Set]                                    Encounter
'- RetrieveEnc
'====================================================
Public Function GetEncounter(ByVal PatientUnitNumber As String, _
                             ByVal FacilityNCID As String, _
                             ByVal EncounterNumber As String) As Encounter
Attribute GetEncounter.VB_Description = "Returns a single (existing) encounter based on patient and encounter numbers. "
  Dim oEnc As Encounter
  Dim oDas As DasCtrl
  
  If Not DebugMode Then On Error GoTo errHandler
  
  #If debugon Then
    Dim DR As DebugRec
    If Not DL Is Nothing Then DL.DebugLog Exe, Module, "GetEncounter", "Begin", DR
  #End If
  
  Call mobjComm.setstatustext("Retrieving encounter record " & EncounterNumber)

  Set oDas = New DasCtrl
  oDas.ClearAll
  oDas.GeneralCriteria.ClearAll
  oDas.GeneralCriteria.UnitNumber = PatientUnitNumber
  oDas.GeneralCriteria.Facility.ValueConcept = FacilityNCID
  oDas.GeneralCriteria.EncounterNumber = EncounterNumber
  oDas.GeneralCriteria.QueryQualifier = "EVENT_TYPE_NCID = '" & NCID_ENCEVENT & "'"
  oDas.GeneralCriteria.ReturnMultiMedia = True
  
  oDas.DoSearch
  
  If Not oDas.EOF Then
    Set oEnc = New Encounter
    oEnc.Init mobjUser
    
    If CStr(mobjPatient.UnitNumber) <> PatientUnitNumber Then
        mobjComm.Message cwiSELECT_PATIENT, PatientUnitNumber, NCID_CW_APPLICATION, mlProgID
    End If
    
    Set oEnc.Patient = mobjPatient
    
    If oEnc.RetrieveEnc(oDas) Then
      Call oEnc.SetCurrentRTF
      Set GetEncounter = oEnc
      RaiseEvent ChangedEncounter(oEnc.EncounterID, oEnc.FacilityNCID)
      RaiseEvent OpenEncounter(oEnc)
    End If
  
  End If

  GoTo Cleanup

errHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "Enctr.GetEncounter ", "Enc", vbCritical)
Cleanup:
#If debugon Then
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "GetEncounter", "End", DR
#End If
   Set oEnc = Nothing
   Set oDas = Nothing
End Function

'*******************************************************************************************
'  Description: This method is called when a client wants to know if this OLE Server
'               supports a specific "view type".
'
'  Parameters:
'    viViewType - Contains the "type" of view being queried for.
'
'  Return Value:
'    Returns TRUE if this OLE Server supports the passed in "view type".
'    Returns FALSE if this OLE Server does not support the passed in "view type".
'*******************************************************************************************
Public Function QueryView(ByVal viViewType As Integer) As Boolean
Attribute QueryView.VB_Description = "As mandated by CW"

#If debugon Then
    Dim DR As DebugRec
    If Not DL Is Nothing Then DL.DebugLog Exe, Module, "QueryView", "Begin", DR
#End If

  If Not DebugMode Then On Error GoTo errHandler

    Select Case viViewType
        Case cwiREGULAR_VIEW
            QueryView = True    '--this view is NOT optional--
        Case cwiPATIENT_SUMMARY_VIEW
            QueryView = True
        Case cwiCLINICIAN_SUMMARY_VIEW
            QueryView = False
        Case cwiENCOUNTER_SUMMARY_VIEW
            QueryView = False
        Case Else   '--unknown "view type", not supported--
            QueryView = False
    End Select
#If debugon Then
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "QueryView", "End", DR
#End If

  Exit Function
errHandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "Enctr.QueryView", "EncounterParent", vbExclamation)
    
    Exit Function

End Function

'**************************************************************************************************
'  Description: This method is called when the CW Core attempts to open a "view" provided by this
'               OLE server.
'
'  Parameters:
'    viViewType - Contains the "type" of view being requested by the CW Core to open.
'    vlParenthWnd - Contains the hWnd of the client area window to be used by this OLE Server.
'    rhMainForm - If "viViewType" is a "regular view", upon return this will contain
'                 the window "handle" of the main "regular" form for this OLE Server.
'
'  Return Value:
'    Returns TRUE if the "view" was opened successfully.
'    Returns FALSE if an error occurred openning the "view".
'
' Comment: Returning False will cause core to not continue loading the module.
'**************************************************************************************************
Public Function OpenView(ByVal viViewType As Integer, _
                         ByVal vlParenthWnd As Long, _
                               rhMainForm As Long) As Boolean
Attribute OpenView.VB_Description = "As mandated by CW"
                               
     
    Dim msg As String    '- message string holder
    Dim s   As String    '- generic string
    Dim oChild As EncChild
    
  Set mcolEncounters = New Collection
 
#If debugon Then
  Dim DR As DebugRec
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "OpenView", "Begin", DR
#End If

'-01.21.00 WAS - this is the first place we can say 'No' to core. So check for all required parameters.
'  see if each of the following is set:
  msg = ""
  If mobjUser Is Nothing Then msg = msg & Chr(10) & "Security User Object"
  If mobjPatient Is Nothing Then msg = msg & Chr(10) & "Patient Object"
  If mobjConfig Is Nothing Then msg = msg & Chr(10) & "Configuration Object"
  If mobjLogonEnvironment Is Nothing Then msg = msg & Chr(10) & "Logon Object"
  If mobjComm Is Nothing Then msg = msg & Chr(10) & "Core Communication Object"
  If msg <> "" Then
    msg = Mid$(msg, 2)   'drop the leading 'lf'
    MsgBox "The following object(s) were not handed-off to the Encounter Object:" & Chr(10) & Chr(10) _
       & msg & Chr(10) & Chr(10) & "Encounter cannot continue to load and is shutting down.", vbCritical, "Open Encounter Error"
    Exit Function
  End If
  msg = ""
'~end ~~

'-- are you allowed in here?
  If Not mobjUser.HasPrivilegeEx(Priv_Current_Encounter, Priv_Read) Then
    MsgBox "You do not have permission to view this information." & Chr$(10) _
      & "Please contact your security administrator if you believe you have received this message in error.", vbExclamation
    Exit Function
  End If
  On Error Resume Next
  Err.Clear
  s = mobjUser.GetAttribute("USER_Ncid")
  If Err Then
    MsgBox "Your Identification was not found in the security file. Please notify your security administrator.", vbCritical, "Security Error"
    Exit Function
    Err.Clear
  End If
  
  If Not DebugMode Then On Error GoTo errHandler
    
    glParentWindow = vlParenthWnd
    
    '---perform actions based on the "view type"---
    Select Case viViewType
    
    Case cwiREGULAR_VIEW
      OpenView = OpenRegularView(vlParenthWnd, rhMainForm)
      
    Case cwiPATIENT_SUMMARY_VIEW
      SetParent frmEncSOC.hwnd, vlParenthWnd
      frmEncSOC.ParentObject = Me
      Call GetEncounterInfo    'RetrievePreviousEncountersEx
      DoEvents
      frmEncSOC.Encounters = mcolEncounters
      rhMainForm = frmEncSOC.hwnd
      bFirstPatientSummaryResize = True
      bPatientSummaryViewRunning = True
      OpenView = True
      
    Case Else    '--unsupported view "type"--
      OpenView = False
      Exit Function
    End Select
        
#If debugon Then
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "OpenView", "End", DR
#End If

    Exit Function
    
errHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "Enctr.OpenView", "Enc", vbCritical)
    OpenView = False
End Function

Private Function OpenRegularView(ByVal vlParenthWnd As Long, _
                                       rhMainForm As Long) As Boolean

#If debugon Then
  Dim DR As DebugRec
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "OpenRegularView", "Begin", DR
#End If

  If Not DebugMode Then On Error GoTo errHandler

      Debug.Print "Open - Pat: " & (mobjPatient.Name) & " Enc Num: " & (mobjPatient.EncounterNumber)
      
      If Val(mobjPatient.EncounterNumber) > 0 Then
        
        Select Case True
          Case mobjCurrEnc Is Nothing
            If Not OpenViewGetEncounter(False) Then Exit Function
          
          Case mobjCurrEnc.EncounterID <> Val(mobjPatient.EncounterNumber) _
            Or mobjCurrEnc.FacilityNCID <> mobjPatient.EncounterFacilityNCID
            If Not OpenViewGetEncounter(False) Then Exit Function
            
          Case Not mobjCurrEnc.FullyLoaded
            If Not OpenViewGetEncounter(False) Then Exit Function
        End Select
        
      Else
        If Not OpenViewGetEncounter(True) Then Exit Function
      End If
    
 '-Mod 25e - moved from OpenViewGetEncounter
  Select Case mobjCurrEnc.Status
'-Mod 37e
  Case Complete, PatientLeftWithoutBeingSeen, Updated, NeedsCoSignature
  '-- since I don't show these in the 'enc select' window, I must be being asked programmatically to display them.
  '   Send the user to the prev enc module.
  '-- First ensure the patient is pointing to this enc and launch prev enc, he'll open to it
      mobjPatient.EncounterNumber = mobjCurrEnc.EncounterID
      mobjPatient.EncounterFacilityNCID = mobjCurrEnc.FacilityNCID
      
      mobjComm.Message cwiSTART_OLE_SERVER, NCID_PREVIOUS_ENCOUNTERS_APPLICATION, NCID_CW_APPLICATION, NCID_ENCOUNTER_APPLICATION
      Set mobjCurrEnc = Nothing
      Set mcolEncounters = New Collection
      OpenRegularView = False
      Call mobjComm.setstatustext("")
      Exit Function
  End Select
    
     
      
      
      '---set global flags---
      bFirstRegularResize = True
      bMainViewRunning = True

        '---add a menu item "EncounterParent" on CW's main menu---
      Call AddMenus
     
        '-- add a reference so the form can talk to me
      frmEncParent.ParentObject = Me
      
'-Mod 45e  Swapped these two calls.  Need to have the notes in place so children can add to them in amends
        '-- display any notes that we have in the db
      Call DisplayNotes
      
        '-- load all the child objects for the form
      Call LoadChildren
      
        '--set input parameter for handle back to form
      rhMainForm = frmEncParent.hwnd
      
      SetParent frmEncParent.hwnd, vlParenthWnd
      OpenRegularView = True
#If debugon Then
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "OpenRegularView", "End", DR
#End If

    Exit Function
    
errHandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "Enctr.OpenRegularView", "Enc", vbCritical)

End Function

'----------------------------------------------------------------------------
'- Purpose: get the encounter to be displayed into the mobjCurrEnc object and prepare to display it
'
' Input: ShowForm - Should we show the "encounter select" form to the user, or is there enough info
'                   to know which one enc to display?
' Output - True on success
'----------------------------------------------------------------------------
Private Function OpenViewGetEncounter(ShowForm As Boolean) As Boolean
  
#If debugon Then
  Dim DR As DebugRec
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "OpenViewGetEncounter", "Begin", DR, "ShowForm", ShowForm
#End If
  
  If Not DebugMode Then On Error GoTo errHandler
    
  If ShowForm Then
    Call GetEncounterInfo
    
    Debug.Print , "enc count: " & mcolEncounters.count
    frmEncSelect.ParentObject = Me
    frmEncSelect.Encounters = mcolEncounters
    frmEncSelect.Top = (Screen.Height - frmEncSelect.Height) / 3
    frmEncSelect.Left = (Screen.Width - frmEncSelect.Width) / 2
    
    mobjComm.Message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_ENCOUNTER_APPLICATION
    frmEncSelect.Show vbModal
    mobjComm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_ENCOUNTER_APPLICATION
    DoEvents
'-Mod 71e
    Set mcolEncounters = Nothing
    
    If mbSelectionCanceled Or Val(mobjPatient.EncounterNumber) = 0 Then
      Set mobjCurrEnc = Nothing
      OpenViewGetEncounter = False
      Exit Function
    End If
  
    mobjPatient.EncounterNumber = mobjCurrEnc.EncounterID
    mobjPatient.EncounterFacilityNCID = mobjCurrEnc.FacilityNCID
  
  Else  '- do not showform
    '- do i have to get the current enconter?
    Dim bNeedToGetCurrEnc As Boolean
    If mobjCurrEnc Is Nothing Then
      bNeedToGetCurrEnc = True
    '- check the curr enc info to see if it matches the patient info
    ElseIf Not (mobjCurrEnc.EncounterID = mobjPatient.EncounterNumber And mobjCurrEnc.FacilityNCID = mobjPatient.EncounterFacilityNCID) Then
      bNeedToGetCurrEnc = True
    End If
    
    If bNeedToGetCurrEnc Then
      Call mobjComm.setstatustext("Retrieving Encounter...")

      Set mobjCurrEnc = GetEncounter(mobjPatient.UnitNumber, mobjPatient.EncounterFacilityNCID, mobjPatient.EncounterNumber)
      Call mobjComm.setstatustext("")
    End If
    DoEvents
  
  End If  '- showform
  
'-- did we get someting? if not, get out of here
  If mobjCurrEnc Is Nothing Then
    MsgBox "Internal data error. Unable to open selected encounter.", vbOKOnly, "Encounter Summmary Error"
    Set mobjCurrEnc = Nothing
    OpenViewGetEncounter = False
    Exit Function
  End If
  
  
''-Mod 20e - handle the navigation based on the enc.status
'  Select Case mobjCurrEnc.Status
'
'  Case Complete, PatientLeftWithoutBeingSeen, Updated
'  '-- since I don't show these in the 'enc select' window, I must be being asked programmatically to display them.
'  '   Send the user to the prev enc module.
'  '-- First ensure the patient is pointing to this enc and launch prev enc, he'll open to it
'      mobjPatient.EncounterNumber = mobjCurrEnc.EncounterID
'      mobjPatient.EncounterFacilityNCID = mobjCurrEnc.FacilityNCID
'
'      mobjComm.Message cwiSTART_OLE_SERVER, NCID_PREVIOUS_ENCOUNTERS_APPLICATION, NCID_CW_APPLICATION, NCID_ENCOUNTER_APPLICATION
'      Set mcolEncounters = Nothing
'      MyState = es_Initialized
'      Set mobjCurrEnc = Nothing
'      Set mcolEncounters = New Collection
'      OpenViewGetEncounter = False
'      Exit Function
'
'  Case Updating
'  '-- we are being called from prev enc or enc cos
'  '   The rest of the program will handle this based on the status flag
'
'  End Select
'
  OpenViewGetEncounter = True
'~end mod 20e

#If debugon Then
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "OpenViewGetEncounter", "End", DR
#End If
  
    Exit Function
    
errHandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "Enctr.OpenViewGetEncounter", "Enc", vbCritical)
  
End Function


'*******************************************************************************************
'  Description: This method is called by the CW Core as a close command
'
'  Parameters:
'    viViewType - Contains the "type" of view that the CW Core is closing.
'*******************************************************************************************
Public Sub CloseView(ByVal viViewType As Integer)
Attribute CloseView.VB_Description = "CW's Mandated method to direct the closing of the Enctr module"

  If Not DebugMode Then On Error GoTo errHandler

#If debugon Then
  Dim DR As DebugRec
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "CloseView", "Begin", DR
#End If

  If viViewType = cwiALL_VIEWS Or viViewType = cwiREGULAR_VIEW Then
    
    If bMainViewRunning Then
      bMainViewRunning = False
      frmEncParent.bCloseMyself = False
      Unload frmEncParent
    End If
       
  End If
  If viViewType = cwiALL_VIEWS Or viViewType = cwiPATIENT_SUMMARY_VIEW Then
 
    If bPatientSummaryViewRunning Then
       bPatientSummaryViewRunning = False
        frmEncSOC.bCloseMyself = False
        Unload frmEncSOC
    End If
   
    
  End If      '- viViewType = cwiALL_VIEWS Or viViewType = cwiREGULAR_VIEW Then
  Call UnloadEnctr
'-- remove patient enc
  If Not mobjPatient Is Nothing Then
    mobjPatient.EncounterNumber = ""
  End If

#If debugon Then
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "CloseView", "End", DR
#End If
    
    Exit Sub
errHandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "Enctr.CloseView", "EncounterParent", vbExclamation)
    
    Exit Sub

End Sub

'*******************************************************************************************
'  Description: This method is called when the user has selected a menu item (from the CW
'               Core's main menu), which is "owned" by this OLE Server.
'
'               NOTE: This is an OPTIONAL method.
'
'  Parameters:
'    vlMenuId - Contains the menu id of the menu item the user selected.
'*******************************************************************************************
Public Sub MenuItemSelected(ByVal vlMenuId As Long)
Attribute MenuItemSelected.VB_Description = "As mandated by CW"

  Dim mEncChild As EncChild      '-- child module
  Dim ThisMenu  As I_EncMenu     '-- menu object

#If debugon Then
  Dim DR As DebugRec
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "MenuItemSelected", "Begin", DR
#End If
    
  If Not DebugMode Then On Error GoTo errHandler
   
   On Error Resume Next
   
   Set ThisMenu = mcolMenus(CStr(vlMenuId))
   If ThisMenu Is Nothing Then
      Err.Clear
      If Not DebugMode Then On Error GoTo errHandler

      For Each mEncChild In mcolEncChild             '- find the active one
        If mEncChild.RegularViewRunning Then
          mEncChild.Module.MenuItemSelected vlMenuId '- pass it on
          Exit For
        End If
      Next mEncChild
      Set mEncChild = Nothing
   Else     '- ThisMenu Is Nothing Then
     If Not DebugMode Then On Error GoTo errHandler
     ThisMenu.Selected
   End If   '- ThisMenu Is Nothing Then
   
#If debugon Then
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "MenuItemSelected", "End", DR
#End If

  Exit Sub

errHandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "Enctr.MenuItemSelected", "EncounterParent", vbExclamation)
    
    Exit Sub
End Sub

'**************************************************************************************************************
'  Description: This method gets called when a new patient has been selected from the CW Core.
'
'               NOTE: This is an OPTIONAL method.
'
'  Parameters:
'    01.1.00 WAS: added optioanl ForceRefresh for Dental to allow them to change the encounter and gui
'**************************************************************************************************************
Public Sub PatientSelected(Optional ForceRefresh As Boolean = False)
Attribute PatientSelected.VB_Description = "As mandated by CW"


  Dim frmwide As Long       '- current size of form
  Dim frmhigh As Long       '- current size of form
  Dim ldummy  As Long       '- dummy return var
  Dim oChild  As EncChild   '- child object from collection
  'Dim oTempE As ICHCSII_EncSum4
  
#If debugon Then
  Dim DR As DebugRec
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "PatientSelected", "Begin", DR, "Patient", mobjPatient.Name, "Enc", mobjPatient.EncounterNumber
#End If
  
  If Not DebugMode Then On Error GoTo errHandler
  
  Debug.Print "New Patient Selected:" & mobjPatient.Name & "/" & mobjPatient.UnitNumber & "   Enc:" & mobjPatient.EncounterNumber
  
  If Not ForceRefresh Then
    If mvarCurrPatientUnitNumber = mobjPatient.UnitNumber Then
      Exit Sub
    End If
  End If
  
  mvarCurrPatientUnitNumber = mobjPatient.UnitNumber

'-- tell the kids
'  For Each oChild In mcolEncChild
'    oChild.Module.PatientSelected
'  Next oChild
'  DoEvents
  
  If bPatientSummaryViewRunning Then
    Set mcolEncounters = New Collection
    If GetEncounterInfo() Then
      frmEncSOC.Encounters = mcolEncounters
      frmEncSOC.Display
     Else
      frmEncSOC.grdSOC.Clear
      frmEncSOC.grdSOC.rows = 1
     End If
  End If

'-- if no encounter, kill myself...
  If Val(mobjPatient.EncounterNumber) = 0 Then
    If bMainViewRunning Then
        bMainViewRunning = False
        frmEncParent.bCloseMyself = False   '- catch it later, below
        Unload frmEncParent
        If UnloadFailed Then Exit Sub

        mobjComm.CloseView cwiREGULAR_VIEW, mlProgID
        
        Set mobjCurrEnc = Nothing
        Set mcolEncounters = New Collection
        Exit Sub
    End If
  End If

  Call ReloadData

   Call mobjComm.setstatustext("")
#If debugon Then
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "PatientSelected", "End", DR
#End If
    
    
      Exit Sub
errHandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "Enctr.PatientSelected", "EncounterParent", vbExclamation)
    
    Exit Sub

End Sub

Friend Function ReloadData() As Boolean
Dim ldummy As Long

'-- redo the openview, and resize events.
Dim EncNum As String
Dim EncFacNum As String
Dim i As Integer

#If debugon Then
  Dim DR As DebugRec
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "ReloadData", "Begin", DR, "Main View Running", bMainViewRunning
#End If
 
  '- clear the form
  If bMainViewRunning Then
  frmEncParent.Visible = False
    With frmEncParent
      .cmds(0).Enabled = True
      .pics(0).Enabled = True
      .pics(0).Height = 500
      .RTBAC(0).Text = ""
      .RTBAC(0).Height = 500
      .RTBNote(0).Text = ""
      .RTBNote(0).Height = 500
      
      For i = .RTBNote.count - 1 To 1 Step -1
        Unload .RTBNote(i)
      Next i
      .picNote.Height = 500
      
      For i = 1 To 7
        .cmds(i).Enabled = True
        .pics(i).Enabled = True
        .RTB(i).Text = ""
        .RTB(i).Height = 500
        .pics(i).Height = 500
        DoEvents
      Next i
  
    End With
'-- clear all the current stuff
'-Mod 34e
    If Not mobjCurrEnc Is Nothing Then RaiseEvent CloseEncounter(mobjCurrEnc)

    Set mobjCurrEnc = Nothing   '-- this has to be done after the form unload so we can save the current enc
    Set mcolEncounters = New Collection
    
    bFirstRegularResize = True
    
    OpenView cwiREGULAR_VIEW, glParentWindow, ldummy
    frmEncParent.Display 0, 0, 0, 0
    frmEncParent.Visible = True
  End If
  
  ReloadData = True

#If debugon Then
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "ReloadData", "End", DR
#End If

End Function

'*******************************************************************************************
'  Description: This method is called when a change has occurred in the "client area"
'               used by this OLE Server, resulting in a need to resize a particular "view".
'
'  Parameters:
'    viViewType - Contains the "type" of view being requested to resize by the CW Core.
'    vlLeft - Contains the "left" coordinate for the OLE server.
'    vlTop - Contains the "top" coordinate for the OLE server.
'    vlWidth - Contains the "width" for the OLE server.
'    vlHeight - Contains the "height" for the OLE server.
'*******************************************************************************************
Public Sub ResizeView(ByVal viViewType As Integer, _
                      ByVal vlLeft As Long, _
                      ByVal vlTop As Long, _
                      ByVal vlWidth As Long, _
                      ByVal vlHeight As Long)
                      
  Dim oChild        As EncChild        '- which child object


#If debugon Then
  Dim DR As DebugRec
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "ResizeView", "Begin", DR
#End If


    If Not DebugMode Then On Error GoTo errHandler
    
    '---perform different actions based upon the "view type"---
    
    Select Case viViewType
    Case cwiREGULAR_VIEW
'-- not the first call, just a real resize
        If Not bFirstRegularResize Then
          If frmEncParent.Left = vlLeft _
            And frmEncParent.Top = vlTop _
            And frmEncParent.Width = vlWidth _
            And frmEncParent.Height = vlHeight Then
            #If debugon Then
              If Not DL Is Nothing Then DL.DebugLog Exe, Module, "ResizeView", "End WOA", DR
            #End If
            Exit Sub
          End If
        
            frmEncParent.WindowState = vbNormal
            
            '---hide the form to so that annyoing "screen blinking" will not be seen while controls are moved and resized---
            frmEncParent.Visible = False
            
            '---resize form to fill entire "client area", NOTE: This generates a "resize" event---
            frmEncParent.Move vlLeft, vlTop, vlWidth, vlHeight
            DoEvents
            
            For Each oChild In mcolEncChild
              If oChild.RegularViewRunning Then
                Call oChild.Module.ResizeView(cwiREGULAR_VIEW, 0, 0, frmEncParent.vseChildClient.Width, frmEncParent.vseChildClient.Height)
                Exit For
              End If
            Next oChild
            frmEncParent.Visible = True

        Else    '--first "regular" resize, need to perform "initailization" procedures--
            '---reset flag so this code will not be executed again---
            bFirstRegularResize = False
        
            '----------------------------------------------------------------------------
            '  NOTE: Here is where you "Show" the "Regular View" form for the FIRST time.
            '----------------------------------------------------------------------------
            frmEncParent.Display vlLeft, vlTop, vlWidth, vlHeight
            DoEvents
        End If
    

    Case cwiPATIENT_SUMMARY_VIEW
        If Not bFirstPatientSummaryResize Then
            frmEncSOC.WindowState = vbNormal
            frmEncSOC.Visible = False
            frmEncSOC.Move vlLeft, vlTop, vlWidth, vlHeight
            frmEncSOC.Visible = True
        Else
            bFirstPatientSummaryResize = False
            frmEncSOC.Display
            DoEvents
            frmEncSOC.WindowState = vbNormal
            frmEncSOC.Visible = False
            frmEncSOC.Move vlLeft, vlTop, vlWidth, vlHeight
            frmEncSOC.Visible = True

        End If
       
    End Select
#If debugon Then
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "ResizeView", "End", DR
#End If
    Exit Sub
    
errHandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "Enctr.ResizeView", "EncounterParent", vbExclamation)
    Exit Sub
End Sub

'-----------------------------------------------------
'Purpose:  Built-in class initialize
'
'Parameters: none
'
'Comments:  setup collections and global constants
'----------------------------------------------------
Private Sub Class_Initialize()

#If debugon Then
  Set DL = GetObject("", "dlog.cdebuglog")
  DL.debugStart
  Dim DR As DebugRec
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "Class_Initialize", "Begin", DR
#End If

    '---create shared object used to show VB errors, etc.---
    Set gobjShared = New CWShared
    Set mcolEncChild = New Collection
    Set mcolEncounters = New Collection
'    Set mobjCurrEnc = New Encounter  'placeholder for the current active enc
'        mobjCurrEnc.Init mobjUser
'    Set mobjWorkFlow = New EncWorkFlow
'    Set mobjWorkFlow.Enctr = Me
    Set mobjUtil = New EncUtil
    Set mcolMenus = New Collection
    Set mobjProperties = New EncProperties
    
    '---set the help file for this application---
    App.HelpFile = gobjShared.CWHelpFile
   
    glAppNCID(0) = NCID_MEDICATIONS_APPLICATION   '- psudo for autocite
    glAppNCID(1) = NCID_REASON_FOR_VISIT
    glAppNCID(2) = NCID_VITAL_SIGNS_APPLICATION
    glAppNCID(3) = NCID_BROWSER_APPLICATION
    glAppNCID(6) = NCID_ASSESMENT_AND_PLAN_APPLICATION
    glAppNCID(4) = NCID_LAB_APPLICATION
    glAppNCID(5) = NCID_RADIOLOGY_RESULTS_APPLICATION
    glAppNCID(7) = NCID_DISPOSITION_APPLICATION
    glAppNCID(8) = NCID_MEDCIN_APPLICATION
    
    Sect(0) = "AutoCite"
    Sect(1) = "Screening"
    Sect(2) = "Vitals"
    Sect(3) = "S/O"
    Sect(6) = "A/P"
    Sect(4) = "Labs"
    Sect(5) = "Rads"
    Sect(7) = "Disposition"
    Sect(8) = "ListTool"
    
    gsSecAreaNames(1) = "Health_History"                       '- screening
    gsSecAreaNames(2) = "Vital_Signs"
    gsSecAreaNames(3) = ""                        '- chart notes
    gsSecAreaNames(6) = "Encounter_Coding"
    gsSecAreaNames(4) = "Results_Retrieval_Non-sensitive"
    gsSecAreaNames(5) = "Results_Retrieval_Non-sensitive"
    gsSecAreaNames(7) = "Patient_Disposition"
    gsSecAreaNames(8) = ""
    
'------------------------------
   gsEncStatus(11) = "Complete"
   gsEncStatus(1) = "Incomplete"
   gsEncStatus(2) = "Needs Signature"
   gsEncStatus(3) = "Needs Co-Signature"
   gsEncStatus(4) = "Pending"
   gsEncStatus(5) = "Checked-In"
   gsEncStatus(6) = "Waiting For VS"
   gsEncStatus(7) = "Waiting"
   gsEncStatus(8) = "In Progress"
   gsEncStatus(9) = "Needs Disposition"
   gsEncStatus(10) = "Updated"
   gsEncStatus(12) = "Updating"
   gsEncStatus(13) = "Left Without Being Seen"
   
   
  'NCIDs for Domain: 14510304 (Encounter Status)
  '---------------------------------------------
  gsEncStatusNCID(1) = "75657"
  gsEncStatusNCID(2) = "14510305"
  gsEncStatusNCID(3) = "14510306"
  gsEncStatusNCID(4) = "662"
  gsEncStatusNCID(5) = "14510307"
  gsEncStatusNCID(6) = ""
  gsEncStatusNCID(7) = "14510308"
  gsEncStatusNCID(8) = "11992"
  gsEncStatusNCID(9) = ""
  gsEncStatusNCID(10) = "14510309"
  gsEncStatusNCID(11) = "205657"
  gsEncStatusNCID(12) = "14510310"
  gsEncStatusNCID(13) = "14508309"

   
   giEncApptStatus(11) = 4
   giEncApptStatus(1) = 0
   giEncApptStatus(2) = 0
   giEncApptStatus(3) = 5
   giEncApptStatus(4) = 0
   giEncApptStatus(5) = 1
   giEncApptStatus(6) = 0
   giEncApptStatus(7) = 2
   giEncApptStatus(8) = 3
   giEncApptStatus(9) = 0
   giEncApptStatus(10) = 6
   giEncApptStatus(12) = 0
   giEncApptStatus(13) = 7
   
   gsEnctype(0) = "In-Patient"
   gsEnctype(1) = "Out-Patient"
   gsEnctype(2) = "Ambulatory"
   gsEnctype(3) = "TeleCons"
   gsEnctype(4) = "Dental"
   
   gsRTFStatus(1) = "Incomplete"
   gsRTFStatus(3) = "Needs Co-Signature"
   gsRTFStatus(4) = "Pending"
   gsRTFStatus(10) = "Updated"
   gsRTFStatus(12) = "New"
   gsRTFStatus(11) = "Complete"
   
   
 MENU_TXT(0) = "&Close"
 MENU_BMP(0) = "close.bmp"
 MENU_TIP(0) = "Close Encounter Summary"
 MENU_NAME(0) = "Close"

 MENU_TXT(1) = "&Sign"
 MENU_BMP(1) = "dispsign.bmp"
 MENU_TIP(1) = "Sign the completed encounter"
 MENU_NAME(1) = "Sign"

 MENU_TXT(2) = "&Add Note"
 MENU_BMP(2) = "new.bmp"
 MENU_TIP(2) = "Add A Note to Encounter"
 MENU_NAME(2) = "AddNote"

 MENU_TXT(3) = "Change Encounter"
 MENU_BMP(3) = "patenct.bmp"
 MENU_TIP(3) = "Change to a different Encounter"
 MENU_NAME(3) = "Change"

 MENU_TXT(4) = "Close Encounter Sub-Section"
 MENU_BMP(4) = "close.bmp"
 MENU_TIP(4) = "dispClose the displayed module"
 MENU_NAME(4) = "CloseChild"

 MENU_TXT(5) = "Add Providers"
 MENU_BMP(5) = "link.bmp"
 MENU_TIP(5) = "Add Providers"
 MENU_NAME(5) = "Add Prov"

 MENU_TXT(6) = "Print"
 MENU_BMP(6) = "print.bmp"
 MENU_TIP(6) = "Print this encounter document"
 MENU_NAME(6) = "Print"

 MENU_TXT(7) = ""
 MENU_BMP(7) = ""
 MENU_TIP(7) = ""
 MENU_NAME(7) = ""

 MENU_TXT(8) = "Templates"
 MENU_BMP(8) = "edit.bmp"
 MENU_TIP(8) = "Select an Encounter Template"
 MENU_NAME(8) = "Template"

 MENU_TXT(9) = "Next Screen"
 MENU_BMP(9) = "forward.bmp"
 MENU_TIP(9) = "Navigate to next screen"
 MENU_NAME(9) = "Next"

 MENU_TXT(10) = "Refresh"
 MENU_BMP(10) = "refresh.bmp"
 MENU_TIP(10) = "Refresh Screen"
 MENU_NAME(10) = "Refresh"

 MENU_TXT(11) = "Preview"
 MENU_BMP(11) = "reorder.bmp"
 MENU_TIP(11) = "Preview this encounter document"
 MENU_NAME(11) = "Preview"

    #If debugon Then
      If Not DL Is Nothing Then DL.DebugLog Exe, Module, "Class_Initialize", "End", DR
    #End If

End Sub

'-----------------------------------------------------
'Purpose: Built-in class terminate
'
'Parameters: none
'
'Comments: destroy references to objects
'----------------------------------------------------
Private Sub Class_Terminate()

    '---release reference to objects---
'-Mod 4e
    mobjProperties.ResetParentObj   '- get rid of this reference
    Set mcolEncounters = Nothing
    Set mcolEncChild = Nothing
'-Mod 34e
    If Not mobjCurrEnc Is Nothing Then RaiseEvent CloseEncounter(mobjCurrEnc)
    
    Set mobjCurrEnc = Nothing
    Set mcolMenus = Nothing
    Set mobjPatient = Nothing
    Set mobjLogonEnvironment = Nothing
    Set mobjComm = Nothing
    Set mobjWorkFlow = Nothing
    Set mobjProperties = Nothing
    Set gobjShared = Nothing
    Set mobjComm = Nothing
    Set mobjLogonEnvironment = Nothing
    Set mobjConfig = Nothing
    Set mobjPatient = Nothing
    Set mobjUser = Nothing
    Set mobjUtil = Nothing
    
    
#If debugon Then
  Dim DR As DebugRec
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "Class_Terminate", "End", DR
  Set DL = Nothing
#End If
    
End Sub

'**************************************************************************************************************
'  Description: This method is used by the CW Core to send messages to this OLE Server.  The
'               messages can be from the CW Core or from another OLE Server.
'
'  Parameters:
'    viMessageType - Contains the "type" of message to process.
'    vsMessageText - Contains any "data" assoicated with this message.
'    vsSourceServerNCID - Contains the NCID of the OLE Server that "sent" the message.
'**************************************************************************************************************
Public Sub Message(ByVal viMessageType As Integer, ByVal vsMessageText As String, ByVal vsSourceServerNCID As String)
Attribute Message.VB_Description = "As mandated by CW"
  
  Dim buff()      As String           '- buffer to hold parts of message
  Dim DisplayMode As Boolean          '- for AddNote: display the note editor or not
  Dim oChild      As EncChild         '- which child module to address with message
  Dim oMenu       As I_EncMenu        '- which menu item to address
  Dim FoundIt     As Boolean          '- loop flag
  
#If debugon Then
  Dim DR As DebugRec
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "Message", "Begin", DR, "Msg Type", viMessageType, "Message", vsMessageText
#End If
    
    If Not DebugMode Then On Error GoTo errHandler
    
    Select Case viMessageType
    
    Case cwiCLOSE_OLE_SERVER
        '--this server has just shut down. so close out the sub application if it's yours
        If mcolEncChild.Item(1).ProgID = vsMessageText Then
          CloseEncChild (vsMessageText)
          frmEncParent.vseChildClient.Visible = False
          frmEncParent.VP1.Visible = True
        End If
    Case cwiREFRESH, cwiREFRESH_DATA
      '- pass it down to all the kids
'-mod 27e handled by core
'      For Each oChild In mcolEncChild
'        If oChild.ModuleLoaded Then
'          Call oChild.Module.Message(viMessageType, vsMessageText, vsSourceServerNCID)
'        End If
'      Next
'-- what if all the kids haven't yet been refreshed????
      'Call LoadChildren
      
    Case cwiPROPERTIES  '-- show the properties page
      Set mobjProperties.Parent = Me
      mobjProperties.Display
'-Mod 38e
      mobjProperties.ResetParentObj
    Case cwiENCOUNTER_ADD_NOTE
      buff = Split(vsMessageText, "|")
      If UBound(buff) <> 2 Then
        vsMessageText = "Error. Too Few Parameters in message ('" & vsMessageText & "')"
        Exit Sub
      End If
      If UCase$(buff(2)) = "TRUE" Then
        DisplayMode = True
      Else
        DisplayMode = False
      End If
       Call AddEncounterNote(buff(0), buff(1), DisplayMode)
       
    Case cwiGENERIC       '-- parse the message to see what needs to be done
       Call MessageGeneric(vsMessageText)
       
    Case cwiPRINT_REGULAR_VIEW
       Call mobjCurrEnc.PrintEncounter(False)
       
    Case cwiPRINT_PREVIEW
       Call mobjCurrEnc.PrintEncounter(True)
       
    Case Else   '--do nothing--
      vsMessageText = "Unknown message type"
    
    End Select
    
    Set oChild = Nothing
    
#If debugon Then
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "Message", "End", DR, "Message", vsMessageText
#End If

    Exit Sub
    
errHandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "Enctr.Message", "EncounterParent", vbExclamation)
    
    Exit Sub
End Sub

Private Sub MessageGeneric(vsMessageText As String)
  Dim buff() As String
  Dim oChild As EncChild
  Dim oChild2 As EncChild
  Dim oMenu As I_EncMenu
  Dim FoundIt As Boolean
  Dim i As Integer
  Dim txt As String
  
  buff = Split(vsMessageText, "|")
  Select Case UCase$(buff(0))
  Case "TREESELECTED"  '-- someone clicked on a child of mine in a tree view at core
   If buff(1) = "" Then Exit Sub   '-- dummy check - should contain an appNCID
   '--else
  
  Case "VIEWOPENED"
    '- buff(1) holds the app ncid of the form.  See if
    Call CoreTookMyForm(buff(1))
    
  Case "AMENDENC"
    Call AmendEnc
    
  Case "OPENCHILD"
    '  part 2 should be the ncid of the child to open
    If Not bMainViewRunning Then Exit Sub  'need the frmencparent
    If buff(1) = "" Then Exit Sub
    For Each oChild In mcolEncChild
     If oChild.AppNCID = buff(1) Then
       Call OpenEncChild(oChild.ServerIndex)
       Exit For
     End If
    Next oChild
    
  Case "MENU"          '  MENU|PRINTINCOMPLETE|enable
    If UBound(buff) <> 2 Then Exit Sub
    Select Case UCase$(buff(1))
    Case "PRINTINCOMPLETE"
      For Each oMenu In mcolMenus
        If oMenu.MenuName = MENU_NAME(M_PRINT) Then  '- print
         If UCase$(buff(2)) = "ENABLE" Then
           oMenu.Enable
         Else
           oMenu.Disable
         End If
         Exit For
        End If
      Next oMenu
      Set oMenu = Nothing
    Case Else
    End Select
  
  Case "LAUNCH"
    '-- open child in buff(1), send rest of string to them
    If Not bMainViewRunning Then Exit Sub  'need the frmencparent
    If buff(1) = "" Then Exit Sub
    
'-Mod 49e  see if we're talking about charnotes or medcin then hit the s/o button
    If buff(1) = NCID_BROWSER_APPLICATION Or buff(1) = NCID_MEDCIN_APPLICATION Then
      If Not frmEncParent.cmds(SECT_SO).Enabled Then Exit Sub
      frmEncParent.cmds(SECT_SO).Value = True
      Exit Sub
    End If
    
    For Each oChild In mcolEncChild
      If oChild.AppNCID = buff(1) Then
      
       If Not frmEncParent.cmds(oChild.ServerIndex).Enabled Then Exit Sub
       
        If Not oChild.RegularViewRunning Then
          For Each oChild2 In mcolEncChild
            If oChild2.RegularViewRunning And oChild2.AppNCID <> oChild.AppNCID Then Call oChild2.Module.CloseView(cwiREGULAR_VIEW)
          Next oChild2

          Call OpenEncChild(oChild.ServerIndex)
          DoEvents
        End If
        
         If UBound(buff) > 1 Then
           '-- do a join Something to pass ALL elements, from 2 to end of array, to child
'           txt = ""
'           For i = 2 To UBound(buff)
'             txt = txt & "|" & buff(i)
'           Next i
           
           If oChild.RegularViewRunning Then
             oChild.Module.Message cwiGENERIC, "LAUNCH" & "|" & buff(2), NCID_ENCOUNTER_APPLICATION
            ' oChild.Module.Message cwiGENERIC, "LAUNCH" & txt, NCID_ENCOUNTER_APPLICATION
           End If
         End If     '- UBound(Buff) > 1 Then
         Exit For
     
     End If         '- oChild.AppNCID = Buff(1) Then
    Next oChild
  
  Case "CLOSEMODULE"   '- the little x in the corner of the core frame was hit; close child
'--12.29.99 WAS - added a search for running module

'-Mod 42e First things: check to see if chart notes is up. if so, just close it.
    If mcolEncChild(CStr(SECT_SO)).RegularViewRunning Then
      mcolEncChild(CStr(SECT_SO)).Module.CloseView (cwiREGULAR_VIEW)
      Exit Sub
    End If
'~end mod 42e

    For Each oChild In mcolEncChild
     If oChild.RegularViewRunning Then
       Call oChild.Module.CloseView(cwiREGULAR_VIEW)
       Exit For
     End If
    Next oChild
    
    If oChild Is Nothing Then
       If bMainViewRunning Then
''-Mod 65e Give the user a chance to not save the current enc
'-------CAN NOT do this.  data is already saved to CDR by children

'         If mobjCurrEnc.FlagChanged Then
'          Select Case MsgBox("Save the current encounter document?", vbQuestion + vbYesNoCancel + vbDefaultButton1, "Close Encounter")
'          Case vbCancel
'            Exit Sub
'          Case vbNo
'            mobjCurrEnc.FlagChanged = False    '- keeps from saving
'          End Select
'        End If
''~ end mod 65e
         bMainViewRunning = False
         frmEncParent.bCloseMyself = True
         Unload frmEncParent
       End If
       
       If UnloadFailed Then
         bMainViewRunning = True
         Exit Sub
       End If
     '-- remove patient enc
       If Not mobjCurrEnc Is Nothing Then
         mobjCurrEnc.Patient.EncounterNumber = ""
       End If
    Else
      Set oChild = Nothing
    End If
    
  Case Else
    vsMessageText = "Unknown message type"
  End Select

End Sub


'****************************************************************************************************
'  Description: This method is called when the CW Core creates an "instance" of this OLE Server.  It
'               is used to create and initialize global variables that are shared across all
'               "views" supported by this OLE Server, i.e. variables that only need to be
'               initialized ONCE, upon object creation.
'
'  Parameters:
'    vlMyProgId - Contains the "ProgId" that the CW Core has assigned to this OLE Server.
'    robjComm - Contains a communications object that is used to handle all communication
'               between this OLE Server and the CW Core.
'    robjLogonEnvironment - Contains the logon environment object.  Contains information such
'                           as who logged on, where they logged on, etc.
'    robjConfig - Contains the configuration object that is used to get and set configuration
'                 information for a specific application.
'    robjPatient - Contains the patient object containing the currently selected patient.
'****************************************************************************************************

Public Sub Initialize(ByVal vlMyProgId As Long, _
                            robjComm As Object, _
                            robjLogonEnvironment As Object, _
                            robjConfig As Object, _
                            robjPatient As Object)

  Dim oChild As EncChild
  'Dim oTempE As ICHCSII_EncSum4
  
#If debugon Then
  Dim DR As DebugRec
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "Initialize", "Begin", DR, "Initialize", robjPatient.Name, "Enc", robjPatient.EncounterNumber
#End If


  Debug.Print "Initialize :" & robjPatient.Name & "   Enc:" & robjPatient.EncounterNumber
  
    If Not DebugMode Then On Error GoTo errHandler
        
'-Mod 32e - Instantiate Help-file object
    App.HelpFile = gobjShared.CWHelpFile
           
'---save the "ProgId" assigned to me by the CW Core---
    mlProgID = vlMyProgId
    
'---create references to passed in objects---
    Set mobjComm = robjComm
    'Set cwobjCWComm = robjComm
    
'-- core calls initialize every time the tree is  hit.  So don't if I'm still alive and well
'   and using the proper patient.
    If Not (mobjLogonEnvironment Is Nothing Or mobjConfig Is Nothing Or mobjPatient Is Nothing) Then
        If mobjLogonEnvironment = robjLogonEnvironment _
          And mobjConfig = robjConfig _
          And mobjPatient = robjPatient Then Exit Sub
    End If
'    Set cwobjLogonEnviron = robjLogonEnvironment
    Set mobjLogonEnvironment = robjLogonEnvironment
    Set mobjConfig = robjConfig
    Set mobjPatient = robjPatient
    mvarCurrPatientUnitNumber = mobjPatient.UnitNumber

    Call mobjComm.setstatustext("Retrieving User Preferences...")
    Set mobjProperties.Parent = Me
    mobjProperties.RetrieveProperties
'-Mod 38e
    mobjProperties.ResetParentObj
'-mod 11e trying to get a handle on the persisting enc data (one enc to anothing)
'    If Val(mobjPatient.EncounterNumber) > 0 Then
'      Set mobjCurrEnc = GetEncounter(mobjPatient.UnitNumber, mobjPatient.EncounterFacilityNCID, mobjPatient.EncounterNumber)
'    End If


    '---set global flags---
    bMainViewRunning = False
    bPatientSummaryViewRunning = False

#If debugon Then
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "Initialize", "End", DR, "vlMyProgId", vlMyProgId
#End If
    Call mobjComm.setstatustext("")
    Exit Sub

errHandler:
    '---show the VB error---
    
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "Enctr.Initialize", "EncounterParent", vbExclamation)
    Call mobjComm.setstatustext("")
    Exit Sub
End Sub

'**************************************************************************************************************
'  Description: This method is called by a client when they want to modify the "properties", or the
'               configuration of this OLE Server.
'
'               NOTE: This is an OPTIONAL method.
'
'  Parameters:
'    robjLogonEnvironment - Contains the logon environment object.
'    robjConfig - Contains the configuration object that is used to get and set configuration
'                 information for a specific application.
'    vsOwnerNCID - Contains the NCID of the "owner" to use in storing modifications to the configuration.
'    vbCalledByDesigner - TRUE if this method is being called by the CW Designer, FALSE otherwise.
'**************************************************************************************************************
Public Sub Properties(robjLogonEnvironment As Object, _
                      robjConfig As Object, _
                ByVal vsOwnerNCID As String, _
                ByVal vbCalledByDesigner As Boolean)

    If Not DebugMode Then On Error GoTo errHandler
    
    
    
    Exit Sub
    
errHandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "Enctr.Properties", "EncounterParent", vbExclamation)
    
    Exit Sub
End Sub


'-----------------------------------------------------
'Purpose: opens (full-screen) an enc child module
'
'Parameters:  idx            - the index of the child to launch (based on form cmd button index)
'             CallerAppNCID  - who called this? used to determine location to send data on closeEncChild event
'             CNDataID       - ChartNotes data id (to setup chartnotes if needed)
'             CNTemplate     - ChartNotes template number (to setup chartnotes if needed)
'
'Returns:  True on success
'
'Comments:  Calls the openView/resizeView methods in the requested child object
'----------------------------------------------------
Friend Function OpenEncChild(idx As Integer, _
                    Optional CallerAppNCID As String, _
                    Optional CNDataID As String, _
                    Optional CNTemplate As String) As Boolean
  
  Dim NewMenu   As I_EncMenu   '- temp menu object
  Dim mEncChild As EncChild    '- which child module
  Dim hwnd      As Long        '- window handle
'  Dim oTempE    As ICHCSII_EncSum '- temporary object to access private methods

  If Not DebugMode Then On Error GoTo errHandler

#If debugon Then
  Dim DR As DebugRec
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "OpenEncChild", "Begin", DR, "Idx", idx
#End If


  Set mEncChild = mcolEncChild(CStr(idx))
 '-- tell it to openview and resize
  If mEncChild.ModuleLoaded Then
    mEncChild.CallerAppNCID = CallerAppNCID
      
'--1.6.00 WAS: told to remove my menu items with a child is up
   Call RemoveMenus
   
'-- per request from Robert Bell (vitals, screening, et al) this has been moved to before openview
'-- just to be sure your 'daddy' is set... and take it down in closeencchild
   Call mEncChild.SetParentObject(Me)   '-- me is enctr obj
        
    If mEncChild.Module.OpenView(cwiREGULAR_VIEW, frmEncParent.vseChildClient.hwnd, hwnd) Then
'--just to make Jonathan happy...
      frmEncParent.vseChildClient.Visible = True
      
      Call mEncChild.Module.ResizeView(cwiREGULAR_VIEW, 0, 0, frmEncParent.vseChildClient.Width, frmEncParent.vseChildClient.Height)
      mEncChild.RegularViewRunning = True
      
      If idx = SECT_SO Then
        
'--01.18.00 WAS - set autosave timer interval
          mEncChild.Module.Message cwiGENERIC, "AUTOSAVE|" & mobjProperties.AutoSaveInt, NCID_ENCOUNTER_APPLICATION
          
'-Mod 36e  No Enc-Related if called from other than s/o
          If CallerAppNCID = "" Then
            Call mEncChild.Module.UpdateEncounterInfo(mobjCurrEnc.StartDTS, _
                                    mobjLogonEnvironment.UserNcid, _
                                    mobjLogonEnvironment.FacilityNCID, _
                                    "", _
                                    mobjCurrEnc.EncounterID, _
                                    mobjCurrEnc.FacilityNCID, _
                                    "")
          End If
         '---send a message to the "Browser" module to edit the selected datastring---
          If CNDataID <> "" Or CNTemplate <> "" Then
            DoEvents
            mEncChild.Module.Message cwiEDIT_STRING, CNDataID & "|" & CNTemplate, NCID_ENCOUNTER_APPLICATION
            DoEvents
          End If
      End If   '-- idx = SECT_SO Then
'~end ~~~~~~
      
      OpenEncChild = True
      
    Else

Dim oTemp4    As ICHCSII_EncSum4
'-Mod 4 (in ChartNotes) "type mismatch in open view of chartnotes"
'-- reset the parent obj on the child
'      If TypeOf mEncChild.Module Is ICHCSII_EncSum Then
'        Set oTemp4 = mEncChild.Module
'        Call oTemp4.ResetParentObject
'        Set oTemp4 = Nothing
'      End If
      mEncChild.ResetParentObject
      frmEncParent.vseChildClient.Visible = False
      frmEncParent.VP1.Visible = True
      OpenEncChild = False
      frmEncParent.vseChildClient.ZOrder 1
      Exit Function
    
    End If
  End If   '--mEncChild.ModuleLoaded

  frmEncParent.vseChildClient.ZOrder 0
  frmEncParent.vseChildClient.Visible = True
  frmEncParent.VP1.Visible = False
  
  OpenEncChild = True
  
#If debugon Then
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "OpenEncChild", "End", DR, "Idx", idx
#End If
  
  Exit Function
  
errHandler:
      '---show the VB error---
  Call gobjShared.ShowVBError(Err.Number, Err.Description, "Enctr.OpenEncChild: " & idx, "EncounterParent", vbExclamation)
  
  frmEncParent.vseChildClient.Visible = False
  frmEncParent.VP1.Visible = True
  frmEncParent.vseChildClient.ZOrder 1
  
  OpenEncChild = False
  Exit Function

End Function

'-----------------------------------------------------
'Purpose:
'
'Parameters:
'
'Returns:
'
'Comments:
'----------------------------------------------------
Public Sub CloseEncChild(ProgID As Long)
Attribute CloseEncChild.VB_Description = "Called by Enctr's child modules to inform Enctr that they are closing their data entry form.  Similar to CW's CloseView method."
  
  Call CloseEncChildEx(ProgID)

End Sub

'-----------------------------------------------------
'Purpose: child module has closed and is telling me so
'
'Parameters: child module's program id (from cw)
'
'Returns:  True on success
'
'Comments: called by the regular view child to tell me he's closing and I need to clean up
'-- 01.01.00 WAS: made this a function to hold the child till I get what I need from him
'----------------------------------------------------
Public Function CloseEncChildEx(ProgID As Long) As Boolean
  
  If Not DebugMode Then On Error GoTo errHandler

#If debugon Then
  Dim DR As DebugRec
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "CloseEncChild", "Begin", DR, "ProgID", ProgID
#End If

  Dim DataDate   As Date             '- data the data was created
  Dim DataOwner  As String           '- who created the data
  Dim DataStatus As String           '- complete or not
  Dim CaseSuccess As Boolean         '- flag for success in the case stmt
  Static NoRecurs As Boolean         '- don't come in here twice
  Dim oMenu      As I_EncMenu        '- menu object
  Dim oEncChild  As EncChild         '- which child module
  Dim oTempc     As ICHCSII_EncSum2  '- temp object to talk to private methods
  Dim oTempE4    As ICHCSII_EncSum4  '- opposite of _encsum
  Dim RTFText    As String           '- the text string
  Dim oConcept   As ConceptCtrl
  Set oConcept = New ConceptCtrl
  
  If NoRecurs Then Exit Function
  NoRecurs = True
  
  '-- take down the 'close child' and seperator menu items
  'For Each oMenu In mcolMenus
  '  If oMenu.MenuName = MENU_NAME(M_CLOSECHILD) Then
  '    oMenu.Delete
  '    mcolMenus.Remove CStr(oMenu.MenuID)
  '    Exit For
  '  End If
  'Next oMenu
  'DoEvents
  
  For Each oMenu In mcolMenus
    If oMenu.MenuName = MENU_NAME(M_SEP) Then
      oMenu.Delete
      mcolMenus.Remove CStr(oMenu.MenuID)
      Exit For
    End If
  Next oMenu
  DoEvents
  For Each oMenu In mcolMenus
    If oMenu.MenuName = MENU_NAME(M_NAV) Then
      oMenu.Delete
      mcolMenus.Remove CStr(oMenu.MenuID)
      Exit For
    End If
  Next oMenu
  DoEvents
  
'-Mod 42e we hve to move this down a bit.  See below
  'Call AddMenus

Dim FoundIt As Boolean

  CaseSuccess = False
  
  '-- get the child that just closed and ask for new information.
  For Each oEncChild In mcolEncChild
    If oEncChild.ProgID = ProgID Then
      FoundIt = True
      Exit For
    End If
  Next oEncChild
  If Not FoundIt Then
    frmEncParent.lblChild.Caption = ""
    frmEncParent.vseChildClient.Visible = False
    frmEncParent.VP1.Visible = True
    CloseEncChildEx = False
    Exit Function
  End If
  
  If Not (oEncChild.ServerIndex = SECT_SO) Then
   
    RTFText = oEncChild.GetRTFNote()
    DataDate = oEncChild.GetDataDate()
    DataStatus = oEncChild.GetStatus()
    DataOwner = oEncChild.GetDataOwner()
    If Val(DataOwner) > 0 Then
      oConcept.UniqueId = DataOwner
      DataOwner = oConcept.PrefRep("2000").Representation
    End If
  End If
  
  If Not RTFText = "" Then
    ' If the user saved any data, and if the encounter was just checkedin, then
    ' update the status ourselves
    If mobjCurrEnc.Status = CheckedIn Then
      If oEncChild.ServerIndex = SECT_VITALS _
      Or oEncChild.ServerIndex = SECT_RFV Then
        mobjCurrEnc.Status = Waiting
      Else
        mobjCurrEnc.Status = InProgress
      End If
    End If
  End If
  
  frmEncParent.lblStatus.Caption = gsEncStatus(CInt(mobjCurrEnc.Status))
  If mobjCurrEnc.Status = CheckedIn Then
    frmEncParent.lblDate = "(Not Yet Started)"
  Else
    frmEncParent.lblDate.Caption = mobjUtil.FormatDTG(mobjCurrEnc.StartDTS)
  End If

'--012200 WAS - added to let vitals know he's free (he's used outside of me)
  If oEncChild.ServerIndex = SECT_VITALS Then
    Call oEncChild.ResetParentObject
  End If

  oEncChild.RegularViewRunning = False
  sChildMenuIDs = ""
  
'-moved from below
    #If debugon Then
      If Not DL Is Nothing Then DL.DebugLog Exe, Module, "CloseEncChild", "Mid", DR, "Server", oEncChild.ServerName, _
"Text", RTFText, "Date", DataDate, "Owner", DataOwner, "Status", DataStatus
    #End If
    
'-- handle special cases based on who it is
  Select Case oEncChild.ServerIndex
  Case SECT_VITALS
    
  Case SECT_AP
    Call RetrieveDiagnoses(oEncChild.Module, mobjCurrEnc)
    
  Case SECT_DISP
'-01.08.00 WAS: Disp is going to push this to me whenever it needs to be changed.
'     Call RetrieveDisposition(oEncChild.Module, mobjCurrEnc)
  Case SECT_SO
    If oEncChild.CallerAppNCID = "" Then
      Call RefreshSO(mobjCurrEnc)
'-Mod 42e
    Else
      '- just get out of here, whomever called cn is lying just below it and should be on top now
      NoRecurs = False
      CloseEncChildEx = True
      Exit Function
'~end mod 42
    End If
  Case SECT_MEDCIN
      Call RefreshMedcin(RTFText, DataDate, DataOwner)
  End Select
  
'-Mod 36e  Do not add this text to the enc if it is Sect_so and called from another app
'  If oEncChild.CallerAppNCID = "" Then
'--01.14.00 WAS: Added labs and rads exception. these will/may have already pasted data into their sectoins
    If Not (oEncChild.ServerIndex = SECT_SO Or oEncChild.ServerIndex = SECT_MEDCIN) Then    '-- dropped the labs and rads part
      Call ReplaceEncounterData(oEncChild.AppNCID, RTFText, DataDate, DataStatus, mobjUser, DataOwner)
    End If
'  Else
'    Call AddToEncounterData(oEncChild.CallerAppNCID, RTFText, DataDate, DataStatus, mobjUser, DataOwner)
'  End If
  
  
'-Mod 42e moved from above
  Call AddMenus
  
  NoRecurs = False

#If debugon Then
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "CloseEncChild", "End", DR, "ProgID", ProgID
#End If
  
  frmEncParent.lblChild.Caption = ""
  frmEncParent.vseChildClient.Visible = False
  frmEncParent.VP1.Visible = True
    
  mobjCurrEnc.FlagChanged = True
  CloseEncChildEx = True
  Exit Function

errHandler:
      '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "Enctr.CloseEncChild", "EncounterParent", vbExclamation)
    NoRecurs = False
    Exit Function

End Function

'-----------------------------------------------------
'Purpose:  Get the diagnoses from A&P
'
'Parameters: oANP   - the a&p object
'            oEnc   - the current encounter
'
'Returns: true on success
'
'Comments: Stored in a collection that is later pushed into semantic links in enc event model
'----------------------------------------------------
Friend Function RetrieveDiagnoses(oANP As Object, _
                                  oEnc As Encounter) As Boolean

  Dim i      As Integer       '- generic counter
  Dim Obj    As Object        '- generic object holder
  Dim Diag() As String        '- diagnosis array
  Dim Pri()  As Integer       '- priority flag
  Dim oDiag  As EncDiagnosis  '- diagnosis object
  
  
#If debugon Then
  Dim DR As DebugRec
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "RetrieveDiagnoses", "Begin", DR, "EncID", oEnc.EncounterID
#End If
  
  i = oANP.DiagnosisCollection.count
  
  #If debugon Then
    If Not DL Is Nothing Then DL.DebugLog Exe, Module, "RetrieveDiagnoses", "Mid", DR, "Diagnosis Count", i
  #End If
  
  If i = 0 Then Exit Function

  Set oEnc.RelatedDiagnoses = New Collection
  
  For Each Obj In oANP.DiagnosisCollection
    Set oDiag = New EncDiagnosis
    oDiag.NCID = Obj.NCID
    oDiag.DataID = Obj.DataID
    oDiag.Priority = Obj.Priority
    oEnc.RelatedDiagnoses.Add oDiag
  Next Obj
    
 '-- 01.02.00 WAS: per conversaion with Jon, they will come back in order.
  '-- the primary diag
  
  oEnc.PrimaryDiagnosisNCID = oEnc.RelatedDiagnoses(1).NCID
  
#If debugon Then
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "RetrieveDiagnoses", "end", DR, "PrimaryDiag", oEnc.PrimaryDiagnosisNCID
#End If
  
  Set oDiag = Nothing
  Exit Function
errHandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "Enctr.RetrieveDiagnoses ", "Enc", vbCritical)
    Set oDiag = Nothing

    Exit Function



End Function

'-----------------------------------------------------
'Purpose:  Get disposition information from Disp obj
'          NOTE:  OBE: disp is responsibile for pushing this data to me
'
'----------------------------------------------------
Public Function RetrieveDisposition(oDisp As Object, _
                                    oEnc As Encounter) As Boolean
  
#If debugon Then
  Dim DR As DebugRec
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "RetrieveDisposition2", "Begin", DR, "EncID", oEnc.EncounterID
#End If
  
  oEnc.DispositionText = oDisp.disposition.disposition
  oEnc.DispositionNCID = oDisp.disposition.dispositionnum
  oEnc.DispAdminOption = oDisp.disposition.AdminOptions
  oEnc.EandMNCID = oDisp.disposition.emcode
  
    Exit Function
errHandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "Enctr.RetrieveDisposition2 ", "Enc", vbCritical)

    Exit Function

#If debugon Then
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "RetrieveDisposition2", "End", DR, "EncID", oEnc.EncounterID
#End If



End Function

'-----------------------------------------------------
'Purpose: Read ChartNotes data from CDR and push to form
'
'Parameters:  oEnc - current encounter object
'
'Returns:  True on success
'
'Comments: Get a new set of textnote objects and compare their contents.
' If the status is Compete, Updated, Updating, then add a note describing the change as:
'     "This is what it was before it was changed."
'----------------------------------------------------
Friend Function RefreshSO(oEnc As Encounter) As Boolean
  
  Dim FoundIt           As Boolean      '- flag for searching loop
  Dim OutString         As String       '- RTF output string
  Dim TNotes            As Collection   '- textnote object collection
  Dim i                 As Integer      '- generic counter
  Dim msgFooter         As String       '- generic RTF lead-in
  Dim msgHeader         As String       '- generic RTF closing
  Dim newCounted()      As String       '- control modified v. new strings
  Dim oClinicalText     As ClinicalText '- DAS object subtype
  Dim oConcept          As ConceptCtrl  '- DAS object subtype
  Dim oDas2             As DasCtrl      '- DAS object
  Dim oDataString       As DataString   '- DAS object subtype
  Dim oTNote            As TextNote     '- from DAS
  Dim oldCounted()      As String       '- control modified v. new strings
  Dim oldTNote          As TextNote     '- from DAS
  Dim rasTextStrings()  As String       '- text as it returns from "convert HTML to text" function
  Dim rnNumTextStrings  As Integer      '- Number of above

#If debugon Then
  Dim DR As DebugRec
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "RefreshSO", "Begin", DR, "EncID", oEnc.EncounterID
#End If

  Set oDas2 = New DasCtrl
  Set TNotes = New Collection
  
  
  Set oConcept = New ConceptCtrl
  
  
  OutString = ""
  msgHeader = "{\rtf1\ansi\ansicpg1252\uc0\deff0" _
      & "{\fonttbl{\f0\fswiss\fcharset0\fprq2 Arial;}}" _
      & "{\colortbl;\red0\green0\blue0;}" _
      & "\deftab720 \plain "
  msgFooter = "}"
  
  oDas2.ClearAll
  oDas2.GeneralCriteria.ClearAll
  oDas2.GeneralCriteria.UnitNumber = mobjPatient.UnitNumber
  oDas2.GeneralCriteria.ReturnMultiMedia = False
  oDas2.GeneralCriteria.QueryQualifier = "EVENT_TYPE_NCID = '" & NCID_DAS_CLINICAL_TEXT & "'"
  oDas2.GeneralCriteria.EncounterNumber = oEnc.EncounterID
  oDas2.GeneralCriteria.Facility.ValueConcept = oEnc.FacilityNCID
  oDas2.DoSearch
  DoEvents
  
  Do While Not oDas2.EOF
    Set oDataString = oDas2.Events
    Set oClinicalText = oDataString.PatientData
    
    If Not oDataString.DataEntryModel.IsEmpty Then
      Set oTNote = New TextNote
      If Not oDas2.Events.ModifyTime = NULL_DATE Then
        oTNote.DataDate = oDataString.ModifyTime
        oTNote.ClinicianNCID = oDataString.Clinician.ValueConcept
'-Mod 43e default to the clinician that created it
        If oTNote.ClinicianNCID = "" Then
          oTNote.ClinicianNCID = oDataString.CreateClinician.ValueConcept
        End If
      Else
        oTNote.DataDate = oDataString.CreateTime
        oTNote.ClinicianNCID = oDataString.CreateClinician.ValueConcept
      End If
      oTNote.DataID = oDataString.DataID
      oTNote.ProfileNCID = oDataString.DataEntryModel.ValueConcept.UniqueId
      
      oTNote.sText = oClinicalText.ReportText.TextObject.Text.Text
      oTNote.Source = TN_FORMTOOL
      TNotes.Add oTNote

    End If
    oDas2.GoToNext
  Loop
  Set oTNote = Nothing
  Set oDas2 = Nothing
  Set oClinicalText = Nothing
  Set oDataString = Nothing
'-Mod 45e  Have to  have these set
  ReDim Preserve oldCounted(0)
  ReDim Preserve newCounted(0)
'~
  
  Select Case oEnc.Status
  Case Complete, Updated
  
  Case Updating   '-- this is actually an amendment of an rtf, so...
  '-- build a string to add as a note on the bottom
    For Each oTNote In TNotes
      FoundIt = False
      For Each oldTNote In oEnc.TextNotes
        
        If oldTNote.DataID = oTNote.DataID And oldTNote.Source = TN_FORMTOOL Then
          If oldTNote.sText <> oTNote.sText Then
          '--modified
            ReDim Preserve oldCounted(UBound(oldCounted) + 1)
            oldCounted(UBound(oldCounted)) = oldTNote.DataID
            ReDim Preserve newCounted(UBound(newCounted) + 1)
            newCounted(UBound(newCounted)) = oTNote.DataID
            
  '-- catch the new string for the so section
            OutString = OutString & "{\plain\f0\fs16\b\u Modified: \plain\f0\fs16 " & oldTNote.ProfileName & "\par}"
'-Mod 45e  spelling error
            OutString = OutString & "{\plain\f0\fs16\b Original Text\plain\f0\fs16 as entered on  " & oldTNote.DataDate & "\par}"
'-Mod 45e used the wrong object (otnote rather than oldtnote)
            Call mobjUtil.ConvertHTMLtoRTF(oldTNote.sText, rnNumTextStrings, rasTextStrings)
            For i = 0 To rnNumTextStrings - 1
              OutString = OutString & "{\plain\f0\fs16 " & rasTextStrings(i) & "\par\par}"
            Next i
            
            FoundIt = True
            Exit For
            
          End If
        End If
      Next oldTNote
        
    Next oTNote
    
    '-- now see if we deleted any.  Look for any olds that are not in the oldcounted list
    
    For Each oldTNote In oEnc.TextNotes
      For i = 1 To UBound(oldCounted)
        If oldTNote.DataID = oldCounted(i) And oldTNote.Source = TN_FORMTOOL Then
          Exit For
        End If
      Next i
      If i > UBound(oldCounted) Then
      '-- didn't find it, so it's been deleted
        OutString = OutString & "{\plain\f0\fs16\b\u Deleted: \plain\f0\fs16 " & oldTNote.ProfileName & "\par}"
        OutString = OutString & "{\plain\f0\fs16\b Original Text\plain as entered on  " & oldTNote.DataDate & "\par}"
'-MOd 45e
        Call mobjUtil.ConvertHTMLtoRTF(oldTNote.sText, rnNumTextStrings, rasTextStrings)

        For i = 0 To rnNumTextStrings - 1
          OutString = OutString & "{\plain\f0\fs16 " & rasTextStrings(i) & "\par\par}"
        Next i
        
      End If
    Next oldTNote
    If OutString <> "" Then
      OutString = msgHeader & OutString & msgFooter
      AddEncounterNote "AMEND", OutString, False
    End If
'-Mod 44e pushed the 'case else' stuff out to where every run will hit it
  End Select
  
    '--convert all the otnotes.stext's
   For i = oEnc.TextNotes.count To 1 Step -1
     If oEnc.TextNotes(i).Source = TN_FORMTOOL Then
       oEnc.TextNotes.Remove i
     End If
   Next i
   For Each oTNote In TNotes
     oEnc.TextNotes.Add oTNote
   Next oTNote
   
'    Set oEnc.TextNotes = TNotes
    OutString = ""
'-Mod 26e
   For Each oTNote In oEnc.TextNotes
     
       OutString = OutString & "{\par\plain\f0\fs16\b\tx" & frmEncParent.RTB(SECT_SO).Width \ 2 & oTNote.ProfileName & "\plain\f0\fs16 \tab  [" & oTNote.ClinicianName & ", " & mobjUtil.FormatDTG(oTNote.DataDate) & "] \par}"
     If oTNote.Source = TN_FORMTOOL Then
       Call mobjUtil.ConvertHTMLtoRTF(oTNote.sText, rnNumTextStrings, rasTextStrings)
       For i = 0 To rnNumTextStrings - 1
         OutString = OutString & "{\pard\plain\f0\fs16 " & rasTextStrings(i) & "\par\par}"
       Next i
     Else
       OutString = OutString & "{\pard\plain\f0\fs16 " & oTNote.sText & "\par\par}"
     End If
   Next oTNote
   
   oEnc.SectionRTFs(CStr(SECT_SO)).Doc = msgHeader & OutString & msgFooter
   frmEncParent.RTB(SECT_SO).Text = ""
   frmEncParent.RTB(SECT_SO).Height = 500
   frmEncParent.pics(SECT_SO).Height = 500
   frmEncParent.RTB(SECT_SO).RTFSelText = msgHeader & OutString & msgFooter
   DoEvents
   frmEncParent.pics(SECT_SO).Height = frmEncParent.RTB(SECT_SO).Height
   DoEvents
  ' frmEncParent.RepositionPicsBelow (SECT_SO - 1)
  
  
#If debugon Then
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "RefreshSO", "End", DR, "EncID", oEnc.EncounterID
#End If
  
  Set TNotes = Nothing
  Set oClinicalText = Nothing
  Set oConcept = Nothing
  Set oDas2 = Nothing
  Set oDataString = Nothing
  Set oTNote = Nothing
  Set oldTNote = Nothing

  
    Exit Function

errHandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "Enctr.refreshSO ", "Enc", vbCritical)
  Set TNotes = Nothing
  Set oClinicalText = Nothing
  Set oConcept = Nothing
  Set oDas2 = Nothing
  Set oDataString = Nothing
  Set oTNote = Nothing
  Set oldTNote = Nothing

    Exit Function
 
  
End Function

'****************************************************************************************************
'  Description: This method is called when the CW Core creates an "instance" of this OLE Server.  It
'               is used to create and initialize global variables that are shared across all
'               "views" supported by this OLE Server, i.e. variables that only need to be
'               initialized ONCE, upon object creation.
'
'  Parameters:
'    vlMyProgId - Contains the "ProgId" that the CW Core has assigned to this OLE Server.
'    robjComm - Contains a communications object that is used to handle all communication
'               between this OLE Server and the CW Core.
'    robjLogonEnvironment - Contains the logon environment object.  Contains information such
'                           as who logged on, where they logged on, etc.
'    robjConfig - Contains the configuration object that is used to get and set configuration
'                 information for a specific application.
'    robjPatient - Contains the patient object containing the currently selected patient.
'    robjUser - Security User information object
'****************************************************************************************************
Public Sub ICHCSII_CoreAPI2_Initialize2(ByVal vlMyProgId As Long, _
                                              robjComm As Object, _
                                              robjLogonEnvironment As Object, _
                                              robjConfig As Object, _
                                              robjPatient As Object, _
                                              robjUser As Object)
  
#If debugon Then
  Dim DR As DebugRec
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "CoreAPI2_Initialize2", "Begin", DR, "vlMyProgId", vlMyProgId
#End If
    
  If Not DebugMode Then On Error GoTo errHandler
  
  Set mobjUser = robjUser
 
  Call Initialize(vlMyProgId, robjComm, robjLogonEnvironment, robjConfig, robjPatient)
 
       
#If debugon Then
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "CoreAPI2_Initialize2", "End", DR, "vlMyProgId", vlMyProgId
#End If
    
    Exit Sub

errHandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "Enctr.Initialize2", "EncounterParent", vbExclamation)
    
    Exit Sub
End Sub

'-----------------------------------------------------
'Purpose:  add this module's menus to the core's Action menu and tool bar
'
'Parameters: none
'
'Comments:
'----------------------------------------------------
Private Sub AddMenus()

  If Not DebugMode Then On Error GoTo errHandler

#If debugon Then
  Dim DR As DebugRec
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "AddMenus", "Begin", DR
#End If
    
    Dim NewMenu As I_EncMenu      '- generic menu object
'--01.20.00 WAS - See if the menu is "up" before adding it
'  Build a new colection indexed on the name
    Dim cTemp As New Collection
    Dim oTemp As I_EncMenu
    For Each oTemp In mcolMenus
      cTemp.Add oTemp, oTemp.MenuName
    Next oTemp
    
    On Error Resume Next
    
    Set NewMenu = New EncMenuRefresh
    Set oTemp = cTemp(NewMenu.MenuName)
    If Err Then
      NewMenu.Enc = Me
      mcolMenus.Add NewMenu, CStr(NewMenu.MenuID)
      Err.Clear
    End If
    
    Set NewMenu = New EncMenuAddNote
    Set oTemp = cTemp(NewMenu.MenuName)
    If Err Then
      NewMenu.Enc = Me
      mcolMenus.Add NewMenu, CStr(NewMenu.MenuID)
      Err.Clear
    End If
    
    Set NewMenu = New EncMenuAddProvider
    Set oTemp = cTemp(NewMenu.MenuName)
    If Err Then
      NewMenu.Enc = Me
      mcolMenus.Add NewMenu, CStr(NewMenu.MenuID)
      Err.Clear
    End If
    
    Set NewMenu = New encMenuTemplateSelect
    Set oTemp = cTemp(NewMenu.MenuName)
    If Err Then
      NewMenu.Enc = Me
      mcolMenus.Add NewMenu, CStr(NewMenu.MenuID)
      Err.Clear
    End If
    
    
    Set NewMenu = New EncMenuSign
    Set oTemp = cTemp(NewMenu.MenuName)
    If Err Then
      NewMenu.Enc = Me
      mcolMenus.Add NewMenu, CStr(NewMenu.MenuID)
      Err.Clear
    End If
    

    Set NewMenu = Nothing
    Set oTemp = Nothing
    Set cTemp = Nothing
    
    
#If debugon Then
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "AddMenus", "End", DR
#End If

  Exit Sub
errHandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "Enctr.AddMenus", "EncounterParent", vbExclamation)
    
    Exit Sub

End Sub

'-----------------------------------------------------
'Purpose:  Removes this module's menus from the core's Action menu and tool bar
'
'Parameters: none
'
'Comments:
'----------------------------------------------------
Friend Sub RemoveMenus()
  If Not DebugMode Then On Error GoTo errHandler

#If debugon Then
  Dim DR As DebugRec
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "RemoveMenus", "Begin", DR
#End If

  Dim i       As Integer    '- generic counter
  Dim EncMenu As I_EncMenu  '- menu object
'-Mod 8e
  If mcolMenus Is Nothing Then Exit Sub

  For i = mcolMenus.count To 1 Step -1
     mcolMenus.Item(1).Delete
     mcolMenus.Remove 1
     DoEvents
  Next i
  
  Set EncMenu = Nothing
    
#If debugon Then
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "RemoveMenus", "End", DR
#End If

  Exit Sub
errHandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "Enctr.RemoveMenus", "EncounterParent", vbExclamation)
    
    Exit Sub

End Sub

'-----------------------------------------------------
'Purpose:  initialize encounter child objects
'
'Parameters: none
'
'Comments:  02.10.00 With the addition of Medcin as an option in the S/O secion,
'           we have to provide special dispensation for SECT_MEDCIN
'----------------------------------------------------
Friend Sub LoadChildren()
  
  Dim i       As Integer          '- generic counter
  Dim oAC     As EncAutoCite      '- autoCite object
  Dim RTF     As String           '- rtf string
  Dim RTFText As String           '- generic rtf string
  Dim Svr     As Object           '- child module object
  Dim sRTF    As EncRTF           '- rtf object
  
  If Not DebugMode Then On Error GoTo errHandler
  
#If debugon Then
  Dim DR As DebugRec
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "LoadChildren", "Begin", DR
#End If
  
  Load frmEncChildLoad
  frmEncChildLoad.Top = (Screen.Height - frmEncChildLoad.Height) / 3
  frmEncChildLoad.Left = (Screen.Width - frmEncChildLoad.Width) / 2
  frmEncChildLoad.Show
  frmEncChildLoad.lblEncounter = "Loading Encounter: " & mobjCurrEnc.EncounterID & "..."
  
  Call mobjComm.setstatustext("Encounter: " & mobjCurrEnc.EncounterID & ", Patient: " & mobjPatient.Name)
  
  frmEncChildLoad.imgArrow(0).Visible = True
  frmEncChildLoad.Refresh
  DoEvents
    
'jf todo: figure out if RefreshAutoCite() should be called from LoadChildren or from somewhere else
  Call RefreshAutoCite(mobjProperties.AutoCite)
'-Mod 56e
  sOriginals(SECT_AC).Doc = frmEncParent.RTBAC(SECT_AC).RTFSelText
  sOriginals(SECT_LABS).Doc = frmEncParent.RTB(SECT_LABS).RTFSelText
  sOriginals(SECT_RADS).Doc = frmEncParent.RTB(SECT_RADS).RTFSelText
'~
  '-- be sure the mcolEncChild is empty
  Set mcolEncChild = Nothing
  Set mcolEncChild = New Collection
              
 frmEncChildLoad.imgArrow(0).Visible = False

Dim iControl As Integer

      For i = 1 To UBound(Sect)
          
          If i = SECT_MEDCIN Then
            iControl = SECT_SO
          Else
            iControl = i
          End If
          
          frmEncChildLoad.imgArrow(i).Visible = True
          frmEncChildLoad.Refresh
          DoEvents
          
          Set mEncChild = New EncChild

          mEncChild.ServerIndex = i
          mEncChild.ClientControl = frmEncParent.RTB(iControl)
'-Mod 35e moved from encChild.startmodule
          mEncChild.AppNCID = CStr(glAppNCID(i))
          mEncChild.ServerName = Sect(i)
'~
'-- launch the child
          Call mEncChild.StartModule(Me, i, False, sOriginals(i))
          
          frmEncChildLoad.Refresh
          DoEvents

          mcolEncChild.Add mEncChild, CStr(mEncChild.ServerIndex)
          
'-Mod 2e - Disable them all, turn them on only if they have write privilages
          frmEncParent.cmds(iControl).Enabled = False
          
          If gsSecAreaNames(i) = "" And mobjUser.HasPrivilegeEx(Priv_Current_Encounter, Priv_Write) Then
             frmEncParent.cmds(iControl).Enabled = True
          Else
            If Not gsSecAreaNames(i) = "" Then
              
              If Not mobjUser.HasPrivilegeEx(gsSecAreaNames(i), Priv_Read) Then
                If Not iControl = SECT_MEDCIN Then
                  frmEncParent.RTB(iControl).Text = ""
                End If
                frmEncParent.RTB(iControl).RTFSelText = "{\rtf1\ansi\ansicpg1252\uc1 \deff0\deflang1033\deflangfe1033" _
                      & "{\fonttbl{\f0\fswiss\fcharset0\fprq2 Arial;}}" _
                      & "{\colortbl;\red0\green0\blue0;}" _
                      & "\pard\plain \widctlpar\adjustright\cgrid\f0\fs16 Insufficient privilages to view this data}"
                frmEncParent.cmds(iControl).Enabled = False
              End If
            
'-- check privilages, turn off buttons not writable
              If Not (mobjUser.HasPrivilegeEx(gsSecAreaNames(i), Priv_Write) And mobjUser.HasPrivilegeEx(Priv_Current_Encounter, Priv_Write)) Then
                frmEncParent.cmds(iControl).Enabled = False
              Else
                frmEncParent.cmds(iControl).Enabled = True
              End If
            End If
          End If
'~end Mod 2e

          Set mEncChild = Nothing
          frmEncChildLoad.imgArrow(i).Visible = False
          
NextModule:
'-Mod 2e
        If Err Then
          frmEncParent.cmds(iControl).Enabled = False
          Err.Clear
        End If
      Next i
'-Mod 56e
      sOriginals(SECT_MEDCIN).Doc = mobjCurrEnc.SectionRTFs(CStr(SECT_MEDCIN)).Doc
      sOriginals(SECT_MEDCIN).CreatedBy = mobjCurrEnc.SectionRTFs(CStr(SECT_MEDCIN)).OwnerName
      sOriginals(SECT_MEDCIN).CreatedOn = mobjCurrEnc.SectionRTFs(CStr(SECT_MEDCIN)).DTS
      '~
      Unload frmEncChildLoad
      
      #If debugon Then
        If Not DL Is Nothing Then DL.DebugLog Exe, Module, "LoadChildren", "End", DR
      #End If
      
    Call mobjComm.setstatustext("Encounter: " & mobjCurrEnc.EncounterID & ", Patient: " & mobjPatient.Name)
          
    Exit Sub

errHandler:
    '---show the VB error---
    gobjShared.ShowVBError Err.Number, Err.Description, "Enc.LoadChildren", App.Title, vbCritical
    Unload frmEncChildLoad
    Resume NextModule
    
End Sub

'-----------------------------------------------------
'Purpose:  Refresh the RTF strings in the autocite section
'
'Parameters: colAutoCites - a collection of autocite objects
'
'Comments:
'----------------------------------------------------

Public Sub RefreshAutoCite(colAutoCites As Collection)
Dim RTF     As String           '- Generic RTF string
Dim RTFRads As String           '- Labs and Rads always go to their own section, so we
Dim RTFLabs As String           '   have separate string vars for them
Dim RTFAC   As String
Dim oAC     As EncAutoCite      '- an individual Autocite object
Dim Svr     As Object           '- encounter child module
Dim I_AC    As ICHCSII_EncSum3  '- autocite-support interface instance
Dim RTFText As String           '- generic rtf string
Dim sRTF    As EncRTF           '- encounter sectional rtf object
Dim bReloadAutocite As Boolean  '- do we need to reload?
Dim bSelRads As Boolean
Dim bSelLabs As Boolean

If Not DebugMode Then On Error GoTo errHandler
  
  
#If debugon Then
  Dim DR As DebugRec
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "RefreshAutoCite", "Begin", DR
#End If
  
  '**12.28.99 WAS Check the autocite stuff to see if we have anything, if so, don't do it again
  frmEncParent.RTBAC(SECT_AC).AutoExpand = True
  frmEncParent.RTBAC(SECT_AC).Text = ""
  
'-Mod 2e Added permission check
  If Not mobjUser.HasPrivilegeEx(Priv_Current_Encounter, Priv_Write) Then
    frmEncParent.cmds(SECT_AC).Enabled = False
  End If
'~end mod 2
  
  On Error Resume Next
  frmEncParent.RTBAC(SECT_AC).RTFSelText = mobjCurrEnc.SectionRTFs(CStr(SECT_AC)).Doc
  If Err Then
    Err.Clear
  End If
  
  If Trim$(frmEncParent.RTBAC(SECT_AC).Text) = "" Then
    bReloadAutocite = True
  End If
  If Not DebugMode Then On Error GoTo errHandler
  
  RTF = ""
  RTFLabs = ""
  RTFRads = ""
'-Mod 30e
  RTFAC = ""
  
'-- transfer the set to the current ac
  For Each oAC In colAutoCites
      If oAC.Selected Then

        Set Svr = mobjComm.InitializeOLEServer(oAC.AppNCID)
        If Not Svr Is Nothing Then
          If TypeOf Svr Is ICHCSII_EncSum3 Then
            Set I_AC = Svr
            RTFText = ""
            RTFText = I_AC.GetAutoCiteNarrative("", oAC.NumberPeriods, oAC.Period)
            If RTFText <> "" Then
'-Mod 6e
              Select Case oAC.AppNCID
              Case NCID_LAB_APPLICATION
                bSelLabs = True
                Call RefreshAutoCiteSendToForm(SECT_LABS, RTFText)

              Case NCID_RADIOLOGY_RESULTS_APPLICATION
                bSelRads = True
                Call RefreshAutoCiteSendToForm(SECT_RADS, RTFText)
                
              Case Else
                If bReloadAutocite Then
                  RTFAC = RTFAC & "{\pard\plain\f0\fs16\b\ul " & oAC.AppName & "\par}" & RTFText & "\par "
                End If
              End Select
'~end mod 6e
            End If
          End If  '- TypeOf Svr Is ICHCSII_EncSum3 Then
        End If    '- Not Svr Is Nothing Then
      Debug.Print "Autocite", oAC.AppName, RTFText
      #If debugon Then
        If Not DL Is Nothing Then DL.DebugLog Exe, Module, "RefreshAutoCite", "Init Svr", DR, "Server", oAC.AppName, "text", RTFText
      #End If
      
      End If      '- oAC.Selected Then
    Next oAC

'-Mod 30e
    If bReloadAutocite Then
        RTFAC = "{\rtf1\ansi\ansicpg1252\uc1 \deff0\deflang1033\deflangfe1033" _
                 & "{\fonttbl{\f0\fswiss\fcharset0\fprq2 Arial;}}" _
                 & "{\colortbl;\red0\green0\blue0;}" _
                 & "\pard\plain \widctlpar\adjustright \cgrid" _
                & RTFAC & "}"
    End If
    
'-Mod 6e Consolidated a block of code into a single function
    Call RefreshAutoCiteSendToForm(SECT_AC, RTFAC)
    
'jf todo: fix flat out setting of .FlagChanged in RefreshAutoCite()
    mobjCurrEnc.FlagChanged = True

#If debugon Then
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "RefreshAutoCite", "end", DR
#End If
    Set oAC = Nothing
    Set I_AC = Nothing
    Set sRTF = Nothing
'-Mod 38e
    Set Svr = Nothing
    Exit Sub

errHandler:
    '---show the VB error---
    gobjShared.ShowVBError Err.Number, Err.Description, "Enctr.RefreshAutoCite", App.Title, vbCritical
    Set oAC = Nothing
    Set I_AC = Nothing
    Set sRTF = Nothing
'-Mod 38e
    Set Svr = Nothing
    Exit Sub
    
End Sub
'-Mod 6e Added function as a consolidation of code from RefreshAutoCite(), above
Private Function RefreshAutoCiteSendToForm(idx As Integer, sText As String) As Boolean
  Dim sRTF As EncRTF
  Dim RTFText As String
  Dim TxControl As Control
  Dim sDoc As String
  
  If Not DebugMode Then On Error GoTo errHandler
  
  Set sRTF = mobjCurrEnc.SectionRTFs(CStr(idx))
  If sRTF Is Nothing Then
    Set sRTF = New EncRTF
    sRTF.SectionIndex = idx
    mobjCurrEnc.SectionRTFs.Add sRTF, CStr(idx)
  End If
  
  Select Case True
  Case Not sText = ""      '- use what they gave you
    sDoc = sText
    sRTF.Doc = sText
    sRTF.DTS = Now
  Case Not sRTF.Doc = ""   '- use what you have
    sDoc = sRTF.Doc
  Case Else
    Set sRTF = Nothing
    Exit Function
  End Select
  
  RTFText = "{\rtf1\ansi\ansicpg1252\uc1 \deff0\deflang1033\deflangfe1033" _
          & "{\fonttbl{\f0\fswiss\fcharset0\fprq2 Arial;}}" _
          & "{\colortbl;\red0\green0\blue0;}" _
          & "\pard\plain \widctlpar\adjustright \cgrid"
  If Not idx = SECT_AC Then
    RTFText = RTFText & "{\pard\plain\f0\fs16 \b \ul " & Sect(idx) & "\par}"
    RTFText = RTFText & "{\pard\plain \li500 \f0\fs16 " & sDoc & "}}"
  Else
    RTFText = RTFText & "{\pard\plain \f0\fs16 " & sDoc & "}}"
  End If
   
  If Not RTFText = "" Then
    If idx = 0 Then
      Set TxControl = frmEncParent.RTBAC(idx)
    Else
      Set TxControl = frmEncParent.RTB(idx)
    End If
    DoEvents
    TxControl.AutoExpand = True
    TxControl.Text = ""
    TxControl.RTFSelText = RTFText
    Set TxControl = Nothing
  End If
  
  Set sRTF = Nothing
  
  Exit Function

errHandler:
    '---show the VB error---
    gobjShared.ShowVBError Err.Number, Err.Description, "Enctr.RefreshAutoCiteSendToForm", App.Title, vbCritical
    Exit Function
   
End Function

Friend Function ReloadAutoCites() As Boolean
'-- go and get a new set of autocite parameters

End Function

'-----------------------------------------------------
'Purpose:  OBE  (see RetrievePreviousEncountersEx)
'----------------------------------------------------
Friend Function RetrievePreviousEncounters() As Boolean

  RetrievePreviousEncounters = RetrievePreviousEncountersEx

End Function
'-----------------------------------------------------
'Purpose: Main Encounter retreival method
'
'Parameters: none
'
'Returns: True on success
'
'Comments:  Repeadedly call RetreiveEnc for each item returned in Das.search
'----------------------------------------------------
Public Function RetrievePreviousEncountersEx() As Boolean
    Dim oEnc          As Encounter    '- an encounter object
    Dim oEnc2         As Encounter    '- a second one
    Dim oDas          As DasCtrl      '- Das control object
    Dim oConcept      As ConceptCtrl  '- converter of ncids to prepresentation 2000
    Dim j             As Integer      '- generic counter
    Dim FoundIt       As Boolean      '- loop flag
    
#If debugon Then
  Dim DR As DebugRec
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "RetrievePreviousEncounters", "Begin", DR
#End If
    
'--01.18.00 WAS: Do we really need to do this?  If the current col of enc's exists and is for this
'                patient, the answer is 'no.'
    If mcolEncounters.count > 0 Then
      If mcolEncounters.Item(1).Patient.UnitNumber = mobjPatient.UnitNumber Then
        RetrievePreviousEncountersEx = True
        Exit Function
      Else
'--01.27.00 added
        Set mcolEncounters = New Collection
      End If
    End If
'~end ~
    Set oConcept = New ConceptCtrl
    Set oDas = New DasCtrl
    oDas.ClearAll
    oDas.GeneralCriteria.ClearAll
    
    oDas.GeneralCriteria.UnitNumber = mobjPatient.UnitNumber
    oDas.GeneralCriteria.QueryQualifier = "EVENT_TYPE_NCID = '" & NCID_ENCEVENT & "'"
    oDas.GeneralCriteria.ReturnMultiMedia = True
    
    oDas.DoSearch
    
  On Error GoTo ehNextEvent
  j = 0
  Do While Not oDas.EOF
'-- there are some "E###" codes used by others, skip them
    Do While (oDas.Events.EncounterNumber = "0" Or Not IsNumeric(oDas.Events.EncounterNumber)) _
    And Not oDas.EOF
      oDas.GoToNext
    Loop
    If oDas.EOF Then Exit Do
    j = j + 1
    Call mobjComm.setstatustext("Retrieving Encounters..." & String(j, ".."))
    Set oEnc = New Encounter
    oEnc.Init mobjUser
    Set oEnc.Patient = mobjPatient
    
    If oEnc.RetrieveEnc(oDas) Then
    
      On Error Resume Next
      FoundIt = False
      For Each oEnc2 In mcolEncounters
        If oEnc2.EncounterID = oEnc.EncounterID Then
           FoundIt = True
           Exit For
        End If
      Next oEnc2
      If Not FoundIt Then
        mcolEncounters.Add oEnc, CStr(oEnc.StorageKey)
      End If
      Set oEnc = Nothing
      Set oEnc2 = Nothing
      On Error GoTo ehNextEvent
    End If
    oDas.GoToNext
ehNextEvent:
   Loop   ' while not odas.eof
   On Error GoTo errHandler
   Set oDas = Nothing
   Set oConcept = Nothing
   Call mobjComm.setstatustext("")
    
#If debugon Then
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "RetrievePreviousEncounters", "End", DR
#End If

    RetrievePreviousEncountersEx = True
    Exit Function

errHandler:
    '---show the VB error---
    gobjShared.ShowVBError Err.Number, Err.Description, "Enc.RetrievePreviousEncounters", App.Title, vbCritical
    
    Set oEnc = Nothing
    Set oEnc2 = Nothing
    Set oConcept = Nothing
    Set oDas = Nothing
End Function

'-----------------------------------------------------
'Purpose: Summary Encounter retrieval method
'
'Parameters: none
'
'Returns: True on success
'
'Comments:  Repeadedly call RetreiveEnc for each item returned in Das.search
'----------------------------------------------------
Public Function GetEncounterInfo() As Boolean
'jf todo: move this stuff to the encounter object - check first to see if the code is essentially already there.
    Dim oEnc          As Encounter    '- an encounter object
    Dim oEnc2         As Encounter    '- a second one
    Dim oDas          As DasCtrl      '- Das control object
    Dim oConcept      As ConceptCtrl  '- converter of ncids to prepresentation 2000
    Dim j             As Integer      '- generic counter
    Dim FoundIt       As Boolean      '- loop flag
    
#If debugon Then
  Dim DR As DebugRec
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "GetEncounterInfo", "Begin", DR
#End If
    
   Set mcolEncounters = New Collection
    Set oConcept = New ConceptCtrl
    Set oDas = New DasCtrl
    oDas.ClearAll
    oDas.GeneralCriteria.ClearAll
    oDas.GeneralCriteria.UnitNumber = mobjPatient.UnitNumber
    oDas.GeneralCriteria.QueryQualifier = "EVENT_TYPE_NCID = '" & NCID_ENCEVENT & "'"
    oDas.GeneralCriteria.ReturnMultiMedia = True
    
    oDas.DoSearch
    
  On Error GoTo ehNextEvent
  j = 0
  Do While Not oDas.EOF
'-- there are some "E###" codes used by others, skip them
    Do While Not IsNumeric(oDas.Events.EncounterNumber) And Not oDas.EOF
      oDas.GoToNext
    Loop
    If oDas.EOF Then Exit Do
    
    j = j + 1
    Call mobjComm.setstatustext("Retrieving Encounters..." & String(j, ".."))
    Set oEnc = New Encounter
    oEnc.Init mobjUser
    Set oEnc.Patient = mobjPatient
'-Mod 47e
    oEnc.Loading = True
                             
  Dim Posn          As Integer                '- holds position from an instr call
  Dim txt           As String                 '- generic text holder
  Dim i             As Integer                '- generic counter
  Dim lTempStatus   As Long
  Dim oClinEvent    As ClinicalEvent          '- Das control subpart
  Dim oClinHdr      As ClinObsHeader          '- Das control subpart
  Dim oClinObs      As ClinicalObservation    '- Das control subpart
  Dim oCommentGroup As CommentGroup           '- Das control subpart
  Dim oCommentItem  As CommentItem            '- Das control subpart
  Dim oDiag         As EncDiagnosis           '- encounter diagnosis object (filled from sematic links)
  Dim oProv         As EncProvider            '- provider object
  Dim oProv2        As EncProvider
  
    oEnc.DataID = oDas.Events.DataID
    oEnc.EncounterID = oDas.Events.EncounterNumber
    oEnc.PatientID = oDas.Events.UnitNumber
'    Debug.Print "loading encounter " & oEnc.EncounterID
'-- DDSA wants this in the "lite" retrieve
    If (Not oDas.Events.EncounterFacility.IsEmpty) Then
      oEnc.FacilityNCID = oDas.Events.EncounterFacility.ValueConcept
      oEnc.Facility = oDas.Events.EncounterFacility.ValueConcept.PrefRep("2000").Representation
    End If
    
'--020800 WAS more checks for older datasets
    If Not oDas.Events.CreatePointOfCare Is Nothing Then
      If (Not oDas.Events.CreatePointOfCare.IsEmpty) Then
        If oDas.Events.CreatePointOfCare.PhysicalLocation.InfoType = "nursingDivision" Then
          oEnc.ClinicNCID = oDas.Events.CreatePointOfCare.PhysicalLocation.NursingDivision.ValueConcept
          oEnc.Clinic = oDas.Events.CreatePointOfCare.PhysicalLocation.NursingDivision.ValueConcept.PrefRep("2000").Representation
        End If
      End If
    End If
    Set oDiag = Nothing
        
        Set oClinEvent = oDas.Events.PatientData
        Set oClinHdr = oClinEvent.Header
      
        With oClinHdr
          If (Not .ObsDateTime.IsEmpty) Then
            oEnc.StartDTS = .ObsDateTime.StartTime.Value.DateTime
          End If
          
          lTempStatus = 0
          If Not .TestStatus.IsEmpty Then
            lTempStatus = .TestStatus.Value.CodedWOSform.ValueConcept
            For i = LBound(gsEncStatusNCID) To UBound(gsEncStatusNCID)
              If Val(gsEncStatusNCID(i)) = lTempStatus Then
                lTempStatus = i
                Exit For
              End If
            Next i
          End If
          
          If (Not .Comments.IsEmpty) Then
            For Each oCommentGroup In .Comments
              For Each oCommentItem In oCommentGroup.Comments
                txt = oCommentItem.Text
                Posn = InStr(txt, "=")
                
                If Posn > 1 And Len(txt) > Posn Then
                  Select Case UCase$(Left$(txt, Posn - 1))
                  Case "ENCSTATUS"
                    If lTempStatus = 0 Then lTempStatus = Val(Mid$(txt, Posn + 1))
                  Case "ENCCAT"
                    oEnc.Category2 = Mid$(txt, Posn + 1)
                  Case "ENCTYPE"
                    oEnc.Enc_Type = Mid$(txt, Posn + 1)
                  Case "PRIMPROV"
                    Set oProv = New EncProvider
                    oProv.NCID = Mid$(txt, Posn + 1)
                     oConcept.UniqueId = Mid$(txt, Posn + 1)
                    oProv.FullName = oConcept.PrefRep("2000").Representation
                    
                    For Each oProv2 In oEnc.Providers
                      If oProv2.NCID = oProv.NCID Then
                       Exit For
                      End If
                    Next oProv2
                    If oProv2 Is Nothing Then
                      oEnc.Providers.Add oProv, CStr(oProv.NCID)
                    End If
                    
                    Set oProv = Nothing
                    Set oProv2 = Nothing
                    
                  
                  End Select
                End If
              
              Next oCommentItem
            Next oCommentGroup
          End If    '-- If (Not .Comments.IsEmpty) Then
      End With      '-- oClinHdr

Dim oClinObs2 As ClinicalObservation

      For Each oClinObs In oClinEvent.ClinObs
        If Not oClinObs.ObsId.IsEmpty Then
          If oClinObs.ObsId.Value.InfoType = "codedWOSform" Then
            Select Case oClinObs.ObsId.Value.CodedWOSform.ValueConcept
            Case NCID_DIAGNOSIS
                oEnc.PrimaryDiagnosisNCID = oClinObs.ObsValue.Value.CodedWOSform.ValueConcept
                Exit For
            Case NCID_ADDITIONALPROVIDER
             Set oProv = New EncProvider
             oProv.NCID = oClinObs.ObsValue.Value.CodedWOSform.ValueConcept
             oProv.FullName = oClinObs.ObsValue.Value.CodedWOSform.ValueConcept.PrefRep("2000").Representation
             
             '-- look for role
             If Not oClinObs.ObsMods.IsEmpty Then
               For Each oClinObs2 In oClinObs.ObsMods
                 Select Case oClinObs2.ObsId.Value.CodedWOSform.ValueConcept
                 Case NCID_PROVIDER_ROLE
                   oProv.RoleNCID = oClinObs2.ObsValue.Value.CodedWOSform.ValueConcept
                   oProv.Role = oClinObs2.ObsValue.Value.CodedWOSform.ValueConcept.PrefRep("2000").Representation
                 End Select
               Next oClinObs2
               Set oClinObs2 = Nothing
             End If
             
             For Each oProv2 In oEnc.Providers
               If oProv2.NCID = oProv.NCID Then
                Exit For
               End If
             Next oProv2
             If oProv2 Is Nothing Then
               oEnc.Providers.Add oProv, CStr(oProv.NCID)
             End If
              
            End Select
          End If
        End If
     Next oClinObs
     Set oClinObs = Nothing
 '-- now update the status field and cascade the update to the appointment object
     oEnc.Status = lTempStatus
'-Mod 47e
    oEnc.Loading = False
    
'===================================================================================
      On Error Resume Next
      FoundIt = False
      For Each oEnc2 In mcolEncounters
        If oEnc2.EncounterID = oEnc.EncounterID Then
           FoundIt = True
           Exit For
        End If
      Next oEnc2
      If Not FoundIt Then
        mcolEncounters.Add oEnc, CStr(oEnc.StorageKey)
      End If
'Debug.Print "Finished loading " & oEnc.EncounterID
    oDas.GoToNext
ehNextEvent:
   Loop   ' while not odas.eof
   
   On Error GoTo errHandler
   Set oDas = Nothing
   Set oConcept = Nothing
   Set oEnc = Nothing
   Set oEnc2 = Nothing

   Call mobjComm.setstatustext("")
    
#If debugon Then
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "GetEncounterInfo", "End", DR
#End If

    GetEncounterInfo = True
    Exit Function

errHandler:
    '---show the VB error---
    gobjShared.ShowVBError Err.Number, Err.Description, "Enc.GetEncounterInfo", App.Title, vbCritical
  Set oConcept = Nothing
  Set oEnc = Nothing
  Set oEnc2 = Nothing
 
  Set oDas = Nothing
End Function

Public Function SetOneEncounter(DataID As String) As Boolean
  '-- read in one comlete enocunter record
  '- set the mobjcurrenc to the one selected
  
  Dim oEnc          As Encounter    '- an encounter object
  Dim oDas          As DasCtrl      '- Das control object
  Dim oConcept      As ConceptCtrl  '- converter of ncids to prepresentation 2000
  
#If debugon Then
  Dim DR As DebugRec
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "SetOneEncounter", "Begin", DR, "Data ID", DataID
#End If
  
  Call mobjComm.setstatustext("Retrieving Encounter...")
  
'  Set mcolEncounters = New Collection
  Set oConcept = New ConceptCtrl
  Set oDas = New DasCtrl
  oDas.ClearAll
  oDas.GeneralCriteria.ClearAll
  oDas.GeneralCriteria.EventID = DataID
  oDas.GeneralCriteria.ReturnMultiMedia = True
  oDas.DoSearch
  
  If oDas.EOF Then
    Call mobjComm.setstatustext("Error retrieving encounter.")
    Exit Function
  End If
  
  Set oEnc = New Encounter
  oEnc.Init mobjUser
  Set oEnc.Patient = mobjPatient
  oEnc.DataID = DataID

  
  If oEnc.RetrieveEnc(oDas) Then
  
    oEnc.Current = True
    
    Set mobjCurrEnc = oEnc
    Call mobjCurrEnc.SetCurrentRTF
    mobjCurrEnc.FlagNew = False
    
    '-- for DDSA
    RaiseEvent ChangedEncounter(mobjCurrEnc.EncounterID, mobjCurrEnc.FacilityNCID)
    Call TellChildrenEncounterChanged
    
    '-- ask kids for new data
    If bMainViewRunning Then
      SetOneEncounter = ReloadData()
    Else
      SetOneEncounter = True
    End If
  
    #If debugon Then
      If Not DL Is Nothing Then DL.DebugLog Exe, Module, "SetOneEncounter", "End", DR
    #End If
  End If
  
  Set oEnc = Nothing
  Set oConcept = Nothing
 Set oDas = Nothing
End Function

'-----------------------------------------------------
'Purpose:  Sets the current (active) encounter, based on current patient encoutner nubmer,
'            into mobjCurrEnc
'
'Parameters: None
'
'Returns:  True on success
'
'Comments:  01.26.00 WAS:Just for J Flowers, If the col is empty,use info in the patient.
'-Mod 63e bail-out if the status is not right (must not be: NeedsCoSignature,Updated,complete,LWOBS)
'----------------------------------------------------
Public Function SetCurrentEncounter() As Boolean
Attribute SetCurrentEncounter.VB_Description = "Based on Patient.EncounterNumber, creates a CurrentEncounter object and sets it to the patient-specified encounter."
  Dim oEnc As Encounter
  Dim oChild As EncChild

#If debugon Then
  Dim DR As DebugRec
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "SetCurrentEncounter", "Begin", DR
#End If
  
  If Not DebugMode Then On Error GoTo errHandler

'-- based on the encounter number in the current patient, make that the current one
'-- dummy check
  If Val(mobjPatient.EncounterNumber) = 0 Then
    SetCurrentEncounter = False
    Exit Function
  End If
 '-Mod 71e  don't show this while changing it
    If bMainViewRunning Then
      frmEncParent.Visible = False
    End If
  
'-- see if we're even on the right patient.
  Call PatientSelected  '-- returns right away if everything's OK
  DoEvents
  
  If Not mobjCurrEnc Is Nothing Then
  
    With mobjCurrEnc
      If .FacilityNCID = mobjPatient.EncounterFacilityNCID _
       And .EncounterID = Val(mobjPatient.EncounterNumber) Then
          SetCurrentEncounter = True
          Call mobjCurrEnc.SetCurrentRTF
 '-Mod 71e
          If bMainViewRunning Then
            frmEncParent.Visible = True
          End If
                
          Exit Function
      End If
    End With
 '--  we're going to change enc, so save the current one
    Call mobjCurrEnc.Save
    
  Else
    Set mobjCurrEnc = New Encounter
  End If
  
'-Mod 17e
  If mcolEncounters Is Nothing Then
    Set mcolEncounters = New Collection
  End If
  
'--01.26.00 WAS: use patient data if enc col is empty

'==============================
'-Mod 71e
SetCurrentEncounter = ReloadData
'mod 71e ------delete all below.  It was optimization and it's in the reload data call.
'
''--  If mcolEncounters.count = 0 Then
'
'    Set mobjCurrEnc = GetEncounter(mobjPatient.UnitNumber, mobjPatient.EncounterFacilityNCID, mobjPatient.EncounterNumber)
'    Call mobjComm.setstatustext("") '- was leaveing a dribble
'
'    If mobjCurrEnc Is Nothing Then
'      SetCurrentEncounter = False
'    Else
''-Mod 63e
''      If Not AllowedInCurrentEncounter(mobjCurrEnc.Status) Then
''        SetCurrentEncounter = False
''        Exit Function
''      End If
'
'      RaiseEvent ChangedEncounter(mobjCurrEnc.EncounterID, mobjCurrEnc.FacilityNCID)
'      Call TellChildrenEncounterChanged
'      If bMainViewRunning Then
'        Call UnloadChildren
'        Call DisplayNotes
'        Call LoadChildren
'        Call ResizeView(cwiREGULAR_VIEW, frmEncParent.Left, frmEncParent.Top, frmEncParent.Width, frmEncParent.Height)
'        SetCurrentEncounter = True
'      Else
'       SetCurrentEncounter = True
'      End If
'      Call mobjCurrEnc.SetCurrentRTF
'    End If
'    Call mobjComm.setstatustext("") '- was leaveing a dribble
'    Exit Function
' '-- End If
'
''-- look through all the encounters for this enc num
'  For Each oEnc In mcolEncounters
'    Select Case mobjPatient.EncounterFacilityNCID
'    Case ""
'      If oEnc.EncounterID = Val(mobjPatient.EncounterNumber) Then
'         mobjPatient.EncounterFacilityNCID = oEnc.FacilityNCID
''-Mod 63e
''        If Not AllowedInCurrentEncounter(mobjCurrEnc.Status) Then
''          SetCurrentEncounter = False
''          Exit Function
''        End If
'
'        oEnc.Current = True
'
'        Set mobjCurrEnc = oEnc
'        Call mobjCurrEnc.SetCurrentRTF
'        mobjCurrEnc.FlagNew = False
'
''-- for DDSA
'        RaiseEvent ChangedEncounter(mobjCurrEnc.EncounterID, mobjCurrEnc.FacilityNCID)
'        Call TellChildrenEncounterChanged
''-- ask kids for new data
'        If bMainViewRunning Then
'          SetCurrentEncounter = ReloadData()
'        Else
'          SetCurrentEncounter = True
'        End If
'
'          #If debugon Then
'            If Not DL Is Nothing Then DL.DebugLog Exe, Module, "SetCurrentEncounter", "End", DR
'          #End If
'
'        Set oEnc = Nothing
'        Call mobjComm.setstatustext("") '- was leaveing a dribble
'        Exit Function
'      End If
'
'    Case Else
'      If oEnc.EncounterID = Val(mobjPatient.EncounterNumber) _
'        And oEnc.FacilityNCID = mobjPatient.EncounterFacilityNCID Then
''-Mod 63e
''        If Not AllowedInCurrentEncounter(mobjCurrEnc.Status) Then
''          SetCurrentEncounter = False
''          Exit Function
''        End If
'
'        oEnc.Current = True
'        Set mobjCurrEnc = oEnc
'
'        Call mobjCurrEnc.SetCurrentRTF
''-- for DDSA
'        RaiseEvent ChangedEncounter(mobjCurrEnc.EncounterID, mobjCurrEnc.FacilityNCID)
'        Call TellChildrenEncounterChanged
'
'        mobjCurrEnc.FlagNew = False
'
''-- ask kids for new data
'        If bMainViewRunning Then
'          SetCurrentEncounter = ReloadData
'        Else
'          SetCurrentEncounter = True
'        End If
'
''-- for DDSA
'        RaiseEvent ChangedEncounter(mobjCurrEnc.EncounterID, mobjCurrEnc.FacilityNCID)
'          #If debugon Then
'            If Not DL Is Nothing Then DL.DebugLog Exe, Module, "SetCurrentEncounter", "End", DR
'          #End If
'        Set oEnc = Nothing
'        Call mobjComm.setstatustext("") '- was leaveing a dribble
'        Exit Function
'      End If
'    End Select
'
'  Next oEnc
'
'  Set oEnc = Nothing
'  SetCurrentEncounter = False
  Call mobjComm.setstatustext("") '- was leaveing a dribble
#If debugon Then
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "SetCurrentEncounter", "End", DR
#End If
  
      Exit Function

errHandler:
  '---show the VB error---
  gobjShared.ShowVBError Err.Number, Err.Description, "Enc.SetCurrentEncounter", App.Title, vbCritical
  Set oEnc = Nothing
  SetCurrentEncounter = False
  
End Function


'-----------------------------------------------------
'Purpose:  remove my references to all child objects
'
'Parameters: none
'
'Comments:  since each child may have a reference to me, the parent, they have
'   to be killed before Core tries to kill me, or he will fail
'----------------------------------------------------
Friend Sub UnloadChildren()
  
  If Not DebugMode Then On Error GoTo errHandler

#If debugon Then
  Dim DR As DebugRec
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "UnloadChildren", "Begin", DR
#End If

Dim i      As Integer   '- generic counter
Dim EChild As EncChild  '- encounter child object

  '-- close all encchild
  For i = mcolEncChild.count To 1 Step -1
      Set EChild = mcolEncChild.Item(i)
'-- if the module loaded, unload it
'-- don't kill these kids, they might be up elsewhere
      If Not (EChild.AppNCID = NCID_MEDICATIONS_APPLICATION Or _
              EChild.AppNCID = NCID_PROBLEMS_APPLICATION Or _
              EChild.AppNCID = NCID_LAB_APPLICATION Or _
              EChild.AppNCID = NCID_RADIOLOGY_RESULTS_APPLICATION) Then
          If EChild.ModuleLoaded Then
            EChild.CloseChild
            DoEvents
          End If
      End If
      mcolEncChild.Remove i
      DoEvents
  Next i
  Set EChild = Nothing
  DoEvents

#If debugon Then
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "UnloadChildren", "End", DR
#End If

  Exit Sub
errHandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "Enctr.UnloadChildren", "EncounterParent", vbExclamation)
    
    Exit Sub

End Sub

'-----------------------------------------------------
'Purpose: Add a new narrative section to the current (open) encounter document.
'
'Parameters:
'  ModuleName - The source (calling) module name. Used as a title for the narrative.
'  NoteText - The actual text of the narrative to be appended.
'  DisplayForm - Flag indicating whether or not to display the note edit/entry form.
'    True - Display the Note Entry form
'    False - Don't display the form
'
'Returns:
'  True - Successfully added narrative
'  False - Either was not able to add the narrative, or there was no current encounter selected.'
'
'Note: If the CurrentEncounter object is not set, this method will exit without action.
'
'This method is used by interally by Enctr to display previously stored narratives
'    and is made available to external objects to "Cite" data to the current encounter.
'
'----------------------------------------------------
Public Function AddEncounterNote(ModuleName As String, _
                                 NoteText As String, _
                        Optional DisplayForm As Boolean = False) As Boolean
  Dim objNote  As EncNote
  Dim oConcept As ConceptCtrl
  Set oConcept = New ConceptCtrl
  
  If Not DebugMode Then On Error GoTo errHandler
  
#If debugon Then
  Dim DR As DebugRec
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "AddEncounterNote", "Begin", DR
#End If
  
  If mobjCurrEnc Is Nothing Then Exit Function
  
  Set objNote = New EncNote
  
  objNote.EncounterID = mobjCurrEnc.EncounterID
  objNote.DateAdded = Now
  objNote.DataDate = Now
  objNote.IsComplete = True

  If ModuleName = "AMEND" Then
    objNote.Source = "Changes Made to Signed Document"
    objNote.UserNcid = "ADMIN"
  ElseIf mobjCurrEnc.Status = Updating Then
    objNote.Source = "Added After Document Was Signed"
    objNote.UserNcid = mobjUser.GetAttribute("USER_Ncid")
  Else
    objNote.Source = ModuleName
    objNote.UserNcid = mobjUser.GetAttribute("USER_Ncid")
  End If
  oConcept.UniqueId = mobjUser.GetAttribute("USER_Ncid")
  objNote.UserName = oConcept.PrefRep("2000").Representation
  
  '-- get notetext in rtf format
  If Left$(NoteText, 5) <> "{\rtf" Then
    frmPrintPreview.TXTemp.Text = NoteText
    NoteText = frmPrintPreview.TXTemp.RTFSelText
    Unload frmPrintPreview
  End If
  
  objNote.Note = NoteText
  objNote.FlagNew = True
  
'-- only show the form if requested
  If DisplayForm Then
    objNote.FlagChanged = False
    Call frmEncNote.Initialize(objNote)
   Set frmEncNote.objEnctr = Me
   
    frmEncNote.Top = (Screen.Height - frmEncNote.Height) / 3
    frmEncNote.Left = (Screen.Width - frmEncNote.Width) / 2
    
    mobjComm.Message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_ENCOUNTER_APPLICATION
    frmEncNote.Show vbModal
    mobjComm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_ENCOUNTER_APPLICATION
    
    If objNote.FlagDeleted Then Exit Function
    
  Else
    objNote.FlagChanged = True
  End If
  
  If objNote.FlagChanged Then
    '-- it is to be saved and shown as a new note
    'noteid is one-based so since we haven't added it yet, we add one to the count
    objNote.NoteID = mobjCurrEnc.EncNotes.count + 1
    Dim CurrNote As Integer   ' the container index (noteID - 1)
    CurrNote = mobjCurrEnc.EncNotes.count
    mobjCurrEnc.EncNotes.Add objNote, CStr(objNote.NoteID)
    mobjCurrEnc.FlagChanged = True
    
 '-- add a control if needed
      If CurrNote > 0 Then
        Load frmEncParent.RTBNote(CurrNote)
        frmEncParent.RTBNote(CurrNote).Top = frmEncParent.RTBNote(CurrNote - 1).Top + frmEncParent.RTBNote(CurrNote - 1).Height
        frmEncParent.RTBNote(CurrNote).Visible = True
      End If
      
  '-- update sensitivity of encounter based on note
  '  mobjCurrEnc.Sensitivity = objNote.SensitivityLevel
      If mobjCurrEnc.Sensitivity > 0 Then
         frmEncParent.lblSensitivity.Caption = "Sensitive Document (" & gsEncSensitivity(mobjCurrEnc.Sensitivity) & ")"
      Else
        frmEncParent.lblSensitivity.Caption = ""
      End If
'-- update view in encsum
      Call frmEncParent.FillEncNote(CurrNote, objNote)
    
  Else
    'nothing to do
  End If
  
  DoEvents
  Set objNote = Nothing
  
  AddEncounterNote = True
  
#If debugon Then
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "AddEncounterNote", "End", DR
#End If

  Exit Function
errHandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "Enctr.AddEncounterNote", "EncounterParent", vbExclamation)
    
    Exit Function


End Function
'-----------------------------------------------------
'Purpose:  Pushes all enc notes onto the main form
'
'Parameters: None
'
'Returns: True on success
'
'Comments:
'----------------------------------------------------
Friend Function DisplayNotes() As Boolean
'displays all notes in the curr enc note collection

  Dim i As Integer
  Dim objNote As EncNote
  If Not DebugMode Then On Error GoTo errHandler
  
#If debugon Then
  Dim DR As DebugRec
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "AddEncounterNote", "Begin", DR
#End If
  
  '-- drop all the loaded rtb's over '0'
'-Mod 45e
  On Error Resume Next
  For i = frmEncParent.RTBNote.count - 1 To 1 Step -1
    Unload frmEncParent.RTBNote(i)
  Next i
  If Not DebugMode Then On Error GoTo errHandler
'~end mod 45e

'-Mod 2e Added permission check
  If Not mobjUser.HasPrivilegeEx(Priv_Current_Encounter, Priv_Write) Then
    frmEncParent.cmdNotes.Enabled = False
  End If
'~end mod 2

i = 0
For Each objNote In mobjCurrEnc.EncNotes
  If Not objNote.FlagDeleted Then
   objNote.FlagChanged = False
   objNote.FlagNew = False
   
  '-- add a control if needed
     If i > 0 Then
       Load frmEncParent.RTBNote(i)
       frmEncParent.RTBNote(i).Top = frmEncParent.RTBNote(i - 1).Top + frmEncParent.RTBNote(i - 1).Height
       frmEncParent.RTBNote(i).Visible = True
     End If
     
     Call frmEncParent.FillEncNote(i, objNote)
     
   i = i + 1
  End If       '-Not objNote.FlagDeleted Then
Next objNote

  Set objNote = Nothing
  DoEvents
  
#If debugon Then
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "DisplayNotes", "End", DR
#End If

  DisplayNotes = True

  Exit Function
errHandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "Enctr.DisplayNotes", "EncounterParent", vbExclamation)
    DisplayNotes = False
    Exit Function

End Function

'-----------------------------------------------------
'Purpose: (DEAD CODE)
'
'Parameters:
'
'Returns:
'
'Comments:
'----------------------------------------------------
Public Function EditEncounterNote(objNote As EncNote, _
                                  DisplayForm As Boolean) As Boolean
  
#If debugon Then
  Dim DR As DebugRec
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "EditEncounter", "Begin", DR
#End If
  If Not DebugMode Then On Error GoTo errHandler
  
  objNote.FlagNew = False
  
'-- only show the form if requested
  If DisplayForm Then
    objNote.FlagChanged = False
    Call frmEncNote.Initialize(objNote)
    frmEncNote.objEnctr = Me
    
    frmEncNote.Top = (Screen.Height - frmEncNote.Height) / 3
    frmEncNote.Left = (Screen.Width - frmEncNote.Width) / 2
    
    mobjComm.Message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_ENCOUNTER_APPLICATION
    frmEncNote.Show vbModal
    mobjComm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_ENCOUNTER_APPLICATION
  Else
    objNote.FlagChanged = True
  End If
  
  If objNote.FlagChanged Then
    '-- it is to be saved and shown as a new note
'    mobjCurrEnc.EncNotes.Add objNote
  Else
    'nothing to do
  End If
  
  Set objNote = Nothing
#If debugon Then
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "EditEncounter", "End", DR
#End If

  Exit Function
errHandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "Enctr.EditEncounterNote", "EncounterParent", vbExclamation)
    
    Exit Function

End Function

'-----------------------------------------------------
'Purpose:  (Stub for legacy apps)
'
'Parameters:  (see CreateEncounterEx)
'
'Returns:
'
'Comments:
'----------------------------------------------------
Public Function CreateEncounter(ApptID As Long, _
                                 ApptType As String, _
                                 Enc_Type As EncType, _
                                 FacilityNCID As String, _
                                 ClinicNCID As String, _
                                 PrimaryProviderNCID As String, _
                        Optional ClerkNote As String = "") As Encounter
                       
Set CreateEncounter = CreateEncounterEx(ApptID, _
                                        ApptType, _
                                        Enc_Type, _
                                        FacilityNCID, _
                                        ClinicNCID, _
                                        PrimaryProviderNCID, _
                                        ClerkNote)
End Function

'-----------------------------------------------------
'Purpose:  Creates a new encounter based on information supplied and the current patient
'
'Parameters:  AppID - Called by appt so this is pointer back to appt
'             Appttype - In/Out/Amb/Telecons/Dental
'             Facilityncid -
'             ClinicNCID -
'             Clerknote - any note passed by a clerk (used for telecons)
'Returns:  The encounter created
'
'Comments: NOTE: it gets an encounter id from the add method
'----------------------------------------------------
Public Function CreateEncounterEx(ApptID As Long, _
                                 ApptType As String, _
                                 Enc_Type As EncType, _
                                 FacilityNCID As String, _
                                 ClinicNCID As String, _
                                 PrimaryProviderNCID As String, _
                        Optional ClerkNote As String = "") As Encounter
Dim oEnc As Encounter
  Set oEnc = New Encounter
  oEnc.Init mobjUser
  
Dim oOcc As Occurrence
  Set oOcc = New Occurrence
  If oOcc Is Nothing Then
    'we failed to created and enc
    MsgBox "Failed to create an encounter for this appointment", vbCritical, "Encounter Creation Error"
  End If
  
  oOcc.UnitNumber = mobjPatient.UnitNumber
  oOcc.Facility = FacilityNCID
  oOcc.ArrivalTime = Now

  oEnc.EncounterID = Val(oOcc.Add)
  
  Set oOcc = Nothing
  If oEnc.EncounterID = 0 Then Exit Function
  Set oEnc.Patient = mobjPatient
  
  If oEnc.Create(ApptID, ApptType, Enc_Type, FacilityNCID, ClinicNCID, PrimaryProviderNCID, ClerkNote) Then
    Set CreateEncounterEx = oEnc
    RaiseEvent CreatedEncounter(oEnc)
  
  '-020400 WAS: patient enc num has not been set  yet.  what's eric doing with the parent obj?
    Dim patEnc As String
    Dim patEncFac As String
    patEnc = mobjPatient.EncounterNumber
    patEncFac = mobjPatient.EncounterFacilityNCID
    
    mobjPatient.EncounterNumber = oEnc.EncounterID
    mobjPatient.EncounterFacilityNCID = oEnc.FacilityNCID
    
  '-- 12.26.99 WAS - Added insurance-copy from Eric Monk
    Dim oIns As Object
    
    Set oIns = CreateObject("CHCSII_Insurance.InsuranceOps")
    On Error GoTo InsErr
    oIns.CopyPatientInsuranceToEncounter mobjPatient, mobjLogonEnvironment
    
    mobjPatient.EncounterNumber = patEnc
    mobjPatient.EncounterFacilityNCID = patEncFac
InsErr:
    Set oEnc = Nothing
    Set oIns = Nothing
  End If
 
End Function

Private Function AddNoteToNewEncounter(oEnc As Encounter, Source As String, NoteText As String)
  Dim objNote  As EncNote
  Dim oConcept As ConceptCtrl
  Set oConcept = New ConceptCtrl
  
  If Not DebugMode Then On Error GoTo errHandler
  
#If debugon Then
  Dim DR As DebugRec
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "AddNoteToNewEncounter", "Begin", DR
#End If
  
  If oEnc Is Nothing Then Exit Function
  
  Set objNote = New EncNote
  
  objNote.EncounterID = oEnc.EncounterID
  objNote.DateAdded = Now
  objNote.DataDate = Now
  objNote.IsComplete = True
  objNote.Source = Source
  objNote.UserNcid = mobjUser.GetAttribute("USER_Ncid")
  oConcept.UniqueId = mobjUser.GetAttribute("USER_Ncid")
  objNote.UserName = oConcept.PrefRep("2000").Representation
  
  '-- get notetext in rtf format
  
  If Left$(NoteText, 5) <> "{\rtf" Then
    NoteText = mobjUtil.ConvertTextToRTF(NoteText)
  End If
  objNote.Note = NoteText
  objNote.FlagNew = True
  objNote.FlagChanged = True
  objNote.NoteID = oEnc.EncNotes.count + 1
  oEnc.EncNotes.Add objNote, CStr(objNote.NoteID)
  oEnc.FlagChanged = True
  DoEvents
  Set objNote = Nothing
  AddNoteToNewEncounter = True
  
#If debugon Then
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "AddNoteToNewEncounter", "End", DR
#End If

  Exit Function
errHandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "Enctr.AddNoteToNewEncounter", "EncounterParent", vbExclamation)
    Set objNote = Nothing
    Exit Function

End Function
Friend Function UnloadEnctr() As Boolean

'-- get rid of all the enctr stuff
    
    mobjPatient.EncounterNumber = ""
    mobjPatient.EncounterFacilityNCID = ""
    
    Call RemoveMenus
    Call UnloadChildren
'-Mod 7e
    If Not mobjCurrEnc Is Nothing Then
      If mobjCurrEnc.FlagChanged Then
        mobjCurrEnc.Save
      End If
    End If
    
    RaiseEvent CloseEncounter(mobjCurrEnc)
    Set gobjShared = New CWShared
    Set mcolEncChild = New Collection
    Set mcolEncounters = New Collection
    Set mobjCurrEnc = Nothing
    Set mcolMenus = New Collection
    
End Function

'-----------------------------------------------------
'Purpose: (Legacy stub; see UpdateEncounterEX)
'
'Parameters:
'
'Returns:
'
'Comments:
'----------------------------------------------------
Public Function UpdateEncounter(objEnc As Encounter) As Boolean
Attribute UpdateEncounter.VB_Description = "The ""Save"" command for the specifid encounter object"
  UpdateEncounter = UpdateEncounterEx(objEnc)
End Function
  

'-----------------------------------------------------
'Purpose: Saves an encounter object back to cdr
'
'Parameters: the encounter to save
'
'Returns:  True on success
'
'Comments: Assumes oenc.flagchanged is 'True'
'----------------------------------------------------
Public Function UpdateEncounterEx(oEnc As Encounter) As Boolean
Attribute UpdateEncounterEx.VB_Description = "Saves the specified encounter to the CDR+ database"
'-Mod 18e  for the others who don't k now about my flagchanged
     oEnc.FlagChanged = True
     UpdateEncounterEx = oEnc.Save
End Function
    


'-----------------------------------------------------
'Purpose: (Dead Code. described in ichcsii_encsum)
'
'Parameters:
'
'Returns:
'
'Comments:
'----------------------------------------------------
Public Sub SummaryFormHeightChanged(ServerName As String, _
                                   NewHeight As Integer)
                                   
'--12.29.99 WAS - this is now OBE
  '-- the specific child object has changed it's height.
  '-- Look for this hWnd and adjust as needed.
  If Not DebugMode Then On Error GoTo errHandler

#If debugon Then
  Dim DR As DebugRec
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "EncSumFormHeightChanged", "Begin", DR
#End If
    
    Dim oChild As EncChild
    If bMainViewRunning Then
      '-- tell all children about this
      For Each oChild In mcolEncChild
        If oChild.ServerName = ServerName Then
          'found it
          oChild.ClientControl.Height = NewHeight
          Set oChild = Nothing
          Exit Sub
        End If
      Next oChild
    End If
    Set oChild = Nothing
    
#If debugon Then
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "EncSumFormHeightChanged", "End", DR
#End If
 
   Exit Sub
errHandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "Enctr.SummaryFormHeightChanged", "EncounterParent", vbExclamation)
    
    Exit Sub

End Sub



'-----------------------------------------------------
'Purpose:  Appends a string to an enc sum section
'
'Parameters:  sappNCID - The section affected
'             rtfText - the text to append
'             dataData - the date to show with this data
'             dataStatus - status of the data
'             DataOwner - who did this
'             sDataOwner - just the name (used internallay and by s/o)
'
'Returns:  True on success
'
'Comments:
'----------------------------------------------------
Public Function AddToEncounterData(sAppNCID As String, _
                                   RTFText As String, _
                                   DataDate As Date, _
                                   DataStatus As String, _
                                   DataOwner As ICHCSII_User, _
                          Optional sDataOwner As String) As Boolean
'----------------------------------------------------
' added to allow the child sections to control the contents of the encounter document
'----------------------------------------------------
  Dim oEncChild As EncChild
  Dim oSectRTF As EncRTF
  
  Dim idx As Integer
  
  If Not DebugMode Then On Error GoTo errHandler
#If debugon Then
  Dim DR As DebugRec
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "AddToEncounterData", "Begin", DR, "Data from Child", RTFText
#End If
  
  If RTFText = "" Then Exit Function
  
  idx = -1
  
  For Each oEncChild In mcolEncChild
    If oEncChild.AppNCID = sAppNCID Then
      idx = oEncChild.ServerIndex
      Exit For
    End If
  Next oEncChild
  
  If idx = -1 Then
    MsgBox "Program Error: Unable to find Encounter Control for '" & sAppNCID & "'"
    GoTo ExitProc
  End If
  
On Error Resume Next
  Set oSectRTF = mobjCurrEnc.SectionRTFs(CStr(idx))
  If oSectRTF Is Nothing Then
    Set oSectRTF = New EncRTF
    oSectRTF.SectionIndex = idx
    mobjCurrEnc.SectionRTFs.Add oSectRTF, CStr(idx)
  End If

  
'  If DataOwner Is Nothing Then
'    Set oSectRTF.OwnerNCID = mobjUser
'  Else
'    Set oSectRTF.OwnerNCID = DataOwner
'  End If

  If IsNumeric(DataStatus) Then
    oSectRTF.Status = DataStatus
  Else
    oSectRTF.Status = rtf_Complete
  End If
  
  If IsDate(DataDate) And DataDate > 0 Then
    oSectRTF.DTS = DataDate
  Else
    oSectRTF.DTS = Now
  End If
  
Dim RTF As String
Dim UName As String
Dim UID As String
Dim oConcept As New ConceptCtrl

  If sDataOwner = "" Then
      UID = DataOwner.GetAttribute("USER_Ncid")
      oConcept.UniqueId = UID
      UName = oConcept.PrefRep("2000").Representation
  Else
    UName = sDataOwner
    UID = sDataOwner
  End If
  
'--020200 WAS- allow regular strings
    If Left$(RTFText, 5) <> "{\rtf" Then
       RTFText = mobjUtil.ConvertTextToRTF(RTFText)
    End If
  
 '--020200 WAS - removed rtf doc-info section
'-- build the wrapper for the rtf text
    RTF = "{\rtf1\ansi\ansicpg1252\uc1 \deff0\deflang1033\deflangfe1033" _
    & "{\fonttbl{\f0\fswiss\fcharset0\fprq2 Arial;}}" _
    & "{\colortbl;\red0\green0\blue0;}"
    RTF = RTF & "\pard\plain \widctlpar\adjustright \fs16\cgrid \tx" & frmEncParent.RTB(idx).Width \ 2 _
    & "{\f0\fs16\b\ul " & Sect(oSectRTF.SectionIndex) & "\plain\f0\fs16 \tab [" & UName & ", " & mobjUtil.FormatDTG$(oSectRTF.DTS) & "]\par }" _
    & "{\pard\plain\f0\fs16 {" & oSectRTF.Doc & " \par}{" & RTFText & "}}}"
  
  
  Dim tmp As String
  
  Select Case idx
'-- shouldn't ever be zero, either
  Case 0
    frmEncParent.RTBAC(0).Text = ""
    frmEncParent.RTBAC(0).RTFSelText = RTF
    oSectRTF.Doc = frmEncParent.RTBAC(0).RTFSelText
    
  Case Else
    frmEncParent.RTB(idx).AutoExpand = True
    frmEncParent.RTB(idx).Text = ""
    frmEncParent.RTB(idx).Height = 500
    
    frmEncParent.RTB(idx).RTFSelText = RTF
    frmEncParent.pics(idx).Height = frmEncParent.RTB(idx).Height
    DoEvents
    oSectRTF.Doc = oSectRTF.Doc & " \par" & RTFText
  End Select
  
ExitProc:
  Set oSectRTF = Nothing
  Set oEncChild = Nothing
#If debugon Then
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "AddToEncounterData", "End", DR
#End If
  
  
    AddToEncounterData = True
    Exit Function
errHandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "Enctr.AddToEncounterData", "EncounterParent", vbExclamation)
    
    Exit Function

End Function

'-----------------------------------------------------
'Purpose:  Replaces current text in a section with that supplied
'
'Parameters:  sappNCID - The section affected
'             rtfText - the text to append
'             dataData - the date to show with this data
'             dataStatus - status of the data
'             DataOwner - who did this
'             sDataOwner - just the name (used internallay and by s/o)
'
'Returns:  True on success
'
'Comments:
'----------------------------------------------------
Public Function ReplaceEncounterData(sAppNCID As String, _
                                     RTFText As String, _
                                     DataDate As Date, _
                                     DataStatus As String, _
                                     DataOwner As ICHCSII_User, _
                            Optional sDataOwner As String) As Boolean

' allows the child sections to control the contents of the encounter document
  
  Dim oEncChild As EncChild
  Dim oSectRTF As EncRTF
  Dim RTF As String
  Dim idx As Integer
#If debugon Then
  Dim DR As DebugRec
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "ReplaceEncounterData", "Begin", DR
#End If
  
  If Not DebugMode Then On Error GoTo errHandler
  
  idx = -1
  For Each oEncChild In mcolEncChild
    If oEncChild.AppNCID = sAppNCID Then
      idx = oEncChild.ServerIndex
      Exit For
    End If
  Next oEncChild
  
  If idx = -1 Then
      MsgBox "Program Error: Unable to find Encounter Control for : " & sAppNCID
  End If
  
On Error Resume Next
   If sDataOwner = "" Then
Dim UName As String
Dim UID As String
Dim oConcept As New ConceptCtrl
   
    UID = DataOwner.GetAttribute("USER_Ncid")
    If UID <> "" Then
      oConcept.UniqueId = UID
      UName = oConcept.PrefRep("2000").Representation
    End If
   Else
     UName = sDataOwner
'-Mod 3e   UID = sDataOwner
   End If     '--If sDataOwner = "" Then

'-- get the proper rtf object
  Set oSectRTF = mobjCurrEnc.SectionRTFs(CStr(idx))
  If oSectRTF Is Nothing Then
    Set oSectRTF = New EncRTF
    oSectRTF.SectionIndex = idx
    mobjCurrEnc.SectionRTFs.Add oSectRTF, CStr(idx)
  End If

'--111099 always reset the owner
    oSectRTF.OwnerNCID = UID
'-Mod 3e
    oSectRTF.OwnerName = UName

  If IsDate(DataDate) And DataDate > 0 Then
    oSectRTF.DTS = DataDate
  Else
    oSectRTF.DTS = Now
  End If

'--020200 WAS- allow regular strings
    If Left$(RTFText, 5) <> "{\rtf" Then
       RTFText = mobjUtil.ConvertTextToRTF(RTFText)
    End If
'-Mod 60e don't do this if there's no data
  If RTFText = "" Then
    RTF = ""
  Else
'-- build the wrapper for the rtf text
    RTF = "{\rtf1\ansi\ansicpg1252\uc1 \deff0\deflang1033\deflangfe1033" _
    & "{\fonttbl{\f0\fswiss\fcharset0\fprq2 Arial;}}" _
    & "{\colortbl;\red0\green0\blue0;}"
    RTF = RTF & "\pard\plain \widctlpar\adjustright \fs16\cgrid \tx" & frmEncParent.RTB(idx).Width \ 2 _
    & "{\par\f0\fs16\b\ul " & Sect(oSectRTF.SectionIndex) & "\plain\f0\fs16 \tab [" & UName & "; " & mobjUtil.FormatDTG$(oSectRTF.DTS) & "]\par }" _
    & "{\pard\f0\fs16 {" & RTFText & "}}}"
  End If
'-- 021500 WAS: add case for "updating" status
'-- no, this happens in validatencounter during close (see "recordamendments")
'  If mobjCurrEnc.Status = Updating And RTFtext <> oSectRTF.Doc Then
'
'  End If
  Select Case idx
  Case 0
    frmEncParent.RTBAC(0).AutoExpand = True
    frmEncParent.RTBAC(0).Text = ""
    frmEncParent.RTBAC(0).Height = 500
    frmEncParent.RTBAC(0).RTFSelText = RTF
    oSectRTF.Doc = frmEncParent.RTBAC(0).TRTFSelText
  Case Else
    frmEncParent.RTB(idx).AutoExpand = True
    frmEncParent.RTB(idx).Text = ""
    frmEncParent.RTB(idx).Height = 500
    frmEncParent.RTB(idx).RTFSelText = RTF
    frmEncParent.pics(idx).Height = frmEncParent.RTB(idx).Height
    DoEvents
    oSectRTF.Doc = RTFText
  End Select
'-- 1.1.00 WAS - special case. need to specifically reset the Note section
  If idx = SECT_DISP Then
    
  End If
  
  Set oSectRTF = Nothing
  Set oEncChild = Nothing
#If debugon Then
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "ReplaceEncounterData", "End", DR
#End If
  
  Exit Function
errHandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "Enctr.ReplaceEncData", "EncounterParent", vbExclamation)
    
    Exit Function

End Function


'-----------------------------------------------------
'Purpose:
'
'Parameters:
'
'Returns:
'
'Comments:
'----------------------------------------------------
Public Function RunChartNotes(sAppNCID As String, _
                              sTemplate As String, _
                     Optional DisplayInSO As Boolean) As Boolean
  '--sAppNCID -who 's asking it to run,
  '--stemplate - which template
 
#If debugon Then
  Dim DR As DebugRec
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "RunChartNotes", "Begin", DR
#End If
  Screen.MousePointer = vbHourglass
  
If DisplayInSO Then
  Call OpenEncChild(SECT_SO, "", "", sTemplate)
Else
  Call OpenEncChild(SECT_SO, sAppNCID, "", sTemplate)
End If

  Screen.MousePointer = vbDefault
Exit Function

#If debugon Then
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "RunChartNotes", "End", DR
#End If
  
  Exit Function
  
errHandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "Enctr.RunChartNotes", "EncounterParent", vbExclamation)
    Call mobjComm.setstatustext("")
    Exit Function

End Function

'-----------------------------------------------------
'Purpose:
'
'Parameters:
'
'Returns:
'
'Comments:
'----------------------------------------------------
Friend Function ValidateEncounterData(oEnc As Encounter) As Boolean
'-- do a dirty read to check for save options
'Dim AppendChanges As Boolean
'Dim AppendChangesText As String
'Dim i As Integer
'Dim Row As Long
'Dim bDataChanged As Boolean
'Dim bNoteChanged As Boolean
'Dim rows As Long
'Dim CurrStatus As EncStatus
Dim oConcept As ConceptCtrl
Set oConcept = New ConceptCtrl

#If debugon Then
  Dim DR As DebugRec
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "ValidateEncounterData", "Begin", DR
#End If

'-Mod 61e  check the current actionInfo on the enc.
  If Not HandleChangedEncounter(oEnc) Then
    ValidateEncounterData = False
    Exit Function
  End If
'-- if the status is updating then we have to sign the changes.
If oEnc.Status = Updating Then
  
'-- first record any amendments that were made
  Call RecordAmendments(oEnc)
  
  Dim mEncChild As EncChild     '- each section of the enc doc is a child obj
  Dim objSign As EncSignOps
  Set objSign = New EncSignOps
  Set objSign.calledby = Me
  
'-- the objsign needs a sigblock object (build one based on current user)
  Dim objSigBlock As SigBlock
  Set objSigBlock = New SigBlock
  objSigBlock.DTS = Now
  objSigBlock.UserNcid = mobjUser.GetAttribute("USER_Ncid")

  oConcept.UniqueId = mobjUser.GetAttribute("USER_Ncid")
  objSigBlock.Line1 = oConcept.PrefRep("2000").Representation
  objSigBlock.Line2 = mobjLogonEnvironment.RoleName
  objSigBlock.Line3 = mobjLogonEnvironment.FacilityName
  
'-- setup the objsign and fire the signEncounter method
  objSign.SigBlock = objSigBlock
  objSign.SigCoSignReq = False
  objSign.Signer = mobjUser.GetAttribute("USER_Ncid")
  objSign.SigUser = mobjUser
'-Mod 41e set the amend flag so we use the proper sig form
  objSign.AmendOnly = True

'-- in SignEncounter, we group the sectional RTFs into one RTF and add signature information.
'   Then add that to the collection of signed rtfs for this encounter.
  If objSign.SignEncounter Then  'it was signed and the enc status must change
'-Mod 37e  - Not always true
'    oEnc.Status = Updated
'    oEnc.EndDTS = Now
'~
    ValidateEncounterData = True
  
  Else
    ValidateEncounterData = False
  End If     '-- objSign.SignEncounter Then  'it was signed and the enc status must change

  Set objSign = Nothing
'-Mod 37e still heed the positive return
Else
    ValidateEncounterData = True
End If   '-- oEnc.Status = Updating Then
'-Mod 37e  This was causing a "save" operation even if the user hit cancel on the sign dialog
'ValidateEncounterData = True

#If debugon Then
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "ValidateEncounterData", "End", DR
#End If
  Set oConcept = Nothing

  
      Exit Function
errHandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "Enctr.ValidateEncounterData ", "Enc", vbCritical)
  Set oConcept = Nothing

    Exit Function

End Function


'-----------------------------------------------------
'Purpose:   Adds a menu item to either Action or View
'
'Parameters:
'  iMenu  - one of 0 (Action) or 1 (View)
'  sMenuText - the text for the menu item
'  sMenuBMPFileName - the filename of the bitmap icon file (empty string for 'view')
'  sTip - the tool tip text string
'
'Returns:   The menu ID of the new menu item
'
'Comments:  called by encounter's child objects to put things on the core menu
'       masgarading As Encounter
'
'----------------------------------------------------
Public Function AddMenuItem(iMenu As MenuNames, _
                            sMenuText As String, _
                            sMenuBMPFileName As String, _
                            sTip As String) As Long
  
  If Not DebugMode Then On Error GoTo errHandler

#If debugon Then
  Dim DR As DebugRec
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "AddMenuItem", "Begin", DR
#End If


  If iMenu = mnAction Then
    AddMenuItem = mobjComm.AddActionMenuItem(sMenuText, sMenuBMPFileName, mlProgID, sTip)
  Else
    AddMenuItem = mobjComm.AddViewmenuitem(sMenuText, mlProgID, sTip)
  End If

'-- just so I can take them down by myself...

'-- add to the childmenu list  (used to eliminate the child menus when core takes over)
 Dim MenuIDs() As String
 Dim x As Integer
 Dim AddOne As Boolean
 
 MenuIDs = Split(sChildMenuIDs, "|")
 x = UBound(MenuIDs)
 If x = -1 Then
  ReDim MenuIDs(0)
  x = 0
 End If
 If x > 0 Then
   AddOne = True
 Else
  If MenuIDs(0) <> "" Then
    AddOne = True
  End If
 End If
 If AddOne Then
  ReDim Preserve MenuIDs(x + 1)
  x = x + 1
 End If
  
 MenuIDs(x) = AddMenuItem
 '--put back into string
 sChildMenuIDs = Join(MenuIDs, "|")
  
#If debugon Then
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "AddMenuItem", "End", DR
#End If

  
  Exit Function
errHandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "Enctr.AddMenuItem", "EncounterParent", vbExclamation)
    Exit Function
  
End Function
 
'-----------------------------------------------------
'Purpose:  core just told me it's opening a form.  See if it's the one that I had up.
'   If so, drop it's menu's (using the schildmenuIDs thing here), put away the child playground.
'
'Parameters:  waaahhh - application ncid
'
'Comments:
'----------------------------------------------------
Private Sub CoreTookMyForm(waaahhh As String)

Dim oTempc As ICHCSII_EncSum2
Dim oTempE4 As ICHCSII_EncSum4
Dim oChild As EncChild
Dim MenuIDs() As String
Dim i As Integer
Dim RTFText As String
Dim DataDate As Date
Dim DataStatus As String
Dim DataOwner As String
  
  If Not DebugMode Then On Error GoTo errHandler
#If debugon Then
  Dim DR As DebugRec
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "CoreTookMyForm", "Begin", DR
#End If
  
  For Each oChild In mcolEncChild
    If oChild.AppNCID = waaahhh And oChild.RegularViewRunning Then
      '-- drop menus
      MenuIDs = Split(sChildMenuIDs, "|")
      
      For i = LBound(MenuIDs) To UBound(MenuIDs)
        If MenuIDs(i) <> "" Then
          mobjComm.deletemenuitem MenuIDs(i)
        End If
      Next i
      
      sChildMenuIDs = ""
            
      frmEncParent.lblChild.Caption = ""
      frmEncParent.vseChildClient.Visible = False
      frmEncParent.VP1.Visible = True
            
      RTFText = oChild.GetRTFNote()
      DataDate = oChild.GetDataDate()
      DataStatus = oChild.GetStatus()
      DataOwner = oChild.GetDataOwner()
      
'--012200 WAS - added to let child know he's free
      Call oChild.ResetParentObject
    
      oChild.RegularViewRunning = False
      Call ReplaceEncounterData(oChild.AppNCID, RTFText, DataDate, DataStatus, mobjUser, DataOwner)

      Exit For
    End If    '-- oChild.AppNcid = AppNcid And oChild.RegularViewRunning
  Next oChild
  Set oChild = Nothing
  
#If debugon Then
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "CoreTookMyForm", "End", DR
#End If

  Exit Sub
errHandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "Enctr.AddMenuItem", "EncounterParent", vbExclamation)
    Exit Sub

End Sub

'-----------------------------------------------------
'Purpose:
'
'Parameters:
'
'Returns:
'
'Comments:
'----------------------------------------------------

Private Function RecordAmendments(oEnc As Encounter) As Boolean
'-Mod 56e - need this to be public
  RecordAmendments = RecordAmendmentsEx(oEnc)
End Function
'-Mod 56e This is the same code as above, just public so it can be hit from the sign menu/encamendsign form
Public Function RecordAmendmentsEx(oEnc As Encounter) As Boolean

'-- use sOriginals() to look for changes.
'-- was filled in encchild.startmodule()
  Dim msg As String
  msg = ""
  Dim msgHeader As String
  Dim msgFooter As String
  Dim oConcept As ConceptCtrl
  
#If debugon Then
  Dim DR As DebugRec
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "RecordAmendmentsEx", "Begin", DR
#End If
    
  msgHeader = "{\rtf1\ansi\ansicpg1252\uc0\deff0" _
      & "{\fonttbl{\f0\fswiss\fcharset0\fprq2 Arial;}}" _
      & "{\colortbl;\red0\green0\blue0;}" _
      & "\deftab720 \plain"
  msgFooter = "}"
  
  Dim oSectRTF As EncRTF
  Dim OldText As String
  Dim newText As String
  Dim CreatedBy As String
  Set oConcept = New ConceptCtrl
  
  For Each oSectRTF In oEnc.SectionRTFs
  
    frmEncParent.TXTemp.Text = ""
    frmEncParent.TXTemp.RTFSelText = oSectRTF.Doc
    OldText = frmEncParent.TXTemp.Text
    frmEncParent.TXTemp.Text = ""
    frmEncParent.TXTemp.RTFSelText = sOriginals(oSectRTF.SectionIndex).Doc
    newText = frmEncParent.TXTemp.Text
    frmEncParent.TXTemp.Text = ""
  
    If newText <> OldText And OldText <> "" Then
      '-- something has to be done
      ' say: the section named (sect) was (original) written by (original user) on (original date)
      '  to (current) by (user) on (now)
      oSectRTF.OwnerNCID = mobjUser.GetAttribute("USER_Ncid")
      
      oConcept.UniqueId = oSectRTF.OwnerNCID
      oSectRTF.OwnerName = oConcept.PrefRep("2000").Representation
      
      If sOriginals(oSectRTF.SectionIndex).CreatedBy <> "" Then
        CreatedBy = "by " & sOriginals(oSectRTF.SectionIndex).CreatedBy
      End If
      
      msg = msg & "{\plain\f0\fs16\b\ul \pard Signed Document Amended." _
      & "\plain\f0\fs16  by \plain\f0\fs16 " & oSectRTF.OwnerName _
      & "\plain\f0\fs16  on \plain\f0\fs16 " & Format$(oSectRTF.DTS, gsDefaultDateTime) _
      & "\plain\f0\fs16 \par \plain\f0\fs16\ul " & Sect(oSectRTF.SectionIndex) _
      & "\plain\f0\fs16  (as entered " & CreatedBy _
      & "\plain\f0\fs16  on \plain\f0\fs16 " & sOriginals(oSectRTF.SectionIndex).CreatedOn _
      & "\plain\f0\fs16 )\par Original:\par " & sOriginals(oSectRTF.SectionIndex).Doc _
      & "\par}"

    End If
    
  Next oSectRTF
  
  If msg <> "" Then
    '-- send to enc
    msg = msgHeader & msg & msgFooter
    
    AddEncounterNote "AMEND", msg, False
  End If
  
  RecordAmendmentsEx = True
  
#If debugon Then
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "RecordAmendmentsex", "End", DR
#End If
  
End Function

'-----------------------------------------------------
'Purpose:  Prev Enc comes here to let user amend an enc
'
'Parameters: none
'
'Comments:
'----------------------------------------------------
Private Sub AmendEnc()

  Dim frmwide As Long
  Dim frmhigh As Long
  Dim ldummy As Long
  Dim bUnloadedForm As Boolean
  
#If debugon Then
  Dim DR As DebugRec
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "AmendEnc", "Begin", DR
#End If
   
'-- if we are not displaying the proper encounter, then close and open on that one
  
'-Mod 53e check for valid curr enc object
If bMainViewRunning Then
  If Not mobjCurrEnc Is Nothing Then
    If Not (Val(mobjPatient.EncounterNumber) = mobjCurrEnc.EncounterID) Then
        frmwide = frmEncParent.Width
        frmhigh = frmEncParent.Height
        frmEncParent.bCloseMyself = False  '--a lie; so we don't tell core about this
        Unload frmEncParent
        DoEvents
        bUnloadedForm = True
        bFirstRegularResize = True
        OpenView cwiREGULAR_VIEW, glParentWindow, ldummy
        ResizeView cwiREGULAR_VIEW, 0, 0, frmwide, frmhigh
    Else
     '-nop
    End If
  Else
    frmwide = frmEncParent.Width
    frmhigh = frmEncParent.Height
    frmEncParent.bCloseMyself = False  '--a lie; so we don't tell core about this
    Unload frmEncParent
    bUnloadedForm = True
    bFirstRegularResize = True
    OpenView cwiREGULAR_VIEW, glParentWindow, ldummy
    ResizeView cwiREGULAR_VIEW, 0, 0, frmwide, frmhigh
  
  End If
Else
  Call mobjComm.Message(cwiSTART_OLE_SERVER, NCID_ENCOUNTER_APPLICATION, NCID_CW_APPLICATION, NCID_ENCOUNTER_APPLICATION)
  DoEvents
End If

'-- if status was "Needs cosignature", now it's "updating"
   mobjCurrEnc.Status = Updating
   If bMainViewRunning Then
    frmEncParent.lblStatus.Caption = gsEncStatus(mobjCurrEnc.Status)
   End If
   
#If debugon Then
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "AmendEnc", "End", DR
#End If

End Sub


Public Sub TellAnPComplete()
'===== functionality moved to mobjCurrEnc_EncStatusChanged()
End Sub


'-----------------------------------------------------
'Purpose:
'
'Parameters:
'
'Returns:
'
'Comments:
'----------------------------------------------------
Public Function AddProvider(Optional EncounterID As String) As Boolean
'-- adds the selected user as a provider to the encounter's collection of providers.
#If debugon Then
  Dim DR As DebugRec
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "AddProvider", "Begin", DR
#End If

  On Error GoTo errHandler
  
  Select Case True
  Case EncounterID = "" And mobjCurrEnc Is Nothing
    MsgBox "No Current Encounter"
    Exit Function
  Case mobjPatient Is Nothing
    MsgBox "No Current Patient"
    Exit Function
  Case EncounterID <> mobjCurrEnc.EncounterID
    Set mobjCurrEnc = GetEncounter(mobjPatient.UnitNumber, mobjPatient.FacilityNCID, EncounterID)
    If mobjCurrEnc Is Nothing Then
      MsgBox "Encounter " & EncounterID & " Not Found"
      Exit Function
    End If
  End Select
'-Mod 28e
  frmProviderList.Display Me

  frmProviderList.Top = (Screen.Height - frmProviderList.Height) / 3
  frmProviderList.Left = (Screen.Width - frmProviderList.Width) / 2

  mobjComm.Message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_ENCOUNTER_APPLICATION
  frmProviderList.Show vbModal
  mobjComm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_ENCOUNTER_APPLICATION

#If debugon Then
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "AddProvider", "End", DR
#End If
  
  Exit Function
  
errHandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "Enctr.AddProvider", "EncounterParent", vbExclamation)
    Exit Function

  
End Function

Public Function LeftWithoutBeingSeen() As Boolean
'  assumes the patient.encounternumber is the encounter in question
'Rules: if anything has been done in A&P then return FALSE
'else, Set Disposition status to;;;;;
' and  set enc SADRStatus to needs coding
' and  Set Enc status to LeftWithoutBeingSeen
' and build rtf w/o signature block (but with user who did this)

'1. get the enc
'2. check the A&P string
'3. if not "" then return false
'4. else set the status's
'5. build rtf
'6. return true.
Dim oEnc As Encounter

#If debugon Then
  Dim DR As DebugRec
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "LeftWithoutBeingSeen", "Begin", DR
#End If

' assume the enctr was initialized with the patient set

  If Not mobjCurrEnc Is Nothing Then
    If Not (Val(mobjPatient.EncounterNumber) = mobjCurrEnc.EncounterID _
          And mobjPatient.EncounterFacilityNCID = mobjCurrEnc.FacilityNCID) Then    ' we have to switch to that enc
      Set oEnc = GetEncounter(mobjPatient.UnitNumber, mobjPatient.EncounterFacilityNCID, mobjPatient.EncounterNumber)
    Else
      Set oEnc = mobjCurrEnc
    End If
  Else
    Set oEnc = GetEncounter(mobjPatient.UnitNumber, mobjPatient.EncounterFacilityNCID, mobjPatient.EncounterNumber)
  End If
    If oEnc Is Nothing Then
        MsgBox "Failed to load requested encounter", vbCritical, "Encounters"
        Exit Function
    End If
'-Mod 62e be sure we have a fullyloaded enc
    If Not oEnc.FullyLoaded Then
      Call oEnc.LoadEncounter
    End If
    
Debug.Print "enc #", oEnc.EncounterID, "Enc status", oEnc.Status

  '-- we have the proper encounter loaded in currenc
Dim apStatus As EncStatus
Dim SpecialChild As EncChild
    Set SpecialChild = New EncChild
    Set SpecialChild.Module = mobjComm.InitializeOLEServer(CStr(glAppNCID(SECT_AP)))
    If SpecialChild.Module Is Nothing Then
        MsgBox "Unable to initialize modules; Unable to check workflow status.", vbCritical, "Encounters"
        Exit Function
    End If
    Call SpecialChild.EncounterChanged
    
  Call SpecialChild.SetParentObject(Me)
  apStatus = SpecialChild.GetStatus()

Debug.Print "got a/p status of", apStatus
  
  If apStatus = Complete Then
    LeftWithoutBeingSeen = False
    Set SpecialChild.Module = Nothing
    Set SpecialChild = Nothing
    Exit Function
  End If
  
  oEnc.Status = PatientLeftWithoutBeingSeen
  oEnc.SADRStatus = NCID_SADR_NEEDS_NEW_PROCESSING
  oEnc.DispositionNCID = 546
  oEnc.DispositionText = ""
  oEnc.DispAdminOption = ""
  
  '-- build the enc rtf for this one.
  Call BuildRTFforCanceledEncounter(oEnc)
  oEnc.FlagChanged = True
  
  If Not oEnc.Save Then
    LeftWithoutBeingSeen = False
  Else
    LeftWithoutBeingSeen = True
  End If

#If debugon Then
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "LeftWithoutBeingSeen", "End", DR
#End If
Debug.Print "New Enc Status", oEnc.Status

'-03.23.00 Noticed the status line wasn't cleared when completed. so...
  Call mobjComm.setstatustext("")
'--01.28.00
  Set oEnc = Nothing
      
  Exit Function
  
errHandler:
  '---show the VB error---
  Call gobjShared.ShowVBError(Err.Number, Err.Description, "Enctr.LeftWithoutBeingSeen ", "Enc", vbCritical)
  Set oEnc = Nothing
  Exit Function

End Function

Friend Function BuildRTFforCanceledEncounter(oEnc As Encounter) As Boolean
  Dim oConcept As ConceptCtrl
  Dim objSign As EncSignOps
  Dim oRTF As EncRTF
   
  
  Set oConcept = New ConceptCtrl
  
'-- the objsign needs a sigblock object (build one based on current user)
  Dim objSigBlock As SigBlock
  Set objSigBlock = New SigBlock
  objSigBlock.DTS = Now
  oConcept.UniqueId = mobjUser.GetAttribute("USER_Ncid")
  objSigBlock.UserNcid = mobjUser.GetAttribute("USER_Ncid")
  objSigBlock.Line1 = oConcept.PrefRep("2000").Representation
  objSigBlock.Line2 = mobjLogonEnvironment.RoleName
  objSigBlock.Line3 = mobjLogonEnvironment.FacilityName
  
  
  Set oRTF = New EncRTF
  oRTF.DTS = Now
  
  Set oRTF.FinalSignature = objSigBlock
  oRTF.Status = rtf_Complete
  
    
  oRTF.Doc = oEnc.BuildRTF(0, objSigBlock)
  
  Call oRTF.Sign(objSigBlock, 0)
  
  oEnc.EncRTFs.Add oRTF

  Set objSigBlock = Nothing
  Set objSign = Nothing
End Function



Public Sub EncAddProvider()
  frmProviderList.Display Me
  frmProviderList.Top = (Screen.Height - frmProviderList.Height) / 3
  frmProviderList.Left = (Screen.Width - frmProviderList.Width) / 2
    
  mobjComm.Message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_ENCOUNTER_APPLICATION
  frmProviderList.Show vbModal
  mobjComm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_ENCOUNTER_APPLICATION
  
End Sub

Friend Function TellChildrenEncounterChanged() As Boolean

Dim oChild As EncChild
Dim oTempE  As ICHCSII_EncSum4
#If debugon Then
  Dim DR As DebugRec
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "TellChildrenEncounterChanged", "Begin", DR
#End If

  If mcolEncChild Is Nothing Then Exit Function
  
  For Each oChild In mcolEncChild
    If oChild.ModuleLoaded Then
      If oChild.EncounterChanged Then
    #If debugon Then
      If Not DL Is Nothing Then DL.DebugLog Exe, Module, "TellChildrenEncounterChanged", "Mid", DR, "Playing EncSum4", oChild.ServerName
    #End If
      Else
    #If debugon Then
      If Not DL Is Nothing Then DL.DebugLog Exe, Module, "TellChildrenEncounterChanged", "Mid", DR, "NOT Playing EncSum4", oChild.ServerName
    #End If
      oChild.Module.PatientSelected
      End If
    End If
  Next oChild
  
  TellChildrenEncounterChanged = True
  
#If debugon Then
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "TellChildrenEncounterChanged", "End", DR
#End If
  
End Function


Private Function NewClinObs(oClinEvent As ClinicalEvent, _
                            ObsIDValue As String, _
                            obsValValue As String, _
                   Optional ObsIDType As String = "codedWOSform", _
                   Optional obsValType As String = "codedWOSform") As ClinicalObservation
  
  Dim oClinicalObservation As ClinicalObservation
  
  Set oClinicalObservation = oClinEvent.ClinObs.Add
    With oClinicalObservation
        ' ObsId
        '------------------------------------------------
        .ObsId.InfoType = "att"
        .ObsId.Value.InfoType = ObsIDType
        .ObsId.Value.CodedWOSform.ValueConcept.UniqueId = ObsIDValue
        ' ObsValue
        '------------------------------------------------
        .ObsValue.InfoType = "att"
        .ObsValue.Value.InfoType = obsValType
        Select Case obsValType
        Case "codedWOSform"
          .ObsValue.Value.CodedWOSform.ValueConcept = obsValValue
        Case "basicString"
          .ObsValue.Value.BasicString = obsValValue
        End Select
    End With
  
  Set NewClinObs = oClinicalObservation
  
End Function

Private Function NewClinObsMod(oClinObs As ClinicalObservation, _
                               ObsIDValue As String, _
                               obsValValue As String, _
                      Optional ObsIDType As String = "codedWOSform", _
                      Optional obsValType As String = "codedWOSform") As ClinicalObservation
  
  Dim oClinicalObservation As ClinicalObservation
  
  Set oClinicalObservation = oClinObs.ObsMods.Add
    With oClinicalObservation
        ' ObsId
        '------------------------------------------------
        .ObsId.InfoType = "att"
        .ObsId.Value.InfoType = ObsIDType
        .ObsId.Value.CodedWOSform.ValueConcept.UniqueId = ObsIDValue
        ' ObsValue
        '------------------------------------------------
        .ObsValue.InfoType = "att"
        .ObsValue.Value.InfoType = obsValType
        Select Case obsValType
        Case "codedWOSform"
          .ObsValue.Value.CodedWOSform.ValueConcept = obsValValue
        Case "basicString"
          .ObsValue.Value.BasicString = obsValValue
        End Select
    End With
  
  Set NewClinObsMod = oClinicalObservation
  
End Function

Public Function RefreshMedcin(RTFText As String, DataDate As Date, DataOwner As String) As Boolean
  '-- find the tnote that has the medcin data
  Dim oTNote            As TextNote
  Dim i                 As Integer
  Dim OutString         As String
  Dim rasTextStrings()  As String       '- text as it returns from "convert HTML to text" function
  Dim rnNumTextStrings  As Integer      '- Number of above
  Dim msgHeader         As String
  Dim msgFooter         As String
  Dim oEnc              As Encounter   '- current enc
  Dim oConcept          As ConceptCtrl
  Dim sMedcinText       As String
  
  
 ' Set oConcept = New ConceptCtrl
 Set oEnc = mobjCurrEnc
  For i = oEnc.TextNotes.count To 1 Step -1
    Set oTNote = oEnc.TextNotes.Item(i)
    If oTNote.Source = TN_ListTool Then
      oEnc.TextNotes.Remove i
      Set oTNote = Nothing
      Exit For
    End If
  Next i
'-Mod 40e    Look for content in the string
'- temp fix for medcin
  If Left$(RTFText, 5) = "\pard" Then
    RTFText = "{\rtf1\ansi\ansicpg1252\uc0\deff0 {\fonttbl{\f0\fswiss\fcharset0\fprq2 Arial;}}{\colortbl;\red0\green0\blue0;} \deftab720 \plain " & RTFText & "}"
  End If
  frmPrintPreview.TXTemp.RTFSelText = RTFText
  OutString = Trim$(frmPrintPreview.TXTemp.Text)
  '-really clean it up looking for any data
  OutString = Replace(OutString, Chr(10), "")
  OutString = Replace(OutString, Chr(13), "")
  OutString = Replace(OutString, Chr(9), "")
  If OutString = "" Then Exit Function
'~ end mod 40e

  Set oTNote = New TextNote
  
'-Mod 26e  (Deleted line)
  
  oTNote.DataDate = Now
  oTNote.Source = TN_ListTool
  oTNote.sText = RTFText
  oTNote.ClinicianNCID = mobjUser.GetAttribute("USER_Ncid")
  
  oEnc.TextNotes.Add oTNote
  Set oTNote = Nothing
  
      sMedcinText = ""
      OutString = ""
      msgHeader = "{\rtf1\ansi\ansicpg1252\uc0\deff0 {\fonttbl{\f0\fswiss\fcharset0\fprq2 Arial;}}{\colortbl;\red0\green0\blue0;} " & "\deftab720 \plain "
      msgFooter = "}"

    For Each oTNote In oEnc.TextNotes
      OutString = OutString & "{\par\plain\f0\fs16\b\tx" & frmEncParent.RTB(SECT_SO).Width \ 2 & oTNote.ProfileName & "\plain\f0\fs16 \tab  [" & oTNote.ClinicianName & ", " & mobjUtil.FormatDTG(oTNote.DataDate) & "] \par}"
      If oTNote.Source = TN_FORMTOOL Then
        Call mobjUtil.ConvertHTMLtoRTF(oTNote.sText, rnNumTextStrings, rasTextStrings)
        For i = 0 To rnNumTextStrings - 1
          OutString = OutString & "{\pard\plain\f0\fs16 " & rasTextStrings(i) & "\par}"
        Next i
      Else
        OutString = OutString & "{\pard\plain\f0\fs16 " & oTNote.sText & "\par}"
'-Mod 26e
       ' sMedcinText = sMedcinText & "{\pard\plain\f0\fs16 " & oTNote.sText & "\par}"
        sMedcinText = oTNote.sText
      End If
    Next oTNote
'-Mod 26e
    oEnc.SectionRTFs(CStr(SECT_MEDCIN)).Doc = msgHeader & sMedcinText & msgFooter
    
    
    oEnc.SectionRTFs(CStr(SECT_SO)).Doc = msgHeader & OutString & msgFooter
    frmEncParent.RTB(SECT_SO).Text = ""
    frmEncParent.RTB(SECT_SO).Height = 500
    frmEncParent.pics(SECT_SO).Height = 500
    frmEncParent.RTB(SECT_SO).RTFSelText = msgHeader & OutString & msgFooter
    DoEvents
    frmEncParent.pics(SECT_SO).Height = frmEncParent.RTB(SECT_SO).Height
    DoEvents

   Set oTNote = Nothing
  
End Function

Private Sub mobjCurrEnc_EncounterClosed()
'-Mod 34e
  RaiseEvent CloseEncounter(mobjCurrEnc)
End Sub

Private Sub mobjCurrEnc_EncStatusChanged(mvarStatus As EncStatus)
   Dim oEncChild As EncChild
'-Mod 33e
   RaiseEvent CurrEncStatusChanged(mvarStatus)
   
'--Tell a/p about the status change if it affects him
   If mvarStatus = Complete Or mvarStatus = Updated Or mvarStatus = NeedsCoSignature Then
      If Not mcolEncChild Is Nothing Then
        Set oEncChild = mcolEncChild(SECT_AP)
        If oEncChild.ModuleLoaded Then
          Call oEncChild.SetParentObject(Me)     '-- be sure a/p knows who I am (they need to get data from me)
          Call oEncChild.Module.CompletedEncounter  '-- call his function to update diagnoses to problems
          Set oEncChild = Nothing
        End If
      End If
   End If

End Sub

Public Function GetEandMCoding(mCol As Collection) As Boolean
Attribute GetEandMCoding.VB_Description = "mCol contains items of type EncEandMCoding.cls"
'-- fill the incoming collection with enc data
'- Assume mCol contains items of type EncEandMCoding
  Dim oEnM        As EncEandMCoding
  Dim oDas        As DasCtrl
  Dim oRTF        As EncRTF
  Dim oClinObs    As ClinicalObservation
  Dim oClinObs2   As ClinicalObservation
  Dim oClinEvent  As ClinicalEvent            '- Das control subpart
'-Mod 64e
  Dim oClinHdr    As ClinObsHeader
  Dim txt         As String
  Dim Posn        As Long
  Dim oCommentGroup As CommentGroup
  Dim oCommentItem As CommentItem
'~
  Dim colRTFs     As Collection
  Dim oAction     As EventActionInfo          '- Das section we're searching
  Dim sCmts       As String                   '- string cmts
  Dim asCmts()    As String
  Dim i           As Integer                  '- generic counter
  Dim j           As Integer                  '- generic counter 2
  Dim ByteArray() As Byte
  Dim sqz         As New EncZLib
  Dim lresult     As Long
'-Mod 51e
  Dim txtComment  As datastratllib.ITextComment
  Dim txtString   As datastratllib.IStringText
  Dim bFoundRTF   As Boolean

  If Not DebugMode Then On Error GoTo errHandler

  Set sqz = New EncZLib
  Set oDas = New DasCtrl
  Set colRTFs = New Collection
#If debugon Then
  Dim DR As DebugRec
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "GetEandMCoding", "Begin", DR, "Appt Cnt", mCol.count
#End If

  
  For Each oEnM In mCol
    If Not oEnM.CriteriaEncounterNumber = "" Then
      #If debugon Then
        If Not DL Is Nothing Then DL.DebugLog Exe, Module, "GetEandMCoding", "Loading:", DR, _
        "Unitnumber", oEnM.CriteriaPatientUnitNumber, "EncounterNumber", oEnM.CriteriaEncounterNumber, _
        "FacilityNCID", oEnM.CriteriaEncounterFacilityNCID
      #End If
  
      oDas.ClearAll
      oDas.GeneralCriteria.ClearAll
      oDas.GeneralCriteria.UnitNumber = oEnM.CriteriaPatientUnitNumber
      oDas.GeneralCriteria.EncounterNumber = oEnM.CriteriaEncounterNumber
      If Not oEnM.CriteriaEncounterFacilityNCID = "" Then
        oDas.GeneralCriteria.Facility.ValueConcept.UniqueId = oEnM.CriteriaEncounterFacilityNCID
      End If
      oDas.GeneralCriteria.QueryQualifier = "EVENT_TYPE_NCID = '" & NCID_ENCEVENT & "'"
      oDas.GeneralCriteria.ReturnMultiMedia = True
      
      oDas.DoSearch
          
      If Not oDas.EOF Then
      
        Set oClinEvent = oDas.Events.PatientData
    
        #If debugon Then
          If Not DL Is Nothing Then DL.DebugLog Exe, Module, "GetEandMCoding", "Mid", DR, "ObsCount", oClinEvent.ClinObs.count
        #End If
'-Mod 64e
        Set oClinHdr = oClinEvent.Header
        With oClinHdr
          If (Not .Comments.IsEmpty) Then
            For Each oCommentGroup In .Comments
              For Each oCommentItem In oCommentGroup.Comments
                txt = oCommentItem.Text
                Posn = InStr(txt, "=")
                
                If Posn > 1 And Len(txt) > Posn Then
                  If UCase$(Left$(txt, Posn - 1)) = "EANDMCODEREVIEW" Then
                    oEnM.CodeReview = Mid$(txt, Posn + 1)
                  End If
                End If
                
              Next oCommentItem
            Next oCommentGroup
          End If    '-- If (Not .Comments.IsEmpty) Then
        End With      '-- oClinHdr
'~end mod 64e
         
        For Each oClinObs In oClinEvent.ClinObs
        
          On Error GoTo ehNextObs
          
          Select Case oClinObs.ObsId.Value.CodedWOSform.ValueConcept
          Case NCID_EANDMCODE
            oEnM.EandMNCID = oClinObs.ObsValue.Value.CodedWOSform.ValueConcept
             If Not oClinObs.ObsMods.IsEmpty Then
               For Each oClinObs2 In oClinObs.ObsMods
                  Select Case oClinObs2.ObsId.Value.CodedWOSform.ValueConcept
                  Case NCID_REVIEWED_BY_CODER
                      oEnM.EandMReviewed = oClinObs2.ObsValue.Value.CodedWOSform.ValueConcept
                  End Select
          '-- 01.01.00 WAS: added
                  If (Not oClinObs2.ActionsInfo.IsEmpty) Then
                    For Each oAction In oClinObs2.ActionsInfo
                      With oAction
                        If .ActionId.ValueConcept.UniqueId = NCID_VERIFIED Then
                          oEnM.EandMReviewedBy = .Clinician.ValueConcept.UniqueId  '-- added 1.1.00 WAS:
                          oEnM.EandMReviewedDate = .EffectiveTime
                          
                         End If
                      End With
                    Next oAction
                  End If           '--If (Not oClinObs2.ActionsInfo.IsEmpty) Then
               Next oClinObs2
               Set oClinObs2 = Nothing
             End If            '--If Not oClinObs.ObsMods.IsEmpty Then
            
            Set oClinObs2 = Nothing
            Set oAction = Nothing
          
          Case NCID_ENCOUNTERRTF
     '     debug.print "inside rtf/note"
    '-- this may be an rtf or a note.  figure it out via the text comments in the eai
                  '-- concatenate all the text stuff into a joined string so we  can see what we have
    
            bFoundRTF = False
            If (Not oClinObs.ActionsInfo.IsEmpty) Then
              For Each oAction In oClinObs.ActionsInfo
                With oAction
                  If (Not .TextComments.IsEmpty) Then
    '-Mod 51e drop the for i = 1 to count stuff
                    For Each txtComment In .TextComments          'i = 1 To .TextComments.count
                      For Each txtString In txtComment.Comments           'j = 1 To .TextComments(i).Comments.count
                        If txtString.Text = "#RTF#" Then  'we have an rtf obs
                          bFoundRTF = True
                          Exit For
                        End If
                      Next
                      If bFoundRTF Then Exit For
                    Next
                  End If  '-(Not .TextComments.IsEmpty) Then
                End With
                  
                If bFoundRTF Then Exit For
              Next oAction
            End If
            If bFoundRTF Then
              
              Set oRTF = New EncRTF
              
              If (Not oClinObs.ActionsInfo.IsEmpty) Then
                'Set oActions = oClinObs.ActionsInfo
                For Each oAction In oClinObs.ActionsInfo
                  With oAction
                      sCmts = ""
                      If (Not .TextComments.IsEmpty) Then
    '-Mod 51e drop the for i = 1 to count stuff
                        For Each txtComment In .TextComments          'i = 1 To .TextComments.count
                          For Each txtString In txtComment.Comments           'j = 1 To .TextComments(i).Comments.count
                            sCmts = sCmts & "|" & txtString.Text
                          Next
                        Next
                      End If
                      sCmts = Mid$(sCmts, 2)  '--drop leading pipe
                      asCmts = Split(sCmts, "|")
                      '-- see which section we're in
                      If InStr(sCmts, "STATUS") > 0 Then
                      
    '-Mod 51e Need the original size and other stuff, too
                        For i = 0 To UBound(asCmts)
                          j = InStr(asCmts(i), "=")
                          If j > 1 And Len(asCmts(i)) > j Then
                            Select Case UCase$(Left$(asCmts(i), j - 1))
                            Case "STATUS":         oRTF.Status = Mid$(asCmts(i), j + 1)
                            Case "TITLE":          oRTF.Title = Mid$(asCmts(i), j + 1)
                            Case "SENSITIVITY":    oRTF.SensitivityLevel = Mid$(asCmts(i), j + 1)
                            Case "ORIGINALSIZE":   oRTF.OriginalSize = Mid$(asCmts(i), j + 1)
                            End Select
                          End If   '-- j > 1
                        Next i
    
                        If (.EffectiveTime <> NULL_DATE) Then
                          oRTF.DTS = .EffectiveTime
                        End If
                      End If
                  End With     'oAction
                Next           '--   oAction In oClinObs.ActionsInfo
              End If           '--If (Not oClinObs.ActionsInfo.IsEmpty) Then
            
              
              ByteArray = oClinObs.ObsValue.Value.MultiMedia.Data
              If oRTF.OriginalSize = 0 Then
                oRTF.OriginalSize = Val(oClinObs.ObsValue.Value.MultiMedia.UncompressedLength)
              End If
              
              If oRTF.OriginalSize > 0 Then
                '-- expand it
                  lresult = sqz.DecompressData(ByteArray, oRTF.OriginalSize)
              End If
             
              If ByteArray(0) = Asc("{") Then
                oRTF.Doc = StrConv(ByteArray, vbUnicode)
              Else
                If ByteArray(0) = 120 And ByteArray(1) = 156 Then '-- one last try
                  lresult = sqz.DecompressData(ByteArray, 599998)
                  If ByteArray(0) = Asc("{") Then
                  
    '-Mod 51e "Use the function, Luke!"
                    oRTF.Doc = StrConv(ByteArray, vbUnicode)
                  End If
                End If
              End If
             If Not oRTF Is Nothing Then
               colRTFs.Add oRTF
             End If
    
           End If       '-bfoundit
         End Select     '-oClinObs.ObsId.Value.CodedWOSform.ValueConcept
          
ehNextObs:
        Next oClinObs
        
        ' sort the enc rtfs and find the most recent
        Dim MostRecentDate As Date
        If Not DebugMode Then On Error GoTo errHandler
        MostRecentDate = 0
        For Each oRTF In colRTFs
          If oRTF.DTS > MostRecentDate Then
              MostRecentDate = oRTF.DTS
              Set oEnM.CurrentRTF = oRTF
          End If
        Next oRTF
        
      End If     '-If Not oDas.EOF Then
    End If       '-If Not oEnM.CriteriaEncounterNumber = "" Then
  Next oEnM
 
  Set oRTF = Nothing
  Set oDas = Nothing
  Set oEnM = Nothing
  Set oDas = Nothing
  Set oClinObs = Nothing
  Set oClinObs2 = Nothing
  Set oClinEvent = Nothing
  Set colRTFs = Nothing
  Set oAction = Nothing
  Set sqz = Nothing
  Set txtComment = Nothing
  Set txtString = Nothing
  
#If debugon Then
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "GetEandMCoding", "End", DR
#End If

  GetEandMCoding = True
  Exit Function
  
errHandler:
  Set oRTF = Nothing
  Set oDas = Nothing
  Set oEnM = Nothing
  Set oDas = Nothing
  Set oClinObs = Nothing
  Set oClinObs2 = Nothing
  Set oClinEvent = Nothing
  Set colRTFs = Nothing
  Set oAction = Nothing
  Set sqz = Nothing
  Set txtComment = Nothing
  Set txtString = Nothing
    
  GetEandMCoding = False
    
  Call gobjShared.ShowVBError(Err.Number, Err.Description, "Enctr.GetEandMCoding", "EncounterParent", vbExclamation)
  Exit Function

End Function

Public Function TestOne()
'Dim oEnM As EncEandMCoding
'Set oEnM = New EncEandMCoding
'oEnM.CriteriaPatientUnitNumber = "121795"
'oEnM.CriteriaEncounterNumber = "1659"
''oEnM.CriteriaEncounterFacilityNCID = mobjCurrEnc.FacilityNCID
'
'Dim mcol As Collection
'Set mcol = New Collection
'mcol.Add oEnM
'Call GetEandMCoding(mcol)

End Function
'-Mod 63e
Private Function AllowedInCurrentEncounter(Status As EncStatus) As Boolean
  Select Case Status
  Case NeedsCoSignature, Updated, Complete, PatientLeftWithoutBeingSeen
    AllowedInCurrentEncounter = False
  Case Else
    AllowedInCurrentEncounter = True
  End Select
End Function

'-Mod 61e Replaced this function from archived files
'         Check the current actionInfo on the enc.
Private Function HandleChangedEncounter(oEnc As Encounter) As Boolean

Dim AppendChanges     As Boolean
Dim AppendChangesText As String
Dim i                 As Integer
Dim Row               As Long
Dim bDataChanged      As Boolean
Dim bNoteChanged      As Boolean
Dim rows              As Long
Dim CurrStatus        As EncStatus
Dim oConcept          As ConceptCtrl
Dim colRTFs           As Collection
Dim oRTF              As EncRTF
Dim colNotes          As Collection
Dim sCurrClin         As String

#If debugon Then
  Dim DR As DebugRec
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "HandleChangedEncounter", "Begin", DR
#End If

  On Error GoTo errHandler
  
  If Not oEnc.StoredVersionChanged(True) Then
    HandleChangedEncounter = True
    Exit Function
  End If

'-- things have changed go get the current data
  Set oConcept = New ConceptCtrl
  oConcept.UniqueId = oEnc.UpdateClinicianNCID
  sCurrClin = oConcept.PrefRep("2000")
  
'==================================
  rows = UBound(Sect) - 1  '- for medcin
  
  '-Mod 61e  check the current actionInfo on the enc.
  'Set oConcept = New ConceptCtrl
  
  'dummy-up a set of rtf sections (with sec numbers)
'  Set colRTFs = New Collection
'  For i = 1 To rows
'    Set oRTF = New EncRTF
'    oRTF.SectionIndex = i
'    colRTFs.Add oRTF, CStr(i)
'  Next i
'  Set colNotes = New Collection

'-- fill these colls with   data from the db
'- just get another enc object
Dim oEnc2 As Encounter
  Set oEnc2 = New Encounter
  Call oEnc2.Init(mobjUser)
  oEnc2.DataID = oEnc.DataID
  Set oEnc2.Patient = oEnc.Patient
  Call oEnc2.LoadEncounter

'======================== status has changed to complete or updated =========================

  If (oEnc2.Status = Complete And oEnc.Status <> Complete And oEnc.Status <> Updating) Then
  '-- only let what they did be added as an amendment
   Select Case MsgBox("The encounter document has been completed by the primary provider.  Append your changes as a note? (Selecting 'No' will discard you changes.)", vbQuestion + vbYesNo, "Encounter Status has changed")
    Case vbYes
      AppendChanges = True
    Case vbNo
      oEnc.Status = NoSave    'temp status to flag the frmparent.unload process no to save this enc
      HandleChangedEncounter = True
      Exit Function
   End Select
  End If

'===================================================

'-- we're only interested in notes not belonging to user
  For i = mobjCurrEnc.OriginalNotes.count To 1 Step -1
    If oEnc.OriginalNotes.Item(i).UserNcid = mobjUser.GetAttribute("USER_Ncid") Then
      oEnc.OriginalNotes.Remove (i)
    End If
  Next i
  For i = oEnc2.EncNotes.count To 1 Step -1
    If oEnc2.EncNotes.Item(i).UserNcid = mobjUser.GetAttribute("USER_Ncid") Then
      oEnc2.EncNotes.Remove (i)
    End If
  Next i

'-- now see if there's been any changes  (orig to now)
  If oEnc.OriginalNotes.count = oEnc2.EncNotes.count And oEnc2.EncNotes.count > 0 Then
    For i = 1 To oEnc.OriginalNotes.count
      If oEnc2.EncNotes.Item(i).Note <> oEnc.OriginalNotes.Item(i).Note Then
         bNoteChanged = True
         Exit For
      End If
    Next i
  End If

  
'============================= notes have changed ==========================================
  If bNoteChanged Then
    MsgBox "The collection of notes has been changed while you were working in this encounter document.  Review the 'Notes' section before continuing.", vbExclamation, "Data Integrity Error"
   '-- reset the notes collection to the current set plus his

    '-- get rid of any that are NOT his
    For i = oEnc.EncNotes.count To 1 Step -1
      If oEnc.EncNotes.Item(i).UserNcid <> mobjUser.GetAttribute("USER_Ncid") Then
        oEnc.EncNotes.Remove (i)
      End If
    Next i
    '-- add the current notes (that are not his)
     For i = 1 To oEnc2.EncNotes.count
      oEnc2.EncNotes.Item(i).NoteID = oEnc.EncNotes.count + 1
      oEnc.EncNotes.Add oEnc2.EncNotes.Item(i), CStr(oEnc2.EncNotes.Item(i).NoteID)
     Next i
     '-- make the 'original' read into the current on
     Set oEnc.OriginalNotes = New Collection
     
Dim objNote As EncNote
     For Each objNote In oEnc.EncNotes
       oEnc.OriginalNotes.Add objNote
     Next objNote
     
  '-- refresh the encounter note section and display it.
    Call DisplayNotes

    HandleChangedEncounter = False
    UnloadFailed = True
    Exit Function

  End If    '--If bNoteChanged Then
'=======================================================================

 
  frmDataSelect.FGrid.Row = 0
  frmDataSelect.FGrid.Col = 1
  frmDataSelect.FGrid.Text = "Your Data"
  frmDataSelect.FGrid.Col = 2
  frmDataSelect.FGrid.Text = "Current Database Data"

  frmDataSelect.FGrid.rows = 8
  bDataChanged = False
'--push the data into the display grid

  If oEnc.SectionRTFs.count < 7 Then
    HandleChangedEncounter = False
    Exit Function
  End If
  Row = 0
  For i = 1 To 7
    frmDataSelect.TxTemp1.Text = ""
    frmDataSelect.TxTemp2.Text = ""
    frmDataSelect.TxTemp1.RTFSelText = sOriginals(i).Doc   'oEnc.SectionRTFs(CStr(i)).Doc
    frmDataSelect.TxTemp2.RTFSelText = oEnc2.SectionRTFs(CStr(i)).Doc
'--12.29.99.WAS - added not sect so
    If frmDataSelect.TxTemp1.Text <> frmDataSelect.TxTemp2.Text _
        And frmDataSelect.TxTemp2.Text <> "" And i <> SECT_SO Then
      bDataChanged = True
      AppendChangesText = AppendChangesText & "\par " & frmDataSelect.TxTemp1.Text
    End If
    Row = Row + 1
    frmDataSelect.TxTemp1.Text = ""
    frmDataSelect.TxTemp1.RTFSelText = oEnc.SectionRTFs(CStr(i)).Doc
    frmDataSelect.FGrid.TextMatrix(Row, 0) = Sect(i)
    frmDataSelect.FGrid.TextMatrix(Row, 1) = frmDataSelect.TxTemp1.Text
    frmDataSelect.FGrid.TextMatrix(Row, 2) = frmDataSelect.TxTemp2.Text
  Next i

 Dim sSelected As String
 
'============================== only append as a note
  
  If AppendChanges Then   '-user only wants to append their changes as a note
     If AppendChangesText <> "" Then
        AppendChangesText = "{\rtf1\ansi\ansicpg1252\uc0\deff0" _
      & "{\fonttbl{\f0\fswiss\fcharset0\fprq2 Arial;}}" _
      & "{\colortbl;\red0\green0\blue0;}" _
      & "\deftab720 \plain " & AppendChangesText & "}"
      Call AddEncounterNote("Changed/Added after signature", AppendChangesText, False)
     End If
     HandleChangedEncounter = True
     Exit Function
  End If
  
'====================================
' If bDataChanged Then
    frmDataSelect.FGrid.ColWidth(1) = (frmDataSelect.FGrid.Width - frmDataSelect.FGrid.ColWidth(0) - 150) / 2
    frmDataSelect.FGrid.WordWrap = True
    frmDataSelect.FGrid.AutoSize 0, 2

    frmDataSelect.ParentObject = Me
    frmDataSelect.Label1.Caption = sCurrClin & " modified this encounter on " & mobjUtil.FormatDTG(oEnc.UpdateDate) & "." _
       & "Please select the data for this document by clicking on one column in each section below."
Screen.MousePointer = vbDefault
    mobjComm.Message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_ENCOUNTER_APPLICATION
    frmDataSelect.Show vbModal
    mobjComm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_ENCOUNTER_APPLICATION

'- user hit 'cancel'
  If frmDataSelect.Canceled Then
    HandleChangedEncounter = False
    UnloadFailed = True
    Exit Function
  End If

 '--now handle the user's selections
  sSelected = frmDataSelect.Selected
  If InStr(sSelected, "2") Then
    For i = 1 To 7
      If Mid$(sSelected, i, 1) = "2" Then
        oEnc.SectionRTFs(CStr(i)).Doc = oEnc2.SectionRTFs(CStr(i)).Doc
      End If
    Next i
  End If
'  End If    ' bdatachanged

  Unload frmDataSelect
  
  HandleChangedEncounter = True
  Set oEnc2 = Nothing
  Set oRTF = Nothing
  Exit Function
  
errHandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "Enctr.HandleChangedEncounter ", "Enc", vbCritical)
    HandleChangedEncounter = True
    Set oEnc2 = Nothing
    Set oRTF = Nothing
    Exit Function
  
End Function

Friend Function ReadSectionRTFs(mcolRTFs As Collection, _
                                mcolNotes As Collection, _
                                oEnc As Encounter, _
                                CurrStatus As EncStatus) As Boolean
                                
End Function

Public Function UserHasSignPrivilege(sForeignLogin) As Boolean
    Dim sMsg As String
    sMsg = "HasPrivilege|" & sForeignLogin & "|" & Priv_Current_Encounter & "|" & Priv_Sign
    Call mobjComm.Message(cwiGENERIC, sMsg, NCID_CW_APPLICATION, NCID_ENCOUNTER_APPLICATION)
    If sMsg = "TRUE" Then
        UserHasSignPrivilege = True
    Else
        If sMsg = "FALSE" Then
          MsgBox "The selected user does not have 'Co-Sign' privileges"
        Else
          MsgBox "Unknown error checking co-signing privileges for provider."
        End If
    End If
End Function

