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
Attribute VB_Name = "OLEServers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"OLEServer"
Attribute VB_Ext_KEY = "Member0" ,"OLEServer"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Const MODULE_NAME As String = "OLEServers"
Private Const APPSECTION As String = "Additional Applications"
Private Const BUFFER_LENGTH As Long = 254

'osh data positions
Private Const OSH_MENU_TEXT         As Integer = 0
Private Const OSH_BUTTON_TEXT       As Integer = 1
Private Const OSH_BUTTON_BITMAP     As Integer = 2
Private Const OSH_PATIENT_REQIRED   As Integer = 3
Private Const OSH_TOOL_TIP          As Integer = 4
Private Const OSH_REG_NAME          As Integer = 5
Private Const OSH_APP_NCID          As Integer = 6
Private Const OSH_PATIENT_SUMMARY   As Integer = 7
Private Const OSH_CMD_LINE          As Integer = 8
Private Const OSH_APP_TYPE          As Integer = 9
Private Const OSH_PRINT             As Integer = 10
Private Const OSH_PRINT_PREVIEW     As Integer = 11
Private Const OSH_PROPERTY_VIEW     As Integer = 12
Private Const OSH_SELECTED          As Integer = 13
Private Const OSH_PRIVILEGE         As Integer = 14

Private mServers As Collection                 'All OLEServer object that are loaded from OSH
Private mRunningServers As Collection          'loaded oleservers, in order of use, all servers with their forms 'up'.

Private Sub Class_Initialize()
  Set mServers = New Collection
  Set mRunningServers = New Collection
End Sub

Private Sub Class_Terminate()
  On Error Resume Next
  Set mServers = Nothing
  Set mRunningServers = Nothing
End Sub


Public Property Get Count() As Long
  On Error Resume Next
  Count = mServers.Count
End Property

'This subroutine extracts information about the passed in OLE Server (EnterpriseOLEServerInfo)
'and stores it in the passed in OLEServer object.
'Parameters:
' OLESvr                    OLEServer object to initialize
' EnterpriseOLEServerInfo   the ";" delimited string from which to parse the server info (created by the osh untility)
'
'  Return Value:
'    Returns TRUE if the OLE Server information was extracted successfully.
'    Returns FALSE if an error occurred in extracting the OLE Server information.
Private Function ExtractOLEServerInfo(ByRef OLESvr As OLEServer, ByVal EnterpriseOLEServerInfo As String) As Boolean
Dim OLESvrConfig() As String
Dim Valid As Boolean

  On Error GoTo ErrHandler
    
  OLESvrConfig = Split(EnterpriseOLEServerInfo, ";")
  
  If UBound(OLESvrConfig) < 13 Then
    Screen.MousePointer = vbDefault
    Beep
    MsgBox "Unable to retrieve the enterprise information for OLE Server: " & OLESvr.sApplicationName & " in ExtractOLEServerInfo.  The CW configuration database may have become corrupted, or is incompatible.", vbCritical, App.Title

    ExtractOLEServerInfo = False
    Exit Function
  End If

  'See If Command line switch is set that is specified in the OSH
'  If Len(OLESvrConfig(OSH_CMD_LINE)) > 0 Then
'    Valid = (gshared.CmdLineSwitches(UCase$(OLESvrConfig(OSH_CMD_LINE))) = UCase$(OLESvrConfig(OSH_CMD_LINE)))
'  Else
    Valid = True
'  End If

 ' If (Left$(UCase$(OLESvrConfig(OSH_SELECTED)), 1) = "Y" And Valid) Then
    With OLESvr
      .sCommandLine = OLESvrConfig(OSH_CMD_LINE)
      .bSelected = (UCase$(Left$(OLESvrConfig(OSH_SELECTED), 1)) = "Y")
      .sMenuString = OLESvrConfig(OSH_MENU_TEXT)
      .sIconString = OLESvrConfig(OSH_BUTTON_TEXT)
      .sButtonBitmap = OLESvrConfig(OSH_BUTTON_BITMAP)
      .sToolTip = OLESvrConfig(OSH_TOOL_TIP)
      .sExecutable = OLESvrConfig(OSH_REG_NAME)
      .sAppNCID = OLESvrConfig(OSH_APP_NCID)
      .bSupportsPatSummary = (Left$(UCase$(OLESvrConfig(OSH_PATIENT_SUMMARY)), 1) = "Y")
      .sAppType = UCase$(OLESvrConfig(OSH_APP_TYPE))
      .bPrint = (Left$(UCase$(OLESvrConfig(OSH_PRINT)), 1) = "Y")
      .bPrintPreview = (Left$(UCase$(OLESvrConfig(OSH_PRINT_PREVIEW)), 1) = "Y")
      .bDisablePropertyView = Not (Left$(UCase$(OLESvrConfig(OSH_PROPERTY_VIEW)), 1) = "Y")
                                        
      'Set Patient required value
      Select Case UCase$(OLESvrConfig(OSH_PATIENT_REQIRED))
        Case "YES", "Y"
          .iPatientRequired = PATIENT_REQUIRED_YES
          .bEnabled = False
        
        Case "NO", "N"
          .iPatientRequired = PATIENT_REQUIRED_NO
          .bEnabled = True
        
        Case "OPTIONAL", "O"
          .iPatientRequired = PATIENT_REQUIRED_NO
          .bEnabled = True
        
        Case Else    '--default to not patient required--
          .iPatientRequired = PATIENT_REQUIRED_NO
          .bEnabled = True
      End Select
      
      'Set Priv
      If UBound(OLESvrConfig) >= 14 Then
         .sAccessPriv = OLESvrConfig(OSH_PRIVILEGE)
      End If
      If Len(.sAccessPriv) = 0 Then
         .sAccessPriv = SetAccessPrivileges(.sAppNCID)
      End If
    
      ExtractOLEServerInfo = True
    End With
  'Else
  '  ExtractOLEServerInfo = False
  'End If

  Exit Function

ErrHandler:
  gshared.ShowVBError Err.Number, Err.Description, "ExtractOLEServerInfo", App.Title, vbCritical

End Function

Public Property Get Item(ByVal IndexKey As Variant) As OLEServer
Attribute Item.VB_UserMemId = 0
  On Error Resume Next
  Set Item = mServers(IndexKey)
End Property
Public Function AddOLEServer(oSvr As OLEServer)
    mServers.Add oSvr, CStr(oSvr.sAppNCID)
End Function
'This function loads information that indicates which OLE Servers
'the logged on user will have access to from CW.
'Return Value:
' Returns TRUE if the OLE Server information was loaded successfully.
' Returns FALSE if an error occurred in loading the OLE Server information.
Public Function LoadOLEServerInfo() As Boolean
Dim NextProgramID As Long
Dim OLESvr As OLEServer
Dim SqlStmt As String
Dim SQLOps As ICHCSII_SQL.ISqlOps
    
  On Error GoTo ErrHandler
    
  #If DEBUGON Then
    If Not DL Is Nothing Then DL.DebugLog Exe, MODULE_NAME, "LoadOLEServerInfo", "Begin", mDR
  #End If

  'initialize variable
  NextProgramID = PROG_ID_SEED + 1

  'Load Patient Servers if logged user is a patient


  NextProgramID = NextProgramID + 1
  
  '  Read in all of the information about each OLE Server application the user specified
  '    to use in CW.
  
  Set SQLOps = goConn.CHCSII_SQLOPS
  
  SqlStmt = "Select * from OSH_INI_SETTINGS where APPLICATION_NCID = '" & NCID_CW_APPLICATION & "' and APSECTION = 'cw applications'"
  SQLOps.Execute SqlStmt
   
  Do While Not SQLOps.EOF
    Set OLESvr = New OLEServer
    OLESvr.bInitial = True
    OLESvr.sApplicationName = SQLOps("apkey")

    If ExtractOLEServerInfo(OLESvr, SQLOps.value("apvalue")) Then
      NextProgramID = NextProgramID + 1
      
      With OLESvr
        '.lProgId = NextProgramID
        .bUserHasAccess = True
'        Select Case .sAppNCID
'          Case NCID_IMMUNIZATION_SYNCH_APPLICATION
'            .bUserHasAccess = (gshared.CmdLineSwitches("TRAINING") = vbNullString)
'
'          Case NCID_IMMUNIZATION_SYNCH_APPLICATION
'            .bUserHasAccess = (gshared.CmdLineSwitches("CITA") = vbNullString)
'
'          Case Else
'            .bUserHasAccess = UserHasAppAccess(.sAppNCID, OLESvr)
'        End Select
        
        'scr 20971 Pharmacy spt
        If gshared.pbMTFSupportsBDM Then
          If .sAppNCID = NCID_OE_MED_APPLICATION Then
            .sApplicationName = "CHCSII_RxPOE.mclsRxPOE"
          End If
        End If
        
        .bRegularViewRunning = False
        .sLaunchSequence = vbNullString
        .bNew = False
        
        
        #If DEBUGON Then
          If Not DL Is Nothing Then DL.DebugLog Exe, MODULE_NAME, "LoadOLEServerInfo", "", mDR, "ServerName", .sApplicationName, "sAccessPriv", .sAccessPriv, "bUserHasAccess", .bUserHasAccess
        #End If
      End With
      OLESvr.bInitial = False
      mServers.Add OLESvr, CStr(OLESvr.sAppNCID)
    End If
    
    Set OLESvr = Nothing
    SQLOps.MoveNext
  Loop

  
  LoadOLEServerInfo = True

  #If DEBUGON Then
    If Not DL Is Nothing Then DL.DebugLog Exe, MODULE_NAME, "LoadOLEServerInfo", "End", mDR
  #End If

  Exit Function

ErrHandler:
    gshared.ShowVBError Err.Number, Err.Description, "OLEServers.LoadOLEServerInfo", App.Title, vbCritical

End Function

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mServers.[_NewEnum]
End Property

Public Sub Remove(ByVal IndexKey As Variant)
  On Error Resume Next
  mServers.Remove IndexKey
End Sub

'If the Priv is NULL in the OSH then check here to see if there is a priv for the App Server.
Private Function SetAccessPrivileges(ByVal AppNCID As String) As String
  Select Case AppNCID
      Case NCID_DDSA_APPLICATION, NCID_DENTAL_BIGTOOTH_APPLICATION, _
           NCID_DENTAL_DATAVIEWS_APPLICATION, NCID_DENTAL_FORENSIC_APPLICATION, _
           NCID_DENTAL_FULLMOUTH_APPLICATION, NCID_DENTAL_LABS_APPLICATION, _
           NCID_DENTAL_NOTES_APPLICATION, NCID_DENTAL_NOTESPLUS_APPLICATION
        SetAccessPrivileges = Priv_Dental_Access
        
      Case NCID_PATIENT_SEARCH_APPLICATION
        SetAccessPrivileges = Priv_Patient_Search
        
      Case NCID_CLINIC_SCHEDULE_APPLICATION
        SetAccessPrivileges = Priv_Appointments
        
      Case NCID_TELCON_APPLICATION
        SetAccessPrivileges = Priv_Telephone_Consult_Appointment
        
      Case NCID_CONSULT_TRACKING_APPLICATION
        SetAccessPrivileges = Priv_Orders_Consults
        
      Case NCID_DEMOGRAPHICS_APPLICATION
        SetAccessPrivileges = Priv_Registration_Patient_Demographics
        
      Case NCID_PATIENT_SUMMARY_APPLICATION
        SetAccessPrivileges = Priv_Health_History
        
      Case NCID_PROBLEMS_APPLICATION
        SetAccessPrivileges = Priv_Medical_Problem_List
        
      Case NCID_MEDICATIONS_APPLICATION
        SetAccessPrivileges = Priv_Medication_List
        
      Case NCID_ALLERGIES_APPLICATION
        SetAccessPrivileges = Priv_Allergies
        
      Case NCID_HEALTH_AND_WELLNESS
        SetAccessPrivileges = Priv_Wellness
        
      Case NCID_IMMUNIZATION_SYNCH_APPLICATION
        SetAccessPrivileges = Priv_Imm_Multiple_Entry
        
      Case NCID_VITAL_SIGNS_APPLICATION
        SetAccessPrivileges = Priv_Vital_Signs
        
      Case NCID_PREVIOUS_ENCOUNTERS_APPLICATION, NCID_CLINICAL_NOTES_APPLICATION
        SetAccessPrivileges = Priv_Previous_Encounter
        
      Case NCID_HEAR_APPLICATION
        SetAccessPrivileges = Priv_Coupler_Provider & "|" & Priv_Coupler_Patient & "|" & Priv_Coupler_Clerk
        
      Case NCID_READINESS_APPLICATION
        SetAccessPrivileges = Priv_Readiness_Information
        
      Case NCID_LAB_APPLICATION
        SetAccessPrivileges = Priv_Laboratory
        
      Case NCID_RADIOLOGY_RESULTS_APPLICATION
        SetAccessPrivileges = Priv_Radiology
        
      Case NCID_FLOWSHEETS_APPLICATION
        SetAccessPrivileges = Priv_Flowsheets
        
      Case NCID_ENCOUNTER_APPLICATION
        SetAccessPrivileges = Priv_Current_Encounter
        
      Case NCID_MEDCIN_APPLICATION
        SetAccessPrivileges = Priv_S_O
        
      Case NCID_REASON_FOR_VISIT
        SetAccessPrivileges = Priv_Reason_For_Visit
        
      Case NCID_ASSESMENT_AND_PLAN_APPLICATION
        SetAccessPrivileges = Priv_Encounter_Coding
        
      Case NCID_ANP_DIAGNOSIS_APPLICATION
        SetAccessPrivileges = Priv_Diagnosis
        
      Case NCID_ANP_PROCEDURES_APPLICATION
        SetAccessPrivileges = Priv_Procedure
        
      Case NCID_OE_LAB_APPLICATION
        SetAccessPrivileges = Priv_Orders_Laboratory_Orders
        
      Case NCID_OE_RAD_APPLICATION
        SetAccessPrivileges = Priv_Orders_Radiology_Orders
        
      Case NCID_OE_MED_APPLICATION
        SetAccessPrivileges = Priv_Orders_Pharmacy_Orders
        
      Case NCID_CONSULT_TRACKING_APPLICATION
        SetAccessPrivileges = Priv_Orders_Consults
        
      Case NCID_DISPOSITION_APPLICATION
        SetAccessPrivileges = Priv_Patient_Disposition
        
      Case NCID_ALERT_APPLICATION
        SetAccessPrivileges = Priv_Alerts
        
      Case NCID_PATIENT_LIST_APPLICATION
        SetAccessPrivileges = Priv_Patient_List
        
      Case NCID_ASM_APPLICATION
        SetAccessPrivileges = Priv_Alerts_Management
        
      Case NCID_COSIGNATURE_APPLICATION
        SetAccessPrivileges = Priv_Encounter_Coding
        
      Case NCID_SIGN_ORDERS_APPLICATION
        SetAccessPrivileges = Priv_Sign_Orders
      
      Case NCID_WEBBROWSER_APPLICATION
        SetAccessPrivileges = Priv_Web_Browser
        
      Case NCID_GENERAL_REPORTS_APPLICATION
        SetAccessPrivileges = Priv_Reports_Standard & "|" & Priv_Reports_Preventive & "|" & Priv_Reports_Customized
        
      Case NCID_TEMPLATE_MANAGEMENT
        SetAccessPrivileges = Priv_Templates_Personal_Template
        
      Case NCID_LIST_MGMT_APPLICATION
        SetAccessPrivileges = Priv_List_Management
        
      Case NCID_NEW_RESULTS_APPLICATION
        SetAccessPrivileges = Priv_Results_Retrieval_Non_Sensitive
        
      Case NCID_CODING_REVIEW
        SetAccessPrivileges = Priv_Encounter_Coding
        
      Case NCID_PATIENTINSTRUCTION_APPLICATION
        SetAccessPrivileges = Priv_Other_Therapies
        
      Case NCID_ORDER_SETS_APPLICATION
        SetAccessPrivileges = Priv_Order_Sets_Personal_Order_Sets
        
      Case NCID_SRTS_OE_APPLICATION
        SetAccessPrivileges = Priv_SRTS_OE
        
      Case NCID_SRTS_OM_APPLICATION
        SetAccessPrivileges = Priv_SRTS_OM
        
      Case NCID_THEATER_ADMIN_APPLICATION
        'SetAccessPrivileges = Priv_User_Administration
        
      Case NCID_DATA_MANAGER
        ' SetAccessPrivileges = Priv_system_Administration
        
      Case NCID_CITA_APPLICATION
        SetAccessPrivileges = Priv_Imm_GiveVacc
        
      Case NCID_CITA_ADMIN_APPLICATION
        SetAccessPrivileges = Priv_Imm_VaccAdmin
  End Select
End Function

'Determines if user has priv to access app module.
Public Function UserHasAppAccess(ByRef AppNCID As String, Optional ByRef OLESvr As OLEServer) As Boolean
    
  On Error GoTo ErrHandler
  
  If Len(Trim$(AppNCID)) = 0 And OLESvr Is Nothing Then
    Exit Function
  End If

  'get the server if not passed in
  If OLESvr Is Nothing Then
    If Not gServers.FindOLEServer(sbNCID, AppNCID, OLESvr) Then
      Exit Function
    End If
    
    AppNCID = OLESvr.sAppNCID
  End If
    
  If Len(Trim$(OLESvr.sAccessPriv)) > 0 Then
    UserHasAppAccess = CheckPriv(OLESvr.sAccessPriv)
  Else
    UserHasAppAccess = CheckPrivSpecialCases(OLESvr.sAppNCID)
  End If

  Exit Function

ErrHandler:
  gshared.ShowVBError Err.Number, Err.Description, "Global.UserHasAppAcess", App.Title, vbCritical
  
End Function

Public Function FindOLEServer(ByVal SearchBy As enumSearchBy, ByVal Criteria As Variant, ByRef OLESvr As OLEServer) As Boolean
    
  On Error Resume Next
  
  Set OLESvr = Nothing
  
  Select Case SearchBy
    Case sbNCID
      Set OLESvr = mServers.Item(CStr(Criteria))
      FindOLEServer = Not (OLESvr Is Nothing)
    
    Case sbProgID
      For Each OLESvr In mServers
        If Criteria = OLESvr.lProgId Then
          FindOLEServer = True
          Exit Function
        End If
      Next
    
    Case sbServerName
      For Each OLESvr In mServers
        If UCase$(Criteria) = UCase$(OLESvr.sApplicationName) Then
          FindOLEServer = True
          Exit Function
        End If
      Next
  End Select
    
End Function


