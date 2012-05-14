VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'-------------------------------------------------------------------------------------------------------
'this class is the entry point class for the CAC_Options DLL
'M.Broaddus
'2/6/2002
'-------------------------------------------------------------------------------------------------------

Implements ICHCSII_CoreAPI2
Implements ICHCSII_CoreAPI3

'Notes
'1. What menu options are supposed to be up?
'2. Is this supposed to be a modal form?
'3. this component will handle all of the decisions that need to be made concerning CAC, so
'   Core will need to launch this upon startup of the application.

'Operations
'1. Check who out of the 4 servers that we care about is showing forms.
'2. if the modules are just open, then we need to make a decision about who to query.

Private cwlMyProgId As Long 'this is the progId assigned to me by Core
Private mlTopMostServerID As Long

Public Property Let TopMostServerID(vData As Long)
    mlTopMostServerID = vData
End Property

Public Function StartCACProcessing(XMLString As String) As Boolean
    Dim retVal As String
    Static bProcessing As Boolean
On Error GoTo ErrorHandler
    
    If bProcessing Then Exit Function
    
    bProcessing = True
    
    gobb.CData.Deserialize XMLString
    
    'BR: If we do not atleast have a SSN and/or lastname for the card data, then this transaction is bogus.
    If gobb.CData.SSN = "" And gobb.CData.Surname = "" Then
        bProcessing = False
        Exit Function
    Else
      If gobb.CData.SponsorSSN = "" Then
        gobb.CData.SponsorSSN = gobb.CData.SSN
      End If
    End If
    
    If mlTopMostServerID > 0 Then
        retVal = CStr(mlTopMostServerID)
    Else
        Call gobb.Comm.Message(cwiFOREGROUND_APPLICATION, retVal, NCID_CW_APPLICATION, 0)
    End If
    
    Select Case retVal
        Case "" 'open the cac options window
            'BR: This is where I perform the check to see if the patient exisits
            If gobb.Patient.SearchPatientCAC(gobb.CData) Then
                gobb.Comm.Message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_CAC_OPTIONS_APPLICATION
                frmCAC_Options.Show vbModal, gobb.Comm.MainWindow
                gobb.Comm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_CAC_OPTIONS_APPLICATION
            End If
        Case NCID_PATIENT_LIST_APPLICATION
            gobb.PatList.ProcessCAC gobb.CData
            
        Case NCID_CITA_ADMIN_APPLICATION
            'BR: This is where I perform the check to see if the patient exisits
            If gobb.Patient.SearchPatientCAC(gobb.CData) Then
                'Immunizations has to be able to support the CAC and the Uniformed Service Card for adding immunizations. If the CAC switch is on
                'this code will execute the same as when the Uniformed Service Card is read and Core sends this message based on an event from the
                'Event Server.
                '<< Begin:SCR #23216; 23217;   Developer: Brian Mowbray 08/04/2003 12:55 PM
                gobb.Comm.Message cwiBarCode, XMLString, NCID_BROADCAST_MSG, NCID_EVENT_SERVER_APPLICATION, False
                '>> End: SCR #23216; 23217;
            End If
        Case NCID_CLINIC_SCHEDULE_APPLICATION 'appointments
            'BR: This is where I perform the check to see if the patient exisits
            If gobb.Patient.SearchPatientCAC(gobb.CData) Then
                gobb.Comm.Message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_CAC_OPTIONS_APPLICATION
                frmCAC_Options.Show vbModal, gobb.Comm.MainWindow
                gobb.Comm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_CAC_OPTIONS_APPLICATION
            End If
        Case NCID_PATIENT_SEARCH_APPLICATION
            If gobb.Patient.SearchPatientCAC(gobb.CData) Then
            End If
            
        Case NCID_DATA_MANAGER
            If gobb.Patient.SearchPatientCAC(gobb.CData) Then
            End If
        
        Case Else
            'BR: This is where I perform the check to see if the patient exisits
            If gobb.Patient.SearchPatientCAC(gobb.CData) Then
                'need to see if it is ok to change patient
                gobb.Comm.Message cwiOK_TO_CLOSE_VIEW, retVal, NCID_CW_APPLICATION, 0
                
                If retVal Then
                    'then show the cac options screen
                    gobb.Comm.Message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_REASON_FOR_VISIT
                    frmCAC_Options.Show vbModal, gobb.Comm.MainWindow
                    gobb.Comm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_REASON_FOR_VISIT
                End If
            End If
    End Select

    bProcessing = False
    Exit Function
ErrorHandler:
    bProcessing = False

End Function

Public Sub Initialize(ByVal ProgId As Long, CommObject As Object, LogonEnvironment As Object, Config As Object, Patient As Object, User As Object)
    Dim tStr As String
    Dim sErrMessage As String
On Error GoTo ErrorHandler
    
    cwlMyProgId = ProgId '---save the "ProgId" assigned to me by CW---b
    
    Set gobb.Comm = CommObject
    Set gobb.Logon = LogonEnvironment
    Set gobb.Config = Config
    Set gobb.Patient = Patient
    Set gobb.User = User
    
    '---Instantiate Help-file object
    App.HelpFile = gobb.SharedObj.CWHelpFile 'SCR-HELP
    
    gobb.Comm.SubscribeToMessage NCID_EVENT_SERVER_APPLICATION, cwiCAC_Card_Message, NCID_VITAL_SIGNS_APPLICATION
    
    Exit Sub
ErrorHandler:
    Call gobb.SharedObj.ShowVBError(Err.Number, Err.Description, "CAC.Initialize", "CAC Options", vbExclamation)
 End Sub

Public Function OpenView(ByVal ViewType As Integer, ByVal ParenthWnd As Long, ByRef MainFormhWnd As Long) As Boolean
On Error GoTo errhandler
    
    If ViewType = cwiREGULAR_VIEW Then
        SetParent frmCAC_Options.hwnd, ParenthWnd
    Else
        Screen.MousePointer = vbDefault
        '*** MsgBox Fix - Code_Validator auto change made 3/21/2003
        MsgBxARMd "Unsupported view type: " & CStr(ViewType) & " in CAC.OpenView..", vbCritical, "Vital Signs"
        
        OpenView = False
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    
    '---set return value---
    OpenView = True
    Screen.MousePointer = vbDefault
    Exit Function
errhandler:
    Call gobb.SharedObj.ShowVBError(Err.Number, Err.Description, "Vitals.OpenView", "Vital Signs", vbCritical)
    OpenView = False
    Screen.MousePointer = vbDefault
End Function

Public Sub ResizeView(ByVal ViewType As Integer, ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long)
    Static firstresize As Boolean
On Error GoTo ErrorHandler
    
    '---perform different actions based upon the "view type"---
    If ViewType = cwiREGULAR_VIEW Then
        If Not firstresize Then

            frmCAC_Options.WindowState = vbNormal
            frmCAC_Options.Visible = False
            
            frmCAC_Options.Move Left, Top, Width, Height

            frmCAC_Options.Visible = True
        Else    '--first "regular" resize, need to perform "initailization" procedures--
            '---reset flag so this code will not be executed again---
            firstresize = False
            
            '---display the "main" form---
            frmCAC_Options.Show vbModal
       End If
       
    Else
        Screen.MousePointer = vbDefault
        '*** MsgBox Fix - Code_Validator auto change made 3/21/2003
        MsgBxARMd "Unsupported resize view type: " & CStr(ViewType) & " in CAC.ResizeView.", vbExclamation, "CAC Options"
    End If
    
    Exit Sub
ErrorHandler:
    Call gobb.SharedObj.ShowVBError(Err.Number, Err.Description, "CAC.ResizeView", "CAC Options", vbExclamation)
End Sub

Public Sub CloseView(ByVal ViewType As Integer)
    'not implemented - look at closeview2
End Sub

Public Sub Message(ByVal viMessageType As Integer, ByVal vsMessageText As String, ByVal vsSourceServerNCID As String)
    Dim sReturn As String
On Error GoTo errhandler
    
    'BR: If the message is not from Encounter, then I am going to ignore it. SCR 23028
    If vsSourceServerNCID <> NCID_ENCOUNTER_APPLICATION Then
        If viMessageType <> 14 And viMessageType <> 15 And viMessageType <> 16 And viMessageType <> 17 And viMessageType <> 19 Then
            Exit Sub
        End If
    End If
    
    Select Case viMessageType
        Case cwiOK_TO_CHANGE_PATIENT
            'Tell Encounter that it's ok
            gobb.Comm.Message cwiOK_TO_CHANGE_PATIENT, "Y", NCID_ENCOUNTER_APPLICATION, NCID_VITAL_SIGNS_APPLICATION
        Case cwiHISTORY
'            gobb.vitalsform.DisplayHistory Me
        Case cwiREFRESH
            '---"refresh" or repaint all running "views"---
            If cwbMainViewRunning Then
                frmCAC_Options.Visible = False
                frmCAC_Options.Visible = True
            End If
        
        Case cwiREFRESH_DATA
        
            'The return from this call should provide me with the NCID of the server that is on top
            gobb.Comm.Message cwiFOREGROUND_APPLICATION, sReturn, NCID_CW_APPLICATION, NCID_ENCOUNTER_APPLICATION
            
        
'        Case cwiCAC_Card_Message
'            StartCACProcessing vsMessageText
        
        Case Else
            'do nothing
    End Select

    Exit Sub
errhandler:
    Call gobb.SharedObj.ShowVBError(Err.Number, Err.Description, "CAC.Message", "CAC Options", vbExclamation)
End Sub

Public Function QueryView(ByVal viViewType As Integer) As Boolean
    Select Case viViewType
        Case cwiREGULAR_VIEW
            QueryView = True    '--this view is NOT optional--
        Case cwiPATIENT_SUMMARY_VIEW
            QueryView = False
        Case cwiPROPERTY_VIEW
            QueryView = False
        Case Else   '--unknown "view type", not supported--
            QueryView = False
    End Select
End Function

'currently this window will not implement any menus
Public Sub MenuItemSelected(ByVal vlMenuId As Long)
On Error GoTo errhandler
    
    Screen.MousePointer = vbHourglass
    Select Case vlMenuId
        Case Else
            '---do nothing
    End Select
    Screen.MousePointer = vbDefault
    
    Exit Sub
errhandler:
    gobb.SharedObj.ShowVBError Err.Number, Err.Description, "CAC.MenuItemSelected", "CAC Options", vbExclamation
    Screen.MousePointer = vbDefault
End Sub

Public Sub PatientSelected()
    Static lPrevPat As Long
On Error GoTo ErrorHandler

    If lPrevPat <> gobb.Patient.UnitNumber Then
        
        'need to refresh
        
        '---update variable---
        lPrevPat = gobb.Patient.UnitNumber
    End If
    
    Exit Sub
ErrorHandler:
    gobb.SharedObj.ShowVBError Err.Number, Err.Description, "CAC.PatientSelected", "CAC Options", vbExclamation
End Sub

Public Sub Properties(LogonEnviron As Object, Config As Object, ByVal OwnerNCID As String, ByVal CalledByDesigner As Boolean)
    'not implemented
End Sub

Public Sub ICHCSII_CoreAPI2_Initialize2(ByVal ProgId As Long, Comm As Object, LogonEnvironment As Object, Config As Object, Patient As Object, User As Object)
    
    Set gobb.User = User
    
    Initialize ProgId, Comm, LogonEnvironment, Config, Patient, User
    
End Sub

Private Sub Class_Initialize()
    If gobb Is Nothing Then Set gobb = New Blackboard
    mlTopMostServerID = 0
End Sub

Private Sub Class_Terminate()
    Set gobb = Nothing
End Sub

Private Function ICHCSII_CoreAPI3_CloseView2(ByVal viViewType As Integer) As Boolean
    Dim i As Long
On Error Resume Next

    ICHCSII_CoreAPI3_CloseView2 = True
    
    'RaiseEvent ShutDown
    
    If Err.Number <> 0 Then

    End If

    'Cleanup
    gobb.ClearBlackBoard
End Function
