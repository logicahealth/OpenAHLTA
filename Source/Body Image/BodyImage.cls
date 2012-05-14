VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBodyImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements ICHCSII_CoreAPI2
Implements ICHCSII_CoreAPI3
Private mlSaveMenuItem   As Long
'Private mlCancelMenuItem As Long
Private mlClearMenuItem  As Long
Private mlVisitMenuItem As Long
Private id As Date

Public Function Clear() As Boolean
    Set gRFVComplaints = Nothing
    gsEncounterNumber = vbNullString
    Set gcolInjury = New Collection
End Function

'*******************************************************************************************
'  Function Name: QueryView
'
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
    Select Case viViewType
        Case cwiREGULAR_VIEW
            QueryView = True    '--this view is NOT optional--
        Case cwiPATIENT_SUMMARY_VIEW
            QueryView = False
        Case cwiPRINT                   '- do you want the print menu available
            QueryView = False
        Case cwiPRINTPREVIEW            '- do you want the print preview menu available
            QueryView = False
        Case cwiPROPERTY_VIEW           '- do you want the properties/options button available
            QueryView = False
        Case Else   '--unknown "view type", not supported--
            QueryView = False
    End Select
End Function

'**************************************************************************************************
'  Function Name: OpenView
'
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
'**************************************************************************************************
Public Function OpenView(ByVal viViewType As Integer, ByVal vlParenthWnd As Long, rhMainForm As Long) As Boolean
    On Error GoTo ErrHandler
    Dim Signor As String
    
    If viViewType = cwiREGULAR_VIEW Then
        If Not CheckProvider Then
            OpenView = False
            Exit Function
        End If
        
        cwbFirstRegularResize = True
        cwbMainViewRunning = True
        Call Menus_Add
        Set gfrmBodyImage = New frmBodyImage
        
        
        Signor = GetSignator
        If Signor = vbNullString Then
            MsgBox "You are not permitted to be assigned as a primary provider and the default provider ('.Unassigned') is not available.", vbCritical, "Unable To Open Module"
            OpenView = False
            Exit Function
        Else
            gfrmBodyImage.Signator = Signor
        End If
        SetParent gfrmBodyImage.hwnd, vlParenthWnd
                
'-- no bi if this is an inpatient enc. (type 5)
        If Not SetMode Then
            OpenView = False
            Exit Function
        End If
        
        rhMainForm = gfrmBodyImage.hwnd
    Else
        OpenView = False
        Exit Function
    End If
        
    OpenView = True
    Exit Function
    
ErrHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "BodyImage.OpenView", "BodyImage", vbCritical)
    OpenView = False
    Exit Function
End Function

'*******************************************************************************************
'  Sub Name: CloseView
'
'  Description: This method is called when the CW Core closes a "view" provided by this OLE
'               Server.
'
'  Parameters:
'    viViewType - Contains the "type" of view that the CW Core is closing.
'*******************************************************************************************
Public Sub CloseView(ByVal viViewType As Integer)
    If viViewType = cwiALL_VIEWS Or viViewType = cwiREGULAR_VIEW Then
        If cwbMainViewRunning Then
            Call Menus_Remove
            cwbMainViewRunning = False
            cwbCloseViewMyself = True
            cwbFirstRegularResize = True
            If Not gfrmBodyImage Is Nothing Then
                Unload gfrmBodyImage          'SCR 53414 5/6/04 Phil Crowder
                Set gfrmBodyImage = Nothing
            End If
        End If
    End If
End Sub

'*******************************************************************************************
'  Sub Name: MenuItemSelected
'
'  Description: This method is called when the user has selected a menu item (from the CW
'               Core's main menu), which is "owned" by this OLE Server.
'
'               NOTE: This is an OPTIONAL method.
'
'  Parameters:
'    vlMenuId - Contains the menu id of the menu item the user selected.
'*******************************************************************************************
Public Sub MenuItemSelected(ByVal vlMenuId As Long)
    On Error GoTo ErrHandler
    Dim oInj As cInjury
    Dim bNewData As Boolean
    
    Select Case vlMenuId
        Case mlVisitMenuItem
            gfrmBodyImage.cmdVisit.Value = True
                        
        Case mlClearMenuItem
            Call gfrmBodyImage.cmdClear
            ' 107375 1/15/06 Quan
            bInjuryChanges = True
            
        Case glPtTrkMenuItem
            gfrmBodyImage.cmdOK.Value = True
            
        Case Else   '--do nothing--
    End Select

    Exit Sub

ErrHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "BodyImage.MenuItemSelected", "BodyImage", vbExclamation)
    Exit Sub
    Resume
End Sub

'**************************************************************************************************************
'  Sub Name: PatientSelected
'
'  Description: This method gets called when a new patient has been selected from the CW Core.
'
'               NOTE: This is an OPTIONAL method.
'
'  Parameters:
'    none
'**************************************************************************************************************
Public Sub PatientSelected()
    If cwbMainViewRunning Then
    
    End If
End Sub

'*******************************************************************************************
'  Sub Name: ResizeView
'
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
Public Sub ResizeView(ByVal viViewType As Integer, ByVal vlLeft As Long, ByVal vlTop As Long, ByVal vlWidth As Long, ByVal vlHeight As Long)
    On Error GoTo ErrHandler

    '---reset global "client area" parameters---
    cwlLeft = vlLeft
    cwlTop = vlTop
    cwlWidth = vlWidth
    cwlHeight = vlHeight
    
    '---perform different actions based upon the "view type"---
    If viViewType = cwiREGULAR_VIEW Then
        If Not cwbFirstRegularResize Then
            gfrmBodyImage.WindowState = vbNormal
            gfrmBodyImage.Visible = False
            gfrmBodyImage.Move cwlLeft, cwlTop, cwlWidth, cwlHeight
                
            gfrmBodyImage.Visible = True
        Else    '--first "regular" resize, need to perform "initailization" procedures--
            cwbFirstRegularResize = False
            Set gfrmBodyImage.mobjParent = Me
    '__ you have to have a patient, if not yet, select now
            If gobjPatient.UnitNumber = 0 Then
                Call gobjPatient.SearchPatientEx(True)
                DoEvents
                If gobjPatient.UnitNumber = 0 Then
                    Call CloseView(cwiREGULAR_VIEW)
                End If
            End If
            gfrmBodyImage.Display
        End If
    End If
    
    Exit Sub
    
ErrHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "BodyImage.ResizeView", "BodyImage", vbExclamation)
    
    Exit Sub
    Resume
End Sub

Private Sub Class_Initialize()
    Set gobjShared = New CWShared
    gbApptCheckedIn = False
    id = Now
    Debug.Print id
    sPriorityRTF(3) = "\highlight1\ul\b\f0\fs20 Priority:  EMERGENT\highlight0\ulnone\b0\par"
    sPriorityRTF(2) = "\highlight2\ul\b\f0\fs20 Priority:  URGENT\highlight0\ulnone\b0\par"
    sPriorityRTF(1) = "\highlight3\ul\b\f0\fs20 Priority:  MINOR\highlight0\ulnone\b0\par"
    sPriorityRTF(0) = "\highlight0\ul\b\f0\fs20 Priority:  (None)\highlight0\ulnone\b0\par"
    Set gcolInjury = New Collection
    Call LoadMedcinLookup
    App.HelpFile = gobjShared.CWHelpFile
      
End Sub

Private Sub Class_Terminate()
    Debug.Print id
    Set gobjShared = Nothing
    Set gobjComm = Nothing
    
End Sub

'**************************************************************************************************************
'  Sub Name: Message
'
'  Description: This method is used by the CW Core to send messages to this OLE Server.  The
'               messages can be from the CW Core or from another OLE Server.
'
'  Parameters:
'    viMessageType - Contains the "type" of message to process.
'    vsMessageText - Contains any "data" assoicated with this message.
'    vsSourceServerNCID - Contains the NCID of the OLE Server that "sent" the message.
'**************************************************************************************************************
Public Sub Message(ByVal viMessageType As Integer, _
                   ByVal vsMessageText As String, _
                   ByVal vsSourceServerNCID As String)
    On Error GoTo ErrHandler
   ' Debug.Print MsgText(viMessageType), vsMessageText, vsSourceServerNCID
    '-------------------------------------------------------------------------------
    '  Take corresponding actions based upon the message "type".
    '
    '  NOTE: This method is required, but currently there are NO REQUIRED messages
    '        to process.
    '-------------------------------------------------------------------------------
    Select Case viMessageType
    
        Case cwiOK_TO_CHANGE_PATIENT
            'SCR 80768 - Jane Shen 1/24/2006, save data before change patient
            If cwbMainViewRunning Then
                '--- SCR 88339     Sherry Wang   2/17/2006
                If Not gfrmBodyImage Is Nothing Then
                    gfrmBodyImage.cmdOK.Value = True
                End If
            End If
            gobjComm.Message cwiOK_TO_CHANGE_PATIENT, "Y", vsSourceServerNCID, 1999963
        
        Case cwiOK_TO_CLOSE_VIEW
            If OKToCloseView Then
                gobjComm.Message cwiOK_TO_CLOSE_VIEW, "Y", vsSourceServerNCID, 1999963
            Else
                gobjComm.Message cwiOK_TO_CLOSE_VIEW, "N", vsSourceServerNCID, 1999963
            End If
            
        Case cwiCURRENT_APPLICATION_NCID
            'SCR 90983 - JQ, 6/18/2006, save inputs before allowing another T2 module to launch
            If frmBodyImage.txtDoc.Text <> vbNullString Then
                ' we have some input, so save; using same code as change patient above
                If cwbMainViewRunning Then
                    If Not gfrmBodyImage Is Nothing Then
                        gfrmBodyImage.cmdOK.Value = True
                    End If
                End If
            End If  ' we have input
            
        
         Case Else   '--do nothing--
                        ' for TESTING:
            ' let's see what other messages we get and when:
            'MsgBox "Received message # " & Str$(viMessageType) & ", " & vsMessageText & " from " _
            ' & vsSourceServerNCID
            
            
    End Select

    Exit Sub
    
ErrHandler:

    Call gobjShared.ShowVBError(Err.Number, Err.Description, "BodyImage.Message", "BodyImage", vbExclamation)
    
    Exit Sub
End Sub

'****************************************************************************************************
'  Sub Name: Initialize
'
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
    On Error GoTo ErrHandler
    
    '---save the "ProgId" assigned to me by the CW Core---
    cwlMyProgId = vlMyProgId
    
    '---create references to passed in objects---
    Set gobjComm = robjComm
    Set gobjLogon = robjLogonEnvironment
    Set gobjconfig = robjConfig
    Set gobjPatient = robjPatient
    
    
    cwbMainViewRunning = False
    
    Call TestMatrix
    Exit Sub

ErrHandler:

    Call gobjShared.ShowVBError(Err.Number, Err.Description, "BodyImage.Initialize", "BodyImage", vbExclamation)
    
    Exit Sub
End Sub

'**************************************************************************************************************
'  Sub Name: Properties
'
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
    On Error GoTo ErrHandler
    
    '*****************************************************************
    '  TO DO:
    '    Add code here to invoke this OLE Server's properties page.
    '*****************************************************************
    Exit Sub
    
ErrHandler:

    Call gobjShared.ShowVBError(Err.Number, Err.Description, "BodyImage.Properties", "BodyImage", vbExclamation)
    
    Exit Sub
End Sub

Private Sub ICHCSII_CoreAPI2_Initialize2(ByVal vlMyProgId As Long, _
                                               robjComm As Object, _
                                               robjLogonEnvironment As Object, _
                                               robjConfig As Object, _
                                               robjPatient As Object, _
                                               robjUser As Object)
    Set gobjUser = robjUser
    Call Initialize(vlMyProgId, robjComm, robjLogonEnvironment, robjConfig, robjPatient)
End Sub

Private Function ICHCSII_CoreAPI3_CloseView2(ByVal viViewType As Integer) As Boolean
'__ core is asking us to close
'<< scr 40440
'__ deleted the cancel button and moved this code here from the menu-item-selected method
    Dim oInj As cInjury
    Dim bNewData As Boolean

    '__ see if anything is new
    bNewData = False
    For Each oInj In gcolInjury
        If oInj.NotSaved Then
            bNewData = True
            Exit For
        End If
    Next oInj
'    __if you found something new, ask to save
    If bNewData Or bInjuryChanges Then ' SCR 105401 9/26/06 Quan
        Select Case MsgBox("Save changes before closing?", vbYesNoCancel + vbQuestion, "Unsaved Data")
        Case vbYes
            gfrmBodyImage.cmdOK.Value = True
        Case vbNo
            '__ just let it fall through
        Case vbCancel
            ICHCSII_CoreAPI3_CloseView2 = False
            Exit Function
        End Select
    End If
  
  
  Call CloseView(viViewType)
  
  
  ICHCSII_CoreAPI3_CloseView2 = True
End Function

Friend Sub Menus_Add()
    'mlSaveMenuItem = gobjComm.AddActionMenuItem("&Pt Tracking", "add.bmp", cwlMyProgId, "Save Input and Close")
    mlClearMenuItem = gobjComm.AddActionMenuItem("Clear &All", "clear.bmp", cwlMyProgId, "Clear All injuries")
    mlVisitMenuItem = gobjComm.AddActionMenuItem("&Enc. Summary", "summary.bmp", cwlMyProgId, "Go to the Encounter Summary View")
    glPtTrkMenuItem = gobjComm.AddActionMenuItem("&Patient Tracking", "sched.bmp", cwlMyProgId, "Go to the patient tracking View")
    'mlCancelMenuItem = gobjComm.AddActionMenuItem("&Cancel", "cancel.bmp", cwlMyProgId, "Cancel Input")
    
End Sub
Friend Sub Menus_Remove()
    'gobjComm.DeleteMenuItem mlCancelMenuItem
    'gobjComm.DeleteMenuItem mlSaveMenuItem
    gobjComm.DeleteMenuItem mlClearMenuItem
    gobjComm.DeleteMenuItem mlVisitMenuItem
    gobjComm.DeleteMenuItem glPtTrkMenuItem
End Sub

Public Function ComplaintsfromNewVisit(colInjuries As Collection, _
                                    ByVal sApptID As String, _
                                    ByVal sEncID As String)
    Dim oInj As cInjury
    Dim v As Variant
    Dim txt() As String
    Set gcolInjury = New Collection
    For Each v In colInjuries
        gcolInjury.Add v
    Next v
    gsEncounterNumber = sEncID
    gsAppointmentID = sApptID
    If Trim(sApptID) <> vbNullString Then
        gbApptCheckedIn = True
    End If
    
End Function

Public Function SaveData() As Boolean
    On Error GoTo ErrHandler
    Dim oSQL As ICHCSII_SQL.ISqlOpsEx
    Dim oInj As cInjury
    Dim sDataID As String
    Dim sSql As String
    Dim sSysdate As String
    
    If gobjShared.UseMySQLDB Then
        sSysdate = "sysdate()"
    Else
        sSysdate = "sysdate"
    End If
        
    Set oSQL = GetSqlObj
    oSQL.Execute "select * from body_image where encounternumber = " & gsEncounterNumber
    '--- SCR 88110   Sherry Wang   2/15/2006
    If oSQL.EOF Then
        sDataID = oSQL.GetNextID
        sSql = "insert into BODY_IMAGE (DATAID, PRIORITY, SEX, ENCOUNTERNUMBER, FACILITYNCID, CREATEDBY, CREATEDON) "
        sSql = sSql & "values (" & sDataID & ", "
        sSql = sSql & "'" & giPriority & "', '" & gobjPatient.Sex & "', "
        sSql = sSql & gsEncounterNumber & ", " & gobjLogon.FacilityNCID & ", "
        sSql = sSql & gobjLogon.UserNCID & ", " & sSysdate & ") "
    Else
        sDataID = oSQL("dataid")
        sSql = "UPDATE BODY_IMAGE "
        sSql = sSql & "SET  updatedby = " & gobjLogon.UserNCID
        sSql = sSql & ", updatedon = " & sSysdate
        sSql = sSql & ", priority = '" & giPriority & "'"
        sSql = sSql & ", sex = '" & gobjPatient.Sex & "'"
        sSql = sSql & " WHERE DATAID = " & sDataID
    End If
    oSQL.Execute sSql
    
    oSQL.Execute "delete from bi_injury where dataid = " & sDataID
    
    If gcolInjury.Count > 0 Then
        For Each oInj In gcolInjury
            If oInj.FromBI Then
                sSql = "insert into BI_INJURY (DATAID, INJURY_ID, PART, SIDE_LR, SIDE_FB, "
                sSql = sSql & "INJURY_TYPE, DESCRIPTION, TYPE_ID, RFVCOMMENT, NEWFOLLOWUP, "
                sSql = sSql & "POSN_LEFT, POSN_TOP, SnoID, MEDCINDESC) values ("
                sSql = sSql & sDataID & ", " & oInj.lblIndex & ", "
                sSql = sSql & "'" & oInj.Part & "', '" & oInj.SideLR & "', "
                sSql = sSql & "'" & oInj.SideFB & "', '" & oInj.Injury & "', "
                sSql = sSql & "'" & oInj.Desc & "', " & oInj.TypeID & ", "
                sSql = sSql & "'" & oInj.RFVcomment & "', '" & oInj.NewFollowup & "', "
                sSql = sSql & oInj.lblLeft & ", " & oInj.lblTop & ", "
                sSql = sSql & "'" & oInj.Sno.SnoID & "', '" & oInj.Sno.Desc & "') "

                oSQL.Execute sSql
            End If
            oInj.NotSaved = False
        Next oInj
    End If
    
    SaveData = True
    Exit Function
ErrHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "BodyImage.SaveData", "BodyImage", vbExclamation)
    Exit Function
    Resume
End Function

Public Function RetrieveData(ByVal EncNum As String, _
                    Optional ByVal FacilityNCID As String) As Boolean
    On Error GoTo ErrHandler
    Dim ocomp As MMMHISReasonForVisit.Complaint
    Dim oSQL As ICHCSII_SQL.ISqlOpsEx
    Dim oInj As cInjury
    Dim sDataID As String
    Set oSQL = GetSqlObj
    
    Set gcolInjury = New Collection
    
    'SCR 172985 load complaints first to make sure Reason for Visit shows up when revisting the Triage module
    If gRFVComplaints Is Nothing Then
        Set gRFVComplaints = GetComplaints(EncNum)
    End If

    For Each ocomp In gRFVComplaints
        Set oInj = New cInjury
        With oInj
            .NotSaved = False
            .FromBI = False
            .DataID = ocomp.DataID
            .RFVDataID = ocomp.DataID
            .lblIndex = ocomp.BIIndex
            .Desc = vbNullString & ocomp.ComplaintText
            .Sno.SnoID = ocomp.SnoID
            .Sno.Desc = .Desc
        End With
        gcolInjury.Add oInj
    Next ocomp
    
    oSQL.Execute "select bi_injury.*, body_image.priority from bi_injury, body_image " _
    & "where body_image.dataid = bi_injury.dataid " _
    & " and encounternumber = " & EncNum _
    & " order by injury_id"
    
    If oSQL.EOF Then
        Exit Function
    End If
    
    
    giPriority = oSQL("priority")
    Dim cnt As Integer
    Do While Not oSQL.EOF
        cnt = cnt + 1
        'Set oInj = New cInjury
        For Each oInj In gcolInjury
            If oInj.Sno.SnoID = oSQL("SnoID") And cnt <= oInj.lblIndex Then
                With oInj
                    .NotSaved = False
                    .FromBI = True
                    .DataID = oSQL("DataID")
                    '.lblIndex = cnt
                    .Part = vbNullString & oSQL("part")
                    .SideLR = vbNullString & oSQL("side_lr")
                    .SideFB = vbNullString & oSQL("side_fb")
                    .Injury = vbNullString & oSQL("injury_type")
                    .Desc = vbNullString & oSQL("deScription")
                    .TypeID = vbNullString & oSQL("type_id")
                    .RFVcomment = vbNullString & oSQL("RFVComment")
                    .NewFollowup = vbNullString & oSQL("newfollowup")
                    .lblLeft = vbNullString & oSQL("posn_left")
                    .lblTop = vbNullString & oSQL("posn_top")
                    .Sno.SnoID = oSQL("SnoID")
                    .Sno.Desc = oSQL("MedcinDesc")
                End With
            End If
        Next oInj
        'gcolInjury.Add oInj
        oSQL.MoveNext
    Loop

    'SCR 172985 code was commented out and moved functionality to top of function
'__ now look for previous screening records
'    If gRFVComplaints Is Nothing Then
'        Set gRFVComplaints = GetComplaints(EncNum)
'    End If

'__ and map the injuries you just collected to the entries in the grfvcomplaints collection
'    For Each oInj In gcolInjury
'        For Each ocomp In gRFVComplaints
'<< scr 50439 we were not deleting all the right complaints.  we had multiple medcin id's and were not relating these
'   two sets of objects correctly.
'   I added the biindex condition below.
'            If ocomp.SnoID = oInj.Sno.SnoID And ocomp.BIIndex = oInj.lblIndex Then
'                oInj.RFVDataID = ocomp.DataID
'                Exit For
'            End If
'        Next ocomp
'    Next oInj
    
    RetrieveData = True
    Exit Function
ErrHandler:

End Function

Public Function FindMedcinObject(cInj As Object) As cBISno
    Set FindMedcinObject = FindSno(cInj)
End Function

Private Function SetMode() As Boolean
    On Error GoTo ErrHandler
    Dim oEnc As EncounterParent
    
    '__ if encounter visible = false,
    If Not gobjComm.Message(cwiOLE_SERVER_RUNNING, NCID_ENCOUNTER_APPLICATION, _
     NCID_CW_APPLICATION, NCID_BODYIMAGE_APPLICATION) Then
        geAppMode = 0
    
    Else    ' if encounter visible = true,
        '__ get the encounter number
        
        Set oEnc = gobjComm.InitializeOLEServer(NCID_ENCOUNTER_APPLICATION)
'<< SCR 47348 no bi when encounter type is inpatientNote
        If oEnc.CurrentEncounter.EncounterType = 5 Then   '-- inpatientNote
            '--- SCR62312 Sherry Wang  2/3/2006
            MsgBox "The Triage Module is not available when an Inpatient Note Encounter is open.", vbInformation, "Triage"
            SetMode = False
            Exit Function
        End If
'        oEnc.CurrentEncounter.EncounterType = "Inpatient"
        gsEncounterNumber = oEnc.CurrentEncounter.encounterid
        gsAppointmentID = oEnc.CurrentEncounter.AppointmentId
            
        Call RetrieveData(gsEncounterNumber)
        
            '__ if readinjuries returns no records
        If gcolInjury.Count = 0 Then
            geAppMode = 1
        Else         '__ if read injuries returns records
            geAppMode = 2
        End If
        
    End If
    SetMode = True
    Exit Function
ErrHandler:
    SetMode = False
End Function

Private Function CheckProvider() As Boolean
    On Error GoTo ErrHandler
    '__ can the current user be a provider in this clinic/facility?
    
    Dim mcolMyClinics As DTOs.Clinics
    Dim objClinicOps As CHCSII_ClinicClient.ClinicOpsEx
    If objClinicOps Is Nothing Then Set objClinicOps = New CHCSII_ClinicClient.ClinicOpsEx
     
    Set mcolMyClinics = objClinicOps.GetUserClinics(gobjLogon.FacilityNCID, gobjLogon.UserNCID)
    If mcolMyClinics.Count < 1 Then
        Screen.MousePointer = vbDefault
        MsgBox "You have not been authorized to view or change appointments for any clinics."
        Exit Function
    End If

    CheckProvider = True
    Exit Function
ErrHandler:
    Screen.MousePointer = vbDefault
    MsgBox "You have not been authorized to view or change appointments for any clinics."
    CheckProvider = True

End Function

Friend Function OKToCloseView() As Boolean
    '__if you found something new, ask to save
    Dim oInj As cInjury
    Dim bNewData As Boolean
    
    If gcolInjury Is Nothing Or gfrmBodyImage Is Nothing Then
        OKToCloseView = True
        Exit Function
    End If
    
    bNewData = False
    For Each oInj In gcolInjury
        If oInj.NotSaved Then
            bNewData = True
            Exit For
        End If
    Next oInj
    
    If bNewData Then
        Screen.MousePointer = vbDefault
        Select Case MsgBox("Save changes before closing?", vbYesNoCancel + vbQuestion, "Unsaved Data")
        Case vbYes
            Call gfrmBodyImage.MoveDataToAppt
            OKToCloseView = True
        Case vbNo
            '__ mark them all for discard
            For Each oInj In gcolInjury
                oInj.NotSaved = False
            Next oInj
            
            OKToCloseView = True
        Case vbCancel
            OKToCloseView = False
        End Select
    Else
        OKToCloseView = True
    End If
End Function
Friend Function ControlMenus(Active As Boolean) As Boolean
    gfrmBodyImage.cmdOK.Enabled = Active
    gfrmBodyImage.cmdVisit.Enabled = Active
    If Active Then
        Call gobjComm.UpdateMenuItem(mlVisitMenuItem, "enable")
        Call gobjComm.UpdateMenuItem(glPtTrkMenuItem, "enable")
    Else
        Call gobjComm.UpdateMenuItem(mlVisitMenuItem, "disable")
        Call gobjComm.UpdateMenuItem(glPtTrkMenuItem, "disable")
    End If

End Function


