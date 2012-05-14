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
Attribute VB_Name = "NewResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'*****************************************************************************************************************
'  Class Name: New Results
'
'  Description: This class is used to Manage the two lists for Results
'
'  Properties:
'
'  Methods:  Each allows an application to...
'    CloseView                 -  ICHCSII_CoreAPI Method that closes server.
'    DiscardReult              -  deletes result from list without signing.
'    ForwardResult             -  Sends on a result to another provider.
'    GetNumberOfSavedResults   -  Used to control menu display...as items are added and deleted from list
'    Initialize                -  ICHCSII_CoreAPI method.
'    MenuItemSelected          -  ICHCSII_CoreAPI method that used to relay menu actions.
'    Message                   -  ICHCSII_CoreAPI method.
'    MoveToNewResults          -  Result item is moved from saved list to new list.
'    MoveToSavedResults        -  Result item is moved from new list to saved list.
'    OpenView                  -  ICHCSII_CoreAPI method.
'    PatientSelected           -  ICHCSII_CoreAPI method not implemented.
'    Properties                -  This method allows a user the ability to modify the "properties" of this OLE
'                                 Server (i.e. configuration values).
'    QueryView                 -  ICHCSII_CoreAPI method.
'    ResizeView                -  ICHCSII_CoreAPI method.
'    Toss Result               -  Update of result causes removal from list and signing.
'    UpdateNewResult           -  Changed status from not viewed to has been viewed for a given result.
'
'
'---- Begin:SCR #269;   Developer: MAG 09/04/2002 02:53 PM
'---- Note: Function Changed: Added multiple functions to handle actions that provider takes.
'---- Do not have concurrency issues in theatre of sqlserver.
'---- So we do not need to do some of the conditional checking we do for the CDR.
'---- No egate updating is going on unbeknownst to the applciation...
'---- So we were abe to dumb these functions down.
'---- In fact, So much so in the training tool, mainly due to the fact we are
'---- not interested in generating outbound messages.
'---- IN ITT mode we did not port into SQLServer any of the PL/SQL functionality to generate outbound HL7 synchronization messages.
'---- Impact of Change: Lets us carry out supported RNR actions in SQLServer and Theatre
'---- End: SCR #269;
'
'    Gems_SQLSrv_ForwardResult   -  Sends on a result to another provider.
'    Gems_SQLSvr_DiscardResult   -  Deletes result from list without signing.
'    Gems_SQLSvr_NewResults      -  Result item is moved from saved list to new list.
'    Gems_SQLSvr_SaveResult      -  Result item is moved from new list to saved list.
'    Gems_SQLSvr_TossResult      -  Update of result causes removal from list and signing.
'*****************************************************************************************************************
Option Explicit
Implements ICHCSII_CoreAPI2
Implements ICHCSII_CoreAPI3


#If DEBUGON Then
    Private DR As DebugRec
    Private DL As cDebugLog
    Private Const Module = "RNResults"
#End If


'*******************************************************************************************
'  Sub Name: MenuItemSelected
'
'  Description: This method is called when the user has selected a menu item (from the CW
'               Core's main menu), which is "owned" by this OLE Server.
'
'
'  Parameters:
'    vlMenuId - Contains the menu id of the menu item the user selected.
'
'  HISTORY:
'  MBorsody - added new menu actions items and logic to make sure an item is slectted...a dn that there are results
'             an action upon  9/20/2001
'*******************************************************************************************
Public Sub MenuItemSelected(ByVal vlMenuId As Long)
    Dim objGrid As VSFlexGrid
    Dim sTableName As String
    Dim intNumRows As Integer
    
    On Error GoTo errhandler
    
    'get appropriate grid
    If gbIsNewResultsList Then
      Set objGrid = gfrmResults.grdResults
    Else
      Set objGrid = gfrmResults.grdSavedResults
    End If
    
    'used to test whether more than justt the grid header exist
    intNumRows = objGrid.Rows

    '---if there are no results in grid do not allow menu action function call
    If intNumRows > 1 Then
    
            Select Case vlMenuId
                Case glMenuView
                      'function accesses appropriate grid
                       gfrmResults.ViewResults
                Case glMenuToss
                      'access function depending on appmode
                        If gbCDR Then
                           TossResult
                       Else
                          Gems_SQLSvr_TossResult
                       End If
                Case glMenuSave
                        If gbCDR Then
                          MoveToSavedResults
                       Else
                          Gems_SQLSvr_SaveResult
                       End If
                Case glMenuForward
                       If gbCDR Then
                          ForwardResult
                       Else
                          Gems_SQLSrv_ForwardResult
                       End If
                Case glMenuDiscard
                       If gbCDR Then
                          DiscardResult
                       Else
                          Gems_SQLSvr_DiscardResult
                       End If
                Case glMenuNew
                      If gbCDR Then
                          MoveToNewResults
                       Else
                          Gems_SQLSvr_NewResult
                       End If
                Case Else
                      '--do nothing--
            End Select
    Else
      MsgBox " There are no results to take the selected action upon."
    End If

GoTo CleanUp

errhandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "in New Results.MenuItemSelected", "New Results", vbExclamation)
CleanUp:
    Set objGrid = Nothing
End Sub
'*******************************************************************************************
'  Sub Name: TossResult
'
'  Description: This method is called when the user has opted to sign and remove a new result
'               from their New Results list... an update causes a  triggers to
'               generate an HL7 msg, and write a signature record into the signature table.
'               As well as delete the updated record from New Results table.
'               Business rule is that a user must first view a result before they can toss it.
'  Parameters:
'
'
'  HISTORY:
'  MBorsody - Created 9/2001
'---- Begin:SCR #46294;   Developer: MAG 02/12/2004
'---- Note: Function Changed:Changed functions that return return Chcsii or Gems object to use CWShared.appmode to return the appropriate object wrapped by CHSII_Conn.
'---- Impact of Change: This change was needed to accomadate the transition over to CHCSII_Conn; CHCSII_DAS and CHCSII_SQL DLLs are going away.

'*******************************************************************************************
Public Sub TossResult()
  
    Dim sSQLResultIDs As String
    Dim sProviderIEN As String
    Dim bCancel As Boolean
    Dim sOrder_Task_No As String
    Dim sOrder_Number As String
    Dim Index As Integer
    Dim Count As Long
    Dim colRowsToDiscard As New Collection
    Dim element
    Dim objGrid As VSFlexGrid
    Dim sTableName As String
    Dim intCurrentRowIndex As Integer
    Dim blnConfirmToss As Boolean
    Dim strList As String
    Dim strPatientName As String
    Dim strCriticalAbnormalValue As String
    Dim objSQLDAS As ICHCSII_DAS
    Dim objRecset As ADODB.Recordset
    ' new stuff
    Dim strGridTestDates As String
    Dim strGridTestIDs As String
    Dim strOutcome As String
    Dim blnHideRow As Boolean
    Dim strUpdates As String
    Dim arrStrNewInfo() As String
    
    On Error GoTo errhandler
    

    
    'get variables used in call using grid
    If gbIsNewResultsList Then
      Set objGrid = gfrmResults.grdResults
      strList = gRNRnewlist
      sTableName = "Result_Event"
    Else
      Set objGrid = gfrmResults.grdSavedResults
      strList = gRNRsavelist
      sTableName = "Result_Event_Save"
    End If
    
    sOrder_Task_No = ""
    sOrder_Number = ""
    sProviderIEN = gsUserNCID
    strGridTestDates = ""
    strGridTestIDs = ""
    strOutcome = ""
    strCriticalAbnormalValue = ""
    gfrmResults.gDisableChange = True
    Count = 0
    
     ' get rowIndexes for selected rows, and do test to make sure the individual rows have ben viewed
     ' working premis - business rule- here is that they must be viewed before they can be tossed
     For Index = 1 To objGrid.Rows - 1
        If objGrid.IsSelected(Index) Then
            If objGrid.Cell(flexcpText, Index, rnrCOLresult_new) = "Y" Then
                  MsgBox ("You have selected a result that has not been viewed yet." & vbCr & vbCr & "You cannot Toss a result without first viewing it.")
                  Exit Sub
            End If
            colRowsToDiscard.Add Index
        End If
    Next Index
      
    ' only enter if have 1..n selected rows
    If colRowsToDiscard.Count > 0 Then
         
        'ask for confirmation, and then toss each individual result
        blnConfirmToss = False
        For Each element In colRowsToDiscard
            intCurrentRowIndex = CInt(element - Count)
            objGrid.Row = intCurrentRowIndex
            'changed to ask question only once for multiple rows selected
            If Not blnConfirmToss Then
                If MsgBox("Do you want to toss this result(s)?", vbYesNo) = vbNo Then
                    bCancel = True
                Else
                    blnConfirmToss = True
                End If
            End If
            
            '--- User confirms selected action
            If Not bCancel Then
                ' get stored proc parameter info
                sOrder_Number = objGrid.Cell(flexcpText, intCurrentRowIndex, rnrCOLorder_number)
                sOrder_Task_No = objGrid.Cell(flexcpText, intCurrentRowIndex, rnrCOLorder_task_no)
                strGridTestDates = objGrid.Cell(flexcpText, intCurrentRowIndex, rnrCOLcertified_dates)
                strGridTestIDs = objGrid.Cell(flexcpText, intCurrentRowIndex, rnrCOLresults_ids)
                strOutcome = RNR_StoredProc("CDRPLUS.RNR_PROVIDER_ACTIONS_PKG.RNR_PROVIDER_ACTION", sOrder_Number, sOrder_Task_No, sProviderIEN, strGridTestIDs, strGridTestDates, strList, gHL7MSGtoss)
                                
                blnHideRow = False
                Select Case strOutcome
                   Case gRNRprocCompleted
                          blnHideRow = True
                   Case gRNRprocRecordUpdated
                         'use patient name in message
                         strPatientName = objGrid.Cell(flexcpText, intCurrentRowIndex, rnrCOLpatient_name)
   
                         If MsgBox("In the time since you originally viewed this result, new test data has been received for " & strPatientName & ". " & vbCr & "The highlighted result can be tossed and removed from your new list once you have viewed the test's updated information." & vbCr & vbCr & "Click OK, If you would like to view the new information. " & vbCr & "Or " & vbCr & "Click Cancel, If you would like to continue working with your new results list and view the information later.", vbOKCancel) = vbOK Then
                            ' opted to view results
                            gfrmResults.ViewResults
                            Exit For
                         Else
                             Set objSQLDAS = SelectCHCSII_DAS
                             Set objRecset = New ADODB.Recordset
                              
                             'record has changed in DB, and user has elected not to view it... update grid with latest info
                             sSQLResultIDs = "Select result_ids,certified_dates, critical_abnormal_value from " & sTableName & " where " & _
                                             "CLINICIAN_NCID = " & sProviderIEN & " AND ORDER_NUMBER = '" & _
                                              sOrder_Number & "' AND ORDER_TASK_NO = " & sOrder_Task_No

                             Set objRecset = objSQLDAS.OpenRecordset(sSQLResultIDs)

                             If Not objRecset.EOF And Not objRecset.BOF Then
                                ' put in new info from DB record
                                 objGrid.Cell(flexcpText, intCurrentRowIndex, rnrCOLresults_ids) = objRecset.Fields(0).Value
                                 objGrid.Cell(flexcpText, intCurrentRowIndex, rnrCOLcertified_dates) = objRecset.Fields(1).Value
                                 objGrid.Cell(flexcpText, intCurrentRowIndex, rnrCOLcritical_abnormal) = objRecset.Fields(2).Value
                                 
                                 ' now do appropriate highlighting/or changing to back to normal color the text depending on the field's value
                                  HighlightAbnormalValue intCurrentRowIndex
                             End If ' empty RS
    
                             ' set viewed flag back to unviewed
                             objGrid.Cell(flexcpText, intCurrentRowIndex, rnrCOLresult_new) = "Y"
                                 
                             ' rebold row to connote unviewed status
                             objGrid.Cell(flexcpFontBold, intCurrentRowIndex, 0, objGrid.Row, objGrid.Cols - 1) = True

                             ' need to resize to accomadate larger bolded font
                             objGrid.AutoSize 0, 8, False, 1

                             ' disable option to take a toss or save... signature actions that require user has viewed the information.. row now has new unviewed data
                             gbDisableMenuItem glMenuSave, gsDisable
                             gbDisableMenuItem glMenuToss, gsDisable

                         End If ' view option
                   Case gRNRprocRecordNotFound
                          MsgBox "In the time since you originally brought up your results list " & vbCr & _
                                  "the highlighted result has been deleted out of the database by another process. " & _
                                  "The highlighted result is going to be removed from the grid so that " & vbCr & "your display will reflect the current state of the lists in the CHCS and the CHCSII data repositories."
                          blnHideRow = True
                   Case gRNRprocRecordLocked
                          MsgBox "Record is locked by another process/user. " & vbCr & "Please refresh results to ensure you are working with the current set of results"

                   Case gRNRprocErrorInExecution
                          MsgBox "Error occurred in database stored procedure when attempting to toss the highlighted result." & vbCr & "Please refresh results to ensure" & _
                                 " you are working with the current set of results"

                   End Select

                    ' now hide and remove row for processed action
                    If blnHideRow = True Then
                        objGrid.RowHidden(intCurrentRowIndex) = False
                        objGrid.RemoveItem (intCurrentRowIndex)
                        objGrid.Refresh
                        Count = Count + 1
                        If Not gbIsNewResultsList Then
                            gLTotalSavedResults = gLTotalSavedResults - 1
                        End If
                    End If
            End If ' not vb cancel
        Next
        
        'inactivate action menu if have removed last result from grid
        If objGrid.Rows = 1 Then
            gfrmResults.DeactivateMenus
            ' disable saved list tab, because is now empty
            If Not gbIsNewResultsList Then
                gfrmResults.tabResults.TabEnabled(rnrSavedTab) = False
            End If
        End If
       
       ' #25219 sometimes even the grid was failing to repaint or even appear inside the window
        If gbIsNewResultsList Then
           gfrmResults.elRNRTab.Refresh
        Else
          gfrmResults.elSTab.Refresh
        End If
        gfrmResults.Move cwlLeft, cwlTop, cwlWidth, cwlHeight
        gfrmResults.Refresh
    Else
      ' ask user to select a result to take an action upon
       MsgBox "Please select a result to take the action upon."
    End If

GoTo CleanUp

errhandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "in New Results.TossResult", "New Results", vbExclamation)

CleanUp:
    If Not objSQLDAS Is Nothing Then Set objSQLDAS = Nothing
    If Not objRecset Is Nothing Then Set objRecset = Nothing
    Set colRowsToDiscard = Nothing
    Set objGrid = Nothing
    gfrmResults.gDisableChange = False
End Sub

'*******************************************************************************************
'  Sub Name: Initialize
'
'  Description: This method is called when CW "creates" an instance of this OLE Server.  It
'               is used to create and initialize global variables that are shared across all
'               "views" supported by this OLE Server, i.e. variables that only need to be
'               initialized ONCE, upon object creation.
'
'  Parameters:
'    vlMyProgId       - Contains the "ProgId" that CW has assigned to this OLE Server.
'    robjComm         - Contains a communications object that is used to handle all communication
'                       between this OLE Server and the CW core.
'    robjLogonEnviron - Contains the logon environment object.
'    robjConfig       - Contains the configuration object used for getting and setting configuration
'                       information for a specific application.
'    robjPatient      - Contains the patient object.
'
'  History
'  MGumbert modified -10/2002 encapsultated other work done on start up into the routines called
'
'*******************************************************************************************
Public Sub Initialize(ByVal vlMyProgId As Long, robjComm As Object, robjLogonEnviron As Object, robjConfig As Object, robjPatient As Object)
    
    
    #If DEBUGON Then
        Dim DR As DebugRec
    #End If
        
    On Error GoTo errhandler
    
    #If DEBUGON Then
        DL.DebugLog Exe, Module, "Initialize", "P-Begin Initialize", DR
    #End If
    
    '---save the "ProgId" assigned to me by CW---
    cwlMyProgId = vlMyProgId
     
    '---create pointers to objects---
    Set cwobjCWComm = robjComm
    Set cwobjLogonEnviron = robjLogonEnviron
    Set cwobjConfig = robjConfig
    Set cwobjPatient = robjPatient
    
    ' pare out specific
    gsUserNCID = cwobjLogonEnviron.UserNCID
    gsUserRole = cwobjLogonEnviron.RoleNCID
    gsEnterpriseNCID = cwobjLogonEnviron.EnterpriseNCID
    gsFacilityNCID = cwobjLogonEnviron.FacilityNCID
    gsShortNameContext = cwobjLogonEnviron.ShortNameContext
    gsLongNameContext = cwobjLogonEnviron.LongNameContext
    
    '---used to control whether second tab is enabled,a nd which tab is up on view
    gLTotalSavedResults = 0
    
    gbIsNewResultsList = True
        
    '--set which views are used by dll
    cwbMainViewRunning = False
    cwbPatientSummaryViewRunning = False
    
   '---- Begin:SCR #28560;   Developer: MAG 10/16/2002 05:26 PM
   '---- Note: Function Changed: Compartmentalize work done in initialize function.
   '---- Impact of Change: Segregate task so that we can determine if network connection problems are occurring.
   
    Call SetADODBProvider
    Call GetSetOfCriticalAbnormalValues
    
    'release 2
    'stored proc no longer use logical locks so this function call is now unecessary
    'Call RemoveLocks

    '---- End: SCR #28560;
    #If DEBUGON Then
        DL.DebugLog Exe, Module, "Initialize", "P-End Initialize", DR
    #End If
    

Exit Sub
errhandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "in Results Review.Initialize", _
                                "New Results", vbExclamation)
End Sub

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
        Case cwiCLINICIAN_SUMMARY_VIEW
            QueryView = False
        Case cwiENCOUNTER_SUMMARY_VIEW
            QueryView = False
        Case Else   '--unknown "view type", not supported--
            QueryView = False
    End Select
End Function


'*******************************************************************************************
'  Function Name: OpenView
'
'  Description: This method is called when CW attempts to open a "view" provided by this
'               OLE server.
'
'  Parameters:
'    viViewType - Contains the "type" of view being requested by CW to open.
'    vlParenthWnd - Contains the hWnd of the client area window to be used by this OLE
'                   server.
'    rhMainForm (OPTIONAL) - If "viViewType" is a "regular view", upon return this will contain
'                            the window "handle" of the main "regular" form for this OLE Server.
'
'  Return Value:
'    Returns TRUE if openning the "view" was successfully.
'    Returns FALSE if an error occurred openning the "view".
'*******************************************************************************************
Public Function OpenView(ByVal viViewType As Integer, ByVal vlParenthWnd As Long, ByRef rhMainForm As Long) As Boolean
    Dim bSuccess As Boolean
    Dim lSep As Long
    
    On Error GoTo errhandler
    
        '---set global flag---
        cwbFirstRegularResize = True
            
        cwbMainViewRunning = True
    
        '----------------------------------------------------------------------
        '  Reset the "parent window" for this OLE server's "main" regular form.
        '
        '  NOTE: This invokes the form's "Load" event, but the form is NOT
        '        shown.  Make sure that you DO NOT perform a "Me.Show" in the
        '        form's "Load" event.
        '----------------------------------------------------------------------
        
        Set gfrmResults = New frmRNResults
        SetParent gfrmResults.hWnd, vlParenthWnd
        
        '---set parameter---
        If Not IsMissing(rhMainForm) Then
            rhMainForm = gfrmResults.hWnd
        End If
            
        OpenView = True

GoTo CleanUp
    
errhandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "in New Results.OpenView", _
                                "New Results", vbExclamation)
    OpenView = False
CleanUp:
    Screen.MousePointer = vbDefault
    Exit Function
End Function


'*******************************************************************************************
'  Sub Name: ResizeView
'
'  Description: This method is called when a change has occurred in the "client area"
'               used by this OLE server and it needs to resize a particular "view".
'
'  Parameters:
'    viViewType - Contains the "type" of view being requested to resize by CW.
'    vlLeft - Contains the "left" coordinate for the OLE server.
'    vlTop - Contains the "top" coordinate for the OLE server.
'    vlWidth - Contains the "width" for the OLE server.
'    vlHeight - Contains the "height" for the OLE server.
'
'  HISTORY:
'  MBorsody - Modified added reinitailizd constants 9/2001
'
'*******************************************************************************************
Public Sub ResizeView(ByVal viViewType As Integer, ByVal vlLeft As Long, ByVal vlTop As Long, ByVal vlWidth As Long, ByVal vlHeight As Long)
    On Error GoTo errhandler

    '---reset global "client area" parameters---
    cwlLeft = vlLeft
    cwlTop = vlTop
    cwlWidth = vlWidth
    cwlHeight = vlHeight
    
    Set gObjMe = Me

    '---perform different actions based upon the "view type"---
    If viViewType = cwiREGULAR_VIEW Then
                
        If Not cwbFirstRegularResize Then
            '---------------------------------------------------------------------------------
            '  Reset the "WindowState" to "Normal", so the form can be resized.  This is
            '    needed because a window CANNOT be moved or resized if it is "minimized" or
            '    "maximized".
            '
            '  NOTE: This generates a "resize" event that is ignored.
            '---------------------------------------------------------------------------------
            If gfrmResults.WindowState <> vbNormal Then
                gfrmResults.WindowState = vbNormal
            End If
            
            '---resize form to fill entire "client area", NOTE: This generates a "resize" event---
            gfrmResults.Move cwlLeft, cwlTop, cwlWidth, cwlHeight
            
            DoEvents
            
        Else
            '--first "regular" resize, need to perform "initailization" procedures--
            '---reset flag so this code will not be executed again---
            cwbFirstRegularResize = False
        
            '---resize form to fill entire "client area", NOTE: This generates a "resize" event---
            gfrmResults.Move cwlLeft, cwlTop, cwlWidth, cwlHeight
            
            '---enterring back in after unloading form need to set variable
            '---for Saved grid's results retrieval back to no results in grid... so will reload
            gbLoadedSavedResults = False
            gbIsNewResultsList = True
            
            '---need to determine if saved tab is enabled
            Call GetNumberOfSavedResults
            
            '---display the "main" form---
             gfrmResults.Display
        End If
    
    End If
    
    Exit Sub
    
errhandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "in Results Review.ResizeView", _
                                "New Results", vbExclamation)
End Sub

'**************************************************************************************************************
'  Sub Name: Message
'
'  Description:
'
'  Parameters:
'    viMessageType - Contains the "type" of message to process.
'    vsMessageText - Contains any "data" assoicated with this message.
'
'  HISTORY:
'  MBorsody - Modified minimize number of refreshes in data refresh 10/2001
'  MBorsody - Modified it to respond notOK to close when the class is
'             doing a data refresh ( eliminates problem in tracker 18604)11/12/2001
'
'**************************************************************************************************************
Public Sub Message(ByVal viMessageType As Integer, ByVal vsMessageText As String, ByVal vsSourceServer As String)
    
    On Error GoTo errhandler
    
    Select Case viMessageType
        Case cwiPRINT_PREVIEW
            cwbPrintPreview = True
            cwbPrintPreview = False
        Case cwiPRINT_REGULAR_VIEW
            cwbPrintPreview = False
        Case cwiPRINT_PATIENT_SUMMARY_VIEW
        Case cwiPRINT_ENCOUNTER_SUMMARY_VIEW
        Case cwiPATIENT_SUMMARY_DRILL_DOWN
        Case cwiENCOUNTER_SUMMARY_DRILL_DOWN
        Case cwiREFRESH
            '---"refresh" or repaint all main "views"---
            If cwbMainViewRunning Then
               gbBusy = True
               gfrmResults.Refresh
               gbBusy = False
            End If
        Case cwiABOUT
            Dim objAbout As CWAbout
            Set objAbout = New CWAbout
            cwobjCWComm.Message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_NEW_RESULTS_APPLICATION
            objAbout.Display App, gfrmResults
            cwobjCWComm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_NEW_RESULTS_APPLICATION
            Set objAbout = Nothing
        Case cwiREFRESH_DATA
           If vsSourceServer = "1801" Then
           'only respond to data refreshes eminating out of core
                If Not gfrmResults Is Nothing Then
                   'server is hidden by another server is up but has not been shown yet
                   If Not gfrmResults.Visible Then
                      '--do nothing, not initially shown when first loading
                   Else
                      '--this can be called as a result of a patient change, initiated by the
                      '--RNR module, we do not want do to a refetch of data for our patient changes...
                      '--doing as we are opening up a lab result which also on return from lab does a
                      '--refresh data... so the act of viewing a result has two net refreshes...
                      '--turned off one
                     If gbPatientChangeInProgress Then
                        Call gfrmResults.Show
                     Else
                       ' is a refresh or selection from main menu... or a patient change originating elsewhere in app, refetch data
                       gbBusy = True
                       Call gfrmResults.RefreshDisplay
                       gbBusy = False
                     End If
                  End If
                 End If
          End If
        Case cwiOK_TO_CHANGE_PATIENT
            cwobjCWComm.Message cwiOK_TO_CHANGE_PATIENT, "Y", NCID_CW_APPLICATION, NCID_NEW_RESULTS_APPLICATION
        ' tracker 18604 MB added an OK to close response based on whether class is busy re-fetching data
        ' and now module wiill not close until data refresh has concluded.
        Case cwiOK_TO_CLOSE_VIEW
            If Not gbBusy Then
                cwobjCWComm.Message cwiOK_TO_CLOSE_VIEW, "Y", NCID_CW_APPLICATION, NCID_NEW_RESULTS_APPLICATION
            Else
                cwobjCWComm.Message cwiOK_TO_CLOSE_VIEW, "N", NCID_CW_APPLICATION, NCID_NEW_RESULTS_APPLICATION
                  MsgBox "Processing data refresh command."
            End If
         End Select

    Exit Sub

errhandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "in Results Message.NewResults", _
                                "New Results", vbExclamation)
End Sub

'**************************************************************************************************************
'  Sub Name: Class_Initialize
'
'  Description: Used to set initial debugging information and create global objects.
'
'  Parameters:
'**************************************************************************************************************

Private Sub Class_Initialize()
    #If ciDEBUG = 1 Then
        If cwbDebugLogon = False Then
            Exit Sub
        End If
    #End If
    #If DEBUGON Then
        Set DL = GetObject("", "DLog.cDebugLog")
        DL.debugStart
    #End If
    gbIsNewResultsList = True
    
    Set gobjShared = New CWShared
    Set gobjCHCSIIConn = New CHCSII_CONN.Conn
    App.HelpFile = gobjShared.CWHelpFile
    
    Exit Sub

errhandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "in Results Class_Initialize.NewResults", _
                                "New Results", vbExclamation)
End Sub

'**************************************************************************************************************
'  Sub Name: Class_Terminate
'
'  Description: Used to clean up the global objects as classes destructor is called.
'
'  Parameters:
'
'  HISTORY:
'  MBorsody - Modified added clean up of global objects 9/2001
'
'**************************************************************************************************************

Private Sub Class_Terminate()

    #If ciDEBUG = 1 Then
        cwbDebugLogOff
    #End If
        

    
    #If DEBUGON Then
        DL.debugStop
        Set DL = Nothing
    #End If
    
GoTo CleanUp

errhandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "in Results Class_Terminate.NewResults", _
                                "New Results", vbExclamation)
CleanUp:
    ' added clean up MAB 9/6/2001
    If Not cwobjCWComm Is Nothing Then Set cwobjCWComm = Nothing
    If Not cwobjLogonEnviron Is Nothing Then Set cwobjLogonEnviron = Nothing
    If Not cwobjConfig Is Nothing Then Set cwobjConfig = Nothing
    If Not cwobjPatient Is Nothing Then Set cwobjPatient = Nothing
    If Not gcolRadAbnormalValues Is Nothing Then Set gcolRadAbnormalValues = Nothing
    If Not gobjShared Is Nothing Then Set gobjShared = Nothing
    If Not gobjCHCSIIConn Is Nothing Then Set gobjCHCSIIConn = Nothing
    
  
End Sub

'**************************************************************************************************************
'  Sub Name: PatientSelected
'
'  Description: This method gets called when a new patient has been selected.
'
'  Parameters:
'    none
'**************************************************************************************************************
Public Sub PatientSelected()

End Sub

'*******************************************************************************************
'  Sub Name: Properties
'
'  Description: This method allows a user the ability to modify the "properties" of this OLE
'               Server (i.e. configuration values).
'
'  Parameters:
'    robjLogonEnviron    - Contains the logon environment object.
'    robjConfig          - Contains the configuration object.
'    vsOwnerNCID         - Contains the "owner" for which to store the new/updated OLE Server
'                          property information.
'    vbCalledByDesigner  - TRUE if this method is being called by the CW Designer Tool,
'                          FALSE otherwise.
'*******************************************************************************************
Public Sub Properties(robjLogonEnviron As Object, robjConfig As Object, ByVal vsOwnerNCID As String, ByVal vbCalledByDesigner As Boolean)
    
    If vbCalledByDesigner Then
        '---create pointers to objects---
        Set cwobjLogonEnviron = robjLogonEnviron
        Set cwobjConfig = robjConfig
    End If
      
End Sub
'**************************************************************************************************************
'  Sub Name: ICHCSII_CoreAPI2_Initialize2
'
'  Description: Delegates call onto Core API1 method.
'
'  Parameters:
'    vlMyProgId           - Contains the Program ID for the server.
'    robjComm             - Object Contains the communication object the server will use to access Core Functionality.
'    robjLogonEnvironment - Object contains logon information about the user
'    robjConfig           - Object contains configuration information
'    robjPatient          - object contains information about the current patient for the app
'    robjUser             - object canotains user information
'**************************************************************************************************************
Public Sub ICHCSII_CoreAPI2_Initialize2(ByVal vlMyProgId As Long, robjComm As Object, robjLogonEnvironment As Object, robjConfig As Object, robjPatient As Object, robjUser As Object)
'// handle this guy , robjSecurity As Object)

 Set gobjUser = robjUser
 Call Initialize(vlMyProgId, robjComm, robjLogonEnvironment, robjConfig, robjPatient)

End Sub

'*******************************************************************************************
'  Sub Name: UpdateNewResultsFlag
'
'  Description: This method updates the flag that designates whether or not a result has yet been viewed.
'               We do not allow providers to Save or Toss results if they have not yet viewed.
'
'
'  HISTORY:
'  MBorsody - Modified to work with multiple grids and enumertaion constants for column numbers 9/2001
'
'*******************************************************************************************

Public Sub UpdateNewResultsFlag()
    '--- Method Added 2/25/2000 - JMS - SCR #466
    Dim sSQL As String, sProviderIEN As String
    Dim sEventID As String
    Dim bCancel As Boolean
    Dim sOrderNumber As String, sOrder_Task_No As String
    Dim sFacility_Ncid As String
    Dim sTableName As String
    Dim objGrid As VSFlexGrid
    Dim objSQL As ICHCSII_SQL.ISqlOps


    Set objSQL = SelectSQL_ISqlOps
    sProviderIEN = gsUserNCID
    sSQL = ""
    sTableName = ""
    
    On Error GoTo errhandler
    
    ' using one of two tables/grids for synchronization
    If gbIsNewResultsList Then
        sTableName = "result_event"
        Set objGrid = gfrmResults.grdResults
    Else
       sTableName = "result_event_save"
       Set objGrid = gfrmResults.grdSavedResults
    End If

    '--- Run update SQL
    If Not bCancel Then
        ' now we are using a hidden field to say whether is a viewed result or not... not bolded txt
        ' so set it to connote fact it has been viewed
        objGrid.Cell(flexcpText, objGrid.Row, rnrCOLresult_new) = "N"

        ' get fields that uniquely identify record
        sOrder_Task_No = objGrid.Cell(flexcpText, objGrid.Row, rnrCOLorder_task_no)
        sOrderNumber = objGrid.Cell(flexcpText, objGrid.Row, rnrCOLorder_number)
        
        sFacility_Ncid = objGrid.Cell(flexcpText, objGrid.Row, rnrCOLfacility_ncid)
        sSQL = "UPDATE " & sTableName & " set NEW_RESULTS_FLAG = 'N' WHERE " & _
               "CLINICIAN_NCID = " & sProviderIEN & " AND ORDER_NUMBER = '" & _
                sOrderNumber & "' AND ORDER_TASK_NO   = " & sOrder_Task_No
                
        Call objSQL.Execute(sSQL)
        objGrid.Cell(flexcpFontBold, objGrid.Row, 0, objGrid.Row, objGrid.Cols - 1) = False
    End If
    
GoTo CleanUp
   
errhandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "in New Results.UpdateNewResultsFlag", "New Results", vbExclamation)
CleanUp:
    Set objSQL = Nothing
    Set objGrid = Nothing
End Sub

'*******************************************************************************************
'  Sub Name: DiscardResult
'
'  Description: This method is called when the user has opted to discard a result from
'               their new results list without signing it.  Discard unlike toss and save does
'               not require viewing first.
'
'  Parameters:
'
'  HISTORY:
'  MBorsody - Created 9/2001
'---- Begin:SCR #46294;   Developer: MAG 02/12/2004
'---- Note: Function Changed:Changed functions that return return Chcsii or Gems object to use CWShared.appmode to return the appropriate object wrapped by CHSII_Conn.
'---- Impact of Change: This change was needed to accomadate the transition over to CHCSII_Conn; CHCSII_DAS and CHCSII_SQL DLLs are going away.
'*******************************************************************************************
Public Sub DiscardResult()

    Dim sProviderIEN As String
    Dim bCancel As Boolean
    Dim sOrder_Task_No As String
    Dim sOrder_Number As String
    Dim Index As Integer
    Dim colRowsToDiscard As Collection
    Dim element
    Dim intCount As Integer
    Dim objGrid As VSFlexGrid
    Dim sTableName As String
    Dim intCurrentRowIndex As Integer
    Dim blnConfirmDiscard As Boolean
    Dim objSQL As ICHCSII_SQL.ISqlOps
    Dim objRecset As New ADODB.Recordset
    Dim sSQLDiscard As String
    Dim blnHideRow As Boolean
    Dim strDTsResult As String
    Dim strPatientName As String
    Dim strList As String
    Dim strGridTestDates As String
    Dim strGridTestIDs As String
    Dim strOutcome As String

    
    On Error GoTo errhandler
    
    Set objSQL = SelectSQL_ISqlOps
    Set objRecset = New ADODB.Recordset
    
    'get variables used for specific list
    If gbIsNewResultsList Then
      Set objGrid = gfrmResults.grdResults
      strList = gRNRnewlist
      sTableName = "Result_Event"
    Else
      Set objGrid = gfrmResults.grdSavedResults
      strList = gRNRsavelist
      sTableName = "Result_Event_Save"
    End If

    sProviderIEN = gsUserNCID

    sOrder_Task_No = ""
    sOrder_Number = ""
    strGridTestDates = ""
    strGridTestIDs = ""
    strOutcome = ""
    gfrmResults.gDisableChange = True

    ' get list of selected indices
    Set colRowsToDiscard = gfrmResults.SelectedRowsIndex
    If colRowsToDiscard.Count > 0 Then
        intCount = 0
        blnConfirmDiscard = False
     
         For Each element In colRowsToDiscard
            intCurrentRowIndex = CInt(element - intCount)
            objGrid.Row = intCurrentRowIndex
                    
            'only ask question once for multiple results
            If Not blnConfirmDiscard Then
                If MsgBox("Do you want to discard this result(s)?", vbYesNo) = vbNo Then
                    bCancel = True
                Else
                   blnConfirmDiscard = True
                End If
            End If

            If Not bCancel Then
             '  get stored proc paramter info
                sOrder_Number = objGrid.Cell(flexcpText, intCurrentRowIndex, rnrCOLorder_number)
                sOrder_Task_No = objGrid.Cell(flexcpText, intCurrentRowIndex, rnrCOLorder_task_no)
                strGridTestDates = objGrid.Cell(flexcpText, intCurrentRowIndex, rnrCOLcertified_dates)
                strGridTestIDs = objGrid.Cell(flexcpText, intCurrentRowIndex, rnrCOLresults_ids)
                strOutcome = RNR_StoredProc("CDRPLUS.RNR_PROVIDER_ACTIONS_PKG.RNR_PROVIDER_ACTION", sOrder_Number, sOrder_Task_No, sProviderIEN, strGridTestIDs, strGridTestDates, strList, gHL7MSGdiscard)
                
                blnHideRow = False
                Select Case strOutcome
                   Case gRNRprocCompleted
                          blnHideRow = True
                   Case gRNRprocRecordUpdated
                        'use patient name in message
                        strPatientName = objGrid.Cell(flexcpText, intCurrentRowIndex, rnrCOLpatient_name)
                        If MsgBox("In the time since you originally loaded your result list, the highlighted result for " & strPatientName & "  has been updated. " & vbCr & vbCr & "Click OK, If you would like to view the result's new information." & vbCr & "Or " & vbCr & "Click Cancel, If you would like to continue with your discard.", vbOKCancel) = vbOK Then
                            ' opted to view results
                            gfrmResults.ViewResults
                            Exit For
                         Else
                            ' record has changed in DB, and user has elected to continue with discard result
                            sSQLDiscard = "UPDATE " & sTableName & " set STATUS = '" & gHL7MSGdiscard & "', GENERATE_HL7 = " & rnrMSGoutbound & ", IN_USE_FLAG = 0 WHERE " & _
                                          "CLINICIAN_NCID = " & sProviderIEN & " AND ORDER_NUMBER = '" & _
                                          sOrder_Number & "' AND ORDER_TASK_NO = " & sOrder_Task_No
                            Call objSQL.Execute(sSQLDiscard)
                            blnHideRow = True
                         End If ' view option
                   Case gRNRprocRecordNotFound
                          ' not considered an error or notable cause user was opting to discard result without viewing
                           blnHideRow = True
                   
                   Case gRNRprocRecordLocked
                          MsgBox "Record is locked by another process/user. " & vbCr & "Please refresh results to ensure you are working with the current set of results."

                   Case gRNRprocErrorInExecution
                          MsgBox "Error occurred in database stored procedure when attempting to discard the highlighted result." & vbCr & "Please refresh results to ensure" & _
                                 " you are working with the current set of results"
                   End Select

                    ' now hide and remove row for processed action
                    If blnHideRow = True Then
                        objGrid.RowHidden(intCurrentRowIndex) = False
                        objGrid.RemoveItem (intCurrentRowIndex)
                        objGrid.Refresh
                        intCount = intCount + 1
                        If Not gbIsNewResultsList Then
                            gLTotalSavedResults = gLTotalSavedResults - 1
                        End If
                    End If
            End If ' not vb cancel
        Next
       
        'inactivate action menu
        If objGrid.Rows = 1 Then
           'just have grid header
            gfrmResults.DeactivateMenus
            ' disable saved list tab... the next time after clicking off of it, the tab will be disabled...  Because it is now empty.
            If Not gbIsNewResultsList Then
                gfrmResults.tabResults.TabEnabled(rnrSavedTab) = False
            End If
        End If
            
       ' #25219 sometimes even the grid was failing to repaint or even appear inside the window
        If gbIsNewResultsList Then
           gfrmResults.elRNRTab.Refresh
        Else
          gfrmResults.elSTab.Refresh
        End If
        
        gfrmResults.Move cwlLeft, cwlTop, cwlWidth, cwlHeight
        gfrmResults.Refresh

    Else
        ' ask user to select a row to take desired action upon
       MsgBox "Please select a result to take the action upon."
    End If
        
GoTo CleanUp

errhandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "in New Results.DiscardResult", "New Results", vbExclamation)
CleanUp:

    If Not objSQL Is Nothing Then Set objSQL = Nothing
    If Not objRecset Is Nothing Then Set objRecset = Nothing
  
    Set colRowsToDiscard = Nothing
    Set objGrid = Nothing
    gfrmResults.gDisableChange = False
End Sub

'*******************************************************************************************
'  Sub Name: MoveToSavedResults
'
'  Description: This method is called when the user has opted to sign and remove a new result
'               from their New Results list and move it to their saved list... does an update on effected record,
'               that causes a trigger to generate an HL7 msg, and it writes a signature record into
'               the signature table.  As well as delete the effected record from New Results table.
'               Business rule - a user must first view a result before they can save it.
'
'  Parameters:
'
'  HISTORY:
'  MBorsody - Created 9/2001
'---- Begin:SCR #46294;   Developer: MAG 02/12/2004
'---- Note: Function Changed:Changed functions that return return Chcsii or Gems object to use CWShared.appmode to return the appropriate object wrapped by CHSII_Conn.
'---- Impact of Change: This change was needed to accomadate the transition over to CHCSII_Conn; CHCSII_DAS and CHCSII_SQL DLLs are going away.
'*******************************************************************************************
Public Sub MoveToSavedResults()

    Dim sSQLResultIDs As String, sProviderIEN As String
    Dim bCancel As Boolean, strMsg As String
    Dim sOrder_Task_No As String, sOrder_Number As String
    Dim Index As Integer
    Dim Count As Long
    Dim colSelectedRows As New Collection
    Dim element
    Dim blnConfirmSave As Boolean
    Dim strPatientName As String
    Dim intCurrentIndex As Integer
    Dim strCriticalAbnormalValue As String
    Dim objSQLDAS As ICHCSII_DAS
    Dim objRecset As ADODB.Recordset

    ' new stuff
    Dim strGridTestDates As String
    Dim strGridTestIDs As String
    Dim strOutcome As String
    Dim blnHideRow As Boolean
    Dim strUpdates As String

  
    On Error GoTo errhandler

    sProviderIEN = gsUserNCID
    gfrmResults.gDisableChange = True
    strCriticalAbnormalValue = ""
    
    
    
    ' get rowIndexes for selected rows, and do test to make sure the individual rows have ben viewed
    ' working premis here is that they must be viewed before they can be saved
    For Index = 1 To gfrmResults.grdResults.Rows - 1
        If gfrmResults.grdResults.IsSelected(Index) Then
            If gfrmResults.grdResults.Cell(flexcpText, Index, rnrCOLresult_new) = "Y" Then
              MsgBox ("You have selected a result that has not been viewed yet." & vbCr & vbCr & "You cannot Save a result without first viewing it.")
              Exit Sub
            End If
            colSelectedRows.Add Index
        End If
    Next Index
    
   'test to ensure a row is selected
   If colSelectedRows.Count > 0 Then
     
        Count = 0
        blnConfirmSave = False
        ' use the hide and remove functionality of grid instead of doing a
        For Each element In colSelectedRows
            intCurrentIndex = (element - Count)
            gfrmResults.grdResults.Row = intCurrentIndex
            
            If Not blnConfirmSave Then
                If MsgBox("Do you want to save this result(s)?", vbYesNo) = vbNo Then
                    bCancel = True
                Else
                   blnConfirmSave = True
                End If
            End If
            
             '--- User confirms selected action
            If Not bCancel Then
                ' get stored proc paramter info
                sOrder_Number = gfrmResults.grdResults.Cell(flexcpText, intCurrentIndex, rnrCOLorder_number)
                sOrder_Task_No = gfrmResults.grdResults.Cell(flexcpText, intCurrentIndex, rnrCOLorder_task_no)
                strGridTestDates = gfrmResults.grdResults.Cell(flexcpText, intCurrentIndex, rnrCOLcertified_dates)
                strGridTestIDs = gfrmResults.grdResults.Cell(flexcpText, intCurrentIndex, rnrCOLresults_ids)
                'invoke stored proc
                strOutcome = RNR_StoredProc("CDRPLUS.RNR_PROVIDER_ACTIONS_PKG.RNR_PROVIDER_ACTION", sOrder_Number, sOrder_Task_No, sProviderIEN, strGridTestIDs, strGridTestDates, gRNRnewlist, gHL7MSGsave)
                
                blnHideRow = False
                Select Case strOutcome
                   Case gRNRprocCompleted
                          blnHideRow = True
                   Case gRNRprocRecordUpdated
                         'use patient name in message at times
                         '----tracker#29021 MAG 11/05/2002
                         '----Function Changed():MovedToSavedResults was changed to use the new list's grid... it was mistakenly using the save list when trying to retrieve the patient's name.  I just recently rewrote this part of the function in  to display appropriate messages which convey the outcomes of the stored procedure invocation.
                         '----Impact of Change: It stops an array out of bounds error from occurring.
                         strPatientName = gfrmResults.grdResults.Cell(flexcpText, intCurrentIndex, rnrCOLpatient_name)

                         If MsgBox("In the time since you originally viewed this result, new result data has been received for " & strPatientName & ". " & vbCr & "The highlighted result can be saved once you have viewed the result's updated information." & vbCr & vbCr & "Click OK, If you would like to view the new information. " & vbCr & "Or " & vbCr & "Click Cancel, If you would like to continue working with your new results list and view the information later.", vbOKCancel) = vbOK Then
                            ' opted to view results
                            gfrmResults.ViewResults
                            Exit For
                         Else
                             Set objSQLDAS = SelectCHCSII_DAS()
                             Set objRecset = New ADODB.Recordset
                              
                             'record has changed in DB, and user has elected not to view it... update grid with latest info
                             sSQLResultIDs = "Select result_ids,certified_dates, critical_abnormal_value from result_event where " & _
                                             "CLINICIAN_NCID = " & sProviderIEN & " AND ORDER_NUMBER = '" & _
                                              sOrder_Number & "' AND ORDER_TASK_NO = " & sOrder_Task_No

                             Set objRecset = objSQLDAS.OpenRecordset(sSQLResultIDs)

                             If Not objRecset.EOF And Not objRecset.BOF Then
                                
                                ' put in new info from DB record
                                 gfrmResults.grdResults.Cell(flexcpText, intCurrentIndex, rnrCOLresults_ids) = objRecset.Fields(0).Value
                                 gfrmResults.grdResults.Cell(flexcpText, intCurrentIndex, rnrCOLcertified_dates) = objRecset.Fields(1).Value
                                 gfrmResults.grdResults.Cell(flexcpText, intCurrentIndex, rnrCOLcritical_abnormal) = objRecset.Fields(2).Value
                                 
                                 ' now do appropriate highlighting/or changing to back to normal color the text depending on the field's value
                                 HighlightAbnormalValue intCurrentIndex
                                 
                             End If ' empty RS
                            
                             
                            '  set viewed flag back to unviewed
                            gfrmResults.grdResults.Cell(flexcpText, intCurrentIndex, rnrCOLresult_new) = "Y"
                                 
                            ' rebold row to connote unviewed status
                            gfrmResults.grdResults.Cell(flexcpFontBold, intCurrentIndex, 0, gfrmResults.grdResults.Row, gfrmResults.grdResults.Cols - 1) = True

                            ' resize to accomadate larger bolded font
                            gfrmResults.grdResults.AutoSize 0, 8, False, 1

                            ' disable option to take a toss or save...signature actions that require user has viewed the information for this row... has new unviewed data
                            gbDisableMenuItem glMenuSave, gsDisable
                            gbDisableMenuItem glMenuToss, gsDisable

                         End If 'view option
                   Case gRNRprocRecordNotFound
                          MsgBox "In the time since you originally brought up your new results list " & vbCr & _
                          "the highlighted result has been deleted out of the database by another process. " & _
                          "The highlighted result is going to be removed from the grid so that " & vbCr & "your display will reflect the current state of the lists in the CHCS and the CHCSII data repositories."

                          blnHideRow = True
                   Case gRNRprocRecordLocked
                          MsgBox "Record is locked by another process/user. " & vbCr & "Please refresh results to ensure you are working with the current set of results"

                   Case gRNRprocErrorInExecution
                          MsgBox "Error occurred in database stored procedure when attempting to save the result to your saved list." & vbCr & "Please refresh results to ensure" & _
                                 " you are working with the current set of results"

                   End Select

                    ' now hide and remove row for processed action
                    If blnHideRow = True Then
                        gfrmResults.grdResults.RowHidden(intCurrentIndex) = False
                        gfrmResults.grdResults.RemoveItem (intCurrentIndex)
                        gfrmResults.grdResults.Refresh
                        Count = Count + 1
                        gLTotalSavedResults = gLTotalSavedResults + 1
                    End If
                    
            End If ' not vb cancel
        Next
        
        ' test to see if we need to activate tab because we now have saved results
        If gLTotalSavedResults > 0 Then
            gfrmResults.tabResults.TabEnabled(rnrSavedTab) = True
        End If
        
        ' inactivate action menu if grid that is up empties of results
        If gfrmResults.grdResults.Rows = 1 Then
           ' have no data just header
            gfrmResults.DeactivateMenus
        End If
        
         ' #25219 sometimes even the grid was failing to repaint or even appear inside the window
        If gbIsNewResultsList Then
           gfrmResults.elRNRTab.Refresh
        Else
          gfrmResults.elSTab.Refresh
        End If
        
        gfrmResults.Move cwlLeft, cwlTop, cwlWidth, cwlHeight
        gfrmResults.Refresh
    Else
       MsgBox "Please select a result to take the action upon."
    End If
    
    GoTo CleanUp
  
errhandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "in New Results.MoveToSavedResults", "New Results", vbExclamation)
CleanUp:
    If Not objSQLDAS Is Nothing Then Set objSQLDAS = Nothing
    If Not objRecset Is Nothing Then Set objRecset = Nothing
    Set colSelectedRows = Nothing
    gfrmResults.gDisableChange = False
End Sub

'*******************************************************************************************
'  Sub Name: ForwardResult
'
'  Description: This method is called when the user has opted to forward  a result onto another provider. A function
'               is called and it will make a copy using the selected result as a template. Result stays on provider's list.
'               A result can be forwarded from either the New or saved list.  But it can only be forwarded onto a provider
'               new list(business rule).
'
'  Parameters:
'
'  HISTORY:
'  MBorsody - Created 9/2001
'---- Begin:SCR #46294;   Developer: MAG 02/12/2004
'---- Note: Function Changed:Changed functions that return return Chcsii or Gems object to use CWShared.appmode to return the appropriate object wrapped by CHSII_Conn.
'---- Impact of Change: This change was needed to accomadate the transition over to CHCSII_Conn; CHCSII_DAS and CHCSII_SQL DLLs are going away.
'*******************************************************************************************
Public Sub ForwardResult()

    Dim sSQL As String
    Dim sProviderIEN As String
    Dim bCancel As Boolean
    Dim sOrder_Task_No As String
    Dim sOrder_Number As String
    Dim Index As Integer
    Dim colRowsToForward As Collection
    Dim element
    Dim sProviderName As String
    Dim sorigProviderIEN As String
    Dim objGrid As VSFlexGrid
    Dim sTableName  As String
    Dim sFwdProviderIEN As String
    Dim blnConfirmForward As Boolean
    Dim intCurrentRowIndex As Integer
    Dim strPatientName As String
    Dim strGridTestDates As String
    Dim strGridTestIDs As String
    Dim strOutcome As String
    Dim blnHideRow As Boolean
    Dim strList As String
    Dim intCount As Integer
   
    On Error GoTo errhandler
    
    sProviderIEN = gsUserNCID
    sSQL = ""
    sProviderName = ""
    sFwdProviderIEN = ""
    gfrmResults.gDisableChange = True
    
    sProviderName = gfrmResults.txtProvider.Text
    
    'get variables used in call using grid
    If gbIsNewResultsList Then
      Set objGrid = gfrmResults.grdResults
      strList = gRNRnewlist
      sTableName = "Result_Event"
    Else
      Set objGrid = gfrmResults.grdSavedResults
      strList = gRNRsavelist
      sTableName = "Result_Event_Save"
    End If
    
    sorigProviderIEN = gsUserNCID
        
    ' make sure doctor is selected... get their name... cannot forward without it
    If sProviderName = "" Then
       MsgBox ("Please select the provider you would like to send the result to.")
    Else
         ' get collection of selected indices
         Set colRowsToForward = gfrmResults.SelectedRowsIndex
         
         'ensure have one or more results to work on
         If colRowsToForward.Count > 0 Then
             intCount = 0
             blnConfirmForward = False
             For Each element In colRowsToForward
                'set working row
                intCurrentRowIndex = CInt(element - intCount)
                objGrid.Row = intCurrentRowIndex
                
                ' ask once for multiple results
                If Not blnConfirmForward Then
                    If MsgBox("Do you want to forward this result(s) to provider " & sProviderName & "?", vbYesNo) = vbNo Then
                        bCancel = True
                    Else
                       blnConfirmForward = True
                    End If
                 End If
                
                '--- actual forward is done by trigger invoked by update in stored proc
                If Not bCancel Then
                    sOrder_Number = objGrid.Cell(flexcpText, intCurrentRowIndex, rnrCOLorder_number)
                    sOrder_Task_No = objGrid.Cell(flexcpText, intCurrentRowIndex, rnrCOLorder_task_no)
                    strPatientName = objGrid.Cell(flexcpText, intCurrentRowIndex, rnrCOLpatient_name)
                    sFwdProviderIEN = Trim(gfrmResults.ForwardingProvider)
                    
                    strOutcome = RNR_StoredProc("CDRPLUS.RNR_PROVIDER_ACTIONS_PKG.RNR_FORWARD_RESULT", sOrder_Number, sOrder_Task_No, sorigProviderIEN, strGridTestIDs, strGridTestDates, strList, gHL7MSGforward, sFwdProviderIEN)
                    
                    blnHideRow = False
                    Select Case strOutcome
                       Case gRNRprocRecordNotFound
                             'use patient name in message
                              strPatientName = objGrid.Cell(flexcpText, intCurrentRowIndex, rnrCOLpatient_name)
                              MsgBox "In the time since you originally brought up your new results list " & vbCr & _
                              "the highlighted result for " & strPatientName & " has been deleted out of the database by another process. " & _
                              "The highlighted result is going to be removed from the grid so that " & vbCr & "your display will reflect the current state of the lists in the CHCS and the CHCSII data repositories."
                              blnHideRow = True
                              intCount = intCount + 1
                       Case gRNRprocRecordLocked
                              MsgBox "Record is locked by another process/user. " & vbCr & "Please refresh results to ensure you are working with the current set of results"

                       Case gRNRprocErrorInExecution
                              MsgBox "Error occurred in database stored procedure when attempting to forward the highlighted result." & vbCr & "Please refresh results to ensure" & _
                                     " you are working with the current set of results"
                       Case gRNRprocRecordAlreadyExist
                               MsgBox "The provider already has the highlighted result on one of their result lists."
                   End Select

                    ' now hide and remove row for processed action
                    If blnHideRow = True Then
                        objGrid.RowHidden(intCurrentRowIndex) = False
                        objGrid.RemoveItem (intCurrentRowIndex)
                        objGrid.Refresh
                        intCount = intCount + 1
                        If Not gbIsNewResultsList Then
                            gLTotalSavedResults = gLTotalSavedResults - 1
                        End If
                    End If
                    
                End If ' cancel
             Next
           
             ' now set the provider input box and
             ' the local variable that stores that tab's provider state back to empty
             ' so refresh elastic that holds the grid and the text input box
             If gbIsNewResultsList Then
                 gfrmResults.NewListProvider = ""
                 gfrmResults.elSTab.Refresh
             Else
                gfrmResults.SavedListProvider = ""
                gfrmResults.elSTab.Refresh
             End If
           
             gfrmResults.txtProvider.Text = ""
            ' #25219 sometimes even the grid was failing to repaint or even appear inside the window
            gfrmResults.Move cwlLeft, cwlTop, cwlWidth, cwlHeight
            gfrmResults.Refresh
            
         Else
             ' no item is selected ask user to select one
            MsgBox "Please select a result to take the action upon."
         End If
   End If

   GoTo CleanUp
    
errhandler:
   '---show the VB error---
   Call gobjShared.ShowVBError(Err.Number, Err.Description, "in New Results.ForwardResult", "New Results", vbExclamation)
CleanUp:
    Set objGrid = Nothing
    Set colRowsToForward = Nothing
    gfrmResults.gDisableChange = False
End Sub

'*******************************************************************************************
'  Sub Name: MoveToNewResults
'
'  Description: This method is called when the user has opted to delete a result from
'               their saved results list which has previously been signed, and
'               move it back to the new list.
'
'               Updated results are moved from the saved list automatically back to the new list
'               by a stored proc on the database.  So the fact the result is still on the saved list
'               attest to it not having been updated (no state change comaprisons needed).
'  Parameters:
'
'  HISTORY:
'  MBorsody - Created 9/2001
'---- Begin:SCR #46294;   Developer: MAG 02/12/2004
'---- Note: Function Changed:Changed functions that return return Chcsii or Gems object to use CWShared.appmode to return the appropriate object wrapped by CHSII_Conn.
'---- Impact of Change: This change was needed to accomadate the transition over to CHCSII_Conn; CHCSII_DAS and CHCSII_SQL DLLs are going away.
'*******************************************************************************************

Public Sub MoveToNewResults()

    Dim sProviderIEN As String
    Dim bCancel As Boolean
    Dim sOrder_Task_No As String
    Dim sOrder_Number As String
    Dim strPatientName As String
    Dim Count As Long
    Dim colRowsSelected As Collection
    Dim element
    Dim intCurrentIndex As Integer
    Dim blnConfirmNew As Boolean
    Dim strGridTestDates As String
    Dim strGridTestIDs As String
    Dim strOutcome As String
    Dim blnHideRow As Boolean
  

    On Error GoTo errhandler
    sProviderIEN = gsUserNCID

    gfrmResults.gDisableChange = True
    
    ' get collection on selected rows
    Set colRowsSelected = gfrmResults.SelectedRowsIndex
    
    If colRowsSelected.Count > 0 Then
         Count = 0

         blnConfirmNew = False
         'ask for confirmation, and then move back to saved list
         For Each element In colRowsSelected
            intCurrentIndex = (element - Count)
            gfrmResults.grdSavedResults.Row = intCurrentIndex
            If Not blnConfirmNew Then
                If MsgBox("Do you want to move this result(s) back to your new result list?", vbYesNo) = vbNo Then
                    bCancel = True
                Else
                    blnConfirmNew = True
                End If
            End If
            
            '--- user wants to proceed with selected action
            If Not bCancel Then
                '  get stored proc paramter info
                sOrder_Number = gfrmResults.grdSavedResults.Cell(flexcpText, intCurrentIndex, rnrCOLorder_number)
                sOrder_Task_No = gfrmResults.grdSavedResults.Cell(flexcpText, intCurrentIndex, rnrCOLorder_task_no)
                strGridTestDates = gfrmResults.grdSavedResults.Cell(flexcpText, intCurrentIndex, rnrCOLcertified_dates)
                strGridTestIDs = gfrmResults.grdSavedResults.Cell(flexcpText, intCurrentIndex, rnrCOLresults_ids)
                strOutcome = RNR_StoredProc("CDRPLUS.RNR_PROVIDER_ACTIONS_PKG.RNR_PROVIDER_ACTION", sOrder_Number, sOrder_Task_No, sProviderIEN, strGridTestIDs, strGridTestDates, gRNRsavelist, gHL7MSGnew)
 
                blnHideRow = False
                Select Case strOutcome
                   Case gRNRprocCompleted
                             blnHideRow = True
                   Case gRNRprocRecordUpdated
                            ' is a non sequitor in theory cannot happen... all updates to records found on saved list result in those
                            ' records being moved back behind the scenes to the new list
                            blnHideRow = True
                   Case gRNRprocRecordNotFound
                           'use patient name in message
                           strPatientName = gfrmResults.grdSavedResults.Cell(flexcpText, intCurrentIndex, rnrCOLpatient_name)
                           MsgBox "In the time since you originally brought up your saved results list " & vbCr & _
                                 "the highlighted result for " & strPatientName & " has been deleted out of the database by another process. " & _
                                 "The highlighted result is going to be removed from the grid so that " & vbCr & "your display will reflect the current state of the list in the CHCS and CHCSII datastores."

                           blnHideRow = True
                   Case gRNRprocRecordLocked
                          MsgBox "Record is locked by another process/user. " & vbCr & "Please refresh results to ensure you are working with the current set of results."

                   Case gRNRprocErrorInExecution
                          MsgBox "Error occurred in database stored procedure when attempting to move the result back to the new list." & vbCr & "Please refresh results to ensure" & _
                                 " you are working with the current set of results"
                   End Select

                    ' now hide and remove row for processed action
                    If blnHideRow = True Then
                        gfrmResults.grdSavedResults.RowHidden(intCurrentIndex) = False
                        gfrmResults.grdSavedResults.RemoveItem (intCurrentIndex)
                        gfrmResults.grdSavedResults.Refresh
                        Count = Count + 1
                        gLTotalSavedResults = gLTotalSavedResults - 1
                    End If
            End If
        Next
        
        'test to see if we need to inactivate action menu
        If gfrmResults.grdSavedResults.Rows = 1 Then
           ' have no data just header
            gfrmResults.DeactivateMenus
            gfrmResults.tabResults.TabEnabled(rnrSavedTab) = False
        End If
        
        ' sometimes grey block where window confriming action was, does not disappear, refresh of grid
        ' #25219 sometimes even the grid was failing to repaint or even appear inside the window
        gfrmResults.elSTab.Refresh
        gfrmResults.Move cwlLeft, cwlTop, cwlWidth, cwlHeight
        gfrmResults.Refresh
    Else
       MsgBox "Please select a result to take the action upon."
    End If
        
    GoTo CleanUp

errhandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "in New Results.MoveToNewResults", "New Results", vbExclamation)
CleanUp:
    Set colRowsSelected = Nothing
    gfrmResults.gDisableChange = False
End Sub

'*******************************************************************************************
'  Sub Name: GetNumberOfSavedResults
'
'  Description: This method returns the number of saved results for a provider. We enable/disable the
'               saved result tab when the form first loads based on whether or not he has results.
'  Parameters:
'
'  HISTORY:
'  MBorsody - Created 9/2001
'
'*******************************************************************************************

Private Sub GetNumberOfSavedResults()
  Dim sSQL As String
  Dim sProviderIEN As String
  Dim objSQL As ICHCSII_DAS
  Dim objRecset As ADODB.Recordset
  Dim ret As Integer
  
    On Error GoTo errhandler
    
    sProviderIEN = gsUserNCID
    sSQL = ""
  
    sSQL = "Select count(*) from result_event_save where result_event_save.clinician_ncid =  " & sProviderIEN

    Set objSQL = SelectCHCSII_DAS
    Set objRecset = New ADODB.Recordset
    Set objRecset = objSQL.OpenRecordset(sSQL)
    
    'get number of records in saved table
     gLTotalSavedResults = CLng(objRecset(0))
    GoTo CleanUp

errhandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "in New Results.GetNumberOfSavedResults", "New Results", vbExclamation)
CleanUp:
    If Not objSQL Is Nothing Then Set objSQL = Nothing
    If Not objRecset Is Nothing Then Set objRecset = Nothing
End Sub

'*******************************************************************************************
'  Function Name: ICHCSII_CoreAPI3_CloseView2
'
'  Description: This is aprt of the extended core API. A
'               It checks to see, if the class is busy doing a data refresh, if so sends
'                back a negative response
'
'  Parameters:  viViewType As Integer
'
'  HISTORY:
'  MBorsody - Added Function to address problem in tracker 18604
'
'*******************************************************************************************

Private Function ICHCSII_CoreAPI3_CloseView2(ByVal viViewType As Integer) As Boolean

Dim bCanClose As Boolean
On Error GoTo errhandler
' tracker 18604 added this so can tell core  not okay to close when in process of refreshing data
' core changes may make this obsolete... was having a timing problem.  Time is now locking itself for an interval, to circumvent
' problem described in tracker
bCanClose = True

    If viViewType = cwiALL_VIEWS Then
        If Not gbBusy Then
             '---this indicates that CW is closing the "main view", NOT the OLE Server---
             cwbCloseViewMyself = False
            
             If cwbMainViewRunning Then
                Call gfrmResults.RemoveMenus
                Unload gfrmResults
                '---release memory reference to form level variables---
                Set gfrmResults = Nothing
             End If
            
             '-- set a global reference for the class, set it to nothing, initially set in resizeview of class
             '-- free reference so can fire classes terminate event... and clean up
             Set gObjMe = Nothing
    
             '---set global flags---
             cwbMainViewRunning = False
             cwbPatientSummaryViewRunning = False
        Else
                bCanClose = False
        End If
     ElseIf viViewType = cwiREGULAR_VIEW Then
        If Not gbBusy Then
            '---this indicates that CW is closing the "main view", NOT the OLE Server---
            cwbCloseViewMyself = False

           If cwbMainViewRunning Then
                '---this indicates that CW is closing the "main view", NOT the OLE Server---
                Call gfrmResults.RemoveMenus
                Unload gfrmResults
                '---release memory reference to form level variables---
                Set gfrmResults = Nothing
             End If
       
             '-- set a global reference for the class, set it to nothing, initially set in resizeview of class
             '-- free reference so can fire classes terminate event... and clean up
             Set gObjMe = Nothing
    
             '---set global flag---
             cwbMainViewRunning = False
        Else
                bCanClose = False
        End If
    ElseIf viViewType = cwiPATIENT_SUMMARY_VIEW Then
          ' no action
    End If
 
ICHCSII_CoreAPI3_CloseView2 = bCanClose
Exit Function
errhandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "in New Results.ICHCSII_CoreAPI3_CloseView2", _
                                "New Results", vbExclamation)
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

Call ICHCSII_CoreAPI3_CloseView2(viViewType)

End Sub

'*******************************************************************************************
'  Function Name(): RNR_StoredProc
'
'  Description:
'
'  Input Parameters:
'
'
'  Output Parameters:
'               Possible Return statuses from the execution of this stored procedure
'                 1 - Successfully Completed
'                 2 - Received New Data Requires Viewing
'                 3 - Record Not Found
'                 4 - Record in Use by Another User
'                 5 - Error in Stored Procedure
'                 6 - Record already exist
'
'---- Begin:SCR #28730;   Developer: MAG 11/01/2002 03:10 PM
'---- Note: Function Changed:Pushed code to logically lock record and effect user RNR action onto database.
'---- Utilizing 3M's SQLServ.dll  we invoke a stored procedure.
'---- The store procedure's return value is used to apprise application's user of the outcome.
'---- For example if the record on the database has had a state change from what the application user has viewed...
'---- we inform them they have new information that requires their viewing.
'---- Impact of Change:Eliminates many of the concurrency issues unique to the new results module.
'---- End: SCR #28730;
'
'*******************************************************************************************
Private Function RNR_StoredProc(ByVal storedProcName As String, ByVal sOrder_Number As String, ByVal sOrder_Task_No As String, ByVal sProviderIEN As String, ByVal strGridTestIDs As String, ByVal strGridTestDates As String, ByVal strList As String, ByVal strAction As String, Optional fowardingProvNCID As String) As String
   
   ' new objects
    Dim oSqlConnect As SQLSERVATLLib.SqlServConnect
    Dim oStatement As SQLSERVATLLib.SqlServStatement
    Dim oBindVar As SQLSERVATLLib.SqlServBindVar
    Dim i As Integer
    Dim sOutcome As String
    
    On Error GoTo errhandler
       
       Set oSqlConnect = New SQLSERVATLLib.SqlServConnect

       oSqlConnect.ClearAll
       Set oStatement = oSqlConnect.Transaction.Add
       oStatement.ProcedureName = storedProcName
       
       Set oBindVar = oStatement.BindVariables.Add
       oBindVar.IsCursor = False
       oBindVar.IsOutput = False
       oBindVar.Value.InfoType = "textString"
       oBindVar.Value.TextString = sOrder_Task_No

       Set oBindVar = oStatement.BindVariables.Add
       oBindVar.IsCursor = False
       oBindVar.IsOutput = False
       oBindVar.Value.InfoType = "textString"
       oBindVar.Value.TextString = sOrder_Number
       
       Set oBindVar = oStatement.BindVariables.Add
       oBindVar.IsCursor = False
       oBindVar.IsOutput = False
       oBindVar.Value.InfoType = "textString"
       oBindVar.Value.TextString = sProviderIEN
       
       If strAction <> gHL7MSGforward Then
            Set oBindVar = oStatement.BindVariables.Add
            oBindVar.IsCursor = False
            oBindVar.IsOutput = False
            oBindVar.Value.InfoType = "textString"
            oBindVar.Value.TextString = strGridTestIDs
       End If
       
       If strAction <> gHL7MSGforward Then
            Set oBindVar = oStatement.BindVariables.Add
            oBindVar.IsCursor = False
            oBindVar.IsOutput = False
            oBindVar.Value.InfoType = "textString"
            oBindVar.Value.TextString = strGridTestDates
       End If
       
       Set oBindVar = oStatement.BindVariables.Add
       oBindVar.IsCursor = False
       oBindVar.IsOutput = False
       oBindVar.Value.InfoType = "textString"
       oBindVar.Value.TextString = strList
       
       If strAction <> gHL7MSGforward Then
            Set oBindVar = oStatement.BindVariables.Add
            oBindVar.IsCursor = False
            oBindVar.IsOutput = False
            oBindVar.Value.InfoType = "textString"
            oBindVar.Value.TextString = strAction
       End If
       
       If strAction = gHL7MSGforward Then
            Set oBindVar = oStatement.BindVariables.Add
            oBindVar.IsCursor = False
            oBindVar.IsOutput = False
            oBindVar.Value.InfoType = "textString"
            oBindVar.Value.TextString = fowardingProvNCID
       End If
       
       Set oBindVar = oStatement.BindVariables.Add
       oBindVar.IsCursor = False
       oBindVar.IsOutput = True
       oBindVar.Value.InfoType = "textString"
       oBindVar.Value.TextString = "123456"
       
       oSqlConnect.ExecuteTransaction
       
       Set oStatement = oSqlConnect.Transaction(1)
       With oStatement.BindVariables
           For i = 1 To .Count
               ' we have a single output return value
               If (.Item(i).IsOutput) Then
                  sOutcome = .Item(i).Value.TextString
               End If
           Next
        End With

       ' free up system resources
       Set oStatement = Nothing
       Set oBindVar = Nothing
      
       RNR_StoredProc = sOutcome

GoTo CleanUp

errhandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "in New Results.RNR_Provider_Action", "New Results", vbExclamation)

CleanUp:
    'free database 3M objects
    If Not oStatement Is Nothing Then Set oStatement = Nothing
    If Not oBindVar Is Nothing Then Set oBindVar = Nothing
    If Not oSqlConnect Is Nothing Then Set oSqlConnect = Nothing

End Function
' End: SCR #22928;

'---- Begin:SCR #24639;   Developer: MAB 06/14/2002 01:26 PM
'---- Note: FunctionChanged:  Added HighlightAbnormalValue funtcion that will take the critical abnormal value and massage it into the correct color for display based on what its value is.  We now retrieve that value for updatain gour grid if its value has been changed in database.      Impact of Change:  We are in a save and toss , when the user elects not to view the result, setting it back to an unviewed state.      We rebold the row.... but the critical abnormal status may have changed.  The net effect of this is we are correctly pulling in the changed value for a critical_abnormal_value and refreshing the display with that value.    Correcting the update problem.  The grid was correctly doing date comparisons to tell whether or not to proceed with the providre's action.  I did not realize that that when a result is amended, the critical ablnormal value too could be changed, and not just the date of the result.

Private Sub HighlightAbnormalValue(currentRowIndex As Integer)
Dim objGrid As VSFlexGrid
Dim i As Integer
Dim element
Dim strCompareValue As String
Dim blnFound As Boolean
    i = currentRowIndex
    'determine which grid is active for highlighting cell
    If gbIsNewResultsList Then
      Set objGrid = gfrmResults.grdResults
    Else
      Set objGrid = gfrmResults.grdSavedResults
    End If
    'display critical labs and rads with color
       ' one of three lab types
        '---- Begin:SCR #25135;   Developer: MAG 07/01/2002 10:42 AM
        '---- Note: Function Changed: Changed label to generic label for lab chems.      Impact of Change: We now do not assign a weighted significance by showing a high or low attacched tO the label see accompanying note for additional information.
        '---- Begin:SCR #  Developer: MAG 07/01/2002 11:12 AM
        '---- Note: Function Changed:Genericized critical abnormal values for lab chems, and removed erroneous normal value inserted when lab chem had no abnormal value.      Impact of Change:  Generic labels have no assignment of hierarchy attesting to result severity.  (no order of precedence, which we cannot determine given the information we have on hand).  see attached not e for enlargded explantaion.
        If objGrid.TextMatrix(i, rnrCOLresult_type) <> "RAD" Then
            If Len(objGrid.TextMatrix(i, rnrCOLcritical_abnormal)) > 0 Then
                Select Case UCase(objGrid.TextMatrix(i, rnrCOLcritical_abnormal))
                Case "CRITICAL", "ABNORMAL"
                    objGrid.Cell(flexcpForeColor, i, rnrCOLcritical_abnormal) = RGB(255, 0, 0) 'vbRed
                 Case Else
                    objGrid.Cell(flexcpForeColor, i, rnrCOLcritical_abnormal) = RGB(0, 0, 0) 'vbBlack
                End Select
            End If
        Else 'is a rad
             strCompareValue = UCase(objGrid.TextMatrix(i, rnrCOLcritical_abnormal))
             strCompareValue = Trim(strCompareValue)
             For Each element In gcolRadAbnormalValues
               ' if reports critical value is amongst values requiring physician be notified it is
               ' put into a critical color(red)
                If strCompareValue = Trim(UCase(element)) Then
                   objGrid.Cell(flexcpForeColor, i, rnrCOLcritical_abnormal) = RGB(255, 0, 0) 'vbRed
                   blnFound = True
                End If
            Next
            ' we may want to set the font back to black cannot be sure if it was red
            If Not blnFound Then
              objGrid.Cell(flexcpForeColor, i, rnrCOLcritical_abnormal) = RGB(0, 0, 0) 'vbBlack
            End If
        '---- End: SCR #25135;
        '---- End: SCR #
        End If
End Sub
' End: SCR #24639;

Public Sub Gems_SQLSrv_ForwardResult()

    Dim sSQL As String, sProviderIEN As String
    Dim bCancel As Boolean
    Dim sFacility_Ncid As String, sOrder_Number As String
    Dim Index As Integer
    Dim colRowsToForward As Collection
    Dim objSQL As ICHCSII_SQL.ISqlOps
    Dim objSQLDAS As ICHCSII_DAS
    Dim element
    Dim sProviderName As String
    Dim objGrid As VSFlexGrid
    Dim sTableName  As String
    Dim sSendingList As String
    Dim sFwdProviderIEN As String
    Dim intCurrentRowIndex As Integer
    Dim sOrder_Task_No As String
    Dim iRecCount As Integer
    Dim iRecCountSave As Integer
    Dim objRecset As ADODB.Recordset
    Dim objRecsetCntSave As ADODB.Recordset
    Dim blnConfirmForward As Boolean
    'used in insert statement
    Dim strCOLResultType As String
    Dim strCOLpriority As String
    Dim strCOLcritical_abnormal As String
    Dim strCOLunit_number As String
    Dim strCOLfacility_ncid As String
    Dim strCOLresults_ids As String
    Dim strCOLaccess_exam_no As String
    Dim strCOLmmtf_code As String
    Dim strCOLcertified_dates As String
    Dim sSQLInsert As String
    Dim sSQLValues As String

  
    Set objSQLDAS = SelectCHCSII_DAS
    Set objSQL = SelectSQL_ISqlOps
    Set objRecset = New ADODB.Recordset
    Set objRecsetCntSave = New ADODB.Recordset
    
    On Error GoTo errhandler
    
        sProviderIEN = gsUserNCID
        sSQL = ""
        sProviderName = ""
        sFwdProviderIEN = ""
        gfrmResults.gDisableChange = True
    
        sProviderName = gfrmResults.txtProvider.Text
        'get variables used to pivot against which list is being worked upon
        If gbIsNewResultsList Then
          Set objGrid = gfrmResults.grdResults
          sTableName = "RESULT_EVENT"
        Else
          Set objGrid = gfrmResults.grdSavedResults
          sTableName = "RESULT_EVENT_SAVE"
        End If
        
        ' make sure doctor is selected... get their name... cannot forward without it
        If sProviderName = "" Then
           MsgBox ("A provider must be selected to designate who you would like to receive the Result.")
        Else
             ' get collection of selected indices
             Set colRowsToForward = gfrmResults.SelectedRowsIndex
        
             ' iterate through and ask if want to individually send result
             For Each element In colRowsToForward
                 'set working row
                 intCurrentRowIndex = CInt(element)
                 objGrid.Row = intCurrentRowIndex
                
                 If Not blnConfirmForward Then
                     If MsgBox("Do You Want To forward this result to provider: " & sProviderName & "?", vbYesNo) = vbNo Then
                        bCancel = True
                    Else
                       blnConfirmForward = True
                    End If
                 End If
                
                '--- Run forward function in CDR+
                If Not bCancel Then
  
                    sFwdProviderIEN = Trim(gfrmResults.ForwardingProvider)
                    sOrder_Number = objGrid.Cell(flexcpText, intCurrentRowIndex, rnrCOLorder_number)
                    sOrder_Task_No = objGrid.Cell(flexcpText, intCurrentRowIndex, rnrCOLorder_task_no)
                    
                   ' since an order can get into a result list by dint of the fact that another
                   ' provider has transferred it there... we need to make sure DNE there already
                   sSQL = "Select count(*) from RESULT_EVENT WHERE " & _
                          "CLINICIAN_NCID = " & sFwdProviderIEN & " AND ORDER_NUMBER = '" & _
                           sOrder_Number & "' AND ORDER_TASK_NO = " & sOrder_Task_No
                           
                   Set objRecset = objSQLDAS.OpenRecordset(sSQL)
                   iRecCount = objRecset.Fields(0).Value
                   objRecset.Close
                   
                   ' may have an earlier copy and saved it so check saved list
                   sSQL = "Select count(*) from RESULT_EVENT_SAVE WHERE " & _
                          "CLINICIAN_NCID = " & sFwdProviderIEN & " AND ORDER_NUMBER = '" & _
                           sOrder_Number & "' AND ORDER_TASK_NO = " & sOrder_Task_No
                           
                    Set objRecsetCntSave = objSQLDAS.OpenRecordset(sSQL)
                    iRecCountSave = objRecsetCntSave.Fields(0).Value
                    objRecsetCntSave.Close
                    If iRecCountSave = 1 Or iRecCount = 1 Then
                        MsgBox "The provider already has the result in one of their result lists."
                    Else
                       If gbSQLServer Then
                            ' manually build insert of row  in SQL server, not handled by triggers
                            ' we have all the information for the insert in the grid so just build insert statement
                            strCOLResultType = objGrid.Cell(flexcpText, intCurrentRowIndex, rnrCOLresult_type)
                            strCOLpriority = objGrid.Cell(flexcpText, intCurrentRowIndex, rnrCOLpriority)
                            strCOLcritical_abnormal = objGrid.Cell(flexcpText, intCurrentRowIndex, rnrCOLcritical_abnormal)
                            If strCOLcritical_abnormal = "" Then
                                strCOLcritical_abnormal = "NULL"
                            End If
                            strCOLunit_number = objGrid.Cell(flexcpText, intCurrentRowIndex, rnrCOLunit_number)
                            strCOLfacility_ncid = objGrid.Cell(flexcpText, intCurrentRowIndex, rnrCOLfacility_ncid)
                            strCOLresults_ids = objGrid.Cell(flexcpText, intCurrentRowIndex, rnrCOLresults_ids)
                            strCOLaccess_exam_no = objGrid.Cell(flexcpText, intCurrentRowIndex, rnrCOLaccess_exam_no)
                            strCOLmmtf_code = objGrid.Cell(flexcpText, intCurrentRowIndex, rnrCOLmmtf_code)
                            strCOLcertified_dates = objGrid.Cell(flexcpText, intCurrentRowIndex, rnrCOLcertified_dates)
   
                           'business rule forwards are only ever inserted into the provider's New List even if it originates off a provider's saved list.
                           sSQLInsert = "INSERT INTO RESULT_EVENT ( FACILITY_NCID, ORDER_NUMBER, EVENT_ID, CLINICIAN_NCID, " & _
                                        " PATIENT_UNIT_NUMBER, RESULT_TYPE, CRITICAL_ABNORMAL_VALUE, NEW_RESULTS_FLAG, PRIORITY, MTF_CODE, " & _
                                        " ORDER_TASK_NO, STATUS, GENERATE_HL7, FWD_CLINICIAN_NCID, RESULT_IDS, FWD_FLAG, " & _
                                        " FILLER_DOCUMENT_NUMBER, IN_USE_FLAG, CERTIFIED_DATES ) VALUES ( "
    
                           sSQLValues = strCOLfacility_ncid & ",'" & sOrder_Number & "', NULL, " & sFwdProviderIEN & ", " & _
                                     strCOLunit_number & ", '" & strCOLResultType & "', '" & strCOLcritical_abnormal & "', 'Y', '" & _
                                     strCOLpriority & "', '" & strCOLmmtf_code & "', " & sOrder_Task_No & ",NULL, 1 ,NULL, '" & _
                                     strCOLresults_ids & "',1, '" & strCOLaccess_exam_no & "',0, '" & strCOLcertified_dates & "')"
                                     
                           sSQL = sSQLInsert & sSQLValues
                       Else
                           ' triggers on update will create new record for forwarding provider
                           sSQL = "UPDATE " & sTableName & " set STATUS = '" & gHL7MSGforward & "', " & _
                                   "FWD_CLINICIAN_NCID = " & sFwdProviderIEN & ", GENERATE_HL7 = " & rnrMSGoutbound & ", IN_USE_FLAG = 0  WHERE " & _
                                   "CLINICIAN_NCID = " & sProviderIEN & " AND ORDER_NUMBER = '" & _
                                   sOrder_Number & "' AND ORDER_TASK_NO = " & sOrder_Task_No
                       End If
                       Call objSQL.Execute(sSQL)
                   End If
                End If
            Next
           
           ' now set the provider input box and
           ' the local variable that stores that tab's provider state back to empty
           If gbIsNewResultsList Then
              gfrmResults.NewListProvider = ""
           Else
             gfrmResults.SavedListProvider = ""
           End If
           gfrmResults.txtProvider.Text = ""
        End If
    
GoTo CleanUp

errhandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "in New Results.Gems_SQLSrv_ForwardResult", "New Results", vbExclamation)
CleanUp:
    If Not objSQL Is Nothing Then Set objSQL = Nothing
    If Not colRowsToForward Is Nothing Then Set colRowsToForward = Nothing
    If Not objGrid Is Nothing Then Set objGrid = Nothing
    gfrmResults.gDisableChange = False
End Sub

'*******************************************************************************************
'  Sub Name: Gems_SQLSvr_DiscardResult
'
'  Description: This method is called when the user has opted to delete a result from
'               their new/save results list without signing it.
'
'  Parameters:
'
'  HISTORY:
'  MGumbert - Created 7/2002
'---- Begin:SCR #46294;   Developer: MAG 02/12/2004
'---- Note: Function Changed:Changed functions that return return Chcsii or Gems object to use CWShared.appmode to return the appropriate object wrapped by CHSII_Conn.
'---- Impact of Change: This change was needed to accomadate the transition over to CHCSII_Conn; CHCSII_DAS and CHCSII_SQL DLLs are going away.
'*******************************************************************************************
Public Sub Gems_SQLSvr_DiscardResult()

    Dim sSQL As String, sProviderIEN As String
    Dim bCancel As Boolean
    Dim sOrder_Task_No As String, sOrder_Number As String
    Dim Index As Integer
    Dim colRowsToDiscard As Collection
    Dim element
    Dim intCount As Integer
    Dim objGrid As VSFlexGrid
    Dim sTableName As String
    Dim intCurrentRowIndex As Integer
    Dim blnConfirmDiscard As Boolean
    Dim objSQL As ICHCSII_SQL.ISqlOps
    Dim strSQLAction As String

    On Error GoTo errhandler
    
    'get variables used in call using grid
    If gbIsNewResultsList Then
      Set objGrid = gfrmResults.grdResults
      sTableName = "Result_event"
    Else
      Set objGrid = gfrmResults.grdSavedResults
      sTableName = "Result_event_save"
    End If

    Set objSQL = SelectSQL_ISqlOps
    sProviderIEN = gsUserNCID
    sSQL = ""
    If gbGEMS Then
      strSQLAction = "UPDATE " & sTableName & " set STATUS = '" & gHL7MSGdiscard & "', GENERATE_HL7 = " & rnrMSGoutbound
    Else
      strSQLAction = "Delete from " & sTableName
    End If
      
    gfrmResults.gDisableChange = True
    ' get list of selected indices
    Set colRowsToDiscard = gfrmResults.SelectedRowsIndex
    If colRowsToDiscard.Count > 0 Then
        intCount = 0
         
         blnConfirmDiscard = False
         For Each element In colRowsToDiscard
            intCurrentRowIndex = CInt(element - intCount)
            objGrid.Row = intCurrentRowIndex
                    
            If Not blnConfirmDiscard Then
                If MsgBox("Do you want to discard this result(s)?", vbYesNo) = vbNo Then
                    bCancel = True
                Else
                   blnConfirmDiscard = True
                End If
            End If

            If Not bCancel Then
                sOrder_Number = objGrid.Cell(flexcpText, intCurrentRowIndex, rnrCOLorder_number)
                sOrder_Task_No = objGrid.Cell(flexcpText, intCurrentRowIndex, rnrCOLorder_task_no)
                Debug.Print "Discard Result #" & objGrid.Row & " order task no. " & sOrder_Task_No
    
                ' use PK of record to insure uniquely identified
                sSQL = strSQLAction & "  WHERE " & _
                       "CLINICIAN_NCID = " & sProviderIEN & " AND ORDER_NUMBER = '" & _
                        sOrder_Number & "' AND ORDER_TASK_NO = " & sOrder_Task_No
                        
                Call objSQL.Execute(sSQL)
              
                ' now hide and remove row(s) from grid that is up and being viewed
                objGrid.RowHidden(intCurrentRowIndex) = False
                objGrid.RemoveItem (intCurrentRowIndex)
                
                ' if saved list is being discarded from, decrement global counter
                If Not gbIsNewResultsList Then
                    gLTotalSavedResults = gLTotalSavedResults - 1
                End If
                ' this is the counter used to track the number of previous deletes, used to adjust row index
                intCount = intCount + 1
            End If
        Next
        
        'refresh once the twos have been deleted
        objGrid.Refresh
        'inactivate action menu
        If objGrid.Rows = 1 Then
           'just have grid header
            gfrmResults.DeactivateMenus
            ' disable saved list tab... so next time after clicking off of it
            ' it will be disabled.  Because it is now empty.
            If Not gbIsNewResultsList Then
                gfrmResults.tabResults.TabEnabled(rnrSavedTab) = False
            End If
        End If
  
    Else
       ' ask user to select a row to take desired action upon
       MsgBox "Please select a result to take the action upon."
    End If
        
GoTo CleanUp

errhandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "in New Results.Gems_SQLSvr_DiscardResult", "New Results", vbExclamation)
CleanUp:
    If Not objSQL Is Nothing Then Set objSQL = Nothing
    If Not colRowsToDiscard Is Nothing Then Set colRowsToDiscard = Nothing
    If Not objGrid Is Nothing Then Set objGrid = Nothing
    gfrmResults.gDisableChange = False
End Sub


'*******************************************************************************************
'  Sub Name: Gems_SQLSvr_NewResults
'
'  Description: This method is called when the user has opted to delete a result from
'               their saved results list which has previously been signed, and
'               move it back to the new list.
'
'  Parameters:
'
'  HISTORY:
'  MGumbert - Created 7/2002
'---- Begin:SCR #46294;   Developer: MAG 02/12/2004
'---- Note: Function Changed:Changed functions that return return Chcsii or Gems object to use CWShared.appmode to return the appropriate object wrapped by CHSII_Conn.
'---- Impact of Change: This change was needed to accomadate the transition over to CHCSII_Conn; CHCSII_DAS and CHCSII_SQL DLLs are going away.
'*******************************************************************************************

Public Sub Gems_SQLSvr_NewResult()

    Dim sSQLUpdate As String, sProviderNCID As String
    Dim bCancel As Boolean
    Dim sOrder_Task_No As String, sOrder_Number As String
    Dim Index As Integer
    Dim Count As Long
    Dim colRowsSelected As Collection
    Dim element
    Dim intCurrentRowIndex As Integer
    Dim objSQL As ICHCSII_SQL.ISqlOps
    Dim blnConfirmNew As Boolean
    'used in insert statement
    Dim strCOLResultType As String
    Dim strCOLpriority As String
    Dim strCOLcritical_abnormal As String
    Dim strCOLunit_number As String
    Dim strCOLfacility_ncid As String
    Dim strCOLresults_ids As String
    Dim strCOLaccess_exam_no As String
    Dim strCOLmmtf_code As String
    Dim strCOLcertified_dates As String
    Dim sSQLInsert As String
    Dim sSQLValues As String
    Dim sSQLDelete As String
    Dim strCOLresult_new As String
    
    On Error GoTo errhandler


    Set objSQL = SelectSQL_ISqlOps
    sProviderNCID = gsUserNCID

    
    gfrmResults.gDisableChange = True
    ' get collection of selected rows
    Set colRowsSelected = gfrmResults.SelectedRowsIndex
    Count = 0
    If colRowsSelected.Count > 0 Then

         'ask for confirmation, and then move back to save d list
         For Each element In colRowsSelected
            intCurrentRowIndex = (element - Count)
            gfrmResults.grdSavedResults.Row = intCurrentRowIndex
             If Not blnConfirmNew Then
                If MsgBox("Do you want to move this result(s) back to your new result list?", vbYesNo) = vbNo Then
                    bCancel = True
                Else
                    blnConfirmNew = True
                End If
            End If
            
            '--- Run delete query to CDR+
            If Not bCancel Then
  
                ' Get parts of PK for record cause can have multiple results for an order number
                sOrder_Number = gfrmResults.grdSavedResults.Cell(flexcpText, intCurrentRowIndex, rnrCOLorder_number)
                sOrder_Task_No = gfrmResults.grdSavedResults.Cell(flexcpText, intCurrentRowIndex, rnrCOLorder_task_no)
                
                If gbSQLServer Then
                    ' manually build insert of row  in SQL server, not handles by triggers
                    strCOLResultType = gfrmResults.grdSavedResults.Cell(flexcpText, intCurrentRowIndex, rnrCOLresult_type)
                    strCOLpriority = gfrmResults.grdSavedResults.Cell(flexcpText, intCurrentRowIndex, rnrCOLpriority)
                    strCOLcritical_abnormal = gfrmResults.grdSavedResults.Cell(flexcpText, intCurrentRowIndex, rnrCOLcritical_abnormal)
                    
                    strCOLunit_number = gfrmResults.grdSavedResults.Cell(flexcpText, intCurrentRowIndex, rnrCOLunit_number)
                    strCOLfacility_ncid = gfrmResults.grdSavedResults.Cell(flexcpText, intCurrentRowIndex, rnrCOLfacility_ncid)
                    strCOLresults_ids = gfrmResults.grdSavedResults.Cell(flexcpText, intCurrentRowIndex, rnrCOLresults_ids)
                    strCOLaccess_exam_no = gfrmResults.grdSavedResults.Cell(flexcpText, intCurrentRowIndex, rnrCOLaccess_exam_no)
                    strCOLmmtf_code = gfrmResults.grdSavedResults.Cell(flexcpText, intCurrentRowIndex, rnrCOLmmtf_code)
                    strCOLcertified_dates = gfrmResults.grdSavedResults.Cell(flexcpText, intCurrentRowIndex, rnrCOLcertified_dates)
                    strCOLresult_new = gfrmResults.grdSavedResults.Cell(flexcpText, intCurrentRowIndex, rnrCOLresult_new)

                    sSQLInsert = "INSERT INTO RESULT_EVENT ( FACILITY_NCID, ORDER_NUMBER, EVENT_ID, CLINICIAN_NCID, " & _
                                 " PATIENT_UNIT_NUMBER, RESULT_TYPE, CRITICAL_ABNORMAL_VALUE, NEW_RESULTS_FLAG, PRIORITY, MTF_CODE, " & _
                                 " ORDER_TASK_NO, STATUS, GENERATE_HL7, FWD_CLINICIAN_NCID, RESULT_IDS, FWD_FLAG, " & _
                                 " FILLER_DOCUMENT_NUMBER, IN_USE_FLAG, CERTIFIED_DATES ) VALUES ( "

                    sSQLValues = strCOLfacility_ncid & ",'" & sOrder_Number & "', NULL, " & sProviderNCID & ", " & _
                                 strCOLunit_number & ", '" & strCOLResultType & "', '" & strCOLcritical_abnormal & "', '" & strCOLresult_new & "', '" & _
                                 strCOLpriority & "', '" & strCOLmmtf_code & "', " & sOrder_Task_No & ",NULL, 1 ,NULL, '" & _
                                 strCOLresults_ids & "',1, '" & strCOLaccess_exam_no & "',0, '" & strCOLcertified_dates & "')"
                
                           
                     sSQLInsert = sSQLInsert & sSQLValues
                     objSQL.Execute (sSQLInsert)
                     'Debug.Print sSQLInsert
                                 
                     ' use PK of record to insure uniquely identified, deleting old record
                     sSQLDelete = "Delete RESULT_EVENT_SAVE WHERE CLINICIAN_NCID = " & sProviderNCID & " AND ORDER_NUMBER = '" & _
                                   sOrder_Number & "' AND ORDER_TASK_NO = " & sOrder_Task_No
                                   
                     'Debug.Print sSQLDelete
                      objSQL.Execute (sSQLDelete)

                Else
                  ' use PK of record to insure uniquely identified releies on triggers to delete and
                  ' move record into saved list table
                   sSQLUpdate = "UPDATE RESULT_EVENT_SAVE set STATUS = '" & gHL7MSGnew & "', GENERATE_HL7 = " & rnrMSGoutbound & "  WHERE " & _
                                "CLINICIAN_NCID = " & sProviderNCID & " AND ORDER_NUMBER = '" & _
                                sOrder_Number & "' AND ORDER_TASK_NO = " & sOrder_Task_No
            
                  Debug.Print sSQLUpdate
                  Call objSQL.Execute(sSQLUpdate)
                End If
                ' this is the counter used to adjust the index of the next row
                Count = Count + 1
                
                ' subtract from global variable that helps regulate action menu
                gLTotalSavedResults = gLTotalSavedResults - 1
                
                ' now hide and remove row for processed action
                gfrmResults.grdSavedResults.RowHidden(intCurrentRowIndex) = False
                gfrmResults.grdSavedResults.RemoveItem (intCurrentRowIndex)
            End If
        Next
        'refresh grid after removind all the rows
        gfrmResults.grdSavedResults.Refresh
        'test to see if we need to inactivate action menu
        If gfrmResults.grdSavedResults.Rows = 1 Then
           ' have no data just header
            gfrmResults.DeactivateMenus
        End If
   Else
       ' ask user to select a row to take desired action upon
       MsgBox "Please select a result to take the action upon."
   End If
    
  GoTo CleanUp

errhandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "in New Results.Gems_SQLSvr_NewResult", "New Results", vbExclamation)
CleanUp:
    If Not objSQL Is Nothing Then Set objSQL = Nothing
    If Not colRowsSelected Is Nothing Then Set colRowsSelected = Nothing
    gfrmResults.gDisableChange = False
End Sub


'*******************************************************************************************
'  Sub Name: Gems_SQLSvr_SaveResult
'
'  Description: This method is called when the user has opted to sign and remove a new result
'               from their New Results list and move it to their saved list... does an update on effected record,
'               that causes a trigger to generate an HL7 msg, and it writes a signature record into
'               the signature table.  As well as delete the effected record from New Results table.
'               Working assumption is that a user must first view a result before they can save it.
'
'  Parameters:
'
'  HISTORY:
'  MGumbert - Created 7/2002
'---- Begin:SCR #46294;   Developer: MAG 02/12/2004
'---- Note: Function Changed:Changed functions that return return Chcsii or Gems object to use CWShared.appmode to return the appropriate object wrapped by CHSII_Conn.
'---- Impact of Change: This change was needed to accomadate the transition over to CHCSII_Conn; CHCSII_DAS and CHCSII_SQL DLLs are going away.
'*******************************************************************************************
Public Sub Gems_SQLSvr_SaveResult()

    Dim sSQL As String, sProviderNCID As String
    Dim bCancel As Boolean
    Dim sOrder_Task_No As String, sOrder_Number As String
    Dim Index As Integer
    Dim Count As Long
    Dim colSelectedRows As New Collection
    Dim element
    Dim objSQL As ICHCSII_SQL.ISqlOps
    Dim intCurrentRowIndex As Integer
    Dim blnConfirmSave As Boolean
    'used in insert statement
    Dim strCOLResultType As String
    Dim strCOLpriority As String
    Dim strCOLcritical_abnormal As String
    Dim strCOLunit_number As String
    Dim strCOLfacility_ncid As String
    Dim strCOLresults_ids As String
    Dim strCOLaccess_exam_no As String
    Dim strCOLmmtf_code As String
    Dim strCOLcertified_dates As String
    Dim strCOLresult_new As String
    Dim sSQLInsert As String
    Dim sSQLValues As String
    Dim sSQLDelete As String
    Dim colSQLStatements As Collection

    On Error GoTo errhandler
      
    Set objSQL = SelectSQL_ISqlOps
    sProviderNCID = gsUserNCID
    sSQL = ""
    
    '--- Run update SQL
    gfrmResults.gDisableChange = True
    
    ' get rowIndexes for selected rows, and do test to make sure the individual rows have ben viewed
    ' working premis here is that they must be viewed before they can be saved
    For Index = 1 To gfrmResults.grdResults.Rows - 1
        If gfrmResults.grdResults.IsSelected(Index) Then
            If gfrmResults.grdResults.Cell(flexcpText, Index, rnrCOLresult_new) = "Y" Then
              MsgBox ("You have selected a result that has not been viewed yet." & vbCr & vbCr & "You cannot Save or Toss a result without first viewing it.")
              Exit Sub
            End If
            colSelectedRows.Add Index
        End If
    Next Index

    Count = 0
    
    If colSelectedRows.Count > 0 Then

        For Each element In colSelectedRows
            intCurrentRowIndex = (element - Count)
                    
            If Not blnConfirmSave Then
                If MsgBox("Do you want to move this result(s) to your saved result list?", vbYesNo) = vbNo Then
                    bCancel = True
                Else
                    blnConfirmSave = True
                End If
            End If
                    
            If Not bCancel Then
                            
                  ' Get parts of PK for record cause can have multiple results for an order number
                  sOrder_Number = gfrmResults.grdResults.Cell(flexcpText, intCurrentRowIndex, rnrCOLorder_number)
                  sOrder_Task_No = gfrmResults.grdResults.Cell(flexcpText, intCurrentRowIndex, rnrCOLorder_task_no)
                  If gbSQLServer Then
                       ' manually build insert of row  in SQL server, not handled by triggers
                       ' we have no need to genrate mesages in training so net effect of provider's
                       ' action is that we copy record to the saved list and delte the record f the
                        strCOLResultType = gfrmResults.grdResults.Cell(flexcpText, intCurrentRowIndex, rnrCOLresult_type)
                        strCOLpriority = gfrmResults.grdResults.Cell(flexcpText, intCurrentRowIndex, rnrCOLpriority)
                        strCOLcritical_abnormal = gfrmResults.grdResults.Cell(flexcpText, intCurrentRowIndex, rnrCOLcritical_abnormal)
                        
                        strCOLunit_number = gfrmResults.grdResults.Cell(flexcpText, intCurrentRowIndex, rnrCOLunit_number)
                        strCOLfacility_ncid = gfrmResults.grdResults.Cell(flexcpText, intCurrentRowIndex, rnrCOLfacility_ncid)
                        strCOLresults_ids = gfrmResults.grdResults.Cell(flexcpText, intCurrentRowIndex, rnrCOLresults_ids)
                        strCOLaccess_exam_no = gfrmResults.grdResults.Cell(flexcpText, intCurrentRowIndex, rnrCOLaccess_exam_no)
                        strCOLmmtf_code = gfrmResults.grdResults.Cell(flexcpText, intCurrentRowIndex, rnrCOLmmtf_code)
                        strCOLcertified_dates = gfrmResults.grdResults.Cell(flexcpText, intCurrentRowIndex, rnrCOLcertified_dates)
                        strCOLresult_new = gfrmResults.grdResults.Cell(flexcpText, intCurrentRowIndex, rnrCOLresult_new)

                        'business rule forwards are only ever inserted into the provider's New List even if it originates off a provider's saved list.
                        sSQLInsert = "INSERT INTO RESULT_EVENT_SAVE ( FACILITY_NCID, ORDER_NUMBER, EVENT_ID, CLINICIAN_NCID, " & _
                                     " PATIENT_UNIT_NUMBER, RESULT_TYPE, CRITICAL_ABNORMAL_VALUE, NEW_RESULTS_FLAG, PRIORITY, MTF_CODE, " & _
                                     " ORDER_TASK_NO, STATUS, GENERATE_HL7, FWD_CLINICIAN_NCID, RESULT_IDS, FWD_FLAG, " & _
                                     " FILLER_DOCUMENT_NUMBER, IN_USE_FLAG, CERTIFIED_DATES ) VALUES ( "
    
                        sSQLValues = strCOLfacility_ncid & ",'" & sOrder_Number & "', NULL, " & sProviderNCID & ", " & _
                                     strCOLunit_number & ", '" & strCOLResultType & "', '" & strCOLcritical_abnormal & "', '" & strCOLresult_new & "', '" & _
                                     strCOLpriority & "', '" & strCOLmmtf_code & "', " & sOrder_Task_No & ",NULL, 1 ,NULL, '" & _
                                     strCOLresults_ids & "',1, '" & strCOLaccess_exam_no & "',0, '" & strCOLcertified_dates & "')"
                    
                         sSQLInsert = sSQLInsert & sSQLValues
                         objSQL.Execute (sSQLInsert)
                         'Debug.Print sSQLInsert
                                     
                         ' use PK of record to insure uniquely identified
                         sSQLDelete = "Delete RESULT_EVENT WHERE CLINICIAN_NCID = " & sProviderNCID & " AND ORDER_NUMBER = '" & _
                                       sOrder_Number & "' AND ORDER_TASK_NO = " & sOrder_Task_No
                                       
                         'Debug.Print sSQLDelete
                         objSQL.Execute (sSQLDelete)

                  Else
                        ' triggers in oracle database arre used to delte saved message and to copy the record to the
                        ' result_event_saved table.  Also, it invokes package procedure to generate an outbound
                        ' new results synchronization HL7 message
                        ' use PK of record to insure uniquely identified
                        sSQL = "UPDATE RESULT_EVENT set STATUS = '" & gHL7MSGsave & "', GENERATE_HL7 = " & rnrMSGoutbound & "  WHERE " & _
                                "CLINICIAN_NCID = " & sProviderNCID & " AND ORDER_NUMBER = '" & _
                                 sOrder_Number & "' AND ORDER_TASK_NO = " & sOrder_Task_No
                                
                        Debug.Print sSQL
                        Call objSQL.Execute(sSQL)
                  End If
            
                  ' this is the counter we use to tell how many rows we have deleted so we can adjust index
                  Count = Count + 1
                
                  'when adding a saved result we need to increment to global variable that controls the save
                  'tab's display... if no saved results tab is inactive
                  gLTotalSavedResults = gLTotalSavedResults + 1
                
                 ' now hide and remove row for processed action
                 gfrmResults.grdResults.RowHidden(intCurrentRowIndex) = False
                 gfrmResults.grdResults.RemoveItem (intCurrentRowIndex)
         End If
        Next
        
        'refresh grid after removing all the rows
        gfrmResults.grdResults.Refresh
        
        'test to see if we need to activate tab because we now have saved results
        If gLTotalSavedResults > 0 Then
            gfrmResults.tabResults.TabEnabled(rnrSavedTab) = True
        End If
        
        'inactivate action menu if grid that is up empties of results
        If gfrmResults.grdResults.Rows = 1 Then
           ' have no data just header
            gfrmResults.DeactivateMenus
        End If
    Else
       ' ask user to select a row to take desired action upon
       MsgBox "Please select a result to take the action upon."
    End If
    
    GoTo CleanUp

errhandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "in New Results.Gems_SQLSvr_SaveResult", "New Results", vbExclamation)
CleanUp:
    If Not objSQL Is Nothing Then Set objSQL = Nothing
    If Not colSelectedRows Is Nothing Then Set colSelectedRows = Nothing
    gfrmResults.gDisableChange = False
End Sub


'*******************************************************************************************
'  Sub Name: Gems_SQLSvr_TossResult
'
'  Description: This method is called when the user has opted to sign and remove a new result
'               from their New Results list... an update causes a  triggers to
'               generate an HL7 msg, and write a signature record into the signature table.
'               As well as delete the updated record from New Results table.
'               Working assumption is that a user must first view a result before they can toss it.
'  Parameters:
'
'
'  HISTORY:
'  MGumbert - Created 7/2002
'---- Begin:SCR #46294;   Developer: MAG 02/12/2004
'---- Note: Function Changed:Changed functions that return return Chcsii or Gems object to use CWShared.appmode to return the appropriate object wrapped by CHSII_Conn.
'---- Impact of Change: This change was needed to accomadate the transition over to CHCSII_Conn; CHCSII_DAS and CHCSII_SQL DLLs are going away.
'*******************************************************************************************
Public Sub Gems_SQLSvr_TossResult()
  
    Dim sSQL As String, sProviderIEN As String
    Dim bCancel As Boolean
    Dim sOrder_Task_No As String, sOrder_Number As String
    Dim Index As Integer
    Dim Count As Long
    Dim colRowsToToss As New Collection
    Dim element
    Dim objGrid As VSFlexGrid
    Dim sTableName As String
    Dim intCurrentRowIndex As Integer
    Dim objSQL As ICHCSII_SQL.ISqlOps
    Dim blnConfirmToss As Boolean
    Dim strSQLAction As String

    On Error GoTo errhandler
      
    Set objSQL = SelectSQL_ISqlOps
    
    'get variables used in calls to access correct grid/result
    If gbIsNewResultsList Then
      Set objGrid = gfrmResults.grdResults
      sTableName = "Result_Event"
    Else
      Set objGrid = gfrmResults.grdSavedResults
      sTableName = "Result_Event_Save"
    End If


    sProviderIEN = gsUserNCID
    sSQL = ""
    ' here dependign on which database we are going against
    ' we are either doing an update via oracleetriggers, or a delete on SQLServer
    If gbGEMS Then
      strSQLAction = "UPDATE " & sTableName & " set STATUS = '" & gHL7MSGtoss & "', GENERATE_HL7 = " & rnrMSGoutbound
    Else
      strSQLAction = "Delete from " & sTableName
    End If
      
    
    gfrmResults.gDisableChange = True
    
    Count = 0
    
     ' get rowIndexes for selected rows, and do test to make sure the individual rows have ben viewed
     ' working premis -rule- here is that they must be viewed before they can be tossed
     For Index = 1 To objGrid.Rows - 1
        If objGrid.IsSelected(Index) Then
            If objGrid.Cell(flexcpText, Index, rnrCOLresult_new) = "Y" Then
                  MsgBox ("You have selected a result that has not been viewed yet." & vbCr & vbCr & "You cannot Save or Toss a result without first viewing it.")
                  Exit Sub
            End If
            colRowsToToss.Add Index
        End If
    Next Index
 
   If colRowsToToss.Count > 0 Then
 
        'ask for confirmation, and then toss each individual result
        For Each element In colRowsToToss
            intCurrentRowIndex = CInt(element - Count)
            objGrid.Row = intCurrentRowIndex
            
            If Not blnConfirmToss Then
                If MsgBox("Do you want to Toss this result(s)?", vbYesNo) = vbNo Then
                    bCancel = True
                Else
                    blnConfirmToss = True
                End If
            End If
            
            '--- Run Toss Update in CDR+
            If Not bCancel Then
                
                sOrder_Number = objGrid.Cell(flexcpText, intCurrentRowIndex, rnrCOLorder_number)
                sOrder_Task_No = objGrid.Cell(flexcpText, intCurrentRowIndex, rnrCOLorder_task_no)
                             
               ' use PK of record to insure uniquely identified
                sSQL = strSQLAction & "  WHERE " & _
                       "CLINICIAN_NCID = " & sProviderIEN & " AND ORDER_NUMBER = '" & _
                        sOrder_Number & "' AND ORDER_TASK_NO = " & sOrder_Task_No
                        
                        
               'Debug.Print sSQL
                Call objSQL.Execute(sSQL)
              
                ' now hide and remove row(s) from grid that is up and being viewed
                objGrid.RowHidden(intCurrentRowIndex) = False
                objGrid.RemoveItem (intCurrentRowIndex)
               
                Count = Count + 1
                If Not gbIsNewResultsList Then
                    gLTotalSavedResults = gLTotalSavedResults - 1
                End If
            End If
        Next
        ' do a grid refresh after we remove all rows
         objGrid.Refresh
        'inactivate action menu
        If objGrid.Rows = 1 Then
           ' have no data just header
            gfrmResults.DeactivateMenus
        End If
    Else
       ' ask user to select a row to take desired action upon
       MsgBox "Please select a result to take the action upon."
    End If
        
GoTo CleanUp

errhandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "in New Results.Gems_SQLSvr_TossResult", "New Results", vbExclamation)

CleanUp:
    If Not objSQL Is Nothing Then Set objSQL = Nothing
    If Not colRowsToToss Is Nothing Then Set colRowsToToss = Nothing
    If Not objGrid Is Nothing Then Set objGrid = Nothing
    
    gfrmResults.gDisableChange = False
End Sub

'*******************************************************************************************
'  Sub Name: GetSetOfCriticalAbnormalValues
'
'  Description: Method is called on start up by initialize and it unlocks any previously locked records
'
'  Parameters:
'
'  History
'  MGumbert - created 10/15/2002 used to be in initialize function broken out to
'             be able to detect errror DB connectivity issue's
'*******************************************************************************************
Private Sub GetSetOfCriticalAbnormalValues()
  Dim sSQL As String
  Dim sProviderIEN As String
  Dim sAbnormal As String
  Dim objSQL As ICHCSII_DAS
  Dim objRecset As ADODB.Recordset
  
    On Error GoTo errhandler
    Set gcolRadAbnormalValues = New Collection

    sSQL = ""
    
    
    'retrieve rad values that require a physician be contacted
    sSQL = "SELECT distinct abnormality_value FROM Abnormal_Codes WHERE Notify_HCP = 'YES' and facility_ncid = " & gsFacilityNCID
    Set objRecset = New ADODB.Recordset
    
    Set objSQL = SelectCHCSII_DAS
    Set objRecset = objSQL.OpenRecordset(sSQL)

    'put into a collection for use later to do comparisons
    'the abnormal lab/rad values that require a physician be contacted
    'test to see if empty result set is returned
    If Not objRecset.EOF And Not objRecset.BOF Then
        Do While Not objRecset.EOF
            sAbnormal = objRecset.Fields(0).Value
            gcolRadAbnormalValues.Add sAbnormal
            objRecset.MoveNext
        Loop
    End If
  
    GoTo CleanUp

errhandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "in New Results.GetSetOfCriticalAbnormalValues", "New Results", vbExclamation)
CleanUp:
    If Not objSQL Is Nothing Then Set objSQL = Nothing
    If Not objRecset Is Nothing Then Set objRecset = Nothing
End Sub

