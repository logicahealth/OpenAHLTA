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
Attribute VB_Name = "PerioPL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ICHCSII_CoreAPI2
Implements ICHCSII_CoreAPI3
Implements ICHCSII_Dental

Private Enum mnumDCT_DentalComponentTab
   lDCT_BigTooth = 0
   lDCT_Notes = 1
   lDCT_Labs = 2
   lDCT_DataViews = 3
   lDCT_Perio = 4
End Enum

Const msMODULE_NAME           As String = "PerioPL."
Private mobjDentalParent   As Object '''CHCSII_Dental.Dental
Private mlngBusyVotes As Long


Public Sub SaveDefaultView(strValue As String)
Dim blnRet As Boolean

    cwobjConfig.SetValue NCID_DENTAL_PERIOCHART_APPLICATION, _
                        cwobjLogonEnviron.UserNCID, _
                        "DefaultSettings", _
                        "ViewType", _
                        strValue
    
    blnRet = cwobjConfig.SaveChanges(NCID_DENTAL_PERIOCHART_APPLICATION, cwobjLogonEnviron.UserNCID)
End Sub
Public Function GetDefaultView() As String
Dim strValue As String

    cwobjConfig.GetValue NCID_DENTAL_PERIOCHART_APPLICATION, _
                                    "DefaultSettings", _
                                    "ViewType", _
                                    strValue
    GetDefaultView = strValue
    
End Function

Public Function GetSVGDirectory() As String
Dim strValue As String
    
    strValue = gobjShared.ImagesPath
    
    GetSVGDirectory = strValue
End Function


Public Function QueryView(ByVal viViewType As Integer) As Boolean
    Select Case viViewType
        Case cwiREGULAR_VIEW
            QueryView = True
        Case Else
            QueryView = False
    End Select
End Function


'This method is called when the CW Core attempts to open a "view" provided by this OLE Server
Public Function OpenView(ByVal viViewType As Integer, _
                         ByVal vlParenthWnd As Long, _
                         rhMainForm As Long) As Boolean

    On Error GoTo errhandler
    
   'Dim tmStart As SYSTEMTIME
   'Dim tmEnd As SYSTEMTIME
   
   'Call GetLocalTime(tmStart)
    'cwobjCWComm.SetStatusText "Opening Dental Perio Charting View..."
    
    'Perform actions based on the view type
    If viViewType = cwiREGULAR_VIEW Then
        'Set global flags
        cwbFirstRegularResize = True
        cwbMainViewRunning = True
        'Here is where we would add menu items
        'add a menu item "template" on CW's main menu
        'ghtemplate = cwobjCWComm.addpopupmenutocw("&Template", cwlMyProgId)
        'add a menu item "Ole Server Menu Item" to the menu "template"
        'glolemenuitem = cwobjcwcomm.addmenuitem("&Ole Server Mneu Item", ghtemplate, cwlmyprogid, "Template menu item")
        
        'Reset the "parent window" fo rthis OLE server's "main" regular form.
        'takes practically 100% of OpenView
        SetParent PerioEntry.hwnd, vlParenthWnd
   

        Set PerioEntry.Parent = Me
        
        'Set Parameter
        rhMainForm = PerioEntry.hwnd
        
    Else 'Unsupported view type
        'set return value
        OpenView = False
        Exit Function
    End If

    'cwobjCWComm.SetStatusText ""

    'set Return Value
    OpenView = True
   
   'Call GetLocalTime(tmEnd)
   'Debug.Print "OpenView: " & TimeDiff(tmStart, tmEnd)

    Exit Function
    
errhandler:
    'show the VB Error
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "PerioOps.OpenView", "PerioOps", vbCritical)
    'Set Return Value
    OpenView = False
    Exit Function
End Function


'Called when the CW Core closes a "view" provided by this OLE Server
Public Sub CloseView(ByVal viViewType As Integer)
    cwobjCWComm.SetStatusText "Closing Dental Perio Charting View..."
    Select Case viViewType
        Case cwiALL_VIEWS, cwiREGULAR_VIEW
            'Close the Main View
            If cwbMainViewRunning Then
                'Set Global Flags
                cwbMainViewRunning = False
                'Indicate that Core is closing the view, but not the OLE Server
                cwbCloseViewMyself = False
                'Close all forms that are associated with all views
                Unload PerioEntry
                Set PerioEntry = Nothing
                Set mobjDentalParent = Nothing
            End If
        Case Else
    End Select
    cwobjCWComm.SetStatusText ""
End Sub

'Called when the user has selected a menu item (from Core's main menu), which is owned by
'this OLE Server
Public Sub MenuItemSelected(ByVal vlMenuId As Long)
    On Error GoTo errhandler
        
    Select Case vlMenuId
        Case glOleMenuItem
            'TODO: take actions needed to process this menu item
        Case Else 'do nothing
    End Select
    Exit Sub
        
errhandler:
    'Show the error
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "PerioOps.MenuItemSelected", "PerioOps", vbExclamation)
    Exit Sub
End Sub

'Called when a new patient has been selected from Core
Public Sub PatientSelected()
    'update the views with the new patient
    If cwbMainViewRunning Then
        'TODO: add code here to update the data in all view with this new patient
    End If
End Sub


'Called when a change has occurred in the client area used by this OLE Server, resulting in a
'need to resize a particular view
Public Sub ResizeView(ByVal viViewType As Integer, ByVal vlLeft As Long, ByVal vlTop As Long, _
                      ByVal vlWidth As Long, ByVal vlHeight As Long)

    On Error GoTo errhandler
    
    'Dim tmStart As SYSTEMTIME
    'Dim tmEnd As SYSTEMTIME
    
    'Call GetLocalTime(tmStart)
    'cwobjCWComm.SetStatusText "Resizing Dental Perio Charting View..."
    
    'reset global client area parameters
    cwlLeft = vlLeft
    cwlTop = vlTop
    cwlWidth = vlWidth
    cwlHeight = vlHeight

    'Perform different actions based upon the view type
    If viViewType = cwiREGULAR_VIEW Then
        If Not cwbFirstRegularResize Then
            'Reset the WindowState to "Normal", so the form can be resized.  This is neeeded
            'because a window cannot be moved ore resized if it is minimized or maximized
            'Note: this generates a form "resize" event
            PerioEntry.WindowState = vbNormal
            'hide the form to avoid screen refreshes
            PerioEntry.Visible = False
            'resize the form to fill the client area
            PerioEntry.Move cwlLeft, cwlTop, cwlWidth, cwlHeight
            PerioEntry.Visible = True
        Else  'first resize, need to perform init procedures
            'reset flag
            cwbFirstRegularResize = False
            'NOTE: Show the form for the first time
            Set PerioEntry.Parent = Me
            PerioEntry.Display
        End If
    End If
    'cwobjCWComm.SetStatusText ""
    'Call GetLocalTime(tmEnd)
    'Debug.Print "ResiveView: " & TimeDiff(tmStart, tmEnd)
    Exit Sub
    
errhandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "PerioOps.ResizeView", "PerioOps", vbExclamation)
    Exit Sub
End Sub
                      

Private Sub Class_Initialize()
    'create shared object used to show vb errors, etc
    Set gobjShared = New CWShared
    'set the help file for this application
    App.HelpFile = gobjShared.CWHelpFile
End Sub


Private Sub Class_Terminate()
    'release references
    Set gobjShared = Nothing
    Set cwobjCWComm = Nothing
    Set cwobjLogonEnviron = Nothing
    Set cwobjConfig = Nothing
    Set cwobjPatient = Nothing
    
    'TODO: release references specific to Perio
    
End Sub


'Used by Core to send messsages to this OLE Server.
Public Sub Message(ByVal viMessageType As Integer, _
                   ByVal vsMessageText As String, _
                   ByVal vsSourceServerNCID As String)
    On Error GoTo errhandler
        
    'Take action based on the message type
    Select Case viMessageType
        Case cwiOK_TO_CHANGE_PATIENT
            'need to define NCID_DENTAL_PERIO
            'cwobjCWComm.Message cwiOK_TO_CHANGE_PATIENT, "Y", vsSourceServerNCID, NCID_DENTAL_PERIO
        Case cwiOK_TO_CLOSE_VIEW
            'cwobjCWComm.Message cwiok_toclose_view, "Y", vsSourceServerNCID, NCID_DENTAL_PERIO
        Case Else 'do nothing
    End Select
    Exit Sub
    
errhandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "PerioOps.Message", "PerioOps", vbExclamation)
    Exit Sub
End Sub

'Called when Core creates an instance of this OLE Server.  It is also used to create and
'initialize global variables that are shared across all views supported by this OLE Server,
'i.e. variables that only need to be initialized once, upon object creation.
Public Sub Initialize(ByVal vlMyProgId As Long, _
                            robjComm As Object, _
                            robjLogonEnvironment As Object, _
                            robjConfig As Object, _
                            robjPatient As Object)
'Dim tmStart As SYSTEMTIME
'Dim tmEnd As SYSTEMTIME

    On Error GoTo errhandler
    
    'Call GetLocalTime(tmStart)
    'Save the ProgID assigned by Core
    cwlMyProgId = vlMyProgId
    'Create references to objects
    Set cwobjCWComm = robjComm
    
    'cwobjCWComm.SetStatusText "Initializing Dental Perio Charting..."
    
    Set cwobjLogonEnviron = robjLogonEnvironment
    Set cwobjConfig = robjConfig
    Set cwobjPatient = robjPatient
    'set global flags
    cwbMainViewRunning = False
    
    'TODO: Add code here to create objects, init variables, etc
    
    'cwobjCWComm.SetStatusText ""
    
    'Call GetLocalTime(tmEnd)
    'Debug.Print "PerioPL.Initialize: " & TimeDiff(tmStart, tmEnd)
    Exit Sub
errhandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "PerioOps.Initialize", "PerioOps", vbExclamation)
    Exit Sub
End Sub

'Called by a client when they want to modify the properties, or the configuration of this
'OLE Server
Public Sub Properties(robjLogonEnvironment As Object, _
                      robjConfig As Object, _
                      ByVal vsOwnerNCID As String, _
                      ByVal vbCalledByDesigner As Boolean)
                      
    On Error GoTo errhandler

    'TODO: add code here to invoke this OLE Server's properties page
    
    Exit Sub
errhandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "PerioOps.Properties", "PerioOps", vbExclamation)
End Sub

'Called after Core creates an instance of this OLE Server.  It is used to create and
'initialize global variables that are shared acrosss all views supported by this OLE Server,
'i.e. variables that only need to be initialized once, upon object creation
Private Sub ICHCSII_CoreAPI2_Initialize2(ByVal vlMyProgId As Long, _
                                               robjComm As Object, _
                                               robjLogonEnvironment As Object, _
                                               robjConfig As Object, _
                                               robjPatient As Object, _
                                               robjUser As Object)
    Set gobjUser = robjUser
    Call Initialize(vlMyProgId, robjComm, robjLogonEnvironment, robjConfig, robjPatient)
End Sub


'Called when Core closes a view provided by this OLE Server.  It differs from the CloseView
'method in that the child module is permitted to say "NO" to Core
Private Function ICHCSII_CoreAPI3_CloseView2(ByVal viViewType As Integer) As Boolean
   Call CloseView(viViewType)
   ICHCSII_CoreAPI3_CloseView2 = True
End Function


'Copied ICHCSII_DENTAL implementation from Dental\Notes\Notes.cls
Private Property Get ICHCSII_Dental_EncounterDental() As Variant

End Property

Private Property Set ICHCSII_Dental_EncounterDental(RHS As Collection)

End Property

Private Property Get ICHCSII_Dental_EncounterDentalLabs() As Variant

End Property

Private Property Set ICHCSII_Dental_EncounterDentalLabs(RHS As Collection)

End Property

Private Property Get ICHCSII_Dental_EncounterDentalNotes() As Variant

End Property

Private Property Set ICHCSII_Dental_EncounterDentalNotes(RHS As Collection)

End Property

Private Property Get ICHCSII_Dental_EncounterDx() As Variant

End Property

Private Property Set ICHCSII_Dental_EncounterDx(RHS As Collection)

End Property

Private Property Get ICHCSII_Dental_EncounterOps() As Variant

End Property

Private Property Get ICHCSII_Dental_EncounterTeeth() As Variant

End Property

Private Property Set ICHCSII_Dental_EncounterTeeth(RHS As Collection)

End Property

Private Property Get ICHCSII_Dental_EncounterTx() As Variant

End Property

Private Property Set ICHCSII_Dental_EncounterTx(RHS As Collection)

End Property

Private Function ICHCSII_Dental_SetDentalParent(objDentalParent As Object) As Boolean
'''************************************************************************
'''Called By:     Any
'''
'''Arguments
'''
'''Description:   Set reference to Dental parent Object
'''
'''Note:
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
    Const sPROC_NAME     As String = "ICHCSII_Dental_SetDentalParent"

'Dim tmStart As SYSTEMTIME
'Dim tmEnd As SYSTEMTIME

   'Call GetLocalTime(tmStart)
    On Error GoTo ErrHnd

    Set mobjDentalParent = objDentalParent
    Set PerioEntry.DentalParent = objDentalParent
    Set PerioEntry.Parent = Me
    
    ICHCSII_Dental_SetDentalParent = True
   'Call GetLocalTime(tmEnd)
   'Debug.Print "PerioPL.SetDentalParent: " & TimeDiff(tmStart, tmEnd)
    Exit Function
ErrHnd:
    ICHCSII_Dental_SetDentalParent = False
    'Call LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")
    Call gobjShared.ShowVBError(Err.Number, Err.Description, sPROC_NAME, "PerioOps", vbExclamation)

    Exit Function
End Function

Private Function ICHCSII_Dental_SetEncounterDentalBT(objBigToothOps As Object) As Boolean

End Function


Private Function ICHCSII_Dental_SetEncounterDentalDataViews(objDataViewsOps As Object) As Boolean

End Function


Private Function ICHCSII_Dental_SetEncounterDentalFM(objFullMouthOps As Object) As Boolean

End Function


Private Function ICHCSII_Dental_SetEncounterDentalLabs(objLabOps As Object) As Boolean

End Function


Private Function ICHCSII_Dental_SetEncounterDentalNotes(objNoteOps As Object) As Boolean

End Function


Private Function ICHCSII_Dental_SetEncounterOps(objEncounterOps As Object) As Boolean
'
End Function


Public Function Mtd_bDentalTabChange(ByVal lFrom As Long, ByVal lTo As Long) As Boolean
'''************************************************************************
'''Called By:     Dental Parent
'''
'''Arguments:
'''lFrom          Identifier of the dental the currently active dental module
'''lTo            Identifier of the dental the about to became active
'''
'''
'''Description:   if lFrom equal to your Tab
'''                  Focus is changing from your tab to lTo Module
'''                  if you ready to release focus: save current data...
'''                     Return true
'''                  else
'''                     return false
'''               Elseif lTo equal to your Tab
'''                  Focus is changing from lTo  tab to your Module
'''                  if you ready to recive focus: Objects are initialized...
'''                     Return true
'''                  else
'''                     return false
'''               Else
'''                  Note: May not be called
'''                  Return true
'''
'''Note:
'''
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:       If successful  - true
'''               Else           - false
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
    Const sPROC_NAME     As String = "Mtd_bDentalTabChange"

    Dim bResult          As Boolean
    
    Static blnNotFirstTabChange As Boolean
    
    'Dim tmStart As SYSTEMTIME
    'Dim tmEnd As SYSTEMTIME
    
    
    On Error GoTo ErrHnd
   
   'Call GetLocalTime(tmStart)
    If lFrom = lDCT_Perio Then
        SaveView
        
        ClosePerioPopups
        
        '''If was able to save current notes
        'bResult = DTNotes_bSave(True, True)
        bResult = True
    ElseIf lTo = lDCT_Perio Then
'        If DTNotes_bMainViewRunning Then
'            '''If was able to update notes with data that can be modified outside of the Templates
'            bResult = DTNotes_bReadExternalItems()
'        Else
'            '''Called prior to open view for some reason
'            bResult = True
'            '''If false is returned Open view is not going to be called
'        End If
        
        'May still need to refresh Notes Prognosis values here
        'PerioEntry.Display
        
        'Don't update prognosis if we're just starting the view for the 1st time
        If blnNotFirstTabChange Then
            PerioEntry.UpdateNotesData
        End If
        
        blnNotFirstTabChange = True
        bResult = True
    Else
        bResult = True
    End If
   
    Mtd_bDentalTabChange = bResult

    'Call GetLocalTime(tmEnd)
    'Debug.Print "Mtd_bDentalTabChange: " & TimeDiff(tmStart, tmEnd)
    Exit Function
ErrHnd:
    Mtd_bDentalTabChange = False
    'Call LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "", True)
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "PerioOps.Mtd_bDentalTabChange", "PerioOps", vbExclamation)
   
    Exit Function
End Function

Public Function SaveView() As Boolean
'''************************************************************************
'''Called By:     Dental Parent
'''
'''Arguments:
'''
'''Description:   Save current data
'''
'''Note:
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:       If successful  - true
'''               Else           - false
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "SaveView"
   
   Dim bResult          As Boolean
   
On Error GoTo ErrHnd

    cwobjCWComm.SetStatusText "Saving Dental Perio Charting View..."
   'TODO: Insert code to save data
   bResult = True
   bResult = PerioEntry.SaveExams
'   If DTNotes_bMainViewRunning Then
'      '''If was able to save current notes
'      If DTNotes_bSave(True, True) Then
'         If DTNTemplates_sDisplayNotes = "" Then
'            bResult = True
'         End If
'      End If
'   Else
'      bResult = True
'   End If
   
   SaveView = bResult
    cwobjCWComm.SetStatusText ""
Exit Function
ErrHnd:
   SaveView = False
   'Call LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")
   Call gobjShared.ShowVBError(Err.Number, Err.Description, msMODULE_NAME & "." & sPROC_NAME, msMODULE_NAME, vbExclamation)
End Function

Private Sub ClosePerioPopups()
   PerioEntry.ClosePopups
End Sub

Private Function PerioNote() As String
   PerioNote = PerioEntry.StatisticsNote
End Function


Public Sub VoteBusy()
   mlngBusyVotes = mlngBusyVotes + 1
   If mlngBusyVotes = 1 Then
      cwobjCWComm.Message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_DENTAL_PERIOCHART_APPLICATION
   End If
End Sub
Public Sub VoteNotBusy()
   mlngBusyVotes = mlngBusyVotes - 1
   If mlngBusyVotes = 0 Then
      cwobjCWComm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_DENTAL_PERIOCHART_APPLICATION
   End If
End Sub
