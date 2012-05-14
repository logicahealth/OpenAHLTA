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
Attribute VB_Name = "clsNotesPlus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit


'************************************************************************
'Module:        Notes/Notes.cls
'
'Description:   Dental Notes CHCSII interface class
'
'Modules Required:
'
'************************************************************************

Const msMODULE_NAME           As String = "Notes."

Private XMLPersist As clsXMLPersist

Private Const cADULT_AGE As Integer = 13

'The variable for the GUI form
Private ofrmNotes As frmNotes
Private oInitUtils As clsInitUtils

Private Enum mnumDCT_DentalComponentTab
   lDCT_BigTooth = 0
   lDCT_Notes = 1
   lDCT_Labs = 2
   lDCT_DataViews = 3
End Enum

Public Enum DirtyFlagEnum
    'For objects only
    dfNoAction = 0
    dfUpdate = 1
    dfAdd = 2
    dfdelete = 3
End Enum

'Enum for Perio Exams
Public Enum eEvalType
    perInitial
    perReEval
    perPostSurgical
    perMaintenance
    perConsultation
End Enum


''''''''''''''' Objects Section '''''''''''''''''''''''''''''''''''''''''

Implements ICHCSII_CoreAPI2
Implements ICHCSII_CoreAPI3
Implements ICHCSII_User
Implements ICHCSII_Dental
Implements ICHCSII_Dental2

Private mobjDentalParent   As Object '''CHCSII_Dental.Dental
Private mobjDentEnc As IntrfcCHCSII_Dental.IDental_Enc
Private lParentHandle As Long
Private bPatientIsAdult As Boolean






Private Sub InitNotesPlus()

   Dim dSysDate As Date
   Dim bResult          As Boolean

        Set ofrmNotes = New frmNotes

        'Set the patientid, encounterid and providerid
        #If debugmode = 1 Then
            'Test values
            ofrmNotes.lPatientID = 104520
            ofrmNotes.lDentalEncounterID = 1717
            '        ofrmNotes.lPatientID = 155018
            '        ofrmNotes.lDentalEncounterID = 999
            
            ofrmNotes.lProviderID = 456
            ofrmNotes.bPatientIsAdult = True
            oInitUtils.PatientAge = 18
        #Else
        
            ofrmNotes.lPatientID = cwobjPatient.UnitNumber
            ofrmNotes.lDentalEncounterID = mobjDentalParent.psDentalEncounterID
            ofrmNotes.lProviderID = cwobjLogonEnviron.UserNCID
            dSysDate = GetSysDate
            'KDunne 1/9/2004
            'SCR 47838 - Age was incorrectly being rounded up.
            ofrmNotes.bPatientIsAdult = Val(goCWShared.CalcAge(cwobjPatient.DateOfBirth)) >= cADULT_AGE
            oInitUtils.PatientAge = Val(goCWShared.CalcAge(cwobjPatient.DateOfBirth))

        #End If
        
        'Init the XMLPersist class and load the XML data
        Set XMLPersist = New clsXMLPersist
        XMLPersist.ExamXML.Load mobjDentalParent.psDentalEncounterID

        Dim ltinit As Double
        ltinit = Timer
        bResult = ofrmNotes.SetReftoDentalParent(mobjDentalParent)
        Set ofrmNotes.DentalEncounter = mobjDentEnc
        
        ofrmNotes.Initialize True
        
        Debug.Print "init notes " & Timer - ltinit

End Sub

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
'''Programmer:    Scott Foster
'''************************************************************************
   Const sPROC_NAME     As String = "Mtd_bDentalTabChange"
Debug.Print "Mtd_bDentalTabChange"

   Dim bResult          As Boolean

On Error GoTo ErrHnd
   
    If lFrom = lDCT_Notes Then
        If Not ofrmNotes Is Nothing Then
            bResult = ofrmNotes.SaveAll
        Else
            bResult = True
        End If

      
   ElseIf lTo = lDCT_Notes Then
      'TODO: Reload data that was able to be modified outside of notes, like the grid.
        If Not ofrmNotes Is Nothing Then
                ofrmNotes.RefreshCurrentTab

'            ofrmNotes.Initialize True
'            ofrmNotes.Initialize False
        End If
        
        bResult = True
      
   Else
      bResult = True
   End If
   
   Mtd_bDentalTabChange = bResult

Exit Function

ErrHnd:
   goCWShared.ShowVBError Err.Number, Err.Description, Err.Source & ":" & sPROC_NAME, App.Title, vbCritical

End Function

Public Property Get PrimProviderID() As Long
    
    PrimProviderID = glPrimProviderID
    
End Property

Public Property Let PrimProviderID(vData As Long)

    glPrimProviderID = vData
    
End Property


Public Function QueryView(ByVal viViewType As Integer) As Boolean
'*******************************************************************************************
'''  Description: This method is called when a client wants to know if this OLE Server
'''               supports a specific "view type".
'''
'''  Parameters:
'''    viViewType - Contains the "type" of view being queried for.
'''
'''  Return Value:
'''    Returns TRUE if this OLE Server supports the passed in "view type".
'''    Returns FALSE if this OLE Server does not support the passed in "view type".
'*******************************************************************************************
   Const sPROC_NAME     As String = "QueryView"
   
   Dim bResult          As Boolean
   
    Select Case viViewType
       'Supported
       Case cwiREGULAR_VIEW
          bResult = True                '''This view is NOT optional--
        'Not supported
        'Case cwiPATIENT_SUMMARY_VIEW
        'Case cwiCLINICIAN_SUMMARY_VIEW
        'Case cwiENCOUNTER_SUMMARY_VIEW
        'Case cwiPROPERTY_VIEW
        'Turn OFF the "Options" button on CW
      Case Else
         '''unknown "view type" - Not supported
   End Select

   QueryView = bResult
   
End Function

Public Function OpenView(ByVal viViewType As Integer, _
                          ByVal vlParenthWnd As Long, _
                          rhMainForm As Long) _
                          As Boolean
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
   Const sPROC_NAME     As String = "OpenView"
   Dim dSysDate As Date
   
Debug.Print "OpenView"

   Dim bResult          As Boolean
'   Dim sErrMsg          As String
   
'   Dim objDLog          As New clsLibDebug

On Error GoTo ErrHandler

'   Call objDLog.Mtd_bLogMessage(msMODULE_NAME, sPROC_NAME)
   
   If viViewType = cwiREGULAR_VIEW Then
        'Load the notes form and get the data
        Screen.MousePointer = vbHourglass
        'Load notes/controls set active tab etc.
        
        'Determine basic things about the patient
'        ofrmNotes.Initialize False
        
        Dim ltshow As Double
        ltshow = Timer
        'Set the parent of ofrmNotes to the handle provided by core
        lParentHandle = vlParenthWnd
        Call SetParent(ofrmNotes.hwnd, lParentHandle)
        rhMainForm = ofrmNotes.hwnd
        
        ofrmNotes.Initialize False
        'ofrmNotes.RefreshCurrentTab
        
        OpenView = True
        ofrmNotes.Show
        
        Debug.Print "show notes " & Timer - ltshow
        
    Else '--unsupported view "type"--
        goCWShared.ShowVBError 0, "The view type of: " & viViewType & " is not supported.", "NotesPlus-OpenView", App.Title, vbInformation
    End If
    
    
    Screen.MousePointer = vbDefault
    
Exit Function

ErrHandler:
    goCWShared.ShowVBError Err.Number, Err.Description, Err.Source & ": " & sPROC_NAME, App.Title, vbCritical
    Screen.MousePointer = vbDefault
   
End Function

Public Sub CloseView(ByVal viViewType As Integer)
'*******************************************************************************************
'  Sub Name: CloseView
'
'  Description: This method is called when the CW Core closes a "view" provided by this OLE
'               Server.
'
'  Parameters:
'    viViewType - Contains the "type" of view that the CW Core is closing.
'*******************************************************************************************
   Const sPROC_NAME     As String = "CloseView"

'   Dim objDLog          As New clsLibDebug

On Error GoTo ErrHandler
    
    If Not ofrmNotes Is Nothing Then
        Unload ofrmNotes
        Set ofrmNotes = Nothing
    End If
    
    Debug.Print "CloseView:" & viViewType
    
Exit Sub

ErrHandler:
   goCWShared.ShowVBError Err.Number, Err.Description, Err.Source, App.Title, vbCritical
   
End Sub

Public Sub MenuItemSelected(ByVal vlMenuId As Long)
'*******************************************************************************************
'  Description: This method is called when the user has selected a menu item (from the CW
'               Core's main menu), which is "owned" by this OLE Server.
'
'               NOTE: This is an OPTIONAL method.
'
'  Parameters:
'    vlMenuId - Contains the menu id of the menu item the user selected.
'*******************************************************************************************
   Const sPROC_NAME     As String = "MenuItemSelected"
   
On Error GoTo ErrHandler
   
   '-- --------------------------------------------------------------------------------
   '--  NOTE: You don't have to create menu items, these are optional, that is why
   '--        the code is commented out.
   '-- --------------------------------------------------------------------------------
'    '---perform the actions associated with the menu item that was selected---
    Select Case vlMenuId
   
'            '******************************************************************************
'            '  TO DO:
'            '    Take actions needed to process this menu item.
'            '******************************************************************************
'
'
'
'        Case Else   '--do nothing--
    End Select
   
Exit Sub

ErrHandler:
   goCWShared.ShowVBError Err.Number, Err.Description, Err.Source, App.Title, vbCritical
   
End Sub

Public Sub PatientSelected()
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
   Const sPROC_NAME     As String = "PatientSelected"
   'Debug.Print "Mtd_bDentalTabChange"

'   Dim objDLog          As New clsLibDebug

On Error GoTo ErrHandler

'   Call objDLog.Mtd_bLogMessage(msMODULE_NAME, sPROC_NAME)

    'TODO: SAVE AND CLOSE THE FORM
    
Exit Sub

ErrHandler:
   goCWShared.ShowVBError Err.Number, Err.Description, Err.Source & ": " & sPROC_NAME, App.Title, vbCritical

End Sub

Public Sub ResizeView(ByVal viViewType As Integer, _
                       ByVal vlLeft As Long, _
                       ByVal vlTop As Long, _
                       ByVal vlWidth As Long, _
                       ByVal vlHeight As Long)
'''*******************************************************************************************
'''  Description: This method is called when a change has occurred in the "client area"
'''               used by this OLE Server, resulting in a need to resize a particular "view".
'''
'''  Parameters:
'''    viViewType - Contains the "type" of view being requested to resize by the CW Core.
'''    vlLeft - Contains the "left" coordinate for the OLE server.
'''    vlTop - Contains the "top" coordinate for the OLE server.
'''    vlWidth - Contains the "width" for the OLE server.
'''    vlHeight - Contains the "height" for the OLE server.
'*******************************************************************************************
   Const sPROC_NAME     As String = "ResizeView"

'   Dim objDLog          As New clsLibDebug
   Dim sErrMsg          As String
Debug.Print "ResizeView"

On Error GoTo ErrHandler

'   Call objDLog.Mtd_bLogMessage(msMODULE_NAME, sPROC_NAME)
   
   'If supported view and main form and required objects are initialized and loaded
    If viViewType = cwiREGULAR_VIEW Then
        If Not ofrmNotes Is Nothing Then
            If ofrmNotes.bFormInitialized = True Then
                ofrmNotes.Visible = False
                ofrmNotes.Move vlLeft, vlTop, vlWidth, vlHeight
                ofrmNotes.Visible = True
            End If
        Else
'          sErrMsg = "The notes plus form is not initialized."
        End If
    Else
       sErrMsg = "The view type is not supported."
    End If
   
    If sErrMsg <> "" Then
        goCWShared.ShowVBError 0, sErrMsg, Err.Source & ": " & sPROC_NAME, App.Title, vbCritical
    End If
   
Exit Sub

ErrHandler:
   goCWShared.ShowVBError Err.Number, Err.Description, Err.Source & ": " & sPROC_NAME, App.Title, vbCritical

   
End Sub

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
'''************************************************************************
   Const sPROC_NAME     As String = "SaveView"
   Debug.Print "SaveView"

   Dim bResult          As Boolean
   
On Error GoTo ErrHnd
    
    If Not ofrmNotes Is Nothing Then
        bResult = ofrmNotes.SaveAll
        SaveView = bResult
    Else
        SaveView = True
    End If
    

Exit Function

ErrHnd:
   goCWShared.ShowVBError Err.Number, Err.Description, Err.Source, App.Title, vbCritical

End Function

Private Sub Class_Terminate()
'************************************************************************
'Called By:     None
'
'Arguments:
'Description:   Initialize interface
'
'Note:
'
'Warning:
'
'Updates Globals:
'
'Returns:
'
'************************************************************************
   
   Const sPROC_NAME     As String = "Class_Terminate"

'   Dim objDLog          As New clsLibDebug
   
On Error Resume Next

'    Call objDLog.Mtd_bLogMessage(msMODULE_NAME, sPROC_NAME)
    Unload ofrmNotes
    XMLPersist.TerminateObjects
    Set XMLPersist = Nothing
    Set ofrmNotes = Nothing
    Set oInitUtils = Nothing
    Set cwobjCWComm = Nothing
    Set cwobjLogonEnviron = Nothing
    Set cwobjConfig = Nothing
    Set cwobjPatient = Nothing
    Set mobjDentalParent = Nothing
    Set goCWShared = Nothing
    Set oInitUtils = Nothing
    Set gobjUser = Nothing
    Set goconn = Nothing
    Set mobjDentEnc = Nothing
    
End Sub

Public Sub Message(ByVal viMessageType As Integer, _
                    ByVal vsMessageText As String, _
                    ByVal vsSourceServerNCID As String)
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
Const sPROC_NAME     As String = "Message"
Debug.Print "Message"
On Error GoTo ErrHandler
   
   '-- -----------------------------------------------------------------------------
   '--  Take corresponding actions based upon the message "type".
   '--
   '--  NOTE: This method is required, but currently there are NO REQUIRED messages
   '--        to process.
   '-- -----------------------------------------------------------------------------
   If Not cwobjCWComm Is Nothing Then
      Select Case viMessageType
         Case cwiOK_TO_CHANGE_PATIENT
            Call cwobjCWComm.Message _
               (cwiOK_TO_CHANGE_PATIENT, "Y", vsSourceServerNCID, NCID_DENTAL_NOTES_APPLICATION)
               'Save the form?
         Case cwiOK_TO_CLOSE_VIEW
            Call cwobjCWComm.Message _
               (cwiOK_TO_CLOSE_VIEW, "Y", vsSourceServerNCID, NCID_DENTAL_NOTES_APPLICATION)
               'Save the form?
         Case cwiGENERIC
            Select Case UCase(vsMessageText)
               Case "EXIT_APP"
                  'TODO: Unload the form and class?
            End Select
      End Select
   End If
   
Exit Sub

ErrHandler:
   goCWShared.ShowVBError Err.Number, Err.Description, Err.Source, App.Title, vbCritical
   
End Sub

Public Sub Initialize(ByVal vlMyProgId As Long, _
                       robjComm As Object, _
                       robjLogonEnvironment As Object, _
                       robjConfig As Object, _
                       robjPatient As Object)
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
   Const sPROC_NAME     As String = "Initialize"
    Debug.Print "Initialize"

'   Dim objDLog          As New clsLibDebug

On Error GoTo ErrHandler

    'Open the database connection
    Call OpenDB
'    'Init the XMLPersist class and load the XML data
'    Dim XMLPersist As New clsXMLPersist
'    XMLPersist.ExamXML.Load DentalEncounterID
'    Set XMLPersist = Nothing
    
'   Call objDLog.Mtd_bLogMessage(msMODULE_NAME, sPROC_NAME)
   
    '---Instantiate Help-file object
    App.HelpFile = goCWShared.CWHelpFile
    '-- -save the "ProgId" assigned to me by the CW Core---
    cwlMyProgId = vlMyProgId
    
    '-- -create references to passed in objects---
    
    Set oInitUtils = New clsInitUtils
    
#If debugmode = 1 Then
    'Instantiate a new comm object because I am not in core
    Set cwobjCWComm = New Comm
    Set oInitUtils.CWComm = cwobjCWComm
    oInitUtils.MyProgID = 0
#Else
    'Get the comm object from core
   Set cwobjCWComm = robjComm
   Set oInitUtils.CWComm = cwobjCWComm
   oInitUtils.MyProgID = cwlMyProgId
#End If

   Set cwobjLogonEnviron = robjLogonEnvironment
   Set cwobjConfig = robjConfig
   Set cwobjPatient = robjPatient
   
   '-- -set global flags---

   '-- ****************************************************************************
   '--  TO DO:
   '--    Add code here to create objects, initialize global variables, etc.
   '-- ****************************************************************************

Exit Sub

ErrHandler:
    goCWShared.ShowVBError Err.Number, Err.Description, Err.Source & "." & sPROC_NAME, App.Title, vbCritical
   
End Sub

 

Public Sub ICHCSII_CoreAPI2_Initialize2(ByVal vlMyProgId As Long, _
                                        robjComm As Object, _
                                        robjLogonEnvironment As Object, _
                                        robjConfig As Object, _
                                        robjPatient As Object, _
                                        robjUser As Object)
'******************************************************************************************
'  Sub Name: Initialize
'
'  Description: This method is called when CW "creates" an instance of this OLE Server.  It
'               is used to create and initialize global variables that are shared across all
'               "views" supported by this OLE Server, i.e. variables that only need to be
'               initialized ONCE, upon object creation.
'
'  Parameters:
'    vlMyProgId - Contains the "ProgId" that CW has assigned to this OLE Server.
'    robjComm - Contains a communications object that is used to handle all communication
'               between this OLE Server and the CW core.
'    robjLogonEnviron - Contains the logon environment object.
'    robjConfig - Contains the configuration object used for getting and setting configuration
'                 information for a specific application.
'    robjPatient - Contains the patient used to do patient searches and selections
'*******************************************************************************************
   Const sPROC_NAME     As String = "ICHCSII_CoreAPI2_Initialize2"
   
   Dim objRoles      As ICHCSII_Roles
   Dim objResources  As ICHCSII_Resources
'   Dim objDLog          As New clsLibDebug

On Error GoTo ErrHandler

'   Call objDLog.Mtd_bLogMessage(msMODULE_NAME, sPROC_NAME)

   Set cwobjPatient = robjPatient
   
   Call Initialize(ByVal vlMyProgId, robjComm, robjLogonEnvironment, robjConfig, robjPatient)
   
   Set gobjUser = robjUser

Exit Sub

ErrHandler:
   goCWShared.ShowVBError Err.Number, Err.Description, Err.Source & ":" & sPROC_NAME, App.Title, vbCritical

End Sub

Public Sub Properties(robjLogonEnvironment As Object, _
                       robjConfig As Object, _
                       ByVal vsOwnerNCID As String, _
                       ByVal vbCalledByDesigner As Boolean)
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
   Const sPROC_NAME     As String = "Properties"
   
On Error GoTo ErrHandler
   
   '-- *****************************************************************
   '--  TO DO:
   '--    Add code here to invoke this OLE Server's properties page.
   '-- *****************************************************************
   
Exit Sub

ErrHandler:
   goCWShared.ShowVBError Err.Number, Err.Description, Err.Source & ": " & sPROC_NAME, App.Title, vbCritical
   
End Sub



Private Function ICHCSII_CoreAPI3_CloseView2(ByVal viViewType As Integer) As Boolean

    If viViewType = cwiREGULAR_VIEW Or viViewType = cwiALL_VIEWS Then
        'TODO: SAVE THE FORM
    End If
   
End Function

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

On Error GoTo ErrHnd

   Set mobjDentalParent = objDentalParent
   ICHCSII_Dental_SetDentalParent = True
   
Exit Function

ErrHnd:
   ICHCSII_Dental_SetDentalParent = False
   goCWShared.ShowVBError Err.Number, Err.Description, Err.Source & ":" & sPROC_NAME, App.Title, vbCritical

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
'*** MsgBox Fix - Code_Validator auto change made 3/21/2003
MsgBxARMd "here"
End Function


Private Function ICHCSII_Dental_SetEncounterOps(objEncounterOps As Object) As Boolean
'
End Function


Private Property Set ICHCSII_Dental2_DentalEncounter(RHS As IntrfcCHCSII_Dental.IDental_Enc)
    
    Set mobjDentEnc = RHS
    
    InitNotesPlus
    
End Property

Private Function ICHCSII_User_CACPresent() As Boolean

End Function

Private Function ICHCSII_User_CACSignText(ByVal PIN As String, ByVal Text As String, Signature As String, Optional ErrCode As Long, Optional ErrMsg As String) As Boolean

End Function

Private Function ICHCSII_User_CACValidatePIN(ByVal PIN As String) As Boolean

End Function

Private Function ICHCSII_User_GetAttribute(AttributeName As String) As String
   
End Function

Private Function ICHCSII_User_GetAttributeNames() As Collection
   
End Function

Private Function ICHCSII_User_HasPrivilege(NCID As String) As Boolean
   
End Function

Private Function ICHCSII_User_HasPrivilegeEx(PrivilegeArea As String, _
                                              PrivilegeType As CHCSII_ClientSecurity.CHCSII_PrivilegeType, _
                                              Optional OtherType As Integer) _
                                              As Boolean
   
End Function

Private Property Get ICHCSII_User_LoginMode() As CHCSII_ClientSecurity.CHCSII_LoginMode

End Property

Private Property Get ICHCSII_User_Resources() As CHCSII_ClientSecurity.ICHCSII_Resources
   
End Property

Private Property Get ICHCSII_User_Roles() As CHCSII_ClientSecurity.ICHCSII_Roles
   
End Property

Private Property Get ICHCSII_User_UserId() As String
   
End Property

Private Property Get ICHCSII_User_UserName() As String
   
End Property

Private Function ICHCSII_User_ValidatePassword(Password As String) As Boolean
   
End Function






