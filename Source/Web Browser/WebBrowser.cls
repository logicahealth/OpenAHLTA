VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Browser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Implements ICHCSII_CoreAPI2

Private cwobjUser As CHCSII_ClientSecurity.ICHCSII_User
Private ghTemplate0, ghTemplate1, ghTemplate2, ghTemplate3, test, StartCount, nTotalFav, CurrentCap
Public gobjShared As CWShared
Dim objPics(9) As Picture
Dim bResult As Boolean
Dim sEnterpriseNCID As String
Dim sFacilityNCID As String
Dim sRoleNCID As String
Dim sUserNCID As String
Dim rsValue As String
Dim OshCount As String
Dim sFavoriteKey As String
Dim sMessage, sTitle, sDefault, sFavCap As String
Dim sEmpty As String
Dim uid As String
Dim pwd As String
Dim sVMSUser As String
Dim sVMSPassword As String
Dim sVMSScript As String
Dim sCHCSIPAddress As String
Dim lMyProgId As Long
Dim StartingAddress As String
Dim mainForm As frmBrowser

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
    
    '---perform actions based on the "view type"---
    If viViewType = cwiREGULAR_VIEW Then
        '---set global flags---
        cwbFirstRegularResize = True
        cwbMainViewRunning = True
                                
        '---------------------------------------------------------------------------------------
        '  If this OLE Server needs it's own menu items, here is a good place to add these
        '    menu items onto the CW Core's main menu.
        '
        '  NOTE: This is OPTIONAL.  You do NOT have to create menu items. Thats why they are
        '        commented out.
        '---------------------------------------------------------------------------------------
        '---add a menu item "Template" on CW's main menu---
        'ghTemplate = cwobjCWComm.AddPopupMenuToCW("&Template", lMyProgId)
        'ghTemplate0 = cwobjCWComm.AddViewMenuItem("Home URL", lMyProgId, "test1")
        'ghTemplate1 = cwobjCWComm.AddViewMenuItem("www.lycos.com", lMyProgId, "test1")
        'ghTemplate2 = cwobjCWComm.AddViewMenuItem("www.altavista.com", lMyProgId, "test1")
        'ghTemplate3 = cwobjCWComm.AddViewMenuItem("www.cnn.com", lMyProgId, "test1")
                
        '---add a menu item "Ole Server Menu Item" to the menu "Template"---
        
        glOleMenuItem0 = cwobjCWComm.AddActionMenuItem("Back", "left.bmp", lMyProgId, "Back")
        glOleMenuItem1 = cwobjCWComm.AddActionMenuItem("Forward", "right.bmp", lMyProgId, "Forward")
        glOleMenuItem2 = cwobjCWComm.AddActionMenuItem("Home", "home.bmp", lMyProgId, "Home")
        glOleMenuItem3 = cwobjCWComm.AddActionMenuItem("Search", "search.bmp", lMyProgId, "Search")
        glOleMenuItem4 = cwobjCWComm.AddActionMenuItem("Refresh", "Refresh.bmp", lMyProgId, "Refresh")
        glOleMenuItem5 = cwobjCWComm.AddActionMenuItem("Stop", "stop.bmp", lMyProgId, "Stop")
        glOleMenuItem6 = cwobjCWComm.AddActionMenuItem("Favorites", "new.bmp", lMyProgId, "Favorites")
        glOleMenuItem7 = cwobjCWComm.AddActionMenuItem("Add", "favorites.bmp", lMyProgId, "Add")
        glOleMenuItem8 = cwobjCWComm.AddActionMenuItem("Organize", "OrganizeFav.bmp", lMyProgId, "Organize")
        glOleMenuItem9 = cwobjCWComm.AddActionMenuItem("Import", "ImportFav.bmp", lMyProgId, "Import")
        glOleMenuItem10 = cwobjCWComm.AddActionMenuItem("Export", "ExportFav.bmp", lMyProgId, "Export")
        glOleMenuItem11 = cwobjCWComm.AddActionMenuItem("Print WebPage", "PrinterD.bmp", lMyProgId, "Print WebPage")
                 
        '-----------------------------------------------------------------------------------------
        '  Reset the "parent window" for this OLE server's "main" regular form.  This is the main
        '    "trick" that makes the OLE server look like it is part of the CW Core.
        '
        '  NOTE: This action invokes the main form's "Load" event, but the form is NOT shown.
        '        Make sure that you DO NOT perform a "Me.Show" in the form's "Load" event, as this
        '        will mess up the "SetParent" API call.  Make sure that you DO NOT "Show" the
        '        "main" regular form in this method, that will happen in the "ResizeView" method.
        '-----------------------------------------------------------------------------------------
        SetParent mainForm.hwnd, vlParenthWnd
        
        '---set parameter---
        rhMainForm = mainForm.hwnd
    Else    '--unsupported view "type"--
        '---set return value---
        OpenView = False
        Screen.MousePointer = vbDefault
        Exit Function
    End If
        
    '---set return value---
    OpenView = True
    Screen.MousePointer = vbDefault
    Exit Function
    
ErrHandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "Template.OpenView", "Template", vbCritical)
    
    '---set return value---
    OpenView = False
    Screen.MousePointer = vbDefault
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
    If viViewType = cwiALL_VIEWS Then
        '---close the "main view"---
        If cwbMainViewRunning Then
            '---set global flags---
            cwbMainViewRunning = False
            
            '---this indicates that the CW Core is closing the "view", NOT this OLE Server---
            cwbCloseViewMyself = False
            
            '---close ALL forms that are associated with all "views"---
            Unload mainForm
        End If
    ElseIf viViewType = cwiREGULAR_VIEW Then
        '---close the "main view"---
        If cwbMainViewRunning Then
            '---set global flags---
            cwbMainViewRunning = False
            
            '---this indicates that the CW Core is closing the "view", NOT this OLE Server---
            cwbCloseViewMyself = False
            
            '---close ALL forms that are associated with all "views"---
            Unload mainForm
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
Static intcount As Integer

sEnterpriseNCID = cwobjLogonEnviron.EnterpriseNCID
sFacilityNCID = cwobjLogonEnviron.FacilityNCID
sRoleNCID = cwobjLogonEnviron.RoleNCID
sUserNCID = cwobjLogonEnviron.UserNCID

Set objPics(0) = LoadPicture(gobjShared.ImagesPath & "\frame01.gif")
Set objPics(1) = LoadPicture(gobjShared.ImagesPath & "\frame02.gif")
Set objPics(2) = LoadPicture(gobjShared.ImagesPath & "\frame03.gif")
Set objPics(3) = LoadPicture(gobjShared.ImagesPath & "\frame04.gif")
Set objPics(4) = LoadPicture(gobjShared.ImagesPath & "\frame05.gif")
Set objPics(5) = LoadPicture(gobjShared.ImagesPath & "\frame06.gif")
Set objPics(6) = LoadPicture(gobjShared.ImagesPath & "\frame07.gif")
Set objPics(7) = LoadPicture(gobjShared.ImagesPath & "\frame08.gif")
Set objPics(8) = LoadPicture(gobjShared.ImagesPath & "\frame09.gif")
    
    On Error GoTo ErrHandler
    
    '----------------------------------------------------------------------------------
    '  NOTE: You don't have to create menu items, these are optional, that is why
    '        the code is commented out.
    '----------------------------------------------------------------------------------
'    '---perform the actions associated with the menu item that was selected---
'    Select Case vlMenuId
'        Case glOleMenuItem
'            '******************************************************************************
'            '  TO DO:
'            '    Take actions needed to process this menu item.
'            '******************************************************************************
        
        mainForm.timTimer.Enabled = True
         
        Select Case vlMenuId 'Button.Key
            Case glOleMenuItem0  'Back
                 On Error GoTo Out
                 mainForm.brwWebBrowser.GoBack
            Case glOleMenuItem1  'Forward
                On Error GoTo Out
                mainForm.brwWebBrowser.GoForward
            Case glOleMenuItem2  'Home
                mainForm.brwWebBrowser.GoHome
            Case glOleMenuItem3      'Search
                mainForm.brwWebBrowser.GoSearch
            Case glOleMenuItem4  'Refresh
                mainForm.brwWebBrowser.Refresh
            Case glOleMenuItem5      'Stop
                mainForm.timTimer.Enabled = False
                mainForm.brwWebBrowser.Stop
                cwobjCWComm.SetStatusText (mainForm.brwWebBrowser.LocationName)
                intcount = 0
                mainForm.Timer1.Enabled = False
                mainForm.Picture1.PaintPicture objPics(intcount), 0, 0
            Case glOleMenuItem6  'Favorites, SCR 22490 - The system shall display web pages in CHCSII browser upon user
                                 '                       selecting a web page from favorites list
                cwobjCWComm.Message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_BROWSER_APPLICATION
                Dim FavDialog As GotoFavDialog
                Set FavDialog = New GotoFavDialog
                Load FavDialog
                FavDialog.Show vbModal
                Set FavDialog = Nothing
                cwobjCWComm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_BROWSER_APPLICATION
                mainForm.cboAddress.SetFocus
                SendKeys "{ENTER}"
            Case glOleMenuItem7  'Add, SCR 22482 - Add Internet web pages to favorites
                cwobjCWComm.Message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_BROWSER_APPLICATION
                mnuFavAdd
                cwobjCWComm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_BROWSER_APPLICATION
            Case glOleMenuItem8  'Organize, SCR 22483 - Delete Internet web pages from favorites
                                 '          SCR 22484 - Provide a file folder structure for favorite web pages
                                 '          SCR 22486 - Allow user to move web pages between favorites file folders
                                 
                cwobjCWComm.Message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_BROWSER_APPLICATION
                mainForm.brwWebBrowser.Navigate2 CStr("javascript:window.external.ShowBrowserUI('OrganizeFavorites', null)")
                cwobjCWComm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_BROWSER_APPLICATION
            Case glOleMenuItem9  'Import, SCR 22487 - Import web pages favorites into CHCSII web browser
                cwobjCWComm.Message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_BROWSER_APPLICATION
                mnuFavImport
                cwobjCWComm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_BROWSER_APPLICATION
            Case glOleMenuItem10  'Export, SCR 22487 - Import web pages favorites into CHCSII web browser
                cwobjCWComm.Message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_BROWSER_APPLICATION
                mnuFavExport
                cwobjCWComm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_BROWSER_APPLICATION
            Case glOleMenuItem11    'Print, SCR 22491 - Enable user to print web page content
                cwobjCWComm.Message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_BROWSER_APPLICATION
                mnuFavPrint
                cwobjCWComm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_BROWSER_APPLICATION
Out:
            End Select


    Exit Sub

ErrHandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "Template.MenuItemSelected", "Template", vbExclamation)
    'Resume Next
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
    '---update the all running "views" with the new patient---
    If cwbMainViewRunning Then
        '****************************************************************************
        '  TO DO:
        '    Add code here to update the data in all "running" views with this new
        '    patient.
        '****************************************************************************
        
        
    
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
            '---------------------------------------------------------------------------------
            '  Reset the "WindowState" to "Normal", so the form can be resized.  This is
            '    needed because a window CANNOT be moved or resized if it is "minimized" or
            '    "maximized".
            '
            '  NOTE: This generates a form "resize" event.
            '---------------------------------------------------------------------------------
            mainForm.WindowState = vbNormal
            
            '---hide the form to so that annyoing "screen blinking" will not be seen while controls are moved and resized---
            mainForm.Visible = False
            
            '---resize form to fill entire "client area", NOTE: This generates a "resize" event---
            mainForm.Move cwlLeft, cwlTop, cwlWidth, cwlHeight
                
            mainForm.Visible = True
        Else    '--first "regular" resize, need to perform "initailization" procedures--
            '---reset flag so this code will not be executed again---
            cwbFirstRegularResize = False
        
            '----------------------------------------------------------------------------
            '  NOTE: Here is where you "Show" the "Regular View" form for the FIRST time.
            '----------------------------------------------------------------------------
            mainForm.Move cwlLeft, cwlTop, cwlWidth, cwlHeight
            mainForm.Show
            mainForm.cboAddress.Text = StartingAddress
            mainForm.timTimer.Enabled = True
            mainForm.brwWebBrowser.Navigate (StartingAddress)
     End If
    End If
    
    Exit Sub
    
ErrHandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "Template.ResizeView", "Template", vbExclamation)
    
    Exit Sub
End Sub

Private Sub Class_Initialize()
    '---create shared object used to show VB errors, etc.---
    Set gobjShared = New CWShared
    
    '---set the help file for this application---
    App.HelpFile = gobjShared.CWHelpFile
    
    '------------------------------------
    '  This statement is needed to debug
    '    this OLE Server.
    '------------------------------------
    #If ciDEBUG = 1 Then
        If cwbDebugLogon = False Then
            Exit Sub
        End If
    #End If
End Sub

Private Sub Class_Terminate()
    '---release reference to objects---
    Set gobjShared = Nothing
    Set cwobjCWComm = Nothing
    Set cwobjLogonEnviron = Nothing
    Set cwobjConfig = Nothing
    Set cwobjPatient = Nothing
    Set cwobjUser = Nothing
    
    Set mainForm = Nothing
    
    '********************************************************************
    '  TO DO:
    '    Add code here to release reference to all other global objects.
    '********************************************************************
    
    
    
    
    '----------------------------------------
    '  If debugging, make sure to "log off".
    '----------------------------------------
    #If ciDEBUG = 1 Then
        cwbDebugLogOff
    #End If
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
Public Sub Message(ByVal viMessageType As Integer, ByVal vsMessageText As String, ByVal vsSourceServerNCID As String)
    On Error GoTo ErrHandler
    
    '-------------------------------------------------------------------------------
    '  Take corresponding actions based upon the message "type".
    '
    '  NOTE: This method is required, but currently there are NO REQUIRED messages
    '        to process.
    '-------------------------------------------------------------------------------
    Dim Buff() As String


    Select Case viMessageType


    Case cwiGENERIC                 '-- parse the message to see what needs to be done
        Buff = Split(vsMessageText, "|")  '-- split the msg at the pipe
              Select Case UCase(Buff(0))   '-- based on the first part...
        
              Case "NAVIGATE"              '-- someone wants me to navigate to a web site
                   If Buff(1) = "" Then Exit Sub   '-- dummy check - should contain a URL
                   mainForm.Show
                   mainForm.cboAddress.Text = Buff(1)
                   mainForm.timTimer.Enabled = True
                   mainForm.brwWebBrowser.Navigate (Buff(1))
              Case Else
                   vsMessageText = "Unknown message type"
              End Select
    Case cwiABOUT
            '---inform CW that I am showing a modal form---
            cwobjCWComm.Message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_BROWSER_APPLICATION
            Dim objAbout As New CWAbout
            Call objAbout.Display(App, mainForm)
            Set objAbout = Nothing
            '---inform CW that I am done showing a modal form---
            cwobjCWComm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_BROWSER_APPLICATION
    End Select
    Exit Sub
    
ErrHandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "Template.Message", "Template", vbExclamation)
    
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
Public Sub Initialize(ByVal vlMyProgId As Long, robjComm As Object, robjLogonEnvironment As Object, robjConfig As Object, robjPatient As Object)
    On Error GoTo ErrHandler
    
    '---save the "ProgId" assigned to me by the CW Core---
    lMyProgId = vlMyProgId
    
    '---create references to passed in objects---
    Set cwobjCWComm = robjComm
    Set cwobjLogonEnviron = robjLogonEnvironment
    Set cwobjConfig = robjConfig
    Set cwobjPatient = robjPatient
    
    '---set global flags---
    cwbMainViewRunning = False
    
    '****************************************************************************
    '  TO DO:
    '    Add code here to create objects, initialize global variables, etc.
    '****************************************************************************
    
    sEnterpriseNCID = cwobjLogonEnviron.EnterpriseNCID
    sFacilityNCID = cwobjLogonEnviron.FacilityNCID
    sRoleNCID = cwobjLogonEnviron.RoleNCID
    sUserNCID = cwobjLogonEnviron.UserNCID
    
    Set mainForm = New frmBrowser
    bResult = cwobjConfig.GetValue(lMyProgId, "StartURL", "URL", StartingAddress)
'    bResult = cwobjConfig.RetrieveConfiguration(sEnterpriseNCID, sFacilityNCID, sRoleNCID, sUserNCID)
'    bResult = cwobjConfig.DeleteBin("UserFavorites")
'    bResult = cwobjConfig.SetValue("184921", sUserNCID, "UserFavorites", "Count", "0")
'    bResult = cwobjConfig.SaveChanges("184921", sUserNCID)
'    bResult = cwobjConfig.GetValue("184921", "UserFavorites", "Favorites", rsValue)
'    bResult = cwobjConfig.DeleteValue("184921", sUserNCID, "UserFavorites", "Favorites")

    Exit Sub

ErrHandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "Template.Initialize", "Template", vbExclamation)
    
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
Public Sub Properties(robjLogonEnvironment As Object, robjConfig As Object, ByVal vsOwnerNCID As String, ByVal vbCalledByDesigner As Boolean)
    On Error GoTo ErrHandler
    
    '*****************************************************************
    '  TO DO:
    '    Add code here to invoke this OLE Server's properties page.
    '*****************************************************************
    


    Exit Sub
    
ErrHandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "Template.Properties", "Template", vbExclamation)
    
    Exit Sub
End Sub

'**********************************************************************
'1. SCR 22482 - Add Internet web pages to favorites
'2. SCR 22483 - Delete Internet web pages from favorites
'3. SCR 22484 - Provide a file folder structure for favorite web pages
'4. SCR 22486 - Allow user to move web pages between favorites file folders
'5. SCR 22487 - Import web pages favorites into CHCSII web browser
'6. SCR 22490 - The system shall display web pages in CHCSII browser upon user selecting a web page from favorites list
'7. SCR 22491 - Enable user to print web page content
'**********************************************************************
Private Sub mnuFavAdd()
'SCR 22482 - Add Internet web pages to favorites
On Error Resume Next
  Dim ShHlper As ShellUIHelper
  Set ShHlper = New ShellUIHelper
  ShHlper.AddFavorite mainForm.brwWebBrowser.LocationURL, mainForm.brwWebBrowser.LocationName
CleanUp:
  Set ShHlper = Nothing
End Sub

Private Sub mnuFavImport()
'SCR 22487 - Import web pages favorites into CHCSII web browser
On Error GoTo ErrHandler
  
  'Get the favorite Import Folder and Push
  Dim CurrImportFavFldr As String
  Dim strSelectedImportFile As String

  CurrImportFavFldr = GetIEImportFavDir
  
  'Ask user to select a file to import
  mainForm.CommonDialog1.FileName = ""
  mainForm.CommonDialog1.Filter = "HTML (*.html; *.htm)|*.html;*.htm|All (*.*)|*.*"
  mainForm.CommonDialog1.DialogTitle = "Import Bookmarks HTML into Favorites"
  mainForm.CommonDialog1.ShowOpen
  
  If mainForm.CommonDialog1.FileName = "" Then GoTo CleanUp
  
  'Save the import file name
  strSelectedImportFile = mainForm.CommonDialog1.FileName
 
  'Ask user to select a folder to import to
  Dim ImportFolder As String
  Dim SelFolderDlg As SelectFolderDialog
  Set SelFolderDlg = New SelectFolderDialog
  SelFolderDlg.Label1.Caption = "Select destination folder where imported favorities will be placed. Please make sure folder is open to select."
  SelFolderDlg.Caption = "Browser for folder"
  Load SelFolderDlg
  SelFolderDlg.Show vbModal, mainForm
  ImportFolder = SelFolderDlg.strSelFldr
  Set SelFolderDlg = Nothing
 
  If ImportFolder = "" Then GoTo CleanUp
  
  'Change the favorite Import Folder temporary to user selection
  SetIEImportFavDir ImportFolder
  
  'Invoke the webbrowser import function
  Dim ShHlper As ShellUIHelper
  Set ShHlper = New ShellUIHelper
  ShHlper.ImportExportFavorites True, mainForm.CommonDialog1.FileName
  
  'mainForm.WebBrowser1.Navigate2 CStr("javascript:window.external.ImportExportFavorites(-1,'')")
  GoTo CleanUp
  
ErrHandler:
CleanUp:
  Set ShHlper = Nothing
  Set SelFolderDlg = Nothing
  'Restore/Pop the favorite Import Folder
  SetIEImportFavDir CurrImportFavFldr
End Sub

Private Sub mnuFavExport()
'SCR 22487 - Import web pages favorites into CHCSII web browser
On Error GoTo ErrHandler

  'Ask user to select a folder to export to
  Dim ExportFolder As String
  Dim SelFolderDlg As SelectFolderDialog
  Set SelFolderDlg = New SelectFolderDialog
  SelFolderDlg.Label1.Caption = "Select Export Favorites Source Folder"
  SelFolderDlg.Caption = "Browser for folder"
  Load SelFolderDlg
  SelFolderDlg.Show vbModal, mainForm
  ExportFolder = SelFolderDlg.strSelFldr
  Set SelFolderDlg = Nothing
 
  If ExportFolder = "" Then GoTo CleanUp

  'MsgBox Favfldr.Self.Path
  BuildExportHtml ExportFolder, True
  
  GoTo CleanUp
ErrHandler:
CleanUp:
  Set SelFolderDlg = Nothing
End Sub

Private Sub mnuFavPrint()
'SCR 22491 - Enable user to print web page content
  On Error Resume Next
  mainForm.brwWebBrowser.SetFocus
  SendKeys "^p"
End Sub

Private Function BuildExportHtml(CurrentFldr As String, Optional IsRoot As Boolean = False) As Boolean
'SCR 22487 - Import web pages favorites into CHCSII web browser
  Static iIndentLevel As Integer
  Static sExportFile As String
  Static fso As Scripting.FileSystemObject
  Dim sIndentStr As String
  'Dim fso As Scripting.FileSystemObject, fl As Scripting.File

  BuildExportHtml = False   'if we go thru the process successfully, we'll set this to True

  If IsRoot Then
    Set fso = New Scripting.FileSystemObject
    iIndentLevel = 0    'Set indent level
    sExportFile = ""    'Blank it

    'Ask user to select a file to export to
    mainForm.CommonDialog1.Filter = "HTML (*.html; *.htm)|*.html;*.htm|All (*.*)|*.*"
    mainForm.CommonDialog1.DialogTitle = "Export Favorites Into Bookmark HTML File"
    mainForm.CommonDialog1.ShowSave

    If mainForm.CommonDialog1.FileName = "" Then GoTo CleanUp

    'Check if that files exists and if user like to overwrite it
    Dim Result As VbMsgBoxResult
    If fso.FileExists(mainForm.CommonDialog1.FileName) Then
      Result = MsgBox(mainForm.CommonDialog1.FileName & " already exists. Overwrite?", vbOKCancel)
    End If
    If Result = vbCancel Then
      'BuildExportHtml = False       'Not overwriting file
      GoTo CleanUp                  'Cancel the operation
    End If

    sExportFile = mainForm.CommonDialog1.FileName  'Save the Export File

    'Now write the Bookmarks' header part out
    Open sExportFile For Output As #1
      Print #1, "<!DOCTYPE NETSCAPE-Bookmark-file-1>"
      Print #1, "<!--This is an automatically generated file."
      Print #1, "It will be read and overwritten."
      Print #1, "Do Not Edit! -->"
      Print #1, "<Title>Bookmarks</Title>"
      Print #1, "<H1>Bookmarks</H1>"
      Print #1, "<DL><p>"
    Close #1
  End If

  'Set indent string
  iIndentLevel = iIndentLevel + 1  'Each recursive call will increment indent level by 1
  Dim iCnt As Integer
  For iCnt = 0 To iIndentLevel
    sIndentStr = sIndentStr + constStrIndent
  Next iCnt
  
  Dim fdrFolder As Scripting.Folder
  Dim fdrSubFolder As Scripting.Folder
  Dim fdrFile As Scripting.File
  Set fdrFolder = fso.GetFolder(CurrentFldr)

  'Dim MyFldrItem As FolderItem
  For Each fdrSubFolder In fdrFolder.SubFolders  'Loop Through the sub folders
    'If MyFldrItem.IsFolder Then

      Open sExportFile For Append As #1
        'Writing out... <DT><H3 FOLDED ADD_DATE="{date}">{title}</H3>
        Dim fldrCreateDate As Variant
        'Dim subFldr As Folder
        'Set subFldr = fso.GetFolder(MyFldrItem.Path)
        fldrCreateDate = DateDiff("s", "00:00:00 01/01/70", fdrSubFolder.DateCreated)
        'Set subFldr = Nothing
        Print #1, sIndentStr & "<DT><H3 FOLDED ADD_DATE=""" & _
          fldrCreateDate & """>" & fdrSubFolder.Name & "</H3>"
        Print #1, sIndentStr & "  <DL><p>"
      Close #1
      iIndentLevel = iIndentLevel + 1
      If Not (BuildExportHtml(fdrSubFolder.Path)) Then GoTo CleanUp  'Drill down to next Level
      iIndentLevel = iIndentLevel - 1
      Open sExportFile For Append As #1
        Print #1, sIndentStr & "  </DL><p>"
      Close #1
  Next fdrSubFolder

  For Each fdrFile In fdrFolder.Files
  
      If fdrFile.Type = "Internet Shortcut" Or _
         fdrFile.Type = "InternetShortcut" Then
      
      'If MyFldrItem.Type = "Internet Shortcut" Then

'        Dim f As Scripting.File
        Dim s As Variant

        'Set f = fso.GetFile(MyFldrItem.Path)
        s = fdrFile.DateCreated
        Dim CreateDt As Variant
        Dim LstAccDt As Variant
        Dim LstModDt As Variant

        CreateDt = DateDiff("s", "00:00:00 01/01/70", fdrFile.DateCreated)
        LstAccDt = DateDiff("s", "00:00:00 01/01/70", fdrFile.DateLastAccessed)
        LstModDt = DateDiff("s", "00:00:00 01/01/70", fdrFile.DateLastModified)

'Writing out...  <DT><A HREF="{url}" ADD_DATE="{date}" LAST_VISIT="{date}" LAST_MODIFIED="{date}">{title}</A>
        Open sExportFile For Append As #1
          Print #1, sIndentStr & "<DT><A HREF=""" & _
            GetURLfromURLFile(fdrFile.Path) & _
            """ ADD_DATE=""" & CreateDt & _
            """ LAST_VISIT=""" & LstAccDt & _
            """ LAST_MODIFIED=""" & LstModDt & """>" & Left(fdrFile.Name, (Len(fdrFile.Name) - 4)) & "</A>"
        Close #1

        'Set f = Nothing

      End If 'MyFldrItem.Type = "Internet ShortCut"
  Next fdrFile
  If IsRoot Then MsgBox "Successfully exported favorities file: " & sExportFile
  BuildExportHtml = True     'Success
  GoTo CleanUp
ErrHandler:
CleanUp:
  iIndentLevel = iIndentLevel - 1
  Set fdrFile = Nothing
  Set fdrFolder = Nothing
  Set fdrSubFolder = Nothing
  If IsRoot Then
    If sExportFile <> "" Then
      Open sExportFile For Append As #1
        Print #1, "</DL><p>"
      Close #1
    End If
    Set fso = Nothing
  End If    'IsRoot

End Function


