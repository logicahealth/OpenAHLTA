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
Attribute VB_Name = "CITA_Init"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit


Implements ICHCSII_CoreAPI2
Implements ICHCSII_CoreAPI3
Implements I_RSTransfer

Private sUser As String
Private sUserPW As String

Private bNewPatient As Boolean

Private oDMBE As IDMBE.I_DMBE
Private ImportExport As ImportExport
Private iTMIPStatus As TMIPDataRetrievalResult


'    This routine builds recordsets for all the required tables
'    Used by TMIP

Public Function BuildImm_Rst(sID As String) As ADODB.Recordset
  Dim rst As ADODB.Recordset
  Dim rstExport As ADODB.Recordset
  Dim rstTemp As ADODB.Recordset
  Dim sSQL As String
'  Dim rstObject As IRSObject.I_RSObjects
  Dim lVaccE As Long
 
 
  
  On Local Error GoTo Func_Err
  
  
  'This will be changed when I start using UnitNumber as my primarykey
  'sID = sID & "20"
    
    
  'Create ado recordset to hold data for export
  CreateRst Theater_Export, rstExport
   
  '>> Begin:SCR #22584;   Developer: Brian Mowbray 03/29/2002 04:13 PM
  sSQL = "SELECT Count(*) As Nbr FROM Imm_People_Vaccine WHERE SPON_SSAN_FMP = '" & sID & "'"
  Set rst = gobjSQL.OpenRecordset(sSQL, adOpenDynamic, adLockBatchOptimistic)
  
  If rst Is Nothing Then
    Set BuildImm_Rst = Nothing
    iTMIPStatus = No_Data
    GoTo Func_Exit
  Else
    If rst.BOF And rst.EOF Then
    Else
      If rst![Nbr] <= 0 Then
        iTMIPStatus = No_Data
        GoTo Func_Exit
      End If
    End If
  End If
  '<< End: SCR #22584;
        
    
  'Get all immunization data from Imm_People_Vaccine table
  sSQL = "SELECT * FROM Imm_People_Vaccine WHERE SPON_SSAN_FMP = '" & sID & "' AND TMIP_STATUS IS NULL"
  Set rst = gobjSQL.OpenRecordset(sSQL, adOpenDynamic, adLockBatchOptimistic)
  If rst Is Nothing Then
    Set BuildImm_Rst = Nothing
    '>> Begin:SCR #22584;   Developer: Brian Mowbray 03/29/2002 04:14 PM
    iTMIPStatus = No_Data
    '<< End: SCR #22584;
    GoTo Func_Exit
  End If
  
  If rst.BOF And rst.EOF Then
    Set BuildImm_Rst = Nothing
    '>> Begin:SCR #22584;   Developer: Brian Mowbray 03/29/2002 04:14 PM
    iTMIPStatus = All_Data_Have_Been_Sent
    '<< End: SCR #22584;
    GoTo Func_Exit
  Else
    
    'Start looping through records for all immunizations found
    Do Until rst.EOF
      rstExport.AddNew
      UpdateRecord rst, rstExport
      
'Get all vaccine specific info from tblVaccines
AddVaccineInfo:
      
      sSQL = "SELECT Vaccine, HL7_ID, CPT, ICD, NbrInSeries, Schedule, Booster, ReportHeader, MPF_DIN, GenericName, Display, Vacc_NCID  FROM tblVaccines WHERE VaccineID = " & CInt(rst![VaccineID])
      Set rstTemp = gobjSQL.OpenRecordset(sSQL, adOpenDynamic, adLockBatchOptimistic)
      If rstTemp.BOF And rstTemp.EOF Then
        GoTo AddVaccExempt
      Else
        UpdateRecord rstTemp, rstExport
      End If
          
'Get all vaccine exemption information
AddVaccExempt:
      
      lVaccE = utNullToZero(rst![Vacc_Exempt])
      sSQL = "Select * FROM tblVacc_Exempt WHERE Vacc_Exempt = " & lVaccE
      Set rstTemp = gobjSQL.OpenRecordset(sSQL, adOpenDynamic, adLockBatchOptimistic)
      If rstTemp.BOF And rstTemp.EOF Then
        GoTo AddProvider
      Else
        UpdateRecord rstTemp, rstExport
      End If

'Get all vaccine provider information
AddProvider:
      
      If Len(rst![ProviderID] & vbNullString) <> 0 Then
        sSQL = "Select * FROM IMM_Provider WHERE IMM_ProviderID = '" & rst![ProviderID] & "'"
        Set rstTemp = gobjSQL.OpenRecordset(sSQL, adOpenDynamic, adLockBatchOptimistic)
        If rstTemp.BOF And rstTemp.EOF Then
          GoTo AddMfg
        Else
          UpdateRecord rstTemp, rstExport
        End If
      End If
      
'Get all vaccine manufacturer information
AddMfg:

      If Len(rst![Mfg_Code] & vbNullString) <> 0 Then
        sSQL = "Select * FROM tblVaccine_MFG WHERE MFG_Code = '" & rst![Mfg_Code] & "'"
        Set rstTemp = gobjSQL.OpenRecordset(sSQL, adOpenDynamic, adLockBatchOptimistic)
        If rstTemp.BOF And rstTemp.EOF Then
          GoTo AddStatus
        Else
          UpdateRecord rstTemp, rstExport
        End If
      End If
      
'Get patients status
AddStatus:

      'If Len(rst![PT_Status] & vbNullString) <> 0 Then
      '  sSQL = "Select * FROM tblPatientStatus WHERE Status = '" & rst![PT_Status] & "'"
      '  Set rstTemp = gobjSQL.OpenRecordset(sSQL, adOpenDynamic, adLockBatchOptimistic)
      '  If rstTemp.BOF And rstTemp.EOF Then
      '    GoTo AddService
      '  Else
      '    UpdateRecord rstTemp, rstExport
      '  End If
      'End If
      
'Get patients service
AddService:

      'If Len(rst![PT_SERVICE] & vbNullString) <> 0 Then
      '  sSQL = "Select * FROM tblService WHERE PT_Service = '" & rst![PT_SERVICE] & "'"
      '  Set rstTemp = gobjSQL.OpenRecordset(sSQL, adOpenDynamic, adLockBatchOptimistic)
      '  If rstTemp.BOF And rstTemp.EOF Then
      '    GoTo NextRecord
      '  Else
      '    UpdateRecord rstTemp, rstExport
      '  End If
      'End If

NextRecord:
    
      rst.MoveNext
    Loop
  End If
  rstExport.MoveFirst
  
  'Use to get temporary xml file for testing, not for use in
  On Error Resume Next
  Kill App.Path & "\ImmunizationsExport.xml"
  On Local Error GoTo Func_Err
  rstExport.Save App.Path & "\ImmunizationsExport.xml", adPersistXML
'  Set rstObject = New RSObject_Ops
'  With rstObject
'     .TableName = "Immunization"
'     Set .DBRecordSet = rstExport
'  End With
  '...Add to collection NOTE key this quarantees uniqueness and _
      gives future capabiltiy for multiple patients being downloaded _
      at once.
      
  
  Set BuildImm_Rst = rstExport
  '>> Begin:SCR #22584;   Developer: Brian Mowbray 03/29/2002 04:14 PM
  iTMIPStatus = SUCCESS
  '<< End: SCR #22584;
  
Func_Exit:
  On Error Resume Next
  rst.Close
  Set rst = Nothing
'  rstExport.Close
'  Set rstExport = Nothing
  rstTemp.Close
  Set rstTemp = Nothing
'  Set rstObject = Nothing
'  Set ImportExport = Nothing
  Exit Function

Func_Err:
   ErrMsgBox "CITA_Init.BuildImm_Rst", Error$, Err, "SSN : " & sID
   Debug.Print Err.Number, Err.Description
   '>> Begin:SCR #22584;   Developer: Brian Mowbray 03/29/2002 04:14 PM
   iTMIPStatus = Application_Error
   '<< End: SCR #22584;
   Resume Func_Exit
End Function






Private Function ConvertService(myService As String) As String
'we need to convert the service code *FROM* CHCSII's idea of a service *TO* CITA's.
'CITA thinks:
'  "A" = "Army"
'  "B" = "Navy"
'  "C" = "Air Force"
'  "D" = "Marine"
'  "E" = "Coast Guard"
'  "F" = "PHS"
'  "G" = "NOAA"
'  "O" = "Other"
'CHCSII thinks:
'a=army
'n=usn
'f=usaf
'm=usmc
'c=uscg
'k=other
'b=noaa
'p=phs
'
  Select Case UCase(myService)
    Case "A"
      ConvertService = "A"
    Case "N"
      ConvertService = "B"
    Case "F"
      ConvertService = "C"
    Case "M"
      ConvertService = "D"
    Case "C"
      ConvertService = "E"
    Case "P"
      ConvertService = "F"
    Case "B"
      ConvertService = "G"
    Case "K", "R"
      ConvertService = "O"
    Case Else
      ConvertService = ""
  End Select
End Function

Private Function ConvertStatus(myStatus As String) As String
'we need to convert the status code *FROM* CHCSII's idea of a status *TO* CITA's.
'CITA thinks:
'"Active Duty"= "AD"
'"Cadet"= "CAD"
'"Civilian"= "CIV"
'"Contractor"= "CON"
'"Family Member"= "DEP"
'"Foreign National"= "FN"
'"Guard"= "NG"
'"Other" = "OTH"
'"Retired" = "RET"
'"Reserves"= "RV"
'CHCSII thinks:
'11=Active Duty
'15=National Guard
'??=Reserves
'41=Dependent
'61=Unknown
  Select Case UCase(myStatus)
    Case "11", "21"
      ConvertStatus = "AD"
    Case "15"
      ConvertStatus = "NG"
    Case "12"
      ConvertStatus = "RV"
    Case "25", "27", "28", "29", "41", "43", "45", "47", "48"
      ConvertStatus = "DEP"
    Case "31", "22", "23", "32", "33"
      ConvertStatus = "RET"
    Case "14", "21"
      ConvertStatus = "CAD"
    Case "51", "52", "53", "54", "55", "56", "57", "59", "61", "64", "65", "68", "69"
      ConvertStatus = "CIV"
    Case "71", "72", "73", "74", "75", "76", "77", "78", "79"
      ConvertStatus = "FN"
    Case "83", "84", "91"
      ConvertStatus = "OTH"
    Case Else
      ConvertStatus = ""
  End Select
End Function
Private Function convertUIC(myUIC As String) As String
'we need to get the UIC out of a pair of parentheses

Dim i As Integer
Dim j As Integer

i = InStr(myUIC, "(")
If i = 0 Then
  convertUIC = myUIC
Else
  j = InStr(myUIC, ")")
  If j < i Then
    convertUIC = ""
  Else
    convertUIC = Mid(myUIC, i + 1, j - (i + 1))
  End If
End If

End Function




Private Function GetFMP() As String
  
  Dim sFMP As String
  
  If Len(gobjPatient.FMP & vbNullString) = 0 Then
      sFMP = "00"
  Else
    sFMP = Format(gobjPatient.FMP, "00")
  End If
  
  GetFMP = sFMP

End Function

'<< SCR #37777;   Developer: Brian Mowbray 06/16/2003 04:22 PM
Public Function GetImmStatus(ByRef rst As ADODB.Recordset, Optional bRequiredOnly As Boolean = False) As Boolean
  Dim sErrorMsg As String
  Dim sStatus As String
  Dim sService As String
  Dim sSpon_ssn_fmp As String
  On Error GoTo Func_Err
  
  sErrorMsg = vbNullString
  
  glCurPat = gobjPatient.UnitNumber
  sSpon_ssn_fmp = GetSponSSN_FMP
  
  'Should not happen but good to check for anyway
  If glCurPat = 0 Then
    sErrorMsg = "No Patient Loaded"
    GoTo Func_Err
  End If
  
  'Need service info to properly calculate immunization needed
  If Len(gobjPatient.PatientCategory) > 2 Then
    sService = ConvertService(Left(gobjPatient.PatientCategory, 1))
    sStatus = ConvertStatus(Mid(gobjPatient.PatientCategory, 2, 2))
  End If
  
  'This only works for current baseline, for now
  '<< SCR #42805; 42809;   Developer: Brian Mowbray 10/02/2003 03:23 PM
  If Not gbITT Then
  
    'This will get called is the patient is changed or the first trip
    If IMM_Tables Is Nothing Then
      Set IMM_Tables = New IMM_Tables
    End If
    
    'Load patient immunization data
    If Not PatientIMM Is Nothing Then
      With PatientIMM
            .Reset
            .GetPatientsVaccineSchedule False, sSpon_ssn_fmp, _
                                        convertUIC(gobjPatient.UnitIdentificationCode), _
                                        CVDate(gobjPatient.DateOfBirth), _
                                        CVDate(Now), _
                                        sStatus, _
                                        sService
      End With
      
      'Recalc the patients nextdue dates to ensure it reflects the current schedule
      If oData Is Nothing Then
        Set oData = New Imm_Data
      End If
      oData.Imm_People_Recalc sSpon_ssn_fmp, CVDate(gobjPatient.DateOfBirth), convertUIC(gobjPatient.UnitIdentificationCode)
      
      'PatientIMM is initialized so get the recordset and pass on
      Set rst = PatientIMM.BuildGenericRecordSet(bRequiredOnly)
    Else
      'PatientIMM not Initialized so create the empty generic rst and pass back
      sErrorMsg = "Error when loading Immunization data"
      GoTo Func_Err
    End If
  Else
    Set rst = oData.GetLatestGenericVaccines(sSpon_ssn_fmp, bRequiredOnly, gobjPatient.UnitIdentificationCode, CVDate(gobjPatient.DateOfBirth))
  End If
  
  GetImmStatus = True

Func_Exit:
  On Error Resume Next
  Exit Function

Func_Err:
  'Creates the generic recordset just like the function call above does
  If Err.Number = 0 Then
    sErrorMsg = Err.Description
  End If
  'Adds the error record to the recordset so the user knows an error occured
  CreateRst Generic_rst, rst
  With rst
    .AddNew
    ![GenericName] = Left$(sErrorMsg, 50) & vbNullString
    .Update
  End With
  Resume Func_Exit
  Resume 'Debug
End Function

'<< SCR #37777;   Developer: Brian Mowbray 06/16/2003 04:22 PM
Private Function GetSponSSN_FMP() As String

  Dim sFMP As String
  
  If Len(gobjPatient.FMP & vbNullString) = 0 Then
      sFMP = "00"
  Else
    sFMP = Format(gobjPatient.FMP, "00")
  End If
  
  GetSponSSN_FMP = RemoveDashes(CStr(gobjPatient.SponsorSSN)) & sFMP

End Function

'**************************************************************
'  Description: This method is called when a client wants to know if this OLE Server
'               supports a specific "view type".
'  Parameters:
'    viViewType - Contains the "type" of view being queried for.
'  Return Value:
'    Returns TRUE if this OLE Server supports the passed in "view type".
'    Returns FALSE if this OLE Server does not support the passed in "view type".
'************************************************************
Public Function QueryView(ByVal viViewType As Integer) As Boolean
    Select Case viViewType
        Case cwiREGULAR_VIEW_I
            QueryView = True    '--this view is NOT optional--
        Case cwiPATIENT_SUMMARY_VIEW_I
            QueryView = False
        Case cwiCLINICIAN_SUMMARY_VIEW_I
            QueryView = False
        Case cwiENCOUNTER_SUMMARY_VIEW_I
            QueryView = False
        Case Else   '--unknown "view type", not supported--
            QueryView = False
    End Select
    glCallingNode = 0
End Function
'*******************************************************************
'  Description: This method is called when the CW Core attempts to open a "view" provided by this
'               OLE server.
'  Parameters:
'    viViewType - Contains the "type" of view being requested by the CW Core to open.
'    vlParenthWnd - Contains the hWnd of the client area window to be used by this OLE Server.
'    rhMainForm - If "viViewType" is a "regular view", upon return this will contain
'                 the window "handle" of the main "regular" form for this OLE Server.
'  Return Value:
'    Returns TRUE if the "view" was opened successfully.
'    Returns FALSE if an error occurred openning the "view".
'************************************************************************
Public Function OpenView(ByVal viViewType As Integer, _
                         ByVal vlParenthWnd As Long, _
                               rhMainForm As Long) As Boolean
    Dim sMsg As String
    On Error GoTo ErrHandler
    '---perform actions based on the "view type"---
    If viViewType = cwiREGULAR_VIEW_I Then
        '---set global flags---
        cwbFirstRegularResize_I = True
        cwbMainViewRunning_I = True
        '---------------------------------------------------------------------------------------
        '  If this OLE Server needs it's own menu items, here is a good place to add these
        '    menu items onto the CW Core's main menu.
        '  NOTE: This is OPTIONAL.  You do NOT have to create menu items. Thats why they are
        '        commented out.
        '---------------------------------------------------------------------------------------
        '---add a menu item "Template" on CW's main menu---
        'ghTemplate = gobjComm.AddPopupMenuToCW("&Template", cwlMyProgId)
        '---add a menu item "Ole Server Menu Item" to the menu "Template"---
        'glOleMenuItem = gobjComm.AddMenuItem("&Ole Server Menu Item", ghTemplate, cwlMyProgId, "Template menu item")
        
        glMenuRefresh_I = gobjComm.AddActionMenuItem("Refresh", "Refresh.bmp", cwlMyProgId_I, "Recalculate Immunizations")

        
        '-----------------------------------------------------------------------------------------
        '  Reset the "parent window" for this OLE server's "main" regular form.  This is the main
        '    "trick" that makes the OLE server look like it is part of the CW Core.
        '  NOTE: This action invokes the main form's "Load" event, but the form is NOT shown.
        '        Make sure that you DO NOT perform a "Me.Show" in the form's "Load" event, as this
        '        will mess up the "SetParent" API call.  Make sure that you DO NOT "Show" the
        '        "main" regular form in this method, that will happen in the "ResizeView" method.
        '-----------------------------------------------------------------------------------------
        '---set parameter---
        '...This code required because CITA_Admin_Init could also
        '...Instantiate Imm_Tables and we only want 1 copy
        '...Similar code exists in the CloseDataObjects
        
        
          'SCR #20068
          'Begin:SCR #20521 - Immunizations - VB error 91 when log in as Imm_nurse - can't access module  Developer: Brian Mowbray 01/16/2002 03:51 PM
          gsMTF_ID = SayMTF_ID()
          If Len(gsMTF_ID) > 0 Then
            ' Begin:SCR #20852;   Developer: Brian Mowbray 05/28/2002 12:25 PM
            SetParent frmImm_People.hwnd, vlParenthWnd
            glCurPat = gobjPatient.UnitNumber
          
            rhMainForm = frmImm_People.hwnd
            ' End: SCR #20852;
          
            '<< SCR #42805; 42809;   Developer: Brian Mowbray 10/02/2003 03:23 PM
            If Not gbITT Then
              If Imm_TablesCounter <= 0 And IMM_Tables Is Nothing Then
                Set IMM_Tables = New IMM_Tables
               End If
              Imm_TablesCounter = Imm_TablesCounter + 1
            End If
          Else
            sMsg = "Please log off the application to complete the initialization process for the immunization module. You are currently setup as Facility Name/NCID: " & gobjLogon.FacilityName & "/" & gobjLogon.FacilityNCID & " and Clinic Name/NCID: " & gobjLogon.NursingDivisionName & "/" & gobjLogon.NursingDivisionNCID & ", you will not be able to use the immunizations module until you have restarted the application." & vbCrLf & "Immunization Module Initialization completed!"
            MsgBxARMd sMsg, vbInformation, "Immunization Module Initialization"
            utLogError "basCITA_Lite:SayMTF_ID", sMsg, 999
            ' Begin:SCR #20852;   Developer: Brian Mowbray 05/28/2002 12:24 PM
            '---set return value---
            OpenView = False
            Screen.MousePointer = vbDefault
            Exit Function
          ' End: SCR #20852;
          End If
          'End: SCR #20521 - Immunizations - VB error 91 when log in as Imm_nurse - can't access module
        
        glCallingNode = 0
        
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
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "CITA_Init.OpenView", "CHCSII_Immun", vbCritical)
    '---set return value---
    OpenView = False
    Exit Function
End Function
'****************************************************************
'  Description: This method is called when the CW Core closes a "view" provided by this OLE
'               Server.
'  Parameters:
'    viViewType - Contains the "type" of view that the CW Core is closing.
'************************************************************************
Public Sub CloseView(ByVal viViewType As Integer)
  '...T.Noren  08/17/01
  Dim CloseThemBoys As Boolean
    If viViewType = cwiALL_VIEWS_I Then
        If cwbMainViewRunning_I Then CloseThemBoys = True
    ElseIf viViewType = cwiREGULAR_VIEW_I Then
        If cwbMainViewRunning_I Then CloseThemBoys = True
    End If
    If CloseThemBoys Then
        '---close the "main view"---
        '---set global flags---
        cwbMainViewRunning_I = False
        '---this indicates that the CW Core is closing the "view", NOT this OLE Server---
        cwbCloseViewMyself_I = False
        '---close ALL forms that are associated with all "views"---
        Unload frmImm_People
        ' T.Noren  11/01/01 make sure we set objects to nothing that are not being used anymore
        ' Following instruction was moved to Terminate code
        CloseDataObjects
    End If
End Sub
'*******************************************************************************************
'  Description: This method is called when the user has selected a menu item (from the CW
'               Core's main menu), which is "owned" by this OLE Server.
'               NOTE: This is an OPTIONAL method.
'  Parameters:
'    vlMenuId - Contains the menu id of the menu item the user selected.
'*******************************************************************************************
Public Sub MenuItemSelected(ByVal vlMenuId As Long)
    On Error GoTo ErrHandler
    '----------------------------------------------------------------------------------
    '  NOTE: You don't have to create menu items, these are optional, that is why
    '        the code is commented out.
    '----------------------------------------------------------------------------------
    Select Case vlMenuId 'Button.Key
        Case glMenuRefresh_I
            SayStatus "Recalculate Immunizations"
            '<< SCR #42805; 42809;   Developer: Brian Mowbray 10/02/2003 03:23 PM
            If Not gbITT Then
              PatientIMM.GetPatientsVaccineSchedule True
            End If
            
            frmImm_People.ImmRecalc
    End Select

      
    Exit Sub
ErrHandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "CITA_Init.MenuItemSelected", "CHCSII_Immun", vbExclamation)
    Exit Sub
End Sub

'**********************************************************************
'  Description: This method gets called when a new patient has been selected from the CW Core.
'               NOTE: This is an OPTIONAL method.
'  Parameters:
'    none
'*************************************************************************
Public Sub PatientSelected()
    '---update the all running "views" with the new patient---
    If cwbMainViewRunning_I Then
        '---Set flag to refresh immunization data for the new patient---
        If glCurPat <> gobjPatient.UnitNumber Then
          bNewPatient = True
          glCurPat = gobjPatient.UnitNumber
          Init
        End If
    End If
    
    If cwbPatientSummaryViewRunning_I Then
        '---Set flag to refresh immunization data for the new patient---
        If glCurPat <> gobjPatient.UnitNumber Then
          bNewPatient = True
          glCurPat = gobjPatient.UnitNumber
          Init
        End If
    End If
End Sub
'***********************************************************
'  Description: This method is called when a change has occurred in the "client area"
'               used by this OLE Server, resulting in a need to resize a particular "view".
'  Parameters:
'    viViewType - Contains the "type" of view being requested to resize by the CW Core.
'    vlLeft - Contains the "left" coordinate for the OLE server.
'    vlTop - Contains the "top" coordinate for the OLE server.
'    vlWidth - Contains the "width" for the OLE server.
'    vlHeight - Contains the "height" for the OLE server.
'**************************************************************
Public Sub ResizeView(ByVal viViewType As Integer, ByVal vlLeft As Long, ByVal vlTop As Long, ByVal vlWidth As Long, ByVal vlHeight As Long)
    Dim sMsg As String
    Dim iCnt As Integer
    
    On Error GoTo ErrHandler
    '---reset global "client area" parameters---
    cwlLeft = vlLeft
    cwlTop = vlTop
    cwlWidth = vlWidth
    cwlHeight = vlHeight
    
    '---perform different actions based upon the "view type"---
    If viViewType = cwiREGULAR_VIEW_I Then
        If Not cwbFirstRegularResize_I Then
            '---------------------------------------------------------------------------------
            '  Reset the "WindowState" to "Normal", so the form can be resized.  This is
            '    needed because a window CANNOT be moved or resized if it is "minimized" or
            '    "maximized".
            '  NOTE: This generates a form "resize" event.
            '---------------------------------------------------------------------------------
            
            ' Begin:SCR #23107;   Developer: Brian Mowbray 04/17/2002 04:27 PM
            utCloseImmForms
            ' End: SCR #23107;

            frmImm_People.WindowState = vbNormal
            '---hide the form to so that annyoing "screen blinking"
            '  will not be seen while controls are moved and resized---
            frmImm_People.Visible = False
            '---resize form to fill entire "client area", NOTE: This generates a "resize" event---
            frmImm_People.Move cwlLeft, cwlTop, cwlWidth, cwlHeight
            frmImm_People.ImmRecalc
            frmImm_People.Visible = True
                
        Else    '--first "regular" resize, need to perform "initailization" procedures--
        
        
          'Begin:SCR #20521 - Immunizations - VB error 91 when log in as Imm_nurse - can't access module  Developer: Brian Mowbray 01/16/2002 03:52 PM
          If Len(gsMTF_ID) = 0 Then
            CloseView cwiREGULAR_VIEW_I
          Else
            '---reset flag so this code will not be executed again---
            cwbFirstRegularResize_I = False
            '----------------------------------------------------------------------------
            '  NOTE: Here is where you "Show" the "Regular View" form for the FIRST time.
            '----------------------------------------------------------------------------
            Init
          End If
          'End: SCR #20521 - Immunizations - VB error 91 when log in as Imm_nurse - can't access module
        End If
    End If
    Exit Sub
ErrHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "CITA_Init.ResizeView", "CHCSII_Immun", vbExclamation)
    Exit Sub
End Sub

'<< SCR #37777;   Developer: Brian Mowbray 06/16/2003 04:31 PM
Public Function SayVaccExemptName(lVaccExmptID As Long) As String
  On Error Resume Next
  SayVaccExemptName = SayExemptionText(lVaccExmptID)
End Function

Private Sub Class_Initialize()
    '--- COMMON TO MULTIPLE SCR-ARMs: MSolano 9/13/01 SCR-ARM-16110, 16130
    '--- Init routine should be called from first top level creatable coclass
    '--- It can be called multiple times, but will actually run only once.
    '--- Instrument all creatable classes in case dll's multiple entry
    '--- points are actually used
    Call ARM_InitAllTransactionsForImmunizations
    Set gobjShared = New CWShared
    App.HelpFile = gobjShared.CWHelpFile
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
  'These should only get set to nothing when the class terminates or a change patient occurs.  - B Mowbray
    
  '<< Begin:SCR #24709; 25036;   Developer: Brian Mowbray 06/24/2002 03:17 PM
  'gbChangePatient Boolean set when the Admin Module is in front and a change patient occurs. Otherwise, these will be set to
  'nothing when the class terminates and the Admin Module will no longer function
    
    If Not gbChangePatient Then
'These should not be destroyed because the admin module does not get destroyed when the patient is changed
'      Set gobjConfig = Nothing
'      Set gobjComm = Nothing
'      Set gobjLogon = Nothing
'      Set gobjUser = Nothing
'      Set gobjShared = Nothing
'      Set gobjSQL = Nothing
'      Set gobjDas = Nothing
      Set IMM_Tables = Nothing
    End If
    
    gbChangePatient = False
    Set PatientIMM = Nothing
'    Set gobjPatient = Nothing
    Set gobjAllergy = Nothing
'    CloseDataObjects
'>> End: SCR #24709; 25036;
End Sub

'*******************************************************
'  Description: This method is used by the CW Core to send messages to this OLE Server.  The
'               messages can be from the CW Core or from another OLE Server.
'  Parameters:
'    viMessageType - Contains the "type" of message to process.
'    vsMessageText - Contains any "data" assoicated with this message.
'    vsSourceServerNCID - Contains the NCID of the OLE Server that "sent" the message.
'****************************************************************************
Public Sub message(ByVal viMessageType As Integer, _
                   ByVal vsMessageText As String, _
                   ByVal vsSourceServerNCID As String)
    On Error GoTo ErrHandler
    '-------------------------------------------------------------------------------
    '  Take corresponding actions based upon the message "type".
    '  NOTE: This method is required, but currently there are NO REQUIRED messages
    '        to process.
    '-------------------------------------------------------------------------------
    
    Exit Sub
ErrHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "CITA_Init.Message", "CHCSII_Immun", vbExclamation)
    Exit Sub
End Sub

'***********************************************************************
'  Description: This method is called when the CW Core creates an "instance" of this OLE Server.  It
'               is used to create and initialize global variables that are shared across all
'               "views" supported by this OLE Server, i.e. variables that only need to be
'               initialized ONCE, upon object creation.
'  Parameters:
'    vlMyProgId - Contains the "ProgId" that the CW Core has assigned to this OLE Server.
'    robjComm - Contains a communications object that is used to handle all communication
'               between this OLE Server and the CW Core.
'    robjLogonEnvironment - Contains the logon environment object.  Contains information such
'                           as who logged on, where they logged on, etc.
'    robjConfig - Contains the configuration object that is used to get and set configuration
'                 information for a specific application.
'    robjPatient - Contains the patient object containing the currently selected patient.
'***************************************************************************
Public Sub Initialize(ByVal vlMyProgId As Long, _
                            robjComm As Object, _
                            robjLogonEnvironment As Object, _
                            robjConfig As Object, _
                            robjPatient As Object, robjUser As ICHCSII_User)
    Dim sMsg As String
    On Error GoTo ErrHandler
    'Saves which node called immunizations, 1 = CITA Admin Module, 2 = CITA Immunization Module
    
    '---save the "ProgId" assigned to me by the CW Core---
    cwlMyProgId_I = vlMyProgId
    '---create references to passed in objects---
    Set gobjComm = robjComm
    Set gobjLogon = robjLogonEnvironment
    Set gobjConfig = robjConfig
    Set gobjPatient = robjPatient
    Set gobjUser = robjUser
    Set PatientIMM = New Patient_Imm
    '---set global flags---
    cwbMainViewRunning_I = False
    
    Set gobjDas = New DasCtrl
    
    'Sets the database object to CHCSII or GEMS
    Set gobjSQL = SelectCHCSII_DAS()
    
    'Sets the ADODB Provider for the backend database
    SetADODBProvider
    
    'Set the DMIS ID
'  SCR 20852 1/25/2002 Brian Mowbray
    gsMTF_ID = SayMTF_ID()
    
    '<< Begin:SCR #24806;   Developer: Brian Mowbray 06/19/2002 05:23 PM
    ' Get the object reference from Core
    Set gobjAllergy = gobjComm.InitializeOLEServer(NCID_ALLERGIES_APPLICATION)
    '>> End: SCR #24806;
    
    

    
    Exit Sub
ErrHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "CITA_Init.Initialize", "CHCSII_Immun", vbExclamation)
    Exit Sub
End Sub
Friend Function Init() As Boolean
  Dim rstAllergies As ADODB.Recordset
  Dim comma As Integer
  Dim sSQL As String
  Dim objRecset As ADODB.Recordset
  Dim rstPers As ADODB.Recordset
  Dim colAllergies As Collection
  Dim sRank As String
  Dim sLogonUser As String
  Dim Service As String
  Dim Status As String
  Dim sRankNCID As String
  Dim sID As String
  Dim sFMP As String
  
  On Error GoTo Err_Handler
  Set oData = New Imm_Data
  If gobjSQL Is Nothing Then
    Set gobjSQL = SelectCHCSII_DAS()
  End If
  'Set the Current Logged in User
  sLogonUser = Replace(gobjLogon.UserName, ",", "")
  SetUser sLogonUser
  
  CreateRst Imm_People, rstPers
  comma = InStr(gobjPatient.Name, ",")
  With rstPers
    .AddNew
    If comma > 0 Then
      !PT_LAST_NAME = Left(gobjPatient.Name, (comma - 1))
      !PT_First_Name = Trim(Mid(gobjPatient.Name, (comma + 1)))
    Else
      !PT_LAST_NAME = gobjPatient.Name
    End If
    !PT_SSAN = gobjPatient.SSN
    !PT_DOB = gobjPatient.DateOfBirth
    'lIEN = gobjPatient.UnitNumber
    !PT_GENDER = gobjPatient.Sex
    !PT_Status = Left(gobjPatient.PatientCategory, 1)
    '<<Begin:SCR #20068 - Brian Mowbray 01/03/2002
    '<<SCR Description: Performance - Open Immunizations fails the ORD6 second req in the field
    'change to grade
    'sRank = StrConv(gobjPatient.Rank, vbProperCase)
    sRank = StrConv(gobjPatient.PayGrade, vbProperCase)
    '>>End: SCR #20068
    
    '<< Begin:SCR #37777;   Developer: Brian Mowbray 06/16/2003 04:23 PM
    sID = GetSponSSN_FMP
    sFMP = right$(sID, 2)
    '>> End: SCR #37777;
    
    If Len(gobjPatient.PatientCategory) > 2 Then
      !PT_Service = ConvertService(Left(gobjPatient.PatientCategory, 1))
      !PT_Status = ConvertStatus(Mid(gobjPatient.PatientCategory, 2, 2))
    End If
    
    Service = !PT_Service & vbNullString
    Status = !PT_Status & vbNullString
    
    '<< Begin:SCR #26306;   Developer: Brian Mowbray 08/14/2002 05:43 PM
    If Not gobjPatient.Sponsor Is Nothing Then
      sRankNCID = gobjPatient.Sponsor.GetProperty("1015", True)
    End If
                
    If Len(sRankNCID) <> 0 Then
  
      sSQL = "SELECT abbrev_rank from rank_abbrev WHERE ncid = '" & sRankNCID & "'"

      Set objRecset = gobjSQL.OpenRecordset(sSQL, adOpenDynamic, adLockBatchOptimistic)
      If Not (objRecset.EOF And objRecset.BOF) Then
        !PT_RANK = objRecset![ABBREV_RANK]
        objRecset.Close
      Else
        !PT_RANK = "UNK"
        objRecset.Close
      End If
    Else
      !PT_RANK = "UNK"
    End If
    '>> End: SCR #26306;
    
    !SPON_SSAN_FMP = sID
    
    If Len(gobjPatient.UnitIdentificationCode) > 2 Then
'      !Pt_Unit = convertUIC(gobjPatient.UnitIdentificationCode)
      '<< SCR #42779;   Developer: Brian Mowbray 10/21/2003 03:53 PM
      !Pt_Unit = Left$(gobjPatient.UnitIdentificationCode, 100)
    End If
    .Update
    
    '<< Begin:SCR #19937; 23657;   Developer: Brian Mowbray 06/11/2002 04:04 PM
    'Get a collection of allergies from the allergies module
    GetAllergiesFromCDR sFMP, !PT_SSAN, gobjPatient.UnitNumber, colAllergies
    '>> End: SCR #19937; 23657;
  
  End With

  
  If rstPers.BOF And rstPers.EOF Then
  Else
'Begin:SCR #20068 - Brian Mowbray 01/03/2002
'SCR Description: Performance Tunning - Open Immunizations fails the ORD6 second req in the field
'This process moved to the backend so the client does not have to call it except when in Theater
'    ClearExpirations rstPers!SPON_SSAN_FMP

    '<< SCR #42805; 42809;   Developer: Brian Mowbray 10/02/2003 03:23 PM
    If gbGEMS Or gbITT Then ClearExpirations rstPers!SPON_SSAN_FMP

'End: SCR #20068
    oData.GetPerson rstPers
    AddImmProvider gobjLogon.UserName, gobjLogon.UserNCID
    If bNewPatient Then
      frmImm_People.RefreshNewPatient
      Set frmImm_People.Allergies = colAllergies
      '<< SCR #42805; 42809;   Developer: Brian Mowbray 10/02/2003 03:23 PM
      If Not gbITT Then
        With PatientIMM
          .Reset
          .GetPatientsVaccineSchedule False, sID, _
                                      convertUIC(gobjPatient.UnitIdentificationCode), _
                                      CVDate(gobjPatient.DateOfBirth), _
                                      CVDate(Now), _
                                      Status, _
                                      Service
        End With
      End If
      Set frmImm_People.Person = rstPers
      bNewPatient = False
      
    Else
      Set frmImm_People.Allergies = colAllergies
      '<< SCR #42805; 42809;   Developer: Brian Mowbray 10/02/2003 03:23 PM
      If Not gbITT Then
        With PatientIMM
          .Reset
          .GetPatientsVaccineSchedule False, sID, _
                                      convertUIC(gobjPatient.UnitIdentificationCode), _
                                      CVDate(gobjPatient.DateOfBirth), _
                                      CVDate(Now), _
                                      Status, _
                                      Service
        End With
      End If
      Set frmImm_People.Person = rstPers
      frmImm_People.Show
      '<< Begin:SCR #39477;   Developer: Brian Mowbray 08/06/2003 11:03 AM
      frmImm_People.SizeVaccineGrid
      '>> End: SCR #39477;
      
    End If
  End If
  Init = True
Func_Exit:
  On Error Resume Next
  objRecset.Close
  Set objRecset = Nothing
  Exit Function
    
Err_Handler:
  ErrMsgBox "CITA_Init:Init", Error$, Err
  Resume Func_Exit
  Resume 'Debug
End Function


'**************************************************************
'  Description: This method is called by a client when they want to modify the "properties", or the
'               configuration of this OLE Server.
'               NOTE: This is an OPTIONAL method.
'  Parameters:
'    robjLogonEnvironment - Contains the logon environment object.
'    robjConfig - Contains the configuration object that is used to get and set configuration
'                 information for a specific application.
'    vsOwnerNCID - Contains the NCID of the "owner" to use in storing modifications to the configuration.
'    vbCalledByDesigner - TRUE if this method is being called by the CW Designer, FALSE otherwise.
'****************************************************************
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
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "CITA_Init.Properties", "CHCSII_Immun", vbExclamation)
    Exit Sub
End Sub


Private Function I_RSTransfer_Transfer(DMBEObj As IDMBE.I_DMBE) As Variant
 

   '     This routine will take records from:
   '
   '           1.  The outside world anything other than CDR & LDDB
   '                and save them to LDDB
   '                It will then move to CDR from LDDB if the original
   '                request was from the outside world to CDR
   
   '
   '           2.  The inside world either CDR or LDDB
   '               and save them to the outside world
   '               CDR and LDDB (if it exists) will be Identical after
   '               running this routine
   

 On Error GoTo haveerror

 Dim patient As Variant
 Dim Data_LDDB As Boolean
 Dim Count As Long
 Dim MergeWithLDDB As Boolean
 Dim GoodReturnCode As Long         ' returns 0 if good to be compatible with DMBE
 Dim OperatorResponse As VbMsgBoxResult
 Dim OrigUseLocalDB As Boolean
 
 '...Kluge fix for not logged on to both databases
 Dim NotBothCDRandLDDB As Boolean
 
 '...Save Original Value
 OrigUseLocalDB = gobjShared.UseLocalDB
  
 NotBothCDRandLDDB = Not gobjShared.CDRConnected Or Not gobjShared.LocalDBConnected
     
 Set oDMBE = DMBEObj
 
 '...Assume a bad deal
 GoodReturnCode = -15  ' no records for dataset
 
 
 patient = BuildPatientID(oDMBE)
    
 If Len(patient & vbNullString) = 0 Then
     GoTo exitRoutine
 End If
   
 '...Module defined variable
 Set ImportExport = New ImportExport
 
 With oDMBE
      
    If .Source = eDM_LDDB And Not (gobjShared.LocalDBConnected Or gobjShared.MySQLDBConnected) Then
       MsgBxARMd "Local Database (LDDB) not Connected", vbInformation + vbOKOnly, "Terminating Request"
       Exit Function
    End If

    If .Source = eDM_CDR And Not gobjShared.CDRConnected Then
       MsgBxARMd "CHCSII Database (CDR) not Connected", vbInformation + vbOKOnly, "Terminating Request"
       Exit Function
    End If
      
    If .Destination = eDM_LDDB And Not (gobjShared.LocalDBConnected Or gobjShared.MySQLDBConnected) Then
       MsgBxARMd "Local Database (LDDB) not Connected", vbInformation + vbOKOnly, "Terminating Request"
       Exit Function
    ElseIf .Destination = eDM_CDR And Not gobjShared.CDRConnected Then
       MsgBxARMd "CHCSII Database (CDR) not Connected", vbInformation + vbOKOnly, "Terminating Request"
       Exit Function
    End If

    '...Where will data come from
    Select Case .Source
    
      Case eDM_LDDB, eDM_CDR    '...Get Data from CDR or LDDB
          
           Data_LDDB = (.Source = eDM_LDDB)
           
           '...Patient List should have only 1 entry in it _
               however, if they decide to do multiple patients _
               at a time this code can handle it
'            patient = .PatientObj.UnitNumber
'           For Each Patient In .PatientList
           
            '...Reset Merged Records in Memory if necessary
'            Erase AllRecordsInMemory
'            NumAllRecordsInMemory = 0
           
            '...Indicate not merging To LDDB
            MergeWithLDDB = False
            
            '...Always bring in the most recent data if CDR connected
            If gobjShared.CDRConnected And (Not Data_LDDB And NotBothCDRandLDDB _
               Or (Data_LDDB And Not NotBothCDRandLDDB And .Destination = eDM_CDR) Or .Source = eDM_CDR) Then
                '...Get rquired records in CDR (user must already be logged on)
                If BuildRecordSetsGEMSorCDR(CStr(patient), oDMBE, "CDR") Then
                  '...If LDDB Exists then always Update LDDB
                  If gobjShared.LocalDBConnected Then
                   If .Destination <> eDM_LDDB Then
'                   OperatorResponse = MsgBxARMd("Obtained Records from CDR..." & vbCrLf & "Do you wish to merge records to LDDB also?", vbInformation + vbYesNoCancel, "Update LDDB")
                    ' Remove this statement and undo above comment if sync both lddb and cdr
                     OperatorResponse = vbNo
                   Else
                     OperatorResponse = vbYes
                   End If
                   Select Case OperatorResponse
                     Case vbYes
                       MergeWithLDDB = True
                     Case vbNo
                     Case Else
                       '...Clean up stuff here before exiting
                       GoTo getout
                    End Select
                  End If
                End If
            End If

            If Data_LDDB Or MergeWithLDDB Then
              '...Get required records in LDDB (User must already be logged on)
              BuildRecordSetsGEMSorCDR CStr(patient), oDMBE
            End If
            
            If MergeWithLDDB Then
              '...Save Data to LDDB
              MSaveDataToCDRorGEMS CStr(patient), False
              
'              '...Put Records in Memory
'              MergeRecords CStr(patient), True
              
              '...Update CDR
              MSaveDataToCDRorGEMS CStr(patient), True
              
              '...remove all recordsets built
              ImportExport.ClearRSCollection
              
              '...Get rquired records in LDDB
              BuildRecordSetsGEMSorCDR CStr(patient), oDMBE
              
            End If
           
'           Next Patient
 
          
          '...Where does this data go
'          If ImportExport.RecordSetsTransfered <> 0 Then
          
                Select Case oDMBE.Destination
                
                   Case eDM_CDR
                     If Not MergeWithLDDB Then
                      If Not gobjShared.CDRConnected Then
                        MsgBxARMd "CHCSII Database (CDR) not Connected", vbInformation + vbOKOnly, "Terminating Request"
                        Exit Function
                      Else
                        '...Update CDR
                        MSaveDataToCDRorGEMS CStr(patient), True
                        GoodReturnCode = 0
                      End If
                     End If
                       
                   Case eDM_LDDB
                      ' Process already completed
                      GoodReturnCode = 0
                      Exit Function
                       
                   Case Else            ' Outside world
                   
                      '>> Begin:SCR #22584;   Developer: Brian Mowbray 03/29/2002 04:17 PM
                      'GoodReturnCode = ImportExport.ExecuteDMBE(DMBEObj)
                      GoodReturnCode = ImportExport.ExecuteDMBE(DMBEObj, iTMIPStatus)
                      '<< End: SCR #22584;
                      
                End Select
            
'          End If
       
      Case Else      '...Get data from Outside World

          '...Get Records from outside world and save in either CDR or LDDB
          If (ImportExport.ExecuteDMBE(DMBEObj)) = 0 Then
              '...0 is a valid return code from Execute

               Data_LDDB = (.Destination = eDM_LDDB)
              
               '...Patient List should have only 1 entry in it _
                   however, if they decide to do multiple patients _
                   THIS CODE WILL NOT WORK _
                     Changes Required _
                        1.  Copy of TransferRecordSets must be Created _
                        2.  IRSObject needs to have a Variable _
                            Indicating the patient associated with _
                            that recordset and Table so that the _
                            routines SaveREcordSetsXXX can load the _
                            correct Recordset associated with a given _
                            Table.
                        
'               For Each Patient In .PatientList

                '...Read records from LDDB
                If BuildRecordSetsGEMSorCDR(CStr(patient), oDMBE) Then
                
                      '...Update LDDB
                      If gobjShared.LocalDBConnected And Data_LDDB Then
                    
                       '...Save Data to LDDB
                       MSaveDataToCDRorGEMS CStr(patient), False, True
                       
                       '...Clear record sets
                       '...In the future need to make a copy _
                           of TransferRecordsets
                       ImportExport.ClearRSCollection
                      
                      End If
                 
                 End If
                 
                 '...If CDR connected update CDR
                 
                '<< Begin:SCR #24298;   Developer: Brian Mowbray 06/24/2002 01:02 PM
                'This was causing an error. This should code should not be execuated because there is no current functionality
                'to communicate with the CDR when it is the destination. The CDR is currently always the source.
                'If gobjShared.CDRConnected And Not NotBothCDRandLDDB Then
                If gobjShared.CDRConnected And (.Destination = eDM_CDR) Then
                '>> End: SCR #24298;
                
                    '...Get Data from CDR
                    BuildRecordSetsGEMSorCDR CStr(patient), oDMBE, "CDR"
                    If gobjShared.LocalDBConnected Then
                      '...Get rquired records in LDDB
                      BuildRecordSetsGEMSorCDR CStr(patient), oDMBE
                    End If
                    
'                    '...Merge Records in Memory
'                    MergeRecords CStr(patient), True
                
                    '...Update CDR
                    MSaveDataToCDRorGEMS CStr(patient), True
                    
                 End If
              
                '...Need to restore TRansferRecordsets Here
                '   If we do multiple patients
               
'               Next Patient
               
               '...Indicate everything is ok
               GoodReturnCode = 0
               
          End If
          
    End Select
getout:
 End With
 
exitRoutine:

 I_RSTransfer_Transfer = GoodReturnCode
 
 Set oDMBE = Nothing
 Set ImportExport = Nothing
 
 '...Restore Original Value
 gobjShared.UseLocalDB = OrigUseLocalDB
 Set gobjSQL = Nothing
 Set gobjSQL = SelectCHCSII_DAS()
 
 Exit Function
 
haveerror:

 GoodReturnCode = Err.Number
 
 MsgBxARMd "Error Number:  " & CStr(Err.Number) & vbCrLf & _
        "Description:   " & Err.Description, vbExclamation, "Allergies"
        
 Debug.Print Err.Number, Err.Description
 Resume exitRoutine
 Resume
End Function


'***********************************************************************
'  Description: This method is called after the CW Core creates an "instance" of this OLE Server.  It
'               is used to create and initialize global variables that are shared across all
'               "views" supported by this OLE Server, i.e. variables that only need to be
'               initialized ONCE, upon object creation.
'  Parameters:
'    vlMyProgId - Contains the "ProgId" that the CW Core has assigned to this OLE Server.
'    robjComm - Contains a communications object that is used to handle all communication
'               between this OLE Server and the CW Core.
'    robjLogonEnvironment - Contains the logon environment object.  Contains information such
'                           as who logged on, where they logged on, etc.
'    robjConfig - Contains the configuration object that is used to get and set configuration
'                 information for a specific application.
'    robjPatient - Contains the patient object containing the currently selected patient.
'    robjUser - Security User object
'***************************************************************************
Private Sub ICHCSII_CoreAPI2_Initialize2(ByVal vlMyProgId As Long, _
                                               robjComm As Object, _
                                               robjLogonEnvironment As Object, _
                                               robjConfig As Object, _
                                               robjPatient As Object, _
                                               robjUser As Object)
    ' SCR 16643 T.Noren 11/01/01
'    Call Initialize(vlMyProgId, robjComm, robjLogonEnvironment, robjConfig, robjPatient)
'    Set gobjUser = robjUser
    Call Initialize(vlMyProgId, robjComm, robjLogonEnvironment, robjConfig, robjPatient, robjUser)
End Sub

'****************************************************************
'  Description: This method is called when the CW Core closes a "view" provided by this OLE
'               Server.  It differs from the CloseView method in that the child module is
'               is permitted to say 'No' to Core.
'  Parameters:
'    viViewType - Contains the "type" of view that the CW Core is closing.
'************************************************************************
Private Function ICHCSII_CoreAPI3_CloseView2(ByVal viViewType As Integer) As Boolean
  Call CloseView(viViewType)
  ICHCSII_CoreAPI3_CloseView2 = True
End Function




Public Function BuildRecordSetsGEMSorCDR(ByVal patient As String, DMBEObj As IDMBE.I_DMBE, Optional ByVal Source As String = "GEMS") As Boolean

  '    Since Immunizations recordsets are the same for CDR and GEMS
  '    there is no need to have a BuildRecordSetsGEMS and BuildRecordSetsCDR
  
  '    This routine builds recordsets for all the required tables
  
   On Local Error GoTo haveerror
   Dim sql As String
   Dim WherePart As String
   Dim SelectPart As String
   Dim FromPart As String
   Dim FileNumber As Integer
   Dim AdditionalSort As String
   Dim ReturnStatus As Boolean
   Dim xEnd As Integer
   Dim rs As ADODB.Recordset
   Dim rsbs As ADODB.Recordset
 
 ' -------------------------------------------------------------------
     Dim Fix_CDR_rs As Boolean          ' Used to fix CDR RS ...
 ' -------------------------------------------------------------------
   
   Set gobjSQL = Nothing
   If Source = "GEMS" Then
    gobjShared.UseLocalDB = True
   Else
    gobjShared.UseLocalDB = False
    Fix_CDR_rs = True                  ' Improper Field Definition
   End If
   
   '<< Begin:SCR #47279;   Developer: Brian Mowbray 12/18/2003 06:00 PM
   Set gobjSQL = SelectCHCSII_DAS(IIf(gobjShared.UseLocalDB, GEMS, Auto))
   '>> End: SCR #47279;
   
   '...Assume valid return code
   ReturnStatus = True
        
   If DMBEObj.Destination = eDM_TMIP Then
   
        '...Build Record Data Object used by Data Manager Back End
        Set rs = BuildImm_Rst(patient)
        
        If Source = "GEMS" Then
         ImportExport.SaveGEMSrs rs, patient, "Immunization"
        Else
         ImportExport.SaveCDRrs rs, patient, "Immunization"
        End If
        
   Else
   
        '...Do we include IMM_PEOPLE_VACCINE_DELETE
        xEnd = IIf(DMBEObj.Source <> eDM_CDR And Source <> "CDR", 3, 2)
           
        '...SQL statement used for all tables
        SelectPart = "Select * "
        WherePart = " Where SPON_SSAN_FMP = '" & patient & "'"
        
        '...Define all the tables here
        For FileNumber = 1 To xEnd
        
          AdditionalSort = ""
          Select Case FileNumber
          
            Case 1
              FromPart = "IMM_PEOPLE"
              AdditionalSort = "CreatedOn ASC"
            Case 2
              FromPart = "IMM_PEOPLE_VACCINE"
              AdditionalSort = "Vacc_Date ASC, vaccineid ASC, vacc_series ASC"
              
            Case 3
              FromPart = "IMM_PEOPLE_VACCINE_DELETE"
              AdditionalSort = "Vacc_Date ASC, vaccineid ASC, vacc_series ASC"
              
'            Case 4
'              FromPart = "IMM_CDR_WORK"
'              AdditionalSort = "CreatedOn ASC"
              
          End Select
          
          On Error Resume Next
          '...Get all data from this file for this patient
          Set rs = gobjSQL.OpenRecordset(SelectPart & "From " & FromPart & WherePart & " Order By " & AdditionalSort)
          
          If Err.Number <> 0 And rs Is Nothing Then
             Err.Clear
             '...Create Dummy recordset so merge operation will work"
             Set rs = New ADODB.Recordset
             rs.Fields.Append "SPON_SSAN_FMP", adVarChar, 50
             rs.Open
          ElseIf Err.Number <> 0 Then
            MsgBxARMd "Error Occurred reading File: " & FromPart & " in " & Source & vbCrLf & vbCrLf & _
                   "Error #: " & CStr(Err.Number) & vbCrLf & _
                   "Description: " & Err.Description, vbExclamation + vbOKOnly, "Error Reading Data"
            Err.Clear
          ElseIf Fix_CDR_rs Then
            If FromPart = "IMM_PEOPLE" Or FromPart = "IMM_PEOPLE_VACCINE" Then
                Select Case FromPart
                  Case "IMM_PEOPLE"
                    Set rsbs = convertBS(rs, Imm_People)
                  Case "IMM_PEOPLE_VACCINE"
                    Set rsbs = convertBS(rs, Imm_People_Vaccine)
                End Select
                Set rs = rsbs
                Set rsbs = Nothing
            End If
          End If
          On Local Error GoTo haveerror
          
          '...Build Record Data Object used by Data Manager Back End
          If Source = "GEMS" Then
            ImportExport.SaveGEMSrs rs, patient, FromPart
          Else
            ImportExport.SaveCDRrs rs, patient, FromPart
          End If
'          Set rs = Nothing
          
        Next FileNumber
   End If
exitRoutine:
   BuildRecordSetsGEMSorCDR = ReturnStatus
   Exit Function
haveerror:
   Debug.Print Err.Number, Err.Description
   ReturnStatus = False
   Resume exitRoutine:
   Resume

End Function

'Public Sub MergeRecords(ByVal patient As String, ByVal MemoryMerge As Boolean)
'
'End Sub

Public Function MSaveDataToCDRorGEMS(ByVal patient As String, ByVal MoveToCDR As Boolean, Optional ExtData As Boolean = False) As Boolean
 
  On Local Error GoTo haveerror
  Dim Count As Long
  Dim rsTo As ADODB.Recordset
  Dim rsFrom As ADODB.Recordset
  Dim ReturnStatus As Boolean
  Dim xFunc As String
  Dim NextRec As String
  Dim ToCrDate As ADODB.field
  Dim FromCrDate As ADODB.field
  Dim RecordsetType As Integer
  Dim xdone As Boolean
  Dim FilesToMove As Integer
  Dim FromExtKey As String
  Dim ToExtKey As String
  Dim Process As Boolean
  Dim TableName As String
  Dim sSQL As String
  Dim ToEOF As Integer
  Dim FromEOF As Integer
  
  '-------------------------------------------------------------------------
  
      ' Establish Dirrection of Move
      
    Set gobjSQL = Nothing
    gobjShared.UseLocalDB = Not MoveToCDR
    
    If MoveToCDR Then
       '...Data can only come from a Pic device or LDDB
       FromExtKey = IIf(ExtData, "", patient & "...LDDB\")
       
    Else
       '...Data can only come from a Pic device or CDR
       FromExtKey = IIf(ExtData, "", patient & "...CDR\")
       
    End If
    
    '...Always merge to CDR or LDDB
    ToExtKey = patient & IIf(MoveToCDR, "...CDR\", "...LDDB\")
    
    '...Connect to Destination
    '<< Begin:SCR #47279;   Developer: Brian Mowbray 12/18/2003 06:00 PM
    Set gobjSQL = SelectCHCSII_DAS(IIf(gobjShared.UseLocalDB, GEMS, Auto))
    '>> End: SCR #47279;
  
  '-------------------------------------------------------------------------


  '-------------------------------------------------------------------------

   ' process IMM_PEOPLE
   
   Set rsTo = ImportExport.GetATable(ToExtKey & "Imm_People")
   
   If rsTo.State = 0 Then
     ' No Table
     ToEOF = -1
   Else
     ' Set EOF Status
     ToEOF = IIf(rsTo.EOF, 1, 0)
   End If
   
   Set rsFrom = ImportExport.GetATable(FromExtKey & "Imm_People")
   If rsFrom.State = 0 Then
     ' No Table
     FromEOF = -1
   Else
     ' Set EOF Status
     FromEOF = IIf(rsFrom.EOF, 1, 0)
   End If
   
   If ToEOF <> 0 Then
    '...Patient doesn't exist in To Table
    If FromEOF <> 0 Then
      '...No IMM data in From Table  I.E.  no vaccinations either
      GoTo exitfunction
    Else
      sSQL = "Insert INTO IMM_PEOPLE " & uBuildInsertSqlFromRS(rsFrom)
      gobjSQL.ExecuteSQL sSQL
    End If
   Else
     If Not uIsRecordsEqual(rsFrom, rsTo) Then
      If FromEOF = 0 Then
       sSQL = "Update IMM_PEOPLE " & uBuildUpdateSqlFromRS("Imm_People", rsFrom, rsTo)
       gobjSQL.ExecuteSQL sSQL
      Else
        ' From Data does not exist
        GoTo exitfunction
      End If
     End If
   End If
   
  '-------------------------------------------------------------------------
  
     ' Delete Records from IMM_PEOPLE_VACINE
     ' that are in IMM_PEOPLE_VACINE_DELETE
 
   '...This is the table we are removing records from
   TableName = "Imm_People_Vaccine"
   
   Set rsTo = ImportExport.GetATable(FromExtKey & "Imm_People_Vaccine_Delete")
   If Not rsTo Is Nothing Then
    If rsTo.State = 0 Then
      ' No Table
      ToEOF = -1
    Else
      ' Set EOF Status
      ToEOF = IIf(rsTo.EOF, 1, 0)
    End If
   Else
    ToEOF = -1
   End If
   If ToEOF = 0 Then
    Dim ImmData As Imm_Data
    Set ImmData = New Imm_Data
    rsTo.MoveFirst
    On Error Resume Next
    Do While ToEOF = 0
    
      sSQL = "DELETE FROM " & TableName & " " & uBuildDeleteSqlFromRs(TableName, rsTo)
      gobjSQL.ExecuteSQL sSQL
      
      '...Need to clear error message if record does not exist on CDR
      If Err.Number <> 0 Then
        Err.Clear
      Else
       ImmData.DeleteImm_People_Vaccine rsTo![SPON_SSAN_FMP], rsTo![VaccineID], rsTo![Vacc_Date], False, rsTo
      End If
      
      rsTo.MoveNext
      ToEOF = IIf(rsTo.EOF, 1, 0)
    Loop
    
    rsTo.MoveFirst
    ToEOF = 0
    '--------------------------------------------------------------------------
       
       ' If data is being moved from LDDB To CDR
       ' Delete records existing in LDDB from Imm_People_Vaccine_Delete
       
    If MoveToCDR And Not ExtData Then
    
        
        '...Connect to LDDB
        Set gobjSQL = Nothing
        gobjShared.UseLocalDB = True
        '<< Begin:SCR #47279;   Developer: Brian Mowbray 12/18/2003 06:00 PM
        Set gobjSQL = SelectCHCSII_DAS(Auto)
        '>> End: SCR #47279;
        
        sSQL = "DELETE FROM Imm_People_Vaccine_Delete Where SPON_SSAN_FMP = '" & rsTo![SPON_SSAN_FMP] & "'"
        gobjSQL.ExecuteSQL sSQL
          
        
        '...Connect to CDR
        Set gobjSQL = Nothing
        gobjShared.UseLocalDB = False
        '<< Begin:SCR #47279;   Developer: Brian Mowbray 12/18/2003 06:00 PM
        Set gobjSQL = SelectCHCSII_DAS(CDR)
        '>> End: SCR #47279;
        
    ElseIf ExtData Then
       ' If data is being moved from EXT To LDDB
       ' Delete Table existing on External device
       
       ' Do Something HERE
       
    End If
    
    Set ImmData = Nothing
    
    '--------------------------------------------------------------------------
   End If
   
  '-------------------------------------------------------------------------
  
   On Local Error GoTo haveerror
   
   '  Imm_People_Vaccine_Delete automaticly gets updated
   '  when vaccine is deleted if vaccine is not deleted then
   '  that means it was added in theater and then deleted in theater
'   FilesToMove = IIf(MoveToCDR, 2, 1)
   FilesToMove = IIf(MoveToCDR, 1, 1)
  
   For RecordsetType = 1 To FilesToMove
   
     Process = True
     
     Select Case RecordsetType
        Case 1
            ' process IMM_PEOPLE_VACINE
            TableName = "Imm_People_Vaccine"
            
        Case 2
            ' process IMM_PEOPLE_VACINE_DELETE
            TableName = "Imm_People_Vaccine_Delete"
            
     End Select
   
     Set rsTo = ImportExport.GetATable(ToExtKey & TableName)
     
     Set rsFrom = ImportExport.GetATable(FromExtKey & TableName)
     
     If rsTo Is Nothing Or rsFrom Is Nothing Then Process = False
     
     If Process Then
     
        If rsFrom.State = 0 Then
          FromEOF = -1
        Else
          FromEOF = IIf(rsFrom.EOF, 1, 0)
        End If
        
        If rsTo.State = 0 Then
         ToEOF = -1
        Else
         ToEOF = IIf(rsTo.EOF, 1, 0)
        End If
        
        xdone = False
     
        With rsTo
          
          If ToEOF = 0 Then
            Set ToCrDate = .Fields("Vacc_Date")
            .MoveFirst
          End If
          If FromEOF = 0 Then
            Set FromCrDate = rsFrom.Fields("Vacc_Date")
            rsFrom.MoveFirst
          End If
          Do
           '...more records in To Table
           If ToEOF <> 0 Then
            ' NO
            If FromEOF <> 0 Then
              xdone = True
              xFunc = ""
              NextRec = 0
            Else
             xFunc = "ADD"
             NextRec = 1
            End If
           Else
             '...more records in From
             If FromEOF <> 0 Then
              ' No
              ' Skip rest of records in To Table
              xdone = True
              xFunc = ""
              NextRec = 0
              
             Else
               ' Is To Table record CREATED more recently
               If ToCrDate > FromCrDate Then
               ' YES
                 xFunc = "ADD"
                 NextRec = 1
                 
               ' Is To Table Older
               ElseIf ToCrDate < FromCrDate Then
               ' YES
                 xFunc = ""
                 NextRec = 2
                 
               ' Records created on same date
               Else
               
                ' Date Created is Equal
               
                If Not uIsRecordsEqual(rsFrom, rsTo) Then
                
                  ' Records Don't Match
                
                  If .Fields("VACCINEID").Value < rsFrom.Fields("VACCINEID") Then
                       xFunc = ""          ' Keep record in To Table
                       NextRec = 2         ' Move to next record from To Table
                  ElseIf .Fields("VACCINEID").Value > rsFrom.Fields("VACCINEID") Then
                       xFunc = "ADD"       ' Keep record to To Table
                       NextRec = 1         ' Move to next record from From Table
                  Else
                  
                    ' VaccineID are Equal
                    
                    If .Fields("VACC_SERIES").Value < rsFrom.Fields("VACC_SERIES") Then
                       xFunc = ""          ' Keep record in To Table
                       NextRec = 2         ' Move to next record from To Table
                    ElseIf .Fields("VACC_SERIES").Value > rsFrom.Fields("VACC_SERIES") Then
                       xFunc = "ADD"       ' Keep record to To Table
                       NextRec = 1         ' Move to next record from From Table
                    Else
                    
                      '...Series are Equal
                      
                      If .Fields("LastEdited").Value < rsFrom.Fields("LastEdited") Then
                        xFunc = "UPDATE"          ' Use From Table Record
                        NextRec = 3         ' Move to next record from To Table
                      ElseIf .Fields("LastEdited").Value > rsFrom.Fields("LastEdited") Then
                        xFunc = ""          ' Keep record in To Table
                        NextRec = 3         ' Move to next record from To Table
                      Else
                      
                       '...Last Edit date is Equal
                       '...Will always keep the most recent Vac-Date
                       
                       If .Fields("VACC_DATE").Value < rsFrom.Fields("VACC_DATE") Then
                          xFunc = "UPDATE"          ' Keep record in To Table
                          NextRec = 3         ' Move to next record from To Table
                       ElseIf .Fields("VACC_DATE").Value > rsFrom.Fields("VACC_DATE") Then
                          xFunc = ""          ' Keep record in To Table
                          NextRec = 3         ' Move to next record from To Table
                       Else
                        ' Udate To Table record with From Table record
                        xFunc = "UPDATE"
                        NextRec = 3
                       End If
                      End If
                    End If
                  End If
                Else
                  xFunc = ""      ' both records are the same
                  NextRec = 3
                End If
               End If
             End If
           End If
           Select Case xFunc
             Case "ADD"
              sSQL = "Insert INTO " & TableName & " " & uBuildInsertSqlFromRS(rsFrom)
              
             Case "UPDATE"
              sSQL = "Update " & TableName & " " & uBuildUpdateSqlFromRS(TableName, rsFrom, rsTo)
              
             Case "DELETE"
              sSQL = "DELETE FROM " & TableName & " " & uBuildDeleteSqlFromRs(TableName, rsTo)
             Case Else
               ' Do Nothing
           End Select
           If xFunc <> "" Then gobjSQL.ExecuteSQL sSQL
           If NextRec = 1 Or NextRec = 3 Then
             rsFrom.MoveNext
             FromEOF = IIf(rsFrom.EOF, 1, 0)
           End If
           If NextRec = 2 Or NextRec = 3 Then
             .MoveNext
             ToEOF = IIf(rsTo.EOF, 1, 0)
           End If
          Loop While Not xdone
        End With
        
     End If
     
   Next RecordsetType
   
   ' Process IMM_To Table_WORK
   ' It was decided not to move this table
'   If MoveToCDR Then
'
'        Set rsTo = ImportExport.GetATable(FromExtKey & "Imm_CDR_Work")
'        Set rsFrom = ImportExport.GetATable(ToExtKey & "Imm_CDR_Work")
'
'   End If
   
GoodReturn:
   ReturnStatus = True
exitfunction:
   Set ToCrDate = Nothing
   Set FromCrDate = Nothing
   Set rsTo = Nothing
   Set rsFrom = Nothing
   MSaveDataToCDRorGEMS = ReturnStatus
   Exit Function
haveerror:
   Debug.Print Err.Number, Err.Description
   Resume exitfunction
   Resume
End Function


Public Function BuildPatientID(DMBEObj As IDMBE.I_DMBE) As String
  Dim sFMP As String

    With DMBEObj

        If Len(.PatientObj.FMP & vbNullString) = 0 Then
          sFMP = "00"
        Else
          sFMP = Format(.PatientObj.FMP, "00")
        End If
        
        BuildPatientID = RemoveDashes(CStr(.PatientObj.SponsorSSN)) & sFMP
   End With
End Function

Public Function convertBS(rs As ADODB.Recordset, ByVal typeoftable As Long) As ADODB.Recordset

   ' T.Noren  10/09/01
   
   ' This routine is used to fix the problem created by CHCSII_DAS
   ' when createing recordsets from CDR+
   ' .....certain field types are created incorrectly I.E  VARCHAR instead of NUMERIC
   
   '  Solution is to create an emptry recordset with proper field definitions
   '  and then move the records from the incorrect recordset to the correct recordset
   '  by moving each field in every record.
   
   
   Dim NewRs As ADODB.Recordset
   Dim xfield As ADODB.field
   
   CreateRst typeoftable, NewRs, True
   Do While Not rs.EOF
    NewRs.AddNew
    UpdateRecord rs, NewRs
    NewRs.Update
    rs.MoveNext
   Loop
   Set convertBS = NewRs
   
   
End Function

Public Function PersistRS(ByVal sSQL As String) As ADODB.Recordset
  Set PersistRS = gobjSQL.OpenRecordset(sSQL)
End Function
