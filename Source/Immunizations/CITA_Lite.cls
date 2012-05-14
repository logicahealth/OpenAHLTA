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
Attribute VB_Name = "CITA_Lite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Implements ICHCSII_CoreAPI2
Implements ICHCSII_CoreAPI3

Dim myPatient As Object ' CISMTypes.Patient
Dim myAllergies As Object 'CISMTypes.Allergies
Dim oCITA As cCITA_LITE

Private glOleMenuItem0, glOleMenuItem1, glOleMenuItem2, glOleMenuItem3, glOleMenuItem4, glOleMenuItem5
Private ghTemplate0, ghTemplate1, ghTemplate2, ghTemplate3

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
  convertUIC = ""
Else
  j = InStr(myUIC, ")")
  If j < i Then
    convertUIC = ""
  Else
    convertUIC = Mid(myUIC, i + 1, j - (i + 1))
  End If
End If

End Function

Private Function RemoveDashes(mySponSSN As String) As String
'it seems that sometimes an SSN will have dashes in it already. This is bad.
'so, we'll remove them.
Dim sTemp1 As String
Dim sTemp2 As String
Dim sTemp3 As String
sTemp1 = mySponSSN
While Len(sTemp1) > 0
  sTemp2 = Left(sTemp1, 1)
  sTemp1 = Mid(sTemp1, 2)
  If Asc(sTemp2) >= 48 And Asc(sTemp2) <= 57 Then 'ascii codes for 0-9
    sTemp3 = sTemp3 & sTemp2
  End If
Wend
RemoveDashes = sTemp3
End Function


Private Function GetFMP(lUnitNumber As Long) As Integer
On Error GoTo err_Handler
   '---OLE objects---
    Dim mobjMember As Member
    Dim mobjMMI As MmiDoc
    
    Set mobjMMI = New MmiDoc
    
    mobjMMI.CreateMember ("Patient")
    '---assign the MMI number of the patient to retrieve---
    mobjMMI.Member.UnitNumber = lUnitNumber

    mobjMMI.Get (False)      '--if "false", then "read-only", else writeable --
    
    '---point to the member object---
    Set mobjMember = mobjMMI.Member
    
    '--- Generic MMI Objects
    If Not mobjMember.GenericId Is Nothing Then
        mobjMember.GenericId.GotoFirst
        If Err Then
            Err.Clear
        Else
            Do While Not mobjMember.GenericId.EOF  '--- Or mobjMember.GenericData…, etc., depending on your needs
                If mobjMember.GenericId.Type = NCID_FMP Then
                    GetFMP = mobjMember.GenericId.Value
                    Exit Do
                Else
                    mobjMember.GenericId.GoToNext
                End If
            Loop
        End If
    End If
  Set mobjMember = Nothing
  Set mobjMMI = Nothing
  Exit Function

err_Handler:
  Call gobjShared.ShowVBError(Err.Number, Err.Description, "Immunisync.GetFMP", "ImmSync", vbExclamation)

  Set mobjMember = Nothing
  Set mobjMMI = Nothing
  Exit Function
Resume Next
End Function

Private Sub GetAllergiesFromCDR(sFMP As String, sSSN As String, lCurPat As Long, colAllergies As Object) 'CISMTypes.Allergies)

Dim mCISMAllergy As Object 'CISMTypes.Allergy
Dim objAllergy As ALLERGYATLLib.AllergyInfo
Dim i As Integer
On Error GoTo Cleanup
' If Not colAllergies Is Nothing Then
'  For i = 0 To colAllergies.Count
'    colAllergies.Remove (1)
'  Next i
' Else
  Set colAllergies = New Allergies ' CreateObject("immunisync.Allergies")
' End If
'do whatever you have to do to get the patient
'get the clinical events?
     With gObjDas
        .ClearAll
        .GeneralCriteria.ClearAll
        .GeneralCriteria.UnitNumber = lCurPat
        .GeneralCriteria.EventType.ValueConcept = NCID_PATIENT_ALLERGIES
        .DoSearch
      
        While Not .EOF
          Set mCISMAllergy = colAllergies.Add(sFMP, sSSN)
          Set objAllergy = gObjDas.Events.PatientData
          If Not objAllergy.CompositeSubstance.IsEmpty Then
               If objAllergy.CompositeSubstance.IsText Then
                   mCISMAllergy.sAllergen = "Other: " & objAllergy.CompositeSubstance.Value(NCID_3M_DEFAULT_CONTEXT)
               Else
                   mCISMAllergy.sAllergen = objAllergy.CompositeSubstance.Value(NCID_3M_DEFAULT_CONTEXT)
               End If
          End If
          mCISMAllergy.sSeverity = ""
          mCISMAllergy.lIEN = lCurPat
          .GoToNext
        Wend
    End With
    'Set gobjDas = Nothing
Exit Sub
Cleanup:

If Not colAllergies Is Nothing Then
  Set colAllergies = Nothing
End If
If Not mCISMAllergy Is Nothing Then
  Set mCISMAllergy = Nothing
End If
If Not objAllergy Is Nothing Then
  Set objAllergy = Nothing
End If
'
Call gobjShared.ShowVBError(Err.Number, Err.Description, "Immunisync.GetAllergiesFromCDR", "ImmSync", vbExclamation)
'MsgBox Erl & ": error " & Err.Number & ": " & Err.Description
End Sub
Private Function GetSponSSN() As String
Dim bFoundMember As Boolean
On Error GoTo Cleanup

If gObjMMI Is Nothing Then
  Set gObjMMI = New MmiDoc
End If

GetSponSSN = 0
gObjMMI.CreateMember ("Patient")
gObjMMI.Member.UnitNumber = glCurPat

gObjMMI.Get (False)      '--if "false", then "read-only", else writeable --
    
If Not gObjMMI.Member.GenericId Is Nothing Then gObjMMI.Member.GenericId.GotoFirst
Do While Not gObjMMI.Member.GenericId.EOF

    If gObjMMI.Member.GenericId.Type = NCID_SPONSOR_SSN Then
        GetSponSSN = RemoveDashes(gObjMMI.Member.GenericId.Value)
        Exit Do
    End If
    gObjMMI.Member.GenericId.GoToNext
Loop

Exit Function
Cleanup:
  GetSponSSN = 0
  'Set gobjMember = Nothing
End Function

Private Sub PrepCitaInfo()

 Dim MyPatient2 As MMMHISPatient.Patient
 Dim tmp As Object 'CISMTypes.Allergy
 Dim comma As Integer
 Dim sSQL As String
 Dim objSQL As CHCSII_DAS
 Dim objRecset As ADODB.Recordset

 On Error GoTo Cleanup
  
 Set MyPatient2 = cwobjPatient

  Set myPatient = New Patient 'CreateObject("immunisync.Patient")
  comma = InStr(MyPatient2.Name, ",")
  With myPatient
    If comma > 0 Then
      .sLastName = Left(MyPatient2.Name, (comma - 1))
      .sFirstName = Trim(Mid(MyPatient2.Name, (comma + 1)))
    Else
      .sLastName = cwobjPatient.Name
    End If
    .sSSN = MyPatient2.SSN
    .dDOB = MyPatient2.DateOfBirth
    .lIEN = MyPatient2.UnitNumber
    .sSex = MyPatient2.Sex
    .sRank = MyPatient2.RANK
    .sService = Left(MyPatient2.PatientCategory, 1)
    .sRank = StrConv(.sRank, vbProperCase)
    sSQL = "select abbrev_rank from rank_abbrev where branch = " & "'" & .sService & "'" & " and full_rank Like " & "'" & .sRank & "'"
    Set objSQL = New CHCSII_DAS
    Set objRecset = objSQL.OpenRecordset(sSQL)
    If Not (objRecset.EOF And objRecset.BOF) Then
        .sRank = objRecset![ABBREV_RANK]
        objRecset.Close
    Else
        .sRank = "UNK"
        objRecset.Close
    End If
    .sSponSSN = GetSponSSN()
    If Len(MyPatient2.PatientCategory) > 2 Then
      .sService = ConvertService(Left(MyPatient2.PatientCategory, 1))
      .sStatus = ConvertStatus(Mid(MyPatient2.PatientCategory, 2, 2))
    End If
    If Len(MyPatient2.UnitIdentificationCode) > 2 Then
      .sUIC = convertUIC(MyPatient2.UnitIdentificationCode)
    End If
    .sFMP = Format(GetFMP(MyPatient2.UnitNumber), "00")
    .sProvider = cwobjLogonEnviron.UserName
    .sProviderNCID = cwobjLogonEnviron.UserNCID
'    sSQL = "Select SSN from provider where name like " & "'" & .sProvider & "%" & "'"
'    Set objRecset = objSQL.OpenRecordset(sSQL)
'    If Not (objRecset.EOF And objRecset.BOF) Then
'        .sProviderSSN = objRecset![SSN]
'        objRecset.Close
'    Else
'        .sProviderSSN = ""
'        objRecset.Close
'    End If
    
    GetAllergiesFromCDR .sFMP, .sSSN, .lIEN, myAllergies
'MsgBox "about to create oconnect object"
  End With
Exit Sub
Cleanup:
'Set MyCITA = Nothing
Set MyPatient2 = Nothing
Set myPatient = Nothing
objRecset.Close
Set objRecset = Nothing
Call gobjShared.ShowVBError(Err.Number, Err.Description, "ShowMe", "FrmImmSync", vbCritical)
Resume Next
End Sub





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
    On Error GoTo ErrHandler
    '---perform actions based on the "view type"---
    If viViewType = cwiREGULAR_VIEW Then
        '---set global flags---
        cwbFirstRegularResize = True
        cwbMainViewRunning = True
        '---------------------------------------------------------------------------------------
        '  If this OLE Server needs it's own menu items, here is a good place to add these
        '    menu items onto the CW Core's main menu.
        '  NOTE: This is OPTIONAL.  You do NOT have to create menu items. Thats why they are
        '        commented out.
        '---------------------------------------------------------------------------------------
        '---add a menu item "Template" on CW's main menu---
        'ghTemplate = cwobjCWComm.AddPopupMenuToCW("&Template", cwlMyProgId)
        '---add a menu item "Ole Server Menu Item" to the menu "Template"---
        'glOleMenuItem = cwobjCWComm.AddMenuItem("&Ole Server Menu Item", ghTemplate, cwlMyProgId, "Template menu item")
        '-----------------------------------------------------------------------------------------
        '  Reset the "parent window" for this OLE server's "main" regular form.  This is the main
        '    "trick" that makes the OLE server look like it is part of the CW Core.
        '  NOTE: This action invokes the main form's "Load" event, but the form is NOT shown.
        '        Make sure that you DO NOT perform a "Me.Show" in the form's "Load" event, as this
        '        will mess up the "SetParent" API call.  Make sure that you DO NOT "Show" the
        '        "main" regular form in this method, that will happen in the "ResizeView" method.
        '-----------------------------------------------------------------------------------------
        SetParent frmImm_People.hwnd, vlParenthWnd
        
        glCurPat = cwobjPatient.UnitNumber

        '---set parameter---
        
        rhMainForm = frmImm_People.hwnd
        Set oMySelf = Me
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
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "Immunization.OpenView", "Immunization", vbCritical)
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
    If viViewType = cwiALL_VIEWS Then
        '---close the "main view"---
        If cwbMainViewRunning Then
            '---set global flags---
            cwbMainViewRunning = False
            '---this indicates that the CW Core is closing the "view", NOT this OLE Server---
            cwbCloseViewMyself = False
            '---close ALL forms that are associated with all "views"---
            Unload frmImm_People
        End If
    ElseIf viViewType = cwiREGULAR_VIEW Then
        '---close the "main view"---
        If cwbMainViewRunning Then
            '---set global flags---
            cwbMainViewRunning = False
            '---this indicates that the CW Core is closing the "view", NOT this OLE Server---
            cwbCloseViewMyself = False
            '---close ALL forms that are associated with all "views"---
            Unload frmImm_People
        End If
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
          Case glOleMenuItem0
              cwobjCWComm.DeleteMenuItem glOleMenuItem0
              'cwobjCWComm.DeleteMenuItem glOleMenuItem1
              'cwobjCWComm.DeleteMenuItem glOleMenuItem2
              cwobjCWComm.SetStatusText ("Close Immunizations")
              cwobjCWComm.CloseView cwiREGULAR_VIEW, cwlMyProgId
              CloseView cwiALL_VIEWS
          Case glOleMenuItem1
              cwobjCWComm.SetStatusText ("Download Patient Data")
          Case glOleMenuItem2
              cwobjCWComm.SetStatusText ("Extract Immunizations")
          Case glOleMenuItem3
              cwobjCWComm.SetStatusText ("Selected Item 3")
          Case glOleMenuItem4
              cwobjCWComm.SetStatusText ("Selected Item 4")
          Case glOleMenuItem5
              cwobjCWComm.SetStatusText ("Selected Item 5")
      End Select

      Select Case vlMenuId 'Button.Key
          Case ghTemplate0
              cwobjCWComm.CloseView cwiALL_VIEWS, cwlMyProgId
          Case ghTemplate1
              '
          Case ghTemplate2
              '
          Case ghTemplate3
              '
      End Select
      
    Exit Sub
ErrHandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "Template.MenuItemSelected", "Template", vbExclamation)
    Exit Sub
End Sub

'**********************************************************************
'  Description: This method gets called when a new patient has been selected from the CW Core.
'               NOTE: This is an OPTIONAL method.
'  Parameters:
'    none
'*************************************************************************
Public Sub PatientSelected()
    If glCurPat <> cwobjPatient.UnitNumber Then
        
        glCurPat = cwobjPatient.UnitNumber
        gsCurPatName = cwobjPatient.Name
'        If MyCITA Is Nothing Then
'          frmImmSync.ShowMe
'          oMySelf.CloseView cwiREGULAR_VIEW
'        End If
        'frmImmSync.lblPatientSrch.Caption = gsCurPatName & _
        '    ": " & cwobjPatient.UnitNumber & _
        '    " " & cwobjPatient.EncounterNumber
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
            '  NOTE: This generates a form "resize" event.
            '---------------------------------------------------------------------------------
            frmImm_People.WindowState = vbNormal
            '---hide the form to so that annyoing "screen blinking"
            '  will not be seen while controls are moved and resized---
            frmImm_People.Visible = False
            '---resize form to fill entire "client area", NOTE: This generates a "resize" event---
            frmImm_People.Move cwlLeft, cwlTop, cwlWidth, cwlHeight
            frmImm_People.Visible = True
        Else    '--first "regular" resize, need to perform "initailization" procedures--
            '---reset flag so this code will not be executed again---
            cwbFirstRegularResize = False
            '----------------------------------------------------------------------------
            '  NOTE: Here is where you "Show" the "Regular View" form for the FIRST time.
            '----------------------------------------------------------------------------
          'PrepCitaInfo
          'StartCita
          frmImm_People.Show
        End If
    End If
    Exit Sub
ErrHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "Immunization.ResizeView", "Immunization", vbExclamation)
    Exit Sub
End Sub

Private Function StartCita()

  Set oCITA = New cCITA_LITE
  
  oCITA.Init "CHCSIICW", "CHCSIICW", myPatient, myAllergies

 ' Set oCITA = oConnect.New_CITA_LITE("CHCSIICW", "CHCSIICW", MyPatient, myAllergies)
''MsgBox "created CitaLite object"
'  Set oConnect = Nothing
'  Set oCITA = Nothing
'  DoEvents

End Function
Private Sub Class_Initialize()
    '---create shared object used to show VB errors, etc.---
    Set gobjShared = New CWShared
    '---set the help file for this application---
    App.HelpFile = gobjShared.CWHelpFile
End Sub

Private Sub Class_Terminate()
    '---release reference to objects---
    Set gobjShared = Nothing
    Set cwobjCWComm = Nothing
    Set cwobjLogonEnviron = Nothing
    Set cwobjConfig = Nothing
    Set cwobjPatient = Nothing
    '**************************************************
    '  TO DO:
    '    Add code here to release reference to all other global objects.
    '******************************************************
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
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "Immunization.Message", "Immunization", vbExclamation)
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
                            robjPatient As Object)
    On Error GoTo ErrHandler
    '---save the "ProgId" assigned to me by the CW Core---
    cwlMyProgId = vlMyProgId
    '---create references to passed in objects---
    Set cwobjCWComm = robjComm
    Set cwobjLogonEnviron = robjLogonEnvironment
    Set cwobjConfig = robjConfig
    Set cwobjPatient = robjPatient
    '---set global flags---
    cwbMainViewRunning = False
    '************************************************************
    '  TO DO:
    '    Add code here to create objects, initialize global variables, etc.
    '****************************************************************
    
    ' Add Menu Items
    glOleMenuItem0 = cwobjCWComm.AddActionMenuItem("Close", "Close.bmp", cwlMyProgId, "Exit Immunizations")
    Set gObjDas = New DasCtrl
    Set gobjConTrans = New ConTrans
    
    Exit Sub
ErrHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "Immunization.Initialize", "Immunization", vbExclamation)
    Exit Sub
End Sub

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
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "Template.Properties", "Template", vbExclamation)
    Exit Sub
End Sub

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
    Set gobjUser = robjUser
    Call Initialize(vlMyProgId, robjComm, robjLogonEnvironment, robjConfig, robjPatient)
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


