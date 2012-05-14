VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDXLstCnt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''************************************************************************
'''msMODULE_NAME:          clsDXLstCnt/DXLstCnt.cls
'''
'''Description:            Class for processing requests associated with Dx pick Lists
'''
'''Modules Required:       Form containing Dx Pick lists and associated controls
'''                        e.g. command ADD to A&P , command search...
'''                        Currently frmDiagnosis and frmTeleconView
'''Author         :
'''Modified       :
'''               :
'''************************************************************************
Option Explicit
''''''''''''''' Constanst Section '''''''''''''''''''''''''''''''''''''''''

Const msMODULE_NAME                    As String = "clsDXLstCnt."

'''Maximum diagnosis allowed per encounter
Private Const mlMAX_DX_PER_ENC         As Long = 15

'''Property Keys for DX Items returned by psmdDx control
Private Const SMD_DX_PROP_Description  As String = "Description"
Private Const SMD_DX_PROP_NCID         As String = "NCID"
Private Const SMD_DX_PROP_SNO_ID    As String = "SnoID"
Private Const SMD_DX_PROP_ICD9         As String = "ICD9"
Private Const SMD_DX_PROP_PREFIX       As String = "Prefix"

'''Error messages
Private Const msSELECTION_REQUIRED     As String = _
              "Please select one or more valid diagnoses prior to initiating this operation.  Valid diagnoses are displayed with an ICD code."
'JT SCR 45625 - added "the" to max dx message
Private Const msMAX_DX_ALLOWED         As String = _
              "the maximum number (" & mlMAX_DX_PER_ENC & ") of diagnoses allowed"
'''Data source
Public Enum pnumDXPLM_DxPickListMode
   lDXPLM_Template         '''Template List
   lDXPLM_ProblemAcute     '''Problem List(Acute)
   lDXPLM_ProblemChronic   '''Problem List(Chronic)
   lDXPLM_Favorites        '''Clinic & User Favorite Lists
   lDXPLM_SnoSearch     '''Sno Search
End Enum
Private mlDXPLM_DxPickListMode         As pnumDXPLM_DxPickListMode


''''''''''''''' Variables Section '''''''''''''''''''''''''''''''''''''''''

Private mlLastSearchResult             As actxSearchSnoMed.enumSearchResult
Private mlLastSearchPattern            As String

''''''''''''''' Objects Section '''''''''''''''''''''''''''''''''''''''''

Public pobjParentDx                    As Diagnosis

Private mobjListManagementOps          As ListManagementOps

'ARM SCR 19539 PJS 12.19.2001
Public psCallingFormName               As String

Public pddcDODCentricCodes             As actxDODCentricCodes.DOD_Centric_Codes

'''GUI
Public poptPickListSearch              As OptionButton   '''Allowes to deselect Persisted lists
Public WithEvents pdxgDiagnosisGrid    As GridX
Attribute pdxgDiagnosisGrid.VB_VarHelpID = -1
Public WithEvents psmdDx               As actxSearchSnoMed.searchSnoMed2
Attribute psmdDx.VB_VarHelpID = -1
Public WithEvents pcmdAddToFavorites   As CommandButton
Attribute pcmdAddToFavorites.VB_VarHelpID = -1
Public WithEvents pcmdSearch           As CommandButton
Attribute pcmdSearch.VB_VarHelpID = -1
Public WithEvents pcboSearchDiagnosis  As ComboBox
Attribute pcboSearchDiagnosis.VB_VarHelpID = -1
Public WithEvents pcmdAddToAP          As CommandButton
Attribute pcmdAddToAP.VB_VarHelpID = -1

''''''''''''''' Events Section '''''''''''''''''''''''''''''''''''''''''

Public Event pevItemProcessedAP()
Public Event pevAllItemProcessedAP()

Private Sub M_AddSelectedDxItems(ByVal bAddToAP As Boolean)
'''************************************************************************
'''Called By:        cmdAddFavorite_Click
'''                  cmdDiagAdd_Click
'''
'''Arguments:
''' bAddToAP         Indicates whether to add requested Items to A&P or Favorites
''' sCaption         Title of the current operation
'''
'''Description:      Depending on bAddToAP add currently selected items to A&P or Favorites
'''
'''Note:
'''Warning:
'''Updates Globals:
'''Returns:
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "M_AddSelectedDxItems"

   Dim lLine            As Long
   
On Error GoTo ErrHnd
                                                                                          
   
   Screen.MousePointer = vbHourglass            'SCR 46002 - Toggle hourglass during this operation
                                                                                          lLine = 20
   gobjAnPParent.LockAnP sPROC_NAME
                                                                                          lLine = 40
   If psmdDx.Visible Then
      If psmdDx.SelectedItems.Count <> 0 Then
        Call M_DXSearchItemsAdd(bAddToAP)
      Else
        'SCR 19141 3/29/02  Phil Crowder
        MsgBxARMd "Please select a diagnosis with an ICD9 code.", , "Diagnosis"
      End If
                                                                                          lLine = 80
   Else
                                                                                          lLine = 100
      Call M_DxPersistedItemsAdd(bAddToAP)
   End If
                                                                                          lLine = 200
   gobjAnPParent.UnlockAnP sPROC_NAME

   If bAddToAP Then pcboSearchDiagnosis.SetFocus    'SCR 22260  Phil Crowder 3/16/02

ErrHnd:
   
   Screen.MousePointer = vbDefault            'SCR 46002 - Toggle hourglass during this operation
   
   If CBool(Err) Then
      Call gobjShared.ShowVBError _
      (Err.Number, Err.Description & vbCrLf & "Line: " & lLine, msMODULE_NAME & sPROC_NAME, Exe, vbCritical)
   End If
   
End Sub
Private Function M_bAddDxItems(colDxItems As Collection, ByVal bProblemListItems As Boolean, ByVal bAddToAP As Boolean) As Boolean
'''************************************************************************
'''Called By:        M_DXSearchItemsAdd
'''                  M_DxPersistedItemsAdd
'''
'''Arguments:
''' colDxItems       Collection of Dx Items requested to be added to A&P or Favorites
''' bAddToAP         Indicates whether to add requested Items to A&P or Favorites
''' bProblemListItem Indicates whether specified DX items are part of the known patient problems
'''                  (patient was already diagnosed with it)
'''
'''Description:      Process specified Dx Items
'''                  If Maxixum number allowed has not beed reached and Item is valid
'''                     Add it to specified repository (A&P or Favorites)
'''
'''Note:
'''Warning:
'''Updates Globals:
'''Returns:
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "M_bAddDxItems"

   Dim objDxI           As clsDXI_DxItem
   Dim sErrMsg          As String
   Dim sErrReport       As String
   Dim lLine            As Long
   Dim bCancel          As Boolean
   Dim sMsgTitle        As String
   
On Error GoTo ErrHnd
                                                                                             lLine = 10
  If bAddToAP Then
   sMsgTitle = "Add Diagnosis to Encounter"
  Else
    sMsgTitle = "Add Diagnosis to Favorites"
  End If
                                                                                             lLine = 20
   'SCR 22521 3/29/02  Phil Crowder
   If colDxItems.Count > 0 And gobjAnPParent.DiagnosisCount <> mlMAX_DX_PER_ENC Then
                                                                                             lLine = 30
      If bAddToAP Then
         If (colDxItems.Count + gobjAnPParent.DiagnosisCount) > mlMAX_DX_PER_ENC Then
                                                                                             lLine = 40
            sErrMsg = "Total number of existing and selected diagnosis exceeds " _
                    & msMAX_DX_ALLOWED & vbCrLf _
                    & "Some of the selected items will not be included"
                                                                                             lLine = 50
            '''If user cancel the operation
            If Not MsgBxARMd(sErrMsg, vbQuestion + vbOKCancel, sMsgTitle) = vbOK Then
                                                                                             lLine = 60
               bCancel = True
            End If
         End If
      '''else
         '''Currenly no limit for adding to favorites
      End If
'SCR 19141 3/29/02  Phil Crowder
'   Else
'                                                                                             lLine = 100
'      sErrReport = msSELECTION_REQUIRED
   End If
   '''If no errors and user did not cancel
                                                                                             lLine = 200
   If Not bCancel And sErrReport = "" Then
                                                                                             lLine = 210
      For Each objDxI In colDxItems
         If bAddToAP Then
            If gobjAnPParent.DiagnosisCount >= mlMAX_DX_PER_ENC Then
                                                                                             lLine = 220
               sErrReport = sErrReport & vbCrLf & "You have reached " & msMAX_DX_ALLOWED
               bCancel = True
            End If
         End If
         If Not bCancel Then
            '''If Item is valid
            If M_bDXItemValid(objDxI, bAddToAP, sErrMsg) Then
                                                                                             lLine = 300
               If bAddToAP Then
                                                                                             lLine = 320
                  '''Convert it into A&P format and add it to A&P collection of diagnosis
                  Call M_DXItemSendToAP(objDxI, bProblemListItems)
               Else
                                                                                             lLine = 340
                  Call M_DXItemSendToFavorites(objDxI, bProblemListItems)     'JT SCR 42399 - indicate if item came from problem list
                  
               End If
            End If
         Else
                                                                                             lLine = 500
            Exit For
         End If
      Next
   End If
   
   If sErrReport <> "" Then
                                                                                             lLine = 600
      Call MsgBxARMd(sErrReport, vbInformation + vbOKOnly, sMsgTitle)
      
   End If
                                                                                             lLine = 700
   If bAddToAP Then
                                                                                             lLine = 720
      RaiseEvent pevAllItemProcessedAP
   End If
   
   M_bAddDxItems = sErrReport = ""
   
Exit Function
ErrHnd:
   M_bAddDxItems = False
   sErrReport = vbCrLf & "Line: " & lLine & vbCrLf & sErrReport
   Call gobjShared.ShowVBError _
   (Err.Number, Err.Description & sErrReport, msMODULE_NAME & sPROC_NAME, Exe, vbCritical)

End Function

'REQ 44501 and 43255 - Don't Allow certain codes for RNDS appointments
Private Function ValidRoundsDX(ByVal sICD9Code As String, ByRef sErrMsg As String) As Boolean

On Error GoTo ErrrorHandler

Dim oListMgr As ListManagerOps

   'Assume it's valid
   ValidRoundsDX = True
   
   'Only implemented for certain modes
   If Len(gobjShared.CmdLineSwitches("HIPAA837")) > 0 And Not gobjShared.IsAppMode(modeITT) _
      And (gobjShared.IsAppMode(modeCHCSI_GUI) Or gobjShared.IsAppMode(modeCDR)) Then
      
      If UCase$(Left(gobjAnPParent.EncounterApptType, 4)) = "RNDS" Then
         'Only perform this lookup for rounds appointments
         
         Set oListMgr = New ListManagerOps
         
         If oListMgr.IsCodeInDomain(cdtInvalidRNDs, cstDiagnosis, sICD9Code) Then
            cwobjCWComm.Message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_ANP_DIAGNOSIS_APPLICATION
            sErrMsg = "This ICD Code is invalid for the RNDS Appointment Type.  Please choose another."
            MsgBxARMd sErrMsg, vbInformation, "Invalid ICD Code"
            cwobjCWComm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_ANP_DIAGNOSIS_APPLICATION
            ValidRoundsDX = False
         End If
         
      End If
   End If
   
ErrrorHandler:
   Set oListMgr = Nothing
   
   If CBool(Err) Then Call gobjShared.ShowVBError(Err.Number, Err.Description, msMODULE_NAME & "ValidRoundsDX", "Diagnosis", vbCritical)
   
End Function
Private Function M_bDXItemValid(objDxI As clsDXI_DxItem, ByVal bAddToAP As Boolean, ByRef sErrMsg As String) As Boolean
'''************************************************************************
'''Called By:        M_DXSearchItemsAdd
'''
'''Arguments:
''' objDxI Dx item requested to be evaluated
''' bAddToAP Indicates whether to add requested Items to A&P or Favorites
''' sErrMsg Error information if any to be returned byref
'''
'''Description:      Evaluate specified DX item and if neccessary populate extended
'''                  properties and replace description with synonym
'''
'''Note:
'''Warning:
'''Updates Globals:
'''Returns:          True  - If no errors and specified Item is valid
'''                  False - Else
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "M_bDXItemValid"

   
On Error GoTo ErrHnd
    
Dim bRefresh As Boolean       'JT SCR 35220 - Added module refresh flag
Dim tExtCode As ICD9Extender
Dim oConvertSno   As IConvert

   sErrMsg = ""
   bRefresh = False           'JT SCR 35220 - initialize module refresh flag to false
   Set oConvertSno = New IConvert
   
   If Not ValidRoundsDX(objDxI.psICD9, sErrMsg) Then
      'Failed the rounds diagnosis check
      M_bDXItemValid = False
      bRefresh = True
      Exit Function
   End If
   
    '''If ICD code of the specified item is valid DOD specific code
    'JT SCR 36572 - Updated code to check IsAppmode
    
    'SCR 31268  Do not make 3M calls if False  Phil Crowder 1/21/02
    'SCR 47646  Do not make 3M calls in Theater
    pddcDODCentricCodes.Use3M = False ' Not (gobjShared.IsAppMode(modeCHCSI_GUI) Or gobjShared.IsAppMode(modeTheater))
    
    If pddcDODCentricCodes.isValidICDCode(objDxI.psICD9) Then
       'JT BEGIN SCR 42310 - use the DODCentricCodes control to process extender selection - Also add message to core about modal dialog
       
       'Show the extenders form
       cwobjCWComm.Message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_ANP_DIAGNOSIS_APPLICATION
       sErrMsg = pddcDODCentricCodes.LoadExtender(objDxI.psICD9, tExtCode)
       cwobjCWComm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_ANP_DIAGNOSIS_APPLICATION
       
       'Process the selection if any
       If Len(sErrMsg) = 0 And Len(tExtCode.Code) > 0 Then
          'ICD Extended was selected - update diagnosis info
          objDxI.psExtICD_Code = tExtCode.Code
          objDxI.psExtICD_Name = tExtCode.Name
          objDxI.psExtICD_NCID = tExtCode.NCID
       End If
       bRefresh = True      'Force a refresh since a dialog may have been displayed
       'JT END SCR 42310
       
    Else
       'JT SCR 35220 - Invalid ICD code results in the display of an error message in the DODCentricCodes control
                      'Since the treeview doesn't refresh properly when a message box displays over it, set the refresh flag so we instruct the
                      'module to refresh itself.
       bRefresh = True
       
       If Not AutoSubCode(objDxI) Then
         sErrMsg = "Invalid DOD Centric ICD Code."
       End If
    End If
   If sErrMsg = "" Then
      '''If valid synonym '''<DSN should never happened
      If Len(objDxI.psDescription) < 2 Then
         sErrMsg = "Failed to find valid synonym"
      End If
   End If
    
   'JT SCR 35220 - Refresh the form based upon the refresh flag
   If bRefresh Then
      'JT SCR 43551 - Don't refresh unless diagnosis for is active in AnP
      If StrComp(psCallingFormName, "frmDiagnosis", vbTextCompare) = 0 Then
         pobjParentDx.ResizeView cwiREGULAR_VIEW, cwlLeft, cwlTop, cwlWidth, cwlHeight
         DoEvents             'Allow time to repaint
      End If
      'JT END SCR 43551
   End If
   
   M_bDXItemValid = sErrMsg = ""
   
Exit Function
ErrHnd:
  If Err.Number = -2147220604 Then
   'SCR 24083  Phil Crowder 7/29/02
   MsgBxARMd "Unable to Validate DOD Specific codes because of Server Timeout.  Please save your work and try again.", , "Diagnosis"
  Else
   M_bDXItemValid = False
   sErrMsg = Err.Description
   Call gobjShared.ShowVBError(Err.Number, sErrMsg, msMODULE_NAME & sPROC_NAME, Exe, vbCritical)
  End If
  
End Function

Private Function AutoSubCode(ByRef objDxI As clsDXI_DxItem) As Boolean
Dim NewCode As String

  NewCode = pddcDODCentricCodes.AutoSubstitute(objDxI.psICD9)
  
  If Len(NewCode) > 0 Then
    objDxI.psICD9 = NewCode
    objDxI.psExtICD_Code = NewCode
    AutoSubCode = True
  End If
  
End Function

Private Sub M_DXItemSendToFavorites(objDxI As clsDXI_DxItem, ByVal bProblemListItems As Boolean)
'''************************************************************************
'''Called By:        M_bAddDxItems
'''
'''Arguments:
''' objDxI Dx item requested to be added to user Favorites
'''
'''Description:      Add Favorites
'''
'''Note:
'''Warning:
'''Updates Globals:
'''Returns:
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "M_DXItemSendToFavorites"

   Dim lLine            As Long
   
On Error GoTo ErrHnd

   'SCR-50350 RBELL 2/24/2004 Always set object to insure that we have the latest entries from the List Management module.
   'If mobjListManagementOps Is Nothing Then
   Call M_InitializeListMgmt
   'End If
   
   'JT BEGIN SCR 42399 - Added support for items from the problem list
   If gobjShared.IsAppMode(modeCHCSI_GUI) And bProblemListItems Then
      'Sno id will actually be the ICD9 code instead of the Sno id - add a . to the end so list management can properly identify it
      Call mobjListManagementOps.User.Diagnosis.Add(Trim(objDxI.psSnoID) & ".", objDxI.psPrefix)
   Else
      Call mobjListManagementOps.User.Diagnosis.Add(CLng(objDxI.psSnoID), objDxI.psPrefix)
   End If
   'JT END SCR 42399
    
   'SCR-50350 RBELL 2/24/2004 Reset object to trigger save event
   'SCR-50350 Added calls to new save/message funcitonality in List Management to redisplay the data
   mobjListManagementOps.SaveUserList
    cwobjCWComm.Message cwiREFRESH_DATA, "", NCID_CW_APPLICATION, NCID_ANP_DIAGNOSIS_APPLICATION, False
    Set mobjListManagementOps = Nothing

Exit Sub
ErrHnd:
   Call gobjShared.ShowVBError _
   (Err.Number, Err.Description & vbCrLf & "Line: " & lLine, msMODULE_NAME & sPROC_NAME, Exe, vbCritical)

End Sub

Private Sub M_DXSearchItemsAdd(ByVal bAddToAP As Boolean)
'''************************************************************************
'''Called By:        cmdDiagAdd_Click
'''
'''Arguments:
''' bAddToAP Indicates whether to add requested Items to A&P or Favorites
'''
'''Description:      Process selected Dx search Items
'''                  If Item is not a duplicate (not already added to A&P) and
'''                  Maxixum number allowed has not beed reached
'''                     Add it to the A&P
'''
'''Note:
'''Warning:
'''Updates Globals:
'''Returns:
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "M_DXSearchItemsAdd"

   Dim objDx            As actxSearchSnoMed.ListItem
   Dim objDxI           As clsDXI_DxItem
   Dim colDxItems       As Collection
   Dim lLine            As Long
   
On Error GoTo ErrHnd
                                                                                    lLine = 20
   Set colDxItems = New Collection
                                                                                    lLine = 40
   For Each objDx In psmdDx.SelectedItems
      '''Extract item properties
                                                                                    lLine = 60
      Set objDxI = New clsDXI_DxItem
                                                                                    lLine = 80
      objDxI.psDescription = objDx.Items(SMD_DX_PROP_Description).value
                                                                                    lLine = 100
      objDxI.psSnoID = objDx.Items(SMD_DX_PROP_SNO_ID).value
      '''<DSN For now - to maintain compatibility
      '''     - assign SnoID to NCID as it was done in the previous code
                                                                                    lLine = 120
      '''objDxI.psNCID = objDx.Items(SMD_DX_PROP_NCID).value
      objDxI.psNCID = objDxI.psSnoID
      '''>DSN
                                                                                    lLine = 140
      objDxI.psICD9 = objDx.Items(SMD_DX_PROP_ICD9).value
                                                                                    lLine = 180
      objDxI.psPrefix = objDx.Items(SMD_DX_PROP_PREFIX).value
                                                                                    lLine = 200
      Call colDxItems.Add(objDxI)
                                                                                    
   Next
                                                                                    lLine = 300
   Call M_bAddDxItems(colDxItems, False, bAddToAP)

   If bAddToAP Then pcboSearchDiagnosis.SetFocus      'SCR 22260  Phil Crowder 3/16/02

Exit Sub
ErrHnd:
   Call gobjShared.ShowVBError _
   (Err.Number, Err.Description & vbCrLf & "Line: " & lLine, msMODULE_NAME & sPROC_NAME, Exe, vbCritical)

End Sub
Private Sub M_InitializeListMgmt()
'''************************************************************************
'''Called By:
'''
'''Arguments:
'''
'''Description:
'''
'''
'''Note:
'''Warning:
'''Updates Globals:
'''Returns:
'''
'''Programmer:       Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME        As String = "M_InitializeListMgmt"

   Dim mobjListManagement  As Object
   Dim lLine               As Long
   
On Error GoTo ErrHandler
                                                                                             lLine = 10
    Set mobjListManagement = cwobjCWComm.InitializeOLEServer(NCID_LIST_MGMT_APPLICATION)
                                                                                             lLine = 20
    Set mobjListManagementOps = mobjListManagement.ListManagementOps
    
Exit Sub
ErrHandler:
   Call gobjShared.ShowVBError _
   (Err.Number, Err.Description & vbCrLf & "Line: " & lLine, msMODULE_NAME & sPROC_NAME, Exe, vbCritical)
Resume Next
End Sub
Private Sub M_DxPersistedItemsAdd(ByVal bAddToAP As Boolean)
'''************************************************************************
'''Called By:        cmdDiagAdd_Click
'''
'''Arguments:
''' bAddToAP Indicates whether to add requested Items to A&P or Favorites
'''
'''Description:      Process selected Dx Items
'''                  If Item represents a Dx
'''                     Include it for submission
'''
'''
'''Note:
'''Warning:
'''Updates Globals:
'''Returns:
'''
'''Programmer:       Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME        As String = "M_DxPersistedItemsAdd"

   Dim lIdx                As Long
   Dim lSelRow             As Long
   Dim objDxI              As clsDXI_DxItem
   Dim bProblemListItem    As Boolean
   Dim colDxItems          As Collection
   Dim lLine               As Long
   
On Error GoTo ErrHnd
                                                                                          
   Screen.MousePointer = vbHourglass                  'SCR 46002 - Toggle hourglass during this operation
                                                                                          lLine = 20
   gobjAnPParent.LockAnP sPROC_NAME                   'SCR 46002 - Make sure A/P locking is being used during this operation

   Set colDxItems = New Collection
                                                                                          lLine = 40
   bProblemListItem = mlDXPLM_DxPickListMode = lDXPLM_ProblemAcute Or _
                      mlDXPLM_DxPickListMode = lDXPLM_ProblemChronic
                                                                                          lLine = 60
   With pdxgDiagnosisGrid.objGeneralGrid
                                                                                          lLine = 80
      For lIdx = 0 To .SelectedRows - 1
                                                                                          lLine = 100
         lSelRow = .SelectedRow(lIdx)
                                                                                          lLine = 120
         '''If selected row contains actual Dx Item (not a group title or separator...)
         If Trim$(.TextMatrix(lSelRow, 0)) <> LBL_CLINIC_FAV And _
            Trim$(.TextMatrix(lSelRow, 0)) <> LBL_USER_FAV And _
            Len(.TextMatrix(lSelRow, 3)) > 4 _
         Then
                                                                                          lLine = 160
            Set objDxI = New clsDXI_DxItem
                                                                                          lLine = 180
            objDxI.psDescription = .TextMatrix(lSelRow, 1)
                                                                                          lLine = 220
            objDxI.psICD9 = .TextMatrix(lSelRow, 0)
                                                                                          lLine = 260
            objDxI.psNCID = .TextMatrix(lSelRow, 2)
                                                                                          lLine = 280
            objDxI.psSnoID = .TextMatrix(lSelRow, 2)
                                                                                          lLine = 300
            objDxI.psPrefix = Trim$(.TextMatrix(lSelRow, 4))   'scr 21770  3/29/02  Phil Crowder
                                                                                          
            If Len(Trim$(.TextMatrix(lSelRow, 8))) > 0 Then    'JT SCR 42399 - track if we have a valid Sno id
               objDxI.psNotFromMedcin = CBool(Trim$(.TextMatrix(lSelRow, 8)))
            End If
            
            Call colDxItems.Add(objDxI)
         End If
      Next lIdx
   End With

   Call M_bAddDxItems(colDxItems, bProblemListItem, bAddToAP)
   
   gobjAnPParent.UnlockAnP sPROC_NAME              'SCR 46002 - Make sure A/P locking is being used during this operation
   
   If bAddToAP Then pcboSearchDiagnosis.SetFocus            'SCR 22260  Phil Crowder 3/16/02
   
ErrHnd:
   Screen.MousePointer = vbDefault                  'SCR 46002 - Toggle hourglass during this operation
   
   If CBool(Err) Then
      Call gobjShared.ShowVBError _
      (Err.Number, Err.Description & vbCrLf & "Line: " & lLine, msMODULE_NAME & sPROC_NAME, Exe, vbCritical)
   End If
   
End Sub
Private Function M_bSearchSno() As Boolean
'''************************************************************************
'''Called By:     pcmdSearch_Click
'''               pcboSearchDiagnosis_KeyPress
'''Arguments:
'''
'''Description:   Initiate and perform diagnosis search based on specified
'''               pattern
'''
'''Note:
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:       True  - If successful
'''               False - Else
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "M_bSearchSno"

   Dim bResult          As Boolean
   Dim lPatientAge      As Long
   Dim lIdx             As Long
   Dim bPatternExist    As Boolean
   Dim sMsg             As String
   Dim sSearchPattern   As String
   Static bIn_static    As Boolean
   
On Error GoTo ErrHnd
    '///ARM-SCR 19539 PJS 12.18.2001
   Dim lARM_handle As Long
   If psCallingFormName = "frmDiagnosis" Then
        lARM_handle = goARM.StartTx(gsARMTx_SearchMedcinForDX)
   End If
   '\\\ARM-SCR 19539 PJS 12.18.2001
   
   If Not bIn_static Then
      bIn_static = True
      sSearchPattern = pcboSearchDiagnosis.Text
      sMsg = "Begin Function - Search HDD"
      Screen.MousePointer = vbHourglass
      sSearchPattern = Trim$(sSearchPattern)
    #If DEBUGON Then
        Dim DR  As DebugRec
        DL.DebugLog Exe, msMODULE_NAME, "GetDiagnosis", "P-Before Sno Search", DR
    #End If
      sMsg = "Evaluating Age of the Patient"
      If (cwobjPatient.DateOfBirth = DateValue("31-Dec-9999")) Then
         lPatientAge = 0
      Else
         '''Get the patient demographic data available
         '''Need Patient Age in minutes
         lPatientAge = 1440 * DateDiff("d", cwobjPatient.DateOfBirth, gobjAnPParent.StartDTS)
      End If
      
      With psmdDx
         '''<DEV need clear method
         Call .ClearSelection
         .MultiSelect = True
                  
         .ListMode = mlmDiagnosis_Search
         sMsg = "Start Search"
         mlLastSearchPattern = sSearchPattern
         
         'JT BEGIN SCR 47437 - Added modal dialog messages since a dialog could display in not records are found
         cwobjCWComm.Message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_ANP_DIAGNOSIS_APPLICATION
         Call .SnoMedSearch(sSearchPattern)
         cwobjCWComm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_ANP_DIAGNOSIS_APPLICATION
         'JT END SCR 47437
                  
         '''<DSN Should return true if at least one is found - currently this functionality is not available
         If mlLastSearchResult = msrSuccess Then
            bResult = True
         End If
         '''>DSN
         sMsg = "End Search"
      End With
       
      sMsg = "Exit Function"
      Screen.MousePointer = vbDefault
                                    #If DEBUGON Then
                                       DL.DebugLog Exe, msMODULE_NAME, sPROC_NAME, "P-After Medcin Search", DR
                                    #End If
      If bResult Then
         Call M_UpdateSearchPatternList
      End If
      bIn_static = False
   End If
   
   M_bSearchSno = bResult
   
   '///ARM-SCR 19539 PJS 12.18.2001
   If psCallingFormName = "frmDiagnosis" Then
      Select Case mlLastSearchResult
        Case actxSearchSnoMed.enumSearchResult.msrNoneFound, _
             actxSearchSnoMed.enumSearchResult.msrSuccess
             goARM.StopTx lARM_handle, enmARMTxStatusOk
        Case actxSearchSnoMed.enumSearchResult.msrError
             goARM.StopTx lARM_handle, enmARMTxStatusAborted
        Case Else  'bad input or more than 500 results give a user prompt and skew timing
            goARM.StopTx lARM_handle, enmARMTxStatusFailed
      End Select
   End If
   '\\\ARM-SCR 19539 PJS 12.18.2001

Exit Function
ErrHnd:
   Screen.MousePointer = vbDefault
   Call gobjShared.ShowVBError _
   (Err.Number, Err.Description & " - " & sMsg, msMODULE_NAME & sPROC_NAME, Exe, vbCritical)
   bIn_static = False
   
End Function
Private Sub M_DXItemSendToAP(objDxI As clsDXI_DxItem, ByVal bProblemListItem As Boolean)
'''************************************************************************
'''Called By:     M_bAddDxItems
'''
'''Arguments:
''' objDxI Dx item requested to be added to A&P
''' bProblemListItem Indicates whether specified DX is one of the known problems
'''                  for current patient (patient was already diagnosed with this problem)
'''
'''Description:   Populate A&P DX object using specified DX Search object
'''               and add to A&P diagnosis collection
'''
'''Note:
'''Warning:
'''Updates Globals:
'''Returns:
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "M_DXItemSendToAP"

   Dim lIdx             As Long
   Dim Chronicity       As String
   Dim DiagType         As String
   Dim bNewProblem As Boolean
   Dim bAcuteProblem As Boolean
   
On Error GoTo ErrHnd
   
   'SCR 33766 3/18/03  Phil Crowder
   If Len(gobjShared.CmdLineSwitches("GREEN")) > 0 Then
    bAcuteProblem = IsAcuteProblem(objDxI)
    
    If Not bAcuteProblem Then
       Chronicity = "Chronic"
    Else
       Chronicity = "Acute"
    End If
   End If
   
   'SCR 33766 3/18/03  Phil Crowder
   If Not bProblemListItem Then
    bNewProblem = IsNewProblem(objDxI)
    
    If bNewProblem Then
      DiagType = "New"
      If Len(gobjShared.CmdLineSwitches("GREEN")) = 0 Then Chronicity = "Acute"
    Else
      DiagType = "Follow-Up"
      If Len(gobjShared.CmdLineSwitches("GREEN")) = 0 Then Chronicity = "Chronic"
    End If
   Else
    DiagType = "Follow-Up"
    If Len(gobjShared.CmdLineSwitches("GREEN")) = 0 Then Chronicity = "Chronic"
   End If
   
   With objDxI
      'JT SCR 40955 - add new property to indicate if diagnosis selected from problem list
      'JT SCR 42399 - added support for problem list items selected from favorites list
      gobjAnPParent.AddDiagnosis .psDescription, .psPrefix, .psICD9, .psNCID, .psSnoID, Chronicity, _
                                DiagType, .psExtICD_Code, .psExtICD_Name, .psExtICD_NCID, bProblemListItem Or .psNotFromMedcin
   End With
   
   RaiseEvent pevItemProcessedAP
    
Exit Sub

ErrHnd:
   Call gobjShared.ShowVBError _
   (Err.Number, Err.Description, msMODULE_NAME & sPROC_NAME, Exe, vbCritical)

End Sub

'SCR 33239  3/13/03  Phil Crowder
'JT BEGIN SCR 36572 - Modified to use CHCSII_CONN object
Private Function IsAcuteProblem(ByRef objDxI As clsDXI_DxItem) As Boolean

On Error GoTo ErrHnd
  
Dim adoRS   As Recordset
Dim sSql    As String
  
   sSql = "SELECT Count(*) FROM ACUTE_DIAGNOSES_CODES WHERE DIAGNOSIS_CODE = '" & Trim$(objDxI.psICD9) & "'"
   Set adoRS = OpenRecordset(sSql)

   If Not adoRS.EOF And Not adoRS.BOF Then
      IsAcuteProblem = adoRS.Fields(0) > 0
   End If
  
ErrHnd:

   'Clean Up Objects
   If Not adoRS Is Nothing Then If adoRS.State = adStateOpen Then adoRS.Close
   Set adoRS = Nothing
  
End Function
'JT END SCR 36572

'SCR 33766 3/18/03  Phil Crowder
Private Function IsNewProblem(ByRef objDxI As clsDXI_DxItem) As Boolean
Dim objProblem       As Object

  On Error GoTo ErrHnd
  
  'JT SCR 36572 - Updated code to check IsAppmode
  If Not gobjShared.IsAppMode(modeCHCSI_GUI) Then
     '''Load the Patients AcuteproblemList
     '''SCR 16163  Phil Crowder 10/22/01 Added active filter to search for duplicates
     'scr 21222 Phil Crowder 2/6/02
     Set objProblem = gobjAnPParent.RetrieveAcuteProblemList.Problems.FindSnoID(objDxI.psSnoID, objDxI.psPrefix)
     
     If Not (objProblem Is Nothing) Then
        IsNewProblem = Not (objProblem.StatusNCID = NCID_ACTIVE)
     Else
        IsNewProblem = True
     End If
  Else
    IsNewProblem = True
  End If

  Exit Function

ErrHnd:

End Function

Private Sub M_UpdateSearchPatternList()
'''************************************************************************
'''Called By:     M_bSearchSno
'''
'''Arguments:
'''
'''Description:   Store search patterns to allow user
'''               reusing them without retyping
'''
'''Note:
'''Warning:
'''Updates Globals:
'''Returns:
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "M_UpdateSearchPatternList"

   Dim lIdx            As Long
   Dim bPatternExist   As Boolean

On Error GoTo ErrHnd
    
   With pcboSearchDiagnosis
      '''SCR 13294 Only list search once
      '''If current search pattern is not stored add it to the list
      For lIdx = 0 To .ListCount - 1
           If .List(lIdx) = .Text Then
               bPatternExist = True
               Exit For
           End If
       Next
       If Not bPatternExist Then
           .AddItem .Text
       End If
      .SelStart = 0
      .SelLength = Len(pcboSearchDiagnosis.Text)
   End With
    
Exit Sub
ErrHnd:
   Call gobjShared.ShowVBError(Err.Number, Err.Description, msMODULE_NAME & sPROC_NAME, Exe, vbCritical)

End Sub


Public Property Let plDXPLM_DxPickListMode(ByVal lVal As pnumDXPLM_DxPickListMode)
'''************************************************************************
'''Called By:     pcboSearchDiagnosis_Click, pcboSearchDiagnosis_GotFocus or any
'''
'''Arguments:
'''
'''Description:   Update GUI according to the specified Mode
'''               If search Mode Show Search grid and hide Persistent Items Grid
'''               for all other modes show Persistent Items Grid and hide Search grid
'''
'''Note:
'''Warning:
'''Updates Globals:
'''Returns:
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "Let plDXPLM_DxPickListMode"

On Error GoTo ErrHnd

   mlDXPLM_DxPickListMode = lVal
   If mlDXPLM_DxPickListMode = lDXPLM_SnoSearch Then
      
      poptPickListSearch.value = True     '''Deselect all Persistent Items pick list options
      psmdDx.Visible = True               '''Show Search Grid
      pdxgDiagnosisGrid.Visible = False   '''Hide Persistent Items Grid
      pcmdSearch.Enabled = True           '''Indicate that search option is available
      
      With pcboSearchDiagnosis
         '''If user did not already select one of the available patterns
         If .Text = "" Then
            '''Restore(show user) search pattern associated with currently displayed search output
            .Text = mlLastSearchPattern
            '''Allow deleting the pattern with a single key press
            .SelStart = 0
            .SelLength = Len(.Text)
         End If
      End With
      
      
   Else
      psmdDx.Visible = False              '''Hide Search Grid
      pdxgDiagnosisGrid.Visible = True    '''Show Persistent Items Grid
      
      'SCR 22260  Phil Crowder 3/16/02
      'pcmdSearch.Enabled = False          '''Indicate that search option is not available
      'pcboSearchDiagnosis.Text = ""       '''Hide search pattern
      

   End If
     
Exit Property
ErrHnd:
   Call gobjShared.ShowVBError(Err.Number, Err.Description, msMODULE_NAME & sPROC_NAME, Exe, vbCritical)
   
End Property

Public Property Get plDXPLM_DxPickListMode() As pnumDXPLM_DxPickListMode
   plDXPLM_DxPickListMode = mlDXPLM_DxPickListMode
End Property

Private Sub Class_Terminate()

On Error Resume Next

   Set pobjParentDx = Nothing
   Set mobjListManagementOps = Nothing
   Set pddcDODCentricCodes = Nothing
'   Set pcvmConvertMedcin = Nothing
   Set poptPickListSearch = Nothing
   Set pdxgDiagnosisGrid = Nothing
   Set psmdDx = Nothing
   Set pcmdAddToFavorites = Nothing
   Set pcmdSearch = Nothing
   Set pcboSearchDiagnosis = Nothing
   Set pcmdAddToAP = Nothing

End Sub

Private Sub pcboSearchDiagnosis_KeyPress(KeyAscii As Integer)
'''************************************************************************
'''Called By:     None
'''Arguments:
'''Description:   Interpret enter key as request to initiate
'''               Search Diagnosis operation
'''Note:
'''Warning:
'''Updates Globals:
'''Returns:
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "pcboSearchDiagnosis_KeyPress"

On Error GoTo ErrHnd

   pcmdSearch.Enabled = True
   If KeyAscii = vbKeyReturn Then
      Call M_bSearchSno
      plDXPLM_DxPickListMode = lDXPLM_SnoSearch
   End If

Exit Sub
ErrHnd:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, msMODULE_NAME & sPROC_NAME, Exe, vbCritical)

End Sub


Private Sub pcmdAddToAP_Click()
'''************************************************************************
'''Called By:     None
'''
'''Arguments:
'''
'''Description:   Add selected Items to A&P
'''
'''
'''Note:
'''Warning:
'''Updates Globals:
'''Returns:
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "pcmdAddToAP_Click"

On Error GoTo ErrHnd

   Call M_AddSelectedDxItems(True)
   
Exit Sub
ErrHnd:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, msMODULE_NAME & sPROC_NAME, Exe, vbCritical)

End Sub

Private Sub pcmdAddToFavorites_Click()
'''************************************************************************
'''Called By:     None
'''
'''Arguments:
'''
'''Description:   Functionality for R1+ SCR-13882
'''               Allow user to add to the favorite list in A/P
'''
'''Note:
'''Warning:
'''Updates Globals:
'''Returns:
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "pcmdAddToFavorites_Click"

On Error GoTo ErrHnd

   Call M_AddSelectedDxItems(False)

Exit Sub
ErrHnd:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, msMODULE_NAME & sPROC_NAME, Exe, vbCritical)

End Sub


Private Sub pcmdSearch_Click()
'''************************************************************************
'''Called By:     None
'''
'''Arguments:
'''
'''Description:   Initiate Search Diagnosis operation
'''
'''
'''Note:
'''Warning:
'''Updates Globals:
'''Returns:
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "pcmdSearch_Click"
   
On Error GoTo ErrHnd
    
   Call M_bSearchSno
   plDXPLM_DxPickListMode = lDXPLM_SnoSearch
    
Exit Sub
ErrHnd:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, msMODULE_NAME & sPROC_NAME, Exe, vbCritical)

End Sub

Private Sub pdxgDiagnosisGrid_gridCollapsed(ByVal row As Long, ByVal Col As Long)
   'JT BEGIN SCR 35298 - Clear grid selection when collapsing to prevent inadvertently adding
   '                     non visible items when doubleclicking on [+] or [-] in the grid
   '                     This only happens when the mouse column is 0
   pdxgDiagnosisGrid.gridSelect 0, 0
   'JT END SCR 35298
End Sub

Private Sub pdxgDiagnosisGrid_gridDblClick(row As Long, Col As Long)
   Call M_DxPersistedItemsAdd(True)
End Sub


Private Sub psmdDx_DblClick()
   Call M_AddSelectedDxItems(True)
End Sub

Private Sub psmdDx_SearchBegin()
   Screen.MousePointer = vbHourglass
End Sub

Private Sub psmdDx_SearchComplete(ByVal Result As actxSearchSnoMed.enumSearchResult, ByVal Message As String, CancelMessage As Boolean)
   Screen.MousePointer = vbDefault
   mlLastSearchResult = Result
End Sub
