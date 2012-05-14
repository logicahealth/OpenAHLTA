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
Attribute VB_Name = "Dental"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'A9999
Private Const A9999_TxCode As Long = 506

Private mbSVGDrawn As Boolean
Private mbEnable603aPrint As Boolean

Public mcolDentalChild                    As Collection
Public mDentalChild                       As DentalChild
Public mLastDentalDate                    As String

Public sPtnEncID        As Long
Public sEncID           As Long
Public sEncDTS          As Date
Public sFclID           As String
Public sPtnID           As String
Public sPrimProvID      As Long

Public svgCommObj As srvSVGComm.clsSVGComm
Public NotesObj As CHCSII_Dental_DXTX.clsNotes

Private mobjDentalPatient As Object 'DentalPatient
Private mobjSBE As IntrfcCHCSII_Dental.IDENTAL_SBE
Private mobjDentalEncounter As IntrfcCHCSII_Dental.IDental_Enc

Private mblnStoredProc As Boolean
Private mblnDRCChanged As Boolean

Private mblnDebugLogging As Boolean
'Public svgCommFactory As New srvSVGComm.SingletonFactory 'hwcc 2/16/2002 we only need the factory once in initialize_objects

Const NCID_DENTAL_READINESS_CLASSIFICATION = "14501725"
Const NCID_LAST_DENTAL_READINESS_CLASSIFICATION_CHANGE_DATE = "14501728"
Const NCID_LAST_DENTAL_EXAM_DATE = "14501714"

Const cDefaultDate As String = "1/1/1920"
Const cNull As String = "NULL"

#If DEBUGON Then
Private Const Module = "Dental.cls"
#End If
Private Const errAppTitle = "CHCSII_Dental"

Implements ICHCSII_CoreAPI2
Implements ICHCSII_CoreAPI3
Implements ICHCSII_User
Implements ICHCSII_Dental
Implements IChild

Public Enum childTab
   BigTooth = 0
'   FullMouth = 1
   Notes = 1
   Labs = 2
   DataViews = 3
   PerioChart = 4
   Forensics = 5
   childTabCount = 6
End Enum

'Private Dental_Log As Boolean
Private msDentalEncounterID As String
Private mlDentalEncounterID As Long   'This is set in M_bOpenDentalEncounter
Private mobjCariesProtocol As CHCSII_Dental_NotesPlus.clsCariesRisk
Private mstrProvideFullName As String  'This is safed to display Workload, use before Encounter is signed
Private mbNewDentalEncounter As Boolean       'Flag it if this is a new dental encounter
Private mbDisplayRequiredSBE As Boolean         'Flag to display the string "SBE Prophylaxis required for future encounters"

Private mPatientLevelDRC As Integer
Private mPatLevelLstExmDte As String
Private mPatLevelLstChgDte As String

Private mEncLevelDRC As Integer
Private mDRCLstExamDte As String 'scr 20759
Private mDRCLstChgDte As String 'scr 20759

Private mblnDentalEncLoaded As Boolean
Private mstrApptDTS As String
Private mstrMiscRTF As String

Private mblnSuppressDRCDisp As Boolean 'scr 26121
Private mblnUseStoredProc As Boolean

'

Private Sub CWSafeMsgBox(Message As String)
  cwobjCWComm.Message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_DDSA_APPLICATION
  MsgBxARMd Message
  cwobjCWComm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_DDSA_APPLICATION
End Sub

Private Sub MyMsgBox(MyPrompt, Optional MyButton As VbMsgBoxStyle = vbOKOnly, Optional MyTitle = "", Optional MyHelpFile = 0, Optional MyContext As VbMsgBoxResult = vbOK)
  MsgBxARMd MyPrompt, MyButton, MyTitle, MyHelpFile, MyContext
End Sub

Public Function GetDentalEncounterIDStr() As String
  GetDentalEncounterIDStr = msDentalEncounterID
End Function

Public Function GetDentalEncounterIDLong() As Long
  GetDentalEncounterIDLong = mlDentalEncounterID
End Function

Private Sub ReplaceExtractWithMissingAndUpdateStatuses(lDntID As Long)

'This sub replaces completed extracts tx with an ongoing missing dx.

Dim oGroup As CHCSII_Dental_DXTX.clsGroup
Dim oDx As CHCSII_Dental_DXTX.clsDx
Dim oTx As CHCSII_Dental_DXTX.clsTx
Dim oSS As CHCSII_Dental_DXTX.clsSurfaceSet
Dim oTooth As CHCSII_Dental_DXTX.clsTooth
Dim sMissingTeeth As String

'    sMissingTeeth = ","
    
    'Go through each group
    For Each oGroup In NotesObj.Groups
        'Go through each Dx in the group
        For Each oDx In oGroup.Dxs
            If oDx.DentalEncounterId = lDntID Then
                'Sign the dxs created in this encounter
                oDx.Signed = True
            End If
            
            'Change the status from new to ongoing for dxs modified in this encounter
            If oDx.ModifiedInDentalEncounterID = lDntID Then
                If oDx.StatusID = dxActiveNew Then
                    oDx.StatusID = dxActiveOngoing
                End If
            End If
            
'            'Get a string of teeth ids that have an active missing dx
'            If oDx.Chartable = True Then
'                If oDx.DxICDCodeID = 1476 And (oDx.StatusID = dxActiveFinding Or oDx.StatusID = dxActiveNew Or oDx.StatusID = dxActiveObserve Or oDx.StatusID = dxActiveOngoing Or oDx.StatusID = dxActiveProvisional) Then
'                    sMissingTeeth = sMissingTeeth & oDx.SurfaceSets(1).Teeth(1).ToothID & ","
'                End If
'            End If
            
        Next
        
        'Go through each Tx in the group
        For Each oTx In oGroup.Txs
            
'            'Replace extract with missing
'            If (oTx.TxADACodeID = 422 Or oTx.TxADACodeID = 122 Or oTx.TxADACodeID = 176) And oTx.StatusID = txCompleted And oTx.DentalEncounterId = lDntID Then
'                'Found a completed extract tx
'                'Add a Dx of missing with a status of ongoing to that group with the same location
'                Set oDx = oGroup.Dxs.Add(sCharts, 1476, NotesObj.ProviderID)
'                oDx.StatusID = dxActiveOngoing
'                oDx.Chartable = True
'                Set oSS = oDx.SurfaceSets.Add()
'                oSS.SurfaceSet = "777"
'                oSS.State = False
'                oSS.SurfaceAbbr = " "
'                oSS.SymbolID = 151
'                Set oTooth = oSS.Teeth.Add(NotesObj.DentalEncounterId)
'                oTooth.ToothID = oTx.SurfaceSets(1).Teeth(1).ToothID
'                oTooth.ToothName = oTx.SurfaceSets(1).Teeth(1).ToothName
'            End If
            
            If oTx.DentalEncounterId = lDntID Then
                'Sign the txs created in this encounter.
                oTx.Signed = True
            End If
            
            If oTx.ModifiedInDentalEncounterID = lDntID Then
                'Change the status from Completed to Pre-Existing for the txs modified in this encounter
                If oTx.StatusID = txCompleted Then
                    oTx.StatusID = txPreExisting
                End If
            End If
        Next
    Next

    NotesObj.SaveGroups
    
    Set oGroup = Nothing
    Set oDx = Nothing
    Set oTx = Nothing
    Set oSS = Nothing
    Set oTooth = Nothing
    
End Sub

Public Sub SBE_EntryFormDisplay(DentalEncounterId As String, _
                                PatientUnitId As String, Optional CallFromOutSideParent As Boolean = True)
     Dim stMousePointer As VBRUN.MousePointerConstants
     stMousePointer = Screen.MousePointer
     Screen.MousePointer = vbDefault
     
     frmSBE.SBEForceDisplay_Set CallFromOutSideParent
     frmSBE.Dental_Encounter_ID = DentalEncounterId
     frmSBE.Patient_Unit_ID = PatientUnitId
     'frmSBE.DentalParent_Set Me
     cwobjCWComm.Message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_DDSA_APPLICATION
     frmSBE.InitSBEForm
     If frmSBE.SBERequireFlag_Get Then
        Load frmSBE
        frmSBE.Show vbModal
     End If
     cwobjCWComm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_DDSA_APPLICATION
     frmSBE.DentalParent_Clear
     frmSBE.SBEForceDisplay_Clear
     
     mbDisplayRequiredSBE = False
     If CallFromOutSideParent Then mbDisplayRequiredSBE = True
            
     Screen.MousePointer = stMousePointer
End Sub

Public Function GetSBENotes(DentalEncounterId As String) As String
    Dim objDAS As ICHCSII_DAS
    Dim objRS As Recordset
    Dim sqlStr As String
    Dim vNotes As String
    Dim blnDisplayFutureSBE As Boolean
    
    'KDunne 6/18/2002
    'SCR 24690
    blnDisplayFutureSBE = True
    
    Set objDAS = GetDasObj()
    sqlStr = "Select SBE, SBE_COMMENTS from DENTAL_ENCOUNTER where DENTAL_ENCOUNTER_ID =" & DentalEncounterId
    Set objRS = objDAS.OpenRecordset(sqlStr)
    If objRS.RecordCount > 0 Then
      'Add SBE and SBE Comments
      vNotes = vNotes & "\viewkind4\uc1\pard\ul\b\f0\fs20 SBE Information:\ulnone\b0\par SBE: "
      Select Case CInt(objRS("SBE"))
        Case 1
          vNotes = vNotes & frmSBE.OptionTaken.Caption
        Case 2
          vNotes = vNotes & frmSBE.OptionNotTaken.Caption
        Case 3
          vNotes = vNotes & frmSBE.OptionNotRequired.Caption
        Case 4
          vNotes = vNotes & frmSBE.OptionNoLongerRequired.Caption
          blnDisplayFutureSBE = False
        Case Else
            blnDisplayFutureSBE = False
      End Select
      If "" <> objRS("SBE_COMMENTS") Then
        vNotes = vNotes & "\par SBE Comments: " & objRS("SBE_COMMENTS") & "\par "
      Else
        vNotes = vNotes & "\par "    'No SBE Comment
      End If
      If blnDisplayFutureSBE Then vNotes = vNotes & "SBE Requirement: SBE Prophylaxis required for future encounters. \par"
    End If
    GetSBENotes = vNotes
    Set objRS = Nothing
    Set objDAS = Nothing
 
End Function

Public Sub PreEncounterSign()
  'CWSafeMsgBox "PreEncounterSign got called"     'hwcc 2/27/2002 change MsgBox to CWSafeMsgBox
  On Error GoTo ErrHandler
  
  'KDunne 7/30/2002
  'SCR 26121
  mblnSuppressDRCDisp = True
  
  If mobjCariesProtocol Is Nothing Then
    InitCariesProtocol
  End If

  'SF 12-16-03 SCR 47086
  mblnSuppressDRCDisp = False

  mobjCariesProtocol.CheckProtocolComplete 'SCR 17133
  
  'Unlike Notes, DataViews may not be initialized before Encounter Signed
  ' SCR 17424 and 17821
  Dim objMod As Object
  Set objMod = cwobjCWComm.InitializeOLEServer(glServerInfo(DataViews))

  If objMod Is Nothing Then Exit Sub
  Call objMod.DisplayUnsignedWorkload(CStr(sEncID), mstrProvideFullName)
  
FuncExit:
  Set objMod = Nothing
  Exit Sub
  
ErrHandler:
  Call gobjShared.ShowVBError(Err.Number, Err.Description, "Dental.PreEncounterSign", errAppTitle, vbExclamation)
  GoTo FuncExit
End Sub

Private Sub InitCariesProtocol()
    Dim objNotesSF As CHCSII_Dental_NotesPlus.clsSingleton
    Dim objSection As CHCSII_Dental_NotesPlus.clsISection
    Dim objMod As Object

    On Error GoTo ErrHandler
    
    If M_bOpenDentalEncounter Then
        'KDunne 7/18/2002
        'SCR 25763 - Sometiimes when signing the encounter the datalayer and svg comm
        '            objects are set to nothing. Make sure that they are initialized.
        '
        If NotesObj Is Nothing Then
            If svgCommObj Is Nothing Then
                initialize_objects
            End If
            If Not (svgCommObj Is Nothing) Then Set NotesObj = svgCommObj.DataObject
        End If
        'End scr 25763
        
        Set objMod = cwobjCWComm.InitializeOLEServer(glServerInfo(Notes))
        Set objNotesSF = New CHCSII_Dental_NotesPlus.clsSingleton
        Set mobjCariesProtocol = objNotesSF.Caries_Risk
        Set mobjCariesProtocol.DataLayer = NotesObj
        Set mobjCariesProtocol.DentalEncounter = mobjDentalEncounter
        objMod.PrimProviderID = sPrimProvID
        mobjCariesProtocol.ProviderID = sPrimProvID
        
        Set objSection = mobjCariesProtocol
        Call objSection.LoadSection(sPtnEncID, mlDentalEncounterID)
    End If

FuncExit:
    Set objNotesSF = Nothing
    Set objSection = Nothing
    Set objMod = Nothing
    Exit Sub
    
ErrHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "Dental.InitCariesProtocol", errAppTitle, vbExclamation)
    GoTo FuncExit
    Resume
End Sub
Private Sub DisplayCariesRiskProtocolScrn()
  'SCR 17096, 17098, 17099, 17119
  Dim oSection As CHCSII_Dental_NotesPlus.clsISection
  Dim objNotesSF As CHCSII_Dental_NotesPlus.clsSingleton
  
  InitCariesProtocol
  
  mobjCariesProtocol.LoadProtocolForm
  'Make sure we save any changes that may have been made by the user
    Set oSection = mobjCariesProtocol
    oSection.SaveSection
  
  Set oSection = Nothing
  Set objNotesSF = Nothing
End Sub

Private Sub SetEncDentalNotes(msDentalEncounterID As String)
    Dim vNotes As String
    Dim objDAS As ICHCSII_DAS
    Dim objRS As Recordset
    Dim sqlStr As String
    Dim strCaries As String
    Dim strPreventive As String
    Dim objProph As CHCSII_Dental_NotesPlus.clsProph
    Dim objSection As CHCSII_Dental_NotesPlus.clsISection
    
    On Error GoTo ErrHandler
    
    vNotes = mobjSBE.RTFComments
    
    If Not (NotesObj Is Nothing) Then vNotes = vNotes & NotesObj.GetNote & vbCrLf
    
    Set objSection = mobjCariesProtocol
    strCaries = objSection.GetRTFNote
    'KDunne 12/18/2003
    'SCR 46797, 46798
    '
    If Trim(strCaries) <> "" Then
        strCaries = "\ul " & "Caries Risk Protocol Status:" & "\ulnone\par " & strCaries & "\par "
    End If
    
    If Trim(mstrMiscRTF) <> Empty Then
        vNotes = vNotes & mstrMiscRTF
    End If

    Dim lngTimer As Long
    lngTimer = Timer
If mblnUseStoredProc Then
    Dim objCmd As CHCSII_Command.CHCSII_Cmd
    
    'Set objDAS = New CHCSII_DataAccess.CHCSII_DASsP
    Set objDAS = GetDasObj
    Set objCmd = New CHCSII_Command.CHCSII_Cmd
    
    objCmd.SPName = "DENT_NP_GET_NARR_SUMM"
    'We must add paramaters in the same order as the appear in the Stored Proc Signature
    Call objCmd.AddParam(False, False, "textString", msDentalEncounterID)

    'This Parameter is defined as a cursor, used for output
    Call objCmd.AddParam(True, True, "", "")
    Set objRS = objDAS.OpenRecordsetSP(objCmd)
Else
    Set objDAS = GetDasObj()

    sqlStr = "SELECT EXAM_EVAL, TREATMENT, RADIOLOGY, MATERIALS, PREVENTIVE, Exam_General, Exam_SOAP, Exam_Endo, Exam_Perio, Exam_Pros, Perio_Charting " & _
              " From TEMPL_NARRATIVE_SUMMARY WHERE " & _
                   " DENTAL_ENCOUNTER_ID =" & msDentalEncounterID & _
                   " ORDER BY ID"
                   
    Set objRS = objDAS.OpenRecordset(sqlStr)

End If
    Debug.Print "Time To Load TEMPL_NARRATIVE_SUMMARY = " & (lngTimer - Timer) & " seconds "
    Dim strExamEval As String
    Dim strTreatment As String
    Dim strRadiology As String
    Dim strMaterials As String
    Dim strPSR As String
    Dim strOrientation As String
    Dim strExamGeneral As String
    Dim strExamSOAP As String
    Dim strExamEndo As String
    Dim strExamPerio As String
    Dim strExamPros As String
    Dim strPerioCharting As String
    
    If objRS.RecordCount > 0 Then
      Do Until objRS.EOF
        strExamEval = strExamEval & objRS("EXAM_EVAL")
        strPreventive = strPreventive & objRS("PREVENTIVE")
        'KDunne 4/26/2002
        'SCR 8570
        strTreatment = strTreatment & objRS("TREATMENT")
        strRadiology = strRadiology & objRS("RADIOLOGY")
        strMaterials = strMaterials & objRS("MATERIALS")
        
        strExamGeneral = strExamGeneral & objRS("Exam_General")
        strExamSOAP = strExamSOAP & objRS("Exam_SOAP")
        strExamEndo = strExamEndo & objRS("Exam_Endo")
        strExamPerio = strExamPerio & objRS("Exam_Perio")
        strExamPros = strExamPros & objRS("Exam_Pros")
        strPerioCharting = strPerioCharting & objRS("Perio_Charting")
        
        objRS.MoveNext
      Loop
      vNotes = vNotes & strExamGeneral & strExamSOAP & strExamEndo & strExamPerio & strExamPros & strPreventive & strTreatment & strRadiology & strMaterials & strPerioCharting
    End If
    
    strPSR = getPSR
'    strOrientation = getOrientation
    vNotes = vNotes & strPSR & strOrientation & strCaries

    If Not IsEmpty(vNotes) Then
        Call gobjEnc.SetRTF(vNotes)
    End If
    
FuncExit:
    Set objRS = Nothing
    Set objDAS = Nothing
    Set objCmd = Nothing
    
    Exit Sub
    
ErrHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "Dental.SetEncDentalNotes", errAppTitle, vbExclamation)
    GoTo FuncExit
End Sub

Friend Function MtdF_bReadyToChangeTab(ByVal lTo As childTab) As Boolean
   Dim vNotes        As String 'Variant
   Dim iDRC         As Integer
   Dim childServer    As DentalChild
   Dim bResult        As Boolean
   
On Error GoTo ErrHnd
     
    'SCR-ARM-30553 RWalker 12/20/02
    If lTo < 0 Then     'RWalker 12/20/02
        lTo = BigTooth  'RWalker 12/20/02
    End If              'RWalker 12/20/02
     
   If lTo = BigTooth Or lTo = DataViews Then
        svgCommObj.TeethObject.clearSelectedTeeth
   End If
      
    'KDunne 8/1/2002
    'SCR 26236 - When changing to Dataviews, must set Encounter note. At this point,
    '            the previous dental child has already saved his information.
'    If lTo = DataViews Then SetEncDentalNotes msDentalEncounterID
    If lTo = DataViews Then SaveEncounter False, False
   Set childServer = mcolDentalChild(CStr(giCurrentServer))
   'Set childServer = mcolDentalChild(CStr(lTo))
   '''If ready to release
    
    'SCR 45991 RWalker 12/12/2003 Since the original Open BigTooth
    'has been replaced with either the FullMouthGC or BigToothGC
    'transaction we need to fire these when the BigTooth tab is clicked,
    'but only when clicked on the Dental form NOT in the Core Folder List
    If (childServer.ServerIndex = BigTooth) Then
      bResult = childServer.Module.Mtd_bDentalTabChange(giCurrentServer, lTo, gbFromCore)
    Else
      bResult = childServer.Module.Mtd_bDentalTabChange(giCurrentServer, lTo)
    End If
   
   Set childServer = Nothing  'hwcc 2/26/2002
   
   'If lTo = 3 Then
   If giCurrentServer = Notes Then      'hwcc we are leaving notes
        ' get notes from notes and save
     SetEncDentalNotes msDentalEncounterID
   End If
      
   If bResult Then
      Set childServer = mcolDentalChild(CStr(lTo))
      'Set childServer = mcolDentalChild(CStr(giCurrentServer))
      '''If ready to acccept
        
      'SCR 45991 RWalker 12/12/2003 Since the original Open BigTooth
      'has been replaced with either the FullMouthGC or BigToothGC
      'transaction we need to fire these when the BigTooth tab is clicked,
      'but only when clicked on the Dental form NOT in the Core Folder List
      If (childServer.ServerIndex = BigTooth) Then
        bResult = childServer.Module.Mtd_bDentalTabChange(giCurrentServer, lTo, gbFromCore)
      Else
        bResult = childServer.Module.Mtd_bDentalTabChange(giCurrentServer, lTo)
      End If
      Set childServer = Nothing 'hwcc 2/26/2002
   End If
   
   'KDunne scr 31172
'   'Some Dental Children will update core menu's, make sure core is  refreshed.
'   cwobjCWComm.Message cwiGENERIC, "Refresh", NCID_CW_APPLICATION, NCID_DENTAL_LABS_APPLICATION

   MtdF_bReadyToChangeTab = bResult

Exit Function
ErrHnd:
   MtdF_bReadyToChangeTab = False
   Call gobjShared.ShowVBError(Err.Number, Err.Description, "Dental.MtdF_bReadyToChangeTab", errAppTitle, vbExclamation)
   
Exit Function
End Function


Public Property Get psDentalEncounterID() As String
   psDentalEncounterID = msDentalEncounterID
End Property

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
         QueryView = True              '--this view is NOT optional--
      Case cwiPATIENT_SUMMARY_VIEW
         QueryView = False
      Case cwiCLINICIAN_SUMMARY_VIEW
         QueryView = False
      Case cwiENCOUNTER_SUMMARY_VIEW
         QueryView = False
      Case Else                        '--unknown "view type", not supported--
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
Public Function OpenView(ByVal viViewType As Integer, _
                          ByVal vlParenthWnd As Long, _
                          rhMainForm As Long) _
                          As Boolean
                          
   Dim objReg      As RegObj.Registry
   Dim objKey      As RegObj.RegKey
   Dim sKeyName    As String
    Dim oTx As clsTx
    Dim bFoundTx As Boolean
    Dim oGroup As clsGroup
    
   On Error Resume Next

    If gobjUser.HasPrivilegeEx(Priv_Dental_Access, Priv_Write) = False Then
        MsgBxARMd ("The provider does not have dental access privileges.")
        Exit Function
    End If
   sKeyName = "\HKEY_LOCAL_MACHINE\SOFTWARE\CHCSII\Dental_Log"
   Set objReg = New RegObj.Registry
   Set objKey = objReg.RegKeyFromString(sKeyName)
   If Not objKey Is Nothing Then
       Dental_Log = True
   Else
       Dental_Log = False
   End If
    
   Set objKey = Nothing
   Set objReg = Nothing

   Dim iDRC As Integer
   Dim bResult As Boolean
   'Dim bSave As Boolean
   Dim stMousePointer As Variant
   Dim oEncDoc As Encounter
   
   On Error GoTo ErrHandler
      
   mPatientLevelDRC = 0
    

   
    If Not gobjEnc Is Nothing Then
   
        mbEnable603aPrint = True
        EnablePrintIcon
'        ' 7/08/2002
'        'WR IParent.CurrentEncounter() now returns IEncDoc interface SCR 25225
'        Set oEncDoc = gobjEnc.CurrentEncounter
'
'        'KDunne 7/15/2002
'        'SCR 25095
'        cwobjCWComm.DeleteMenuItem (glMenuPrintID)
'        glMenuPrintID = gobjEnc.AddMenuItem(0, "&SF603A", "Print.bmp", "Preview SF603A")
    End If
   
   '-- -perform actions based on the "view type"---
   If viViewType = cwiREGULAR_VIEW Then
   
      '''''''
      If Dental_Log Then
        Open "c:\temp\dental.log" For Append As #1
        Print #1, ""
        Print #1, CStr(Now) & ",OpenView(),Patient-" & sPtnEncID & ",Dental Encounter-" & msDentalEncounterID
        Close #1
      End If
      '''''''
      
      mbNewDentalEncounter = False       'Don't know if this is a new Dental Encounter
      mblnDentalEncLoaded = False        ' Flag get set inside M_bOpenDentalEncounter
      If M_bOpenDentalEncounter Then
      
        initialize_objects
                
        If Not svgCommObj Is Nothing Then
           svgCommObj.DentalEncounterId = mlDentalEncounterID
           svgCommObj.ProviderID = sPrimProvID
'Notes extracted from Jon's e-mail - This line will cause the TeethDocTx to have to refresh everytime
' we go into and out of dental.  Only unless we change patients should this line
'  be executed.  This is handled by CHCSII_Dental_SVG .
'    Jonathan
' hwcc 3/4/02
'           Set svgCommObj.TeethObject.TeethDocTx = Nothing
        End If
  
        If Not gobjShared.CmdLineSwitches("DENTALTRIAGE") = "DENTALTRIAGE" Then
            If mobjDentalEncounter.NewEncounter Then
                mobjDentalEncounter.SBE.GetUserInput msDentalEncounterID, sPtnID, False
            End If
        End If
      
        'SF 10/21/03
        'Add the A9999 workload code to every encounter, since the create encounter is also in the SP we can't add the Tx in the SP, so we need to check if the A9999 already exists.
        For Each oTx In NotesObj.AllTxs
            If oTx.TxADACodeID = A9999_TxCode And oTx.DentalEncounterId = mlDentalEncounterID Then
                bFoundTx = True
                Exit For
            End If
        Next
        Set oTx = Nothing
    
        If bFoundTx = False Then
            'Add the Tx
            Set oGroup = svgCommObj.DataObject.Groups.Add
            Set oTx = oGroup.Txs.Add(sNotes, A9999_TxCode, svgCommObj.DataObject.ProviderID)
            oTx.StatusID = txCompleted
            Set oGroup = Nothing
        End If
    
      
        Dim objMod As Object
        Set objMod = cwobjCWComm.InitializeOLEServer(glServerInfo(Notes))
      
         DisplayCariesRiskProtocolScrn
         
        Set objMod = Nothing   'hwcc 2/26/2002
           
         '-- -set global flags---
         cwbFirstRegularResize = True
         cwbMainViewRunning = True
         
         '-- -------------------------------------------------------------------------------------
         '--  If this OLE Server needs it's own menu items, here is a good place to add these
         '--    menu items onto the CW Core's main menu.
         '--
         '--  NOTE: This is OPTIONAL.  You do NOT have to create menu items. Thats why they are
         '--        commented out.
         '-- -------------------------------------------------------------------------------------
         '-- -add a menu item "Dental" on CW's main menu---
         '-- ghDental = cwobjCWComm.AddPopupMenuToCW("&Dental", cwlMyProgId)
         '-- -add a menu item "Ole Server Menu Item" to the menu "Dental"---
         '-- glOleMenuItem = cwobjCWComm.AddMenuItem("&Ole Server Menu Item", ghDental, cwlMyProgId, "Dental menu item")
         
         '-- ---------------------------------------------------------------------------------------
         '--  Reset the "parent window" for this OLE server's "main" regular form.  This is the main
         '--    "trick" that makes the OLE server look like it is part of the CW Core.
         '--
         '--  NOTE: This action invokes the main form's "Load" event, but the form is NOT shown.
         '--        Make sure that you DO NOT perform a "Me.Show" in the form's "Load" event, as this
         '--        will mess up the "SetParent" API call.  Make sure that you DO NOT "Show" the
         '--        "main" regular form in this method, that will happen in the "ResizeView" method.
         '-- ---------------------------------------------------------------------------------------
         
         Set frmDental = New frmDental
         SetParent frmDental.hwnd, vlParenthWnd
         Set frmDental.ParentObject = Me
         
         '-- -set parameter---
         rhMainForm = frmDental.hwnd
      Else
         OpenView = False
         Exit Function
      End If
   Else                                '--unsupported view "type"--
      '-- -set return value---
      OpenView = False
      Exit Function
   End If
   
   '-- -set return value---
   OpenView = True
   
   Set oEncDoc = Nothing
   Exit Function
   
ErrHandler:
   '-- -show the VB error---
   Call gobjShared.ShowVBError(Err.Number, Err.Description, "Dental<main>.OpenView", errAppTitle, vbCritical)
   '-- -set return value---
   OpenView = False
   Exit Function

End Function

Public Function InitializeDental(Optional IsEncounter As Boolean) As Boolean
    InitializeDental = False
    If IsEncounter Then
        ' prepare for signing dental encounter
        InitializeDental = True
    End If
                          
End Function

Public Function CompletedEncounter() As Boolean
   
    Dim vDntID           As Variant
    Dim bReread          As Boolean
    Dim lCount           As Long
   
    Dim strSQL As String
    Dim objDAS As ICHCSII_DAS
    Dim objRS As ADODB.Recordset
           
    Dim OverallDRC As Integer
    Dim sOverallLastChangeDate As String
    Dim iEncDRC As Integer
    Dim sOralExam As String
    Dim strLogMsg As String
    
    Dim objCmd As CHCSII_Command.CHCSII_Cmd
    
    On Error GoTo errHandle
    
    'KDunne 10/22/2002
    'SCR 28687
    If mblnDebugLogging Then
        strLogMsg = "Dental Encounter ID = " & mlDentalEncounterID & _
            ", Patient Level DRC = " & mPatientLevelDRC & ", Patient Level Last Exam Date = " & _
            mPatLevelLstExmDte & ", Patient Level Last Change Date = " & mPatLevelLstChgDte & _
            ", Encounter Level DRC = " & mEncLevelDRC & ", Encounter Level Last Exam Date = " & _
            mDRCLstExamDte & ", Encounter Level Last Change Date = " & mDRCLstChgDte & vbCrLf & _
            "Appointment Date/Time = " & mstrApptDTS
        Log_Event "Begin Save Encounter", strLogMsg
    End If
    
    CompletedEncounter = False
    
    initialize_objects
        
'    strSQL = "Select DENTAL_ENCOUNTER_ID, DRC from DENTAL_ENCOUNTER Where " _
'             & "Encounter_Number=" & Val(sEncID) & " AND " _
'             & "Encounter_Facility_NCID=" & sFclID & " AND " _
'             & "Patient_Unit_ID=" & sPtnID
'
'    Set objRS = objDAS.OpenRecordset(strSQL)
    
'    If Not objRS.EOF Then
'        vDntID = objRS(0)
'        iEncDRC = objRS.Fields("DRC")
'    End If
'    Set objRS = Nothing  'hwcc 2/26/2002
    
    vDntID = mobjDentalEncounter.DentalEncounterId
    iEncDRC = mEncLevelDRC
    
    
    
 
    If CLng(vDntID) > 0 Then
        
        'Replaces completed extract with missing
        ReplaceExtractWithMissingAndUpdateStatuses CLng(vDntID)
        
        If mblnUseStoredProc Then
            'Set objDAS = New CHCSII_DataAccess.CHCSII_DASsP
            Set objDAS = GetDasObj
            Set objCmd = New CHCSII_Command.CHCSII_Cmd
            
            objCmd.SPName = "DENT_PARENT_SIGN_ENCOUNTER"
            'We must add paramaters in the same order as the appear in the Stored Proc Signature
            Call objCmd.AddParam(False, False, "textString", mobjDentalEncounter.PatientID)
            Call objCmd.AddParam(False, False, "textString", vDntID)

            objDAS.ExecuteSP objCmd
        Else
            Set objDAS = GetDasObj()
            strSQL = "Update DENTAL_ENCOUNTER Set SIGNED=-1 Where DENTAL_ENCOUNTER_ID=" & CLng(vDntID)
            objDAS.ExecuteSQL (strSQL)
            
            'On Error Resume Next 'hwcc 3/4/02 fix for SCR 21211
            ' This block is specific for Inserting DENTAL_ENCOUNTER_DX/TX_NOTE
            '  Oracle split out an row not found error while deleting or inserting
            '   In this case we don't really need to do anything. In case the database
            '    backend is down, there will be problems with other SQL hit everywhere
            '     so for that case it can be spotted pretty easy and we don't worry about it too much
                    'Dx Delete
            strSQL = "DELETE FROM DENTAL_ENCOUNTER_DX_NOTE Where SIGNED_IN_DENTAL_ENCOUNTER_ID =" & CLng(vDntID)
            objDAS.ExecuteSQL (strSQL)
            
            'Tx Delete
            strSQL = "DELETE FROM DENTAL_ENCOUNTER_TX_NOTE Where SIGNED_IN_DENTAL_ENCOUNTER_ID =" & CLng(vDntID)
            objDAS.ExecuteSQL (strSQL)
            
            'Dx insert
            strSQL = EncSign_DxInsert_Insert & EncSign_DxInsert_Sel
            If goCHCSIIConn.UseSQLServer Then
                strSQL = strSQL & EncSign_DxInsert_From_SQLServer & EncSign_DxInsert_Where_SQLServer
            Else
                strSQL = strSQL & EncSign_DxInsert_From & EncSign_DxInsert_Where
            End If
            strSQL = strSQL & " ((DENTAL_ENCOUNTER.PATIENT_UNIT_ID = " & sPtnID & " ) AND " & "(DENTAL_ENCOUNTER.DENTAL_ENCOUNTER_ID = " & CLng(vDntID) & " ) AND " & _
                "(DENTAL_ENCOUNTER_DX.STATUS_ID <> 21) AND (DENTAL_ENCOUNTER_DX.SIGNED = 0) OR " & _
                "(DENTAL_ENCOUNTER.PATIENT_UNIT_ID = " & sPtnID & " ) AND " & "(DENTAL_ENCOUNTER.DENTAL_ENCOUNTER_ID = " & CLng(vDntID) & " ) AND " & _
                "(DENTAL_ENCOUNTER_DX.STATUS_ID BETWEEN 10 AND 14) AND (DENTAL_ENCOUNTER_DX.SIGNED = - 1)) "
            objDAS.ExecuteSQL (strSQL)
            
            'Tx
            strSQL = EncSign_TxInsert_Insert & EncSign_TxInsert_Sel
            If goCHCSIIConn.UseSQLServer Then
                strSQL = EncSign_TxInsert_From_SQLServer & EncSign_TxInsert_Where_SQLServer
            Else
                strSQL = EncSign_TxInsert_From & EncSign_TxInsert_Where
            End If
            strSQL = " ((DENTAL_ENCOUNTER.PATIENT_UNIT_ID = " & sPtnID & " ) AND " & "(DENTAL_ENCOUNTER.DENTAL_ENCOUNTER_ID = " & CLng(vDntID) & " ) AND " & _
                    "(DENTAL_ENCOUNTER_TX.STATUS_ID <> 104) AND (DENTAL_ENCOUNTER_TX.SIGNED = 0) OR " & _
                    "(DENTAL_ENCOUNTER.PATIENT_UNIT_ID = " & sPtnID & " ) AND " & "(DENTAL_ENCOUNTER.DENTAL_ENCOUNTER_ID = " & CLng(vDntID) & " ) AND " & _
                    "(DENTAL_ENCOUNTER_TX.STATUS_ID IN (100, 101, 102)) AND (DENTAL_ENCOUNTER_TX.SIGNED = - 1)) "
            
            'Clipboard.SetText strSQL
            objDAS.ExecuteSQL (strSQL)
        End If
           
        CompletedEncounter = True

    End If
       
'ASamedov 02/26/02
'Update Patient level DRC if encounter signed successfully
'DRC BizRules (Sign Encounter)
' # here readiness dt
sOverallLastChangeDate = mDRCLstChgDte

If sOverallLastChangeDate = "" Or sOverallLastChangeDate = cNull Then
    sOverallLastChangeDate = cDefaultDate
End If

If CompletedEncounter = True Then

    'KDunne 10/22/2002
    'SCR 28687
    If mblnDebugLogging Then
        strLogMsg = "Dental Encounter ID = " & mlDentalEncounterID & _
            ", Patient Level DRC = " & mPatientLevelDRC & ", Patient Level Last Exam Date = " & _
            mPatLevelLstExmDte & ", Patient Level Last Change Date = " & mPatLevelLstChgDte & _
            ", Encounter Level DRC = " & mEncLevelDRC & ", Encounter Level Last Exam Date = " & _
            mDRCLstExamDte & ", Encounter Level Last Change Date = " & mDRCLstChgDte
        Log_Event "Completed Encounter = True", strLogMsg
    End If

    'Update Overall (Patient level) DRC
    'First to check if the encounter's appt.date/time is later
    'than the Overall.LastChangeDate
    
    'KDunne 7/9/2002
    'SCR 20759
    'dEncDate = getEncApptDate
    
'    If DateDiff("s", Format(sOverallLastChangeDate, gcDATE_FORMAT_VB_TIME), mstrApptDTS) >= 0 Then
        '(to ensure that this encounter is the latest
    If (mPatLevelLstChgDte <> mDRCLstChgDte And mDRCLstChgDte <> "") Or _
        CDate(mPatLevelLstChgDte) = CDate(mstrApptDTS) Then
        OverallDRC = iEncDRC
        
        sOverallLastChangeDate = mstrApptDTS 'getEncApptDate

        sOralExam = getOralExam
        'If Encounter has Initial or Periodic Exam performed then
        'LastExamChangeDate should be also updated
        If sOralExam = "1" Or sOralExam = "2" Then
            Call SaveDentalPatient(OverallDRC, sPtnID, mstrApptDTS, mstrApptDTS)
        Else
            Call SaveDentalPatient(OverallDRC, sPtnID, "", mstrApptDTS)
        End If
        
        'to refresh DRC in Readiness module
        'send appropriate message to Readiness module
        cwobjCWComm.Message cwiREFRESH_DATA, "DRC|" & Str(OverallDRC), NCID_READINESS_APPLICATION, NCID_DDSA_APPLICATION, False
        
    End If

End If
'

    'hwcc 2/22/02 - for SCRs 21534, 8346, 17467
    If Not (NotesObj Is Nothing) Then
       Dim RtnCd As Boolean
       'SF Do not reload the datalayer because GC does that below
       'RtnCd = NotesObj.LoadGroups(CLng(sPtnID), CLng(vDntID), sPrimProvID)
       
       'SF Reload GC and the datalayer
       'cwobjCWComm.Message cwiREFRESH_DATA, "REFRESHPATIENT", NCID_DENTAL_SVG, NCID_DDSA_APPLICATION
       
       'CWSafeMsgBox "Called loadgroups:" & RtnCd   'This is for testing 'hwcc 2/27/2002 Change to CWSafeMsgBox
    Else
        CWSafeMsgBox "Dental parent did not call loadgroups"
    End If
    
    'KDunne 7/30/2002
    'SCR 26121
    mblnSuppressDRCDisp = False
    Call DisplayDentalPatient(0, sPtnEncID)
    
    terminate_Objects
    Set objDAS = Nothing
    Set objRS = Nothing
    
    Exit Function
   
errHandle:
    'SF 3-07-02 Added error handling. Before, it was just on error resume next for the whole function.
    If Err.Number = -2147221503 Then
        Resume Next
    Else
        Call gobjShared.ShowVBError(Err.Number, Err.Description, "Dental.CompletedEncounter", errAppTitle, vbCritical)
        Resume Next
    End If
                          
End Function

Public Function LockDentalEncounter() As Boolean
    LockDentalEncounter = False
    If Not gobjEnc Is Nothing Then
        LockDentalEncounter = gobjEnc.StartSave()
    End If

End Function

Public Function UnLockDentalEncounter() As Boolean
    UnLockDentalEncounter = False
    If Not gobjEnc Is Nothing Then
        UnLockDentalEncounter = gobjEnc.EndSave()
    End If
   
End Function

Public Sub LoadChildren(Optional bChildServer As Boolean = False, _
                         Optional iChildServer As Integer)

   Dim i              As Integer
   Dim iStart         As Integer
   Dim iEnd           As Integer
   Dim childServer    As DentalChild
   Dim bFound         As Boolean
   Dim oldStatusText  As String

#If DEBUGON Then
   Dim DR  As DebugRec

   If Not DL Is Nothing Then DL.DebugLog Exe, Module, "LoadChildren", "Begin", DR
#End If
   'Screen.MousePointer = vbHourglass
'   BigTooth = 0
'   FullMouth = 1
'   Notes = 2
'   Labs = 3
'   DataViews = 4

   '-- If Not DebugMode Then On Error GoTo ErrHandler

   If bChildServer Then
      '-- Exit Sub ' for testing
      iStart = iChildServer            ' for testing
      iEnd = iChildServer
   Else
      iStart = BigTooth
      iEnd = Forensics
   End If

   For i = iStart To iEnd
      '-- open the summary form for each module
      Set mDentalChild = New DentalChild
      mDentalChild.ServerIndex = i
      Select Case i
         Case BigTooth
            mDentalChild.ViewType = cwiREGULAR_VIEW
            mDentalChild.childHeight = 5100
            mDentalChild.childWidth = 11000
'         Case FullMouth
'            mDentalChild.ViewType = cwiREGULAR_VIEW
'            mDentalChild.childHeight = 5100
'            mDentalChild.childWidth = 11000
         Case Notes
            mDentalChild.ViewType = cwiREGULAR_VIEW
            mDentalChild.childHeight = 5100
            mDentalChild.childWidth = 11000
         Case Labs
            mDentalChild.ViewType = cwiREGULAR_VIEW
            mDentalChild.childHeight = 5100
            mDentalChild.childWidth = 11000
         Case DataViews
            mDentalChild.ViewType = cwiREGULAR_VIEW
            mDentalChild.childHeight = 5100
            mDentalChild.childWidth = 11280
         Case Forensics
            mDentalChild.ViewType = cwiREGULAR_VIEW
            mDentalChild.childHeight = 5100
            mDentalChild.childWidth = 11280
         Case PerioChart
            mDentalChild.ViewType = cwiREGULAR_VIEW
            mDentalChild.childHeight = 5100
            mDentalChild.childWidth = 11000

      End Select
      mDentalChild.Comm = cwobjCWComm
      mDentalChild.Config = cwobjConfig
      mDentalChild.LogonEnvironment = cwobjLogonEnviron
      mDentalChild.User = gobjUser
      mDentalChild.ModuleOpenView = False

      If mcolDentalChild Is Nothing Then
         Set mcolDentalChild = New Collection
      End If

      For Each childServer In mcolDentalChild

         If childServer.ServerIndex = iChildServer Then
            bFound = True
            Exit For

         Else
            bFound = False
         End If

      Next

      If Not bFound Then
         oldStatusText = cwobjCWComm.GetStatusText
         DoEvents
         cwobjCWComm.SetStatusText "Initializing Server: " & Sect(iChildServer)
         
        '''''''
        If Dental_Log Then
          Open "c:\temp\dental.log" For Append As #1
          Print #1, ""
          Print #1, CStr(Now) & ",LoadChildren(),Initializing Server: " & Sect(iChildServer)
          Close #1
        End If
        '''''''

         If mDentalChild.StartModule(Me, i, mobjDentalEncounter) Then
            '''''''
            If Dental_Log Then
              Open "c:\temp\dental.log" For Append As #1
              Print #1, ""
              Print #1, CStr(Now) & ",LoadChildren(),mDentalChild.StartModule(Me, i) OK."
              Close #1
            End If
            '''''''
            mDentalChild.ModuleLoaded = True
            mcolDentalChild.Add mDentalChild, CStr(mDentalChild.ServerIndex)
         Else '<bjr100501 SCR17713
            mDentalChild.ModuleLoaded = False
         End If
         cwobjCWComm.SetStatusText oldStatusText
      Else
         Set mDentalChild = Nothing
      End If

NextModule:
      
   Next i

   DoEvents
   'Screen.MousePointer = vbDefault
#If DEBUGON Then
   If Not DL Is Nothing Then DL.DebugLog Exe, Module, "LoadChildren", "End", DR
#End If
   
   Exit Sub
   
ErrHandler:
   '-- -show the VB error---
   gobjShared.ShowVBError Err.Number, Err.Description, "Dental.LoadChildren", errAppTitle, vbCritical
   Resume NextModule
End Sub
Public Function OpenDentalChild(intServer As Integer) As Boolean

   Dim rhMainForm     As Long
   Static playArea    As Long
   Static oldParentHandle As Long
   Dim ret            As Boolean
   Dim i              As Integer
   Dim childServer    As DentalChild
   Dim oldStatusText  As String
   Dim childWidth     As Long
   Dim childHeigth    As Long
   
#If DEBUGON Then
   Dim DR  As DebugRec

'   If Not DL Is Nothing Then DL.DebugLog Exe, Module, "OpenDentalChild " & intServer, "Begin", DR
#End If
   
   OpenDentalChild = True
   
   i = 0
   playArea = frmDental.ocxChildPlayArea.hwnd
   'Screen.MousePointer = vbHourglass
'   BigTooth = 0
'   FullMouth = 1
'   Notes = 2
'   Labs = 3
'   DataViews = 4


   Set childServer = mcolDentalChild(CStr(intServer))

   If Not childServer.ModuleOpenView Then
      oldStatusText = cwobjCWComm.GetStatusText
      oldParentHandle = childServer.parentHandle
      
      DoEvents
      cwobjCWComm.SetStatusText "Opening View of Server: " & childServer.ServerName
      
      '''''''
      If Dental_Log Then
        Open "c:\temp\dental.log" For Append As #1
        Print #1, ""
        Print #1, CStr(Now) & ",OpenDentalChild(),Opening View of Server: " & childServer.ServerName
        Close #1
      End If
      '''''''
    
      childServer.parentHandle = playArea
      ret = childServer.Module.OpenView(childServer.ViewType, childServer.parentHandle, rhMainForm)
      If Not ret Then
        DoEvents
        OpenDentalChild = False
        childServer.parentHandle = oldParentHandle
               
        cwobjCWComm.SetStatusText oldStatusText
        Screen.MousePointer = vbDefault
        frmDental.setCurrentTab (False)
        frmDental.ocxChildrenTab.CurrTab = giCurrentServer
        Set childServer = Nothing   'hwcc 2/26/2002
        Exit Function
      End If
      
      childServer.ModuleOpenView = True
      cwobjCWComm.SetStatusText oldStatusText
   End If
   
   frmDental.ocxChildPlayArea.Refresh
   oldStatusText = cwobjCWComm.GetStatusText
   
   '''''''
    If Dental_Log Then
      Open "c:\temp\dental.log" For Append As #1
      Print #1, ""
      Print #1, CStr(Now) & ",OpenDentalChild(), Resizing View of Server: " & childServer.ServerName
      Close #1
    End If
    '''''''

   DoEvents
   cwobjCWComm.SetStatusText "Resizing View of Server: " & childServer.ServerName
   Call childServer.Module.ResizeView(childServer.ViewType, 0, 0, frmDental.ocxChildPlayArea.Width, frmDental.ocxChildPlayArea.Height)
   cwobjCWComm.SetStatusText oldStatusText
   'Screen.MousePointer = vbDefault
   
   Set childServer = Nothing  'hwcc 2/26/2002
   
#If DEBUGON Then
   If Not DL Is Nothing Then DL.DebugLog Exe, Module, "OpenDentalChild " & intServer, "End", DR
   
#End If

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

    Static bPrinting As Boolean 'Determines if the user is currently printing.
    
   Dim vNotes           As String       'Used by SCR 24426 RWalker 12/11/02 'Variant
   Dim iDRC             As Integer
   Dim bResult          As Boolean
              
   Dim childServer      As DentalChild  'Used by SCR 24426 RWalker 12/11/02
   Dim childMenuItem    As childMenus
   Dim oEncDoc          As Encounter
   
   'Dim objDAS          As ICHCSII_DAS    'hwcc 2/26/2002 don't need this no more
   'Dim objRS            As Recordset     'hwcc 2/26/2002 don't need this no more
   'Dim sqlString        As String        'hwcc 2/26/2002 don't need this no more
   'Set objDAS = GetDasObj()    'hwcc 2/26/2002 don't need this no more
   
   
    On Error GoTo ErrHandler
   
    Screen.MousePointer = vbHourglass
               
   ' 7/08/2002
    'WR IParent.CurrentEncounter() now returns IEncDoc interface SCR 25225
    Set oEncDoc = gobjEnc.CurrentEncounter
   
   '-- --------------------------------------------------------------------------------
   '--  NOTE: You don't have to create menu items, these are optional, that is why
   '--        the code is commented out.
   '-- --------------------------------------------------------------------------------
'    '---perform the actions associated with the menu item that was selected---
    Select Case vlMenuId
        Case glMenuSave
                        
            ''''''''''''''''''''
            '<bjr 071801 [14411]
            ''''''''''''''''''''
            'problem: save action from notes issued save action to inactive
            'module (bigtooth), causing data integrity errors
            'fix: save action should only issue save action to current server
            
            '<bjr added
     
            Call SaveEncounter(False, False)
            
            'SCR 24890 RWalker 12/11/02 - MEP (multiple entry point) solution (Xref SCR 29085)
            Set childServer = mcolDentalChild(CStr(giCurrentServer))
            vNotes = childServer.ServerName
            
            goARM.StopCmplxTx NCID_ENCOUNTER_APPLICATION & gsARMTx_Is_Dental_Caption_Save & CStr(glMenuSave), _
                              enmARMTxStatusOk, enmARMStop_MeOnly, _
                              vNotes

        Case glMenuNext
     
            'SCR 24426 RWalker 12/11/02 - MEP (multiple entry point) solution (Xref SCR 29085)
            Set childServer = mcolDentalChild(CStr(giCurrentServer))
            vNotes = childServer.ServerName
            
            If Not gobjEnc Is Nothing Then
               Call oEncDoc.OpenSection(eDisp)
            End If
            
            goARM.StopCmplxTx NCID_ENCOUNTER_APPLICATION & gsARMTx_Is_Dental_Caption_Disposition & CStr(glMenuNext), _
                              enmARMTxStatusOk, enmARMStop_MeOnly, _
                              vNotes

        Case glMenuPrintID
            
            'If the user is already printing, then don't try to do it again.
            If bPrinting = True Then
                Exit Sub
            End If
            
            bPrinting = True
            
            'KDunne 8/1/2002
            'SCR 26236 - When printing, first we save, then we update the Encounter
            '            RTF Note.
            Call SaveEncounter(False, False)
            SetEncDentalNotes msDentalEncounterID
            'End 26236
            
            If giCurrentServer = BigTooth Or giCurrentServer = DataViews Then
                Set mDentalChild = mcolDentalChild(CStr(giCurrentServer))
                mDentalChild.Module.PrintSF603A True

            Else
                oEncDoc.PrintEncounter True
            End If
            
            'Finished printing
            bPrinting = False
            
        Case Else   '--send to children--
         Dim i       As Integer
         Dim bFound  As Boolean

         For Each childServer In mcolDentalChild

            For Each childMenuItem In childServer.childMenuItems

               If childMenuItem.menuItem = vlMenuId Then
                  childServer.Module.MenuItemSelected (vlMenuId)
                  bFound = True
                  Exit For

               End If

            Next

         Next

         If Not bFound Then

            For i = 1 To mcolDentalChild.Count
               mcolDentalChild(i).Module.MenuItemSelected (vlMenuId)
            Next

         End If
    End Select
   
FuncExit:
    Screen.MousePointer = vbDefault
    
   Exit Sub
   
ErrHandler:
   '-- -show the VB error---
   Call gobjShared.ShowVBError(Err.Number, Err.Description, "Dental.MenuItemSelected", errAppTitle, vbExclamation)
   
   bPrinting = False
   
   GoTo FuncExit

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

   '-- -update the all running "views" with the new patient---

   If cwbMainViewRunning Then
      '-- ****************************************************************************
      '--  TO DO:
      '--    Add code here to update the data in all "running" views with this new
      '--    patient.
      '-- ****************************************************************************
      
      
      
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
Public Sub ResizeView(ByVal viViewType As Integer, _
                       ByVal vlLeft As Long, _
                       ByVal vlTop As Long, _
                       ByVal vlWidth As Long, _
                       ByVal vlHeight As Long)

   On Error GoTo ErrHandler
   
   '-- -reset global "client area" parameters---
   cwlLeft = vlLeft
   cwlTop = vlTop
   cwlWidth = vlWidth
   cwlHeight = vlHeight
   
   frmDental.ocxChildPlayArea.Height = cwlHeight
   frmDental.ocxChildPlayArea.Width = cwlWidth
   
   '-- -perform different actions based upon the "view type"---

   If viViewType = cwiREGULAR_VIEW Then
            
      '--- SCR-ARM-42428 RWalker 11/05/2003
      gbFromCore = True

      If Not cwbFirstRegularResize Then
         '-- -------------------------------------------------------------------------------
         '--  Reset the "WindowState" to "Normal", so the form can be resized.  This is
         '--    needed because a window CANNOT be moved or resized if it is "minimized" or
         '--    "maximized".
         '--
         '--  NOTE: This generates a form "resize" event.
         '-- -------------------------------------------------------------------------------
         frmDental.WindowState = vbNormal
         
         '-- -hide the form to so that annyoing "screen blinking" will not be seen while controls are moved and resized---
         frmDental.Visible = False
         
         '-- -resize form to fill entire "client area", NOTE: This generates a "resize" event---
         frmDental.Move cwlLeft, cwlTop, cwlWidth, cwlHeight
         
         '*** Added by Scott 2/1/01
         'Call the resize event of the children
         '*** Modified by Quan 4/27/01
         Dim i As Integer
         i = 0
                 
         If cwlWidth < 400 Then
            cwlWidth = 400
         End If
         
         If cwlHeight < 150 Then
            cwlHeight = 150
         End If

'SCR 33645
'         For Each mDentalChild In mcolDentalChild
'            If mDentalChild.ServerIndex <> giCurrentServer Then
'                mDentalChild.Module.ResizeView viViewType, cwlLeft, cwlTop, cwlWidth - 400, cwlHeight - 150
'            End If
'         Next

'This is completely wrong
'         For Each mDentalChild In mcolDentalChild
'            If i <> giCurrentServer Then
'                mDentalChild.Module.ResizeView viViewType, cwlLeft, cwlTop, cwlWidth - 400, cwlHeight - 150
'            End If
'            i = i + 1
'         Next
         Set mDentalChild = mcolDentalChild(CStr(giCurrentServer))
         
        '43306
        If cwlWidth - 300 < 0 Then
           cwlWidth = 300
        End If
        If cwlHeight - 150 < 0 Then
            cwlHeight = 150
        End If
        mDentalChild.Module.ResizeView viViewType, cwlLeft, cwlTop, cwlWidth - 300, cwlHeight - 150
         '*** End of addition
         
         frmDental.Visible = True
      Else                             '--first "regular" resize, need to perform "initailization" Dentals--
         '-- -reset flag so this code will not be executed again---
         cwbFirstRegularResize = False
         
         '-- --------------------------------------------------------------------------
         '--  NOTE: Here is where you "Show" the "Regular View" form for the FIRST time.
         '-- --------------------------------------------------------------------------
         If Not gobjEnc Is Nothing Then
            'SCR 24426/24890 RWalker 10/28/2002 - Replaced Caption parm's literal with constant so
            'Core can start this ARM transaction when the user selects the action on the CW Main Form
            glMenuNext = gobjEnc.AddMenuItem(mnAction, gsARMTx_Is_Dental_Caption_Disposition, "Next.bmp", "Go to the Disposition screen")
            glMenuSave = gobjEnc.AddMenuItem(mnAction, gsARMTx_Is_Dental_Caption_Save, "save.bmp", "Save Dental")
            '--- SCR-ARM-24426 RWalker 12/9/02 - MEP (multiple entry point) solution (Xref SCR 29085)
            goARM.InitCmplxTxCls gsARMTx_OpenDisposition, _
                                 NCID_ENCOUNTER_APPLICATION & gsARMTx_Is_Dental_Caption_Disposition & CStr(glMenuNext), _
                                 gsARM_ACT_FROM
            '--- SCR-ARM-24890 RWalker 12/9/02 - MEP (multiple entry point) solution (Xref SCR 29085)
            goARM.InitCmplxTxCls gsARMTx_SaveDentalData, _
                                 NCID_ENCOUNTER_APPLICATION & gsARMTx_Is_Dental_Caption_Save & CStr(glMenuSave), _
                                 gsARM_ACT_FROM
                                 
            '--- SCR-ARM-42428 RWalker 11/05/2003 - Initialize ARM control
            frmDental.ocxChildrenTab.CurrTab = frmDental.ocxChildrenTab.FirstTab

            frmDental.Display
            
            'Set the default tab
            frmDental.ChangeTab BigTooth

         Else
            CloseView cwiREGULAR_VIEW
            Exit Sub
         End If
      End If
      
      '--- SCR-ARM-42428 RWalker 11/05/2003
      gbFromCore = False


   End If
   
   Exit Sub
   
ErrHandler:
   '-- -show the VB error---
   Call gobjShared.ShowVBError(Err.Number, Err.Description, "Dental.ResizeView", errAppTitle, vbExclamation)
   
   Exit Sub

End Sub


Private Sub Class_Initialize()

#If DEBUGON Then
   Set DL = GetObject("", "dlog.cdebuglog")
   DL.debugStart
   Dim DR  As DebugRec

   DL.DebugLog Exe, Module, "Class_Initialize", "Begin", DR
#End If
   '-- -create shared object used to show VB errors, etc.---
   Set gobjShared = New CWShared
   
   Set mobjDentalPatient = New DentalPatient
   
   '-- -set the help file for this application---
   App.HelpFile = gobjShared.CWHelpFile
   
   '--- SCR 28289 RWalker 10/10/02 Instantiate ARM / Init transactions
   Call ARM_InitAllTransactionsForDental_Parent
      
   '-- ----------------------------------
   '--  This statement is needed to debug
   '--    this OLE Server.
   '-- ----------------------------------
#If ciDEBUG = 1 Then

   If cwbDebugLogon = False Then
      Exit Sub

   End If

#End If
   
#If DEBUGON Then
   DL.DebugLog Exe, Module, "Class_Initialize", "End", DR
#End If

End Sub
Public Sub initialize_objects()

#If DEBUGON Then
    Dim DR  As DebugRec

    If Not DL Is Nothing Then DL.DebugLog Exe, Module, "initialize_objects", "Begin", DR
#End If
    Dim svgCommFactory As srvSVGComm.SingletonFactory    'hwcc 2/26/2002
    Set svgCommFactory = New srvSVGComm.SingletonFactory  'hwcc 2/26/2002
    
    
    
    If Not (svgCommFactory Is Nothing) Then Set svgCommObj = svgCommFactory.Singleton
    If Not (svgCommObj Is Nothing) Then Set NotesObj = svgCommObj.DataObject
    
    Set svgCommFactory = Nothing    'hwcc 2/26/2002
   
    Set mobjDentalPatient = New DentalPatient
    
    'KDunne 5/17/2002
    'SCR 22441 - Make sure global Current Server varibable is intialized to BigTooth
     giCurrentServer = BigTooth
    
#If DEBUGON Then
   If Not DL Is Nothing Then DL.DebugLog Exe, Module, "initialize_objects", "End", DR
   
#End If

End Sub



Private Sub terminate_Objects()

    glMenuPrintID = 0

    Set svgCommObj = Nothing
    Set NotesObj = Nothing
    Set mobjDentalPatient = Nothing
    Set mobjSBE = Nothing
    Set mobjDentalEncounter = Nothing
    
    If Not mobjCariesProtocol Is Nothing Then
        mobjCariesProtocol.Destroy_Objects
        Set mobjCariesProtocol = Nothing
    End If
    
    DestroyDasObj
    
    'KDunne 5/17/2002
    'SCR 22441 - Make sure we set the CurrentServer to BigTooth when we are done.
    giCurrentServer = BigTooth
    
    'KDunne 7/9/2002
    'SCR 20759
    mDRCLstExamDte = Empty
    mDRCLstChgDte = Empty
    'end 20759
    
    'KDunne 7/30/2002
    'SCR 26121
    mblnSuppressDRCDisp = False
    mblnDentalEncLoaded = False
    
    'Set svgCommFactory = Nothing   'hwcc 2/26/2002 we release svgCommFactory reference in initialize_objects
End Sub

Private Sub Class_Terminate()

#If DEBUGON Then
   Dim DR  As DebugRec
   Dim objNotesSF As CHCSII_Dental_NotesPlus.clsSingleton
   
   DL.DebugLog Exe, Module, "Class_Terminate", "Begin", DR
#End If
   '-- -release reference to objects---
   '-- Set gobjAnPCollection = Nothing
   terminate_Objects
   Set objNotesSF = New CHCSII_Dental_NotesPlus.clsSingleton
   
   Set mDentalChild = Nothing    'hwcc 2/26/2002
   Set mobjCariesProtocol = Nothing 'hwcc 2/26/2002
   Set mcolDentalChild = Nothing
   Set gobjShared = Nothing
   Set cwobjCWComm = Nothing
   Set cwobjLogonEnviron = Nothing
   Set cwobjConfig = Nothing
   Set cwobjPatient = Nothing
   Set gobjUser = Nothing

   objNotesSF.Destroy_Objects
   Set objNotesSF = Nothing
   
  
   '-- ********************************************************************
   '--  TO DO:
   '--    Add code here to release reference to all other global objects.
   '-- ********************************************************************
   
   
   
   
   '-- --------------------------------------
   '--  If debugging, make sure to "log off".
   '-- --------------------------------------
#If ciDEBUG = 1 Then
   cwbDebugLogOff
#End If
#If DEBUGON Then
   DL.DebugLog Exe, Module, "Class_Terminate", "End", DR
   Set DL = Nothing
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
Public Sub Message(ByVal viMessageType As Integer, _
                    ByVal vsMessageText As String, _
                    ByVal vsSourceServerNCID As String)
    Dim strOKToChange As String
    
   On Error GoTo ErrHandler
   
   '-- -----------------------------------------------------------------------------
   '--  Take corresponding actions based upon the message "type".
   '--
   '--  NOTE: This method is required, but currently there are NO REQUIRED messages
   '--        to process.
   '-- -----------------------------------------------------------------------------
   Select Case viMessageType
        
    Case cwiGENERIC
         MessageGeneric vsMessageText
    Case cwiOK_TO_CHANGE_PATIENT
        If SaveEncounter(True, False) Then
            strOKToChange = "Y"
        Else
            strOKToChange = "N"
        End If
        
        cwobjCWComm.Message cwiOK_TO_CHANGE_PATIENT, strOKToChange, NCID_CW_APPLICATION, NCID_DDSA_APPLICATION, False
    
    Case cwiREFRESH_DATA
        frmDental.Refresh
        
    Case cwiFOREGROUND_APPLICATION
        'KDunne 1/13/2002
        'SCR 31404
        'Send Message to current Dental child only
        If Not mcolDentalChild Is Nothing Then
            mcolDentalChild(CStr(giCurrentServer)).Module.Message viMessageType, _
                vsMessageText, vsSourceServerNCID
        End If
    Case Else '--do nothing-- and tell children
      If Not mcolDentalChild Is Nothing Then
         For Each mDentalChild In mcolDentalChild
            mDentalChild.Module.Message viMessageType, vsMessageText, vsSourceServerNCID
         Next
      End If
   End Select
   
   Exit Sub
   
ErrHandler:
   '-- -show the VB error---
   Call gobjShared.ShowVBError(Err.Number, Err.Description, "Dental.Message", errAppTitle, vbExclamation)
   
   Exit Sub
    Resume
End Sub
Private Sub MessageGeneric(vsMessageText As String)

   Dim childServer  As DentalChild
   Dim i            As Integer
   Dim buff()       As String

'''   If giCurrentServer <> frmDental.getCurrentTab Then
'''        Exit Sub
'''   End If
   
   ''mbKeepMeAlive= True
   buff = Split(vsMessageText, "|")
   Select Case UCase$(buff(0))
      Case "LAUNCH"
         '-- open child in buff(1), send rest of string to them
         If Not cwbMainViewRunning Then _
            Exit Sub  'need the frmAnp

         If buff(1) = "" Then _
            Exit Sub

         For i = LBound(glServerInfo) To UBound(glServerInfo)

            If glServerInfo(i) = buff(1) Then
                '''''''
                If Dental_Log Then
                    Open "c:\temp\dental.log" For Append As #1
                    Print #1, CStr(Now) & ",MessageGeneric(), Server - " & buff(1) & ", Tab - " & CStr(i)
                    ' & ",Notes from Notes Ops-" & vNotes
                    Close #1
                End If
                '''''''
                '12/20/02 RWalker Set in MessageGeneric procedure
                gbFromCore = True

                ChangeTab (i)
                
                'SCR-ARM-30553 RWalker 12/20/02 Set in MessageGeneric procedure
                gbFromCore = False
                
                Exit For

            End If

         Next i

         For Each childServer In mcolDentalChild

            If childServer.AppNCID = buff(1) Then

               If UBound(buff) > 1 Then
                  '-- ///// do a jointSome
                  childServer.Module.Message cwiGENERIC, "LAUNCH|" & buff(2), NCID_DDSA_APPLICATION
                  Exit For

               End If

            End If                     '- oChild.AppNCID = Buff(1) Then

         Next childServer

      Case "DELETE_SAVE"
            cwobjCWComm.DeleteMenuItem (glMenuSave)
            
      Case "ADD_SAVE"
            '--- SCR-ARM-30766 RWalker 12/12/02 - MEP (multiple entry point) solution (Xref SCR 29085)
            '1) Replaced hard-coded "&Save" with ARMSharedTxNames constant to maintain consistency with
            'the ARM transaction's nickname.  The constant id'd in the OpenView method was "Save", thus
            'preventing the ARM Stop from finding the related Start after this line of code executed. I
            'changed the constant literal to contain the ampersand.
            '2) The menu id is the last part of the ARM transaction's nickname which Core uses to start
            'the ARM transaction. So the global variable needs to change when the menu id changes. If it
            'is different then we need to define a new ARM transaction.
            Dim lNew As Long
            lNew = gobjEnc.AddMenuItem(0, gsARMTx_Is_Dental_Caption_Save, "save.bmp", "Save Dental")
            If lNew <> glMenuSave Then
                glMenuSave = lNew
                goARM.InitCmplxTxCls gsARMTx_SaveDentalData, _
                                     NCID_ENCOUNTER_APPLICATION & gsARMTx_Is_Dental_Caption_Save & CStr(glMenuSave), _
                                     gsARM_ACT_FROM
            End If
      
      Case "SVGDRAWN"
            mbSVGDrawn = True
            EnablePrintIcon
            
      Case Else
         vsMessageText = "Unknown message type"
   End Select
   'mbKeepMeAlive= False
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
      
      
      
      '---Instantiate Help-file object
    App.HelpFile = gobjShared.CWHelpFile
    '-- -save the "ProgId" assigned to me by the CW Core---
    cwlMyProgId = vlMyProgId
    
    '-- -create references to passed in objects---
    Set cwobjCWComm = robjComm
    Set cwobjLogonEnviron = robjLogonEnvironment
    Set cwobjConfig = robjConfig
    Set cwobjPatient = robjPatient
      
    ReDim glServerInfo(6)     'hwcc 2/26/2002 change from 8 to 6
    glServerInfo(BigTooth) = NCID_DENTAL_BIGTOOTH_APPLICATION
    '   glServerInfo(FullMouth) = NCID_DENTAL_FULLMOUTH_APPLICATION
    'hwcc 2/1/01 - we are not full fetch NotePlus
    glServerInfo(Notes) = NCID_DENTAL_NOTESPLUS_APPLICATION
    glServerInfo(Labs) = NCID_DENTAL_LABS_APPLICATION
    glServerInfo(DataViews) = NCID_DENTAL_DATAVIEWS_APPLICATION
    glServerInfo(PerioChart) = NCID_DENTAL_PERIOCHART_APPLICATION
    glServerInfo(Forensics) = NCID_DENTAL_FORENSIC_APPLICATION
    
    ReDim Sect(6)     'hwcc 2/26/2002 change from 8 to 6
    Sect(BigTooth) = "Big Tooth"
    '   Sect(FullMouth) = "Full Mouth"
    Sect(Notes) = "Dental Notes"
    Sect(Labs) = "Dental Labs"
    Sect(DataViews) = "Dental Data Views"
    Sect(PerioChart) = "Dental Perio Chartings"
    Sect(Forensics) = "Dental Forensics"
    '-- -set global flags---
    cwbMainViewRunning = False
   
    If gobjShared.CmdLineSwitches("DENTLOG") = "DENTLOG" Then
        mblnDebugLogging = True
        Log_Event "Dental Parent Initialize", "Dental Logging Enabled"
    End If

    '-- ****************************************************************************
    '--  TO DO:
    '--    Add code here to create objects, initialize global variables, etc.
    '-- ****************************************************************************
       
    Exit Sub
   
ErrHandler:
    '-- -show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "OrderSet.Initialize", errAppTitle, vbExclamation)
    
    Exit Sub
    
End Sub

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
Public Sub ICHCSII_CoreAPI2_Initialize2(ByVal vlMyProgId As Long, _
                                        robjComm As Object, _
                                        robjLogonEnvironment As Object, _
                                        robjConfig As Object, _
                                        robjPatient As Object, _
                                        robjUser As Object)

   On Error GoTo ErrHandler
   
   Set gobjUser = robjUser
   
   Dim objRoles      As ICHCSII_Roles
   Dim objResources  As ICHCSII_Resources
   
    If gobjUser.HasPrivilegeEx(Priv_Dental_Access, Priv_Write) = True Then
        'Only init if the user has dental privileges
        Call Initialize(ByVal vlMyProgId, robjComm, robjLogonEnvironment, robjConfig, robjPatient)
    End If
   
   Exit Sub
   
ErrHandler:
   Call gobjShared.ShowVBError(Err.Number, Err.Description, "Dental.Initialize", errAppTitle, vbExclamation)
   
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
   
   '-- *****************************************************************
   '--  TO DO:
   '--    Add code here to invoke this OLE Server's properties page.
   '-- *****************************************************************

   Exit Sub
   
ErrHandler:
   '-- -show the VB error---
   Call gobjShared.ShowVBError(Err.Number, Err.Description, "Dental.Properties", errAppTitle, vbExclamation)
   
   Exit Sub

End Sub

Private Function ICHCSII_CoreAPI3_CloseView2(ByVal viViewType As Integer) As Boolean
  'save DRC here, ask for highest DRC from Big Tooth and notes
  Dim vNotes        As String 'Variant
  Dim bResult As Boolean
  Dim blnCloseEncounter As Boolean
  
  Debug.Print Screen.MousePointer
  
  'frmDental.setCursor (vbHourglass)
  'Screen.MousePointer = vbHourglass
  
  'DoEvents
    
  blnCloseEncounter = True
  
  If viViewType = cwiALL_VIEWS Then
      '-- -close the "main view"---

      If cwbMainViewRunning Then
        '-- -set global flags---
        cwbMainViewRunning = False
         
        '-- -this indicates that the CW Core is closing the "view", NOT this OLE Server---
        cwbCloseViewMyself = False
         
         '-- -close ALL forms that are associated with all "views"---
        If gobjEnc Is Nothing Then
            Unload frmDental
        Else
                        
            If Not SaveEncounter(True, blnCloseEncounter) Then GoTo FuncExit
                      
        End If
      End If

   ElseIf viViewType = cwiREGULAR_VIEW Then
      '-- -close the "main view"---

      If cwbMainViewRunning Then
         '-- -set global flags---
         cwbMainViewRunning = False
         
         '-- -this indicates that the CW Core is closing the "view", NOT this OLE Server---
         cwbCloseViewMyself = False
         
         '-- -close ALL forms that are associated with all "views"---
        If gobjEnc Is Nothing Then
            Unload frmDental
        Else
        
            If Not SaveEncounter(True, blnCloseEncounter) Then GoTo FuncExit
                                    
        End If
         
      End If

   End If
   
FuncExit:
   cwobjCWComm.DeleteMenuItem (glMenuPrintID)
   mblnDentalEncLoaded = False
   ICHCSII_CoreAPI3_CloseView2 = True
   terminate_Objects
   'frmDental.setCursor (vbDefault)
   Screen.MousePointer = vbDefault
End Function

Private Function SaveDentalEncounter(ByVal iDRC As Integer, _
                                     ByVal msDentalEncounterID As String, _
                                     Optional ByVal strLastExamDate As String = "", _
                                     Optional ByVal strLastChangeDate As String = "") _
                                     As Boolean
    
    ' save DRC to Dental Encounter Table
    Dim strSQL As String
    Dim objDAS As ICHCSII_DAS
       
    Dim sDRC_LastChangeDate As String
    Dim sDRC_LastExamDate As String
    '

    On Error GoTo errHandle
    
    SaveDentalEncounter = False
    If iDRC > 0 And iDRC < 5 And msDentalEncounterID > 0 Then
        'KDunne scr 27059 - Combined all steps into one SQL statement.
        Set objDAS = GetDasObj()
        strSQL = "UPDATE DENTAL_ENCOUNTER SET DRC=" & iDRC '''& " WHERE DENTAL_ENCOUNTER_ID=" & msDentalEncounterID
        
        'ASamedov - 02/21/02
        If strLastExamDate <> "" And strLastExamDate <> cDefaultDate And strLastExamDate <> cNull Then
            sDRC_LastExamDate = goCHCSIIConn.FormatDateForDBInsert(strLastExamDate)
            'sDRC_LastExamDate = "TO_DATE('" & Format(strLastExamDate, gcDATE_FORMAT_VB_TIME) & "'," & gcDATE_FORMAT_ORA_TIME & ")"
            strSQL = strSQL & ", DRC_LastExamDate = " & sDRC_LastExamDate
            mDRCLstExamDte = strLastExamDate
        End If
        
        If strLastChangeDate <> "" And strLastChangeDate <> cDefaultDate And strLastChangeDate <> cNull Then
            sDRC_LastChangeDate = goCHCSIIConn.FormatDateForDBInsert(strLastChangeDate)
            'sDRC_LastChangeDate = "TO_DATE('" & Format(strLastChangeDate, gcDATE_FORMAT_VB_TIME) & "'," & gcDATE_FORMAT_ORA_TIME & ")"
            strSQL = strSQL & ",  DRC_LastChangeDate = " & sDRC_LastChangeDate
            mDRCLstChgDte = strLastChangeDate
        End If

        strSQL = strSQL & " WHERE DENTAL_ENCOUNTER_ID=" & msDentalEncounterID
        objDAS.ExecuteSQL (strSQL)
        SaveDentalEncounter = True
    End If
    
    If Not objDAS Is Nothing Then
        Set objDAS = Nothing
    End If
    
    If SaveDentalEncounter Then mEncLevelDRC = iDRC
    '''''''
    If Dental_Log Then
        Open "c:\temp\dental.log" For Append As #1
        Print #1, CStr(Now) & ",SaveDentalEncounter(),Patient-" & sPtnEncID & ",Dental Encounter-" & msDentalEncounterID & ",Input DRC-" & CStr(iDRC)
        Close #1
    End If
    '''''''
    Exit Function
  
errHandle:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "Dental.SaveDentalEncounter", errAppTitle, vbCritical)
    If Not objDAS Is Nothing Then
        Set objDAS = Nothing
    End If
End Function
Private Function SaveDentalPatient(ByVal iDRC As Integer, _
                                   ByVal PatientID As String, _
                                   Optional ByVal strLastExamDate As String = "", _
                                   Optional ByVal strLastChangeDate As String = "" _
                                   ) As Boolean
    ' save SRC to MMI
    Dim iDRC1, iDRC2 As Integer
    
    'ASamedov - 02/11/02
    Dim sLastExamDate As String
    Dim sLastChangeDate As String
    Dim objEncounter As CHCSIIEncounterOps.Encounter
    Dim strLogMsg As String
    
    sLastExamDate = ""
    sLastChangeDate = ""
    '
    SaveDentalPatient = False
    
    
    'KDunne 10/22/2002
    'SCR 28687
    If mblnDebugLogging Then
        strLogMsg = " iDRC = " & iDRC & _
            ", PatientID = " & PatientID & ", Patient Level Last Exam Date = " & _
            mPatLevelLstExmDte & ", strLastExamDate = " & strLastExamDate & _
            ", strLastChangeDate = " & strLastChangeDate
        Log_Event "SaveDentalPatient - Begining", strLogMsg
    End If
    
    
    If iDRC > 0 And iDRC < 5 Then
       
        iDRC1 = mEncLevelDRC
               
        Call mUpdateGenericMMICDR(PatientID, NCID_DENTAL_READINESS_CLASSIFICATION, CStr(iDRC), "")
        mPatientLevelDRC = iDRC
        
        ' for Readiness
        iDRC2 = iDRC
        Set objEncounter = gobjEnc.CurrentEncounter
        
        Call mUpdateGenericMMICDR(PatientID, NCID_LAST_DENTAL_READINESS_CLASSIFICATION_CHANGE_DATE, mstrApptDTS, "")
        mPatLevelLstChgDte = mstrApptDTS
        
        'ASamedov - 02/11/02
        If strLastExamDate <> "" And strLastExamDate <> cNull Then
            Call mUpdateGenericMMICDR(PatientID, NCID_LAST_DENTAL_EXAM_DATE, strLastExamDate, "")
            mPatLevelLstExmDte = strLastExamDate
        End If
        SaveDentalPatient = True
        
    ElseIf iDRC = 0 Then
        Call mUpdateGenericMMICDR(PatientID, NCID_DENTAL_READINESS_CLASSIFICATION, "4", "")
        ' for Readiness
        iDRC2 = 4
        
        Call mUpdateGenericMMICDR(PatientID, NCID_LAST_DENTAL_READINESS_CLASSIFICATION_CHANGE_DATE, cDefaultDate, "")
        Call mUpdateGenericMMICDR(PatientID, NCID_LAST_DENTAL_EXAM_DATE, cDefaultDate, "")
        SaveDentalPatient = True
        
        mPatientLevelDRC = 4
        mPatLevelLstChgDte = cDefaultDate
        mPatLevelLstExmDte = cDefaultDate
    End If
    
    'KDunne 10/22/2002
    'SCR 28687
    If mblnDebugLogging Then
        strLogMsg = " iDRC2 = " & iDRC2 & _
            ", PatientID = " & PatientID & ", Patient Level Last Exam Date = " & _
            mPatLevelLstExmDte & ", Patient Level Last Change Date = " & mPatLevelLstChgDte
        Log_Event "SaveDentalPatient - After Updating MMI", strLogMsg
    End If
    
     ' update DRC in Readiness module, see SCR 22051
    If iDRC2 > 0 Then
        Call SaveDentalReadiness(iDRC2, PatientID, strLastExamDate, strLastChangeDate)
    End If
    
End Function

Private Function RetrieveDentalPatient(ByRef iDRC As Integer, ByVal PatientID As String) As Boolean
    'retrieve SRC to MMI
    iDRC = 0
    'hwcc 2/7/02 - mRetrieveGenericMMICDR now return the value string
    Dim RtnStr As String
    'RtnStr = mRetrieveGenericMMICDR(PatientID, NCID_DENTAL_READINESS_CLASSIFICATION)
    RtnStr = ReadDentalReadiness(PatientID, NCID_DENTAL_READINESS_CLASSIFICATION)
    
    If RtnStr <> "" Then iDRC = CInt(RtnStr)   'SCR 21472
         
    '''''''
    If Dental_Log Then
        Open "c:\temp\dental.log" For Append As #1
        Print #1, CStr(Now) & ",RetrieveDentalPatient(),Patient-" & sPtnEncID & ",Dental Encounter-" & msDentalEncounterID & ",output DRC-" & CStr(iDRC)
        Close #1
    End If
    '''''''
    
    RetrieveDentalPatient = True
End Function

Private Function DisplayDentalPatient(ByVal iDRC As Integer, ByVal PatientID As String) As Boolean

    Dim strDRC As String
    ' call name bar interface
    ' send message to cw to display DRC
    DisplayDentalPatient = False
    
    'KDunne 7/30/2002
    'SCR 26121 - If we're signing the app, then we're out of the dental encounter,
    '            don't display the DRC Info.
    If mblnSuppressDRCDisp = True Then
        DisplayDentalPatient = False
        Exit Function
    End If
    
    If Not cwobjCWComm Is Nothing Then
        If iDRC = 0 Then
            strDRC = ""
        Else
            strDRC = CStr(iDRC)
        End If
        cwobjCWComm.Message cwiGENERIC, "DRC|" & strDRC, NCID_CW_APPLICATION, NCID_DDSA_APPLICATION
        DisplayDentalPatient = True
    End If
  
End Function

Private Function getCurrentDate() As Date
 
'    Dim strSQL As String
'    Dim objRS As Recordset
'    Dim objDAS As ICHCSII_DAS
    
    On Error GoTo errHandle
     
    getCurrentDate = Now
    
'    Set objDAS = GetDasObj()
'    strSQL = "SELECT TO_CHAR(sysdate, 'MM/DD/YYYY hh24:mi:ss') FROM DUAL"
'
'    Set objRS = objDAS.OpenRecordset(strSQL)
'    If Not objRS.EOF Then
'         getCurrentDate = CDate(objRS(0))
'    End If
    getCurrentDate = cwobjCWComm.GmtToLocal(cwobjCWComm.GetDateTime)

    
    'TEMP (for future date testing )
    'getCurrentDate = CDate(Format("04/04/2003 10:10:10 AM", gcDATE_FORMAT_VB_TIME))
    '
        
'    Set objDAS = Nothing  'hwcc 2/26/2002
'    Set objRS = Nothing     'hwcc 2/26/2002
        
'    If Not objDAS Is Nothing Then
'       Set objDAS = Nothing
'    End If
'    If Not objRS Is Nothing Then
'       Set objRS = Nothing
'    End If
    
    Exit Function
    
errHandle:
'    Set objDAS = Nothing  'hwcc 2/26/2002
'    Set objRS = Nothing     'hwcc 2/26/2002
    
'    If Not objDAS Is Nothing Then
'       Set objDAS = Nothing
'    End If
'    If Not objRS Is Nothing Then
'       Set objRS = Nothing
'    End If
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "Dental.getCurrentDate", errAppTitle, vbCritical)
  
End Function

Private Sub mUpdateGenericMMICDR(ByVal PatientID As String, _
   sNCID As String, _
   sValue As String, _
   Optional sValueNCID As Variant)
            
   On Error GoTo errHandle
   
    Dim mobjGEMS_MMIOps As GEMS_MMIObj.GEMS_MMIOps
    Dim mobjGEMS_MMI As GEMS_MMI
    Set mobjGEMS_MMIOps = New GEMS_MMIOps
    Set mobjGEMS_MMI = mobjGEMS_MMIOps.GetPatient(CLng(PatientID))
    Call mobjGEMS_MMIOps.UpdateGenericData(sNCID, sValue, CStr(sValueNCID))
    mobjGEMS_MMIOps.PutPatient
    Set mobjGEMS_MMIOps = Nothing   'hwcc 2/26/2002
    Set mobjGEMS_MMI = Nothing       'hwcc 2/26/2002

   Exit Sub
errHandle:

    Call gobjShared.ShowVBError(Err.Number, Err.Description, "Dental.mUpdateGenericMMICDR", errAppTitle, vbCritical)
   
End Sub

Private Function mRetrieveGenericMMICDR(ByVal PatientID As String, _
   ByVal sNCID As String) As String
'hwcc 2/7/02 - Change this from subroutine to function and return
' the Value Text in string

   On Error GoTo errHandle
     
    Dim mobjGEMS_MMIOps As GEMS_MMIObj.GEMS_MMIOps
    Dim mobjGEMS_MMI As GEMS_MMI
    Set mobjGEMS_MMIOps = New GEMS_MMIObj.GEMS_MMIOps
    Set mobjGEMS_MMI = mobjGEMS_MMIOps.GetPatient(CLng(PatientID))
    mRetrieveGenericMMICDR = mobjGEMS_MMIOps.GetGenericData(CStr(sNCID))
    Set mobjGEMS_MMIOps = Nothing  'hwcc 2/26/2002
    Set mobjGEMS_MMI = Nothing    'hwcc 2/26/2002
    Exit Function

errHandle:

   Call gobjShared.ShowVBError(Err.Number, _
      Err.Description, _
      "Dental.mRetrieveGenericMMICDR", _
      errAppTitle, _
      vbCritical)
   
End Function
Private Function ICHCSII_Dental_SetDentalParent(objDentalParent As Object) As Boolean

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

End Function

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



Private Function IChild_GetAutoCiteNarrative(TestType As String, NumPeriods As Integer, Period As AC_Period) As String

End Function

Private Function ICHild_SetParent(Parent As IParent) As Boolean
    Dim oEncDoc As Encounter

    If gobjUser.HasPrivilegeEx(Priv_Dental_Access, Priv_Write) = False Then
        Exit Function
    End If
   
   ICHild_SetParent = True

   Set gobjEnc = Parent
   
   If Not gobjEnc Is Nothing Then
        ' 7/08/2002
        'WR IParent.CurrentEncounter() now returns IEncDoc interface SCR 25225
        Set oEncDoc = gobjEnc.CurrentEncounter
       sPrimProvID = oEncDoc.Providers(1).NCID
    
   End If

    Set oEncDoc = Nothing
End Function
Public Function AddMenuItem(iMenu As MenuNames, _
                            sMenuText As String, _
                            sMenuBMPFileName As String, _
                            sTip As String, _
                            Optional progID As Long) _
                            As Long

'------------------------------------------------------
' Purpose: Adds a menu item to either Action or View
' Rreturns: The menu ID of the new menu item
' Parameters:
'  iMenu  - one of 0 (Action) or 1 (View)
'  sMenuText - the text for the menu item
'  sMenuBMPFileName - the filename of the bitmap icon file (empty string for 'view')
'  sTip - the tool tip text string
'--------------------------------------------------
   
   
'  If iMenu = mnAction Then
'    AddMenuItem = cwobjCWComm.AddActionmenuitem(sMenuText, sMenuBMPFileName, cwlMyProgId, sTip)
'  Else
'    AddMenuItem = cwobjCWComm.AddViewmenuitem(sMenuText, cwlMyProgId, sTip)
'  End If
   Dim childServer  As DentalChild
   Dim childMenu    As childMenus
   
   Set childMenu = New childMenus
   
   If iMenu = mnAction Then
      AddMenuItem = gobjEnc.AddMenuItem(0, sMenuText, sMenuBMPFileName, sTip)
   Else
      AddMenuItem = gobjEnc.AddMenuItem(1, sMenuText, "", sTip)
   End If

   If progID > 0 Then

      For Each childServer In mcolDentalChild

         If childServer.progID = progID Then
            childMenu.menuItem = AddMenuItem
            childServer.childMenuItems.Add childMenu, CStr(AddMenuItem)
            Exit For

         End If

      Next

   End If

   Set childServer = Nothing
   Set childMenu = Nothing
   Exit Function

ErrHandler:
   '-- -show the VB error---
   Call gobjShared.ShowVBError(Err.Number, Err.Description, "Dental.AddMenuItem", errAppTitle, vbExclamation)
   Exit Function
   
End Function
Public Sub ChangeTab(child As childTab)

   frmDental.ChangeTab child
   
End Sub

Public Sub UpdateDentalEncDRC()
    'Dim objDTNOps       As New CHCSII_Dental_Notes.clsDTNotesOps
    Dim iDRC            As Integer
  
    'ASamedov [SCR 8268, 8857, 14116] - 01/28/02
    Dim strSQL As String
    Dim objDAS As ICHCSII_DAS
    Dim objRS As Recordset
    Dim strPatientLastExamDate As String
    Dim bExamChecked As Boolean
    Dim sOralExam As String
    Dim MaxDxDRC As Integer
    Dim OverallDRC As Integer
    Dim MaxDRC As Integer
    Dim dteApptDate As Date
    Dim blnDirty As Boolean
    
    On Error GoTo ErrHandler
          
    'ASamedov [SCR 8268, 8857, 18441] - 01/28/02
    'DRC BizRules implementation
    '(to calculate and update DRC value
    'based on Paul's matrix)
           
    sOralExam = getOralExam
    If sOralExam = "1" Or sOralExam = "2" Then
        bExamChecked = True
        If mPatLevelLstExmDte = mDRCLstExamDte Then blnDirty = True
    Else
        bExamChecked = False
        If mPatLevelLstExmDte <> mDRCLstExamDte Then blnDirty = True
    End If
        
        
    MaxDxDRC = getMaxDxDRC
        
    If mPatientLevelDRC = 4 And Not bExamChecked Then
        iDRC = 4
    Else
        If MaxDxDRC > 0 Then
            'MaxDRC = 1, 2, or 3
'                If DateDiff("m", Format(mDRCLstExamDte, gcMONTH_FORMAT_VB), getCurrentDate) < 14 _
               Or mDRCLstExamDte = cDefaultDate Then
            If Not IsDRCExpired(mDRCLstExamDte) Or mDRCLstExamDte = cDefaultDate Then
                'less than 13 month
                'or it is first encounter with DX's DRC (MaxDRC > 0)
                iDRC = MaxDxDRC
                
            Else
                'greater than 13 month
                'if Initial/Periodic exam checked
                'then DRC is calculated based on MaxDxDRC
                If bExamChecked Then
                    iDRC = MaxDxDRC
                Else
                    iDRC = 4
                End If
            End If
        Else
            'MaxDxDRC = 0
'                If DateDiff("m", Format(mDRCLstExamDte, gcMONTH_FORMAT_VB), getCurrentDate) < 14 Then
            If Not IsDRCExpired(mDRCLstExamDte) Then
                    iDRC = 1
            ElseIf mDRCLstExamDte = cDefaultDate Then
                If bExamChecked Then
                    'periodic/initial exam checked first time
                    iDRC = 1
                Else
                    iDRC = 4
                End If
            ElseIf Not bExamChecked Then
                'greater than 13 month,
                'and LastExam date is not default date,
                'and exam not checked
                iDRC = 4
            ElseIf bExamChecked Then
                'greater than 13 month,
                'and LastExam date is not default date,
                'and exam checked
                iDRC = 1
            End If
        End If

    End If
    
    If iDRC <> mEncLevelDRC Then blnDirty = True
    
    'KDunne 7/10/2002
    'SCR 20759 - Only save data if drc has changed.
    If blnDirty Then
    
        If gobjEnc.StartSave Then
            'ASamedov - 02/07/02
            'Save DRC value in Dental_Encounter table
            If bExamChecked Then
                Call SaveDentalEncounter(iDRC, msDentalEncounterID, mstrApptDTS, mstrApptDTS)
            Else
                Call SaveDentalEncounter(iDRC, msDentalEncounterID, mPatLevelLstExmDte, mstrApptDTS)
            End If
            '
            gobjEnc.EndSave
        End If
    End If
    
    
FuncExit:
    'KDunne 7/16/2002
    'SCR 25736 - Always call display function.
    If blnDirty Then Call DisplayDentalPatient(iDRC, sPtnEncID)
    
    Exit Sub
    
ErrHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "Dental.UpdateDentalEncDRC", errAppTitle, vbExclamation)
    GoTo FuncExit

End Sub

Private Function M_bOpenDentalEncounter() As Boolean
'''************************************************************************
'''Called By:     ANY
'''
'''Arguments
'''
'''Description:   Attempt to find Dental encounter record for Current Encounter,
'''               Current Encounter Facility and Patient
'''               If found
'''                  Open this dental encounter
'''               Else
'''                  Create new dental encounter record and open it
'''
'''               Set Dental Encounter ID property to ID of the currently open Dental Encounter
'''Note:
'''
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:       True  - If successful
'''               False - Else
'''
'''Programmer:    Alexander Alberstein
'''Modified by:   ASamedov
'''Date:          02/13/2002
'''************************************************************************
   Dim objDAS           As ICHCSII_DAS
   Dim objRS            As Recordset
   Dim sqlString        As String
   Dim vDntID           As Variant
   Dim bReread          As Boolean
   Dim lCount           As Long
   Dim strLogMsg        As String
   
   Dim strSubSQL As String
   Dim iDRC As Integer
   Dim vMaxID As Long
   Dim iPSRLL As String
   Dim iPSRLC As String
   Dim iPSRLR As String
   Dim iPSRRL As String
   Dim iPSRRC As String
   Dim iPSRRR As String
   Dim strSQL As String
     
   'ASamedov
    Dim strApptDateTime As String
    Dim strLastChangeDate As String
    Dim strLastExamDate As String
    
    Dim sDentEncID As String
    Dim sLastExamDate As String
    Dim sLastChangeDate As String
    Dim sApptDateTime As String
    Dim oEncDoc     As Encounter
    Dim objSBE As SBE
    
    Dim lSBE As Integer
    Dim strSBEComments As String
    Dim oGroup As clsGroup
    Dim oTx As clsTx
    Dim bFoundTx As Boolean
    
    '

   On Error GoTo ErrHnd
    ' 7/08/2002
    'WR IParent.CurrentEncounter() now returns IEncDoc interface SCR 25225
    Set oEncDoc = gobjEnc.CurrentEncounter
  
   'KDunne 6/7/2002
   'Caries Protocol needs to have an open dental encounter when signing the app. During normal
   'loads, this would result in the dental encounter being loaded twice. this flag is used to
   'prevent that from occurring.
    If mblnDentalEncLoaded Then
        M_bOpenDentalEncounter = True
        Exit Function
    End If
   
   mbNewDentalEncounter = False
   
   msDentalEncounterID = ""
   Set objDAS = GetDasObj()
   Set objSBE = New SBE
   
   sEncID = oEncDoc.EncounterID
   sFclID = oEncDoc.FacilityNCID
   sPtnID = cwobjPatient.UnitNumber
   sPtnEncID = oEncDoc.PatientID
   sPrimProvID = oEncDoc.Providers(1).NCID
   sEncDTS = oEncDoc.StartDTS
   mstrProvideFullName = oEncDoc.PrimaryProvider.FullName 'Use for displaying workload before signing encounter
   
   mDRCLstChgDte = ""
   mDRCLstExamDte = ""
   sApptDateTime = ""
       
   'get Encouner appointment date/time
   'KDunne 7/9/2002
   'SCR 20759 - Appointment date is stored in the Encounter object, no need to hit
   'the db.
   
   mstrApptDTS = oEncDoc.ApptDTS
    If mblnUseStoredProc Then

        Set mobjDentalEncounter = New DentalEncounter
        
        M_bOpenDentalEncounter = mobjDentalEncounter.LoadDentalEncounter(cwobjPatient.UnitNumber, oEncDoc.EncounterID, _
                oEncDoc.FacilityNCID, oEncDoc.ApptDTS)
                
        With mobjDentalEncounter.DRC
            mEncLevelDRC = .EncLvlDRC
            mDRCLstExamDte = ConvertSPDate(.EncLvlLstExmDte)
            mDRCLstChgDte = ConvertSPDate(.EncLvlLstChgDte)
            mPatientLevelDRC = .PatLvlDRC

            mPatLevelLstChgDte = ConvertSPDate(.PatLvlLstChgDate)
            mPatLevelLstExmDte = ConvertSPDate(.PatLvlLstExmDate)
            
        End With
        
        msDentalEncounterID = mobjDentalEncounter.DentalEncounterId
        mlDentalEncounterID = mobjDentalEncounter.DentalEncounterId

    Else
        Dim objDentEncounter As DentalEncounter
        
        Set objDentEncounter = New DentalEncounter
        
        Call LoadReadinessDRCInfo(CLng(sPtnID))
    
        '''Attempt to find existing dental encounter record
        'sqlString = "Select to_char(LASTPERIODICEXAMDATE,'MM/DD/YYYY'),DENTAL_ENCOUNTER_ID, "
        sqlString = "Select LASTPERIODICEXAMDATE,DENTAL_ENCOUNTER_ID, " _
                  & " DRC, DRC_LastExamDate, DRC_LastChangeDate, SBE, SBE_COMMENTS, " _
                  & " PSR_LOWER_LEFT, PSR_LOWER_CENTER, PSR_LOWER_RIGHT, " _
                  & " PSR_UPPER_LEFT , PSR_UPPER_CENTER, PSR_UPPER_RIGHT, " _
                  & " ELIGIBILITYFLAG, ELIGIBILITYDATE " _
                  & " from DENTAL_ENCOUNTER Where " _
                  & "Encounter_Number = " & Val(sEncID) & " AND " _
                  & "Encounter_Facility_NCID = " & sFclID & " AND " _
                  & "Patient_Unit_ID = " & sPtnID
        Set objRS = objDAS.OpenRecordset(sqlString)
        '''If found
        If Not objRS.RecordCount = 0 Then
           ' found an existing encounter of patient
           'mLastDentalDate = objRS(0)
           With objRS
                vDntID = .Fields("DENTAL_ENCOUNTER_ID")
    '                mPatientLevelDRC = .Fields("DRC") & ""
                mEncLevelDRC = .Fields("DRC") & ""
                strLastExamDate = .Fields("DRC_LastExamDate") & ""
                strLastChangeDate = .Fields("DRC_LastChangeDate") & ""
                'objSBE.Set_SBE_Info .Fields("SBE"), .Fields("SBE_COMMENTS")
                lSBE = Val(.Fields("SBE") & "")
                strSBEComments = .Fields("SBE_COMMENTS") & ""
            End With
            
            'Null dates in the db come across as 12/31/9999. If our date isn't null,
            'use the encounter level value
            If strLastExamDate <> "12/31/9999" Or strLastExamDate <> Empty Then
                mDRCLstExamDte = strLastExamDate
            Else
                'problem with the value in db, use patient level value
                mDRCLstExamDte = mPatLevelLstExmDte
            End If
            
            If strLastChangeDate <> "12/31/9999" Or strLastExamDate <> Empty Then
                mDRCLstChgDte = strLastChangeDate
            Else
                mDRCLstChgDte = mPatLevelLstChgDte
            End If
            
           Set objRS = Nothing  'hwcc 2/26/2002
           
    '''        ElseIf Not bReread Then
        Else
           '''Create new encounter
           'to create new dental_encounter_id
           Set objRS = Nothing   'hwcc 2/27/2002
           
           'SF
           sDentEncID = objDAS.GetNextID(esDENTAL_ENCOUNTER_SEQ)
'           sqlString = "SELECT DENTAL_ENCOUNTER_SEQ.NEXTVAL from dual"
'           Set objRS = objDAS.OpenRecordset(sqlString)
'           sDentEncID = objRS!NextVal
           vDntID = sDentEncID
           Set objRS = Nothing 'hwcc 2/26/2002
                       
           mDRCLstChgDte = mPatLevelLstChgDte
           mDRCLstExamDte = mPatLevelLstExmDte
           
           sApptDateTime = goCHCSIIConn.FormatDateForDBInsert(mstrApptDTS)
           'sApptDateTime = "TO_DATE('" & Format(mstrApptDTS, gcDATE_FORMAT_VB_TIME) & "'," & gcDATE_FORMAT_ORA_TIME & ")"
                       
           'Set SBE to 0, the SBE Form will update it later
           lSBE = 0
                      
           sqlString = "Insert into DENTAL_ENCOUNTER (" _
                     & "Dental_Encounter_ID,Encounter_Number,Encounter_Facility_NCID," _
                     & "Patient_Unit_ID,EligibilityFlag,Encounter_Date," _
                     & "SBE, DRC_LastExamDate, DRC_LastChangeDate, DRC)" _
                     & "Values (" & Val(sDentEncID) & ", " & Val(sEncID) & ", " & sFclID & ", " _
                     & sPtnID & ", " & "1," & sApptDateTime & "," & lSBE & _
                     ", " & goCHCSIIConn.FormatDateForDBInsert(mDRCLstExamDte) & "" & _
                     ", " & goCHCSIIConn.FormatDateForDBInsert(mDRCLstChgDte) & ", " & _
                      mPatientLevelDRC & ")"
           
           'SF and Kieran 3-5-03 'SCR 33185
           mEncLevelDRC = mPatientLevelDRC
           
           Call objDAS.ExecuteSQL(sqlString)
           mbNewDentalEncounter = True    'Flag it, its a new dental encounter, so SBE form will pop up
    
        End If
        
        objDentEncounter.SetDentalEncounter mlDentalEncounterID, cwobjPatient.UnitNumber, oEncDoc.EncounterID, _
            oEncDoc.FacilityNCID, oEncDoc.ApptDTS, mbNewDentalEncounter, mEncLevelDRC, mDRCLstExamDte, _
            mDRCLstChgDte, mPatientLevelDRC, mPatLevelLstExmDte, mPatLevelLstChgDte, lSBE, _
            strSBEComments
            
        Set mobjDentalEncounter = objDentEncounter
        If Not IsEmpty(vDntID) And Not IsNull(vDntID) Then
           msDentalEncounterID = CStr(vDntID)
           mlDentalEncounterID = vDntID
        End If
    End If 'mblnUseStoredProc

    'KDunne 10/22/2002
    'SCR 28687
    If mblnDebugLogging Then
        strLogMsg = "Dental Encounter ID = " & mlDentalEncounterID & _
            ", Patient Level DRC = " & mPatientLevelDRC & ", Patient Level Last Exam Date = " & _
            mPatLevelLstExmDte & ", Patient Level Last Change Date = " & mPatLevelLstChgDte & _
            ", Encounter Level DRC = " & mEncLevelDRC & ", Encounter Level Last Exam Date = " & _
            mDRCLstExamDte & ", Encounter Level Last Change Date = " & mDRCLstChgDte & vbCrLf & _
            "Appointment Date/Time = " & mstrApptDTS
        Log_Event "After Open Encounter", strLogMsg
    End If
    
    If mPatientLevelDRC <> 4 Then

       If mDRCLstExamDte = "" Then
          mDRCLstExamDte = cDefaultDate
       End If
       
       If mPatLevelLstExmDte = "" Then mPatLevelLstExmDte = cDefaultDate 'scr 25725
       
'        If DateDiff("m", Format(mPatLevelLstExmDte, gcMONTH_FORMAT_VB), getCurrentDate) > 13 _
           Or mPatientLevelDRC = 0 Then
        If IsDRCExpired(mPatLevelLstExmDte) Or mPatientLevelDRC = 0 Then
          '(indicating that the monthly stored procedure did not properly update this patient info)
          'OR iDRC is empty (blank) (i.e. iDRC is "0" since "0" is default value for Integer)
          'Therefore OverallDRC and OverallLastChangeDate should be updated
           Call SaveDentalPatient("4", sPtnID, "", mstrApptDTS)
           'Update encounter level
           Call SaveDentalEncounter(4, msDentalEncounterID, "", mstrApptDTS)
        End If
    End If

FuncExit:
   Set mobjSBE = mobjDentalEncounter.SBE

   Set objRS = Nothing
   Set objDAS = Nothing
   Set oEncDoc = Nothing
   Set objSBE = Nothing
   Set objDentEncounter = Nothing
   
   'KDunne 7/16/2002
   'SCR 25736 - Always call display function.
   If mEncLevelDRC <> 0 Then Call DisplayDentalPatient(mEncLevelDRC, sPtnEncID)
   
   mblnDentalEncLoaded = Not msDentalEncounterID = ""
   
   M_bOpenDentalEncounter = Not msDentalEncounterID = ""
   
    Exit Function
ErrHnd:
   M_bOpenDentalEncounter = False
   Set objRS = Nothing
   Set objDAS = Nothing
   
   Call gobjShared.ShowVBError(Err.Number, Err.Description, "Dental.M_bOpenDentalEncounter", errAppTitle, vbExclamation)

    GoTo FuncExit
End Function

Private Function getOralExam() As String

    Dim strSQL As String
    Dim objDAS As ICHCSII_DAS
    Dim objRS As Recordset
    Dim sOralExam As String
    
    On Error GoTo errHandle
    getOralExam = ""
    Set objDAS = GetDasObj()

    strSQL = "SELECT OralExam from TEMPL_EXAMGENERAL WHERE " _
           & "Dental_Encounter_ID = " & msDentalEncounterID
    
    Set objRS = objDAS.OpenRecordset(strSQL)
    
    If Not objRS.EOF Then
        sOralExam = objRS(0) & ""
    End If
    
    getOralExam = sOralExam
           
    Set objRS = Nothing  'hwcc 2/26/2002
    Set objDAS = Nothing  'hwcc 2/26/2002
        
    Exit Function
  
errHandle:
    
    Set objRS = Nothing  'hwcc 2/26/2002
    Set objDAS = Nothing  'hwcc 2/26/2002
    
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "Dental.getOralExam", errAppTitle, vbCritical)
       
End Function

Private Function getMaxDxDRC() As Integer
'KDunne 7/9/2002
'SCR 20759 - Instead of loading the max Dx DRC from the database, we're now
'            loading it from the datalayer.
'
    Dim objGroups As CHCSII_Dental_DXTX.colGroups
    Dim objGroup As CHCSII_Dental_DXTX.clsGroup
    Dim objDXs As CHCSII_Dental_DXTX.colDxs
    Dim objDx As CHCSII_Dental_DXTX.clsDx
    Dim intMaxDRC As Integer
    
    
    Dim strSQL As String
    Dim objDAS As ICHCSII_DAS
    Dim objRS As Recordset
    
    On Error GoTo errHandle
    
    getMaxDxDRC = 0
    intMaxDRC = 0
    
    If NotesObj Is Nothing Then
        GoTo FuncExit
    End If
    
    Set objGroups = NotesObj.Groups
    
    If objGroups Is Nothing Then
         GoTo FuncExit
    End If
    
    For Each objGroup In objGroups
        Set objDXs = objGroup.Dxs
    
        For Each objDx In objDXs
            With objDx
                If .StatusID = dxActiveNew Or .StatusID = dxActiveOngoing Or .StatusID = dxActiveProvisional Then
                    If .DRC > intMaxDRC Then
                        intMaxDRC = .DRC
                    End If
                End If
            
            End With 'objDx
        
        Next 'objDX
    Next 'objGroup
    
'    Set objDAS = GetDasObj()
'
'    strSQL = "SELECT MAX(DENTAL_ENCOUNTER_DX.DRC) AS MAXDXDRC " _
'           & "FROM DENTAL_ENCOUNTER, DENTAL_ENCOUNTER_DX " _
'           & "WHERE DENTAL_ENCOUNTER.DENTAL_ENCOUNTER_ID = DENTAL_ENCOUNTER_DX.DENTAL_ENCOUNTER_ID " _
'           & "AND (DENTAL_ENCOUNTER.PATIENT_UNIT_ID = " & sPtnEncID & " ) AND " _
'           & " (DENTAL_ENCOUNTER_DX.STATUS_ID = 10 OR " _
'           & "DENTAL_ENCOUNTER_DX.STATUS_ID = 11 OR " _
'           & "DENTAL_ENCOUNTER_DX.STATUS_ID = 12)"
'
'    Set objRS = objDAS.OpenRecordset(strSQL)
'
'    Dim vMaxDRC
'    vMaxDRC = objRS.Fields("MAXDXDRC")
'    If Not objRS.EOF And Not IsNull(vMaxDRC) Then
'        getMaxDxDRC = CInt(vMaxDRC)
'    End If
'
'    Set objRS = Nothing  'hwcc 2/26/2002
'    Set objDAS = Nothing  'hwcc 2/26/2002
    
'    If Not objRS Is Nothing Then
'       Set objRS = Nothing
'    End If
'    If Not objDAS Is Nothing Then
'       Set objDAS = Nothing
'    End If
    
FuncExit:
    getMaxDxDRC = intMaxDRC
    
    Set objGroups = Nothing
    Set objGroup = Nothing
    Set objDXs = Nothing
    Set objDx = Nothing
    
    Set objRS = Nothing  'hwcc 2/26/2002
    Set objDAS = Nothing  'hwcc 2/26/2002
    
    Exit Function
  
errHandle:
           
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "Dental.getMaxDxDRC", errAppTitle, vbCritical)
    GoTo FuncExit
End Function


Private Function getPatientLastChangeDate(ByVal PatientID As String) As String

    'ASamedov - 02/11/02
    Dim sLastChangeDate As String
    '
    
    On Error GoTo errHandle
    sLastChangeDate = ""
    
    'sLastChangeDate = mRetrieveGenericMMICDR(PatientID, NCID_LAST_DENTAL_READINESS_CLASSIFICATION_CHANGE_DATE)
    sLastChangeDate = ReadDentalReadiness(PatientID, NCID_LAST_DENTAL_READINESS_CLASSIFICATION_CHANGE_DATE)


    'KDunne 5/29/2002
    'SCR 24191 - For new patients, Last Change date is coming across as 12/31/9999.
    If Trim(sLastChangeDate) = "12/31/9999" Then
        sLastChangeDate = cDefaultDate
    End If

    getPatientLastChangeDate = sLastChangeDate
       
    '''''''
    If Dental_Log Then
        Open "c:\temp\dental.log" For Append As #1
        'Print #1, CStr(Now) & ",getPatientLastChangeDate(),Patient-" & sPtnEncID & ",Dental Encounter-" & msDentalEncounterID & ",The highest DRC for the patient-" & getPatientLastChangeDate(PatientID)
        Close #1
    End If
    '''''''
    
    Exit Function
  
errHandle:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "Dental.getPatientLastChangeDate", errAppTitle, vbCritical)
    
End Function

Private Function getEncApptDate() As Date
 
    Dim strSQL As String
    Dim objRS As Recordset
    Dim objDAS As ICHCSII_DAS
    
    On Error GoTo errHandle
     
    getEncApptDate = Now
    
    Set objDAS = GetDasObj()

    strSQL = "SELECT " & goCHCSIIConn.FormatDateForDBSelect("APPT_DATE_TIME", MM_DD_YYYY_HH_MM_SS) _
              & " FROM APPOINTMENT" _
              & " WHERE ENCOUNTER_NUMBER = " & Val(sEncID) _
              & " AND PATIENT_UNIT_NUMBER = " & sPtnID _
              & " AND FACILITY_NCID = " & sFclID

'    strSQL = "SELECT TO_CHAR(APPT_DATE_TIME, 'MM/DD/YYYY hh24:mi:ss')" _
'              & " FROM APPOINTMENT" _
'              & " WHERE ENCOUNTER_NUMBER = " & Val(sEncID) _
'              & " AND PATIENT_UNIT_NUMBER = " & sPtnID _
'              & " AND FACILITY_NCID = " & sFclID
                  
    Set objRS = objDAS.OpenRecordset(strSQL)
    If Not objRS.EOF Then
        getEncApptDate = CDate(objRS(0)) 'CDate(objRS.Fields("APPT_DATE_TIME"))
    End If
           
    'TEMP (for future date testing)
    'getEncApptDate = CDate(Format("04/04/2003 10:10:10 AM", gcDATE_FORMAT_VB_TIME))
    '
                      
    Set objRS = Nothing  'hwcc 2/26/2002
    Set objDAS = Nothing  'hwcc 2/26/2002
    
'    If Not objDAS Is Nothing Then
'       Set objDAS = Nothing
'    End If
'    If Not objRS Is Nothing Then
'       Set objRS = Nothing
'    End If
    
    Exit Function
    
errHandle:
    Set objRS = Nothing  'hwcc 2/26/2002
    Set objDAS = Nothing  'hwcc 2/26/2002
    
'    If Not objDAS Is Nothing Then
'       Set objDAS = Nothing
'    End If
'    If Not objRS Is Nothing Then
'       Set objRS = Nothing
'    End If
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "Dental.getEncApptDate", errAppTitle, vbCritical)
  
End Function

Private Function SaveDentalReadiness(ByVal iDRC As Integer, _
                                   ByVal PatientID As String, _
                                   Optional ByVal strLastExamDate As String = "", _
                                   Optional ByVal strLastChangeDate As String = "" _
                                   ) As Boolean

    Dim sSQL As String
    Dim sCurDate As String
    Dim objDAS As ICHCSII_DAS
    Dim sMaxLastRDT As String
    Dim objRS As ADODB.Recordset
    
    Dim HIV_TEST_DT, HIV_TEST_RSLT, G6PD_TEST_DT, G6PD_TEST_RSLT, SICKLE_CELL_TEST_DT
    Dim SICKLE_CELL_TEST_RSLT, BLOOD_TYPE_TEST_DT, BLOOD_TYPE_TEST_RSLT, HEAR_EXAM_DT
    Dim HEAR_AID_REQUIRED, HEAR_AID_ISSUED, HEAR_AID_ISSUED_DT, LAST_DENTAL_EXAM_DT
    Dim DENTAL_READINESS_DT, MEDIC_ASSESSMENT_DT, TEMP_PROFILE, TEMP_PROFILE_START_DT
    Dim TEMP_PROFILE_END_DT, TEMP_PROFILE_DIGNOS_CD, TEMP_PROFILE_DIGNOS_DSC, TEMP_PROFILE_LIMITATION
    Dim LAST_VISUAL_ACUITY_DT, UNCORRECTED_OD, CORRECTED_OD, UNCORRECTED_OS, CORRECTED_OS
    Dim UNCORRECTED_OU, CORRECTED_OU, PROTECT_INSERT_REQUIRED, PROTECT_INSERT_ISSUED
    Dim PROTECT_INSERT_EXAM_DT, DNA_ON_FILE
    Dim strLogMsg As String
    
    On Error GoTo errHandle
    
    Set objDAS = GetDasObj()
    sCurDate = goCHCSIIConn.FormatDisplayDate(getCurrentDate, MM_DD_YYYY_HH_MM_SS)
    'sCurDate = Format(getCurrentDate, "mm/dd/yyyy hh:mm:ss")
    
    sSQL = "SELECT MAX(LAST_READINESS_DT) FROM READINESS WHERE PATIENT_UNIT_NUMBER=" & CLng(PatientID)
    Set objRS = objDAS.OpenRecordset(sSQL)
    If Not objRS.EOF Then
        sMaxLastRDT = goCHCSIIConn.FormatDisplayDate(objRS(0), MM_DD_YYYY_HH_MM_SS)
        'sMaxLastRDT = Format(objRS(0), "mm/dd/yyyy hh:mm:ss")
    Else
        'KDunne 10/22/2002
        'SCR 28687
        If mblnDebugLogging Then
            strLogMsg = " objRS.EOF = True"
            Log_Event "SaveDentalReadiness - Early Exit", strLogMsg
        End If
        
        SaveDentalReadiness = False
        Exit Function
    End If
    
    'KDunne 10/22/2002
    'SCR 28687
    If mblnDebugLogging Then
        strLogMsg = " sCurDate = " & sCurDate & ", sMaxLastRDT = " & sMaxLastRDT
        Log_Event "SaveDentalReadiness - After Getting Max LAST_READINESS_DT", strLogMsg
    End If
    
    If sCurDate <> sMaxLastRDT Then
        sCurDate = goCHCSIIConn.FormatDateForDBInsert(sCurDate)
        'sCurDate = "TO_DATE('" & sCurDate & "','MM/DD/YYYY HH24:MI:SS')"
    ElseIf mPatientLevelDRC <> iDRC Then
        sSQL = "UPDATE READINESS SET DENTAL_READINESS_CD=" & CStr(iDRC) & " WHERE PATIENT_UNIT_NUMBER=" & PatientID & " And LAST_READINESS_DT=" & goCHCSIIConn.FormatDateForDBInsert(sMaxLastRDT)
        objDAS.ExecuteSQL (sSQL)
        
        'KDunne 10/22/2002
        'SCR 28687
        If mblnDebugLogging Then
            strLogMsg = " SQL string used: " & sSQL
            Log_Event "SaveDentalReadiness - Early Exit 2", strLogMsg
        End If
        SaveDentalReadiness = True
        Exit Function
    End If
        
    sSQL = "SELECT HIV_TEST_DT,HIV_TEST_RSLT," _
            & " G6PD_TEST_DT,G6PD_TEST_RSLT,SICKLE_CELL_TEST_DT,SICKLE_CELL_TEST_RSLT,BLOOD_TYPE_TEST_DT," _
            & " BLOOD_TYPE_TEST_RSLT,HEAR_EXAM_DT,HEAR_AID_REQUIRED,HEAR_AID_ISSUED,HEAR_AID_ISSUED_DT," _
            & " LAST_DENTAL_EXAM_DT,DENTAL_READINESS_DT, MEDIC_ASSESSMENT_DT," _
            & " TEMP_PROFILE, TEMP_PROFILE_START_DT, TEMP_PROFILE_END_DT,TEMP_PROFILE_DIGNOS_CD," _
            & " TEMP_PROFILE_DIGNOS_DSC, TEMP_PROFILE_LIMITATION, LAST_VISUAL_ACUITY_DT, UNCORRECTED_OD," _
            & " CORRECTED_OD, UNCORRECTED_OS, CORRECTED_OS, UNCORRECTED_OU, CORRECTED_OU," _
            & " PROTECT_INSERT_REQUIRED, PROTECT_INSERT_ISSUED, PROTECT_INSERT_EXAM_DT,DNA_ON_FILE" _
            & " FROM READINESS WHERE PATIENT_UNIT_NUMBER=" & PatientID & " And LAST_READINESS_DT=" & goCHCSIIConn.FormatDateForDBInsert(sMaxLastRDT)
        
'    sSQL = "SELECT HIV_TEST_DT,HIV_TEST_RSLT," _
'            & " G6PD_TEST_DT,G6PD_TEST_RSLT,SICKLE_CELL_TEST_DT,SICKLE_CELL_TEST_RSLT,BLOOD_TYPE_TEST_DT," _
'            & " BLOOD_TYPE_TEST_RSLT,HEAR_EXAM_DT,HEAR_AID_REQUIRED,HEAR_AID_ISSUED,HEAR_AID_ISSUED_DT," _
'            & " LAST_DENTAL_EXAM_DT,DENTAL_READINESS_DT, MEDIC_ASSESSMENT_DT," _
'            & " TEMP_PROFILE, TEMP_PROFILE_START_DT, TEMP_PROFILE_END_DT,TEMP_PROFILE_DIGNOS_CD," _
'            & " TEMP_PROFILE_DIGNOS_DSC, TEMP_PROFILE_LIMITATION, LAST_VISUAL_ACUITY_DT, UNCORRECTED_OD," _
'            & " CORRECTED_OD, UNCORRECTED_OS, CORRECTED_OS, UNCORRECTED_OU, CORRECTED_OU," _
'            & " PROTECT_INSERT_REQUIRED, PROTECT_INSERT_ISSUED, PROTECT_INSERT_EXAM_DT,DNA_ON_FILE" _
'            & " FROM READINESS WHERE PATIENT_UNIT_NUMBER=" & PatientID & " And LAST_READINESS_DT=TO_DATE('" & sMaxLastRDT & "', 'MM/DD/YYYY HH24:MI:SS')"
                      
    'KDunne 10/22/2002
    'SCR 28687
    If mblnDebugLogging Then
        strLogMsg = " SQL string used: " & sSQL
        Log_Event "SaveDentalReadiness - Select Previous Readiness Info", strLogMsg
    End If
    
    Set objRS = objDAS.OpenRecordset(sSQL)
       
    If Not objRS.EOF Then
        HIV_TEST_DT = goCHCSIIConn.FormatDateForDBInsert(objRS("HIV_TEST_DT"))
'        HIV_TEST_DT = Format(objRS("HIV_TEST_DT"), "mm/dd/yyyy hh:mm:ss")
'        HIV_TEST_DT = "TO_DATE('" & HIV_TEST_DT & "','MM/DD/YYYY HH24:MI:SS')"
        HIV_TEST_RSLT = objRS("HIV_TEST_RSLT") & ""
        
        G6PD_TEST_DT = goCHCSIIConn.FormatDateForDBInsert(objRS("G6PD_TEST_DT"))
'        G6PD_TEST_DT = Format(objRS("G6PD_TEST_DT"), "mm/dd/yyyy hh:mm:ss")
'        G6PD_TEST_DT = "TO_DATE('" & G6PD_TEST_DT & "','MM/DD/YYYY HH24:MI:SS')"
        
        G6PD_TEST_RSLT = objRS("G6PD_TEST_RSLT") & ""
        SICKLE_CELL_TEST_DT = goCHCSIIConn.FormatDateForDBInsert(objRS("SICKLE_CELL_TEST_DT"))
'        SICKLE_CELL_TEST_DT = Format(objRS("SICKLE_CELL_TEST_DT"), "mm/dd/yyyy hh:mm:ss")
'        SICKLE_CELL_TEST_DT = "TO_DATE('" & SICKLE_CELL_TEST_DT & "','MM/DD/YYYY HH24:MI:SS')"
        
        SICKLE_CELL_TEST_RSLT = objRS("SICKLE_CELL_TEST_RSLT") & ""
        BLOOD_TYPE_TEST_DT = goCHCSIIConn.FormatDateForDBInsert(objRS("BLOOD_TYPE_TEST_DT"))
'        BLOOD_TYPE_TEST_DT = Format(objRS("BLOOD_TYPE_TEST_DT"), "mm/dd/yyyy hh:mm:ss")
'        BLOOD_TYPE_TEST_DT = "TO_DATE('" & BLOOD_TYPE_TEST_DT & "','MM/DD/YYYY HH24:MI:SS')"
        
        BLOOD_TYPE_TEST_RSLT = objRS("BLOOD_TYPE_TEST_RSLT") & ""
        
        HEAR_EXAM_DT = goCHCSIIConn.FormatDateForDBInsert(objRS("HEAR_EXAM_DT"))
'        HEAR_EXAM_DT = Format(objRS("HEAR_EXAM_DT"), "mm/dd/yyyy hh:mm:ss")
'        HEAR_EXAM_DT = "TO_DATE('" & HEAR_EXAM_DT & "','MM/DD/YYYY HH24:MI:SS')"
                
        HEAR_AID_REQUIRED = objRS("HEAR_AID_REQUIRED") & ""
        HEAR_AID_ISSUED = objRS("HEAR_AID_ISSUED") & ""
        HEAR_AID_ISSUED_DT = goCHCSIIConn.FormatDateForDBInsert(objRS("HEAR_AID_ISSUED_DT"))
'        HEAR_AID_ISSUED_DT = Format(objRS("HEAR_AID_ISSUED_DT"), "mm/dd/yyyy hh:mm:ss")
'        HEAR_AID_ISSUED_DT = "TO_DATE('" & HEAR_AID_ISSUED_DT & "','MM/DD/YYYY HH24:MI:SS')"
         
        MEDIC_ASSESSMENT_DT = goCHCSIIConn.FormatDateForDBInsert(objRS("MEDIC_ASSESSMENT_DT"))
'        MEDIC_ASSESSMENT_DT = Format(objRS("MEDIC_ASSESSMENT_DT"), "mm/dd/yyyy hh:mm:ss")
'        MEDIC_ASSESSMENT_DT = "TO_DATE('" & MEDIC_ASSESSMENT_DT & "','MM/DD/YYYY HH24:MI:SS')"
         
        TEMP_PROFILE = objRS("TEMP_PROFILE") & ""
        TEMP_PROFILE_START_DT = goCHCSIIConn.FormatDateForDBInsert(objRS("TEMP_PROFILE_START_DT"))
'        TEMP_PROFILE_START_DT = Format(objRS("TEMP_PROFILE_START_DT"), "mm/dd/yyyy hh:mm:ss")
'        TEMP_PROFILE_START_DT = "TO_DATE('" & TEMP_PROFILE_START_DT & "','MM/DD/YYYY HH24:MI:SS')"
         
        TEMP_PROFILE_END_DT = goCHCSIIConn.FormatDateForDBInsert(objRS("TEMP_PROFILE_END_DT"))
'        TEMP_PROFILE_END_DT = Format(objRS("TEMP_PROFILE_END_DT"), "mm/dd/yyyy hh:mm:ss")
'        TEMP_PROFILE_END_DT = "TO_DATE('" & TEMP_PROFILE_END_DT & "','MM/DD/YYYY HH24:MI:SS')"
         
        TEMP_PROFILE_DIGNOS_CD = objRS("TEMP_PROFILE_DIGNOS_CD") & ""
        TEMP_PROFILE_DIGNOS_DSC = objRS("TEMP_PROFILE_DIGNOS_DSC") & ""
        TEMP_PROFILE_LIMITATION = objRS("TEMP_PROFILE_LIMITATION") & ""
        LAST_VISUAL_ACUITY_DT = goCHCSIIConn.FormatDateForDBInsert(objRS("LAST_VISUAL_ACUITY_DT"))
'        LAST_VISUAL_ACUITY_DT = Format(objRS("LAST_VISUAL_ACUITY_DT"), "mm/dd/yyyy hh:mm:ss")
'        LAST_VISUAL_ACUITY_DT = "TO_DATE('" & LAST_VISUAL_ACUITY_DT & "','MM/DD/YYYY HH24:MI:SS')"
         
        UNCORRECTED_OD = objRS("UNCORRECTED_OD") & ""
        CORRECTED_OD = objRS("CORRECTED_OD") & ""
        UNCORRECTED_OS = objRS("UNCORRECTED_OS") & ""
        CORRECTED_OS = objRS("CORRECTED_OS") & ""
        UNCORRECTED_OU = objRS("UNCORRECTED_OU") & ""
        CORRECTED_OU = objRS("CORRECTED_OU") & ""
        PROTECT_INSERT_REQUIRED = objRS("PROTECT_INSERT_REQUIRED") & ""
        PROTECT_INSERT_ISSUED = objRS("PROTECT_INSERT_ISSUED") & ""
        PROTECT_INSERT_EXAM_DT = goCHCSIIConn.FormatDateForDBInsert(objRS("PROTECT_INSERT_EXAM_DT"))
'        PROTECT_INSERT_EXAM_DT = Format(objRS("PROTECT_INSERT_EXAM_DT"), "mm/dd/yyyy hh:mm:ss")
'        PROTECT_INSERT_EXAM_DT = "TO_DATE('" & PROTECT_INSERT_EXAM_DT & "','MM/DD/YYYY HH24:MI:SS')"
         
        DNA_ON_FILE = objRS("DNA_ON_FILE") & ""
        
        If Len(strLastExamDate) > 0 Then
            LAST_DENTAL_EXAM_DT = goCHCSIIConn.FormatDisplayDate(strLastExamDate, MM_DD_YYYY_HH_MM_SS)
            'LAST_DENTAL_EXAM_DT = Format(strLastExamDate, "mm/dd/yyyy hh:mm:ss")
        Else
            LAST_DENTAL_EXAM_DT = goCHCSIIConn.FormatDisplayDate(objRS("LAST_DENTAL_EXAM_DT"), MM_DD_YYYY_HH_MM_SS)
            'LAST_DENTAL_EXAM_DT = Format(objRS("LAST_DENTAL_EXAM_DT"), "mm/dd/yyyy hh:mm:ss")
        End If
        LAST_DENTAL_EXAM_DT = goCHCSIIConn.FormatDateForDBInsert(LAST_DENTAL_EXAM_DT)
        'LAST_DENTAL_EXAM_DT = "TO_DATE('" & LAST_DENTAL_EXAM_DT & "','MM/DD/YYYY HH24:MI:SS')"
    
        If Len(strLastChangeDate) > 0 Then
            DENTAL_READINESS_DT = goCHCSIIConn.FormatDisplayDate(strLastChangeDate, MM_DD_YYYY_HH_MM_SS)
            'DENTAL_READINESS_DT = Format(strLastChangeDate, "mm/dd/yyyy hh:mm:ss")
        Else
            DENTAL_READINESS_DT = goCHCSIIConn.FormatDisplayDate(objRS("DENTAL_READINESS_DT"), MM_DD_YYYY_HH_MM_SS)
            'DENTAL_READINESS_DT = Format(objRS("DENTAL_READINESS_DT"), "mm/dd/yyyy hh:mm:ss")
        End If
        DENTAL_READINESS_DT = goCHCSIIConn.FormatDateForDBInsert(DENTAL_READINESS_DT)
        'DENTAL_READINESS_DT = "TO_DATE('" & DENTAL_READINESS_DT & "','MM/DD/YYYY HH24:MI:SS')"
        
        sSQL = "INSERT INTO READINESS (LAST_READINESS_DT,PATIENT_UNIT_NUMBER,HIV_TEST_DT,HIV_TEST_RSLT," _
            & " G6PD_TEST_DT,G6PD_TEST_RSLT,SICKLE_CELL_TEST_DT,SICKLE_CELL_TEST_RSLT,BLOOD_TYPE_TEST_DT," _
            & " BLOOD_TYPE_TEST_RSLT,HEAR_EXAM_DT,HEAR_AID_REQUIRED,HEAR_AID_ISSUED,HEAR_AID_ISSUED_DT," _
            & " LAST_DENTAL_EXAM_DT, DENTAL_READINESS_CD, DENTAL_READINESS_DT, MEDIC_ASSESSMENT_DT," _
            & " TEMP_PROFILE, TEMP_PROFILE_START_DT, TEMP_PROFILE_END_DT,TEMP_PROFILE_DIGNOS_CD," _
            & " TEMP_PROFILE_DIGNOS_DSC, TEMP_PROFILE_LIMITATION, LAST_VISUAL_ACUITY_DT, UNCORRECTED_OD," _
            & " CORRECTED_OD, UNCORRECTED_OS, CORRECTED_OS, UNCORRECTED_OU, CORRECTED_OU," _
            & " PROTECT_INSERT_REQUIRED, PROTECT_INSERT_ISSUED, PROTECT_INSERT_EXAM_DT,DNA_ON_FILE)" _
            & " VALUES(" & sCurDate & "," & CLng(PatientID) & "," & HIV_TEST_DT & ",'" & HIV_TEST_RSLT _
            & "'," & G6PD_TEST_DT & ",'" & G6PD_TEST_RSLT & "'," & SICKLE_CELL_TEST_DT & ",'" & SICKLE_CELL_TEST_RSLT & "'," & BLOOD_TYPE_TEST_DT _
            & ",'" & BLOOD_TYPE_TEST_RSLT & "'," & HEAR_EXAM_DT & ",'" & HEAR_AID_REQUIRED & "','" & HEAR_AID_ISSUED & "'," & HEAR_AID_ISSUED_DT _
            & "," & LAST_DENTAL_EXAM_DT & "," & CStr(iDRC) & "," & DENTAL_READINESS_DT & "," & MEDIC_ASSESSMENT_DT _
            & ",'" & TEMP_PROFILE & "'," & TEMP_PROFILE_START_DT & "," & TEMP_PROFILE_END_DT & ",'" & TEMP_PROFILE_DIGNOS_CD _
            & "','" & TEMP_PROFILE_DIGNOS_DSC & "','" & TEMP_PROFILE_LIMITATION & "'," & LAST_VISUAL_ACUITY_DT & ",'" & UNCORRECTED_OD _
            & "','" & CORRECTED_OD & "','" & UNCORRECTED_OS & "','" & CORRECTED_OS & "','" & UNCORRECTED_OU & "','" & CORRECTED_OU _
            & "','" & PROTECT_INSERT_REQUIRED & "','" & PROTECT_INSERT_ISSUED & "'," & PROTECT_INSERT_EXAM_DT & ",'" & DNA_ON_FILE & "')"
    Else
        If Len(strLastExamDate) > 0 And Len(strLastChangeDate) > 0 Then
            LAST_DENTAL_EXAM_DT = goCHCSIIConn.FormatDateForDBInsert(strLastExamDate)
'            LAST_DENTAL_EXAM_DT = Format(strLastExamDate, "mm/dd/yyyy hh:mm:ss")
'            LAST_DENTAL_EXAM_DT = "TO_DATE('" & LAST_DENTAL_EXAM_DT & "','MM/DD/YYYY HH24:MI:SS')"
            
            DENTAL_READINESS_DT = goCHCSIIConn.FormatDateForDBInsert(strLastChangeDate)
'            DENTAL_READINESS_DT = Format(strLastChangeDate, "mm/dd/yyyy hh:mm:ss")
'            DENTAL_READINESS_DT = "TO_DATE('" & DENTAL_READINESS_DT & "','MM/DD/YYYY HH24:MI:SS')"
            
            sSQL = "INSERT INTO READINESS (LAST_READINESS_DT,PATIENT_UNIT_NUMBER, LAST_DENTAL_EXAM_DT, DENTAL_READINESS_CD, DENTAL_READINESS_DT)" _
            & " VALUES(" & sCurDate & "," & CLng(PatientID) _
            & "," & LAST_DENTAL_EXAM_DT & "," & CStr(iDRC) & "," & DENTAL_READINESS_DT & ")"
        ElseIf Len(strLastExamDate) > 0 And Len(strLastChangeDate) = 0 Then
            LAST_DENTAL_EXAM_DT = goCHCSIIConn.FormatDateForDBInsert(strLastExamDate)
'            LAST_DENTAL_EXAM_DT = Format(strLastExamDate, "mm/dd/yyyy hh:mm:ss")
'            LAST_DENTAL_EXAM_DT = "TO_DATE('" & LAST_DENTAL_EXAM_DT & "','MM/DD/YYYY HH24:MI:SS')"
            
            sSQL = "INSERT INTO READINESS (LAST_READINESS_DT,PATIENT_UNIT_NUMBER, LAST_DENTAL_EXAM_DT, DENTAL_READINESS_CD)" _
            & " VALUES(" & sCurDate & "," & CLng(PatientID) _
            & "," & LAST_DENTAL_EXAM_DT & "," & CStr(iDRC) & ")"
        
        ElseIf Len(strLastExamDate) = 0 And Len(strLastChangeDate) > 0 Then
            DENTAL_READINESS_DT = goCHCSIIConn.FormatDateForDBInsert(strLastChangeDate)
'            DENTAL_READINESS_DT = Format(strLastChangeDate, "mm/dd/yyyy hh:mm:ss")
'            DENTAL_READINESS_DT = "TO_DATE('" & DENTAL_READINESS_DT & "','MM/DD/YYYY HH24:MI:SS')"
            sSQL = "INSERT INTO READINESS (LAST_READINESS_DT,PATIENT_UNIT_NUMBER, DENTAL_READINESS_CD, DENTAL_READINESS_DT)" _
            & " VALUES(" & sCurDate & "," & CLng(PatientID) _
            & "," & CStr(iDRC) & "," & DENTAL_READINESS_DT & ")"
                    
        Else
            sSQL = "INSERT INTO READINESS (LAST_READINESS_DT,PATIENT_UNIT_NUMBER, DENTAL_READINESS_CD)" _
            & " VALUES(" & sCurDate & "," & CLng(PatientID) _
            & ",'" & CStr(iDRC) & "')"
        End If
    End If
        
    'KDunne 10/22/2002
    'SCR 28687
    If mblnDebugLogging Then
        strLogMsg = " SQL string used: " & sSQL
        Log_Event "SaveDentalReadiness - Final Update Statement", strLogMsg
    End If
    
    objDAS.ExecuteSQL (sSQL)
    SaveDentalReadiness = True
    Exit Function
    
errHandle:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "Dental.SaveDentalReadiness", errAppTitle, vbCritical)
      
End Function

Private Function ReadDentalReadiness(ByVal PatientID As String, ByVal NCID As String) As String
    Dim sSQL As String
    Dim objRS As ADODB.Recordset
    Dim objDAS As ICHCSII_DAS
    Dim varValue
    
    On Error GoTo errHandle
    
    Set objDAS = GetDasObj()
    If NCID = NCID_DENTAL_READINESS_CLASSIFICATION Then
        sSQL = "SELECT DENTAL_READINESS_CD FROM READINESS WHERE PATIENT_UNIT_NUMBER=" & CLng(PatientID)
        sSQL = sSQL & " AND LAST_READINESS_DT=(SELECT MAX(LAST_READINESS_DT) FROM READINESS WHERE PATIENT_UNIT_NUMBER=" & CLng(PatientID) & ")"
    ElseIf NCID = NCID_LAST_DENTAL_READINESS_CLASSIFICATION_CHANGE_DATE Then
        sSQL = "SELECT DENTAL_READINESS_DT FROM READINESS WHERE PATIENT_UNIT_NUMBER=" & CLng(PatientID)
        sSQL = sSQL & " AND LAST_READINESS_DT=(SELECT MAX(LAST_READINESS_DT) FROM READINESS WHERE PATIENT_UNIT_NUMBER=" & CLng(PatientID) & ")"
    ElseIf NCID = NCID_LAST_DENTAL_EXAM_DATE Then
        sSQL = "SELECT LAST_DENTAL_EXAM_DT FROM READINESS WHERE PATIENT_UNIT_NUMBER=" & CLng(PatientID)
        sSQL = sSQL & " AND LAST_READINESS_DT=(SELECT MAX(LAST_READINESS_DT) FROM READINESS WHERE PATIENT_UNIT_NUMBER=" & CLng(PatientID) & ")"
    Else
        Exit Function
    End If
    Set objRS = objDAS.OpenRecordset(sSQL)
    
    If Not objRS.EOF Then
        varValue = objRS(0) & ""
    End If
    ReadDentalReadiness = CStr(varValue)
    
    Exit Function
    
errHandle:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "Dental.ReadDentalReadiness", errAppTitle, vbCritical)
                                   
End Function

Private Function getPSR() As String
    Dim strLowerLeft As String
    Dim strLowerCenter As String
    Dim strLowerRight As String
    Dim strUpperLeft As String
    Dim strUpperCenter As String
    Dim strUpperRight As String
    
    Dim sNotes As String
    
    Dim sSQL As String
    Dim objRS As ADODB.Recordset
    Dim objDAS As ICHCSII_DAS
    
    On Error GoTo errHandle
    
    If mblnUseStoredProc Then
        With svgCommObj.mObjPSR
            strLowerLeft = formatPSR(.lowerLeft)
            strLowerCenter = formatPSR(.lowerCenter)
            strLowerRight = formatPSR(.lowerRight)
            
            strUpperLeft = formatPSR(.upperLeft)
            strUpperCenter = formatPSR(.upperCenter)
            strUpperRight = formatPSR(.upperRight)

        End With
    
    Else
        Set objDAS = GetDasObj()
        
        sSQL = "SELECT PSR_LOWER_LEFT, PSR_LOWER_CENTER, PSR_LOWER_RIGHT, PSR_UPPER_LEFT," _
                & " PSR_UPPER_CENTER , PSR_UPPER_RIGHT From DENTAL_ENCOUNTER" _
                & " WHERE DENTAL_ENCOUNTER_ID = " & msDentalEncounterID
                
        Set objRS = objDAS.OpenRecordset(sSQL)
        
        If Not objRS.EOF Then
            strLowerLeft = formatPSR(objRS("PSR_LOWER_LEFT") & "")
            strLowerCenter = formatPSR(objRS("PSR_LOWER_CENTER") & "")
            strLowerRight = formatPSR(objRS("PSR_LOWER_RIGHT") & "")
            
            strUpperLeft = formatPSR(objRS.Fields("PSR_UPPER_LEFT") & "")
            strUpperCenter = formatPSR(objRS.Fields("PSR_UPPER_CENTER") & "")
            strUpperRight = formatPSR(objRS.Fields("PSR_UPPER_RIGHT") & "")
        End If
        
    End If
    
    If Trim(strLowerLeft) <> "" Or Trim(strLowerCenter) <> "" Or Trim(strLowerRight) <> "" _
        Or Trim(strUpperLeft) <> "" Or Trim(strUpperCenter) <> "" Or Trim(strUpperRight) <> "" Then
        sNotes = "\viewkind4\uc1\pard\ul\b\f0\fs20 PSR:\ulnone\b0\par "
        sNotes = sNotes & "    " & strUpperRight & "|" & strUpperCenter & "|" & strUpperLeft & "    \par " & vbCrLf
        sNotes = sNotes & "R ------------------- L" & "\par " & vbCrLf
        sNotes = sNotes & "    " & strLowerRight & "|" & strLowerCenter & "|" & strLowerLeft & "    \par " & vbCrLf
    End If
            

        
    'If sNotes <> "" Then
    '    sNotes = "\viewkind4\uc1\pard\ul\b\f0\fs20 PSR Information:\ulnone\b0\par PSR: " & sNotes
    'End If
    
    getPSR = sNotes
    Exit Function
    
errHandle:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "Dental.getPSR", errAppTitle, vbCritical)

End Function

Private Function getOrientation() As String
    Dim sSQL As String
    Dim objRS As ADODB.Recordset
    Dim objDAS As ICHCSII_DAS
    Dim sToothName As String
    Dim sCurToothName As String
    Dim sOrientation As String
    Dim sNotes As String
    Dim bMesial, bDistal, bFacial, bLingual
    Dim bMesialRot, bDistalRot, bIntrude, bExtrude
    
    On Error GoTo errHandle
    
    Set objDAS = GetDasObj()
    
    sSQL = "SELECT TOOTH_NAME,MESIAL_VERSION,DISTAL_VERSION,FACIAL_VERSION,LINGUAL_VERSION,ROTATED_MESIAL,ROTATED_DISTAL,INTRUDED,EXTRUDED FROM DENTAL_ENCOUNTER_TEETH WHERE DENTAL_ENCOUNTER_ID="
    sSQL = sSQL & msDentalEncounterID & " ORDER BY DENTAL_ENCOUNTER_TEETH_ID DESC"
    Set objRS = objDAS.OpenRecordset(sSQL)
    
    Do While Not objRS.EOF
        sCurToothName = objRS("TOOTH_NAME") & ""
        If sToothName <> objRS("TOOTH_NAME") Then
            sToothName = objRS("TOOTH_NAME") & ""
            
            bFacial = objRS("FACIAL_VERSION")
            bDistal = objRS("LINGUAL_VERSION")
                        
            bMesialRot = objRS("ROTATED_MESIAL")
            bDistalRot = objRS("ROTATED_DISTAL")
                        
            bMesial = objRS("MESIAL_VERSION")
            bDistal = objRS("DISTAL_VERSION")
         
            bIntrude = objRS("INTRUDED")
            bExtrude = objRS("EXTRUDED")
                        
            If bMesial Then
                sOrientation = "Mesial" 'mesiofacial
            ElseIf bDistal Then
                sOrientation = "Distal" 'distofacial
            End If
            
            If bFacial Then
                sOrientation = sOrientation & "Facial"
            ElseIf bLingual Then
                sOrientation = sOrientation & "Lingual"
            End If
            
            If bMesialRot Then
                If sOrientation <> "" Then
                    sOrientation = sOrientation & "," & "Rotated Mesially"
                Else
                    sOrientation = "Rotated Mesially"
                End If
            ElseIf bDistalRot Then
                If sOrientation <> "" Then
                    sOrientation = sOrientation & "," & "Rotated Distally"
                Else
                    sOrientation = "Rotated Distally"
                End If
            End If
            
            If bIntrude Then
                If sOrientation <> "" Then
                    sOrientation = sOrientation & "," & "Intruded"
                Else
                    sOrientation = "Intruded"
                End If
            ElseIf bExtrude Then
                If sOrientation <> "" Then
                    sOrientation = sOrientation & "," & "Extruded"
                Else
                    sOrientation = "Extruded"
                End If
            End If

        End If
        
        If sOrientation <> "" Then
            'Displacement - Tooth 3 ; Rotation:Distal ; Tilt:Mesial-Facial ; Height:intruded
            sNotes = sNotes & "Displacement - Tooth " & sToothName & "; " & sOrientation & "\par" & vbCrLf
        End If
                
        sOrientation = ""
        objRS.MoveNext
    Loop
    
    If sNotes <> "" Then
        sNotes = "\viewkind4\uc1\pard\ul\b\f0\fs20 Orientation:\ulnone\b0\par " & sNotes
    End If
    
    getOrientation = sNotes
    Exit Function
errHandle:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "Dental.getOrientation", errAppTitle, vbCritical)

End Function

Private Function formatPSR(ByRef sInput As String) As String
    If Len(sInput) = 0 Then
        formatPSR = "    "
    ElseIf Len(sInput) = 1 Then
        formatPSR = "  " & sInput & "  "
    Else
        formatPSR = " " & sInput & "  "
    End If
End Function

'*********************************************************************************
'Sub:Unload_Dental
'Author:KDunne
'Date:5/21/2002
'Purpose: Unloads Dental form and all dental children
'*********************************************************************************
Sub Unload_Dental()
    
    For Each mDentalChild In mcolDentalChild
        mDentalChild.CloseChild
    Next
    
    Unload frmDental
    
    Set mcolDentalChild = Nothing
    
    If Not gobjEnc Is Nothing Then
          gobjEnc.ClosedView
    End If
    
End Sub

'*********************************************************************************
'Function: GetLockedByUserName
'Author:KDunne
'Date:5/22/2002
'Purpose: If the encounter has been locked, get the name of the user who locked it
'*********************************************************************************
Private Function GetLockedByMsg(ByVal lngEncounterNumber As Long) As String
    Dim objDAS As ICHCSII_DAS
    Dim objRS As ADODB.Recordset
    Dim strSQL As String
    Dim strMsg As String
    Dim strLockUserNcid As String
    Dim strLockSectionNumber As String
    Dim strLockEventID As String
    Dim strLockItemName As String
    Dim strLockWorkstation As String
    Dim strLockUserName As String
    Dim strLockTime As String
    
    On Error GoTo ErrHandler
    
    Set objDAS = GetDasObj()

    strSQL = "select EL.*, 24 * 60 * (sysdate - date_time) as lock_minutes" & _
        " from encounter_lock EL " & _
        " where " & _
        "  EL.encounter_number = " & sEncID & _
        " and EL.section_number = 11 " 'KDunne 5/22/2002 - 11 is Dental's section number
    
    Set objRS = objDAS.OpenRecordset(strSQL)
    
    If Not objRS Is Nothing Then
        If objRS.EOF And objRS.BOF Then GoTo FuncExit
        
        With objRS
            strLockUserNcid = .Fields("USER_NCID")
            strLockSectionNumber = .Fields("SECTION_NUMBER")
            strLockEventID = .Fields("EVENT_ID")
            strLockItemName = .Fields("ITEM_NAME")
            strLockWorkstation = .Fields("WORKSTATION")
            strLockTime = Format$(.Fields("DATE_TIME"), gsDefaultDateTime)
        End With 'objRS
    
        strSQL = "Select distinct Provider_Name " & _
              " From Clinic_Provider CP " & _
              " Where CP.CLINICIAN_NCID = " & strLockUserNcid
              
        Set objRS = objDAS.OpenRecordset(strSQL)
        
        If Not objRS Is Nothing Then
            If objRS.EOF And objRS.BOF Then Exit Function
            
            strLockUserName = objRS.Fields("Provider_Name")
        End If
    
    End If
    
    strMsg = "This encounter was locked by " & strLockUserName & " on workstation '" _
        & strLockWorkstation & "' at " & strLockTime & "." & vbCrLf
    
FuncExit:
    
    'If we weren't able to get our info from the encounter_lock table, give the user the old
    'message.
    '
    If Trim(strMsg) = "" Then
        strMsg = "Your lock on the Dental section of the current encounter has been taken away by another user."
    End If
    
    GetLockedByMsg = strMsg
    Set objRS = Nothing
    Set objDAS = Nothing
    Exit Function
ErrHandler:
    strMsg = Empty
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "Dental.GetLockedByMsg", errAppTitle, vbExclamation)

    GoTo FuncExit
End Function

'*****************************************************************************
'Function:LoadReadinessDRCInfo
'Author:KDunne
'Date: 7/9/2002
'Purpose: Loads DRC Info from the readiness table.
'Notes: SCR 20759 - Cominbing three seperate queries into one. Function replaces
'                   ReadDentalReadiness table.
'*****************************************************************************
Private Function LoadReadinessDRCInfo(lngPatientID As Long) As Boolean
    Dim strSQL As String
    Dim objDAS As ICHCSII_DAS
    Dim objRS As ADODB.Recordset
    Dim blnReturn As Boolean
    
    On Error GoTo ErrHandler
    
    InitDRCInfo
    
    strSQL = "SELECT DENTAL_READINESS_CD, DENTAL_READINESS_DT,LAST_DENTAL_EXAM_DT  " & _
                " FROM READINESS " & _
                " WHERE PATIENT_UNIT_NUMBER=" & lngPatientID & _
                " AND LAST_READINESS_DT=(SELECT MAX(LAST_READINESS_DT) FROM READINESS WHERE PATIENT_UNIT_NUMBER=" & lngPatientID & ")"
    
    Set objDAS = GetDasObj
    Set objRS = objDAS.OpenRecordset(strSQL)
    
    If Not objRS Is Nothing Then
        If objRS.RecordCount > 0 Then
            mPatientLevelDRC = Val(objRS.Fields("DENTAL_READINESS_CD") & "")
            mPatLevelLstChgDte = objRS.Fields("DENTAL_READINESS_DT") & ""
            mPatLevelLstExmDte = objRS.Fields("LAST_DENTAL_EXAM_DT") & ""
            
            If mPatientLevelDRC = 0 Then
                'ASamedov - 02/08/02
                'Update Patient level
                mPatientLevelDRC = 4
                Call SaveDentalPatient(mPatientLevelDRC, lngPatientID)
            End If
            
            'KDunne 5/29/2002
            'SCR 24191 - For new patients, dates are coming across as 12/31/9999.
            If Trim(mPatLevelLstChgDte) = "12/31/9999" Or Trim(mPatLevelLstChgDte) = "" Then
                mPatLevelLstChgDte = cDefaultDate
            End If
            
            If Trim(mPatLevelLstExmDte) = "12/31/9999" Or Trim(mPatLevelLstExmDte) = "" Then
                mPatLevelLstExmDte = cDefaultDate
            End If
            'End 24191
            
        End If
    End If
    blnReturn = True
FuncExit:
    Set objDAS = Nothing
    Set objRS = Nothing
    
    LoadReadinessDRCInfo = blnReturn
    
    Exit Function
ErrHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "Dental.LoadReadinessDRCInfo", errAppTitle, vbExclamation)
    blnReturn = False
    InitDRCInfo
    GoTo FuncExit
End Function

Private Sub InitDRCInfo()
    mPatLevelLstExmDte = cDefaultDate
    mPatLevelLstChgDte = cDefaultDate
    mPatientLevelDRC = 4
End Sub

'*********************************************************************************
'Function:SaveEncounter
'Author:KDunne
'Date:7/25/2002
'Purpose: Save Functionality was spread out throughout Dental Parent. This function
'         will store it one place.
'**********************************************************************************
Private Function SaveEncounter(ByVal blnSaveAll As Boolean, ByVal blnClose As Boolean) As Boolean
    Dim blnReturn As Boolean
    Dim bResult As Boolean
    
    'KDunne 8/22/2002
    'SCR 26879 - Disable Core while saving.
'    cwobjCWComm.Message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_DDSA_APPLICATION
    
    'SF 11/22/03 Fix for ITT, gobjEnc is nothing when switching patients. It may be due to 832 core or encounter?
    If Not gobjEnc Is Nothing Then
        If Not gobjEnc.StartSave() Then
            'KDunne 5/21/2002
            'SCR 22001
            CWSafeMsgBox GetLockedByMsg(sEncID) & "Your unsaved changes will be discarded." & vbCrLf
            If blnClose Then
                Unload_Dental
            End If
            blnReturn = False
            GoTo FuncExit
            'End scr 22001
        End If
    Else
        blnReturn = True
        GoTo FuncExit
    End If
    
    If blnSaveAll Then
        For Each mDentalChild In mcolDentalChild
            bResult = mDentalChild.Module.SaveView()
            If bResult = False Then
                CWSafeMsgBox "error - mDentalChild.Module.SaveView()"  '2/27/2002
            End If
        Next
    Else
        bResult = mcolDentalChild(CStr(giCurrentServer)).Module.SaveView()
        If bResult = False Then
            CWSafeMsgBox "error - mDentalChild.Module.SaveView()"   'hwcc 2/27/2002
        End If
    End If
                
    ' get iDRC from Big Tooth and notes
    SetEncDentalNotes msDentalEncounterID
            
    ' update name bar on the screen
    
    gobjEnc.EndSave

    If blnClose Then
        'KDunne 5/21/2002
        'SCR 22001
        Unload_Dental
        
        svgCommObj.TeethObject.clearSelectedTeeth
        Call DisplayDentalPatient(0, sPtnEncID)

    End If
    blnReturn = True
    
FuncExit:
    SaveEncounter = blnReturn
    
    'KDunne 8/22/2002
    'SCR 26879 - Re-enable Core.
'    cwobjCWComm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_DDSA_APPLICATION
    Exit Function
    
ErrHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "Dental.SaveEncounter", errAppTitle, vbExclamation)
    blnReturn = False
    'KDunne 12/10/2002
    'SCR 30602 - Make sure we unlock the encounter in the event of an error.
    If Not gobjEnc Is Nothing Then
        gobjEnc.EndSave
    End If
    GoTo FuncExit
End Function

'*******************************************************************************
'Property: Misc_RTF
'Author:KDunne
'Date:7/25/2002
'Purpose:BigTooth does not have a way of updating the Encounter Document directly.
'        This function will allow Dental Children to write to the encounter document
'        directly.
'********************************************************************************
Public Property Let Misc_RTF(ByVal strValue As String)
    mstrMiscRTF = strValue
End Property

Public Property Get Misc_RTF() As String
    Misc_RTF = mstrMiscRTF
End Property

Public Property Let HelpId(ByVal lngHelpId As Long)
    frmDental.HelpContextID = lngHelpId
End Property

Public Property Get HelpId() As Long
    HelpId = frmDental.HelpContextID
End Property

'**********************************************************************
'Function:IsDRCExpired
'Author:KDunne
'Date:8/30/2002
'Purpose:Given the patient's last exam date, function checks to see if
'        patient's drc has expired.
'***********************************************************************
Private Function IsDRCExpired(ByVal dteLastExamDate As Date) As Boolean
    Dim blnReturn As Boolean
    
    If DateDiff("m", Format(dteLastExamDate, gcMONTH_FORMAT_VB), getCurrentDate) > 13 Then
        blnReturn = True
    Else
        blnReturn = False
    End If

    IsDRCExpired = blnReturn
    
End Function

Public Property Get SBE() As IntrfcCHCSII_Dental.IDENTAL_SBE
    Set SBE = mobjSBE
End Property

Public Sub EnablePrintIcon()

Dim oEncDoc As Encounter
    
    If mbEnable603aPrint = True And mbSVGDrawn = False Then
        'Send a message to SVG to determine if the SVG has drawn
        cwobjCWComm.Message cwiGENERIC, "ISSVGDRAWN|", NCID_DENTAL_SVG, NCID_DDSA_APPLICATION
        'This is called recursively, so exit
        Exit Sub
    End If
    
    If mbEnable603aPrint = True And mbSVGDrawn = True Then
        If glMenuPrintID > 0 Then
            Call cwobjCWComm.UpdateMenuItem(glMenuPrintID, "ENABLE")
        Else
            Set oEncDoc = gobjEnc.CurrentEncounter
            glMenuPrintID = gobjEnc.AddMenuItem(0, "&SF603A", "Print.bmp", "Preview SF603A")
        End If
    Else
        'Add it but disable it because it shouldn't show yet. This is to "reserve" it's place in the toolbar.
        Set oEncDoc = gobjEnc.CurrentEncounter
        glMenuPrintID = gobjEnc.AddMenuItem(0, "&SF603A", "Print.bmp", "Preview SF603A")
        cwobjCWComm.UpdateMenuItem glMenuPrintID, "DISABLE"
    End If
    
End Sub

Public Sub DisablePrintIcon()
        cwobjCWComm.UpdateMenuItem glMenuPrintID, "DISABLE"
End Sub

'*******************************************************************************************
'Sub:Log_Event
'Author:Kdunne
'Date 10/22/2002
'Purpose: Used for logging debug messages. Currently saving messages to event viewer.
'*******************************************************************************************
Private Sub Log_Event(ByVal strMsgLocation As String, ByVal strMessage As String, _
    Optional ByVal lngEventType As Long = vbLogEventTypeInformation)
    With App
        .LogEvent strMsgLocation & vbCrLf & strMessage, lngEventType
    End With
    
'
'    Open "c:\temp\dental.log" For Append As #1
'    Print #1, ""
'    Print #1, CStr(Now) & "," & strMsgLocation & "," & strMessage
'    Close #1

End Sub

Private Function ConvertSPDate(ByVal dteDate As Date) As Date
    Dim dteReturn As Date
    

    dteReturn = cwobjCWComm.GmtToLocal(dteDate)

    
    ConvertSPDate = dteReturn
End Function
