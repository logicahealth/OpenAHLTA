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
Attribute VB_Name = "clsCariesRisk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Implements clsISection

Private mlngDentalEncounterId As Long
Private mlngPatUnitId As Long
Private mlngDirtyFlag As DirtyFlagEnum
Private mlngId As Long
Private mlngProviderId As Long

Private mobjProtocol As clsCRProtocol
Private mobjRecallVisit As clsCRRecall

'SCR 8499
'Private mobjExamGeneral As clsExamGeneral
Private mobjExamSOAP As clsExamSOAP

Private mobjPreventive As clsPreventive
Private mobjDataLayer As CHCSII_Dental_DXTX.clsNotes
Private mlngPrevDentEnc As Long
Private mlngProtocolState As CariesRiskProtocol_CariesStateEnum
Private mdteRecallDate As Date
Private mlngPrevState As CariesRiskProtocol_CariesStateEnum
Private mdteEncDate As Date

Private mobjDentEnc As IntrfcCHCSII_Dental.IDental_Enc

Public Enum CariesRiskProtocol_CariesStateEnum
    crpCS_None
    crpCS_Low
    crpCS_Moderate
    crpCS_High
    crpCS_RecallWait
    crpCS_HighRecall
End Enum

Private Const cADULT_AGE As Integer = 13
Private Const const_ModuleName = "clsCariesRisk" '*** Insert Module Name Here

Private Sub Class_Terminate()
    Set mobjProtocol = Nothing
    Set mobjRecallVisit = Nothing
    'Set mobjExamGeneral = Nothing
    Set mobjExamSOAP = Nothing
End Sub
Public Property Get HasData() As Boolean
    If mobjProtocol.HasData Or mobjRecallVisit.HasData Then
        HasData = True
    Else
        HasData = False
    End If
End Property

'****************************************************************
'Function:CheckProtocolComplete
'Author:KDunne
'Date:1/7/2002
'Purpose: Called From Dental Parent. Checks to see if Protocol was completed
'         during this encounter. If it wasn't, display the appropriate Form,
'         if needed.
'*****************************************************************
Public Sub CheckProtocolComplete()
    Dim objCurrSection As clsISection
    Dim blnUsePrevState As Boolean
    Dim objProtoSection As clsISection
    Dim lngSaveState As CariesRiskProtocol_CariesStateEnum
    
    Set objCurrSection = Me
        
    Call objCurrSection.LoadSection(mlngPatUnitId, mlngDentalEncounterId)
    
    Select Case mlngProtocolState
        Case CariesRiskProtocol_CariesStateEnum.crpCS_High, CariesRiskProtocol_CariesStateEnum.crpCS_Moderate
            If mobjProtocol.Protocol_Completed Then
                blnUsePrevState = True
            Else
                blnUsePrevState = False
            End If
            Set objProtoSection = mobjProtocol
            
        Case CariesRiskProtocol_CariesStateEnum.crpCS_HighRecall
            If mobjRecallVisit.Completed Then
                blnUsePrevState = True
            Else
                blnUsePrevState = False
            End If
            Set objProtoSection = mobjRecallVisit
        Case CariesRiskProtocol_CariesStateEnum.crpCS_Low
            If DateDiff("d", mdteRecallDate, Now()) = 0 Then
                blnUsePrevState = True
                Set objProtoSection = mobjProtocol
            Else
                blnUsePrevState = False
            End If
        Case CariesRiskProtocol_CariesStateEnum.crpCS_RecallWait
            If DateDiff("d", mdteRecallDate, Now()) = 0 Then
                blnUsePrevState = True
                If mlngPrevState = crpCS_High Then
                    Set objProtoSection = mobjProtocol
                Else
                    Set objProtoSection = mobjRecallVisit
                End If
            Else
                blnUsePrevState = False
            End If

    End Select
    
    If objProtoSection Is Nothing = False Then
        objProtoSection.LoadSection mlngPatUnitId, mlngDentalEncounterId
    End If
    If LoadProtocolForm(blnUsePrevState) Then
        objCurrSection.SaveSection
    End If
    
End Sub

Public Property Get Caries_Risk_State() As CariesRiskProtocol_CariesStateEnum
    Caries_Risk_State = mlngProtocolState
End Property
Public Property Get Preventive() As clsPreventive
    Set Preventive = mobjPreventive
End Property
Public Property Set Preventive(ByVal objObject As clsPreventive)
    Set mobjPreventive = objObject
End Property

Public Property Get Examsoap() As clsExamSOAP
    Set Examsoap = mobjExamSOAP
End Property


Public Property Set Examsoap(ByVal objObject As clsExamSOAP)
    Set mobjExamSOAP = objObject
End Property

Public Property Get DataLayer() As CHCSII_Dental_DXTX.clsNotes

    Set DataLayer = mobjDataLayer
End Property

Public Property Set DataLayer(ByVal objDataLayer As CHCSII_Dental_DXTX.clsNotes)
    Set mobjDataLayer = objDataLayer
End Property

Public Property Get ProviderId() As Long
    ProviderId = mlngProviderId
End Property

Public Property Let ProviderId(ByVal lngValue As Long)
    mlngProviderId = lngValue
End Property

'***********************************************************************************
'Function:GetCariesStateEnumDesc
'Author:KDunne
'Date: 12/21/2001
'Purpose: Returns the description for a given CariesRiskProtocol_CariesStateEnum value
'***********************************************************************************
Public Function GetCariesStateEnumDesc(ByVal lngValue As CariesRiskProtocol_CariesStateEnum) As String
    Dim strReturn As String
    
    strReturn = Empty
    On Error GoTo ErrHandler
    Select Case lngValue
        Case CariesRiskProtocol_CariesStateEnum.crpCS_High
            strReturn = "High"
        Case CariesRiskProtocol_CariesStateEnum.crpCS_HighRecall
            strReturn = "High Recall"
        Case CariesRiskProtocol_CariesStateEnum.crpCS_Low
            strReturn = "Low"
        Case CariesRiskProtocol_CariesStateEnum.crpCS_Moderate
            strReturn = "Moderate"
        Case CariesRiskProtocol_CariesStateEnum.crpCS_None
            strReturn = Empty
        Case CariesRiskProtocol_CariesStateEnum.crpCS_RecallWait
            strReturn = "Recall Wait"
        Case Else
            strReturn = Empty
    End Select
    
    GetCariesStateEnumDesc = strReturn
    Exit Function
ErrHandler:
    GetCariesStateEnumDesc = Empty
    goCWShared.ShowVBError Err.Number, Err.Description, const_ModuleName & ".GetCariesStateEnumDesc()", "", vbCritical
End Function

'**********************************************************************************
'Function: CariesRiskChanged
'Author: KDunne
'Date: 12/31/2001
'Purpose: Loads the Caries Risk Protocol screens when the user manually changes
'         the Caries Risk.
'***********************************************************************************
Public Function ChangeCariesRiskState(ByVal lngNewState As CariesRiskProtocol_CariesStateEnum) As Boolean
    Dim objCurrSect As clsISection
    
    mlngProtocolState = lngNewState
    UpdateDirtyFlag
    'We've changed our Protocol state, reset the Protocols.
    If mobjProtocol.HasData = False Then
        'We need to have the Protocol Id, if a previous protocol has
        'been associated with this encounter
        Set objCurrSect = mobjProtocol
        Call objCurrSect.LoadSection(mlngPatUnitId, mlngDentalEncounterId)
    End If
    
    If mobjRecallVisit.HasData = False Then
        'We need to have the Recall Id, if a previous protocol has
        'been associated with this encounter
        Set objCurrSect = mobjRecallVisit
        Call objCurrSect.LoadSection(mlngPatUnitId, mlngDentalEncounterId)
    End If
    
    
    mobjProtocol.Reset_Values
    mobjRecallVisit.Reset_Values
    
    'Load the Protocol Form.
    ChangeCariesRiskState = LoadProtocolForm
End Function
'**********************************************************************************
'Function: LoadProtocolForm
'Author: KDunne
'Date: 12/31/2001
'Purpose: Checks the current protocol state, and loads the necessary protocol screen.
'***********************************************************************************
Public Function LoadProtocolForm(Optional ByVal blnPrevState As Boolean = False) As Boolean
    Dim X As Integer
    Dim blnReturn As Boolean
    Dim lngFormState As CariesRiskProtocol_ProtocolFormStateEnum
    Dim objCurrSection As clsISection
    Dim objRS As ADODB.Recordset
    Dim lngProtocolState As CariesRiskProtocol_CariesStateEnum
    Dim blnPatIsAdult As Boolean
    
    On Error GoTo ErrHandler
    
    If cwobjCWComm Is Nothing Then
        blnReturn = False
        GoTo FuncExit
    End If
    If blnPrevState Then
        lngProtocolState = mlngPrevState
    Else
        lngProtocolState = mlngProtocolState
    End If
    
     Select Case lngProtocolState
        Case CariesRiskProtocol_CariesStateEnum.crpCS_Low, CariesRiskProtocol_CariesStateEnum.crpCS_None
            LoadProtocolForm = False
            Exit Function
        Case CariesRiskProtocol_CariesStateEnum.crpCS_Moderate
            lngFormState = crpPFS_Moderate
        Case CariesRiskProtocol_CariesStateEnum.crpCS_High
            lngFormState = crpPFS_High
        Case CariesRiskProtocol_CariesStateEnum.crpCS_HighRecall
            lngFormState = crpPFS_Recall
        Case CariesRiskProtocol_CariesStateEnum.crpCS_RecallWait
            'We're in the recall wait state. Don't show protocol
            'form until patient is back in the High Risk recall state.
            LoadProtocolForm = False
            Exit Function
        Case Else
            'Unknown state
            LoadProtocolForm = False
            Exit Function
    End Select
       
    '--- SCR-ARM-25468 RWalker 10/21/03
    If Not goARM Is Nothing Then goARM.StartCmplxTx gsARM_Tx_CorrUT_ProtocolComplete, , , , , , , , eAnyFamily
       
    'KDunne 10/2/2002
    'SCR 20759 - We're only loading the Preventive object here.
    '
    'KDunne 1/9/2004
    'SCR 47838 Age calculation was incorrectly rounding up.
    '
    blnPatIsAdult = Val(goCWShared.CalcAge(cwobjPatient.DateOfBirth)) >= cADULT_AGE


    If mobjPreventive Is Nothing Then
        Set mobjPreventive = New clsPreventive
        mobjPreventive.PatientIsAdult = blnPatIsAdult
        Set objCurrSection = mobjPreventive
        Call objCurrSection.LoadSection(mlngPatUnitId, mlngDentalEncounterId)
    End If
    
    With frmCariesRiskDialog
        .Form_State = lngFormState
        .LoadCariesRiskProtocol mobjProtocol
        .LoadDataLayer mobjDataLayer
        
        'SCR 8499
        .LoadExamSOAP mobjExamSOAP
        '.LoadExamGeneral mobjExamGeneral
        
        .LoadPreventive mobjPreventive
        .LoadRecallVisit mobjRecallVisit
        .DentalEncounterID = mlngDentalEncounterId
        .PatientUnitId = mlngPatUnitId
        .EncounterDate = mdteEncDate
        
        'Only one tab on the Caries Risk Protocol can be displayed at a time. So it
        'will either be the High Risk Recall screen, or one of the protocol screens.
        If lngFormState = crpPFS_Recall Then
            Set objCurrSection = mobjRecallVisit
            
            .chkVisitDiet.Value = ConvertBoolToCB(mobjRecallVisit.Diet_Counseling)
            .chkVisitTest.Value = ConvertBoolToCB(mobjRecallVisit.Strep_Mutans_Test)
            .chkVisitSalivaTest.Value = ConvertBoolToCB(mobjRecallVisit.Salivary_Test)

        Else 'MOderate or High Risk Tab
        
            mobjProtocol.Protocol_State = mlngProtocolState
            mobjProtocol.Form_State = lngFormState
            
            'Dental Nutritional Counseling section
            .chkDentNutr(lngFormState).Value = ConvertBoolToCB(mobjProtocol.Dental_Nutrional_Counseling)
            .chkExpEtiology(lngFormState).Value = ConvertBoolToCB(mobjProtocol.Exlain_Etiology_Caries)
            .chkSucroseSrc(lngFormState).Value = ConvertBoolToCB(mobjProtocol.Identify_Sucrose_Sources)
            .chkDietMod(lngFormState).Value = ConvertBoolToCB(mobjProtocol.Recommend_Diet_Mods)
            .chkGum(lngFormState).Value = ConvertBoolToCB(mobjProtocol.Xylitol_Chewing_Gum)
            .ucNoteDentNut(lngFormState).Comments = mobjProtocol.Nutrition_Comments
            .chkOralHygiene(lngFormState).Value = ConvertBoolToCB(mobjProtocol.Oral_Hygiene_Instructions)
            
            .chkCariesElim(lngFormState).Value = ConvertBoolToCB(mobjProtocol.CheckForCariesElimination)
            .chkFlourAppsCmplt(lngFormState).Value = ConvertBoolToCB(mobjProtocol.Fluoride_Complete)
            .chkAddFlourTx(lngFormState).Value = ConvertBoolToCB(mobjProtocol.Additional_Fluoride_Tx)
            .chkSealants(lngFormState).Value = ConvertBoolToCB(mobjProtocol.Sealants_Considered)
                        
            'Meds Section
            
            .chkDentifrice(lngFormState).Value = ConvertBoolToCB(mobjProtocol.Rx_Dentifrice)
            .chkOTC(lngFormState).Value = ConvertBoolToCB(mobjProtocol.Rx_OTC)
            .chkChlor(lngFormState).Value = ConvertBoolToCB(mobjProtocol.Rx_Chlorhexidine)
        
            'Check for Template workload credits
            If mobjPreventive.PRV_WC_Nutritional Then
                mobjProtocol.Dental_Counseling_Template_WL = True
            Else
                mobjProtocol.Dental_Counseling_Template_WL = False
            End If
            
            If mobjPreventive.PRV_OH_Brushing Or mobjPreventive.PRV_OH_Flossing Then
                mobjProtocol.Oral_Hygiene_Template_WL = True
            Else
                mobjProtocol.Oral_Hygiene_Template_WL = False
            End If
        End If
        
        'Disable all of the tabs.
        .sstCariesRisk.TabEnabled(CariesRiskProtocol_ProtocolFormStateEnum.crpPFS_High) = False
        .sstCariesRisk.TabEnabled(CariesRiskProtocol_ProtocolFormStateEnum.crpPFS_Moderate) = False
        .sstCariesRisk.TabEnabled(CariesRiskProtocol_ProtocolFormStateEnum.crpPFS_Recall) = False
        
        'Now, just enable the tab that we'll be working with.
        .sstCariesRisk.TabEnabled(lngFormState) = True
        
        'Set the tab
        .sstCariesRisk.Tab = lngFormState
        
        'Show the form. Broadcast the appropriate messages.
        cwobjCWComm.Message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, cwlMyProgId
        .Show vbModal
        cwobjCWComm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, cwlMyProgId
        
    
        If .User_Completed Then
            mdteRecallDate = Now()
            mlngPrevState = mlngProtocolState
            mlngProtocolState = .Caries_Risk_Protocol_State
            'Update Caries Risk dirty flag (current object)
            UpdateDirtyFlag
        End If
    
    End With
    
    
    Unload frmCariesRiskDialog
    Set objCurrSection = mobjRecallVisit
    objCurrSection.SaveSection
    
    Set objCurrSection = mobjProtocol
    objCurrSection.SaveSection
    
    Set objCurrSection = mobjPreventive
    objCurrSection.SaveSection
    
    Set objCurrSection = Me
    objCurrSection.SaveSection
    
    If Not mobjExamSOAP Is Nothing Then
        Set objCurrSection = mobjExamSOAP
        objCurrSection.SaveSection
    End If
    
    blnReturn = True
    
    '--- SCR-ARM-25468 RWalker 10/21/03
    If Not goARM Is Nothing Then goARM.StopCmplxTx gsARM_Tx_CorrUT_ProtocolComplete, enmARMTxStatusOk, enmARMStop_MeOnly

FuncExit:
    LoadProtocolForm = blnReturn
    Set objCurrSection = Nothing
    
    Exit Function
       
ErrHandler:
    goCWShared.ShowVBError Err.Number, Err.Description, const_ModuleName & ".LoadProtocolForm()", "", vbCritical
    blnReturn = False
    GoTo FuncExit
End Function
Private Property Get clsISection_DentalEncounterID() As Long
    clsISection_DentalEncounterID = mlngDentalEncounterId
End Property

Private Function clsISection_GetNote() As String
    Dim strNote As String
    Dim objCurrSection As clsISection
    
    strNote = Empty
    
    If mobjProtocol Is Nothing = False Then
        Set objCurrSection = mobjProtocol
        strNote = objCurrSection.GetNote()
    End If
    
    If mobjRecallVisit Is Nothing = False Then
        Set objCurrSection = mobjRecallVisit
        strNote = strNote & objCurrSection.GetNote()
    End If
    
    clsISection_GetNote = strNote
  
FuncExit:
    Set objCurrSection = Nothing
    Exit Function
       
ErrHandler:
    clsISection_GetNote = Empty
    goCWShared.ShowVBError Err.Number, Err.Description, const_ModuleName & ".clsISection_GetNote()", "", vbCritical
    GoTo FuncExit
End Function
Private Function clsISection_GetRTFNote() As String
    Dim strNote As String
    Dim objCurrSection As clsISection
    
    strNote = Empty
    
    If mobjProtocol Is Nothing = False Then
        Set objCurrSection = mobjProtocol
        strNote = objCurrSection.GetRTFNote()
    End If
    
    If mobjRecallVisit Is Nothing = False Then
        Set objCurrSection = mobjRecallVisit
        strNote = strNote & objCurrSection.GetRTFNote()
    End If
    
    clsISection_GetRTFNote = strNote
  
FuncExit:
    Set objCurrSection = Nothing
    Exit Function
       
ErrHandler:
    clsISection_GetRTFNote = Empty
    goCWShared.ShowVBError Err.Number, Err.Description, const_ModuleName & ".clsISection_GetRTFNote()", "", vbCritical
    GoTo FuncExit
End Function
Private Function clsISection_LoadSection(ByVal PatientID As Long, ByVal DentalEncounterID As Long) As Boolean
    Dim blnReturn As Boolean
    Dim objRS As ADODB.Recordset
    Dim objCurrSection As clsISection
    Dim blnNewRecall As Boolean
    Dim strDate As String
    
    On Error GoTo ErrHandler
    
    mlngPatUnitId = PatientID
    mlngDentalEncounterId = DentalEncounterID
            
    If goconn Is Nothing Then
        OpenDB
        'Init the XMLPersist class and load the XML data
        Dim XMLPersist As New clsXMLPersist
        XMLPersist.ExamXML.Load DentalEncounterID
    End If
    
    Set objRS = DB_Load_Caries_Risk(PatientID)
    
    With objRS

        If .EOF Then
            mlngDirtyFlag = dfAdd
            mlngDentalEncounterId = DentalEncounterID
            mlngProtocolState = CariesRiskProtocol_CariesStateEnum.crpCS_None
            mlngPrevState = CariesRiskProtocol_CariesStateEnum.crpCS_None
        Else
            
            mlngDirtyFlag = dfNoAction
            
            mlngPrevDentEnc = Val(.Fields("CARIES_LST_DENTAL_ENCOUNTER_ID") & "")
            mlngProtocolState = Val(.Fields("CARIES_PROTOCOL_STATE") & "")
            mlngPrevState = Val(.Fields("CARIES_PREV_STATE") & "")
            strDate = .Fields("CARIES_RECALL_DATE") & ""
            'KDunne 6/28/2002
            'SCR 25171 - In theater, the IIF statement evaluates both the true part and the
            '            false part at the same time. This does not happen in cdr mode. WIth
            '            a new patient, our date is an empty string, when you do a CDate on an
            '            empty string, you get a Type mismatch error.
'''            mdteRecallDate = IIf(IsDate(strDate), CDate(strDate), CDate(0))
            If IsDate(strDate) Then
                mdteRecallDate = CDate(strDate)
            Else
                mdteRecallDate = CDate(0)
            End If
            'Check to see if its time to come out of Recall Wait state
            If mlngProtocolState = CariesRiskProtocol_CariesStateEnum.crpCS_RecallWait Then
                'We re-evaluate the patient after 3 months has elapsed.
                If DateDiff("m", mdteRecallDate, Now()) >= 3 Then
                    mlngPrevState = mlngProtocolState
                    mlngProtocolState = CariesRiskProtocol_CariesStateEnum.crpCS_HighRecall
                    blnNewRecall = True
                    UpdateDirtyFlag
                End If
            
            End If
            
            'KDunne 12/19/2003
            'SCR 47071 - If this is a new encounter for this protocol, we want to save.
            If mlngPrevDentEnc <> mlngDentalEncounterId Then
                UpdateDirtyFlag
            End If
        End If
        
        mlngPatUnitId = PatientID
    End With

    If mobjDentEnc Is Nothing Then
        mdteEncDate = GetEncounterDate(mlngDentalEncounterId)
    Else
        mdteEncDate = mobjDentEnc.EncounterDate
    End If
        
    Set mobjProtocol = New clsCRProtocol
    Set mobjRecallVisit = New clsCRRecall
    
    mobjProtocol.Previous_Dental_Encounter_ID = mlngPrevDentEnc
    mobjProtocol.Set_Dental_Encounter_Id mlngDentalEncounterId
    
    If mlngProtocolState = crpCS_High Or mlngProtocolState = crpCS_Moderate Then
        Set objCurrSection = mobjProtocol
        Call objCurrSection.LoadSection(mlngPatUnitId, mlngDentalEncounterId)
        mobjProtocol.Protocol_State = mlngProtocolState
    End If
    
    mobjRecallVisit.Previous_Dental_Encounter_ID = mlngPrevDentEnc
    mobjRecallVisit.Set_Dental_Encounter_Id mlngDentalEncounterId
    
    If blnNewRecall = False And mlngProtocolState = crpCS_HighRecall Then
        Set objCurrSection = mobjRecallVisit
        Call objCurrSection.LoadSection(mlngPatUnitId, mlngDentalEncounterId)
    End If
      
    
    Set mobjProtocol.DataLayer = mobjDataLayer
    
    If mobjDataLayer Is Nothing Then
        Set mobjDataLayer = New CHCSII_Dental_DXTX.clsNotes
        Call mobjDataLayer.LoadGroups(cwobjPatient, cwobjLogonEnviron, cwobjConfig, cwobjCWComm, mlngDentalEncounterId, mlngProviderId)
        'Call mobjDataLayer.LoadGroups(mlngPatUnitId, mlngDentalEncounterId, mlngProviderId)
    End If
    
    
FuncExit:
    clsISection_LoadSection = blnReturn
    Set objRS = Nothing
    Exit Function
ErrHandler:
   blnReturn = False
#If debugmode = 0 Then
   goCWShared.ShowVBError Err.Number, Err.Description, const_ModuleName & ".clsISection_LoadSection()", "", vbCritical
#End If
   GoTo FuncExit
End Function

Private Property Get clsISection_PatientID() As Long
    clsISection_PatientID = mlngPatUnitId
End Property

Private Function clsISection_SaveSection() As Boolean
'NOtes:
'SCR 21679 - Caries State is now stored in Dental_Patient.
    Dim strSql As String
    Dim blnReturn As Boolean
    Dim objRS As ADODB.Recordset
    Dim objCurrSection As clsISection
    
    On Error GoTo ErrHandler
    
    strSql = Empty
    Select Case mlngDirtyFlag
        Case dfNoAction
            blnReturn = True
            GoTo FuncExit
        Case dfAdd

            strSql = "INSERT INTO DENTAL_PATIENT " & _
                     "(PATIENT_UNIT_ID, CARIES_LST_DENTAL_ENCOUNTER_ID, " & _
                     " CARIES_PROTOCOL_STATE, CARIES_RECALL_DATE, CARIES_PREV_STATE) " & _
                     "VALUES (" & mlngPatUnitId & ", " & mlngDentalEncounterId & _
                     ", " & mlngProtocolState & ", '" & Format(mdteRecallDate, "DDMMMYYYY") & "', " & mlngPrevState & ")"

        Case dfUpdate
            strSql = "UPDATE DENTAL_PATIENT " & _
                     " SET CARIES_LST_DENTAL_ENCOUNTER_ID = " & mlngDentalEncounterId & _
                     ", CARIES_PROTOCOL_STATE = " & mlngProtocolState & _
                     ", CARIES_RECALL_DATE = '" & Format(mdteRecallDate, "DDMMMYYYY") & _
                     "', CARIES_PREV_STATE = " & mlngPrevState & _
                     " WHERE PATIENT_UNIT_ID = " & mlngPatUnitId
        Case dfdelete
            strSql = "DELETE FROM TEMPL_EXAM_ENDO_FINDINGS WHERE " & _
                     " PATIENT_UNIT_ID = " & mlngPatUnitId
    End Select
    
    If strSql <> "" Then
        goconn.ExecuteSQL (strSql)
    End If
    mlngDirtyFlag = dfNoAction
    
    'If there's a collection, loop through it here.
    If mobjProtocol Is Nothing = False Then
        Set objCurrSection = mobjProtocol
        objCurrSection.SaveSection
    End If
    
    If mobjRecallVisit Is Nothing = False Then
        Set objCurrSection = mobjRecallVisit
        objCurrSection.SaveSection
    End If

FuncExit:
    Set objRS = Nothing
    Set objCurrSection = Nothing
    Exit Function
    
ErrHandler:
    blnReturn = False
    goCWShared.ShowVBError Err.Number, Err.Description, const_ModuleName & ".clsISection_SaveSection()", "", vbCritical
    GoTo FuncExit
End Function

'****************************************************
'Sub: UpdateDirtyFlag
'Author: KDunne
'Date: 11/5/2001
'Purpose: Updates module level dirty flag field
'****************************************************
Private Sub UpdateDirtyFlag()

    If mlngDirtyFlag = dfNoAction Then
        mlngDirtyFlag = dfUpdate
    End If
    
End Sub

'******************************************************************************
'Function: DB_Load_Recall_Info
'Author: KDunne
'Date: 12/27/2001
'Purpose: Loads data from the TEMPL_CARIES_RISK_RECALL_VISIT table
'NOtes:
'        2/21/2002 SCR 21679 - Caries State is now stored in Dental_Patient.
'*******************************************************************************

Private Function DB_Load_Caries_Risk(ByVal lngPatUnitId As Long) As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo ErrHandler:
    strSql = "Select PATIENT_UNIT_ID, CARIES_LST_DENTAL_ENCOUNTER_ID, CARIES_PROTOCOL_STATE, " & _
        "CARIES_RECALL_DATE, CARIES_PREV_STATE " & _
        "From DENTAL_PATIENT WHERE PATIENT_UNIT_ID = " & lngPatUnitId
        
    Set DB_Load_Caries_Risk = goconn.OpenRecordset(strSql)
    
    Exit Function
ErrHandler:
    goCWShared.ShowVBError Err.Number, Err.Description, const_ModuleName & ".DB_Load_Caries_Risk()", "", vbCritical
    Set DB_Load_Caries_Risk = Nothing
End Function

'*************************************************************************
'Function:GetEncounterDate
'Author:KDunne
'Date:3/8/2002
'Purpose:Gets the encounter date
'**************************************************************************
Private Function GetEncounterDate(ByVal lngEncId As Long) As Date
    Dim objRS As ADODB.Recordset
    Dim strSql As String
    Dim dteReturn As Date
    
    On Error GoTo ErrHandler
    
    strSql = "Select ENCOUNTER_DATE from Dental_Encounter where DENTAL_ENCOUNTER_ID = " & lngEncId
    
    Set objRS = goconn.OpenRecordset(strSql)
    
    If objRS Is Nothing Then
        dteReturn = 0
    Else
        dteReturn = objRS.Fields("ENCOUNTER_DATE")
    End If
    
FuncExit:
    GetEncounterDate = dteReturn
    Set objRS = Nothing
    Exit Function
ErrHandler:
    dteReturn = 0
    goCWShared.ShowVBError Err.Number, Err.Description, const_ModuleName & ".GetEncounterDate()", "", vbCritical
    GoTo FuncExit
End Function


'************************************************************************
'Sub:KDunne
'Author:KDunne
'Date 1/18/2002
'Purpose: Checks the current protocol caries risk state, and updates the
'         ExamGeneral object with the proper value for its Caries Risk State
'*************************************************************************
Public Sub SetExamEvalCariesRiskState()
    Select Case mlngProtocolState
        Case CariesRiskProtocol_CariesStateEnum.crpCS_Low
            mobjExamSOAP.RiskAssessEnum = RiskAssessEN.low
        Case CariesRiskProtocol_CariesStateEnum.crpCS_Moderate
            mobjExamSOAP.RiskAssessEnum = RiskAssessEN.Moderate
        
        Case CariesRiskProtocol_CariesStateEnum.crpCS_High, CariesRiskProtocol_CariesStateEnum.crpCS_HighRecall, _
            CariesRiskProtocol_CariesStateEnum.crpCS_RecallWait
            mobjExamSOAP.RiskAssessEnum = RiskAssessEN.High
        Case Else
            mobjExamSOAP.RiskAssessEnum = RiskAssessEN.NoRiskAssessValue
    End Select
    
    mobjExamSOAP.RiskAssess_Caries = mobjExamSOAP.RiskAssessEnum

End Sub

Public Property Set DentalEncounter(objDentEnc As IntrfcCHCSII_Dental.IDental_Enc)
    Set mobjDentEnc = objDentEnc
End Property
Public Sub Destroy_Objects()
    Set mobjPreventive = Nothing
    
    'SCR 8499
    Set mobjExamSOAP = Nothing
    'Set mobjExamGeneral = Nothing
    
    Set mobjProtocol = Nothing
    Set mobjRecallVisit = Nothing
    Set mobjDentEnc = Nothing
End Sub

