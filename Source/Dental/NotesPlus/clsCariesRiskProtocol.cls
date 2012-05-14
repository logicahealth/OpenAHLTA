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
Attribute VB_Name = "clsCRProtocol"
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
Private mlngPrevDentEncounterId As Long
Private mobjDataLayer As CHCSII_Dental_DXTX.clsNotes
Private mlngFormState As CariesRiskProtocol_ProtocolFormStateEnum
Private mblnDentalCounseling As Boolean
Private mblnExplainEtiology As Boolean
Private mblnSucroseSrcs As Boolean
Private mblnDietMods As Boolean
'Private mblnWrittenInstructs As Boolean
Private mblnGum As Boolean
Private mblnCariesEliminated As Boolean
Private mblnSealants As Boolean
Private mblnAddFluoride As Boolean
Private mblnFluorideCmplt As Boolean
Private mblnRxDenifrice As Boolean
Private mblnRxOTC As Boolean
Private mblnRxChlorhexidine As Boolean
Private mblnProtocolCmplt As Boolean
Private mlngProtocolState As CariesRiskProtocol_CariesStateEnum
Private mstrNutritionCmnts As String
Private mblnOralHygInstruct As Boolean

Private mlngWLDentCounseling As Long
Private mlngWLOralHygiene As Long
Private mlngPrevWLDentCounsel As Long
Private mlngPrevWLOralHygiene As Long

Private mblnWLDentCounsel As Boolean
Private mblnWLOralHygiene As Boolean
Private mblnHasData As Boolean

Private Const const_ModuleName = "clsCRProtocol" '*** Insert Module Name Here

Private Const cDentalCounselingWLCode = 257
Private Const cOralHygieneWLCode = 283
Public Property Get HasData() As Boolean
    HasData = mblnHasData
End Property
Public Sub Set_Dental_Encounter_Id(ByVal lngValue As Long)
    mlngDentalEncounterId = lngValue
End Sub
Public Property Let Form_State(ByVal lngValue As Long)
    mlngFormState = lngValue
End Property
Public Property Get Form_State() As Long
    Form_State = mlngFormState
End Property
'*****************************************************************************
'Sub:Reset_Values
'Author:KDunne
'Date:1/4/2002
'Purpose: Resets the object to a "fresh" state. Retains the id. Used when Protocol
'         is completed.
'*****************************************************************************
Public Sub Reset_Values()
    mblnDentalCounseling = False
    mblnExplainEtiology = False
    mblnSucroseSrcs = False
    mblnDietMods = False
'    mblnWrittenInstructs = False
    mblnGum = False
    mblnCariesEliminated = False
    mblnSealants = False
    mblnAddFluoride = False
    mblnFluorideCmplt = False
    mblnRxDenifrice = False
    mblnRxOTC = False
    mblnRxChlorhexidine = False
    mblnProtocolCmplt = False
    mstrNutritionCmnts = Empty
    mblnOralHygInstruct = False

    If mlngWLOralHygiene > 0 Then
        mblnWLDentCounsel = True
    Else
        mblnWLDentCounsel = False
    End If
    
    mlngWLDentCounseling = 0
    
    If mlngWLDentCounseling > 0 Then
        mblnWLDentCounsel = True
    Else
        mblnWLDentCounsel = False
    End If

    mblnHasData = False
    
    UpdateWorkLoadCodes
    If mlngId = 0 Then
        mlngDirtyFlag = dfAdd
    Else
        mlngDirtyFlag = dfUpdate
    End If
    
End Sub

'Property is used to deterimine wheter or not provider has received
'workload credit from Dental Counseling in the Preventive template.
Public Property Get Dental_Counseling_Template_WL() As Boolean
    Dental_Counseling_Template_WL = mblnWLDentCounsel
End Property
Public Property Let Dental_Counseling_Template_WL(ByVal blnValue As Boolean)
    mblnWLDentCounsel = blnValue
End Property

'Property is used to determine whether or not provider has received
'workload credit from Oral Hygiene in the preventive template.
Public Property Get Oral_Hygiene_Template_WL() As Boolean
    Oral_Hygiene_Template_WL = mblnWLOralHygiene
End Property
Public Property Let Oral_Hygiene_Template_WL(ByVal blnValue As Boolean)
    mblnWLOralHygiene = blnValue
End Property
Public Property Get Prev_WL_Dental_NutritionalCounseling() As Long
    Prev_WL_Dental_NutritionalCounseling = mlngPrevWLDentCounsel
End Property
Public Property Get Prev_WL_Oral_Hygiene_Instructions() As Long
    Prev_WL_Oral_Hygiene_Instructions = mlngPrevWLOralHygiene
End Property
Public Property Get WorkLoad_Dental_NutrionalCounseling() As Long
    WorkLoad_Dental_NutrionalCounseling = mlngWLDentCounseling
End Property

Public Property Let WorkLoad_Dental_NutrionalCounseling(ByVal lngValue As Long)
    mlngWLDentCounseling = lngValue
End Property

Public Property Get WorkLoad_Oral_Hygiene_Instructions() As Long
    WorkLoad_Oral_Hygiene_Instructions = mlngWLOralHygiene
End Property

Public Property Let WorkLoad_Oral_Hygiene_Instructions(ByVal lngValue As Long)
    mlngWLOralHygiene = lngValue
End Property

Public Property Get DataLayer() As CHCSII_Dental_DXTX.clsNotes
    Set DataLayer = mobjDataLayer
End Property

Public Property Set DataLayer(ByVal objValue As CHCSII_Dental_DXTX.clsNotes)
    Set mobjDataLayer = objValue
End Property

Public Property Get Previous_Dental_Encounter_ID() As Long
    Previous_Dental_Encounter_ID = mlngPrevDentEncounterId
End Property
Public Property Let Previous_Dental_Encounter_ID(ByVal lngValue As Long)
    mlngPrevDentEncounterId = lngValue
End Property

Public Property Get Oral_Hygiene_Instructions() As Boolean
    Oral_Hygiene_Instructions = mblnOralHygInstruct
End Property

Public Property Let Oral_Hygiene_Instructions(ByVal blnValue As Boolean)
    mblnOralHygInstruct = blnValue
End Property
Public Property Get Nutrition_Comments() As String
    Nutrition_Comments = mstrNutritionCmnts
End Property
Public Property Let Nutrition_Comments(ByVal strValue As String)
    mstrNutritionCmnts = strValue
End Property
Public Property Let Protocol_State(ByVal lngData As Long)

    mlngProtocolState = lngData
    UpdateDirtyFlag
End Property

Public Property Get Protocol_State() As Long
    Protocol_State = mlngProtocolState
End Property

Public Property Let Protocol_Completed(ByVal blnValue As Boolean)
    mblnProtocolCmplt = blnValue
    UpdateDirtyFlag
End Property


Public Property Get Protocol_Completed() As Boolean
    Protocol_Completed = mblnProtocolCmplt
End Property

Public Property Let Rx_Chlorhexidine(ByVal blnValue As Boolean)
    mblnRxChlorhexidine = blnValue
    UpdateDirtyFlag
End Property

Public Property Get Rx_Chlorhexidine() As Boolean
    Rx_Chlorhexidine = mblnRxChlorhexidine
End Property

Public Property Let Rx_OTC(ByVal blnValue As Boolean)
    mblnRxOTC = blnValue
    UpdateDirtyFlag
End Property

Public Property Get Rx_OTC() As Boolean
    Rx_OTC = mblnRxOTC
End Property

Public Property Let Rx_Dentifrice(ByVal blnValue As Boolean)
    mblnRxDenifrice = blnValue
    UpdateDirtyFlag
End Property


Public Property Get Rx_Dentifrice() As Boolean


    Rx_Dentifrice = mblnRxDenifrice
End Property



Public Property Let Fluoride_Complete(ByVal blnValue As Boolean)
    mblnFluorideCmplt = blnValue
    UpdateDirtyFlag
End Property


Public Property Get Fluoride_Complete() As Boolean
    Fluoride_Complete = mblnFluorideCmplt
End Property



Public Property Let Additional_Fluoride_Tx(ByVal blnValue As Boolean)
    mblnAddFluoride = blnValue
    UpdateDirtyFlag
End Property


Public Property Get Additional_Fluoride_Tx() As Boolean


    Additional_Fluoride_Tx = mblnAddFluoride
End Property



Public Property Let Sealants_Considered(ByVal blnValue As Boolean)
    mblnSealants = blnValue
    UpdateDirtyFlag
End Property


Public Property Get Sealants_Considered() As Boolean


    Sealants_Considered = mblnSealants
End Property



Public Property Let Caries_Eliminated(ByVal blnValue As Boolean)
    mblnCariesEliminated = blnValue
    UpdateDirtyFlag
End Property


Public Property Get Caries_Eliminated() As Boolean


    Caries_Eliminated = mblnCariesEliminated
End Property

Public Property Let Xylitol_Chewing_Gum(ByVal blnValue As Boolean)
    mblnGum = blnValue
    UpdateDirtyFlag
End Property

Public Property Get Xylitol_Chewing_Gum() As Boolean
    Xylitol_Chewing_Gum = mblnGum
End Property



'Public Property Let Written_Instructions(ByVal blnValue As Boolean)
'    mblnWrittenInstructs = blnValue
'    UpdateDirtyFlag
'End Property
'
'
'Public Property Get Written_Instructions() As Boolean
'
'
'    Written_Instructions = mblnWrittenInstructs
'End Property



Public Property Let Recommend_Diet_Mods(ByVal blnValue As Boolean)
    mblnDietMods = blnValue
    UpdateDirtyFlag
End Property


Public Property Get Recommend_Diet_Mods() As Boolean


    Recommend_Diet_Mods = mblnDietMods
End Property



Public Property Let Identify_Sucrose_Sources(ByVal blnValue As Boolean)
    mblnSucroseSrcs = blnValue
    UpdateDirtyFlag
End Property

Public Property Get Identify_Sucrose_Sources() As Boolean

    Identify_Sucrose_Sources = mblnSucroseSrcs

End Property



Public Property Let Exlain_Etiology_Caries(ByVal blnValue As Boolean)
    mblnExplainEtiology = blnValue
    UpdateDirtyFlag
End Property


Public Property Get Exlain_Etiology_Caries() As Boolean


    Exlain_Etiology_Caries = mblnExplainEtiology
End Property



Public Property Let Dental_Nutrional_Counseling(ByVal blnValue As Boolean)
    mblnDentalCounseling = blnValue
    UpdateDirtyFlag
End Property


Public Property Get Dental_Nutrional_Counseling() As Boolean
    Dental_Nutrional_Counseling = mblnDentalCounseling
End Property

Private Sub Class_Initialize()
    mblnHasData = False
End Sub

Private Property Get clsISection_DentalEncounterID() As Long
    clsISection_DentalEncounterID = mlngDentalEncounterId
End Property

Private Function clsISection_GetNote() As String
    Dim strNote As String
    Dim strTemp As String
    
    strNote = Empty
    strTemp = Empty
    If mblnDentalCounseling Then
        strNote = strNote & " Nutritional counseling performed"
    End If
        
    If mblnExplainEtiology And strNote <> "" Then
        strNote = strNote & ", etiology of caries as Strep " & _
            "Mutans infection explained; consider giving written " & _
            "instructions to educate family members and reinforce education"
    ElseIf mblnExplainEtiology Then
        strNote = strNote & " etiology of caries as Strep " & _
            "Mutans infection explained; consider giving written " & _
            "instructions to educate family members and reinforce education"
    End If
    
    If mblnSucroseSrcs And strNote <> "" Then
        strNote = strNote & ", Patient's sucrose sources identified"
    ElseIf mblnSucroseSrcs Then
        strNote = strNote & "Patient's sucrose sources identified"
    End If
    
    If mblnDietMods And strNote <> "" Then
        strNote = strNote & ", Diet modifications recommended"
    ElseIf mblnDietMods Then
        strNote = strNote & "Diet modifications recommended"
    End If
            
    If mblnGum And strNote <> "" Then
        strNote = strNote & ", Xylitol chewing gum recommended, Chew 2 pieces " & _
            "for 5 minutes three times a day"
    End If
    
    If strNote <> "" Then
        strNote = strNote & "." & vbCrLf
    End If
    
    If mstrNutritionCmnts <> "" Then
        strNote = strNote & "Nutrition Comments: " & mstrNutritionCmnts & "; "
    End If

    If mblnOralHygInstruct Then
        strNote = strNote & "Oral hygiene instructions given; "
    End If
    
    If mblnCariesEliminated Then
        strNote = strNote & vbTab & "Caries elimination completed"
    End If
    
    If mblnSealants Then
        strNote = strNote & ", sealants considered." & vbCrLf
    Else
        strNote = strNote & "." & vbCrLf
    End If
    
    If mblnAddFluoride Then
        strNote = strNote & vbTab & "Fluoride series complete. " & vbCrLf
    ElseIf mblnFluorideCmplt Then
        strNote = strNote & vbTab & "Return to clinic for additional Fluoride Treatment. " & vbCrLf
    End If

    If mblnRxDenifrice Then
        strNote = strNote & vbTab & "Rx for 1.1% neutral NaFl dentifrice/gel." & vbCrLf
    End If
    
    If mblnRxOTC Then
        strNote = strNote & vbTab & "Recommend use of NaFl 0.05% OTC rinse." & vbCrLf
    End If
    
    If mblnRxChlorhexidine Then
        If mlngProtocolState = 0 Then
            strNote = strNote & vbTab & "Rx written for Chlorhexidine 0.12% rinse - Once a week." & vbCrLf
        Else
            strNote = strNote & vbTab & "Rx written for Chlorhexidine 0.12% rinse - Twice a day " & vbCrLf
        End If
    End If
    
    If mblnProtocolCmplt Then
        If mlngProtocolState = 0 Then
            strNote = strNote & vbTab & "Moderate Protocol Completed Reassess at Next Periodic Exam." & vbCrLf
        Else
            strNote = strNote & vbTab & "High Risk Protocol Completed - 3 Month Recall." & vbCrLf
        End If
    
    
    End If
    If strNote <> "" Then
        If mlngProtocolState = crpCS_Moderate Then
            strNote = "Moderate Risk Protocol" & vbCrLf & strNote
        ElseIf mlngProtocolState = crpCS_High Then
            strNote = "High Risk Protocol" & vbCrLf & strNote
        End If
    End If
    clsISection_GetNote = strNote
End Function
Private Function clsISection_GetRTFNote() As String
    Dim strNote As String
    Dim strTemp As String
    Dim blnWriteNote As Boolean
    
    strNote = Empty
    strTemp = Empty
    blnWriteNote = False
    
    If mblnDentalCounseling Then
        strNote = strNote & " Nutritional counseling performed"
        blnWriteNote = True
    End If
        
    If mblnExplainEtiology And strNote <> "" Then
        blnWriteNote = True
        strNote = strNote & ", etiology of caries as Strep " & _
            "Mutans infection explained; consider giving written " & _
            "instructions to educate family members and reinforce education"
    ElseIf mblnExplainEtiology Then
        blnWriteNote = True
        strNote = strNote & " etiology of caries as Strep " & _
            "Mutans infection explained; consider giving written " & _
            "instructions to educate family members and reinforce education"
    End If
    
    If mblnSucroseSrcs And strNote <> "" Then
        blnWriteNote = True
        strNote = strNote & ", Patient's sucrose sources identified"
    ElseIf mblnSucroseSrcs Then
        blnWriteNote = True
        strNote = strNote & "Patient's sucrose sources identified"
    End If
    
    If mblnDietMods And strNote <> "" Then
        blnWriteNote = True
        strNote = strNote & ", Diet modifications recommended"
    ElseIf mblnDietMods Then
        blnWriteNote = True
        strNote = strNote & "Diet modifications recommended"
    End If
            
    If mblnGum And strNote <> "" Then
        blnWriteNote = True
        strNote = strNote & ", Xylitol chewing gum recommended, Chew 2 pieces " & _
            "for 5 minutes three times a day"
    ElseIf mblnGum Then
        blnWriteNote = True
        strNote = strNote & "Xylitol chewing gum recommended, Chew 2 pieces for 5 minutes three times a day"
    End If

    
    If strNote <> "" Then
        strNote = strNote & "; "
    End If
    
    If mstrNutritionCmnts <> "" Then
        blnWriteNote = True
        strNote = strNote & "Dental Nutritional Counseling Comments: " & mstrNutritionCmnts & "; "
    End If
    
    If mblnOralHygInstruct Then
        blnWriteNote = True
        strNote = strNote & "Oral hygiene instructions given; "
    End If
    
    If mblnCariesEliminated Then
        strNote = strNote & "Caries elimination completed"
    End If
    
    If mblnSealants And mblnCariesEliminated Then
        strNote = strNote & ", sealants considered; "
    ElseIf mblnSealants Then
        blnWriteNote = True
        strNote = strNote & "Sealants considered; "
    ElseIf strNote <> "" And mblnCariesEliminated Then
        strNote = strNote & "; "
    End If
    
    If mblnFluorideCmplt Then
        blnWriteNote = True
        strNote = strNote & "Fluoride series complete; "
    ElseIf mblnAddFluoride Then
        blnWriteNote = True
        strNote = strNote & "Return to clinic for additional Fluoride Treatment; "
    End If

    If mblnRxDenifrice Then
        blnWriteNote = True
        strNote = strNote & "Rx for 1.1% neutral NaFl dentifrice/gel; "
    End If
    
    If mblnRxOTC Then
        blnWriteNote = True
        strNote = strNote & "Recommend use of NaFl 0.05% OTC rinse; "
    End If
    
    If mblnRxChlorhexidine Then
        If mlngFormState = crpPFS_Moderate Then
            strNote = strNote & "Rx written for Chlorhexidine 0.12% rinse - Once a week; "
        Else
            strNote = strNote & "Rx written for Chlorhexidine 0.12% rinse - Twice a day; "
        End If
        blnWriteNote = True
    End If
    
    If mblnProtocolCmplt Then
        If mlngFormState = crpPFS_Moderate Then
            strNote = strNote & "Moderate Protocol Completed Reassess at Next Periodic Exam; "
        Else
            strNote = strNote & "High Risk Protocol Completed - 3 Month Recall; "
        End If
    
    
    End If
    If blnWriteNote Then
        If mlngFormState = crpPFS_Moderate Then
            strNote = crtfStartControlName & "Moderate Risk Protocol" & _
                crtfEndControlName & " " & strNote
        Else
            strNote = crtfStartControlName & "High Risk Protocol " & _
                crtfEndControlName & " " & strNote
        End If
    End If
    clsISection_GetRTFNote = strNote
End Function
Private Function clsISection_LoadSection(ByVal PatientID As Long, ByVal DentalEncounterID As Long) As Boolean
    Dim blnReturn As Boolean
    Dim objRS As ADODB.Recordset
    Dim blnCarryForward As Boolean
    
    On Error GoTo ErrHandler
    
    mlngPatUnitId = PatientID
    mlngDentalEncounterId = DentalEncounterID
    
    'Check the Previous Dental Encounter Id, if it doesn't  match with our current
    'dental encounter id, then we've started a new encounter. Load the previous
    'encounter, and check to see if it was completed. If it was, then we start
    'a fresh High Risk Recall Session. Otherwise, we load any previous data from
    'the last dental encounter.
    
    If (mlngPrevDentEncounterId <> mlngDentalEncounterId) And _
        (mlngPrevDentEncounterId > 0) Then
        
        'Load info from the previous encounter
        
        Set objRS = DB_Load_Protocol_Info(mlngPrevDentEncounterId)
        
        'Make sure that we have info from the previous encounter, if we don't
        'load the current encounter.
        If objRS.EOF Then
            Set objRS = DB_Load_Protocol_Info(mlngDentalEncounterId)
        Else
            If CBool(objRS.Fields("PROTOCOL_CMPLT")) Then
            'The Previous Caries Protocol was completed. No need to
            'carry the previous information forward. Load current
            'encounter
                Set objRS = DB_Load_Protocol_Info(mlngDentalEncounterId)
            Else
                'We're carrying forward info from a previous encounter
                blnCarryForward = True
            End If
        End If
    Else
        Set objRS = DB_Load_Protocol_Info(mlngDentalEncounterId)
    End If
   
    With objRS
        If .EOF Then
            mlngId = 0
            mlngDirtyFlag = dfAdd
            Reset_Values
        Else
            If blnCarryForward Then
                mlngDirtyFlag = dfAdd
            Else
                mlngDirtyFlag = dfUpdate
            End If
            mlngId = Val(.Fields("CRP_ID") & "")
            mblnDentalCounseling = CBool(Val(.Fields("DENTAL_NUTRIT_COUNSEL") & ""))
            mblnExplainEtiology = CBool(Val(.Fields("EXPLAIN_ETIOLOTY") & ""))
            mblnSucroseSrcs = CBool(Val(.Fields("IDENT_SUCROSE_SRC") & ""))
            mblnDietMods = CBool(Val(.Fields("REC_DIET_MODS") & ""))
'            mblnWrittenInstructs = CBool(.Fields("WRITTEN_INSTRUCTS"))
            mblnGum = CBool(Val(.Fields("GUM") & ""))
            mblnCariesEliminated = CBool(Val(.Fields("CARIES_ELIMIN") & ""))
            mblnSealants = CBool(Val(.Fields("SEALANTS") & ""))
            mblnAddFluoride = CBool(Val(.Fields("ADD_FLUORIDE_REQ") & ""))
            mblnFluorideCmplt = CBool(Val(.Fields("ALL_FLUORIDE_CMPLT") & ""))
            mblnRxDenifrice = CBool(Val(.Fields("Rx_Dentifrice") & ""))
            mblnRxOTC = CBool(Val(.Fields("Rx_OTC") & ""))
            mblnRxChlorhexidine = CBool(Val(.Fields("RX_REMINERAL") & ""))
            mblnProtocolCmplt = CBool(Val(.Fields("PROTOCOL_CMPLT") & ""))
            mstrNutritionCmnts = (.Fields("NUTRITION_CMNTS")) & ""
            mblnOralHygInstruct = CBool(Val(.Fields("ORAL_HYGIENE_INSTRUCT") & ""))
            mlngWLDentCounseling = Val(.Fields("WLID_DENT_NUTR_COUNSEL") & "")
            mlngWLOralHygiene = Val(.Fields("WLID_ORAL_HYGIENE_INSTR") & "")
            
            If blnCarryForward Then
                If mlngPrevWLDentCounsel <= 0 Then
                    mlngPrevWLDentCounsel = mlngWLDentCounseling
                End If
                
                If mlngPrevWLOralHygiene <= 0 Then
                    mlngPrevWLOralHygiene = mlngWLOralHygiene
                End If
            Else
                mlngPrevWLDentCounsel = Val(.Fields("WL_PREV_DENT_NUTR_COUNSEL") & "")
                mlngPrevWLOralHygiene = Val(.Fields("WL_PREV_ORAL_HYGIENE") & "")
            End If
        End If
    End With
    
FuncExit:
    clsISection_LoadSection = blnReturn
    mblnHasData = blnReturn
    Set objRS = Nothing
    Exit Function
ErrHandler:
   blnReturn = False
   goCWShared.ShowVBError Err.Number, Err.Description, const_ModuleName & ".clsISection_LoadSection()", "", vbCritical
   GoTo FuncExit
End Function

Private Property Get clsISection_PatientID() As Long
    clsISection_PatientID = mlngPatUnitId
End Property

Private Function clsISection_SaveSection() As Boolean
    Dim strSql As String
    Dim blnReturn As Boolean
    
    On Error GoTo ErrHandler
    
    If mlngDirtyFlag = dfNoAction Then
        blnReturn = True
        GoTo FuncExit
    Else
        UpdateWorkLoadCodes
    End If
    
    If mlngId = 0 Then mlngDirtyFlag = dfAdd
    
    strSql = Empty
    Select Case mlngDirtyFlag
        Case dfNoAction
            blnReturn = True
            GoTo FuncExit
        Case dfAdd
            mlngId = goconn.GetNextID(esTEMPL_CARIES_RISK_PROTOCOL_SEQ)
'            strSql = "SELECT TEMPL_CARIES_RISK_PROTOCOL_Seq.NEXTVAL From Dual"
'            Set objRS = goconn.OpenRecordset(strSql)
'            mlngId = objRS.Fields("NEXTVAL")

            strSql = "INSERT INTO TEMPL_CARIES_RISK_PROTOCOL " & _
                     "(CRP_ID, DENTAL_ENCOUNTER_ID, PROTOCOL_CMPLT, " & _
                        "EXPLAIN_ETIOLOTY, " & _
                        "IDENT_SUCROSE_SRC, REC_DIET_MODS, " & _
                        "GUM, DENTAL_NUTRIT_COUNSEL, " & _
                        "ORAL_HYGIENE_INSTRUCT, SEALANTS, CARIES_ELIMIN, " & _
                        "ADD_FLUORIDE_REQ, ALL_FLUORIDE_CMPLT, " & _
                        "Rx_Dentifrice , Rx_OTC, RX_REMINERAL, Nutrition_Cmnts, " & _
                        "WLID_DENT_NUTR_COUNSEL, WLID_ORAL_HYGIENE_INSTR,  " & _
                        "WL_PREV_DENT_NUTR_COUNSEL, WL_PREV_ORAL_HYGIENE) " & _
                     "VALUES (" & mlngId & ", " & mlngDentalEncounterId & ", " & _
                    CInt(mblnProtocolCmplt) & ", " & CInt(mblnExplainEtiology) & ", " & _
                    CInt(mblnSucroseSrcs) & ", " & CInt(mblnDietMods) & ", " & _
                    CInt(mblnGum) & ", " & _
                    CInt(mblnDentalCounseling) & ", " & CInt(mblnOralHygInstruct) & ", " & _
                    CInt(mblnSealants) & ", " & CInt(mblnCariesEliminated) & ", " & _
                    CInt(mblnAddFluoride) & ", " & CInt(mblnFluorideCmplt) & ", " & _
                    CInt(mblnRxDenifrice) & ", " & CInt(mblnRxOTC) & ", " & _
                    CInt(mblnRxChlorhexidine) & ", '" & Scrub_Apostrophe(mstrNutritionCmnts) & "', " & _
                    mlngWLDentCounseling & ", " & mlngWLOralHygiene & ", " & _
                    mlngPrevWLDentCounsel & ", " & mlngPrevWLOralHygiene & ")"

        Case dfUpdate
            strSql = "UPDATE TEMPL_CARIES_RISK_PROTOCOL " & _
                     " SET PROTOCOL_CMPLT = " & CInt(mblnProtocolCmplt) & ", EXPLAIN_ETIOLOTY = " & _
                     CInt(mblnExplainEtiology) & ", IDENT_SUCROSE_SRC = " & _
                     CInt(mblnSucroseSrcs) & ", REC_DIET_MODS = " & CInt(mblnDietMods) & _
                     ", GUM = " & _
                     CInt(mblnGum) & ", DENTAL_NUTRIT_COUNSEL = " & CInt(mblnDentalCounseling) & ", " & _
                     "ORAL_HYGIENE_INSTRUCT = " & CInt(mblnOralHygInstruct) & ", SEALANTS = " & _
                     CInt(mblnSealants) & ", CARIES_ELIMIN = " & CInt(mblnCariesEliminated) & ", " & _
                     "ADD_FLUORIDE_REQ = " & CInt(mblnAddFluoride) & ", ALL_FLUORIDE_CMPLT = " & _
                     CInt(mblnFluorideCmplt) & ", Rx_Dentifrice = " & CInt(mblnRxDenifrice) & ", " & _
                     "Rx_OTC = " & CInt(mblnRxOTC) & ", RX_REMINERAL = " & CInt(mblnRxChlorhexidine) & _
                     " WHERE CRP_ID = " & mlngId
        Case dfdelete
            strSql = "DELETE FROM TEMPL_CARIES_RISK_PROTOCOL WHERE " & _
                     " CRP_ID = " & mlngId
    End Select
    
    If strSql <> "" Then
        goconn.ExecuteSQL (strSql)
    End If
    mlngDirtyFlag = dfNoAction
    
    'If there's a collection, loop through it here.
FuncExit:
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

'**************************************************************
'Function: CheckForCariesElimination
'Author: KDunne
'Date:1/1/2002
'Purpose: Checks for Caries elimination for the entire mouth.
'         Checks for any unrestored caries. If there are, Caries
'         Elimination is false.
'**************************************************************
Public Function CheckForCariesElimination() As Boolean

    Dim objGroups As CHCSII_Dental_DXTX.colGroups
    Dim objGroup As CHCSII_Dental_DXTX.clsGroup
    Dim objDxs As CHCSII_Dental_DXTX.colDxs
    Dim objDx As CHCSII_Dental_DXTX.clsDx
    Dim objTeeth As CHCSII_Dental_DXTX.colTeeth
    Dim objTooth As CHCSII_Dental_DXTX.clsTooth
    Dim objSSs As CHCSII_Dental_DXTX.colSurfaceSets
    Dim objSS As CHCSII_Dental_DXTX.clsSurfaceSet
    Dim intIndex As Integer
    Dim intCount As Integer
    Dim blnReturn As Boolean
    Dim blnFoundCaries As Boolean
    
    On Error GoTo ErrHandler:
    
    If mobjDataLayer Is Nothing Then
        blnReturn = False
        Err.Raise 1000, , "Data layer object was not set"
    End If
    blnReturn = True
    
    Set objGroups = mobjDataLayer.Groups
    
    If objGroups Is Nothing Then
        Exit Function
    End If
    
    For Each objGroup In objGroups
        'In each Group, we have a collection of DX and TX's.
        Set objDxs = objGroup.Dxs
        
        'Loop through the Diagnosis'
        For Each objDx In objDxs
            'SCR 42724 RWalker 10/20/2003
            'Also check for codes:
            'cICDCodeK02_0 ("Caries, with minimal extension into dentin") 1418
            'cICDCodeK02_13 ("Caries, moderate to advanced into dentin ") 1420
            If (objDx.DxICDCodeID = cICDCodeK02_2 Or objDx.DxICDCodeID = cICDCodeK02_12 Or _
                objDx.DxICDCodeID = cICDCodeK02_14 Or objDx.DxICDCodeID = cICDCodeK02_3 Or _
                objDx.DxICDCodeID = cICDCodeK02_4 Or objDx.DxICDCodeID = cICDCodeK02_5 Or _
                objDx.DxICDCodeID = cICDCodeK02_51 Or _
                objDx.DxICDCodeID = cICDCodeK02_0 Or objDx.DxICDCodeID = cICDCodeK02_13) _
                And _
                (objDx.StatusID = dxActiveNew Or objDx.StatusID = dxActiveOngoing) _
                And _
                (objDx.DRC = 2 Or objDx.DRC = 3) Then '10-28-03 SF Added the DRC logic
                blnFoundCaries = True
                Exit For
            End If
        Next
        
        If blnFoundCaries Then Exit For
    Next
    
    blnReturn = Not blnFoundCaries
    
FuncExit:
    CheckForCariesElimination = blnReturn
    mblnCariesEliminated = blnReturn
    Exit Function
ErrHandler:
    blnReturn = False
    goCWShared.ShowVBError Err.Number, Err.Description, "clsExamEndoTest.CheckForCariesElimination()", "", vbCritical
    GoTo FuncExit
    
End Function

'******************************************************************************
'Function: DB_Load_Recall_Info
'Author: KDunne
'Date: 12/27/2001
'Purpose: Loads data from the TEMPL_CARIES_RISK_RECALL_VISIT table
'*******************************************************************************
Private Function DB_Load_Protocol_Info(ByVal lngDentalEncounterId As Long) As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo ErrHandler
    
    strSql = "Select CRP_ID, DENTAL_ENCOUNTER_ID, PROTOCOL_CMPLT, " & _
        "EXPLAIN_ETIOLOTY, NUTRITION_CMNTS, " & _
        "IDENT_SUCROSE_SRC, REC_DIET_MODS, " & _
        "GUM, DENTAL_NUTRIT_COUNSEL, " & _
        "ORAL_HYGIENE_INSTRUCT, SEALANTS, CARIES_ELIMIN, " & _
        "ADD_FLUORIDE_REQ, ALL_FLUORIDE_CMPLT, " & _
        "Rx_Dentifrice , Rx_OTC, RX_REMINERAL, " & _
        "WLID_DENT_NUTR_COUNSEL, WLID_ORAL_HYGIENE_INSTR, " & _
        "WL_PREV_DENT_NUTR_COUNSEL, WL_PREV_ORAL_HYGIENE " & _
        "From TEMPL_CARIES_RISK_PROTOCOL " & _
        "Where DENTAL_ENCOUNTER_ID = " & lngDentalEncounterId
        
    Set DB_Load_Protocol_Info = goconn.OpenRecordset(strSql)
    
    Exit Function
ErrHandler:
    goCWShared.ShowVBError Err.Number, Err.Description, const_ModuleName & ".DB_Load_Protocol_Info()", "", vbCritical
    Set DB_Load_Protocol_Info = Nothing
End Function

'*****************************************************************************
'Function: UpdateWorkLoadCodes
'Author:KDunne
'Date:1/2/2001
'Purpose: Updates the workload codes for the form.
'*****************************************************************************
Private Function UpdateWorkLoadCodes() As Boolean
    Dim lngDirtyFlag As DirtyFlagEnum
    Dim lngTempWLId As Long
    
    If mblnDentalCounseling And (mblnWLDentCounsel = False) Then
        'Dental Nutrional Conseling was completed.
        
        'User received workload credit in a previous encounter. They
        'can't get it twice.
        If mlngPrevWLDentCounsel > 0 Then GoTo CheckOralHygiene
        
        If mlngWLDentCounseling > 0 Then
            lngDirtyFlag = dfUpdate
        Else
            'For Dental Counseling, first we check to see if they've received
            'workload credit for this item elsewhere this encounter.
            lngTempWLId = FindWorkloadId(cDentalCounselingWLCode, mlngDentalEncounterId)
            
            If lngTempWLId > 0 Then
                lngDirtyFlag = dfUpdate
                mlngWLDentCounseling = lngTempWLId
            Else
                lngDirtyFlag = dfAdd
            End If
        End If
        
        If UpdateWorkload(mlngWLDentCounseling, lngDirtyFlag, mlngDentalEncounterId, cDentalCounselingWLCode, 1) = False Then
            UpdateWorkLoadCodes = False
            Exit Function
        End If
    ElseIf mblnDentalCounseling = False And mblnWLDentCounsel = False Then
        lngTempWLId = FindWorkloadId(cDentalCounselingWLCode, mlngDentalEncounterId)
        Call UpdateWorkload(lngTempWLId, dfdelete, mlngDentalEncounterId, cDentalCounselingWLCode, 1)
    End If
    
CheckOralHygiene:

    If mblnOralHygInstruct And (mblnWLOralHygiene = False) Then
        'Oral Hygiene Instructions
        
        'User received workload credit for Oral Hygiene in a previous
        'encounter. They can't get it twice.
        If mlngPrevWLOralHygiene > 0 Then Exit Function
        
        If mlngWLOralHygiene > 0 Then
            lngDirtyFlag = dfUpdate
        Else
            lngDirtyFlag = dfAdd
        End If
        If UpdateWorkload(mlngWLOralHygiene, lngDirtyFlag, mlngDentalEncounterId, cOralHygieneWLCode, 1) = False Then
            UpdateWorkLoadCodes = False
            Exit Function
        End If
    
    ElseIf mlngWLOralHygiene > 0 And mblnWLOralHygiene = False Then
        Call UpdateWorkload(mlngWLOralHygiene, dfdelete, mlngDentalEncounterId, cOralHygieneWLCode, 1)
    End If
End Function

