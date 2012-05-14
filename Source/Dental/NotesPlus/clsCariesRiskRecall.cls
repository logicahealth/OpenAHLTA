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
Attribute VB_Name = "clsCRRecall"
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

Private mblnDiet_Counseling As Boolean
Private mblnStrep_Mutans_Test As Boolean
Private mblnStrep_Mutans_Low As Boolean
Private mblnStrep_Mutans_High As Boolean
Private mblnSalivary_Test As Boolean
Private mblnCompleted As Boolean
Private mlngFluorideTx As CariesRiskRecall_FluroideTreatmentsEnum
Private mlngStrepMutansLevel As CariesRiskRecall_StrepMutansLevelsEnum
Private mlngWLFluoride As Long
Private mlngWLStrepMutansTest As Long
Private mlngPrevWLFluoride As Long
Private mlngPrevWLSMTest As Long
Private mblnHasData As Boolean

Public Enum CariesRiskRecall_FluroideTreatmentsEnum
    crrFT_Not_Addressed = 0
    crrFT_NaFlVarnish = 1
    crrFT_APF = 2
    crrFT_NaFlNeutral = 3
    crrFT_SnFl = 4
End Enum

Public Enum CariesRiskRecall_StrepMutansLevelsEnum
    crrSML_Not_Addressed = 0
    crrSML_Low = 1
    crrSML_High = 2
End Enum

Private Const const_ModuleName = "clsCRRecall" '*** Insert Module Name Here
Private Const cFluoride5NAFLWLCode = 472
Private Const cFluoride123APFWLCode = 474
Private Const cFluoride2NaFlWLCode = 473
Private Const cFluoride8SnFlWLCode = 471
Private Const cStrepMutansWLCode = 41

Public Property Get HasData() As Boolean
    HasData = mblnHasData
End Property
'*****************************************************************************
'Sub:Reset_Values
'Author:KDunne
'Date:1/4/2002
'Purpose: Resets the object to a "fresh" state. Retains the id. Used when Protocol
'         is completed.
'*****************************************************************************
Public Sub Reset_Values()
    mblnDiet_Counseling = False
    mblnStrep_Mutans_Test = False
    mblnStrep_Mutans_Low = False
    mblnStrep_Mutans_High = False
    mblnSalivary_Test = False
    mblnCompleted = False
    mlngFluorideTx = crrFT_Not_Addressed
    mblnHasData = False
    
    If mlngId = 0 Then mlngDirtyFlag = dfAdd
End Sub
Public Property Get Prev_WorkLoad_Strep_Mutans_Test() As Long
    Prev_WorkLoad_Strep_Mutans_Test = mlngPrevWLSMTest
End Property
Public Property Get WorkLoad_Fluoride_Application() As Long
    WorkLoad_Fluoride_Application = mlngWLFluoride
End Property

Public Property Let WorkLoad_Fluoride_Application(ByVal lngValue As Long)
    mlngWLFluoride = lngValue
End Property

Public Property Get WorkLoad_Strep_Mutans_Test() As Long
    WorkLoad_Strep_Mutans_Test = mlngWLStrepMutansTest
End Property

Public Property Let WorkLoad_Strep_Mutans_Test(ByVal lngValue As Long)
    mlngWLStrepMutansTest = lngValue
End Property

Public Property Get Strep_Mutans_Level() As CariesRiskRecall_StrepMutansLevelsEnum
    Strep_Mutans_Level = mlngStrepMutansLevel
End Property

Public Property Let Strep_Mutans_Level(ByVal lngValue As CariesRiskRecall_StrepMutansLevelsEnum)
    mlngStrepMutansLevel = lngValue
    UpdateDirtyFlag
End Property

Public Property Get Previous_Dental_Encounter_ID() As Long
    Previous_Dental_Encounter_ID = mlngPrevDentEncounterId
End Property
Public Property Let Previous_Dental_Encounter_ID(ByVal lngValue As Long)
    mlngPrevDentEncounterId = lngValue
End Property

Public Property Get Flouride_Treatment() As CariesRiskRecall_FluroideTreatmentsEnum
    Flouride_Treatment = mlngFluorideTx
End Property
Public Property Let Flouride_Treatment(ByVal lngValue As CariesRiskRecall_FluroideTreatmentsEnum)
    mlngFluorideTx = lngValue
    UpdateDirtyFlag
End Property
Public Property Let Completed(ByVal blnValue As Boolean)
    mblnCompleted = blnValue
    UpdateDirtyFlag
End Property


Public Property Get Completed() As Boolean
    Completed = mblnCompleted
End Property

Public Property Let Salivary_Test(ByVal blnValue As Boolean)
    mblnSalivary_Test = blnValue
    UpdateDirtyFlag
End Property


Public Property Get Salivary_Test() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Salivary_Test
    Salivary_Test = mblnSalivary_Test
End Property

Public Property Let Strep_Mutans_Test(ByVal blnValue As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Strep_Mutans_Test = 5
    mblnStrep_Mutans_Test = blnValue
    UpdateDirtyFlag
End Property


Public Property Get Strep_Mutans_Test() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Strep_Mutans_Test
    Strep_Mutans_Test = mblnStrep_Mutans_Test
End Property



Public Property Let Diet_Counseling(ByVal blnValue As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Diet_Counseling = 5
    mblnDiet_Counseling = blnValue
    UpdateDirtyFlag
End Property


Public Property Get Diet_Counseling() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Diet_Counseling
    Diet_Counseling = mblnDiet_Counseling
End Property




Private Sub Class_Initialize()
    mlngDirtyFlag = dfAdd
End Sub

Private Property Get clsISection_DentalEncounterID() As Long
    clsISection_DentalEncounterID = mlngDentalEncounterId
End Property

Private Function clsISection_GetNote() As String
    Dim strNote As String
    
    strNote = Empty
    
    If mblnDiet_Counseling Then
        strNote = strNote & vbTab & "Reinforce diet counseling, home fluoride use, xylitol use, and oral hygiene instructions." & vbCrLf
    End If
    
    If mlngFluorideTx <> CariesRiskRecall_FluroideTreatmentsEnum.crrFT_Not_Addressed Then
        strNote = strNote & vbTab & GetFluorideTreatmentEnumDesc(mlngFluorideTx) & vbCrLf
    End If
    
    If mblnStrep_Mutans_Test Then
        If mlngStrepMutansLevel = CariesRiskRecall_StrepMutansLevelsEnum.crrSML_Low Then
            strNote = strNote & vbTab & "Strep Mutans testing performed levels are < " & _
                "100,000 CFU in High Risk Recall.  Change to Low Risk. " & vbCrLf
        Else
            strNote = strNote & vbTab & "Strep Mutans testing performed levels are > " & _
                "repeat office fluoride and chlorhexidine rinse protocols (1 oz " & _
                "twice a day x 14 days) Reexamine for undiscovered caries, Reevaluate " & _
                "diet, and consider reinfection by spouse in High Risk Recall. " & _
                "Recall in 3 months" & vbCrLf
        End If
    
    End If
    
    If mblnCompleted Then
        strNote = strNote & vbTab & "High Risk Protocol Recall Completed - Continue 3 Month Recall" & vbCrLf
    End If
    
    If strNote <> "" Then
        strNote = "High Risk Recall Visit" & vbCrLf & strNote
    End If
    
    clsISection_GetNote = strNote
End Function
Private Function clsISection_GetRTFNote() As String
    Dim strNote As String
    
    strNote = Empty
    
    If mblnDiet_Counseling Then
        strNote = strNote & "Reinforce diet counseling, home fluoride use, xylitol use, and oral hygiene instructions; "
    End If
    
    If mlngFluorideTx <> CariesRiskRecall_FluroideTreatmentsEnum.crrFT_Not_Addressed Then
        strNote = strNote & GetFluorideTreatmentEnumDesc(mlngFluorideTx) & "; "
    End If
    
    If mblnStrep_Mutans_Test Then
        If mlngStrepMutansLevel = CariesRiskRecall_StrepMutansLevelsEnum.crrSML_Low Then
            strNote = strNote & "Strep Mutans testing performed levels are < " & _
                "100,000 CFU in High Risk Recall.  Change to Low Risk; "
        Else
            strNote = strNote & "Strep Mutans testing performed levels are > 100,000.00 CFU, " & _
                "repeat office fluoride and chlorhexidine rinse protocols (1 oz " & _
                "twice a day x 14 days) Reexamine for undiscovered caries, Reevaluate " & _
                "diet, and consider reinfection by spouse in High Risk Recall. " & _
                "Recall in 3 months; "
        End If
    ElseIf mblnSalivary_Test Then
        strNote = strNote & "Salivary bacterial tests are unavailable, patients judged clinically " & _
            "to be at continued high risk should repeat chlorhexidine " & _
            "rinse protocol (1 oz twice a day x 14 days) at 3 month " & _
            "intervals until risk judged to be low or moderate; "
    End If
    
    If mblnCompleted Then
        strNote = strNote & "High Risk Protocol Recall Completed - Continue 3 Month Recall; "
    End If
    
    If strNote <> "" Then
        strNote = crtfStartControlName & "High Risk Recall Visit" & _
            crtfEndControlName & " " & strNote
    End If
    
    clsISection_GetRTFNote = strNote

End Function
Private Function clsISection_LoadSection(ByVal PatientID As Long, ByVal DentalEncounterID As Long) As Boolean
    Dim blnReturn As Boolean
    Dim objRS As ADODB.Recordset
    Dim blnNewEncounter As Boolean
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
        Set objRS = DB_Load_Recall_Info(mlngPrevDentEncounterId)
        
        'Make sure that we have info from the previous encounter, if we don't
        'load the current encounter.
        If objRS.EOF Then
            Set objRS = DB_Load_Recall_Info(mlngDentalEncounterId)
        Else
            If CBool(Val(objRS.Fields("PROTOCOL_CMPLT") & "")) Then
            'The Previous Caries Protocol was completed. No need to
            'carry the previous information forward. Load current
            'encounter
                Set objRS = DB_Load_Recall_Info(mlngDentalEncounterId)
            Else
                blnCarryForward = True
            End If
        End If
    Else
        Set objRS = DB_Load_Recall_Info(mlngDentalEncounterId)
    End If
    
    With objRS
        If .EOF Then
            mlngDirtyFlag = dfAdd
            Reset_Values
        Else
            If blnCarryForward Then
                mlngDirtyFlag = dfAdd
            Else
                mlngDirtyFlag = dfUpdate
            End If
            
            mlngId = Val(.Fields("CR_RECALL_ID") & "")
            mblnDiet_Counseling = CBool(Val(.Fields("DIET_COUNSELING") & ""))
            mblnStrep_Mutans_Test = CBool(Val(.Fields("STREP_MUTANS_TEST") & ""))
            mlngStrepMutansLevel = Val(.Fields("STREP_MUTANS_LEVEL") & "")
            mblnSalivary_Test = CBool(Val(.Fields("SB_TEST_UNAVAIL") & ""))
            mblnCompleted = CBool(Val(.Fields("PROTOCOL_CMPLT") & ""))
            mlngFluorideTx = Val(.Fields("OFFICE_FLUORIDE") & "")
            mlngStrepMutansLevel = Val(.Fields("STREP_MUTANS_LEVEL") & "")
            mlngWLFluoride = Val(.Fields("WLID_OFFICE_FLUORIDE") & "")
            mlngWLStrepMutansTest = Val(.Fields("WLID_STREP_MUTANS_TEST") & "")
            
            If blnCarryForward Then
                mlngPrevWLFluoride = mlngWLFluoride
                mlngPrevWLSMTest = mlngWLStrepMutansTest
            Else
                mlngPrevWLFluoride = Val(.Fields("WL_PREV_OFFICE_FLUORIDE") & "")
                mlngPrevWLSMTest = Val(.Fields("WL_PREV_STREP_MUTANS_TEST") & "")
            End If
        End If
    End With
    
    blnReturn = True
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

Public Sub Set_Dental_Encounter_Id(ByVal lngValue As Long)
    mlngDentalEncounterId = lngValue
End Sub
Private Property Get clsISection_PatientID() As Long
    clsISection_PatientID = mlngPatUnitId
End Property

Private Function clsISection_SaveSection() As Boolean
    Dim strSql As String
    Dim blnReturn As Boolean
    Dim objRS As ADODB.Recordset
    
    On Error GoTo ErrHandler
    
    strSql = Empty
    
    If mlngDirtyFlag = dfNoAction Then
        blnReturn = True
        GoTo FuncExit
    Else
        UpdateWorkLoadCodes
    End If
    If mlngId = 0 Then mlngDirtyFlag = dfAdd
    Select Case mlngDirtyFlag
        Case dfNoAction
            blnReturn = True
            GoTo FuncExit
        Case dfAdd
            mlngId = goconn.GetNextID(esTEMPL_CARIES_RISK_RECALL_V_SEQ)
'            strSql = "SELECT TEMPL_CARIES_RISK_RECALL_V_seq.NEXTVAL From Dual"
'            Set objRS = goconn.OpenRecordset(strSql)
'            mlngId = objRS.Fields("NEXTVAL")

            strSql = "INSERT INTO TEMPL_CARIES_RISK_RECALL_VISIT " & _
                        "(CR_RECALL_ID, DENTAL_ENCOUNTER_ID,  " & _
                        "DIET_COUNSELING, STREP_MUTANS_TEST, " & _
                        "STREP_MUTANS_LEVEL,  WLID_OFFICE_FLUORIDE, WLID_STREP_MUTANS_TEST, " & _
                        "WL_PREV_OFFICE_FLUORIDE,  WL_PREV_STREP_MUTANS_TEST, " & _
                        "OFFICE_FLUORIDE, SB_TEST_UNAVAIL, PROTOCOL_CMPLT) " & _
                     "VALUES (" & mlngId & ", " & mlngDentalEncounterId & ", " & _
                     CInt(mblnDiet_Counseling) & ", " & CInt(mblnStrep_Mutans_Test) & _
                     ", " & mlngStrepMutansLevel & ", " & mlngWLFluoride & ", " & _
                     mlngWLStrepMutansTest & ", " & mlngPrevWLFluoride & ", " & mlngPrevWLSMTest & _
                     ", " & mlngFluorideTx & ", " & CInt(mblnSalivary_Test) & _
                     ", " & CInt(mblnCompleted) & ")"
        Case dfUpdate
            strSql = "UPDATE TEMPL_CARIES_RISK_RECALL_VISIT " & _
                     " SET DIET_COUNSELING = " & CInt(mblnDiet_Counseling) & ", " & _
                     "STREP_MUTANS_TEST = " & CInt(mblnStrep_Mutans_Test) & ", " & _
                     "STREP_MUTANS_LEVEL = " & mlngStrepMutansLevel & ", " & _
                     "WLID_STREP_MUTANS_TEST = " & mlngWLStrepMutansTest & ", " & _
                     "WLID_OFFICE_FLUORIDE = " & mlngWLFluoride & ", " & _
                     "WL_PREV_OFFICE_FLUORIDE = " & mlngPrevWLFluoride & ", " & _
                     "WL_PREV_STREP_MUTANS_TEST = " & mlngPrevWLSMTest & ", " & _
                     "OFFICE_FLUORIDE = " & mlngFluorideTx & ", SB_TEST_UNAVAIL = " & _
                     CInt(mblnSalivary_Test) & ", PROTOCOL_CMPLT = " & CInt(mblnCompleted) & _
                     " WHERE CR_RECALL_ID = " & mlngId
        Case dfdelete
            strSql = "DELETE FROM TEMPL_CARIES_RISK_RECALL_VISIT WHERE " & _
                     " CR_RECALL_ID = " & mlngId
    End Select
    
    If strSql <> "" Then
        goconn.ExecuteSQL (strSql)
    End If
    
    mlngDirtyFlag = dfNoAction
    
    'If there's a collection, loop through it here.
FuncExit:
    Set objRS = Nothing
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

'**********************************************************************************************
'Function:
'Author: KDunne
'Date:12/21/2001
'Purpose: Returns the description for a given CariesRiskRecall_FluroideTreatmentsEnum  value
'**********************************************************************************************
Public Function GetFluorideTreatmentEnumDesc(ByVal lngValue As CariesRiskRecall_FluroideTreatmentsEnum) As String
    Dim strReturn As String
    
    Select Case lngValue
        Case CariesRiskRecall_FluroideTreatmentsEnum.crrFT_NaFlVarnish
            strReturn = "5% NaFl varnish"
        Case CariesRiskRecall_FluroideTreatmentsEnum.crrFT_APF
            strReturn = "1.23% APF"
        Case CariesRiskRecall_FluroideTreatmentsEnum.crrFT_NaFlNeutral
            strReturn = "2% neutral NaFl"
        Case CariesRiskRecall_FluroideTreatmentsEnum.crrFT_SnFl
            strReturn = "8% SnFl"
        Case CariesRiskRecall_FluroideTreatmentsEnum.crrFT_Not_Addressed
            strReturn = Empty
        Case Else
            strReturn = Empty
    End Select
    
    GetFluorideTreatmentEnumDesc = strReturn
    Exit Function
ErrHandler:
    GetFluorideTreatmentEnumDesc = Empty
    goCWShared.ShowVBError Err.Number, Err.Description, const_ModuleName & ".GetFluorideTreatmentEnumDesc()", "", vbCritical

End Function

'******************************************************************************
'Function: DB_Load_Recall_Info
'Author: KDunne
'Date: 12/27/2001
'Purpose: Loads data from the TEMPL_CARIES_RISK_RECALL_VISIT table
'*******************************************************************************
Private Function DB_Load_Recall_Info(ByVal lngDentalEncounterId As Long) As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo ErrHandler
    
    strSql = "SELECT CR_RECALL_ID, DENTAL_ENCOUNTER_ID, " & _
             "DIET_COUNSELING, STREP_MUTANS_TEST, " & _
             "OFFICE_FLUORIDE, SB_TEST_UNAVAIL, PROTOCOL_CMPLT, " & _
             "STREP_MUTANS_LEVEL, WLID_OFFICE_FLUORIDE, " & _
             "WLID_STREP_MUTANS_TEST, WL_PREV_OFFICE_FLUORIDE, " & _
             "WL_PREV_STREP_MUTANS_TEST " & _
             "FROM TEMPL_CARIES_RISK_RECALL_VISIT " & _
             "WHERE DENTAL_ENCOUNTER_ID = " & lngDentalEncounterId
        
    Set DB_Load_Recall_Info = goconn.OpenRecordset(strSql)
    
    Exit Function
ErrHandler:
    goCWShared.ShowVBError Err.Number, Err.Description, const_ModuleName & ".DB_Load_Recall_Info()", "", vbCritical
    Set DB_Load_Recall_Info = Nothing
End Function
'*****************************************************************************
'Function: GetFluorideWLCode
'Author:KDunne
'Date:1/2/2001
'Purpose: Retrieves the Proper Workload code for the Selected Fluoride treatment
'*****************************************************************************
Private Function GetFluorideWLCode() As Long
    Dim lngReturn As Long
    
    Select Case mlngFluorideTx
        Case CariesRiskRecall_FluroideTreatmentsEnum.crrFT_NaFlVarnish
            lngReturn = cFluoride5NAFLWLCode
        Case CariesRiskRecall_FluroideTreatmentsEnum.crrFT_APF
            lngReturn = cFluoride123APFWLCode
        Case CariesRiskRecall_FluroideTreatmentsEnum.crrFT_NaFlNeutral
            lngReturn = cFluoride2NaFlWLCode
        Case CariesRiskRecall_FluroideTreatmentsEnum.crrFT_SnFl
            lngReturn = cStrepMutansWLCode
    End Select
    
    GetFluorideWLCode = lngReturn
End Function
'*****************************************************************************
'Function: UpdateWorkLoadCodes
'Author:KDunne
'Date:1/2/2001
'Purpose: Updates the workload codes for the form.
'*****************************************************************************
Private Function UpdateWorkLoadCodes() As Boolean
    Dim lngDirtyFlag As DirtyFlagEnum
    
    If mlngFluorideTx = crrFT_Not_Addressed Then
        If mlngWLFluoride <> 0 Then
            'User previously had selected a Fluoride Tx, but has now unselected it.
            'Remove workload credit
            Call UpdateWorkload(mlngWLFluoride, dfdelete, mlngDentalEncounterId, 0, 1)
        End If
    Else
        If mlngPrevWLFluoride > 0 Then GoTo CheckStrepMutans
        
        If mlngWLFluoride > 0 Then
            lngDirtyFlag = dfUpdate
        Else
            lngDirtyFlag = dfAdd
        End If
        
        If UpdateWorkload(mlngWLFluoride, lngDirtyFlag, mlngDentalEncounterId, GetFluorideWLCode(), 1) = False Then
            UpdateWorkLoadCodes = False
            Exit Function
        End If
    End If
    
CheckStrepMutans:
    If mblnStrep_Mutans_Test Or mblnSalivary_Test Then
    
        If mlngPrevWLSMTest Then Exit Function
        
        If mlngWLStrepMutansTest > 0 Then
            lngDirtyFlag = dfUpdate
        Else
            lngDirtyFlag = dfAdd
        End If
        
        If UpdateWorkload(mlngWLStrepMutansTest, lngDirtyFlag, mlngDentalEncounterId, cStrepMutansWLCode, 1) = False Then
            UpdateWorkLoadCodes = False
            Exit Function
        End If
    Else
        If mlngWLStrepMutansTest <> 0 Then
            'User previously had selected a Strep Mutans test, but has now unselected it.
            'Remove workload credit
            Call UpdateWorkload(mlngWLStrepMutansTest, dfdelete, mlngDentalEncounterId, 0, 1)
        End If
    
    End If
    
End Function

