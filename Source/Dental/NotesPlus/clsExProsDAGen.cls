VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsExProsDAGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Implements clsISection

Private mlngDentalEncounterId As Long
Private mlngPatUnitId As Long
Private mlngDirtyFlag As DirtyFlagEnum

Private mlngId As Long
Private mstrExpectations As String
Private mstrComments As String
Private mlngPersonality As DentAssess_PersonalityEnum
Private mlngMandibularAss As DentAssess_ProviderAssessEnum
Private mlngMaxillaryAss As DentAssess_ProviderAssessEnum
Private mobjSOAP As clsExamSOAP

Private Const const_ModuleName = "clsExProsDAGen"

Public Enum DentAssess_PersonalityEnum
    daPer_Not_Addressed
    daPer_Philosophical
    daPer_Exacting
    daPer_Hysterical
    daPer_Indifferent
End Enum

Public Enum DentAssess_ProviderAssessEnum
    daAss_Not_Addressed
    daAss_Acceptable
    daAss_Unacceptable
End Enum


Public Property Get Dental_Assessment_Id() As Long
    Dental_Assessment_Id = mlngId
End Property
Public Property Let Dental_Assessment_Id(ByVal lngValue As Long)
    mlngId = lngValue
End Property
Public Property Get Expectations() As String
    Expectations = mstrExpectations
End Property
Public Property Let Expectations(ByVal strValue As String)
    mstrExpectations = strValue
    UpdateDirtyFlag
End Property

Public Property Get Comments() As String
    Comments = mstrComments
End Property
Public Property Let Comments(ByVal strValue As String)
    mstrComments = strValue
    UpdateDirtyFlag
End Property

Public Property Get Personality() As DentAssess_PersonalityEnum
    Personality = mlngPersonality
End Property

Public Property Let Personality(ByVal lngValue As DentAssess_PersonalityEnum)
    mlngPersonality = lngValue
    UpdateDirtyFlag
End Property

Public Property Get Current_Mandibular_Assessement() As DentAssess_ProviderAssessEnum
    Current_Mandibular_Assessement = mlngMandibularAss
End Property
Public Property Let Current_Mandibular_Assessement(ByVal lngValue As DentAssess_ProviderAssessEnum)
    mlngMandibularAss = lngValue
    UpdateDirtyFlag
End Property

Public Property Get Current_Maxillary_Assessement() As DentAssess_ProviderAssessEnum
    Current_Maxillary_Assessement = mlngMaxillaryAss
End Property
Public Property Let Current_Maxillary_Assessement(ByVal lngValue As DentAssess_ProviderAssessEnum)
    mlngMaxillaryAss = lngValue
    UpdateDirtyFlag
End Property

Public Property Get SOAP() As clsExamSOAP
    Set SOAP = mobjSOAP
End Property
Public Property Set SOAP(ByVal objSOAP As clsExamSOAP)
    Set mobjSOAP = objSOAP
End Property



Private Sub Class_Terminate()
    Set mobjSOAP = Nothing
End Sub

Private Property Get clsISection_DentalEncounterID() As Long
    clsISection_DentalEncounterID = mlngDentalEncounterId
End Property

Private Function clsISection_GetNote() As String
    Dim strNote As String
    
    strNote = Empty
          
    If Trim(mstrExpectations) <> "" Then
        strNote = strNote & "Expectectations: " & mstrExpectations & "; "
    End If
    
    If mlngPersonality <> DentAssess_PersonalityEnum.daPer_Not_Addressed Then
        strNote = strNote & "Personality: " & _
            GetPersonalityEnumDesc(mlngPersonality) & "; "
        
    End If
    
    If mlngMandibularAss <> DentAssess_ProviderAssessEnum.daAss_Not_Addressed Then
        strNote = strNote & "Provider's assessment of the patient's current " & _
            "mandibular denture is " & GetAssessmentEnumDesc(mlngMandibularAss) & "; "
    End If
    
    If mlngMaxillaryAss <> DentAssess_ProviderAssessEnum.daAss_Not_Addressed Then
        strNote = strNote & "Provider's assessment of the patient's current " & _
            "maxillary denture is " & GetAssessmentEnumDesc(mlngMaxillaryAss) & "; "
    End If
    
    If Trim(mstrComments) <> "" Then
        strNote = strNote & "Comments: " & mstrComments & "; "
    End If
    
    If strNote <> "" Then
        strNote = "Denture Assessment - General " & strNote
    End If
    clsISection_GetNote = strNote
End Function

Private Function clsISection_GetRTFNote() As String
    Dim strNote As String
    
    strNote = Empty
          
    If Trim(mstrExpectations) <> "" Then
        strNote = strNote & "Expectectations: " & mstrExpectations & "; "
    End If
    
    If mlngPersonality <> DentAssess_PersonalityEnum.daPer_Not_Addressed Then
        strNote = strNote & "Personality: " & _
            GetPersonalityEnumDesc(mlngPersonality) & "; "
        
    End If
    
    If mlngMandibularAss <> DentAssess_ProviderAssessEnum.daAss_Not_Addressed Then
        strNote = strNote & "Provider's assessment of the patient's current " & _
            "mandibular denture is " & GetAssessmentEnumDesc(mlngMandibularAss) & "; "
    End If
    
    If mlngMaxillaryAss <> DentAssess_ProviderAssessEnum.daAss_Not_Addressed Then
        strNote = strNote & "Provider's assessment of the patient's current " & _
            "maxillary denture is " & GetAssessmentEnumDesc(mlngMaxillaryAss) & "; "
    End If
    
    If Trim(mstrComments) <> "" Then
        strNote = strNote & "Comments: " & mstrComments & "; "
    End If
    
    If strNote <> "" Then
        strNote = crtfStartControlName & "Denture Assessment General: " & _
            crtfEndControlName & strNote
    End If
    clsISection_GetRTFNote = strNote
End Function

Private Function clsISection_LoadSection(ByVal PatientID As Long, ByVal DentalEncounterID As Long) As Boolean
    Dim blnReturn As Boolean
    Dim objRS As ADODB.Recordset
    
    On Error GoTo ErrHandler
    
    mlngPatUnitId = PatientID
    mlngDentalEncounterId = DentalEncounterID
    
    Set objRS = DB_Load_Denture_Assessment(mlngDentalEncounterId)
    
    With objRS
        If .EOF Then
            mlngDirtyFlag = dfAdd
            mlngId = 0
            mstrExpectations = Empty
            mstrComments = Empty
            mlngPersonality = DentAssess_PersonalityEnum.daPer_Not_Addressed
            mlngMandibularAss = DentAssess_ProviderAssessEnum.daAss_Not_Addressed
            mlngMaxillaryAss = DentAssess_ProviderAssessEnum.daAss_Not_Addressed
        
        Else
            mlngDirtyFlag = dfUpdate
            mlngId = .Fields("EP_DA_ID")
            mstrExpectations = .Fields("EXPECTATIONS") & ""
            mstrComments = .Fields("COMMENTS") & ""
            mlngPersonality = .Fields("PERSONALITY") & ""
            mlngMandibularAss = .Fields("CURRENT_MANDIBULAR")
            mlngMaxillaryAss = .Fields("CURRENT_MAXILLARY")
        End If
    End With
    
FuncExit:
    clsISection_LoadSection = blnReturn
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
    
    strSql = Empty
    Select Case mlngDirtyFlag
        Case dfNoAction
            blnReturn = True
            GoTo FuncExit
        Case dfAdd
            mlngId = goconn.GetNextID(esTEMPL_EXAM_PROS_DENTURE_SEQ)
'            strSql = "SELECT TEMPL_EXAM_PROS_DENTURE_Seq.NEXTVAL From Dual"
'            Set objRS = goconn.OpenRecordset(strSql)
'            mlngId = objRS.Fields("NEXTVAL")

            strSql = "INSERT INTO TEMPL_EXAM_PROS_DENTURE " & _
                     "(EP_DA_ID, DENTAL_ENCOUNTER_ID, EXPECTATIONS, " & _
                     "PERSONALITY, CURRENT_MANDIBULAR, " & _
                     "CURRENT_MAXILLARY, COMMENTS) " & _
                     "VALUES (" & mlngId & ", " & mlngDentalEncounterId & ", '" & _
                    Scrub_Apostrophe(mstrExpectations) & "', " & mlngPersonality & ", " & _
                    mlngMandibularAss & ", " & mlngMaxillaryAss & ", '" & _
                    Scrub_Apostrophe(mstrComments) & "')"

        Case dfUpdate
            strSql = "UPDATE TEMPL_EXAM_PROS_DENTURE " & _
                     " SET DENTAL_ENCOUNTER_ID = " & mlngDentalEncounterId & _
                     " , EXPECTATIONS = '" & Scrub_Apostrophe(mstrExpectations) & "', " & _
                     " PERSONALITY = " & mlngPersonality & ", CURRENT_MANDIBULAR = " & mlngMandibularAss & _
                     ", CURRENT_MAXILLARY = " & mlngMaxillaryAss & ", COMMENTS = '" & _
                     Scrub_Apostrophe(mstrComments) & "' " & _
                     " WHERE EP_DA_ID = " & mlngId
        Case dfdelete
            strSql = "DELETE FROM TEMPL_EXAM_PROS_DENTURE WHERE " & _
                     " EP_DA_ID = " & mlngId
    End Select
    
    If strSql <> "" Then
        goconn.ExecuteSQL (strSql)
    End If
    mlngDirtyFlag = dfNoAction
    
    'If there's a collection, loop through it here.
    blnReturn = True
FuncExit:
    clsISection_SaveSection = blnReturn
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

Private Function DB_Load_Denture_Assessment(ByVal lngDentalEncounterId As Long) As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo ErrHandler
    
    strSql = "Select EP_DA_ID, DENTAL_ENCOUNTER_ID, EXPECTATIONS, " & _
                     "PERSONALITY, CURRENT_MANDIBULAR, " & _
                     "CURRENT_MAXILLARY, COMMENTS " & _
                     "FROM TEMPL_EXAM_PROS_DENTURE " & _
                     "WHERE DENTAL_ENCOUNTER_ID = " & lngDentalEncounterId
    
    Set DB_Load_Denture_Assessment = goconn.OpenRecordset(strSql)
    
    Exit Function
ErrHandler:
    goCWShared.ShowVBError Err.Number, Err.Description, const_ModuleName & ".DB_Load_Denture_Assessment()", "", vbCritical
    Set DB_Load_Denture_Assessment = Nothing
End Function
'*******************************************************************
'Function: GetAssessmentEnumDesc
'Author: KDunne
'Date: 12/11/2001
'Purpose. Returns the description for a given DentAssess_ProviderAssessEnum value.
'*******************************************************************
Public Function GetAssessmentEnumDesc(ByVal lngValue As _
    DentAssess_ProviderAssessEnum) As String
    
    Dim strReturn As String
    
    Select Case lngValue
        Case DentAssess_ProviderAssessEnum.daAss_Acceptable
            strReturn = "Acceptable"
        Case DentAssess_ProviderAssessEnum.daAss_Unacceptable
            strReturn = "Unacceptable"
        Case DentAssess_ProviderAssessEnum.daAss_Not_Addressed
            strReturn = Empty
        Case Else
            strReturn = Empty
    End Select
    
    GetAssessmentEnumDesc = strReturn
End Function

'*******************************************************************
'Function: GetPersonalityEnumDesc
'Author: KDunne
'Date: 12/11/2001
'Purpose. Returns the description for a given DentAssess_PersonalityEnum value.
'*******************************************************************
Public Function GetPersonalityEnumDesc(ByVal lngValue As _
            DentAssess_PersonalityEnum) As String

    Dim strReturn As String
    
    Select Case lngValue
        Case DentAssess_PersonalityEnum.daPer_Exacting
            strReturn = "Exacting"
        Case DentAssess_PersonalityEnum.daPer_Hysterical
            strReturn = "Hysterical"
        Case DentAssess_PersonalityEnum.daPer_Indifferent
            strReturn = "Indifferent"
        Case DentAssess_PersonalityEnum.daPer_Not_Addressed
            strReturn = Empty
        Case DentAssess_PersonalityEnum.daPer_Philosophical
            strReturn = "Philosophical"
        Case Else
            strReturn = Empty
    
    End Select
    
    GetPersonalityEnumDesc = strReturn
End Function
