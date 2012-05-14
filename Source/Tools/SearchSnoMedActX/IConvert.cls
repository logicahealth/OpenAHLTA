VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IConvert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'''''''''''''''''''''''''''''
'   3M Conversion
'''''''''''''''''''''''''''''
Public Enum enum3M_Domain
    DIAGNOSIS_DOMAIN = 658
    PROCEDURE_DOMAIN = 41
    HCPCS_DOMAIN = 17273
End Enum

Public Enum enum3M_CONTEXT
    NCID_DIAGNOSIS_COLUMN_CONTEXT = 1257
    NCID_PROCEDURE_COLUMN_CONTEXT = 17274
    NCID_HCPCS_COLUMN_CONTEXT = 78828
End Enum
Private m_3M_NCID As Long
Private m_3M_LongName As String
Private m_3M_ShortName As String
Private m_3M_Code As String

'SF SCR 72733
Public Enum eTermType
    ettSymptom = 1
    ettPhysical = 2
    ettTest = 3
    ettHistory = 5
    ettDisease = 6
    ettRx = 7
    ettRef = 8
    ettAll = 0
End Enum

''____________________________
''Sno Search Options
'Public Enum enumMedcinSearch
'   Diagnosis_Search = 0
'   Procedure_Search = 1
'   Patient_Instruction_Search = 2
'   All_Search = 3
'   E_M_Search = 4
'End Enum


'____________________________
'E&M Code related
Public Enum EnumEMCodeSelected
    esDefault = 1
    esOverride
    esNone
End Enum

Public Enum EnumSetting
    sEmergency_Dept = 100
    sHome
    sLong_Term_Facility
    sHospital_Inpatient
    sHospital_Observation
    sNursing_Facility
    sOutpatient
End Enum
Public Enum EnumService
    sPrev_Med_Admin_Assess = 200
    sConsult_Confirmatory
    sExt_Dr_Svc_W_Patient
    sPrev_Med_Eval_Mgt
    sCase_Mgt_Conf
    sPrev_Med_Group_Couns
    sExt_Dr_No_Pat_Contact
    sPrev_Med_Indiv_Couns
    sLife_Disab_Eval_Basic
    sWork_Med_Disab_Dr
    sNeonatal_IC
    sPrev_Med_Other
    sCare_Plan_Coord
    sCritical_Care
    sDr_Standby
    sCase_Mgt_Phone
    sOther_Unlisted_E_M
    sWork_Med_Disab_Other
    sDirected_Care
    sNewborn_H_P
    sNewborn_Non_Hospital
    sNewborn_Sub_Hosp_Care
    sNewborn_H_P_W_Dischg
    sNewborn_Attend_Deliv
    sNewborn_Resuscitation
    sPhysician_Care
    sObservation_Discharge
    sInitial_Observation
    sConsult_Initial_Inpatient
    sFinal_Discharge
    sConsult_Follow_Up_Inpatient
    sInitial_Hospital_Care
    sInpatient_Observ_Same_Day
    sSubsequent_Hospital_Care
    sLong_Term_Facility_Visit
    sHome_Visit
    sNursing_Assessment
    sNursing_Discharge
    sConsult_Initial_Patient
    sConsult_Follow_Up_Inpatient2
    sNursing_Subsequent_Care
    sOutpatient_Consult
    sOutpatient_Visit

End Enum
Public Enum EnumPatientStatus
    psNew_Patient = 300
    psExisting_Patient
End Enum
Public Enum EnumExamType
    etGeneral_Multi_System = 400
    etCardiovascular
    etEar_Nose_Mouth_And_Throat
    etEyes
    etGenitourinary
    etHematologic_Lymphatic_Immunologic
    etMusculoskeletal
    etNeurological
    etPsychiatric
    etRespiratory
    etSkin
End Enum
Public Enum EnumCodeDetail
    cdHPI = 1
    cdROS
    cdPFSH
    cdOverall_History
    cdExam
    cdOvarall_MDM
    cdDX_MGT_Options
    cdCompexity_Of_Data
    cdOverall_Risk
    cdPresenting_Problem
    cdDiagnostic_Procedures_Ordered
    cdManagement_Options_Selected
End Enum

Private Function VerifyCPTCode(ID As String) As String

    Dim oRS As Recordset
    Dim sSQL As String
    
    sSQL = "Select CPT_CODE from CPT_CODES_CHCS where CPT_CODE = '" & ID & "'"
    
    Set oRS = OpenRecordset(sSQL)
    
    If Not oRS.EOF Then
        VerifyCPTCode = ID
    Else
        MsgBox "Code " & ID & " is not a CPT code."
    End If
    
    oRS.Close
    Set oRS = Nothing
    
End Function

Public Function SnoTermType(SnoID As String) As eTermType
    
    'Only called from encounter parent, enctransfer.TransferSymptoms
    
    Dim oRS As Recordset
    Dim sSQL As String
    
    sSQL = "Select ConceptType from SNO_Search where DescriptionID = " & SnoID
    
    Set oRS = OpenRecordset(sSQL)
    
    If Not oRS.EOF Then
        Select Case oRS("ConceptType")
            
            Case "(finding)"
                SnoTermType = ettSymptom
            
        End Select
    End If
    
    oRS.Close
    Set oRS = Nothing
    
End Function


''''''''''''''''''''''''
'Medcin Conversions
''''''''''''''''''''''''
Public Function SnoTranscribeFamHx(ByVal SnoID As String) As String
Attribute SnoTranscribeFamHx.VB_Description = "Returns a Family History description."
    
    SnoTranscribeFamHx = SnoTranscribe(SnoID)
    
End Function

'Used by Couplers to determine which Medcin Header an item belongs to SCR-38863
Public Function SnoTranscribe(ByVal SnoID As String) As String

    Dim oRS As Recordset
    Dim sSQL As String
    
    sSQL = "SELECT   SNO_CONCEPTS.FULLYSPECIFIEDNAME" _
            & " From SNO_CONCEPTS, SNO_DESCRIPTIONS" _
            & " Where (SNO_DESCRIPTIONS.ConceptId = SNO_CONCEPTS.ConceptId)" _
            & " AND (SNO_DESCRIPTIONS.DESCRIPTIONID = " & SnoID & ")"

    Set oRS = OpenRecordset(sSQL)
    
    If Not oRS.EOF Then
        SnoTranscribe = oRS("FULLYSPECIFIEDNAME")
        'Just use the part in parenthesis
        'Entire posterior auricular vein (body structure)
        If InStr(SnoTranscribe, "(") > 0 Then
            SnoTranscribe = Mid(SnoTranscribe, InStr(SnoTranscribe, "("))
            SnoTranscribe = Left(SnoTranscribe, Len(SnoTranscribe) - 1)
        End If
    End If
    
    oRS.Close
    Set oRS = Nothing
    
End Function

Public Function SnoICD9(ByVal SnoID As String) As String
Attribute SnoICD9.VB_Description = "Returns an ICD9 from MedcinID"

On Error GoTo ErrHnd

    SnoICD9 = SnoICDCode(SnoID)

Exit Function

ErrHnd:
    MsgBxARMd "Error in IConvert.SnoICD9: " & Err.Description

End Function
Public Function SnoCPT4(ByVal SnoID As String) As String
Attribute SnoCPT4.VB_Description = "Returns an CPT4 from MedcinID"
    'Change to CPT4 and ID
    
    SnoCPT4 = VerifyCPTCode(SnoID)
    
End Function

'Returns known possible descriptions
'Used by AnP for validating encounters
Public Function AnP_PossibleSnoDescriptions(ByVal SnoID As String) As Collection

    Dim c As Collection
    Set c = New Collection
    
    c.Add SnoDescription(SnoID)
    c.Add SnoTranscribe(SnoID)

    Set AnP_PossibleSnoDescriptions = c
    
End Function

Public Function SnoDescription(ByVal SnoID As String) As String
Attribute SnoDescription.VB_Description = "Returns an Description from MedcinID"
    
    Dim oRS As Recordset
    Dim sSQL As String
    
    sSQL = "Select Term from SNO_Descriptions where DescriptionID = " & SnoID
    
    Set oRS = OpenRecordset(sSQL)
    
    If Not oRS.EOF Then
        SnoDescription = oRS("Term")
    End If
    
    oRS.Close
    Set oRS = Nothing
    
End Function
Public Function CPTDescription(ByVal CPTID As String) As String
    
    Dim oRS As Recordset
    Dim sSQL As String
    
    sSQL = "Select FULL_NARRATIVE from CPT_CODES_CHCS where CPT_CODE = '" & CPTID & "'"
    
    Set oRS = OpenRecordset(sSQL)
    
    If Not oRS.EOF Then
        CPTDescription = oRS("FULL_NARRATIVE")
    End If
    
    oRS.Close
    Set oRS = Nothing
    
End Function

Private Function SnoICDCode(SnoID As String) As String

    Dim oRS As Recordset
    Dim sSQL As String
    
                    'Min SnoID
    If SnoID = "0" Then
        Exit Function
    ElseIf Val(SnoID) < 101013 Then
        MsgBox "Trying to convert medcinid: " & SnoID
        Exit Function
    End If
        
    'Fast search for active codes
    sSQL = "SELECT TARGETCODES FROM SNO_Search Where DESCRIPTIONID = " & SnoID
    
    Set oRS = OpenRecordset(sSQL)
    
    If oRS.EOF Then
        'Perform longer search covering all codes.
        sSQL = "SELECT SNO_CROSS_MAP_TARGETS.TARGETCODES" _
            & " From SNO_DESCRIPTIONS, SNO_CROSS_MAPS, SNO_CROSS_MAP_TARGETS" _
            & " Where (SNO_DESCRIPTIONS.ConceptId = SNO_CROSS_MAPS.MAPCONCEPTID)" _
            & " AND (SNO_CROSS_MAPS.MAPTARGETID = SNO_CROSS_MAP_TARGETS.TARGETID)" _
            & " AND (SNO_DESCRIPTIONS.DESCRIPTIONID = " & SnoID & ")"
        Set oRS = OpenRecordset(sSQL)
    End If
    
    If Not oRS.EOF Then
        SnoICDCode = oRS("TARGETCODES")
        If InStr(SnoICDCode, "|") Then
            SnoICDCode = Split(SnoICDCode, "|")(0)
        End If
    Else
        Err.Raise 1, "SnoICDCode", "Could not find an ICD code for SnoID = " & SnoID
    End If
    
    oRS.Close
    Set oRS = Nothing

End Function

