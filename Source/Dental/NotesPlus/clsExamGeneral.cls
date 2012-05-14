VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsExamGeneral"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
  'Class Module Created by Sean McAvoy
  'Hold values for the Exam/Eval General Tab for the Dental Templates.
  
Option Explicit

Implements clsISection

Private Const XML_ELEMENT_NAME As String = "clsExamGeneral"
'Private mvarXMLPersist As New clsXMLPersist
'Private mvarXMLElement As IXMLDOMElement

'Used to determine if the data should be inserted or updated when modified.
Private bSavedToDB As Boolean
Private dfDirtyFlag As DirtyFlagEnum

'default from template
Private mvarPatientID As Long
Private mvarDentalEncounterID As Long
Private mobjSBE As IntrfcCHCSII_Dental.IDENTAL_SBE

'***********************************
'Evaluation section
Private mintSick_Call As Integer
Private mintOral_Surgery As Integer
Private mintPeriodontal As Integer
Private mintProsthodontics As Integer
Private mintEndodontics As Integer
Private mintOrthodontics As Integer
Private mintPediatric As Integer
Private mintOrofacialPain As Integer
Private mintAfter_HoursCare As Integer
Private mintConsultation_Response As Integer

Private mlSick_Call_WLID As Long
Private mbSick_Call_WL_chg As Boolean
Private mlAfter_HoursCare_WLID As Long
Private mbAfter_HoursCare_chg As Boolean
Private mlConsultation_Response_WLID As Long
Private mbConsultation_Response_chg As Boolean

'Workload codes
Private Const Sick_Call_WLcode As Long = 705
Private Const After_HoursCare_WLcode As Long = 686
Private Const mintConsultation_Response_WLcode As Long = 78
'***********************************



'********
'Oral Exam section
Private mintOralExam As OralExamEN
Private mstrOralExamNote As String
Private mlOralExamWLID As Long
Private mbOralExam_WLID_chg As Boolean
'Note! the OralExamEN is coordinated with Index in Notes chkOE_MouseUp
Public Enum OralExamEN
    NoOralExamValue = 0
    Periodic = 1
    Comprehensive = 2
    DetailedExtensive = 3
    LimitedProblemFocused = 4
    LimitedReevaluation = 5
    PreOrtho = 6
    PeriodicOrtho = 7
    PerioComprehensive = 8
    Postmortem = 9
    RoutinePostOp = 10
    NonRoutinePostOp = 11
End Enum

'*********
'Tobacco section
Private mintTobaccoUse As TobaccoUseEN
Private mintTobaccoUseCounseled As Integer
Private mstrTobaccoUseNote As String
Private mlTobacco_Counseled_WLID As Long
Private mbTobacco_Counseled_WL_chg As Boolean
Private Const TU_Counseled_WLcode As Integer = 465
Public Enum TobaccoUseEN
    NoTobaccoUseValue = 0
    DontSmoke = 1
    Smoke = 2
    smokeless = 3
    UseBoth = 4
End Enum

'**********

'Other vars
Private mstrGCSupp_Procedures As String 'SCR 19977
Private mbPatientIsAdult As Boolean
Private mdblDRC As Double

'SCRs 8649, 8652, 8655, 8656, 8657, 8658
Private mstrChief_Complaint As String
Private mstrMedical_History As String
Private mintMedical_History_NC As Integer
Private mstrDental_History As String
Private mintDental_History_NC As Integer
Private mstrPresent_Problem As String
Private mstrPostOperative As String
Private mstrProgressNote As String

'****************************
Private mintSBE_Require As Integer
'yes = -1
'no = 0
'Field in table Dental_Parent
Private mbUpdateSBE As Boolean
Private mbSBEcontrolLocked As Boolean
Private mstrSBE_Text As String
'****************************




Public Property Let Chief_Complaint(ByVal vData As String)
    mstrChief_Complaint = vData
    UpdateDirtyFlag
End Property
Public Property Get Chief_Complaint() As String
    Chief_Complaint = mstrChief_Complaint
End Property
Private Function CalcOralExamWorkloadCode(OralExamType As OralExamEN) As Long

'This function translates the OralExamType into a TxADA Code ID from the Tx_ADA_Code_Definition table

    Select Case OralExamType
    
        Case OralExamEN.Periodic
            CalcOralExamWorkloadCode = 313
            
        Case OralExamEN.Comprehensive
            CalcOralExamWorkloadCode = 69
            
        Case OralExamEN.DetailedExtensive
            CalcOralExamWorkloadCode = 760
        
        Case OralExamEN.LimitedProblemFocused
            CalcOralExamWorkloadCode = 215
        
        Case OralExamEN.LimitedReevaluation
            CalcOralExamWorkloadCode = 540
        
        Case OralExamEN.PreOrtho
            CalcOralExamWorkloadCode = 338
        
        Case OralExamEN.PeriodicOrtho
            CalcOralExamWorkloadCode = 314
        
        Case OralExamEN.PerioComprehensive
            CalcOralExamWorkloadCode = 701
        
        Case OralExamEN.Postmortem
            CalcOralExamWorkloadCode = 329
    
        Case OralExamEN.RoutinePostOp
            CalcOralExamWorkloadCode = 271
    
        Case OralExamEN.NonRoutinePostOp
            CalcOralExamWorkloadCode = 476
            
    End Select

End Function

Public Property Let PostOperative(ByVal vData As String)
    mstrPostOperative = vData
    UpdateDirtyFlag
End Property
Public Property Get PostOperative() As String
    PostOperative = mstrPostOperative
End Property
Public Property Let Medical_History_NC(ByVal vData As Integer)
    mintMedical_History_NC = vData
    UpdateDirtyFlag
End Property
Public Property Get Medical_History_NC() As Integer
    Medical_History_NC = mintMedical_History_NC
End Property
Public Property Let Medical_History(ByVal vData As String)
    mstrMedical_History = vData
    UpdateDirtyFlag
End Property
Public Property Get Medical_History() As String
    Medical_History = mstrMedical_History
End Property
Public Property Let Dental_History_NC(ByVal vData As Integer)
    mintDental_History_NC = vData
    UpdateDirtyFlag
End Property
Public Property Get Dental_History_NC() As Integer
    Dental_History_NC = mintDental_History_NC
End Property
Public Property Let Dental_History(ByVal vData As String)
    mstrDental_History = vData
    UpdateDirtyFlag
End Property
Public Property Get Dental_History() As String
    Dental_History = mstrDental_History
End Property
Public Property Let ProgressNote(ByVal vData As String)
    mstrProgressNote = vData
    UpdateDirtyFlag
End Property
Public Property Get ProgressNote() As String
    ProgressNote = mstrProgressNote
End Property
Public Property Let Present_Problem(ByVal vData As String)
    mstrPresent_Problem = vData
    UpdateDirtyFlag
End Property
Public Property Get Present_Problem() As String
    Present_Problem = mstrPresent_Problem
End Property

Public Property Set SBE(objSBE As IntrfcCHCSII_Dental.IDENTAL_SBE)
    Set mobjSBE = objSBE
End Property
Public Property Let SBE_Text(ByVal vData As String)
    mstrSBE_Text = vData
End Property

Public Property Get SBE_Text() As String
    SBE_Text = mstrSBE_Text
End Property

Public Property Let GC_Supporting_Procedures(ByVal strValue As String)
    mstrGCSupp_Procedures = strValue
End Property

Public Property Get GC_Supporting_Procedures() As String
    If mstrGCSupp_Procedures = "" Then
        mstrGCSupp_Procedures = gcSupportingProcs
    End If
    GC_Supporting_Procedures = mstrGCSupp_Procedures
End Property

'** END ************************************************************
'*******************************************************************
'*******************************************************************

Public Property Let SBEcontrolLocked(ByVal vData As Boolean)
    mbSBEcontrolLocked = vData
End Property
Public Property Get SBEcontrolLocked() As Boolean
    SBEcontrolLocked = mbSBEcontrolLocked
End Property

Public Property Let SBE_Require(ByVal vData As Integer)
    mintSBE_Require = vData
    mbUpdateSBE = True
End Property
Public Property Get SBE_Require() As Integer
    SBE_Require = mintSBE_Require
End Property



Public Property Let PatientIsAdult(ByVal vData As Boolean)
    mbPatientIsAdult = vData
End Property
Public Property Let After_HoursCare(ByVal vData As Integer)
    mintAfter_HoursCare = vData
    UpdateDirtyFlag
    mbAfter_HoursCare_chg = True
End Property
Public Property Get After_HoursCare() As Integer
    After_HoursCare = mintAfter_HoursCare
End Property
'**** end of workload vars



Public Property Let Sick_Call(ByVal vData As Integer)
    mintSick_Call = vData
    UpdateDirtyFlag
    mbSick_Call_WL_chg = True
End Property
Public Property Get Sick_Call() As Integer
    Sick_Call = mintSick_Call
End Property
Public Property Let Oral_Surgery(ByVal vData As Integer)
    
    mintOral_Surgery = vData
    UpdateDirtyFlag

End Property
Public Property Get Oral_Surgery() As Integer
    Oral_Surgery = mintOral_Surgery
End Property
Public Property Let Periodontal(ByVal vData As Integer)
    mintPeriodontal = vData
    UpdateDirtyFlag
End Property
Public Property Get Periodontal() As Integer
    Periodontal = mintPeriodontal
End Property
Public Property Let Pediatric(ByVal vData As Integer)
    mintPediatric = vData
    UpdateDirtyFlag
End Property
Public Property Get Pediatric() As Integer
    Pediatric = mintPediatric
End Property

Public Property Let Prosthodontics(ByVal vData As Integer)
    mintProsthodontics = vData
    UpdateDirtyFlag
End Property
Public Property Get Prosthodontics() As Integer
    Prosthodontics = mintProsthodontics
End Property
Public Property Let Endodontics(ByVal vData As Integer)
    mintEndodontics = vData
    UpdateDirtyFlag
End Property
Public Property Get Endodontics() As Integer
    Endodontics = mintEndodontics
End Property
Public Property Let Orthodontics(ByVal vData As Integer)
    mintOrthodontics = vData
    UpdateDirtyFlag
End Property
Public Property Get Orthodontics() As Integer
    Orthodontics = mintOrthodontics
End Property
Public Property Let OrofacialPain(ByVal vData As Integer)
    mintOrofacialPain = vData
    UpdateDirtyFlag
End Property
Public Property Get OrofacialPain() As Integer
    OrofacialPain = mintOrofacialPain
End Property
Public Property Let Consultation_Response(ByVal vData As Integer)
    mintConsultation_Response = vData
    UpdateDirtyFlag
    mbConsultation_Response_chg = True
End Property
Public Property Get Consultation_Response() As Integer
    Consultation_Response = mintConsultation_Response
End Property

Public Property Get PatientIsAdult() As Boolean
    PatientIsAdult = mbPatientIsAdult
End Property

Public Property Let OralExam(ByVal vData As OralExamEN)
    mintOralExam = vData
    mbOralExam_WLID_chg = True
    UpdateDirtyFlag
End Property
Public Property Get OralExam() As OralExamEN
    OralExam = mintOralExam
End Property

Public Property Let TobaccoUseCounseled(ByVal vData As Integer)
    mintTobaccoUseCounseled = vData
    mbTobacco_Counseled_WL_chg = True
    UpdateDirtyFlag
End Property

Public Property Get TobaccoUseCounseled() As Integer
    TobaccoUseCounseled = mintTobaccoUseCounseled
End Property

Public Property Let TobaccoUseNote(ByVal vData As String)
    mstrTobaccoUseNote = vData
    UpdateDirtyFlag
End Property

Public Property Get TobaccoUseNote() As String
    TobaccoUseNote = mstrTobaccoUseNote
End Property

Public Property Let TobaccoUse(ByVal vData As Integer)
    mintTobaccoUse = vData
    UpdateDirtyFlag
End Property

Public Property Get TobaccoUse() As Integer
    TobaccoUse = mintTobaccoUse
End Property


Public Property Let OralExamNote(ByVal vData As String)
    mstrOralExamNote = vData
    UpdateDirtyFlag
End Property

Public Property Get OralExamNote() As String
    OralExamNote = mstrOralExamNote
End Property



Private Sub Class_Terminate()
    
'    Set mvarXMLElement = Nothing
'    Set mvarXMLPersist = Nothing

End Sub

Private Property Get clsISection_DentalEncounterID() As Long
    
    clsISection_DentalEncounterID = mvarDentalEncounterID

End Property



Private Property Get clsISection_PatientID() As Long
    
    clsISection_PatientID = mvarPatientID

End Property



Private Function clsISection_GetRTFNote() As String

    'clsISection_GetRTFNote = clsISection_GetNote
 Dim sTxt As String
 Dim sNote As String
 Const sPROC_NAME = "clsISection_GetRTFNote"
 On Error GoTo errTrap
 
 sNote = Empty
 
If mintSick_Call = 0 And mintOral_Surgery = 0 And mintPeriodontal = 0 And _
        mintProsthodontics = 0 And mintEndodontics = 0 And _
        mintOrthodontics = 0 And mintOrofacialPain = 0 And _
        mintAfter_HoursCare = 0 And mintConsultation_Response = 0 And _
        mintOralExam = NoOralExamValue And _
        mstrOralExamNote = "" And _
        mintTobaccoUse = NoTobaccoUseValue And _
        mintTobaccoUseCounseled = 0 And _
        mstrTobaccoUseNote = "" And _
        mstrChief_Complaint = "" And _
        mstrMedical_History = "" And mintMedical_History_NC = 0 And mstrDental_History = "" And _
        mintDental_History_NC = 0 And mstrPresent_Problem = "" And mstrPostOperative = "" And _
        mintPediatric = 0 And mstrProgressNote = "" Then
    sNote = ""
    
    clsISection_GetRTFNote = sNote
    Exit Function
 End If
 
    If mdblDRC = 1 Then
        sNote = sNote & crtfStartControlName & "Class -1 Patient Readiness" & crtfEndControlName
    End If
    
    If mintOralExam > 0 Or mstrOralExamNote <> "" Then
        sNote = sNote & crtfStartControlName & "Oral Exam: " '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
        
        Select Case mintOralExam
           Case OralExamEN.Periodic
               sNote = sNote & "Periodic; "
           Case OralExamEN.Comprehensive
               sNote = sNote & "Comprehensive; "
           Case OralExamEN.DetailedExtensive
               sNote = sNote & "Detailed/Extensive; "
           Case OralExamEN.LimitedProblemFocused
               sNote = sNote & "Limited Problem Focused; "
           Case OralExamEN.LimitedReevaluation
               sNote = sNote & "Limited Reevaluation; "
           Case OralExamEN.PreOrtho
               sNote = sNote & "Pre-Ortho; "
           Case OralExamEN.PeriodicOrtho
               sNote = sNote & "Periodic Ortho; "
           Case OralExamEN.PerioComprehensive
               sNote = sNote & "Perio Comp; "
           Case OralExamEN.Postmortem
               sNote = sNote & "Post Mortem; "
           Case OralExamEN.RoutinePostOp
               sNote = sNote & "Routine Post-Op; "
           Case OralExamEN.NonRoutinePostOp
               sNote = sNote & "Non Routine Post-Op; "
        End Select
        
        If mstrOralExamNote <> "" Then
            sNote = sNote & "Oral Exam Comment: " & mstrOralExamNote & "; "
        End If
        'Remove the last "; "
        sNote = Left(sNote, Len(sNote) - Len("; "))
        sNote = sNote & crtfEndControlName
    End If
    
    If mintTobaccoUse > NoTobaccoUseValue Or mstrTobaccoUseNote <> "" Then
        sNote = sNote & crtfStartControlName & "Tobacco Use: " '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
        Select Case mintTobaccoUse
            Case TobaccoUseEN.UseBoth
                sNote = sNote & "Smokes tobacco and uses smokeless tobacco; "
             Case TobaccoUseEN.DontSmoke
                sNote = sNote & "Non-smoker; "
            Case TobaccoUseEN.Smoke
                sNote = sNote & "Smokes tobacco; "
            Case TobaccoUseEN.smokeless
                sNote = sNote & "Uses smokeless tobacco; "
        End Select
        If mintTobaccoUseCounseled = -1 Then
            sNote = sNote & "Counseled about tobacco cessation; "
        End If
        If mstrTobaccoUseNote <> "" Then
            sNote = sNote & "Tobacco Use Comment: " & mstrTobaccoUseNote & "; "
        End If
                'Remove the last "; "
        sNote = Left(sNote, Len(sNote) - Len("; "))
        sNote = sNote & crtfEndControlName
    End If
    
    If mstrChief_Complaint <> "" Or mstrMedical_History <> "" Or mstrDental_History <> "" Or mintMedical_History_NC <> 0 Or _
        mintDental_History_NC <> 0 Or mstrPresent_Problem <> "" Or mstrPostOperative <> "" Or mstrProgressNote <> "" Then
            sNote = sNote & crtfStartControlName & "Subjective:" & crtfEndControlName '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
            
            If mstrChief_Complaint <> "" Then
                sNote = sNote & crtfStartControlName & crtfStartControlName & "Chief Complaint Comment: " & mstrChief_Complaint & "; " & crtfEndControlName
            End If
            
            If mstrMedical_History <> "" And mintMedical_History_NC <> 0 Then
                sNote = sNote & crtfStartControlName & crtfStartControlName & "Non-contributory Medical History Comment: " & mstrMedical_History & "; " & crtfEndControlName
            End If
            If mstrMedical_History <> "" And mintMedical_History_NC = 0 Then
                sNote = sNote & crtfStartControlName & crtfStartControlName & "Medical History Comment: " & mstrMedical_History & "; " & crtfEndControlName
            End If
            If mstrMedical_History = "" And mintMedical_History_NC <> 0 Then
                sNote = sNote & crtfStartControlName & crtfStartControlName & "Medical History Non-contributory; " & crtfEndControlName
            End If
                
                
            If mstrDental_History <> "" And mintDental_History_NC <> 0 Then
                sNote = sNote & crtfStartControlName & crtfStartControlName & "Non-contributory Dental History Comment: " & mstrDental_History & "; " & crtfEndControlName
            End If
            If mstrDental_History <> "" And mintDental_History_NC = 0 Then
                sNote = sNote & crtfStartControlName & crtfStartControlName & "Dental History Comment: " & mstrDental_History & "; " & crtfEndControlName
            End If
            If mstrDental_History = "" And mintDental_History_NC <> 0 Then
                sNote = sNote & crtfStartControlName & crtfStartControlName & "Dental History Non-contributory; " & crtfEndControlName
            End If
    
            If mstrPresent_Problem <> "" Then
                sNote = sNote & crtfStartControlName & crtfStartControlName & "Present Problem Comment: " & mstrPresent_Problem & "; " & crtfEndControlName
            End If
            If mstrPostOperative <> "" Then
                sNote = sNote & crtfStartControlName & crtfStartControlName & "Post Operative Comment: " & mstrPostOperative & "; " & crtfEndControlName
            End If
            If mstrProgressNote <> "" Then
                sNote = sNote & crtfStartControlName & crtfStartControlName & "Progress Comment: " & mstrProgressNote & "; " & crtfEndControlName
            End If
            'Remove the last "; "
            sNote = Left(sNote, Len(sNote) - Len("; "))
            sNote = sNote & crtfEndControlName
    
    End If

    'Evaluation
    If mintSick_Call <> 0 Or mintOral_Surgery <> 0 Or mintPeriodontal <> 0 Or _
        mintProsthodontics <> 0 Or mintEndodontics <> 0 Or _
        mintOrthodontics <> 0 Or mintOrofacialPain <> 0 Or _
        mintAfter_HoursCare <> 0 Or mintConsultation_Response <> 0 Or mintPediatric <> 0 Then
            sNote = sNote + crtfStartControlName & "Evaluation: " '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
            
           If mintAfter_HoursCare = -1 Or mintConsultation_Response = -1 Then
'                sNote = sNote + crtfStartControlName
                If mintAfter_HoursCare = -1 And mintConsultation_Response = -1 Then
                   sNote = sNote + "After Hour Care for a Consulation Response" + ", "
                ElseIf mintAfter_HoursCare = -1 Then
                   sNote = sNote + "After Hour Care" + ", "
                ElseIf mintConsultation_Response = -1 Then
                   sNote = sNote + "Consulation Response" + ", "
                End If
'                sNote = sNote + crtfEndControlName
            End If
           If mintSick_Call <> 0 Then sTxt = sTxt + "Sick Call; "
           If mintOral_Surgery <> 0 Then sTxt = sTxt + "Oral Surgery; "
           If mintPeriodontal <> 0 Then sTxt = sTxt + "Periodontal; "
           If mintProsthodontics <> 0 Then sTxt = sTxt + "Prosthodontics, "
           If mintEndodontics <> 0 Then sTxt = sTxt + "Endodontics; "
           If mintOrthodontics <> 0 Then sTxt = sTxt + "Orthodontics; "
           If mintPediatric <> 0 Then sTxt = sTxt + "Pediatric; "
           If mintOrofacialPain <> 0 Then sTxt = sTxt + "OrofacialPain; "
           sNote = sNote + sTxt
            
            'Remove the last "; "
            sNote = Left(sNote, Len(sNote) - Len("; "))
            sNote = sNote + crtfEndControlName
    End If

'    If mintSBE_Require <> 0 Then
'        sNote = sNote & crtfStartControlName & "" & "SBE Prophylaxis required for dental treatment" & crtfEndControlName
'    End If
'    If mstrSBE_Text <> "" Then
'        sNote = sNote & mstrSBE_Text
'    End If
    
    clsISection_GetRTFNote = sNote
    Exit Function
errTrap:
    clsISection_GetRTFNote = "Error in generating note for ExamEval-General section."
    goCWShared.ShowVBError Err.Number, Err.Description, Err.Source & "." & sPROC_NAME, App.Title, vbCritical
End Function
'
''*******************************************************************
''Function: clsISection_GetNote
''Author: Sean McAvoy
''Date: 12/4/2001
''Returns: A string that is displayed in the narrative summary for this encounter
''Purpose: To compare the values from this encounter and the previous encounter and
''return a text note.
''*******************************************************************
Private Function clsISection_GetNote() As String
 Dim sTxt As String
 Dim sNote As String
 Const sPROC_NAME = "clsISection_GetNote"
 On Error GoTo errTrap

If mintSick_Call = 0 And mintOral_Surgery = 0 And mintPeriodontal = 0 And _
        mintProsthodontics = 0 And mintEndodontics = 0 And _
        mintOrthodontics = 0 And mintOrofacialPain = 0 And _
        mintAfter_HoursCare = 0 And mintConsultation_Response = 0 And _
        mintOralExam = NoOralExamValue And _
        mstrOralExamNote = "" And _
        mintTobaccoUse = TobaccoUseEN.NoTobaccoUseValue And _
        mintTobaccoUseCounseled = 0 And _
        mstrTobaccoUseNote = "" And _
        mstrChief_Complaint = "" And _
        mstrMedical_History = "" And mintMedical_History_NC = 0 And mstrDental_History = "" And _
        mintDental_History_NC = 0 And mstrPresent_Problem = "" And mstrPostOperative = "" And _
        mstrProgressNote = "" Then
    sNote = ""
    clsISection_GetNote = sNote
    Exit Function
 End If

 sNote = "Exam Evaluation - General" & vbCrLf
 Select Case mintOralExam
    Case OralExamEN.Periodic
        sNote = sNote & vbTab & "Oral Exam Type: Periodic" & vbCrLf
    Case OralExamEN.Comprehensive
        sNote = sNote & vbTab & "Oral Exam Type: Comprehensive" & vbCrLf
    Case OralExamEN.DetailedExtensive
        sNote = sNote & vbTab & "Oral Exam Type: Detailed/Extensive" & vbCrLf
    Case OralExamEN.LimitedProblemFocused
        sNote = sNote & vbTab & "Oral Exam Type: Limited Problem Focused" & vbCrLf
    Case OralExamEN.LimitedReevaluation
        sNote = sNote & vbTab & "Oral Exam Type: Limited Reevaluation" & vbCrLf
    Case OralExamEN.PreOrtho
        sNote = sNote & vbTab & "Oral Exam Type: Pre-Ortho" & vbCrLf
    Case OralExamEN.PeriodicOrtho
        sNote = sNote & vbTab & "Oral Exam Type: Periodic Ortho" & vbCrLf
    Case OralExamEN.PerioComprehensive
        sNote = sNote & vbTab & "Oral Exam Type: Perio Comp" & vbCrLf
    Case OralExamEN.Postmortem
        sNote = sNote & vbTab & "Oral Exam Type: Post Mortem" & vbCrLf
    Case OralExamEN.RoutinePostOp
        sNote = sNote & vbTab & "Oral Exam Type: Routine Post-Op" & vbCrLf
    Case OralExamEN.NonRoutinePostOp
        sNote = sNote & vbTab & "Oral Exam Type: Non Routine Post-Op" & vbCrLf
 End Select
 
 If Trim(mstrOralExamNote) <> "" Then
    sNote = sNote & vbTab & "Oral Exam Note: " & mstrOralExamNote & vbCrLf
 End If

 Select Case mintTobaccoUse
    Case TobaccoUseEN.UseBoth
        sNote = sNote & vbTab & "Patient smokes tobacco and uses smokeless tobacco." & vbCrLf
     Case TobaccoUseEN.DontSmoke
        sNote = sNote & vbTab & "Patient is a non-smoker." & vbCrLf
    Case TobaccoUseEN.Smoke
        sNote = sNote & vbTab & "Patient smokes tobacco." & vbCrLf
    Case TobaccoUseEN.smokeless
        sNote = sNote & vbTab & "Patient uses smokeless tobacco." & vbCrLf
End Select
If mintTobaccoUseCounseled = -1 Then
    sNote = sNote & vbTab & "Patient has been counseled about tobacco cessation." & vbCrLf
End If
If Trim(mstrTobaccoUseNote) <> "" Then
        sNote = sNote & vbTab & "Tobacco Use Note: " & mstrTobaccoUseNote & vbCrLf
Else
    sNote = sNote & vbCrLf
End If

    If mstrChief_Complaint <> "" Then sNote = sNote & vbTab & "Chief Complaint Note: " & mstrChief_Complaint & vbCrLf
    If mstrMedical_History <> "" And mintMedical_History_NC <> 0 Then
        sNote = sNote & vbTab & "Medical History Non-contributory." & vbCrLf
    End If
    If mstrMedical_History <> "" And mintMedical_History_NC = 0 Then
        sNote = sNote & vbTab & "Medical History Note: " & mstrMedical_History & vbCrLf
    End If
'    If mstrMedical_History = "" And mintMedical_History_NC <> 0 Then
'        sNote = sNote & vbTab & "Medical History is Non-contributory." & vbCrLf
'    End If
        
    If mstrDental_History <> "" And mintDental_History_NC <> 0 Then
        sNote = sNote & vbTab & "Dental History Non-contributory." & vbCrLf
    End If
    If mstrDental_History <> "" And mintDental_History_NC = 0 Then
        sNote = sNote & vbTab & "Dental History Note: " & mstrDental_History & vbCrLf
    End If
'    If mstrDental_History = "" And mintDental_History_NC <> 0 Then
'        sNote = sNote & vbTab & "Dental History is Non-contributory." & vbCrLf
'    End If
        
    If mstrPresent_Problem <> "" Then sNote = sNote & vbTab & "Present Problem Note: " & mstrPresent_Problem & vbCrLf
    If mstrPostOperative <> "" Then sNote = sNote & vbTab & "Post Operative Note: " & mstrPostOperative & vbCrLf
    If mstrProgressNote <> "" Then sNote = sNote & vbTab & "Progress Note: " & mstrProgressNote & vbCrLf

    'Evaluation
    sNote = "Exam Evaluation - SOAP" + vbCrLf
    If mintAfter_HoursCare = -1 And mintConsultation_Response = -1 Then
       sNote = sNote + vbTab + "Patient encounter is during After Hour Care and is a Consulation Response." + vbCrLf
    ElseIf mintAfter_HoursCare = -1 Then
       sNote = sNote + vbTab + "Patient encounter is during After Hour Care." + vbCrLf
    ElseIf mintConsultation_Response = -1 Then
       sNote = sNote + vbTab + "Patient encounter is for a Consulation Response." + vbCrLf
    End If
 
    If mintSick_Call = 0 And mintOral_Surgery = 0 And mintPeriodontal = 0 And _
        mintProsthodontics = 0 And mintEndodontics = 0 And _
        mintOrthodontics = 0 And mintOrofacialPain = 0 Then
        
        'then skip this section
    Else
        sNote = sNote + vbTab + "Patient Evaluation is for: "
        
        If mintSick_Call <> 0 Then sTxt = sTxt + "Sick Call, "
        If mintOral_Surgery <> 0 Then sTxt = sTxt + "Oral Surgery, "
        If mintPeriodontal <> 0 Then sTxt = sTxt + "Periodontal, "
        If mintProsthodontics <> 0 Then sTxt = sTxt + "Prosthodontics, "
        If mintEndodontics <> 0 Then sTxt = sTxt + "Endodontics, "
        If mintOrthodontics <> 0 Then sTxt = sTxt + "Orthodontics, "
        If mintOrofacialPain <> 0 Then sTxt = sTxt + "OrofacialPain, "
        
        sTxt = Trim(sTxt)
        If sTxt <> "" Then
            If Mid(sTxt, Len(sTxt), 1) = "," Then
                sTxt = Mid(sTxt, 1, Len(sTxt) - 1)
            End If
            If sTxt <> "" Then sNote = sNote + sTxt + vbCrLf
        End If
            
    End If

sNote = sNote & vbCrLf

    clsISection_GetNote = sNote
    Exit Function
errTrap:
    clsISection_GetNote = "Error in generating note for ExamEval-General section."
    goCWShared.ShowVBError Err.Number, Err.Description, Err.Source & "." & sPROC_NAME, App.Title, vbCritical
End Function

'*******************************************************************
'Function: GenerateGeneralWL_ID
'Author: Sean McAvoy
'Date: 12/11/2001
'Returns: True if successful.
'Purpose: Called from SaveSection, it's main purpose is to set the mlXXX_WL variable.
'This variable is represents the workload code for each item on the screen that
'requires workload credit.
'*******************************************************************

Private Function GenerateGeneralWL_ID(sTmpMsg As String) As Boolean
Dim lTmpOE_WLcode As Long
Dim sTmp
Const sPROC_NAME As String = "GenerateGeneralWL_ID"

On Error GoTo ErrHnd

    If dfDirtyFlag = dfNoAction Then  '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        GenerateGeneralWL_ID = True
        Exit Function
    Else
        'Update the workload
        If mvarDentalEncounterID = 0 Then
            GenerateGeneralWL_ID = False
            Exit Function
        End If
        
        '**************************************
        'Update the Workload for Tobacco Counseled
        If mbTobacco_Counseled_WL_chg = True Then
            If mintTobaccoUseCounseled = -1 And mlTobacco_Counseled_WLID = 0 Then
                'set dfDirtyflag = dfAdd
                'The tobacco use Counseled has been selected in the GUI but no record
                '  has been entered into DB to represent this.
                If Not UpdateWorkload(mlTobacco_Counseled_WLID, dfAdd, mvarDentalEncounterID, TU_Counseled_WLcode, 1) Then
                    GenerateGeneralWL_ID = False
                    sTmpMsg = "Error. Could not correctly update the Tobacco Counseled Workload Credits. dfAdd failed. "
                End If
            ElseIf mintTobaccoUseCounseled = -1 And mlTobacco_Counseled_WLID > 0 Then
                'set dfDirtyflag = upDate
                'Tobacco use Counseled has already been selected previously, data has already
                '  been entered in to DB. Updates the Deleted column in the Templ_Workload table.
                If Not UpdateWorkload(mlTobacco_Counseled_WLID, dfUpdate, mvarDentalEncounterID, TU_Counseled_WLcode, 1) Then
                    GenerateGeneralWL_ID = False
                    sTmpMsg = "Error. Could not correctly update the Tobacco Counseled Workload Credits. dfUpdate failed. "
                End If
            ElseIf mintTobaccoUseCounseled = 0 And mlTobacco_Counseled_WLID > 0 Then
                'set dfDirtyflag = Delete
                'Tobacco use Counseled has been deselected (unchecked)
                'Data is already in the Templ_workload table, but we need to mark the
                '  selection as deleted.
                If Not UpdateWorkload(mlTobacco_Counseled_WLID, dfdelete, mvarDentalEncounterID, TU_Counseled_WLcode, 1) Then
                    GenerateGeneralWL_ID = False
                    sTmpMsg = "Error. Could not correctly update the Tobacco Counseled Workload Credits. "
                End If
            End If
        End If
        mbTobacco_Counseled_WL_chg = False
        
        
        '***************************************
        'Update the Workload for Oral Exam
        
        If mbOralExam_WLID_chg = True Then
        
            lTmpOE_WLcode = CalcOralExamWorkloadCode(mintOralExam)
              
            If mintOralExam <> NoOralExamValue And mlOralExamWLID = 0 Then
                'One of the OralExam types has been selected in the GUI but no record
                '  has been entered into DB to represent this.
                If Not UpdateWorkload(mlOralExamWLID, dfAdd, mvarDentalEncounterID, lTmpOE_WLcode, 1) Then
                    GenerateGeneralWL_ID = False
                    sTmpMsg = sTmpMsg & "Error. Could not correctly update the OralExam Workload Credits. dfAdd failed. "
                End If
            ElseIf mintOralExam <> NoOralExamValue And mlOralExamWLID > 0 Then
                'One of the OralExam types has already been selected previously, data has already
                '  been entered in to DB. No Change
                If Not UpdateWorkload(mlOralExamWLID, dfUpdate, mvarDentalEncounterID, lTmpOE_WLcode, 1) Then
                    GenerateGeneralWL_ID = False
                    sTmpMsg = sTmpMsg & "Error. Could not correctly update the OralExam Workload Credits. dfUpdate failed. "
                End If
            
            'Changed the line below for SCR 35144
            'ElseIf mlOralExamWLID > 0 And mbOralExam_WL_Apply = False And mintOralExam = 0 Then
            ElseIf mintOralExam = NoOralExamValue And mlOralExamWLID > 0 Then
                'Data is already in the Templ_workload table, but we need to mark the
                '  selection as deleted.
                If Not UpdateWorkload(mlOralExamWLID, dfdelete, mvarDentalEncounterID, lTmpOE_WLcode, 1) Then
                    GenerateGeneralWL_ID = False
                    sTmpMsg = sTmpMsg & "Error. Could not correctly update the OralExam Workload Credits. dfDelete failed.  "
                End If
            End If
            
            mbOralExam_WLID_chg = False
        End If
''### 1 ################################################################################
        'Evaluation workload
        
        'Sick call
        If mbSick_Call_WL_chg = True Then
            sTmpMsg = "Error. Could not correctly update the Sick Call Workload Credits."
            If mintSick_Call = -1 And mlSick_Call_WLID = 0 Then
                'set dfDirtyflag = dfAdd
                'Has been selected in the GUI but no record has been entered into DB to represent this.
                If Not UpdateWorkload(mlSick_Call_WLID, dfAdd, mvarDentalEncounterID, Sick_Call_WLcode, 1) Then
                    GenerateGeneralWL_ID = False
                    sTmpMsg = sTmpMsg + "Error. Could not correctly update the Sick Call Workload Credits. dfAdd failed."
                End If
            ElseIf mintSick_Call = -1 And mlSick_Call_WLID > 0 Then
                'set dfDirtyflag = dfUpdate
                'Has already been selected previously, data has already been entered in to DB. Updates the Deleted column in the Templ_Workload table (as not deleted).
                If Not UpdateWorkload(mlSick_Call_WLID, dfUpdate, mvarDentalEncounterID, Sick_Call_WLcode, 1) Then
                    GenerateGeneralWL_ID = False
                    sTmpMsg = sTmpMsg + "Error. Could not correctly update the Sick Call Workload Credits. dfUpdate failed."
                End If
            ElseIf mlSick_Call_WLID > 0 Then
                'set dfDirtyflag = Delete
                'Has been deselected (unchecked), Data is already in the Templ_workload table, but we need to mark the selection as deleted.
                If Not UpdateWorkload(mlSick_Call_WLID, dfdelete, mvarDentalEncounterID, Sick_Call_WLcode, 1) Then
                    GenerateGeneralWL_ID = False
                    sTmpMsg = sTmpMsg + "Error. Could not correctly update the Sick Call Workload Credits. dfDelete failed."
                End If
            End If
            'Note: The condition of: ElseIf mintSick_Call = 0 And mlSick_Call_WLID < 0 should never
            'happen. This would mean that the checkbox was deselected and the _WL var had no value.
            'If the checkbox was at one point -1, means it was selected, then _WL had to have a value.
            '_WL get set thru UpdateWorkload function when a dfAdd is used. This
            'method applies to all WorkLoad check boxes listed below.
            
            mbSick_Call_WL_chg = False
        End If
        
        'After hours care
        If mbAfter_HoursCare_chg = True Then
            sTmpMsg = "Error. Could not correctly update the After Hours Care Workload Credits."
            If mintAfter_HoursCare = -1 And mlAfter_HoursCare_WLID = 0 Then
                'set dfDirtyflag = dfAdd
                'Has been selected in the GUI but no record has been entered into DB to represent this.
                If Not UpdateWorkload(mlAfter_HoursCare_WLID, dfAdd, mvarDentalEncounterID, After_HoursCare_WLcode, 1) Then
                    GenerateGeneralWL_ID = False
                    sTmpMsg = sTmpMsg + "Error. Could not correctly update the After HoursCare Workload Credits. dfAdd failed."
                End If
            ElseIf mintAfter_HoursCare = -1 And mlAfter_HoursCare_WLID > 0 Then
                'set dfDirtyflag = dfUpdate
                'Has already been selected previously, data has already been entered in to DB. Updates the Deleted column in the Templ_Workload table (as not deleted).
                If Not UpdateWorkload(mlAfter_HoursCare_WLID, dfUpdate, mvarDentalEncounterID, After_HoursCare_WLcode, 1) Then
                    GenerateGeneralWL_ID = False
                    sTmpMsg = sTmpMsg + "Error. Could not correctly update the After HoursCare Workload Credits. dfUpdate failed."
                End If
            ElseIf mintAfter_HoursCare = 0 And mlAfter_HoursCare_WLID > 0 Then
                'set dfDirtyflag = Delete
                'Has been deselected (unchecked), Data is already in the Templ_workload table, but we need to mark the selection as deleted.
                If Not UpdateWorkload(mlAfter_HoursCare_WLID, dfdelete, mvarDentalEncounterID, After_HoursCare_WLcode, 1) Then
                    GenerateGeneralWL_ID = False
                    sTmpMsg = sTmpMsg + "Error. Could not correctly update the After HoursCare Workload Credits. dfDelete failed."
                End If
            End If
            
            mbAfter_HoursCare_chg = False
        End If
    
        'Consultation Response
        If mbConsultation_Response_chg = True Then
            sTmpMsg = "Error. Could not correctly update the Consultation Response Workload Credits."
            If mintConsultation_Response = -1 And mlConsultation_Response_WLID = 0 Then
                'set dfDirtyflag = dfAdd
                'Has been selected in the GUI but no record has been entered into DB to represent this.
                If Not UpdateWorkload(mlConsultation_Response_WLID, dfAdd, mvarDentalEncounterID, mintConsultation_Response_WLcode, 1) Then
                    GenerateGeneralWL_ID = False
                    sTmpMsg = sTmpMsg + "Error. Could not correctly insert the Consultation Response Workload Credits. dfAdd failed."
                End If
            ElseIf mintConsultation_Response = -1 And mlConsultation_Response_WLID > 0 Then
                'set dfDirtyflag = dfUpdate
                'Has already been selected previously, data has already been entered in to DB. Updates the Deleted column in the Templ_Workload table (as not deleted).
                If Not UpdateWorkload(mlConsultation_Response_WLID, dfUpdate, mvarDentalEncounterID, mintConsultation_Response_WLcode, 1) Then
                    GenerateGeneralWL_ID = False
                    sTmpMsg = sTmpMsg + "Error. Could not correctly update the Consultation Response Workload Credits. dfUpdate failed."
                End If
            ElseIf mintConsultation_Response = 0 And mlConsultation_Response_WLID > 0 Then
                'set dfDirtyflag = Delete
                'Has been deselected (unchecked), Data is already in the Templ_workload table, but we need to mark the selection as deleted.
                If Not UpdateWorkload(mlConsultation_Response_WLID, dfdelete, mvarDentalEncounterID, mintConsultation_Response_WLcode, 1) Then
                    GenerateGeneralWL_ID = False
                    sTmpMsg = sTmpMsg + "Error. Could not correctly delete the Consultation Response Workload Credits. dfDelete failed."
                End If
            End If
            
            mbConsultation_Response_chg = False
        End If

    End If

    GenerateGeneralWL_ID = True

Exit Function

ErrHnd:
    mbTobacco_Counseled_WL_chg = False
    mbOralExam_WLID_chg = False
    GenerateGeneralWL_ID = False
    goCWShared.ShowVBError Err.Number, Err.Description & ": " & sTmpMsg, Err.Source & "." & sPROC_NAME, App.Title, vbCritical

End Function
'*******************************************************************
'Function: clsISection_SaveSection
'Author: Sean McAvoy
'Date: 12/4/2001
'Returns: True/False
'Purpose: Saves the data to the table
'*******************************************************************
Private Function clsISection_SaveSection() As Boolean

Dim sSQL As String
Dim mstrOralExamNoteTMP As String
Dim mstrTobaccoUseNoteTMP As String
Dim mstrRevealsNoteTMP As String
Dim sTmpMsg As String

Dim mstrPostOperativeTMP As String
Dim mstrProgressNoteTMP As String
Dim mstrChief_ComplaintTMP As String
Dim mstrMedical_HistoryTMP As String
Dim mstrDental_HistoryTMP As String
Dim mstrPresent_ProblemTMP As String

Const sPROC_NAME As String = "SaveSection"

On Error GoTo ErrHnd

    If mvarDentalEncounterID = 0 Or mvarPatientID = 0 Then
        clsISection_SaveSection = False
        Exit Function
    End If

    If dfDirtyFlag = dfNoAction Then
        clsISection_SaveSection = True
        Exit Function
    Else
        'Set the WorkLoad Values
        If Not GenerateGeneralWL_ID(sTmpMsg) Then
            MsgBxARMd sTmpMsg, vbCritical, "Workload Credit Update Failed"
        End If
    End If
        
    'Format the String variables
    mstrPostOperativeTMP = Trim(mstrPostOperative)
    mstrProgressNoteTMP = Trim(mstrProgressNote)
    mstrChief_ComplaintTMP = Trim(mstrChief_Complaint)
    mstrMedical_HistoryTMP = Trim(mstrMedical_History)
    mstrDental_HistoryTMP = Trim(mstrDental_History)
    mstrPresent_ProblemTMP = Trim(mstrPresent_Problem)
    
    mstrOralExamNoteTMP = Trim(mstrOralExamNote)
    mstrTobaccoUseNoteTMP = Trim(mstrTobaccoUseNote)
    
    If mstrOralExamNoteTMP <> "" Then mstrOralExamNoteTMP = RQ(mstrOralExamNoteTMP)
    If mstrTobaccoUseNoteTMP <> "" Then mstrTobaccoUseNoteTMP = RQ(mstrTobaccoUseNoteTMP)
    If mstrPostOperativeTMP <> "" Then mstrPostOperativeTMP = RQ(mstrPostOperativeTMP)
    If mstrProgressNoteTMP <> "" Then mstrProgressNoteTMP = RQ(mstrProgressNoteTMP)
    If mstrChief_ComplaintTMP <> "" Then mstrChief_ComplaintTMP = RQ(mstrChief_ComplaintTMP)
    If mstrMedical_HistoryTMP <> "" Then mstrMedical_HistoryTMP = RQ(mstrMedical_HistoryTMP)
    If mstrDental_HistoryTMP <> "" Then mstrDental_HistoryTMP = RQ(mstrDental_HistoryTMP)
    If mstrPresent_ProblemTMP <> "" Then mstrPresent_ProblemTMP = RQ(mstrPresent_ProblemTMP)

'    'Save the data to the xml doc
'    If mvarXMLElement Is Nothing Then
'        Set mvarXMLElement = mvarXMLPersist.ExamXML.XMLDoc.selectSingleNode("//" & XML_ELEMENT_NAME)
'
'        If mvarXMLElement Is Nothing Then
'            'Create the node
'            Set mvarXMLElement = mvarXMLPersist.ExamXML.XMLDoc.createElement(XML_ELEMENT_NAME)
'            mvarXMLPersist.ExamXML.XMLDoc.documentElement.appendChild mvarXMLElement
'        End If
'    End If
'
'    'Set the attributes
'    mvarXMLElement.setAttribute "ORALEXAM", CStr(mintOralExam)
'    mvarXMLElement.setAttribute "ORALEXAM", CStr(mintOralExam)
'    mvarXMLElement.setAttribute "ORALEXAM_NOTE", XMLRQ(mstrOralExamNoteTMP)
'    mvarXMLElement.setAttribute "RISK_CARIES", CStr(mintRiskAssess_Caries)
'    mvarXMLElement.setAttribute "RISK_PERIO", CStr(mintRiskAssess_Perio)
'    mvarXMLElement.setAttribute "RISK_ORALCA", CStr(mintRiskAssess_OralCA)
'    mvarXMLElement.setAttribute "RISK_NOTE", XMLRQ(mstrRiskAssess_NoteTMP)
'    mvarXMLElement.setAttribute "TOBACCO_USE", CStr(mintTobaccoUse)
'    mvarXMLElement.setAttribute "TOBACCO_COUNSELED", CStr(mintTobaccoUseCounseled)
'    mvarXMLElement.setAttribute "TOBACCO_NOTE", XMLRQ(mstrTobaccoUseNoteTMP)
'    mvarXMLElement.setAttribute "STE_WITHINLIMITS", CStr(mintSTE_WithinLimits)
'    mvarXMLElement.setAttribute "REVEAL_NOTE", XMLRQ(mstrRevealsNoteTMP)
'    mvarXMLElement.setAttribute "ORALEXAM_WL", CStr(mlOralExamWLID)
'    mvarXMLElement.setAttribute "TOBACCO_COUNSELED_WL", CStr(mlTobacco_Counseled_WLID)
'    mvarXMLElement.setAttribute "GC_SUPP_PROCS", XMLRQ(mstrGCSupp_Procedures)

    If dfDirtyFlag = dfAdd Then
        If mvarDentalEncounterID = 0 Then
            clsISection_SaveSection = False
            Exit Function
        End If

        'The record does not exist, so create an insert statement
        sSQL = "INSERT INTO TEMPL_EXAMGENERAL (" _
            & " DENTAL_ENCOUNTER_ID, ORALEXAM, ORALEXAM_NOTE," _
            & " CHIEF_COMPLAINT, MEDICAL_HISTORY, MEDICAL_HISTORY_NC, DENTAL_HISTORY, " _
            & " DENTAL_HISTORY_NC, PRESENT_PROBLEM, POSTOPERATIVE, " _
            & " PROGRESSNOTE," _
            & " TOBACCO_USE, TOBACCO_COUNSELED, TOBACCO_NOTE, ORALEXAM_WL, TOBACCO_COUNSELED_WL," _
            & " SICK_CALL, ORAL_SURGERY, PERIODONTAL, PROSTHODONTICS, ENDODONTICS, ORTHODONTICS,  " _
            & " OROFACIAL_PAIN, AFTER_HOURSCARE, CONSULTATION_RESPONSE, Pediatric, SICK_CALL_WL, AFTER_HOURSCARE_WL, " _
            & " Consultation_Response_WL, GC_SUPP_PROCS) "
            
        sSQL = sSQL & " VALUES" _
            & " (" & CStr(mvarDentalEncounterID) & ", " _
            & CStr(mintOralExam) & ", " _
            & " '" & mstrOralExamNoteTMP & "', " _
            & " '" & mstrChief_ComplaintTMP & "', " _
            & " '" & mstrMedical_HistoryTMP & "', " _
            & CStr(mintMedical_History_NC) & ", " _
            & " '" & mstrDental_HistoryTMP & "', " _
            & CStr(mintDental_History_NC) & ", " _
            & " '" & mstrPresent_ProblemTMP & "', " _
            & " '" & mstrPostOperativeTMP & "', " _
            & " '" & mstrProgressNoteTMP & "', " _
            & CStr(mintTobaccoUse) & ", " _
            & CStr(mintTobaccoUseCounseled) & ", " _
            & " '" & mstrTobaccoUseNoteTMP & "', " _
            & CStr(mlOralExamWLID) & ", " _
            & CStr(mlTobacco_Counseled_WLID) & ", "
            
         sSQL = sSQL _
            & CStr(mintSick_Call) & ", " _
            & CStr(mintOral_Surgery) & ", " _
            & CStr(mintPeriodontal) & ", " _
            & CStr(mintProsthodontics) & ", " _
            & CStr(mintEndodontics) & ", " _
            & CStr(mintOrthodontics) & ", " _
            & CStr(mintOrofacialPain) & ", " _
            & CStr(mintAfter_HoursCare) & ", " _
            & CStr(mintConsultation_Response) & ", " _
            & CStr(mintPediatric) & ", " _
            & CStr(mlSick_Call_WLID) & ", " _
            & CStr(mlAfter_HoursCare_WLID) & ", " _
            & CStr(mlConsultation_Response_WLID) & ", "
            
        sSQL = sSQL _
            & " '" & mstrGCSupp_Procedures & "') "
            
    ElseIf dfDirtyFlag = dfUpdate Then
        'The record exists, so create an update statement
        sSQL = "UPDATE TEMPL_EXAMGENERAL SET " & _
                "ORALEXAM = " & CStr(mintOralExam) & ", " & _
                "ORALEXAM_NOTE = '" & mstrOralExamNoteTMP & "', " & _
                "CHIEF_COMPLAINT = '" & mstrChief_ComplaintTMP & "', " & _
                "MEDICAL_HISTORY = '" & mstrMedical_HistoryTMP & "', " & _
                "MEDICAL_HISTORY_NC = " & CStr(mintMedical_History_NC) & ", " & _
                "DENTAL_HISTORY = '" & mstrDental_HistoryTMP & "', " & _
                "DENTAL_HISTORY_NC = " & CStr(mintDental_History_NC) & ", " & _
                "PRESENT_PROBLEM = '" & mstrPresent_ProblemTMP & "', " & _
                "POSTOPERATIVE = '" & mstrPostOperativeTMP & "', " & _
                "PROGRESSNOTE = '" & mstrProgressNoteTMP & "', " & _
                "TOBACCO_USE = " & CStr(mintTobaccoUse) & ", " & _
                "TOBACCO_COUNSELED = " & CStr(mintTobaccoUseCounseled) & ", " & _
                "TOBACCO_NOTE = '" & mstrTobaccoUseNoteTMP & "', " & _
                "ORALEXAM_WL = " & CStr(mlOralExamWLID) & ", " & _
                "TOBACCO_COUNSELED_WL = " & CStr(mlTobacco_Counseled_WLID) & ", "
                
        sSQL = sSQL & _
                "SICK_CALL = " + CStr(mintSick_Call) + ", " + _
                "ORAL_SURGERY = " + CStr(mintOral_Surgery) + ", " + _
                "PERIODONTAL = " + CStr(mintPeriodontal) + ", " + _
                "PROSTHODONTICS = " + CStr(mintProsthodontics) + ", " + _
                "ENDODONTICS = " + CStr(mintEndodontics) + ", " + _
                "ORTHODONTICS = " + CStr(mintOrthodontics) + ", " + _
                "Orofacial_Pain = " + CStr(mintOrofacialPain) + ", " + _
                "AFTER_HOURSCARE = " + CStr(mintAfter_HoursCare) + ", " + _
                "CONSULTATION_RESPONSE = " + CStr(mintConsultation_Response) + ", " & _
                "Pediatric = " + CStr(mintPediatric) + ", " & _
                "SICK_CALL_WL = " + CStr(mlSick_Call_WLID) + ", " + _
                "After_HoursCare_WL = " + CStr(mlAfter_HoursCare_WLID) & ", " & _
                "Consultation_Response_WL = " + CStr(mlConsultation_Response_WLID) & ", "

        sSQL = sSQL & _
                "GC_SUPP_PROCS = '" & mstrGCSupp_Procedures & "' " & _
                "Where (DENTAL_ENCOUNTER_ID = " & CStr(mvarDentalEncounterID) & ")"

    ElseIf dfDirtyFlag = dfdelete Then
        'Mark the record as deleted
'        sSQL = "DELETE FROM TEMPL_EXAMGENERAL " & _
'               "Where (DENTAL_ENCOUNTER_ID = " & CStr(mvarDentalEncounterID) & ")"
    End If

    If sSQL <> "" Then goconn.ExecuteSQL sSQL
    
    bSavedToDB = True
    dfDirtyFlag = dfNoAction
    
    clsISection_SaveSection = True

Exit Function

ErrHnd:
    clsISection_SaveSection = False
    goCWShared.ShowVBError Err.Number, Err.Description, Err.Source & "." & sPROC_NAME, App.Title, vbCritical
End Function
Public Function ResetSBE(ByVal PatientID As Long, ByVal DentalEncounterID) As Boolean
Const sPROC_NAME = "ResetSBE"
On Error GoTo ErrHnd
Dim sSQL As String

sSQL = "UPDATE DENTAL_PATIENT SET SBE_REQUIRE_FLAG = 0 where PATIENT_UNIT_ID = " & CStr(PatientID)
goconn.ExecuteSQL sSQL

sSQL = "UPDATE DENTAL_ENCOUNTER SET SBE = 4, SBE_COMMENTS = '' WHERE PATIENT_UNIT_ID = " & CStr(PatientID) & _
    " AND DENTAL_ENCOUNTER_ID = " & CStr(DentalEncounterID)
goconn.ExecuteSQL sSQL


ResetSBE = True
Exit Function
ErrHnd:
ResetSBE = False
    goCWShared.ShowVBError Err.Number, Err.Description, Err.Source & "." & sPROC_NAME, App.Title, vbCritical
End Function

'*******************************************************************
'Function: clsISection_LoadSection
'Author: Sean McAvoy
'Date: 12/4/2001
'Returns: True/False
'Purpose: This loads the data from the db
'*******************************************************************
Private Function clsISection_LoadSection(ByVal PatientID As Long, ByVal DentalEncounterID As Long) As Boolean

Dim sSQL As String, rs As Recordset, rsPrev As Recordset, rs2 As Recordset, rs3 As Recordset
Dim XMLDoc As MSXML2.DOMDocument30

Const sPROC_NAME = "clsExamGeneral.clsISection_LoadSection"
Static bSBE_FirstLoad As Boolean
On Error GoTo ErrHnd

    mvarPatientID = PatientID
    mvarDentalEncounterID = DentalEncounterID
    If mvarDentalEncounterID = 0 Or mvarPatientID = 0 Then
        clsISection_LoadSection = False
        Exit Function
    End If
        
    'KDunne 4/22/2002
    'SCR 8525 - Load DRC Info
    LoadDRC mvarDentalEncounterID
    
    'SCR 17481
    
'''        sSQL = "Select SBE_REQUIRE_FLAG from DENTAL_PATIENT where PATIENT_UNIT_ID = " & CStr(mvarPatientID)
'''        Set rs = goconn.OpenRecordset(sSQL)
        
        mbUpdateSBE = False
        If mobjSBE Is Nothing Then
             mintSBE_Require = 0
'           Commended the following out for SCR 21783
'            If DentalEncounterID <> 1717 Then 'NotesPlus test encounter, can remove later
'                cwobjCWComm.Message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, cwlMyProgId
'                MsgBxARMd "Warning. The Patient ID could not be found in the Dental_Parent table in " & _
'                    "order to update the SBE_REQUIRED_FLAG.", , "SBE Error"
'                cwobjCWComm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, cwlMyProgId
'            End If
        Else
'''            mintSBE_Require = Val(rs("SBE_REQUIRE_FLAG") & "")
            mintSBE_Require = mobjSBE.SBERequiredFlag
            If bSBE_FirstLoad = False Then
                bSBE_FirstLoad = True
                If mintSBE_Require = -1 Then
                    mbSBEcontrolLocked = True
                End If
            End If
        End If
'''        rs.Close
        
    'Load the values
    
'    'Load the XML
'    Set mvarXMLElement = mvarXMLPersist.ExamXML.XMLDoc.selectSingleNode("//" & XML_ELEMENT_NAME)
'    If mvarXMLElement Is Nothing Then
'        'Init the vars
'        mintOralExam = 0
'        mstrOralExamNote = ""
'        mintRiskAssess_Caries = 0
'        mintRiskAssess_Perio = 0
'        mintRiskAssess_OralCA = 0
'        mstrRiskAssess_Note = ""
'        mintTobaccoUse = 0
'        mintTobaccoUseCounseled = 0
'        mstrTobaccoUseNote = ""
'        mintSTE_WithinLimits = 0
'        mstrRevealsNote = ""
'        mlOralExamWLID = 0
'        mlTobacco_Counseled_WLID = 0
'        mstrGCSupp_Procedures = ""
'    Else
'        mintOralExam = mvarXMLElement.Attributes.getNamedItem("ORALEXAM").nodeValue
'        mstrOralExamNote = XMLRRQ(mvarXMLElement.Attributes.getNamedItem("ORALEXAM_NOTE").nodeValue)
'        mintRiskAssess_Caries = mvarXMLElement.Attributes.getNamedItem("RISK_CARIES").nodeValue
'        mintRiskAssess_Perio = mvarXMLElement.Attributes.getNamedItem("RISK_PERIO").nodeValue
'        mintRiskAssess_OralCA = mvarXMLElement.Attributes.getNamedItem("RISK_ORALCA").nodeValue
'        mstrRiskAssess_Note = XMLRRQ(mvarXMLElement.Attributes.getNamedItem("RISK_NOTE").nodeValue)
'        mintTobaccoUse = mvarXMLElement.Attributes.getNamedItem("TOBACCO_USE").nodeValue
'        mintTobaccoUseCounseled = mvarXMLElement.Attributes.getNamedItem("TOBACCO_COUNSELED").nodeValue
'        mstrTobaccoUseNote = XMLRRQ(mvarXMLElement.Attributes.getNamedItem("TOBACCO_NOTE").nodeValue)
'        mintSTE_WithinLimits = mvarXMLElement.Attributes.getNamedItem("STE_WITHINLIMITS").nodeValue
'        mstrRevealsNote = XMLRRQ(mvarXMLElement.Attributes.getNamedItem("REVEAL_NOTE").nodeValue)
'        mlOralExamWLID = mvarXMLElement.Attributes.getNamedItem("ORALEXAM_WL").nodeValue
'        mlTobacco_Counseled_WLID = mvarXMLElement.Attributes.getNamedItem("TOBACCO_COUNSELED_WL").nodeValue
'        mstrGCSupp_Procedures = XMLRRQ(mvarXMLElement.Attributes.getNamedItem("GC_SUPP_PROCS").nodeValue)
'    End If
    
    If gbUseStoredProc = True Then
        Dim oCmd As New CHCSII_Command.CHCSII_Cmd

        oCmd.SPName = "Dent_Notes_ExamEval_Load"

        Call oCmd.AddParam(False, False, "textString", mvarDentalEncounterID)
        Call oCmd.AddParam(True, True, "", "")

        Set rs = goconn.OpenRecordsetSP(oCmd)

        If rs.EOF Then
            bSavedToDB = False
            dfDirtyFlag = dfAdd
            mintOralExam = 0
            mstrOralExamNote = ""
            mintMedical_History_NC = 0
            mintDental_History_NC = 0
            mstrPostOperative = ""
            mstrProgressNote = ""
            mintTobaccoUse = NoTobaccoUseValue
            mintTobaccoUseCounseled = 0
            mstrTobaccoUseNote = ""
            mlOralExamWLID = 0
            mlTobacco_Counseled_WLID = 0
            mstrTobaccoUseNote = ""
            mstrGCSupp_Procedures = ""
            mstrChief_Complaint = ""
            mstrMedical_History = ""
            
            mstrDental_History = ""
            mstrPresent_Problem = ""
            
            mstrPostOperative = ""
            mstrProgressNote = ""
        
            mintSick_Call = 0
            mintOral_Surgery = 0
            mintPeriodontal = 0
            mintProsthodontics = 0
            mintEndodontics = 0
            mintOrthodontics = 0
            mintOrofacialPain = 0
            mintAfter_HoursCare = 0
            mintConsultation_Response = 0
            mlSick_Call_WLID = 0
            mlAfter_HoursCare_WLID = 0
            mlConsultation_Response_WLID = 0
            mintPediatric = 0
            
            rs.Close
            Set rs = Nothing
            clsISection_LoadSection = True
            Exit Function
        Else
            'RECORD FOUND...
            bSavedToDB = True
            dfDirtyFlag = dfNoAction

            'Set the values from the db including "" to allow for nulls

            mintOralExam = Val(rs("ORALEXAM") & "")
            mstrOralExamNote = rs("ORALEXAM_NOTE") & ""
            mintMedical_History_NC = Val(rs("MEDICAL_HISTORY_NC") & "")
            mintDental_History_NC = Val(rs("DENTAL_HISTORY_NC") & "")
            mintTobaccoUse = Val(rs("TOBACCO_USE") & "")
            mintTobaccoUseCounseled = Val(rs("TOBACCO_COUNSELED") & "")
            mlOralExamWLID = Val(rs("ORALEXAM_WL") & "")
            mlTobacco_Counseled_WLID = Val(rs("TOBACCO_COUNSELED_WL") & "")
            mstrTobaccoUseNote = rs("TOBACCO_NOTE") & ""
            mstrGCSupp_Procedures = rs("GC_SUPP_PROCS") & ""
            mstrChief_Complaint = rs("CHIEF_COMPLAINT") & ""
            mstrMedical_History = rs("MEDICAL_HISTORY") & ""
            
            mstrDental_History = rs("DENTAL_HISTORY") & ""
            mstrPresent_Problem = rs("PRESENT_PROBLEM") & ""
            
            mstrPostOperative = rs("POSTOPERATIVE") & ""
            mstrProgressNote = rs("PROGRESSNOTE") & ""
            
            mintSick_Call = Val(rs("SICK_CALL") & "")
            mintOral_Surgery = Val(rs("ORAL_SURGERY") & "")
            mintPeriodontal = Val(rs("PERIODONTAL") & "")
            mintProsthodontics = Val(rs("PROSTHODONTICS") & "")
            mintEndodontics = Val(rs("ENDODONTICS") & "")
            mintOrthodontics = Val(rs("ORTHODONTICS") & "")
            mintOrofacialPain = Val(rs("Orofacial_Pain") & "")
            mintAfter_HoursCare = Val(rs("AFTER_HOURSCARE") & "")
            mintConsultation_Response = Val(rs("CONSULTATION_RESPONSE") & "")
            mlSick_Call_WLID = Val(rs("SICK_CALL_WL") & "")
            mlAfter_HoursCare_WLID = Val(rs("After_HoursCare_WL") & "")
            mlConsultation_Response_WLID = Val(rs("Consultation_Response_WL") & "")
            mintPediatric = Val(rs("Pediatric") & "")
            
            rs.Close
        End If

    Else

        'Because of the issues with Tuxedo (4k limit), three sql calls will
        'need to be made to retrieve the several 2k note fields.
        sSQL = "SELECT DENTAL_ENCOUNTER_ID, ORALEXAM, ORALEXAM_NOTE, TOBACCO_USE, TOBACCO_COUNSELED," _
                & " ORALEXAM_WL, TOBACCO_COUNSELED_WL, TOBACCO_NOTE, SICK_CALL, ORAL_SURGERY, PERIODONTAL," _
                & " PROSTHODONTICS, ENDODONTICS, ORTHODONTICS, Pediatric, Orofacial_Pain, AFTER_HOURSCARE," _
                & " CONSULTATION_RESPONSE, GC_SUPP_PROCS, CHIEF_COMPLAINT, MEDICAL_HISTORY, DENTAL_HISTORY," _
                & " PRESENT_PROBLEM, POSTOPERATIVE, ProgressNote , Medical_History_NC, Dental_History_NC," _
                & " Sick_Call_WL, After_HoursCare_WL, Consultation_Response_WL" _
                & " FROM TEMPL_EXAMGENERAL Where " _
                & " DENTAL_ENCOUNTER_ID = " & CStr(mvarDentalEncounterID)

        Set rs = goconn.OpenRecordset(sSQL)

        If rs.EOF Then
            bSavedToDB = False
            dfDirtyFlag = dfAdd
            mintOralExam = 0
            mstrOralExamNote = ""
            mintMedical_History_NC = 0
            mintDental_History_NC = 0
            mintTobaccoUse = NoTobaccoUseValue
            mintTobaccoUseCounseled = 0
            mstrTobaccoUseNote = ""
            mlOralExamWLID = 0
            mlTobacco_Counseled_WLID = 0
            mstrPostOperative = ""
            mstrProgressNote = ""
            mintSick_Call = 0
            mintOral_Surgery = 0
            mintPeriodontal = 0
            mintProsthodontics = 0
            mintEndodontics = 0
            mintOrthodontics = 0
            mintOrofacialPain = 0
            mintAfter_HoursCare = 0
            mintConsultation_Response = 0
            mlSick_Call_WLID = 0
            mlAfter_HoursCare_WLID = 0
            mlConsultation_Response_WLID = 0
            mintPediatric = 0
            
            rs.Close
            Set rs = Nothing
            clsISection_LoadSection = True
            Exit Function
        Else
            'RECORD FOUND...
            bSavedToDB = True
            dfDirtyFlag = dfNoAction

            'Set the values from the db including "" to allow for nulls

            mintOralExam = Val(rs("ORALEXAM") & "")
            mstrOralExamNote = rs("ORALEXAM_NOTE") & ""
            mintMedical_History_NC = Val(rs("MEDICAL_HISTORY_NC") & "")
            mintDental_History_NC = Val(rs("DENTAL_HISTORY_NC") & "")
            mintTobaccoUse = Val(rs("TOBACCO_USE") & "")
            mintTobaccoUseCounseled = Val(rs("TOBACCO_COUNSELED") & "")
            mlOralExamWLID = Val(rs("ORALEXAM_WL") & "")
            mlTobacco_Counseled_WLID = Val(rs("TOBACCO_COUNSELED_WL") & "")
            mstrTobaccoUseNote = rs("TOBACCO_NOTE") & ""
            mstrGCSupp_Procedures = rs("GC_SUPP_PROCS") & ""
            mstrChief_Complaint = rs("CHIEF_COMPLAINT") & ""
            mstrMedical_History = rs("MEDICAL_HISTORY") & ""
            
            mstrDental_History = rs("DENTAL_HISTORY") & ""
            mstrPresent_Problem = rs("PRESENT_PROBLEM") & ""
            
            mstrPostOperative = rs("POSTOPERATIVE") & ""
            mstrProgressNote = rs("PROGRESSNOTE") & ""
            
            mintSick_Call = Val(rs("SICK_CALL") & "")
            mintOral_Surgery = Val(rs("ORAL_SURGERY") & "")
            mintPeriodontal = Val(rs("PERIODONTAL") & "")
            mintProsthodontics = Val(rs("PROSTHODONTICS") & "")
            mintEndodontics = Val(rs("ENDODONTICS") & "")
            mintOrthodontics = Val(rs("ORTHODONTICS") & "")
            mintOrofacialPain = Val(rs("Orofacial_Pain") & "")
            mintAfter_HoursCare = Val(rs("AFTER_HOURSCARE") & "")
            mintConsultation_Response = Val(rs("CONSULTATION_RESPONSE") & "")
            mlSick_Call_WLID = Val(rs("SICK_CALL_WL") & "")
            mlAfter_HoursCare_WLID = Val(rs("After_HoursCare_WL") & "")
            mlConsultation_Response_WLID = Val(rs("Consultation_Response_WL") & "")
            mintPediatric = Val(rs("Pediatric") & "")
            
            rs.Close

            sSQL = "SELECT TOBACCO_NOTE, GC_SUPP_PROCS FROM TEMPL_EXAMGENERAL Where " & _
            "DENTAL_ENCOUNTER_ID = " & CStr(mvarDentalEncounterID)
            Set rs = goconn.OpenRecordset(sSQL)
            If rs.EOF Then
                mstrTobaccoUseNote = ""
                mstrGCSupp_Procedures = ""
            Else
                mstrTobaccoUseNote = rs("TOBACCO_NOTE") & ""
                mstrGCSupp_Procedures = rs("GC_SUPP_PROCS") & ""
            End If
            rs.Close
        
        End If
    End If

    Set rs = Nothing
    
'    'Set the Enum to be in sync with variables...
'    Select Case mintOralExam
'        Case 0
'            Me.OralExamEnum = OralExamEN.NoOralExamValue
'        Case 1
'            Me.OralExamEnum = OralExamEN.Periodic
'        Case 2
'            Me.OralExamEnum = OralExamEN.InitialCharting
'        Case 3
'            Me.OralExamEnum = OralExamEN.Comprehensive
'    End Select
    
'    Select Case mintTobaccoUse
'        Case 0
'            Me.TobaccoUseEnum = TobaccoUseEN.NoTobaccoUseValue
'        Case 1
'            Me.TobaccoUseEnum = TobaccoUseEN.DontSmoke
'        Case 2
'            Me.TobaccoUseEnum = TobaccoUseEN.Smoke
'        Case 3
'            Me.TobaccoUseEnum = TobaccoUseEN.smokeless
'        Case 4
'            Me.TobaccoUseEnum = TobaccoUseEN.UseBoth
'    End Select
    
    clsISection_LoadSection = True
    
Exit Function

ErrHnd:
    Set rs = Nothing
    clsISection_LoadSection = False
    goCWShared.ShowVBError Err.Number, Err.Description, Err.Source & "." & sPROC_NAME, App.Title, vbCritical

End Function
Private Function UpdateDirtyFlag()

    If bSavedToDB = False Then
        dfDirtyFlag = dfAdd
    ElseIf dfDirtyFlag = dfNoAction Then
        dfDirtyFlag = dfUpdate
    End If
    
End Function

Private Sub Class_Initialize()

    dfDirtyFlag = dfNoAction
    mintOralExam = NoOralExamValue
    mstrOralExamNote = ""
    mintTobaccoUse = NoTobaccoUseValue
    mintTobaccoUseCounseled = 0
    mstrTobaccoUseNote = ""
    mlOralExamWLID = 0
    mlTobacco_Counseled_WLID = 0
    mintSBE_Require = 0
    mbUpdateSBE = False
    mbSBEcontrolLocked = False
    mstrSBE_Text = ""
    mstrChief_Complaint = ""
    mstrMedical_History = ""
    mintMedical_History_NC = 0
    mstrDental_History = ""
    mintDental_History_NC = 0
    mstrPresent_Problem = ""
    mstrPostOperative = ""
    mstrProgressNote = ""
    mintSick_Call = 0
    mintOral_Surgery = 0
    mintPeriodontal = 0
    mintProsthodontics = 0
    mintEndodontics = 0
    mintOrthodontics = 0
    mintOrofacialPain = 0
    mintAfter_HoursCare = 0
    mintConsultation_Response = 0
    mlSick_Call_WLID = 0
    mlAfter_HoursCare_WLID = 0
    mlConsultation_Response_WLID = 0

    
End Sub


'##################################################################
'##################################################################

'Private Function SetPSRvalues() As Boolean
'    On Error GoTo ErrTrap

'Public Type PSR_Control
'    Upper As New Collection
'    Lower As New Collection
'End Type

'
'    'Clear all previous values
'    Do While mPSR.Lower.Count > 0
'        mPSR.Lower.Remove 1
'    Loop
'    Do While mPSR.Upper.Count > 0
'        mPSR.Upper.Remove 1
'    Loop
'    With frmNotes
'        mPSR.Lower.Add .psrDTN.PSR.lowerRight, "R"
'        mPSR.Lower.Add .psrDTN.PSR.lowerCenter, "C" 'Labeled as Anterior
'        mPSR.Lower.Add .psrDTN.PSR.lowerLeft, "L"
'        mPSR.Upper.Add .psrDTN.PSR.upperRight, "R"
'        mPSR.Upper.Add .psrDTN.PSR.upperCenter, "C" 'Labeled as Anterior
'        mPSR.Upper.Add .psrDTN.PSR.upperLeft, "L"
'    End With
'    'Syntax is psr.Lower("R")... psr.Lower("L") to get values
'
'    SetPSRvalues = True
'    Exit Function
'ErrTrap:
'    SetPSRvalues = False
'End Function


'*************************************************************************************
'Sub:LoadDRC
'Author:KDunne
'Date:4/22/2002
'Purpose: Queries the Dental Encounter table and loads the DRC.
'*************************************************************************************
Public Sub LoadDRC(ByVal lngDentalEncounterId As Long)
    Dim strSql As String
    Dim objRS As ADODB.Recordset
    
    If gbUseStoredProc = True Then
        Dim oCmd As New CHCSII_Command.CHCSII_Cmd
        
        oCmd.SPName = "Dent_Notes_LoadDRC"

        Call oCmd.AddParam(False, False, "textString", lngDentalEncounterId)
        Call oCmd.AddParam(True, True, "", "")
        
        Set objRS = goconn.OpenRecordsetSP(oCmd)
        
    Else
    
        strSql = "Select DRC From Dental_Encounter Where Dental_Encounter_ID = " & lngDentalEncounterId
        
        Set objRS = goconn.OpenRecordset(strSql)
    End If
    
    If Not objRS.BOF And Not objRS.EOF Then
        mdblDRC = Val(objRS.Fields("DRC") & "")
    End If


FuncExit:
    Set objRS = Nothing
    Exit Sub
    
ErrHandler:
    GoTo FuncExit
End Sub

