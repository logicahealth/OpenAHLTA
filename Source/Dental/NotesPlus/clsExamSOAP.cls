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
Attribute VB_Name = "clsExamSOAP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
  'Class Module Created by Sean McAvoy
  'Hold values for the Exam/Eval SOAP Tab for the Dental Templates.

Option Explicit

Implements clsISection

'Used to determine if the data should be inserted or updated when modified.
Private bSavedToDB As Boolean

Private mvarPatientID As Long
Private mvarDentalEncounterID As Long

Private dfDirtyFlag As DirtyFlagEnum
'local variable(s) to hold property value(s)

Private mstrMax_Impressions As String
Private mstrMan_Impressions As String
Private mstrSummary As String
Private mintDiagnostic_Mounting As Integer

'SCR 8485
Private mintSTE_WithinLimits As Integer
Private mstrRevealsNote As String

Private mbPatientIsAdult As Boolean

'SCR 8499
Public Enum RiskAssessEN
    NoRiskAssessValue = 0
    low = 1
    Moderate = 2
    High = 3
End Enum
Public RiskAssessEnum As RiskAssessEN
Private mintRiskAssess_Caries As Integer
Private mintRiskAssess_Perio As Integer
Private mintRiskAssess_OralCA As Integer
Private mstrRiskAssess_Note As String
Public Property Let RiskAssess_Note(ByVal vData As String)
    mstrRiskAssess_Note = vData
    UpdateDirtyFlag
End Property

Public Property Get RiskAssess_Note() As String
    RiskAssess_Note = mstrRiskAssess_Note
End Property

Public Property Let RiskAssess_OralCA(ByVal vData As Integer)
    mintRiskAssess_OralCA = vData
    UpdateDirtyFlag
End Property

Public Property Get RiskAssess_OralCA() As Integer
    RiskAssess_OralCA = mintRiskAssess_OralCA
End Property

Public Property Let RiskAssess_Perio(ByVal vData As Integer)
    mintRiskAssess_Perio = vData
    UpdateDirtyFlag
End Property

Public Property Get RiskAssess_Perio() As Integer
    RiskAssess_Perio = mintRiskAssess_Perio
End Property

Public Property Let RiskAssess_Caries(ByVal vData As Integer)
    If mintRiskAssess_Caries = vData Then
        Exit Property
    End If
    mintRiskAssess_Caries = vData
    UpdateDirtyFlag
End Property

Public Property Get RiskAssess_Caries() As Integer
    RiskAssess_Caries = mintRiskAssess_Caries
End Property

Public Property Let RevealsNote(ByVal vData As String)
    mstrRevealsNote = vData
    UpdateDirtyFlag
End Property

Public Property Get RevealsNote() As String
    RevealsNote = mstrRevealsNote
End Property

Public Property Let STE_WithinLimits(ByVal vData As Integer)
    mintSTE_WithinLimits = vData
    UpdateDirtyFlag
End Property

Public Property Get STE_WithinLimits() As Integer
    STE_WithinLimits = mintSTE_WithinLimits
End Property


Public Property Let PatientIsAdult(ByVal vData As Boolean)
    mbPatientIsAdult = vData
End Property
Public Property Get PatientIsAdult() As Boolean
    PatientIsAdult = mbPatientIsAdult
End Property
Public Property Let Diagnostic_Mounting(ByVal vData As Integer)
    mintDiagnostic_Mounting = vData
    UpdateDirtyFlag
End Property
Public Property Get Diagnostic_Mounting() As Integer
    Diagnostic_Mounting = mintDiagnostic_Mounting
End Property
Public Property Let Summary(ByVal vData As String)
    mstrSummary = vData
    UpdateDirtyFlag
End Property
Public Property Get Summary() As String
    Summary = mstrSummary
End Property
Public Property Let Man_Impressions(ByVal vData As String)
    mstrMan_Impressions = vData
    UpdateDirtyFlag
End Property
Public Property Get Man_Impressions() As String
    Man_Impressions = mstrMan_Impressions
End Property
Public Property Let Max_Impressions(ByVal vData As String)
    mstrMax_Impressions = vData
    UpdateDirtyFlag
End Property
Public Property Get Max_Impressions() As String
    Max_Impressions = mstrMax_Impressions
End Property
Private Property Get clsISection_DentalEncounterID() As Long
    clsISection_DentalEncounterID = mvarDentalEncounterID
End Property


Private Property Get clsISection_PatientID() As Long
    
    clsISection_PatientID = mvarPatientID

End Property



Private Function clsISection_GetRTFNote() As String

Dim sTemp As String
 Dim sNote As String
 Dim sTxt As String
 
 Const sPROC_NAME = "clsISection_GetRTFNote"
 On Error GoTo errTrap
    
    If mstrMax_Impressions = "" And mstrMan_Impressions = "" And _
        mintSTE_WithinLimits = 0 And _
        mstrRevealsNote = "" And _
        mintRiskAssess_Caries = 0 And _
        mintRiskAssess_Perio = 0 And _
        mintRiskAssess_OralCA = 0 And _
        mstrRiskAssess_Note = "" And _
        mstrSummary = "" And mintDiagnostic_Mounting = 0 Then
    
        'No need to update note..
        sNote = ""
        clsISection_GetRTFNote = sNote
        Exit Function
    End If

    'SCR 8499 Risk Assessment
    If mintRiskAssess_Caries > 0 Or mintRiskAssess_Perio > 0 Or _
    mintRiskAssess_OralCA > 0 Or mstrRiskAssess_Note <> "" Then
        sNote = sNote + crtfStartControlName & "Risk Assessment: " '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
        Select Case mintRiskAssess_Caries
            Case 1
                sNote = sNote + "Caries risk is LOW; "
            Case 2
                sNote = sNote + "Caries risk is MODERATE; "
            Case 3
                sNote = sNote + "Caries risk is HIGH; "
        End Select
        Select Case mintRiskAssess_Perio
            Case 1
                sNote = sNote + "Perio risk is LOW; "
            Case 2
                sNote = sNote + "Perio risk is MODERATE; "
            Case 3
                sNote = sNote + "Perio risk is HIGH; "
        End Select
        Select Case mintRiskAssess_OralCA
            Case 1
                sNote = sNote + "Oral CA risk is LOW; "
            Case 2
                sNote = sNote + "Oral CA risk is MODERATE; "
            Case 3
                sNote = sNote + "Oral CA risk is HIGH; "
        End Select
        If mstrRiskAssess_Note <> "" Then
            sNote = sNote & "Risk Assess Comment: " + mstrRiskAssess_Note + "; "
        End If
        'Remove the last "; "
        sNote = Left(sNote, Len(sNote) - Len("; "))
        sNote = sNote & crtfEndControlName
    End If

    If mstrSummary <> "" Then
        sNote = sNote + crtfStartControlName & "Objective: " '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
        sNote = sNote + "Summary of Findings: " + mstrSummary
        sNote = sNote + crtfEndControlName
    End If
    
    If mstrMax_Impressions <> "" Or mstrMan_Impressions <> "" Or mintDiagnostic_Mounting <> 0 Then
        sNote = sNote + crtfStartControlName & "Impressions made for: " '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
        sTxt = ""
        If mstrMax_Impressions <> "" Then
            sTxt = Replace(mstrMax_Impressions, ",", "; Max ")
            sNote = sNote + "Max " + sTxt + "; "
            sTxt = ""
        End If
        If mstrMan_Impressions <> "" Then
            sTxt = Replace(mstrMan_Impressions, ",", "; Man ")
            sNote = sNote + "Man " + sTxt + "; "
        End If
        If mintDiagnostic_Mounting <> 0 Then
            sNote = sNote + "Diagnostic Mountings made; "
        End If
        'Remove the last "; "
        sNote = Left(sNote, Len(sNote) - Len("; "))
        sNote = sNote + crtfEndControlName
    End If
    
    'SCR 8485
    If mintSTE_WithinLimits = -1 Or mstrRevealsNote <> "" Then
        sNote = sNote + crtfStartControlName & "Soft Tissue Exam: " '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
        If mintSTE_WithinLimits = -1 Then
            sNote = sNote + "Within normal limits; "
        End If
        If mstrRevealsNote <> "" Then
            sNote = sNote & "Reveals Comment: " + mstrRevealsNote + "; "
        End If
        'Remove the last "; "
        sNote = Left(sNote, Len(sNote) - Len("; "))
        sNote = sNote & crtfEndControlName
    End If
    
    clsISection_GetRTFNote = sNote
    Exit Function
errTrap:
    clsISection_GetRTFNote = "Error in generating note for ExamEval-General section."
    goCWShared.ShowVBError Err.Number, Err.Description, Err.Source & "." & sPROC_NAME, App.Title, vbCritical

End Function

'*******************************************************************
'Function: GetNote
'Author: Sean McAvoy
'Date: 12/6/2001
'Returns: A string that is displayed in the narrative summary for this encounter
'Purpose: To compare the values from this encounter and the previous encounter and
'return a text note.
'*******************************************************************
Private Function clsISection_GetNote() As String

Dim sTemp As String
 Dim sNote As String
 Dim sTxt As String
 
 Const sPROC_NAME = "clsISection_GetNote"
 On Error GoTo errTrap
    
    If mstrMax_Impressions = "" And mstrMan_Impressions = "" And _
        mintSTE_WithinLimits = 0 And _
        mstrRevealsNote = "" And _
        mintRiskAssess_Caries = 0 And _
        mintRiskAssess_Perio = 0 And _
        mintRiskAssess_OralCA = 0 And _
        mstrRiskAssess_Note = "" And _
        mstrSummary = "" And mintDiagnostic_Mounting = 0 Then
    
        'No need to update note..
        sNote = ""
        clsISection_GetNote = sNote
        Exit Function
    End If
        
    Select Case mintRiskAssess_Caries
       Case 1
           sNote = sNote + vbTab + "The Caries risk assessment for this patient is LOW." + vbCrLf
       Case 2
           sNote = sNote + vbTab + "The Caries risk assessment for this patient is MODERATE." + vbCrLf
       Case 3
           sNote = sNote + vbTab + "The Caries risk assessment for this patient is HIGH." + vbCrLf
    End Select
    Select Case mintRiskAssess_Perio
       Case 1
           sNote = sNote + vbTab + "The Perio risk assessment for this patient is LOW." + vbCrLf
       Case 2
           sNote = sNote + vbTab + "The Perio risk assessment for this patient is MODERATE." + vbCrLf
       Case 3
           sNote = sNote + vbTab + "The Perio risk assessment for this patient is HIGH." + vbCrLf
    End Select
    Select Case mintRiskAssess_OralCA
       Case 1
           sNote = sNote + vbTab + "The Oral CA risk assessment for this patient is LOW." + vbCrLf
       Case 2
           sNote = sNote + vbTab + "The Oral CA risk assessment for this patient is MODERATE." + vbCrLf
       Case 3
           sNote = sNote + vbTab + "The Oral CA risk assessment for this patient is HIGH." + vbCrLf
    End Select
    
    If Trim(mstrRiskAssess_Note) <> "" Then
       sNote = sNote + vbTab + "Risk Assessment Note: " + mstrRiskAssess_Note + vbCrLf
    End If
        
    If mstrSummary <> "" Then sNote = sNote + vbTab + "Summary Note: " + mstrSummary + vbCrLf
    
    If mstrMax_Impressions <> "" Then
        sNote = sNote + vbTab + "Max Impressions Made For: " + mstrMax_Impressions + vbCrLf
    End If
    If mstrMan_Impressions <> "" Then
        sNote = sNote + vbTab + "Man Impressions Made For: " + mstrMan_Impressions + vbCrLf
    End If
    If mintDiagnostic_Mounting <> 0 Then
        sNote = sNote + vbTab + "Diagnostic Mountings Impressions were made for the patient." + vbCrLf + vbCrLf
    End If
    
    If mintSTE_WithinLimits = -1 Then
        sNote = sNote + vbTab + "The patient's soft tissue Exam is within normal limits." + vbCrLf
    End If
    If Trim(mstrRevealsNote) <> "" Then
        sNote = sNote + vbTab + "Reveals Note: " + mstrRevealsNote + vbCrLf
    End If
    
    clsISection_GetNote = sNote
    Exit Function
errTrap:
    clsISection_GetNote = "Error in generating note for ExamEval-General section."
    goCWShared.ShowVBError Err.Number, Err.Description, Err.Source & "." & sPROC_NAME, App.Title, vbCritical


    
End Function
'*******************************************************************
'Function: GenerateSOAPWL_ID
'Author: Sean McAvoy
'Date: 12/11/2001
'Returns: True if successful.
'Purpose: Called from SaveSection, it's main purpose is to set the mlXXX_WL variable.
'This variable is represents the workload code for each item on the screen that
'requires workload credit.
'*******************************************************************

'*******************************************************************
'Function: SaveSection
'Author: Sean McAvoy
'Date: 12/5/2001
'Returns: True/False
'Purpose: Saves the data to the table
'*******************************************************************
Private Function clsISection_SaveSection() As Boolean
Dim mstrMax_ImpressionsTMP As String
Dim mstrMan_ImpressionsTMP As String
Dim mstrSummaryTMP As String
Dim mstrOtherTMP As String
Dim mstrRevealsNoteTMP As String
Dim sTmpMsg As String
Dim sSQL As String
Dim mstrRiskAssess_NoteTMP As String

Const sPROC_NAME As String = "SaveSection"

On Error GoTo ErrHnd

    If dfDirtyFlag = dfNoAction Then
        clsISection_SaveSection = True
        Exit Function
    Else
        If mvarDentalEncounterID = 0 Then
            clsISection_SaveSection = False
            Exit Function
        End If
    End If
        
    'Clean up and ready string fields for db...
    mstrMax_ImpressionsTMP = Trim(mstrMax_Impressions)
    mstrMan_ImpressionsTMP = Trim(mstrMan_Impressions)
    mstrSummaryTMP = Trim(mstrSummary)
    mstrRevealsNoteTMP = Trim(mstrRevealsNote)
    mstrRiskAssess_NoteTMP = Trim(mstrRiskAssess_Note)

    'deal with "  '  " if any...
    If mstrMax_ImpressionsTMP <> "" Then mstrMax_ImpressionsTMP = RQ(mstrMax_ImpressionsTMP)
    If mstrMan_ImpressionsTMP <> "" Then mstrMan_ImpressionsTMP = RQ(mstrMan_ImpressionsTMP)
    If mstrSummaryTMP <> "" Then mstrSummaryTMP = RQ(mstrSummaryTMP)
    If mstrOtherTMP <> "" Then mstrOtherTMP = RQ(mstrOtherTMP)
    If mstrRevealsNoteTMP <> "" Then mstrRevealsNoteTMP = RQ(mstrRevealsNoteTMP)
    If mstrRiskAssess_NoteTMP <> "" Then mstrRiskAssess_NoteTMP = RQ(mstrRiskAssess_NoteTMP)

    If dfDirtyFlag = dfAdd Then
        If mvarDentalEncounterID = 0 Then
            clsISection_SaveSection = False
            Exit Function
        End If

        'The record does not exist, so create an insert statement
        sSQL = "INSERT INTO TEMPL_EXAMSOAP (" _
        & "DENTAL_ENCOUNTER_ID, MAX_IMPRESSIONS, MAN_IMPRESSIONS, SUMMARY, RISK_CARIES, RISK_PERIO, RISK_ORALCA, RISK_NOTE, " _
        & "REVEAL_NOTE, STE_WITHINLIMITS, " _
        & "DIAGNOSTIC_MOUNTING) " _
            & " VALUES" _
            & " (" & CStr(mvarDentalEncounterID) & ", " _
            & " '" & mstrMax_ImpressionsTMP & "', " _
            & " '" & mstrMan_ImpressionsTMP & "', " _
            & " '" & mstrSummaryTMP & "', " _
            & CStr(mintRiskAssess_Caries) & ", " _
            & CStr(mintRiskAssess_Perio) & ", " _
            & CStr(mintRiskAssess_OralCA) & ", " _
            & " '" & mstrRiskAssess_NoteTMP & "', " _

        sSQL = sSQL _
            & " '" & mstrRevealsNoteTMP & "', " _
            & CStr(mintSTE_WithinLimits) & ", " _
            & CStr(mintDiagnostic_Mounting) & ") " _

    ElseIf dfDirtyFlag = dfUpdate Then
        'The record exists, so create an update statement
        sSQL = "UPDATE TEMPL_EXAMSOAP SET " + _
        "MAX_IMPRESSIONS = '" + mstrMax_ImpressionsTMP + "', " + _
        "MAN_IMPRESSIONS = '" + mstrMan_ImpressionsTMP + "', "
    
        'SCR 8485
        sSQL = sSQL & _
            "STE_WITHINLIMITS = " + CStr(mintSTE_WithinLimits) + ", " + _
            "REVEAL_NOTE = '" + mstrRevealsNoteTMP + "', "
        
        'SCR 8499
        sSQL = sSQL & _
            "RISK_CARIES = " + CStr(mintRiskAssess_Caries) + ", " + _
            "RISK_PERIO = " + CStr(mintRiskAssess_Perio) + ", " + _
            "RISK_ORALCA = " + CStr(mintRiskAssess_OralCA) + ", " + _
            "RISK_NOTE = '" + mstrRiskAssess_NoteTMP + "', "

        sSQL = sSQL + _
        "SUMMARY = '" + mstrSummaryTMP + "', " + _
        "DIAGNOSTIC_MOUNTING = " + CStr(mintDiagnostic_Mounting) + _
        " Where (DENTAL_ENCOUNTER_ID = " & CStr(mvarDentalEncounterID) & ")"
        
 
    ElseIf dfDirtyFlag = dfdelete Then
        'Mark the record as deleted
        sSQL = "DELETE FROM TEMPL_EXAMSOAP " + _
               "Where (DENTAL_ENCOUNTER_ID = " & CStr(mvarDentalEncounterID) & ")"
    End If
    
    goconn.ExecuteSQL sSQL
    
    bSavedToDB = True
    dfDirtyFlag = dfNoAction
    
    clsISection_SaveSection = True
    
Exit Function

ErrHnd:
    clsISection_SaveSection = False
    goCWShared.ShowVBError Err.Number, Err.Description + "  " + sTmpMsg, Err.Source & "." & sPROC_NAME, App.Title, vbCritical

End Function



'*******************************************************************
'Function: LoadSection
'Author: Sean McAvoy
'Date: 12/4/2001
'Returns: True/False
'Purpose: This loads the data from the db
'*******************************************************************
Private Function clsISection_LoadSection(ByVal PatientID As Long, ByVal DentalEncounterID As Long) As Boolean

Dim sSQL As String, rs As Recordset, rsPrev As Recordset
Const sPROC_NAME = "clsExamSOAP.clsISection_LoadSection"

On Error GoTo ErrHnd

    mvarPatientID = PatientID
    mvarDentalEncounterID = DentalEncounterID
    
    If gbUseStoredProc = True Then
        
        'Load the values using the SP
        Dim oCmd As New CHCSII_Command.CHCSII_Cmd
        
        oCmd.SPName = "Dent_Notes_SOAP_Load"

        Call oCmd.AddParam(False, False, "textString", mvarNarrSummDentalEncounterID)
        Call oCmd.AddParam(True, True, "", "")
        
        Set rs = goconn.OpenRecordsetSP(oCmd)
            
        If rs.EOF Then
            bSavedToDB = False
            dfDirtyFlag = dfAdd
        
            mstrMax_Impressions = ""
            mstrMan_Impressions = ""
            mstrSummary = ""
            mintDiagnostic_Mounting = 0
            
            'SCR 8485
            mintSTE_WithinLimits = 0
            mstrRevealsNote = ""

            'SCR 8499
            mintRiskAssess_Caries = 0
            mintRiskAssess_Perio = 0
            mintRiskAssess_OralCA = 0
            mstrRiskAssess_Note = ""

            mstrSummary = ""
        Else
        
            bSavedToDB = True
            dfDirtyFlag = dfNoAction
            
            mstrMax_Impressions = rs("MAX_IMPRESSIONS") & ""
            mstrMan_Impressions = rs("MAN_IMPRESSIONS") & ""
            mintDiagnostic_Mounting = Val(rs("DIAGNOSTIC_MOUNTING") & "")
            
            'SCR 8485
            mintSTE_WithinLimits = Val(rs("STE_WITHINLIMITS") & "")
            mstrRevealsNote = rs("REVEAL_NOTE") & ""
            
            'SCR 8499
            mintRiskAssess_Caries = Val(rs("RISK_CARIES") & "")
            mintRiskAssess_Perio = Val(rs("RISK_PERIO") & "")
            mintRiskAssess_OralCA = Val(rs("RISK_ORALCA") & "")
            mstrRiskAssess_Note = rs("RISK_NOTE") & ""

            mstrSummary = rs("SUMMARY") & ""
            
        End If
            
    Else

        'Load the values
        'Because of the issues with Tuxedo (4k limit), three sql calls will
        'need to be made to retrieve the several 2k note fields.
    
        sSQL = "SELECT MAX_IMPRESSIONS, MAN_IMPRESSIONS, DIAGNOSTIC_MOUNTING, SICK_CALL_WL, AFTER_HOURSCARE_WL," _
            & " RISK_CARIES, RISK_PERIO, RISK_ORALCA, STE_WITHINLIMITS " _
            & " FROM TEMPL_EXAMSOAP Where DENTAL_ENCOUNTER_ID = " & CStr(mvarDentalEncounterID)
        
        Set rs = goconn.OpenRecordset(sSQL)
        
        If rs.EOF Then
            bSavedToDB = False
            dfDirtyFlag = dfAdd
        
            mstrMax_Impressions = ""
            mstrMan_Impressions = ""
            mstrSummary = ""
            mintDiagnostic_Mounting = 0
            
            'SCR 8485
            mintSTE_WithinLimits = 0
            mstrRevealsNote = ""

            'SCR 8499
            mintRiskAssess_Caries = 0
            mintRiskAssess_Perio = 0
            mintRiskAssess_OralCA = 0
            mstrRiskAssess_Note = ""

            mstrSummary = ""
                        
        Else
            bSavedToDB = True
            dfDirtyFlag = dfNoAction
        
            mstrMax_Impressions = rs("MAX_IMPRESSIONS") & ""
            mstrMan_Impressions = rs("MAN_IMPRESSIONS") & ""
            mintDiagnostic_Mounting = Val(rs("DIAGNOSTIC_MOUNTING") & "")
            
            'SCR 8485
            mintSTE_WithinLimits = Val(rs("STE_WITHINLIMITS") & "")
            
            'SCR 8499
            mintRiskAssess_Caries = Val(rs("RISK_CARIES") & "")
            mintRiskAssess_Perio = Val(rs("RISK_PERIO") & "")
            mintRiskAssess_OralCA = Val(rs("RISK_ORALCA") & "")

                    
            sSQL = "SELECT RISK_NOTE, REVEAL_NOTE, SUMMARY " + _
                "FROM TEMPL_EXAMSOAP Where " + _
                "DENTAL_ENCOUNTER_ID = " + CStr(mvarDentalEncounterID)
            rs.Close
            Set rs = goconn.OpenRecordset(sSQL)
            If rs.EOF Then
                mstrRiskAssess_Note = ""
                mstrRevealsNote = ""
                mstrSummary = ""
            Else
                mstrRiskAssess_Note = rs("RISK_NOTE") & ""
                mstrRevealsNote = rs("REVEAL_NOTE") & ""
                mstrSummary = rs("SUMMARY") & ""
            End If
            
        End If
    End If
    
    rs.Close
    Set rs = Nothing
    
    clsISection_LoadSection = True
    
Exit Function

ErrHnd:
    Set rs = Nothing
    clsISection_LoadSection = False
    goCWShared.ShowVBError Err.Number, Err.Description, Err.Source & "." & sPROC_NAME, App.Title, vbCritical
    Exit Function
    Resume
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
    mstrMax_Impressions = ""
    mstrMan_Impressions = ""
    mstrSummary = ""
    mintDiagnostic_Mounting = 0

    'SCR 8485
    mintSTE_WithinLimits = 0
    mstrRevealsNote = ""
    
    'SCR 8499
    mintRiskAssess_Caries = 0
    mintRiskAssess_Perio = 0
    mintRiskAssess_OralCA = 0
    mstrRiskAssess_Note = ""

End Sub






