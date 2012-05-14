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
Attribute VB_Name = "clsNarrSumm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Implements clsISection



Public Property Let MaterialsTabNote(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MaterialsTabNote = 5
    
    If mvarNarrSummMaterialsTabNote = vData Then
        Exit Property
    End If
    
    mvarNarrSummMaterialsTabNote = vData
    UpdateDirtyFlag
End Property


Public Property Get MaterialsTabNote() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MaterialsTabNote
    MaterialsTabNote = mvarNarrSummMaterialsTabNote
End Property



Public Property Get PerioChartingNote() As String

    PerioChartingNote = mvarNarrSummPerioChartingNote
    
End Property


Public Property Let PerioChartingNote(ByVal vData As String)
    
    If mvarNarrSummPerioChartingNote = vData Then
        Exit Property
    End If
    
    mvarNarrSummPerioChartingNote = vData
    UpdateDirtyFlag

End Property

Public Property Let RadiologyTabNote(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RadiologyTabNote = 5
    
    If mvarNarrSummRadiologyTabNote = vData Then
        Exit Property
    End If
    
    mvarNarrSummRadiologyTabNote = vData
    UpdateDirtyFlag
End Property


Public Property Get RadiologyTabNote() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RadiologyTabNote
    RadiologyTabNote = mvarNarrSummRadiologyTabNote
End Property



Public Property Let TreatmentTabNote(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TreatmentTabNote = 5
    
    If mvarNarrSummTreatmentTabNote = vData Then
        Exit Property
    End If
    
    mvarNarrSummTreatmentTabNote = vData
    UpdateDirtyFlag
End Property


Public Property Get TreatmentTabNote() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TreatmentTabNote
    TreatmentTabNote = mvarNarrSummTreatmentTabNote
End Property



Public Property Let PreventiveTabNote(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PreventiveTabNote = 5
    
    If mvarNarrSummPreventiveTabNote = vData Then
        Exit Property
    End If
    
    mvarNarrSummPreventiveTabNote = vData
    UpdateDirtyFlag
End Property


Public Property Get PreventiveTabNote() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PreventiveTabNote
    PreventiveTabNote = mvarNarrSummPreventiveTabNote
End Property



Private Property Let ExamTabNote(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ExamTabNote = 5
    
    If mvarNarrSummExamTabNote = vData Then
        Exit Property
    End If
    
    mvarNarrSummExamTabNote = vData
    UpdateDirtyFlag
End Property


Private Property Get ExamTabNote() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ExamTabNote
    ExamTabNote = mvarNarrSummExamTabNote
End Property

Public Property Let ExamProsTabNote(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ExamTabNote = 5
    
    If mvarNarrSummExamProsTabNote = vData Then
        Exit Property
    End If
    
    mvarNarrSummExamProsTabNote = vData
    UpdateDirtyFlag
End Property


Public Property Get ExamProsTabNote() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ExamTabNote
    ExamProsTabNote = mvarNarrSummExamProsTabNote
End Property
Public Property Let ExamPerioTabNote(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ExamTabNote = 5
    
    If mvarNarrSummExamPerioTabNote = vData Then
        Exit Property
    End If
    
    mvarNarrSummExamPerioTabNote = vData
    UpdateDirtyFlag
End Property


Public Property Get ExamPerioTabNote() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ExamTabNote
    ExamPerioTabNote = mvarNarrSummExamPerioTabNote
End Property

Public Property Let ExamEndoTabNote(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ExamTabNote = 5

    If mvarNarrSummExamEndoTabNote = vData Then
        Exit Property
    End If
    
    mvarNarrSummExamEndoTabNote = vData
    UpdateDirtyFlag
End Property


Public Property Get ExamEndoTabNote() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ExamTabNote
    ExamEndoTabNote = mvarNarrSummExamEndoTabNote
End Property

Public Property Let ExamSOAPTabNote(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ExamTabNote = 5
    
    If mvarNarrSummExamSOAPTabNote = vData Then
        Exit Property
    End If
    
    mvarNarrSummExamSOAPTabNote = vData
    UpdateDirtyFlag
End Property


Public Property Get ExamSOAPTabNote() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ExamTabNote
    ExamSOAPTabNote = mvarNarrSummExamSOAPTabNote
End Property


Public Property Let ExamGeneralTabNote(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ExamTabNote = 5

    If mvarNarrSummExamGeneralTabNote = vData Then
        Exit Property
    End If
    
    mvarNarrSummExamGeneralTabNote = vData
    UpdateDirtyFlag
End Property


Public Property Get ExamGeneralTabNote() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ExamTabNote
    ExamGeneralTabNote = mvarNarrSummExamGeneralTabNote
End Property



Public Property Let WorkloadId(vData As Long)
    
    mvarNarrSummWorkloadID = vData
    UpdateDirtyFlag
    
End Property

Public Property Get WorkloadId() As Long
    
    WorkloadId = mvarNarrSummWorkloadID
    
End Property

Private Property Get clsISection_DentalEncounterID() As Long
    
    clsISection_DentalEncounterID = mvarNarrSummDentalEncounterID

End Property



Private Function GetWorkloadCode() As String

    'Add code here to figure out the workload code for this instance and return it.
    'Used in SaveSection

End Function


Private Property Get clsISection_PatientID() As Long
    
    clsISection_PatientID = mvarNarrSummPatientID

End Property



Private Function clsISection_GetRTFNote() As String

    clsISection_GetRTFNote = clsISection_GetNote

End Function

'*******************************************************************
'Function: GetNote
'Author: Your name here
'Date:
'Returns: A string that is displayed in the narrative summary for this encounter
'Purpose: To compare the values from this encounter and the previous encounter and
'return a text note.
'*******************************************************************
Private Function clsISection_GetNote() As String

Dim sTemp As String
    
'    For Each oIV In oIVs
'        sTemp = sTemp & oIV.GetNote
'    Next
'
'    If sTemp <> "" Then
'        clsISection_GetNote = "Anesthetics" & vbCrLf & sTemp
'    End If
    
End Function

'*******************************************************************
'Function: SaveSection
'Author: Your name here
'Date:
'Returns: True/False
'Purpose: Saves the data to the table
'*******************************************************************
Private Function clsISection_SaveSection() As Boolean

Dim sSQL As String

Const sPROC_NAME As String = "SaveSection"

On Error GoTo ErrHnd

    If dfNarrSummDirtyFlag = dfNoAction Then
        clsISection_SaveSection = True
        Exit Function
    Else
        'Update the workload
        'UpdateWorkload mvarNarrSummWorkloadID, dfNarrSummDirtyFlag, mvarNarrSummDentalEncounterID, GetWorkloadCode, 1
    End If
        
    If dfNarrSummDirtyFlag = dfAdd Then
        'The record does not exist, so create an insert statement
        InsertRecords
        
    ElseIf dfNarrSummDirtyFlag = dfUpdate Then
        'The record exists, so delete the rows and re-add them
        If gbUseStoredProc = True Then
            'Do nothing
        Else
            sSQL = "DELETE FROM TEMPL_NARRATIVE_SUMMARY Where (DENTAL_ENCOUNTER_ID = " & mvarNarrSummDentalEncounterID & ")"
            goconn.ExecuteSQL sSQL
        End If
        
        InsertRecords
        
    ElseIf dfNarrSummDirtyFlag = dfdelete Then
        'Mark the record as deleted
        sSQL = "DELETE FROM TEMPL_NARRATIVE_SUMMARY Where (DENTAL_ENCOUNTER_ID = " & mvarNarrSummDentalEncounterID & ")"
        goconn.ExecuteSQL sSQL
        
    End If
    
    bNarrSummSavedToDB = True
    dfNarrSummDirtyFlag = dfNoAction
    
    clsISection_SaveSection = True

Exit Function

ErrHnd:
    goCWShared.ShowVBError Err.Number, Err.Description, Err.Source & "." & sPROC_NAME, App.Title, vbCritical

End Function

'*******************************************************************
'Function: InsertRecords
'Author: Scott Foster
'Date: 11/28/01
'Returns: Nothing
'Purpose: To insert the tab records into the database, this allows for an unlimited note.
'*******************************************************************
Private Sub InsertRecords()

On Error GoTo ErrHnd

Dim lSeq As Long
Dim sSQL As String
Dim sExamNote As String

Dim sExamGeneralNote As String
Dim sExamSOAPNote As String
Dim sExamEndoNote As String
Dim sExamPerioNote As String
Dim sExamProsNote As String

Dim sPreventiveNote As String
Dim sTreatmentNote As String
Dim sRadiologyNote As String
Dim sMaterialsNote As String

Dim sPerioChartingNote As String

    sExamNote = mvarNarrSummExamTabNote
    sPreventiveNote = mvarNarrSummPreventiveTabNote
    sTreatmentNote = mvarNarrSummTreatmentTabNote
    sRadiologyNote = mvarNarrSummRadiologyTabNote
    sMaterialsNote = mvarNarrSummMaterialsTabNote

    sExamGeneralNote = mvarNarrSummExamGeneralTabNote
    sExamSOAPNote = mvarNarrSummExamSOAPTabNote
    sExamEndoNote = mvarNarrSummExamEndoTabNote
    sExamPerioNote = mvarNarrSummExamPerioTabNote
    sExamProsNote = mvarNarrSummExamProsTabNote

    sPerioChartingNote = mvarNarrSummPerioChartingNote
    
    'Always use SQL for now because of SCR 40930, until SCR 33942 is fixed.
    If False Then
    'If gbUseStoredProc = True Then
        Dim oCmd As New CHCSII_Command.CHCSII_Cmd
        
        oCmd.SPName = "Dent_Notes_NarrSumm_Save"

        Call oCmd.AddParam(False, False, "textString", mvarNarrSummDentalEncounterID)
        Call oCmd.AddParam(False, False, "textString", RQ(sExamNote))
        Call oCmd.AddParam(False, False, "textString", RQ(sExamGeneralNote))
        Call oCmd.AddParam(False, False, "textString", RQ(sExamSOAPNote))
        Call oCmd.AddParam(False, False, "textString", RQ(sExamEndoNote))
        Call oCmd.AddParam(False, False, "textString", RQ(sExamPerioNote))
        Call oCmd.AddParam(False, False, "textString", RQ(sExamProsNote))
        Call oCmd.AddParam(False, False, "textString", RQ(sPreventiveNote))
        Call oCmd.AddParam(False, False, "textString", RQ(sTreatmentNote))
        Call oCmd.AddParam(False, False, "textString", RQ(sRadiologyNote))
        Call oCmd.AddParam(False, False, "textString", RQ(sMaterialsNote))
        Call oCmd.AddParam(False, False, "textString", RQ(sPerioChartingNote))
        
        goconn.ExecuteSP oCmd
            
    Else
        'Always delete
        sSQL = "DELETE FROM TEMPL_NARRATIVE_SUMMARY Where (DENTAL_ENCOUNTER_ID = " & mvarNarrSummDentalEncounterID & ")"
        goconn.ExecuteSQL sSQL

        Do
            lSeq = goconn.GetNextID(esTEMPL_NARRATIVE_SUMMARY_SEQ)
            
            'Create a insert statement with the first 500 chars
            sSQL = "INSERT INTO TEMPL_NARRATIVE_SUMMARY" _
                & " (ID," _
                & " DENTAL_ENCOUNTER_ID," _
                & " EXAM_EVAL," _
                & " EXAM_General," _
                & " EXAM_SOAP," _
                & " EXAM_Endo," _
                & " EXAM_Perio," _
                & " EXAM_Pros," _
                & " PREVENTIVE," _
                & " TREATMENT," _
                & " RADIOLOGY," _
                & " MATERIALS," _
                & " Perio_Charting)" _
                & " VALUES"
            sSQL = sSQL & " (" & lSeq & "," _
                & " " & mvarNarrSummDentalEncounterID & "," _
                & " '" & RQ(Left(sExamNote, 500)) & "'," _
                & " '" & RQ(Left(sExamGeneralNote, 500)) & "'," _
                & " '" & RQ(Left(sExamSOAPNote, 500)) & "'," _
                & " '" & RQ(Left(sExamEndoNote, 500)) & "'," _
                & " '" & RQ(Left(sExamPerioNote, 500)) & "'," _
                & " '" & RQ(Left(sExamProsNote, 500)) & "'," _
                & " '" & RQ(Left(sPreventiveNote, 500)) & "'," _
                & " '" & RQ(Left(sTreatmentNote, 500)) & "'," _
                & " '" & RQ(Left(sRadiologyNote, 500)) & "'," _
                & " '" & RQ(Left(sMaterialsNote, 500)) & "'," _
                & " '" & RQ(Left(sPerioChartingNote, 500)) & "')"
            goconn.ExecuteSQL sSQL
            
            'Remove the first 500 that was saved to the db
            sExamNote = Mid(sExamNote, 501)
            sExamGeneralNote = Mid(sExamGeneralNote, 501)
            sExamSOAPNote = Mid(sExamSOAPNote, 501)
            sExamEndoNote = Mid(sExamEndoNote, 501)
            sExamPerioNote = Mid(sExamPerioNote, 501)
            sExamProsNote = Mid(sExamProsNote, 501)
            sPreventiveNote = Mid(sPreventiveNote, 501)
            sTreatmentNote = Mid(sTreatmentNote, 501)
            sRadiologyNote = Mid(sRadiologyNote, 501)
            sMaterialsNote = Mid(sMaterialsNote, 501)
            sPerioChartingNote = Mid(sPerioChartingNote, 501)
            
        Loop Until sExamNote = "" And sExamGeneralNote = "" And sExamSOAPNote = "" _
            And sExamEndoNote = "" And sExamPerioNote = "" And sExamProsNote = "" _
            And sPreventiveNote = "" And sTreatmentNote = "" And sRadiologyNote = "" _
            And sMaterialsNote = "" And sPerioChartingNote = ""
    
    End If

Exit Sub

ErrHnd:
    If Err.Number = -2147221503 Then
        Resume Next
    Else
        goCWShared.ShowVBError Err.Number, Err.Description, Err.Source & "." & "InsertRecords", App.Title, vbCritical
    End If
    
End Sub
'*******************************************************************
'Function: LoadSection
'Author: Scott Foster
'Date: 10-22-01
'Returns: True/False
'Purpose: This loads the data from the db
'*******************************************************************
Private Function clsISection_LoadSection(ByVal PatientID As Long, ByVal DentalEncounterID As Long) As Boolean

Dim sSQL As String, rs As Recordset, rsPrev As Recordset, rs2 As Recordset, rs3 As Recordset

Const sPROC_NAME = "LoadSection"

On Error GoTo ErrHnd

    If mvarNarrSummDentalEncounterID = DentalEncounterID Then
        Exit Function
    End If
    
    mvarNarrSummPatientID = PatientID
    mvarNarrSummDentalEncounterID = DentalEncounterID

    If gbUseStoredProc = True Then
        Dim oCmd As New CHCSII_Command.CHCSII_Cmd
        
        oCmd.SPName = "Dent_Notes_NarrSumm_Load"

        Call oCmd.AddParam(False, False, "textString", mvarNarrSummDentalEncounterID)
        Call oCmd.AddParam(True, True, "", "")
        
        Set rs = goconn.OpenRecordsetSP(oCmd)
        
        If rs.EOF Then
             'Init the values here to ""
             mvarNarrSummExamTabNote = ""
             mvarNarrSummExamGeneralTabNote = ""
             mvarNarrSummExamSOAPTabNote = ""
             mvarNarrSummExamEndoTabNote = ""
             mvarNarrSummExamPerioTabNote = ""
             mvarNarrSummExamProsTabNote = ""
             mvarNarrSummPreventiveTabNote = ""
             mvarNarrSummTreatmentTabNote = ""
             mvarNarrSummRadiologyTabNote = ""
             mvarNarrSummMaterialsTabNote = ""
             mvarNarrSummPerioChartingNote = ""
             bNarrSummSavedToDB = False
             dfNarrSummDirtyFlag = dfAdd
        Else
             'Set the values from the db including "" to allow for nulls
             mvarNarrSummExamTabNote = ""
             mvarNarrSummExamGeneralTabNote = ""
             mvarNarrSummExamSOAPTabNote = ""
             mvarNarrSummExamEndoTabNote = ""
             mvarNarrSummExamPerioTabNote = ""
             mvarNarrSummExamProsTabNote = ""
             mvarNarrSummPreventiveTabNote = ""
             mvarNarrSummTreatmentTabNote = ""
             mvarNarrSummRadiologyTabNote = ""
             mvarNarrSummMaterialsTabNote = ""
             mvarNarrSummPerioChartingNote = ""
             
             Do Until rs.EOF
                 mvarNarrSummExamTabNote = mvarNarrSummExamTabNote & rs("EXAM_EVAL") & ""
                 mvarNarrSummPreventiveTabNote = mvarNarrSummPreventiveTabNote & rs("PREVENTIVE") & ""
                 mvarNarrSummTreatmentTabNote = mvarNarrSummTreatmentTabNote & rs("TREATMENT") & ""
                 mvarNarrSummRadiologyTabNote = mvarNarrSummRadiologyTabNote & rs("RADIOLOGY") & ""
                 mvarNarrSummMaterialsTabNote = mvarNarrSummMaterialsTabNote & rs("MATERIALS") & ""
                 mvarNarrSummExamGeneralTabNote = mvarNarrSummExamGeneralTabNote & rs("EXAM_General") & ""
                 mvarNarrSummExamSOAPTabNote = mvarNarrSummExamSOAPTabNote & rs("EXAM_SOAP") & ""
                 mvarNarrSummExamEndoTabNote = mvarNarrSummExamEndoTabNote & rs("EXAM_Endo") & ""
                 mvarNarrSummExamPerioTabNote = mvarNarrSummExamPerioTabNote & rs("EXAM_Perio") & ""
                 mvarNarrSummExamProsTabNote = mvarNarrSummExamProsTabNote & rs("EXAM_Pros") & ""
                 mvarNarrSummPerioChartingNote = mvarNarrSummPerioChartingNote & rs("Perio_Charting") & ""
                 
                 rs.MoveNext
             Loop
             
             bNarrSummSavedToDB = True
             dfNarrSummDirtyFlag = dfNoAction
        End If
        
    Else
        'Load the values can only get a max of 2500 chars in one rs that is why we have 2 sql statements
        sSQL = "SELECT ID, EXAM_EVAL, PREVENTIVE, TREATMENT, RADIOLOGY, " _
            & " MATERIALS From TEMPL_NARRATIVE_SUMMARY" _
            & " Where (DENTAL_ENCOUNTER_ID = " & mvarNarrSummDentalEncounterID & ")" _
            & " ORDER BY ID"
        
        Set rs = goconn.OpenRecordset(sSQL)
        
        'Load the values can only get a max of 2500 chars in one rs that is why we have 2 sql statements
        sSQL = "SELECT ID, EXAM_General, EXAM_SOAP, EXAM_Endo, EXAM_Perio, " _
            & " EXAM_Pros From TEMPL_NARRATIVE_SUMMARY" _
            & " Where (DENTAL_ENCOUNTER_ID = " & mvarNarrSummDentalEncounterID & ")" _
            & " ORDER BY ID"
        Set rs2 = goconn.OpenRecordset(sSQL)
        
        'Load the values can only get a max of 2500 chars in one rs that is why we have 2 sql statements
        sSQL = "SELECT ID, Perio_Charting From TEMPL_NARRATIVE_SUMMARY" _
            & " Where (DENTAL_ENCOUNTER_ID = " & mvarNarrSummDentalEncounterID & ")" _
            & " ORDER BY ID"
        Set rs3 = goconn.OpenRecordset(sSQL)
    
        If rs.EOF Then
             'Init the values here to ""
             mvarNarrSummExamTabNote = ""
             mvarNarrSummExamGeneralTabNote = ""
             mvarNarrSummExamSOAPTabNote = ""
             mvarNarrSummExamEndoTabNote = ""
             mvarNarrSummExamPerioTabNote = ""
             mvarNarrSummExamProsTabNote = ""
             mvarNarrSummPreventiveTabNote = ""
             mvarNarrSummTreatmentTabNote = ""
             mvarNarrSummRadiologyTabNote = ""
             mvarNarrSummMaterialsTabNote = ""
             mvarNarrSummPerioChartingNote = ""
             dfNarrSummDirtyFlag = dfAdd
        Else
             'Set the values from the db including "" to allow for nulls
             mvarNarrSummExamTabNote = ""
             mvarNarrSummExamGeneralTabNote = ""
             mvarNarrSummExamSOAPTabNote = ""
             mvarNarrSummExamEndoTabNote = ""
             mvarNarrSummExamPerioTabNote = ""
             mvarNarrSummExamProsTabNote = ""
             mvarNarrSummPreventiveTabNote = ""
             mvarNarrSummTreatmentTabNote = ""
             mvarNarrSummRadiologyTabNote = ""
             mvarNarrSummMaterialsTabNote = ""
             mvarNarrSummPerioChartingNote = ""
             
             Do Until rs.EOF
                 mvarNarrSummExamTabNote = mvarNarrSummExamTabNote & rs("EXAM_EVAL") & ""
                 mvarNarrSummPreventiveTabNote = mvarNarrSummPreventiveTabNote & rs("PREVENTIVE") & ""
                 mvarNarrSummTreatmentTabNote = mvarNarrSummTreatmentTabNote & rs("TREATMENT") & ""
                 mvarNarrSummRadiologyTabNote = mvarNarrSummRadiologyTabNote & rs("RADIOLOGY") & ""
                 mvarNarrSummMaterialsTabNote = mvarNarrSummMaterialsTabNote & rs("MATERIALS") & ""
                 
                 rs.MoveNext
             Loop
             
             Do Until rs2.EOF
                 mvarNarrSummExamGeneralTabNote = mvarNarrSummExamGeneralTabNote & rs2("EXAM_General") & ""
                 mvarNarrSummExamSOAPTabNote = mvarNarrSummExamSOAPTabNote & rs2("EXAM_SOAP") & ""
                 mvarNarrSummExamEndoTabNote = mvarNarrSummExamEndoTabNote & rs2("EXAM_Endo") & ""
                 mvarNarrSummExamPerioTabNote = mvarNarrSummExamPerioTabNote & rs2("EXAM_Perio") & ""
                 mvarNarrSummExamProsTabNote = mvarNarrSummExamProsTabNote & rs2("EXAM_Pros") & ""
                 
                 rs2.MoveNext
             Loop
             
             Do Until rs3.EOF
                 mvarNarrSummPerioChartingNote = mvarNarrSummPerioChartingNote & rs3("Perio_Charting") & ""
                 
                 rs3.MoveNext
             Loop
             
             dfNarrSummDirtyFlag = dfNoAction
        End If
    
    End If
    
    rs.Close
        
    Set rs = Nothing
    
    clsISection_LoadSection = True
    
Exit Function

ErrHnd:
    goCWShared.ShowVBError Err.Number, Err.Description, Err.Source & "." & sPROC_NAME, App.Title, vbCritical

End Function

Private Function UpdateDirtyFlag()

    If bNarrSummSavedToDB = False Then
        dfNarrSummDirtyFlag = dfAdd
    ElseIf dfNarrSummDirtyFlag = dfNoAction Then
        dfNarrSummDirtyFlag = dfUpdate
    End If
    
End Function

Private Sub Class_Initialize()

    dfNarrSummDirtyFlag = dfNoAction
    
End Sub






