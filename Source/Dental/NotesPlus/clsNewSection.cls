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
Attribute VB_Name = "clsNewSection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Implements clsISection

Private mvarPatientID As Long
Private mvarDentalEncounterID As Long
Private mvarWorkloadID As Long

Private dfDirtyFlag As DirtyFlagEnum



Public Property Let WorkloadID(vData As Long)
    
    mvarWorkloadID = vData
    UpdateDirtyFlag
    
End Property

Public Property Get WorkloadID() As Long
    
    WorkloadID = mvarWorkloadID
    
End Property

Private Property Get clsISection_DentalEncounterID() As Long
    
    clsISection_DentalEncounterID = mvarDentalEncounterID

End Property



Private Function GetWorkloadCode() As String

    'Add code here to figure out the workload code for this instance and return it.
    'Used in SaveSection

End Function


Private Property Get clsISection_PatientID() As Long
    
    clsISection_PatientID = mvarPatientID

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

    If dfDirtyFlag = dfNoAction Then
        clsISection_SaveSection = True
        Exit Function
    Else
        'Update the workload
        UpdateWorkload mvarWorkloadID, dfDirtyFlag, mvarDentalEncounterID, GetWorkloadCode, 1
    End If
        
    If dfDirtyFlag = dfAdd Then
        'The record does not exist, so create an insert statement
        sSQL = ""
    ElseIf dfDirtyFlag = dfUpdate Then
        'The record exists, so create an update statement
        sSQL = ""
    ElseIf dfDirtyFlag = dfDelete Then
        'Mark the record as deleted
        sSQL = ""
    End If
    
    goconn.ExecuteSQL sSQL
    
    dfDirtyFlag = dfNoAction
    
    clsISection_SaveSection = True

Exit Function

ErrHnd:
    goCWShared.ShowVBError Err.Number, Err.Description, Err.Source & "." & sPROC_NAME, App.Title, vbCritical

End Function


'*******************************************************************
'Function: LoadSection
'Author: Scott Foster
'Date: 10-22-01
'Returns: True/False
'Purpose: This loads the data from the db
'*******************************************************************
Private Function clsISection_LoadSection(ByVal PatientID As Long, ByVal DentalEncounterID As Long) As Boolean

Dim sSQL As String, rs As Recordset, rsPrev As Recordset
Const sPROC_NAME = "LoadSection"

On Error GoTo ErrHnd

    mvarPatientID = PatientID
    mvarDentalEncounterID = DentalEncounterID
    
    'Load the values
    sSQL = ""
    
    Set rs = goconn.OpenRecordset(sSQL)
    
    If rs.EOF Then
        dfDirtyFlag = dfAdd
        'Init the values here to ""
        mvarWorkloadID = 0
        
        Dim mvarExample As String 'Delete this
        Dim mvarAnestheticUsed As Boolean 'Delete this
        mvarExample = ""
        mvarAnestheticUsed = True
    Else
        dfDirtyFlag = dfNoAction
        'Set the values from the db including "" to allow for nulls
        mvarWorkloadID = Val(rs("Workload_ID") & "")
        
        'Example of a text value
        mvarExample = rs("Example") & ""
        'Example of a boolean value
        mvarAnestheticUsed = CBool(Val(rs("ANESTH_USED") & ""))
    End If
    
    rs.Close
    
'   'This is only relevant for patient level sections.
'    'Find the previous encounter and load it's values to compare with the given encounter
'    ' to determine what values need to go into the note.
'    sSQL = "SELECT " _
'        & " From " _
'        & " Where " _
'        & " AND (DENTAL_ENCOUNTER.PATIENT_UNIT_ID = " & mvarPatientID & ") AND" _
'        & " (DENTAL_ENCOUNTER.DENTAL_ENCOUNTER_ID < " & mvarDentalEncounterID & ")" _
'        & " ORDER BY DENTAL_ENCOUNTER.DENTAL_ENCOUNTER_ID DESC"
'    Set rsPrev = goconn.OpenRecordset(sSQL)
'
'    If rsPrev.EOF Then
'        'The prev value is ""
'
'    Else
'        'Populate the previous encounter's values
'
'    End If
'
'    rsPrev.Close
'    Set rsPrev = Nothing
    
    Set rs = Nothing
    
    clsISection_LoadSection = True
    
Exit Function

ErrHnd:
    goCWShared.ShowVBError Err.Number, Err.Description, Err.Source & "." & sPROC_NAME, App.Title, vbCritical
    
End Function

Private Function UpdateDirtyFlag()

    If dfDirtyFlag = dfNoAction Then
        dfDirtyFlag = dfUpdate
    End If
    
End Function

Private Sub Class_Initialize()

    dfDirtyFlag = dfNoAction
    
End Sub




