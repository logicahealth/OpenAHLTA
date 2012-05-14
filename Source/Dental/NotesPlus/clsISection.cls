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
Attribute VB_Name = "clsISection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Private mvarPatientID As Long
'Private mvarDentalEncounterID As Long
'
Private dfDirtyFlag As DirtyFlagEnum
'Public Function GetWorkloadCode() As String
'
'    'Add code here to figure out the workload code for this instance and return it.
'
'End Function

Public Property Get DentalEncounterID() As Long
    
'    DentalEncounterID = mvarDentalEncounterID

End Property



Public Property Get PatientID() As Long
    
'    PatientID = mvarPatientID

End Property



Public Function GetRTFNote() As String

'    GetRTFNote = GetNote

End Function

'*******************************************************************
'Function: GetNote
'Author: Scott Foster
'Date: 10/19/01
'Returns: A string that is displayed in the narrative summary for this encounter
'Purpose: To compare the values from this encounter and the previous encounter and
'return a text note.
'*******************************************************************
Public Function GetNote() As String

'Dim sTemp As String
'
''    If mvarShortTerm <> "" Then
''        sTemp = sTemp & vbTab & "Short term: " & mvarShortTerm & vbCrLf
''    End If
''
''    If mvarLongTerm <> "" Then
''        sTemp = sTemp & vbTab & "Long term: " & mvarLongTerm & vbCrLf
''    End If
'
'
'    GetNote = sTemp

End Function

'*******************************************************************
'Function: SaveSection
'Author: Scott Foster
'Date: 10-22-01
'Returns: True/False
'Purpose: Saves the data to the table
'*******************************************************************
Public Function SaveSection() As Boolean

'Dim sSQL As String
'
'Const sPROC_NAME As String = "SaveSection"
'
'On Error GoTo ErrHnd
'
'    If dfDirtyFlag = dfNoAction Then
'        SaveSection = True
'        Exit Function
'
'    ElseIf dfDirtyFlag = dfAdd Then
'        'The record does not exist, so create an insert statement
'        sSQL = ""
'    ElseIf dfDirtyFlag = dfUpdate Then
'        'The record exists, so create an update statement
'        sSQL = ""
'    End If
'
'    goconn.ExecuteSQL sSQL
'
'    dfDirtyFlag = dfNoAction
'
'    SaveSection = True
'
'Exit Function
'
'ErrHnd:
'    goCWShared.ShowVBError Err.Number, Err.Description, Err.Source & "." & sPROC_NAME, App.Title, vbCritical

End Function


'*******************************************************************
'Function: LoadSection
'Author: Scott Foster
'Date: 10-22-01
'Returns: True/False
'Purpose: This loads the data from the db
'*******************************************************************
Public Function LoadSection(ByVal PatientID As Long, ByVal DentalEncounterID As Long) As Boolean

'Dim sSQL As String, rs As Recordset, rsPrev As Recordset
'Const sPROC_NAME = "LoadSection"
'
'On Error GoTo ErrHnd
'
'    mvarPatientID = PatientID
'    mvarDentalEncounterID = DentalEncounterID
'
'    'Load the values
'    sSQL = ""
'
'    Set rs = goconn.OpenRecordset(sSQL)
'
'    If rs.EOF Then
'        dfDirtyFlag = dfAdd
'        'Init the values here to ""
'    Else
'        dfDirtyFlag = dfNoAction
'        'Set the values from the db including "" to allow for nulls
'    End If
'
'    rs.Close
'
''    'Find the previous encounter and load it's values to compare with the given encounter
''    ' to determine what values need to go into the note.
''    sSQL = "SELECT " _
''        & " From " _
''        & " Where " _
''        & " AND (DENTAL_ENCOUNTER.PATIENT_UNIT_ID = " & mvarPatientID & ") AND" _
''        & " (DENTAL_ENCOUNTER.DENTAL_ENCOUNTER_ID < " & mvarDentalEncounterID & ")" _
''        & " ORDER BY DENTAL_ENCOUNTER.DENTAL_ENCOUNTER_ID DESC"
''    Set rsPrev = goconn.OpenRecordset(sSQL)
''
''    If rsPrev.EOF Then
''        'The prev value is ""
''
''    Else
''        'Populate the previous encounter's values
''
''    End If
''
''    rsPrev.Close
''    Set rsPrev = Nothing
'
'    Set rs = Nothing
'
'    LoadSection = True
'
'Exit Function
'
'ErrHnd:
'    goCWShared.ShowVBError Err.Number, Err.Description, Err.Source & "." & sPROC_NAME, App.Title, vbCritical
End Function

Private Function UpdateDirtyFlag()

'    If dfDirtyFlag = dfNoAction Then
'        dfDirtyFlag = dfUpdate
'    End If
    
End Function

Private Sub Class_Initialize()

'    dfDirtyFlag = dfNoAction
    
End Sub


