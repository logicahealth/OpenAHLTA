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
Attribute VB_Name = "clsTreatGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements clsISection

Private mvarPatientID As Long
Private mvarDentalEncounterID As Long
Private mvarWorkloadID As Long
Private mvarComments As String
Private mvarSupportingProcs As String

Private dfDirtyFlag As DirtyFlagEnum


Public Property Get Comments() As String

    Comments = mvarComments
    
End Property

Public Property Let Comments(vData As String)
    
    mvarComments = vData
    UpdateDirtyFlag

End Property

Public Property Let SupportingProcs(vData As String)

    mvarSupportingProcs = vData
    UpdateDirtyFlag
    
End Property

Public Property Get SupportingProcs() As String
    
    If mvarSupportingProcs = "" Then
        mvarSupportingProcs = gcSupportingProcs
    End If
    
    SupportingProcs = mvarSupportingProcs
    
End Property

Public Property Let WorkloadId(vData As Long)
    
    mvarWorkloadID = vData
    UpdateDirtyFlag
    
End Property

Public Property Get WorkloadId() As Long
    
    WorkloadId = mvarWorkloadID
    
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
'*******************************************************************
'Function: SupportingProcsDisplay
'Author: Sean McAvoy
'Date: 1/2/02
'Returns: String
'Purpose: To get the display string for supporting procedures
'*******************************************************************
Private Function SupportingProcsDisplay() As String

Dim oItems As New CMultiListItems

    oItems.ListType = Standard
    oItems.LoadFromValueString mvarSupportingProcs
    SupportingProcsDisplay = oItems.Value_String_Display
Set oItems = Nothing
End Function



Private Function clsISection_GetNote() As String

    clsISection_GetNote = "Section No Longer Valid."
End Function

'*******************************************************************
'Function: GetRTFNote
'Author: Sean McAvoy
'Date: 2/15/02
'Returns: A string that is displayed in the narrative summary for this encounter
'Purpose: To compare the values from this encounter and the previous encounter and
'return a text note.
'*******************************************************************
Private Function clsISection_GetRTFNote() As String
Dim sTxt As String
Dim sTemp As String
    
    If mvarComments <> "" Then
        sTemp = crtfStartControlName & "Procedure Comments: " & mvarComments & crtfEndControlName
    End If
    sTxt = SupportingProcsDisplay
    If sTxt <> "" Then
        sTemp = sTemp + crtfStartControlName & "Supporting Procedures: " & sTxt & crtfEndControlName
    End If

    If sTemp <> "" Then
        clsISection_GetRTFNote = sTemp
    End If
    
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
        'UpdateWorkload mvarWorkloadID, dfDirtyFlag, mvarDentalEncounterID, GetWorkloadCode, 1
    End If
        
    If dfDirtyFlag = dfAdd Then
        'The record does not exist, so create an insert statement
        sSQL = "INSERT INTO TEMPL_USER_COMMENTS (DENTAL_ENCOUNTER_ID, TREAT_GEN, TREAT_GEN_SUPP_PROC)" _
            & " VALUES (" & mvarDentalEncounterID & ", '" & RQ(mvarComments) & "', '" & RQ(mvarSupportingProcs) & "')"
        
    ElseIf dfDirtyFlag = dfUpdate Then
        'The record exists, so create an update statement
        sSQL = "UPDATE TEMPL_USER_COMMENTS SET TREAT_GEN = '" & RQ(mvarComments) & "', TREAT_GEN_SUPP_PROC = '" & RQ(mvarSupportingProcs) & "'" _
            & " WHERE (DENTAL_ENCOUNTER_ID = " & mvarDentalEncounterID & ")"
            
    ElseIf dfDirtyFlag = dfdelete Then
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
    sSQL = "SELECT TREAT_GEN, TREAT_GEN_SUPP_PROC From TEMPL_USER_COMMENTS WHERE (DENTAL_ENCOUNTER_ID = " & mvarDentalEncounterID & ")"
    
    Set rs = goconn.OpenRecordset(sSQL)
    
    If rs.EOF Then
        dfDirtyFlag = dfAdd
        'Init the values here to ""
        mvarWorkloadID = 0
        mvarComments = ""
        mvarSupportingProcs = ""
        
    Else
        dfDirtyFlag = dfNoAction
        'Set the values from the db including "" to allow for nulls
        mvarWorkloadID = 0 ' Val(rs("Workload_ID") & "")
        mvarSupportingProcs = rs("TREAT_GEN_SUPP_PROC") & ""
        mvarComments = rs("Treat_Gen") & ""
    End If
    
    rs.Close
    
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






