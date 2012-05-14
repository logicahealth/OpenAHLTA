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
Attribute VB_Name = "clsExProsGenPro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements clsISection

Private mvarDentalEncounterID As Long
Private mvarPatientID As Long

Private dfDirtyFlag As DirtyFlagEnum

Private mvarTxDiscussed As Boolean
Private mvarTxAccepted As Boolean
Private mvarGenProsNotes As String
Private mvarSupportingProcs As String

'Max length for text fields
Private Const cGenProsNotesLen As Integer = 2000

Public Property Let SupportingProcs(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SupportingProcs = 5
    mvarSupportingProcs = vData
    UpdateDirtyFlag
End Property


Public Property Get SupportingProcs() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SupportingProcs
    
    If mvarSupportingProcs = "" Then
        mvarSupportingProcs = gcSupportingProcs
    End If
    
    SupportingProcs = mvarSupportingProcs
End Property



'*******************************************************************
'Function: SupportingProcsDisplay
'Author: Scott Foster
'Date: 10/31/01
'Returns: String
'Purpose: To get the display string for supporting procedures
'*******************************************************************
Public Function SupportingProcsDisplay() As String

Dim oItems As New CMultiListItems
Dim oItem As CMultiListItem

    oItems.ListType = Standard
    oItems.LoadFromValueString mvarSupportingProcs
    SupportingProcsDisplay = oItems.Value_String_Display

End Function

Public Property Let GenProsNotes(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SurgicalNotes = 5
    mvarGenProsNotes = Left(vData, cGenProsNotesLen)
    UpdateDirtyFlag
End Property


Public Property Get GenProsNotes() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SurgicalNotes
    GenProsNotes = mvarGenProsNotes
End Property


Public Property Let TxDiscussed(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.HistoryReviewed = 5
    mvarTxDiscussed = vData
    UpdateDirtyFlag
End Property


Public Property Get TxDiscussed() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.HistoryReviewed
    TxDiscussed = mvarTxDiscussed
End Property



Public Property Let TxAccepted(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RadsReviewed = 5
    mvarTxAccepted = vData
    UpdateDirtyFlag
End Property


Public Property Get TxAccepted() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RadsReviewed
    TxAccepted = mvarTxAccepted
End Property


Private Sub Class_Initialize()

    dfDirtyFlag = dfNoAction
    mvarSupportingProcs = gcSupportingProcs
    
End Sub


Private Property Get clsISection_DentalEncounterID() As Long

    clsISection_DentalEncounterID = mvarDentalEncounterID

End Property


'*******************************************************************
'Function: GetRTFNote
'Author: Scott Foster
'Modified by: ASamedov
'Date: 12/26/01
'Returns: A string that is displayed in the narrative summary for this encounter
'Purpose: To compare the values from this encounter and the previous encounter and
'return a text note.
'*******************************************************************
Private Function clsISection_GetRTFNote() As String

Dim sTemp As String
    
    If mvarTxDiscussed = True Then
        sTemp = sTemp & crtfStartControlName & "" & "Treatment plan discussed with patient;" & crtfEndControlName
    End If
    
    If mvarTxAccepted = True Then
        sTemp = sTemp & crtfStartControlName & "" & "Treatment plan accepted by patient;" & crtfEndControlName
    End If
       
    If SupportingProcsDisplay <> "" Then
        sTemp = sTemp & crtfStartControlName & "Supporting procs: " & SupportingProcsDisplay & "; " & crtfEndControlName
    End If
    
    If mvarGenProsNotes <> "" Then
        sTemp = sTemp & crtfStartControlName & "Comment: " & mvarGenProsNotes & "; " & crtfEndControlName
    End If
    clsISection_GetRTFNote = sTemp
    
End Function


Private Function clsISection_GetNote() As String

    clsISection_GetNote = "Section no longer valid."
    
End Function


'*******************************************************************
'Function: clsISection_LoadSection
'Author: ASamedov
'Date: 12-24-01
'Returns: True/False
'Purpose: This loads the data from the db
'*******************************************************************
Private Function clsISection_LoadSection(ByVal PatientID As Long, ByVal DentalEncounterID As Long) As Boolean
Dim sSQL As String, rs As Recordset, rsPrev As Recordset
Const sPROC_NAME = "LoadSection"

On Error GoTo ErrHnd

    mvarDentalEncounterID = DentalEncounterID
    
    'Load the values
    sSQL = "Select TxDiscussed, TxAccepted, Comments, SupportingProcedures" _
        & " From TEMPL_EXAM_PROS_GENPROS" _
        & " Where TEMPL_EXAM_PROS_GENPROS.DENTAL_ENCOUNTER_ID = " & mvarDentalEncounterID

    Set rs = goconn.OpenRecordset(sSQL)
    
    If rs.EOF Then
        dfDirtyFlag = dfAdd
        'Init the values here to ""
        mvarTxDiscussed = False
        mvarTxAccepted = False
        mvarGenProsNotes = ""
        
    Else
        dfDirtyFlag = dfNoAction
        'Set the values from the db including "" to allow for nulls
        mvarTxDiscussed = CBool(Val(rs("TXDISCUSSED") & ""))
        mvarTxAccepted = CBool(Val(rs("TXACCEPTED") & ""))
        mvarGenProsNotes = rs("COMMENTS") & ""
        mvarSupportingProcs = rs("SUPPORTINGPROCEDURES") & ""
    End If
    
    rs.Close
    Set rs = Nothing
    
    clsISection_LoadSection = True
    
Exit Function

ErrHnd:
    goCWShared.ShowVBError Err.Number, Err.Description, Err.Source & "." & sPROC_NAME, App.Title, vbCritical
End Function



Private Property Get clsISection_PatientID() As Long

    clsISection_PatientID = mvarPatientID

End Property


'*******************************************************************
'Function: clsISection_SaveSection
'Author: ASamedov
'Date: 12-24-01
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
        
    ElseIf dfDirtyFlag = dfAdd Then
        'The record does not exist, so create an insert statement
        sSQL = "Insert into TEMPL_EXAM_PROS_GENPROS (DENTAL_ENCOUNTER_ID,TXDISCUSSED, TXACCEPTED, COMMENTS, SUPPORTINGPROCEDURES) Values" _
            & " (" & mvarDentalEncounterID & ", " & CInt(mvarTxDiscussed) & ", " & CInt(mvarTxAccepted) & ", '" & mvarGenProsNotes & "', '" & mvarSupportingProcs & "')"
        
    ElseIf dfDirtyFlag = dfUpdate Then
        'The record exists, so create an update statement
        sSQL = "Update TEMPL_EXAM_PROS_GENPROS Set" _
            & " TXDISCUSSED = " & CInt(mvarTxDiscussed) & ", " _
            & " TXACCEPTED = " & CInt(mvarTxAccepted) & ", " _
            & " COMMENTS = '" & RQ(mvarGenProsNotes) & "', " _
            & " SUPPORTINGPROCEDURES = '" & mvarSupportingProcs & "'" _
            & " Where DENTAL_ENCOUNTER_ID = " & mvarDentalEncounterID
            
    End If
    
    goconn.ExecuteSQL sSQL
    
    dfDirtyFlag = dfNoAction
    
    clsISection_SaveSection = True

Exit Function

ErrHnd:
    goCWShared.ShowVBError Err.Number, Err.Description, Err.Source & "." & sPROC_NAME, App.Title, vbCritical

End Function

Private Function UpdateDirtyFlag()

    If dfDirtyFlag = dfNoAction Then
        dfDirtyFlag = dfUpdate
    End If
    
End Function



