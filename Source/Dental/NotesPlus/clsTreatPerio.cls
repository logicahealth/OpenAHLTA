VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTreatPerio"
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

Private dfDirtyFlag As DirtyFlagEnum

Private mvarConsentGiven As Boolean
Private mvarRadsReviewed As Boolean
Private mvarHistoryReviewed As Boolean
Private mvarPostOpInstructions As String
Private mvarSurgicalNotes As String
Private mvarSupportingProcs As String

'Max length for text fields
Private Const cPostOpInstructionsLen As Integer = 2000
Private Const cSurgicalNotesLen As Integer = 2000

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

Public Property Let SurgicalNotes(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SurgicalNotes = 5
    mvarSurgicalNotes = Left(vData, cSurgicalNotesLen)
    UpdateDirtyFlag
End Property


Public Property Get SurgicalNotes() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SurgicalNotes
    SurgicalNotes = mvarSurgicalNotes
End Property



Public Property Let PostOpInstructions(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PostOpInstructions = 5
    mvarPostOpInstructions = Left(vData, cPostOpInstructionsLen)
    UpdateDirtyFlag
End Property


Public Property Get PostOpInstructions() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PostOpInstructions
    PostOpInstructions = mvarPostOpInstructions
End Property



Public Property Let HistoryReviewed(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.HistoryReviewed = 5
    mvarHistoryReviewed = vData
    UpdateDirtyFlag
End Property


Public Property Get HistoryReviewed() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.HistoryReviewed
    HistoryReviewed = mvarHistoryReviewed
End Property



Public Property Let RadsReviewed(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RadsReviewed = 5
    mvarRadsReviewed = vData
    UpdateDirtyFlag
End Property


Public Property Get RadsReviewed() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RadsReviewed
    RadsReviewed = mvarRadsReviewed
End Property



Public Property Let ConsentGiven(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ConsentGiven = 5
    mvarConsentGiven = vData
    UpdateDirtyFlag
End Property


Public Property Get ConsentGiven() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ConsentGiven
    ConsentGiven = mvarConsentGiven
End Property




Private Sub Class_Initialize()

    dfDirtyFlag = dfNoAction
    mvarSupportingProcs = gcSupportingProcs
    
End Sub


Private Property Get clsISection_DentalEncounterID() As Long

    clsISection_DentalEncounterID = mvarDentalEncounterID

End Property


'*******************************************************************
'Function: GetNote
'Author: Scott Foster
'Date: 10/24/01
'Returns: A string that is displayed in the narrative summary for this encounter
'Purpose: To compare the values from this encounter and the previous encounter and
'return a text note.
'*******************************************************************
Private Function clsISection_GetNote() As String

Dim sTemp As String
    
    If mvarConsentGiven = True Then
        sTemp = sTemp & vbTab & "Patient consent given." & vbCrLf
    End If
    
    If mvarRadsReviewed = True Then
        sTemp = sTemp & vbTab & "Radiographs reviewed." & vbCrLf
    End If
    
    If mvarHistoryReviewed = True Then
        sTemp = sTemp & vbTab & "Medical history reviewed." & vbCrLf
    End If
    
    If SupportingProcsDisplay <> "" Then
        sTemp = sTemp & vbTab & "Supporting procs: " & SupportingProcsDisplay & vbCrLf
    End If
    
    If mvarPostOpInstructions <> "" Then
        sTemp = sTemp & vbTab & "Post-op instructions: " & mvarPostOpInstructions & vbCrLf
    End If
    
    If mvarSurgicalNotes <> "" Then
        sTemp = sTemp & vbTab & "Surgical notes: " & mvarSurgicalNotes & vbCrLf
    End If
    
    If sTemp <> "" Then
            sTemp = "Perio Treatment: " & vbCrLf & sTemp
    End If
    
    clsISection_GetNote = sTemp
    
End Function


Private Function clsISection_GetRTFNote() As String

Dim sTemp As String
    
    If mvarConsentGiven = True Then
        sTemp = sTemp & "Patient consent given." & "; "
    End If
    
    If mvarRadsReviewed = True Then
        sTemp = sTemp & "Radiographs reviewed." & "; "
    End If
    
    If mvarHistoryReviewed = True Then
        sTemp = sTemp & "Medical history reviewed." & "; "
    End If
    
    'Remove the last "; "
    If sTemp <> "" Then
        sTemp = crtfStartControlName & Left(sTemp, Len(sTemp) - Len("; ")) & crtfEndControlName
    End If
    
    If SupportingProcsDisplay <> "" Then
        sTemp = sTemp & crtfStartControlName & "Supporting procedures: " & SupportingProcsDisplay & crtfEndControlName
    End If
    
    If mvarPostOpInstructions <> "" Then
        sTemp = sTemp & crtfStartControlName & "Post-op instructions: " & mvarPostOpInstructions & crtfEndControlName
    End If
    
    If mvarSurgicalNotes <> "" Then
        sTemp = sTemp & crtfStartControlName & "Surgical Comment: " & mvarSurgicalNotes & crtfEndControlName
    End If
    
    
    
'    If sTemp <> "" Then
'            sTemp = "Perio Treatment: " & vbCrLf & sTemp
'    End If
    
    clsISection_GetRTFNote = sTemp
    
End Function


'*******************************************************************
'Function: clsISection_LoadSection
'Author: Scott Foster
'Date: 10-24-01
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
    sSQL = "Select CONSENT_GIVEN, RADIO_REV, HISTORY_REV, POST_OP_COMMENTS, SURGICAL_NOTES, SUPPORTING_PROCS" _
        & " From TEMPL_TREAT_PERIO" _
        & " Where DENTAL_ENCOUNTER_ID = " & mvarDentalEncounterID

    Set rs = goconn.OpenRecordset(sSQL)
    
    If rs.EOF Then
        dfDirtyFlag = dfAdd
        'Init the values here to ""
        mvarConsentGiven = False
        mvarRadsReviewed = False
        mvarHistoryReviewed = False
        mvarPostOpInstructions = ""
        mvarSurgicalNotes = ""
        
    Else
        dfDirtyFlag = dfNoAction
        'Set the values from the db including "" to allow for nulls
        mvarConsentGiven = CBool(Val(rs("CONSENT_GIVEN") & ""))
        mvarRadsReviewed = CBool(Val(rs("RADIO_REV") & ""))
        mvarHistoryReviewed = CBool(Val(rs("HISTORY_REV") & ""))
        mvarPostOpInstructions = rs("POST_OP_COMMENTS") & ""
        mvarSurgicalNotes = rs("SURGICAL_NOTES") & ""
        mvarSupportingProcs = rs("SUPPORTING_PROCS") & ""
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
'Author: Scott Foster
'Date: 10-22-01
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
        sSQL = "Insert into TEMPL_TREAT_PERIO (DENTAL_ENCOUNTER_ID,CONSENT_GIVEN, RADIO_REV, HISTORY_REV, POST_OP_COMMENTS, SURGICAL_NOTES, SUPPORTING_PROCS) Values" _
            & " (" & mvarDentalEncounterID & ", " & CInt(mvarConsentGiven) & ", " & CInt(mvarRadsReviewed) & ", " & CInt(mvarHistoryReviewed) & ", '" & mvarPostOpInstructions & "', '" & mvarSurgicalNotes & "', '" & mvarSupportingProcs & "')"
        
    ElseIf dfDirtyFlag = dfUpdate Then
        'The record exists, so create an update statement
        sSQL = "Update TEMPL_TREAT_PERIO Set" _
            & " CONSENT_GIVEN = " & CInt(mvarConsentGiven) & ", " _
            & " RADIO_REV = " & CInt(mvarRadsReviewed) & ", " _
            & " HISTORY_REV = " & CInt(mvarHistoryReviewed) & ", " _
            & " POST_OP_COMMENTS = '" & RQ(mvarPostOpInstructions) & "', " _
            & " SURGICAL_NOTES = '" & RQ(mvarSurgicalNotes) & "', " _
            & " SUPPORTING_PROCS = '" & RQ(mvarSupportingProcs) & "'" _
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

