VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEtiology"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Implements clsISection

'Used to determine if the data should be inserted or updated when modified.
Private bSavedToDB As Boolean

Private mvarPatientID As Long
Private mvarDentalEncounterID As Long
Private dfDirtyFlag As DirtyFlagEnum
Private mvarValue As String
Private sValuePrev As String
    
'Constant used to define Etiology selections
'The format is Key,Description,primary,secondary|[next record]
Private Const cEtiologyValues As String = "Plaque^Plaque^false^false|" _
    & "Calculus^Calculus^false^false|Tobacco related-current smoker^Tobacco related-current smoker^false^false|" _
    & "Tobacco related-smokeless user^Tobacco related-smokeless user^false^false|" _
    & "Tooth brush trauma^Tooth brush trauma^false^false|Open contact^Open contact^false^false|" _
    & "Defective restoration^Defective restoration^false^false|Systemic disease^Systemic disease^false^false|" _
    & "Trauma from occlusion^Trauma from occlusion^false^false|Root fracture^Root fracture^false^false|" _
    & "Root proximity^Root proximity^false^false|Factitious injury^Factitious injury^false^false|" _
    & "Bruxism^Bruxism^false^false|Developmental anomalies^Developmental anomalies^false^false|" _
    & "Frenum^Frenum^false^false|Other^Other^false^false|Unknown^Unknown^false^false"

'*******************************************************************
'Function: GetDelimitedDisplay
'Author: Scott Foster
'Date: 10/19/2001
'Returns:
'Purpose: To return a comma delimited list of descriptions with primary first and
'secondary last.
'*******************************************************************
'
Public Function GetDelimitedDisplay() As String

Dim oItems As New CMultiListItems
Dim oItem As CMultiListItem
Dim sTemp As String
Dim bFound As Boolean

    oItems.ListType = PrimarySecondaryList
    oItems.LoadFromValueString mvarValue
    
    'Get all the primary items first
    bFound = False
    sTemp = "Primary: "
    For Each oItem In oItems
        If oItem.Primary = True Then
            sTemp = sTemp & oItem.Description & ", "
            bFound = True
        End If
    Next
    'Remove the last ", "
    If bFound = True Then
        sTemp = Mid(sTemp, 1, Len(sTemp) - Len(", "))
    End If
    
    'Get all the secondary items next
    bFound = False
    sTemp = sTemp & vbCrLf & "Secondary: "
    For Each oItem In oItems
        If oItem.Secondary = True Then
            sTemp = sTemp & oItem.Description & ", "
            bFound = True
        End If
    Next
    'Remove the last ", "
    If bFound = True Then
        sTemp = Mid(sTemp, 1, Len(sTemp) - Len(", "))
    End If
    
    GetDelimitedDisplay = sTemp
    
End Function


Public Property Let Value(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Value = 5
    
    UpdateDirtyFlag
    mvarValue = vData

End Property


Public Property Get Value() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Value
    Value = mvarValue
End Property






Private Function UpdateDirtyFlag()

    If bSavedToDB = False Then
        dfDirtyFlag = dfAdd
    ElseIf dfDirtyFlag = dfNoAction Then
        dfDirtyFlag = dfUpdate
    End If
    
End Function

Private Sub Class_Initialize()

    dfDirtyFlag = dfNoAction
    
End Sub




Private Property Get clsISection_DentalEncounterID() As Long
    
    clsISection_DentalEncounterID = mvarDentalEncounterID

End Property


'*******************************************************************
'Function: GetNote
'Author: Scott Foster
'Date: 10/19/01
'Returns: A string that is displayed in the narrative summary for this encounter
'Purpose: To compare the values from this encounter and the previous encounter and
'return a text note.
'*******************************************************************
Private Function clsISection_GetNote() As String

Dim oItems As New CMultiListItems
Dim oItem As CMultiListItem
Dim sPriTemp As String
Dim sSecTemp As String
Dim bFound As Boolean
Dim sSearch As String

    oItems.ListType = PrimarySecondaryList
    oItems.LoadFromValueString mvarValue
    
    bFound = False
    'Get the primaries first
    For Each oItem In oItems
        If oItem.Primary = True Then
            sPriTemp = sPriTemp & " " & oItem.Description & ", "
            bFound = True
        End If
    Next
    'Remove the last ", "
    If bFound = True Then
        sPriTemp = Mid(sPriTemp, 1, Len(sPriTemp) - Len(", "))
        sPriTemp = vbTab & "Primary: " & sPriTemp & vbCrLf
    End If
    
    bFound = False
    'Get the secondaries next
    For Each oItem In oItems
        If oItem.Secondary = True Then
            sSecTemp = sSecTemp & " " & oItem.Description & ", "
            bFound = True
        End If
    Next
    'Remove the last ", "
    If bFound = True Then
        sSecTemp = Mid(sSecTemp, 1, Len(sSecTemp) - Len(", "))
        sSecTemp = vbTab & "Secondary " & sSecTemp
    End If
   
    If sPriTemp <> "" Or sSecTemp <> "" Then
        clsISection_GetNote = "Etiology:" & vbCrLf & sPriTemp & sSecTemp
   End If
   
End Function


Private Function clsISection_GetRTFNote() As String

Dim oItems As New CMultiListItems
Dim oItem As CMultiListItem
Dim sPriTemp As String
Dim sSecTemp As String
Dim bFound As Boolean
Dim sSearch As String

    oItems.ListType = PrimarySecondaryList
    oItems.LoadFromValueString mvarValue
    
    bFound = False
    'Get the primaries first
    For Each oItem In oItems
        If oItem.Primary = True Then
            sPriTemp = sPriTemp & oItem.Description & ", "
'            bFound = True
        ElseIf oItem.Secondary = True Then
            sSecTemp = sSecTemp & oItem.Description & ", "
'            bFound = True
        End If
    Next
    'Remove the last ", "
    If sPriTemp <> "" Then
        sPriTemp = Mid(sPriTemp, 1, Len(sPriTemp) - Len(", "))
        sPriTemp = "Primary: " & sPriTemp
    End If
    
'    bFound = False
'    'Get the secondaries next
'    For Each oItem In oItems
'        If oItem.Secondary = True Then
'            sSecTemp = sSecTemp & " " & oItem.Description & ", "
'            bFound = True
'        End If
'    Next
    'Remove the last ", "
    If sSecTemp <> "" Then
        sSecTemp = Mid(sSecTemp, 1, Len(sSecTemp) - Len(", "))
        sSecTemp = "Secondary: " & sSecTemp
    End If
   
    If sPriTemp <> "" Or sSecTemp <> "" Then
        If sSecTemp <> "" And sPriTemp <> "" Then
            'Insert a "; " if both have a value
            sPriTemp = sPriTemp & "; "
        End If
        clsISection_GetRTFNote = crtfStartControlName & "Etiology: " & sPriTemp & sSecTemp & crtfEndControlName
   End If
    
End Function


'*******************************************************************
'Function: clsISection_LoadSection
'Author: Scott Foster
'Date: 10-22-01
'Returns: True/False
'Purpose: This loads the data from the db
'*******************************************************************
Private Function clsISection_LoadSection(ByVal PatientID As Long, ByVal DentalEncounterID As Long) As Boolean

'This loads the data from the db

Dim sSQL As String, rs As Recordset, rsPrev As Recordset
Const sPROC_NAME = "LoadSection"

On Error GoTo ErrHnd

    mvarPatientID = PatientID
    mvarDentalEncounterID = DentalEncounterID
    
    'Load the values
    sSQL = "SELECT VALUE FROM TEMPL_ETIOLOGY WHERE DENTAL_ENCOUNTER_ID = " & mvarDentalEncounterID
    
    Set rs = goconn.OpenRecordset(sSQL)
    
    If rs.EOF Then
        bSavedToDB = False
        dfDirtyFlag = dfAdd
        'Init the values here to the default
        mvarValue = cEtiologyValues
    Else
        bSavedToDB = True
        dfDirtyFlag = dfNoAction
        'Set the values from the db including "" to allow for nulls
        mvarValue = rs("Value") & ""
    End If
    
    rs.Close
    
'    'Find the previous encounter and load it's values to compare with the given encounter
'    ' to determine what values need to go into the note.
'    sSQL = "SELECT TEMPL_ETIOLOGY.VALUE From TEMPL_ETIOLOGY, DENTAL_ENCOUNTER" _
'        & " Where TEMPL_ETIOLOGY.DENTAL_ENCOUNTER_ID = DENTAL_ENCOUNTER.DENTAL_ENCOUNTER_ID" _
'        & " AND (DENTAL_ENCOUNTER.PATIENT_UNIT_ID = " & mvarPatientID & ") AND" _
'        & " (DENTAL_ENCOUNTER.DENTAL_ENCOUNTER_ID < " & mvarDentalEncounterID & ")" _
'        & " ORDER BY DENTAL_ENCOUNTER.DENTAL_ENCOUNTER_ID DESC"
'
'    Set rsPrev = goconn.OpenRecordset(sSQL)
'
'    If rsPrev.EOF Then
'        'The prev value is ""
'        sValuePrev = ""
'    Else
'        'Populate the previous encounter's values
'        sValuePrev = rsPrev("Value") & ""
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
        sSQL = "Insert into TEMPL_ETIOLOGY (DENTAL_ENCOUNTER_ID, VALUE) VALUES" _
            & " (" & mvarDentalEncounterID & ", '" & mvarValue & "')"
            
    ElseIf dfDirtyFlag = dfUpdate Then
        'The record exists, so create an update statement
        sSQL = "Update TEMPL_ETIOLOGY SET VALUE = '" & mvarValue & "' WHERE DENTAL_ENCOUNTER_ID = " _
            & mvarDentalEncounterID
    End If
    
    goconn.ExecuteSQL sSQL
    
    bSavedToDB = True
    dfDirtyFlag = dfNoAction
    
    clsISection_SaveSection = True

Exit Function

ErrHnd:
    goCWShared.ShowVBError Err.Number, Err.Description, Err.Source & "." & sPROC_NAME, App.Title, vbCritical

End Function


