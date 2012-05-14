VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsExamRadFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Implements clsISection

'Used to determine if the data should be inserted or updated when modified.
Private bSavedToDB As Boolean

Private mvarPatientID As Long
Private mvarDentalEncounterID As Long
Private mvarComments As String 'local copy
Private mvarRadiographicFindingsData As String 'local copy
Private mvarRadiographicFindingsDisplay As String 'local copy

Private dfDirtyFlag As DirtyFlagEnum
Public Property Let RadiographicFindingsDisplay(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RadiographicFindingsDisplay = 5
    mvarRadiographicFindingsDisplay = vData
    UpdateDirtyFlag
End Property


Public Property Get RadiographicFindingsDisplay() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RadiographicFindingsDisplay
    RadiographicFindingsDisplay = mvarRadiographicFindingsDisplay
End Property



Public Property Let RadiographicFindingsData(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RadiographicFindingsData = 5
    mvarRadiographicFindingsData = vData
    UpdateDirtyFlag
End Property


Public Property Get RadiographicFindingsData() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RadiographicFindingsData
    RadiographicFindingsData = mvarRadiographicFindingsData
End Property



Public Property Let Comments(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Comments = 5
    mvarComments = vData
    UpdateDirtyFlag
End Property


Public Property Get Comments() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Comments
    Comments = mvarComments
End Property



Private Property Get clsISection_DentalEncounterID() As Long
    
    clsISection_DentalEncounterID = mvarDentalEncounterID

End Property



Private Property Get clsISection_PatientID() As Long
    
    clsISection_PatientID = mvarPatientID

End Property



Private Function clsISection_GetRTFNote() As String

Dim sTemp As String
    
    If mvarRadiographicFindingsDisplay <> "" Then
        sTemp = sTemp & Trim(mvarRadiographicFindingsDisplay) & "; "
    End If
    
    If mvarComments <> "" Then
        sTemp = sTemp & "Comments: " & mvarComments & "; "
    End If
    
    
    If sTemp <> "" Then
        'Remove the last "; "  Moved inside If statement. SM.
        sTemp = Left(sTemp, Len(sTemp) - Len("; "))
        clsISection_GetRTFNote = crtfStartControlName & "Radiographic findings: " & sTemp & crtfEndControlName
    End If

End Function

'*******************************************************************
'Function: GetNote
'Author: Scott Foster
'Date: 10/19/01
'Returns: A string that is displayed in the narrative summary for this encounter
'Purpose: To compare the values from this encounter and the previous encounter and
'return a text note.
'*******************************************************************
Private Function clsISection_GetNote() As String

Dim sTemp As String
    
    If mvarRadiographicFindingsDisplay <> "" Then
        sTemp = sTemp & vbTab & "Radiographic findings: " & mvarRadiographicFindingsDisplay & vbCrLf
    End If
    
    If mvarComments <> "" Then
        sTemp = sTemp & vbTab & "Comments: " & mvarComments & vbCrLf
    End If
    
    If sTemp <> "" Then
        clsISection_GetNote = "Radiographic findings" & vbCrLf & sTemp
    End If
    
End Function

'*******************************************************************
'Function: SaveSection
'Author: Scott Foster
'Date: 10-31-01
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
        sSQL = "INSERT INTO TEMPL_RADIO_FINDINGS" _
            & " (RADIOFINDDATA, RADIOFINDDISPLAY," _
            & " COMMENTS, DENTAL_ENCOUNTER_ID)" _
            & " VALUES ('" & mvarRadiographicFindingsData & "', '" & mvarRadiographicFindingsDisplay & "'," _
            & " '" & RQ(mvarComments) & "', " & mvarDentalEncounterID & ")"
            
    ElseIf dfDirtyFlag = dfUpdate Then
        'The record exists, so create an update statement
        sSQL = "UPDATE TEMPL_RADIO_FINDINGS" _
            & " SET RADIOFINDDATA = '" & mvarRadiographicFindingsData & "'," _
            & " RADIOFINDDISPLAY = '" & mvarRadiographicFindingsDisplay & "'," _
            & " COMMENTS = '" & RQ(mvarComments) & "'" _
            & " Where (DENTAL_ENCOUNTER_ID = " & mvarDentalEncounterID & ")"

    End If
    
    goconn.ExecuteSQL sSQL
    
    bSavedToDB = True
    dfDirtyFlag = dfNoAction
    
    clsISection_SaveSection = True

Exit Function

ErrHnd:
    goCWShared.ShowVBError Err.Number, Err.Description, Err.Source & "." & sPROC_NAME, App.Title, vbCritical

End Function


'*******************************************************************
'Function: LoadSection
'Author: Scott Foster
'Date: 10-31-01
'Returns: True/False
'Purpose: This loads the data from the db
'*******************************************************************
Private Function clsISection_LoadSection(ByVal PatientID As Long, ByVal DentalEncounterID As Long) As Boolean

Dim sSQL As String, rs As Recordset, rsPrev As Recordset
Const sPROC_NAME = "clsISection_LoadSection"

On Error GoTo ErrHnd

    mvarPatientID = PatientID
    mvarDentalEncounterID = DentalEncounterID
    
    If gbUseStoredProc = True Then
        Dim oCmd As New CHCSII_Command.CHCSII_Cmd
        
        oCmd.SPName = "Dent_Notes_RadFind_Load"

        Call oCmd.AddParam(False, False, "textString", mvarDentalEncounterID)
        Call oCmd.AddParam(True, True, "", "")
        
        Set rs = goconn.OpenRecordsetSP(oCmd)
        
        If rs.EOF Then
            bSavedToDB = False
            dfDirtyFlag = dfAdd
            'Init the values here to ""
            mvarRadiographicFindingsData = ""
            mvarRadiographicFindingsDisplay = ""
            mvarComments = ""
        Else
            bSavedToDB = True
            dfDirtyFlag = dfNoAction
            'Set the values from the db including "" to allow for nulls
            mvarComments = rs("COMMENTS") & ""
            mvarRadiographicFindingsData = rs("RADIOFINDDATA") & ""
            mvarRadiographicFindingsDisplay = rs("RADIOFINDDISPLAY") & ""
        End If
            
    Else
    
        'KDunne 5/2/2002
        'SCR 23430 - Size of the db fields RADIOFINDDATA and RADIOFINDDISPLAY were increased to
        '            4000. Tuxedo can only handle 4000 total bytes of data coming across so the
        '            select statement was broken down into 3 seperate parts.
        '
        
        'Load the values
        
        sSQL = "SELECT DENTAL_ENCOUNTER_ID, COMMENTS From TEMPL_RADIO_FINDINGS" _
            & " WHERE (DENTAL_ENCOUNTER_ID = " & mvarDentalEncounterID & ")"
    
        sSQL = "SELECT RADIOFINDDATA, RADIOFINDDISPLAY, COMMENTS From TEMPL_RADIO_FINDINGS" _
            & " WHERE (DENTAL_ENCOUNTER_ID = " & mvarDentalEncounterID & ")"
        
        Set rs = goconn.OpenRecordset(sSQL)
        
        If rs.EOF Then
            bSavedToDB = False
            dfDirtyFlag = dfAdd
            'Init the values here to ""
            mvarRadiographicFindingsData = ""
            mvarRadiographicFindingsDisplay = ""
            mvarComments = ""
        Else
            bSavedToDB = True
            dfDirtyFlag = dfNoAction
            'Set the values from the db including "" to allow for nulls
            mvarComments = rs("COMMENTS") & ""
            
            sSQL = "SELECT RADIOFINDDATA From TEMPL_RADIO_FINDINGS" _
                & " WHERE (DENTAL_ENCOUNTER_ID = " & mvarDentalEncounterID & ")"
            
            Set rs = goconn.OpenRecordset(sSQL)
    
            If Not rs.EOF Then
                mvarRadiographicFindingsData = rs("RADIOFINDDATA") & ""
            End If
            
            sSQL = "SELECT RADIOFINDDISPLAY From TEMPL_RADIO_FINDINGS" _
                & " WHERE (DENTAL_ENCOUNTER_ID = " & mvarDentalEncounterID & ")"
            
            Set rs = goconn.OpenRecordset(sSQL)
            
            If Not rs.EOF Then
                mvarRadiographicFindingsDisplay = rs("RADIOFINDDISPLAY") & ""
            End If
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

    If bSavedToDB = False Then
        dfDirtyFlag = dfAdd
    ElseIf dfDirtyFlag = dfNoAction Then
        dfDirtyFlag = dfUpdate
    End If
    
End Function

Private Sub Class_Initialize()

    dfDirtyFlag = dfNoAction
    
End Sub

