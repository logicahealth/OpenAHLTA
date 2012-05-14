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
Attribute VB_Name = "clsGingiva"
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

Private mvarColor As String 'local copy
Private mvarTexture As String 'local copy
Private mvarConsistency As String 'local copy
Private mvarContour As String 'local copy
Private mvarMargins As String 'local copy
Private mvarPapillae As String 'local copy
Private mvarBleeding As String 'local copy
Private mvarSuppuration As String 'local copy
Private mvarTreatPhase As String 'local copy
'Private mvarPerioNote As String 'local copy

Private mvarDentalEncounterID As Long 'local copy
Private mvarPatientID As Long 'local copy

'Constant to define the possible treatment phases
Private Const cTreatPhases As String = " |Initial Therapy|Surgical Therapy|Reevaluation|Therapy Completed|Maintenance"

'Constants used to define the columns in the Exam-Perio-Gingiva grid
Private Const cColorCol As Integer = 0
Private Const cTextureCol As Integer = 1
Private Const cConsistencyCol As Integer = 2
Private Const cContourCol As Integer = 3
Private Const cMarginsCol As Integer = 4
Private Const cPapillaeCol As Integer = 5
Private Const cBleedingCol As Integer = 6
Private Const cSuppurationCol As Integer = 7

'Constants used to define options in the Exam-Perio-Gingiva grid
Private Const cColorValues As String = " |Coral pink|Pink|Light pink|Erythematous(red)|Magenta(purple)|Cyanotic(blue)|Pigment(light)|Pigment(dark)"
Private Const cTextureValues As String = " |Stippled|Smooth|Pebbly"
Private Const cConsistencyValues As String = " |Firm|Fibrotic|Edematous|Boggy"
Private Const cContourValues As String = " |Scalloped|Flat|Irregular"
Private Const cMarginsValues As String = " |Knife edged|Rolled|Rounded"
Private Const cPapillaeValues As String = " |Fill embrasure|Blunted|Swollen|Punched out|Hyperplastic"
Private Const cBleedingValues As String = " |Absent|Localized|Generalized|Slight|Moderate|Heavy"
Private Const cSuppurationValues As String = " |Absent|Localized|Generalized"

Private dfDirtyFlag As DirtyFlagEnum

Public Function GetTreatPhases() As String

    GetTreatPhases = cTreatPhases
    
End Function


Public Property Let PerioNote(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PerioNote = 5
    mvarPerioNote = Left(vData, 2000)
    UpdateDirtyFlag
End Property


Public Property Get PerioNote() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PerioNote
    PerioNote = mvarPerioNote
End Property



Public Property Let TreatPhase(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TreatPhase = 5
    mvarTreatPhase = Trim(vData)
    UpdateDirtyFlag
End Property


Public Property Get TreatPhase() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TreatPhase
    TreatPhase = mvarTreatPhase
End Property




Public Property Get ColColor() As Integer
    
    ColColor = cColorCol
    
End Property
Public Property Get ColTexture() As Integer
    
    ColTexture = cTextureCol
    
End Property
Public Property Get ColConsistency() As Integer
    
    ColConsistency = cConsistencyCol
    
End Property
Public Property Get ColContour() As Integer
    
    ColContour = cContourCol
    
End Property
Public Property Get ColMargins() As Integer
    
    ColMargins = cMarginsCol
    
End Property
Public Property Get ColPapillae() As Integer
    
    ColPapillae = cPapillaeCol
    
End Property
Public Property Get ColBleeding() As Integer
    
    ColBleeding = cBleedingCol
    
End Property
Public Property Get ColSuppuration() As Integer
    
    ColSuppuration = cSuppurationCol
    
End Property
Public Property Get ColorValues() As String
    
    ColorValues = cColorValues
    
End Property

Public Property Get TextureValues() As String
    
    TextureValues = cTextureValues
    
End Property
Public Property Get ConsistencyValues() As String
    
    ConsistencyValues = cConsistencyValues
    
End Property
Public Property Get ContourValues() As String
    
    ContourValues = cContourValues
    
End Property
Public Property Get MarginsValues() As String
    
    MarginsValues = cMarginsValues
    
End Property
Public Property Get PapillaeValues() As String
    
    PapillaeValues = cPapillaeValues
    
End Property
Public Property Get BleedingValues() As String
    
    BleedingValues = cBleedingValues
    
End Property
Public Property Get SuppurationValues() As String
    
    SuppurationValues = cSuppurationValues
    
End Property

Public Property Let Suppuration(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Suppuration = 5
    UpdateDirtyFlag
    mvarSuppuration = vData
End Property


Public Property Get Suppuration() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Suppuration
    Suppuration = mvarSuppuration
End Property



Public Property Let Bleeding(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Bleeding = 5
    UpdateDirtyFlag
    mvarBleeding = vData
End Property


Public Property Get Bleeding() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Bleeding
    Bleeding = mvarBleeding
End Property



Public Property Let Papillae(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Papillae = 5
    UpdateDirtyFlag
    mvarPapillae = vData
End Property


Public Property Get Papillae() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Papillae
    Papillae = mvarPapillae
End Property



Public Property Let Margins(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Margins = 5
    UpdateDirtyFlag
    mvarMargins = vData
End Property


Public Property Get Margins() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Margins
    Margins = mvarMargins
End Property



Public Property Let Contour(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Contour = 5
    UpdateDirtyFlag
    mvarContour = vData
End Property


Public Property Get Contour() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Contour
    Contour = mvarContour
End Property



Public Property Let Consistency(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Consistency = 5
    UpdateDirtyFlag
    mvarConsistency = vData
End Property


Public Property Get Consistency() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Consistency
    Consistency = mvarConsistency
End Property



Public Property Let Texture(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Texture = 5
    UpdateDirtyFlag
    mvarTexture = vData
End Property


Public Property Get Texture() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Texture
    Texture = mvarTexture
End Property



Public Property Let Color(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Color = 5
    UpdateDirtyFlag
    mvarColor = vData
End Property


Public Property Get Color() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Color
    Color = mvarColor
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
'Purpose: This function creates the note string for Gingiva
'return a text note.
'*******************************************************************
Private Function clsISection_GetNote() As String

Dim sNote As String
    
    If mvarColor = "" And mvarTexture = "" And mvarConsistency = "" And mvarContour = "" _
        And mvarMargins = "" And mvarPapillae = "" And mvarBleeding = "" And mvarSuppuration = "" Then
        'If everything is blank, then don't add a note.
        sNote = ""
    Else
        sNote = vbTab & "Gingiva: "
        
        If mvarColor <> "" Then
            sNote = sNote & "Color: " & mvarColor & ";  "
        End If
        
        If mvarTexture <> "" Then
            sNote = sNote & "Texture: " & mvarTexture & "; "
        End If
        
        If mvarConsistency <> "" Then
            sNote = sNote & "Consistency: " & mvarConsistency & "; "
        End If
        
        If mvarContour <> "" Then
            sNote = sNote & "Contour: " & mvarContour & "; "
        End If
        
        If mvarMargins <> "" Then
            sNote = sNote & "Margins: " & mvarMargins & "; "
        End If
        
        If mvarPapillae <> "" Then
            sNote = sNote & "Papillae: " & mvarPapillae & "; "
        End If
        
        If mvarBleeding <> "" Then
            sNote = sNote & "Bleeding: " & mvarBleeding & "; "
        End If
        
        If mvarSuppuration <> "" Then
            sNote = sNote & "Suppuration: " & mvarSuppuration
        End If
        
        sNote = sNote & vbCrLf
    End If
    
    If mvarTreatPhase <> "" Then
        sNote = sNote & vbTab & "Treatment Phase: " & mvarTreatPhase & vbCrLf
    End If
    
    If mvarPerioNote <> "" Then
        sNote = sNote & vbTab & "Perio Note: " & mvarPerioNote & vbCrLf
    End If
    
    clsISection_GetNote = sNote
    
End Function


Private Function clsISection_GetRTFNote() As String

Dim sNote As String
    
    If mvarColor = "" And mvarTexture = "" And mvarConsistency = "" And mvarContour = "" _
        And mvarMargins = "" And mvarPapillae = "" And mvarBleeding = "" And mvarSuppuration = "" Then
        'If everything is blank, then don't add a note.
        sNote = ""
    Else
        sNote = crtfStartControlName & "Gingiva: "
        
        If mvarColor <> "" Then
            sNote = sNote & "Color: " & mvarColor & ";  "
        End If
        
        If mvarTexture <> "" Then
            sNote = sNote & "Texture: " & mvarTexture & "; "
        End If
        
        If mvarConsistency <> "" Then
            sNote = sNote & "Consistency: " & mvarConsistency & "; "
        End If
        
        If mvarContour <> "" Then
            sNote = sNote & "Contour: " & mvarContour & "; "
        End If
        
        If mvarMargins <> "" Then
            sNote = sNote & "Margins: " & mvarMargins & "; "
        End If
        
        If mvarPapillae <> "" Then
            sNote = sNote & "Papillae: " & mvarPapillae & "; "
        End If
        
        If mvarBleeding <> "" Then
            sNote = sNote & "Bleeding: " & mvarBleeding & "; "
        End If
        
        If mvarSuppuration <> "" Then
            sNote = sNote & "Suppuration: " & mvarSuppuration & "; "
        End If
        
        'Remove the last "; "
        sNote = Left(sNote, Len(sNote) - Len("; "))
        
        sNote = sNote & crtfEndControlName
        
    End If
    
    If mvarTreatPhase <> "" Then
        sNote = sNote & crtfStartControlName & "Treatment Phase: " & mvarTreatPhase & crtfEndControlName
    End If
    
    If mvarPerioNote <> "" Then
        sNote = sNote & crtfStartControlName & "Perio Note: " & mvarPerioNote & crtfEndControlName
    End If
    
    clsISection_GetRTFNote = sNote
    
End Function


'*******************************************************************
'Function: clsISection_LoadSection
'Author: Scott Foster
'Date: 10-22-01
'Returns: True/False
'Purpose: This loads the data from the db
'*******************************************************************
Private Function clsISection_LoadSection(ByVal PatientID As Long, ByVal DentalEncounterID As Long) As Boolean

Dim sSQL As String, rs As Recordset, rsPrev As Recordset
Dim lLastPerioExamDE As Long

Const sPROC_NAME = "LoadSection"

On Error GoTo ErrHnd

    mvarPatientID = PatientID
    mvarDentalEncounterID = DentalEncounterID
    
    'Load the values
    sSQL = "SELECT TEMPL_GINGIVA.COLOR, TEMPL_GINGIVA.TEXTURE," _
        & " TEMPL_GINGIVA.CONSISTENCY, TEMPL_GINGIVA.CONTOUR," _
        & " TEMPL_GINGIVA.MARGINS, TEMPL_GINGIVA.PAPILLAE, TEMPL_GINGIVA.BLEEDING," _
        & " TEMPL_GINGIVA.Suppuration, TreatPhase, PerioNote" _
        & " From TEMPL_GINGIVA, DENTAL_ENCOUNTER" _
        & " Where TEMPL_GINGIVA.DENTAL_ENCOUNTER_ID = DENTAL_ENCOUNTER.DENTAL_ENCOUNTER_ID" _
        & " AND (DENTAL_ENCOUNTER.PATIENT_UNIT_ID = " & mvarPatientID & ") AND" _
        & " (DENTAL_ENCOUNTER.DENTAL_ENCOUNTER_ID = " & mvarDentalEncounterID & ")"
    
    Set rs = goconn.OpenRecordset(sSQL)
    
    If rs.EOF Then
        bSavedToDB = False
        dfDirtyFlag = dfAdd
        mvarColor = ""
        mvarTexture = ""
        mvarConsistency = ""
        mvarContour = ""
        mvarMargins = ""
        mvarPapillae = ""
        mvarBleeding = ""
        mvarSuppuration = ""
        mvarTreatPhase = ""
        mvarPerioNote = ""
    Else
        bSavedToDB = True
        dfDirtyFlag = dfNoAction
        mvarColor = rs("Color") & ""
        mvarTexture = rs("Texture") & ""
        mvarConsistency = rs("Consistency") & ""
        mvarContour = rs("Contour") & ""
        mvarMargins = rs("Margins") & ""
        mvarPapillae = rs("Papillae") & ""
        mvarBleeding = rs("Bleeding") & ""
        mvarSuppuration = rs("Suppuration") & ""
        mvarTreatPhase = rs("TreatPhase") & ""
        mvarPerioNote = rs("PerioNote") & ""
    End If
    
    rs.Close
    
    If mvarTreatPhase = "" Then
        'Find out if I need to default the Treatment phase to anything based on the
        'patients last perio exam
        sSQL = "SELECT PERIO_EXAMS.DENTAL_ENCOUNTER_ID, PERIO_EXAMS.EVALUATION_TYPE_ID" _
            & " From PERIO_EXAMS, DENTAL_ENCOUNTER" _
            & " WHERE PERIO_EXAMS.DENTAL_ENCOUNTER_ID = DENTAL_ENCOUNTER.DENTAL_ENCOUNTER_ID AND ((PERIO_EXAMS.DENTAL_ENCOUNTER_ID < " & mvarDentalEncounterID & ")" _
            & " AND (DENTAL_ENCOUNTER.PATIENT_UNIT_ID = " & mvarPatientID & ") AND (PERIO_EXAMS.EVALUATION_TYPE_ID = " & eEvalType.perReEval & ") OR" _
            & " (PERIO_EXAMS.DENTAL_ENCOUNTER_ID < " & mvarDentalEncounterID & ") AND (DENTAL_ENCOUNTER.PATIENT_UNIT_ID = " & mvarPatientID & ") AND" _
            & " (PERIO_EXAMS.EVALUATION_TYPE_ID = " & eEvalType.perInitial & "))" _
            & " ORDER BY PERIO_EXAMS.DENTAL_ENCOUNTER_ID DESC"

        Set rs = goconn.OpenRecordset(sSQL)
        
        If Not rs.EOF Then
            lLastPerioExamDE = rs("DENTAL_ENCOUNTER_ID")
            
            'Filter on the lLastPerioExamDE for initial, because it is the last perio exam
            rs.Filter = "DENTAL_ENCOUNTER_ID = " & lLastPerioExamDE & " and EVALUATION_TYPE_ID = " & eEvalType.perInitial
            If Not rs.EOF Then
                'The last perio exam had an Initial exam
                'so default the treatment phase to Initial Therapy.
                mvarTreatPhase = "Initial Therapy"
            End If
            
            'Filter on the lLastPerioExamDE for re-eval, because it is the last perio exam
            rs.Filter = "DENTAL_ENCOUNTER_ID = " & lLastPerioExamDE & " and EVALUATION_TYPE_ID = " & eEvalType.perReEval
            If Not rs.EOF Then
                'The last perio exam had a Re-Evaluation exam
                'so default the treatment phase to Surgical Therapy.
                mvarTreatPhase = "Surgical Therapy"
            End If
        End If
    End If
    
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

'Save the data to the TemplGingiva table

Dim sSQL As String

Const sPROC_NAME As String = "SaveSection"

On Error GoTo ErrHnd

    If dfDirtyFlag = dfNoAction Then
        clsISection_SaveSection = True
        Exit Function
        
    ElseIf dfDirtyFlag = dfAdd Then
        'The record does not exist, so create an insert statement
        sSQL = "INSERT INTO TEMPL_GINGIVA (" _
            & " DENTAL_ENCOUNTER_ID, COLOR," _
            & " TEXTURE, CONSISTENCY," _
            & " CONTOUR, MARGINS," _
            & " PAPILLAE, BLEEDING," _
            & " SUPPURATION, TREATPHASE, PERIONOTE)" _
            & " VALUES" _
            & " (" & mvarDentalEncounterID & ", '" & mvarColor & "'," _
            & " '" & mvarTexture & "', '" & mvarConsistency & "'," _
            & " '" & mvarContour & "', '" & mvarMargins & "'," _
            & " '" & mvarPapillae & "', '" & mvarBleeding & "'," _
            & " '" & mvarSuppuration & "', '" & mvarTreatPhase & "', '" & RQ(mvarPerioNote) & "')"
    ElseIf dfDirtyFlag = dfUpdate Then
        'The record exists, so create an update statement
        sSQL = "UPDATE TEMPL_GINGIVA" _
            & " SET COLOR = '" & mvarColor & "', TEXTURE = '" & mvarTexture & "'," _
            & " CONSISTENCY = '" & mvarConsistency & "', CONTOUR = '" & mvarContour & "'," _
            & " MARGINS = '" & mvarMargins & "', PAPILLAE = '" & mvarPapillae & "'," _
            & " BLEEDING = '" & mvarBleeding & "', SUPPURATION = '" & mvarSuppuration & "'," _
            & " TREATPHASE = '" & mvarTreatPhase & "', PERIONOTE = '" & RQ(mvarPerioNote) & "'" _
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


