VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAnesthLocal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarCarpules As String 'local copy
Private mvarLocalAnesthetic As String 'local copy
Private mvarLocation As String 'local copy
Private mvarLocationDisplay As String
Private mvarSupplementalOption As String 'local copy
Private mvarGivenBy As String 'local copy
Private mvarGivenByID As Long 'local copy
Private mvarComments As String 'local copy
Private mvarAnesthID As Long 'local copy
Private mvarDentalEncounterID As Long 'local copy
Private mvarDirtyFlag As DirtyFlagEnum
Private mvarDeleted As Boolean 'local copy
Private mvarLocalAnestheticType As Integer 'local copy
Private mvarWorkloadID As Long 'local copy
Private mvarTechniqueType As Integer 'local copy

'Used to capture properties of the techniques
Private Type TechniqueType
    TechniqueText As String
    WorkloadCode As String
End Type
Private arTechnique(5) As TechniqueType

'Used to capture all properties of the anesthetic
Private Type AnesthType
    AnesthNoteText As String
    VasoNoteText As String
    AnesthMG As Integer
    AnesthML As Double
    VasoMGOrIU As Double
    VasoUnits As String
End Type
Private arAnesth(13) As AnesthType

Public Function GetWorkloadCode() As String

    GetWorkloadCode = arTechnique(mvarTechniqueType).WorkloadCode

End Function

Public Property Let TechniqueType(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TechniqueType = 5
    mvarTechniqueType = vData
    UpdateDirtyFlag
End Property


Public Property Get TechniqueType() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TechniqueType
    TechniqueType = mvarTechniqueType
End Property





Public Property Let WorkloadId(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WorkLoadID = 5
    mvarWorkloadID = vData
End Property


Public Property Get WorkloadId() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.WorkLoadID
    WorkloadId = mvarWorkloadID
End Property





'*******************************************************************
'Function: GetNote
'Author: Scott Foster
'Date: 10/19/01
'Returns: A string that is displayed in the narrative summary for this encounter
'Purpose: To return a text note.
'*******************************************************************
Public Function GetNote() As String

Dim sTemp As String
    
    'Test one of the required fields
    If mvarLocalAnesthetic <> "" And mvarDeleted = False Then
        sTemp = vbTab & Me.LocalAnestheticML & " ml " & arAnesth(mvarLocalAnestheticType).AnesthNoteText & " (" & Me.LocalAnestheticMG & " mg)"
        
            If arAnesth(mvarLocalAnestheticType).VasoNoteText <> "" Then
                sTemp = sTemp & ", with " & arAnesth(mvarLocalAnestheticType).VasoNoteText & " (" & Me.VasoconstrictorMG & " " & arAnesth(mvarLocalAnestheticType).VasoUnits & ")"
            End If
            
        'Technique
        If arTechnique(mvarTechniqueType).TechniqueText <> "" Then
            sTemp = sTemp & ", Technique: " & arTechnique(mvarTechniqueType).TechniqueText
        End If
        
        'Location
        If mvarLocationDisplay <> "" Then
            sTemp = sTemp & ", Location: " & mvarLocationDisplay
        End If
        
        'Supplemental option
        If mvarSupplementalOption <> "" Then
            sTemp = sTemp & ", Supplemental: " & mvarSupplementalOption
        End If
        
        'Given by
        If mvarGivenBy <> "" Then
            sTemp = sTemp & ", Given by: " & mvarGivenBy
        End If
        
        'Comments
        If mvarComments <> "" Then
            sTemp = sTemp & ", Comments: " & mvarComments
        End If
        
    End If
    
    GetNote = sTemp & vbCrLf
    
End Function
Public Function GetRTFNote() As String

Dim sTemp As String
    
    'Test one of the required fields
    If mvarLocalAnesthetic <> "" And mvarDeleted = False Then
        sTemp = crtfStartControlName & Me.LocalAnestheticML & " ml " & arAnesth(mvarLocalAnestheticType).AnesthNoteText & " (" & Me.LocalAnestheticMG & " mg)"
            If arAnesth(mvarLocalAnestheticType).VasoNoteText <> "" Then
                sTemp = sTemp & ", with " & arAnesth(mvarLocalAnestheticType).VasoNoteText & " (" & Me.VasoconstrictorMG & " " & arAnesth(mvarLocalAnestheticType).VasoUnits & ")"
            End If
            
        'Technique
        If arTechnique(mvarTechniqueType).TechniqueText <> "" Then
            sTemp = sTemp & ", Technique: " & arTechnique(mvarTechniqueType).TechniqueText
        End If
        
        'Location
        If mvarLocationDisplay <> "" Then
            sTemp = sTemp & ", Location: " & mvarLocationDisplay
        End If
        
        'Supplemental option
        If mvarSupplementalOption <> "" Then
            sTemp = sTemp & ", Supplemental: " & mvarSupplementalOption
        End If
        
        'Given by
        If mvarGivenBy <> "" Then
            sTemp = sTemp & ", Given by: " & mvarGivenBy
        End If
        
        'Comments
        If mvarComments <> "" Then
            sTemp = sTemp & ", Comments: " & mvarComments
        End If
        
        If sTemp <> "" Then
            sTemp = sTemp & crtfEndSection
        End If
        
    End If
    
    GetRTFNote = sTemp

End Function

Public Property Get LocalAnestheticType() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LocalAnestheticType
    LocalAnestheticType = mvarLocalAnestheticType
End Property


Public Property Let LocalAnestheticType(vData As Integer)
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LocalAnestheticType
    mvarLocalAnestheticType = vData
    UpdateDirtyFlag
End Property

Public Property Let Deleted(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Deleted = 5
    'Setting deleted to false does nothing
    mvarDeleted = vData
    If mvarDeleted = True Then
        mvarDirtyFlag = dfdelete
    End If
    
End Property


Public Property Get Deleted() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Deleted
    Deleted = mvarDeleted
End Property



Public Property Get VasoconstrictorMG() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.VasoconstrictorMG
    
    VasoconstrictorMG = Val(mvarCarpules) * arAnesth(mvarLocalAnestheticType).VasoMGOrIU * 1.8
    
End Property



Public Property Get LocalAnestheticML() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LocalAnestheticML
    LocalAnestheticML = Val(mvarCarpules) * arAnesth(mvarLocalAnestheticType).AnesthML
    
End Property




Public Property Get LocalAnestheticMG() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LocalAnestheticMG
    'Calculate the anesthetic in mg
    
    LocalAnestheticMG = Val(mvarCarpules) * arAnesth(mvarLocalAnestheticType).AnesthMG * 1.8
    
End Property



Private Function UpdateDirtyFlag()

    If mvarDirtyFlag = dfNoAction Then
        mvarDirtyFlag = dfUpdate
    End If
    
End Function

Public Property Get DirtyFlag() As DirtyFlagEnum
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DirtyFlag
    DirtyFlag = mvarDirtyFlag
End Property
Public Property Let DirtyFlag(ByVal vData As DirtyFlagEnum)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.DirtyFlag = Form1
    mvarDirtyFlag = vData
End Property



'*******************************************************************
'Function: SaveSection
'Author: Scott Foster
'Date: 11-7-01
'Returns: True/False
'Purpose: Saves the data to the table
'*******************************************************************
Public Function SaveSection() As Boolean

Dim sSQL As String
Dim rs As Recordset

Const sPROC_NAME As String = "SaveSection"

On Error GoTo ErrHnd

    If mvarDirtyFlag = dfNoAction Then
        SaveSection = True
        Exit Function
    Else
        'Update the workload
        If UpdateWorkload(mvarWorkloadID, mvarDirtyFlag, mvarDentalEncounterID, GetWorkloadCode, 1, mvarGivenByID) = False Then
            Exit Function
        End If
    End If
    
    If mvarDirtyFlag = dfdelete Then
        'Delete the record
        If mvarAnesthID <> 0 Then
            sSQL = "UPDATE TEMPL_ANESTH SET DELETED = 1 Where (ANESTH_ID = " & mvarAnesthID & ")"
        Else
            SaveSection = True
            Exit Function
        End If
        
    ElseIf mvarDirtyFlag = dfAdd Then
        'The record does not exist, so create an insert statement
        mvarAnesthID = goconn.GetNextID(esTEMPL_ANESTH_SEQ)
'        sSQL = "Select TEMPL_ANESTH_SEQ.NEXTVAL ID from Dual"
'        Set rs = goconn.OpenRecordset(sSQL)
'        If Not rs.EOF Then
'            mvarAnesthID = rs("ID")
'        Else
'            goCWShared.ShowVBError 0, "Could not get the NextVal from Anesth_ID_Seq sequence", "clsIV." & sPROC_NAME, App.Title, vbCritical
'            Exit Function
'        End If
'        rs.Close
        
        sSQL = "Insert into TEMPL_ANESTH (ANESTH_ID, DENTAL_ENCOUNTER_ID" _
            & ", CARPULES, LOCAL_ANESTH, LOCAL_ANESTH_TYPE, TECHNIQUE, LOCATION" _
            & ", SUPPLEMENTAL_OPTION, GIVEN_BY, COMMENTS, IS_LOCAL_ANESTH, DELETED, WORKLOAD_ID, LOCATION_DISPLAY) Values (" _
            & "" & mvarAnesthID & ", " & mvarDentalEncounterID _
            & ", '" & mvarCarpules & "', '" & mvarLocalAnesthetic & "', " & mvarLocalAnestheticType & ", '" & mvarTechniqueType & "', '" & mvarLocation & "'" _
            & ", '" & mvarSupplementalOption & "', '" & mvarGivenBy & "', '" & mvarComments & "',1,0," & mvarWorkloadID & ", '" & mvarLocationDisplay & "')"
            
    ElseIf mvarDirtyFlag = dfUpdate Then
        'The record exists, so create an update statement
        sSQL = "Update TEMPL_ANESTH" _
            & " Set CARPULES='" & mvarCarpules & "'," _
            & " LOCAL_ANESTH='" & mvarLocalAnesthetic & "'," _
            & " LOCAL_ANESTH_TYPE=" & mvarLocalAnestheticType & "," _
            & " TECHNIQUE='" & mvarTechniqueType & "'," _
            & " LOCATION='" & mvarLocation & "'," _
            & " SUPPLEMENTAL_OPTION='" & mvarSupplementalOption & "'," _
            & " GIVEN_BY='" & mvarGivenBy & "'," _
            & " COMMENTS='" & mvarComments & "'," _
            & " DELETED=" & Int(mvarDeleted) & "," _
            & " WORKLOAD_ID=" & mvarWorkloadID & "," _
            & " LOCATION_DISPLAY='" & mvarLocationDisplay & "'" _
            & " Where Anesth_ID = " & mvarAnesthID
            
    End If
    
    goconn.ExecuteSQL sSQL
        
    mvarDirtyFlag = dfNoAction
    
    SaveSection = True

Exit Function

ErrHnd:
    goCWShared.ShowVBError Err.Number, Err.Description, Err.Source & "." & sPROC_NAME, App.Title, vbCritical
    
    'Delete the workload credit if it didn't update correctly
    UpdateWorkload mvarWorkloadID, dfdelete, mvarDentalEncounterID, "", 1
    
End Function


Public Property Let DentalEncounterID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DentalEncounterID = 5
    mvarDentalEncounterID = vData
    UpdateDirtyFlag
End Property


Public Property Get DentalEncounterID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DentalEncounterID
    DentalEncounterID = mvarDentalEncounterID
End Property



Public Property Let AnesthID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AnesthID = 5
    mvarAnesthID = vData
    UpdateDirtyFlag
End Property


Public Property Get AnesthID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AnesthID
    AnesthID = mvarAnesthID
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



Public Property Let GivenBy(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.GivenBy = 5
    mvarGivenBy = vData
    UpdateDirtyFlag
End Property

Public Property Let GivenByID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.GivenByID = 5
    mvarGivenByID = vData
    UpdateDirtyFlag
End Property


Public Property Get GivenBy() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.GivenBy
    GivenBy = mvarGivenBy
End Property

Public Property Get GivenByID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.GivenByID
    GivenByID = mvarGivenByID
End Property


Public Property Let SupplementalOption(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SupplementalOption = 5
    mvarSupplementalOption = vData
    UpdateDirtyFlag
End Property


Public Property Get SupplementalOption() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SupplementalOption
    SupplementalOption = mvarSupplementalOption
End Property



Public Property Let Location(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Location = 5
    mvarLocation = vData
    UpdateDirtyFlag
End Property


Public Property Let LocationDisplay(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Location = 5
    mvarLocationDisplay = vData
    
    'Don't update the dirty flag because the user can't directly modify this, they
    ' must modify Location
    'UpdateDirtyFlag
    
End Property

Public Property Get Location() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Location
    Location = mvarLocation
End Property



Public Property Get LocationDisplay() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Location
    
    LocationDisplay = mvarLocationDisplay
    
End Property



Public Property Get TechniqueText() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Technique

    'Return the TechniqueText based on the technique type
    TechniqueText = arTechnique(mvarTechniqueType).TechniqueText
    
End Property



Public Property Let LocalAnesthetic(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LocalAnesthetic = 5
    mvarLocalAnesthetic = vData
    UpdateDirtyFlag
End Property


Public Property Get LocalAnesthetic() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LocalAnesthetic
    LocalAnesthetic = mvarLocalAnesthetic
End Property



Public Property Let Carpules(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Carpules = 5
    mvarCarpules = vData
    UpdateDirtyFlag
End Property


Public Property Get Carpules() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Carpules
    Carpules = mvarCarpules
End Property



Private Sub Class_Initialize()

    'Init the technique values
    'Note, this is synched with the values in frmAnesthDialog
    arTechnique(0).TechniqueText = ""
    arTechnique(0).WorkloadCode = ""
    arTechnique(1).TechniqueText = "Local Anesthesia"
    arTechnique(1).WorkloadCode = "222" '"D9215"
    arTechnique(2).TechniqueText = "Local Anesthesia-not in conjunction with procedure"
    arTechnique(2).WorkloadCode = "223" '"D9210"
    arTechnique(3).TechniqueText = "Regional block anesthesia"
    arTechnique(3).WorkloadCode = "364" '"D9211"
    arTechnique(4).TechniqueText = "Trigeminal division block anesthesia"
    arTechnique(4).WorkloadCode = "477" '"D9212"
    arTechnique(5).TechniqueText = "Local Anesthesia-No workload"
    arTechnique(5).WorkloadCode = "NULL"

    'Init the calulation values for arAnesth
    arAnesth(1).AnesthNoteText = "Lidocaine, 2%"
    arAnesth(1).VasoNoteText = "1:100,000 epinephrine"
    arAnesth(1).AnesthMG = 20
    arAnesth(1).AnesthML = 1.8
    arAnesth(1).VasoMGOrIU = 0.01
    arAnesth(1).VasoUnits = "mg"
    arAnesth(2).AnesthNoteText = "Mepivicaine HCL, 3%"
    arAnesth(2).VasoNoteText = ""
    arAnesth(2).AnesthMG = 30
    arAnesth(2).AnesthML = 1.8
    arAnesth(2).VasoMGOrIU = 0
    arAnesth(2).VasoUnits = "mg"
    arAnesth(3).AnesthNoteText = "Bupivicaine HCL, 0.5%"
    arAnesth(3).VasoNoteText = "1:200,000 epinephrine"
    arAnesth(3).AnesthMG = 5
    arAnesth(3).AnesthML = 1.8
    arAnesth(3).VasoMGOrIU = 0.005
    arAnesth(3).VasoUnits = "mg"
    arAnesth(4).AnesthNoteText = "Prilocaine, 4%"
    arAnesth(4).VasoNoteText = ""
    arAnesth(4).AnesthMG = 40
    arAnesth(4).AnesthML = 1.8
    arAnesth(4).VasoMGOrIU = 0
    arAnesth(4).VasoUnits = "mg"
    arAnesth(5).AnesthNoteText = "Lidocaine, 2%"
    arAnesth(5).VasoNoteText = "1:50,000 epinephrine"
    arAnesth(5).AnesthMG = 20
    arAnesth(5).AnesthML = 1.8
    arAnesth(5).VasoMGOrIU = 0.02
    arAnesth(5).VasoUnits = "mg"
    arAnesth(6).AnesthNoteText = "Prilocaine, 4%"
    arAnesth(6).VasoNoteText = "1:200,000 epinephrine"
    arAnesth(6).AnesthMG = 40
    arAnesth(6).AnesthML = 1.8
    arAnesth(6).VasoMGOrIU = 0.05
    arAnesth(6).VasoUnits = "mg"
    
    arAnesth(7).AnesthNoteText = "Articaine, 4%"
    arAnesth(7).VasoNoteText = "1:100,000 epinephrine"
    arAnesth(7).AnesthMG = 40
    arAnesth(7).AnesthML = 1.8
    arAnesth(7).VasoMGOrIU = 0.01
    arAnesth(7).VasoUnits = "mg"
    
    arAnesth(8).AnesthNoteText = "Articaine, 4%"
    arAnesth(8).VasoNoteText = "1:200,000 epinephrine"
    arAnesth(8).AnesthMG = 40
    arAnesth(8).AnesthML = 1.8
    arAnesth(8).VasoMGOrIU = 0.005
    arAnesth(8).VasoUnits = "mg"
    
    arAnesth(9).AnesthNoteText = "Etidocaine, 1.5%"
    arAnesth(9).VasoNoteText = "1:200,000 epinephrine"
    arAnesth(9).AnesthMG = 15
    arAnesth(9).AnesthML = 1.8
    arAnesth(9).VasoMGOrIU = 0.005
    arAnesth(9).VasoUnits = "mg"
    
    arAnesth(10).AnesthNoteText = "Mepivacaine, 2%"
    arAnesth(10).VasoNoteText = "1:20,000 Levonordefrin"
    arAnesth(10).AnesthMG = 20
    arAnesth(10).AnesthML = 1.8
    arAnesth(10).VasoMGOrIU = 0.05
    arAnesth(10).VasoUnits = "mg"
    
    arAnesth(11).AnesthNoteText = "Prilocaine, 3%"
    arAnesth(11).VasoNoteText = "Octapressin"
    arAnesth(11).AnesthMG = 30
    arAnesth(11).AnesthML = 1.8
    arAnesth(11).VasoMGOrIU = 0.03
    arAnesth(11).VasoUnits = "iu"
    
    arAnesth(12).AnesthNoteText = "Mepivacaine, 2%"
    arAnesth(12).VasoNoteText = "1:100,000 epinephrine"
    arAnesth(12).AnesthMG = 20
    arAnesth(12).AnesthML = 1.8
    arAnesth(12).VasoMGOrIU = 0.01
    arAnesth(12).VasoUnits = "mg"
    
    arAnesth(13).AnesthNoteText = "Other"
    arAnesth(13).VasoNoteText = ""
    arAnesth(13).AnesthMG = 0
    arAnesth(13).AnesthML = 0
    arAnesth(13).VasoMGOrIU = 0
    arAnesth(13).VasoUnits = "mg"

End Sub


