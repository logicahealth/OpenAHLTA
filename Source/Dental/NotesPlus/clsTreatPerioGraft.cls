VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTreatPGraft"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarGraftMaterial As String
Private mvarMaterialComments As String
Private mvarMaterialManuf As String
Private mvarMaterialLot As String
Private mvarMembrane As String
Private mvarMembraneManuf As String
Private mvarMembraneLot As String
Private mvarMembraneComments As String
Private mvarSutureType As String
Private mvarSutureNumber As String
Private mvarSutureTech As String
Private mvarRootCond As String
Private mvarRootCondComments As String
Private mvarPerioDressing As String
Private mvarPerioDressingComments As String
Private mvarDentalEncounterID As Long 'local copy
Private mvarDirtyFlag As DirtyFlagEnum 'local copy
Private mvarGraftID As Long 'local copy

'Options for GraftMaterials
Private Const cGraftMaterialValues As String = _
    "Autograft (intraoral)^Autograft (intraoral)^false|" _
    & "Autograft (extraoral)^Autograft (extraoral)^false|" _
    & "Allograft (FDBA)^Allograft (FDBA)^false|" _
    & "Allograft (DFDBA)^Allograft (DFDBA)^false|" _
    & "Alloplast (HA)^Alloplast (HA)^false|" _
    & "Alloplast (Calcium sulfate)^Alloplast (Calcium sulfate)^false|" _
    & "Alloplast (Bioactive granules)^Alloplast (Bioactive granules)^false|" _
    & "Alloplast (Other)^Alloplast (Other)^false|" _
    & "Xerograft (Bovine)^Xerograft (Bovine)^false|" _
    & "Isograft^Isograft^false|" _
    & "Biologics (Enamel matrix derived)^Biologics (Enamel matrix derived)^false|" _
    & "Biologics (Platelet rich plasma)^Biologics (Platelet rich plasma)^false|" _
    & "Biologics (Growth factors)^Biologics (Growth factors)^false|" _
'    & " ^ ^false"

'Options for Suture types
Private Const cGraftSutureValues As String = _
    "Black silk (3-0)^Black silk (3-0)^false|" _
    & "Black silk (4-0)^Black silk (4-0)^false|" _
    & "Black silk (5-0)^Black silk (5-0)^false|" _
    & "Black silk (6-0)^Black silk (6-0)^false|" _
    & "Chromic (3-0)^Chromic (3-0)^false|" _
    & "Chromic (4-0)^Chromic (4-0)^false|" _
    & "Chromic (5-0)^Chromic (5-0)^false|" _
    & "Chromic (6-0)^Chromic (6-0)^false|" _
    & "Gut (3-0)^Gut (3-0)^false|" _
    & "Gut (4-0)^Gut (4-0)^false|" _
    & "Gut (5-0)^Gut (5-0)^false|" _
    & "Gut (6-0)^Gut (6-0)^false|" _
    & "Gortex^Gortex^false|" _
    & "Vicryl^Vicryl^false|" _
    & "Nylon (4-0)^Nylon (4-0)^false|" _
    & "Nylon (5-0)^Nylon (5-0)^false|" _
    & "Nylon (6-0)^Nylon (6-0)^false|"
    
'    & " ^ ^false"

'Constants for field lengths
Private Const cGraftMaterialLen As Integer = 1000
Private Const cMaterialCommentsLen As Integer = 2000
Private Const cMaterialLotLen As Integer = 50
Private Const cMaterialManufLen As Integer = 1000
Private Const cMembraneLen As Integer = 1000
Private Const cMembraneCommentsLen As Integer = 2000
Private Const cMembraneLotLen As Integer = 50
Private Const cMembraneManufLen As Integer = 1000
Private Const cPerioDressingLen As Integer = 50
Private Const cPerioDressingCommentsLen As Integer = 2000
Private Const cRootCondLen As Integer = 50
Private Const cRootCondCommentsLen As Integer = 2000
Private Const cSutureNumberLen As Integer = 50
Private Const cSutureTechLen As Integer = 2000
Private Const cSutureTypeLen As Integer = 1000

'*******************************************************************
'Function: GetNote
'Author: Scott Foster
'Date: 10-30-01
'Returns: The note
'Purpose: To get a text string that represents the note
'*******************************************************************
Public Function GetNote() As String

Dim sNote As String
Dim sTemp As String

    'Graft material
    sTemp = GraftMaterialDisplay
    If sTemp <> "" Then
        sNote = sNote & "Graft material: " & sTemp & vbCrLf
        sNote = sNote & vbTab & "Manufacturer: " & mvarMaterialManuf & vbCrLf
        sNote = sNote & vbTab & "Lot: " & mvarMaterialLot & vbCrLf
        sNote = sNote & vbTab & "Comments: " & mvarMaterialComments & vbCrLf
    End If
    
    'Membrane
    If mvarMembrane <> "" Then
        sNote = sNote & "Membrane: " & mvarMembrane & vbCrLf
        sNote = sNote & vbTab & "Manufacturer: " & mvarMembraneManuf & vbCrLf
        sNote = sNote & vbTab & "Lot: " & mvarMembraneLot & vbCrLf
        sNote = sNote & vbTab & "Comments: " & mvarMembraneComments & vbCrLf
    End If
    
    'Sutures
    sTemp = SutureTypeDisplay
    If sTemp <> "" Then
        sNote = sNote & "Suture type: " & sTemp & vbCrLf
        sNote = sNote & vbTab & "Number of sutures: " & mvarSutureNumber & vbCrLf
        sNote = sNote & vbTab & "Technique: " & mvarSutureTech & vbCrLf
    End If
    
    'Root conditioning
    If mvarRootCond <> "" Then
        sNote = sNote & "Root conditioning: " & mvarRootCond & vbCrLf
        sNote = sNote & vbTab & "Comments: " & mvarRootCondComments & vbCrLf
    End If
    
    'Perio Dressing
    If mvarPerioDressing <> "" Then
        sNote = sNote & "Perio dressing: " & mvarPerioDressing & vbCrLf
        sNote = sNote & vbTab & "Comments: " & mvarPerioDressingComments & vbCrLf
    End If
    
    GetNote = sNote
    
End Function

Public Function GetRTFNote() As String

Dim sNote As String
Dim sTemp As String

    'Graft material
    sTemp = GraftMaterialDisplay
    If sTemp <> "" Then
        sNote = sNote & crtfStartControlName & "Graft material: " & sTemp & "; "
        If mvarMaterialManuf <> "" Then
            sNote = sNote & "Manufacturer: " & mvarMaterialManuf & "; "
        End If
        If mvarMaterialLot <> "" Then
            sNote = sNote & "Lot: " & mvarMaterialLot & "; "
        End If
        If mvarMaterialComments <> "" Then
            sNote = sNote & "Comments: " & mvarMaterialComments & "; "
        End If
        'Remove the last "; "
        sNote = Left(sNote, Len(sNote) - Len("; "))
        sNote = sNote & crtfEndControlName
    End If
    
    'Membrane
    If mvarMembrane <> "" Then
        sNote = sNote & crtfStartControlName & "Membrane: " & mvarMembrane & "; "
        If mvarMembraneManuf <> "" Then
            sNote = sNote & "Manufacturer: " & mvarMembraneManuf & "; "
        End If
        If mvarMembraneLot <> "" Then
            sNote = sNote & "Lot: " & mvarMembraneLot & "; "
        End If
        If mvarMembraneComments <> "" Then
            sNote = sNote & "Comments: " & mvarMembraneComments & "; "
        End If
        'Remove the last "; "
        sNote = Left(sNote, Len(sNote) - Len("; "))
        sNote = sNote & crtfEndControlName
    End If
    
    'Sutures
    sTemp = SutureTypeDisplay
    If sTemp <> "" Then
        sNote = sNote & crtfStartControlName & "Suture type: " & sTemp & "; "
        If mvarSutureNumber <> "" Then
            sNote = sNote & "Number of sutures: " & mvarSutureNumber & "; "
        End If
        If mvarSutureTech <> "" Then
            sNote = sNote & "Technique: " & mvarSutureTech & "; "
        End If
        'Remove the last "; "
        sNote = Left(sNote, Len(sNote) - Len("; "))
        sNote = sNote & crtfEndControlName
    End If
    
    'Root conditioning
    If mvarRootCond <> "" Then
        sNote = sNote & crtfStartControlName & "Root conditioning: " & mvarRootCond & "; "
        If mvarRootCondComments <> "" Then
            sNote = sNote & "Comments: " & mvarRootCondComments & "; "
        End If
        'Remove the last "; "
        sNote = Left(sNote, Len(sNote) - Len("; "))
        sNote = sNote & crtfEndControlName
    End If
    
    'Perio Dressing
    If mvarPerioDressing <> "" Then
        sNote = sNote & crtfStartControlName & "Perio dressing: " & mvarPerioDressing & "; "
        If mvarPerioDressingComments <> "" Then
            sNote = sNote & "Comments: " & mvarPerioDressingComments & "; "
        End If
        'Remove the last "; "
        sNote = Left(sNote, Len(sNote) - Len("; "))
        sNote = sNote & crtfEndControlName
    End If
    
    GetRTFNote = sNote
    
End Function

'*******************************************************************
'Function: GraftMaterialDisplay
'Author: Scott Foster
'Date: 10-26-01
'Returns: A comma delimited list of the selected graft materials
'Purpose: To return a comma delimited list of the selected graft materials
'*******************************************************************
Public Function GraftMaterialDisplay() As String
    
Dim oItems As New CMultiListItems
Dim oItem As CMultiListItem

    oItems.ListType = Standard
    oItems.LoadFromValueString mvarGraftMaterial
    GraftMaterialDisplay = oItems.Value_String_Display
    
End Function


'*******************************************************************
'Function: SutureTypeDisplay
'Author: Scott Foster
'Date: 10-29-01
'Returns: A comma delimited list of the selected graft materials
'Purpose: To return a comma delimited list of the selected graft materials
'*******************************************************************
Public Function SutureTypeDisplay() As String
    
Dim oItems As New CMultiListItems
Dim oItem As CMultiListItem

    oItems.ListType = Standard
    oItems.LoadFromValueString mvarSutureType
    SutureTypeDisplay = oItems.Value_String_Display
    
End Function

Friend Property Let GraftID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.GraftID = 5
    mvarGraftID = vData
End Property


Public Property Get GraftID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.GraftID
    GraftID = mvarGraftID
End Property



Friend Property Let DirtyFlag(ByVal vData As DirtyFlagEnum)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.DirtyFlag = Form1
    mvarDirtyFlag = vData
End Property


Public Property Get DirtyFlag() As DirtyFlagEnum
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DirtyFlag
    DirtyFlag = mvarDirtyFlag
End Property



'*******************************************************************
'Function: clsISection_SaveSection
'Author: Scott Foster
'Date: 10-25-01
'Returns: True/False
'Purpose: Saves the data to the table
'*******************************************************************
Public Function SaveSection() As Boolean

Dim sSQL As String
Dim rs As Recordset

Const sPROC_NAME = "SaveSection"

    Select Case mvarDirtyFlag
    
        Case dfNoAction
            'Do nothing
            
        Case dfAdd
            'Add the data to the db
            
            'Get the sequence
            mvarGraftID = goconn.GetNextID(esGRAFT_ID_SEQ)
'            sSQL = "Select Graft_ID_SEQ.NextVal as ID from Dual"
'            Set rs = goconn.OpenRecordset(sSQL)
'            If Not rs.EOF Then
'                mvarGraftID = rs("ID")
'            Else
'                'Could not find the sequence
'                goCWShared.ShowVBError Err.Number, "Could not find sequence: " & sSQL, Err.Source & "." & sPROC_NAME, App.Title, vbCritical
'                SaveSection = False
'            End If
'            rs.Close
            
            'Insert the data
            sSQL = "INSERT INTO TEMPL_TREAT_PERIO_GRAFT" _
                & " (GRAFT_ID, GRAFT_MATERIAL, MATERIAL_MANUF, MATERIAL_LOT, MATERIAL_COMMENTS," _
                & " MEMBRANE, MEMBRANE_MANUF, MEMBRANE_LOT, MEMBRANE_COMMENTS," _
                & " SUTURE_TYPE, SUTURE_NUMBER, SUTURE_TECH," _
                & " ROOT_COND, ROOT_COND_COMMENTS," _
                & " PERIO_DRESSING, PERIO_DRESSING_COMMENTS, DENTAL_ENCOUNTER_ID)" _
                & " VALUES (" & mvarGraftID & ", '" & RQ(mvarGraftMaterial) & "', '" & RQ(mvarMaterialManuf) & "', '" & RQ(mvarMaterialLot) & "', '" & RQ(mvarMaterialComments) & "'," _
                & " '" & RQ(mvarMembrane) & "', '" & RQ(mvarMembraneManuf) & "', '" & RQ(mvarMembraneLot) & "', '" & RQ(mvarMembraneComments) & "'," _
                & " '" & RQ(SutureType) & "', '" & RQ(mvarSutureNumber) & "', '" & RQ(mvarSutureTech) & "'," _
                & " '" & RQ(mvarRootCond) & "', '" & RQ(mvarRootCondComments) & "'," _
                & " '" & RQ(mvarPerioDressing) & "', '" & RQ(mvarPerioDressing) & "', " & mvarDentalEncounterID & ")"
            
            goconn.ExecuteSQL (sSQL)
            
        Case dfUpdate
            'Update data in the db
            sSQL = "UPDATE TEMPL_TREAT_PERIO_GRAFT" _
                & " SET GRAFT_MATERIAL = '" & RQ(mvarGraftMaterial) & "', MATERIAL_MANUF = '" & RQ(mvarMaterialManuf) & "'," _
                & " MATERIAL_LOT = '" & RQ(mvarMaterialLot) & "', MATERIAL_COMMENTS = '" & RQ(mvarMaterialComments) & "'," _
                & " MEMBRANE = '" & RQ(mvarMembrane) & "', MEMBRANE_MANUF = '" & RQ(mvarMembraneManuf) & "'," _
                & " MEMBRANE_LOT = '" & RQ(mvarMembraneLot) & "', MEMBRANE_COMMENTS = '" & RQ(mvarMembraneComments) & "'," _
                & " SUTURE_TYPE = '" & RQ(mvarSutureType) & "', SUTURE_NUMBER = '" & RQ(mvarSutureNumber) & "'," _
                & " SUTURE_TECH = '" & RQ(SutureTech) & "', ROOT_COND = '" & RQ(mvarRootCond) & "'," _
                & " ROOT_COND_COMMENTS = '" & RQ(RootCondComments) & "', PERIO_DRESSING = '" & RQ(mvarPerioDressing) & "'," _
                & " PERIO_DRESSING_COMMENTS = '" & RQ(mvarPerioDressingComments) & "'" _
                & " Where (GRAFT_ID = " & mvarGraftID & ")"
                
            goconn.ExecuteSQL (sSQL)

        Case dfdelete
            'Delete the record in the db
            If mvarGraftID <> 0 Then
                sSQL = "DELETE FROM TEMPL_TREAT_PERIO_GRAFT" _
                    & " Where (GRAFT_ID = " & mvarGraftID & ")"
                goconn.ExecuteSQL (sSQL)
            End If
            
    End Select
    
    mvarDirtyFlag = dfNoAction
    SaveSection = True
    
End Function

Private Function UpdateDirtyFlag()

    If mvarDirtyFlag = dfNoAction Then
        mvarDirtyFlag = dfUpdate
    End If
    
End Function


Friend Property Let DentalEncounterID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DentalEncounterID = 5
    mvarDentalEncounterID = vData
End Property


Public Property Get DentalEncounterID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DentalEncounterID
    DentalEncounterID = mvarDentalEncounterID
End Property




Public Property Let PerioDressingComments(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PerioDressingComments = 5
    mvarPerioDressingComments = Left(vData, cPerioDressingCommentsLen)
    UpdateDirtyFlag
End Property


Public Property Get PerioDressingComments() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PerioDressingComments
    PerioDressingComments = mvarPerioDressingComments
End Property



Public Property Let PerioDressing(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PerioDressing = 5
    mvarPerioDressing = Left(vData, cPerioDressingLen)
    UpdateDirtyFlag
End Property


Public Property Get PerioDressing() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PerioDressing
    PerioDressing = mvarPerioDressing
End Property



Public Property Let RootCondComments(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RootCondComments = 5
    mvarRootCondComments = Left(vData, cRootCondCommentsLen)
    UpdateDirtyFlag
End Property


Public Property Get RootCondComments() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RootCondComments
    RootCondComments = mvarRootCondComments
End Property



Public Property Let RootCond(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RootCond = 5
    mvarRootCond = Left(vData, cRootCondLen)
    UpdateDirtyFlag
End Property


Public Property Get RootCond() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RootCond
    RootCond = mvarRootCond
End Property



Public Property Let SutureTech(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SutureTech = 5
    mvarSutureTech = Left(vData, cSutureTechLen)
    UpdateDirtyFlag
End Property


Public Property Get SutureTech() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SutureTech
    SutureTech = mvarSutureTech
End Property



Public Property Let SutureNumber(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SutureNumber = 5
    mvarSutureNumber = Left(vData, cSutureNumberLen)
    UpdateDirtyFlag
End Property


Public Property Get SutureNumber() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SutureNumber
    SutureNumber = mvarSutureNumber
End Property



Public Property Let SutureType(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SutureType = 5
    mvarSutureType = Left(vData, cSutureTypeLen)
    UpdateDirtyFlag
End Property


Public Property Get SutureType() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SutureType
    If mvarSutureType = "" Then
        mvarSutureType = cGraftSutureValues
    End If
        
    SutureType = mvarSutureType
    
End Property



Public Property Let MembraneComments(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MembraneComments = 5
    mvarMembraneComments = Left(vData, cMembraneCommentsLen)
    UpdateDirtyFlag
End Property


Public Property Get MembraneComments() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MembraneComments
    MembraneComments = mvarMembraneComments
End Property



Public Property Let MembraneLot(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MembraneLot = 5
    mvarMembraneLot = Left(vData, cMembraneLotLen)
    UpdateDirtyFlag
End Property


Public Property Get MembraneLot() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MembraneLot
    MembraneLot = mvarMembraneLot
End Property



Public Property Let MembraneManuf(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MembraneManuf = 5
    mvarMembraneManuf = Left(vData, cMembraneManufLen)
    UpdateDirtyFlag
End Property


Public Property Get MembraneManuf() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MembraneManuf
    MembraneManuf = mvarMembraneManuf
End Property



Public Property Let Membrane(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Membrane = 5
    mvarMembrane = Left(vData, cMembraneLen)
    UpdateDirtyFlag
End Property


Public Property Get Membrane() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Membrane
    Membrane = mvarMembrane
End Property



Public Property Let MaterialLot(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MaterialLot = 5
    mvarMaterialLot = Left(vData, cMaterialLotLen)
    UpdateDirtyFlag
End Property


Public Property Get MaterialLot() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MaterialLot
    MaterialLot = mvarMaterialLot
End Property



Public Property Let MaterialManuf(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MaterialManuf = 5
    mvarMaterialManuf = Left(vData, cMaterialManufLen)
    UpdateDirtyFlag
End Property


Public Property Get MaterialManuf() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MaterialManuf
    MaterialManuf = mvarMaterialManuf
End Property



Public Property Let MaterialComments(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MaterialComments = 5
    mvarMaterialComments = Left(vData, 2000)
    UpdateDirtyFlag
End Property


Public Property Get MaterialComments() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MaterialComments
    MaterialComments = mvarMaterialComments
End Property



Public Property Let GraftMaterial(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.GraftMaterial = 5
    mvarGraftMaterial = Left(vData, cGraftMaterialLen)
    UpdateDirtyFlag
End Property


Public Property Get GraftMaterial() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.GraftMaterial

    If mvarGraftMaterial = "" Then
        mvarGraftMaterial = cGraftMaterialValues
    End If
    
    GraftMaterial = mvarGraftMaterial
    
End Property



Private Sub Class_Initialize()

'Set the default
    mvarGraftMaterial = cGraftMaterialValues
    mvarSutureType = cGraftSutureValues
    mvarDirtyFlag = dfNoAction

End Sub


