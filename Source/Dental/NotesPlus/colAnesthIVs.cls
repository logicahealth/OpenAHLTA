VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colAnesthIVs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsIV"
Attribute VB_Ext_KEY = "Member0" ,"clsIV"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mCol As Collection

Private mvarPatientID As Long
Private mvarDentalEncounterID As Long



Public Property Get DentalEncounterID() As Long
    DentalEncounterID = mvarDentalEncounterID
End Property
'*******************************************************************
'Function: LoadSection
'Author: Scott Foster
'Date: 1-05-01
'Returns: True/False
'Purpose: This loads the data from the db
'*******************************************************************
Public Function LoadSection(ByVal PatientID As Long, ByVal DentalEncounterID As Long) As Boolean

Dim sSQL As String, rs As Recordset, rsPrev As Recordset
Dim objNewMember As clsAnesthIV

Const sPROC_NAME = "LoadSection"

On Error GoTo ErrHnd

    mvarPatientID = PatientID
    mvarDentalEncounterID = DentalEncounterID
    
    'Load the values
    sSQL = "Select Anesth_ID, QTY_PROCEDURES, TECHNIQUE, GIVEN_BY, COMMENTS, DELETED, WORKLOAD_ID FROM TEMPL_ANESTH WHERE" _
        & " IS_LOCAL_ANESTH = 0 and DENTAL_ENCOUNTER_ID = " & mvarDentalEncounterID & " AND DELETED = 0 ORDER BY ANESTH_ID"
    
    Set rs = goconn.OpenRecordset(sSQL)
    
    Do Until rs.EOF
        'Set the values from the db including "" to allow for nulls
        Set objNewMember = Add(mvarDentalEncounterID)
        objNewMember.AnesthID = rs("Anesth_ID") & ""
        objNewMember.QtyOfProcedures = rs("QTY_PROCEDURES") & ""
        objNewMember.TechniqueType = Val(rs("TECHNIQUE") & "")
        objNewMember.GivenBy = rs("GIVEN_BY") & ""
        objNewMember.Comments = rs("COMMENTS") & ""
        objNewMember.Deleted = CBool(Val(rs("DELETED") & ""))
        objNewMember.WorkloadID = Val(rs("WORKLOAD_ID"))
        objNewMember.DirtyFlag = dfNoAction
        rs.MoveNext
    Loop
    
    rs.Close
    
    Set rs = Nothing
    
    LoadSection = True
    
Exit Function

ErrHnd:
    goCWShared.ShowVBError Err.Number, Err.Description, Err.Source & "." & sPROC_NAME, App.Title, vbCritical
End Function


'*******************************************************************
'Function: SaveSection
'Author: Scott Foster
'Date: 1-4-01
'Returns: True/False
'Purpose: Saves the IV objects
'*******************************************************************
Public Function SaveSection() As Boolean

Dim oIV As clsAnesthIV
    
    SaveSection = True
    
    For Each oIV In mCol
        If oIV.SaveSection = False Then
            SaveSection = False
        End If
    Next
    
End Function


Public Function Add(DentalEncounter As Long) As clsAnesthIV
    
    'create a new object
    Dim objNewMember As clsAnesthIV
    Set objNewMember = New clsAnesthIV

    'set the properties passed into the method
    objNewMember.DentalEncounterID = DentalEncounter
    objNewMember.DirtyFlag = dfAdd
    
    'Add the member to the collection
    mCol.Add objNewMember

    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As clsAnesthIV
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)
    
    If mCol(vntIndexKey).AnesthID <> 0 Then
        'Set the dirtyflag to deleted
        mCol(vntIndexKey).DirtyFlag = dfDelete
    Else
        'The object has not been saved, so remove it.
        mCol.Remove vntIndexKey
    End If
    
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub

'*******************************************************************
'Function: GetNote
'Author: Scott Foster
'Date: 10-22-01
'Returns: Text note
'Purpose: Gets the note for the narrative summary
'*******************************************************************
Public Function GetNote() As String

Dim oIV As clsAnesthIV

    'Get the note from each oIV object
    For Each oIV In mCol
        GetNote = GetNote & oIV.GetNote
    Next

End Function
Public Function GetRTFNote() As String
    
    GetRTFNote = GetNote
    
End Function

