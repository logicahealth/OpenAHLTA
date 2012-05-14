VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAnesthIV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarQtyOfProcedures As String 'local copy
Private mvarGivenBy As String 'local copy
Private mvarGivenByID As Long 'local copy
Private mvarComments As String 'local copy
Private mvarDirtyFlag As DirtyFlagEnum
Private mvarDentalEncounterID As Long 'local copy
Private mvarAnesthID As Long 'local copy
Private mvarDeleted As Boolean 'local copy
Private mvarWorkloadID As Long 'local copy
Private mvarTechniqueType As Integer 'local copy

Private Type TechniqueType
    TechniqueText As String
    WorkloadCode As String
End Type
Private arTechnique(6) As TechniqueType

Public Property Get TechniqueText() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Technique

    'Return the TechniqueText based on the technique type
    TechniqueText = arTechnique(mvarTechniqueType).TechniqueText
    
End Property

Public Function GetRTFNote() As String

Dim sTemp As String
    
    'Test one of the required fields
    If mvarDeleted = False Then
        sTemp = crtfStartControlName & "Qty of Procedures " & mvarQtyOfProcedures & " " & Me.TechniqueText & " given by " & mvarGivenBy & ", Comments: " & mvarComments & crtfEndSection
    End If
    
    GetRTFNote = sTemp

End Function
'*******************************************************************
'Function: GetNote
'Author: Scott Foster
'Date: 11/15/01
'Returns: A string that is displayed in the narrative summary for this encounter
'Purpose: To compare the values from this encounter and the previous encounter and
'return a text note.
'*******************************************************************
Public Function GetNote() As String

Dim sTemp As String
    
    'Test one of the required fields
    If mvarDeleted = False Then
        sTemp = vbTab & "Qty of Procedures " & mvarQtyOfProcedures & " " & Me.TechniqueText & " given by " & mvarGivenBy & ", Comments: " & mvarComments & vbCrLf
    End If
    
    GetNote = sTemp

End Function

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




Friend Property Let AnesthID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AnesthID = 5
    mvarAnesthID = vData
End Property


Public Property Get AnesthID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AnesthID
    AnesthID = mvarAnesthID
End Property



Public Property Let DentalEncounterID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DentalEncounterID = 5
    mvarDentalEncounterID = vData
End Property


Public Property Get DentalEncounterID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DentalEncounterID
    DentalEncounterID = mvarDentalEncounterID
End Property




'*******************************************************************
'Function: SaveSection
'Author: Scott Foster
'Date: 11-4-01
'Returns: True/False
'Purpose: Saves the data to the table
'*******************************************************************
Public Function SaveSection() As Boolean

Dim sSQL As String
Dim rs As Recordset
Dim iMultiplier As Integer

Const sPROC_NAME As String = "SaveSection"

On Error GoTo ErrHnd

    If mvarDirtyFlag = dfNoAction Then
        SaveSection = True
        Exit Function
    Else
        'Update the workload
        If mvarTechniqueType = 2 Or mvarTechniqueType = 4 Then
            iMultiplier = mvarQtyOfProcedures
        Else
            iMultiplier = 1
        End If
        
        If UpdateWorkload(mvarWorkloadID, mvarDirtyFlag, mvarDentalEncounterID, GetWorkloadCode, iMultiplier, mvarGivenByID) = False Then
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
            & ", QTY_PROCEDURES, TECHNIQUE" _
            & ", GIVEN_BY, COMMENTS, IS_LOCAL_ANESTH, DELETED, WORKLOAD_ID) Values (" _
            & "" & mvarAnesthID & ", " & mvarDentalEncounterID _
            & ", '" & mvarQtyOfProcedures & "', '" & mvarTechniqueType & "'" _
            & ", '" & mvarGivenBy & "', '" & mvarComments & "',0,0," & mvarWorkloadID & ")"
            
    ElseIf mvarDirtyFlag = dfUpdate Then
        'The record exists, so create an update statement
        sSQL = "Update TEMPL_ANESTH" _
            & " Set QTY_PROCEDURES='" & mvarQtyOfProcedures & "'," _
            & " TECHNIQUE='" & mvarTechniqueType & "'," _
            & " GIVEN_BY='" & mvarGivenBy & "'," _
            & " COMMENTS='" & mvarComments & "'," _
            & " DELETED=" & CInt(mvarDeleted) & "," _
            & " WORKLOAD_ID=" & mvarWorkloadID _
            & " Where Anesth_ID = " & mvarAnesthID
            
    End If
    
    goconn.ExecuteSQL sSQL
    
    mvarDirtyFlag = dfNoAction
    
    SaveSection = True

Exit Function

ErrHnd:
    goCWShared.ShowVBError Err.Number, Err.Description, Err.Source & "." & sPROC_NAME, App.Title, vbCritical
    Resume Next
End Function

Public Function GetWorkloadCode() As String

    GetWorkloadCode = arTechnique(mvarTechniqueType).WorkloadCode

End Function

Private Function UpdateDirtyFlag()

    If mvarDirtyFlag = dfNoAction Then
        mvarDirtyFlag = dfUpdate
    End If
    
End Function


Public Property Let DirtyFlag(ByVal vData As DirtyFlagEnum)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.DirtyFlag = Form1
    mvarDirtyFlag = vData
End Property


Public Property Get DirtyFlag() As DirtyFlagEnum
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DirtyFlag
    DirtyFlag = mvarDirtyFlag
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



Public Property Let QtyOfProcedures(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.QtyOfProcedures = 5
    mvarQtyOfProcedures = vData
    UpdateDirtyFlag
End Property


Public Property Get QtyOfProcedures() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.QtyOfProcedures
    QtyOfProcedures = mvarQtyOfProcedures
End Property



Private Sub Class_Initialize()

    'Init the technique values
    arTechnique(0).TechniqueText = ""
    arTechnique(0).WorkloadCode = ""
    arTechnique(1).TechniqueText = "Analgesia, Nitrous Oxide"
    arTechnique(1).WorkloadCode = "20" '"D9230"
    arTechnique(2).TechniqueText = "Deep sedation/general anesthesia - Each additional 15 Minutes"
    arTechnique(2).WorkloadCode = "156" '"D9221"
    arTechnique(3).TechniqueText = "Deep sedation/general anesthesia - First 30 Minutes"
    arTechnique(3).WorkloadCode = "157" '"D9220"
    arTechnique(4).TechniqueText = "IV conscious sedation/analgesia - Each additional 15 Minutes"
    arTechnique(4).WorkloadCode = "503" '
    arTechnique(5).TechniqueText = "IV conscious sedation/analgesia - First 30 Minutes"
    arTechnique(5).WorkloadCode = "511" '"D9241"
    arTechnique(6).TechniqueText = "Non-Intravenous conscious sedation"
    arTechnique(6).WorkloadCode = "514" '"D9248"

End Sub


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

