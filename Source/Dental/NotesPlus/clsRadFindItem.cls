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
Attribute VB_Name = "clsRadFindItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarDirtyFlag As DirtyFlagEnum

Private mvarDental_Encounter_ID As Long 'local copy
Private mvarFindingsID As Long 'local copy
Private mvarToothNumber As String 'local copy
Private mvarFindingsNormal As Boolean
Private mvarCariesNoted As Boolean 'local copy
Private mvarPeriapicalPathology As Boolean 'local copy
Private mvarBoneLoss As Boolean 'local copy
Private mvarImpactedTooth As Boolean 'local copy
Private mvarOtherFindings As Boolean 'local copy
Private mvarDeleted As Boolean

Public Property Let FindingsNormal(vData As Boolean)
    
    mvarFindingsNormal = vData
    
End Property

Public Property Get FindingsNormal() As Boolean
    
    FindingsNormal = mvarFindingsNormal
    
End Property

Public Property Let OtherFindings(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.OtherFindings = 5
    mvarOtherFindings = vData
    UpdateDirtyFlag
End Property


Public Property Get OtherFindings() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.OtherFindings
    OtherFindings = mvarOtherFindings
End Property

Public Property Let ImpactedTooth(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ImpactedTooth = 5
    mvarImpactedTooth = vData
    UpdateDirtyFlag
    
End Property

Public Property Get ImpactedTooth() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ImpactedTooth
    ImpactedTooth = mvarImpactedTooth
End Property

Public Property Let BoneLoss(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BoneLoss = 5
    mvarBoneLoss = vData
    UpdateDirtyFlag
    
End Property


Public Property Get BoneLoss() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BoneLoss
    BoneLoss = mvarBoneLoss
End Property

Public Property Let PeriapicalPathology(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PeriapicalPathology = 5
    mvarPeriapicalPathology = vData
    UpdateDirtyFlag
    
End Property

Public Property Get PeriapicalPathology() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PeriapicalPathology
    PeriapicalPathology = mvarPeriapicalPathology
End Property



Public Property Let CariesNoted(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CariesNoted = 5
    mvarCariesNoted = vData
    UpdateDirtyFlag
    
End Property

Public Property Get CariesNoted() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CariesNoted
    CariesNoted = mvarCariesNoted
End Property



Public Property Let ToothNumber(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Tooth_Number = 5
    mvarToothNumber = vData
    UpdateDirtyFlag
    
End Property


Public Property Get ToothNumber() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Tooth_Number
    ToothNumber = mvarToothNumber
End Property



Public Property Let FindingsID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Finding_ID = 5
    mvarFindingsID = vData
End Property


Public Property Get FindingsID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Finding_ID
    FindingsID = mvarFindingsID
End Property



Public Property Let Dental_Encounter_ID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Dental_Encounter_ID = 5
    mvarDental_Encounter_ID = vData
End Property


Public Property Get Dental_Encounter_ID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Dental_Encounter_ID
    Dental_Encounter_ID = mvarDental_Encounter_ID
End Property

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


Public Property Let Deleted(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CariesNoted = 5
    'Setting deleted to false does nothing
    mvarDeleted = vData
    If mvarDeleted = True Then
        mvarDirtyFlag = dfdelete
    End If
End Property


Public Property Get Deleted() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CariesNoted
    Deleted = mvarDeleted
End Property
'*******************************************************************
'Function: GetNote
'Author: ASamedov
'Date: 01/08/2002
'Purpose: Builds a narrative summary of the user's selections.
'*******************************************************************
Public Function GetNote() As String
    Dim strNote As String

    strNote = Empty
    
    If Not mvarDeleted Then
        
        If Trim(mvarToothNumber) <> "" Then
            strNote = strNote & vbTab & "Tooth Number = " & mvarToothNumber & ". "
        Else
            Exit Function
        End If
    
        If mvarFindingsNormal = True Then
            strNote = strNote & "Findings normal. "
        End If
    
        If mvarCariesNoted = True Then
            strNote = strNote & "Caries noted. "
        End If
    
        If mvarPeriapicalPathology = True Then
            strNote = strNote & "Periapical pathology. "
        End If
    
        If mvarBoneLoss = True Then
            strNote = strNote & "Bone loss. "
        End If
    
        If mvarImpactedTooth = True Then
            strNote = strNote & "Impacted tooth. "
        End If
    
        If mvarOtherFindings = True Then
            strNote = strNote & "Other findings. "
        End If
    
        If strNote <> "" Then
            strNote = strNote & vbCrLf
        End If
    
        GetNote = strNote
    
    End If

End Function

Public Function GetRTFNote() As String

    GetRTFNote = GetNote

End Function


Private Function UpdateDirtyFlag()

    If mvarDirtyFlag = dfNoAction Then
        mvarDirtyFlag = dfUpdate
    End If
    
End Function

'*******************************************************************
'Function: SaveSection
'Author: ASamedov
'Date: 01/09/2002
'Returns: True/False
'Purpose: Saves the data to the table
'*******************************************************************
Public Function SaveSection() As Boolean

Dim sSQL As String

Const sPROC_NAME As String = "SaveSection"

On Error GoTo ErrHnd

    If mvarDirtyFlag = dfNoAction Then
        SaveSection = True
        Exit Function
    Else
'        'Update the workload
'        If UpdateWorkload(mvarWorkloadID, mvarDirtyFlag, mvarDentalEncounterID, GetWorkloadCode, 1) = False Then
'            Exit Function
'        End If
    End If
    
    If mvarDirtyFlag = dfdelete Then
        'Delete the record
        If mvarFindingsID <> 0 Then
            sSQL = "UPDATE TEMPL_RADIOLOGY_FINDINGS SET DELETED = 1 Where (FINDINGS_ID = " & mvarFindingsID & ")"
        Else
            SaveSection = True
            Exit Function
        End If
        
    ElseIf mvarDirtyFlag = dfAdd Then
        'The record does not exist, so create an insert statement
        mvarFindingsID = goconn.GetNextID(esTEMPL_RADIOLOGY_FINDINGS_SEQ)
'        sSQL = "Select TEMPL_RADIOLOGY_FINDINGS_SEQ.NEXTVAL ID from Dual"
'        Set rs = goconn.OpenRecordset(sSQL)
'        If Not rs.EOF Then
'            mvarFindingsID = rs("ID")
'        Else
'            goCWShared.ShowVBError 0, "Could not get the NextVal from TEMPL_RADIOLOGY_FINDINGS_SEQ sequence", "clsIV." & sPROC_NAME, App.Title, vbCritical
'            Exit Function
'        End If
'        rs.Close
        
        sSQL = "Insert into TEMPL_RADIOLOGY_FINDINGS (FINDINGS_ID, DENTAL_ENCOUNTER_ID" _
            & ", TOOTH_NUMBER, FINDINGS_NORMAL, CARIES_NOTED, PERIAPICAL_PATHOLOGY, BONE_LOSS, IMPACTED_TOOTH" _
            & ", OTHER_FINDINGS, DELETED) Values (" _
            & "" & mvarFindingsID & ", " & mvarDental_Encounter_ID _
            & ",'" & mvarToothNumber & "', " & CInt(mvarFindingsNormal) & ", " & CInt(mvarCariesNoted) & ", " & CInt(mvarPeriapicalPathology) & ", " & CInt(mvarBoneLoss) & ", " & CInt(mvarImpactedTooth) & "" _
            & ", " & CInt(mvarOtherFindings) & " ," & "0" & ")"
            
    ElseIf mvarDirtyFlag = dfUpdate Then
        'The record exists, so create an update statement
        sSQL = "Update TEMPL_RADIOLOGY_FINDINGS" _
            & " Set TOOTH_NUMBER ='" & mvarToothNumber & "'," _
            & " FINDINGS_NORMAL = " & CInt(mvarFindingsNormal) & "," _
            & " CARIES_NOTED =" & CInt(mvarCariesNoted) & "," _
            & " PERIAPICAL_PATHOLOGY =" & CInt(mvarPeriapicalPathology) & "," _
            & " BONE_LOSS = " & CInt(mvarBoneLoss) & "," _
            & " IMPACTED_TOOTH=" & CInt(mvarImpactedTooth) & "," _
            & " OTHER_FINDINGS =" & CInt(mvarOtherFindings) & "," _
            & " DELETED=" & Abs(CInt(mvarDeleted)) _
            & " Where FINDINGS_ID = " & mvarFindingsID
            
    End If
    goconn.ExecuteSQL sSQL
        
    mvarDirtyFlag = dfNoAction
    
    SaveSection = True

Exit Function

ErrHnd:
    goCWShared.ShowVBError Err.Number, Err.Description, Err.Source & "." & sPROC_NAME, App.Title, vbCritical
    
    'Delete the workload credit if it didn't update correctly
    'UpdateWorkload mvarWorkloadID, dfDelete, mvarDental_Encounter_ID, "", 1
    
End Function
