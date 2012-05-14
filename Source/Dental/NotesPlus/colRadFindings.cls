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
Attribute VB_Name = "colRadFindings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mlngDentalEncounterId As Long
Private mlngPatientUnitId As Long

Private mlngFindingsId As Long
Private mlngDirtyFlag As DirtyFlagEnum
Private mcolcls As Collection

Public Property Get FindingsID() As Long
    FindingsID = mlngFindingsId
End Property

Private Sub Class_Initialize()
    Set mcolcls = New Collection
    mlngDirtyFlag = dfNoAction
End Sub

Public Property Get DentalEncounterID() As Long
    DentalEncounterID = mlngDentalEncounterId
End Property

Public Function GetNote() As String
    Dim strNote As String
    Dim strTemp As String
    Dim objItem As clsRadFindItem
        
    strTemp = Empty
    
    If Not mcolcls Is Nothing Then
        For Each objItem In mcolcls
            If Not objItem.Deleted Then
                strTemp = strTemp & vbTab & objItem.GetNote
            End If
        Next
    End If
    
    If strTemp <> "" Then
        strNote = strNote & vbTab & "Periapical Tooth Findings:" & vbCrLf & strTemp
    End If
    
    GetNote = strNote
    
End Function

Private Function GetRTFNote() As String

End Function



Public Property Get Item(vntIndexKey As Variant) As clsRadFindItem
  Set Item = mcolcls(vntIndexKey)
End Property
Public Property Get Count() As Long
    If Not mcolcls Is Nothing Then
        Count = mcolcls.Count
    End If
End Property
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mcolcls.[_NewEnum]
End Property


Private Sub UpdateDirtyFlag()

    If mlngDirtyFlag = dfNoAction Then
        mlngDirtyFlag = dfUpdate
    End If
    
End Sub
'**************************************************
'Sub: LoadSection
'Author: ASamedov
'Returns: True/False
'Purpose: This loads the data from the db
'***************************************************
Public Function LoadSection(ByVal PatientID As Long, ByVal DentalEncounterID As Long) As Boolean
    Dim sSQL As String
    Dim objRS As ADODB.Recordset
    Dim objRadFindItem As clsRadFindItem
    
Const sPROC_NAME = "LoadSection"

On Error GoTo ErrHnd
    
    'mvarPatientID = PatientID
    mlngDentalEncounterId = DentalEncounterID
    
    sSQL = "Select FINDINGS_ID, TOOTH_NUMBER, FINDINGS_NORMAL, CARIES_NOTED, PERIAPICAL_PATHOLOGY," _
           & " BONE_LOSS, IMPACTED_TOOTH, OTHER_FINDINGS, DELETED" _
           & " FROM TEMPL_RADIOLOGY_FINDINGS WHERE DENTAL_ENCOUNTER_ID = " & mlngDentalEncounterId & " AND DELETED = 0 ORDER BY FINDINGS_ID"
             
    Set objRS = goconn.OpenRecordset(sSQL)
    
    Set mcolcls = New Collection
    
        Do Until objRS.EOF
            'Set the values from the db including "" to allow for nulls
            Set objRadFindItem = Add(mlngDentalEncounterId)
            objRadFindItem.Dental_Encounter_ID = mlngDentalEncounterId
            objRadFindItem.FindingsID = Val(objRS("FINDINGS_ID") & "")
            objRadFindItem.ToothNumber = objRS("TOOTH_NUMBER") & ""
            objRadFindItem.FindingsNormal = Abs(Val(objRS("FINDINGS_NORMAL") & ""))
            objRadFindItem.CariesNoted = Abs(Val(objRS("CARIES_NOTED") & ""))
            objRadFindItem.PeriapicalPathology = Abs(Val(objRS("PERIAPICAL_PATHOLOGY") & ""))
            objRadFindItem.BoneLoss = Abs(Val(objRS("BONE_LOSS") & ""))
            objRadFindItem.ImpactedTooth = Abs(Val(objRS("IMPACTED_TOOTH") & ""))
            objRadFindItem.OtherFindings = Abs(Val(objRS("OTHER_FINDINGS") & ""))
            objRadFindItem.Deleted = Abs(Val(objRS("DELETED") & ""))
            objRadFindItem.DirtyFlag = dfNoAction
            objRS.MoveNext
            
        Loop
    
    objRS.Close
    
    Set objRS = Nothing
    
    LoadSection = True
        
Exit Function

ErrHnd:
    goCWShared.ShowVBError Err.Number, Err.Description, Err.Source & "." & sPROC_NAME, App.Title, vbCritical
    
End Function

'*******************************************************************
'Function: SaveSection
'Author: ASamedov
'Date: 01/09/2002
'Returns: True/False
'Purpose: Saves the oLocal objects
'*******************************************************************
Public Function SaveSection() As Boolean

Dim oRadiologyFindItem As clsRadFindItem
    
    SaveSection = True
    
    For Each oRadiologyFindItem In mcolcls
        If oRadiologyFindItem.SaveSection = False Then
            SaveSection = False
        End If
    Next
    
End Function

Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)
    
    If mcolcls(vntIndexKey).FindingsID <> 0 Then
        'Set the dirtyflag to deleted
        mcolcls(vntIndexKey).Deleted = True
    Else
        'The object has not been saved, so remove it.
        mcolcls.Remove vntIndexKey
    End If
    
End Sub


Public Function Add(DentalEncounter As Long) As clsRadFindItem
    
    'create a new object
    Dim oRadiologyFindItem As clsRadFindItem
    Set oRadiologyFindItem = New clsRadFindItem

    'set the properties passed into the method
    oRadiologyFindItem.Dental_Encounter_ID = DentalEncounter
    oRadiologyFindItem.Deleted = False
    oRadiologyFindItem.DirtyFlag = dfAdd
    
    'Add the member to the collection
    mcolcls.Add oRadiologyFindItem

    'return the object created
    Set Add = oRadiologyFindItem
    Set oRadiologyFindItem = Nothing

End Function

Private Sub Class_Terminate()
    Set mcolcls = Nothing
End Sub
