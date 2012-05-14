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
Attribute VB_Name = "colTeeth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsTooth"
Attribute VB_Ext_KEY = "Member0" ,"clsTooth"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable to hold collection
Private mCol As Collection
Private mvarDentalEncounterTeethID As Long
Friend Function RefreshTeethFromRS(DentalEncounterTeethID As Long, rsSSTeeth As Recordset) As Boolean

'This function populates the teeth collection with the teeth
'that have an active status

'Dim rs As New Recordset, sSQL As String
Dim oTempTooth As clsTooth

'Note: there is only going to be one tooth, because the DentalEncounterTeethID is the PK
'for the teeth table

    rsSSTeeth.Filter = ""
    
    If Not rsSSTeeth.EOF Then
        
        rsSSTeeth.Filter = "DENTAL_ENCOUNTER_TEETH_ID=" & DentalEncounterTeethID
        
        'This is used for saving
        mvarDentalEncounterTeethID = DentalEncounterTeethID
    End If
    
    Do Until rsSSTeeth.EOF
        'Add the Tooths to the Teeth collection
        Set oTempTooth = Me.Add(rsSSTeeth("DENTAL_ENCOUNTER_ID"), rsSSTeeth("DENTAL_ENCOUNTER_TEETH_ID"))
        oTempTooth.DRC = rsSSTeeth("DRC")
        oTempTooth.ToothID = rsSSTeeth("TOOTH_ID")
        oTempTooth.ToothName = rsSSTeeth("TOOTH_NAME") & ""
        oTempTooth.MesialVersion = Val(rsSSTeeth("MESIAL_VERSION") & "")
        oTempTooth.DistalVersion = Val(rsSSTeeth("DISTAL_VERSION") & "")
        oTempTooth.FacialVersion = Val(rsSSTeeth("FACIAL_VERSION") & "")
        oTempTooth.LingualVersion = Val(rsSSTeeth("LINGUAL_VERSION") & "")
        oTempTooth.RotatedMesial = Val(rsSSTeeth("ROTATED_MESIAL") & "")
        oTempTooth.RotatedDistal = Val(rsSSTeeth("ROTATED_DISTAL") & "")
        oTempTooth.Intruded = Val(rsSSTeeth("INTRUDED") & "")
        oTempTooth.Extruded = Val(rsSSTeeth("Extruded") & "")
        oTempTooth.Deleted = Val(rsSSTeeth("TEETH_DELETED") & "")
        'This tooth is clean
        oTempTooth.DirtyFlag = dfNoAction
        Set oTempTooth = Nothing

        rsSSTeeth.MoveNext
    Loop
    
End Function

Public Function Add(DentalEncounterID As Long, Optional DentalEncounterTeethID As Long = 0) As clsTooth
    'create a new object
    Dim objNewMember As clsTooth
    Set objNewMember = New clsTooth

    'set the primary key
    objNewMember.DentalEncounterTeethID = DentalEncounterTeethID
    objNewMember.DentalEncounterID = DentalEncounterID
    
    'If the DentalEncounterTeethID = 0 then the user is adding a new tooth
    If DentalEncounterTeethID = 0 Then
        objNewMember.DirtyFlag = dfAdd
        mCol.Add objNewMember
    Else
        'Not needed because it is the default for this object
        objNewMember.DirtyFlag = dfNoAction
        mCol.Add objNewMember, CStr(DentalEncounterTeethID)
    End If
    
    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As clsTooth
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

    If mCol(vntIndexKey).ToothID = 0 Then
        'The tooth was never saved so remove it from the collection
        mCol.Remove vntIndexKey
    Else
        'The tooth was saved, so set the deleted property to deleted
        mCol(vntIndexKey).Deleted = True
    End If

End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property


Friend Function SaveTeeth(ByRef DentalEncounterTeethID As Long) As Boolean

'This function loops through the collection and saves each tooth

Dim oTooth As clsTooth

    SaveTeeth = True
    
    For Each oTooth In mCol
        'If one tooth doesn't save right then SaveTeeth returns false
        If oTooth.SaveTooth(DentalEncounterTeethID) = False Then
            SaveTeeth = False
        End If
    Next

End Function

Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub

