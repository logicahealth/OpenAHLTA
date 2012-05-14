VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colDxs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsDx"
Attribute VB_Ext_KEY = "Member0" ,"clsDx"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable to hold collection
Private mCol As Collection
Private mvarTxCodeID As Long
Private mvarToothID As Long
Private mvarCurrentDentalEncounterID As Long
Private mvarGroupID As Long

'Collection to hold all Txs to ensure unique data integrity
Private mvarTxsRef As Collection

'Collection to hold all Dxs to ensure unique data integrity
Private mvarDxsRef As Collection



Friend Property Let GroupID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.GroupID = 5
    mvarGroupID = vData
End Property


Friend Property Get GroupID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.GroupID
    GroupID = mvarGroupID
End Property




Friend Property Let CurrentDentalEncounterID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CurrentDentalEncounterID = 5
    mvarCurrentDentalEncounterID = vData
    
    'This is used when the current encounter changes and the datalayer is not reloaded
    Dim oDx As clsDx
    For Each oDx In mCol
        oDx.CurrentDentalEncounterID = mvarCurrentDentalEncounterID
    Next
        
End Property


Friend Property Get CurrentDentalEncounterID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CurrentDentalEncounterID
    CurrentDentalEncounterID = mvarCurrentDentalEncounterID
End Property




Private Function AddInternal(DentalEncounterID As Long, DxICDCodeID As Long, DxCodeID As Long) As clsDx

    'create a new object
    Dim objNewMember As clsDx
    Set objNewMember = New clsDx
    
    Dim oTest As New clsDx
    
    'set the properties passed into the method
    objNewMember.DentalEncounterID = DentalEncounterID
    objNewMember.CurrentDentalEncounterID = mvarCurrentDentalEncounterID
    objNewMember.DxCodeID = DxCodeID
   
'Removing source
'    objNewMember.Source = Source
    
    'This one already has an id in the database
    
    'Check if this object exists somewhere else in the project, if so
    ' then add the reference to the existing object
    On Error GoTo AddErr
    Set oTest = mvarDxsRef(CStr(DxCodeID))
    On Error GoTo 0
    If oTest.DxCodeID = DxCodeID Then
        'The object already exists so just add the reference
        mCol.Add oTest, CStr(DxCodeID)
        Set AddInternal = Nothing
        Set oTest = Nothing
        Exit Function
    End If
    
    'The object does not exist in mvarDxsRef so add it
    mvarDxsRef.Add objNewMember
    
    'Set the mvarDxsRef and mvarTxsRef in the new Dx object
    objNewMember.DxsRef = mvarDxsRef
    
    'Add the rest of the properties
    objNewMember.InternalAdd = True
    objNewMember.DxICDCodeID = DxICDCodeID
    objNewMember.InternalAdd = False
    objNewMember.GroupID = mvarGroupID
    objNewMember.DirtyFlag = dfNoAction
    mCol.Add objNewMember, CStr(DxCodeID)

    'return the object created
    Set AddInternal = objNewMember
    Set objNewMember = Nothing
    Set oTest = Nothing

Exit Function

AddErr:
    If Err.Number = 5 Or Err.Number = 9 Then
        'Invalid assignment or Out of range error, so continue
        Resume Next
    Else
        Err.Raise Err.Number, Err.Source, Err.Description
    End If

End Function

Public Function Add(Chartable As SourceEnum, DxICDCodeID As Long, ProviderID As Long) As clsDx
    
    'create a new object
    Dim objNewMember As clsDx
    Set objNewMember = New clsDx
    
    Dim oTest As New clsDx
    
    'set the dentalencounterid of the new object to the CurrentDentalEncounterID
    objNewMember.DentalEncounterID = mvarCurrentDentalEncounterID
    objNewMember.CurrentDentalEncounterID = mvarCurrentDentalEncounterID
    objNewMember.ModifiedInDentalEncounterID = mvarCurrentDentalEncounterID
    
    
    'The user is adding a new one
    objNewMember.DxICDCodeID = DxICDCodeID
    
    'Set the defaults
    objNewMember.ProviderID = ProviderID
    objNewMember.CreateDateTime = GetSysDate
    objNewMember.StatusID = dxActiveNew
    objNewMember.Signed = False
    objNewMember.PerioStateID = DxNotPerio
    objNewMember.GroupID = mvarGroupID

' Removing source
'    objNewMember.Source = Source

'    'Notes can only create other notable Dxs, and charts can only create other chartable Dxs
'    If objNewMember.Source = sCharts Then
        objNewMember.Chartable = True
'    Else
'        objNewMember.Chartable = False
'    End If
    objNewMember.Chartable = Chartable

    'The object does not exist in mvarDxsRef so add it
    mvarDxsRef.Add objNewMember
    
    'Set the mvarDxsRef and mvarTxsRef in the new Dx object
    objNewMember.DxsRef = mvarDxsRef
    
    objNewMember.DirtyFlag = dfAdd
    mCol.Add objNewMember

    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing
    Set oTest = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As clsDx
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property



Public Function AddObj(Dx As clsDx) As Boolean

'This only works (now) if the dx is added to the new collection AND removed from the old collection.

'This function adds the passed in Dx to this collection

    AddObj = False
    
    If Dx.DxCodeID = 0 Then
        'The Dx does not have a primary key yet, so add it to the collection
        'without a key
        mCol.Add Dx
    Else
        'The Dx has a key so add it to the collection with a key
        mCol.Add Dx, CStr(Dx.DxCodeID)
    End If
    
    AddObj = True
    
End Function


Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property


Friend Function RefreshDxsFromRS(ByRef DxsRs As Recordset, GroupID As Long, CurrentDentalEncounterID As Long) As Boolean

'This function populates the Dxs given the preloaded Dxs recordset
Dim oTempDx As clsDx
    
    mvarCurrentDentalEncounterID = CurrentDentalEncounterID
    mvarGroupID = GroupID
    
    DxsRs.Filter = ""
    If Not DxsRs.EOF Then
        DxsRs.Filter = "Group_ID = " & GroupID
        
        'Release the previous collection
        Set mCol = New Collection
        
        'Add the Dx objects to the Dx collection
        Do Until DxsRs.EOF
            Set oTempDx = AddInternal(DxsRs("DENTAL_ENCOUNTER_ID"), DxsRs("DX_ICD_CODE_ID"), DxsRs("DX_CODE_ID"))
            
            'oTempDx will be nothing if the object already exists, because the add method added the reference
            If Not oTempDx Is Nothing Then
                'Add the additional fields
                oTempDx.StatusID = Val(DxsRs("STATUS_ID") & "")
                oTempDx.Priority = Val(DxsRs("PRIORITY") & "")
                oTempDx.ProviderID = Val(DxsRs("PROVIDER_ID") & "")
                If Not IsNull(DxsRs("CREATE_DATE_TIME")) Then
                    'Convert the GMT time to local
                    oTempDx.CreateDateTime = cwobjCWComm.GmtToLocal(DxsRs("CREATE_DATE_TIME"))
                End If
                oTempDx.Chartable = CBool(Val(DxsRs("CHARTABLE") & ""))
                oTempDx.DRC = Val(DxsRs("DRC") & "")
                oTempDx.Signed = CBool(Val(DxsRs("SIGNED") & ""))
                If Val(DxsRs("Perio_State") & "") = 0 Then
                    'For backwards compatility with old data.
                    oTempDx.PerioStateID = DxNotPerio
                Else
                    oTempDx.PerioStateID = Val(DxsRs("Perio_State") & "")
                End If
                'These are not needed because they are populated by setting the DxICDCodeID
                oTempDx.ICDCode = DxsRs("ICD_CODE") & ""
                oTempDx.ICDCategory = DxsRs("ICD_CATEGORY") & ""
                oTempDx.ICDDefinition = DxsRs("ICD_DEFINITION") & ""
                oTempDx.ChartRelated = Val(DxsRs("CHART_RELATED") & "")
                oTempDx.DRC = Val(DxsRs("DRC") & "")
                oTempDx.DRCMin = Val(DxsRs("DRCMIN") & "")
                oTempDx.DRCMax = Val(DxsRs("DRCMAX") & "")
                oTempDx.DeleteReason = Val(DxsRs("DELETE_REASON") & "")
                oTempDx.ToothID = mvarToothID
                If Not IsNull(DxsRs("MODIFIED_DATE_TIME")) Then
                    oTempDx.ModifiedDateTime = cwobjCWComm.GmtToLocal(DxsRs("MODIFIED_DATE_TIME")) & ""
                End If
                'The Dx is clean
                'This is here because it is updated in SetDirtyFlag
                oTempDx.ModifiedInDentalEncounterID = Val(DxsRs("MODIFIED_IN_DENTAL_ENCOUNTER"))
                oTempDx.DirtyFlag = dfNoAction
            End If
            
            DxsRs.MoveNext
        Loop
        
        Set oTempDx = Nothing
    End If
    
    RefreshDxsFromRS = True
    
End Function
Friend Function RefreshLocSubLocFromRS(rsLocSubLoc As Recordset) As Boolean

Dim oDx As clsDx

    'Go through each dx in mvardxs and refresh their SS and Teeth
    For Each oDx In mCol
        If oDx.Chartable = False Then
            oDx.Locations.RefreshLocationsFromDxCodeIDRS oDx.DxCodeID, 0, rsLocSubLoc
        End If
    Next
    
    RefreshLocSubLocFromRS = True

End Function

Friend Function RefreshSSTeethFromRS(rsSSTeeth As Recordset) As Boolean

Dim oDx As clsDx

    'Go through each dx in mvardxs and refresh their SS and Teeth
    For Each oDx In mCol
        If oDx.Chartable = True Then
            oDx.SurfaceSets.RefreshSurfaceSetsFromRS oDx.DxCodeID, 0, rsSSTeeth
        End If
    Next
    
    RefreshSSTeethFromRS = True
    
End Function


Public Function RemoveObj(Dx As clsDx) As Boolean

'This function removes the dx reference from this collection

Dim lCounter As Long

    RemoveObj = False
    
    'Remove the Tx based on the index in the collection
    For lCounter = 1 To mCol.Count
        If mCol(lCounter) Is Dx Then
            mCol.Remove (lCounter)
            Exit For
        End If
    Next
    
    RemoveObj = True

End Function

Friend Property Let TxsRef(vData As Collection)

    Set mvarTxsRef = vData
    
End Property

Friend Property Let DxsRef(vData As Collection)

    Set mvarDxsRef = vData
    
End Property

Friend Property Get TxsRef() As Collection

    Set TxsRef = mvarTxsRef

End Property

Friend Property Get DxsRef() As Collection

    Set DxsRef = mvarDxsRef

End Property

Private Function RefreshDxsBySQL(SQL As String, Source As SourceEnum) As Boolean

'*********************** NOT USED ANYMORE - See RefreshDxsFromRS

''This function populates the Dxs collection given a sql statement
'
'Dim rs As Recordset
'Dim oTempDx As clsDx
'
'    Set rs = New Recordset
'    Set rs = goconn.OpenRecordset(SQL)
'
'    'Release the previous collection
'    Set mCol = New Collection
'
'    'Add the Dx objects to the Dx collection
'    Do Until rs.EOF
'        Set oTempDx = AddInternal(rs("DENTAL_ENCOUNTER_ID"), Source, rs("DX_ICD_CODE_ID"), rs("DX_CODE_ID"))
'
'        'oTempDx will be nothing if the object already exists, because the add method added the reference
'        If Not oTempDx Is Nothing Then
'            'Add the additional fields
'            oTempDx.StatusID = Val(rs("STATUS_ID") & "")
'            oTempDx.Priority = Val(rs("PRIORITY") & "")
'            oTempDx.ProviderID = Val(rs("PROVIDER_ID") & "")
'            oTempDx.CreateDateTime = rs("CREATE_DATE_TIME") & ""
'            oTempDx.Chartable = CBool(Val(rs("CHARTABLE") & ""))
'            oTempDx.DRC = Val(rs("DRC") & "")
'            oTempDx.Signed = CBool(Val(rs("SIGNED") & ""))
'            If Val(rs("Perio_State") & "") = 0 Then
'                'For backwards compatility with old data.
'                oTempDx.PerioStateID = DxNotPerio
'            Else
'                oTempDx.PerioStateID = Val(rs("Perio_State") & "")
'            End If
'            'These are not needed because they are populated by setting the DxICDCodeID
'            oTempDx.ICDCode = rs("ICD_CODE") & ""
'            oTempDx.ICDCategory = rs("ICD_CATEGORY") & ""
'            oTempDx.ICDDefinition = rs("ICD_DEFINITION") & ""
'            oTempDx.ChartRelated = Val(rs("CHART_RELATED") & "")
'            oTempDx.DRC = Val(rs("DRC") & "")
'            oTempDx.DRCMin = Val(rs("DRCMIN") & "")
'            oTempDx.DRCMax = Val(rs("DRCMAX") & "")
'            oTempDx.DeleteReason = Val(rs("DELETE_REASON") & "")
'            oTempDx.ToothID = mvarToothID
'            If Not IsNull(rs("MODIFIED_DATE_TIME")) Then
'                oTempDx.ModifiedDateTime = rs("MODIFIED_DATE_TIME") & ""
'            End If
'            'The Dx is clean
'            'This is here because it is updated in SetDirtyFlag
'            oTempDx.ModifiedInDentalEncounterID = Val(rs("MODIFIED_IN_DENTAL_ENCOUNTER"))
'            oTempDx.DirtyFlag = dfNoAction
'        End If
'
'        rs.MoveNext
'    Loop
'
'    Set oTempDx = Nothing
'    rs.Close
'    Set rs = Nothing
'
'    RefreshDxsBySQL = True
    
End Function


Public Sub Remove(vntIndexKey As Variant, DentalEncounterID As Long)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)
    
    If mCol(vntIndexKey).DentalEncounterID = DentalEncounterID Then
        'The Dx is from this encounter
        If mCol(vntIndexKey).Signed = False Then
            'The Dx is unsigned, so remove the Dx from the database
            mCol(vntIndexKey).StatusID = dxNonActiveDeleted
            mCol(vntIndexKey).DirtyFlag = dfdelete
        Else
            'The Dx is signed, so set the statusid to deleted
            mCol(vntIndexKey).StatusID = dxNonActiveDeleted
        End If
    Else
        'The Dx is from an old encounter, and it doesn't matter if the old
        ' encounter is signed or not. So, set the statusid to deleted
        mCol(vntIndexKey).StatusID = dxNonActiveDeleted
    End If
    
'    If mCol(vntIndexKey).DxCodeID = 0 Then
'        'The Dx was never saved so remove it from the collection
'        mCol.Remove vntIndexKey
'    Else
'        'The Dx was saved, so set the statusid to deleted
'        mCol(vntIndexKey).StatusID = dxNonActiveDeleted
'    End If
    
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property


Friend Function SaveDxs(GroupID As Long) As Boolean

'Save each Dx
'TxCodeID is the ID that the Dx belongs to.
'Only applicable when this collection is under a Tx object

Dim oDx As clsDx
Dim lCounter As Long

    SaveDxs = True
    
    'Use for-next (not for-each) because I need the index position
    'Must count backwards to solve the problem with deleting a dx and all the indexes changing.
    For lCounter = mCol.Count To 1 Step -1
        Set oDx = mCol(lCounter)
        
        If oDx.SaveDx(GroupID) = False Then
           'If one dx doesn't save properly, then set SaveDxs = false
            SaveDxs = False
        End If
        
        'Remove all the Dxs that were not saved because they were deleted in this encounter
        If oDx.StatusID = dxNonActiveDeleted And oDx.DxCodeID = 0 Then
            mCol.Remove lCounter
        End If
    Next
    
    Set oDx = Nothing
     
   
End Function

Friend Function RefreshDxsFromPatientID(GroupID As Long, PatientID As Long, Source As SourceEnum, CurrentDentalEncounterID As Long, Optional ToothID As Long = 0) As Boolean

'************* NOT USED ANYMORE - see RefreshDxsFromRS

''This function loads all of the Dxs given the patientid, source and toothid.
'
'Dim sSQL As String
'
'    mvarToothID = ToothID
'    mvarCurrentDentalEncounterID = CurrentDentalEncounterID
'    mvarGroupID = GroupID
'
''    If Source = sCharts Then
''        'The grid will display all active and chartable and unsigned dxs
''        'The graphical tooth will display all active and unsigned dxs
''
''        If mvarToothID = 0 Then
''            'Don't filter on ToothID
''            sSQL = "SELECT DENTAL_ENCOUNTER.PATIENT_UNIT_ID," _
''                & " DENTAL_ENCOUNTER_DX.DX_CODE_ID," _
''                & " DENTAL_ENCOUNTER_DX.DX_ICD_CODE_ID," _
''                & " DENTAL_ENCOUNTER_DX.CREATE_DATE_TIME," _
''                & " DENTAL_ENCOUNTER_DX.MODIFIED_DATE_TIME," _
''                & " DENTAL_ENCOUNTER_DX.PRIORITY," _
''                & " DENTAL_ENCOUNTER_DX.PROVIDER_ID," _
''                & " DENTAL_ENCOUNTER_DX.DENTAL_ENCOUNTER_ID," _
''                & " DENTAL_ENCOUNTER_DX.STATUS_ID," _
''                & " DENTAL_ENCOUNTER_DX.CHARTABLE," _
''                & " DENTAL_ENCOUNTER_DX.DRC," _
''                & " DENTAL_ENCOUNTER_DX.SIGNED," _
''                & " DENTAL_ENCOUNTER_DX.PERIO_STATE," _
''                & " DENTAL_ENCOUNTER_DX.DELETE_REASON," _
''                & " DX_ICD_CODE_DEFINITION.ICD_CATEGORY," _
''                & " DX_ICD_CODE_DEFINITION.ICD_CODE," _
''                & " DX_ICD_CODE_DEFINITION.ICD_DEFINITION," _
''                & " DX_ICD_CODE_DEFINITION.DRC AS DEFAULT_DRC," _
''                & " DX_ICD_CODE_DEFINITION.CHART_RELATED," _
''                & " DX_ICD_CODE_DEFINITION.EFFECTIVE_DATE," _
''                & " DX_ICD_CODE_DEFINITION.EFFECTIVE_END_DATE," _
''                & " DX_ICD_CODE_DEFINITION.DRCMIN," _
''                & " DX_ICD_CODE_DEFINITION.DRCMax" _
''                & " FROM DENTAL_ENCOUNTER_DX, DENTAL_ENCOUNTER," _
''                & " DX_ICD_CODE_DEFINITION, DENTAL_ENCOUNTER_GROUPS"
''            sSQL = sSQL & " Where DENTAL_ENCOUNTER_DX.DENTAL_ENCOUNTER_ID = DENTAL_ENCOUNTER.DENTAL_ENCOUNTER_ID" _
''                & " AND DENTAL_ENCOUNTER_DX.DX_ICD_CODE_ID = DX_ICD_CODE_DEFINITION.DX_ICD_CODE_ID" _
''                & " AND DENTAL_ENCOUNTER_DX.DX_CODE_ID = DENTAL_ENCOUNTER_GROUPS.DX_CODE_ID" _
''                & " AND ((DENTAL_ENCOUNTER_DX.STATUS_ID <> 21) AND" _
''                & " (DENTAL_ENCOUNTER.PATIENT_UNIT_ID = " & PatientID & ") AND" _
''                & " (DENTAL_ENCOUNTER_DX.SIGNED = 0) AND" _
''                & " (DENTAL_ENCOUNTER_DX.CHARTABLE = - 1) AND" _
''                & " (DENTAL_ENCOUNTER_GROUPS.GROUP_ID = " & GroupID & " AND (DENTAL_ENCOUNTER_DX.DX_CODE_ID > 0) ) OR" _
''                & " (DENTAL_ENCOUNTER_DX.STATUS_ID BETWEEN 10 AND" _
''                & " 14) AND (DENTAL_ENCOUNTER.PATIENT_UNIT_ID = " & PatientID & ")" _
''                & " AND (DENTAL_ENCOUNTER_DX.SIGNED = - 1) AND" _
''                & " (DENTAL_ENCOUNTER_DX.CHARTABLE = - 1) AND" _
''                & " (DENTAL_ENCOUNTER_GROUPS.GROUP_ID = " & GroupID & ") AND (DENTAL_ENCOUNTER_DX.DX_CODE_ID > 0) )" _
''                & " ORDER BY DENTAL_ENCOUNTER_DX.CREATE_DATE_TIME"
''
''        Else
''            'Filter on the toothid
''            sSQL = "SELECT DENTAL_ENCOUNTER.PATIENT_UNIT_ID," _
''                & " DENTAL_ENCOUNTER_DX.DX_CODE_ID," _
''                & " DENTAL_ENCOUNTER_DX.DX_ICD_CODE_ID," _
''                & " DENTAL_ENCOUNTER_DX.CREATE_DATE_TIME," _
''                & " DENTAL_ENCOUNTER_DX.MODIFIED_DATE_TIME," _
''                & " DENTAL_ENCOUNTER_DX.PRIORITY," _
''                & " DENTAL_ENCOUNTER_DX.PROVIDER_ID," _
''                & " DENTAL_ENCOUNTER_DX.DENTAL_ENCOUNTER_ID," _
''                & " DENTAL_ENCOUNTER_DX.STATUS_ID," _
''                & " DENTAL_ENCOUNTER_DX.CHARTABLE," _
''                & " DENTAL_ENCOUNTER_DX.DRC," _
''                & " DENTAL_ENCOUNTER_DX.SIGNED," _
''                & " DENTAL_ENCOUNTER_DX.PERIO_STATE," _
''                & " DENTAL_ENCOUNTER_DX.DELETE_REASON," _
''                & " DENTAL_ENCOUNTER_TEETH.TOOTH_ID," _
''                & " DX_ICD_CODE_DEFINITION.ICD_CATEGORY," _
''                & " DX_ICD_CODE_DEFINITION.ICD_CODE," _
''                & " DX_ICD_CODE_DEFINITION.ICD_DEFINITION," _
''                & " DX_ICD_CODE_DEFINITION.DRC AS DEFAULT_DRC," _
''                & " DX_ICD_CODE_DEFINITION.CHART_RELATED," _
''                & " DX_ICD_CODE_DEFINITION.EFFECTIVE_DATE," _
''                & " DX_ICD_CODE_DEFINITION.EFFECTIVE_END_DATE," _
''                & " DX_ICD_CODE_DEFINITION.DRCMIN," _
''                & " DX_ICD_CODE_DEFINITION.DRCMax"
''            sSQL = sSQL & " FROM DENTAL_ENCOUNTER_DX, DENTAL_ENCOUNTER," _
''                & " DENTAL_ENCOUNTER_DXS_SS," _
''                & " DENTAL_ENC_SURFACE_SET, DENTAL_ENCOUNTER_TEETH," _
''                & " DX_ICD_CODE_DEFINITION, DENTAL_ENCOUNTER_GROUPS" _
''                & " Where DENTAL_ENCOUNTER_DX.DENTAL_ENCOUNTER_ID = DENTAL_ENCOUNTER.DENTAL_ENCOUNTER_ID" _
''                & " AND DENTAL_ENCOUNTER_DX.DX_CODE_ID = DENTAL_ENCOUNTER_DXS_SS.DX_CODE_ID" _
''                & " AND DENTAL_ENCOUNTER_DXS_SS.SURFACE_SET_ID = DENTAL_ENC_SURFACE_SET.SURFACE_SET_ID" _
''                & " AND DENTAL_ENC_SURFACE_SET.DENTAL_ENCOUNTER_TEETH_ID" _
''                & " = DENTAL_ENCOUNTER_TEETH.DENTAL_ENCOUNTER_TEETH_ID" _
''                & " AND DENTAL_ENCOUNTER_DX.DX_ICD_CODE_ID = DX_ICD_CODE_DEFINITION.DX_ICD_CODE_ID" _
''                & " AND DENTAL_ENCOUNTER_DX.DX_CODE_ID = DENTAL_ENCOUNTER_GROUPS.DX_CODE_ID" _
''                & " AND ((DENTAL_ENCOUNTER_TEETH.TOOTH_ID = " & mvarToothID & ") AND" _
''                & " (DENTAL_ENCOUNTER.PATIENT_UNIT_ID = " & PatientID & ") AND" _
''                & " (DENTAL_ENCOUNTER_DX.STATUS_ID <> 21) AND" _
''                & " (DENTAL_ENCOUNTER_DX.SIGNED = 0) AND" _
''                & " (DENTAL_ENCOUNTER_DX.CHARTABLE = - 1) AND" _
''                & " (DENTAL_ENCOUNTER_GROUPS.GROUP_ID = " & GroupID & ") (DENTAL_ENCOUNTER_DX.DX_CODE_ID > 0) OR" _
''                & " (DENTAL_ENCOUNTER_TEETH.TOOTH_ID = " & mvarToothID & ") AND" _
''                & " (DENTAL_ENCOUNTER.PATIENT_UNIT_ID = " & PatientID & ") AND" _
''                & " (DENTAL_ENCOUNTER_DX.STATUS_ID BETWEEN 10 AND" _
''                & " 14) AND (DENTAL_ENCOUNTER_DX.SIGNED = - 1) AND" _
''                & " (DENTAL_ENCOUNTER_DX.CHARTABLE = - 1) AND" _
''                & " (DENTAL_ENCOUNTER_GROUPS.GROUP_ID = " & GroupID & ") AND (DENTAL_ENCOUNTER_DX.DX_CODE_ID > 0) )" _
''                & " ORDER BY DENTAL_ENCOUNTER_DX.CREATE_DATE_TIME"
''
''        End If
''
''    ElseIf Source = sNotes Then
'        'Display active dxs
'        sSQL = "SELECT DENTAL_ENCOUNTER.PATIENT_UNIT_ID," _
'            & " DENTAL_ENCOUNTER_DX.DX_CODE_ID," _
'            & " DENTAL_ENCOUNTER_DX.DX_ICD_CODE_ID," _
'            & " DENTAL_ENCOUNTER_DX.CREATE_DATE_TIME," _
'            & " DENTAL_ENCOUNTER_DX.MODIFIED_DATE_TIME," _
'            & " DENTAL_ENCOUNTER_DX.PRIORITY," _
'            & " DENTAL_ENCOUNTER_DX.PROVIDER_ID," _
'            & " DENTAL_ENCOUNTER_DX.DENTAL_ENCOUNTER_ID," _
'            & " DENTAL_ENCOUNTER_DX.STATUS_ID," _
'            & " DENTAL_ENCOUNTER_DX.CHARTABLE," _
'            & " DENTAL_ENCOUNTER_DX.DRC," _
'            & " DENTAL_ENCOUNTER_DX.SIGNED," _
'            & " DENTAL_ENCOUNTER_DX.PERIO_STATE," _
'            & " DENTAL_ENCOUNTER_DX.DELETE_REASON," _
'            & " DENTAL_ENCOUNTER_DX.MODIFIED_IN_DENTAL_ENCOUNTER," _
'            & " DX_ICD_CODE_DEFINITION.ICD_CATEGORY," _
'            & " DX_ICD_CODE_DEFINITION.ICD_CODE," _
'            & " DX_ICD_CODE_DEFINITION.ICD_DEFINITION," _
'            & " DX_ICD_CODE_DEFINITION.DRC AS DEFAULT_DRC," _
'            & " DX_ICD_CODE_DEFINITION.CHART_RELATED," _
'            & " DX_ICD_CODE_DEFINITION.EFFECTIVE_DATE," _
'            & " DX_ICD_CODE_DEFINITION.EFFECTIVE_END_DATE," _
'            & " DX_ICD_CODE_DEFINITION.DRCMIN," _
'            & " DX_ICD_CODE_DEFINITION.DRCMax"
'        sSQL = sSQL & " FROM DENTAL_ENCOUNTER_DX, DENTAL_ENCOUNTER," _
'            & " DX_ICD_CODE_DEFINITION," _
'            & " DENTAL_ENCOUNTER_GROUPS"
'        sSQL = sSQL & " Where DENTAL_ENCOUNTER_DX.DENTAL_ENCOUNTER_ID = DENTAL_ENCOUNTER.DENTAL_ENCOUNTER_ID" _
'            & " AND DENTAL_ENCOUNTER_DX.DX_ICD_CODE_ID = DX_ICD_CODE_DEFINITION.DX_ICD_CODE_ID" _
'            & " AND DENTAL_ENCOUNTER_DX.DX_CODE_ID = DENTAL_ENCOUNTER_GROUPS.DX_CODE_ID" _
'            & " AND (DENTAL_ENCOUNTER.PATIENT_UNIT_ID = " & PatientID & ") AND" _
'            & " (DENTAL_ENCOUNTER_GROUPS.GROUP_ID = " & GroupID & ") AND" _
'            & " (DENTAL_ENCOUNTER_DX.STATUS_ID <> 21) AND" _
'            & " (DENTAL_ENCOUNTER_DX.DX_CODE_ID > 0)" _
'            & " ORDER BY DENTAL_ENCOUNTER_DX.CREATE_DATE_TIME"
'
''    End If
'
'    'Load the Dxs
'    RefreshDxsFromPatientID = RefreshDxsBySQL(sSQL, Source)

End Function

Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    
    'destroys collection when this class is terminated
    Set mCol = Nothing
    Set mvarTxsRef = Nothing
    Set mvarDxsRef = Nothing

End Sub

