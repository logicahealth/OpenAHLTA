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
Attribute VB_Name = "colTxs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsTx"
Attribute VB_Ext_KEY = "Member0" ,"clsTx"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Event TxAdded(ByVal NewProviderID As Long, ByRef TookOwnership As Boolean)

'local variable to hold collection
Private mCol As Collection

Private mvarSurfaceSet As Long
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
    Dim oTx As clsTx
    For Each oTx In mCol
        oTx.CurrentDentalEncounterID = mvarCurrentDentalEncounterID
    Next

End Property


Friend Property Get CurrentDentalEncounterID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CurrentDentalEncounterID
    CurrentDentalEncounterID = mvarCurrentDentalEncounterID
End Property





Private Function AddInternal(DentalEncounterID As Long, TxADACodeID As Long, TxCodeID As Long) As clsTx

'This one already has an id in the database
Dim objNewMember As clsTx
Dim oTest As New clsTx

    Set objNewMember = New clsTx
    
    'set the properties passed into the method
    objNewMember.DentalEncounterID = DentalEncounterID
    objNewMember.CurrentDentalEncounterID = mvarCurrentDentalEncounterID
    objNewMember.TxCodeID = TxCodeID
    
'Removing source
'    'Used then accessing the Dx property of the Tx object
'    objNewMember.Source = Source
   
    'Check if this object exists somewhere else in the project, if so
    ' then add the reference to the existing object
    On Error GoTo AddInternalErr
    Set oTest = mvarTxsRef(CStr(TxCodeID))
    On Error GoTo 0
    If oTest.TxCodeID = TxCodeID Then
        'The object already exists so just add the reference
        mCol.Add oTest, CStr(TxCodeID)
        Set AddInternal = Nothing
        Set oTest = Nothing
        Exit Function
    End If
    
    'The object does not exist in mvarDxsRef so add it
    mvarTxsRef.Add objNewMember, CStr(TxCodeID)
    
    'Set the mvarDxsRef and mvarTxsRef in the new Tx object
    objNewMember.TxsRef = mvarTxsRef
    
    objNewMember.InternalAdd = True
    objNewMember.TxADACodeID = TxADACodeID
    objNewMember.InternalAdd = False
    objNewMember.GroupID = mvarGroupID
    objNewMember.DirtyFlag = dfNoAction
    mCol.Add objNewMember, CStr(TxCodeID)

    'return the object created
    Set AddInternal = objNewMember
    Set objNewMember = Nothing
    Set oTest = Nothing

Exit Function

AddInternalErr:
    If Err.Number = 5 Or Err.Number = 9 Then
        'Invalid assignment or Out of range error, so continue
        Resume Next
    Else
        Err.Raise Err.Number, Err.Source, Err.Description
    End If

End Function

Public Function RemoveObj(Tx As clsTx) As Boolean

'This function removes the tx reference from this collection

Dim lCounter As Long

    RemoveObj = False
    
    'Remove the Tx based on the index in the collection
    For lCounter = 1 To mCol.Count
        If mCol(lCounter) Is Tx Then
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

Public Function Add(Chartable As SourceEnum, TxADACodeID As Long, ProviderID As Long) As clsTx
    
'The user wants to create a new object
Dim objNewMember As clsTx
Dim oTest As New clsTx
Dim bTookOwnership As Boolean

    'Tracker 17486 Raise the event that checks the other Dxs in the same group for ownership
    RaiseEvent TxAdded(ProviderID, bTookOwnership)
    If bTookOwnership = False Then
        'Don't add anything
        Set Add = Nothing
        Exit Function
    End If
    
    Set objNewMember = New clsTx
    
    'set the properties passed into the method
    objNewMember.DentalEncounterID = mvarCurrentDentalEncounterID
    objNewMember.CurrentDentalEncounterID = mvarCurrentDentalEncounterID
    objNewMember.ModifiedInDentalEncounterID = mvarCurrentDentalEncounterID
    
'Removing source
'    'Used when accessing the Dx property of the Tx object
'    objNewMember.Source = Source
    
    'The user is adding a new one
    objNewMember.TxADACodeID = TxADACodeID
    
    'Set the defaults
    objNewMember.ProviderID = ProviderID
    objNewMember.CreateDateTime = GetSysDate
    objNewMember.StatusID = txPlanned
    objNewMember.PrognosisID = pNone
    objNewMember.Signed = False
    objNewMember.PhaseID = tpNone
    objNewMember.GroupID = mvarGroupID
    
    'The object does not exist in mvarDxsRef so add it
    mvarTxsRef.Add objNewMember
    
    'Set the mvarDxsRef and mvarTxsRef in the new Tx object
    objNewMember.TxsRef = mvarTxsRef
    
'Removing source
'    If objNewMember.Source = sCharts Then
'        objNewMember.Chartable = True
'    Else
'        objNewMember.Chartable = False
'    End If
     objNewMember.Chartable = Chartable

    objNewMember.DirtyFlag = dfAdd
    mCol.Add objNewMember
    
    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing
    Set oTest = Nothing

End Function
Public Function AddObj(Tx As clsTx) As Boolean

'This only works (now) if the tx is added to the new collection AND removed from the old collection.

'This function adds the passed in Tx to this collection

    AddObj = False
    
    If Tx.TxCodeID = 0 Then
        'The Tx does not have a primary key yet, so add it to the collection
        'without a key
        mCol.Add Tx
    Else
        'The Tx has a key so add it to the collection with a key
        mCol.Add Tx, CStr(Tx.TxCodeID)
    End If
    
    AddObj = True
    
End Function

Public Property Get Item(vntIndexKey As Variant) As clsTx
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


Friend Function RefreshTxsFromPatientID(GroupID As Long, PatientID As Long, Source As SourceEnum, CurrentDentalEncounterID As Long, Optional ToothID As Long = 0) As Boolean

'************* NOT USED ANYMORE - see RefreshTxsFromRS

''This function loads all of the Txs given the patientid.
'
'Dim sSQL As String
'Dim rs As Recordset
'Dim oTempTx As clsTx
'
'    mvarToothID = ToothID
'    mvarCurrentDentalEncounterID = CurrentDentalEncounterID
'    mvarGroupID = GroupID
'
''    If Source = sCharts Then
''        'The grid will show Planned and Inprogress
''        'The graphical tooth will show planned, completed, preexisting
''
''        If mvarToothID = 0 Then
''            'Don't filter on the toothid
''            sSQL = "SELECT DENTAL_ENCOUNTER_TX.TX_CODE_ID," _
''                & " DENTAL_ENCOUNTER_TX.DENTAL_ENCOUNTER_ID," _
''                & " DENTAL_ENCOUNTER_TX.TX_ADA_CODE_ID," _
''                & " DENTAL_ENCOUNTER_TX.STATUS_ID," _
''                & " DENTAL_ENCOUNTER_TX.CREATE_DATE_TIME," _
''                & " DENTAL_ENCOUNTER_TX.MODIFIED_DATE_TIME," _
''                & " DENTAL_ENCOUNTER_TX.CHARTABLE," _
''                & " DENTAL_ENCOUNTER_TX.COMMENTS," _
''                & " DENTAL_ENCOUNTER_TX.MATERIALS," _
''                & " DENTAL_ENCOUNTER_TX.PROGNOSIS_ID," _
''                & " DENTAL_ENCOUNTER_TX.SIGNED," _
''                & " DENTAL_ENCOUNTER_TX.PHASE," _
''                & " DENTAL_ENCOUNTER_TX.WORKLOAD_ID," _
''                & " DENTAL_ENCOUNTER_TX.DELETE_REASON," _
''                & " TX_ADA_CODE_DEFINITION.ADA_CATEGORY_CODE," _
''                & " TX_ADA_CODE_DEFINITION.ADA_DEFINITION," _
''                & " TX_ADA_CODE_DEFINITION.ADA_ALIAS," _
''                & " TX_ADA_CODE_DEFINITION.CHART_RELATED," _
''                & " TX_ADA_CODE_DEFINITION.LABFLAG," _
''                & " TX_ADA_CODE_DEFINITION.ADA_HELP_NARATIVE," _
''                & " TX_ADA_CODE_DEFINITION.ADA_CODE," _
''                & " TX_ADA_CODE_DEFINITION.PERIO," _
''                & " TX_ADA_CODE_DEFINITION.FPD_FLAG," _
''                & " TX_ADA_CODE_DEFINITION.RPD_FLAG," _
''                & " TX_ADA_CODE_DEFINITION.DENTURE_FLAG,"
''            sSQL = sSQL & " TX_ADA_CODE_DEFINITION.DWV" _
''                & " FROM DENTAL_ENCOUNTER_TX, TX_ADA_CODE_DEFINITION," _
''                & " DENTAL_ENCOUNTER_GROUPS," _
''                & " DENTAL_ENCOUNTER"
''            sSQL = sSQL & " Where DENTAL_ENCOUNTER_TX.TX_ADA_CODE_ID = TX_ADA_CODE_DEFINITION.TX_ADA_CODE_ID" _
''                & " AND DENTAL_ENCOUNTER_TX.TX_CODE_ID = DENTAL_ENCOUNTER_GROUPS.TX_CODE_ID" _
''                & " AND DENTAL_ENCOUNTER_TX.DENTAL_ENCOUNTER_ID = DENTAL_ENCOUNTER.DENTAL_ENCOUNTER_ID" _
''                & " AND ((DENTAL_ENCOUNTER_TX.CHARTABLE = - 1) AND" _
''                & " (DENTAL_ENCOUNTER_TX.STATUS_ID <> 104) AND" _
''                & " (DENTAL_ENCOUNTER_TX.SIGNED = 0) AND" _
''                & " (DENTAL_ENCOUNTER.PATIENT_UNIT_ID = " & PatientID & ") AND" _
''                & " (DENTAL_ENCOUNTER_GROUPS.GROUP_ID = " & GroupID & ") OR" _
''                & " (DENTAL_ENCOUNTER_TX.CHARTABLE = - 1) AND" _
''                & " (DENTAL_ENCOUNTER_TX.STATUS_ID IN (100, 101, 103)) AND" _
''                & " (DENTAL_ENCOUNTER_TX.SIGNED = - 1) AND" _
''                & " (DENTAL_ENCOUNTER.PATIENT_UNIT_ID = " & PatientID & ") AND" _
''                & " (DENTAL_ENCOUNTER_GROUPS.GROUP_ID = " & GroupID & "))" _
''                & " ORDER BY DENTAL_ENCOUNTER_TX.CREATE_DATE_TIME"
''
''        Else
''            'Filter on the toothid
''            sSQL = "SELECT DENTAL_ENCOUNTER_TX.TX_CODE_ID," _
''                & " DENTAL_ENCOUNTER_TX.DENTAL_ENCOUNTER_ID," _
''                & " DENTAL_ENCOUNTER_TX.TX_ADA_CODE_ID," _
''                & " DENTAL_ENCOUNTER_TX.STATUS_ID," _
''                & " DENTAL_ENCOUNTER_TX.CREATE_DATE_TIME," _
''                & " DENTAL_ENCOUNTER_TX.MODIFIED_DATE_TIME," _
''                & " DENTAL_ENCOUNTER_TX.PROVIDER_ID," _
''                & " DENTAL_ENCOUNTER_TX.CHARTABLE," _
''                & " DENTAL_ENCOUNTER_TX.COMMENTS," _
''                & " DENTAL_ENCOUNTER_TX.MATERIALS," _
''                & " DENTAL_ENCOUNTER_TX.PROGNOSIS_ID," _
''                & " DENTAL_ENCOUNTER_TX.SIGNED," _
''                & " DENTAL_ENCOUNTER_TX.PHASE," _
''                & " DENTAL_ENCOUNTER_TX.WORKLOAD_ID," _
''                & " DENTAL_ENCOUNTER_TX.DELETE_REASON,"
''            sSQL = sSQL & " TX_ADA_CODE_DEFINITION.ADA_CATEGORY_CODE," _
''                & " TX_ADA_CODE_DEFINITION.ADA_DEFINITION," _
''                & " TX_ADA_CODE_DEFINITION.ADA_ALIAS," _
''                & " TX_ADA_CODE_DEFINITION.CHART_RELATED," _
''                & " TX_ADA_CODE_DEFINITION.LABFLAG," _
''                & " TX_ADA_CODE_DEFINITION.ADA_HELP_NARATIVE," _
''                & " TX_ADA_CODE_DEFINITION.ADA_CODE," _
''                & " TX_ADA_CODE_DEFINITION.PERIO," _
''                & " TX_ADA_CODE_DEFINITION.FPD_FLAG," _
''                & " TX_ADA_CODE_DEFINITION.RPD_FLAG," _
''                & " TX_ADA_CODE_DEFINITION.DENTURE_FLAG,"
''            sSQL = sSQL & " TX_ADA_CODE_DEFINITION.DWV," _
''                & " DENTAL_ENCOUNTER_TEETH.TOOTH_ID" _
''                & " FROM DENTAL_ENCOUNTER_TX, TX_ADA_CODE_DEFINITION," _
''                & " DENTAL_ENCOUNTER_TXS_SS," _
''                & " DENTAL_ENC_SURFACE_SET, DENTAL_ENCOUNTER_TEETH," _
''                & " DENTAL_ENCOUNTER_GROUPS," _
''                & " DENTAL_ENCOUNTER"
''            sSQL = sSQL & " Where DENTAL_ENCOUNTER_TX.TX_ADA_CODE_ID = TX_ADA_CODE_DEFINITION.TX_ADA_CODE_ID" _
''                & " AND DENTAL_ENCOUNTER_TX.TX_CODE_ID = DENTAL_ENCOUNTER_TXS_SS.TX_CODE_ID" _
''                & " AND DENTAL_ENCOUNTER_TXS_SS.SURFACE_SET_ID = DENTAL_ENC_SURFACE_SET.SURFACE_SET_ID" _
''                & " AND DENTAL_ENC_SURFACE_SET.DENTAL_ENCOUNTER_TEETH_ID" _
''                & " = DENTAL_ENCOUNTER_TEETH.DENTAL_ENCOUNTER_TEETH_ID" _
''                & " AND DENTAL_ENCOUNTER_TX.TX_CODE_ID = DENTAL_ENCOUNTER_GROUPS.TX_CODE_ID" _
''                & " AND DENTAL_ENCOUNTER_TX.DENTAL_ENCOUNTER_ID = DENTAL_ENCOUNTER.DENTAL_ENCOUNTER_ID" _
''                & " AND ((DENTAL_ENCOUNTER_TX.STATUS_ID <> 104) AND" _
''                & " (DENTAL_ENCOUNTER_TX.CHARTABLE = - 1) AND" _
''                & " (DENTAL_ENCOUNTER_TEETH.TOOTH_ID = " & mvarToothID & ") AND" _
''                & " (DENTAL_ENCOUNTER_TX.SIGNED = 0) AND" _
''                & " (DENTAL_ENCOUNTER_GROUPS.GROUP_ID = " & GroupID & ") AND" _
''                & " (DENTAL_ENCOUNTER.PATIENT_UNIT_ID = " & PatientID & ") OR" _
''                & " (DENTAL_ENCOUNTER_TX.STATUS_ID IN (100, 101, 103)) AND" _
''                & " (DENTAL_ENCOUNTER_TX.CHARTABLE = - 1) AND" _
''                & " (DENTAL_ENCOUNTER_TEETH.TOOTH_ID = " & mvarToothID & ") AND" _
''                & " (DENTAL_ENCOUNTER_TX.SIGNED = - 1) AND" _
''                & " (DENTAL_ENCOUNTER_GROUPS.GROUP_ID = " & GroupID & ") AND" _
''                & " (DENTAL_ENCOUNTER.PATIENT_UNIT_ID = " & PatientID & "))" _
''                & " ORDER BY DENTAL_ENCOUNTER_TX.CREATE_DATE_TIME"
''
''        End If
''
''    ElseIf Source = sNotes Then
'        sSQL = "SELECT DENTAL_ENCOUNTER_TX.TX_CODE_ID," _
'            & " DENTAL_ENCOUNTER_TX.DENTAL_ENCOUNTER_ID," _
'            & " DENTAL_ENCOUNTER_TX.TX_ADA_CODE_ID," _
'            & " DENTAL_ENCOUNTER_TX.STATUS_ID," _
'            & " DENTAL_ENCOUNTER_TX.CREATE_DATE_TIME," _
'            & " DENTAL_ENCOUNTER_TX.MODIFIED_DATE_TIME," _
'            & " DENTAL_ENCOUNTER_TX.PROVIDER_ID," _
'            & " DENTAL_ENCOUNTER_TX.CHARTABLE," _
'            & " DENTAL_ENCOUNTER_TX.COMMENTS," _
'            & " DENTAL_ENCOUNTER_TX.MATERIALS," _
'            & " DENTAL_ENCOUNTER_TX.PROGNOSIS_ID," _
'            & " DENTAL_ENCOUNTER_TX.SIGNED," _
'            & " DENTAL_ENCOUNTER_TX.PHASE," _
'            & " DENTAL_ENCOUNTER_TX.WORKLOAD_ID," _
'            & " DENTAL_ENCOUNTER_TX.DELETE_REASON," _
'            & " DENTAL_ENCOUNTER_TX.MODIFIED_IN_DENTAL_ENCOUNTER,"
'        sSQL = sSQL & " TX_ADA_CODE_DEFINITION.ADA_CATEGORY_CODE," _
'            & " TX_ADA_CODE_DEFINITION.ADA_DEFINITION," _
'            & " TX_ADA_CODE_DEFINITION.ADA_ALIAS," _
'            & " TX_ADA_CODE_DEFINITION.CHART_RELATED," _
'            & " TX_ADA_CODE_DEFINITION.LABFLAG," _
'            & " TX_ADA_CODE_DEFINITION.ADA_HELP_NARATIVE," _
'            & " TX_ADA_CODE_DEFINITION.ADA_CODE," _
'            & " TX_ADA_CODE_DEFINITION.PERIO," _
'            & " TX_ADA_CODE_DEFINITION.FPD_FLAG," _
'            & " TX_ADA_CODE_DEFINITION.RPD_FLAG," _
'            & " TX_ADA_CODE_DEFINITION.DENTURE_FLAG,"
'        sSQL = sSQL & " TX_ADA_CODE_DEFINITION.DWV" _
'            & " FROM DENTAL_ENCOUNTER_TX, TX_ADA_CODE_DEFINITION," _
'            & " DENTAL_ENCOUNTER," _
'            & " DENTAL_ENCOUNTER_GROUPS" _
'            & " Where DENTAL_ENCOUNTER_TX.TX_ADA_CODE_ID = TX_ADA_CODE_DEFINITION.TX_ADA_CODE_ID" _
'            & " AND"
'        sSQL = sSQL & " DENTAL_ENCOUNTER_TX.DENTAL_ENCOUNTER_ID = DENTAL_ENCOUNTER.DENTAL_ENCOUNTER_ID" _
'            & " AND" _
'            & " DENTAL_ENCOUNTER_TX.TX_CODE_ID = DENTAL_ENCOUNTER_GROUPS.TX_CODE_ID" _
'            & " AND (((DENTAL_ENCOUNTER_TX.STATUS_ID <> 104) AND" _
'            & " (DENTAL_ENCOUNTER_TX.SIGNED = 0) AND" _
'            & " (DENTAL_ENCOUNTER.PATIENT_UNIT_ID = " & PatientID & ") AND" _
'            & " (DENTAL_ENCOUNTER_GROUPS.GROUP_ID = " & GroupID & ")) OR"
'
'        If gbDentalTriage = False Then
'            sSQL = sSQL & " ((DENTAL_ENCOUNTER_TX.STATUS_ID IN (100, 101, 102, 103)) AND"
'        Else 'If the dentaltriage command line switch is on, then query all statuses except deleted
'            sSQL = sSQL & " ((DENTAL_ENCOUNTER_TX.STATUS_ID <> 104) AND"
'        End If
'
'        sSQL = sSQL & " (DENTAL_ENCOUNTER.PATIENT_UNIT_ID = " & PatientID & ") AND" _
'            & " (DENTAL_ENCOUNTER_TX.SIGNED = -1) AND" _
'            & " (DENTAL_ENCOUNTER_GROUPS.GROUP_ID = " & GroupID & ")))" _
'            & " ORDER BY DENTAL_ENCOUNTER_TX.CREATE_DATE_TIME"
'
''    End If
'
'    RefreshTxsFromPatientID = RefreshTxsBySQL(sSQL, Source)
    
End Function


Friend Function RefreshTxsFromRS(TxsRs As Recordset, GroupID As Long, CurrentDentalEncounterID As Long) As Boolean

Dim oTempTx As clsTx

    mvarCurrentDentalEncounterID = CurrentDentalEncounterID
    mvarGroupID = GroupID
    
    TxsRs.Filter = ""
    If Not TxsRs.EOF Then
        TxsRs.Filter = "Group_ID = " & GroupID
        
        'Release the previous collection
        Set mCol = New Collection
        
        'Add the Tx objects to the Txs collection
        Do Until TxsRs.EOF
            Set oTempTx = AddInternal(TxsRs("DENTAL_ENCOUNTER_ID"), TxsRs("TX_ADA_CODE_ID"), TxsRs("TX_CODE_ID"))
            
            'oTempTx will be nothing if the object already exists, so the add method added the reference
            If Not oTempTx Is Nothing Then
                'Add the additional fields
                oTempTx.StatusID = Val(TxsRs("STATUS_ID") & "")
                If Not IsNull(TxsRs("CREATE_DATE_TIME")) Then
                    oTempTx.CreateDateTime = cwobjCWComm.GmtToLocal(TxsRs("CREATE_DATE_TIME"))
                End If
                oTempTx.ProviderID = Val(TxsRs("PROVIDER_ID") & "")
                oTempTx.Chartable = CBool(Val(TxsRs("CHARTABLE") & ""))
                oTempTx.Comments = TxsRs("COMMENTS") & ""
                oTempTx.Materials = TxsRs("MATERIALS") & ""
                oTempTx.PrognosisID = Val(TxsRs("PROGNOSIS_ID") & "")
                oTempTx.Signed = CBool(Val(TxsRs("SIGNED") & ""))
                oTempTx.IsPerio = CBool(Val(TxsRs("PERIO") & ""))
                oTempTx.IsFPD = CBool(Val(TxsRs("FPD_FLAG") & ""))
                oTempTx.IsRPD = CBool(Val(TxsRs("RPD_FLAG") & ""))
                oTempTx.IsCompleteDenture = CBool(Val(TxsRs("DENTURE_FLAG") & ""))
                oTempTx.ADACategoryCode = TxsRs("ADA_CATEGORY_CODE") & ""
                oTempTx.ADACode = TxsRs("ADA_CODE") & ""
                oTempTx.ADADefinition = TxsRs("ADA_DEFINITION") & ""
                oTempTx.ADAAlias = TxsRs("ADA_ALIAS") & ""
                oTempTx.ADAHelpNarative = TxsRs("ADA_HELP_NARATIVE") & ""
                oTempTx.ChartRelated = Val(TxsRs("CHART_RELATED") & "")
                oTempTx.LabFlag = Val(TxsRs("LABFLAG") & "")
                oTempTx.DWV = Val(TxsRs("DWV") & "")
                oTempTx.PhaseID = Val(TxsRs("PHASE") & "")
                oTempTx.WorkloadID = Val(TxsRs("WORKLOAD_ID") & "")
                oTempTx.DeleteReason = Val(TxsRs("DELETE_REASON") & "")
                If Not IsNull(TxsRs("MODIFIED_DATE_TIME")) Then
                    oTempTx.ModifiedDateTime = cwobjCWComm.GmtToLocal(TxsRs("MODIFIED_DATE_TIME"))
                End If
                'The Tx is clean
                'This is here because it is updated in SetDirtyFlag
                oTempTx.ModifiedInDentalEncounterID = Val(TxsRs("MODIFIED_IN_DENTAL_ENCOUNTER"))
                oTempTx.DirtyFlag = dfNoAction
            End If
            
            TxsRs.MoveNext
        Loop
        
        Set oTempTx = Nothing
    End If
    
    RefreshTxsFromRS = True
    
End Function

Friend Function RefreshLocSubLocFromRS(TxsLocSubLocRS As Recordset) As Boolean
    
    'Load the Loc and Sub Loc for the txs

Dim oTx As clsTx

    For Each oTx In mCol
        oTx.Locations.RefreshLocationsFromTxCodeIDRS oTx.TxCodeID, TxsLocSubLocRS
    Next
    
End Function

Friend Function RefreshSSandTeethFromRS(TxsSSTeethLocRS As Recordset) As Boolean
    
    'Load the SS and Teeth for the txs

Dim oTx As clsTx

    For Each oTx In mCol
        oTx.SurfaceSets.RefreshSurfaceSetsFromRS 0, oTx.TxCodeID, TxsSSTeethLocRS
    Next
    
End Function

Private Function RefreshTxsBySQL(SQL As String, Source As SourceEnum) As Boolean

'************* NOT USED ANYMORE - see RefreshTxsFromRS

'Dim rs As Recordset
'Dim oTempTx As clsTx
'
'    Set rs = New Recordset
'    Set rs = goconn.OpenRecordset(SQL)
'
'    'Release the previous collection
'    Set mCol = New Collection
'
'    'Add the Tx objects to the Txs collection
'    Do Until rs.EOF
'        Set oTempTx = AddInternal(rs("DENTAL_ENCOUNTER_ID"), Source, rs("TX_ADA_CODE_ID"), rs("TX_CODE_ID"))
'
'        'oTempTx will be nothing if the object already exists, so the add method added the reference
'        If Not oTempTx Is Nothing Then
'            'Add the additional fields
'            oTempTx.StatusID = Val(rs("STATUS_ID") & "")
'            oTempTx.CreateDateTime = rs("CREATE_DATE_TIME") & ""
'            oTempTx.ProviderID = Val(rs("PROVIDER_ID") & "")
'            oTempTx.Chartable = CBool(Val(rs("CHARTABLE") & ""))
'            oTempTx.Comments = rs("COMMENTS") & ""
'            oTempTx.Materials = rs("MATERIALS") & ""
'            oTempTx.PrognosisID = Val(rs("PROGNOSIS_ID") & "")
'            oTempTx.Signed = CBool(Val(rs("SIGNED") & ""))
'            oTempTx.IsPerio = CBool(Val(rs("PERIO") & ""))
'            oTempTx.IsFPD = CBool(Val(rs("FPD_FLAG") & ""))
'            oTempTx.IsRPD = CBool(Val(rs("RPD_FLAG") & ""))
'            oTempTx.IsCompleteDenture = CBool(Val(rs("DENTURE_FLAG") & ""))
'            oTempTx.ADACategoryCode = rs("ADA_CATEGORY_CODE") & ""
'            oTempTx.ADACode = rs("ADA_CODE") & ""
'            oTempTx.ADADefinition = rs("ADA_DEFINITION") & ""
'            oTempTx.ADAAlias = rs("ADA_ALIAS") & ""
'            oTempTx.ADAHelpNarative = rs("ADA_HELP_NARATIVE") & ""
'            oTempTx.ChartRelated = Val(rs("CHART_RELATED") & "")
'            oTempTx.LabFlag = Val(rs("LABFLAG") & "")
'            oTempTx.DWV = Val(rs("DWV") & "")
'            oTempTx.PhaseID = Val(rs("PHASE") & "")
'            oTempTx.WorkloadID = Val(rs("WORKLOAD_ID") & "")
'            oTempTx.DeleteReason = Val(rs("DELETE_REASON") & "")
'            If Not IsNull(rs("MODIFIED_DATE_TIME")) Then
'                oTempTx.ModifiedDateTime = rs("MODIFIED_DATE_TIME") & ""
'            End If
'            'The Tx is clean
'            'This is here because it is updated in SetDirtyFlag
'            oTempTx.ModifiedInDentalEncounterID = Val(rs("MODIFIED_IN_DENTAL_ENCOUNTER"))
'            oTempTx.DirtyFlag = dfNoAction
'        End If
'
'        rs.MoveNext
'    Loop
'
'    Set oTempTx = Nothing
'    rs.Close
'    Set rs = Nothing
'
'    RefreshTxsBySQL = True
    
End Function


Friend Function SaveTxs(GroupID As Long) As Boolean

'Save each Dx

Dim oTx As clsTx
Dim lCounter As Long

    SaveTxs = True
    
    'Use for-next (not for-each) because I need the index position
    'Must count backwards to solve the problem with deleting a tx and all the indexes changing.
    For lCounter = mCol.Count To 1 Step -1
        Set oTx = mCol(lCounter)
        
        If oTx.SaveTx(GroupID) = False Then
            'If one Tx doesn't save properly, then set SaveTxs = false
            SaveTxs = False
        End If
        'Remove all the Txs that were not saved because they were deleted in this encounter
        If oTx.StatusID = txdeleted And oTx.TxCodeID = 0 Then
            mCol.Remove lCounter
        End If
    Next
    
    Set oTx = Nothing
    
    
End Function

Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)

    If TypeName(vntIndexKey) = "String" Then
        'The Dx was saved, so set the statusid to Deleted
        mCol(vntIndexKey).StatusID = txdeleted
    Else
        'The Dx was never saved so remove it from the collection by index
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
    Set mvarTxsRef = Nothing
    Set mvarDxsRef = Nothing

End Sub

