VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colSubLocations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsSubLocation"
Attribute VB_Ext_KEY = "Member0" ,"clsSubLocation"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable to hold collection
Private mCol As Collection



Friend Function SaveSubLocations(DxLocID As Long, TxLocID As Long) As Boolean

'Save each Location
'DxCodeID or TxCodeID is the ID that the Location belongs to.

Dim oSubLocation As clsSubLocation

    SaveSubLocations = True
    
    For Each oSubLocation In mCol
        If oSubLocation.SaveSubLocation(DxLocID, TxLocID) = False Then
           'If one Location doesn't save properly, then set SaveLocations = false
            SaveSubLocations = False
        End If
    Next
    
End Function


Public Function Add(SubLocationID As Long, Optional pkid As Long = 0) As clsSubLocation
    'create a new object
    Dim objNewMember As clsSubLocation
    Set objNewMember = New clsSubLocation

    'set the properties passed into the method
    objNewMember.SubLocationID = SubLocationID
    
    
    If pkid = 0 Then
        objNewMember.DirtyFlag = dfAdd
        mCol.Add objNewMember
    Else
        objNewMember.DirtyFlag = dfNoAction
        mCol.Add objNewMember, CStr(pkid)
    End If

    objNewMember.SubLocationUpdatedInEncounter = False
    
    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Friend Function RefreshSubLocationsFromDxLocID(DxLocID As Long) As Boolean

''This function populates the Locations collection given the DxCodeID and the Source
'
'Dim rs As Recordset
'Dim oTempSubLocation As clsSubLocation
'Dim sSQL As String
'
'    sSQL = "SELECT NOTE_SUB_LOCATION.SUB_LOCATION_NAME," _
'        & " NOTE_DX_SUB_LOCATION.DX_SUBLOC_ID," _
'        & " NOTE_DX_SUB_LOCATION.DELETED," _
'        & " NOTE_DX_SUB_LOCATION.DX_LOCATION_ID," _
'        & " NOTE_DX_SUB_LOCATION.SUB_LOCATION_ID" _
'        & " From NOTE_SUB_LOCATION, NOTE_DX_SUB_LOCATION" _
'        & " Where NOTE_SUB_LOCATION.SUB_LOCATION_ID = NOTE_DX_SUB_LOCATION.SUB_LOCATION_ID" _
'        & " AND (NOTE_DX_SUB_LOCATION.DX_LOCATION_ID = " & DxLocID & ")"
'
'    Set rs = New Recordset
'
'    Set rs = goconn.OpenRecordset(sSQL)
'
'    'Release the previous collection
'    Set mCol = New Collection
'
'    'Add the Location objects to the Locations collection
'    Do Until rs.EOF
'        Set oTempSubLocation = Add(rs("SUB_LOCATION_ID"), rs("DX_SUBLOC_ID"))
'
'        'The Location is clean
'        oTempSubLocation.DxSubLocID = rs("DX_SUBLOC_ID")
'        oTempSubLocation.DirtyFlag = dfNoAction
'
'        rs.MoveNext
'    Loop
'
'    Set oTempSubLocation = Nothing
'    rs.Close
'    Set rs = Nothing
    
End Function

Friend Function RefreshSubLocationsFromDxLocIDRS(DxLocID As Long, DxLocSubLoc As Recordset) As Boolean

'This function populates the Locations collection given the TXCodeID and the Source

Dim oTempSubLocation As clsSubLocation

    DxLocSubLoc.Filter = ""
    
    If Not DxLocSubLoc.EOF Then
        DxLocSubLoc.Filter = "DX_LOCATION_ID=" & DxLocID
    End If
    
    'Release the previous collection
    Set mCol = New Collection
    
    'Add the SubLocation objects to the SubLocations collection
    Do Until DxLocSubLoc.EOF
        
        If Val(DxLocSubLoc("SUB_LOCATION_ID") & "") <> 0 Then
            Set oTempSubLocation = Add(DxLocSubLoc("SUB_LOCATION_ID"), DxLocSubLoc("DX_SUBLOC_ID"))
            
            'The Location is clean
            oTempSubLocation.DxSubLocID = DxLocSubLoc("DX_SUBLOC_ID")
            oTempSubLocation.DirtyFlag = dfNoAction
            oTempSubLocation.SubLocationUpdatedInEncounter = False
        End If
        
        DxLocSubLoc.MoveNext
    Loop
    
    Set oTempSubLocation = Nothing

End Function

Friend Function RefreshSubLocationsFromTxLocID(TxLocID As Long) As Boolean

''This function populates the Locations collection given the TXCodeID and the Source
'
'Dim rs As Recordset
'Dim oTempSubLocation As clsSubLocation
'Dim sSQL As String
'
'    sSQL = "SELECT NOTE_SUB_LOCATION.SUB_LOCATION_NAME," _
'        & " NOTE_TX_SUB_LOCATION.TX_SUBLOC_ID," _
'        & " NOTE_TX_SUB_LOCATION.DELETED," _
'        & " NOTE_TX_SUB_LOCATION.TX_LOCATION_ID," _
'        & " NOTE_TX_SUB_LOCATION.SUB_LOCATION_ID" _
'        & " From NOTE_SUB_LOCATION, NOTE_TX_SUB_LOCATION" _
'        & " Where NOTE_SUB_LOCATION.SUB_LOCATION_ID = NOTE_TX_SUB_LOCATION.SUB_LOCATION_ID" _
'        & " AND (NOTE_TX_SUB_LOCATION.TX_LOCATION_ID = " & TxLocID & ")"
'
'    Set rs = New Recordset
'    Set rs = goconn.OpenRecordset(sSQL)
'
'    'Release the previous collection
'    Set mCol = New Collection
'
'    'Add the Location objects to the Locations collection
'    Do Until rs.EOF
'        Set oTempSubLocation = Add(rs("SUB_LOCATION_ID"), rs("TX_SUBLOC_ID"))
'
'        'The Location is clean
'        oTempSubLocation.Deleted = rs("Deleted")
'        oTempSubLocation.TxSubLocID = rs("TX_SUBLOC_ID")
'        oTempSubLocation.DirtyFlag = dfNoAction
'
'        rs.MoveNext
'    Loop
'
'    Set oTempSubLocation = Nothing
'    rs.Close
'    Set rs = Nothing
    
End Function


Friend Function RefreshSubLocationsFromTxLocIDRS(TxLocID As Long, TxLocSubLoc As Recordset) As Boolean

'This function populates the Locations collection given the TXCodeID and the Source

Dim oTempSubLocation As clsSubLocation

    TxLocSubLoc.Filter = ""
    
    If Not TxLocSubLoc.EOF Then
        TxLocSubLoc.Filter = "TX_LOCATION_ID=" & TxLocID
    End If

    'Release the previous collection
    Set mCol = New Collection
    
    'Add the Location objects to the Locations collection
    Do Until TxLocSubLoc.EOF
    
        If Val(TxLocSubLoc("SUB_LOCATION_ID") & "") <> 0 Then
            Set oTempSubLocation = Add(TxLocSubLoc("SUB_LOCATION_ID"), TxLocSubLoc("TX_SUBLOC_ID"))
        
            'The Location is clean
            oTempSubLocation.Deleted = TxLocSubLoc("Deleted")
            oTempSubLocation.TxSubLocID = TxLocSubLoc("TX_SUBLOC_ID")
            oTempSubLocation.DirtyFlag = dfNoAction
            oTempSubLocation.SubLocationUpdatedInEncounter = False
        End If
        
        TxLocSubLoc.MoveNext
    Loop
    
    Set oTempSubLocation = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As clsSubLocation
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

    If mCol(vntIndexKey).SubLocationID = 0 Then
        'The sublocation was never saved so remove it from the collection
        mCol.Remove vntIndexKey
    Else
        'The sublocation was saved, so set the deleted property to deleted
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


Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub

