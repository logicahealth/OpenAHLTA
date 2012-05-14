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
Attribute VB_Name = "colLocations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsLocation"
Attribute VB_Ext_KEY = "Member0" ,"clsLocation"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable to hold collection
Private mCol As Collection


Public Function Add(LocationID As Long, Optional pkid As Long = 0) As clsLocation
    'create a new object
    Dim objNewMember As clsLocation
    Set objNewMember = New clsLocation

    'set the properties passed into the method
    objNewMember.LocationID = LocationID
    
    If pkid = 0 Then
        'Set the default
        objNewMember.Deleted = False
        
        objNewMember.DirtyFlag = dfAdd
        mCol.Add objNewMember
    Else
        objNewMember.DirtyFlag = dfNoAction
        mCol.Add objNewMember, CStr(pkid)
    End If

    objNewMember.LocationUpdatedInEncounter = False
    
    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As clsLocation
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


Friend Function RefreshLocationsFromDxCodeID(DxCodeID As Long, Source As SourceEnum) As Boolean

'This function populates the Locations collection given the DxCodeID and the Source

Dim rs As Recordset
Dim oTempLocation As clsLocation
Dim sSQL As String

    If Source = sCharts Then
        'There will never be locations for a chartable dx
        Exit Function
    End If
    
    'No need to carry the source below this object into SubLocations
    
    sSQL = "SELECT NOTE_LOCATION.LOCATION_ID," _
        & " NOTE_LOCATION.LOCATION_NAME," _
        & " NOTE_DXS_LOCATION.DELETED," _
        & " NOTE_DXS_LOCATION.DX_LOCATION_ID," _
        & " NOTE_DXS_LOCATION.DX_CODE_ID" _
        & " From NOTE_DXS_LOCATION, NOTE_LOCATION" _
        & " Where NOTE_DXS_LOCATION.LOCATION_ID = NOTE_LOCATION.LOCATION_ID" _
        & " AND (NOTE_DXS_LOCATION.DELETED = 0) AND" _
        & " (NOTE_DXS_LOCATION.DX_CODE_ID = " & DxCodeID & ")"
    
    Set rs = New Recordset
    
    Set rs = goconn.OpenRecordset(sSQL)
    
    'Release the previous collection
    Set mCol = New Collection
    
    'Add the Location objects to the Locations collection
    Do Until rs.EOF
        Set oTempLocation = Add(rs("LOCATION_ID"), rs("DX_LOCATION_ID"))
        
        'The Location is clean
        oTempLocation.DxLocID = rs("DX_LOCATION_ID")
        oTempLocation.Deleted = rs("Deleted")
        oTempLocation.DirtyFlag = dfNoAction
        oTempLocation.LocationUpdatedInEncounter = False
        rs.MoveNext
    Loop
    
    Set oTempLocation = Nothing
    rs.Close
    Set rs = Nothing
    
End Function
Friend Function RefreshLocationsFromDxCodeIDRS(DxCodeID As Long, TxCodeID As Long, rsLocSubLoc As Recordset) As Boolean

'This function populates the Locations collection given the DxCodeID and the Source

'Dim rs As Recordset
Dim oTempLocation As clsLocation
Dim sSQL As String

    rsLocSubLoc.Filter = ""
    
    If Not rsLocSubLoc.EOF Then
        If DxCodeID <> 0 Then
            rsLocSubLoc.Filter = "DX_CODE_ID=" & DxCodeID
        ElseIf TxCodeID <> 0 Then
            rsLocSubLoc.Filter = "Tx_code_id=" & TxCodeID
        End If
    End If
    
    
    'Release the previous collection
    Set mCol = New Collection
    
    'Add the Location objects to the Locations collection
    Do Until rsLocSubLoc.EOF
        Set oTempLocation = Add(rsLocSubLoc("LOCATION_ID"), rsLocSubLoc("DX_LOCATION_ID"))
        
        'The Location is clean
        oTempLocation.DxLocID = rsLocSubLoc("DX_LOCATION_ID")
        oTempLocation.Deleted = rsLocSubLoc("Deleted")
        
        oTempLocation.DirtyFlag = dfNoAction
        oTempLocation.LocationUpdatedInEncounter = False
        
        rsLocSubLoc.MoveNext
    Loop
    
    For Each oTempLocation In mCol
        'Load the subloc
        oTempLocation.SubLocations.RefreshSubLocationsFromDxLocIDRS oTempLocation.DxLocID, rsLocSubLoc
    Next
    
    Set oTempLocation = Nothing

End Function
Friend Function RefreshLocationsFromTxCodeIDRS(TxCodeID As Long, rsLocSubLoc As Recordset) As Boolean

'This function populates the Locations collection given the DxCodeID and the Source

'Dim rs As Recordset
Dim oTempLocation As clsLocation
Dim sSQL As String

    rsLocSubLoc.Filter = ""
    
    If Not rsLocSubLoc.EOF Then
            rsLocSubLoc.Filter = "TX_CODE_ID=" & TxCodeID
    End If
    
    'Release the previous collection
    Set mCol = New Collection
    
    'Add the Location objects to the Locations collection
    Do Until rsLocSubLoc.EOF
        Set oTempLocation = Add(rsLocSubLoc("LOCATION_ID"), rsLocSubLoc("TX_LOCATION_ID"))
        
        'The Location is clean
        oTempLocation.TxLocID = rsLocSubLoc("TX_LOCATION_ID")
        oTempLocation.Deleted = rsLocSubLoc("Deleted")
        
        
        oTempLocation.DirtyFlag = dfNoAction
        oTempLocation.LocationUpdatedInEncounter = False
        
        rsLocSubLoc.MoveNext
    Loop
            
    'Load the SubLoc, this must be outside of the loop above so it doesn't interfere with loading the locations
    For Each oTempLocation In mCol
        'Load the subloc
        oTempLocation.SubLocations.RefreshSubLocationsFromTxLocIDRS oTempLocation.TxLocID, rsLocSubLoc
    Next
    
    Set oTempLocation = Nothing

End Function


Friend Function RefreshLocationsFromTxCodeID(TxCodeID As Long, Source As SourceEnum) As Boolean

'This function populates the Locations collection given the TxCodeID and the Source

Dim rs As Recordset
Dim oTempLocation As clsLocation
Dim sSQL As String

    If Source = sCharts Then
        'There will never be locations for a chartable dx
        Exit Function
    End If
    
    'No need to carry the source below this object into SubLocations
    
    sSQL = "SELECT NOTE_LOCATION.LOCATION_NAME," _
        & " NOTE_TXS_LOCATION.DELETED," _
        & " NOTE_TXS_LOCATION.LOCATION_ID," _
        & " NOTE_TXS_LOCATION.TX_LOCATION_ID," _
        & " NOTE_TXS_LOCATION.TX_CODE_ID" _
        & " From NOTE_TXS_LOCATION, NOTE_LOCATION" _
        & " Where NOTE_TXS_LOCATION.LOCATION_ID = NOTE_LOCATION.LOCATION_ID" _
        & " AND (NOTE_TXS_LOCATION.DELETED = 0) AND" _
        & " (NOTE_TXS_LOCATION.TX_CODE_ID = " & TxCodeID & ")"
    
    Set rs = New Recordset
    Set rs = goconn.OpenRecordset(sSQL)
    
    'Release the previous collection
    Set mCol = New Collection
    
    'Add the Location objects to the Locations collection
    Do Until rs.EOF
        Set oTempLocation = Add(rs("LOCATION_ID"), rs("TX_LOCATION_ID"))
        
        'The Location is clean
        oTempLocation.TxLocID = rs("TX_LOCATION_ID")
        oTempLocation.Deleted = rs("Deleted")
        oTempLocation.DirtyFlag = dfNoAction
        oTempLocation.LocationUpdatedInEncounter = False
        
        rs.MoveNext
    Loop
    
    Set oTempLocation = Nothing
    rs.Close
    Set rs = Nothing
    
End Function

Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)

    If mCol(vntIndexKey).LocationID = 0 Then
        'The location was never saved so remove it from the collection
        mCol.Remove vntIndexKey
    Else
        'The location was saved, so set the deleted property to deleted
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


Friend Function SaveLocations(DxCodeID As Long, TxCodeID As Long) As Boolean

'Save each Location
'DxCodeID or TxCodeID is the ID that the Location belongs to.

Dim oLocation As clsLocation

    SaveLocations = True
    
    For Each oLocation In mCol
        If oLocation.SaveLocation(DxCodeID, TxCodeID) = False Then
           'If one Location doesn't save properly, then set SaveLocations = false
            SaveLocations = False
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

