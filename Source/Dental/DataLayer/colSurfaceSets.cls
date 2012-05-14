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
Attribute VB_Name = "colSurfaceSets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsSurfaceSet"
Attribute VB_Ext_KEY = "Member0" ,"clsSurfaceSet"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable to hold collection
Private mCol As Collection
Friend Function RefreshSurfaceSetsFromDxORTxCodeID(DxCodeID As Long, TxCodeID As Long, Source As SourceEnum) As Boolean

'This function populates the surfacesets collection given the DxCodeID

Dim rs As New Recordset, sSQL As String
Dim oTempSS As clsSurfaceSet
    
    If Source = sNotes Then
        'A non-chartable item will never have a surface set.
        Exit Function
    End If
    
    If DxCodeID <> 0 Then
        
        sSQL = "SELECT DENTAL_ENC_SURFACE_SET.SURFACE_SET_ID," _
            & " DENTAL_ENC_SURFACE_SET.SURFACESET," _
            & " DENTAL_ENC_SURFACE_SET.DENTAL_ENCOUNTER_TEETH_ID," _
            & " DENTAL_ENC_SURFACE_SET.COLOR_ID," _
            & " DENTAL_ENC_SURFACE_SET.STATE," _
            & " DENTAL_ENC_SURFACE_SET.HATCH_ID," _
            & " DENTAL_ENC_SURFACE_SET.SYMBOL_ID," _
            & " DENTAL_ENC_SURFACE_SET.SUB_MENU_ID," _
            & " DENTAL_ENC_SURFACE_SET.DELETED," _
            & " DENTAL_ENC_SURFACE_SET.SurfaceAbbr" _
            & " FROM DENTAL_ENC_SURFACE_SET," _
            & " DENTAL_ENCOUNTER_DXS_SS" _
            & " Where DENTAL_ENC_SURFACE_SET.SURFACE_SET_ID = DENTAL_ENCOUNTER_DXS_SS.SURFACE_SET_ID" _
            & " AND (DENTAL_ENC_SURFACE_SET.DELETED = 0) AND" _
            & " (DENTAL_ENCOUNTER_DXS_SS.DX_CODE_ID = " & DxCodeID & ")" _
            & " ORDER BY DENTAL_ENC_SURFACE_SET.SURFACE_SET_ID"
    
    ElseIf TxCodeID <> 0 Then
        
        sSQL = "SELECT DENTAL_ENC_SURFACE_SET.SURFACE_SET_ID," _
            & " DENTAL_ENC_SURFACE_SET.SURFACESET," _
            & " DENTAL_ENC_SURFACE_SET.DENTAL_ENCOUNTER_TEETH_ID," _
            & " DENTAL_ENC_SURFACE_SET.COLOR_ID," _
            & " DENTAL_ENC_SURFACE_SET.STATE," _
            & " DENTAL_ENC_SURFACE_SET.HATCH_ID," _
            & " DENTAL_ENC_SURFACE_SET.SYMBOL_ID," _
            & " DENTAL_ENC_SURFACE_SET.SUB_MENU_ID," _
            & " DENTAL_ENC_SURFACE_SET.DELETED," _
            & " DENTAL_ENC_SURFACE_SET.SurfaceAbbr" _
            & " FROM DENTAL_ENC_SURFACE_SET," _
            & " DENTAL_ENCOUNTER_TXS_SS" _
            & " Where DENTAL_ENC_SURFACE_SET.SURFACE_SET_ID = DENTAL_ENCOUNTER_TXS_SS.SURFACE_SET_ID" _
            & " AND (DENTAL_ENC_SURFACE_SET.DELETED = 0) AND" _
            & " (DENTAL_ENCOUNTER_TXS_SS.TX_CODE_ID = " & TxCodeID & ")" _
            & " ORDER BY DENTAL_ENC_SURFACE_SET.SURFACE_SET_ID"
    
    Else
        'Do nothing
        Exit Function
        
    End If
    
    Set rs = goconn.OpenRecordset(sSQL)
    
    'Release the previous collection
    Set mCol = New Collection
    
    
    'Add the SurfaceSet objects to the SurfaceSets collection
    Do Until rs.EOF
        Set oTempSS = Add(rs("DENTAL_ENCOUNTER_TEETH_ID"), rs("SURFACE_SET_ID"))
        
        oTempSS.State = Val(rs("STATE") & "")
        oTempSS.SurfaceSet = rs("SURFACESET") & ""
        oTempSS.ColorID = Val(rs("COLOR_ID") & "")
        oTempSS.HatchID = Val(rs("HATCH_ID") & "")
        oTempSS.SymbolID = Val(rs("SYMBOL_ID") & "")
        oTempSS.SubMenuID = Val(rs("SUB_MENU_ID") & "")
        oTempSS.SurfaceAbbr = rs("SurfaceAbbr") & ""
        oTempSS.Deleted = CBool(rs("Deleted"))
        
        'The surfaceset is clean
        oTempSS.DirtyFlag = dfNoAction
        oTempSS.SSChangedInEncounter = False
        Set oTempSS = Nothing
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing

End Function

Friend Function RefreshSurfaceSetsFromRS(DxCodeID As Long, TxCodeID As Long, rsSSandTooth As Recordset) As Boolean
                                                                            'THIS CAN BE EITHER DX OR TX SURFACE SETS

'This function populates the surfacesets collection given the RS

Dim rs As New Recordset, sSQL As String
Dim oTempSS As clsSurfaceSet
    
    rsSSandTooth.Filter = ""
    
    'Filter on the correct Dx or Tx
    If Not rsSSandTooth.EOF Then
        If DxCodeID <> 0 Then
            rsSSandTooth.Filter = "DX_CODE_ID = " & DxCodeID
        ElseIf TxCodeID <> 0 Then
            'need to test
            rsSSandTooth.Filter = "TX_CODE_ID = " & TxCodeID
        End If
    End If
    
    'Release the previous collection
    Set mCol = New Collection
    
    
    'Add the SurfaceSet objects to the SurfaceSets collection
    Do Until rsSSandTooth.EOF
    
        Set oTempSS = Add(rsSSandTooth("DENTAL_ENCOUNTER_TEETH_ID"), rsSSandTooth("SURFACE_SET_ID"))
        
        oTempSS.State = Val(rsSSandTooth("STATE") & "")
        oTempSS.SurfaceSet = rsSSandTooth("SURFACESET") & ""
        oTempSS.ColorID = Val(rsSSandTooth("COLOR_ID") & "")
        oTempSS.HatchID = Val(rsSSandTooth("HATCH_ID") & "")
        oTempSS.SymbolID = Val(rsSSandTooth("SYMBOL_ID") & "")
        oTempSS.SubMenuID = Val(rsSSandTooth("SUB_MENU_ID") & "")
        oTempSS.SurfaceAbbr = rsSSandTooth("SurfaceAbbr") & ""
        oTempSS.Deleted = False 'Datalayer doesn't load any deleted SS
        
        'The surfaceset is clean
        oTempSS.DirtyFlag = dfNoAction
        oTempSS.SSChangedInEncounter = False
        
        Set oTempSS = Nothing
        rsSSandTooth.MoveNext
    Loop
    
    'This must be loaded outside of the above loop so it won't interfere with the filter in colTeeth
    For Each oTempSS In mCol
        'Load the teeth for the surface set
        oTempSS.Teeth.RefreshTeethFromRS oTempSS.DentalEncounterTeethID, rsSSandTooth
    Next
    
End Function

Public Function Add(Optional DentalEncounterTeethID As Long = 0, Optional SurfaceSetID As Long = 0) As clsSurfaceSet
    'create a new object
    Dim objNewMember As clsSurfaceSet
    Set objNewMember = New clsSurfaceSet

    'Set the primary key
    objNewMember.SurfaceSetID = SurfaceSetID
    objNewMember.DentalEncounterTeethID = DentalEncounterTeethID
    objNewMember.DirtyFlag = dfAdd
        
    If SurfaceSetID = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, CStr(SurfaceSetID)
    End If
    
    objNewMember.SSChangedInEncounter = False
    
    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As clsSurfaceSet
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

    If mCol(vntIndexKey).SurfaceSetID = 0 Then
        'The SS was never saved so remove it from the collection
        mCol.Remove vntIndexKey
    Else
        'The SS was saved, so set the deleted property to deleted
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


Friend Function SaveSurfaceSets(DxCodeID As Long, TxCodeID As Long) As Boolean

'This function loops through the SurfaceSets and calls the SaveSurfaceSet function

Dim oSS As clsSurfaceSet

    SaveSurfaceSets = True
    
    For Each oSS In mCol
        If oSS.SaveSurfaceSet(DxCodeID, TxCodeID) = False Then
            'If one surfaceset doesn't save correctly then set SaveSurfaceSets = false.
            SaveSurfaceSets = False
        End If
    Next

End Function

Private Function AddObj(SurfaceSet As clsSurfaceSet) As Boolean

'THIS WILL NOT BE USED ANYMORE
'This function adds the passed in Tx to this collection

    AddObj = False
    
    If SurfaceSet.SurfaceSetID = 0 Then
        'The Tx does not have a primary key yet, so add it to the collection
        'without a key
        mCol.Add SurfaceSet
    Else
        'The Tx has a key so add it to the collection with a key
        mCol.Add SurfaceSet, CStr(SurfaceSet.SurfaceSetID)
    End If
    
    AddObj = True
    
End Function


Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub

