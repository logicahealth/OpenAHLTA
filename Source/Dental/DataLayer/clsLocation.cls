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
Attribute VB_Name = "clsLocation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarLocationID As Long
Private mvarLocationName As String
Private mvarSubLocations As colSubLocations
Private mvarDeleted As Boolean
Private mvarDirtyFlag As DirtyFlagEnum
Private mvarTxLocID As Long
Private mvarDxLocID As Long
Private mvarLocationUpdatedInEncounter As Boolean
Private mvarIsToothLoc As Boolean

Public Property Get IsToothLoc() As Boolean

    IsToothLoc = mvarIsToothLoc
    
End Property

Friend Property Let DxLocID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DxLocID = 5
    mvarDxLocID = vData
End Property




Friend Property Get DxLocID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DxLocID
        
    DxLocID = mvarDxLocID

End Property



Friend Property Get LocationUpdatedInEncounter() As Boolean
    
    LocationUpdatedInEncounter = mvarLocationUpdatedInEncounter
    
End Property

Friend Property Let LocationUpdatedInEncounter(ByRef vData As Boolean)

    mvarLocationUpdatedInEncounter = vData
    
End Property

Friend Property Let TxLocID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TxLocID = 5
    mvarTxLocID = vData
End Property


Friend Property Get TxLocID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TxLocID
    TxLocID = mvarTxLocID

End Property



Friend Property Let DirtyFlag(ByVal vData As DirtyFlagEnum)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DirtyFlag = 5
    mvarDirtyFlag = vData
End Property


Friend Property Get DirtyFlag() As DirtyFlagEnum
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DirtyFlag
    DirtyFlag = mvarDirtyFlag
End Property



Public Property Let Deleted(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Deleted = 5
    mvarDeleted = vData
    SetDirtyFlag
End Property


Public Property Get Deleted() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Deleted
    Deleted = mvarDeleted
End Property



Friend Function SaveLocation(DxCodeID As Long, TxCodeID As Long) As Boolean

'This function saves the Location to the Note_Dxs_Loc or Note_Txs_Loc table

Dim sSQL As String
Dim rs As Recordset

    SaveLocation = False

    Select Case mvarDirtyFlag
    
        Case DirtyFlagEnum.dfNoAction
            'The Location has not changed since the last save so do nothing
            
        Case DirtyFlagEnum.dfUpdate
            'Update the database
            If DxCodeID <> 0 Then
                'Update the Note_Dxs_Loc table
                sSQL = "UPDATE NOTE_DXS_LOCATION" _
                    & " SET LOCATION_ID = " & mvarLocationID & "," _
                    & " DELETED = " & CInt(mvarDeleted) & "" _
                    & " Where (DX_LOCATION_ID = " & mvarDxLocID & ")"
            
            ElseIf TxCodeID <> 0 Then
                'Update the Note_Txs_Loc table
                sSQL = "UPDATE NOTE_TXS_LOCATION" _
                    & " SET LOCATION_ID = " & mvarLocationID & "," _
                    & " DELETED = " & CInt(mvarDeleted) & "" _
                    & " Where (TX_LOCATION_ID = " & mvarTxLocID & ")"
                
            End If
            
            goconn.ExecuteSQL sSQL
            
        Case DirtyFlagEnum.dfAdd
            'Add the Dx to the database
            
            If gbUseStoredProc = True Then

                Dim oCmd As CHCSII_Command.CHCSII_Cmd
                Set oCmd = New CHCSII_Command.CHCSII_Cmd
                
                oCmd.SPName = "DL_NOTE_LOCATION_INSERT"
    
                Call oCmd.AddParam(False, False, "textString", mvarLocationID)
                Call oCmd.AddParam(False, False, "textString", CInt(mvarDeleted))
                Call oCmd.AddParam(False, False, "textString", DxCodeID)
                Call oCmd.AddParam(False, False, "textString", TxCodeID)
                Call oCmd.AddParam(True, True, "", "")
                
                Set rs = goconn.OpenRecordsetSP(oCmd)
                
                If Not rs.EOF Then
                    If DxCodeID <> 0 Then
                        mvarDxLocID = rs("ID")
                    ElseIf TxCodeID <> 0 Then
                        mvarTxLocID = rs("ID")
                    End If
                Else
                    MsgBxARMd "Fatal error, could not get the identity value from the DL_NOTE_LOCATION_INSERT SP."
                    Exit Function
                End If
            Else
            
                If DxCodeID <> 0 Then
                    'Add the data to the Note_Dxs_Loc table
                    
'                    'Get the sequence
'                    sSQL = "SELECT NOTE_DXS_LOCATION_SEQ.NEXTVAL AS PKID FROM DUAL"
'                    Set rs = goconn.OpenRecordset(sSQL)
'                    If rs.EOF Then
'                        Exit Function
'                    End If
'                    mvarDxLocID = rs("PKID")
'                    rs.Close
                    mvarDxLocID = goconn.GetNextID(esNOTE_DXS_LOCATION_SEQ)
                
                    sSQL = "INSERT INTO NOTE_DXS_LOCATION" _
                        & " (LOCATION_ID, DELETED, DX_LOCATION_ID, DX_CODE_ID)" _
                        & " VALUES (" _
                        & mvarLocationID & ", " _
                        & CInt(mvarDeleted) & ", " _
                        & mvarDxLocID & "," _
                        & DxCodeID & ")"
                
                ElseIf TxCodeID <> 0 Then
                    'Add the data to the Note_Txs_Loc table
                    
'                    'Get the sequence
'                    sSQL = "SELECT NOTE_TXS_LOCATION_SEQ.NEXTVAL AS PKID FROM DUAL"
'                    Set rs = goconn.OpenRecordset(sSQL)
'                    If rs.EOF Then
'                        Exit Function
'                    End If
'                    mvarTxLocID = rs("PKID")
'                    rs.Close
                    mvarTxLocID = goconn.GetNextID(esNOTE_TXS_LOCATION_SEQ)
                    
                    sSQL = "INSERT INTO NOTE_TXS_LOCATION" _
                        & " (LOCATION_ID, DELETED, TX_LOCATION_ID, TX_CODE_ID)" _
                        & " VALUES (" _
                        & mvarLocationID & ", " _
                        & CInt(mvarDeleted) & ", " _
                        & mvarTxLocID & "," _
                        & TxCodeID & ")"
                    
                End If
                goconn.ExecuteSQL sSQL
            
            End If
            
    End Select
    
    Set rs = Nothing
    SaveLocation = True
    
    'Save the SubLocations if needed
    If Not mvarSubLocations Is Nothing Then
        SaveLocation = mvarSubLocations.SaveSubLocations(mvarDxLocID, mvarTxLocID)
    End If
    
    mvarDirtyFlag = dfNoAction


End Function

Public Property Set SubLocations(ByVal vData As colSubLocations)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.SubLocations = Form1
    Set mvarSubLocations = vData
End Property


Public Property Get SubLocations() As colSubLocations
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SubLocations

    If mvarSubLocations Is Nothing Then
        'Get all of the Txs for this Dx
        Set mvarSubLocations = New colSubLocations

'        If mvarDxLocID <> 0 Then
'            'Load the Dx Sublocations
'            mvarSubLocations.RefreshSubLocationsFromDxLocID mvarDxLocID
'        ElseIf mvarTxLocID <> 0 Then
'            'Load the Tx sublocations
'            mvarSubLocations.RefreshSubLocationsFromTxLocID mvarTxLocID
'        End If
    End If

    Set SubLocations = mvarSubLocations
End Property



Private Property Let LocationName(ByVal vData As String)
'THIS SHOULD NOT BE USED

'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LocationName = 5
    mvarLocationName = vData
End Property


Public Property Get LocationName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LocationName
    LocationName = mvarLocationName
End Property



Public Property Let LocationID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LocationID = 5
    
'Set the location name when the locationid changes
    
Dim sSQL As String
Dim rs As Recordset

    sSQL = "SELECT LOCATION_NAME, Is_Tooth_Loc" _
        & " From NOTE_LOCATION" _
        & " WHERE (LOCATION_ID = " & vData & ")"
    
    Set rs = goconn.OpenRecordset(sSQL)
    
    If rs.EOF Then
        'Could not find the LocationName, so do nothing
    Else
        'Found the location name, so set name and the id
        mvarLocationName = rs("LOCATION_NAME") & ""
        mvarIsToothLoc = Val(rs("Is_Tooth_Loc") & "")
        mvarLocationID = vData
        SetDirtyFlag
    End If
    
    rs.Close
    Set rs = Nothing
    
End Property

Private Sub SetDirtyFlag()
    
    'This is called when a user sets a property
    'If the current flag is dfAdd, then don't change it until SaveTx is called
    'otherwise, set it to sfUpdate
    
    'SF 10-11-02
    If mvarDirtyFlag = dfNoAction Then
        mvarDirtyFlag = dfUpdate
    End If
    
    mvarLocationUpdatedInEncounter = True
    
'    If mvarDirtyFlag = dfAdd Then
'        Exit Sub
'    Else
'        mvarDirtyFlag = dfUpdate
'    End If
    
End Sub


Public Property Get LocationID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LocationID
    LocationID = mvarLocationID
End Property



Private Sub Class_Terminate()

    Set mvarSubLocations = Nothing

End Sub


