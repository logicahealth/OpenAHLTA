VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSubLocation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarSubLocationID As Long
Private mvarSubLocationName As String
Private mvarDeleted As Boolean
Private mvarDirtyFlag As DirtyFlagEnum
'local variable(s) to hold property value(s)
Private mvarDxSubLocID As Long
'local variable(s) to hold property value(s)
Private mvarTxSubLocID As Long
Private mvarSubLocationUpdatedInEncounter As Boolean
Private mvarIsToothLoc As Boolean



Public Property Get IsToothLoc() As Boolean

    IsToothLoc = mvarIsToothLoc
    
End Property


Friend Property Get SubLocationUpdatedInEncounter() As Boolean
    
    SubLocationUpdatedInEncounter = mvarSubLocationUpdatedInEncounter
    
End Property

Friend Property Let SubLocationUpdatedInEncounter(ByRef vData As Boolean)

    mvarSubLocationUpdatedInEncounter = vData
    
End Property

Friend Property Let TxSubLocID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TxSubLocID = 5
    mvarTxSubLocID = vData
End Property


Friend Property Get TxSubLocID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TxSubLocID
    TxSubLocID = mvarTxSubLocID
End Property



Friend Property Let DxSubLocID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DxSubLocID = 5
    mvarDxSubLocID = vData
End Property


Friend Property Get DxSubLocID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DxSubLocID
    DxSubLocID = mvarDxSubLocID
End Property



Public Property Let DirtyFlag(ByVal vData As DirtyFlagEnum)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DirtyFlag = 5
    mvarDirtyFlag = vData
End Property

Friend Function SaveSubLocation(DxLocID As Long, TxLocID As Long) As Boolean

'This function saves the Location to the Note_Dx_Sub_Loc or Note_Tx_Sub_Loc table

Dim sSQL As String
Dim rs As Recordset

    SaveSubLocation = False

    Select Case mvarDirtyFlag
    
        Case DirtyFlagEnum.dfNoAction
            'The SubLocation has not changed since the last save so do nothing
            
        Case DirtyFlagEnum.dfUpdate
            'Update the database
            If DxLocID <> 0 Then
                'Update the Note_Dx_Sub_Loc table
                sSQL = "UPDATE NOTE_DX_SUB_LOCATION" _
                    & " SET SUB_LOCATION_ID = " & mvarSubLocationID & "," _
                    & " DELETED = " & CInt(mvarDeleted) & "" _
                    & " Where (DX_SUBLOC_ID = " & mvarDxSubLocID & ")"
            
            ElseIf TxLocID <> 0 Then
                'Update the Note_Tx_Sub_Loc table
                sSQL = "UPDATE NOTE_TX_SUB_LOCATION" _
                    & " SET SUB_LOCATION_ID = " & mvarSubLocationID & "," _
                    & " DELETED = " & CInt(mvarDeleted) & "" _
                    & " Where (TX_SUBLOC_ID = " & mvarTxSubLocID & ")"
                
            End If
            
            goconn.ExecuteSQL sSQL
            
        Case DirtyFlagEnum.dfAdd
            'Add the Dx to the database
            
            If gbUseStoredProc = True Then

                Dim oCmd As CHCSII_Command.CHCSII_Cmd
                Set oCmd = New CHCSII_Command.CHCSII_Cmd
                
                oCmd.SPName = "Dl_Note_Sub_Location_Insert"
    
                Call oCmd.AddParam(False, False, "textString", DxLocID)
                Call oCmd.AddParam(False, False, "textString", mvarSubLocationID)
                Call oCmd.AddParam(False, False, "textString", CInt(mvarDeleted))
                Call oCmd.AddParam(False, False, "textString", TxLocID)
                Call oCmd.AddParam(True, True, "", "")
                
                Set rs = goconn.OpenRecordsetSP(oCmd)
                
                If Not rs.EOF Then
                    If DxLocID <> 0 Then
                        mvarDxSubLocID = rs("ID")
                    ElseIf TxLocID <> 0 Then
                        mvarTxSubLocID = rs("ID")
                    End If
                Else
                    MsgBxARMd "Fatal error, could not get the identity value from the Dl_Note_Sub_Location_Insert SP."
                    Exit Function
                End If
            Else
            
                If DxLocID <> 0 Then
                    'Add the data to the Note_Dxs_Loc table
                    
'                    'Get the sequence
'                    sSQL = "SELECT NOTE_DX_SUB_LOCATION_SEQ.NEXTVAL AS PKID FROM DUAL"
'                    Set rs = goconn.OpenRecordset(sSQL)
'                    If rs.EOF Then
'                        Exit Function
'                    End If
'                    mvarDxSubLocID = rs("PKID")
'                    rs.Close
                    mvarDxSubLocID = goconn.GetNextID(esNOTE_DX_SUB_LOCATION_SEQ)
                
                    sSQL = "INSERT INTO NOTE_DX_SUB_LOCATION" _
                        & " (DX_LOCATION_ID, SUB_LOCATION_ID, DELETED," _
                        & " DX_SUBLOC_ID)" _
                        & " VALUES (" & DxLocID & ", " _
                        & mvarSubLocationID & ", " _
                        & CInt(mvarDeleted) & ", " _
                        & mvarDxSubLocID & ")"
                    
                ElseIf TxLocID <> 0 Then
                    'Add the data to the Note_Txs_Loc table
                    
'                    'Get the sequence
'                    sSQL = "SELECT NOTE_TXS_LOCATION_SEQ.NEXTVAL AS PKID FROM DUAL"
'                    Set rs = goconn.OpenRecordset(sSQL)
'                    If rs.EOF Then
'                        Exit Function
'                    End If
'                    mvarTxSubLocID = rs("PKID")
'                    rs.Close
                    mvarTxSubLocID = goconn.GetNextID(esNOTE_TXS_LOCATION_SEQ)
                
                    sSQL = "INSERT INTO NOTE_TX_SUB_LOCATION" _
                        & " (TX_LOCATION_ID, SUB_LOCATION_ID, DELETED," _
                        & " TX_SUBLOC_ID)" _
                        & " VALUES (" & TxLocID & ", " _
                        & mvarSubLocationID & ", " _
                        & CInt(mvarDeleted) & ", " _
                        & mvarTxSubLocID & ")"
                    
                End If
                goconn.ExecuteSQL sSQL
                
            End If
    End Select
    
    Set rs = Nothing
    SaveSubLocation = True
    
    mvarDirtyFlag = dfNoAction


End Function


Private Sub SetDirtyFlag()
    
    'This is called when a user sets a property
    'If the current flag is dfAdd, then don't change it until SaveTx is called
    'otherwise, set it to sfUpdate
    
    'SF 10-11-02
    If mvarDirtyFlag = dfNoAction Then
        mvarDirtyFlag = dfUpdate
    End If
    
    mvarSubLocationUpdatedInEncounter = True
    
'    If mvarDirtyFlag = dfAdd Then
'        Exit Sub
'    Else
'        mvarDirtyFlag = dfUpdate
'    End If
    
End Sub


Public Property Get DirtyFlag() As DirtyFlagEnum
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



Friend Property Let SubLocationName(ByVal vData As String)

'This should never be used

'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SubLocationName = 5
    mvarSubLocationName = vData
End Property


Public Property Get SubLocationName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SubLocationName
    SubLocationName = mvarSubLocationName
End Property



Public Property Let SubLocationID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SubLocationID = 5

'Set the location name when the locationid changes
    
Dim sSQL As String
Dim rs As Recordset

    sSQL = "SELECT SUB_LOCATION_NAME, Is_Tooth_Loc" _
        & " From NOTE_SUB_LOCATION" _
        & " WHERE (SUB_LOCATION_ID = " & vData & ")"
    
    Set rs = goconn.OpenRecordset(sSQL)
    
    If rs.EOF Then
        'Could not find the LocationName, so do nothing
    Else
        'Found the location name, so set name and the id
        mvarSubLocationName = rs("SUB_LOCATION_NAME") & ""
        mvarIsToothLoc = Val(rs("Is_Tooth_Loc") & "")
        mvarSubLocationID = vData
        SetDirtyFlag
    End If
    
    rs.Close
    Set rs = Nothing
    
End Property




Public Property Get SubLocationID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SubLocationID
    SubLocationID = mvarSubLocationID
End Property



