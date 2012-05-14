VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarDentalEncounterID As Long
Private mvarCurrentDentalEncounterID As Long
Private mvarICDCode As String
Private mvarICDDefinition As String
Private mvarICDCategory As String
Private mvarDRC As Long
Private mvarDRCMin As Long
Private mvarDRCMax As Long
Private mvarEffectiveDate As Date
Private mvarEffectiveEndDate As Date
Private mvarStatusID As DxStatusEnum
Private mvarPriority As Long
Private mvarProviderID As Long
Private mvarCreateDateTime As Date
Private mvarDxICDCodeID As Long
Private mvarDxCodeID As Long
Private mvarSurfaceSets As colSurfaceSets
Private mvarTxs As colTxs
Private mvarChartRelated As Long 'If the ICD code is a (chart related) chartable code or not
Private mvarChartable As SourceEnum 'If the Dx was added through GC or Notes
Private mvarDirtyFlag As DirtyFlagEnum
Private mvarLocations As colLocations
'Private mvarSource As SourceEnum
Private mvarStatus As String
Private mvarProviderName As String
Private mvarSigned As Boolean
Private mvarGroupID As Long
Private mvarPerioStateID As DxPerioStateEnum
Private mvarPerioState As String
Private mvarDeleteReason As DeleteReasonEnum
Private mvarModifiedDateTime As Date
Private mvarModifiedInDentalEncounterID As Long
'Was a sublocation updated in this Dx, byref down to the sublocation level
Private mvarSubLocationUpdatedInEncounter As Boolean

'This is not saved and is only used by graphic charting, SVG_Manager:clsManage.cls
Private mvarGCIsDirty As Boolean

'For SingleTooth Mode only
Private mvarToothID As Long
Private mvarInternalAdd As Boolean

'Collection to hold all Txs to ensure unique data integrity
Private mvarTxsRef As Collection

'Collection to hold all Dxs to ensure unique data integrity
Private mvarDxsRef As Collection

Public Property Get GCIsDirty() As Boolean
    'This is not saved and is only used by graphic charting, SVG_Manager:clsManage.cls
    GCIsDirty = mvarGCIsDirty
End Property
Public Property Let GCIsDirty(vData As Boolean)
    'This is not saved and is only used by graphic charting, SVG_Manager:clsManage.cls
    mvarGCIsDirty = vData

End Property
Public Property Get DeleteReason() As DeleteReasonEnum
    
    DeleteReason = mvarDeleteReason
    
End Property

Public Property Let DeleteReason(vData As DeleteReasonEnum)
    
    If vData = 0 Then
        mvarDeleteReason = drNone
    Else
        mvarDeleteReason = vData
    End If
    
    SetDirtyFlag
    
End Property

Public Property Get DeleteReasonText() As String
    
    DeleteReason = LoadResString(mvarDeleteReason)
    
End Property

Public Function GetNote() As String

'This function gets the note from the note recordset
'Use the global recordset to obtain the previous values

Dim sTemp As String
Dim bChanged As Boolean 'If the dx changed then display it in the note.

Dim oLoc As clsLocation
Dim oSubLoc As clsSubLocation
Dim oSS As clsSurfaceSet
Dim bSomethingChanged As Boolean

    'Tracker 21622
    'If the status was deleted from this encounter then don't display anything
    If Me.StatusID = dxNonActiveDeleted And Me.DentalEncounterID = Me.CurrentDentalEncounterID Then
        Exit Function
    End If
           
    'Determine if a location or sublocation has been modified in this encounter.
    If Not mvarLocations Is Nothing Then
        For Each oLoc In mvarLocations
            If oLoc.LocationUpdatedInEncounter = True Then
                mvarModifiedInDentalEncounterID = mvarCurrentDentalEncounterID
                Exit For
            End If
            For Each oSubLoc In oLoc.SubLocations
                If oSubLoc.SubLocationUpdatedInEncounter = True Then
                    mvarModifiedInDentalEncounterID = mvarCurrentDentalEncounterID
                    Exit For
                End If
            Next
        Next
    End If
    
    If Not mvarSurfaceSets Is Nothing Then
        For Each oSS In mvarSurfaceSets
            If oSS.SSChangedInEncounter = True Then
                mvarModifiedInDentalEncounterID = mvarCurrentDentalEncounterID
                Exit For
            End If
        Next
    End If
    
    'If the dx was modified in the current encounter, then display it in the note
    If mvarModifiedInDentalEncounterID = mvarCurrentDentalEncounterID Then
        sTemp = sTemp & GetNoteText("LOCATION_NAME")
        sTemp = sTemp & GetNoteText("SUB_LOCATION_NAME")
        sTemp = sTemp & GetNoteText("TOOTH_ID")
        sTemp = sTemp & GetNoteText("SURFACEABBR")
        sTemp = sTemp & GetNoteText("ICD_DEFINITION")
        sTemp = sTemp & GetNoteText("STATUS_ID")
        sTemp = sTemp & GetNoteText("DRC")
        sTemp = sTemp & GetNoteText("PERIO_STATE")
    End If
    
    If sTemp <> "" Then
        'Remove the first ", "
        sTemp = Mid(sTemp, Len(" ") + 1)
    End If
    
    If sTemp <> "" Then
        GetNote = crtfStartControlName & sTemp & crtfEndControlName
    Else
        GetNote = ""
    End If
    
End Function

Private Function GetNoteText(Field As String) As String

'This function gets the text of the note based on the field you want

    Select Case Field
        
        Case "ICD_DEFINITION"
            GetNoteText = ", " & mvarICDDefinition
            
        Case "DRC"
            GetNoteText = ", " & mvarDRC
            
        Case "LOCATION_NAME"
            If mvarChartable = sNotes Then
                'Get the location
                If Me.Locations.Count > 0 Then
                    If Me.Locations(1).IsToothLoc Then
                        GetNoteText = ", #" & Me.Locations(1).LocationName
                    Else
                        GetNoteText = ", " & Me.Locations(1).LocationName
                    End If
                End If
            End If
            
        Case "SUB_LOCATION_NAME"
            If mvarChartable = sNotes Then
                'Get the sublocation
                If Me.Locations.Count > 0 Then
                    If Me.Locations(1).SubLocations.Count > 0 Then
                        If Me.Locations(1).SubLocations(1).SubLocationName <> "" Then
                            If Me.Locations(1).SubLocations(1).SubLocationName <> "" Then
                                GetNoteText = ", " & Me.Locations(1).SubLocations(1).SubLocationName
                            End If
                        End If
                    End If
                End If
            End If
            
        Case "TOOTH_ID"
            If mvarChartable = sCharts Then
                'Chartable Dx
                If mvarSurfaceSets.Count > 0 Then
                    'Get the tooth
                    If mvarSurfaceSets(1).Teeth.Count > 0 Then
                        GetNoteText = " #" & mvarSurfaceSets(1).Teeth(1).ToothName
                    End If
                End If
            End If

        Case "SURFACEABBR"
            If mvarChartable = sCharts Then
                'Chartable Dx
                If mvarSurfaceSets.Count > 0 Then
                    'Get the surfaceset
                    If mvarSurfaceSets(1).SurfaceAbbr <> "" Then
                        If mvarSurfaceSets(1).SurfaceAbbr <> "" Then
                            GetNoteText = ", " & mvarSurfaceSets(1).SurfaceAbbr
                        End If
                    End If
                End If
            End If

        Case "STATUS_ID"
            GetNoteText = ", " & mvarStatus
    
        Case "PERIO_STATE"
            If mvarPerioState <> "" Then
                GetNoteText = ", " & mvarPerioState
            End If
            
        Case Else
            MsgBxARMd "Wrong parameter in Dx.GetNoteText"
            
    End Select
    
End Function


Public Function GeneralLocationName() As String

'This returns the name of the first location or a "" if there is no location

Dim oLoc As clsLocation
    
    If mvarChartable = sCharts Then
        If Not mvarSurfaceSets Is Nothing Then
            If mvarSurfaceSets.Count > 0 Then
                If mvarSurfaceSets(1).Teeth.Count > 0 Then
                    GeneralLocationName = mvarSurfaceSets(1).Teeth(1).ToothName
                End If
            End If
        End If
    Else
        If Not mvarLocations Is Nothing Then
            If mvarLocations.Count > 0 Then
                GeneralLocationName = mvarLocations(1).LocationName
            End If
        End If
    End If
    
End Function

Public Property Get ModifiedDateTime() As Date
    
    ModifiedDateTime = mvarModifiedDateTime
    
End Property

Public Property Get ModifiedInDentalEncounterID() As Long

    ModifiedInDentalEncounterID = mvarModifiedInDentalEncounterID
    
End Property

Friend Property Let ModifiedInDentalEncounterID(vData As Long)

    mvarModifiedInDentalEncounterID = vData
    
End Property

Private Property Let PerioState(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PerioState = 5
    mvarPerioState = vData
End Property


Public Property Get PerioState() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PerioState
    PerioState = mvarPerioState
End Property




Public Property Let PerioStateID(ByVal vData As DxPerioStateEnum)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.PerioStateID = Form1
    
    'Get the PerioState Text from the resource file
    'If vData is 0, then the Dx is not a perio dx, for backwards compatibility with existing data.
    If vData = DxNotPerio Or vData = 0 Then
        vData = DxNotPerio
        mvarPerioState = ""
    Else
        mvarPerioState = LoadResString(vData)
    End If
    
    mvarPerioStateID = vData
    SetDirtyFlag

End Property


Public Property Get PerioStateID() As DxPerioStateEnum
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PerioStateID
    PerioStateID = mvarPerioStateID
End Property



Public Property Let GroupID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.GroupID = 5
    mvarGroupID = vData
    SetDirtyFlag
End Property


Public Property Get GroupID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.GroupID
    GroupID = mvarGroupID
End Property



Private Property Let ProviderName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ProviderName = 5
    mvarProviderName = vData
End Property


Public Property Get ProviderName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ProviderName

'Query the Clinic_Provider table to get the name
'Only used in reports
    
Dim rs As Recordset
Dim sSQL As String

    If Val(mvarProviderID) <> 0 Then
        'Query the DB
        sSQL = "SELECT PROVIDER_NAME" _
            & " From CDRPLUS.CLINIC_PROVIDER" _
            & " Where (CLINICIAN_NCID = " & mvarProviderID & ")"
        Set rs = goconn.OpenRecordset(sSQL)
        If rs.EOF Then
            'If rs.eof then return ""
            mvarProviderName = ""
        Else
            'Set the provider name
            mvarProviderName = rs("PROVIDER_NAME") & ""
        End If
        rs.Close
        Set rs = Nothing
    Else
        mvarProviderName = ""
    End If
    
    ProviderName = mvarProviderName

End Property




Friend Property Let InternalAdd(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.InternalAdd = 5
    mvarInternalAdd = vData
End Property


Friend Property Get InternalAdd() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.InternalAdd
    InternalAdd = mvarInternalAdd
End Property



Friend Property Let ToothID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ToothID = 5
    mvarToothID = vData
End Property


Friend Property Get ToothID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ToothID
    ToothID = mvarToothID
End Property



Friend Property Let DRCMax(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DRCMax = 5
    mvarDRCMax = vData
End Property


Public Property Get DRCMax() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DRCMax
    DRCMax = mvarDRCMax
End Property



Friend Property Let DRCMin(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DRCMin = 5
    mvarDRCMin = vData
End Property


Public Property Get DRCMin() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DRCMin
    DRCMin = mvarDRCMin
End Property



Private Property Let Status(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Status = 5
    mvarStatus = vData
End Property


Public Property Get Status() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Status
    Status = mvarStatus
End Property



Private Property Let Source(ByVal vData As SourceEnum)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Source = Form1
'    mvarSource = vdata
End Property


Private Property Get Source() As SourceEnum
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Source
'    Source = mvarSource
End Property



Public Property Set Locations(ByVal vData As colLocations)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Locations = Form1
    Set mvarLocations = vData
End Property


Public Property Get Locations() As colLocations
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Locations
        
    If mvarLocations Is Nothing Then
        'Get all of the Txs for this Dx
        Set mvarLocations = New colLocations
'        mvarLocations.RefreshLocationsFromDxCodeID mvarDxCodeID, mvarChartable
    End If
    
    Set Locations = mvarLocations
    
End Property










Public Property Set SurfaceSets(ByVal vData As colSurfaceSets)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.SurfaceSets = Form1
    Set mvarSurfaceSets = vData
End Property


Public Property Get SurfaceSets() As colSurfaceSets
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SurfaceSets
    
    If mvarSurfaceSets Is Nothing Then
'        'Get all of the Txs for this Dx
        Set mvarSurfaceSets = New colSurfaceSets
'        mvarSurfaceSets.RefreshSurfaceSetsFromDxORTxCodeID mvarDxCodeID, 0, mvarChartable
    End If
    
    Set SurfaceSets = mvarSurfaceSets
End Property



Public Property Let DirtyFlag(ByVal vData As DirtyFlagEnum)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.DirtyFlag = Form1
    mvarDirtyFlag = vData
End Property


Public Property Get DirtyFlag() As DirtyFlagEnum
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DirtyFlag
    DirtyFlag = mvarDirtyFlag
End Property



Public Property Let Chartable(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Chartable = 5
    mvarChartable = vData
    SetDirtyFlag
End Property


Public Property Let Signed(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Chartable = 5
    mvarSigned = vData
    SetDirtyFlag
End Property

Friend Function SaveDx(GroupID As Long) As Boolean

'This function saves the Dx to the Dental_Encounter_Dx table

Dim sSQL As String
Dim rs As Recordset
Dim sOraTimeFormat As String

On Error GoTo SaveDxErr

    SaveDx = False

'    If InStr(UCase(mvarCreateDateTime), "AM") Or InStr(UCase(mvarCreateDateTime), "PM") Then
'        'The time is in civilian time
'        sOraTimeFormat = "MM/DD/RRRR HH:MI:SS AM"
'    Else
'        'The time is in military time
'        sOraTimeFormat = "MM/DD/RRRR HH24:MI:SS"
'    End If
    
    'Make sure the date is changed to reflect the update date/time
    'mvarCreateDateTime = GetSysDate
    
    Select Case mvarDirtyFlag
    
        Case DirtyFlagEnum.dfNoAction
            'The Dx has not changed since the last save so do nothing
            
        Case DirtyFlagEnum.dfUpdate
            'Update the database
            
            'This is updated when a property of the Dx object is changed.
            'mvarModifiedDateTime = GetSysDate
            
            If mvarDxCodeID = 0 Then
                MsgBxARMd "WARNING: mvarDxCodeID = 0 and it is trying to update the database with it. Aborting update. Please tell Scott how you encountered this. And yes this message will be removed when we solve the problem."
                Exit Function
            End If
            
            sSQL = "UPDATE DENTAL_ENCOUNTER_DX" _
                & " SET DX_ICD_CODE_ID = " & mvarDxICDCodeID & "," _
                & " CREATE_DATE_TIME = " & goChcsiiConn.FormatDateForDBInsert(mvarCreateDateTime) & "," _
                & " MODIFIED_DATE_TIME = " & goChcsiiConn.FormatDateForDBInsert(mvarModifiedDateTime) & "," _
                & " PROVIDER_ID = " & mvarProviderID & "," _
                & " DENTAL_ENCOUNTER_ID = " & mvarDentalEncounterID & "," _
                & " STATUS_ID = " & mvarStatusID & "," _
                & " CHARTABLE = " & CInt(mvarChartable) & "," _
                & " DRC = " & mvarDRC & "," _
                & " SIGNED = " & CInt(mvarSigned) & "," _
                & " PERIO_STATE = " & mvarPerioStateID & "," _
                & " DELETE_REASON = " & mvarDeleteReason & "," _
                & " MODIFIED_IN_DENTAL_ENCOUNTER = " & mvarModifiedInDentalEncounterID & "" _
                & " WHERE (DX_CODE_ID = " & mvarDxCodeID & ")"
            goconn.ExecuteSQL sSQL
            
            'Update the groups table
            sSQL = "Update DENTAL_ENCOUNTER_GROUPS" _
                & " Set GROUP_ID = " & mvarGroupID & "" _
                & " Where (DX_CODE_ID = " & mvarDxCodeID & ")"
            goconn.ExecuteSQL sSQL
            
        Case DirtyFlagEnum.dfAdd
            'Add the Dx to the database
            
            'This is updated when a property of the Dx object is changed.
            'mvarModifiedDateTime = GetSysDate
            
            If gbUseStoredProc = True Then

                Dim oCmd As CHCSII_Command.CHCSII_Cmd
                Set oCmd = New CHCSII_Command.CHCSII_Cmd
                
                oCmd.SPName = "Dl_Note_Dx_Insert"
    
                'Call oCmd.AddParam(False, True, "textString", mvarDxCodeID)
                Call oCmd.AddParam(False, False, "textString", mvarDxICDCodeID)
                Call oCmd.AddParam(False, False, "textString", mvarCreateDateTime)
                Call oCmd.AddParam(False, False, "textString", mvarModifiedDateTime)
                Call oCmd.AddParam(False, False, "textString", mvarProviderID)
                Call oCmd.AddParam(False, False, "textString", mvarDentalEncounterID)
                Call oCmd.AddParam(False, False, "textString", mvarStatusID)
                Call oCmd.AddParam(False, False, "textString", mvarDRC)
                Call oCmd.AddParam(False, False, "textString", CInt(mvarSigned))
                Call oCmd.AddParam(False, False, "textString", mvarPerioStateID)
                Call oCmd.AddParam(False, False, "textString", mvarDeleteReason)
                Call oCmd.AddParam(False, False, "textString", mvarModifiedInDentalEncounterID)
                Call oCmd.AddParam(False, False, "textString", CInt(mvarChartable))
                Call oCmd.AddParam(False, False, "textString", mvarGroupID)
                Call oCmd.AddParam(False, False, "textString", sOraTimeFormat)
                Call oCmd.AddParam(True, True, "", "")
                
                Set rs = goconn.OpenRecordsetSP(oCmd)
                
                If Not rs.EOF Then
                    mvarDxCodeID = rs("ID")
                Else
                    MsgBxARMd "Fatal error, could not get the identity value from the SP."
                    Exit Function
                End If
            Else
                'Get the sequence
'                sSQL = "SELECT DENTAL_ENCOUNTER_DX_SEQ.NEXTVAL AS PKID FROM DUAL"
'                Set rs = goconn.OpenRecordset(sSQL)
'                If rs.EOF Then
'                    Exit Function
'                End If
'                mvarDxCodeID = rs("PKID")
'                rs.Close
                mvarDxCodeID = goconn.GetNextID(esDENTAL_ENCOUNTER_DX_SEQ)
                
                sSQL = "INSERT INTO DENTAL_ENCOUNTER_DX" _
                    & " (DX_CODE_ID," _
                    & " DX_ICD_CODE_ID," _
                    & " CREATE_DATE_TIME," _
                    & " MODIFIED_DATE_TIME," _
                    & " PROVIDER_ID," _
                    & " DENTAL_ENCOUNTER_ID," _
                    & " STATUS_ID," _
                    & " DRC," _
                    & " SIGNED," _
                    & " PERIO_STATE," _
                    & " DELETE_REASON," _
                    & " MODIFIED_IN_DENTAL_ENCOUNTER," _
                    & " CHARTABLE)" _
                    & " VALUES ("
                sSQL = sSQL _
                    & " " & mvarDxCodeID & "," _
                    & " " & mvarDxICDCodeID & "," _
                    & " " & goChcsiiConn.FormatDateForDBInsert(mvarCreateDateTime) & "," _
                    & " " & goChcsiiConn.FormatDateForDBInsert(mvarModifiedDateTime) & "," _
                    & " " & mvarProviderID & "," _
                    & " " & mvarDentalEncounterID & "," _
                    & " " & mvarStatusID & "," _
                    & " " & mvarDRC & "," _
                    & " " & CInt(mvarSigned) & "," _
                    & " " & mvarPerioStateID & "," _
                    & " " & mvarDeleteReason & "," _
                    & " " & mvarModifiedInDentalEncounterID & "," _
                    & " " & CInt(mvarChartable) & ")"
                goconn.ExecuteSQL sSQL
        
                'Add the new dx to the groups table
                sSQL = "INSERT INTO DENTAL_ENCOUNTER_GROUPS" _
                    & " (GROUP_ID, DX_CODE_ID, TX_CODE_ID)" _
                    & " VALUES (" & mvarGroupID & ", " & mvarDxCodeID & ",0)"
                goconn.ExecuteSQL sSQL
            End If
            
        Case DirtyFlagEnum.dfdelete
            If mvarDxCodeID <> 0 Then
                'Delete the Dx and it's associated data
                'Delete the sublocation first
                sSQL = "DELETE FROM NOTE_DX_SUB_LOCATION" _
                    & " WHERE DX_LOCATION_ID IN" _
                    & " (SELECT DX_LOCATION_ID" _
                    & " From NOTE_DXS_LOCATION" _
                    & " WHERE (DX_CODE_ID = " & mvarDxCodeID & "))"
                goconn.ExecuteSQL sSQL
                
                'Delete the location next
                sSQL = "DELETE FROM NOTE_DXS_LOCATION" _
                    & " Where (DX_CODE_ID = " & mvarDxCodeID & ")"
                goconn.ExecuteSQL sSQL
                
                'Delete the Dx group info
                sSQL = "DELETE FROM DENTAL_ENCOUNTER_GROUPS" _
                    & " Where (DX_CODE_ID = " & mvarDxCodeID & ")"
                goconn.ExecuteSQL sSQL
                
                'Delete the Surfaceset relationship to the Dx
                sSQL = "DELETE FROM DENTAL_ENCOUNTER_DXS_SS" _
                    & " Where (DX_CODE_ID = " & mvarDxCodeID & ")"
                goconn.ExecuteSQL sSQL
                
                'Delete the Dx
                sSQL = "DELETE FROM DENTAL_ENCOUNTER_DX" _
                    & " Where (DX_CODE_ID = " & mvarDxCodeID & ")"
                goconn.ExecuteSQL sSQL
                
                'Set the DxCodeID to 0 because it does not exist in the db anymore
                mvarDxCodeID = 0
                
            End If
    End Select

'    'This must be here because if two references are made to it, it only needs to save the
'    'Dx once, but it must add the relationships in both (or more) cases.
'    'Add the Dx to the TxCodeID. Only applicable if this Dx is in a collection that is
'    'under a Tx.
'    If TxCodeID <> 0 Then
'        'Find out if the relationship exists
'        sSQL = "SELECT DX_CODE_ID, TX_CODE_ID" _
'            & " From DENTAL_ENCOUNTER_TXS_DXS" _
'            & " WHERE (DX_CODE_ID = " & mvarDxCodeID & ") AND (TX_CODE_ID = " & TxCodeID & ")"
'        Set rs = goconn.OpenRecordset(sSQL)
'        If rs.EOF Then
'            'Add the relationship
'            sSQL = "INSERT INTO DENTAL_ENCOUNTER_TXS_DXS" _
'                & " (DX_CODE_ID, TX_CODE_ID)" _
'                & " VALUES (" & mvarDxCodeID & ", " & TxCodeID & ")"
'            goconn.ExecuteSQL sSQL
'        End If
'        rs.Close
'    End If
    
    Set rs = Nothing
    SaveDx = True
    
'    'Save the Txs if needed
'    If Not mvarTxs Is Nothing Then
'        SaveDx = mvarTxs.SaveTxs(mvarDxCodeID)
'    End If
'
    'Save the SurfaceSets if needed
    If Not mvarSurfaceSets Is Nothing And mvarDxCodeID <> 0 Then
        SaveDx = mvarSurfaceSets.SaveSurfaceSets(mvarDxCodeID, 0)
    End If
    
    'Save the locations if needed
    If Not mvarLocations Is Nothing And mvarDxCodeID <> 0 Then
        SaveDx = mvarLocations.SaveLocations(mvarDxCodeID, 0)
    End If
    
    mvarDirtyFlag = dfNoAction

Exit Function

SaveDxErr:

    If InStr(Err.Description, "Error Text = ORA-00100: no data found Rollback work ORA-00100: no data found") > 0 Then
    'I can't use the err.number because chcsii sql always returns the same error number for all errors
    'If Err.Number = -2147221503 Then
        'The db server reported "No data found", so continue
        Resume Next
    Else
        MsgBxARMd Err.Description & " -- " & sSQL, vbCritical, Err.Source
        'Err.Raise Err.Number, Err.Source, Err.Description
        Exit Function
        Resume
    End If

End Function

Private Sub SetDirtyFlag()
    
    'This is called when a user sets a property
    'If the current flag is dfAdd, then don't change it until SaveDx is called
    'otherwise, set it to sfUpdate

Static dTempSysDate As Date
Static icounter As Integer
    
    '4-24-02 Jon wants to update the modified date time every time the user changes a property.
    mvarModifiedDateTime = GetSysDate
    
    mvarModifiedInDentalEncounterID = mvarCurrentDentalEncounterID
    
    'SF 10-11-02
    If mvarDirtyFlag = dfNoAction Then
        mvarDirtyFlag = dfUpdate
    End If
    
'    If mvarDirtyFlag = dfAdd Then
'        Exit Sub
'    Else
'        mvarDirtyFlag = dfUpdate
'    End If
    
End Sub

Public Property Get Chartable() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Chartable
    Chartable = mvarChartable
End Property


Public Property Get Signed() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Chartable
    Signed = mvarSigned
End Property






Friend Property Let DxsRef(vData As Collection)

    Set mvarDxsRef = vData
    
End Property

Friend Property Get DxsRef() As Collection

    Set DxsRef = mvarDxsRef

End Property



Friend Property Let DxCodeID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DxCodeID = 5
    mvarDxCodeID = vData
End Property


Public Property Get DxCodeID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DxCodeID
    DxCodeID = mvarDxCodeID
End Property



Public Property Let DxICDCodeID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DxICDCodeID = 5
    
'Refresh the data from the DX_ICD_CODE_DEFINITION table

Dim sSQL As String
Dim rs As Recordset
    
    If mvarInternalAdd = True Then
        'Set the property
        mvarDxICDCodeID = vData
        Exit Property
    Else
        'Update the fields
        sSQL = "SELECT ICD_CODE, ICD_CATEGORY, ICD_DEFINITION," _
            & " CHART_RELATED, DRC, DRCMIN, DRCMAX From DX_ICD_CODE_DEFINITION" _
            & " Where (DX_ICD_CODE_ID = " & vData & ")"
        
        Set rs = New Recordset
        Set rs = goconn.OpenRecordset(sSQL)
    
        If Not rs.EOF Then
            mvarICDCode = rs("ICD_CODE") & ""
            mvarICDCategory = rs("ICD_CATEGORY") & ""
            mvarICDDefinition = rs("ICD_DEFINITION") & ""
            mvarChartRelated = Val(rs("CHART_RELATED") & "")
            mvarDRC = Val(rs("DRC") & "")
            mvarDRCMin = Val(rs("DRCMIN") & "")
            mvarDRCMax = Val(rs("DRCMAX") & "")
        ElseIf vData = -1 Then
            'This is the special code for a Dx that should not show in the grid.
            ' So the user thinks they are adding a Tx without a Dx.
        Else
            'Err.Raise vbObjectError + 514, "Property Let DxICDCodeID", "Invalid ICDCodeID"
            MsgBxARMd "Invalid ICDCodeID", vbCritical, "Property Let DxICDCodeID"
            Exit Property
        End If
        rs.Close
        Set rs = Nothing
        
        mvarDxICDCodeID = vData
        SetDirtyFlag
    End If
    
End Property


Public Property Get DxICDCodeID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DxICDCodeID
    DxICDCodeID = mvarDxICDCodeID
End Property



Public Property Let CreateDateTime(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CreateDateTime = 5
    mvarCreateDateTime = vData
    SetDirtyFlag
End Property


Friend Property Let ModifiedDateTime(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CreateDateTime = 5
    SetDirtyFlag
    mvarModifiedDateTime = vData
End Property

Public Property Get CreateDateTime() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CreateDateTime
    CreateDateTime = mvarCreateDateTime
End Property



Public Property Let ProviderID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ProviderID = 5
    mvarProviderID = vData
    SetDirtyFlag
End Property


Public Property Get ProviderID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ProviderID
    ProviderID = mvarProviderID
End Property



Public Property Let Priority(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Priority = 5
    mvarPriority = vData
    SetDirtyFlag
End Property


Public Property Get Priority() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Priority
    Priority = mvarPriority
End Property



Public Property Let StatusID(ByVal vData As DxStatusEnum)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.StatusID = 5

Dim lCounter As Long
Dim rs As Recordset
Dim sSQL As String
    
    If vData <> -1 Then
        'If the object has been deleted, but the user wants to change the status to
        ' new or something other than deleted, then mark the object as being added
        ' back again.
        If vData <> dxNonActiveDeleted And mvarStatusID = dxNonActiveDeleted And mvarDirtyFlag = dfNoAction Then
            mvarDirtyFlag = dfAdd
        End If

        If vData = dxNonActiveDeleted Then
            If mvarDentalEncounterID = mvarCurrentDentalEncounterID Then
                'The Dx is from this encounter
                If mvarSigned = False Then
                    'If there is a more recent encounter, then don't remove the Dx from
                    ' the database
                    sSQL = "SELECT MAX(DENTAL_ENCOUNTER_ID) AS MAXDE From DENTAL_ENCOUNTER" _
                        & " WHERE (PATIENT_UNIT_ID IN (SELECT PATIENT_UNIT_ID" _
                        & " From DENTAL_ENCOUNTER WHERE (DENTAL_ENCOUNTER_ID = " & mvarCurrentDentalEncounterID & ")))" _
                        & " Having (Max(DENTAL_ENCOUNTER_ID) > " & mvarCurrentDentalEncounterID & ")"

                    Set rs = goconn.OpenRecordset(sSQL)
                    If rs.EOF Then
                        'The Dx is unsigned and does not have a subsequent encounter
                        ' , so remove the Dx from the database
                        
                        mvarDirtyFlag = dfdelete
                    Else
                        'The Dx is unsigned, but does have a subsequent encounter
                        ' , so mark the Dx as deleted
                        'It will be, so do nothing
                    End If
                    rs.Close
                    Set rs = Nothing
                Else
                    'The Dx is signed, so set the statusid to deleted
                    'Get the Status Text from the resource file
                    'It will be, so do nothing
                End If
            Else
                'The Dx is from an old encounter, and it doesn't matter if the old
                ' encounter is signed or not. So, set the statusid to deleted
                'It will be, so do nothing
            End If
        End If
        
        'Get the Status Text from the resource file
        mvarStatus = LoadResString(vData)
        mvarStatusID = vData
        'If the dirtyflag is set to dfdelete, then don't change it
        If mvarDirtyFlag <> dfdelete Then
            SetDirtyFlag
        End If
    
    End If

End Property


Public Property Get StatusID() As DxStatusEnum
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.StatusID
    StatusID = mvarStatusID
End Property












Public Property Let DRC(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DRC = 5
    mvarDRC = vData
    SetDirtyFlag
End Property


Public Property Get DRC() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DRC
    DRC = mvarDRC
End Property



Friend Property Let ChartRelated(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Chart_Related = 5
    mvarChartRelated = vData
End Property


Public Property Get ChartRelated() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Chart_Related
    ChartRelated = mvarChartRelated
End Property



Friend Property Let ICDCategory(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ICDCategory = 5
    mvarICDCategory = vData
End Property


Public Property Get ICDCategory() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ICDCategory
    ICDCategory = mvarICDCategory
End Property



Friend Property Let ICDDefinition(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ICDDefinition = 5
    mvarICDDefinition = vData
End Property


Public Property Get ICDDefinition() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ICDDefinition
    ICDDefinition = mvarICDDefinition
End Property



Friend Property Let ICDCode(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ICDCode = 5
    mvarICDCode = vData
End Property


Public Property Get ICDCode() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ICDCode
    ICDCode = mvarICDCode
End Property



Public Property Let DentalEncounterID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DentalEncounterID = 5
    mvarDentalEncounterID = vData
    SetDirtyFlag
End Property


Friend Property Let CurrentDentalEncounterID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DentalEncounterID = 5
    
Dim oLoc As clsLocation
Dim oSubLoc As clsSubLocation
Dim oSS As clsSurfaceSet
    
    If vData <> mvarCurrentDentalEncounterID Then
    
        mvarCurrentDentalEncounterID = vData
        
        'Reset the LocationModifiedInEncounter, SubLocationModifiedInEncounter and SSModifiedInEncounter values
        If Not mvarLocations Is Nothing Then
            For Each oLoc In mvarLocations
                oLoc.LocationUpdatedInEncounter = False
                
                For Each oSubLoc In oLoc.SubLocations
                    oSubLoc.SubLocationUpdatedInEncounter = False
                Next
            Next
        End If
        
        If Not mvarSurfaceSets Is Nothing Then
            For Each oSS In mvarSurfaceSets
                oSS.SSChangedInEncounter = False
            Next
        End If
        
    End If
    
End Property

Public Property Get DentalEncounterID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DentalEncounterID
    DentalEncounterID = mvarDentalEncounterID
End Property
Friend Property Get CurrentDentalEncounterID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DentalEncounterID
    CurrentDentalEncounterID = mvarCurrentDentalEncounterID
End Property



Private Sub Class_Initialize()
    
    'Set the default
    mvarDeleteReason = drNone
    
End Sub


Private Sub Class_Terminate()

    Set mvarSurfaceSets = Nothing
    Set mvarTxs = Nothing
    Set mvarLocations = Nothing
    Set mvarTxsRef = Nothing
    Set mvarDxsRef = Nothing

End Sub


