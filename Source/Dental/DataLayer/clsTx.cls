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
Attribute VB_Name = "clsTx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Used to calculate workload in SaveTx
Private Const GTRMembraneRemoval As Long = 166

Private mvarTxCodeID As Long
Private mvarDentalEncounterID As Long
Private mvarCurrentDentalEncounterID As Long
Private mvarTxADACodeID As Long
Private mvarStatusID As TxStatusEnum
Private mvarCreateDateTime As Date
Private mvarProviderID As Long
Private mvarADACategoryCode As String
Private mvarADACode As String
Private mvarADADefinition As String
Private mvarADAAlias As String
Private mvarADAHelpNarative As String
Private mvarLabFlag As Long
Private mvarDWV As Long
Private mvarSurfaceSets As colSurfaceSets
Private mvarDxs As colDxs
Private mvarPrognosisID As Long
Private mvarChartRelated As Long 'If the ICD code is a (chart related) chartable code or not
Private mvarChartable As SourceEnum  'If the Dx was added through GC or Notes
Private mvarDirtyFlag As DirtyFlagEnum
Private mvarLocations As colLocations
Private mvarSource As SourceEnum
Private mvarStatus As String
Private mvarPrognosis As String
Private mvarComments As String
Private mvarMaterials As String
Private mvarToothID As Long
Private mvarInternalAdd As Boolean
Private mvarProviderName As String
Private mvarSigned As Boolean
Private mvarGroupID As Long
Private mvarIsPerio As Boolean
Private mvarIsRPD As Boolean
Private mvarIsFPD As Boolean
Private mvarIsCompleteDenture As Boolean
Private mvarPhaseID As TxPhaseEnum
Private mvarPhase As String
Private mvarWorkloadID As Long
Private mvarDeleteReason As DeleteReasonEnum
Private mvarModifiedDateTime As Date
Private mvarModifiedInDentalEncounterID As Long

'This is not saved and is only used by graphic charting, SVG_Manager:clsManage.cls
Private mvarGCIsDirty As Boolean

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
Public Function HasLabCase() As Boolean

'Determine if this Tx has a lab case
Dim sSQL As String
Dim rs As Recordset

    sSQL = "SELECT TX_CODE_ID From LAB_TX_SPECIFICATIONS" _
        & " Where (TX_CODE_ID = " & mvarTxCodeID & ")"
    Set rs = goconn.OpenRecordset(sSQL)
    If Not rs.EOF Then
        HasLabCase = True
    Else
        HasLabCase = False
    End If
    rs.Close

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

Public Property Get ModifiedInDentalEncounterID() As Long

    ModifiedInDentalEncounterID = mvarModifiedInDentalEncounterID
    
End Property

Friend Property Let ModifiedInDentalEncounterID(vData As Long)

    mvarModifiedInDentalEncounterID = vData
    
End Property

Friend Property Let ModifiedDateTime(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CreateDateTime = 5
    SetDirtyFlag
    mvarModifiedDateTime = vData
End Property

Public Property Get ModifiedDateTime() As Date
    
    ModifiedDateTime = mvarModifiedDateTime
    
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
Dim bChanged As Boolean

Dim oLoc As clsLocation
Dim oSubLoc As clsSubLocation
Dim oSS As clsSurfaceSet
Dim bSomethingChanged As Boolean

    'Tracker 21622
    'If the status was deleted from this encounter then don't display anything
    If Me.StatusID = txdeleted And Me.DentalEncounterID = Me.CurrentDentalEncounterID Then
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
        sTemp = sTemp & GetNoteText("PHASE")
        sTemp = sTemp & GetNoteText("LOCATION_NAME")
        sTemp = sTemp & GetNoteText("SUB_LOCATION_NAME")
        sTemp = sTemp & GetNoteText("TOOTH_ID")
        sTemp = sTemp & GetNoteText("SURFACEABBR")
        sTemp = sTemp & GetNoteText("ADA_DEFINITION")
        sTemp = sTemp & GetNoteText("STATUS_ID")
        sTemp = sTemp & GetNoteText("PROGNOSIS_ID")
        sTemp = sTemp & GetNoteText("COMMENTS")
        sTemp = sTemp & GetNoteText("MATERIALS")
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
        
        Case "ADA_DEFINITION"
            GetNoteText = ", " & mvarADADefinition
            
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
                        If Me.Locations(1).SubLocations.Count > 0 Then
                            If Me.Locations(1).SubLocations(1).SubLocationName <> "" Then
                                GetNoteText = ", " & Me.Locations(1).SubLocations(1).SubLocationName
                            End If
                        End If
                    End If
                End If
            End If
            
        Case "TOOTH_ID"
            If mvarChartable = sCharts Then
                'Chartable Tx
                If mvarSurfaceSets.Count > 0 Then
                    'Get the tooth
                    If mvarSurfaceSets(1).Teeth.Count > 0 Then
                        GetNoteText = " #" & mvarSurfaceSets(1).Teeth(1).ToothName
                    End If
                End If
            End If

        Case "SURFACEABBR"
            If mvarChartable = sCharts Then
                'Chartable Tx
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
        
        Case "PROGNOSIS_ID"
            'Don't show "Prognosis:" if the user selected None=<blank>.
            If Trim(mvarPrognosis) <> "" Then
                GetNoteText = " Prognosis: " & mvarPrognosis
            End If
    
       Case "COMMENTS"
            If mvarComments <> "" Then
                GetNoteText = ", " & mvarComments
            End If
       
       Case "MATERIALS"
            If mvarMaterials <> "" Then
                GetNoteText = ", " & mvarMaterials
            End If
       
       Case "PHASE"
            If mvarPhase <> "" Then
                GetNoteText = ", " & mvarPhase
            End If
       Case Else
            MsgBxARMd "Wrong parameter in Tx.GetNoteText"
            
    End Select
    
End Function




Friend Property Let IsPerio(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IsPerio = 5
    
    'The user can't change this.

    mvarIsPerio = vData
    
End Property


Public Property Get IsPerio() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IsPerio
    IsPerio = mvarIsPerio
End Property

Friend Property Let IsFPD(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IsPerio = 5
    
    'The user can't change this.

    mvarIsFPD = vData
    
End Property


Public Property Get IsFPD() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IsPerio
    IsFPD = mvarIsFPD
End Property
Friend Property Let IsRPD(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IsPerio = 5
    
    'The user can't change this.

    mvarIsRPD = vData
    
End Property


Public Property Get IsRPD() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IsPerio
    IsRPD = mvarIsRPD
End Property

Friend Property Let IsCompleteDenture(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IsPerio = 5
    
    'The user can't change this.

    mvarIsCompleteDenture = vData
    
End Property


Public Property Get IsCompleteDenture() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IsPerio
    IsCompleteDenture = mvarIsCompleteDenture
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




Friend Property Let TxsRef(vData As Collection)

    Set mvarTxsRef = vData
    
End Property


Friend Property Get TxsRef() As Collection

    Set TxsRef = mvarTxsRef

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



Public Property Let InternalAdd(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.InternalAdd = 5
    mvarInternalAdd = vData
End Property


Public Property Get InternalAdd() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.InternalAdd
    InternalAdd = mvarInternalAdd
End Property



Friend Property Let ToothID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ToothID = 5
    mvarToothID = vData
End Property


Friend Property Let WorkloadID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ToothID = 5
    mvarWorkloadID = vData
    SetDirtyFlag
End Property

Friend Property Get ToothID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ToothID
    ToothID = mvarToothID
End Property


Friend Property Get WorkloadID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ToothID
    WorkloadID = mvarWorkloadID
End Property

Public Property Let Materials(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Materials = 5
    mvarMaterials = Left(vData, 255)
    SetDirtyFlag
End Property


Public Property Get Materials() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Materials
    Materials = mvarMaterials
End Property



Public Property Let Comments(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Comments = 5
    mvarComments = Left(vData, 255)
    SetDirtyFlag
End Property


Public Property Get Comments() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Comments
    Comments = mvarComments
End Property



Private Property Let Prognosis(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Prognosis = 5
    mvarPrognosis = vData
End Property


Private Property Let Phase(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Prognosis = 5
    mvarPhase = vData
End Property



Public Property Get Prognosis() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Prognosis
    Prognosis = mvarPrognosis
End Property



Public Property Get Phase() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Prognosis
    Phase = mvarPhase
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



Public Property Get Signed() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Status
    Signed = mvarSigned
End Property

Private Property Let Source(ByVal vData As SourceEnum)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Source = Form1
    mvarSource = vData
End Property


Private Property Get Source() As SourceEnum
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Source
    Source = mvarSource
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
'        mvarLocations.RefreshLocationsFromTxCodeID mvarTxCodeID, mvarChartable
    End If
    
    Set Locations = mvarLocations
    
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


Public Property Get Chartable() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Chartable
    Chartable = mvarChartable
End Property



Public Property Let PrognosisID(ByVal vData As PrognosisEnum)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PrognosisID = 5
    
    'Set the prognosis string
    mvarPrognosis = LoadResString(vData)
    
    mvarPrognosisID = vData
    SetDirtyFlag
End Property


Public Property Let PhaseID(ByVal vData As TxPhaseEnum)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PrognosisID = 5
    
    'Set the prognosis string
    If vData = -1 Then
        vData = tpNone
    End If
    
    If vData = tpNone Then
        mvarPhase = ""
    Else
        mvarPhase = LoadResString(vData)
    End If
    
    mvarPhaseID = vData
    SetDirtyFlag
End Property



Public Property Get PrognosisID() As PrognosisEnum
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PrognosisID
    PrognosisID = mvarPrognosisID
End Property



Public Property Get PhaseID() As TxPhaseEnum
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PrognosisID
    PhaseID = mvarPhaseID
End Property


'Friend Property Set Dxs(ByVal vData As colDxs)
''used when assigning an Object to the property, on the left side of a Set statement.
''Syntax: Set x.Dxs = Form1
'    Set mvarDxs = vData
'End Property


'Public Property Get Dxs() As colDxs
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Dxs
'
'    If mvarDxs Is Nothing Then
'        'Get all of the Txs for this Dx
'        Set mvarDxs = New colDxs
'
'        mvarDxs.DxsRef = mvarDxsRef
'        mvarDxs.TxsRef = mvarTxsRef
'
'        mvarDxs.RefreshDxsFromTxCodeID mvarTxCodeID, mvarSource, mvarCurrentDentalEncounterID
'
'    End If
'
'    Set Dxs = mvarDxs
'End Property








Friend Property Let DWV(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DWV = 5
    mvarDWV = vData
    SetDirtyFlag
End Property


Private Sub SetDirtyFlag()
    
    'This is called when a user sets a property
    'If the current flag is dfAdd, then don't change it until SaveTx is called
    'otherwise, set it to sfUpdate
    
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

Public Property Get DWV() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DWV
    DWV = mvarDWV
End Property



Friend Property Let LabFlag(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LabFlag = 5
    mvarLabFlag = vData
End Property


Public Property Get LabFlag() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LabFlag
    LabFlag = mvarLabFlag
End Property



Friend Property Let ChartRelated(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ChartRelated = 5
    mvarChartRelated = vData
End Property


Public Property Get ChartRelated() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ChartRelated
    ChartRelated = mvarChartRelated
End Property



Friend Property Let ADAHelpNarative(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ADAHelpNarative = 5
    mvarADAHelpNarative = vData
End Property


Public Property Get ADAHelpNarative() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ADAHelpNarative
    ADAHelpNarative = mvarADAHelpNarative
End Property



Friend Property Let ADAAlias(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ADAAlias = 5
    mvarADAAlias = vData
End Property


Friend Property Get CurrentDentalEncounterID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DentalEncounterID
    CurrentDentalEncounterID = mvarCurrentDentalEncounterID
End Property

Friend Property Let CurrentDentalEncounterID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DentalEncounterID = 5
Dim oLoc As clsLocation
Dim oSubLoc As clsSubLocation
Dim oSS As clsSurfaceSet
    
    If vData = 0 Then
        MsgBxARMd "Waring!!! The CurrentDentalEncounterID is being set to 0"
    End If
    
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

Public Property Let Signed(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Chartable = 5
    mvarSigned = vData
    SetDirtyFlag
End Property

Public Property Get ADAAlias() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ADAAlias
    ADAAlias = mvarADAAlias
End Property



Friend Property Let ADADefinition(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ADADefinition = 5
    mvarADADefinition = vData
End Property


Public Property Get ADADefinition() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ADADefinition
    ADADefinition = mvarADADefinition
End Property



Friend Property Let ADACode(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ADACode = 5
    mvarADACode = vData
End Property


Public Property Get ADACode() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ADACode
    ADACode = mvarADACode
End Property



Friend Property Let ADACategoryCode(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ADACategoryCode = 5
    mvarADACategoryCode = vData
End Property


Public Property Get ADACategoryCode() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ADACategoryCode
    ADACategoryCode = mvarADACategoryCode
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



Public Property Let CreateDateTime(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CreateDateTime = 5
    mvarCreateDateTime = vData
    SetDirtyFlag
End Property


Public Property Get CreateDateTime() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CreateDateTime
    CreateDateTime = mvarCreateDateTime
End Property



Public Property Let StatusID(ByVal vData As TxStatusEnum)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.StatusID = 5

Dim sSQL As String
Dim rs As Recordset

    If vData <> -1 Then
        'If the object has been deleted, but the user wants to change the status to
        ' new or something other than deleted, then mark the object as being added
        ' back again.
        If vData <> txdeleted And mvarStatusID = txdeleted And mvarDirtyFlag = dfNoAction Then
            mvarDirtyFlag = dfAdd
        End If
        
        If vData = txdeleted Then
            If mvarDentalEncounterID = mvarCurrentDentalEncounterID Then
                'The Tx is from this encounter
                If mvarSigned = False Then
                    'Determine if the Tx has an associated Lab entry
                    If HasLabCase = True Then
                        'There is a lab associated, so don't let the user delete it.
                        MsgBxARMd "You cannot delete a treatment once lab work has been associated with a procedure.", vbCritical, "Delete procedure"
                        Exit Property
                    End If
                    
                    'If there is a more recent encounter, then don't remove the Tx from
                    ' the database
                    sSQL = "SELECT MAX(DENTAL_ENCOUNTER_ID) AS MAXDE From DENTAL_ENCOUNTER" _
                        & " WHERE (PATIENT_UNIT_ID IN (SELECT PATIENT_UNIT_ID" _
                        & " From DENTAL_ENCOUNTER WHERE (DENTAL_ENCOUNTER_ID = " & mvarCurrentDentalEncounterID & ")))" _
                        & " Having (Max(DENTAL_ENCOUNTER_ID) > " & mvarCurrentDentalEncounterID & ")"
                    Set rs = goconn.OpenRecordset(sSQL)
                    If rs.EOF Then
                        'The Dx is unsigned and does not have a subsequent encounter
                        ' , so remove the Tx from the database
                        mvarDirtyFlag = dfdelete
                    Else
                        'The Dx is unsigned, but does have a subsequent encounter
                        ' , so mark the Tx as deleted
                        'It will be, so do nothing
                    End If
                    rs.Close
                    Set rs = Nothing
                    
                Else
                    'The Tx is signed, so set the statusid to deleted
                    'It already will be so do nothing
                End If
            Else
                'The Tx is from an old encounter, and it doesn't matter if the old
                ' encounter is signed or not. So, set the statusid to deleted
                'It already will be so do nothing
            End If
        End If
        
        '--- SCR-ARM-24406 RWalker 1/13/03 - Converted to a Complex Tx type
        If (mvarDentalEncounterID = mvarCurrentDentalEncounterID) And (mvarStatusID <> vData) And (vData <> txdeleted) Then
            goARM.StopCmplxTx gsARMTx_UpdateDentalGCTxStatus, enmARMTxStatusOk, enmARMStop_MeOnly, App.EXEName
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


Public Property Get StatusID() As TxStatusEnum
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.StatusID
    StatusID = mvarStatusID
End Property



Friend Function SaveTx(GroupID As Long) As Boolean

'This function saves the Tx to the Dental_Encounter_Tx table

Dim sSQL As String
Dim rs As Recordset
Dim sOraTimeFormat As String

On Error GoTo SaveTxErr

    SaveTx = False

    If InStr(UCase(mvarCreateDateTime), "AM") Or InStr(UCase(mvarCreateDateTime), "PM") Then
        'The time is in civilian time
        sOraTimeFormat = "MM/DD/RRRR HH:MI:SS AM"
    Else
        'The time is in military time
        sOraTimeFormat = "MM/DD/RRRR HH24:MI:SS"
    End If
    
    'Make sure the date is changed to reflect the update date/time
'     mvarCreateDateTime = GetSysDate
    
    'Calculate workload
    If mvarDirtyFlag = dfAdd Or mvarDirtyFlag = dfUpdate Then
        If mvarStatusID = txCompleted And mvarTxADACodeID <> GTRMembraneRemoval Then 'Don't give workload for GTRMembraneRemoval when it is set to completed
            If mvarWorkloadID = 0 Then
                'Add the workload
                UpdateWorkload mvarWorkloadID, dfAdd, mvarCurrentDentalEncounterID, mvarTxADACodeID, 1
            Else
                'Update the workload
                'This shouldn't happen for a different encounter than the Add workload because
                ' it won't show up in the next encounter
                UpdateWorkload mvarWorkloadID, dfUpdate, mvarCurrentDentalEncounterID, mvarTxADACodeID, 1
            End If
        'For SCR 17174
        ElseIf mvarTxADACodeID = GTRMembraneRemoval And mvarStatusID = txInProgress Then
            If mvarWorkloadID = 0 Then
                'Add the workload
                UpdateWorkload mvarWorkloadID, dfAdd, mvarCurrentDentalEncounterID, mvarTxADACodeID, 1
            Else
                'Update the workload
                UpdateWorkload mvarWorkloadID, dfUpdate, mvarCurrentDentalEncounterID, mvarTxADACodeID, 1
            End If
        'Don't change the workload for a tx that is signed.
        ElseIf mvarSigned = False Then
            'Remove any workload
            UpdateWorkload mvarWorkloadID, dfdelete, mvarCurrentDentalEncounterID, mvarTxADACodeID, 1
        End If
    
    ElseIf mvarDirtyFlag = dfdelete Then
        'Delete any workload
        UpdateWorkload mvarWorkloadID, dfdelete, mvarCurrentDentalEncounterID, mvarTxADACodeID, 1
    End If
    
    
    Select Case mvarDirtyFlag
    
        Case DirtyFlagEnum.dfNoAction
            'The Dx has not changed since the last save so do nothing
            
        Case DirtyFlagEnum.dfUpdate
            'Update the database
            
            'This is updated when a property of the Tx object is changed.
            'mvarModifiedDateTime = GetSysDate
            If mvarTxCodeID = 0 Then
                MsgBxARMd "WARNING: mvarTxCodeID = 0 and it is trying to update the database with it. Aborting update. Please tell Scott how you encountered this. And yes this message will be removed when we solve the problem."
                Exit Function
            End If
            
            sSQL = "UPDATE DENTAL_ENCOUNTER_TX SET" _
                & " TX_ADA_CODE_ID = " & mvarTxADACodeID & "," _
                & " CREATE_DATE_TIME = " & goChcsiiConn.FormatDateForDBInsert(mvarCreateDateTime) & "," _
                & " MODIFIED_DATE_TIME = " & goChcsiiConn.FormatDateForDBInsert(mvarModifiedDateTime) & "," _
                & " PROVIDER_ID = " & mvarProviderID & "," _
                & " DENTAL_ENCOUNTER_ID = " & mvarDentalEncounterID & "," _
                & " STATUS_ID = " & mvarStatusID & "," _
                & " CHARTABLE = " & CInt(mvarChartable) & "," _
                & " COMMENTS = '" & RQ(mvarComments) & "'," _
                & " MATERIALS = '" & RQ(mvarMaterials) & "'," _
                & " PROGNOSIS_ID = " & mvarPrognosisID & "," _
                & " MODIFIED_IN_DENTAL_ENCOUNTER = " & mvarModifiedInDentalEncounterID & "," _
                & " PHASE = " & CInt(mvarPhaseID) & "," _
                & " WORKLOAD_ID = " & IIf(CLng(mvarWorkloadID) = 0, "NULL", CLng(mvarWorkloadID)) & "," _
                & " DELETE_REASON = " & mvarDeleteReason & "," _
                & " SIGNED = " & CInt(mvarSigned) & "" _
                & " Where (TX_CODE_ID = " & mvarTxCodeID & ")"
            goconn.ExecuteSQL sSQL
            
            'Update the groups table
            sSQL = "Update DENTAL_ENCOUNTER_GROUPS" _
                & " Set GROUP_ID = " & mvarGroupID & "" _
                & " Where (TX_CODE_ID = " & mvarTxCodeID & ")"
            goconn.ExecuteSQL sSQL
            
        Case DirtyFlagEnum.dfAdd
            'Add the Tx to the database
            
            'This is updated when a property of the Tx object is changed.
            'mvarModifiedDateTime = GetSysDate
            
            If gbUseStoredProc = True Then
                
                Dim oCmd As CHCSII_Command.CHCSII_Cmd
                Set oCmd = New CHCSII_Command.CHCSII_Cmd
                
                oCmd.SPName = "DL_NOTE_TX_INSERT"
    
                'Call oCmd.AddParam(False, False, "textString", mvarTxCodeID)
                Call oCmd.AddParam(False, False, "textString", mvarTxADACodeID)
                Call oCmd.AddParam(False, False, "textString", mvarCreateDateTime)
                Call oCmd.AddParam(False, False, "textString", mvarModifiedDateTime)
                Call oCmd.AddParam(False, False, "textString", mvarProviderID)
                Call oCmd.AddParam(False, False, "textString", mvarDentalEncounterID)
                Call oCmd.AddParam(False, False, "textString", mvarStatusID)
                Call oCmd.AddParam(False, False, "textString", CInt(mvarChartable))
                Call oCmd.AddParam(False, False, "textString", RQ(mvarComments))
                Call oCmd.AddParam(False, False, "textString", RQ(mvarMaterials))
                Call oCmd.AddParam(False, False, "textString", mvarPrognosisID)
                Call oCmd.AddParam(False, False, "textString", mvarModifiedInDentalEncounterID)
                Call oCmd.AddParam(False, False, "textString", CInt(mvarPhaseID))
                Call oCmd.AddParam(False, False, "textString", IIf(CLng(mvarWorkloadID) = 0, "", CLng(mvarWorkloadID)))
                Call oCmd.AddParam(False, False, "textString", mvarDeleteReason)
                Call oCmd.AddParam(False, False, "textString", CInt(mvarSigned))
                Call oCmd.AddParam(False, False, "textString", mvarGroupID)
                Call oCmd.AddParam(False, False, "textString", sOraTimeFormat)
                Call oCmd.AddParam(True, True, "", "")
                
                Set rs = goconn.OpenRecordsetSP(oCmd)
                'goconn.ExecuteSP oCmd
                
                If Not rs.EOF Then
                    mvarTxCodeID = rs("ID")
                Else
                    MsgBxARMd "Fatal error, could not get the identity value from the SP."
                    Exit Function
                End If

            Else
                'Get the sequence
'                sSQL = "SELECT DENTAL_ENCOUNTER_TX_SEQ.NEXTVAL AS PKID FROM DUAL"
'                Set rs = goconn.OpenRecordset(sSQL)
'                If rs.EOF Then
'                    Exit Function
'                End If
'                mvarTxCodeID = rs("PKID")
'                rs.Close
'                Set rs = Nothing
                mvarTxCodeID = goconn.GetNextID(esDENTAL_ENCOUNTER_TX_SEQ)
            
                sSQL = "INSERT INTO DENTAL_ENCOUNTER_TX" _
                    & " (TX_CODE_ID," _
                    & " TX_ADA_CODE_ID," _
                    & " CREATE_DATE_TIME," _
                    & " MODIFIED_DATE_TIME," _
                    & " PROVIDER_ID," _
                    & " DENTAL_ENCOUNTER_ID," _
                    & " STATUS_ID," _
                    & " CHARTABLE," _
                    & " COMMENTS," _
                    & " MATERIALS," _
                    & " PROGNOSIS_ID," _
                    & " MODIFIED_IN_DENTAL_ENCOUNTER," _
                    & " PHASE," _
                    & " WORKLOAD_ID," _
                    & " DELETE_REASON," _
                    & " SIGNED)"
                sSQL = sSQL & " VALUES (" _
                    & "" & mvarTxCodeID & "," _
                    & " " & mvarTxADACodeID & "," _
                    & " " & goChcsiiConn.FormatDateForDBInsert(mvarCreateDateTime) & "," _
                    & " " & goChcsiiConn.FormatDateForDBInsert(mvarModifiedDateTime) & "," _
                    & " " & mvarProviderID & "," _
                    & " " & mvarDentalEncounterID & "," _
                    & " " & mvarStatusID & "," _
                    & " " & CInt(mvarChartable) & "," _
                    & " '" & RQ(mvarComments) & "'," _
                    & " '" & RQ(mvarMaterials) & "'," _
                    & " " & mvarPrognosisID & "," _
                    & " " & mvarModifiedInDentalEncounterID & "," _
                    & " " & CInt(mvarPhaseID) & "," _
                    & " " & IIf(CLng(mvarWorkloadID) = 0, "NULL", CLng(mvarWorkloadID)) & "," _
                    & " " & mvarDeleteReason & "," _
                    & " " & CInt(mvarSigned) & ")"
                
                goconn.ExecuteSQL sSQL
                
                'Add the new tx to the groups table
                sSQL = "INSERT INTO DENTAL_ENCOUNTER_GROUPS" _
                    & " (GROUP_ID, TX_CODE_ID, DX_CODE_ID)" _
                    & " VALUES (" & mvarGroupID & ", " & mvarTxCodeID & ",0)"
                goconn.ExecuteSQL sSQL
            End If
            
        Case DirtyFlagEnum.dfdelete
            If mvarTxCodeID <> 0 Then
                'Delete the Tx and it's associated data
                'Delete the sublocation first
                
                sSQL = "DELETE FROM NOTE_TX_SUB_LOCATION" _
                    & " WHERE TX_LOCATION_ID IN" _
                    & " (SELECT TX_LOCATION_ID" _
                    & " From NOTE_TXS_LOCATION" _
                    & " WHERE (TX_CODE_ID = " & mvarTxCodeID & "))"
                goconn.ExecuteSQL sSQL
                
                'Delete the location next
                sSQL = "DELETE FROM NOTE_TXS_LOCATION" _
                    & " Where (TX_CODE_ID = " & mvarTxCodeID & ")"
                goconn.ExecuteSQL sSQL
                
                'Delete the Dx group info
                sSQL = "DELETE FROM DENTAL_ENCOUNTER_GROUPS" _
                    & " Where (TX_CODE_ID = " & mvarTxCodeID & ")"
                goconn.ExecuteSQL sSQL
                
                'Delete the Surfaceset relationship to the TX
                sSQL = "DELETE FROM DENTAL_ENCOUNTER_TXS_SS" _
                    & " Where (TX_CODE_ID = " & mvarTxCodeID & ")"
                goconn.ExecuteSQL sSQL
                
                'Delete the TX
                sSQL = "DELETE FROM DENTAL_ENCOUNTER_TX" _
                    & " Where (TX_CODE_ID = " & mvarTxCodeID & ")"
                goconn.ExecuteSQL sSQL
                
                
                'Set the TXCodeID to 0 because it does not exist in the db anymore
                mvarTxCodeID = 0
                
            End If
            
    End Select

'    'This must be here because if two references are made to it, it only needs to save the
'    'Tx once, but it must add the relationships in both (or more) cases.
'    'Add the Tx to the DxCodeID. Only applicable if this Tx is in a collection that is
'    'under a Dx.
'    If DxCodeID <> 0 And mvarTxCodeID <> 0 Then
'        'Find out if the relationship exists
'        sSQL = "SELECT DX_CODE_ID, TX_CODE_ID" _
'            & " From DENTAL_ENCOUNTER_TXS_DXS" _
'            & " WHERE (DX_CODE_ID = " & DxCodeID & ") AND (TX_CODE_ID = " & mvarTxCodeID & ")"
'        Set rs = goconn.OpenRecordset(sSQL)
'        If rs.EOF Then
'            'Add the relationship
'            sSQL = "INSERT INTO DENTAL_ENCOUNTER_TXS_DXS" _
'                & " (DX_CODE_ID, TX_CODE_ID)" _
'                & " VALUES (" & DxCodeID & ", " & mvarTxCodeID & ")"
'            goconn.ExecuteSQL sSQL
'        End If
'        rs.Close
'    End If


    Set rs = Nothing
    SaveTx = True
    
'    'Save the Dxs if needed
'    If Not mvarDxs Is Nothing Then
'        SaveTx = mvarDxs.SaveDxs(mvarTxCodeID)
'    End If
    
    'Save the SurfaceSets if needed
    If Not mvarSurfaceSets Is Nothing And mvarTxCodeID <> 0 Then
        SaveTx = mvarSurfaceSets.SaveSurfaceSets(0, mvarTxCodeID)
    End If

    'Save the Locations if needed
    If Not mvarLocations Is Nothing And mvarTxCodeID <> 0 Then
        SaveTx = mvarLocations.SaveLocations(0, mvarTxCodeID)
    End If

    mvarDirtyFlag = dfNoAction

Exit Function

SaveTxErr:
    If InStr(Err.Description, "ORA-00100: no data found Rollback work ORA-00100: no data found") > 0 Then
    'I can't use the err.number because chcsii sql always returns the same error number for all errors
    'If Err.Number = -2147221503 Then
        'The db server reported "No data found", so continue
        Resume Next
    ElseIf InStr(Err.Description, "ORA-02292: integrity constraint (DENTAL.TX_CODE_ID_FK) violated - child record found") > 0 Then
        'There is a lab associated with this Dx, so the user can't delete it.
        'Don't do anything because the object has not been saved yet.
        MsgBxARMd "A dental lab exists for " & mvarADACode & ". You must first delete the lab before you can delete the Tx. Please change the status of the treatment from 'Deleted'.", vbInformation, "SaveTx"
    Else
        MsgBxARMd Err.Description & " -- " & sSQL, vbCritical, Err.Source
        'Err.Raise Err.Number, Err.Source, Err.Description
        'Resume
    End If

End Function




Public Property Set SurfaceSets(ByVal vData As colSurfaceSets)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.SurfaceSets = Form1
    Set mvarSurfaceSets = vData
End Property


Public Property Get SurfaceSets() As colSurfaceSets
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SurfaceSets
    
    If mvarSurfaceSets Is Nothing Then
        'Get all of the Txs for this Dx
        Set mvarSurfaceSets = New colSurfaceSets
'        mvarSurfaceSets.RefreshSurfaceSetsFromDxORTxCodeID 0, mvarTxCodeID, mvarChartable
    End If
    
    Set SurfaceSets = mvarSurfaceSets
End Property

Public Property Let TxADACodeID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TxADACodeID = 5

'Refresh the data from the DX_ICD_CODE_DEFINITION table

Dim sSQL As String
Dim rs As Recordset

    If mvarInternalAdd = True Then
        'Set the property
        mvarTxADACodeID = vData
        Exit Property
    Else
        'Update the fields
        sSQL = "SELECT ADA_CATEGORY_CODE," _
            & " ADA_DEFINITION, ADA_ALIAS, CHART_RELATED, LABFLAG," _
            & " ADA_HELP_NARATIVE , ADA_CODE, PERIO, DWV, FPD_FLAG, RPD_FLAG, DENTURE_FLAG" _
            & " From TX_ADA_CODE_DEFINITION" _
            & " WHERE (TX_ADA_CODE_ID = " & vData & ")"
        
        Set rs = New Recordset
        Set rs = goconn.OpenRecordset(sSQL)
        
        If Not rs.EOF Then
                mvarADACategoryCode = rs("ADA_CATEGORY_CODE") & ""
                mvarADADefinition = rs("ADA_DEFINITION") & ""
                mvarADAAlias = rs("ADA_ALIAS") & ""
                mvarChartRelated = Val(rs("CHART_RELATED") & "")
                mvarLabFlag = Val(rs("LABFLAG") & "")
                mvarADAHelpNarative = rs("ADA_HELP_NARATIVE") & ""
                mvarADACode = rs("ADA_CODE") & ""
                mvarIsPerio = CBool(Val(rs("PERIO") & ""))
                mvarDWV = Val(rs("DWV") & "")
                mvarIsFPD = CBool(Val(rs("FPD_FLAG") & ""))
                mvarIsRPD = CBool(Val(rs("RPD_FLAG") & ""))
                mvarIsCompleteDenture = CBool(Val(rs("DENTURE_FLAG") & ""))
        
        
        Else
            'Err.Raise vbObjectError + 513, "Property Let TxADACodeID", "Invalid ADACodeID"
            MsgBxARMd "Invalid ADACodeID", vbCritical, "Property Let TxADACodeID"
            Exit Property
        End If
        
        rs.Close
        Set rs = Nothing
        
        mvarTxADACodeID = vData
        SetDirtyFlag
    End If
    
End Property


Public Property Get TxADACodeID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TxADACodeID
    TxADACodeID = mvarTxADACodeID
End Property



Friend Property Let DentalEncounterID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DentalEncounterID = 5
    If vData = 0 Then
        MsgBxARMd "Warning!!! The DentalEncounterID is being set to 0."
    End If
    
    mvarDentalEncounterID = vData
    SetDirtyFlag
    
End Property


Public Property Get DentalEncounterID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DentalEncounterID
    DentalEncounterID = mvarDentalEncounterID
End Property



Friend Property Let TxCodeID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TxCodeID = 5
    mvarTxCodeID = vData
End Property


Public Property Get TxCodeID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TxCodeID
    TxCodeID = mvarTxCodeID
End Property



Private Sub Class_Terminate()

    Set mvarSurfaceSets = Nothing
    Set mvarLocations = Nothing
    Set mvarTxsRef = Nothing
    Set mvarDxsRef = Nothing

End Sub


