VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarPatientID As Long
Private mvarDxs As colDxs
Private mvarTxs As colTxs
Private mvarCurrentDentalEncounterID As Long
Private mvarProviderID As Long
Private mvarGroups As colGroups
Private moRadOrders As colRadOrders

'Collection to hold all Txs to ensure "unique" Tx data integrity
Private mvarTxsRef As New Collection

'Collection to hold all Dxs to ensure "unique" Dx data integrity
Private mvarDxsRef As New Collection

'Enum for the status of Dxs, Txs and Prognosis

Public Enum DxStatusEnum
    'For Dxs in the DENTAL_ENCOUNTER_DX table only
    dxActiveNew = 10
    dxActiveOngoing = 11
    dxActiveProvisional = 12
    dxActiveFinding = 13
    dxActiveObserve = 14
    dxNonActiveResolved = 20
    dxNonActiveDeleted = 21
    'dxNonActiveReplaced = 22
    dxNonActiveConfirmedNegative = 23
    dxNonActiveUnableToConfirm = 24
End Enum

Public Enum TxStatusEnum
    'For Txs in the DENTAL_ENCOUNTER_TX table only
    txPlanned = 100
    txInProgress = 101
    txCompleted = 102
    txPreExisting = 103
    txdeleted = 104
    txRevised = 105
End Enum

Public Enum TxPhaseEnum
    'For the treatment phase in the DENTAL_ENCOUNTER_TX table only
    tpNone = 0
    tpEmergency = 301
    tpDiagnosis = 302
    tpCariesControl = 303
    tpPerioControl = 304
    tpRestorative = 305
    tpMaintenance = 306
End Enum

Public Enum PrognosisEnum
    'For prognosis in the DENTAL_ENCOUNTER_TX table only
    pExcellent = 0
    pGood = 1
    pFair = 2
    pPoor = 3
    pHopeless = 4
    pNone = 5
End Enum

Public Enum DirtyFlagEnum
    'For objects only
    dfNoAction = 0
    dfUpdate = 1
    dfAdd = 2
    dfdelete = 3
End Enum

Public Enum DxPerioStateEnum
    DxPerioPrimary = 200
    DxPerioSecondary = 201
    DxNotPerio = 202
End Enum

'This is to distinguish from charts and notes for the datalayer only
Public Enum SourceEnum
    sCharts = -1
    sNotes = 0
'    sNotesMaterials = 1
'    sReports = 2
End Enum

Public Enum DeleteReasonEnum
    drNone = 5
    drChartedIncorrectly = 6
    drDisagree = 7
    drTimeLapse = 8
    drAdditionalInfo = 9
    drExtraction = 400
End Enum


'Dental Imaging
Private moRads As colRads

'These are coordinated with the constants in Lib.bas
Public Enum eRadStatus
    Ordered = 501
    Taken = 502
    Examined = 503
    Signed = 504
    Deleted = 505
End Enum

'Insync with cRAD_TYPE_VALUES in lib.bas
Public Enum eRadType
    BiteWing = 601
    PA = 602
    Pano = 603
    FMSeries = 604
    Other = 605
    HzBiteWing = 606
    VtBiteWing = 607
End Enum

'This is used in the Find method of colRads
Public Type tRadFilter
    RadStartDate As String
    RadEndDate As String
    Version As String
    RadType As String
    Status As String
    Location As String
    Comments As String
    FacilityNCID As String
    RequestedByNCID As String
    TakenBy As String
End Type

'*****************************************
'Used for Rad Orders
'*****************************************

'In sync with cRAD_ORDER_PRIORITY_VALUES
Public Enum eRadOrderPriority
    Routine = 701
    PreOp = 702
    ASAP = 703
    STAT = 704
    Notify = 705
End Enum

'Used for Rad Order statuses
'In sync with cRAD_ORDER_STATUS_VALUES
Public Enum eRadOrderStatus
    Planned = 801
    Active = 802
    InProgress = 803
    Scheduled = 804
    Deleted = 805
    CompletedAddr = 806
    CompletedUnAddr = 807
    Cancelled = 808
End Enum

'Tied to the Note Location table
Public Enum eRadOrderLocations
    eOralCavity = 1
    eBoth = 109
    eLeft = 110
    eRight = 111
    eOther = 108
End Enum
Public Property Get AllDxs() As Collection

    'Returns a collection of all the Dxs
    
    Set AllDxs = mvarDxsRef
    
End Property

Public Property Get AllTxs() As Collection
    
    'Returns a collection of all the Txs
    
    Set AllTxs = mvarTxsRef
    
End Property


Public Property Get CWComm() As MMMHISComm.Comm
    
    Set CWComm = cwobjCWComm

End Property

Public Property Get CWConfig() As MMMHISConfiguration.Config

    Set CWConfig = cwobjConfig
    
End Property

Public Property Get CWLogonEnviron() As MMMHISLogon.Logon
    
    Set CWLogonEnviron = cwobjLogonEnviron

End Property

Public Property Get CWPatient() As MMMHISPatient.Patient
    
    Set CWPatient = cwobjPatient
    
End Property

Private Function ForensicsUnionQuery( _
    ChartableTxs As String, ChartableDxs As String, _
    NonChartableTxs As String, NonChartableDxs As String, _
    OtherCrit As String, ChartableTxVirginTeeth As String, _
    ChartableDxVirginTeeth As String, RemoveImpossibleTxs As String, _
    RemoveImpossibleDxs As String, HitCount As Integer) As Recordset

'This function creates a union query from all the sql statements passed in
' and opens the recordset

Dim sSQL As String

    'If there is nothing to search for, then don't search
    If ChartableTxs = "" And ChartableDxs = "" And NonChartableTxs = "" _
        And NonChartableDxs = "" And OtherCrit = "" And ChartableTxVirginTeeth = "" _
        And ChartableDxVirginTeeth = "" Then
        
        Set ForensicsUnionQuery = New Recordset
        Exit Function
    End If
    
    'Form the first part of the query
    sSQL = "SELECT Patient_UNIT_ID, COUNT(*) AS Matches" _
        & " FROM ("

    If ChartableTxs <> "" Then
        sSQL = sSQL & ChartableTxs & " Union All "
    End If
    
    If ChartableTxVirginTeeth <> "" Then
        sSQL = sSQL & ChartableTxVirginTeeth & " Union All "
    End If
    
    If ChartableDxs <> "" Then
        sSQL = sSQL & ChartableDxs & " Union All "
    End If
    
    If ChartableDxVirginTeeth <> "" Then
        sSQL = sSQL & ChartableDxVirginTeeth & " Union All "
    End If
    
    If NonChartableTxs <> "" Then
        sSQL = sSQL & NonChartableTxs & " Union All "
    End If
    
    If NonChartableDxs <> "" Then
        sSQL = sSQL & NonChartableDxs & " Union All "
    End If
    
    'Remove the last  " Union All "
    sSQL = Left(sSQL, Len(sSQL) - Len(" Union All "))
    
    'Close the union query
    sSQL = sSQL & ")"
    
    'Start the where clause
    sSQL = sSQL & " Where "
    
    'Add the other criteria part if it exists
    If OtherCrit <> "" Then
        sSQL = sSQL & " patient_unit_id in (" & OtherCrit & ") and "
    End If
    
    'Remove the impossible Txs
    If RemoveImpossibleTxs <> "" Then
        sSQL = sSQL & " (patient_unit_id not in (" & RemoveImpossibleTxs & ")) and "
    End If
    
    'Remove the impossible Dxs
    If RemoveImpossibleDxs <> "" Then
        sSQL = sSQL & " (patient_unit_id not in (" & RemoveImpossibleDxs & ")) and "
    End If
    
    If OtherCrit = "" And RemoveImpossibleTxs = "" And RemoveImpossibleDxs = "" Then
        'Remove the where clause
        sSQL = Left(sSQL, Len(sSQL) - Len(" Where "))
    Else
        'Remove the last " and "
        sSQL = Left(sSQL, Len(sSQL) - Len(" and "))
    End If
    
    'Add the Group by clause
    sSQL = sSQL & " GROUP BY patient_Unit_ID HAVING (COUNT(*) >= " & HitCount & ") ORDER BY Matches DESC"
   
    Set ForensicsUnionQuery = goconn.ExecuteSQL(sSQL)
    
End Function

Public Function GetNextForensicsPatientID(ProviderID As Long, Optional UnlockPatientID As Long = 0) As Long

    'This function searches the Hit_Queue table and gets the next best matched post mortem patient.
    
Dim sSQL As String
Dim rs As Recordset

On Error GoTo Error

    'Unlock the record based on the UnlockPatientID
    If UnlockPatientID <> 0 Then
        sSQL = "UPDATE HIT_QUEUE SET RECORD_LOCK = NULL" _
            & " Where (POSTMORTEMID = " & UnlockPatientID & ")"
        goconn.ExecuteSQL sSQL
    End If
    
    'Select the next best match
    sSQL = "SELECT POSTMORTEMID From HIT_QUEUE WHERE (RECORD_LOCK = 0) AND (REVIEWED = 0) OR" _
        & " (RECORD_LOCK IS NULL) AND (REVIEWED IS NULL) ORDER BY QUALITY DESC"
    
    Set rs = goconn.ExecuteSQL(sSQL)
    
    If rs.EOF Then
        GetNextForensicsPatientID = 0
    Else
        'Lock the record being retrieved
        sSQL = "Update HIT_QUEUE SET RECORD_LOCK = " & ProviderID & "," _
            & " REVIEWED = 1 Where (POSTMORTEMID = " & rs("POSTMORTEMID") & ")"
        goconn.ExecuteSQL sSQL
        
        'Return the best matched patientid
        GetNextForensicsPatientID = rs("POSTMORTEMID")
    End If
    rs.Close
    Set rs = Nothing

Exit Function

Error:
    MsgBxARMd "Error in: " & Err.Source & ", " & Err.Description

End Function

Public Function GetNote() As String

'This function gets the note from the datalayer
    
Dim oGroup As clsGroup
Dim sTemp As String

    For Each oGroup In mvarGroups
        sTemp = sTemp & oGroup.GetNote
    Next
    
    If sTemp <> "" Then
        GetNote = crtfStartTemplateName & "Diagnosis and Treatments:" & crtfEndTemplateName & sTemp
    End If
    
End Function


Private Sub IterateThroughAllObjects()

'This sub iterates through all the objects in the datalayer to "pre load" everything from the db

Dim oGroup As clsGroup
Dim oDx As clsDx
Dim oTx As clsTx
Dim oSS As clsSurfaceSet
Dim oLoc As clsLocation
Dim oSubLoc As clsSubLocation
Dim oTooth As clsTooth

For Each oGroup In mvarGroups
    For Each oDx In oGroup.Dxs
        If oDx.Chartable = True Then
            For Each oSS In oDx.SurfaceSets
                For Each oTooth In oSS.Teeth
                    'Do nothing
                Next
            Next
        Else
            For Each oLoc In oDx.Locations
                For Each oSubLoc In oLoc.SubLocations
                    'Do nothing
                Next
            Next
        End If
    Next
    For Each oTx In oGroup.Txs
        If oTx.Chartable = True Then
            For Each oSS In oTx.SurfaceSets
                For Each oTooth In oSS.Teeth
                    'Do nothing
                Next
            Next
        Else
            For Each oLoc In oTx.Locations
                For Each oSubLoc In oLoc.SubLocations
                    'Do nothing
                Next
            Next
        End If
    Next
Next


End Sub

Public Property Get RadOrders() As colRadOrders
    
    Set RadOrders = moRadOrders
    
End Property


Public Property Get Rads() As colRads

    Set Rads = moRads
    
End Property

Public Function SearchForensics(PostMortemPatientID As Long, _
    ChartableTxs As Boolean, ChartableTxSurfaceSets As Boolean, _
    NonChartableTxs As Boolean, NonChartableTxSurfaceSets As Boolean, _
    ChartableDxs As Boolean, ChartableDxSubLoc As Boolean, _
    NonChartableDxs As Boolean, NonChartableDxSubLoc As Boolean, _
    OtherCrit As String, MatchVirginTxTeeth As Boolean, _
    MatchVirginDxTeeth As Boolean, RemoveImpossibleTxs As Boolean, _
    RemoveImpossibleDxs As Boolean, Optional HitCount As Integer = 0) As Recordset

'This function searches the best match given a PostMortemPatientID

Dim sChartableTxs As String
Dim sChartableTxVirginTeeth As String
Dim sChartableDxs As String
Dim sChartableDxVirginTeeth As String
Dim sNONChartableTxs As String
Dim sNONChartableDxs As String
Dim sRemoveImpossibleDxs As String
Dim sRemoveImpossibleTxs As String

On Error GoTo Error

    Set SearchForensics = New Recordset
    
    If ChartableTxs = True Then
        'Search for Chartable Txs, down to the surfaceset level
        sChartableTxs = SearchForensicsChartableTxs(PostMortemPatientID, ChartableTxSurfaceSets)
    End If
    
    If MatchVirginTxTeeth = True Then
        'Match Tx Virgin teeth
        sChartableTxVirginTeeth = SearchForensicsChartableTxVirgin(PostMortemPatientID)
    End If
    
    If ChartableDxs = True Then
        'Search for Chartable Dxs, down to the surfaceset level
        sChartableDxs = SearchForensicsChartableDxs(PostMortemPatientID, ChartableDxSubLoc)
    End If
    
    If MatchVirginDxTeeth = True Then
        'Match Dx Virgin teeth
        sChartableDxVirginTeeth = SearchForensicsChartableDxVirgin(PostMortemPatientID)
    End If
    
    If NonChartableTxs = True Then
        'Search for NonChartable Txs, down to the sub location level
        sNONChartableTxs = SearchForensicsNONChartableTxs(PostMortemPatientID, NonChartableTxSurfaceSets)
    End If
    
    If NonChartableDxs = True Then
        'Search for NonChartable Dxs, down to the sub location level
        sNONChartableDxs = SearchForensicsNONChartableDxs(PostMortemPatientID, NonChartableDxSubLoc)
    End If
    
    'Remove antemortem patients that have a Dx, but the post mortem is virgin
    If RemoveImpossibleDxs = True Then
        sRemoveImpossibleDxs = SearchForensicsImpossibleDx(PostMortemPatientID)
    End If
    
    'Remove antemortem patients that have a Tx, but the post mortem is virgin
    If RemoveImpossibleTxs = True Then
        sRemoveImpossibleTxs = SearchForensicsImpossibleTx(PostMortemPatientID)
    End If
    
    Set SearchForensics = ForensicsUnionQuery(sChartableTxs, sChartableDxs, sNONChartableTxs, sNONChartableDxs, OtherCrit, sChartableTxVirginTeeth, sChartableDxVirginTeeth, sRemoveImpossibleTxs, sRemoveImpossibleDxs, HitCount)
    
Exit Function

Error:
    MsgBxARMd "Error in: " & Err.Source & ": SearchForensics, " & Err.Description

End Function

Public Property Set Groups(ByVal vData As colGroups)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Groups = Form1
    Set mvarGroups = vData
End Property


Public Property Get Groups() As colGroups
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Groups
    Set Groups = mvarGroups
End Property



Public Function LoadGroups(CWPatient As MMMHISPatient.Patient, CWLogonEnviron As MMMHISLogon.Logon, CWConfig As MMMHISConfiguration.Config, CWComm As MMMHISComm.Comm, CurrentDentalEncounterID As Long, ProviderID As Long) As Boolean

Dim lTempPatientID As Long
Dim lTime As Long
        
    lTime = Timer

    Set cwobjPatient = CWPatient
    Set cwobjLogonEnviron = CWLogonEnviron
    Set cwobjConfig = CWConfig
    Set cwobjCWComm = CWComm
    
    'Open the DB
    OpenDB
'cwobjPatient.Name
    'Load all the Dxs from the PatientID
    lTempPatientID = CWPatient.UnitNumber
    
    'Unload the Txs, Dxs and Refs
    Set mvarDxsRef = Nothing
    Set mvarTxsRef = Nothing
    
    Set mvarGroups = New colGroups
    
    mvarProviderID = ProviderID
    glPrimProviderID = mvarProviderID
    
    mvarGroups.DxsRef = mvarDxsRef
    mvarGroups.TxsRef = mvarTxsRef
    
    If mvarGroups.RefreshGroupsFromPatientID(lTempPatientID, CurrentDentalEncounterID) Then
        
        mvarPatientID = lTempPatientID
        mvarCurrentDentalEncounterID = CurrentDentalEncounterID
        
'        IterateThroughAllObjects
        
        LoadGroups = True
    
        If goShared.CmdLineSwitches("IMAGING") <> "" Then
            'Load imaging
            Set moRads = New colRads
            moRads.Load PatientID
            
            'Load RadOrders
            Set moRadOrders = New colRadOrders
            moRadOrders.Init mvarCurrentDentalEncounterID, CWLogonEnviron.StoringUserNCID, CWLogonEnviron.StoringUserName
            'This is loaded in order entry
            'moRadOrders.LoadRadOrders lTempPatientID, 0, 0, 0, Empty
        End If
        
    End If
    
    If goShared.CmdLineSwitches("DATALAYER") = "DATALAYER" Then
        goShared.ShowVBError 0, "Datalayer took: " & Timer - lTime & " seconds to load.", "", "Datalayer performance", 0
    End If
    
    Debug.Print "Datalayer took: " & Timer - lTime & " seconds to load."
    
End Function


Public Sub OpenDBForForensics()

    'Used by forensics to access the forensics functions without calling loadgroups
    'Open the DB
    OpenDB

End Sub

Public Property Get PatientID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PatientID
    PatientID = mvarPatientID
End Property
Public Property Get ProviderID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PatientID
    ProviderID = mvarProviderID
End Property

Public Property Get DentalEncounterID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PatientID
    DentalEncounterID = mvarCurrentDentalEncounterID
End Property

Public Property Let DentalEncounterID(vData As Long)
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PatientID
    
Dim oGroup As clsGroup
Dim oDx As clsDx
Dim oTx As clsTx
    
    mvarCurrentDentalEncounterID = vData
    
    'Loop through the groups and set the current dental encounter id in the lower objects
    If mvarGroups Is Nothing Then
        Err.Raise vbObjectError + 7, "Datalayer.DentalEncounterID Let", "The SVGComm server must load the datalayer before setting the DentalEncounterID."
        Exit Property
    End If
    
    mvarGroups.CurrentDentalEncounterID = mvarCurrentDentalEncounterID
    
    If Not moRadOrders Is Nothing Then
        moRadOrders.DentalEncounterID = vData
    End If
    
End Property

Private Property Let PatientID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PatientID = 5
    mvarPatientID = vData
End Property




Public Property Let ProviderID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PatientID = 5
    mvarProviderID = vData
    
    glPrimProviderID = mvarProviderID
    
End Property










Public Function SaveGroups() As Boolean

    'If the Groups is not nothing, then save
    If Not mvarGroups Is Nothing Then
        SaveGroups = mvarGroups.SaveGroups
    Else
        SaveGroups = True
    End If
    
    'Save Rads
    If goShared.CmdLineSwitches("IMAGING") = "IMAGING" Then
        'Save the Rads
        If Not moRads Is Nothing Then
            moRads.Save
        End If
        
        'Save the RadOrders
        If Not moRadOrders Is Nothing Then
            moRadOrders.Save
        End If
    End If
    
End Function

Private Function SearchForensicsChartableTxs( _
    PostMortemPatientID As Long, _
    SearchSurfaceSets As Boolean) As String

'This returns a query string that contains a list of patient unit ids that
' match down to the surface level a chartable treatment.

Dim rs As Recordset
Dim sSQL As String

    'Find all the Tx info for the PostMortem patient
    sSQL = "SELECT DENTAL_ENCOUNTER_TX.TX_ADA_CODE_ID, " _
        & " DENTAL_ENCOUNTER_TEETH.TOOTH_ID," _
        & " DENTAL_ENC_SURFACE_SET.SurfaceSet" _
        & " FROM DENTAL_ENCOUNTER, DENTAL_ENCOUNTER_TX," _
        & " DENTAL_ENCOUNTER_TXS_SS," _
        & " DENTAL_ENC_SURFACE_SET," _
        & " DENTAL_ENCOUNTER_TEETH" _
        & " Where DENTAL_ENCOUNTER.DENTAL_ENCOUNTER_ID = DENTAL_ENCOUNTER_TX.DENTAL_ENCOUNTER_ID" _
        & " AND DENTAL_ENCOUNTER_TX.TX_CODE_ID = DENTAL_ENCOUNTER_TXS_SS.TX_CODE_ID" _
        & " AND DENTAL_ENCOUNTER_TXS_SS.SURFACE_SET_ID = DENTAL_ENC_SURFACE_SET.SURFACE_SET_ID" _
        & " AND DENTAL_ENC_SURFACE_SET.DENTAL_ENCOUNTER_TEETH_ID" _
        & " = DENTAL_ENCOUNTER_TEETH.DENTAL_ENCOUNTER_TEETH_ID" _
        & " AND (DENTAL_ENCOUNTER.PATIENT_UNIT_ID = " & PostMortemPatientID & ") AND" _
        & " (DENTAL_ENCOUNTER_TX.STATUS_ID <> 104 AND DENTAL_ENCOUNTER_TX.STATUS_ID <> 100)"
    
    Set rs = goconn.OpenRecordset(sSQL)
    
    If rs.EOF Then
        SearchForensicsChartableTxs = ""
        Exit Function
    End If
    
    'Build the sql statment to find the related data
    sSQL = "SELECT DENTAL_ENCOUNTER.PATIENT_UNIT_ID" _
        & " FROM DENTAL_ENCOUNTER, DENTAL_ENCOUNTER_TX," _
        & " DENTAL_ENCOUNTER_TXS_SS, DENTAL_ENC_SURFACE_SET," _
        & " DENTAL_ENCOUNTER_TEETH, MMI_ID Where DENTAL_ENCOUNTER.DENTAL_ENCOUNTER_ID = DENTAL_ENCOUNTER_TX.DENTAL_ENCOUNTER_ID" _
        & " AND DENTAL_ENCOUNTER_TX.TX_CODE_ID = DENTAL_ENCOUNTER_TXS_SS.TX_CODE_ID" _
        & " AND DENTAL_ENCOUNTER_TXS_SS.SURFACE_SET_ID = DENTAL_ENC_SURFACE_SET.SURFACE_SET_ID" _
        & " AND DENTAL_ENC_SURFACE_SET.DENTAL_ENCOUNTER_TEETH_ID" _
        & " = DENTAL_ENCOUNTER_TEETH.DENTAL_ENCOUNTER_TEETH_ID" _
        & " AND DENTAL_ENCOUNTER.PATIENT_UNIT_ID = MMI_ID.UNIT_NUMBER" _
        & " AND ("
    
    'Loop through and build a sql statement to compare all known Tx data
    Do Until rs.EOF
        If SearchSurfaceSets = True Then 'Search the surface sets
            sSQL = sSQL & "(DENTAL_ENCOUNTER_TEETH.TOOTH_ID = " & rs("TOOTH_ID") & ") AND" _
                & " (DENTAL_ENCOUNTER_TX.TX_ADA_CODE_ID = " & rs("TX_ADA_CODE_ID") & ") AND" _
                & " (DENTAL_ENC_SURFACE_SET.SURFACESET = '" & rs("SURFACESET") & "') AND" _
                & " (DENTAL_ENCOUNTER_TX.STATUS_ID <> 104 AND DENTAL_ENCOUNTER_TX.STATUS_ID <> 100) AND (MMI_ID.UNIDENTIFIED_FLAG = 'A') OR "
        Else 'Don't search the surface sets
            sSQL = sSQL & "(DENTAL_ENCOUNTER_TEETH.TOOTH_ID = " & rs("TOOTH_ID") & ") AND" _
                & " (DENTAL_ENCOUNTER_TX.TX_ADA_CODE_ID = " & rs("TX_ADA_CODE_ID") & ") AND" _
                & " (DENTAL_ENCOUNTER_TX.STATUS_ID <> 104 AND DENTAL_ENCOUNTER_TX.STATUS_ID <> 100) AND (MMI_ID.UNIDENTIFIED_FLAG = 'A') OR "
        End If
        rs.MoveNext
    Loop
    rs.Close
    
    'Remove the last " OR "
    sSQL = Left(sSQL, Len(sSQL) - 4)
    
    'End the sql statement
    sSQL = sSQL & ")"
    
    SearchForensicsChartableTxs = sSQL
    
End Function


Private Function SearchForensicsChartableTxVirgin(PostMortemPatientID As Long) As String

'This returns a query string that contains a list of patient unit ids that
' match the virgin teeth.

Dim sSQLPostVT As String
Dim sSQLAnteVT As String

    'Find all the Virgin teeth for the PostMortem patient
    sSQLPostVT = "SELECT DENTAL_ENC_VIEW_TEETH.TOOTH_VIEW From DENTAL_ENC_VIEW_TEETH" _
        & " Where (PATIENT_UNIT_ID = " & PostMortemPatientID & ") And (TOOTH_VIEW <> 0)" _
        & " AND (TOOTH_VIEW NOT IN (SELECT DENTAL_ENCOUNTER_TEETH.TOOTH_ID" _
        & " FROM DENTAL_ENCOUNTER_TXS_SS, DENTAL_ENCOUNTER_TX," _
        & " DENTAL_ENC_SURFACE_SET, DENTAL_ENCOUNTER_TEETH," _
        & " DENTAL_ENCOUNTER Where DENTAL_ENCOUNTER_TXS_SS.TX_CODE_ID = DENTAL_ENCOUNTER_TX.TX_CODE_ID" _
        & " AND DENTAL_ENCOUNTER_TXS_SS.SURFACE_SET_ID = DENTAL_ENC_SURFACE_SET.SURFACE_SET_ID" _
        & " AND DENTAL_ENC_SURFACE_SET.DENTAL_ENCOUNTER_TEETH_ID" _
        & " = DENTAL_ENCOUNTER_TEETH.DENTAL_ENCOUNTER_TEETH_ID" _
        & " AND DENTAL_ENCOUNTER_TX.DENTAL_ENCOUNTER_ID = DENTAL_ENCOUNTER.DENTAL_ENCOUNTER_ID" _
        & " AND DENTAL_ENCOUNTER_TEETH.DENTAL_ENCOUNTER_ID" _
        & " = DENTAL_ENCOUNTER.DENTAL_ENCOUNTER_ID AND" _
        & " (DENTAL_ENCOUNTER_TX.STATUS_ID <> 104 AND DENTAL_ENCOUNTER_TX.STATUS_ID <> 100) AND" _
        & " (DENTAL_ENCOUNTER.PATIENT_UNIT_ID = " & PostMortemPatientID & ")))"
    
    'Find all virgin teeth for Antemortem patients
    sSQLAnteVT = "SELECT DENTAL_ENC_VIEW_TEETH.PATIENT_UNIT_ID From DENTAL_ENC_VIEW_TEETH, MMI_ID" _
        & " Where DENTAL_ENC_VIEW_TEETH.PATIENT_UNIT_ID = MMI_ID.UNIT_NUMBER" _
        & " AND (DENTAL_ENC_VIEW_TEETH.TOOTH_VIEW <> 0) AND" _
        & " (MMI_ID.UNIDENTIFIED_FLAG = 'A') AND" _
        & " (DENTAL_ENC_VIEW_TEETH.TOOTH_VIEW NOT IN" _
        & " (SELECT DENTAL_ENCOUNTER_TEETH.TOOTH_ID" _
        & " FROM DENTAL_ENCOUNTER_TXS_SS," _
        & " DENTAL_ENCOUNTER_TX, DENTAL_ENC_SURFACE_SET, DENTAL_ENCOUNTER_TEETH," _
        & " DENTAL_ENCOUNTER Where DENTAL_ENCOUNTER_TXS_SS.TX_CODE_ID = DENTAL_ENCOUNTER_TX.TX_CODE_ID" _
        & " AND DENTAL_ENCOUNTER_TXS_SS.SURFACE_SET_ID = DENTAL_ENC_SURFACE_SET.SURFACE_SET_ID" _
        & " AND DENTAL_ENC_SURFACE_SET.DENTAL_ENCOUNTER_TEETH_ID" _
        & " = DENTAL_ENCOUNTER_TEETH.DENTAL_ENCOUNTER_TEETH_ID" _
        & " AND DENTAL_ENCOUNTER_TX.DENTAL_ENCOUNTER_ID = DENTAL_ENCOUNTER.DENTAL_ENCOUNTER_ID" _
        & " AND DENTAL_ENCOUNTER_TEETH.DENTAL_ENCOUNTER_ID = DENTAL_ENCOUNTER.DENTAL_ENCOUNTER_ID AND" _
        & " (DENTAL_ENCOUNTER_TX.STATUS_ID <> 104 AND DENTAL_ENCOUNTER_TX.STATUS_ID <> 100) AND (DENTAL_ENCOUNTER.PATIENT_UNIT_ID = DENTAL_ENC_VIEW_TEETH.PATIENT_UNIT_ID)))"
           
    SearchForensicsChartableTxVirgin = sSQLAnteVT & " AND DENTAL_ENC_VIEW_TEETH.TOOTH_VIEW IN(" & sSQLPostVT & ")"
    
End Function



Private Function SearchForensicsChartableDxVirgin(PostMortemPatientID As Long) As String

'This returns a query string that contains a list of patient unit ids that
' match the virgin teeth.

Dim sSQLPostVT As String
Dim sSQLAnteVT As String

    'Find all the Virgin teeth for the PostMortem patient
    sSQLPostVT = "SELECT DENTAL_ENC_VIEW_TEETH.TOOTH_VIEW From DENTAL_ENC_VIEW_TEETH" _
        & " Where (PATIENT_UNIT_ID = " & PostMortemPatientID & ") And (TOOTH_VIEW <> 0)" _
        & " AND (TOOTH_VIEW NOT IN (SELECT DENTAL_ENCOUNTER_TEETH.TOOTH_ID" _
        & " FROM DENTAL_ENCOUNTER_DXS_SS, DENTAL_ENCOUNTER_DX," _
        & " DENTAL_ENC_SURFACE_SET, DENTAL_ENCOUNTER_TEETH," _
        & " DENTAL_ENCOUNTER Where DENTAL_ENCOUNTER_DXS_SS.DX_CODE_ID = DENTAL_ENCOUNTER_DX.DX_CODE_ID" _
        & " AND DENTAL_ENCOUNTER_DXS_SS.SURFACE_SET_ID = DENTAL_ENC_SURFACE_SET.SURFACE_SET_ID" _
        & " AND DENTAL_ENC_SURFACE_SET.DENTAL_ENCOUNTER_TEETH_ID" _
        & " = DENTAL_ENCOUNTER_TEETH.DENTAL_ENCOUNTER_TEETH_ID" _
        & " AND DENTAL_ENCOUNTER_DX.DENTAL_ENCOUNTER_ID = DENTAL_ENCOUNTER.DENTAL_ENCOUNTER_ID" _
        & " AND DENTAL_ENCOUNTER_TEETH.DENTAL_ENCOUNTER_ID" _
        & " = DENTAL_ENCOUNTER.DENTAL_ENCOUNTER_ID AND" _
        & " (DENTAL_ENCOUNTER_DX.STATUS_ID <> 21) AND" _
        & " (DENTAL_ENCOUNTER.PATIENT_UNIT_ID = " & PostMortemPatientID & ")))"
    
    'Find all virgin teeth for Antemortem patients
    sSQLAnteVT = "SELECT DENTAL_ENC_VIEW_TEETH.PATIENT_UNIT_ID From DENTAL_ENC_VIEW_TEETH, MMI_ID" _
        & " Where DENTAL_ENC_VIEW_TEETH.PATIENT_UNIT_ID = MMI_ID.UNIT_NUMBER" _
        & " AND (DENTAL_ENC_VIEW_TEETH.TOOTH_VIEW <> 0) AND" _
        & " (MMI_ID.UNIDENTIFIED_FLAG = 'A') AND" _
        & " (DENTAL_ENC_VIEW_TEETH.TOOTH_VIEW NOT IN" _
        & " (SELECT DENTAL_ENCOUNTER_TEETH.TOOTH_ID" _
        & " FROM DENTAL_ENCOUNTER_DXS_SS," _
        & " DENTAL_ENCOUNTER_DX, DENTAL_ENC_SURFACE_SET, DENTAL_ENCOUNTER_TEETH," _
        & " DENTAL_ENCOUNTER Where DENTAL_ENCOUNTER_DXS_SS.DX_CODE_ID = DENTAL_ENCOUNTER_DX.DX_CODE_ID" _
        & " AND DENTAL_ENCOUNTER_DXS_SS.SURFACE_SET_ID = DENTAL_ENC_SURFACE_SET.SURFACE_SET_ID" _
        & " AND DENTAL_ENC_SURFACE_SET.DENTAL_ENCOUNTER_TEETH_ID" _
        & " = DENTAL_ENCOUNTER_TEETH.DENTAL_ENCOUNTER_TEETH_ID" _
        & " AND DENTAL_ENCOUNTER_DX.DENTAL_ENCOUNTER_ID = DENTAL_ENCOUNTER.DENTAL_ENCOUNTER_ID" _
        & " AND DENTAL_ENCOUNTER_TEETH.DENTAL_ENCOUNTER_ID = DENTAL_ENCOUNTER.DENTAL_ENCOUNTER_ID AND" _
        & " (DENTAL_ENCOUNTER_DX.STATUS_ID <> 21) AND (DENTAL_ENCOUNTER.PATIENT_UNIT_ID = DENTAL_ENC_VIEW_TEETH.PATIENT_UNIT_ID)))"
           
    SearchForensicsChartableDxVirgin = sSQLAnteVT & " AND DENTAL_ENC_VIEW_TEETH.TOOTH_VIEW IN(" & sSQLPostVT & ")"
    
End Function




Private Function SearchForensicsImpossibleDx(PostMortemPatientID As Long) As String

'This returns a query string that contains a list of patient unit ids that
' must be removed from the overall search because the post mortem patient has a
'virgin tooth where the ante mortem tooth has a dx.

Dim sSQLPostVT As String
Dim sSQLAnteDx As String
Dim rs As Recordset
Dim sPatientIDs As String

    'Find all the Virgin teeth without Dxs for the PostMortem patient
    sSQLPostVT = "SELECT DENTAL_ENC_VIEW_TEETH.TOOTH_VIEW From DENTAL_ENC_VIEW_TEETH" _
        & " Where (PATIENT_UNIT_ID = " & PostMortemPatientID & ") And (TOOTH_VIEW <> 0)" _
        & " AND (TOOTH_VIEW NOT IN (SELECT DENTAL_ENCOUNTER_TEETH.TOOTH_ID" _
        & " FROM DENTAL_ENCOUNTER_DXS_SS, DENTAL_ENCOUNTER_DX," _
        & " DENTAL_ENC_SURFACE_SET, DENTAL_ENCOUNTER_TEETH," _
        & " DENTAL_ENCOUNTER Where DENTAL_ENCOUNTER_DXS_SS.DX_CODE_ID = DENTAL_ENCOUNTER_DX.DX_CODE_ID" _
        & " AND DENTAL_ENCOUNTER_DXS_SS.SURFACE_SET_ID = DENTAL_ENC_SURFACE_SET.SURFACE_SET_ID" _
        & " AND DENTAL_ENC_SURFACE_SET.DENTAL_ENCOUNTER_TEETH_ID" _
        & " = DENTAL_ENCOUNTER_TEETH.DENTAL_ENCOUNTER_TEETH_ID" _
        & " AND DENTAL_ENCOUNTER_DX.DENTAL_ENCOUNTER_ID = DENTAL_ENCOUNTER.DENTAL_ENCOUNTER_ID" _
        & " AND DENTAL_ENCOUNTER_TEETH.DENTAL_ENCOUNTER_ID" _
        & " = DENTAL_ENCOUNTER.DENTAL_ENCOUNTER_ID AND" _
        & " (DENTAL_ENCOUNTER_DX.STATUS_ID <> 21) AND" _
        & " (DENTAL_ENCOUNTER.PATIENT_UNIT_ID = " & PostMortemPatientID & ")))"
    
    'Find a list of patient ids where the postmortem patient has a virgin dx tooth and the
    ' ante mortem tooth has a dx
    
    'Find a list of teeth that have dxs.
    sSQLAnteDx = "SELECT DENTAL_ENCOUNTER.PATIENT_UNIT_ID" _
        & " FROM DENTAL_ENCOUNTER, DENTAL_ENCOUNTER_DX, DENTAL_ENCOUNTER_DXS_SS," _
        & " DENTAL_ENC_SURFACE_SET, DENTAL_ENCOUNTER_TEETH, MMI_ID" _
        & " Where DENTAL_ENCOUNTER.DENTAL_ENCOUNTER_ID = DENTAL_ENCOUNTER_DX.DENTAL_ENCOUNTER_ID" _
        & " AND DENTAL_ENCOUNTER_DX.DX_CODE_ID = DENTAL_ENCOUNTER_DXS_SS.DX_CODE_ID" _
        & " AND DENTAL_ENCOUNTER_DXS_SS.SURFACE_SET_ID = DENTAL_ENC_SURFACE_SET.SURFACE_SET_ID" _
        & " AND DENTAL_ENC_SURFACE_SET.DENTAL_ENCOUNTER_TEETH_ID" _
        & " = DENTAL_ENCOUNTER_TEETH.DENTAL_ENCOUNTER_TEETH_ID" _
        & " AND DENTAL_ENCOUNTER.PATIENT_UNIT_ID = MMI_ID.UNIT_NUMBER" _
        & " AND (DENTAL_ENCOUNTER_DX.STATUS_ID <> 21) AND" _
        & " (MMI_ID.UNIDENTIFIED_FLAG = 'A')"
           
    'Open the recordset and make a simpler one to speed up the search of the main query
    Set rs = goconn.ExecuteSQL(sSQLAnteDx & "AND (DENTAL_ENCOUNTER_TEETH.TOOTH_ID IN (" & sSQLPostVT & "))")
    
    If rs.EOF Then
        Exit Function
        SearchForensicsImpossibleDx = ""
    Else
        Do Until rs.EOF
            sPatientIDs = sPatientIDs & "," & rs("PATIENT_UNIT_ID")
            rs.MoveNext
        Loop
        'Remove the first "'"
        sPatientIDs = Mid(sPatientIDs, 2)
        
        SearchForensicsImpossibleDx = sPatientIDs
    End If
    
    rs.Close
    Set rs = Nothing
   
End Function




Private Function SearchForensicsImpossibleTx(PostMortemPatientID As Long) As String

'This returns a query string that contains a list of patient unit ids that
' must be removed from the overall search because the post mortem patient has a
'virgin tooth where the ante mortem tooth has a Tx.

Dim sSQLPostVT As String
Dim sSQLAnteTx As String
Dim rs As Recordset
Dim sPatientIDs As String

    'Find all the Virgin teeth without TXs for the PostMortem patient
    sSQLPostVT = "SELECT DENTAL_ENC_VIEW_TEETH.TOOTH_VIEW From DENTAL_ENC_VIEW_TEETH" _
        & " Where (PATIENT_UNIT_ID = " & PostMortemPatientID & ") And (TOOTH_VIEW <> 0)" _
        & " AND (TOOTH_VIEW NOT IN (SELECT DENTAL_ENCOUNTER_TEETH.TOOTH_ID" _
        & " FROM DENTAL_ENCOUNTER_TXS_SS, DENTAL_ENCOUNTER_TX," _
        & " DENTAL_ENC_SURFACE_SET, DENTAL_ENCOUNTER_TEETH," _
        & " DENTAL_ENCOUNTER Where DENTAL_ENCOUNTER_TXS_SS.TX_CODE_ID = DENTAL_ENCOUNTER_TX.TX_CODE_ID" _
        & " AND DENTAL_ENCOUNTER_TXS_SS.SURFACE_SET_ID = DENTAL_ENC_SURFACE_SET.SURFACE_SET_ID" _
        & " AND DENTAL_ENC_SURFACE_SET.DENTAL_ENCOUNTER_TEETH_ID" _
        & " = DENTAL_ENCOUNTER_TEETH.DENTAL_ENCOUNTER_TEETH_ID" _
        & " AND DENTAL_ENCOUNTER_TX.DENTAL_ENCOUNTER_ID = DENTAL_ENCOUNTER.DENTAL_ENCOUNTER_ID" _
        & " AND DENTAL_ENCOUNTER_TEETH.DENTAL_ENCOUNTER_ID" _
        & " = DENTAL_ENCOUNTER.DENTAL_ENCOUNTER_ID AND" _
        & " (DENTAL_ENCOUNTER_TX.STATUS_ID <> 21) AND" _
        & " (DENTAL_ENCOUNTER.PATIENT_UNIT_ID = " & PostMortemPatientID & ")))"
    
    'Find a list of patient ids where the postmortem patient has a virgin TX tooth and the
    ' ante mortem tooth has a TX
    
    'Find a list of teeth that have TXs.
    sSQLAnteTx = "SELECT DENTAL_ENCOUNTER.PATIENT_UNIT_ID" _
        & " FROM DENTAL_ENCOUNTER, DENTAL_ENCOUNTER_TX, DENTAL_ENCOUNTER_TXS_SS," _
        & " DENTAL_ENC_SURFACE_SET, DENTAL_ENCOUNTER_TEETH, MMI_ID" _
        & " Where DENTAL_ENCOUNTER.DENTAL_ENCOUNTER_ID = DENTAL_ENCOUNTER_TX.DENTAL_ENCOUNTER_ID" _
        & " AND DENTAL_ENCOUNTER_TX.TX_CODE_ID = DENTAL_ENCOUNTER_TXS_SS.TX_CODE_ID" _
        & " AND DENTAL_ENCOUNTER_TXS_SS.SURFACE_SET_ID = DENTAL_ENC_SURFACE_SET.SURFACE_SET_ID" _
        & " AND DENTAL_ENC_SURFACE_SET.DENTAL_ENCOUNTER_TEETH_ID" _
        & " = DENTAL_ENCOUNTER_TEETH.DENTAL_ENCOUNTER_TEETH_ID" _
        & " AND DENTAL_ENCOUNTER.PATIENT_UNIT_ID = MMI_ID.UNIT_NUMBER" _
        & " AND (DENTAL_ENCOUNTER_TX.STATUS_ID <> 21) AND" _
        & " (MMI_ID.UNIDENTIFIED_FLAG = 'A')"
    
    'Open the recordset and make a simpler one to speed up the search of the main query
    Set rs = goconn.ExecuteSQL(sSQLAnteTx & "AND (DENTAL_ENCOUNTER_TEETH.TOOTH_ID IN (" & sSQLPostVT & "))")
    
    If rs.EOF Then
        Exit Function
        SearchForensicsImpossibleTx = ""
    Else
        Do Until rs.EOF
            sPatientIDs = sPatientIDs & "," & rs("PATIENT_UNIT_ID")
            rs.MoveNext
        Loop
        'Remove the first "'"
        sPatientIDs = Mid(sPatientIDs, 2)
        
        SearchForensicsImpossibleTx = sPatientIDs
    End If
    
    rs.Close
    Set rs = Nothing
    
End Function








Private Function SearchForensicsNONChartableTxs(PostMortemPatientID As Long, SearchSubLoc As Boolean) As String

'This returns a query string that contains a list of patient unit ids that
' match down to the surface level a chartable treatment.

Dim rs As Recordset
Dim sSQL As String

    'Select find all the Tx info for the PostMortem patient
    sSQL = "SELECT DENTAL_ENCOUNTER_TX.TX_ADA_CODE_ID," _
        & " NOTE_TXS_LOCATION.LOCATION_ID, NOTE_TX_SUB_LOCATION.SUB_LOCATION_ID" _
        & " FROM NOTE_TXS_LOCATION, DENTAL_ENCOUNTER, DENTAL_ENCOUNTER_TX , NOTE_TX_SUB_LOCATION" _
        & " Where NOTE_TXS_LOCATION.TX_CODE_ID = DENTAL_ENCOUNTER_TX.TX_CODE_ID" _
        & " AND DENTAL_ENCOUNTER.DENTAL_ENCOUNTER_ID = DENTAL_ENCOUNTER_TX.DENTAL_ENCOUNTER_ID" _
        & " AND NOTE_TXS_LOCATION.TX_LOCATION_ID = NOTE_TX_SUB_LOCATION.TX_LOCATION_ID" _
        & " AND (DENTAL_ENCOUNTER.PATIENT_UNIT_ID = " & PostMortemPatientID & ")" _
        & " AND (DENTAL_ENCOUNTER_TX.STATUS_ID <> 104 AND DENTAL_ENCOUNTER_TX.STATUS_ID <> 100)"
    
    Set rs = goconn.OpenRecordset(sSQL)
    
    If rs.EOF Then
        SearchForensicsNONChartableTxs = ""
        Exit Function
    End If
    
    If SearchSubLoc = True Then 'Search the sub locations
        'Build the sql statment to find the related data
        sSQL = "SELECT DENTAL_ENCOUNTER.PATIENT_UNIT_ID" _
            & " FROM MMI_ID, DENTAL_ENCOUNTER, DENTAL_ENCOUNTER_TX," _
            & " NOTE_TXS_LOCATION , NOTE_TX_SUB_LOCATION" _
            & " Where MMI_ID.UNIT_NUMBER = DENTAL_ENCOUNTER.PATIENT_UNIT_ID" _
            & " AND DENTAL_ENCOUNTER.DENTAL_ENCOUNTER_ID = DENTAL_ENCOUNTER_TX.DENTAL_ENCOUNTER_ID" _
            & " AND DENTAL_ENCOUNTER_TX.TX_CODE_ID = NOTE_TXS_LOCATION.TX_CODE_ID" _
            & " AND NOTE_TXS_LOCATION.TX_LOCATION_ID = NOTE_TX_SUB_LOCATION.TX_LOCATION_ID" _
            & " AND ("
        
        'Loop through and build a sql statement to compare all known Tx data
        Do Until rs.EOF
        sSQL = sSQL & "DENTAL_ENCOUNTER_TX.TX_ADA_CODE_ID = " & rs("TX_ADA_CODE_ID") & ") AND" _
            & " (MMI_ID.UNIDENTIFIED_FLAG = 'A') AND (DENTAL_ENCOUNTER_TX.STATUS_ID <> 104 AND DENTAL_ENCOUNTER_TX.STATUS_ID <> 100) AND" _
            & " (NOTE_TXS_LOCATION.LOCATION_ID = " & rs("LOCATION_ID") & " ) AND" _
            & " (NOTE_TX_SUB_LOCATION.SUB_LOCATION_ID = " & rs("SUB_LOCATION_ID") & " ) OR "
            rs.MoveNext
        Loop
    Else 'Don't search the sub locations
        'Build the sql statment to find the related data
        sSQL = "SELECT DENTAL_ENCOUNTER.PATIENT_UNIT_ID" _
            & " FROM MMI_ID, DENTAL_ENCOUNTER, DENTAL_ENCOUNTER_TX," _
            & " NOTE_TXS_LOCATION" _
            & " Where MMI_ID.UNIT_NUMBER = DENTAL_ENCOUNTER.PATIENT_UNIT_ID" _
            & " AND DENTAL_ENCOUNTER.DENTAL_ENCOUNTER_ID = DENTAL_ENCOUNTER_TX.DENTAL_ENCOUNTER_ID" _
            & " AND DENTAL_ENCOUNTER_TX.TX_CODE_ID = NOTE_TXS_LOCATION.TX_CODE_ID" _
            & " AND ("
        
        'Loop through and build a sql statement to compare all known Tx data
        Do Until rs.EOF
        sSQL = sSQL & "DENTAL_ENCOUNTER_TX.TX_ADA_CODE_ID = " & rs("TX_ADA_CODE_ID") & ") AND" _
            & " (MMI_ID.UNIDENTIFIED_FLAG = 'A') AND (DENTAL_ENCOUNTER_TX.STATUS_ID <> 104 AND DENTAL_ENCOUNTER_TX.STATUS_ID <> 100) AND" _
            & " (NOTE_TXS_LOCATION.LOCATION_ID = " & rs("LOCATION_ID") & " ) OR "
            rs.MoveNext
        Loop
    End If
    rs.Close
    
    'Remove the last " OR "
    sSQL = Left(sSQL, Len(sSQL) - Len(" OR "))
    
    'End the sql statement
    sSQL = sSQL & ")"
    
    SearchForensicsNONChartableTxs = sSQL
    
End Function


Private Function SearchForensicsChartableDxs(PostMortemPatientID As Long, SearchSurfaceSets As Boolean) As String

'This returns a query string that contains a list of patient unit ids that
' match down to the surface level a chartable diagnosis.

Dim rs As Recordset
Dim sSQL As String

    'Select find all the Dx info for the PostMortem patient
    sSQL = "SELECT DENTAL_ENCOUNTER_TEETH.TOOTH_NAME," _
        & " DENTAL_ENC_SURFACE_SET.SURFACESET," _
        & " DENTAL_ENCOUNTER_DX.DX_ICD_CODE_ID" _
        & " FROM DENTAL_ENCOUNTER_DX, DENTAL_ENCOUNTER," _
        & " DENTAL_ENCOUNTER_DXS_SS, DENTAL_ENC_SURFACE_SET, DENTAL_ENCOUNTER_TEETH" _
        & " Where DENTAL_ENCOUNTER_DX.DENTAL_ENCOUNTER_ID = DENTAL_ENCOUNTER.DENTAL_ENCOUNTER_ID" _
        & " AND DENTAL_ENCOUNTER_DX.DX_CODE_ID = DENTAL_ENCOUNTER_DXS_SS.DX_CODE_ID" _
        & " AND DENTAL_ENCOUNTER_DXS_SS.SURFACE_SET_ID = DENTAL_ENC_SURFACE_SET.SURFACE_SET_ID" _
        & " AND DENTAL_ENC_SURFACE_SET.DENTAL_ENCOUNTER_TEETH_ID" _
        & " = DENTAL_ENCOUNTER_TEETH.DENTAL_ENCOUNTER_TEETH_ID" _
        & " AND (DENTAL_ENCOUNTER.PATIENT_UNIT_ID = " & PostMortemPatientID & ") AND" _
        & " (DENTAL_ENCOUNTER_DX.STATUS_ID <> 21)"

    Set rs = goconn.OpenRecordset(sSQL)
    
    If rs.EOF Then
        SearchForensicsChartableDxs = ""
        Exit Function
    End If
    
    'Begin the sql statement
    sSQL = "SELECT DENTAL_ENCOUNTER.PATIENT_UNIT_ID" _
        & " FROM DENTAL_ENCOUNTER_DX, DENTAL_ENCOUNTER," _
        & " DENTAL_ENCOUNTER_DXS_SS, DENTAL_ENC_SURFACE_SET, DENTAL_ENCOUNTER_TEETH, MMI_ID" _
        & " Where DENTAL_ENCOUNTER_DX.DENTAL_ENCOUNTER_ID = DENTAL_ENCOUNTER.DENTAL_ENCOUNTER_ID" _
        & " AND DENTAL_ENCOUNTER_DX.DX_CODE_ID = DENTAL_ENCOUNTER_DXS_SS.DX_CODE_ID" _
        & " AND DENTAL_ENCOUNTER_DXS_SS.SURFACE_SET_ID = DENTAL_ENC_SURFACE_SET.SURFACE_SET_ID" _
        & " AND DENTAL_ENC_SURFACE_SET.DENTAL_ENCOUNTER_TEETH_ID" _
        & " = DENTAL_ENCOUNTER_TEETH.DENTAL_ENCOUNTER_TEETH_ID" _
        & " AND DENTAL_ENCOUNTER.PATIENT_UNIT_ID = MMI_ID.UNIT_NUMBER" _
        & " AND ("
    
    'Loop through and build a sql statement to compare all known Tx data
    Do Until rs.EOF
        If SearchSurfaceSets = True Then 'Search the surface sets
            sSQL = sSQL & "(DENTAL_ENCOUNTER_TEETH.TOOTH_NAME = '" & rs("TOOTH_NAME") & "') AND" _
                & " (DENTAL_ENCOUNTER_DX.DX_ICD_CODE_ID = " & rs("DX_ICD_CODE_ID") & ") AND" _
                & " (DENTAL_ENC_SURFACE_SET.SURFACESET = '" & rs("SURFACESET") & "') AND" _
                & " (DENTAL_ENCOUNTER_DX.STATUS_ID <> 21) AND (MMI_ID.UNIDENTIFIED_FLAG = 'A') OR "
        Else 'Don't search the surface sets
            sSQL = sSQL & "(DENTAL_ENCOUNTER_TEETH.TOOTH_NAME = '" & rs("TOOTH_NAME") & "') AND" _
                & " (DENTAL_ENCOUNTER_DX.DX_ICD_CODE_ID = " & rs("DX_ICD_CODE_ID") & ") AND" _
                & " (DENTAL_ENCOUNTER_DX.STATUS_ID <> 21) AND (MMI_ID.UNIDENTIFIED_FLAG = 'A') OR "
        End If
        rs.MoveNext
    Loop
    
    rs.Close

    'Remove the last " OR "
    sSQL = Left(sSQL, Len(sSQL) - 4)
    
    'End the sql statement
    sSQL = sSQL & ")"
    
    SearchForensicsChartableDxs = sSQL
    
End Function

Private Function SearchForensicsNONChartableDxs(PostMortemPatientID As Long, SearchSubLoc As Boolean) As String

'This returns a query string that contains a list of patient unit ids that
' match down to the surface level a chartable diagnosis.

Dim rs As Recordset
Dim sSQL As String

    'Select find all the Dx info for the PostMortem patient
    sSQL = "SELECT DENTAL_ENCOUNTER_DX.DX_ICD_CODE_ID," _
        & " NOTE_DX_SUB_LOCATION.SUB_LOCATION_ID," _
        & " NOTE_DXS_LOCATION.LOCATION_ID" _
        & " FROM DENTAL_ENCOUNTER_DX, DENTAL_ENCOUNTER," _
        & " NOTE_DXS_LOCATION , NOTE_DX_SUB_LOCATION" _
        & " Where DENTAL_ENCOUNTER_DX.DENTAL_ENCOUNTER_ID = DENTAL_ENCOUNTER.DENTAL_ENCOUNTER_ID" _
        & " AND DENTAL_ENCOUNTER_DX.DX_CODE_ID = NOTE_DXS_LOCATION.DX_CODE_ID" _
        & " AND NOTE_DXS_LOCATION.DX_LOCATION_ID = NOTE_DX_SUB_LOCATION.DX_LOCATION_ID" _
        & " And (DENTAL_ENCOUNTER.PATIENT_UNIT_ID = " & PostMortemPatientID & ")" _
        & " AND (DENTAL_ENCOUNTER_DX.STATUS_ID <> 21)"
    
    Set rs = goconn.OpenRecordset(sSQL)
    
    If rs.EOF Then
        SearchForensicsNONChartableDxs = ""
        Exit Function
    End If

    If SearchSubLoc = True Then 'Search the sublocations
        'Begin the sql statement
        sSQL = "SELECT DENTAL_ENCOUNTER.PATIENT_UNIT_ID" _
            & " FROM MMI_ID, DENTAL_ENCOUNTER, DENTAL_ENCOUNTER_DX," _
            & " NOTE_DXS_LOCATION , NOTE_DX_SUB_LOCATION" _
            & " Where MMI_ID.UNIT_NUMBER = DENTAL_ENCOUNTER.PATIENT_UNIT_ID" _
            & " AND DENTAL_ENCOUNTER.DENTAL_ENCOUNTER_ID = DENTAL_ENCOUNTER_DX.DENTAL_ENCOUNTER_ID" _
            & " AND DENTAL_ENCOUNTER_DX.DX_CODE_ID = NOTE_DXS_LOCATION.DX_CODE_ID" _
            & " AND NOTE_DXS_LOCATION.DX_LOCATION_ID = NOTE_DX_SUB_LOCATION.DX_LOCATION_ID" _
            & " AND ("
        
        'Loop through and build a sql statement to compare all known Tx data
        Do Until rs.EOF
            sSQL = sSQL & "MMI_ID.UNIDENTIFIED_FLAG = 'A') AND" _
                & " (DENTAL_ENCOUNTER_DX.STATUS_ID <> 21) AND" _
                & " (NOTE_DXS_LOCATION.LOCATION_ID = " & rs("LOCATION_ID") & ") AND" _
                & " (NOTE_DX_SUB_LOCATION.SUB_LOCATION_ID = " & rs("SUB_LOCATION_ID") & ") AND" _
                & " (DENTAL_ENCOUNTER_DX.DX_ICD_CODE_ID = " & rs("DX_ICD_CODE_ID") & ") OR "
    
            rs.MoveNext
        Loop
    Else 'Don't search the sub locations
        'Begin the sql statement
        sSQL = "SELECT DENTAL_ENCOUNTER.PATIENT_UNIT_ID" _
            & " FROM MMI_ID, DENTAL_ENCOUNTER, DENTAL_ENCOUNTER_DX," _
            & " NOTE_DXS_LOCATION Where MMI_ID.UNIT_NUMBER = DENTAL_ENCOUNTER.PATIENT_UNIT_ID" _
            & " AND DENTAL_ENCOUNTER.DENTAL_ENCOUNTER_ID = DENTAL_ENCOUNTER_DX.DENTAL_ENCOUNTER_ID" _
            & " AND DENTAL_ENCOUNTER_DX.DX_CODE_ID = NOTE_DXS_LOCATION.DX_CODE_ID" _
            & " AND ("
            
        'Loop through and build a sql statement to compare all known Tx data
        Do Until rs.EOF
            sSQL = sSQL & "MMI_ID.UNIDENTIFIED_FLAG = 'A') AND" _
                & " (DENTAL_ENCOUNTER_DX.STATUS_ID <> 21) AND" _
                & " (NOTE_DXS_LOCATION.LOCATION_ID = " & rs("LOCATION_ID") & ") AND" _
                & " (DENTAL_ENCOUNTER_DX.DX_ICD_CODE_ID = " & rs("DX_ICD_CODE_ID") & ") OR "
    
            rs.MoveNext
        Loop
    End If
    
    rs.Close

    'Remove the last " OR "
    sSQL = Left(sSQL, Len(sSQL) - 4)
    
    'End the sql statement
    sSQL = sSQL & ")"
    
    SearchForensicsNONChartableDxs = sSQL
    
End Function


Private Function SignEncounter() As Boolean

'******* NOT USED ANYMORE ********

'This function scans all the Dxs and Txs and sets
' the object.signed = true that have this dental encounter id

'This only works for LoadDxs now
    
'Dim oDx As clsDx, oTx As clsTx
'
'    'The Dxs have not been loaded
'    If mvarDxs Is Nothing Then
'        SignEncounter = False
'        Exit Function
'    End If
'
'    For Each oDx In mvarDxs
'        'Sign all the Dxs for this encounter
'        If oDx.DentalEncounterID = mvarCurrentDentalEncounterID Then
'            oDx.Signed = True
'        End If
'
'        For Each oTx In oDx.Txs
'            'Sign all the Txs for this Dx and are for this encounter
'            If oTx.DentalEncounterID = mvarCurrentDentalEncounterID Then
'                oTx.Signed = True
'            End If
'        Next
'    Next
'
'    SignEncounter = True

End Function

Public Function UnlockForensicsPatientID(PatientID As Long)

Dim sSQL As String

'Unlock the record based on the UnlockPatientID
        sSQL = "UPDATE HIT_QUEUE SET RECORD_LOCK = NULL" _
            & " Where (POSTMORTEMID = " & PatientID & ")"
        goconn.ExecuteSQL sSQL

End Function


Private Sub Class_Terminate()

    Set mvarDxs = Nothing
    Set mvarTxs = Nothing
    Set mvarGroups = Nothing
    Set mvarTxsRef = Nothing
    Set mvarDxsRef = Nothing
    Set moRads = Nothing
    Set moRadOrders = Nothing
    
End Sub


