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
Attribute VB_Name = "colRadOrders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsRadOrder"
Attribute VB_Ext_KEY = "Member0" ,"clsRadOrder"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable to hold collection
Private mCol As Collection

'Used when creating RadOrders
Private mlDentalEncounterID As Long
Private mlProviderID As Long
Private msProviderName As String

Private morsLocations As Recordset
Private morsTxADACodes As Recordset

Public Function AllLocations() As Recordset
    'Returns a recordset containing all the locations from the Note_Location table
    
    Set AllLocations = morsLocations
    
End Function

Public Property Get DentalEncounterID() As Long
    
    DentalEncounterID = mlDentalEncounterID
    
End Property

Friend Property Let DentalEncounterID(vData As Long)
    
    mlDentalEncounterID = vData
    
End Property

Public Sub ExternalConnect(DB As String, UserID As String, Password As String)

    'This is called from the Worklist SCP to connect to the database from outside of chcsii

Dim oICD As ICD_DAS
Dim sConnectionString As String

On Error GoTo ErrHndDAS

    Set goChcsiiConn = New CHCSII_CONN.Conn
    
    Set goconn = goChcsiiConn.CHCSII_DAS(ICD, NoTimeZoneConversion, Default)
        
    Set oICD = goconn
    
    sConnectionString = "Provider=MSDataShape.1;" & _
                    "Data Provider=MSDAORA.1;" & _
                    "Data Source=" & DB & ";" & _
                    "User ID=" & UserID & ";" & _
                    "Password=" & Password & ";" & _
                    "Persist Security Info=True"
    
On Error GoTo ErrHnd

    oICD.OpenDatabase sConnectionString, ""
    
    Set oICD = Nothing
    
Exit Sub

ErrHndDAS:
    Resume Next
    
Exit Sub

ErrHnd:
    MsgBox Err.Description & ": " & Err.Source
    
End Sub

Public Function Init(DentalEncounterID As Long, ProviderID As Long, ProviderName As String) As Boolean
    
    mlDentalEncounterID = DentalEncounterID
    mlProviderID = ProviderID
    msProviderName = ProviderName
    
    Init = True
    
End Function

Public Function LookupOrderPriorityName(OrderPriority As eRadOrderPriority) As String
    
    LookupOrderPriorityName = LoadResString(OrderPriority)
    
End Function


Public Function LookupRadOrderStatusName(Status As eRadOrderStatus) As String
    
    LookupRadOrderStatusName = LoadResString(Status)
    
End Function

Public Function OrderPriorityValues() As String
    
    'Returns the possible enum values for the order priority
    OrderPriorityValues = cRAD_ORDER_PRIORITY_VALUES
    
End Function

Public Function OrderStatusValues() As String
    
    OrderStatusValues = cRAD_ORDER_STATUS_VALUES
    
End Function



Public Property Get ProviderID() As Long
    
    ProviderID = mlProviderID
    
End Property

Friend Property Let ProviderID(vData As Long)
    
    mlProviderID = vData
    
End Property
Public Property Get ProviderName() As String
    
    ProviderName = msProviderName
    
End Property

Friend Property Let ProviderName(vData As String)
    
    msProviderName = vData
    
End Property

Private Property Get RadOrderTypeValues() As Variant
    
    'Not used yet
    
    'Returns an array of possible rad type enums

    RadOrderTypeValues = Split(cRAD_ORDER_TYPE_VALUES, ",")
    
End Property

Public Function LookupRadOrderTypeName(RadType As eRadType) As String
        
        LookupRadOrderTypeName = LoadResString(RadType)

End Function

Public Function LoadRadOrders(PatientID As Long, ProviderID As Long, QueueID As Long, Status As eRadOrderStatus, DateTime As Date, DICOMAccessionNumber As String) As Boolean

'This loads the rads according to the parameter values passed in
    
Dim oRs As Recordset
Dim sWhere As String
Dim oRadOrder As clsRadOrder

On Error GoTo ErrHnd
    
'    If mlDentalEncounterID = 0 Then
'        MsgBox "You must call Init before calling LoadRadOrders"
'        Exit Function
'    End If
    
    'Build the where clause based on the parameters
    If PatientID > 0 Then
        sWhere = " Patient_Unit_Number = " & PatientID & " and "
    End If
    
    If ProviderID > 0 Then
        sWhere = sWhere & " Provider_ID = " & ProviderID & " and "
    End If
    
    If QueueID > 0 Then
        sWhere = sWhere & " Queue_ID = " & QueueID & " and "
    End If
    
    If Status <> 0 Then
        sWhere = sWhere & " Status_ID = " & Status & " and "
    End If
    
    If DateTime <> Empty Then
        Debug.Assert False 'Need to test this with all days
        sWhere = sWhere & " REQUESTED_BY_DATE_TIME = " & DateTime & " and "
    End If
    
    If DICOMAccessionNumber <> "" Then
        sWhere = sWhere & " DICOM_Accession_Number = '" & DICOMAccessionNumber & "' and "
    End If
    
    If sWhere <> "" Then
        sWhere = " where " & Left(sWhere, Len(sWhere) - Len(" and "))
    End If
    
    Set oRs = goconn.OpenRecordset("Select * from V_DENT_LOAD_RAD_ORDERS " & sWhere)
    
    Set mCol = Nothing
    Set mCol = New Collection
    
    Do Until oRs.EOF
        Set oRadOrder = New clsRadOrder
        With oRadOrder
            .InternalAdd = True
            .RadOrderID = oRs("Rad_Order_ID")
            .DICOMAccessionNumber = oRs("DICOM_ACCESSION_NUMBER") & ""
            
            'This is used when the image store scp creates a new rad.
            .FacilityNCID = Val(oRs("Facility_NCID") & "")
            
            .DentalEncounterID = Val(oRs("Dental_Encounter_ID") & "")
            .RadTypeID = oRs("Rad_Type_ID")
            .ProviderID = oRs("Provider_ID")
            .ProviderName = oRs("Provider_Name")
            .StatusID = oRs("Status_ID")
            .TxADACodeID = Val(oRs("Tx_ADA_Code_ID") & "")
            .TxADACode = oRs("ADA_Code") & ""
            .TxADACodeDefinition = oRs("ADA_Definition") & ""
            .CreatedDateTime = oRs("Create_Date_Time") & ""
            .RequestedByDate = oRs("REQUESTED_BY_DATE_TIME") & ""
            .LocationName = oRs("Location_Name")
            .Quantity = oRs("Quantity") & ""
            .Instructions = oRs("Instructions") & ""
            .PriorityID = Val(oRs("Priority_ID") & "")
            .PatientUnitNumber = oRs("Patient_Unit_Number")
            .PatientName = oRs("Patient_Name")
            .QueueID = oRs("Queue_ID")
            .QueueName = oRs("Queue_Name") & ""
            .InternalAdd = False
            .DirtyFlag = dfNoAction
        End With
        
        mCol.Add oRadOrder, CStr(oRadOrder.RadOrderID)
        oRs.MoveNext
    Loop
    
    Set oRadOrder = Nothing
    
Exit Function

ErrHnd:
    MsgBox Err.Description & ": " & Err.Number & ": " & Err.Source
    
End Function

Public Function Save() As Boolean

Dim oRadOrder As clsRadOrder

    For Each oRadOrder In mCol
        oRadOrder.Save
        
        'If the rad order was deleted, then remove it from the collection
        If oRadOrder.StatusID = eRadOrderStatus.Deleted Then
            mCol.Remove CStr(oRadOrder.RadOrderID)
            Set oRadOrder = Nothing
        End If
    Next
    
End Function


Public Function Add(TypeID As eRadType, LocationName As String, Quantity As Long, PatientUnitNumber As Long, PatientName As String) As clsRadOrder
    
Dim objNewMember As clsRadOrder
Dim oRs As Recordset
Dim sSQL As String
Dim lRadOrderID As Long
Dim sDicomAssessionNumber As Long

On Error GoTo ErrHnd

'    sSQL = "Select Dental_Imaging_Rad_Orders_SEQ.NextVal as RadOrderID from Dual"
'    Set oRs = goconn.OpenRecordset(sSQL)
'    If oRs.EOF Then
'        MsgBox "Cannot " & sSQL
'        Exit Function
'    Else
'        lRadOrderID = oRs("RadOrderID")
'    End If
'    oRs.Close
    lRadOrderID = goconn.GetNextID(esDENTAL_IMAGING_RAD_ORDERS_SEQ)
    
    
'    'Get the DICOM Assession number - this sequence cycles at 99999
'    sSQL = "Select DENTAL_IMG_RAD_Order_ACC_SEQ.NextVal as DicomAssID from Dual"
'    Set oRs = goconn.OpenRecordset(sSQL)
'    If oRs.EOF Then
'        MsgBox "Cannot " & sSQL
'        Exit Function
'    Else
'        sDicomAssessionNumber = Format(cwobjCWComm.GetDateTime, "MMDDYYYY") & "-" & Format(oRs("DicomAssID"), 0)
'    End If
'    oRs.Close
    sDicomAssessionNumber = Format(cwobjCWComm.GetDateTime, "MMDDYYYY") & "-" & Format(goconn.GetNextID(esDENTAL_IMG_RAD_ORDER_ACC_SEQ), 0)
    
    Set objNewMember = New clsRadOrder

    With objNewMember
    
        'Get the RadOrderID from the database
'        Debug.Assert False
        '.RadOrderID = RadOrderID
        
        'Module level vars
        .DentalEncounterID = mlDentalEncounterID
        .ProviderID = mlProviderID
        .ProviderName = msProviderName
        
        Set .TxADACodes = morsTxADACodes
        
        'Required vars
        .RadOrderID = lRadOrderID
        .DICOMAccessionNumber = sDicomAssessionNumber
        .RadTypeID = TypeID
        .Quantity = Quantity
        .LocationName = LocationName
        .PatientUnitNumber = PatientUnitNumber
        .PatientName = PatientName
        .FacilityNCID = cwobjLogonEnviron.StoringFacilityNCID
        
        'Defaults
        .StatusID = Planned
        .RequestedByDate = cwobjCWComm.GetDateTime
        .CreatedDateTime = cwobjCWComm.GetDateTime
        .PriorityID = Routine
        
        '.TxADACodeID = TxADACodeID
        '.LocationID = LocationID
        '.Pregnant = Pregnant
        
        .DirtyFlag = dfAdd
        
        mCol.Add objNewMember, CStr(.RadOrderID)
'        mCol.Add objNewMember

    End With
    
    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing

Exit Function

ErrHnd:
    MsgBox Err.Description & ": " & Err.Number
    'Resume
End Function

Public Property Get Item(vntIndexKey As Variant) As clsRadOrder
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


    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    
On Error GoTo ErrHnd

    If Not goconn Is Nothing Then
        'creates the collection when this class is created
        Set mCol = New Collection
        
        'Init the morsLocations recordset
        Set morsLocations = goconn.OpenRecordset("SELECT LOCATION_ID, LOCATION_NAME, IS_TOOTH_LOC, SORT_ORDER From V_NOTE_LOCATION_RADS order by Sort_Order", adOpenStatic)
        morsLocations.ActiveConnection = Nothing
        
        'Init the morsTxADACodes recordset
        Set morsTxADACodes = goconn.OpenRecordset("SELECT TX_ADA_CODE_ID, ADA_CODE, ADA_DEFINITION, ADA_SHORT From TX_ADA_CODE_DEFINITION")
        morsTxADACodes.ActiveConnection = Nothing
    End If
    
Exit Sub

ErrHnd:
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub

