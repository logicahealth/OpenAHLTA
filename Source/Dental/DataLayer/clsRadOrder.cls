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
Attribute VB_Name = "clsRadOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarRadOrderID As Long
Private mvarDICOMAccessionNumber As String
Private mvarDentalEncounterID As Long
Private mvarRadTypeID As eRadType
Private mvarRadTypeName As String
Private mvarProviderID As Long
Private mvarProviderName As String
Private mvarStatusID As eRadOrderStatus
Private mvarStatusName As String
Private mvarPatientUnitNumber As Long
Private mvarPatientName As String
Private mvarQueueID As Long
Private mvarQueueName As String
Private mlFacilityNCID As Long

Private moRads As colRads

'Can be ADA or CPT codes both in same table
Private mvarTxADACodeID As Long
Private mvarTxADACodeDefinition As String
Private mvarTxADACode As String

Private mvarCreatedDateTime As Date
'Private mvarLocationID As Long
Private mvarLocationName As String
Private mvarQuantity As Long
Private mvarInstructions As String
Private mvarPriorityID As eRadOrderPriority
Private mvarPriorityName As String
Private mvarRequestedByDate As Date
Private mvarPregnant As Boolean

Private morsTxADACodes As Recordset
Private meDirtyFlag As DirtyFlagEnum
Private mbInternalAdd As Boolean

Public Property Let DICOMAccessionNumber(vData As String)
    
    mvarDICOMAccessionNumber = vData
    
End Property

Public Property Let FacilityNCID(vData As Long)
    
    mlFacilityNCID = vData
    
End Property

Public Property Get FacilityNCID() As Long

    FacilityNCID = mlFacilityNCID
    
End Property

Public Property Get DICOMAccessionNumber() As String
    
    'Format the accession number based on the RadOrderID and the CreatedDateTime
    'MMDDYYYY-NNNNN
    'TODO: This is a problem when mvarRadOrderID becomes more than 5 digits
    
    DICOMAccessionNumber = mvarDICOMAccessionNumber
    
End Property


Public Property Get DICOMCreatedDateTime() As String
    
    'This formats the CreatedDateTime to the DICOM format, YYYYMMDD
    
    DICOMCreatedDateTime = Format(mvarCreatedDateTime, "YYYYMMDD")
    
End Property

Public Property Get DICOMPatientName() As String
    
    'Returns the properly DICOM formatted patient name, cannot be larger than 64 chars
    DICOMPatientName = Left(Replace(Replace(mvarPatientName, ",", ""), " ", "^"), 64)
    
End Property


Public Property Get DICOMProviderame() As String
    
    'Returns the properly DICOM formatted patient name, cannot be larger than 64 chars
    DICOMProviderame = Left(Replace(Replace(mvarProviderName, ",", ""), " ", "^"), 64)
    
End Property



Friend Property Get InternalAdd() As Boolean
    
    InternalAdd = mbInternalAdd
    
End Property

Friend Property Let InternalAdd(vData As Boolean)
    
    mbInternalAdd = vData
    
End Property


Public Property Get PatientName() As String
    
    PatientName = mvarPatientName
    
End Property

Friend Property Let PatientName(vData As String)
    
    mvarPatientName = vData
    
End Property

Public Property Get PatientUnitNumber() As Long
    
    PatientUnitNumber = mvarPatientUnitNumber
    
End Property

Friend Property Let PatientUnitNumber(vData As Long)
        
    'Debug.Assert False
    'This should not be changed after it is added.
    
    mvarPatientUnitNumber = vData
    
End Property

Public Property Get QueueID() As Long
    
    QueueID = mvarQueueID
    
End Property

Public Property Let QueueID(vData As Long)
        
    'Get the queuename from the database
    
Dim sSQL As String
Dim oRs As Recordset

    If mbInternalAdd = False Then
        sSQL = "Select Queue_Name from Dental_Imaging_Queues where Queue_ID = " & vData
        Set oRs = goconn.OpenRecordset(sSQL)
        If Not oRs.EOF Then
            mvarQueueName = oRs("Queue_Name")
        Else
            MsgBox "Did not get any results from: " & sSQL
        End If
        oRs.Close
    End If
        
    mvarQueueID = vData
    
End Property
Public Property Get QueueName() As String
    
    QueueName = mvarQueueName
    
End Property

Friend Property Let QueueName(vData As String)
    
    mvarQueueName = vData
    
End Property

Public Property Get Rads() As colRads
    
    'Get a new rads if one does not exist
    '********** NOTE: The rads object is not loaded from the LoadRadOrders method in the colRadOrders class.
    If moRads Is Nothing Then
        Set moRads = New colRads
        moRads.ExternalInit mvarRadOrderID, mvarLocationName, mvarRadTypeID, mvarPatientUnitNumber, mvarProviderID, mlFacilityNCID
    End If
    
    Set Rads = moRads
    
End Property

Public Function Save() As Boolean
    
    'Save the RadOrder to the Dental_Imaging_Rad_Orders table

Dim sSQL As String
Dim oRad As clsRad

On Error GoTo ErrHnd

    Select Case meDirtyFlag
    
        Case DirtyFlagEnum.dfNoAction
        
        Case DirtyFlagEnum.dfAdd
            sSQL = "INSERT INTO DENTAL_IMAGING_RAD_ORDERS" _
                & " (RAD_ORDER_ID, DENTAL_ENCOUNTER_ID, RAD_TYPE_ID," _
                & " PROVIDER_ID, STATUS_ID, TX_ADA_CODE_ID," _
                & " CREATE_DATE_TIME, REQUESTED_BY_DATE_TIME," _
                & " LOCATION_NAME, QUANTITY, INSTRUCTIONS, PRIORITY_ID," _
                & " PATIENT_UNIT_NUMBER, Queue_ID, DICOM_ACCESSION_NUMBER, Facility_NCID)" _
                & " VALUES (" _
                & " " & mvarRadOrderID & ", " & mvarDentalEncounterID & ", " & mvarRadTypeID & ", " _
                & " " & mvarProviderID & ", " & mvarStatusID & ", " & mvarTxADACodeID & ", " _
                & " " & goChcsiiConn.FormatDateForDBInsert(mvarCreatedDateTime) & ", " & goChcsiiConn.FormatDateForDBInsert(mvarRequestedByDate) & ", " _
                & " '" & mvarLocationName & "', " & mvarQuantity & ", '" & RQ(mvarInstructions) & "', " & mvarPriorityID & ", " _
                & " " & mvarPatientUnitNumber & ", " & mvarQueueID & ", '" & mvarDICOMAccessionNumber & "', " & mlFacilityNCID & ")"
                        
                'TO_DATE(7, '" & mvarCreateDateTime & "')
                'TO_DATE(8, '" & mvarRequestedByDate & "')
        Case DirtyFlagEnum.dfUpdate
            sSQL = "UPDATE DENTAL_IMAGING_RAD_ORDERS" _
                    & " SET DENTAL_ENCOUNTER_ID = " & mvarDentalEncounterID & ", RAD_TYPE_ID = " & mvarRadTypeID & "," _
                    & " PROVIDER_ID = " & mvarProviderID & ", STATUS_ID = " & mvarStatusID & ", TX_ADA_CODE_ID = " & mvarTxADACodeID & "," _
                    & " CREATE_DATE_TIME = " & goChcsiiConn.FormatDateForDBInsert(mvarCreatedDateTime) & "," _
                    & " REQUESTED_BY_DATE_TIME = " & goChcsiiConn.FormatDateForDBInsert(mvarRequestedByDate) & "," _
                    & " LOCATION_NAME = '" & mvarLocationName & "', QUANTITY = " & mvarQuantity & ", INSTRUCTIONS = '" & RQ(Replace(mvarInstructions, vbCrLf, "")) & "'," _
                    & " PRIORITY_ID = " & mvarPriorityID & ", PATIENT_UNIT_NUMBER = " & mvarPatientUnitNumber & ", " _
                    & " Queue_ID = " & mvarQueueID & "," _
                    & " DICOM_ACCESSION_NUMBER = '" & mvarDICOMAccessionNumber & "'," _
                    & " Facility_NCID = '" & mlFacilityNCID & "'" _
                    & " Where (RAD_ORDER_ID = " & mvarRadOrderID & ")"

        Case DirtyFlagEnum.dfdelete
            sSQL = "Update DENTAL_IMAGING_RAD_ORDERS SET Status_ID = " & eRadOrderStatus.Deleted & " Where (RAD_ORDER_ID = " & mvarRadOrderID
        
    End Select
    
    If sSQL <> "" Then
        goconn.ExecuteSQL sSQL
    End If
    
    'Save the Rads if any
    For Each oRad In moRads
        oRad.Save
        
        'There is no current logic to remove the rads from the collection
        If oRad.StatusID = eRadStatus.Deleted Then
            'Add logic to remove the rads from the collection if needed
        End If
    Next
    
    meDirtyFlag = DirtyFlagEnum.dfNoAction
    
Exit Function

ErrHnd:

    MsgBox "Could not save the rad order: " & Err.Description & ": " & Err.Source
    
End Function

Private Sub SetDirtyFlag()
    
    'This is called when a user sets a property
    'If the current flag is dfAdd, then don't change it until Save is called
    'otherwise, set it to sfUpdate
    
    If meDirtyFlag = dfNoAction Then
        meDirtyFlag = dfUpdate
    End If
    
End Sub
Friend Property Let DirtyFlag(ByVal vData As DirtyFlagEnum)
    meDirtyFlag = vData
End Property

Friend Property Get DirtyFlag() As DirtyFlagEnum
    DirtyFlag = meDirtyFlag
End Property

Public Property Let Pregnant(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Pregnant = 5
    mvarPregnant = vData
    
    SetDirtyFlag
End Property


Public Property Get Pregnant() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Pregnant
    Pregnant = mvarPregnant
End Property



Public Property Let RequestedByDate(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RequestedByDate = 5
    mvarRequestedByDate = vData
    
    SetDirtyFlag
End Property


Public Property Get RequestedByDate() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RequestedByDate
    RequestedByDate = mvarRequestedByDate
End Property



Private Property Let PriorityName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PriorityName = 5
    mvarPriorityName = vData
End Property


Public Property Get PriorityName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PriorityName
    PriorityName = mvarPriorityName
End Property



Public Property Let PriorityID(ByVal vData As eRadOrderPriority)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PriorityID = 5
    mvarPriorityID = vData
    
    mvarPriorityName = LoadResString(vData)
    
    SetDirtyFlag
End Property


Public Property Get PriorityID() As eRadOrderPriority
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PriorityID
    PriorityID = mvarPriorityID
End Property



Public Property Let Instructions(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Instructions = 5
    mvarInstructions = vData
    
    SetDirtyFlag
End Property


Public Property Get Instructions() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Instructions
    Instructions = mvarInstructions
End Property



Public Property Let Quantity(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Quantity = 5
    mvarQuantity = vData
    
    SetDirtyFlag
End Property


Public Property Get Quantity() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Quantity
    Quantity = mvarQuantity
End Property



Public Property Let LocationName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LocationName = 5
    mvarLocationName = vData
End Property


Public Property Get LocationName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LocationName
    LocationName = mvarLocationName
End Property



'Public Property Let LocationID(ByVal vData As Long)
''used when assigning a value to the property, on the left side of an assignment.
''Syntax: X.LocationID = 5
'    mvarLocationID = vData
'
'    SetDirtyFlag
'End Property


'Public Property Get LocationID() As Long
''used when retrieving value of a property, on the right side of an assignment.
''Syntax: Debug.Print X.LocationID
'    LocationID = mvarLocationID
'End Property



Public Property Let CreatedDateTime(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CreatedDateTime = 5
    mvarCreatedDateTime = vData
    
    SetDirtyFlag
End Property


Public Property Get CreatedDateTime() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CreatedDateTime
    CreatedDateTime = mvarCreatedDateTime
End Property



Friend Property Let TxADACodeDefinition(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TxADACodeDefinition = 5
    mvarTxADACodeDefinition = vData
    
    SetDirtyFlag
End Property

Public Property Get TxADACodeDefinition() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TxADACodeDefinition
    TxADACodeDefinition = mvarTxADACodeDefinition
End Property

Friend Property Let TxADACode(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TxADACodeDefinition = 5
    mvarTxADACode = vData
    
    SetDirtyFlag
End Property


Public Property Get TxADACode() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TxADACodeDefinition
    TxADACode = mvarTxADACode
End Property


Public Property Let TxADACodeID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TxADACodeID = 5
    
    If mbInternalAdd = False Then
        'Get the code
        morsTxADACodes.Filter = "TX_ADA_CODE_ID = " & vData
        
        If Not morsTxADACodes.EOF Then
            mvarTxADACode = morsTxADACodes("ADA_Code")
            mvarTxADACodeDefinition = morsTxADACodes("ADA_DEFINITION")
        Else
            MsgBox "Could not find any records in the TX_ADA_CODE_DEFINITION table with TX_ADA_CODE_ID = " & vData
        End If
    End If
    
    mvarTxADACodeID = vData
    
    SetDirtyFlag
    
End Property


Public Property Get TxADACodeID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TxADACodeID
    TxADACodeID = mvarTxADACodeID
End Property



Private Property Let StatusName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.StatusName = 5
    mvarStatusName = vData
End Property


Public Property Get StatusName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.StatusName
    StatusName = mvarStatusName
End Property



Public Property Let StatusID(ByVal vData As eRadOrderStatus)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.StatusID = 5
    mvarStatusID = vData
    
    mvarStatusName = LoadResString(vData)
    
    SetDirtyFlag
    
End Property


Public Property Get StatusID() As eRadOrderStatus
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.StatusID
    StatusID = mvarStatusID
End Property



Public Property Let ProviderName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ProviderName = 5
    mvarProviderName = vData
End Property


Public Property Get ProviderName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ProviderName
    ProviderName = mvarProviderName
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



Private Property Let RadTypeName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TypeName = 5
    mvarRadTypeName = vData
End Property


Public Property Get RadTypeName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TypeName
    RadTypeName = mvarRadTypeName
End Property



Public Property Let RadTypeID(ByVal vData As eRadType)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TypeID = 5
    mvarRadTypeID = vData
    
    mvarRadTypeName = LoadResString(vData)
    
    SetDirtyFlag
End Property


Public Property Get RadTypeID() As eRadType
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TypeID
    RadTypeID = mvarRadTypeID
End Property



Public Property Let DentalEncounterID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DentalEncounterID = 5
    mvarDentalEncounterID = vData
    
    SetDirtyFlag
End Property


Public Property Get DentalEncounterID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DentalEncounterID
    DentalEncounterID = mvarDentalEncounterID
End Property



Public Property Let RadOrderID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RadOrderID = 5
    mvarRadOrderID = vData
    
    SetDirtyFlag
End Property


Public Property Get RadOrderID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RadOrderID
    RadOrderID = mvarRadOrderID
End Property



Friend Property Set TxADACodes(vData As Recordset)
    
    Set morsTxADACodes = vData
    
End Property

Private Sub Class_Initialize()

    'Default the priority to Routine
    Me.PriorityID = eRadOrderPriority.Routine
    
    'Default the status to Planned
    Me.StatusID = eRadOrderStatus.Planned
    
End Sub


Private Sub Class_Terminate()

    Set morsTxADACodes = Nothing
    Set moRads = Nothing
    
End Sub


