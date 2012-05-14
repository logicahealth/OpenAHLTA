VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MapperTheater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private objShared As CWShared
Implements DL_Interfaces.iMapperInpatientOrder

Private mConn As CHCSII_CONN.Conn
Private SQLOps As ICHCSII_DAS
Private mMapperFactory As DL_Interfaces.iDataMapperFactory
Private mCoreMapper    As DL_Interfaces.iMapperCore


Private Sub iMapperInpatientOrder_Init(Support As DL_Support.iParameters, Cache As DL_Support.CacheMapper)
On Error GoTo ErrHandler

    Set mConn = New CHCSII_CONN.Conn
    Set SQLOps = mConn.CHCSII_DAS(Auto)
    Set objShared = New CWShared

    Set mMapperFactory = Support.Item("DataMapperFactory").Value
    
    Exit Sub

ErrHandler:
  Err.Raise Err.Number, CACHE_NAME & ".MapperTheater.iMapperInpatientOrder_Init", Err.Description, Err.HelpFile, Err.HelpContext
End Sub

'--- SCR 88340  Sherry Wang  2/21/2006
Private Function iMapperInpatientOrder_NewOrders(Params As DL_Support.iParameters) As DTOs.InPatientOrders
    Dim sSelect     As String
    Dim sWhere      As String
    Dim oRs         As ADODB.Recordset
    Dim oParams     As DM_InPatOrder.Retrieve_Parameters
    Dim InPtOrder   As DTOs.InPatientOrder
    Dim objOrders   As DTOs.InPatientOrders
    
    On Error GoTo ErrHandler
    
    Set oParams = Params
       
    sSelect = "Select DATA_ID, UNIT_NUMBER, NURSE_ACK_FALG, ORDER_TYPE" & vbNewLine
    sWhere = " from V_ALL_ORDERS WHERE unit_number in " & oParams.UnitNumber & vbNewLine

    If LenB(oParams.Status) > 0 Then
        sWhere = sWhere & " And Status in " & oParams.Status & vbNewLine
    End If

    sWhere = sWhere & GetPatStatusWhere(oParams) & GetEncNumberWhere(oParams, False) & GetOccurrence
    
    Set oRs = SQLOps.OpenRecordset(sSelect & sWhere)
    
    If oRs.RecordCount > 0 Then
        Set objOrders = New DTOs.InPatientOrders
        oRs.MoveFirst
        Do Until oRs.EOF
            Set InPtOrder = New DTOs.InPatientOrder
            With InPtOrder
                .DataID = ReadStringField(oRs.Fields("DATA_ID"))
                .OrderType = ReadStringField(oRs.Fields("ORDER_TYPE"))
                .UnitNumber = ReadStringField(oRs.Fields("UNIT_NUMBER"))
                .NurseAck = CBool(ReadLongField(oRs.Fields("NURSE_ACK_FALG")))
            End With
            objOrders.Add InPtOrder
            oRs.MoveNext
            Set InPtOrder = Nothing
        Loop
    End If
    
    Set iMapperInpatientOrder_NewOrders = objOrders
    Set oRs = Nothing
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, CACHE_NAME & ".MapperTheater.iMapperInpatientOrder_NewOrders", Err.Description, Err.HelpFile, Err.HelpContext

End Function

Private Function iMapperInpatientOrder_Retrieve(Params As DL_Support.iParameters) As DTOs.InPatientOrders
    On Error GoTo ErrHandler
    Dim MyOrders  As DTOs.InPatientOrders
    Dim param     As DM_InPatOrder.Retrieve_Parameters
    
    Set MyOrders = New DTOs.InPatientOrders
    Set param = Params
    
    If LenB(param.OrderType) = 0 Then 'Get all orders
        Set MyOrders = GetInPatOrderLab(param, MyOrders)
        Set MyOrders = GetInPatOrderRad(param, MyOrders)
        Set MyOrders = GetInPatOrderRx(param, MyOrders)
        Set MyOrders = GetInPatOrderOther(param, MyOrders)
    ElseIf param.OrderType = "Lab" Then 'Get all orders
        Set MyOrders = GetInPatOrderLab(param, MyOrders)
    ElseIf param.OrderType = "Rad" Then 'Get all orders
        Set MyOrders = GetInPatOrderRad(param, MyOrders)
    ElseIf param.OrderType = "Med" Then 'Get all orders
        Set MyOrders = GetInPatOrderRx(param, MyOrders)
    ElseIf param.OrderType = "Other" Then 'Get all orders
        Set MyOrders = GetInPatOrderOther(param, MyOrders)
    End If
    
    Set iMapperInpatientOrder_Retrieve = MyOrders

    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, CACHE_NAME & ".MapperTheater.iMapperInpatientOrder_Retrieve", Err.Description, Err.HelpFile, Err.HelpContext
End Function


Private Function iMapperInpatientOrder_RetrieveInpatientID(Params As DL_Support.iParameters) As String
    On Error GoTo ErrHandler
    
    Dim param    As Retrieve_Parameters
    Dim sSql     As String
    Dim oRs      As ADODB.Recordset
    
    iMapperInpatientOrder_RetrieveInpatientID = vbNullString
    Set param = Params
    
    If Val(param.EncounterNumber) = 0 Then
        Exit Function
    End If
  
    sSql = "SELECT A.INPATIENT_ID"
    sSql = sSql & " FROM ENCOUNTERS E, INPATIENT_ADMISSION A"
    sSql = sSql & " WHERE E.APPTID = A.APPT_ID  "
    sSql = sSql & "       AND  E.ENCOUNTERNUMBER = " & param.EncounterNumber
    
    Set oRs = SQLOps.OpenRecordset(sSql)
    
    If oRs.RecordCount > 0 Then
        oRs.MoveFirst
        iMapperInpatientOrder_RetrieveInpatientID = oRs.Fields("INPATIENT_ID")
    End If
    
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, CACHE_NAME & ".MapperTheater.iMapperInpatientOrder_RetrieveInpatientID", Err.Description, Err.HelpFile, Err.HelpContext

End Function

Private Function iMapperInpatientOrder_Save(Orders As DTOs.InPatientOrders) As Boolean
On Error GoTo ErrHandler
Dim nOrderCount As Integer
Dim MyOrder As DTOs.InPatientOrder
    
    For nOrderCount = 1 To Orders.Count
        Set MyOrder = Orders.Item(nOrderCount)
        iMapperInpatientOrder_Save = UpdateInPtOrder(MyOrder)
    Next nOrderCount
    
    Exit Function

ErrHandler:
  Err.Raise Err.Number, CACHE_NAME & ".MapperTheater.iMapperInpatientOrder_Save", Err.Description, Err.HelpFile, Err.HelpContext
End Function



Private Function UpdateInPtOrder(ByVal InPtOrderData As DTOs.InPatientOrder) As Boolean
On Error GoTo ErrHandler

    UpdateInPtOrder = False
    Select Case UCase(InPtOrderData.OrderType)
        Case "MED"
            SaveMed InPtOrderData
        Case "RAD"
            SaveRad InPtOrderData
        Case "LAB"
            SaveLab InPtOrderData
        Case Else
            SaveOtherOrder InPtOrderData
    End Select
    
    UpdateInPtOrder = True
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, CACHE_NAME & ".MapperTheater.UpdateInPtOrder", Err.Description, Err.HelpFile, Err.HelpContext
End Function

Private Function GetInPatOrderLab(ByVal oParams As DM_InPatOrder.Retrieve_Parameters, _
                                  ByRef objOrders As DTOs.InPatientOrders) As DTOs.InPatientOrders
    Dim sSelect     As String
    Dim sWhere      As String
    Dim sSort       As String
    Dim oRs         As ADODB.Recordset
    Dim InPtOrder   As DTOs.InPatientOrder
    
    On Error GoTo ErrHandler
       
    sSelect = "Select OE_LAB_DATA_ID, ORDER_NAME , " & vbCrLf _
        & " STATUS, PREVIOUS_ORDER_STATUS, COMPLETED_DATE," & vbCrLf _
        & " ORDER_ID, TIMING_DURATION, TIMING_FREQUENCY, " & vbCrLf _
        & " ORDERED_BY, TIME_ORDERED, LAB_COMMENT ," & vbCrLf _
        & " SIGN_FLAG, NURSE_ACK_FALG, NURSE_ACK_NAME, NURSE_ACK_DATE, PATIENT_TYPE " & vbCrLf _

    sWhere = " From oe_lab WHERE unit_number = " & oParams.UnitNumber

    If LenB(oParams.Status) > 0 Then
        sWhere = sWhere & " And Status in " & oParams.Status
    End If

    sWhere = sWhere & GetPatStatusWhere(oParams) & GetEncNumberWhere(oParams, False) & GetOccurrence
    ' 88061 - Quan
    sSort = " Order by TIME_ORDERED ASC"
   
    Set oRs = SQLOps.OpenRecordset(sSelect & sWhere & sSort)
    
    If oRs.RecordCount > 0 Then
        oRs.MoveFirst
        Do Until oRs.EOF
            Set InPtOrder = New DTOs.InPatientOrder
            With InPtOrder
                .OrderType = "Lab"
                .OrderName = ReadStringField(oRs.Fields("ORDER_NAME"))
                .OrderStatus = ReadLongField(oRs.Fields("Status"))
                .OrderId = ReadStringField(oRs.Fields("ORDER_ID"))
                .Provider.Representation = ReadStringField(oRs.Fields("ORDERED_BY"))
                .OrderDate = oRs.Fields("TIME_ORDERED")
                .Comment = ReadStringField(oRs.Fields("LAB_COMMENT"))
                .Signed = ReadLongField(oRs.Fields("SIGN_FLAG"))
                .NurseAck = ReadLongField(oRs.Fields("NURSE_ACK_FALG"))
                .NurseAckName = ReadStringField(oRs.Fields("NURSE_ACK_NAME"))
                .NurseAckDate = ReadDateField(oRs.Fields("NURSE_ACK_DATE"))
                .PreviousOrderStatus = ReadLongField(oRs.Fields("PREVIOUS_ORDER_STATUS"))
                If ReadLongField(oRs.Fields("PATIENT_TYPE")) = 1 Then
                    .PatientType = "Inpatient"
                Else
                    .PatientType = "Outpatient"
                End If
                .Details = "Frequency: " & ReadStringField(oRs.Fields("TIMING_FREQUENCY")) _
                            & ", Duration: " & CStr(ReadLongField(oRs.Fields("TIMING_DURATION"))) & " day(s)"
                If .Details = "Frequency: , Duration: 0 day(s)" Then
                    .Details = vbNullString
                End If
                
                'SCR 79779 -- Frank 1/10/06
                .DataID = ReadStringField(oRs.Fields("OE_LAB_DATA_ID"))
                .IsMultiMedOrder = False
            End With
            objOrders.Add InPtOrder
            oRs.MoveNext
            Set InPtOrder = Nothing
        Loop
    End If
    
    Set GetInPatOrderLab = objOrders
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, CACHE_NAME & ".MapperTheater.GetInPatOrderLab", Err.Description, Err.HelpFile, Err.HelpContext
End Function

Private Function GetInPatOrderRad(ByVal oParams As DM_InPatOrder.Retrieve_Parameters, _
                                  ByRef objOrders As DTOs.InPatientOrders) As DTOs.InPatientOrders
    Dim sSelect     As String
    Dim sWhere      As String
    Dim sSort       As String
    Dim oRs         As ADODB.Recordset
    Dim InPtOrder   As DTOs.InPatientOrder
    
    On Error GoTo ErrHandler
    
    sSelect = "Select OE_RAD_DATA_ID, ORDER_NAME , " & vbCrLf _
        & " status, PREVIOUS_ORDER_STATUS, COMPLETED_DATE," & vbCrLf _
        & " ORDER_ID, TIMING_DURATION, TIMING_FREQUENCY, " & vbCrLf _
        & " ORDERED_BY, TIME_ORDERED, Rad_COMMENT ," & vbCrLf _
        & " SIGN_FLAG, NURSE_ACK_FALG, NURSE_ACK_NAME, NURSE_ACK_DATE, PATIENT_TYPE " & vbCrLf _
    
    sWhere = " From oe_Rad WHERE unit_number =   " & oParams.UnitNumber
    If LenB(oParams.Status) > 0 Then
        sWhere = sWhere & " And Status in " & oParams.Status
    End If
    
    sWhere = sWhere & GetPatStatusWhere(oParams) & GetEncNumberWhere(oParams, False) & GetOccurrence
    ' 88061 - Quan
    sSort = " Order by TIME_ORDERED ASC"
    
    Set oRs = SQLOps.OpenRecordset(sSelect & sWhere & sSort)
    If oRs.RecordCount > 0 Then
        oRs.MoveFirst
        Do Until oRs.EOF
            Set InPtOrder = New DTOs.InPatientOrder
            With InPtOrder
                .OrderType = "Rad"
                .OrderName = ReadStringField(oRs.Fields("ORDER_NAME"))
                .OrderStatus = ReadLongField(oRs.Fields("Status"))
                .OrderId = ReadStringField(oRs.Fields("ORDER_ID"))
                .Provider.Representation = ReadStringField(oRs.Fields("ORDERED_BY"))
                .OrderDate = oRs.Fields("TIME_ORDERED")
                .Comment = ReadStringField(oRs.Fields("Rad_COMMENT"))
                .Signed = ReadLongField(oRs.Fields("SIGN_FLAG"))
                .NurseAck = ReadLongField(oRs.Fields("NURSE_ACK_FALG"))
                .NurseAckName = ReadStringField(oRs.Fields("NURSE_ACK_NAME"))
                .NurseAckDate = ReadDateField(oRs.Fields("NURSE_ACK_DATE"))
                .PreviousOrderStatus = ReadLongField(oRs.Fields("PREVIOUS_ORDER_STATUS"))
                If ReadLongField(oRs.Fields("PATIENT_TYPE")) = 1 Then
                    .PatientType = "Inpatient"
                Else
                    .PatientType = "Outpatient"
                End If
                .Details = "Frequency: " & ReadStringField(oRs.Fields("TIMING_FREQUENCY")) _
                            & ", Duration: " & CStr(ReadLongField(oRs.Fields("TIMING_DURATION"))) & " day(s)"
                If .Details = "Frequency: , Duration: 0 day(s)" Then
                    .Details = vbNullString
                End If
                
                'SCR 79779 -- Frank 1/10/06
                .DataID = ReadStringField(oRs.Fields("OE_RAD_DATA_ID"))
                .IsMultiMedOrder = False
            End With
            objOrders.Add InPtOrder
            oRs.MoveNext
            Set InPtOrder = Nothing
        Loop
    End If
    
    Set GetInPatOrderRad = objOrders
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, CACHE_NAME & ".MapperTheater.GetInPatOrderRad", Err.Description, Err.HelpFile, Err.HelpContext
End Function
Private Function GetInPatOrderRx(ByVal oParams As DM_InPatOrder.Retrieve_Parameters, _
                                  ByRef objOrders As DTOs.InPatientOrders) As DTOs.InPatientOrders
    Dim sSelect     As String
    Dim sWhere      As String
    Dim sSort       As String
    Dim oRs         As ADODB.Recordset
    Dim InPtOrder   As DTOs.InPatientOrder
    
    On Error GoTo ErrHandler
   
    sSelect = "Select OE_RX_DATA_ID, ORDER_NAME , " & vbCrLf _
        & " STATUS_NCID AS status, FILL_DATE as COMPLETED_DATE," & vbCrLf _
        & " ORDER_ID, PREVIOUS_ORDER_STATUS, " & vbCrLf _
        & " ORDERED_BY, TIME_ORDERED, Rx_COMMENT, SIG_CODE, PARENT_DATA_ID, " & vbCrLf _
        & " SIGN_FLAG, NURSE_ACK_FALG, NURSE_ACK_NAME, NURSE_ACK_DATE, PATIENT_TYPE, ORDER_IEN, CHILD_MED_IEN" & vbCrLf
    
    sWhere = " From oe_Rx WHERE unit_number =  " & oParams.UnitNumber
    If LenB(oParams.Status) > 0 Then
        sWhere = sWhere & " And STATUS_NCID in " & GetStatusListforMed(oParams.Status)
    End If
    
    sWhere = sWhere & GetPatStatusWhere(oParams) & GetEncNumberWhere(oParams, False)
    ' 88061 - Quan
    sSort = " Order by TIME_ORDERED ASC"
    
    Set oRs = SQLOps.OpenRecordset(sSelect & sWhere & sSort)
    If oRs.RecordCount > 0 Then
        oRs.MoveFirst
        Do Until oRs.EOF
            Set InPtOrder = New DTOs.InPatientOrder
            With InPtOrder
                .OrderType = "Med"
                .OrderName = ReadStringField(oRs.Fields("ORDER_NAME"))
                .OrderStatus = ConvertFromMedStatus(CStr(ReadDoubleField(oRs.Fields("Status"))))
                .OrderId = ReadStringField(oRs.Fields("ORDER_ID"))
                .Provider.Representation = ReadStringField(oRs.Fields("ORDERED_BY"))
                .OrderDate = oRs.Fields("TIME_ORDERED")
                .Comment = ReadStringField(oRs.Fields("Rx_COMMENT"))
                .Signed = ReadLongField(oRs.Fields("SIGN_FLAG"))
                .NurseAck = ReadLongField(oRs.Fields("NURSE_ACK_FALG"))
                .NurseAckName = ReadStringField(oRs.Fields("NURSE_ACK_NAME"))
                .NurseAckDate = ReadDateField(oRs.Fields("NURSE_ACK_DATE"))
                .PreviousOrderStatus = ConvertFromMedStatus(CStr(ReadDoubleField(oRs.Fields("PREVIOUS_ORDER_STATUS"))))
                
                'SCR 79779 -- Frank 1/10/06
                .DataID = ReadStringField(oRs.Fields("OE_RX_DATA_ID"))
                
                If ReadLongField(oRs.Fields("PATIENT_TYPE")) = 1 Then
                    .PatientType = "Inpatient"
                    .IsPendingRefill = False    '*** 88669 2/24/06 - Frank; Not Applicable for inpatient med order
                Else
                    .PatientType = "Outpatient"
                     '** 88669 2/24/06 - Frank; pass in both OrderId and DataID
                    .IsPendingRefill = IsPendingRefill(.OrderId, .DataID)
                End If
                
                .Details = "SIG: " & ReadStringField(oRs.Fields("SIG_CODE"))
                'SCR 76468
                .OrderIEN = ReadStringField(oRs.Fields("ORDER_IEN"))
                .ChildMedIEN = ReadStringField(oRs.Fields("CHILD_MED_IEN"))
                
                .IsMultiMedOrder = getMultiOrderCount(.OrderIEN)
                                
                .IsContinuousMedOrder = IsContinuousMedOrder(.DataID)
                
                .ParentDataID = ReadStringField(oRs.Fields("PARENT_DATA_ID"))

            End With
            objOrders.Add InPtOrder
            oRs.MoveNext
            Set InPtOrder = Nothing
        Loop
    End If
    
    Set GetInPatOrderRx = objOrders
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, CACHE_NAME & ".MapperTheater.GetInPatOrderRx", Err.Description, Err.HelpFile, Err.HelpContext
End Function

Private Function GetInPatOrderOther(ByRef oParams As DM_InPatOrder.Retrieve_Parameters, _
                                    ByRef objOrders As DTOs.InPatientOrders) As DTOs.InPatientOrders
    Dim sSelect     As String
    Dim sWhere      As String
    Dim oRs         As ADODB.Recordset
    Dim InPtOrder   As DTOs.InPatientOrder
    
    Dim sComment    As String
    
    On Error GoTo ErrHandler
    
    If oParams.PatientStatus = isOutpatient Then
        Set GetInPatOrderOther = objOrders
        Exit Function
    End If
   
    sSelect = "SELECT OE_OTR_DATA_ID, UNIT_NUMBER, " & vbNewLine
    sSelect = sSelect & " FACILITY_NCID, ENCOUNTER_NUMBER, ORDER_TYPE, " & vbNewLine
    sSelect = sSelect & " ORDER_TYPE_DESC, ORDER_COMMENTS, ANCILLARY_COMMENT, ORDER_DATE, " & vbNewLine
    sSelect = sSelect & " ORDERED_BY, ORDER_STATUS, NURSE_ACK_FLAG,  " & vbNewLine
    sSelect = sSelect & " SIGN_FLAG, NURSE_ACK_NAME, NURSE_ACK_DATE, PREVIOUS_ORDER_STATUS " & vbNewLine

    sWhere = " From oe_Other WHERE unit_number =  " & oParams.UnitNumber & vbNewLine
    If LenB(oParams.Status) > 0 Then
        sWhere = sWhere & " And ORDER_STATUS in " & oParams.Status & vbNewLine
    End If
    If LenB(oParams.OtherOrderType) > 0 Then
        sWhere = sWhere & " And ORDER_TYPE = " & mConn.SQLQuote(oParams.OtherOrderType)
    End If
    
    If LenB(oParams.OrderId) > 0 Then
        sWhere = sWhere & " And OE_OTR_DATA_ID = " & oParams.OrderId
    End If
    
    If oParams.EncounterNumber = vbNullString Then
        sWhere = sWhere & GetEncNumberWhere(oParams, True)
    End If
    
    Set oRs = SQLOps.OpenRecordset(sSelect & sWhere)
    If oRs.RecordCount > 0 Then
        oRs.MoveFirst
        Do Until oRs.EOF
            Set InPtOrder = New DTOs.InPatientOrder
            With InPtOrder
                .OrderType = ReadStringField(oRs.Fields("ORDER_TYPE"))
                .OrderName = ReadStringField(oRs.Fields("ORDER_TYPE_DESC"))
                .OrderStatus = ReadLongField(oRs.Fields("ORDER_STATUS"))
                .OrderId = ReadLongField(oRs.Fields("OE_OTR_DATA_ID"))
                .Provider.Representation = ReadStringField(oRs.Fields("ORDERED_BY"))
                .OrderDate = ReadDateField(oRs.Fields("ORDER_DATE"))
                .Comment = ReadStringField(oRs.Fields("ORDER_COMMENTS"))
                .Signed = ReadLongField(oRs.Fields("SIGN_FLAG"))
                .NurseAck = ReadLongField(oRs.Fields("NURSE_ACK_FLAG"))
                .NurseAckDate = ReadDateField(oRs.Fields("NURSE_ACK_DATE"))
                .NurseAckName = ReadStringField(oRs.Fields("NURSE_ACK_NAME"))
                .UnitNumber = ReadLongField(oRs.Fields("UNIT_NUMBER"))
                .FacilityNCID = ReadLongField(oRs.Fields("FACILITY_NCID"))
                
                '*** Pre 1/18/06 ORDER_COMMENTS contains two field separate by |
                sComment = .Comment
                If InStr(1, sComment, "|") > 0 Then
                    .Comment = Mid$(sComment, 1, InStr(1, sComment, "|") - 1)
                    .AncillaryComment = Mid$(sComment, InStr(1, sComment, "|") + 1)
                Else
                     'SCR 85951 -- Frank 1/18/06
                    .AncillaryComment = ReadStringField(oRs.Fields("ANCILLARY_COMMENT"))
                End If
                
                .PreviousOrderStatus = ReadLongField(oRs.Fields("PREVIOUS_ORDER_STATUS"))
                .PatientType = "Inpatient"
                
                'SCR 79779 -- Frank 1/10/06
                .DataID = ReadStringField(oRs.Fields("OE_OTR_DATA_ID"))
                .IsMultiMedOrder = False
                
               
            End With
            objOrders.Add InPtOrder
            oRs.MoveNext
            Set InPtOrder = Nothing
        Loop
    End If
    
    Set GetInPatOrderOther = objOrders
    
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, CACHE_NAME & ".MapperTheater.GetInPatOrderOther", Err.Description, Err.HelpFile, Err.HelpContext
End Function

Private Sub SaveOtherOrder(ByRef oOtherOrder As DTOs.InPatientOrder)
    Dim sDataID As String
    Dim oRs As ADODB.Recordset
    Dim sSql  As String
    
    On Error GoTo ErrHandler
    
    Dim sSysDate As String
    
    If objShared.UseMySQLDB Then
        sSysDate = "sysdate()"
    Else
        sSysDate = "sysdate"
    End If
    
    sSql = "SELECT * FROM OE_OTHER WHERE OE_OTR_DATA_ID = " & Val(oOtherOrder.OrderId)
    Set oRs = SQLOps.OpenRecordset(sSql)
    If oRs.RecordCount = 0 Then
        sDataID = SQLOps.GetNextID
        With oOtherOrder
            sSql = "INSERT INTO OE_OTHER " & vbNewLine
            sSql = sSql & "(OE_OTR_DATA_ID, UNIT_NUMBER, FACILITY_NCID, ENCOUNTER_NUMBER, " & vbNewLine
            sSql = sSql & "ORDER_TYPE, ORDER_TYPE_DESC, ORDER_COMMENTS, ORDER_DATE, ANCILLARY_COMMENT," & vbNewLine
            sSql = sSql & "ORDERED_BY, ORDERED_BY_NCID, ORDER_STATUS, PREVIOUS_ORDER_STATUS, SIGN_FLAG)"
            sSql = sSql & "VALUES (" & sDataID & ", " & .UnitNumber & ", " & vbNewLine
            sSql = sSql & .FacilityNCID & ", " & .EncounterNumber & ", " & vbNewLine
            sSql = sSql & mConn.SQLQuote(.OrderType) & ", " & mConn.SQLQuote(.OrderName) & ", " & vbNewLine
            sSql = sSql & mConn.SQLQuote(.Comment) & ", " & sSysDate & ", " & vbNewLine
            sSql = sSql & mConn.SQLQuote(.AncillaryComment) & ", " & vbNewLine
            sSql = sSql & mConn.SQLQuote(.Provider.Representation) & ", " & vbNewLine
            sSql = sSql & mConn.SQLQuote(.Provider.NCID) & ", " & vbNewLine
            sSql = sSql & .OrderStatus & ", " & .PreviousOrderStatus & "," & CLng(.Signed) & ")" & vbNewLine
        End With
        
        SQLOps.ExecuteSQL sSql
        
        oOtherOrder.OrderId = sDataID
    Else
        With oOtherOrder
            sSql = "UPDATE OE_OTHER " & vbNewLine
            sSql = sSql & " SET ORDER_STATUS = " & .OrderStatus & ", " & vbNewLine
            sSql = sSql & " PREVIOUS_ORDER_STATUS = " & .PreviousOrderStatus & ", " & vbNewLine
            sSql = sSql & " NURSE_ACK_FLAG = " & CLng(.NurseAck) & ", " & vbNewLine
            sSql = sSql & " SIGN_FLAG = " & CLng(.Signed) & ", " & vbNewLine
            sSql = sSql & " NURSE_ACK_NAME = " & mConn.SQLQuote(.NurseAckName) & ", " & vbNewLine
            sSql = sSql & " ORDER_COMMENTS = " & mConn.SQLQuote(.Comment) & ", " & vbNewLine
            sSql = sSql & " ANCILLARY_COMMENT = " & mConn.SQLQuote(.AncillaryComment) & ", " & vbNewLine
            sSql = sSql & " NURSE_ACK_DATE = " & mConn.FormatDateForDBInsert(.NurseAckDate) & vbNewLine
            sSql = sSql & " WHERE OE_OTR_DATA_ID = " & .OrderId & vbNewLine
        End With
        
        SQLOps.ExecuteSQL sSql
    End If
    
    'SCR 64289  - 1/13/05 Sherry Wang.  Remove Unassociated Med Order in Med Order tab
    'when a distinued lab form InpatientOrder module.
    If oOtherOrder.OrderStatus = isDiscontinued Then
        sSql = "Delete from DIAGNOSESRELATIONSHIPS " _
            & "  where RELATEDDATAID = " & oOtherOrder.OrderId

        SQLOps.ExecuteSQL sSql
    End If
    
    Exit Sub
    
ErrHandler:
    Err.Raise Err.Number, CACHE_NAME & ".MapperTheater.SaveOtherOrder", Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Private Sub SaveLab(ByRef oOrder As DTOs.InPatientOrder)
    Dim sSql  As String
    
    On Error GoTo ErrHandler
    
    If oOrder.OrderStatus = isDiscontinued Then
    
        'SCR 83145 - 1/11/2006 Jane Shen. Delete children order if parent status is pending
        sSql = "DELETE FROM OE_LAB WHERE ORDER_ID = '" & oOrder.OrderId & _
                    "' AND STATUS = " & isPending & " AND IS_ORDER_OCCURRENCE = 1"
        
        SQLOps.ExecuteSQL sSql
        
        sSql = "Delete from DIAGNOSESRELATIONSHIPS " _
            & "  where RELATEDDATAID in (select OE_LAB_DATA_ID from  OE_LAB WHERE ORDER_ID = " & mConn.SQLQuote(oOrder.OrderId) & ")"

        SQLOps.ExecuteSQL sSql
    End If
    
    With oOrder
        sSql = "UPDATE OE_LAB " & vbNewLine
        sSql = sSql & " SET SIGN_FLAG = " & CLng(.Signed) & ", " & vbNewLine
        sSql = sSql & " STATUS = " & CLng(.OrderStatus) & ", " & vbNewLine
        sSql = sSql & " PREVIOUS_ORDER_STATUS = " & CLng(.PreviousOrderStatus) & ", " & vbNewLine
        sSql = sSql & " NURSE_ACK_FALG = " & CLng(.NurseAck) & ", " & vbNewLine
        sSql = sSql & " NURSE_ACK_NAME = " & mConn.SQLQuote(.NurseAckName) & ", " & vbNewLine
        sSql = sSql & " NURSE_ACK_DATE = " & mConn.FormatDateForDBInsert(.NurseAckDate) & vbNewLine
        
        'SCR 79779 -- Frank 1/10/06
'        sSQL = sSQL & " WHERE ORDER_ID = " & mConn.SQLQuote(.OrderId) & vbNewLine
        sSql = sSql & " WHERE OE_LAB_DATA_ID = " & .DataID & vbNewLine
        
        ' Quan 81814, 81815
        'sSQL = sSQL & " AND STATUS <> " & CStr(isCompleted_Unaddr) & vbNewLine
        'sSQL = sSQL & " AND STATUS <> " & CStr(isCompleted_Addr) & vbNewLine
        'sSQL = sSQL & " AND STATUS <> " & CStr(isPreliminary) & vbNewLine
    End With
    
    SQLOps.ExecuteSQL sSql
    
'    'SCR 64289  - 1/13/05 Sherry Wang.  Remove Unassociated Lab Order
'    'when a distinued lab form InpatientOrder module.
'    If oOrder.OrderStatus = isDiscontinued Then
'        sSQL = "Delete from DIAGNOSESRELATIONSHIPS " _
'            & "  where RELATEDDATAID in (select OE_LAB_DATA_ID from  OE_LAB WHERE ORDER_ID = " & mConn.SQLQuote(oOrder.OrderId) & ")"
'
'        SQLOps.ExecuteSQL sSQL
'    End If
    
    Exit Sub
    
ErrHandler:
    Err.Raise Err.Number, CACHE_NAME & ".MapperTheater.SaveLab", Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Private Sub SaveRad(ByRef oOrder As DTOs.InPatientOrder)
    Dim sSql  As String
    
    On Error GoTo ErrHandler
    
    If oOrder.OrderStatus = isDiscontinued Then
        'SCR 83145 - 1/11/2006 Jane Shen. Delete children order if parent status is pending
        sSql = "DELETE FROM OE_RAD WHERE ORDER_ID = '" & oOrder.OrderId & _
                    "' AND STATUS = " & isPending & " AND IS_ORDER_OCCURRENCE = 1"
        
        SQLOps.ExecuteSQL sSql
        
        sSql = "Delete from DIAGNOSESRELATIONSHIPS " _
            & "  where RELATEDDATAID in (select OE_RAD_DATA_ID from  OE_RAD WHERE ORDER_ID = " & mConn.SQLQuote(oOrder.OrderId) & ")"

        SQLOps.ExecuteSQL sSql
    End If
    
    With oOrder
        sSql = "UPDATE OE_RAD " & vbNewLine
        sSql = sSql & " SET SIGN_FLAG = " & CLng(.Signed) & ", " & vbNewLine
        sSql = sSql & " STATUS = " & CLng(.OrderStatus) & ", " & vbNewLine
        sSql = sSql & " PREVIOUS_ORDER_STATUS = " & CLng(.PreviousOrderStatus) & ", " & vbNewLine
        sSql = sSql & " NURSE_ACK_FALG = " & CLng(.NurseAck) & ", " & vbNewLine
        sSql = sSql & " NURSE_ACK_NAME = " & mConn.SQLQuote(.NurseAckName) & ", " & vbNewLine
        sSql = sSql & " NURSE_ACK_DATE = " & mConn.FormatDateForDBInsert(.NurseAckDate) & vbNewLine
        
        'SCR 79779 -- Frank 1/10/06
        'sSQL = sSQL & " WHERE ORDER_ID = " & mConn.SQLQuote(.OrderId) & vbNewLine
        sSql = sSql & " WHERE OE_RAD_DATA_ID = " & .DataID & vbNewLine
        
        ' Quan 81814, 81815
        'sSQL = sSQL & " AND STATUS <> " & CStr(isCompleted_Unaddr) & vbNewLine
        'sSQL = sSQL & " AND STATUS <> " & CStr(isCompleted_Addr) & vbNewLine
        'sSQL = sSQL & " AND STATUS <> " & CStr(isPreliminary) & vbNewLine
    End With
    
    SQLOps.ExecuteSQL sSql
    
'    'SCR 64289  - 1/13/05 Sherry Wang.  Remove Unassociated Rad Order
'    'when a distinued lab form InpatientOrder module.
'    If oOrder.OrderStatus = isDiscontinued Then
'        sSQL = "Delete from DIAGNOSESRELATIONSHIPS " _
'            & "  where RELATEDDATAID in (select OE_RAD_DATA_ID from  OE_RAD WHERE ORDER_ID = " & mConn.SQLQuote(oOrder.OrderId) & ")"
'
'        SQLOps.ExecuteSQL sSQL
'    End If
    
    Exit Sub
    
ErrHandler:
    Err.Raise Err.Number, CACHE_NAME & ".MapperTheater.SaveRad", Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Private Sub SaveMed(ByRef oOrder As DTOs.InPatientOrder)
    Dim sSql  As String
    
    On Error GoTo ErrHandler
    
    '--- UPDATE OR_RX TABLE
    With oOrder
        sSql = "UPDATE OE_RX " & vbNewLine
        sSql = sSql & " SET SIGN_FLAG = " & CLng(.Signed) & ", " & vbNewLine
        sSql = sSql & " STATUS_NCID = " & ConvertToMedStatus(.OrderStatus) & ", " & vbNewLine
        sSql = sSql & " PREVIOUS_ORDER_STATUS = " & ConvertToMedStatus(.PreviousOrderStatus) & ", " & vbNewLine
        sSql = sSql & " NURSE_ACK_FALG = " & CLng(.NurseAck) & ", " & vbNewLine
        sSql = sSql & " NURSE_ACK_NAME = " & mConn.SQLQuote(.NurseAckName) & ", " & vbNewLine
        sSql = sSql & " NURSE_ACK_DATE = " & mConn.FormatDateForDBInsert(.NurseAckDate) & vbNewLine
        
        'SCR 79779 -- Frank 1/10/06
        sSql = sSql & " WHERE OE_RX_DATA_ID = " & .DataID & vbNewLine
        
    End With
    
    SQLOps.ExecuteSQL sSql
    
    '--- UPDATE MEDICATIONS TABLE
    With oOrder
        sSql = "UPDATE MEDICATIONS " & vbNewLine
        sSql = sSql & " SET STATUSNCID = " & ConvertToMedStatus(.OrderStatus) & vbNewLine

        'SCR 79779 -- Frank 1/10/06
        sSql = sSql & " WHERE DATAID = " & .DataID
    End With
    
    SQLOps.ExecuteSQL sSql
    
    'SCR 64024 - 1/12/05 Frank.  Remove Unassociated Med Order in Med Order tab
    'when a distinued med form InpatientOrder module.
    If ConvertToMedStatus(oOrder.OrderStatus) = NCID_ORDERSTATUS_DISCONTINUED Then
        sSql = "Delete from DIAGNOSESRELATIONSHIPS " _
            & "  where RELATEDDATAID in (select OE_RX_DATA_ID from  oe_rx WHERE ORDER_ID = " & mConn.SQLQuote(oOrder.OrderId) & ")"

        SQLOps.ExecuteSQL sSql
    End If


    Exit Sub
    
ErrHandler:
    Err.Raise Err.Number, CACHE_NAME & ".MapperTheater.SaveMed", Err.Description, Err.HelpFile, Err.HelpContext
End Sub


Private Function GetPatStatusWhere(ByVal oParams As DM_InPatOrder.Retrieve_Parameters) As String
    Dim sWhere As String
    On Error GoTo ErrHandler
    
    With oParams
        If oParams.PatientStatus = isBoth Then
            sWhere = vbNullString
        ElseIf oParams.PatientStatus = IsInpatient Then
            sWhere = " and PATIENT_TYPE = 1"
        ElseIf oParams.PatientStatus = isOutpatient Then
            sWhere = " and (PATIENT_TYPE <> 1 OR PATIENT_TYPE IS NULL)"
        End If
        
    End With
    
    GetPatStatusWhere = sWhere
    
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, CACHE_NAME & ".MapperTheater.GetPatStatusWhere", Err.Description, Err.HelpFile, Err.HelpContext
End Function
Private Function GetOccurrence() As String
    GetOccurrence = " And IS_ORDER_OCCURRENCE = 0 "
End Function


Private Function GetEncNumberWhere(ByVal oParams As DM_InPatOrder.Retrieve_Parameters, ByVal bOtherOrder As Boolean) As String
    Dim sEncList As String
    Dim sWhere   As String
    Dim sSql     As String
    Dim oRs      As ADODB.Recordset
    
    On Error GoTo ErrHandler
    
    sWhere = vbNullString
    
    If LenB(oParams.InpatientId) = 0 Or oParams.InpatientId = "0" Then
NoEncFound:
        If bOtherOrder = True Then
            GetEncNumberWhere = " AND ENCOUNTER_NUMBER IS NULL"
        Else
            GetEncNumberWhere = " AND ENC_NUM IS NULL"
        End If
        Exit Function
    End If
    
    With oParams
        If LenB(.InpatientId) > 0 And .InpatientId <> "0" Then
            sSql = "SELECT E.ENCOUNTERNUMBER"
            sSql = sSql & " FROM ENCOUNTERS E, INPATIENT_ADMISSION A"
            sSql = sSql & " WHERE E.APPTID = A.APPT_ID  "
            
            If InStr(1, .InpatientId, "(") > 0 Then
                sSql = sSql & "       AND  A.INPATIENT_ID in " & .InpatientId
            Else
                sSql = sSql & "       AND  A.INPATIENT_ID = " & .InpatientId
            End If
            
            Set oRs = SQLOps.OpenRecordset(sSql)
            If oRs.RecordCount > 0 Then
                oRs.MoveFirst
                Do Until oRs.EOF
                    If bOtherOrder = True Then
                        sEncList = sEncList & oRs("ENCOUNTERNUMBER") & ", "
                    Else
                        sEncList = sEncList & "'" & oRs("ENCOUNTERNUMBER") & "', "
                    End If
                    oRs.MoveNext
                Loop
                
                sEncList = Left$(sEncList, Len(sEncList) - 2)
                
                If bOtherOrder = True Then
                    sWhere = " AND ENCOUNTER_NUMBER IN (" & sEncList & ")"
                Else
                    sWhere = " AND ENC_NUM IN (" & sEncList & ")"
                End If
            Else
                GoTo NoEncFound
            End If
        End If
    End With
    
    GetEncNumberWhere = sWhere
    
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, CACHE_NAME & ".MapperTheater.GetEncNumberWhere", Err.Description, Err.HelpFile, Err.HelpContext
End Function

Public Function getMultiOrderCount(ByVal order_ien As String) As Boolean
    
    'SCR 79779 -- Frank 1/10/06
    
    Dim strSQL As String
    Dim rs     As Recordset
    
On Error GoTo ErrHandler
    
    strSQL = "Select count(*) from oe_rx where order_ien = " & order_ien
    strSQL = strSQL & " and STATUS_NCID <> " & NCID_ORDERSTATUS_DISCONTINUED
    
    Set rs = SQLOps.OpenRecordset(strSQL)
    
    getMultiOrderCount = Val(rs.Fields.Item(0).Value) > 1
    rs.Close
    Exit Function

ErrHandler:
    Err.Raise Err.Number, CACHE_NAME & ".MapperTheater.getMultiOrderCount", Err.Description, Err.HelpFile, Err.HelpContext
    
End Function


Public Function IsContinuousMedOrder(DataID As String) As Boolean
    Dim sSql    As String
    Dim rs      As Recordset
    
On Error GoTo ErrHandler
    
    If DataID = "" Then
        IsContinuousMedOrder = False
    Else
        sSql = "Select count(*) from oe_rx "
        sSql = sSql & " where parent_data_id = " & DataID
        
        Set rs = SQLOps.OpenRecordset(sSql)
        
        IsContinuousMedOrder = Val(rs.Fields.Item(0).Value) > 0
        rs.Close
    End If
    
    Exit Function

ErrHandler:
    Err.Raise Err.Number, CACHE_NAME & ".MapperTheater.IsContinuousMedOrder", Err.Description, Err.HelpFile, Err.HelpContext
    
End Function


Public Function IsPendingRefill(ByVal OrderId As String, ByVal DataID As String) As Boolean
    
    '** SCR 88669 - 2/24/06 - Frank (Change pass in arguement from OrderID
    '** Order ID can not be used to uniquely identify the duplicate row
    
    '** Use DataID to uniquely identified the dupliate refill row
    '** Return true, if this record is the Pending outpatient refill med order
    
    Dim rs      As Recordset
    Dim sSql    As String
    
    On Error GoTo ErrHandler
    
    IsPendingRefill = False
    
    
    '*********************************************
    '** SCR 88669 - 2/24/06 - Frank
    '** Check to see if there's duplicate refill exist
    sSql = "SELECT Count(*) FROM OE_RX"
    sSql = sSql & " WHERE ORDER_ID = '" & OrderId & "'"
    sSql = sSql & " AND PATIENT_TYPE = 0"
    
    Set rs = SQLOps.OpenRecordset(sSql)
    
    If Not (rs.BOF And rs.EOF) Then
        '** Find Pending Refill med order
        '** This could also be on hold (PREVIOUS_ORDER_STATUS = PENDING)
        
        '** SCR 88669 - 2/24/06 - Frank
        '** This order has refill, find which one it is (Active or Pending)
        If rs.Fields(0).Value > 1 Then
        
            sSql = "SELECT Count(*) FROM OE_RX"
            sSql = sSql & " WHERE OE_RX_DATA_ID = " & DataID
            sSql = sSql & " AND PATIENT_TYPE = 0 " & " AND (STATUS_NCID = " & NCID_PARTIAL
            sSql = sSql & " OR PREVIOUS_ORDER_STATUS = " & NCID_PARTIAL & ")"
            
            Set rs = SQLOps.OpenRecordset(sSql)
            
            '** Only one pending refill med order, exist
            '** the other one is Active (both have same Order ID)
            If Not (rs.BOF And rs.EOF) Then
                IsPendingRefill = Val(rs.Fields(0).Value & vbNullString) > 0
            End If
        End If
    End If
    
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, CACHE_NAME & ".MapperTheater.IsPendingRefill", Err.Description, Err.HelpFile, Err.HelpContext

End Function






