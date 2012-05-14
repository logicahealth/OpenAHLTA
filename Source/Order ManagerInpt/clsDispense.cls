VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDispense"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"WorkStatus"
Attribute VB_Ext_KEY = "Member0" ,"WorkStatus"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private msComment   As String
Private msOverideReason As String
Private msOverideComment As String

Private msRefills   As String
Private msRefillRem As String
Private msQuantity  As String
Private msName      As String
Private msSig       As String
Private msExpSig    As String
Private msOrderId   As String
Private msMedID     As String
Private mIsInpatient As Boolean
Private mNurseAck As Boolean
Private mOrderStaus As Long
Private mFillDate   As Date
Private msEndDate As String    'SCR 85433 -- 1/10/06 - Frank

Public Property Get EndDate() As String
    EndDate = msEndDate
End Property
Public Property Let EndDate(ByVal vData As String)
    msEndDate = vData
End Property

Public Property Let OrderStaus(ByVal vData As Long)
    mOrderStaus = vData
End Property
Public Property Get OrderStaus() As Long
    OrderStaus = mOrderStaus
End Property

Public Property Let IsInpatient(ByVal vData As Boolean)
    mIsInpatient = vData
End Property
Public Property Get IsInpatient() As Boolean
    IsInpatient = mIsInpatient
End Property

Public Property Let NurseAck(ByVal vData As Boolean)
    mNurseAck = vData
End Property
Public Property Get NurseAck() As Boolean
    NurseAck = mNurseAck
End Property

Public Property Let Name(ByVal vData As String)
    msName = vData
End Property
Public Property Get Name() As String
    Name = msName
End Property

Public Property Let MedID(ByVal vData As String)
    msMedID = vData
End Property
Public Property Get MedID() As String
    MedID = msMedID
End Property

Public Property Let OrderID(ByVal vData As String)
    msOrderId = vData
End Property

Public Property Get OrderID() As String
    OrderID = msOrderId
End Property

Public Property Let Sig(ByVal vData As String)
    msSig = vData
End Property

Public Property Get Sig() As String
    Sig = msSig
End Property

Public Property Let ExpSig(ByVal vData As String)
    msExpSig = vData
End Property

Public Property Get ExpSig() As String
    ExpSig = msExpSig
End Property

Public Property Let Comment(ByVal vData As String)
    msComment = vData
End Property

Public Property Get Comment() As String
    Comment = msComment
End Property

Public Property Let OverideComment(ByVal vData As String)
    msOverideComment = vData
End Property

Public Property Get OverideComment() As String
    OverideComment = msOverideComment
End Property

Public Property Let OverideReason(ByVal vData As String)
    msOverideReason = vData
End Property

Public Property Get OverideReason() As String
    OverideReason = msOverideReason
End Property

Public Property Let Refills(ByVal vData As String)
    msRefills = vData
End Property

Public Property Get Refills() As String
    Refills = msRefills
End Property

Public Property Let RefillRem(ByVal vData As String)
    msRefillRem = vData
End Property

Public Property Get RefillRem() As String
    RefillRem = msRefillRem
End Property

Public Property Let Quantity(ByVal vData As String)
    msQuantity = vData
End Property

Public Property Get Quantity() As String
    Quantity = msQuantity
End Property

Public Property Let FillDate(ByVal vData As Date)
    mFillDate = vData
End Property

Public Property Get FillDate() As Date
    FillDate = msQuantity
End Property

Public Sub saveDispense(ByVal OrderID As String, Optional CallFromGemsOE As Boolean = False, Optional IsFirstTimeDispense As Boolean = False, Optional IsInpatientMedOrder As Boolean, Optional bCreateNextDayMedOrder As Boolean)

    Dim strSQL      As String
    Dim rxSQLStatus As String
    Dim medSQLStatus As String
    Dim whereRXOrder As String
    Dim whereMEDOrder As String

    Dim OriginalDataID As String
    Dim oDR As ADODB.Recordset
    
    On Error GoTo ErrHnd
    

    'SCR 135000
    'Refills = Val(frmDispense.txtRefillsRemaining)
    Quantity = Val(frmDispense.txtQuan)
    '** Synchup the SIG with updated Refills/Quantity
    'Sig = frmDispense.UpdateSigWhenRefillsChanged  ' SCR 107029, jq, 5/24/07: don't change the SIG
    
    rxSQLStatus = " STATUS_NCID = " & NCID_ORDERSTATUS_ACTIVE
    medSQLStatus = " STATUSNCID = " & NCID_ORDERSTATUS_ACTIVE
    
    whereRXOrder = "  ORDER_ID = '" & OrderID & "'"
    whereMEDOrder = "  ORDERNUMBER = '" & OrderID & "'"

    If getMultiOrderCount(OrderID) Then
       
        '** Get ORIGIONAL RELATIONSHIP
        strSQL = "SELECT * FROM DIAGNOSESRELATIONSHIPS " _
                 & "WHERE RELATEDDATAID = (SELECT OE_RX_DATA_ID FROM OE_RX WHERE " _
                         & rxSQLStatus & " and " & whereRXOrder & ")"
        
        Set oDR = goBB.I_DAS.OpenRecordset(strSQL)
        
        If Not (oDR.BOF And oDR.EOF) Then
            OriginalDataID = oDR.Fields("RELATEDDATAID")
        End If
            
        '** This cause the ORIGINAL RELATIONSHIP TO BE DELETED
        strSQL = "Delete oe_rx where "
        strSQL = strSQL & rxSQLStatus & " and " & whereRXOrder
        goBB.I_DAS.ExecuteSQL strSQL

        '*** Preserve the relationship [COMPLETED RECORD DELTED FROM OE_RX, NEW PENDING
        '*** RECORD MUST USE ORIGINAL DATA_ID (otherwise it won't show up in A&P]
        
        '** 89055 - 3/2/06 - Frank; Applicable only when there's diagnosis
        '** relationship exist with the originial med order
        If Val(OriginalDataID) > 0 Then
            strSQL = "update oe_rx set "
            strSQL = strSQL & " OE_RX_DATA_ID = " & Val(OriginalDataID)
            strSQL = strSQL & " where " & whereRXOrder
            goBB.I_DAS.ExecuteSQL strSQL
    
            Set oDR = Nothing
        End If
    End If
   
    strSQL = "update oe_rx set "
    strSQL = strSQL & " Refills = " & Val(RefillRem)
    strSQL = strSQL & ", FILL_DATE = " & goBB.DataConn.FormatDateForDBInsert(mFillDate)
    strSQL = strSQL & ", Quantity = " & Val(Quantity)
    strSQL = strSQL & ", RX_COMMENT = " & goBB.DataConn.SQLQuote(Comment, 255)

    '*** SCR -- 4/13/05 - Frank
    '*** Save updated sig
    '... SCR 107029, jq, 5/24/07: requirement change -- keep original SIG
    '... SCR 134546, jq, 7/5/07; we need to change it when the user updates qty or refills
    strSQL = strSQL & ", Sig_Code = " & goBB.DataConn.SQLQuote(Sig, 100)
    
    '** 70656 - 5/27/05 - Frank
    '** For INPATIENT MED ORDER, when NOT CREATING NEXT DAY MED ORDER, Set Status to COMPLETE
    If IsInpatientMedOrder And (Not bCreateNextDayMedOrder) Then
        strSQL = strSQL & ", STATUS_NCID = " & NCID_ORDERSTATUS_COMPLETE
    Else
        strSQL = strSQL & ", STATUS_NCID = " & NCID_ORDERSTATUS_ACTIVE
    End If
    
    strSQL = strSQL & " where " & whereRXOrder
    
    goBB.I_DAS.ExecuteSQL strSQL
    
    strSQL = "update Medications set "
    'SCR 135000
    strSQL = strSQL & " Refills = " & Val(Refills)  ' jq: this sets original refills to refills remaining
    strSQL = strSQL & ", REFILLSREMAINING = " & Val(RefillRem)
    
    strSQL = strSQL & ", DISPENSINGLOCATIONNCID = " & goBB.Logon.FacilityNCID
    strSQL = strSQL & ", Quantity = " & Val(Quantity)
    strSQL = strSQL & ", MEDICATIONSCOMMENT = " & goBB.DataConn.SQLQuote(Comment, 255)
    strSQL = strSQL & ", FILLDATE =  " & goBB.DataConn.FormatDateForDBInsert(mFillDate)
    
    '** 70656 - 5/27/05 - Frank
    '** For INPATIENT MED ORDER, when NOT CREATING NEXT DAY MED ORDER, Set Status to COMPLETE
    If IsInpatientMedOrder And (Not bCreateNextDayMedOrder) Then
        strSQL = strSQL & ", STATUSNCID = " & NCID_ORDERSTATUS_COMPLETE
    Else
        strSQL = strSQL & ", STATUSNCID = " & NCID_ORDERSTATUS_ACTIVE
    End If
    
    
    '*** SCR -- 4/13/05 - Frank
    '*** Save updated sig
    strSQL = strSQL & ", Sig = " & goBB.DataConn.SQLQuote(Sig, 50)
    
    strSQL = strSQL & " where " & whereMEDOrder
    
    goBB.I_DAS.ExecuteSQL strSQL
        
    If Len(goBB.SharedObj.CmdLineSwitches("INPATIENT2")) > 0 Then
        strSQL = "update rx_list_FDB set current_stock = current_stock - " & Val(Quantity)
        strSQL = strSQL & " where medid = (select distinct order_ien from oe_rx where ORDER_ID = '" & OrderID & "')"
    Else
        strSQL = "update rx_list set current_stock = current_stock - " & Val(Quantity)
        strSQL = strSQL & " where drug_ien = (select distinct order_ien from oe_rx where ORDER_ID = '" & OrderID & "')"
    End If
    
    Call goBB.I_DAS.ExecuteSQL(strSQL)
    
    
    '*** 5/11/05 - Frank
    '*** After dispense inpatient med order, create a new pending med order for next day
    If IsInpatientMedOrder Then
        
        '** Do this for inpatient med order with create next day's order only
        Call UpdatePrevActiveToComplete(OrderID)
    
        If bCreateNextDayMedOrder Then
            Call CreateNextDayMedOrder(OrderID)
        Else
            'This is only used by Med_Fill_History to track inpatient action by Order_Id
            Dim DataId As String
            DataId = goBB.I_DAS.GetNextID()
            frmDispense.ActionOrderID = Right$(CStr(goBB.Patient.UnitNumber), 6) + "-" + Right$(CStr(DataId), 5)
        End If
    End If
    
    'SCR 133400, jq, 6/21/07: CallFromGemsOE is always false and not needed; send msg
    'If CallFromGemsOE Then
        'Send message to Gems_OE to refresh the list
        goBB.Comm.Message cwiREFRESH_DATA, "", NCID_GEMS_OE_MED_APPLICATION, NCID_ORDER_MANAGER_INPATIENT_APPLICATION, False
        ' and meds
        goBB.Comm.Message cwiREFRESH_DATA, "", NCID_MEDICATIONS_APPLICATION, NCID_ORDER_MANAGER_INPATIENT_APPLICATION, False
    'End If
    
    Exit Sub
    
ErrHnd:
     goBB.SharedObj.ShowVBError Err.Number, Err.Description, "clsDispense.saveDispense", "clsDispense", vbExclamation
    Exit Sub
    Resume 'DEBUG
End Sub


Public Sub CreateNextDayMedOrder(ByVal OrderID As String)

    '*** 5/11/05 - FRANK
    '*** Create Next Day Inpatient Med Order w/ status Pending
    '*** When Dispensing with Create Next Day's Order

    Dim DataId As String
    Dim ExistingDataID As String
    
    Dim sParent_DataID As String
    
    Dim sSql As String
    
    Dim OraOrderDate As String
    
    Dim OE_Rx_Where As String
    Dim Medications_Where As String
    
    Dim NewOrderID As String
    Dim oRs As ADODB.Recordset

    On Error GoTo ErrHnd
    
    sSql = "Select * from OE_RX where ORDER_ID = '" & OrderID & "' AND (STATUS_NCID = " & NCID_ORDERSTATUS_ACTIVE
    sSql = sSql & " OR STATUS_NCID = " & NCID_ORDERSTATUS_COMPLETE & ")"
    sSql = sSql & " AND unit_number = " & goBB.Patient.UnitNumber
''    sSql = sSql & " AND ORDER_DATE = " & goBB.DataConn.sqldate(Date)
    Set oRs = goBB.I_DAS.OpenRecordset(sSql)
    
    If Not (oRs.BOF And oRs.EOF) Then

        DataId = goBB.I_DAS.GetNextID()
        NewOrderID = Right$(CStr(goBB.Patient.UnitNumber), 6) + "-" + Right$(CStr(DataId), 5)
        frmDispense.ActionOrderID = NewOrderID
        

        OE_Rx_Where = " WHERE OE_RX_DATA_ID = " & DataId & " and ORDER_ID = '" & NewOrderID & "'"
        Medications_Where = " WHERE DATAID = " & DataId & " and ORDERNUMBER = '" & NewOrderID & "'"

        '**************** OE_RX TABLE **************
        '** Add next day order
        
        OraOrderDate = Replace$(goBB.DataConn.SQLDate(DateAdd("d", 1, Date)), ":00:00:00", ":00:01:00")
        
        sSql = "INSERT INTO oe_rx (OE_RX_DATA_ID,ORDER_ID,UNIT_NUMBER,STATUS_NCID,ORDER_DATE,FILL_DATE) values("
        sSql = sSql & DataId & ",'" & NewOrderID & "'," & GetValueOrNull(oRs.Fields("UNIT_NUMBER") & vbNullString) & ","
        sSql = sSql & NCID_PARTIAL & "," & OraOrderDate & ",NULL)"
        
        goBB.I_DAS.ExecuteSQL sSql
            
        sSql = "UPDATE OE_RX "
        sSql = sSql & "SET ENC_NUM = " & GetValueOrNull(oRs.Fields("ENC_NUM") & vbNullString, True) & ","
        sSql = sSql & " ENC_FAC_NCID = " & GetValueOrNull(oRs.Fields("ENC_FAC_NCID") & vbNullString) & ","
        sSql = sSql & " RX_COMMENT = " & GetValueOrNull(oRs.Fields("RX_COMMENT") & vbNullString, True) & ","
        sSql = sSql & " SIG_CODE = " & GetValueOrNull(oRs.Fields("SIG_CODE") & vbNullString, True) & ","
        sSql = sSql & " QUANTITY = " & GetValueOrNull(oRs.Fields("QUANTITY") & vbNullString) & ","
        sSql = sSql & " REFILLS = 0,"
        sSql = sSql & " PHARMACY_NAME_SENT = " & GetValueOrNull(oRs.Fields("PHARMACY_NAME_SENT") & vbNullString, True) & ","
        sSql = sSql & " PHARMACY_IEN_SENT = " & GetValueOrNull(oRs.Fields("PHARMACY_IEN_SENT") & vbNullString, True) & ","
        sSql = sSql & " CHILD_CAP = " & GetValueOrNull(oRs.Fields("CHILD_CAP") & vbNullString) & ","
        sSql = sSql & " ORDER_IEN = " & GetValueOrNull(oRs.Fields("ORDER_IEN") & vbNullString) & ","
        sSql = sSql & " ORDER_NAME = " & GetValueOrNull(oRs.Fields("ORDER_NAME") & vbNullString, True) & ","
        'SF 9/6/05 SCR 76469
        'Time_Ordered is the Parent date time. Order_Date is the child date time.
        sSql = sSql & " TIME_ORDERED = " & goBB.DataConn.SQLDate(GetValueOrNull(oRs.Fields("TIME_ORDERED") & vbNullString)) & ","
        sSql = sSql & " CHILD_MED_IEN = " & GetValueOrNull(oRs.Fields("ORDER_IEN") & vbNullString)
        
        sSql = sSql & OE_Rx_Where
        
        Call goBB.I_DAS.ExecuteSQL(sSql)
        
        sSql = "UPDATE OE_RX "
        sSql = sSql & "SET SENSITIVE = " & GetValueOrNull(oRs.Fields("SENSITIVE") & vbNullString) & ","
        sSql = sSql & " ORDERED_BY = " & GetValueOrNull(oRs.Fields("ORDERED_BY") & vbNullString, True) & ","
        sSql = sSql & " PROVIDER = " & GetValueOrNull(oRs.Fields("PROVIDER") & vbNullString, True) & ","
        sSql = sSql & " UPDATEDBY = " & GetValueOrNull(oRs.Fields("UPDATEDBY") & vbNullString, True) & ","
        sSql = sSql & " UPDATEDON = " & goBB.DataConn.SQLDate(oRs.Fields("UPDATEDON") & vbNullString) & ","
        sSql = sSql & " CREATEDBY = " & GetValueOrNull(oRs.Fields("CREATEDBY") & vbNullString, True) & ","
        sSql = sSql & " CREATEDON = " & goBB.DataConn.SQLDate(oRs.Fields("CREATEDON") & vbNullString) & ","
        sSql = sSql & " SIGN_FLAG = " & GetValueOrNull(oRs.Fields("SIGN_FLAG") & vbNullString) & ","
        sSql = sSql & " NURSE_ACK_FALG = " & GetValueOrNull(oRs.Fields("NURSE_ACK_FALG") & vbNullString) & ","
        sSql = sSql & " PATIENT_TYPE = " & GetValueOrNull(oRs.Fields("PATIENT_TYPE") & vbNullString) & ","
        sSql = sSql & " PROCESS_PRIORITY = " & GetValueOrNull(oRs.Fields("PROCESS_PRIORITY") & vbNullString) & ","
        sSql = sSql & " NURSE_ACK_NAME = " & GetValueOrNull(oRs.Fields("NURSE_ACK_NAME") & vbNullString, True) & ","
        sSql = sSql & " NURSE_ACK_DATE = " & goBB.DataConn.SQLDate(oRs.Fields("NURSE_ACK_DATE") & vbNullString) & ","
        sSql = sSql & " PREVIOUS_ORDER_STATUS = " & GetValueOrNull(oRs.Fields("PREVIOUS_ORDER_STATUS") & vbNullString)

        sSql = sSql & OE_Rx_Where
        
        Call goBB.I_DAS.ExecuteSQL(sSql)
    
    
            
        '*** 2/14/06 - Frank
        '*** Note: Both OE_RX and Medications table use the same Data_ID
        sParent_DataID = oRs.Fields("PARENT_DATA_ID") & vbNullString        '** Not first child, same parent data_id
        If Len(sParent_DataID) = 0 Then
            sParent_DataID = oRs.Fields("OE_RX_DATA_ID") & vbNullString     '** First child, use data_id of parent order
        End If
        
        sSql = "UPDATE OE_RX "
        sSql = sSql & " SET PARENT_DATA_ID = " & Val(sParent_DataID)
        sSql = sSql & OE_Rx_Where
        Call goBB.I_DAS.ExecuteSQL(sSql)
    
    End If
    
    '**************** MEDICATIONS TABLE **************
    
    sSql = "Select * from MEDICATIONS where ORDERNUMBER = '" & OrderID & "' AND STATUSNCID = " & NCID_ORDERSTATUS_ACTIVE
    Set oRs = goBB.I_DAS.OpenRecordset(sSql)
    If Not (oRs.BOF And oRs.EOF) Then
    
        '** Add next day order
        sSql = "INSERT INTO medications (DATAID, UNIT_NUMBER,MEDICATIONNCID,STARTDATE,ORDERNUMBER,FILLDATE,STATUSNCID) "
        sSql = sSql & " values(" & DataId & "," & GetValueOrNull(oRs.Fields("UNIT_NUMBER") & vbNullString) & ","
        sSql = sSql & GetValueOrNull(oRs.Fields("MEDICATIONNCID") & vbNullString) & ", "
        sSql = sSql & OraOrderDate & ", '" & NewOrderID & "',NULL," & NCID_PARTIAL & ")"
        
        Call goBB.I_DAS.ExecuteSQL(sSql)
        
        sSql = "UPDATE MEDICATIONS "
        sSql = sSql & "SET SIG = " & GetValueOrNull(oRs.Fields("SIG") & vbNullString, True) & ","
        sSql = sSql & " QUANTITY = " & GetValueOrNull(oRs.Fields("QUANTITY") & vbNullString) & ","
        sSql = sSql & " OTC = " & GetValueOrNull(oRs.Fields("OTC") & vbNullString) & ","
        sSql = sSql & " REFILLS = " & GetValueOrNull(oRs.Fields("REFILLS") & vbNullString) & ","
        sSql = sSql & " REFILLSREMAINING = " & GetValueOrNull(oRs.Fields("REFILLSREMAINING") & vbNullString) & ","
        sSql = sSql & " DAYSOFSUPPLY = " & GetValueOrNull(oRs.Fields("DAYSOFSUPPLY") & vbNullString) & ","
        
        '** Next Day Order Date
        sSql = sSql & " ORDERDATE = " & OraOrderDate & ","
        
        sSql = sSql & " ORDEREXPIRATIONDATE = " & goBB.DataConn.SQLDate(oRs.Fields("ORDEREXPIRATIONDATE") & vbNullString) & ","
        
        '*** Same as STARTDATE
        sSql = sSql & " ENDDATE = " & OraOrderDate & ","
        
        sSql = sSql & " ORDERINGPROVIDERNCID = " & GetValueOrNull(oRs.Fields("ORDERINGPROVIDERNCID") & vbNullString) & ","
        sSql = sSql & " ORDERINGPROVIDER = " & GetValueOrNull(oRs.Fields("ORDERINGPROVIDER") & vbNullString, True) & ","
        sSql = sSql & " CLINICNCID = " & GetValueOrNull(oRs.Fields("CLINICNCID") & vbNullString) & ","
        sSql = sSql & " PATIENTTYPE = " & GetValueOrNull(oRs.Fields("PATIENTTYPE") & vbNullString, True) & ","
        sSql = sSql & " DISPENSINGLOCATIONNCID = " & GetValueOrNull(oRs.Fields("DISPENSINGLOCATIONNCID") & vbNullString) & ","
        sSql = sSql & " NATIONALDRUGCODE = " & GetValueOrNull(oRs.Fields("NATIONALDRUGCODE") & vbNullString, True) & ","
        sSql = sSql & " MEDICATIONSCOMMENT = " & GetValueOrNull(oRs.Fields("MEDICATIONSCOMMENT") & vbNullString, True) & ","
        sSql = sSql & " CDR_DATAID = " & GetValueOrNull(oRs.Fields("CDR_DATAID") & vbNullString) & ","
        sSql = sSql & " STORINGAPP = " & GetValueOrNull(oRs.Fields("STORINGAPP") & vbNullString)


        sSql = sSql & Medications_Where
        Call goBB.I_DAS.ExecuteSQL(sSql)
     
        sSql = "UPDATE MEDICATIONS "
        sSql = sSql & "SET UPDATEFLAG = " & GetValueOrNull(oRs.Fields("UPDATEFLAG") & vbNullString, True) & ","
        sSql = sSql & " CREATEDBY = " & GetValueOrNull(oRs.Fields("CREATEDBY") & vbNullString)
        
        sSql = sSql & Medications_Where
        Call goBB.I_DAS.ExecuteSQL(sSql)
        
        sSql = "UPDATE MEDICATIONS "
        sSql = sSql & "SET CREATEDON = " & goBB.DataConn.SQLDate(oRs.Fields("CREATEDON") & vbNullString) & ","
        sSql = sSql & " NON_CODED_MEDICATION_NAME = " & GetValueOrNull(oRs.Fields("NON_CODED_MEDICATION_NAME") & vbNullString, True)
        
        sSql = sSql & Medications_Where
        Call goBB.I_DAS.ExecuteSQL(sSql)
        
        sSql = "UPDATE MEDICATIONS "
        sSql = sSql & "SET Tmip_Status = " & GetValueOrNull(oRs.Fields("Tmip_Status") & vbNullString) & ","
        sSql = sSql & " ENC_NUM = " & GetValueOrNull(oRs.Fields("ENC_NUM") & vbNullString, True) & ","
        sSql = sSql & " FDB_MEDID = " & GetValueOrNull(oRs.Fields("FDB_MEDID") & vbNullString) & ","
        sSql = sSql & " OVERRIDE_REASON = " & GetValueOrNull(oRs.Fields("OVERRIDE_REASON") & vbNullString) & ","
        sSql = sSql & " OVERRIDE_COMMENT = " & GetValueOrNull(oRs.Fields("OVERRIDE_COMMENT") & vbNullString, True) & ","
        sSql = sSql & " WARNINGS = " & GetValueOrNull(oRs.Fields("WARNINGS") & vbNullString, True)
        
        sSql = sSql & Medications_Where
            
        Call goBB.I_DAS.ExecuteSQL(sSql)
    
    
        '** 2/14/06 - Frank
        sSql = "UPDATE MEDICATIONS "
        sSql = sSql & " SET PARENT_DATA_ID = " & Val(sParent_DataID)
        sSql = sSql & Medications_Where
        Call goBB.I_DAS.ExecuteSQL(sSql)
    
    End If
    Exit Sub
    
ErrHnd:
     goBB.SharedObj.ShowVBError Err.Number, Err.Description, "clsDispense.CreateNextDayMedOrder", "clsDispense", vbExclamation
    Exit Sub
    Resume '* DEBUG
End Sub

Public Sub UpdatePrevActiveToComplete(ByVal OrderID As String)

    '*** 5/13/05 - FRANK
    '*** Update PrevMedOrderStat ACTIVE to COMPLETED

    Dim sSql As String
    Dim oRs As ADODB.Recordset
    Dim ActiveOrderID As String
    Dim OrderIEN As String

    On Error GoTo ErrHnd
    
    sSql = "SELECT ORDER_IEN, UNIT_NUMBER FROM OE_RX WHERE ORDER_ID = '" & OrderID & "'"
    Set oRs = goBB.I_DAS.OpenRecordset(sSql)
    If Not (oRs.BOF And oRs.EOF) Then
        '*** Note Multiple encounter may have the same ODER_IEN
        '*** The one we want is part of ORDER_ID contains UNIT_NUMBER
        oRs.MoveFirst
        While Not oRs.EOF
            If oRs.Fields("UNIT_NUMBER") & vbNullString = goBB.Patient.UnitNumber Then
                OrderIEN = oRs.Fields("ORDER_IEN")
                oRs.MoveLast
            End If
            oRs.MoveNext
        Wend
    End If

    sSql = "SELECT ORDER_ID FROM OE_RX "
    sSql = sSql & " WHERE ORDER_IEN = " & OrderIEN
    sSql = sSql & " AND ORDER_ID <> '" & OrderID & "'"
    sSql = sSql & " AND STATUS_NCID = " & NCID_ORDERSTATUS_ACTIVE
    sSql = sSql & " AND UNIT_NUMBER = " & goBB.Patient.UnitNumber
    sSql = sSql & " AND PATIENT_TYPE = 1"
    
    ActiveOrderID = ""
    Set oRs = goBB.I_DAS.OpenRecordset(sSql)
    If Not (oRs.BOF And oRs.EOF) Then
        ActiveOrderID = oRs.Fields("ORDER_ID") & vbNullString
    End If
    
    If Len(ActiveOrderID) > 0 Then
        sSql = "UPDATE OE_RX "
        sSql = sSql & " set STATUS_NCID = " & NCID_ORDERSTATUS_COMPLETE
        sSql = sSql & " WHERE ORDER_ID = '" & ActiveOrderID & "'"
        sSql = sSql & " AND STATUS_NCID = " & NCID_ORDERSTATUS_ACTIVE
        
        Call goBB.I_DAS.ExecuteSQL(sSql)
        
        sSql = "UPDATE MEDICATIONS "
        sSql = sSql & " set STATUSNCID = " & NCID_ORDERSTATUS_COMPLETE
        sSql = sSql & " WHERE ORDERNUMBER = '" & ActiveOrderID & "'"
        sSql = sSql & " and STATUSNCID = " & NCID_ORDERSTATUS_ACTIVE
        
        Call goBB.I_DAS.ExecuteSQL(sSql)
    End If
    
    Exit Sub
    
ErrHnd:
     goBB.SharedObj.ShowVBError Err.Number, Err.Description, "clsDispense.UpdatePrevActiveToComplete", "clsDispense", vbExclamation
    Exit Sub
    Resume '* DEBUG
End Sub

Private Function GetValueOrNull(sValue As String, Optional IsString As Boolean = False) As String

    On Error Resume Next

    If IsString Then
        GetValueOrNull = IIf(sValue = "", "NULL", SQLQuote(sValue))
    Else
        GetValueOrNull = IIf(sValue = "", "NULL", sValue)
    End If
    
End Function

Public Sub saveRefills(ByVal OrderID As String, Optional CallFromGemsOE As Boolean = False)

Dim rs     As Recordset
Dim strSQL As String
Dim DataId As String
Dim FacilityID As String

On Error GoTo ErrHnd

    FacilityID = goBB.Logon.FacilityNCID
    'SCR 46333, is refill, create new pending order by using the exist active order's order_id
    DataId = goBB.I_DAS.GetNextID()

    strSQL = "SELECT * FROM OE_RX WHERE ORDER_ID = '" & OrderID & "'"
    Set rs = goBB.I_DAS.OpenRecordset(strSQL)

    Dim sEnc As String
    
    

    If Not rs.EOF Then
    
        sEnc = rs!ENC_NUM & vbNullString
        
        ' SCR 79695, Quan
        Dim sSig As String
        sSig = frmDispense.UpdateSigWhenRefillsChanged        ' rs!SIG_CODE,
        
        'SCR 50879/50992 - JShen 3/9/2004
        '** SCR 72193 - Frank 6/13/05; Added PROCESS_PRIORITY
        'SCR 135000 changed refills to refillRem
        strSQL = "INSERT INTO oe_rx (oe_rx_data_id, enc_fac_ncid, order_date, " & _
            "unit_number, rx_comment, sig_code, quantity,refills, " & _
            "order_id, status_ncid, order_ien, " & _
            "order_name, time_ordered, ordered_by, provider, " & _
            "PATIENT_TYPE, ENC_NUM, PROCESS_PRIORITY) " & _
            "VALUES (" & _
            DataId & "," & _
            FacilityID & "," & _
            goBB.DataConn.FormatDateForDBInsert("" & rs!ORDER_DATE) & "," & _
            "" & rs!Unit_Number & "," & _
            goBB.DataConn.SQLQuote(Comment, 255) & "," & _
            goBB.DataConn.SQLQuote("" & sSig, 100) & "," & _
            Val(Quantity) & "," & _
            Val(RefillRem) & ",'" & _
            "" & rs!order_id & "'," & _
            NCID_PARTIAL & "," & _
            "" & rs!ORDER_IEN & ", " & _
            goBB.DataConn.SQLQuote("" & rs!ORDER_NAME, 75) & "," & _
            goBB.DataConn.FormatDateForDBInsert(mFillDate) & "," & _
            goBB.DataConn.SQLQuote("" & rs!ordered_by, 100) & ",'" & _
            "" & rs!Provider & "'," & _
            "" & rs!PATIENT_TYPE & "," & _
             SQLQuote(sEnc) & "," & _
            Val("" & rs!PROCESS_PRIORITY) & ")"

        goBB.I_DAS.ExecuteSQL strSQL
    End If
    
    'SCR 50352 - Jane Shen 2/24/2004
    strSQL = "update Medications set "
    'SCR 135000
    strSQL = strSQL & " Refills = " & Val(Refills)
    strSQL = strSQL & ", REFILLSREMAINING = " & Val(msRefillRem)
    strSQL = strSQL & ", Quantity = " & Val(Quantity)
    strSQL = strSQL & ", Sig = " & goBB.DataConn.SQLQuote(Sig, 50)
    strSQL = strSQL & " where ORDERNUMBER = '" & OrderID & "'"
    goBB.I_DAS.ExecuteSQL strSQL
    
    'SCR 133400, jq, 6/21/07: CallFromGemsOE is always false and not needed; send msg
    'If CallFromGemsOE = True Then
        'Send message to Gems_OE to refresh the list
        goBB.Comm.Message cwiREFRESH_DATA, "", NCID_GEMS_OE_MED_APPLICATION, NCID_ORDER_MANAGER_INPATIENT_APPLICATION, False
        ' and meds
        goBB.Comm.Message cwiREFRESH_DATA, "", NCID_MEDICATIONS_APPLICATION, NCID_ORDER_MANAGER_INPATIENT_APPLICATION, False
    'End If
    
    Exit Sub
    
ErrHnd:
     goBB.SharedObj.ShowVBError Err.Number, Err.Description, "clsDispense.saveRefills", "clsDispense", vbExclamation

End Sub

Public Function retrieveData(OrderID As String, Optional isPending As Boolean, _
                        Optional DispenseFromPrintLabel As Boolean = False) As Boolean

    Dim rs      As Recordset
    Dim strSQL  As String
    
    Dim ReasonID As String
    
    On Error GoTo ErrHnd

    '************************ NOTE **************************
    'PRE-FDB implementation
    'select order_ien from oe_rx
    'where order_ien in (select medicationncid from  medications)
    '
    'POST-FDB implementation [INPATIENT2]
    'select order_ien from oe_rx
    'where order_ien in (select fdb_medid from medications)

    '** 1/10/05 - Frank, added: Patient_Type, URSE_ACK_FALG, STATUS_NCID
    strSQL = "select ORDER_NAME, order_ien, Patient_Type, STATUS_NCID, NURSE_ACK_FALG, RX_COMMENT, SIG_CODE, Refills, QUANTITY, PHARMACY_NAME_SENT,FILL_DATE  "
    strSQL = strSQL & " from OE_RX "
    strSQL = strSQL & " where  ORDER_ID = '" & OrderID & "'"
    
    If DispenseFromPrintLabel Then
        '*** Don't care Med Order status.
        '*** User can Dispense when Med Order is: Pending, Active, Complete....
    Else
        If isPending Then 'SCR 49671 - Jane Shen 2/20/2004
            strSQL = strSQL & " and STATUS_NCID = " & NCID_PARTIAL
        Else
            '** Refill
            strSQL = strSQL & " and STATUS_NCID = " & NCID_ORDERSTATUS_ACTIVE
        End If
    End If
    
    Set rs = goBB.I_DAS.OpenRecordset(strSQL)

    If Not rs.EOF Then
        Comment = rs!RX_COMMENT & vbNullString
        Quantity = rs!Quantity
        Name = rs!ORDER_NAME & vbNullString
        Refills = rs!Refills & vbNullString
        Sig = rs!SIG_CODE & vbNullString
        ExpSig = rs!PHARMACY_NAME_SENT & vbNullString
        MedID = rs!ORDER_IEN
        IsInpatient = rs!PATIENT_TYPE = 1
        'SCR 68684 - Frank 4/13/05
        '** Fixed Type Mismatch: NurseAck = "0" & rs!NURSE_ACK_FALG
        NurseAck = Val("" & rs!NURSE_ACK_FALG)
        OrderStaus = "" & rs!STATUS_NCID
        If IsNull(rs!FILL_DATE) Then
            firstFill = True
        Else
            firstFill = False
        End If
        rs.Close
    End If
    
    'SCR 135000 add refills to query
    If Len(goBB.SharedObj.CmdLineSwitches("INPATIENT2")) > 0 Then
        'FRANK - 1/05/04
        
       '85433  -  Medications can be refilled after the stop date
        'Frank 1/10/06 -- Added EdnDate
        strSQL = "select a.REFILLS, a.REFILLSREMAINING, a.Override_Reason, a.Override_Comment, " _
               & "a.FDB_MEDID, a.ENDDATE from MEDICATIONS a " _
               & "where a.ORDERNUMBER = '" & OrderID & "' "
    Else
        strSQL = "select REFILLS, REFILLSREMAINING from MEDICATIONS where ORDERNUMBER = '" & OrderID & "'"
    End If

    
    Set rs = goBB.I_DAS.OpenRecordset(strSQL)
    If Not rs.EOF Then
        'SCR 135000
        Refills = rs!Refills & vbNullString
        RefillRem = rs!RefillsRemaining & vbNullString
        If Len(goBB.SharedObj.CmdLineSwitches("INPATIENT2")) > 0 Then
            OverideComment = rs!OVERRIDE_COMMENT & vbNullString
            ReasonID = rs!OVERRIDE_REASON & vbNullString
            EndDate = rs!EndDate & vbNullString
        End If
    
        rs.Close
    End If
    
    'FRANK - 1/05/04
    If ReasonID <> "" Then
        If Len(goBB.SharedObj.CmdLineSwitches("INPATIENT2")) > 0 Then
            strSQL = "select a.reason_description " _
                       & "from Override_Reason a " _
                       & "where a.id = " & ReasonID
    
            Set rs = goBB.I_DAS.OpenRecordset(strSQL)
            If Not rs.EOF Then
                OverideReason = rs!reason_description & vbNullString
            End If
        End If
        rs.Close
    End If
    
    retrieveData = True
    Exit Function
    
ErrHnd:
    goBB.SharedObj.ShowVBError Err.Number, Err.Description, "clsDispense.retrieveData", "clsDispense", vbExclamation
    retrieveData = False
    Exit Function
    Resume
    
End Function



Public Function GetMedsCollection() As Collection
    'get the active meds for the current patient
    Dim medsCollection As Collection
    Dim objMeds As clsDispense
    Dim strSQL As String
    Dim rs     As Recordset
    
On Error GoTo ErrHandler

    Set medsCollection = New Collection
    
' Begin:SCR #49472;   Developer: Jane Shen 02/11/2004 04:10 PM
' Note: Function Changed:   GetMedsCollection()    Impact of Change:  add condition "o.Unit_number = m.unit_numbe"      Modules Changed:  OrderManager.dll\clsDispense.cls
' End: SCR #49472;
    strSQL = "select o.ORDER_NAME, o.SIG_CODE, m.REFILLSREMAINING, o.QUANTITY, "
    strSQL = strSQL & " m.EndDate, o.ORDER_ID, o.RX_COMMENT, o.STATUS_NCID,  o.PATIENT_TYPE"
    strSQL = strSQL & " from oe_rx o, medications m"
    strSQL = strSQL & " where o.Unit_number = '" & goBB.Patient.UnitNumber & "'"
    strSQL = strSQL & " and o.STATUS_NCID = " & NCID_ORDERSTATUS_ACTIVE
    strSQL = strSQL & " and o.Unit_number = m.unit_number"
'    strSQL = strSQL & " and o.order_ien = m.MedicationNCID"
    strSQL = strSQL & " and o.order_id = m.OrderNumber "

    Set rs = goBB.I_DAS.OpenRecordset(strSQL)
    Do While Not rs.EOF
        Set objMeds = New clsDispense
        
        With objMeds
            If rs!STATUS_NCID = NCID_ORDERSTATUS_ACTIVE Then
                .Quantity = rs!Quantity
                .Name = rs!ORDER_NAME
                .RefillRem = rs!RefillsRemaining
                .Sig = rs!SIG_CODE & vbNullString
                .OrderID = rs!order_id
                .Comment = rs!RX_COMMENT & vbNullString
                .OrderStaus = rs!STATUS_NCID & vbNullString
                .IsInpatient = Val(rs!PATIENT_TYPE & vbNullString) = 1
                
                .EndDate = rs!EndDate & vbNullString    '85433
                
                If Not Val(objMeds.RefillRem) = 0 Then
                    medsCollection.Add objMeds
                End If
            End If
        End With
        Set objMeds = Nothing
        rs.MoveNext
        
        Loop
    
    Set GetMedsCollection = medsCollection
    rs.Close
    
    Exit Function
    
ErrHandler:
    goBB.SharedObj.ShowVBError Err.Number, Err.Description, "clsDispense.GetMedsCollection", "OrderMgt", vbExclamation
    
End Function



