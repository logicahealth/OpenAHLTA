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
Attribute VB_Name = "UnsignedOrders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"UnsignedOrder"
Attribute VB_Ext_KEY = "Member0" ,"UnsignedOrder"
Option Explicit
Private mCol As Collection
Private mvarErrorIsMessage As Boolean 'local copy
Private mvarWarningIndex As Integer
Private mvarbSendingOverrides As Boolean



'------------------------------------------------------------------------------
Public Function AddUsingHL7Segments() As Boolean
    Dim oDAL As ICHCSII_DAS
    Dim oRS  As Recordset
    Set oDAL = pobjConn.CHCSII_DAS '<SCR 37864

    Dim sQuerySelectLabs As String
    Dim sQueryWhereLabs As String
    Dim sQuerySelectRads As String
    Dim sQueryWhereRads As String
    
    sQuerySelectLabs = _
      "Select (n.last_name + ',' + n.first_name ) as Name, " & vbCrLf _
    & " lab.ORDER_NAME, lab.status, lab.ORDER_ID," & vbCrLf _
    & " lab.TIME_ORDERED, lab.LAB_COMMENT as COMMENT," & vbCrLf _
    & " id.ssn, id.BIRTH_DATE  " & vbCrLf _
    
    sQueryWhereLabs = _
      " From  oe_lab lab,  mmi_id id, mmi_name n " & vbCrLf _
    & " Where lab.unit_number = id.unit_number " & vbCrLf _
    & " and   lab.unit_number = n.unit_number " & vbCrLf _
     & " and  lab.sign = 1 " & vbCrLf _
    
    Set oRS = oDAL.OpenRecordset(sQuerySelectLabs & sQueryWhereLabs)
    Call AddUnsignedOrderCollection(oRS, hlLAB_ORDERS_CODE)
    
    sQuerySelectRads = _
      "Select (n.last_name + ',' + n.first_name ) as Name, " & vbCrLf _
    & " rad.ORDER_NAME, rad.status, rad.ORDER_ID," & vbCrLf _
    & " rad.TIME_ORDERED,rad.RAD_COMMENT as COMMENT," & vbCrLf _
    & " id.ssn, id.BIRTH_DATE " & vbCrLf _
    
    sQueryWhereRads = _
      " From  oe_rad rad,  mmi_id id, mmi_name n" & vbCrLf _
    & " Where rad.unit_number = id.unit_number " & vbCrLf _
    & " and   rad.unit_number = n.unit_number " & vbCrLf _
     & " and  rad.sign = 1 " & vbCrLf _
    
    Set oRS = oDAL.OpenRecordset(sQuerySelectRads & sQueryWhereRads)
    Call AddUnsignedOrderCollection(oRS, hlRAD_ORDERS_CODE)
    
    ' Begin:SCR #881;   Developer: Jane Shen 10/24/2002 04:22 PM
    sQuerySelectRads = _
      "Select distinct (n.last_name + ',' + n.first_name ) as Name, " & vbCrLf _
    & " rx.ORDER_NAME, rx.sig_code, med.STATUSNCID, rx.ORDER_ID," & vbCrLf _
    & " rx.Order_Date as TIME_ORDERED,rx.RX_COMMENT as COMMENT," & vbCrLf _
    & " id.ssn, id.BIRTH_DATE " & vbCrLf _
    ' End: SCR #881;
    
    sQueryWhereRads = _
      " From  OE_RX rx, MEDICATIONS med, mmi_id id, mmi_name n" & vbCrLf _
    & " Where rx.unit_number = id.unit_number " & vbCrLf _
    & " and   rx.unit_number = n.unit_number " & vbCrLf _
     & " and   rx.unit_number = med.unit_number and  rx.sign = 1 " & vbCrLf _
    
    Set oRS = oDAL.OpenRecordset(sQuerySelectRads & sQueryWhereRads)
    Call AddUnsignedOrderCollection(oRS, hlRX_ORDERS_CODE)

    oRS.Close
    
End Function

Public Property Get Item(vntIndexKey As Variant) As UnsignedOrder
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

    Dim unsignOrder As GEMS_OE.UnsignedOrder
    
    Set unsignOrder = mCol(vntIndexKey)
    If unsignOrder.UserHasSigned = True Then
        Call goOESession.PerformAction(unsignOrder.OrderType, hlACTION_EXPAND_SIG, unsignOrder.OrderNumber)
    End If
    
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
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub
Public Sub AddFromSummaryCollection(pSum As SummaryType, bClear As Boolean)
'    Dim objNewMember As UnsignedOrder
'    Dim i As Integer
'
'    If bClear = True Then
'        For i = 1 To mCol.Count
'            mCol.Remove 1
'        Next
'    End If
'
'    Set objNewMember = New UnsignedOrder
'    objNewMember.OrderNumber = pSum.OrderId
'    objNewMember.SigningAction = pSum.SigningAction
'    objNewMember.szOrderEventDateHL7Format = pSum.szOrderEventDateHL7Format
'    objNewMember.OrderType = pSum.OrderTypeCode
'    objNewMember.szStatus = pSum.Status
'    If goHl7Constructor.SignatureClass > 1 Then
'        objNewMember.ProviderSignatureInfo.DateTimeHL7Format = Format$(Now, DATE_TIME_ORC7_FORMAT)
'        objNewMember.ProviderSignatureInfo.IEN = gszProviderIEN
'    Else
'        objNewMember.NurseSignatureInfo.DateTimeHL7Format = Format$(Now, DATE_TIME_ORC7_FORMAT)
'        objNewMember.NurseSignatureInfo.IEN = gszProviderIEN
'    End If
'    objNewMember.UserHasSigned = True
'    objNewMember.SigningSuccessful = True
'
'    mCol.Add objNewMember
'    Set objNewMember = Nothing
    
End Sub
Public Sub AddIndividualOrder(OrderId As String, OrderDate As String, OrderType As hlORDERTYPECODE, Optional OrderName As String)
    Dim objNewMember As UnsignedOrder
    Dim i As Integer

    
    'Clear the collection
    For i = 1 To mCol.Count
        mCol.Remove 1
    Next
    
    Set objNewMember = New UnsignedOrder
    objNewMember.OrderNumber = OrderId
    objNewMember.SigningAction = hlNEW
    objNewMember.szOrderEventDateHL7Format = OrderDate
    objNewMember.OrderType = OrderType
    objNewMember.szDisplayText = OrderName
'    'objNewMember.szStatus = pSum.Status
'    If goHl7Constructor.SignatureClass > 1 Then
'        objNewMember.ProviderSignatureInfo.DateTimeHL7Format = Format$(Now, DATE_TIME_ORC7_FORMAT)
'        objNewMember.ProviderSignatureInfo.IEN = gszProviderIEN
'    Else
'        objNewMember.NurseSignatureInfo.DateTimeHL7Format = Format$(Now, DATE_TIME_ORC7_FORMAT)
'        objNewMember.NurseSignatureInfo.IEN = gszProviderIEN
'    End If
    objNewMember.UserHasSigned = True
    objNewMember.SigningSuccessful = True
    
    
    mCol.Add objNewMember
    Set objNewMember = Nothing
    

End Sub


Public Property Let ErrorIsMessage(ByVal vData As Boolean)
    mvarErrorIsMessage = vData
End Property

Public Property Get ErrorIsMessage() As Boolean
    ErrorIsMessage = mvarErrorIsMessage
End Property
Public Property Let WarningIndex(ByVal vData As Integer)
    mvarWarningIndex = vData
End Property

Public Property Get WarningIndex() As Integer
    WarningIndex = mvarWarningIndex
End Property
Public Property Let bSendingOverrides(ByVal vData As Boolean)
    mvarbSendingOverrides = vData
End Property

Public Property Get bSendingOverrides() As Boolean
    bSendingOverrides = mvarbSendingOverrides
End Property


Private Sub AddUnsignedOrderCollection(oRS As Recordset, OrderType As hlORDERTYPECODE)
    Dim szNewStr As String
    Dim sigCode As String
    Dim objNewMember As UnsignedOrder
    
    Do While Not oRS.EOF
        Set objNewMember = New UnsignedOrder
        With objNewMember
            szNewStr = Format$(oRS("TIME_ORDERED"), DATE_FORMAT_VB_DD_MON_YYYY_HH_MM_SS)
            
            .OrderNumber = oRS("ORDER_ID")
            If OrderType = hlRX_ORDERS_CODE Then
                .szStatus = "PENDING"
                sigCode = oRS("SIG_CODE")
                .szDisplayText = oRS("ORDER_NAME") & "  " & sigCode & " on " & szNewStr
            Else
                .szDisplayText = oRS("ORDER_NAME") & " on " & szNewStr
            If Not IsNull(oRS("STATUS")) Then .szStatus = OrderSatusForDisplay(oRS("STATUS"))
            End If
            If Not IsNull(oRS("COMMENT")) Then .OrderComment = oRS("COMMENT")
            If Not IsNull(oRS("TIME_ORDERED")) Then .OrderStartDateTime = oRS("TIME_ORDERED")
            Set .PatientInfo = New PatientType
            szNewStr = oRS("SSN")
            szNewStr = Left$(szNewStr, 3) & "-" & Mid$(szNewStr, 4, 2) & "-" & Right$(szNewStr, 4)
            
            .PatientInfo.Name = oRS("Name")
            .PatientInfo.DOB = Format$(oRS("BIRTH_DATE"), ORC7_FORMAT)
            .PatientInfo.Fmp = "20/" & szNewStr
            
            .szOrderingProviderIEN = ""
            .szOrderingProviderName = "NPOE, DOCTOR Tester"
            .szEnteredByName = "NPOE, NURSE TEN"
            .OrderType = OrderType
            .ProviderSignatureNeeded = hlSigNeeded
            .CompletionStatus = hlBLANK
            mCol.Add objNewMember
        End With
        oRS.MoveNext
    Loop

End Sub
    
