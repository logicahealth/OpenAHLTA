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



'Public Function Add(OrderNumber As String, OrderType As HL7Orders.hlORDERTYPECODE, bModifiable As Boolean, szDisplayText As String, szStatus As String, szExpirationDate As String, szDispenseStatus As String, szRefillsRemaining As String, szLastFillDate As String, szOnHoldUntilDate As String, szEnteredByIEN As String, szEnteredByName As String, szOrderingProviderIEN As String, szOrderingProviderName As String, SigningAction As hlORDERSIGNACTION, szOrderEventDate As String, NurseSignatureNeeded As hlSIGNATURENEEDED, ProviderSignatureNeeded As hlSIGNATURENEEDED, CounterSignatureNeeded As hlSIGNATURENEEDED, AncillarySignatureNeeded As hlSIGNATURENEEDED, AncillarySignatureInfo As SignatureInfo, CounterSignatureInfo As SignatureInfo, NurseSignatureInfo As SignatureInfo, ProviderSignatureInfo As SignatureInfo, PatientInfo As PatientType, UnsignedOrders As UnsignedOrders, Optional sKey As String) As UnsignedOrder
'    'create a new object
'    Dim objNewMember As UnsignedOrder
'    Set objNewMember = New UnsignedOrder
'
'
'    'set the properties passed into the method
'    objNewMember.OrderNumber = OrderNumber
'
'    objNewMember.OrderType = OrderType
'
'    objNewMember.Modifiable = Modifiable
'    objNewMember.szDisplayText = szDisplayText
'    objNewMember.szStatus = szStatus
'    objNewMember.szExpirationDate = szExpirationDate
'    objNewMember.szDispenseStatus = szDispenseStatus
'    objNewMember.szRefillsRemaining = szRefillsRemaining
'    objNewMember.szLastFillDate = szLastFillDate
'    objNewMember.szOnHoldUntilDate = szOnHoldUntilDate
'    objNewMember.szEnteredByIEN = szEnteredByIEN
'    objNewMember.szEnteredByName = szEnteredByName
'    objNewMember.szOrderingProviderIEN = szOrderingProviderIEN
'    objNewMember.szOrderingProviderName = szOrderingProviderName
'
'    objNewMember.SigningAction = SigningAction
'
'    objNewMember.szOrderEventDate = szOrderEventDate
'
'    objNewMember.NurseSignatureNeeded = NurseSignatureNeeded
'
'    objNewMember.ProviderSignatureNeeded = ProviderSignatureNeeded
'
'    objNewMember.CounterSignatureNeeded = CounterSignatureNeeded
'
'    objNewMember.AncillarySignatureNeeded = AncillarySignatureNeeded
'
'    If IsObject(AncillarySignatureInfo) Then
'        Set objNewMember.AncillarySignatureInfo = AncillarySignatureInfo
'    Else
'        objNewMember.AncillarySignatureInfo = AncillarySignatureInfo
'    End If
'    If IsObject(CounterSignatureInfo) Then
'        Set objNewMember.CounterSignatureInfo = CounterSignatureInfo
'    Else
'        objNewMember.CounterSignatureInfo = CounterSignatureInfo
'    End If
'    If IsObject(NurseSignatureInfo) Then
'        Set objNewMember.NurseSignatureInfo = NurseSignatureInfo
'    Else
'        objNewMember.NurseSignatureInfo = NurseSignatureInfo
'    End If
'    If IsObject(ProviderSignatureInfo) Then
'        Set objNewMember.ProviderSignatureInfo = ProviderSignatureInfo
'    Else
'        objNewMember.ProviderSignatureInfo = ProviderSignatureInfo
'    End If
'    If IsObject(PatientInfo) Then
'        Set objNewMember.PatientInfo = PatientInfo
'    Else
'        objNewMember.PatientInfo = PatientInfo
'    End If
'    Set objNewMember.UnsignedOrders = UnsignedOrders
'    If Len(sKey) = 0 Then
'        mCol.Add objNewMember
'    Else
'        mCol.Add objNewMember, sKey
'    End If
'
'
'    'return the object created
'    Set Add = objNewMember
'    Set objNewMember = Nothing
'
'
'End Function
'------------------------------------------------------------------------------
Public Function AddUsingHL7Segments() As Boolean
    Dim szNewStr As String
    Dim i As Long
    Dim iStop As Long
    Dim szTemp As String
    Dim szTemp2 As String
    Dim iStart As Long
    
    
    Dim objNewMember As UnsignedOrder
    Dim objNewZOH As UnsignedAddAction
    
    i = InStr(gszHL7RetMsg, "ZOF")
    If i Then
      szNewStr = Mid$(gszHL7RetMsg, i, Len(gszHL7RetMsg))
      i = InStr(szNewStr, "ZOF")
    Else
      '** no ZOD info ...*** ERROR ***
    End If
    
 
    
    While i
        iStart = i
        Set objNewMember = New UnsignedOrder
        iStop = InStr(iStart + 1, szNewStr, "ZOF")
        objNewMember.OrderNumber = szGetColumn(szNewStr, hlZOFSeg.OrderId + 1, "^")
        objNewMember.OrderType = szGetColumn(szNewStr, hlZOFSeg.OrderType + 1, "^")
        
        szTemp = szGetColumn(szNewStr, hlZOFSeg.OrderDisplayText + 1, "^")
        ParseDisplayText szTemp, szTemp2
        objNewMember.szDisplayText = szTemp
        
        objNewMember.szStatus = szGetColumn(szNewStr, hlZOFSeg.OrderStatus + 1, "^")
        
        'SCR 31655  NPOE - Sign Orders module, canceled LAB & MED orders not marked
        'If objNewMember.OrderType <> hlRAD_ORDERS_CODE Then
            If Not (objNewMember.szStatus = "PENDING" Or objNewMember.szStatus = "ACTIVE") And InStr(objNewMember.szDisplayText, objNewMember.szStatus) = 0 Then
                objNewMember.szDisplayText = objNewMember.szDisplayText & " [" & objNewMember.szStatus & "]"
            End If
        'End If
        If Trim$(szGetColumn(szNewStr, hlZOFSeg.OrderModifiable + 1, "^")) = "1" Then
            objNewMember.Modifiable = hlIS_MODIFIABLE
        Else
            objNewMember.Modifiable = hlIS_NOT_MODIFIABLE
        End If
        objNewMember.szExpirationDate = szGetColumn(szNewStr, hlZOFSeg.ExpirationDate + 1, "^")
        objNewMember.szExpirationDate = FormatHL7DateToMilitary(Left$(objNewMember.szExpirationDate, 8))
        objNewMember.szDispenseStatus = szGetColumn(szNewStr, hlZOFSeg.LastDispenseDate + 1, "^")
  
        If Len(objNewMember.szDispenseStatus) > 1 Then
            objNewMember.szDispenseStatus = Right$((Left$(objNewMember.szDispenseStatus, Len(objNewMember.szDispenseStatus) - 1)), Len(objNewMember.szDispenseStatus) - 2)   'gets rid of ()
        End If
        If InStr(objNewMember.szDispenseStatus, "Expand") > 0 Then
            objNewMember.szDispenseStatus = Left$(objNewMember.szDispenseStatus, 17)
        End If
        
        szTemp = szGetColumn(szNewStr, hlZOFSeg.RefillsRemaining + 1, "^")
        If (Len(szTemp) = 0) Or Not (IsNumeric(szTemp)) Then     'Just in case blank - although shouldn't be
            objNewMember.szRefillsRemaining = "" 'means we did not get a Refills remaining sequence
            objNewMember.szLastFillDate = "" 'means we did not get a Last Fill date sequence
        Else
            objNewMember.szRefillsRemaining = szTemp
            objNewMember.szLastFillDate = FormatHL7DateToMilitary(szGetColumn(szNewStr, hlZOFSeg.LastFillDate + 1, "^"))
        End If
    
        objNewMember.szOnHoldUntilDate = FormatHL7DateToMilitary(szGetColumn(szNewStr, hlZOFSeg.OnHoldUntil + 1, "^"))
        
        szTemp = szGetColumn(szNewStr, hlZOFSeg.EnteredBy + 1, "^")
        objNewMember.szEnteredByIEN = szGetColumn(szTemp, 1, "\")
        objNewMember.szEnteredByName = Trim$(szGetColumn(szTemp, 2, "\") & ", " & szGetColumn(szTemp, 3, "\") & " " & szGetColumn(szTemp, 4, "\"))
        
        szTemp = szGetColumn(szNewStr, hlZOFSeg.OrderedBy + 1, "^")
        If LenB(szTemp) <> 0 Then
            objNewMember.szOrderingProviderIEN = szGetColumn(szTemp, 1, "\")
            objNewMember.szOrderingProviderName = Trim$(szGetColumn(szTemp, 2, "\") & ", " & szGetColumn(szTemp, 3, "\") & " " & szGetColumn(szTemp, 4, "\"))
        Else
            objNewMember.szOrderingProviderIEN = ""
            objNewMember.szOrderingProviderName = ""
        
        End If
        
        objNewMember.SigningAction = GetOrderSigningActionFromText(szGetColumn(szNewStr, hlZOFSeg.OrderSigningAction + 1, "^"))
        
        objNewMember.szOrderEventDate = FormatHL7DateToMilitary(szGetColumn(szNewStr, hlZOFSeg.OrderEventDate + 1, "^"))
        objNewMember.szOrderEventDateHL7Format = szGetColumn(szNewStr, hlZOFSeg.OrderEventDate + 1, "^")
        
        objNewMember.NurseSignatureNeeded = GetSignatureNeededFromText(szGetColumn(szNewStr, hlZOFSeg.NurseSigNeeded + 1, "^"))
        objNewMember.CounterSignatureNeeded = GetSignatureNeededFromText(szGetColumn(szNewStr, hlZOFSeg.CounterSigNeeded + 1, "^"))
        objNewMember.ProviderSignatureNeeded = GetSignatureNeededFromText(szGetColumn(szNewStr, hlZOFSeg.ProviderSigNeeded + 1, "^"))
        objNewMember.AncillarySignatureNeeded = GetSignatureNeededFromText(szGetColumn(szNewStr, hlZOFSeg.AncillarySigNeeded + 1, "^"))
        
        objNewMember.NurseSignatureInfo.DateTimeHL7Format = szGetColumn(szNewStr, hlZOFSeg.NurseSigDate + 1, "^")
        objNewMember.NurseSignatureInfo.DateTime = FormatHL7DateToMilitary(objNewMember.NurseSignatureInfo.DateTimeHL7Format)
        szTemp = szGetColumn(szNewStr, hlZOFSeg.NurseSignature + 1, "^")
        objNewMember.NurseSignatureInfo.IEN = szGetColumn(szTemp, 1, "\")
        objNewMember.NurseSignatureInfo.szName = Trim$(szGetColumn(szTemp, 2, "\") & ", " & szGetColumn(szTemp, 3, "\") & " " & szGetColumn(szTemp, 4, "\"))
       
        objNewMember.ProviderSignatureInfo.DateTimeHL7Format = szGetColumn(szNewStr, hlZOFSeg.ProviderSigDate + 1, "^")
        objNewMember.ProviderSignatureInfo.DateTime = FormatHL7DateToMilitary(objNewMember.ProviderSignatureInfo.DateTimeHL7Format)
        szTemp = szGetColumn(szNewStr, hlZOFSeg.ProviderSignature + 1, "^")
        objNewMember.ProviderSignatureInfo.IEN = szGetColumn(szTemp, 1, "\")
        objNewMember.ProviderSignatureInfo.szName = Trim$(szGetColumn(szTemp, 2, "\") & ", " & szGetColumn(szTemp, 3, "\") & " " & szGetColumn(szTemp, 4, "\"))
        
        objNewMember.CounterSignatureInfo.DateTimeHL7Format = szGetColumn(szNewStr, hlZOFSeg.CounterSigDate + 1, "^")
        objNewMember.CounterSignatureInfo.DateTime = FormatHL7DateToMilitary(objNewMember.CounterSignatureInfo.DateTimeHL7Format)
        szTemp = szGetColumn(szNewStr, hlZOFSeg.CounterSignature + 1, "^")
        objNewMember.CounterSignatureInfo.IEN = szGetColumn(szTemp, 1, "\")
        objNewMember.CounterSignatureInfo.szName = Trim$(szGetColumn(szTemp, 2, "\") & ", " & szGetColumn(szTemp, 3, "\") & " " & szGetColumn(szTemp, 4, "\"))
        
        
        objNewMember.AncillarySignatureInfo.DateTimeHL7Format = szGetColumn(szNewStr, hlZOFSeg.AncillaryApprovalDate + 1, "^")
        objNewMember.AncillarySignatureInfo.DateTime = FormatHL7DateToMilitary(objNewMember.AncillarySignatureInfo.DateTimeHL7Format)
        szTemp = szGetColumn(szNewStr, hlZOFSeg.AncillaryApprovalBy + 1, "^")
        objNewMember.AncillarySignatureInfo.IEN = szGetColumn(szTemp, 1, "\")
        objNewMember.AncillarySignatureInfo.szName = Trim$(szGetColumn(szTemp, 2, "\") & ", " & szGetColumn(szTemp, 3, "\") & " " & szGetColumn(szTemp, 4, "\"))
        
        objNewMember.PatientInfo.IEN = szGetColumn(szGetColumn(szNewStr, hlZOFSeg.PatientIEN + 1, "^"), 1, "\")
        objNewMember.PatientInfo.Name = szGetColumn(szNewStr, hlZOFSeg.PatientNAME + 1, "^")
        objNewMember.PatientInfo.DOB = szGetColumn(szNewStr, hlZOFSeg.DOB + 1, "^")
        objNewMember.PatientInfo.Fmp = szGetColumn(szNewStr, hlZOFSeg.FMP_SSNSponsor + 1, "^")
        
        If LenB(objNewMember.PatientInfo.Fmp) > 0 Then
            objNewMember.PatientInfo.Fmp = Left$(objNewMember.PatientInfo.Fmp, 6) & "-" & Mid$(objNewMember.PatientInfo.Fmp, 7, 2) & "-" & Right$(objNewMember.PatientInfo.Fmp, 4)
        End If
        
        objNewMember.OrderStartDateTime = FormatHL7DateToMilitary(szGetColumn(szNewStr, hlZOFSeg.OrderStartDateTime + 1, "^"))
        objNewMember.OrderComment = szGetColumn(szNewStr, hlZOFSeg.OrderComment + 1, "^")
        objNewMember.CompletionStatus = GetSignatureCompletionFromText(szGetColumn(szNewStr, hlZOFSeg.CompletionStatus + 1, "^"))
        If iStop Then
            szNewStr = Mid$(szNewStr, iStop, Len(szNewStr))
        Else
            szNewStr = ""
        End If
        Debug.Print objNewMember.OrderComment
        i = InStr(szNewStr, "ZOF")
        mCol.Add objNewMember
    Wend
   
    'now do the ZOH segments
    i = InStr(gszHL7RetMsg, "ZOH")
    
    If i Then
      szNewStr = Mid$(gszHL7RetMsg, i, Len(gszHL7RetMsg))
      i = InStr(szNewStr, "ZOH")
    Else
      '** no ZOD info ...*** ERROR ***
    End If
    
 
    Dim objMember As UnsignedOrder
    While i
        iStart = i
        Set objNewZOH = New UnsignedAddAction
        
        iStop = InStr(iStart + 1, szNewStr, "ZOH")
        objNewZOH.OrderNumber = szGetColumn(szNewStr, hlZOHSeg.OrderId + 1, "^")
        
        'first find it corresponding partner in UnsignedOrders
        For Each objMember In mCol
            If objMember.OrderNumber = objNewZOH.OrderNumber Then
                Exit For
            End If
        Next
        
        
        szTemp = szGetColumn(szNewStr, hlZOHSeg.EnteredBy + 1, "^")
        objNewZOH.szEnteredByIEN = szGetColumn(szTemp, 1, "\")
        objNewZOH.szEnteredByName = Trim$(szGetColumn(szTemp, 2, "\") & ", " & szGetColumn(szTemp, 3, "\") & " " & szGetColumn(szTemp, 4, "\"))
        
        szTemp = szGetColumn(szNewStr, hlZOHSeg.OrderedBy + 1, "^")
        If LenB(szTemp) <> 0 Then
            objNewZOH.szOrderingProviderIEN = szGetColumn(szTemp, 1, "\")
            objNewZOH.szOrderingProviderName = Trim$(szGetColumn(szTemp, 2, "\") & ", " & szGetColumn(szTemp, 3, "\") & " " & szGetColumn(szTemp, 4, "\"))
        Else
            objNewZOH.szOrderingProviderIEN = ""
            objNewZOH.szOrderingProviderName = ""
        
        End If
        
        objNewZOH.SigningAction = GetOrderSigningActionFromText(szGetColumn(szNewStr, hlZOHSeg.OrderSigningAction + 1, "^"))
        objNewZOH.szOrderEventDate = FormatHL7DateToMilitary(szGetColumn(szNewStr, hlZOHSeg.OrderEventDate + 1, "^"))
        objNewZOH.szOrderEventDateHL7Format = szGetColumn(szNewStr, hlZOHSeg.OrderEventDate + 1, "^")
        
        objNewZOH.NurseSignatureNeeded = GetSignatureNeededFromText(szGetColumn(szNewStr, hlZOHSeg.NurseSigNeeded + 1, "^"))
        objNewZOH.CounterSignatureNeeded = GetSignatureNeededFromText(szGetColumn(szNewStr, hlZOHSeg.CounterSigNeeded + 1, "^"))
        objNewZOH.ProviderSignatureNeeded = GetSignatureNeededFromText(szGetColumn(szNewStr, hlZOHSeg.ProviderSigNeeded + 1, "^"))
        'objNewZOH.AncillarySignatureNeeded = GetSignatureNeededFromText(szGetColumn(szNewStr, hlZOHSeg.AncillarySigNeeded + 1, "^"))
        
        objNewZOH.NurseSignatureInfo.DateTime = FormatHL7DateToMilitary(szGetColumn(szNewStr, hlZOHSeg.NurseSigDate + 1, "^"))
        szTemp = szGetColumn(szNewStr, hlZOHSeg.NurseSignature + 1, "^")
        objNewZOH.NurseSignatureInfo.IEN = szGetColumn(szTemp, 1, "\")
        objNewZOH.NurseSignatureInfo.szName = Trim$(szGetColumn(szTemp, 2, "\") & ", " & szGetColumn(szTemp, 3, "\") & " " & szGetColumn(szTemp, 4, "\"))
       
        objNewZOH.ProviderSignatureInfo.DateTime = FormatHL7DateToMilitary(szGetColumn(szNewStr, hlZOHSeg.ProviderSigDate + 1, "^"))
        szTemp = szGetColumn(szNewStr, hlZOHSeg.ProviderSignature + 1, "^")
        objNewZOH.ProviderSignatureInfo.IEN = szGetColumn(szTemp, 1, "\")
        objNewZOH.ProviderSignatureInfo.szName = Trim$(szGetColumn(szTemp, 2, "\") & ", " & szGetColumn(szTemp, 3, "\") & " " & szGetColumn(szTemp, 4, "\"))
        
        
        objNewZOH.CounterSignatureInfo.DateTime = FormatHL7DateToMilitary(szGetColumn(szNewStr, hlZOHSeg.CounterSigDate + 1, "^"))
        szTemp = szGetColumn(szNewStr, hlZOHSeg.CounterSignature + 1, "^")
        objNewZOH.CounterSignatureInfo.IEN = szGetColumn(szTemp, 1, "\")
        objNewZOH.CounterSignatureInfo.szName = Trim$(szGetColumn(szTemp, 2, "\") & ", " & szGetColumn(szTemp, 3, "\") & " " & szGetColumn(szTemp, 4, "\"))
        
        objNewZOH.ModToIVDate = szGetColumn(szNewStr, hlZOHSeg.ModToIVRateDate + 1, "^")
        objNewZOH.ModToIVRate = szGetColumn(szNewStr, hlZOHSeg.ModToIVRate + 1, "^")
        objNewZOH.ActionComment = szGetColumn(szNewStr, hlZOHSeg.ActionComment + 1, "^")
'        objNewZOH.AncillarySignatureInfo.DateTime = FormatHL7DateToMilitary(szGetColumn(szNewStr, hlZOHSeg.AncillaryApprovalDate + 1, "^"))
'        szTemp = szGetColumn(szNewStr, hlZOHSeg.AncillaryApprovalBy + 1, "^")
'        objNewZOH.AncillarySignatureInfo.IEN = szGetColumn(szTemp, 1, "\")
'        objNewZOH.AncillarySignatureInfo.szName = szGetColumn(szTemp, 2, "\")
'
             
        If iStop Then
            szNewStr = Mid$(szNewStr, iStop, Len(szNewStr))
        Else
            szNewStr = ""
        End If
        i = InStr(szNewStr, "ZOH")
        objMember.UnsignedAddActions.Add objNewZOH
    Wend


    'set the properties passed into the method
'    objNewMember.OrderNumber = OrderNumber
'    objNewMember.OrderType = OrderType
'
'    objNewMember.bModifiable = bModifiable
'    objNewMember.szDisplayText = szDisplayText
'    objNewMember.szStatus = szStatus
'    objNewMember.szExpirationDate = szExpirationDate
'    objNewMember.szDispenseStatus = szDispenseStatus
'    objNewMember.szRefillsRemaining = szRefillsRemaining
'    objNewMember.szLastFillDate = szLastFillDate
'    objNewMember.szOnHoldUntilDate = szOnHoldUntilDate
'    objNewMember.szEnteredByIEN = szEnteredByIEN
'    objNewMember.szEnteredByName = szEnteredByName
'    objNewMember.szOrderingProviderIEN = szOrderingProviderIEN
'    objNewMember.szOrderingProviderName = szOrderingProviderName
'    objNewMember.SigningAction = SigningAction
'
'    objNewMember.szOrderEventDate = szOrderEventDate
'
'    objNewMember.NurseSignatureNeeded = NurseSignatureNeeded
'
'    objNewMember.ProviderSignatureNeeded = ProviderSignatureNeeded
'
'    objNewMember.CounterSignatureNeeded = CounterSignatureNeeded
'
'    objNewMember.AncillarySignatureNeeded = AncillarySignatureNeeded
'
'    If IsObject(AncillarySignatureInfo) Then
'        Set objNewMember.AncillarySignatureInfo = AncillarySignatureInfo
'    Else
'        objNewMember.AncillarySignatureInfo = AncillarySignatureInfo
'    End If
'    If IsObject(CounterSignatureInfo) Then
'        Set objNewMember.CounterSignatureInfo = CounterSignatureInfo
'    Else
'        objNewMember.CounterSignatureInfo = CounterSignatureInfo
'    End If
'    If IsObject(NurseSignatureInfo) Then
'        Set objNewMember.NurseSignatureInfo = NurseSignatureInfo
'    Else
'        objNewMember.NurseSignatureInfo = NurseSignatureInfo
'    End If
'    If IsObject(ProviderSignatureInfo) Then
'        Set objNewMember.ProviderSignatureInfo = ProviderSignatureInfo
'    Else
'        objNewMember.ProviderSignatureInfo = ProviderSignatureInfo
'    End If
'    If IsObject(PatientInfo) Then
'        Set objNewMember.PatientInfo = PatientInfo
'    Else
'        objNewMember.PatientInfo = PatientInfo
'    End If
'    Set objNewMember.UnsignedOrders = UnsignedOrders
'    If Len(sKey) = 0 Then
'        mCol.Add objNewMember
'    Else
'        mCol.Add objNewMember, sKey
'    End If


    'return the object created
'    Set Add = objNewMember
'    Set objNewMember = Nothing


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
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub
Public Sub AddFromSummaryCollection(pSum As SummaryType, bClear As Boolean)
    Dim objNewMember As UnsignedOrder
    Dim i As Integer
    
    If bClear = True Then
        For i = 1 To mCol.Count
            mCol.Remove 1
        Next
    End If
    
    Set objNewMember = New UnsignedOrder
    objNewMember.OrderNumber = pSum.OrderId
    objNewMember.SigningAction = pSum.SigningAction
    objNewMember.szOrderEventDateHL7Format = pSum.szOrderEventDateHL7Format
    objNewMember.OrderType = pSum.OrderTypeCode
    objNewMember.szStatus = pSum.Status
    If goHl7Constructor.SignatureClass > 1 Then
        objNewMember.ProviderSignatureInfo.DateTimeHL7Format = Format$(Now, DATE_TIME_ORC7_FORMAT)
        objNewMember.ProviderSignatureInfo.IEN = gszProviderIEN
    Else
        objNewMember.NurseSignatureInfo.DateTimeHL7Format = Format$(Now, DATE_TIME_ORC7_FORMAT)
        objNewMember.NurseSignatureInfo.IEN = gszProviderIEN
    End If
    objNewMember.UserHasSigned = True
    objNewMember.SigningSuccessful = True
    
    mCol.Add objNewMember
    Set objNewMember = Nothing
    
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
    'objNewMember.szStatus = pSum.Status
    If goHl7Constructor.SignatureClass > 1 Then
        objNewMember.ProviderSignatureInfo.DateTimeHL7Format = Format$(Now, DATE_TIME_ORC7_FORMAT)
        objNewMember.ProviderSignatureInfo.IEN = gszProviderIEN
    Else
        objNewMember.NurseSignatureInfo.DateTimeHL7Format = Format$(Now, DATE_TIME_ORC7_FORMAT)
        objNewMember.NurseSignatureInfo.IEN = gszProviderIEN
    End If
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


