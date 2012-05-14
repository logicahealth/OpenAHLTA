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
Attribute VB_Name = "OE_SessionManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
     Private moeSearchEngine As OE_SearchEngine 'local copy
    
     Private mvarRxTablet As RxType 'local copy
     Private mvarRadTablet As RadType 'local copy
     Private mvarLabTablet As LabType 'local copy
     
     Private mvarOrderTypeCode As hlORDERTYPECODE 'local copy
     Private mvarszOverrideReason As String 'local copy
     Private mvarSessionError As hlERRORS 'local copy
     
     Private mvarClinicName As String 'local copy      '#12564
     Private mvarClinicIEN As String 'local copy
     Private mvarMEPRSCode As String 'local copy
     Private mvarMEPRSIen As String 'local copy
     
     Private mvarMTFCode As String 'local copy
     Private mvarszDefOrderPage As String 'local copy
     Private mvarszUserIen As String 'local copy
     Private mvarszProviderIen As String 'local copy
     Private mvarszOrderingProviderIEN As String 'may be different than ProviderIEN if a non-doc is ordering
     Private mvarszConnectFailureText As String 'contains text that describes why connectivity to CHCS I failed; either patient or TCP or Access/Verify
     '==============================================
     
     Public goRxTransactionCol As RxTransactionCollection
     Public goLabTransactionCol As LabTransactionCollection
     Public goRadTransactionCol As RadTransactionCollection
     Public goDeletedLabTablet As LabType   'AMW 2/25/00 These hold tablets temporarily
     Public goDeletedRadTablet As RadType   ' for a delete until they can get
     Public goDeletedRxTablet As RxType     ' into the transactioncol
     
     Private mbCalledAfterInit2 As Boolean
     Private mvarLoginStatus As hlERRORS        '2-8-00; AAB PR#191;
     Public LoginStatus As hlERRORS
     Public gcolClinicColl As ClinicCollection
     
     Public SummaryCollection   As Collection
     Public RxCollection   As Collection
     Public RadCollection  As Collection
     Public LabCollection  As Collection
     
     Public unSignOrders As UnsignedOrders
     
     Const c_mszModuleName As String = "OE_SESSIONMANAGER."
     
     Private mobjdas     As ICHCSII_DAS
     Private I_Das       As CHCSII_CONN.Conn
     Private EncId       As String
 
'============================================

     Implements ICHCSII_AnP
     Implements ICHCSII_CoreAPI2
     Implements ICHCSII_CoreAPI3
     

Const mInitLabDuration As Double = 1.5 '56473 - Jane Shen 12/28/2004
Const mInitRadDuration As Double = 1.5

Public Property Get bCalledAfterInit2() As Boolean
    bCalledAfterInit2 = mbCalledAfterInit2
End Property
Public Property Let bCalledAfterInit2(ByVal vData As Boolean)
    mbCalledAfterInit2 = bCalledAfterInit2
End Property

Public Function GetErrorText(ErrorTextSent) As String
                              
Dim lReturnErr As hlERRORS
Dim szErrorText As String

    lReturnErr = ErrorTextSent
    GetErrorText = Trim$(szErrorText)

End Function


Private Sub ICHCSII_AnP_gridDxClick(Row As Long, Col As Long)
'n/a
End Sub

Private Property Get ICHCSII_AnP_primaryProviderIEN() As String

End Property
Private Property Get ICHCSII_AnP_ShowAllMeds() As Boolean
    ICHCSII_AnP_ShowAllMeds = False
End Property

Private Property Get ICHCSII_AnP_ShowAllRads() As Boolean
    ICHCSII_AnP_ShowAllRads = False
End Property

Private Property Get ICHCSII_AnP_ShowAllLabs() As Boolean
    ICHCSII_AnP_ShowAllLabs = False
End Property

Private Function ICHCSII_CoreAPI3_CloseView2(ByVal viViewType As Integer) As Boolean
    
    If gbBackgroundBusy = False Then
        'CloseView viViewType
        ICHCSII_CoreAPI3_CloseView2 = True  'We will shut down
    Else
        ICHCSII_CoreAPI3_CloseView2 = False
    End If
    
End Function

Public Sub ICHCSII_CoreAPI2_Initialize2(ByVal vlMyProgId As Long, robjComm As Object, robjLogonEnvironment As Object, robjConfig As Object, robjPatient As Object, robjUser As Object)

    Set gobjUser = robjUser
    Set gobjLogonEnvironment = robjLogonEnvironment
    Set gobjPatient = robjPatient
    
    Call Initialize(vlMyProgId, robjComm, robjLogonEnvironment, robjConfig, robjPatient)

End Sub

Public Sub Initialize(ByVal vlMyProgId As Long, robjComm As Object, robjLogonEnvironment As Object, robjConfig As Object, robjPatient As Object)
    
    On Error GoTo ErrHandler
    '---create references to passed in objects---
    Set cwobjCWCommSession = robjComm
    Set cwobjLogonEnviron = robjLogonEnvironment
    Set cwobjConfig = robjConfig
    Set cwobjPatient = robjPatient
           
    Exit Sub

ErrHandler:

    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "OE_SessionManager.Initialize", "OE_SessionManager", vbExclamation)
    
End Sub


Public Sub ChangeMEPRSCode(szMeprsCode As String, szMeprsIEN As String)

          mvarMEPRSIen = szMeprsIEN
          mvarMEPRSCode = szMeprsCode
End Sub

Public Property Let MEPRSIen(ByVal vData As String)
    mvarMEPRSIen = vData
End Property

Friend Property Let MTFCode(ByVal vData As String)
    mvarMTFCode = vData
End Property

Friend Property Let DefOrderPage(ByVal vData As String)
    mvarszDefOrderPage = vData
End Property

Friend Property Let UserIen(ByVal vData As String)
    mvarszUserIen = vData
End Property

Friend Property Let ProviderIen(ByVal vData As String)
    mvarszProviderIen = vData
End Property
Friend Property Let OrderingProviderIen(ByVal vData As String)
    'may be different than ProviderIEN if a non-doc is ordering
    mvarszOrderingProviderIEN = vData
End Property
Public Property Get MEPRSIen() As String
    MEPRSIen = mvarMEPRSIen
End Property

Friend Property Let MEPRSCode(ByVal vData As String)
    mvarMEPRSCode = vData
End Property


Public Property Get MEPRSCode() As String
    MEPRSCode = mvarMEPRSCode
End Property

Friend Property Let ClinicName(ByVal vData As String)
    
End Property


Friend Property Let ClinicIEN(ByVal vData As String)
    mvarClinicIEN = vData
End Property


Public Property Get MTFCode() As String
          MTFCode = mvarMTFCode
End Property

Public Property Get DefOrderPage() As String
    DefOrderPage = ""
End Property

Public Property Get ProviderIen() As String
    ProviderIen = mvarszProviderIen
End Property
Public Property Get OrderingProviderIen() As String
    OrderingProviderIen = mvarszOrderingProviderIEN
End Property
Public Property Get UserIen() As String
    UserIen = mvarszUserIen
End Property

Public Property Get ClinicIEN() As String
    ClinicIEN = mvarClinicIEN
End Property

Public Property Get ClinicName() As String
    ClinicName = mvarClinicName
End Property


Public Function GetOrdersOnPage(lngOrderType As hlORDERTYPECODE) As hlERRORS


    Dim hlErr As hlERRORS
    
    #If DEBUGON Then
      Dim DR  As DebugRec
      'If Not DL Is Nothing Then DL.DebugLog Exe, Module, "GetOrdersOnPage_Gems", "Begin", DR
    #End If
     
    On Error GoTo ErrHandler
    
' Begin:SCR #25819;   Developer: Jane Shen 07/31/2002 05:45 PM
'    DoSplash SPLASH_LOAD, SPLASH_LOADING_ORDERS
' End: SCR #25819;
    
    Set SummaryCollection = New Collection
    
    Select Case lngOrderType
        Case hlALL_ORDERS_CODE:
            Call oeSearchEngine.LoadAllRxOrdersIntoCollections(SummaryCollection, RxCollection, hlErr)
            Call oeSearchEngine.LoadAllRadOrdersIntoCollections(SummaryCollection, RadCollection, hlErr)
            Call oeSearchEngine.LoadAllLabOrdersIntoCollections(SummaryCollection, LabCollection, hlErr)
            
        Case hlRX_ORDERS_CODE:
            Set RxCollection = New Collection
            Call oeSearchEngine.LoadAllRxOrdersIntoCollections(SummaryCollection, RxCollection, hlErr)

        Case hlRAD_ORDERS_CODE:
            Set RadCollection = New Collection
            Call oeSearchEngine.LoadAllRadOrdersIntoCollections(SummaryCollection, RadCollection, hlErr)
            
        Case hlLAB_ORDERS_CODE:
            Set LabCollection = New Collection
            Call oeSearchEngine.LoadAllLabOrdersIntoCollections(SummaryCollection, LabCollection, hlErr)
            
    End Select
        
  Exit Function
  
ErrHandler:
  gobjShared.ShowVBError Err.Number, Err.Description, "OE_SessionManager.GetOrdersOnPage_Gems", "OE", vbExclamation
       
End Function
Public Property Get SpecimenChoices() As Collection
    'Collection of all specimens available for the pre-verified lab
    
    If gobjShared.IsAppMode(modeITT) Then   '<SCR 37864
        Set SpecimenChoices = oeSearchEngine.GetSpecimenDataCollection()
    End If
   
End Property
Friend Function GetExpandedSig(SigText As String, ExpandedSig As String) As hlERRORS

    GetExpandedSig = PerformAction(hlRX_ORDERS_CODE, hlACTION_EXPAND_SIG, SigText)
'    ExpandedSig = goOESession.RXTablet.ExpandedSig     'results are in the tablet

End Function


Public Property Get SessionError() As hlERRORS

    SessionError = mvarSessionError
    
End Property

Friend Property Let OverrideReason(ByVal vData As String)
    mvarszOverrideReason = vData
End Property


Friend Property Get OverrideReason() As String
    OverrideReason = mvarszOverrideReason
End Property

Friend Property Let ConnectFailureText(ByVal vData As String)
 mvarszConnectFailureText = vData
End Property


Friend Property Get ConnectFailureText() As String
    ConnectFailureText = mvarszConnectFailureText
End Property

Public Property Get OrderTypeCode() As hlORDERTYPECODE
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: 'Debug.print X.OrderTypeCode
    OrderTypeCode = mvarOrderTypeCode
End Property

Public Property Get LabTablet() As LabType
    Set LabTablet = mvarLabTablet
End Property

Public Property Get RadTablet() As RadType
    Set RadTablet = mvarRadTablet
End Property

Public Property Get RXTablet() As RxType
    Set RXTablet = mvarRxTablet
End Property
Public Property Set LabTablet(ByRef newLab As LabType)
    Set mvarLabTablet = newLab
End Property
Public Property Set RadTablet(ByRef newRad As RadType)
    Set mvarRadTablet = newRad
End Property
Public Property Set RXTablet(ByRef newRx As RxType)
    Set mvarRxTablet = newRx
End Property

Public Function PerformAction(OrderType As hlORDERTYPECODE, OrderAction As hlACTION, Optional OrderId As String) As hlERRORS
        
    Dim tableName As String
    Dim sSql As String
    Dim rs As Recordset
    Dim sValues As String
    
    mvarOrderTypeCode = OrderType
    gbBackgroundBusy = True
    PerformAction = hlERROR_NONE
    On Error GoTo ErrHandler
    
    Select Case OrderType
        Case hlRX_ORDERS_CODE
            tableName = "OE_RX"
        Case hlLAB_ORDERS_CODE
            tableName = "OE_LAB"
        Case hlRAD_ORDERS_CODE
            tableName = "OE_RAD"
    End Select
    
    Dim deleteType As enumItemType
    
    Select Case OrderAction
    Case hlACTION_DELETE
        If OrderType = hlRX_ORDERS_CODE Then
            
            'SCR 64024 - Frank 1/12/05 - enumItemType.itMed
            deleteType = enumItemType.itMed
            
            If gobjShared.IsAppMode(modeITT) Then
                If oAnP Is Nothing Then
                     sSql = "UPDATE MEDICATIONS SET  STATUSNCID = " & NCID_ORDERSTATUS_DISCONTINUED
                     sSql = sSql & " WHERE ORDERNUMBER = '" & OrderId & "'"
                Else
                    sSql = "DELETE FROM MEDICATIONS WHERE ORDERNUMBER = '" & OrderId & "'"
                End If
                mobjdas.ExecuteSQL sSql
            Else
                sSql = "UPDATE MEDICATIONS SET  STATUSNCID = " & NCID_ORDERSTATUS_DISCONTINUED
                sSql = sSql & " WHERE ORDERNUMBER = '" & OrderId & "'"
                mobjdas.ExecuteSQL sSql
                
                'SCR 50379 - Jane Shen 2/25/2004
                If getMultiOrderCount(OrderId) > 1 Then
                    sSql = "DELETE FROM oe_rx WHERE ORDER_ID = '" & OrderId & "'"
                    sSql = sSql & " And STATUS_NCID = " & NCID_PARTIAL
                    mobjdas.ExecuteSQL sSql
                End If

                sSql = "UPDATE oe_rx SET  STATUS_NCID = " & NCID_ORDERSTATUS_DISCONTINUED
                sSql = sSql & " WHERE ORDER_ID = '" & OrderId & "'"
                mobjdas.ExecuteSQL sSql
                
                '** Change to use Parent_Data_ID and OE_Rx_Data_ID
                '** from Order_IEN -- Order_Ien is mapped med id which
                '** can not distinguish between inpatient continuous order
                '** and duplicate med order
                
                '************* This is also in OrderManagerInpt:OrderMgtOpsInpt.DiscontinueOrderStatus
                sSql = "select rx2.order_id from oe_rx rx1 JOIN oe_rx rx2 ON rx1.Unit_NUmber = rx2.Unit_NUmber" _
                    & " and rx2.PARENT_DATA_ID = rx1.OE_RX_DATA_ID where rx1.ORDER_ID = '" & OrderId & "'" _
                    & " and rx1.patient_type=1 and rx2.STATUS_NCID = " & NCID_PARTIAL
                Set rs = mobjdas.OpenRecordset(sSql)
                If Not rs.EOF Then
                    Do Until rs.EOF
                        sValues = sValues & rs("order_id") & ","
                    Loop
                    sValues = Mid(sValues, 1, Len(sValues) - 1)
                    sSql = "update medications set statusncid = " & NCID_ORDERSTATUS_DISCONTINUED & " where ORDERNUMBER in (" & sValues & ")"
                    mobjdas.ExecuteSQL sSql
                    'Update the oe_rx pending order to discontinued. (for continuous med orders)
                    sSql = "update oe_rx set status_ncid = " & NCID_ORDERSTATUS_DISCONTINUED & " where order_id in(" & sValues & ")"
                    mobjdas.ExecuteSQL sSql
                End If
                
           End If
       Else
       
            'SCR 71878 - check the status, if not pending, can't be deleted.
            sSql = "Select STATUS From " & tableName & " WHERE ORDER_ID = " & I_Das.SQLQuote(OrderId)
            
            Set rs = mobjdas.OpenRecordset(sSql)
            If Not rs.EOF Then
                If rs!Status <> isPending Then
                    MsgBxARMd "Only the Pending Order Can be Deleted.", vbInformation, "Deleting Order"
                    PerformAction = hlERROR_LAB_REASON
                    GoTo CleanUp
                End If
            End If
            
            ' Tracker#62880 - J.Shen 12/15/2004
            sSql = "UPDATE " & tableName & " Set STATUS = " & isDiscontinued & _
                   " WHERE ORDER_ID = '" & OrderId & "'" & " AND IS_ORDER_OCCURRENCE  <> 1 " & _
                   " AND STATUS <> " & isCompleted_Addr
            mobjdas.ExecuteSQL sSql
            
            sSql = "DELETE FROM " & tableName & " WHERE ORDER_ID = '" & OrderId & _
                    "' AND STATUS = " & isPending & " AND IS_ORDER_OCCURRENCE = 1"
            mobjdas.ExecuteSQL sSql
            
            Dim tablet As Object
            
            If OrderType = hlLAB_ORDERS_CODE Then
               deleteType = itLab
               Set tablet = goOESession.LabTablet
            ElseIf OrderType = hlRAD_ORDERS_CODE Then
                deleteType = itRad
                Set tablet = goOESession.RadTablet
            End If
       
            Call SaveToEncNote(tablet)
       End If
        
        ' Begin:SCR #1017;   Developer: Jane Shen 11/08/2002 10:27 AM
        If Not oAnP Is Nothing Then
            Call oAnP.RemoveOrder(OrderId, deleteType)
        End If
        ' End: SCR #1017;
    
    Case hlACTION_HOLD, hlACTION_RELEASE
    
        Dim Status As enumItemStatus
        If OrderAction = hlACTION_HOLD Then
            Status = isHold
        Else
            Status = isActive
        End If
        
        If OrderType = hlRX_ORDERS_CODE Then
            If OrderAction = hlACTION_HOLD Then
                sSql = "update Medications set STATUSNCID = " & NCID_ONHOLD
                mobjdas.ExecuteSQL sSql
                sSql = "update oe_ex set STATUS_NCID = " & NCID_ONHOLD
            Else
                sSql = "update Medications set STATUSNCID = " & NCID_ORDERSTATUS_ACTIVE
                mobjdas.ExecuteSQL sSql
                sSql = "update oe_ex set STATUS_NCID = " & NCID_ORDERSTATUS_ACTIVE
            End If
            sSql = sSql & "WHERE ORDERNUMBER = '" & OrderId & "'"
        Else
            sSql = "update " & tableName & " set STATUS = " & Status
            sSql = sSql & "WHERE ORDER_ID = '" & OrderId & "'"
        End If
        mobjdas.ExecuteSQL sSql
          
    Case hlACTION_EXPAND_SIG  'Use for Sign in ITT
        ' Begin:SCR #43019;   Developer: Jane Shen 10/20/2003 06:00 PM
        ' Note: Function Changed:    changes made in OE_SessionManager.cls:PerformAction()
        ' End: SCR #43019;
        If gobjShared.IsAppMode(modeITT) Then
            sSql = "Update " & tableName & " Set SIGN = NULL where Order_id = '" & OrderId & "'"
            mobjdas.ExecuteSQL sSql
        Else
            '*** SCR 65972
            '*** 2/17/05 - Frank -- Moved code to Function
            goOESession.RXTablet.ExpandedSig = GetRxExpandedSig(goOESession.RXTablet.sigCode, goOESession.RXTablet.DefaultQuantity, goOESession.RXTablet.Refills)
        End If
        
    Case hlACTION_NEW, hlACTION_MODIFY, hlACTION_ADD_TO_QUEUE
        PerformAction = PerformActionSave(OrderType, OrderId, OrderAction)
        
    Case hlACTION_PREMODIFY
        If OrderType = hlRX_ORDERS_CODE Then
            Dim rx_type As RxType
            sSql = "select * from rx_list r, medications m"
            sSql = sSql & " Where r.drug_ien = m.MedicationNCID "
            sSql = sSql & " and   m.OrderNumber = '" & OrderId & "'"
            
            Set rs = mobjdas.OpenRecordset(sSql)
            If Not rs.EOF Then
                Call Deserialize_RX(rs, rx_type, True)
            Else
                'get non-code
                sSql = "select * from medications Where MedicationNCID = 0 "
                sSql = sSql & " and unit_number = " & gobjPatient.UnitNumber
                If gobjShared.AppMode = modeITT Then
                    sSql = sSql & " and StatusNCID not in (" & NCID_ORDERSTATUS_DISCONTINUED & ")"
                End If
                
                Set rs = mobjdas.OpenRecordset(sSql)
                If Not rs.EOF Then
                    Call Deserialize_RX(rs, rx_type, False)
                End If
                
            End If
            Set RXTablet = rx_type
        Else
            PerformAction = PerformActionSave(OrderType, OrderId, OrderAction)
       End If
       
    Case hlACTION_PREVERIFY
    
        If gobjShared.IsAppMode(modeTheater) And OrderType = hlRX_ORDERS_CODE Then
          If Len(gobjShared.CmdLineSwitches("INPATIENT2")) > 0 Then
            sSql = "SELECT DEFAULT_SIG, DEFAULT_REFILL, DEFAULT_ROUTE, DEFAULT_DOSE, "
            sSql = sSql & " DEFAULT_UNIT, CONTENT_UNIT, DOSAGE_FORM, DEFAULT_FREQUENCY "
            sSql = sSql & " From  RX_LIST_FDB "
            sSql = sSql & " Where MEDID = " & goOESession.RXTablet.OrderIen
            sSql = sSql & " And DRUG_LIST = 1"
          Else
            sSql = "SELECT DEFAULT_SIG, DEFAULT_REFILL, DEFAULT_ROUTE, "
            sSql = sSql & " DEFAULT_UNIT, CONTENT_UNIT, DOSAGE_FORM "
            sSql = sSql & " From  RX_LIST "
            sSql = sSql & " Where rx_list.drug_ien = " & goOESession.RXTablet.OrderIen
            sSql = sSql & " and FACILITY_NCID = " & cwobjLogonEnviron.FacilityNCID
          End If
        
            Set rs = mobjdas.OpenRecordset(sSql)
            
            If Not rs.EOF Then
                goOESession.RXTablet.DefaultSig = getSigValues(rs)
            End If
       
        End If

    End Select
            
CleanUp:
    gbBackgroundBusy = False
    Set rs = Nothing
    
    Exit Function
    
ErrHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "OE_SessionManager.PerformAction()", "Gems_OE", vbExclamation)
    PerformAction = hlERROR_GET_ORDERS_FAILED
    Exit Function
    Resume
End Function


Public Property Get oeSearchEngine() As OE_SearchEngine
    Set oeSearchEngine = moeSearchEngine
End Property


Public Function Logout() As hlERRORS

End Function

Public Function Login(szAccessCode As String, szVerifyCode As String, szCHCSIpAddr As String, szCHCSIpPort As String) As hlERRORS
    
End Function

Private Sub Class_Initialize()

On Error GoTo ErrorHandler

    mvarSessionError = hlERROR_NONE     'start clean; created 9-2-99; AAB
    mvarLoginStatus = hlERROR_NONE       '2-8-00; AAB; PR#191
    
    '--- SCR-ARM-16113 MSolano 8/24/01; COMMON TO MULTIPLE ARM-SCRs
    '--- Init routine should be called from first top level creatable coclass
    '--- It can be called multiple times, but will actually run only once.
    '--- Instrument all creatable classes in case dll's multiple entry
    '--- points are actually used
    Call ARM_InitAllTransactionsForOE
    Set gobjShared = New CWShared
    
    Set goRxTransactionCol = New RxTransactionCollection
    Set goLabTransactionCol = New LabTransactionCollection
    Set goRadTransactionCol = New RadTransactionCollection
    Set gcolClinicColl = New ClinicCollection
    Set mvarRxTablet = New RxType
    Set mvarRadTablet = New RadType
    Set mvarLabTablet = New LabType
    
    LoginStatus = hlERROR_NONE
    mbCalledAfterInit2 = True
    bHL7SessionCreated = False
    gbBackgroundBusy = False
    
    '05-30-00; AAB
    IsActionViaSubmitButton(hlRX_ORDERS_CODE) = False
    IsActionViaSubmitButton(hlRAD_ORDERS_CODE) = False
    IsActionViaSubmitButton(hlLAB_ORDERS_CODE) = False
    Set I_Das = New CHCSII_CONN.Conn
    Set mobjdas = I_Das.CHCSII_DAS '<SCR 37864
    Set moeSearchEngine = New OE_SearchEngine
    Exit Sub
    
ErrorHandler:
    DisplayLogErr "OE_SessionManager: Initialize()"
End Sub

Private Sub Class_Terminate()

    Set mobjdas = Nothing

End Sub

Private Function GetHl7Warnings(OrderType As hlORDERTYPECODE, bOrderIDProvided As Boolean) As Boolean
    
    Dim bWarnings As Boolean
    
    bWarnings = False
    
    GetHl7Warnings = bWarnings
    
End Function
Public Sub GetHl7WarningText(OrderType As hlORDERTYPECODE, szTextBody As String, szWarningHeader As String, iNumWarnings As Integer, szTestName As String)
    
    szTextBody = vbNullString
    
End Sub


Public Sub SendErrorMessageToCore(OrderType As hlORDERTYPECODE, OrderAction As hlACTION, ErrorCode As hlERRORS)

Dim szOrderId_AsError As String
  
        On Error GoTo ErrorHandler
        
        szOrderId_AsError = "HL7ERR = " & Trim$(Str$(ErrorCode))

        Select Case OrderType
        Case hlRX_ORDERS_CODE:
            cwobjCWCommRx.Message hlACTION_ORDER_SET, szOrderId_AsError, NCID_BROADCAST_MSG, NCID_GEMS_OE_MED_APPLICATION
            oLogger.LogtoFile c_mszModuleName & "SendErrorMessageToCore - RX", 1, "OrderAction", OrderAction, "szOrderId_AsError", szOrderId_AsError
        
        Case hlLAB_ORDERS_CODE:
            cwobjCWCommLab.Message hlACTION_ORDER_SET, szOrderId_AsError, NCID_BROADCAST_MSG, NCID_GEMS_OE_LAB_APPLICATION
            oLogger.LogtoFile c_mszModuleName & "SendErrorMessageToCore - LAB", 1, "OrderAction", OrderAction, "szOrderId_AsError", szOrderId_AsError
            
        Case hlRAD_ORDERS_CODE:
            cwobjCWCommRad.Message hlACTION_ORDER_SET, szOrderId_AsError, NCID_BROADCAST_MSG, NCID_GEMS_OE_RAD_APPLICATION
            oLogger.LogtoFile c_mszModuleName & "SendErrorMessageToCore - RAD", 1, "OrderAction", OrderAction, "szOrderId_AsError", szOrderId_AsError
              
        End Select
        
        Exit Sub
    
ErrorHandler:
    gobjShared.ShowVBError Err.Number, Err.Description, "OE_SessionManager.SendErrorMessageToCore()", "Gems_OE", vbExclamation
End Sub
Public Sub SendResetTimerMessageToCore(OrderType As hlORDERTYPECODE)

    
    Select Case OrderType
        Case hlRX_ORDERS_CODE
            cwobjCWCommRx.Message cwiRESET_TIMEOUT_TIMER, vbNullString, NCID_BROADCAST_MSG, NCID_GEMS_OE_MED_APPLICATION
            oLogger.LogtoFile c_mszModuleName & "SendResetTimerMessageToCore - RX", 1
    
        Case hlLAB_ORDERS_CODE
            cwobjCWCommLab.Message cwiRESET_TIMEOUT_TIMER, vbNullString, NCID_BROADCAST_MSG, NCID_GEMS_OE_LAB_APPLICATION
            oLogger.LogtoFile c_mszModuleName & "SendResetTimerMessageToCore - LAB", 1
        
        Case hlRAD_ORDERS_CODE
            cwobjCWCommRad.Message cwiRESET_TIMEOUT_TIMER, vbNullString, NCID_BROADCAST_MSG, NCID_GEMS_OE_RAD_APPLICATION
            oLogger.LogtoFile c_mszModuleName & "SendResetTimerMessageToCore - RAD", 1
          
    End Select
    
End Sub


Public Function ICHCSII_AnP_SetAnPParent(objAnPParent As Object) As Boolean
    CommonSetAnP objAnPParent
End Function

Public Property Get UnsignedOrdersCollection() As UnsignedOrders

    Set UnsignedOrdersCollection = unSignOrders

End Function

Public Function GetOrdersToSign(lCursor As hlORDERSIGNCURSOR, szNumOrders, iTotalNumberOfOrders As Integer, bMore As Boolean) As hlERRORS
On Error GoTo ErrHandler

    If unSignOrders Is Nothing Then
        Set unSignOrders = New UnsignedOrders
    End If
    Call unSignOrders.AddUsingHL7Segments
    iTotalNumberOfOrders = unSignOrders.Count
    GetOrdersToSign = hlERROR_NONE
    
    Exit Function
    
ErrHandler:
    gobjShared.ShowVBError Err.Number, Err.Description, "OE_SessionManager.GetOrdersToSign()", "Gems_OE", vbExclamation
End Function
Public Function GetOrderTypeTextfromEnum(lngEnumVar As hlORDERTYPECODE) As String
    Select Case lngEnumVar
        Case 9
            GetOrderTypeTextfromEnum = "RX"
        Case 5
            GetOrderTypeTextfromEnum = "RAD"
        Case 4
            GetOrderTypeTextfromEnum = "LAB"
    End Select
End Function
Public Function SendSignOrders() As hlERRORS
    
On Error GoTo ErrHandler

    SendSignOrders = hlERROR_NONE
    Exit Function

ErrHandler:
    gobjShared.ShowVBError Err.Number, Err.Description, "OE_SessionManager.SendSignOrders()", "Gems_OE", vbExclamation
End Function
Public Function GetSignatureClass() As Integer
   GetSignatureClass = 1
End Function
Private Function PerformActionSave(OrderType As hlORDERTYPECODE, OrderId As String, OrderAction As hlACTION) As hlERRORS
    
    Dim tablet      As Object
    Dim objOrder    As clsItem
    Dim saveData    As Boolean
    Dim tempHolder  As SummaryType
    Dim Status      As enumItemStatus
On Error GoTo ErrHandler
    
    PerformActionSave = hlERROR_NONE

    If gobjShared.IsAppMode(modeITT) And OrderAction = hlACTION_NEW Then
        If duplicateOrders(OrderType) Then Exit Function
    End If
    
    saveData = True
    
    Set objOrder = New clsItem
    objOrder.ItemState = getOrderState(OrderAction)
    
    If Not oAnP Is Nothing Then
        EncId = oAnP.EncounterNumber
    Else
        EncId = 0
    End If
   
    'Orders that don't save to database
    If objOrder.ItemState = isQueued Or OrderAction = hlACTION_PREMODIFY Then
        saveData = False
    End If
    
    If LenB(OrderId) = 0 Then
        If checkItemExistens(OrderType) Then
            OrderId = CreateNewOrder(OrderType)
        Else
            PerformActionSave = hlERROR_ORDERID_NOT_FOUND
            MsgBxARMd "This item is not available in the facility formulary. The order for this item cannot be submitted at this time. " & _
             " The item can be entered into the facility formulary by personnel with the role/privileges. ." & vbCrLf
            Exit Function
        End If
    End If

    Select Case OrderType
    
        Case hlRX_ORDERS_CODE:
            goOESession.RXTablet.OrderId = OrderId
            Frm_OE_RX.NewOrderid = OrderId
            Set tempHolder = setRxObject(objOrder)
            Set tablet = goOESession.RXTablet
            If saveData Then
                PerformActionSave = saveRx(OrderId, tempHolder)
            End If

        Case hlRAD_ORDERS_CODE:
            goOESession.RadTablet.OrderId = OrderId
            Set tempHolder = setRadObject(objOrder)
            Set tablet = goOESession.RadTablet
            If saveData Then
                Status = getEnumSatus(tempHolder.Status)
                PerformActionSave = saveRad(OrderId, Status)
                If IsInpatient Then
                    Call CheckFreqencyState(OrderType, OrderId)
                End If
            End If

        Case hlLAB_ORDERS_CODE:
            goOESession.LabTablet.OrderId = OrderId
            Set tempHolder = setLabObject(objOrder)
            Set tablet = goOESession.LabTablet
            If saveData Then
                Status = getEnumSatus(tempHolder.Status)
                PerformActionSave = saveLab(OrderId, Status)
                If IsInpatient Then
                    Call CheckFreqencyState(OrderType, OrderId)
                End If
            End If
    End Select

    If Not oAnP Is Nothing Then
        Select Case objOrder.ItemState
        Case isNew
            Call oAnP.AddOrder(objOrder)
        Case isQueued
            Call oAnP.AddOrder(objOrder, tablet)
        Case isSaved
            Call oAnP.UpdateOrder(objOrder)
        End Select
    End If
    
    Set tempHolder = Nothing
    
    Exit Function
    
    
ErrHandler:
    gobjShared.ShowVBError Err.Number, Err.Description, "OE_SessionManager.PerformActionSave()", "Gems_OE", vbExclamation
    PerformActionSave = hlERROR_GET_ORDERS_FAILED
    
End Function

Function getOrderState(OrderAction As hlACTION)

Select Case OrderAction
    Case hlACTION_NEW
        getOrderState = isNew
    Case hlACTION_ADD_TO_QUEUE
        getOrderState = isQueued
    Case hlACTION_MODIFY, hlACTION_PREMODIFY
        getOrderState = isSaved
End Select

End Function
Public Sub Reset(OrderType As hlORDERTYPECODE)
        mvarSessionError = hlERRORS.hlERROR_NONE
End Sub

Public Sub RequestToUpdateScreen(OrderType As hlORDERTYPECODE, OrderAction As hlACTION, Optional OrderId As String)
    On Error GoTo ErrHandler
    
    Dim Frm_Active As Form

    Select Case OrderType
        Case hlRX_ORDERS_CODE:
            Set Frm_Active = Frm_OE_RX
            Frm_Active.Cbo_Orders(hlRX_ORDERS_CODE).Clear
            Frm_Active.mbInOrderSet = True
        Case hlLAB_ORDERS_CODE:
            Set Frm_Active = Frm_OE_Lab
            Frm_Active.Cbo_Orders(hlLAB_ORDERS_CODE).Clear
            Frm_Active.mbInOrderSet = True
        Case hlRAD_ORDERS_CODE:
            Set Frm_Active = Frm_OE_Rad
            Frm_Active.Cbo_Orders(hlRAD_ORDERS_CODE).Clear
            Frm_Active.mbInOrderSet = True
    End Select
    
    Frm_Active.mhlOrderAction = OrderAction
    
    If OrderAction = hlACTION_NEW Then
        UpdateScreenWithInfo Frm_Active, OrderType
        Frm_Active.Btn_Submit.Enabled = True
        SaveToQueueControl Frm_Active.Btn_SaveToQueue, OrderAction, False, OrderType
        Frm_Active.mbPreVerifyDone = True
        If OrderType = hlRX_ORDERS_CODE Then
            Frm_Active.Btn_SaveToQueue.Enabled = True
        End If
        
    ElseIf OrderAction = hlACTION_RENEW_MODIFY Or OrderAction = hlACTION_MODIFY Then
        Select Case OrderType
            Case hlRX_ORDERS_CODE:
                RXTablet.OrderDate = Now
            Case hlLAB_ORDERS_CODE:
                LabTablet.Timing.StartDate = Now
            Case hlRAD_ORDERS_CODE:
                LabTablet.Timing.StartDate = Now
        End Select
        UpdateScreenWithInfo Frm_Active, OrderType
        Frm_Active.Btn_Submit.Enabled = True
        SaveToQueueControl Frm_Active.Btn_SaveToQueue, OrderAction, False, OrderType
        Frm_Active.mbPreVerifyDone = True
        Frm_Active.mszOrderID = OrderId
        
    Else
        Frm_Active.mszOrderID = OrderId
        DoActionButton Val(OrderAction), OrderType, Frm_Active
        
    End If

    Exit Sub
    
ErrHandler:
    gobjShared.ShowVBError Err.Number, Err.Description, "OE_SessionManager.getSigValues()", "Gems_OE", vbExclamation
End Sub
Public Sub RequestToUpdateOrders()
     RefreshEachUIsOrdersOnPage
End Sub

Private Function getSigValues(rs As Recordset) As String

    Dim sigV As String
    On Error GoTo ErrHandler

    sigV = vbNullString
    If LenB(gobjShared.CmdLineSwitches("INPATIENT2")) > 0 Then
    
        goOESession.RXTablet.DefaultSig = vbNullString & rs("DEFAULT_SIG")   'sig
        goOESession.RXTablet.DefaultDose = vbNullString & rs("DEFAULT_DOSE")   'sig
        goOESession.RXTablet.DefaultUnit = vbNullString & rs("DEFAULT_FREQUENCY")
        goOESession.RXTablet.DefaultRoute = vbNullString & rs("DEFAULT_ROUTE")
        goOESession.RXTablet.DefaultQuantity = vbNullString & rs("DEFAULT_UNIT")
        goOESession.RXTablet.Refills = vbNullString & rs("DEFAULT_REFILL")
        
        If Len(goOESession.RXTablet.DefaultSig) > 0 Then
            sigV = goOESession.RXTablet.DefaultSig
        Else
            '*** Default Dose
            sigV = goOESession.RXTablet.DefaultDose
            
            
            '*** Default Route
            If Len(sigV) > 0 Then sigV = Trim$(sigV) & " "
            sigV = sigV & goOESession.RXTablet.DefaultRoute
                    
            '*** Frequency
            If Len(sigV) > 0 Then sigV = Trim$(sigV) & " "
            sigV = sigV & goOESession.RXTablet.DefaultUnit

            '*** Qty
            If Len(sigV) > 0 Then sigV = Trim$(sigV) & " "
            sigV = sigV & "#" & goOESession.RXTablet.DefaultQuantity
        
        
            '*** Refill
            If Len(sigV) > 0 Then sigV = Trim$(sigV) & " "
            sigV = sigV & "RF" & goOESession.RXTablet.Refills
        
        End If
    Else
        If Not IsNull(rs("CONTENT_UNIT")) Then
            sigV = sigV & rs("CONTENT_UNIT") & " " 'freq
            goOESession.RXTablet.DefaultUnit = vbNullString & rs("CONTENT_UNIT")
        Else
            goOESession.RXTablet.DefaultUnit = vbNullString
        End If
        
        If Not IsNull(rs("DOSAGE_FORM")) Then sigV = sigV & rs("DOSAGE_FORM") & " "   ' dose
            
        If Not IsNull(rs("DEFAULT_ROUTE")) Then
            sigV = sigV & rs("DEFAULT_ROUTE") & " " 'route
            goOESession.RXTablet.DefaultDaysSupply = vbNullString & rs("DEFAULT_ROUTE")
        Else
            goOESession.RXTablet.DefaultDaysSupply = vbNullString
        End If
        
        If Not IsNull(rs("DEFAULT_UNIT")) Then 'quan
            sigV = sigV & "#" & rs("DEFAULT_UNIT") & " " 'quan
            goOESession.RXTablet.DefaultQuantity = vbNullString & rs("DEFAULT_UNIT")
        Else
            goOESession.RXTablet.DefaultQuantity = vbNullString
        End If
        
        If Not IsNull(rs("DEFAULT_REFILL")) Then
            sigV = sigV & "RF" & rs("DEFAULT_REFILL") & " " ' refill
            goOESession.RXTablet.Refills = vbNullString & rs("DEFAULT_REFILL")
        Else
            goOESession.RXTablet.Refills = vbNullString
        End If
        
        If Not IsNull(rs("DEFAULT_SIG")) Then sigV = sigV & rs("DEFAULT_SIG")  'sig
    End If
    
    getSigValues = sigV

    Exit Function
    
ErrHandler:
    gobjShared.ShowVBError Err.Number, Err.Description, "OE_SessionManager.getSigValues()", "Gems_OE", vbExclamation
End Function

Private Function saveRx(ByVal OrderId As String, ByVal tHolder As SummaryType) As hlERRORS
    Dim sSql        As String
    
    Dim sPendingOrderID As String
    Dim whereRXOrder As String
    Dim whereMEDOrder As String
    
    On Error GoTo ErrHandler
            
    With goOESession.RXTablet
    
        '82521 - 1/13/06 - Frank
        .sigCode = UCase$(.sigCode)
        .ExpandedSig = UCase$(.ExpandedSig)
        
        '***********************
        '** For INPATIENT MED ORDER, need to update both ACTIVE and PENDING med orders
        sPendingOrderID = GetPendingMedOrderID(.OrderId)
        
        If LenB(sPendingOrderID) = 0 Then
            whereRXOrder = " where  ORDER_ID = '" & OrderId & "'"
            whereMEDOrder = " where  ORDERNUMBER = '" & OrderId & "'"
        Else
            whereRXOrder = " where ORDER_ID in ('" & OrderId & "', '" & sPendingOrderID & "')"
            whereMEDOrder = " where ORDERNUMBER in ('" & OrderId & "', '" & sPendingOrderID & "')"
        End If

        
        If LenB(.Status) = 0 Then
            .Status = NCID_PARTIAL ' Order from OrderSet
        ElseIf LenB(tHolder.LastFillDate) > 0 Then
            .Status = NCID_ORDERSTATUS_ACTIVE ' the order was dispensed
        End If
        
        '*** SCR - Frank - 5/10/05
        '*** Make Inpatient Med Order go to PENDING on SUBMIT
        '**64327 Set Status for Inpatient Med Order to Active
''        If .InPatientType = "INPATIENT" Then
''            .Status = NCID_ORDERSTATUS_ACTIVE
''        End If
        
        'Update the Medications table
        sSql = "update Medications set Sig = " & I_Das.SQLQuote(.sigCode, 120) ' SCR 110563 2/14/07 Quan 'SCR 50608 - Shen 3/2/2004
        sSql = sSql & ", Quantity = '" & Val(.Quantity)
        sSql = sSql & "', MEDICATIONSCOMMENT = " & I_Das.SQLQuote(.Comment, 255) 'SCR 50758 - Shen 3/5/2004
        sSql = sSql & ", MedicationNCID = '" & .OrderIen
        sSql = sSql & "', NON_CODED_MEDICATION_NAME = " & I_Das.SQLQuote(.OrderName, 125)
        ' SCR 81867
        If Format(Convert_String_TO_Date(.OrderDate), "MM/DD/YYYY") = Format(Now, "MM/DD/YYYY") Then
            sSql = sSql & ", STARTDATE =  " & SQLDateSub(Now)
        Else
            sSql = sSql & ", STARTDATE =  " & SQLDateSub(Convert_String_TO_Date(.OrderDate))
        End If
        
        'SCR 107029, JQ, 6/8/07
        ' refills or refills remaining? this is the same value we're putting in refills remaining... don't do it
        If Val(tHolder.Refill) = 0 Then sSql = sSql & ", Refills =  " & Val(.Refills)
        
        sSql = sSql & ", OrderDate =  " & SQLDateSub(Now)
        sSql = sSql & ", ORDERINGPROVIDER =  '" & cwobjLogonEnviron.UserName
        sSql = sSql & "', ORDERINGPROVIDERNCID =  '" & cwobjLogonEnviron.UserNCID
        sSql = sSql & "', CLINICNCID =  '" & cwobjLogonEnviron.NursingDivisionNCID & "'"

        If .InPatientType = "INPATIENT" Then
            sSql = sSql & ", PATIENTTYPE = 'I' "
        Else
            sSql = sSql & ", PATIENTTYPE = 'O' "
        End If
    
        '** 1/18/06 -- Frank; Do not update status when user do a MODIFY
        If LenB(sPendingOrderID) = 0 Then sSql = sSql & ", STATUSNCID =  " & .Status
        
        sSql = sSql & ", FDB_MEDID = " & .OrderIen
        
        If Len(gobjShared.CmdLineSwitches("INPATIENT2")) > 0 Then
            sSql = sSql & ", ENDDATE =  " & SQLDateSub(.Timing.StopDate)
            
            '** 7/12/05 - Frank, added Expanded_Sig to Medications
            sSql = sSql & ",EXPANDED_SIG = " & SQLQuote(.ExpandedSig)
     
            If .DispensingLocationNCID = "" Then
                sSql = sSql & ", DISPENSINGLOCATIONNCID =  " & cwobjLogonEnviron.FacilityNCID
            Else
                sSql = sSql & ", DISPENSINGLOCATIONNCID =  " & .DispensingLocationNCID
            End If
            If CLng("0" & .OverrideReasonID) > 0 Then
               sSql = sSql & ", Override_Reason = " & .OverrideReasonID
            End If
            sSql = sSql & ", Override_Comment = '" & .OverrideComment & "'"
            sSql = sSql & ", WARNINGS = '" & .WarningInfo & "'"
            
            If .Status = NCID_ORDERSTATUS_ACTIVE Then
                sSql = sSql & ", FILLDATE =  " & SQLDateSub(Now)
            End If
        Else
            
            sSql = sSql & ", DISPENSINGLOCATIONNCID =  " & cwobjLogonEnviron.FacilityNCID
        End If
        
        If (LenB(tHolder.LastFillDate) > 0) Or Val(tHolder.Refill) <> Val(.Refills) Then
            sSql = sSql & ", RefillsREMAINING = '" & Val(.Refills) & "'"
        End If
        
        sSql = sSql & ", ENC_NUM =  " & EncId
        sSql = sSql & whereMEDOrder
        
        mobjdas.ExecuteSQL sSql
              
        If gobjShared.IsAppMode(modeTheater) And .Status = NCID_ORDERSTATUS_ACTIVE Then
        
            'SCR #65435 - Jane Shen 2/10/2005
            If Len(gobjShared.CmdLineSwitches("INPATIENT2")) > 0 Then
                sSql = "update rx_list_FDB set current_stock = current_stock - " & Val(.Quantity)
                sSql = sSql & " where medid = " & .OrderIen
            Else
                sSql = "update rx_list set current_stock = current_stock - " & Val(.Quantity)
                sSql = sSql & " where drug_ien = " & .OrderIen
            End If
            mobjdas.ExecuteSQL sSql
        End If
        
        'update oe_rx table
        sSql = "update OE_RX set SIG_CODE = " & I_Das.SQLQuote(.sigCode, 120) ' SCR 110563 2/14/07 Quan 'SCR 50608 -Shen 3/2/2004
        sSql = sSql & ", order_name = " & I_Das.SQLQuote(.OrderName, 70)
        sSql = sSql & ", ORDER_IEN = " & .OrderIen
        sSql = sSql & ", Quantity = '" & Val(.Quantity)
        If gobjShared.IsAppMode(modeITT) Then   '<SCR 37864
            sSql = sSql & "', order_date = " & SQLDateSub(Now)
            sSql = sSql & ", time_ordered = " & SQLDateSub(Now)
        Else                                    'SF SCR 77234
            ' SCR 81867
            If Format(Convert_String_TO_Date(.OrderDate), "MM/DD/YYYY") = Format(Now, "MM/DD/YYYY") Then
                sSql = sSql & "', order_date = " & SQLDateSub(Now)
            Else
                sSql = sSql & "', order_date = " & SQLDateSub(Convert_String_TO_Date(.OrderDate)) 'SQLDateSub(Now)
            End If
            sSql = sSql & ", time_ordered = " & SQLDateSub(Now)
            If .Status = NCID_ORDERSTATUS_ACTIVE Then
                sSql = sSql & ", FILL_DATE = " & SQLDateSub(Now)
            End If
            sSql = sSql & ", Refills = " & Val(.Refills)
            
             '** 1/18/06 -- Frank; Do not update status when user do a MODIFY
            If LenB(sPendingOrderID) = 0 Then sSql = sSql & ", STATUS_NCID =  " & .Status


        End If
        sSql = sSql & ", RX_COMMENT = " & I_Das.SQLQuote(.Comment, 255)
        sSql = sSql & ", ordered_by = '" & cwobjLogonEnviron.UserName
        sSql = sSql & "', provider = '" & cwobjLogonEnviron.UserNCID
        sSql = sSql & "', unit_number = '" & gobjPatient.UnitNumber
        sSql = sSql & "', PHARMACY_NAME_SENT = " & I_Das.SQLQuote(.PharmacyNameSent, 100)
        sSql = sSql & ", child_cap =  " & .ChildCap
        sSql = sSql & ", enc_num =  " & EncId
        sSql = sSql & ", enc_fac_ncid = " & cwobjLogonEnviron.FacilityNCID
        
        If IsInpatient Then
            sSql = sSql & ", PROCESS_PRIORITY = " & IIf(.Priority = 0, "Null", .Priority)
            sSql = sSql & GetAckSignSQL & GetInpatientStatus(hlRX_ORDERS_CODE)
        End If
        
        '** 1/19/06 - Frank
'        sSql = sSql & "  where OE_RX_DATA_ID = '" & tHolder.data_id & "'"
        sSql = sSql & whereRXOrder
        mobjdas.ExecuteSQL sSql
                    
                    
        '1/6/05 - Frank
        '*** Set Med Order status to COMPLETE
        '*** When Refill = 0
        If Len(gobjShared.CmdLineSwitches("INPATIENT2")) > 0 Then
            If .InPatientType = "OUTPATIENT" Then
               If .Status = NCID_ORDERSTATUS_ACTIVE And Val(.Refills) = 0 Then
                    sSql = "update MEDICATIONS "
                    sSql = sSql & " set  StatusNCID = " & NCID_ORDERSTATUS_COMPLETE
                    sSql = sSql & whereMEDOrder
                    
                    mobjdas.ExecuteSQL sSql
    
                    sSql = "update oe_rx set "
                    sSql = sSql & " STATUS_NCID = " & NCID_ORDERSTATUS_COMPLETE
                    
                    sSql = sSql & whereRXOrder
                    mobjdas.ExecuteSQL sSql
                End If
            End If
            
            '*** 4/5/05 - Frank
            '*** SCR 67437
            Call SaveMedHistory(OrderId, Val(.Quantity), Val(.Refills), .sigCode, .Status, .Comment, .OrderIen, .OrderName)
        End If
        

    End With

    Exit Function
    
ErrHandler:
    gobjShared.ShowVBError Err.Number, Err.Description, "OE_SessionManager.saveRx()", "Gems_OE", vbExclamation

End Function


Private Sub SaveMedHistory(OrderId As String, Qty As Integer, Refills As Integer, Sig As String, Action As String, Comment As String, OrderIen As String, MedName As String)

    '** SCR 67437 - Frank
    '** Save Med History

    Dim oMed  As DTOs.MedFill
    Dim oMeds As DTOs.MedFills
    
    On Error GoTo Err_Handler

    If Len(OrderId) = 0 Then Exit Sub

    Set oMeds = New DTOs.MedFills
    Set oMed = New DTOs.MedFill
    
    With oMed
        .OrderId = OrderId
        .FillDate = Now
        .Name = cwobjLogonEnviron.UserName
        .Quantity = Qty
        .Sig = Sig
        .Comment = Comment
        .OrderIen = OrderIen
        .OrderName = MedName
        .Refills = Refills
        .UnitNumber = gobjPatient.UnitNumber   '--- SCR 79823  Sherry Wang  10/11/2005
        .FacilityNCID = cwobjLogonEnviron.FacilityNCID
        If (Action = NCID_ORDERSTATUS_ACTIVE) Or (Action = NCID_ORDERSTATUS_COMPLETE) Then
            .Action = "Dispensed"
        Else
            .Action = "Ordered"
            .ActionOrderId = OrderId 'other action id is captured in OM Input dispense class
        End If
    End With
    
    oMeds.Add oMed
    Call gobjShared.DataAccessService.FillMedMapper.Save(oMeds)    '&&&& Jane
    
    Set oMed = Nothing
    Set oMeds = Nothing
    
    Exit Sub
Err_Handler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "Frm_OE_Rx.SaveMedHistory", "GEMS_OE", vbExclamation)

End Sub


'Private Function getPatientType() As String
'
'    '*** Default to Outpatient
'    getPatientType = "O"
'
'    Select Case mPatientType
'        Case "OUTPATIENT"
'            getPatientType = "O"
'        Case "INPATIENT"
'            getPatientType = "I"
'    End Select
'
'End Function

Private Function saveRad(ByVal OrderId As String, ByVal Status As enumItemStatus) As hlERRORS
    Dim sSql        As String

On Error GoTo ErrHandler
        
    
    With goOESession.RadTablet
                       
        sSql = "update oe_rad set "
        sSql = sSql & " enc_fac_ncid = " & cwobjLogonEnviron.FacilityNCID
        sSql = sSql & ", enc_num = " & I_Das.SQLQuote(oAnP.EncounterNumber)
        sSql = sSql & ", order_ien = " & .OrderIen
        sSql = sSql & ", order_name = " & I_Das.SQLQuote(.OrderName, 75)
        sSql = sSql & ", rad_comment = " & I_Das.SQLQuote(.Comment, 255)
        sSql = sSql & ", clinical_impression_text = " & I_Das.SQLQuote(.ClinicalImpressionText, 80)
        sSql = sSql & ", process_priority = " & .ProcessPriority
        'SCR 135527 - LCI - 07/19/2007: Use GetStartDateString is to set StartDate time = AdminTime
        sSql = sSql & ", timing_start_date = " & GetStartDateString(.Timing.AdminTime, .Timing.StartDate)
        'sSql = sSql & ", timing_start_date = " & SQLDateSub(.Timing.StartDate)
        sSql = sSql & ", timing_admin_time = " & I_Das.SQLQuote(.Timing.AdminTime)
        sSql = sSql & ", time_ordered = " & SQLDateSub(Now)
        sSql = sSql & ", provider = " & I_Das.SQLQuote(cwobjLogonEnviron.UserNCID, 255)
        sSql = sSql & ", ORDERED_BY = " & I_Das.SQLQuote(cwobjLogonEnviron.UserName, 100)
        sSql = sSql & ", unit_number = " & gobjPatient.UnitNumber
        sSql = sSql & ", STATUS = " & Status
        sSql = sSql & ", TIMING_DURATION = " & .Timing.Duration
        sSql = sSql & ", TIMING_FREQUENCY = " & I_Das.SQLQuote(.Timing.Frequency, 50)
        sSql = sSql & ", TIMING_ONE_X_OR_CONTINUED = " & .Timing.OnexOrContinued
        sSql = sSql & ", TRANSPORT_MODE = " & .TransportMode
        sSql = sSql & ", WET_READING_WANTED = " & .WetReadingWanted
        If IsInpatient Then
            sSql = sSql & GetAckSignSQL & GetInpatientStatus(hlRAD_ORDERS_CODE)
        End If
        sSql = sSql & " where ORDER_ID = '" & .OrderId & "'"
                    
        mobjdas.ExecuteSQL sSql
    
    End With
    
    If LenB(enterResultBuf) > 0 Then
        Call saveResult(hlRAD_ORDERS_CODE, OrderId)
    End If
    
    Exit Function
    
ErrHandler:
    gobjShared.ShowVBError Err.Number, Err.Description, "OE_SessionManager.saveRad()", "Gems_OE", vbExclamation
    Exit Function
    Resume
End Function

Private Function saveLab(ByVal OrderId As String, ByVal Status As enumItemStatus) As hlERRORS
    Dim sSql        As String
    
On Error GoTo ErrHandler
        
    With goOESession.LabTablet

        sSql = "update oe_lab set ENC_FAC_NCID = " & cwobjLogonEnviron.FacilityNCID
        sSql = sSql & ", ENC_NUM = " & I_Das.SQLQuote(oAnP.EncounterNumber)
        sSql = sSql & ", ORDER_NAME = " & I_Das.SQLQuote(.OrderName, 74)
        sSql = sSql & ", ORDER_IEN = " & .OrderIen
        sSql = sSql & ", LAB_COMMENT = " & I_Das.SQLQuote(.Comment, 255)
        sSql = sSql & ", COLLECT_METHOD = " & I_Das.SQLQuote(.CollectMethod)
        sSql = sSql & ", PROCESS_PRIORITY = " & .ProcessPriority
        sSql = sSql & ", LAB_COLLECT_PRIORITY = " & I_Das.SQLQuote(.LabCollectPriority)
        sSql = sSql & ", TIMING_ADMIN_TIME = " & I_Das.SQLQuote(.Timing.AdminTime)
        'SCR 135527 - LCI - 07/19/2007: Use GetStartDateString is to set StartDate time = AdminTime
        sSql = sSql & ", TIMING_START_DATE = " & GetStartDateString(.Timing.AdminTime, .Timing.StartDate)
        sSql = sSql & ", TIMING_ONE_X_OR_CONTINUED = " & .Timing.OnexOrContinued
        sSql = sSql & ", TIMING_FREQUENCY = " & I_Das.SQLQuote(.Timing.Frequency)
        sSql = sSql & ", UNEXPANDED_TIME = " & I_Das.SQLQuote(.Timing.UnexpandedTime)
        sSql = sSql & ", TIMING_DURATION = " & Val(.Timing.Duration)
        sSql = sSql & ", PROVIDER = " & I_Das.SQLQuote(cwobjLogonEnviron.UserNCID, 100)
        sSql = sSql & ", ORDERED_BY = " & I_Das.SQLQuote(cwobjLogonEnviron.UserName, 100)
        sSql = sSql & ", UNIT_NUMBER = " & gobjPatient.UnitNumber
        sSql = sSql & ", TIME_ORDERED = " & I_Das.FormatDateForDBInsert(Now)
        sSql = sSql & ", STATUS = " & Status
        sSql = sSql & ", SPECIMEN_NAME = " & I_Das.SQLQuote(.Specimen.Name)
        sSql = sSql & ", SPECIMEN_IEN = " & I_Das.SQLQuote(.Specimen.IEN)
        sSql = sSql & ", SPECIMEN_TOPOGRAPHY = " & I_Das.SQLQuote(.Specimen.Topography)
        sSql = sSql & ", SPECIMEN_CAPCOLOR_OR_CONTAINER = " & I_Das.SQLQuote(.Specimen.CapColorOrContainer)
        If IsInpatient Then
            sSql = sSql & GetAckSignSQL & GetInpatientStatus(hlLAB_ORDERS_CODE)
        End If
        sSql = sSql & " where ORDER_ID = '" & OrderId & "'"
        
        mobjdas.ExecuteSQL sSql
    End With
    
    If LenB(enterResultBuf) > 0 Then
        Call saveResult(hlLAB_ORDERS_CODE, OrderId)
    End If

    Exit Function
    
ErrHandler:
    gobjShared.ShowVBError Err.Number, Err.Description, "OE_SessionManager.saveLab", "Gems_OE", vbExclamation
    
End Function

Private Function setRxObject(objOrder As clsItem) As SummaryType

    Dim tHolder As SummaryType
    Dim sSql    As String
    Dim rs      As Recordset
    Dim Count   As Integer
    
On Error GoTo ErrHandler

    Set tHolder = New SummaryType
    With goOESession.RXTablet

        If gobjShared.IsAppMode(modeTheater) Then
            Count = getMultiOrderCount(.OrderId)
            If Count = 0 Then
                Dim dataId As String
                dataId = mobjdas.GetNextID()
                sSql = "INSERT INTO oe_rx (OE_RX_DATA_ID, ORDER_ID) values("
                sSql = sSql & dataId & ",'" & .OrderId & "')"
                mobjdas.ExecuteSQL sSql
            End If
            
            If Count > 1 Then
                sSql = "SELECT OE_RX_DATA_ID, FILL_DATE, Refills FROM OE_RX WHERE ORDER_ID = '" & _
                    .OrderId & "'"
                sSql = sSql & " and STATUS_NCID = " & NCID_ORDERSTATUS_ACTIVE
            Else
                sSql = "SELECT OE_RX_DATA_ID, FILL_DATE, Refills FROM OE_RX WHERE ORDER_ID = '" & _
                    .OrderId & "'"
            End If
            Set rs = mobjdas.OpenRecordset(sSql)

            If Not rs.EOF Then
                Set tHolder = New SummaryType
                tHolder.data_id = rs!OE_RX_DATA_ID
                If (rs!FILL_DATE & vbNullString) <> "" Then
                    tHolder.LastFillDate = rs!FILL_DATE
                End If
                tHolder.Refill = rs!Refills & vbNullString
            End If
        Else
            sSql = "SELECT OE_RX_DATA_ID FROM OE_RX WHERE ORDER_ID = '" & .OrderId & "'"
            Set rs = mobjdas.OpenRecordset(sSql)
            If Not rs.EOF Then
                tHolder.data_id = rs!OE_RX_DATA_ID
            End If
       End If

        objOrder.ItemType = itMed
        objOrder.PropValue(FLD_START_DT) = Convert_String_TO_Date(.OrderDate)
        objOrder.PropValue(FLD_COMMENT) = .Comment
        objOrder.PropValue(FLD_SIG) = Left(.sigCode, 120) ' SCR 110563 2/14/07 Quan
        objOrder.PropValue(FLD_QUANTITY) = Val(.Quantity)
        objOrder.PropValue(FLD_REFILLS) = Val(.Refills)
        objOrder.PropValue(FLD_ORDERIEN) = .OrderIen
        objOrder.PropValue(FLD_ORDERNAME) = .OrderName
        objOrder.PropValue(FLD_ORDEREDBY) = cwobjLogonEnviron.UserName
        objOrder.PropValue(FLD_PROVIDER) = cwobjLogonEnviron.UserName
        objOrder.PropValue(FLD_ORDER_DT) = Now
        objOrder.PropValue(FLD_CHILDCAP) = .ChildCap
    
        objOrder.PropValue(FLD_DATAID) = tHolder.data_id
        objOrder.PropValue(FLD_ORDERID) = .OrderId
    End With
       
    Set setRxObject = tHolder
    rs.Close
    Exit Function
    
ErrHandler:
    gobjShared.ShowVBError Err.Number, Err.Description, "OE_SessionManager.setRxObject", "Gems_OE", vbExclamation
End Function
Private Function setRadObject(objOrder As clsItem) As SummaryType

    Dim tHolder As SummaryType
    Dim rs      As Recordset
    Dim sSql    As String
    
On Error GoTo ErrHandler
    
    sSql = "SELECT OE_RAD_DATA_ID, RESULTS, STATUS FROM OE_RAD WHERE ORDER_ID = '" & _
            goOESession.RadTablet.OrderId & "'"
            
    Set rs = mobjdas.OpenRecordset(sSql)
    
    If Not rs.EOF Then
        Set tHolder = New SummaryType
        tHolder.data_id = rs!OE_RAD_DATA_ID
        tHolder.BriefText = rs!Results & vbNullString
        tHolder.Status = OrderSatusForDisplay(Val(rs!Status & vbNullString))
    End If

    With goOESession.RadTablet
        objOrder.ItemType = itRad
        objOrder.PropValue(FLD_ORDERIEN) = .OrderIen
        objOrder.PropValue(FLD_COMMENT) = .Comment
        objOrder.PropValue(FLD_CLIN_IMPRESS) = .ClinicalImpressionText
        objOrder.PropValue(FLD_TIME_ONE_CONT) = .Timing.OnexOrContinued
        objOrder.PropValue(FLD_TRANS_MD) = .TransportMode
        objOrder.PropValue(FLD_TIME_ADMIN) = .Timing.AdminTime
        objOrder.PropValue(FLD_TIME_DUR) = .Timing.Duration
        objOrder.PropValue(FLD_TIME_FREQ) = .Timing.Frequency
        objOrder.PropValue(FLD_TIME_START) = .Timing.StartDate
        objOrder.PropValue(FLD_ORDERNAME) = .OrderName
        objOrder.PropValue(FLD_ORDEREDBY) = cwobjLogonEnviron.UserName
        objOrder.PropValue(FLD_PROVIDER) = cwobjLogonEnviron.UserName
        objOrder.PropValue(FLD_ORDER_DT) = Now
        objOrder.PropValue(FLD_P_PRIORITY) = .ProcessPriority
        objOrder.PropValue(FLD_ORDERID) = .OrderId
        objOrder.PropValue(FLD_DATAID) = tHolder.data_id
        objOrder.PropValue(FLD_RESULTS) = tHolder.BriefText
        objOrder.PropValue(FLD_RESULTS_STATUS) = tHolder.Status
    End With
    
    Set setRadObject = tHolder
    rs.Close
    Exit Function
    
ErrHandler:
    gobjShared.ShowVBError Err.Number, Err.Description, "OE_SessionManager.setRadObject", "Gems_OE", vbExclamation

End Function
Private Function setLabObject(objOrder As clsItem) As SummaryType
    Dim tHolder As SummaryType
    Dim rs      As Recordset
    Dim sSql    As String
    
On Error GoTo ErrHandler
    
    sSql = "SELECT OE_LAB_DATA_ID,RESULTS,STATUS FROM OE_LAB WHERE ORDER_ID = '" & _
            goOESession.LabTablet.OrderId & "'"
            
    Set rs = mobjdas.OpenRecordset(sSql)
    
    If Not rs.EOF Then
        Set tHolder = New SummaryType
        tHolder.data_id = rs!OE_LAB_DATA_ID
        tHolder.BriefText = rs!Results & vbNullString
        tHolder.Status = OrderSatusForDisplay(Val(rs!Status & vbNullString))
    End If
        
    With goOESession.LabTablet
        objOrder.ItemType = itLab
        objOrder.PropValue(FLD_ORDERIEN) = .OrderIen
        objOrder.PropValue(FLD_ORDERNAME) = .OrderName
        objOrder.PropValue(FLD_COMMENT) = .Comment
        objOrder.PropValue(FLD_COLLECT) = .CollectMethod
        
        'SCR 102117, JQ, 7-28-06; the 2 priority fields were previously reversed:
        objOrder.PropValue(FLD_P_PRIORITY) = .ProcessPriority
        objOrder.PropValue(FLD_LAB_PRIORITY) = .LabCollectPriority
        
        objOrder.PropValue(FLD_TIME_ONE_CONT) = .Timing.OnexOrContinued
        objOrder.PropValue(FLD_TIME_ADMIN) = .Timing.AdminTime
        objOrder.PropValue(FLD_TIME_DUR) = .Timing.Duration
        objOrder.PropValue(FLD_TIME_FREQ) = .Timing.Frequency
        objOrder.PropValue(FLD_TIME_START) = .Timing.StartDate
        objOrder.PropValue(FLD_ORDER_DT) = Now
        objOrder.PropValue(FLD_UNEXP_TIME) = .Timing.UnexpandedTime
        objOrder.PropValue(FLD_ORDEREDBY) = cwobjLogonEnviron.UserName
        objOrder.PropValue(FLD_PROVIDER) = cwobjLogonEnviron.UserName
        objOrder.PropValue(FLD_ORDERID) = .OrderId
        objOrder.PropValue(FLD_DATAID) = tHolder.data_id
        objOrder.PropValue(FLD_RESULTS) = tHolder.BriefText
        objOrder.PropValue(FLD_RESULTS_STATUS) = tHolder.Status
    End With
    
    Set setLabObject = tHolder
    rs.Close
    Exit Function
    
ErrHandler:
    gobjShared.ShowVBError Err.Number, Err.Description, "OE_SessionManager.setLabObject", "Gems_OE", vbExclamation
End Function

Private Function checkItemExistens(ByVal OrderType As hlORDERTYPECODE) As Boolean

    Dim sSql As String
    Dim rs     As Recordset

    On Error GoTo ErrHandler

    Select Case OrderType
    
    Case hlRX_ORDERS_CODE:
        If Len(gobjShared.CmdLineSwitches("INPATIENT2")) > 0 Then
            sSql = "Select count(*) from RX_LIST_FDB" _
                  & " where medid = " & goOESession.RXTablet.OrderIen
        Else
          sSql = "Select count(*) from rx_list where drug_ien = '" & goOESession.RXTablet.OrderIen & "'"
        End If

    Case hlRAD_ORDERS_CODE:
        sSql = "Select count(*) from rad_test where rad_test_ien = " & goOESession.RadTablet.OrderIen & " OR  UPPER(rad_test_name) = " & SQLQuote(goOESession.RadTablet.OrderName)

    Case hlLAB_ORDERS_CODE:
        sSql = "Select count(*) from LAB_TEST_OE where lab_test_ien = " & goOESession.LabTablet.OrderIen & " OR  UPPER(lab_test_name) = " & SQLQuote(goOESession.LabTablet.OrderName)

    End Select
    
    Set rs = mobjdas.OpenRecordset(sSql)
     If rs.Fields.Item(0).Value > 0 Then
        checkItemExistens = True
    Else
        checkItemExistens = False
    End If
    
    rs.Close
    Exit Function
    
ErrHandler:
    gobjShared.ShowVBError Err.Number, Err.Description, "OE_SessionManager.checkItemExistens", "Gems_OE", vbExclamation

End Function

Private Function getEnumSatus(ByVal orderStatus As String) As enumItemStatus
On Error GoTo ErrHandler

    Select Case orderStatus
        Case "Pending":
            If gobjShared.IsAppMode(modeTheater) Then
                getEnumSatus = isPending
            Else
                getEnumSatus = isActive
            End If
        Case "Preliminary": getEnumSatus = isPreliminary
        Case "Completed Addressed": getEnumSatus = isCompleted_Addr
        Case "Completed Unaddressed": getEnumSatus = isCompleted_Unaddr
        Case "Active": getEnumSatus = isActive
        Case "ON HOLD": getEnumSatus = isHold
    End Select
     
    Exit Function
    
ErrHandler:
    gobjShared.ShowVBError Err.Number, Err.Description, "OE_SessionManager.getEnumSatus", "Gems_OE", vbExclamation
End Function
Private Function duplicateOrders(ByVal oType As hlORDERTYPECODE) As Boolean

    Dim oName As String
    Dim i As Integer
    Dim dup As Integer
    On Error GoTo ErrHandler
    
    dup = 0
    
    Select Case oType
    
        Case hlRX_ORDERS_CODE:
            oName = goOESession.RXTablet.OrderName
        Case hlRAD_ORDERS_CODE:
            oName = goOESession.RadTablet.OrderName
        Case hlLAB_ORDERS_CODE:
            oName = goOESession.LabTablet.OrderName
    End Select
    
    For i = 1 To goOESession.SummaryCollection.Count
        If SummaryCollection(i).FullText = oName Then
            dup = dup + 1
        End If
    Next
    
    If dup > 1 Then
        duplicateOrders = True
        Frm_Warning.Show vbModal
    Else
        duplicateOrders = False
    End If
    
    Exit Function
    
ErrHandler:
    gobjShared.ShowVBError Err.Number, Err.Description, "OE_SessionManager.duplicateOrders", "Gems_OE", vbExclamation
End Function


Private Function GetAckSignSQL() As String
Dim sSql As String
On Error GoTo ErrHandler
    
    sSql = sSql & ", NURSE_ACK_FALG = 0 "
    sSql = sSql & ", NURSE_ACK_DATE = NULL "
    sSql = sSql & ", NURSE_ACK_NAME = NULL "
    If gobjUser.HasPrivilegeEx(Priv_Current_Encounter, Priv_Sign) Then
        sSql = sSql & ",SIGN_FLAG = 1"
    End If
    
    GetAckSignSQL = sSql
    
    Exit Function
    
ErrHandler:
    gobjShared.ShowVBError Err.Number, Err.Description, "OE_SessionManager.GetAckSignSQL", "Gems_OE", vbExclamation
End Function

Private Function GetPendingMedOrderID(ByVal OrderId As String) As String
    
    Dim rs As ADODB.Recordset
    Dim sSql As String

    On Error GoTo ErrHandler
    
    GetPendingMedOrderID = vbNullString
    
    '** For Inpatient, Continuous med order, change made to Active (displayed)
    '** need to apply to Pending (next day) also
    
    sSql = "SELECT ORDER_ID FROM OE_RX"
    sSql = sSql & " WHERE PARENT_DATA_ID in (select OE_RX_DATA_ID from oe_rx where order_id = '" & OrderId & "')"
    sSql = sSql & " AND ABS(PATIENT_TYPE) = 1"
    sSql = sSql & " AND (STATUS_NCID = " & NCID_PARTIAL
    sSql = sSql & " OR (STATUS_NCID = " & NCID_ONHOLD & " AND PREVIOUS_ORDER_STATUS = " & NCID_PARTIAL & "))"
    
    
    
    Set rs = mobjdas.OpenRecordset(sSql)
    
    '** Only one pending med order for continuous med order at any give time
    If Not (rs.BOF And rs.EOF) Then
        GetPendingMedOrderID = rs.Fields("ORDER_ID") & vbNullString
    End If
    
    Exit Function
    
ErrHandler:
    gobjShared.ShowVBError Err.Number, Err.Description, "OE_SessionManager.GetPendingMedOrderID", "Gems_OE", vbExclamation


End Function
Private Function GetInpatientStatus(ByVal OrderType As hlORDERTYPECODE) As String
Dim sSql    As String
Dim rs      As Recordset
Dim oType   As String

On Error GoTo ErrHandler
    
    Select Case OrderType
        Case hlRX_ORDERS_CODE
            oType = RXTablet.InPatientType
        Case hlLAB_ORDERS_CODE
            oType = LabTablet.InPatientType
        Case hlRAD_ORDERS_CODE
            oType = RadTablet.InPatientType
    End Select
    
    If oType = "INPATIENT" Then
        GetInpatientStatus = ", PATIENT_TYPE = 1 "
    Else
        GetInpatientStatus = ", PATIENT_TYPE = 0 "
    End If
    
    Set rs = Nothing
    Exit Function
    
ErrHandler:
    gobjShared.ShowVBError Err.Number, Err.Description, "OE_SessionManager.GetInpatientStatus", "Gems_OE", vbExclamation
End Function

Private Sub CheckFreqencyState(ByVal OrderType As hlORDERTYPECODE, ByVal OrderId As String)
Dim StartDateCols As Collection
On Error GoTo ErrHandler

    Select Case OrderType
    
        Case hlRAD_ORDERS_CODE:
            With goOESession.RadTablet.Timing
                If Val(.Duration) > 0 And LenB(.Frequency) > 0 Then
                    If Val(.Duration) > mInitRadDuration Then
                        Call StoreProcedures(OrderId, mInitRadDuration, .Frequency, .StartDate, "RAD")
                    Else
                        Call StoreProcedures(OrderId, .Duration, .Frequency, .StartDate, "RAD")
                    End If
                End If
            End With
        Case hlLAB_ORDERS_CODE:
            With goOESession.LabTablet.Timing
                If Val(.Duration) > 0 And LenB(.Frequency) > 0 Then
                    If Val(.Duration) > mInitLabDuration Then
                        Call StoreProcedures(OrderId, mInitLabDuration, .Frequency, .StartDate, "LAB")
                    Else
                        Call StoreProcedures(OrderId, .Duration, .Frequency, .StartDate, "LAB")
                    End If
                End If
            End With
    End Select

    Exit Sub
    
ErrHandler:
    gobjShared.ShowVBError Err.Number, Err.Description, "OE_SessionManager.CheckFreqencyState", "Gems_OE", vbExclamation
End Sub
Private Sub StoreProcedures(ByVal OrderId As String, ByVal Duration As Single, _
                            ByVal Frequency As String, ByVal StartDate As Date, _
                            ByVal OrderType As String)
                            
On Error GoTo ErrHandler
    Dim oParam As ADODB.Parameter
    Dim oCmd As ADODB.Command
    Dim oDAS As IADO_DAS
    Dim rs  As Recordset
    Dim currDay As Date
    
    Set oDAS = I_Das.ADO_DAS
    Set oCmd = New ADODB.Command
 
' Quan 80315 just pass the start date to SP which will process the following logic

'    'Set start date without time
'    If Frequency = "QAM" Then 'every day at 900
'        If StartDate > DateValue(Now) + (9 / 24) Then
'            StartDate = DateValue(StartDate) + 1
'        Else
'            StartDate = DateValue(StartDate)
'        End If
'    ElseIf Frequency = "QHS" Then 'one time at 22 bedtime
'        If StartDate > DateValue(Now) + (22 / 24) Then
'            StartDate = DateValue(StartDate) + 1
'        Else
'            StartDate = DateValue(StartDate)
'        End If
'    ElseIf Frequency = "BID" Then 'two time per day at 9 and 17,
'        If StartDate < DateValue(Now) + (9 / 24) Then
'            StartDate = DateValue(StartDate)
'        ElseIf StartDate > DateValue(Now) + (17 / 24) Then
'            StartDate = DateValue(StartDate) + 1
'        End If
'     ElseIf Frequency = "TID" Then 'three time per day 9,17, 22,
'        If StartDate < DateValue(Now) + (9 / 24) Then
'            StartDate = DateValue(StartDate)
'        ElseIf StartDate > DateValue(Now) + (22 / 24) Then
'            StartDate = DateValue(StartDate) + 1
'        End If
'    ElseIf Frequency = "QID" Then 'four time per day 8,12,17,22,
'        If StartDate < DateValue(Now) + (8 / 24) Then
'            StartDate = DateValue(StartDate)
'        ElseIf StartDate > DateValue(Now) + (22 / 24) Then
'            StartDate = DateValue(StartDate) + 1
'        End If
'    End If
               
    With oCmd
    
        .CommandType = adCmdStoredProc
        .CommandText = "Theater_Cont_Order_Gen_Pkg.prc_CREATE_ORDERS"
        
        Set oParam = .CreateParameter("p_nOrder_ID", adVarChar, adParamInput, 128, OrderId)
        .Parameters.Append oParam
        Set oParam = .CreateParameter("p_nDuration", adNumeric, adParamInput, 128, Duration)
        .Parameters.Append oParam
        Set oParam = .CreateParameter("p_vFrequency", adVarChar, adParamInput, 128, Frequency)
        .Parameters.Append oParam
        Set oParam = .CreateParameter("p_dStart_Date", adDate, adParamInput, 128, StartDate)
        .Parameters.Append oParam
        Set oParam = .CreateParameter("p_vOrder_Type", adVarChar, adParamInput, 128, OrderType)
        .Parameters.Append oParam
        
        Set oParam = .CreateParameter("p_nStatus", adVarChar, adParamOutput, 128, vbNullString)
        .Parameters.Append oParam
        
    End With
    Set rs = oDAS.OpenRecordsetSP(oCmd)

    Exit Sub

ErrHandler:
    gobjShared.ShowVBError Err.Number, Err.Description, "OE_SessionManager.StoreProcedures", "Gems_OE", vbExclamation
    
End Sub

Private Function GetStartDateString(ByVal sTime As String, ByVal sDate As Date) As String
Dim AmPm As String
Dim Timing As String
On Error GoTo ErrHandler
    'SCR 135527 - LCI - 07/19/2007: Used to set Start Date time stamp based on timing selection
    Timing = Format$(Now, "hh:mm")
    If sTime = "PM" Then
        AmPm = Format$(sDate, "m/dd/yyyy")
        If Timing > "20:00" And Not sDate > Now Then 'If it's AFTER 8:00 PM set to next day
            AmPm = DateAdd("d", 1, AmPm)
        End If
        AmPm = AmPm & " 8:00:00 PM"
        GetStartDateString = I_Das.FormatDateForDBInsert(AmPm)
    ElseIf sTime = "AM" Then
        AmPm = Format$(sDate, "m/dd/yyyy")
        If Timing > "08:00" And Not sDate > Now Then 'If it's AFTER 8:00 AM set to next day
            AmPm = DateAdd("d", 1, AmPm)
        End If
        AmPm = AmPm & " 8:00:00 AM"
        GetStartDateString = I_Das.FormatDateForDBInsert(AmPm)
    Else
        GetStartDateString = I_Das.FormatDateForDBInsert(sDate)
    End If
    
    Exit Function
    
ErrHandler:
    gobjShared.ShowVBError Err.Number, Err.Description, "OE_SessionManager.GetStartDateString", "Gems_OE", vbExclamation
End Function



Public Function PerformScreenMed(MedID As String, MedName As String, ByRef IsUseCancel As Boolean, ByRef OverrideReasonID As String, ByRef OverrideComment As String, ByRef WarningInfo As String, Optional bPregnant As Boolean = False) As hlERRORS
        
    '***********************
    '** Note Patient Allergies will be added to the Drug Screen in Drug2Drug.ScreenDrug()
        
    Dim bSkip As Boolean
    
    Dim oSummaryType As SummaryType
    
    Dim ocolNonFDB As Collection
    Dim oNonFDBMed As Drug2Drug.NonFDBMed
    
    Dim oFDB As Drug2Drug.CHCSII_FDB
    
    Dim oColMeds As Collection
    Dim oMed As Drug2Drug.MedData
    
    Dim UserActionText As String

    
    PerformScreenMed = hlERROR_NONE
    
    On Error GoTo ErrHandler
    
    cwobjCWCommRx.SetStatusText "Screening drugs..."

    
    Screen.MousePointer = vbHourglass
    DoEvents
    
    Set oColMeds = New Collection
    
    Set oMed = New Drug2Drug.MedData
    oMed.Description = MedName
    oMed.ID = MedID
    oMed.Prospective = True
    oColMeds.Add oMed

    Set ocolNonFDB = New Collection
    For Each oSummaryType In goOESession.SummaryCollection
        With oSummaryType
            bSkip = True
            'SCR 59824 and 59819
            '1/11/05 - Frank; Checking Active and Pending Drugs
            If (.Status = NCID_ORDERSTATUS_ACTIVE) _
                Or (.Status = NCID_PARTIAL) Then
                
                    bSkip = False
                
            '*** SCR 59824
            '*** 1/26/05 - Frank
            '*** Treat COMPLETE within last 30 days as Active
            ElseIf (.Status = NCID_ORDERSTATUS_COMPLETE) Then
                If .LastFillDate <> "" Then
                    If DateDiff("d", .LastFillDate & vbNullString, Date$) <= 30 Then
                        bSkip = False
                    
                    End If
                End If
            End If
            
            If bSkip Then
                '** Don't check agains med orders that had been completed
                '** over 30 days or Discontinued
            Else
                If Len(Trim$(.FDBMedID)) = 0 Then 'Not a FDB Med
                    Set oNonFDBMed = New NonFDBMed
                    oNonFDBMed.Description = .FullText
                    oNonFDBMed.LastFilled = .LastFillDate
                    oNonFDBMed.Refills = .Refill
                    oNonFDBMed.Status = GetMedOrderStatus(.Status)
                    oNonFDBMed.OrderNumber = .OrderId
                    
                    ocolNonFDB.Add oNonFDBMed
                Else
                    Set oMed = New Drug2Drug.MedData
                    oMed.Description = .FullText
                    oMed.ID = .FDBMedID
                    oMed.Prospective = False
                    oColMeds.Add oMed
                End If
            End If
            
        End With
    Next oSummaryType


    DoEvents
    If oFDB Is Nothing Then
        Set oFDB = New Drug2Drug.CHCSII_FDB
        Set oFDB.Comm = cwobjCWCommRx
    End If
    
    oFDB.ShowCHCSIIMeds ocolNonFDB
     
    Dim sUnitNumber As String
    sUnitNumber = cwobjPatient.UnitNumber
    
    Call oFDB.SetCallInfo(cwobjCWCommRx, NCID_GEMS_OE_MED_APPLICATION)
    Call oFDB.ScreenDrug(sUnitNumber, oColMeds, UserActionText, OverrideReasonID, OverrideComment, WarningInfo, , bPregnant)
    
    If UserActionText = "" Then
        'SCR 59884 12/15/04 Frank Lourng
        'Display "Drug Checking Sucessful" message from CORE Status Panel
        cwobjCWCommRx.SetStatusText "Drug Interaction Checking Successful."
    Else
        cwobjCWCommRx.SetStatusText ""
    
    End If
    
    '** Clear out the GEMS_OE msg
    Call Frm_OE_RX.SendMessageToPanel("Order is ready for submission", 12, YES)    '3
    
    IsUseCancel = InStr(1, UserActionText, "cancel", vbTextCompare) > 0
    
    Screen.MousePointer = vbDefault
    DoEvents
    Exit Function
    
ErrHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "OE_SessionManager.PerformScreenMed()", "Gems_OE", vbExclamation)
    PerformScreenMed = False
    Call Frm_OE_RX.SendMessageToPanel("", 0, YES)    '3
    Exit Function
    Resume
End Function


Private Function GetMedOrderStatus(Status As String) As String

    Dim sStatus As String
    
    Select Case Status
        Case NCID_ORDERSTATUS_SCHEDULED
            sStatus = "SCHEDULED"
            
        Case NCID_ORDERSTATUS_ACTIVE
            sStatus = "ACTIVE"
            
        Case NCID_ORDERSTATUS_DISCONTINUED
            sStatus = "DISCONTINUED"
    
        Case NCID_ORDERSTATUS_COMPLETE
            sStatus = "COMPLETED"
    
        Case NCID_ONHOLD
            sStatus = "ONHOLD"
    
        Case NCID_ORDERSTATUS_EXPIRED, NCID_EXPIRED
            sStatus = "EXPIRED"
            
        Case NCID_ORDERSTATUS_EXPIRED, NCID_EXPIRED
            sStatus = "PENDING"
    End Select
    
    GetMedOrderStatus = sStatus
    
End Function

Private Sub SaveToEncNote(ByVal tablet As Object)

    On Error GoTo ErrHandler

    Dim objEnc As CHCSIIEncounterCurrent.EncounterParent
    Dim RTFNote As String
    
    RTFNote = "Order by: " & tablet.OrderingProviderLname & vbCrLf _
              & "Order: " & tablet.OrderName & vbCrLf _
              & "Order ID: " & tablet.OrderId & vbCrLf _
              & "Status: Discontinued " & vbCrLf

    Set objEnc = cwobjCWComm.InitializeOLEServer(NCID_ENCOUNTER_APPLICATION)
    If objEnc.CurrentEncounter Is Nothing Then
    Else
        Call objEnc.CurrentEncounter.AddTextNote("", RTFNote)
    End If
    Set objEnc = Nothing


    Exit Sub
    
ErrHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "OE_SessionManager.SaveToEncNote()", "Gems_OE", vbExclamation)
End Sub

