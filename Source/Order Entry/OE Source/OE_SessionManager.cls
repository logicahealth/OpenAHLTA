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

     Private oHl7Constructor As Hl7Constructor

'Local variable(s) to hold property value(s)
     Private moeSearchEngine As OE_SearchEngine 'local copy
    
     Private mvarRxTablet As RxType 'local copy
     Private mvarRadTablet As RadType 'local copy
     Private mvarLabTablet As LabType 'local copy
     Private mvarOtherTablet As BaseTablet 'local copy
     
     Private mvarOrderTypeCode As hlORDERTYPECODE 'local copy
     Private mvarszOverrideReason As String 'local copy
     Private mvarSessionError As hlERRORS 'local copy
     Private mvarSignatureClass As Integer 'local copy
     
     Private mvarClinicName As String 'local copy      '#12564
     Private mvarClinicIEN As String 'local copy
     Private mvarMEPRSCode As String 'local copy
     Private mvarMEPRSIen As String 'local copy
     Private mvarApptIen As String 'local copy 'SCR 31315
     
     'Added 02-09-2000; AAB for CT (Consult-Tracking);  = PR#191/SCR#115
     Private mvarMTFCode As String 'local copy
     Private mvarszDefOrderPage As String 'local copy
     Private mvarszUserIen As String 'local copy
     Private mvarszProviderIen As String 'local copy
     Private mvarszOrderingProviderIen As String 'may be different than ProviderIEN if a non-doc is ordering
     Private mvarszConnectFailureText As String 'contains text that describes why connectivity to CHCS I failed; either patient or TCP or Access/Verify
     '==============================================
     
     Public goRxTransactionCol As RxTransactionCollection
     Public goLabTransactionCol As LabTransactionCollection
     Public goRadTransactionCol As RadTransactionCollection
     Public goDeletedLabTablet As LabType   'AMW 2/25/00 These hold tablets temporarily
     Public goDeletedRadTablet As RadType   ' for a delete until they can get
     Public goDeletedRxTablet As RxType     ' into the transactioncol
     
     Private mbCalledAfterInit2 As Boolean
     Private mbCalledFromTabletUIs As Boolean
     Private mvarLoginStatus As hlERRORS        '2-8-00; AAB PR#191;
     
     Public gcolClinicColl As ClinicCollection
     Const c_mszModuleName As String = "OE_SESSIONMANAGER."
     
     'Private mvarPatientRequired As Boolean
     
'============================================

     Implements ICHCSII_AnP   'PR-286; AAB
     Implements ICHCSII_CoreAPI2
     Implements ICHCSII_CoreAPI3
     
     

Friend Property Let SessionError(ByVal vData As hlERRORS)

               'Allow [letting] 11-18-01; AAB: for CANCEL of AP pop-ups
               'mvarSessionError As hlERRORS
               '' SessionError = mvarSessionError
             mvarSessionError = vData
    
End Property
Private Function IsError_Recoverable(OrderType As hlORDERTYPECODE, OrderAction As hlACTION, Optional ErrorNumber As Integer, Optional Hl7ReturnStatusCode As String) As Integer

'NEW 04-05-00; AAB for PR#221

'some error numbers, eg 7 " not available" allow NO recovery
          'We can see if ANY disallowed errors exist,  (THIS VERSION)
          ' or search ONLY for the error number passed in: (FUTURE)

'LOGIC redesign: 4-6-00: AAB:
'         ANY hard = cannot proceed;
'         NO HARD plus ALL soft = Information-icon and OK button with recover allowance;
'         NO HARD but NOT all soft (some = not in soft table) has exclamation box with CHOICE (yes/no)

'RETURN VALUES: 0=SOFT 1=HARD 2=CHOICE
''Public Const giERRORS_ARE_SOFT = 0
''Public Const giERRORS_ARE_HARD = 1
''Public Const giERRORS_ARE_CHOICE = 2

Dim iCount As Integer
Dim iErrCount As Integer
Dim iErrorNumber As Integer
Dim piReturn As Integer
Dim objOrderObject As Object

               Select Case OrderType
               Case hlRX_ORDERS_CODE
                    Set objOrderObject = oHl7Constructor.RxError
               Case hlRAD_ORDERS_CODE
                    Set objOrderObject = oHl7Constructor.RadError
               Case hlLAB_ORDERS_CODE
                    Set objOrderObject = oHl7Constructor.LabError
               Case Else
                    Set objOrderObject = oHl7Constructor.OtherError
               End Select

'Uses a lookup table of know HARD errors

                    IsError_Recoverable = giERRORS_ARE_HARD    'HARD default  = 1
                    iErrCount = objOrderObject.Count
                    If iErrCount > 0 Then
            
                    'ANY know hard error is immediate disqualification
                         For iCount = 1 To iErrCount
                               iErrorNumber = Val(objOrderObject(iCount).TypeCode)
                               piReturn = IsError_In_HardTable(iErrorNumber, OrderAction)
                               If piReturn = YES Then     'HARD error found
                                    IsError_Recoverable = giERRORS_ARE_HARD     '1
                                    Exit Function
                               End If
                         Next iCount
                         
                   'next, ALL errors must match KNOW-SOFT errors
                        '' For iCount = 1 To oHl7Constructor.RxError.Count
                         For iCount = 1 To iErrCount
                         
                               ''iErrorNumber = Val(oHl7Constructor.RxError(iCount).TypeCode)
                               iErrorNumber = Val(objOrderObject(iCount).TypeCode)
                               
                               piReturn = IsError_In_SoftTable(iErrorNumber)
                               If piReturn = NO Then     'ALL soft error codes must be present to be SOFT (else = CHOICE)
                                    IsError_Recoverable = giERRORS_ARE_CHOICE   '2
                                    Exit Function
                               End If
                         Next iCount
                   
                   
                         IsError_Recoverable = giERRORS_ARE_SOFT '0
               
               Else
                    'we should not be here if no errors
                     IsError_Recoverable = giERRORS_ARE_SOFT     '0
                    Exit Function
               End If
             
''          Case hlRad_ORDERS_CODE
''               iErrCount = oHl7Constructor.RadError.Count
''               If oHl7Constructor.RadError.Count > 0 Then
''                  For iCount = 1 To oHl7Constructor.RadError.Count
''                       'szErrMsg = szErrMsg & oHl7Constructor.RadError(iCount).Text & vbCrLf
''                  Next iCount
''                  IsError_Recoverable = True
''               End If
''
''          Case hlLab_ORDERS_CODE
''               iErrCount = oHl7Constructor.LabError.Count
''               If oHl7Constructor.LabError.Count > 0 Then
''                  For iCount = 1 To oHl7Constructor.LabError.Count
''                      ' szErrMsg = szErrMsg & oHl7Constructor.LabError(iCount).Text & vbCrLf
''                  Next iCount
''                  IsError_Recoverable = True
''               End If
''
''          End Select

End Function


Private Function IsError_In_HardTable(ErrorNumber As Integer, OrderAction As hlACTION) As Boolean

'NEW 04-05-00; AAB for PR#220

               IsError_In_HardTable = NO
               
               Select Case ErrorNumber
                Case 7    'NOT available
                    IsError_In_HardTable = YES
                Case 8    'NOT available
                    IsError_In_HardTable = YES
                Case 35    'NO keys
                    IsError_In_HardTable = YES
                Case 26   'Order can not be modified.renewed
                    IsError_In_HardTable = YES
                Case 54   'This procedure does not have an assigned Radiology Location.
                    IsError_In_HardTable = YES
                
                'Added 7-11-00; AAB PR#274
                Case 25   'Order can not be cancelled or put on hold
                    IsError_In_HardTable = YES
                Case 284   'Non Provider cannot override allergies
                    IsError_In_HardTable = YES
                Case 131 'added 4/30/02 AKM. If MEPRS Code, cannot procede on a preverify
                    If OrderAction = hlACTION_PREVERIFY Then
                        IsError_In_HardTable = YES
                    End If
                    
                    'Added 5/13/02 For lab:
                    '287^This order can not be modified.  Specimen has been logged into Lab.
                    '288^This order can not be cancelled.  Specimen has been logged into Lab.
                    '289^This order can not be placed on hold.  Specimen has been logged into Lab.
                Case 287
                    IsError_In_HardTable = YES
                Case 288
                    IsError_In_HardTable = YES
                Case 289
                    IsError_In_HardTable = YES
   
               End Select
     
End Function
Private Function IsError_In_SoftTable(ErrorNumber As Integer) As Boolean

'NEW 04-05-00; AAB for PR#220

               IsError_In_SoftTable = NO
               
               Select Case ErrorNumber
               Case 30    'Inpatient Use Only
                    IsError_In_SoftTable = YES
                Case 46       'missing quantity
                    IsError_In_SoftTable = YES
               Case 52        'sent QTY exceeds suggested-Max QTY
                    IsError_In_SoftTable = YES
               Case 80        '[AR] order timing
                    IsError_In_SoftTable = YES
               Case 51       'rqstd days supply exceeds max
                    IsError_In_SoftTable = YES
              Case 110       'Clinical impression is required for this order
                    IsError_In_SoftTable = YES
               ' Case 52        'sent QTY exceeds suggested-Max QTY
                    'IsError_In_SoftTable = YES
              ' Case 52        'sent QTY exceeds suggested-Max QTY
                    'IsError_In_SoftTable = YES
              End Select
     
End Function


Public Property Get LoginStatus() As hlERRORS

'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LoginStatus

    LoginStatus = mvarLoginStatus
    
End Property

Public Property Get GetCHCSPatientInfo() As PatientType

   'SCR 42781 add Patient Name and FMP to Warnings Screen

    Set GetCHCSPatientInfo = oHl7Constructor.goPatient

End Property




'Public Property Let CHCSPatientName(ByVal vData As String)
''used when assigning a value to the property, on the left side of an assignment.
''Syntax: X.CHCSPatientName = 5
'    mvarCHCSPatientName = vData
'End Property


Public Property Get CHCSPatientName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CHCSPatientName
    CHCSPatientName = oHl7Constructor.goPatient.Name
End Property
Public Property Get bCalledAfterInit2() As Boolean
    bCalledAfterInit2 = mbCalledAfterInit2
End Property
Public Property Let bCalledAfterInit2(ByVal vData As Boolean)
    mbCalledAfterInit2 = vData
End Property
Public Property Get bCalledFromTabletUIs() As Boolean
    bCalledFromTabletUIs = mbCalledFromTabletUIs
End Property
Public Property Let bCalledFromTabletUIs(ByVal vData As Boolean)
    mbCalledFromTabletUIs = vData
End Property
'Public Property Let bPatientRequired(ByVal vData As Boolean)
'    mvarPatientRequired = vData
'End Property
'Public Property Get bPatientRequired() As Boolean
'    bPatientRequired = mvarPatientRequired
'End Property
Public Function GetErrorText(ErrorTextSent) As String
               
     'NEW 02-09-2000; AAB for  = PR#191/SCR#115; oHl7Constructor.GetErrorText (hlErr)
               
Dim lReturnErr As hlERRORS
Dim szErrorText As String

               lReturnErr = ErrorTextSent
               szErrorText = oHl7Constructor.GetErrorText(lReturnErr)
               GetErrorText = Trim$(szErrorText)

End Function

Public Sub PatientSelected()
    goOESession.ReleaseAndSelectPatient
End Sub

Public Sub RequestToUpdateOrders()
     RefreshEachUIsOrdersOnPage         'we provide a public link through SessionManager to the Refresh routines
End Sub

Public Function SendHl7Message(szHL7_Msg As String) As String

Dim szHL7RetMsg As String

'NEW 02-09-2000; AAB for  = PR#191/SCR#115 (Consult Tracking)

               oHl7Constructor.HL7MsgHandler szHL7_Msg


'               oHl7Constructor.oHL7Server.SendHl7Message (szHL7_Msg)
               szHL7RetMsg = Space$(32800)
               
               ''Debug.Print "Returned from SendHL7Message"
               '---------------------------------------------------
               oHl7Constructor.oHL7Server.GetHL7Message (szHL7RetMsg)
               szHL7RetMsg = Trim$(szHL7RetMsg)
               
               
               SendHl7Message = szHL7RetMsg
    
''    ShowHL7InboundDebugMessages
''
''    gszHL7RetCode = szGetSegSeq(gszHL7RetMsg, 1, "MSA", hlMSASeg.AckCode)  'added 10-29-97; AAB
''    gszErrMsg = szGetSegSeq(gszHL7RetMsg, 1, "MSA", hlMSASeg.ErrorCondition)
    
    

End Function

Private Sub ICHCSII_AnP_DisplayEncounterLockMessage()

End Sub

Private Property Get ICHCSII_AnP_EncounterApptIEN() As String

End Property

Private Property Get ICHCSII_AnP_EncounterClinicNCID() As String

End Property

Private Sub ICHCSII_AnP_gridDxClick(Row As Long, Col As Long)
'n/a
End Sub

Private Property Get ICHCSII_AnP_IsEncounterLocked() As Boolean

End Property

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
    On Error Resume Next
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
    Set cwobjCWCommSession = robjComm
    
    Set gobjPatient = robjPatient
    
               'Set gobjComm = robjComm
               'Set goOESession.robjUser = robjUser
    gbResults = CreateHL7Session       'function return tested 2-8-00; AAB PR#191
   
    If gbResults = True Then
          Call Initialize(vlMyProgId, robjComm, robjLogonEnvironment, robjConfig, robjPatient)
''     Else
''           Debug.Print
    End If
    
End Sub

Public Sub Initialize(ByVal vlMyProgId As Long, robjComm As Object, robjLogonEnvironment As Object, robjConfig As Object, robjPatient As Object)
    
    On Error GoTo ErrHandler
    
'    '---save the "ProgId" assigned to me by the CW Core---
'    cwlMyProgId = vlMyProgId
    
    '---create references to passed in objects---
    'Set cwobjCWCommSession = robjComm 'moved up to Initialize2
    Set cwobjLogonEnviron = robjLogonEnvironment
    Set cwobjConfig = robjConfig
    Set cwobjPatient = robjPatient
    
'    '---set global flags---
'    cwbMainViewRunning = False
'
    '****************************************************************************
    '  TO DO: $$$
    '    Add code here to create objects, initialize global variables, etc.
    '****************************************************************************
       
    Exit Sub

ErrHandler:

    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "OE_SessionManager.Initialize", "OE_SessionManager", vbExclamation)
    Exit Sub
    
End Sub


Public Sub ChangeMEPRSCode(szMEPRSCode As String, szMEPRSIEN As String)

     'This routine updates new properties with changed information
          mvarMEPRSIen = szMEPRSIEN
          mvarMEPRSCode = szMEPRSCode
          oHl7Constructor.szMEPRSIEN = szMEPRSIEN
          oHl7Constructor.szMEPRSCode = szMEPRSCode
    
End Sub

Public Property Let MEPRSIen(ByVal vData As String)
' = PR#191/SCR#115
'Was PUBLIC; made PRIVATE; oops, should have left public; fixed 02-22-00; AAB
'
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MEPRSIen = 5
    mvarMEPRSIen = vData
    oHl7Constructor.szMEPRSIEN = vData
End Property
Public Property Let ApptIen(ByVal vData As String) 'SCR 31315
    mvarApptIen = vData
    oHl7Constructor.szApptIEN = vData
End Property

Friend Property Let MTFCode(ByVal vData As String)
' = PR#191/SCR#115
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MTFCode = 5
    mvarMTFCode = vData
End Property

Friend Property Let DefOrderPage(ByVal vData As String)
'= PR#191/SCR#115
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DefOrderPage = 5
    mvarszDefOrderPage = vData
End Property

Friend Property Let UserIen(ByVal vData As String)
' = PR#191/SCR#115
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.UserIen  = 5
    mvarszUserIen = vData
End Property

Friend Property Let ProviderIEN(ByVal vData As String)
' = PR#191/SCR#115
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ProviderIen  = 5
    mvarszProviderIen = vData
End Property
Public Property Let OrderingProviderIen(ByVal vData As String)
    'may be different than ProviderIEN if a non-doc is ordering
    'NPOE needs to make this public for consults
    mvarszOrderingProviderIen = vData
    oHl7Constructor.szOrderingProviderIEN = vData
End Property
Public Property Get MEPRSIen() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: 'Debug.print X.MEPRSIen
    MEPRSIen = mvarMEPRSIen
End Property



Friend Property Let MEPRSCode(ByVal vData As String)
' = PR#191/SCR#115
'Was PUBLIC; made PRIVATE
'
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MEPRSCode = 5
    mvarMEPRSCode = vData
    oHl7Constructor.szMEPRSCode = vData
End Property

Public Property Get MEPRSCode() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: 'Debug.print X.MEPRSCode
    MEPRSCode = mvarMEPRSCode
End Property

'Friend Property Get HL7MSG(Index As Integer) As String
''used when retrieving value of a property, on the right side of an assignment.
''Syntax: 'Debug.print X.MEPRSCode
'    HL7MSG = oHl7Constructor.RetMsg(Index)
'End Property


Friend Property Let ClinicName(ByVal vData As String)
          ' = PR#191/SCR#115
          'Was PUBLIC; made PRIVATE
          '
          'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ClinicIEN = 5
    mvarClinicName = vData
   ' oHl7Constructor.szClinicName = vData
    
End Property


Friend Property Let ClinicIEN(ByVal vData As String)
' = PR#191/SCR#115
'Was PUBLIC; made PRIVATE
'
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ClinicIEN = 5
    mvarClinicIEN = vData
    oHl7Constructor.szClinicIEN = vData
    
End Property


Public Property Get MTFCode() As String
' = PR#191/SCR#115
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: 'Debug.print X.MTFCode
          MTFCode = mvarMTFCode
End Property

Public Property Get DefOrderPage() As String

 '= PR#191/SCR#115
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: 'Debug.print X.DefOrderPage

    ''DefOrderPage = mvarszDefOrderPage
     'recommended fix for SCR-6963 [PCR-4296]; 08-23-00; AAB
     
    If Not oHl7Constructor Is Nothing Then
        DefOrderPage = oHl7Constructor.szDefOrderPage
    Else
        DefOrderPage = ""
    End If
    
End Property


Public Property Get ProviderIEN() As String
' = PR#191/SCR#115
    ProviderIEN = mvarszProviderIen
End Property
Public Property Get OrderingProviderIen() As String
    'may be different from ProviderIEN if non-doc is ordering
    OrderingProviderIen = mvarszOrderingProviderIen
End Property
Public Property Get UserIen() As String
' = PR#191/SCR#115
    UserIen = mvarszUserIen
End Property

Public Property Get ClinicIEN() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: 'Debug.print X.ClinicIEN
    ClinicIEN = mvarClinicIEN
End Property

Public Property Get ClinicName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: 'Debug.print X.ClinicIEN
    ClinicName = mvarClinicName
End Property

Public Property Get SummaryCollection() As Collection

    'SummaryCollection is list of orders on page.  It is filled out after GetOrdersOnPage message
    
          Set SummaryCollection = oHl7Constructor.Summary
    
End Property
Public Property Get UnsignedOrdersCollection() As UnsignedOrders

    'UnsignedOrders is list of orders to be signed.
    
          Set UnsignedOrdersCollection = oHl7Constructor.colUnsignedOrder
    
End Property


Public Function GetOrdersOnPage(lngOrderType As hlORDERTYPECODE) As hlERRORS

    'Use the SummaryCollection after invoking this method to get the list of orders
    
    Dim hlErr As hlERRORS
    Dim iCount As Integer
    Dim szErrMsg As String
    
    'Added 2-10-00; AMW:Used for display of Splash Screen
    DoSplash SPLASH_LOAD, SPLASH_LOADING_ORDERS
    
          ''If mvarLoginStatus <> hlERROR_LOGIN_COMM_ERROR Then     '2-8-00; AAB PR#191
     If mvarLoginStatus = hlERROR_NONE Then    'for PR#191

              hlErr = oHl7Constructor.GetOrders(lngOrderType)
                 'Check timeout errors
               If hlErr <> hlERROR_NONE Then
                         If hlErr = hlERROR_TIMED_OUT Or hlErr = hlERROR_NOT_LOGGED_IN Then
                                   If ReconnectToCHCS <> hlERROR_NONE Then
                                                'MsgBxARMd oHl7Constructor.GetErrorText(hlErr), vbCritical, "CHCS I Error"
                                                ''Exit Function
                                             GoTo ExitFunctionJmp:
                                   Else
                                        'resend the message
                                             hlErr = oHl7Constructor.GetOrders(lngOrderType)
                                   End If
                                
          '               Else  (ALL OTHER ERRORS) ... was disabled; re-enabled 11-01-00 AAB
          '                      GenerateLoginFailureMessage hlErr, gobjPatient.MedicalRecordNumber, gobjPatient.Name
                         Else
                                If oHl7Constructor.OtherError.Count > 0 Then
                                    szErrMsg = "There is a problem getting order information for: " & gobjPatient.MedicalRecordNumber _
                                                & " with IEN of: " & gobjPatient.Name & vbCrLf
                                    For iCount = 1 To oHl7Constructor.SignError.Count
                                        szErrMsg = szErrMsg & oHl7Constructor.SignError(iCount).Text & vbCrLf
                                    Next iCount
                
                                End If
                                If LenB(szErrMsg) = 0 Then
                                    MsgBxARMd oHl7Constructor.GetErrorText(hlErr), vbCritical, "CHCS I Error"
                                Else
                                    MsgBxARMd szErrMsg, vbCritical, "CHCS I Error"
                                End If
                                'Exit Function
                         End If
                  
               'Else  -- removed 11-01-00 AAB; no code
                End If

     End If    '2-8-00; AAB PR#191
    
ExitFunctionJmp:
    
               ''SetBusy lngOrderType, False
          GetOrdersOnPage = hlErr
    
End Function
Public Function GetOrdersToSign(lCursor As hlORDERSIGNCURSOR, szNumOrders, iTotalNumberOfOrders As Integer, bMore As Boolean) As hlERRORS

    Dim hlErr As hlERRORS
    Dim iCount As Integer
    Dim szErrMsg As String
    
    If mvarLoginStatus = hlERROR_NONE Then    'for PR#191

            hlErr = oHl7Constructor.GetOrdersToSign(hlRETURNBYPROVIDER, hlRETURNBATCH, "", Trim$(szNumOrders), lCursor, iTotalNumberOfOrders, bMore)
            
                 'Check timeout errors
            If hlErr <> hlERROR_NONE Then
               If hlErr = hlERROR_TIMED_OUT Or hlErr = hlERROR_NOT_LOGGED_IN Then
                 If ReconnectToCHCS <> hlERROR_NONE Then
                      GoTo ExitFunctionJmp:
                 Else
                       hlErr = oHl7Constructor.GetOrdersToSign(hlRETURNBYPROVIDER, hlRETURNBATCH, "", Trim$(szNumOrders), lCursor, iTotalNumberOfOrders, bMore)

                 End If
              
               Else
                 
                
                If oHl7Constructor.SignError.Count > 0 Then
                    For iCount = 1 To oHl7Constructor.SignError.Count
                        szErrMsg = szErrMsg & oHl7Constructor.SignError(iCount).Text & vbCrLf
                    Next iCount
                
                End If
                If LenB(szErrMsg) = 0 Then
                    MsgBxARMd oHl7Constructor.GetErrorText(hlErr), vbCritical, "CHCS I Error"
                Else
                    MsgBxARMd szErrMsg, vbCritical, "CHCS I Error"
                End If
                
            End If
                
               
            End If

     End If    '2-8-00; AAB PR#191
    
ExitFunctionJmp:
    
               ''SetBusy lngOrderType, False
          GetOrdersToSign = hlErr
    
End Function
Public Function SendPDTSCancel(hlOutcome As hlPDTSOUTCOMES, szReason As String, szHeader As String, szOrderID As String) As hlERRORS

    Dim hlErr As hlERRORS
    Dim iCount As Integer
    Dim szErrMsg As String
    
    If mvarLoginStatus = hlERROR_NONE Then

            hlErr = oHl7Constructor.SendPDTSCancel(hlOutcome, szReason, szHeader, szOrderID)
            
                 'Check timeout errors
            If hlErr <> hlERROR_NONE Then
               If hlErr = hlERROR_TIMED_OUT Or hlErr = hlERROR_NOT_LOGGED_IN Then
                 If ReconnectToCHCS <> hlERROR_NONE Then
                      GoTo ExitFunctionJmp:
                 Else
                       hlErr = oHl7Constructor.SendPDTSCancel(hlOutcome, szReason, szHeader, szOrderID)

                 End If
              
               Else
                 
                
                If oHl7Constructor.SignError.Count > 0 Then
                    For iCount = 1 To oHl7Constructor.SignError.Count
                        szErrMsg = szErrMsg & oHl7Constructor.SignError(iCount).Text & vbCrLf
                    Next iCount
                
                End If
                If LenB(szErrMsg) = 0 Then
                    MsgBxARMd oHl7Constructor.GetErrorText(hlErr), vbCritical, "CHCS I Error"
                Else
                    MsgBxARMd szErrMsg, vbCritical, "CHCS I Error"
                End If
                
            End If
                
               
            End If

     End If
    
ExitFunctionJmp:
    
               
          SendPDTSCancel = hlErr
    
End Function
Public Function SendSignOrders() As hlERRORS

    Dim hlErr As hlERRORS
    Dim iIndex As Integer
    Dim iCount As Integer
    Dim szErrMsg As String
    Dim piReturn As Integer
    Dim bOrderIDProvided As Boolean
    Dim pSign As UnsignedOrder
    Dim bOverrides As Boolean
    Dim bCancels As Boolean
    
    
    goOESession.UnsignedOrdersCollection.bSendingOverrides = False
     If mvarLoginStatus = hlERROR_NONE Then    'for PR#191
            hlErr = oHl7Constructor.SendSignOrders()
                 'Check timeout errors
            If hlErr <> hlERROR_NONE Then
               If hlErr = hlERROR_TIMED_OUT Or hlErr = hlERROR_NOT_LOGGED_IN Then
                 If ReconnectToCHCS <> hlERROR_NONE Then
                      GoTo ExitFunctionJmp:
                 Else
                       hlErr = oHl7Constructor.SendSignOrders()

                 End If
              
               End If
            End If

     End If    '2-8-00; AAB PR#191
    
    
    If oHl7Constructor.SignError.Count > 0 Then
        For iCount = 1 To oHl7Constructor.SignError.Count
            szErrMsg = szErrMsg & oHl7Constructor.SignError(iCount).Text & vbCrLf
        Next iCount

        MsgBxARMd szErrMsg, vbCritical, "CHCS I Errors"
        'GoTo ExitFunctionJmp
    End If
'
'        hlErr = hlERROR_PERFORM_ACTION_FAILED
'    Else
    mvarOrderTypeCode = hlSiGN_ORDERS_CODE
    bOverrides = False
    bCancels = False
    iIndex = 1
    For Each pSign In goOESession.UnsignedOrdersCollection
        If pSign.WarningsProcessed = False And pSign.WarningCollection.Count > 0 Then
            'pSign.WarningsProcessed = True
            goOESession.UnsignedOrdersCollection.WarningIndex = iIndex
            piReturn = GetHl7Warnings(hlSiGN_ORDERS_CODE, bOrderIDProvided)
            If piReturn = True Then
                bOverrides = True
            End If
            If pSign.WarningOverriden = False Then
             '   bOverrides = True
            'Else
                bCancels = True
            End If
                
        End If
        iIndex = iIndex + 1
    Next
    
    If bOverrides Then
        goOESession.UnsignedOrdersCollection.bSendingOverrides = True
        hlErr = oHl7Constructor.SendSignOrders()
        'Check timeout errors
        If hlErr <> hlERROR_NONE Then
            If hlErr = hlERROR_TIMED_OUT Or hlErr = hlERROR_NOT_LOGGED_IN Then
                If ReconnectToCHCS <> hlERROR_NONE Then
                    GoTo ExitFunctionJmp:
                Else
                    hlErr = oHl7Constructor.SendSignOrders()
        
                End If
        
            End If
        End If

    End If
    
    If oHl7Constructor.SignError.Count > 0 Then
        For iCount = 1 To oHl7Constructor.SignError.Count
            szErrMsg = szErrMsg & oHl7Constructor.SignError(iCount).Text & vbCrLf
        Next iCount

        MsgBxARMd szErrMsg, vbCritical, "CHCS I Errors"
    
    End If
    
    If bCancels Then
        'now cancel the orders the users wants to
        For Each pSign In goOESession.UnsignedOrdersCollection
            If pSign.UserHasSigned = True And pSign.WarningOverriden = False And pSign.WarningCollection.Count > 0 Then
            
                'Send cancel message
                
                Dim pOrderTypeCode As hlORDERTYPECODE
                Dim szPatientName As String
                Dim szPatientIEN As String
                pOrderTypeCode = pSign.OrderType
                szPatientName = pSign.PatientInfo.Name
                szPatientIEN = pSign.PatientInfo.IEN
                
                hlErr = goOESession.ReleaseAndSelectPatient(szPatientName, szPatientIEN)
'                If Len(cwobjCWCommSession.CmdLineSwitch("PDTS")) > 0 Then
                
                    If pOrderTypeCode = hlRX_ORDERS_CODE Then
                        hlErr = goOESession.SendPDTSCancel(pSign.WarningPDTSOutcome, pSign.WarningReason, pSign.WarningHeader, pSign.OrderNumber)
                    Else
                        hlErr = goOESession.PerformAction(pOrderTypeCode, hlACTION_DELETE, pSign.OrderNumber)
                    End If
'                Else
'                    hlErr = goOESession.PerformAction(pOrderTypeCode, hlACTION_DELETE, pSign.OrderNumber)
'                End If
                '.PerformAction(pOrderTypeCode, hlACTION_DELETE, pSign.OrderNumber)
                
                'If hlErr = hlERROR_NONE Then
                '    goOESession.UnsignedOrdersCollection(goOESession.UnsignedOrdersCollection.WarningIndex).SigningSuccessful = True
                'Else
                '    goOESession.UnsignedOrdersCollection(goOESession.UnsignedOrdersCollection.WarningIndex).SigningSuccessful = False
                'End If
            
            
            End If
        Next
    End If
ExitFunctionJmp:
    
          SendSignOrders = hlErr
    
End Function


Public Property Get SpecimenChoices() As Collection
    'Collection of all specimens available for the pre-verified lab
    Set SpecimenChoices = oHl7Constructor.Specimen
   
End Property



Friend Function GetExpandedSig(SigText As String, ExpandedSIG As String) As hlERRORS

'This function takes a SIG and expands it to readable english
'It returns any errors from hl7constructor
   
                  '5-10-00; AAB Expanded sig is NOT the function return
                  'Therefore GetExpandedSig() returns an error type
                  '
                  'Link this into PerformAction chain ... 05-10-00; AAB PR#234
                             '' GetExpandedSig = oHl7Constructor.ExpandSig(SigText)
               GetExpandedSig = PerformAction(hlRX_ORDERS_CODE, hlACTION_EXPAND_SIG, SigText)
               ExpandedSIG = oHl7Constructor.RXTablet.ExpandedSIG     'results are in the tablet
  
End Function


Public Property Get SessionError() As hlERRORS

    SessionError = mvarSessionError
    
End Property

Friend Property Let OverrideReason(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.OverrideReason = 5
    mvarszOverrideReason = vData
End Property


Friend Property Get OverrideReason() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: 'Debug.print X.OverrideReason
    OverrideReason = mvarszOverrideReason
End Property

Friend Property Let ConnectFailureText(ByVal vData As String)
'contains text that describes why connectivity to CHCS I failed; either patient or TCP or Access/Verify
 mvarszConnectFailureText = vData
End Property


Public Property Get ConnectFailureText() As String
'contains text that describes why connectivity to CHCS I failed; either patient or TCP or Access/Verify
    ConnectFailureText = mvarszConnectFailureText
End Property

Public Function ReconnectToCHCS() As hlERRORS

    Dim szAccessCode As String, szVerifyCode As String, szCHCSIpAddr As String, szCHCSIpPort As String
    Dim szMTFCode As String
    Dim szPatientIEN As String, szPatientName As String
    Dim szLocalIP As String
    Dim szClinicIEN As String
    Dim szMEPRSCode As String
    Dim szMEPRSIEN As String
    
    'KELLY DEAD 7-27-01 Dim szErr As String
    Dim hlErr As hlERRORS
    
    GetUserInformation szAccessCode, szVerifyCode, szCHCSIpAddr, szCHCSIpPort, szClinicIEN, szMEPRSCode, szMEPRSIEN, szMTFCode
    
    szLocalIP = szGetLocalIPAddress()
               '   oHl7Constructor.szClinicIen = mvarClinicIEN
               '    oHl7Constructor.szMeprsCode = mvarMEPRSCode
               '    oHl7Constructor.szMeprsIEN = mvarMEPRSIen
               '    oHl7Constructor.szMTFCode = szMTFCode
               '
    oLogger.LogtoFile c_mszModuleName & "ReconnectToCHCS", en_Entering

    mvarszConnectFailureText = ""
    
    'ignore the clinic information that GEtUserInformation gives you.  Send down the mvar ones.  The user may have chnaged from default clinic
    '1-7-02 AKM
    If LenB(gobjPatient.MedicalRecordNumber) > 0 Then
        GetPatientInformation szPatientIEN, szPatientName
        hlErr = oHl7Constructor.LoginAndSelectPatient(szAccessCode, szVerifyCode, hlLOGIN.hlUNENCRYPTED, _
               szCHCSIpAddr, szCHCSIpPort, szLocalIP, _
               mvarClinicIEN, mvarMEPRSCode, mvarMEPRSIen, szMTFCode, _
               szPatientIEN, szPatientName)
    Else
        hlErr = oHl7Constructor.Login(szAccessCode, szVerifyCode, hlLOGIN.hlUNENCRYPTED, _
               szCHCSIpAddr, szCHCSIpPort, szLocalIP, _
               mvarClinicIEN, mvarMEPRSCode, mvarMEPRSIen, szMTFCode)
  
    End If
    mvarLoginStatus = hlErr
    If hlErr <> hlERROR_NONE Then
        If hlErr = hlERROR_PATIENT_SELECT_FAILED Then
            mvarszProviderIen = oHl7Constructor.szProviderIEN
            mvarszOrderingProviderIen = oHl7Constructor.szOrderingProviderIEN
            mvarszUserIen = oHl7Constructor.szUserIEN
            mvarszDefOrderPage = oHl7Constructor.szDefOrderPage
        End If
        GenerateLoginFailureMessage hlErr
    Else
        szPatientName = oHl7Constructor.goPatient.Name
        Set mvarRxTablet = oHl7Constructor.RXTablet
        Set mvarRadTablet = oHl7Constructor.RadTablet
        Set mvarLabTablet = oHl7Constructor.LabTablet
        Set mvarOtherTablet = oHl7Constructor.OtherTablet
        
        mvarszProviderIen = oHl7Constructor.szProviderIEN
        mvarszOrderingProviderIen = oHl7Constructor.szOrderingProviderIEN
  
        mvarszUserIen = oHl7Constructor.szUserIEN
        mvarszDefOrderPage = oHl7Constructor.szDefOrderPage
        
        CheckIfUserIsProvider
        
    End If
    ReconnectToCHCS = hlErr
    
    CheckIfUserIsProvider
    
    oLogger.LogtoFile c_mszModuleName & "ReconnectToCHCS", en_Exiting, "hlErr", hlErr, "szPatientIEN", szPatientIEN

End Function



Public Property Get OrderTypeCode() As hlORDERTYPECODE
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: 'Debug.print X.OrderTypeCode
    OrderTypeCode = mvarOrderTypeCode
End Property




Public Property Get LabTablet() As LabType
    Set LabTablet = mvarLabTablet
End Property
Public Property Get OtherTablet() As BaseTablet
    Set OtherTablet = mvarOtherTablet
End Property

Public Property Get RadTablet() As RadType
    Set RadTablet = mvarRadTablet
End Property

Public Property Get RXTablet() As RxType
    Set RXTablet = mvarRxTablet
End Property
Public Property Set OtherTablet(ByRef newOther As BaseTablet)
    Set oHl7Constructor.OtherTablet = newOther
    Set mvarOtherTablet = oHl7Constructor.OtherTablet
End Property
Public Property Set LabTablet(ByRef newLab As LabType)
    Set oHl7Constructor.LabTablet = newLab
    Set mvarLabTablet = oHl7Constructor.LabTablet
End Property

Public Property Set RadTablet(ByRef newRad As RadType)
    Set oHl7Constructor.RadTablet = newRad
    Set mvarRadTablet = oHl7Constructor.RadTablet
End Property
Public Property Set RXTablet(ByRef newRx As RxType)
    Set oHl7Constructor.RXTablet = newRx
    Set mvarRxTablet = oHl7Constructor.RXTablet
End Property



Public Function PerformAction(OrderType As hlORDERTYPECODE, OrderAction As hlACTION, Optional orderid As String) As hlERRORS
    
Dim pszTempA As String
Dim piReturn As Integer
Dim piOneRetry As Integer                              '6-4-01; AAB #11978

Dim lReturnErr As hlERRORS
Dim szErrMsg As String
Dim iErrCount As Integer
Dim WorkingTablet As Object
Dim bErr50MoreWarnings As Boolean
Dim bOrderIDProvided As Boolean     ' TRUE if after sending a new order message, an orderid is provided.
                                                         '  Should then send and override message
          
          mvarOrderTypeCode = OrderType
    
          'SCR 21227 4-10-02 latest is take it out
'          If Not goAnP Is Nothing Then
'            If Not goAnP.IsEncounterLocked Then
'                goAnP.DisplayEncounterLockMessage
'                lReturnErr = hlERROR_ANP_ENCOUNTER_NOT_LOCKED
'                GoTo ExitFunctionJmp
'            End If
'          End If
          
           '' Debug.Print "Start of Perform Action " & OrderType

          piOneRetry = False

          Select Case OrderType
          Case hlRX_ORDERS_CODE:
                    Set WorkingTablet = oHl7Constructor.RXTablet
          Case hlRAD_ORDERS_CODE:
                    Set WorkingTablet = oHl7Constructor.RadTablet
          Case hlLAB_ORDERS_CODE:
                    Set WorkingTablet = oHl7Constructor.LabTablet
          End Select
          
          gbBackgroundBusy = True
          
          '=======================================================================================
          Select Case OrderAction
          Case hlACTION_RENEW_MODIFY, hlACTION_MED_LIST_RENEW:
                    OrderAction = hlACTION_RENEW
                    lReturnErr = oHl7Constructor.PerformAction(orderid, OrderAction, OrderType)
          Case hlACTION_MED_LIST_MODIFY:
                    OrderAction = hlACTION_MODIFY
                    lReturnErr = oHl7Constructor.PerformAction(orderid, OrderAction, OrderType)
          Case hlACTION_ORDER_SET:
                    OrderAction = hlACTION_NEW
                    lReturnErr = oHl7Constructor.SendOrder(OrderType)
          Case hlACTION_DETAIL, hlACTION_HOLD, hlACTION_MODIFY, hlACTION_RELEASE, hlACTION_DELETE, hlACTION_RENEW:
                    lReturnErr = oHl7Constructor.PerformAction(orderid, OrderAction, OrderType)
          Case hlACTION_NEW:
                    lReturnErr = oHl7Constructor.SendOrder(OrderType)
          Case hlACTION_PREMODIFY:
                    lReturnErr = oHl7Constructor.PreModify(orderid, OrderType)
          Case hlACTION_PRESET:
                    lReturnErr = oHl7Constructor.PreSet(OrderType)
          Case hlACTION_EXPAND_SIG     'added 05-10-00; AAB PR#234
                    lReturnErr = oHl7Constructor.ExpandSig(orderid)      'using 'orderid' as string register
          Case hlACTION_ADD_TO_QUEUE:
                    lReturnErr = oHl7Constructor.Sav2Q(OrderType)
          Case hlACTION_RESET_ORDERING_PROVIDER:
          'Reset the ordering provider
                    lReturnErr = oHl7Constructor.Sav2Q(OrderType)
          Case Else
                    lReturnErr = oHl7Constructor.PreVerify(OrderType)
          End Select

          '======================================================================================
          
          gbBackgroundBusy = False
          
          
'6-4-01; AAB #11978
Retry:

     'Check timeout errors
     'NOTE: "lReturnErr" is NOT the (HL7-returned) error 'Collection'; it is a return-status error for the ACTION
     
          If lReturnErr <> hlERROR_NONE Then
                    If lReturnErr = hlERROR_TIMED_OUT Then
                              If ReconnectToCHCS <> hlERROR_NONE Then
                                        ''MsgBxARMd oHl7Constructor.GetErrorText(lReturnErr), vbCritical, "CHCS I Error"
                                        GoTo ExitFunctionJmp:
                              Else
                                   'Re-send the message
                                        gbBackgroundBusy = True
                                        Select Case OrderAction
                                        Case hlACTION_DETAIL, hlACTION_HOLD, hlACTION_MODIFY, hlACTION_RELEASE, hlACTION_DELETE, hlACTION_RENEW:
                                            lReturnErr = oHl7Constructor.PerformAction(orderid, OrderAction, OrderType)
                                        Case hlACTION_NEW:
                                            lReturnErr = oHl7Constructor.SendOrder(OrderType)
                                        Case hlACTION_PREVERIFY      'Fixed 9-22-99; AAB: was triggered via ELSE;
                                            lReturnErr = oHl7Constructor.PreVerify(OrderType)
                                        Case hlACTION_PREMODIFY:
                                            lReturnErr = oHl7Constructor.PreModify(orderid, OrderType)
                                        Case hlACTION_PRESET:
                                            lReturnErr = oHl7Constructor.PreSet(OrderType)
                                        Case hlACTION_EXPAND_SIG     'added 05-10-00; AAB PR#234
                                             lReturnErr = oHl7Constructor.ExpandSig(orderid)        'ExpandSig(SigText)
                                        Case Else
                                            lReturnErr = hlERROR_BAD_ACTION_CODE    '9-22-99; AAB: case ELSE triggers an error
                                            GoTo ExitFunctionJmp:
                                        End Select
                                        gbBackgroundBusy = False
                              End If
                         
                    Else      'All other non-timeout errors; NOT HL7 errors
                    
                               If OrderAction = hlACTION_PREVERIFY Or OrderAction = hlACTION_NEW Then
                                   ' If OrderAction = hlACTION_PREVERIFY Then
                                   'An ACTION failure above could be a missing parameter; 'PR#221 allow to recover and retry
                                   
                                        szErrMsg = oHl7Constructor.GetErrorText(lReturnErr)

                                         If OrderAction = hlACTION_NEW Then
                                                  piReturn = IsError_Soft(OrderType, lReturnErr)
                                                  If piReturn = YES Then  'have a SOFT (recoverable) error
                                                            MsgBxARMd szErrMsg, vbExclamation, "Incomplete or Incorrect Data"        ''PR#221; uses "ErrorIsMessage" for steering
                                                            AskCoreToRefreshAfterModal OrderType
                                                            '' GoTo ExitFunctionJmp:
                                                  Else
                                                            MsgBxARMd szErrMsg, vbCritical, "CHCS I Errors"
                                                            AskCoreToRefreshAfterModal OrderType
                                                            lReturnErr = hlERROR_PERFORM_ACTION_FAILED       'added 9-2-99; AAB
                                                                 ''GoTo ExitFunctionJmp:
                                                  End If
                                        Else
                                        
                                                    szErrMsg = oHl7Constructor.GetErrorText(lReturnErr)
                                                            ''If WorkingTablet.ErrorIsMessage = False Then                      'added 9-2-99; AAB
                                                   MsgBxARMd szErrMsg, vbCritical, "CHCS I Errors"
                                                   AskCoreToRefreshAfterModal OrderType
                                                   lReturnErr = hlERROR_PERFORM_ACTION_FAILED       'added 9-2-99; AAB
                                                            ''GoTo ExitFunctionJmp:
                                        End If
                                        
                                      '===========================================
                                      
                                        GoTo ExitFunctionJmp:
                                       
''                                         Else
''                                              MsgBxARMd szErrMsg, vbExclamation, "Incomplete or Incorrect Data"        ''PR#221; uses "ErrorIsMessage" for steering
''                                              'NOTE: keep 'lReturnErr' intact so caller knows we were not successful
''                                         End If
                                                  ''MsgBxARMd oHl7Constructor.GetErrorText(lReturnErr), vbCritical
                                                  
                                       '' GoTo ExitFunctionJmp:
                                       
                                End If
                               
                     End If
                     
          End If
    
          If OrderAction = hlACTION_EXPAND_SIG Then
               lReturnErr = hlERROR_NONE
                GoTo ExitFunctionJmp:
          End If
          
          '======================================================================================

          Select Case OrderType
          Case hlRX_ORDERS_CODE:
               Set mvarRxTablet = oHl7Constructor.RXTablet
               Set WorkingTablet = oHl7Constructor.RXTablet
          Case hlRAD_ORDERS_CODE:
               Set mvarRadTablet = oHl7Constructor.RadTablet
               Set WorkingTablet = oHl7Constructor.RadTablet
          Case hlLAB_ORDERS_CODE:
               Set mvarLabTablet = oHl7Constructor.LabTablet
               Set WorkingTablet = oHl7Constructor.LabTablet
          Case Else
               Set mvarOtherTablet = oHl7Constructor.OtherTablet
               Set WorkingTablet = oHl7Constructor.OtherTablet
          
          End Select

                         
     '======== HL7 errors from ERR segment ============================================
                        
                         '6-4-01; AAB #11978
                         bErr50MoreWarnings = False
                          szErrMsg = ""
            If OrderAction = hlACTION_ADD_TO_QUEUE Then
                   lReturnErr = hlERROR_NONE
                            ''Select Case OrderType
                            ''     Case hlRX_ORDERS_CODE:
                            ''         oHl7Constructor.RXTablet.orderid = "Q" & Format(Now, "yyyymmddhhmmss")
                            ''
                            ''     Case hlRAD_ORDERS_CODE:
                            ''         oHl7Constructor.RadTablet.orderid = "Q" & Format(Now, "yyyymmddhhmmss")
                            ''
                            ''     Case hlLAB_ORDERS_CODE:
                            ''         oHl7Constructor.LabTablet.orderid = "Q" & Format(Now, "yyyymmddhhmmss")
                            ''
                            '' End Select
                            ''WorkingTablet.orderid = Format(Now, "yyyymmddhhmmss")
                  GoTo JumpToAddToColl:
            End If
          
        '=====================================================================================
     'Check HL7 errors; NOTE error 50 fixes are a temporary work-around pending fix to mumps code
          If GetHl7Errors(OrderType, szErrMsg, iErrCount) Then
                    bErr50MoreWarnings = bAnyErr50(OrderType)
                    If WorkingTablet.ErrorIsMessage = False Then                          'added 9-2-99; AAB
                              If Not bErr50MoreWarnings Then                                    'ERROR 50 is always special case; added 9-13-99 AKM
                                   '04-04-00; AAB for PR#221; allow user to EDIT or CLEAR a hard error MsgBxARMd
                                   '04-05-00; Control as-best-as-possible when we allow choice
                                        piReturn = IsError_Recoverable(OrderType, OrderAction)
                                        If piReturn = giERRORS_ARE_CHOICE Then       'YES Then
                                                  pszTempA = vbCrLf & vbCrLf & "Enter <YES> to CORRECT order, or <NO> to CLEAR order"   ''& & vbCrLf
                                                  piReturn = MsgBxARMd(szErrMsg & pszTempA, vbExclamation + vbYesNo + vbDefaultButton1, "CHCS I Errors")
                                                  If piReturn = vbYes Then
                                                        lReturnErr = hlERROR_SOFT_ERROR                      'caller can tell if this allows the user to correct and retry
                                                   Else
                                                        lReturnErr = hlERROR_PERFORM_ACTION_FAILED       'added 9-2-99; AAB
                                                   End If
                                                        ''GoTo ExitFunctionJmp:
                                                  AskCoreToRefreshAfterModal OrderType
                                                  
                                        ElseIf piReturn = giERRORS_ARE_HARD Then
                                             MsgBxARMd szErrMsg, vbCritical, "CHCS I Errors"
                                             AskCoreToRefreshAfterModal OrderType
                                             lReturnErr = hlERROR_PERFORM_ACTION_FAILED
                                                    ''GoTo ExitFunctionJmp:
                                        Else      'giERRORS_ARE_SOFT
                                                    ''MsgBxARMd szErrMsg, vbInformation, "CHCS I Errors"
                                             MsgBxARMd szErrMsg, vbExclamation, "Please correct and re-submit"
                                             AskCoreToRefreshAfterModal OrderType
                                             lReturnErr = hlERROR_SOFT_ERROR
                                                    '' GoTo ExitFunctionJmp:     '04-06-00; AAB let SOFT result run through rest of this code stream
                                        End If
                                        
                                         GoTo ExitFunctionJmp:
                                        '================================
                              Else
                                        'ERROR-50 Detected:   '6-01-01; AAB #11978
                                      If GetHl7Warnings(OrderType, bOrderIDProvided) Then
                                        If HandleSecondaryWarnings(OrderType, OrderAction, WorkingTablet, bOrderIDProvided, orderid, lReturnErr) Then
                                            GoTo ExitFunctionJmp:
                                        Else
                                            lReturnErr = hlERROR_NONE
                                            
                                        End If
                                     End If
'                                            If bOrderIDProvided = True And OrderAction = hlACTION_NEW Then
'                                                   lReturnErr = oHl7Constructor.SendOverrides(OrderType)
'                                                   Call ShowCHCSErrors(WorkingTablet, OrderType, False, bErr50MoreWarnings, lReturnErr)
'                                                   If lReturnErr = hlERROR_PERFORM_ACTION_FAILED Then
'                                                        GoTo ExitFunctionJmp:
'                                                   End If
'
'
'                                              'SECONDARY warnings disguised as error 50
'                                             ElseIf ((OrderAction = hlACTION_NEW) Or (OrderAction = hlACTION_MODIFY) Or (OrderAction = hlACTION_RELEASE) Or (OrderAction = hlACTION_RENEW)) And (piOneRetry = False) Then
'                                                      piReturn = GetHl7Warnings(OrderType, bOrderIDProvided)
'                                                      If LenB(WorkingTablet.WarningReason) = 0 Then
'                                                           'removed 8-6-01; AAB # 14384
'                                                           ''Call ShowCHCSErrors(WorkingTablet, OrderType, False, bErr50MoreWarnings, lReturnErr)
'                                                           lReturnErr = hlERROR_PERFORM_ACTION_FAILED
'                                                           GoTo ExitFunctionJmp:
'                                                      Else
'
'                                                           gbBackgroundBusy = True
'
'                                                           Select Case OrderAction
'                                                           Case hlACTION_NEW
'                                                                lReturnErr = oHl7Constructor.SendOrder(OrderType)      'RETRY
'                                                           'Case hlACTION_MODIFY etc
'                                                           Case Else
'                                                                lReturnErr = oHl7Constructor.PerformAction(orderid, OrderAction, OrderType)
'                                                           End Select
'
'                                                           gbBackgroundBusy = False
'                                                           piOneRetry = True                  'only allow one retry
'                                                           GoTo Retry:
'                                                      End If
'
'                                            Else
'                                                      Call ShowCHCSErrors(WorkingTablet, OrderType, False, bErr50MoreWarnings, lReturnErr)
'                                                      lReturnErr = hlERROR_PERFORM_ACTION_FAILED
'                                                      GoTo ExitFunctionJmp:
'                                            End If
'
'        '============================================  work-around for error-50         '6-4-01; AAB #11978
'
                              End If
                              
                    Else      'Error is REALLY a MESSAGE ===============================================
                              If Not OrderAction = hlACTION_PREMODIFY Or OrderAction = hlACTION_PRESET Then      '03-31-00; AAB: PR#219(1)
                                        MsgBxARMd szErrMsg, vbInformation, "CHCS I Messages"
                                        AskCoreToRefreshAfterModal OrderType
                                        lReturnErr = hlERROR_NONE
                              End If
                              
                    End If
                              ''GoTo ExitFunctionJmp: 'removed AKM 9/13/99
          End If
    
          '======================================================================================
     
          If OrderAction = hlACTION_PREVERIFY Or OrderAction = hlACTION_PRESET Then                  '9-2-99; AAB
                    If OrderType = hlRX_ORDERS_CODE Then    'RXO(16) only for RX pre-verify; AAB 9-15-99
                           'Added 9-15-99; AAB
                                If oHl7Constructor.RXTablet.DrugIsValid = False Then
                                    Dim szMsgBoxCaption As String
                                    Dim szMsgBoxText As String
                                    
                                    szMsgBoxCaption = "Invalid Medication Code: " & oHl7Constructor.RXTablet.OrderName       'szOrderName
                                    szMsgBoxText = "CHCS I cannot preverify this selection"
                                    MsgBxARMd szMsgBoxText, vbInformation, szMsgBoxCaption
                                    AskCoreToRefreshAfterModal OrderType
                                              ' lReturnErr = $$$ need new error code for RXO-16
                                               '$$$ III need dedicated error message for this + others
                                       lReturnErr = hlERROR_PREVERIFY_RX_FAILED
                                    GoTo ExitFunctionJmp:
                                End If
                    End If
                    
                    If bCheckOrderNamesOK(OrderType) = False Then
                            lReturnErr = hlERROR_PREVERIFY_ORDERNAME_DONT_MATCH
                            GoTo ExitFunctionJmp:
                    End If
         End If
    
          '======================================================================================
    
          If GetHl7Warnings(OrderType, bOrderIDProvided) Then    'is TRUE if Warnings exist, even if NOT endorsed
                If HandleSecondaryWarnings(OrderType, OrderAction, WorkingTablet, bOrderIDProvided, orderid, lReturnErr) Then
'                    If OrderType = hlRX_ORDERS_CODE Then
'                        If WorkingTablet.WarningPDTSOutcome = hlPDTSOutcomes_DiffQty Then
'                            lReturnErr = hlERROR_USER_CANCELLED_NEWQTY
'                        ElseIf WorkingTablet.WarningPDTSOutcome = hlPDTSOutcomes_DiffDirections Then
'                            lReturnErr = hlERROR_USER_CANCELLED_NEWSIG
'                        ElseIf WorkingTablet.WarningPDTSOutcome = hlPDTSOutcomes_DiffDose Then
'                            lReturnErr = hlERROR_USER_CANCELLED_NEWSIG
'                        Else
'                            lReturnErr = hlERROR_USER_CANCELLED
'                        End If
'                    Else
'                        lReturnErr = hlERROR_USER_CANCELLED
'
'                    End If
                   
                   GoTo ExitFunctionJmp:
                End If

               'NOTE: drop through if warnings were-not required
               'If WorkingTablet.WarningReason <> "" And OrderAction <> hlACTION_PREVERIFY Then
'                    If LenB(WorkingTablet.WarningReason) > 0 Then     'Indicates endorsement (Reason for Overriding Warnings)
'                              If bOrderIDProvided = True And OrderAction = hlACTION_NEW Then
'                                        ''#11978; AAB
'                                            lReturnErr = oHl7Constructor.SendOverrides(OrderType)
'                                           Call ShowCHCSErrors(WorkingTablet, OrderType, False, bErr50MoreWarnings, lReturnErr)
'                                            If lReturnErr = hlERROR_PERFORM_ACTION_FAILED Then
'                                                GoTo ExitFunctionJmp:
'                                           End If
'                              Else      'No Order-ID exists at this point ... proceed
'                                     If OrderAction = hlACTION_NEW Then
'                                             lReturnErr = oHl7Constructor.SendOrder(OrderType)
'                                             Call ShowCHCSErrors(WorkingTablet, OrderType, False, bErr50MoreWarnings, lReturnErr)
'                                             If lReturnErr = hlERROR_PERFORM_ACTION_FAILED Then
'                                                  GoTo ExitFunctionJmp:
'                                             End If
'                                     ElseIf OrderAction = hlACTION_PREVERIFY Or OrderAction = hlACTION_PRESET Then
'                                             ''No Action required ... Supplying a reason means we can proceed smoothly ...
'                                     Else
'                                             lReturnErr = oHl7Constructor.PerformAction(orderid, OrderAction, OrderType)
'                                             Call ShowCHCSErrors(WorkingTablet, OrderType, False, bErr50MoreWarnings, lReturnErr)
'                                     End If
'                              End If
'
'                              '===================================================================
'                              If lReturnErr <> hlERROR_NONE Then
'                                       If lReturnErr = hlERROR_TIMED_OUT Then
'                                           If ReconnectToCHCS <> hlERROR_NONE Then
'                                               'MsgBxARMd oHl7Constructor.GetErrorText(lReturnErr), vbCritical, "CHCS I Error"
'                                               GoTo ExitFunctionJmp:
'                                           End If
'                                       Else
'                                           MsgBxARMd oHl7Constructor.GetErrorText(lReturnErr), vbCritical
'                                           GoTo ExitFunctionJmp:
'                                       End If
'                              End If
'
'               'PR#127, PR#134: 02-18-00: AAB
'                              'ElseIf OrderAction = hlACTION_PREVERIFY Then    'blank but required means user aborted
'               '=================================================================================================
'               'Added 9-3-99; AAB see also bug fix in GetWarnings() and add-on to Warnings form
'               'Next 11 lines were deactivated by AKM; re-activated 02-18-00; AAB PR#127 PR#134
'               '
'               'No Reason entered ... These warnings were NOT endorsed ... Cannot send order or otherwise proceed
'                     ElseIf gbOverrideWarnings = False Then  'FALSE means we CANCELLED from the warnings screen
'
'                            If bOrderIDProvided = False Then
'                                    Select Case OrderType
'                                    Case hlRX_ORDERS_CODE:
'                                       lReturnErr = hlERROR_RX_REASON
'                                    Case hlRAD_ORDERS_CODE:
'                                        lReturnErr = hlERROR_RAD_REASON
'                                    Case hlLAB_ORDERS_CODE:
'                                        lReturnErr = hlERROR_LAB_REASON
'                                    End Select
'
'                                    GoTo ExitFunctionJmp:
'                               'end of PR#127 PR#134 re-activated section
'                            End If
'
'                     Else
'                            GoTo ExitFunctionJmp:
'                     End If
'
          End If      'IF warnings exist ...
        
          '=========================================================================
JumpToAddToColl:
                If IsActionViaSubmitButton(OrderType) = False Then     '05-30-00; AAB PR#257
              
                         Select Case OrderAction
                         Case hlACTION_NEW:
                         
                            'Move to individual callers to be as late as possible; 5-24-00; AAB; PR#257
                            'KEEP case statement as 'place-holder' to avoid case-else flow
                                  
                            'SendMessageToCoreAndAddTransaction OrderType, OrderAction
                            '05-30-00; AAB new control design over this logic
                              
                                        '' If IsActionViaSubmitButton(OrderType) = False Then
                                SendMessageToCoreAndAddTransaction OrderType, OrderAction
                                        '' End If
               
                         Case hlACTION_MODIFY, hlACTION_RELEASE, hlACTION_RENEW, hlACTION_HOLD:
                              SendMessageToCoreAndAddTransaction OrderType, OrderAction, orderid
                         Case hlACTION_DELETE:
                              SendMessageToCoreAndAddTransaction OrderType, OrderAction, orderid
                              RemoveFromOrderList OrderType, orderid
                              
                              
                         Case hlACTION_PREMODIFY, hlACTION_PRESET, hlACTION_PREVERIFY, hlACTION_DETAIL, hlACTION_EXPAND_SIG: 'ignore these actions
                              ''IGNORED CASE
                              Debug.Print    'need a hook for break-point
                         
                         Case Else
                              SendMessageToCoreAndAddTransaction OrderType, OrderAction
                         End Select
                         
               End If

ExitFunctionJmp:

               PerformAction = lReturnErr           '09-2-99; AAB
               mvarSessionError = lReturnErr      '09-2-99; AAB  'fixed 05-30-00; AAB was "lReturnErr=lReturnErr"
               
               gbBackgroundBusy = False
               '===========================================================================================
               
               'NEW: 5-30-00; AAB
               If (lReturnErr <> hlERROR_NONE) And (lReturnErr <> hlERROR_SOFT_ERROR) Then

                        Select Case OrderAction
                        Case hlACTION_NEW, hlACTION_MODIFY, hlACTION_RENEW:
                             '12-18-01 VB modal errors with Ordersets AKM
                                 If OrderTypeCode = hlLAB_ORDERS_CODE And goOESession.LabTablet.Subscript = hlAUTOPSY Then
                                         
                                         Unload Frm_Autopsy_AP
                                         
                                 ElseIf OrderTypeCode = hlLAB_ORDERS_CODE And _
                                        (goOESession.LabTablet.Subscript = hlCYTNONGYN Or _
                                         goOESession.LabTablet.Subscript = hlSURGPATH Or _
                                         goOESession.LabTablet.Subscript = hlBONEMARROW) Then
                                         
                                         Unload Frm_CytoNonGyn
                                        
                                 ElseIf OrderTypeCode = hlLAB_ORDERS_CODE And (goOESession.LabTablet.Subscript = hlCYTGYN) Then
                                        
                                        Unload Frm_CytoGyn
        
                                 End If
                             
                                SendErrorMessageToCore OrderType, OrderAction, lReturnErr
                        End Select
                        
               End If
               
 
End Function





'Public Property Set oeSearchEngine(ByVal vData As OE_SearchEngine)
''used when assigning an Object to the property, on the left side of a Set statement.
''Syntax: Set x.oeSearchEngine = Form1
''    Set mvaroeSearchEngine = vData
'End Property


Public Property Get oeSearchEngine() As OE_SearchEngine
    Set oeSearchEngine = moeSearchEngine
End Property

Public Function CommonPatientSelected() As hlERRORS

    Dim szPatientName As String
    Dim szPatientIEN As String

    
    GetPatientInformation szPatientIEN, szPatientName
    oLogger.LogtoFile c_mszModuleName & "CommonPatientSelected", en_Entering, "PatientIEN", szPatientIEN
    
    'Tracker 20191 clear the transaction logs
    If Len(gobjShared.CmdLineSwitches("GREEN")) > 0 Then
        Set goRxTransactionCol = Nothing
        Set goLabTransactionCol = Nothing
        Set goRadTransactionCol = Nothing
    
        Set goRxTransactionCol = New RxTransactionCollection
        Set goLabTransactionCol = New LabTransactionCollection
        Set goRadTransactionCol = New RadTransactionCollection
    End If

    If oHl7Constructor.goPatient.IEN <> szPatientIEN Then
        goOESession.ReleaseAndSelectPatient
        Frm_OE_RX.ClearRx HARD_RESET
        RefreshFormCaption hlRX_ORDERS_CODE, Frm_OE_RX, hlACTION_NEW
        
        Frm_OE_Lab.ClearLab HARD_RESET
        RefreshFormCaption hlLAB_ORDERS_CODE, Frm_OE_Lab, hlACTION_NEW
        
        Frm_OE_Rad.ClearRad HARD_RESET
        RefreshFormCaption hlRAD_ORDERS_CODE, Frm_OE_Rad, hlACTION_NEW
        
        RefreshEachUIsOrdersOnPage
    End If
    
     CommonPatientSelected = hlERROR_NONE
     oLogger.LogtoFile c_mszModuleName & "CommonPatientSelected", en_Exiting
   
End Function


Friend Function CreateHL7Session() As Boolean

Dim szAccessCode As String, szVerifyCode As String, szCHCSIpAddr As String, szCHCSIpPort As String
Dim szMTFCode As String
Dim szPatientIEN As String, szPatientName As String
Dim szLocalIP As String
Dim oCreateHl7Constructor As CreateHl7Constructor
'KELLY DEAD 7-27-01 Dim szErr As String
Dim hlErr As hlERRORS

    oLogger.LogtoFile c_mszModuleName & "CreateHL7Session", en_Entering
'========================================================
'Fast-Exit to prevent multi-instancing:
    
          If Not (oHl7Constructor Is Nothing) Then
                   'RefreshEachUIsOrdersOnPage
               CreateHL7Session = True
                   'bHL7SessionCreated = True
               oLogger.LogtoFile c_mszModuleName & "CreateHL7Session - Already created", en_Exiting, "hlErr", hlErr, "szPatientIEN", szPatientIEN
               
               
               Exit Function
          End If
    
'========================================================
'NORMAL ENTRY:
          gbBackgroundBusy = True
          
          bHL7SessionCreated = True
          oLogger.LogtoFile c_mszModuleName & "CreateHL7Session - Creating Search Engine", en_Entering, "hlErr", hlErr

          Set moeSearchEngine = New OE_SearchEngine
          oLogger.LogtoFile c_mszModuleName & "CreateHL7Session - Creating Search Engine", en_Exiting, "hlErr", hlErr
       
'' Designed to load on first click event  ... 7-11-01 AAB
''Moved to SetRequestingLocation in Hl7Common
''
''            'AKM added 6/27/01 to allow for 1 central location of clinic info
''          moeSearchEngine.LoadAllMeprsIntoCollection gcolClinicColl
                
    'Create a create-hl7constructor object only long enough to get a hl7constructor object amw 8/20/99
        oLogger.LogtoFile c_mszModuleName & "CreateHL7Session - CreateHl7Constructor", en_Entering, "hlErr", hlErr
          Set oCreateHl7Constructor = New CreateHl7Constructor
          Set oHl7Constructor = oCreateHl7Constructor.GetHl7Constructor
          Set oCreateHl7Constructor = Nothing
            oLogger.LogtoFile c_mszModuleName & "CreateHL7Session - CreateHl7Constructor", en_Exiting, "hlErr", hlErr

        oLogger.LogtoFile c_mszModuleName & "CreateHL7Session - GetUserInformation", en_Entering, "hlErr", hlErr
        GetUserInformation szAccessCode, szVerifyCode, szCHCSIpAddr, szCHCSIpPort, mvarClinicIEN, mvarMEPRSCode, mvarMEPRSIen, szMTFCode
        szLocalIP = szGetLocalIPAddress()
        oLogger.LogtoFile c_mszModuleName & "CreateHL7Session - GetUserInformation", en_Exiting, "hlErr", hlErr
            
        If LenB(gobjPatient.MedicalRecordNumber) > 0 Then
            oLogger.LogtoFile c_mszModuleName & "CreateHL7Session - GetPatientInformation", en_Entering, "hlErr", hlErr
            GetPatientInformation szPatientIEN, szPatientName
            oLogger.LogtoFile c_mszModuleName & "CreateHL7Session - GetPatientInformation", en_Exiting, "hlErr", hlErr
            
            oLogger.LogtoFile c_mszModuleName & "CreateHL7Session - LoginAndSelectPatient", en_Entering, "hlErr", hlErr
            hlErr = oHl7Constructor.LoginAndSelectPatient(szAccessCode, szVerifyCode, hlLOGIN.hlUNENCRYPTED, _
                 szCHCSIpAddr, szCHCSIpPort, szLocalIP, _
                 mvarClinicIEN, mvarMEPRSCode, mvarMEPRSIen, szMTFCode, _
                 szPatientIEN, szPatientName)
            oLogger.LogtoFile c_mszModuleName & "CreateHL7Session - LoginAndSelectPatient", en_Exiting, "hlErr", hlErr
        Else
            If mbCalledFromTabletUIs = True Then
                hlErr = hlERROR_PATIENT_SELECT_FAILED
            Else
                
                oLogger.LogtoFile c_mszModuleName & "CreateHL7Session - Login", en_Entering, "hlErr", hlErr
                hlErr = oHl7Constructor.Login(szAccessCode, szVerifyCode, hlLOGIN.hlUNENCRYPTED, _
                   szCHCSIpAddr, szCHCSIpPort, szLocalIP, _
                     mvarClinicIEN, mvarMEPRSCode, mvarMEPRSIen, szMTFCode)
                oLogger.LogtoFile c_mszModuleName & "CreateHL7Session - Login", en_Exiting, "hlErr", hlErr
            End If
             
        End If
        mvarLoginStatus = hlErr    '2-8-00; AAB PR#191
        mvarSessionError = hlErr   '3-8-01; AAB: INC2(10073)
            
          If hlErr <> hlERROR_NONE Then
                If hlErr = hlERROR_PATIENT_SELECT_FAILED Then
                    mvarszProviderIen = oHl7Constructor.szProviderIEN
                    mvarszOrderingProviderIen = oHl7Constructor.szOrderingProviderIEN
  
                    mvarszUserIen = oHl7Constructor.szUserIEN
                    mvarszDefOrderPage = oHl7Constructor.szDefOrderPage
                End If
 
                GenerateLoginFailureMessage hlErr
     
                'szErr = oHl7Constructor.GetErrorText(hlErr)
                CreateHL7Session = False
                bHL7SessionCreated = True
                oLogger.LogtoFile c_mszModuleName & "CreateHL7Session - FAILED2", en_Exiting, "hlErr", hlErr, "szPatientIEN", szPatientIEN
                
                gbBackgroundBusy = False
                Exit Function
    
          End If
          
          CheckIfUserIsProvider
          
    'no errors ... proceed
          szPatientName = oHl7Constructor.goPatient.Name
          Set mvarRxTablet = oHl7Constructor.RXTablet
          Set mvarRadTablet = oHl7Constructor.RadTablet
          Set mvarLabTablet = oHl7Constructor.LabTablet
          Set mvarOtherTablet = oHl7Constructor.OtherTablet
         
         'REDUNDANT: Removed 02-26-00; AAB
               ' RefreshEachUIsOrdersOnPage
         
          CreateHL7Session = True

    'BUG Fix: 02-22-2000; AAB for CT enhancement: failed to occupy new property registers
          mvarszProviderIen = oHl7Constructor.szProviderIEN
          mvarszOrderingProviderIen = oHl7Constructor.szOrderingProviderIEN
          mvarszUserIen = oHl7Constructor.szUserIEN
          mvarszDefOrderPage = oHl7Constructor.szDefOrderPage
          oLogger.LogtoFile c_mszModuleName & "CreateHL7Session", en_Exiting, "hlErr", hlErr, "szPatientIEN", szPatientIEN

          gbBackgroundBusy = False
End Function
'Friend Function CreateHL7SessionForSigning() As Boolean
'
'Dim szAccessCode As String, szVerifyCode As String, szCHCSIpAddr As String, szCHCSIpPort As String
'Dim szMTFCode As String
'Dim szPatientIEN As String, szPatientName As String
'Dim szLocalIP As String
'Dim oCreateHl7Constructor As CreateHl7Constructor
''KELLY DEAD 7-27-01 Dim szErr As String
'Dim hlErr As hlERRORS
'
'    oLogger.LogtoFile c_mszModuleName & "CreateHL7Session", en_Entering
''========================================================
''Fast-Exit to prevent multi-instancing:
'
'          If Not (oHl7Constructor Is Nothing) Then
'                   'RefreshEachUIsOrdersOnPage
'               CreateHL7SessionForSigning = True
'                   'bHL7SessionCreated = True
'
'               oLogger.LogtoFile c_mszModuleName & "CreateHL7SessionForSigning - Already created", en_Exiting, "hlErr", hlErr, "szPatientIEN", szPatientIEN
'               Exit Function
'          End If
'
''========================================================
''NORMAL ENTRY:
'
'          bHL7SessionCreated = True
'          oLogger.LogtoFile c_mszModuleName & "CreateHL7SessionForSigning - Creating Search Engine", en_Entering, "hlErr", hlErr
'
'          Set moeSearchEngine = New OE_SearchEngine
'          oLogger.LogtoFile c_mszModuleName & "CreateHL7SessionForSigning - Creating Search Engine", en_Exiting, "hlErr", hlErr
'
''' Designed to load on first click event  ... 7-11-01 AAB
'''Moved to SetRequestingLocation in Hl7Common
'''
'''            'AKM added 6/27/01 to allow for 1 central location of clinic info
'''          moeSearchEngine.LoadAllMeprsIntoCollection gcolClinicColl
'
'    'Create a create-hl7constructor object only long enough to get a hl7constructor object amw 8/20/99
'        oLogger.LogtoFile c_mszModuleName & "CreateHL7SessionForSigning - CreateHl7Constructor", en_Entering, "hlErr", hlErr
'          Set oCreateHl7Constructor = New CreateHl7Constructor
'          Set oHl7Constructor = oCreateHl7Constructor.GetHl7Constructor
'          Set oCreateHl7Constructor = Nothing
'            oLogger.LogtoFile c_mszModuleName & "CreateHL7SessionForSigning - CreateHl7Constructor", en_Exiting, "hlErr", hlErr
'
'            oLogger.LogtoFile c_mszModuleName & "CreateHL7SessionForSigning - GetUserInformation", en_Entering, "hlErr", hlErr
'
'          GetUserInformation szAccessCode, szVerifyCode, szCHCSIpAddr, szCHCSIpPort, mvarClinicIEN, mvarMEPRSCode, mvarMEPRSIen, szMTFCode
'          szLocalIP = szGetLocalIPAddress()
'            oLogger.LogtoFile c_mszModuleName & "CreateHL7SessionForSigning - GetUserInformation", en_Exiting, "hlErr", hlErr
'
'            oLogger.LogtoFile c_mszModuleName & "CreateHL7SessionForSigning - GetPatientInformation", en_Entering, "hlErr", hlErr
'
'            oLogger.LogtoFile c_mszModuleName & "CreateHL7SessionForSigning - Login", en_Entering, "hlErr", hlErr
'            hlErr = oHl7Constructor.Login(szAccessCode, szVerifyCode, hlLOGIN.hlUNENCRYPTED, _
'               szCHCSIpAddr, szCHCSIpPort, szLocalIP)
'
'
'          oLogger.LogtoFile c_mszModuleName & "CreateHL7SessionForSigning - Login", en_Exiting, "hlErr", hlErr
'
'          mvarLoginStatus = hlErr    '2-8-00; AAB PR#191
'          mvarSessionError = hlErr   '3-8-01; AAB: INC2(10073)
'
'          If hlErr <> hlERROR_NONE Then
'
'
'                GenerateLoginFailureMessage hlErr
'
'                'szErr = oHl7Constructor.GetErrorText(hlErr)
'                CreateHL7SessionForSigning = False
'                bHL7SessionCreated = True
'                oLogger.LogtoFile c_mszModuleName & "CreateHL7SessionForSigning - FAILED2", en_Exiting, "hlErr", hlErr, "szPatientIEN", szPatientIEN
'
'                Exit Function
'
'          End If
'
'    'no errors ... proceed
''          szPatientName = oHl7Constructor.goPatient.Name
''          Set mvarRxTablet = oHl7Constructor.RXTablet
''          Set mvarRadTablet = oHl7Constructor.RadTablet
''          Set mvarLabTablet = oHl7Constructor.LabTablet
'
'
'          CreateHL7SessionForSigning = True
'
'    'BUG Fix: 02-22-2000; AAB for CT enhancement: failed to occupy new property registers
'          mvarszProviderIen = oHl7Constructor.szProviderIEN
'          mvarszOrderingProviderIen = oHl7Constructor.szOrderingProviderIEN
'          mvarszUserIen = oHl7Constructor.szUserIEN
'     '     mvarszDefOrderPage = oHl7Constructor.szDefOrderPage
'          oLogger.LogtoFile c_mszModuleName & "CreateHL7SessionForSigning", en_Exiting, "hlErr", hlErr
'
'End Function


Public Function Logout() As hlERRORS

Dim hlErr As hlERRORS
    
    
    If Not (oHl7Constructor Is Nothing) Then           '06-26-01; TLS: prevent re-logout attempts
          hlErr = oHl7Constructor.Logout()
          mvarLoginStatus = hlERROR_NOT_LOGGED_IN
          Logout = hlErr
          
          '06-26-01; TLS: old constructor = old user ... NG
          'Set oHl7Constructor = Nothing
    End If
    
End Function

Public Function Login(szAccessCode As String, szVerifyCode As String, szCHCSIpAddr As String, szCHCSIpPort As String) As hlERRORS
    
Dim hlErr As hlERRORS
Dim szLocalIP As String
Dim iNull As Integer
    
                              ' Dim oWinsock As Winsock
                              ' Set oWinsock = New Winsock1
                              
           szLocalIP = Space$(100)
           oHl7Constructor.oHL7Server.GetClientAddress szLocalIP
           szLocalIP = Trim$(szLocalIP)
           iNull = InStr(1, szLocalIP, vbNullChar)
           If iNull > 0 Then
                szLocalIP = Left$(szLocalIP, iNull - 1)
           End If
           
           hlErr = oHl7Constructor.Login(szAccessCode, szVerifyCode, hlLOGIN.hlUNENCRYPTED, szCHCSIpAddr, szCHCSIpPort, szLocalIP, mvarClinicIEN, mvarMEPRSCode, mvarMEPRSIen, mvarMTFCode)

           Login = hlErr
           mvarLoginStatus = hlErr             '2-8-00; AAB PR#191
           mvarSessionError = hlErr          '03-05-01; AAB INC2(10073)
            
                              ''Set oWinsock = Nothing

End Function

Private Sub Class_Initialize()

    mvarSessionError = hlERROR_NONE     'start clean; created 9-2-99; AAB
    mvarLoginStatus = hlERROR_NONE       '2-8-00; AAB; PR#191
    
    '--- SCR-ARM-16113 MSolano 8/24/01; COMMON TO MULTIPLE ARM-SCRs
    '--- Init routine should be called from first top level creatable coclass
    '--- It can be called multiple times, but will actually run only once.
    '--- Instrument all creatable classes in case dll's multiple entry
    '--- points are actually used
     Call ARM_InitAllTransactionsForOE

    
    If gobjShared Is Nothing Then Set gobjShared = New CWShared
    Set goRxTransactionCol = New RxTransactionCollection
    Set goLabTransactionCol = New LabTransactionCollection
    Set goRadTransactionCol = New RadTransactionCollection
    Set gcolClinicColl = New ClinicCollection
    
    
    mbCalledAfterInit2 = True
    mbCalledFromTabletUIs = False
    
    bHL7SessionCreated = False
    gbBackgroundBusy = False
    
    '05-30-00; AAB
    IsActionViaSubmitButton(hlRX_ORDERS_CODE) = False
    IsActionViaSubmitButton(hlRAD_ORDERS_CODE) = False
    IsActionViaSubmitButton(hlLAB_ORDERS_CODE) = False
    
    'added 9-5-01 for shutting off/on show all orders
    gbShowAllRads = False
    gbShowAllMeds = False
    gbShowAllLabs = False
    
    'mvarPatientRequired = True
    
          '    If goOESession Is Nothing Then
          '        Set goOESession = Me
          '        goOESession.CreateHL7Session
          '    End If

End Sub

Private Sub Class_Terminate()

    Dim hlErr As hlERRORS
    'KELLY DEAD 7-27-01 Dim i As Integer
    
'Debug.print "Class Terminate OE_SessionManager"
'Debug.print "Try to Logout"

        If Not (oHl7Constructor.oHL7Server Is Nothing) Then
            hlErr = oHl7Constructor.Logout
            mvarLoginStatus = hlERROR_NOT_LOGGED_IN
        End If
        bHL7SessionCreated = False
        
'        moeSearchEngine.bLogoffDatabase
'    For i = 1 To goRadTransactionCol.Count
'        goRadTransactionCol(i).Remove
'    Next i
'    For i = 1 To goLabTransactionCol.Count
'        goLabTransactionCol(i).Remove
'    Next i
'    For i = 1 To goRxTransactionCol.Count
'        goRxTransactionCol(i).Remove
'    Next i
'
    Set gcolClinicColl = Nothing
    
End Sub

Private Sub GetUserInformation(szAccessCode As String, szVerifyCode As String, szCHCSIpAddr As String, szCHCSIpPort As String, szClinicIEN As String, szMEPRSCode As String, szMEPRSIEN As String, szMTFCode As String)

'will use the local block to get these values
'but now hardcode them

               '    szAccessCode = GetAppProfile("USER", "szAccessCode", "")
               '    szVerifyCode = GetAppProfile("USER", "szVerifyCode", "")
               '    szCHCSIpAddr = GetAppProfile("USER", "szCHCSIpAddr", "139.121.192.17")
               '    szCHCSIpPort = GetAppProfile("USER", "szCHCSIpPort", "2204")
               '    szClinicIen = GetAppProfile("USER", "szClinicIen", "21")
               '    szMeprsCode = GetAppProfile("USER", "szMeprsCode", "BAAA")
               '
               '    szMeprsIEN = GetAppProfile("USER", "szMeprsIen", "130")
               '    szMTFCode = GetAppProfile("USER", "szMTFCode", "F1756")
               '
    On Error GoTo ErrorHandler
    
    'new CoreAPI functions
    
    'path:\core stubb\bin
    
    Dim i As Integer
    Dim objResources As ICHCSII_Resources
    Dim objResource As ICHCSII_Resource
 
    Set objResources = gobjUser.Resources
    
    For i = 1 To objResources.Count
        Set objResource = gobjUser.Resources.Item(i)
        If objResource.Name = "CHCS" Then
            szAccessCode = objResource.UserId
            szVerifyCode = objResource.Password
            szCHCSIpAddr = objResource.GetAttribute("OE_IPAddr")
            szCHCSIpPort = objResource.GetAttribute("OE_Port")
            szMTFCode = objResource.GetAttribute("DMIS_ID")
        End If
        
    Next i
    
    Dim objClinicOps As New ClinicOpsEx
    Dim objClinic As DTOs.Clinic
    
    ''Set objClinic = objClinicOps.GetClinic(gobjLogonEnvironment.FacilityNCID)
    Set objClinic = objClinicOps.GetClinic(gobjLogonEnvironment.NursingDivisionNCID)
               
    If (objClinic Is Nothing) Then
        Exit Sub
    End If
    szClinicIEN = objClinic.IEN
    mvarClinicName = objClinic.Name
    szMEPRSCode = objClinic.MEPRSCode
    szMEPRSIEN = objClinic.MEPRSIen
    
    'BUG Fix: 02-22-2000; AAB for CT enhancement: failed to copy these 'local' variable dautm into respective 'mvar' property registers
     mvarMTFCode = szMTFCode
   
    'Determine if User is a Provider
'    Dim szRoleName As String
    
'    gbUserIsProvider = False
'
'    Set mobjProviderSearch = New CHCSII_NonProvider.clsNonProvider
'    Set mRstSearchedProviders = New ADODB.Recordset
'    mobjProviderSearch.OpenNonProvider cwobjLogonEnviron, cwobjCWCommSession, gobjUser
'    mobjProviderSearch.getProviders mRstSearchedProviders
'    If mRstSearchedProviders.State <> adStateClosed Then
'        If Not mRstSearchedProviders.EOF And Not mRstSearchedProviders.BOF Then
'            mRstSearchedProviders.MoveFirst
'            If mRstSearchedProviders.Fields("IEN").Value = oHl7Constructor.szOrderingProviderIEN Then
'                gbUserIsProvider = True
'            End If
'            mRstSearchedProviders.MoveFirst
'        End If
'    End If
    
'    With gobjUser
'        For i = 1 To .Roles.Count
'            szRoleName = LCase$(.Roles.Item(i).Name)
'            If InStr(szRoleName, "provider") > 0 Then
'                gbUserIsProvider = True
'                Exit For
'            End If
'            If InStr(szRoleName, "independent medic") > 0 Then
'                gbUserIsProvider = True
'                Exit For
'            End If
'        Next i
'    End With
'
    Exit Sub
    
ErrorHandler:
    If Err.Number <> -2147220494 Then   'GG 12/10/2002 SCR #29491
        MsgBxARMd "Error :" & Err.Number & "  " & Err.Description
    End If
    Resume Next
    
End Sub
Private Sub GetPatientInformation(szPatientIEN As String, szPatientName As String)
'will use the local block to get these values
'but now hardcode them
    
'    szPatientIEN = GetAppProfile("PATIENT", "szPatientIEN", "84082")
'    szPatientName = GetAppProfile("PATIENT", "szPatientName", "PREBECK, CHERRIE F")
    szPatientIEN = gobjPatient.MedicalRecordNumber
    szPatientName = gobjPatient.Name
    

End Sub
Public Function RefreshPatientInfo() As Boolean

Dim szPatientIEN As String
Dim szPatientName As String
    
          RefreshPatientInfo = False
          GetPatientInformation szPatientIEN, szPatientName
          If mvarLoginStatus = hlERROR_NOT_LOGGED_IN Then
              
              ReleaseAndSelectPatient
              RefreshEachUIsOrdersOnPage
              RefreshPatientInfo = True
              'mvarLoginStatus = hlERROR_NONE
          
          ElseIf szPatientIEN <> oHl7Constructor.goPatient.IEN And (mvarLoginStatus = hlERROR_NONE Or mvarLoginStatus = hlERROR_PATIENT_SELECT_FAILED) Then
              mvarLoginStatus = hlERROR_NONE
              ReleaseAndSelectPatient
              RefreshEachUIsOrdersOnPage
              RefreshPatientInfo = True
          End If
          
End Function

Private Function GetHl7Errors(OrderType As hlORDERTYPECODE, szErrMsg As String, iErrCount As Integer) As Boolean
    
Dim iCount As Integer
    
    GetHl7Errors = False
    
          Select Case OrderType
    
          Case hlRX_ORDERS_CODE
               iErrCount = oHl7Constructor.RxError.Count
               If oHl7Constructor.RxError.Count > 0 Then
                   For iCount = 1 To oHl7Constructor.RxError.Count
                       szErrMsg = szErrMsg & oHl7Constructor.RxError(iCount).Text & vbCrLf
                   Next iCount
                   GetHl7Errors = True
               End If
             
          Case hlRAD_ORDERS_CODE
               iErrCount = oHl7Constructor.RadError.Count
               If oHl7Constructor.RadError.Count > 0 Then
                  For iCount = 1 To oHl7Constructor.RadError.Count
                       szErrMsg = szErrMsg & oHl7Constructor.RadError(iCount).Text & vbCrLf
                  Next iCount
                  GetHl7Errors = True
               End If
               
          Case hlLAB_ORDERS_CODE
               iErrCount = oHl7Constructor.LabError.Count
               If oHl7Constructor.LabError.Count > 0 Then
                  For iCount = 1 To oHl7Constructor.LabError.Count
                       szErrMsg = szErrMsg & oHl7Constructor.LabError(iCount).Text & vbCrLf
                  Next iCount
                  GetHl7Errors = True
               End If
          Case Else
               iErrCount = oHl7Constructor.OtherError.Count
               If oHl7Constructor.OtherError.Count > 0 Then
                  For iCount = 1 To oHl7Constructor.OtherError.Count
                       szErrMsg = szErrMsg & oHl7Constructor.OtherError(iCount).Text & vbCrLf
                  Next iCount
                  GetHl7Errors = True
               End If
          
          End Select
  
    
End Function
Public Function GetHl7Warnings(OrderType As hlORDERTYPECODE, bOrderIDProvided As Boolean) As Boolean
    
    Dim bWarnings As Boolean
    
    bWarnings = False              'set to TRUE if there are warnings ... see below
    bOrderIDProvided = False ' TRUE if there is an order id returned; should then use an override message
    
    'oHl7Constructor.RxWarning(1).WarningTypeCode = 5 is DOSEAGE type
    
    Select Case OrderType
         Case hlRX_ORDERS_CODE
             If oHl7Constructor.RxWarning.Count > 0 Then
                bWarnings = True
                If LenB(oHl7Constructor.RXTablet.orderid) > 0 Then
                    bOrderIDProvided = True
                End If
             End If
             
         Case hlRAD_ORDERS_CODE
            If oHl7Constructor.RadWarning.Count > 0 Then
                bWarnings = True
                If LenB(oHl7Constructor.RadTablet.orderid) > 0 Then
                    bOrderIDProvided = True
                End If
            End If
            
         Case hlLAB_ORDERS_CODE
            If oHl7Constructor.LabWarning.Count > 0 Then
                bWarnings = True
                If LenB(oHl7Constructor.LabTablet.orderid) > 0 Then
                    bOrderIDProvided = True
                End If
            End If
         Case hlCON_ORDERS_CODE
            Set oHl7Constructor.OtherWarning = New Collection
'            Dim szTextBody As String, szWarningHeader As String, szTestname As String
'            Dim iNumWarnings As Integer
            mvarOrderTypeCode = hlCON_ORDERS_CODE
            oHl7Constructor.HL7GetWarnings oHl7Constructor.OtherWarning, oHl7Constructor.OtherTablet
            If oHl7Constructor.OtherWarning.Count > 0 Then
                bWarnings = True
                If LenB(oHl7Constructor.OtherTablet.orderid) > 0 Then
                    bOrderIDProvided = True
                End If
            End If
        Case hlSiGN_ORDERS_CODE
            If oHl7Constructor.colUnsignedOrder(oHl7Constructor.colUnsignedOrder.WarningIndex).WarningCollection.Count > 0 Then
                bWarnings = True
                bOrderIDProvided = True
            End If

     End Select
     
     If bWarnings Then
        'Tracker 26281
        Screen.MousePointer = vbDefault
        If OrderType = hlSiGN_ORDERS_CODE Then
            ShowWarningScreen oHl7Constructor.colUnsignedOrder(oHl7Constructor.colUnsignedOrder.WarningIndex).OrderType, bOrderIDProvided
        Else
            ShowWarningScreen OrderType, bOrderIDProvided
        End If
        AskCoreToRefreshAfterModal OrderType
        'Tracker 26281
        Screen.MousePointer = vbHourglass
    End If
     
     Select Case OrderType
         Case hlRX_ORDERS_CODE
            mvarszOverrideReason = oHl7Constructor.RXTablet.WarningReason
         Case hlRAD_ORDERS_CODE
            mvarszOverrideReason = oHl7Constructor.RadTablet.WarningReason
         Case hlLAB_ORDERS_CODE
            mvarszOverrideReason = oHl7Constructor.LabTablet.WarningReason
         Case hlCON_ORDERS_CODE
            mvarszOverrideReason = oHl7Constructor.OtherTablet.WarningReason
        Case hlSiGN_ORDERS_CODE
            mvarszOverrideReason = oHl7Constructor.colUnsignedOrder(oHl7Constructor.colUnsignedOrder.WarningIndex).WarningReason
            
     End Select
    
    GetHl7Warnings = bWarnings  'bug fix 9-3-99; AAB
    
End Function
Public Sub GetHl7WarningText(ByVal OrderType As hlORDERTYPECODE, szTextBody As String, szWarningHeader As String, iNumWarnings As Integer, szTestname As String)
    
    Dim iCount As Integer
    Dim pWarningCollection As Collection
    
    Select Case OrderType
        Case hlRX_ORDERS_CODE
            Set pWarningCollection = oHl7Constructor.RxWarning
            szWarningHeader = oHl7Constructor.RXTablet.WarningHeader
            szTestname = oHl7Constructor.RXTablet.OrderName
        Case hlLAB_ORDERS_CODE
            Set pWarningCollection = oHl7Constructor.LabWarning
            szWarningHeader = oHl7Constructor.LabTablet.WarningHeader
            szTestname = oHl7Constructor.LabTablet.OrderName
        Case hlRAD_ORDERS_CODE
            Set pWarningCollection = oHl7Constructor.RadWarning
            szWarningHeader = oHl7Constructor.RadTablet.WarningHeader
            szTestname = oHl7Constructor.RadTablet.OrderName
        Case hlSiGN_ORDERS_CODE
            Set pWarningCollection = oHl7Constructor.colUnsignedOrder(oHl7Constructor.colUnsignedOrder.WarningIndex).WarningCollection
            szWarningHeader = oHl7Constructor.colUnsignedOrder(oHl7Constructor.colUnsignedOrder.WarningIndex).WarningHeader
            szTestname = oHl7Constructor.colUnsignedOrder(oHl7Constructor.colUnsignedOrder.WarningIndex).szDisplayText
            OrderType = oHl7Constructor.colUnsignedOrder(oHl7Constructor.colUnsignedOrder.WarningIndex).OrderType
        Case hlCON_ORDERS_CODE
            Set pWarningCollection = oHl7Constructor.OtherWarning
            szWarningHeader = oHl7Constructor.OtherTablet.WarningHeader
            szTestname = oHl7Constructor.OtherTablet.OrderName
    End Select
    szTextBody = ""
    iNumWarnings = pWarningCollection.Count
    Select Case OrderType
         Case hlRX_ORDERS_CODE
            'iNumWarnings = oHl7Constructor.RXWarning.Count
            'szWarningHeader = oHl7Constructor.RXTablet.WarningHeader
            Dim bHitPDTS As Boolean
            
            If pWarningCollection(1).WarningTypeCodePDTS = hlPDTSWarning_Local Then
                szTextBody = "LOCAL WARNINGS" & vbCrLf
                bHitPDTS = False
            Else
                szTextBody = "PDTS WARNINGS" & vbCrLf
                bHitPDTS = True
            End If
            For iCount = 1 To iNumWarnings
                If iCount > 1 Then
                    If pWarningCollection(iCount).WarningTypeCodePDTS <> hlPDTSWarning_Local And bHitPDTS = False Then
                        szTextBody = szTextBody & "PDTS WARNINGS" & vbCrLf
                        bHitPDTS = True
                    End If
                End If
                szTextBody = szTextBody & iCount & ") " & pWarningCollection(iCount).ReasonOrProblem & vbCrLf
                If bHitPDTS = False Then
                    If InStr(pWarningCollection(iCount).ReasonOrProblem, "Allergic Reaction Class") > 0 Then
                        szTextBody = szTextBody & vbTab & "Patient allergies to:"
                    End If
    
                    If LenB(pWarningCollection(iCount).ConflictingOrderName) > 0 Then
                        szTextBody = szTextBody & vbTab & vbTab & pWarningCollection(iCount).ConflictingOrderName & vbCrLf
                    End If
                                
                    If LenB(pWarningCollection(iCount).ConflictSigAndDate) > 0 Then
                        szTextBody = szTextBody & vbTab & vbTab & pWarningCollection(iCount).ConflictSigAndDate & vbCrLf
                    End If
                End If
                If bHitPDTS = True Then
'                    If LenB(pWarningCollection(iCount).PDTSExpandedWarning) > 0 Then
'                        szTextBody = szTextBody & vbTab & vbTab & pWarningCollection(iCount).PDTSExpandedWarning & vbCrLf
'                    End If
                Else
                    If LenB(pWarningCollection(iCount).ConflictsWithOrWhy) > 0 And LenB(pWarningCollection(iCount).ConflictSigAndDate) = 0 Then
                        szTextBody = szTextBody & vbTab & vbTab & pWarningCollection(iCount).ConflictsWithOrWhy & vbCrLf
                    End If
                End If
                
            Next iCount
            'szTestName = oHl7Constructor.RXTablet.OrderName
         Case hlRAD_ORDERS_CODE
            'iNumWarnings = oHl7Constructor.radWarning.Count
            For iCount = 1 To iNumWarnings
                szTextBody = szTextBody & iCount & ") " & pWarningCollection(iCount).ConflictData & vbCrLf
                szTextBody = szTextBody & vbTab & pWarningCollection(iCount).WarningDescription & vbCrLf
               ' szTextBody = szTextBody & vbTab & oHl7Constructor.RadWarning(iCount).SectionTitle & vbCrLf
            Next iCount
            'szTestName = oHl7Constructor.RadTablet.OrderName
         Case hlLAB_ORDERS_CODE
            'iNumWarnings = oHl7Constructor.labWarning.Count
            For iCount = 1 To iNumWarnings
                szTextBody = szTextBody & iCount & ") " & pWarningCollection(iCount).ConflictData & vbCrLf
                szTextBody = szTextBody & vbTab & pWarningCollection(iCount).WarningDescription & ":" & pWarningCollection(iCount).ConflictingOrderName & vbCrLf
'                szTextBody = szTextBody & vbTab & vbTab & oHl7Constructor.LabWarning(iCount).ConflictSigAndDate & vbCrLf
            Next iCount
        Case hlCON_ORDERS_CODE
           
            For iCount = 1 To iNumWarnings
                szTextBody = szTextBody & iCount & ") " & pWarningCollection(iCount).ConflictData & vbCrLf
                szTextBody = szTextBody & vbTab & pWarningCollection(iCount).WarningDescription & ":" & pWarningCollection(iCount).ConflictingOrderName & vbCrLf
            Next iCount
            
     End Select
    
End Sub


Private Function bCheckOrderNamesOK(OrderType As hlORDERTYPECODE) As Boolean
    
    Dim szMsgBoxCaption As String, szMsgBoxText As String
    Dim iReturn As Integer
    Dim WorkingTablet As Object
    
    bCheckOrderNamesOK = True
    
    Select Case OrderType
        Case hlRX_ORDERS_CODE:
            Set WorkingTablet = oHl7Constructor.RXTablet
        Case hlRAD_ORDERS_CODE:
            Set WorkingTablet = oHl7Constructor.RadTablet
        Case hlLAB_ORDERS_CODE:
            Set WorkingTablet = oHl7Constructor.LabTablet
    End Select
 
    ''If WorkingTablet.OrderName <> oHl7Constructor.Order.LastPreverifyNameWas Then
    If WorkingTablet.LastPreverifyNamesMatch = False Then       'changed 9-2-99; AAB
        szMsgBoxCaption = "CHCS I Is Substituting Your Selection"
        szMsgBoxText = "Your selection of " & WorkingTablet.LastPreverifyNameWas & vbCrLf
        szMsgBoxText = szMsgBoxText & "will be substituted with " & WorkingTablet.OrderName
        iReturn = MsgBxARMd(szMsgBoxText, vbInformation + vbOKCancel, szMsgBoxCaption)
        AskCoreToRefreshAfterModal OrderType
        If iReturn <> vbOK Then
            bCheckOrderNamesOK = False
        End If
    End If
    
End Function

Private Function bAnyErr50(OrderType As hlORDERTYPECODE) As Boolean

Dim i As Integer
Dim bErr50MoreWarnings As Boolean
    
    'check to see if there are any Err 50s- non-overriden warnings
    'only return that there are err50 if you get the message and there are warnings to override
    
    
    bErr50MoreWarnings = False
    
    Select Case OrderType
        Case hlRX_ORDERS_CODE:
            If oHl7Constructor.RxWarning.Count > 0 Then
                For i = 1 To oHl7Constructor.RxError.Count
                    If oHl7Constructor.RxError(i).TypeCode = "50" Then
                        bErr50MoreWarnings = True
                    End If
                Next i
            End If
        Case hlRAD_ORDERS_CODE:
            If oHl7Constructor.RadWarning.Count > 0 Then
                For i = 1 To oHl7Constructor.RadError.Count
                    If oHl7Constructor.RadError(i).TypeCode = "50" Then
                        bErr50MoreWarnings = True
                    End If
                Next i
            End If
        Case hlLAB_ORDERS_CODE:
            If oHl7Constructor.LabWarning.Count > 0 Then
                For i = 1 To oHl7Constructor.LabError.Count
                    If oHl7Constructor.LabError(i).TypeCode = "50" Then
                        bErr50MoreWarnings = True
                    End If
                Next i
            End If
    End Select

    bAnyErr50 = bErr50MoreWarnings
    
End Function

Public Function ReleaseAndSelectPatient(Optional szPatientName As String, Optional szPatientIEN As String)

    
    Dim hlErr As hlERRORS
    'KELLY DEAD 7-27-01 Dim szErrMsg As String
    'KELLY DEAD 7-27-01 Dim iCount As Integer
    If Not mvarLoginStatus = hlERROR_NOT_LOGGED_IN Then
        If LenB(oHl7Constructor.goPatient.Name) > 0 Then
            oHl7Constructor.SendPatientReleaseMessage
        End If
    End If
    
    If szPatientIEN = "" Then
        GetPatientInformation szPatientIEN, szPatientName
    End If
    oLogger.LogtoFile c_mszModuleName & "ReleaseAndSelectPatient", en_Entering, "szPatientIEN", szPatientIEN
    
          ''If mvarLoginStatus <> hlERROR_LOGIN_COMM_ERROR Then     'for PR#191
     If mvarLoginStatus = hlERROR_NONE Or mvarLoginStatus = hlERROR_PATIENT_SELECT_FAILED Then    'for PR#191 changed akm 4/3/2000
               
               'removed 06-20-00; AAB per AKM; try only once
               ''For iCount = 1 To 3
                    hlErr = oHl7Constructor.SendPatientSelectMsg(szPatientIEN, szPatientName)
                    If hlErr <> hlERROR_NONE Then
                           If hlErr = hlERROR_TIMED_OUT Or hlErr = hlERROR_NOT_LOGGED_IN Then
                               hlErr = ReconnectToCHCS()
                                   ''hlErr = oHl7Constructor.SendPatientSelectMsg(szPatientIEN, szPatientName)
                           'Else
                               'Exit For
                           End If
                    'Else
                           'Exit For
                    End If
              '' Next iCount
                If hlErr <> hlERROR_NONE Then
                    GenerateLoginFailureMessage hlErr

               End If
    
    End If '2-8-00;AAB; PR#191
    
    'Tracker 20191 moved to CommonPatientSelected on 2/10/2003 - AKM
    If Len(gobjShared.CmdLineSwitches("GREEN")) = 0 Then
        Set goRxTransactionCol = Nothing
        Set goLabTransactionCol = Nothing
        Set goRadTransactionCol = Nothing

        Set goRxTransactionCol = New RxTransactionCollection
        Set goLabTransactionCol = New LabTransactionCollection
        Set goRadTransactionCol = New RadTransactionCollection
    End If
    
    Set mvarRxTablet = oHl7Constructor.RXTablet
    Set mvarRadTablet = oHl7Constructor.RadTablet
    Set mvarLabTablet = oHl7Constructor.LabTablet
    Set mvarOtherTablet = oHl7Constructor.OtherTablet
    
    mvarLoginStatus = hlErr
    oLogger.LogtoFile c_mszModuleName & "ReleaseAndSelectPatient", en_Exiting, "hlErr", hlErr
  
End Function


Public Sub SendMessageToCoreAndAddTransaction(OrderType As hlORDERTYPECODE, OrderAction As hlACTION, Optional szOldOrderID As String)
    
Dim szNewOrderID As String
Dim szOrderID As String

Dim tRxTrans As RxTransaction
Dim tLabTrans As LabTransaction
Dim tRadTrans As RadTransaction
  
          On Error GoTo ErrorHandler
    
          Select Case OrderType
          Case hlRX_ORDERS_CODE:
               szNewOrderID = goOESession.RXTablet.orderid
          Case hlLAB_ORDERS_CODE:
               szNewOrderID = goOESession.LabTablet.orderid
          Case hlRAD_ORDERS_CODE:
               szNewOrderID = goOESession.RadTablet.orderid
          End Select
          
          If OrderAction = hlACTION_ADD_TO_QUEUE Then
            szNewOrderID = "Q" & Format(Now, "yyyymmddhhmmss")
           End If
          
          If LenB(szOldOrderID) > 0 And LenB(szNewOrderID) > 0 Then
                    '06-16-00; AAB PR#259
                    ''szOrderId = szNewOrderID & "\" & szOldOrderID
               szOrderID = szNewOrderID & "|" & szOldOrderID
         ElseIf LenB(szOldOrderID) > 0 Then
              szOrderID = szOldOrderID
          ElseIf LenB(szNewOrderID) > 0 Then
              szOrderID = szNewOrderID
          End If
          
          Select Case OrderType
          Case hlRX_ORDERS_CODE:
                 'Add an rxtransaction to the rxtransactioncollection object
                    With tRxTrans
                        .OldOrderid = szOldOrderID
                        .NewOrderid = szNewOrderID
                        .Action = OrderAction
                        Set .RXTablet = oHl7Constructor.RXTablet.CopyTablet
                        Set .RXWarnings = New Collection
                        Dim cWarningObject As RxWarningType
                        Dim ccWarningObject As RxWarningType
                       
                        
                        For Each cWarningObject In oHl7Constructor.RxWarning
                            Set ccWarningObject = cWarningObject.CopyWarning
                            .RXWarnings.Add ccWarningObject
                        Next
                       
                         'temp amw
                         Debug.Print "oldOrderid: " & .OldOrderid
                         Debug.Print "NewOrderid: " & .NewOrderid
                         Debug.Print "Action:     " & .Action
                         '.RXTablet.DebugPrint
                         
                          goRxTransactionCol.Add tRxTrans
     
                    End With
                    cwobjCWCommRx.Message OrderAction, szOrderID, NCID_BROADCAST_MSG, NCID_OE_MED_APPLICATION
                    oLogger.LogtoFile c_mszModuleName & "SendMessageToCoreAndAddTransaction - RX", 1, "OrderAction", OrderAction, "szOrderId", szOrderID

              Case hlLAB_ORDERS_CODE:
                    With tLabTrans
                        .OldOrderid = szOldOrderID
                        .NewOrderid = szNewOrderID
                        .Action = OrderAction
                        Set .LabTablet = oHl7Constructor.LabTablet.CopyTablet
                        
                                   'temp amw
                                   Debug.Print "OldOrderid: " & .OldOrderid
                                   Debug.Print "NewOrderid: " & .NewOrderid
                                   Debug.Print "Action:     " & .Action
                                   '.LabTablet.DebugPrint

                        
                        Set .LabWarnings = New Collection
                        Dim cLabWarningObject As LabWarningType
                        Dim ccLabWarningObject As LabWarningType
                        For Each cLabWarningObject In oHl7Constructor.LabWarning
                            Set ccLabWarningObject = cLabWarningObject.CopyWarning
                            .LabWarnings.Add ccLabWarningObject
                        Next
                        goLabTransactionCol.Add tLabTrans
                    End With
                    cwobjCWCommLab.Message OrderAction, szOrderID, NCID_BROADCAST_MSG, NCID_OE_LAB_APPLICATION
                    oLogger.LogtoFile c_mszModuleName & "SendMessageToCoreAndAddTransaction - LAB", 1, "OrderAction", OrderAction, "szOrderId", szOrderID
                  
              Case hlRAD_ORDERS_CODE:
                    With tRadTrans
                        .OldOrderid = szOldOrderID
                        .NewOrderid = szNewOrderID
                        .Action = OrderAction
                        Set .RadTablet = oHl7Constructor.RadTablet.CopyTablet
                        
                                   'temp amw
                                   Debug.Print "oldOrderid: " & .OldOrderid
                                   Debug.Print "NewOrderid: " & .NewOrderid
                                   Debug.Print "Action:     " & .Action
                                   '.RadTablet.DebugPrint
                                   
                        
                        Set .RadWarnings = New Collection
                        Dim cRadWarningObject As RadWarningType
                        Dim ccRadWarningObject As RadWarningType
                        For Each cRadWarningObject In oHl7Constructor.RadWarning
                            Set ccRadWarningObject = cRadWarningObject.CopyWarning
                            .RadWarnings.Add ccRadWarningObject
                        Next
                        goRadTransactionCol.Add tRadTrans
                    End With
                    oLogger.LogtoFile c_mszModuleName & "SendMessageToCoreAndAddTransaction - RAD", 1, "OrderAction", OrderAction, "szOrderId", szOrderID
                    cwobjCWCommRad.Message OrderAction, szOrderID, NCID_BROADCAST_MSG, NCID_OE_RAD_APPLICATION
          End Select
          
          Exit Sub
    
ErrorHandler:
    '%%%
End Sub

Public Sub SendErrorMessageToCore(OrderType As hlORDERTYPECODE, OrderAction As hlACTION, ErrorCode As hlERRORS)

'Created new 05-30-00; AAB; taken from SendMessageToCore ...
  ''Public Sub SendErrorMessageToCore(OrderType As hlORDERTYPECODE, OrderAction As hlACTION, Optional szOldOrderID As String)
  '
  '
  'always send OrderAction of hlACTION_ORDER_SET
Dim szOrderId_AsError As String
  
                    On Error GoTo ErrorHandler
                    
                    szOrderId_AsError = "HL7ERR = " & Trim$(Str$(ErrorCode))
          
                    Select Case OrderType
                    Case hlRX_ORDERS_CODE:
                        cwobjCWCommRx.Message hlACTION_ORDER_SET, szOrderId_AsError, NCID_BROADCAST_MSG, NCID_OE_MED_APPLICATION
                        oLogger.LogtoFile c_mszModuleName & "SendErrorMessageToCore - RX", 1, "OrderAction", OrderAction, "szOrderId_AsError", szOrderId_AsError
                    
                    Case hlLAB_ORDERS_CODE:
                        cwobjCWCommLab.Message hlACTION_ORDER_SET, szOrderId_AsError, NCID_BROADCAST_MSG, NCID_OE_LAB_APPLICATION
                        oLogger.LogtoFile c_mszModuleName & "SendErrorMessageToCore - LAB", 1, "OrderAction", OrderAction, "szOrderId_AsError", szOrderId_AsError
                        
                    Case hlRAD_ORDERS_CODE:
                        cwobjCWCommRad.Message hlACTION_ORDER_SET, szOrderId_AsError, NCID_BROADCAST_MSG, NCID_OE_RAD_APPLICATION
                        oLogger.LogtoFile c_mszModuleName & "SendErrorMessageToCore - RAD", 1, "OrderAction", OrderAction, "szOrderId_AsError", szOrderId_AsError
                          
                    End Select
                    
                    Exit Sub
    
ErrorHandler:
    '%%%
End Sub
Public Sub SendResetTimerMessageToCore(OrderType As hlORDERTYPECODE)

    
    Select Case OrderType
        Case hlRX_ORDERS_CODE:
            cwobjCWCommRx.Message cwiRESET_TIMEOUT_TIMER, "", NCID_BROADCAST_MSG, NCID_OE_MED_APPLICATION
            oLogger.LogtoFile c_mszModuleName & "SendResetTimerMessageToCore - RX", 1
    
    Case hlLAB_ORDERS_CODE:
            cwobjCWCommLab.Message cwiRESET_TIMEOUT_TIMER, "", NCID_BROADCAST_MSG, NCID_OE_LAB_APPLICATION
            oLogger.LogtoFile c_mszModuleName & "SendResetTimerMessageToCore - LAB", 1
        
    Case hlRAD_ORDERS_CODE:
            cwobjCWCommRad.Message cwiRESET_TIMEOUT_TIMER, "", NCID_BROADCAST_MSG, NCID_OE_RAD_APPLICATION
            oLogger.LogtoFile c_mszModuleName & "SendResetTimerMessageToCore - RAD", 1
          
    End Select
                    

    
End Sub

Public Sub ShowCHCSErrors(WorkingTablet As Object, OrderType As hlORDERTYPECODE, bIgnoreErr50 As Boolean, bErr50MoreWarnings As Boolean, lReturnErr As hlERRORS)
    
    Dim szErrMsg As String
    Dim iErrCount As Integer
    
''          If GetHl7Errors(OrderType, szErrMsg, iErrCount) Then
''
''                bErr50MoreWarnings = bAnyErr50(OrderType)
''
''                If WorkingTablet.ErrorIsMessage = False Then                 'added 9-2-99; AAB
''                    If Not bErr50MoreWarnings Then     'added 9-13-99 AKM
''                        MsgBxARMd szErrMsg, vbCritical, "CHCS I Errors"
''                        AskCoreToRefreshAfterModal OrderType
''                        lReturnErr = hlERROR_PERFORM_ACTION_FAILED 'added 9-2-99; AAB
''                    Else
''                        lReturnErr = hlERROR_NONE
''                    End If
''
''                Else
''                    MsgBxARMd szErrMsg, vbInformation, "CHCS I Messages"
''                    AskCoreToRefreshAfterModal OrderType
''                    lReturnErr = hlERROR_NONE
''                End If
''
''          End If
''
''
    
    '=============================================
    
              If GetHl7Errors(OrderType, szErrMsg, iErrCount) Then
                    bErr50MoreWarnings = False
                    bErr50MoreWarnings = bAnyErr50(OrderType)
                
                    If WorkingTablet.ErrorIsMessage = False Then                 'added 9-2-99; AAB
                         'ERROR (not a message)
                         
                              If Not bIgnoreErr50 Then     'added 9-13-99 AKM
                                  MsgBxARMd szErrMsg, vbCritical, "CHCS I Errors"
                                  AskCoreToRefreshAfterModal OrderType
                                  lReturnErr = hlERROR_PERFORM_ACTION_FAILED 'added 9-2-99; AAB
                              
                              ElseIf bIgnoreErr50 And bErr50MoreWarnings Then      'added 9-13-99 AKM
                              
                                   lReturnErr = hlERROR_NONE
                              
                              Else
                                  MsgBxARMd szErrMsg, vbCritical, "CHCS I Errors"
                                  AskCoreToRefreshAfterModal OrderType
                                  lReturnErr = hlERROR_PERFORM_ACTION_FAILED 'added 9-2-99; AAB
                              End If
                              
                    Else      'MESSAGE
                              MsgBxARMd szErrMsg, vbInformation, "CHCS I Messages"
                              AskCoreToRefreshAfterModal OrderType
                              lReturnErr = hlERROR_NONE
                    End If
              
          End If
    

End Sub



Private Function szGetLocalIPAddress() As String
    Dim szLocalIP As String
    Dim iNull As Integer
    
    szLocalIP = Space$(100)
    oHl7Constructor.oHL7Server.GetClientAddress szLocalIP
    szLocalIP = Trim$(szLocalIP)
    iNull = InStr(1, szLocalIP, vbNullChar)
    If iNull > 0 Then
        szLocalIP = Left$(szLocalIP, iNull - 1)
    End If
    szGetLocalIPAddress = szLocalIP

End Function

Public Sub RequestToUpdateScreen(OrderType As hlORDERTYPECODE, OrderAction As hlACTION, Optional orderid As String)
    
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
                    ''PerformAction OrderType, hlACTION_PREVERIFY
        UpdateScreenWithInfo Frm_Active, OrderType
        Frm_Active.Btn_Submit.Enabled = True
        'SCR GetOrders
         'SCR 20176
         If OrderType = hlLAB_ORDERS_CODE Then
                        
            Select Case goOESession.LabTablet.Subscript
                Case hlAUTOPSY
                    Frm_OE_Lab.mbPreVerifyDone = YES
                    Frm_Autopsy_AP.Show vbModal
                    Frm_Active.Btn_Submit.Enabled = False
                Case hlCYTGYN
                            
                    Frm_OE_Lab.mbPreVerifyDone = YES
                    Frm_CytoGyn.Show vbModal
                    Frm_Active.Btn_Submit.Enabled = False
                    
                                
                Case hlBONEMARROW, hlCYTNONGYN, hlSURGPATH
                            
                    Frm_OE_Lab.mbPreVerifyDone = YES
                    Frm_CytoNonGyn.Show vbModal
                    Frm_Active.Btn_Submit.Enabled = False
                            
                End Select
        End If
        'SCR 22638
        SaveToQueueControl Frm_Active.Btn_SaveToQueue, OrderAction, False, OrderType

        Frm_Active.mbPreVerifyDone = True
        AskCoreToRefreshAfterModal OrderType
    ElseIf OrderAction = hlACTION_RENEW_MODIFY Or OrderAction = hlACTION_MODIFY Then
        PerformAction OrderType, hlACTION_PREMODIFY, orderid
        If SessionError = hlERROR_NONE Then
            Select Case OrderType
                Case hlRX_ORDERS_CODE:
                    RXTablet.OrderDate = Now
                Case hlLAB_ORDERS_CODE:
                    LabTablet.Timing.StartDate = Now
                Case hlRAD_ORDERS_CODE:
                    LabTablet.Timing.StartDate = Now
            End Select
'            If OrderAction = hlACTION_RENEW_MODIFY Then
'                Frm_Active.mhlOrderAction = hlACTION_RENEW
'            End If
            UpdateScreenWithInfo Frm_Active, OrderType
            Frm_Active.Btn_Submit.Enabled = True
            'SCR GetOrders
             'SCR 20176

            SaveToQueueControl Frm_Active.Btn_SaveToQueue, OrderAction, False, OrderType

            Frm_Active.mbPreVerifyDone = True
            Frm_Active.mszOrderID = orderid
        End If
    Else
    
        Frm_Active.mszOrderID = orderid
        DoActionButton Val(OrderAction), OrderType, Frm_Active
        
    End If

End Sub
Function GetControlID() As String
    GetControlID = oHl7Constructor.GetControlID
End Function

Private Sub GenerateLoginFailureMessage(hlErr As hlERRORS)
    Dim szErr As String
    Dim szAccessCode As String, szVerifyCode As String, szCHCSIpAddr As String, szCHCSIpPort As String, szMTFCode As String
    Dim szLocalIP As String
    Dim szPatientIEN As String, szPatientName As String
    szLocalIP = szGetLocalIPAddress()
    
    GetUserInformation szAccessCode, szVerifyCode, szCHCSIpAddr, szCHCSIpPort, mvarClinicIEN, mvarMEPRSCode, mvarMEPRSIen, szMTFCode
    GetPatientInformation szPatientIEN, szPatientName
    Select Case hlErr
    
        'Added 2-8-00; AAB: put before next CASE else it will never be singled-out; PR#191
        Case hlERROR_LOGIN_COMM_ERROR
            'Debug.Print
            szErr = oHl7Constructor.GetErrorText(hlErr)
            szErr = szErr & vbCrLf & vbCrLf & "Contact system administrator with the following information:" & vbCrLf & _
            "CHCS I IP Address: " & szCHCSIpAddr & vbCrLf & _
            "CHCS I IP Port: " & szCHCSIpPort & vbCrLf & _
            "Client IP Address: " & szLocalIP
  
            'MsgBxARMd "CHCS I Login Failure " & vbCrLf & szErr, vbCritical, cszErrTitle
            
        Case hlERROR_ACCESSVERIFY_CODE To hlERROR_LOGIN_NOT_A_PROVIDER:
            szErr = oHl7Constructor.GetErrorText(hlErr)
            szErr = szErr & vbCrLf & vbCrLf & "Contact system administrator with the following information:" & vbCrLf & _
            "CHCS I IP Address: " & szCHCSIpAddr & vbCrLf & _
            "CHCS I IP Port: " & szCHCSIpPort & vbCrLf & _
            "Client IP Address: " & szLocalIP
         
            
            '& vbCrLf & vbCrLf & "CHCS I Resource is NOT available; see Administrator"
            'MsgBxARMd "CHCS I Login Failure " & vbCrLf & szErr, vbCritical, cszErrTitle
        
        Case hlERROR_PATIENT_SELECT_FAILED:
            If LenB(oHl7Constructor.gszErrMsg) > 0 Then
                If LenB(szPatientIEN) = 0 Then
                    szErr = "The patient you have selected has no IEN for this MTF/Facility. "
                    szErr = szErr & "You must first enter the patient into CHCS for this MTF " & vbCrLf
                    szErr = szErr & "and wait for the HL7 transfer to populate CHCSII before you can create an appointment in this MTF/Facility"
                Else
                    szErr = oHl7Constructor.gszErrMsg
                End If
                
             Else
                If LenB(szPatientIEN) = 0 Then
                    szErr = "The patient you have selected has no IEN for this MTF/Facility. "
                    szErr = szErr & "You must first enter the patient into CHCS for this MTF " & vbCrLf
                    szErr = szErr & "and wait for the HL7 transfer to populate CHCSII before you can create an appointment in this MTF/Facility"
                Else
                    szErr = oHl7Constructor.GetErrorText(hlErr)
                End If
             End If
                
                szErr = "Error from CHCS I :" & szErr & _
                vbCrLf & vbCrLf & "Contact system administrator with the following information:" & _
                vbCrLf & "CHCS II Patient Name: " & szPatientName & vbCrLf
                If LenB(szPatientIEN) > 0 Then
                    szErr = szErr & "CHCS II Patient Medical Record Number: " & szPatientIEN & vbCrLf
                End If
                szErr = szErr & "CHCS I  Provider IEN: " & oHl7Constructor.szProviderIEN & vbCrLf & _
                "CHCS I  User IEN: " & oHl7Constructor.szUserIEN
            'MsgBxARMd "CHCS I Patient Select Failure " & vbCrLf & szErr, vbCritical, cszErrTitle
        
        End Select
        mvarszConnectFailureText = szErr

End Sub

Public Function ICHCSII_AnP_SetAnPParent(objAnPParent As Object) As Boolean
    CommonSetAnP objAnPParent
End Function


Public Sub Reset(OrderType As hlORDERTYPECODE)
    Select Case OrderType
        Case hlRX_ORDERS_CODE:
            Set oHl7Constructor.RxError = New Collection
            Set oHl7Constructor.RxWarning = New Collection
        
        Case hlLAB_ORDERS_CODE:
            Set oHl7Constructor.LabError = New Collection
            Set oHl7Constructor.LabWarning = New Collection
        Case hlRAD_ORDERS_CODE:
            Set oHl7Constructor.RadError = New Collection
            Set oHl7Constructor.RadWarning = New Collection
        End Select
        mvarSessionError = hlERRORS.hlERROR_NONE
End Sub
Public Function GetOrderTypeTextfromEnum(lngEnumVar As hlORDERTYPECODE) As String
    GetOrderTypeTextfromEnum = oHl7Constructor.GetOrderTypeTextfromEnum(lngEnumVar)
End Function
Public Sub CopyWarningsFromSets(OrderType As hlORDERTYPECODE, ByRef PassedCol As Collection)
    Dim cRxWarningObject As RxWarningType
    Dim ccRxWarningObject As RxWarningType
    Dim cRadWarningObject As RadWarningType
    Dim ccRadWarningObject As RadWarningType
    Dim cLabWarningObject As LabWarningType
    Dim ccLabWarningObject As LabWarningType
            
    Select Case OrderType
        Case hlRX_ORDERS_CODE:
            Set oHl7Constructor.RxWarning = New Collection
            For Each cRxWarningObject In PassedCol
                 Set ccRxWarningObject = cRxWarningObject.CopyWarning
                 oHl7Constructor.RxWarning.Add ccRxWarningObject
            Next
        Case hlLAB_ORDERS_CODE:
            Set oHl7Constructor.LabWarning = New Collection
            For Each cLabWarningObject In PassedCol
                 Set ccLabWarningObject = cLabWarningObject.CopyWarning
                 oHl7Constructor.LabWarning.Add ccLabWarningObject
            Next
        Case hlRAD_ORDERS_CODE:
            Set oHl7Constructor.RadWarning = New Collection
            For Each cRadWarningObject In PassedCol
                 Set ccRadWarningObject = cRadWarningObject.CopyWarning
                 oHl7Constructor.RadWarning.Add ccRadWarningObject
            Next
  
    End Select
End Sub

Private Sub CheckIfUserIsProvider()

    gbUserIsProvider = False
    'SCR 22147 3-18-02 AKM
    'If cwobjCWCommSession.CmdLineSwitch("NPOE") = "NPOE" Then 'Switch is there let NPOE show
        Dim lFacility_NCID As Long
        lFacility_NCID = gobjLogonEnvironment.FacilityNCID
        mvarSignatureClass = moeSearchEngine.GetSignatureClass(oHl7Constructor.szProviderIEN, lFacility_NCID)
        oHl7Constructor.SignatureClass = mvarSignatureClass
        If mvarSignatureClass > 1 Then
            gbUserIsProvider = True
        End If
      ' Set mobjProviderSearch = New CHCSII_NonProvider.clsNonProvider
      ' Set mRstSearchedProviders = New ADODB.Recordset
     '  mobjProviderSearch.OpenNonProvider cwobjLogonEnviron, cwobjCWCommSession, gobjUser
     '  mobjProviderSearch.getProviders mRstSearchedProviders
     '  If mRstSearchedProviders.State <> adStateClosed Then
     '      If Not mRstSearchedProviders.EOF And Not mRstSearchedProviders.BOF Then
     '          mRstSearchedProviders.MoveFirst
     '          If mRstSearchedProviders.Fields("IEN").Value = oHl7Constructor.szProviderIEN Then
     '              gbUserIsProvider = True
     '          End If
     '          mRstSearchedProviders.MoveFirst
     '      End If
     '  End If
    'Else
    
'        Dim szRoleName As String
'        Dim i As Integer
'        With gobjUser
'            For i = 1 To .Roles.Count
'                szRoleName = LCase$(.Roles.Item(i).Name)
'                If InStr(szRoleName, "provider") > 0 Then
'                    gbUserIsProvider = True
'                    Exit For
'                End If
'                If InStr(szRoleName, "independent medic") > 0 Then
'                    gbUserIsProvider = True
'                    Exit For
'                End If
'            Next i
'        End With
'    End If
End Sub
Public Function GetSignatureClass() As Integer
   GetSignatureClass = mvarSignatureClass
End Function

Public Function HandleSecondaryWarnings(OrderType As hlORDERTYPECODE, OrderAction As hlACTION, WorkingTablet As Object, bOrderIDProvided As Boolean, orderid As String, lReturnErr As Long) As Boolean
    HandleSecondaryWarnings = False
    
    Dim bErr50MoreWarnings As Boolean
    If WorkingTablet.WarningOverriden Then     'Indicates endorsement (Reason for Overriding Warnings)
        If bOrderIDProvided = True And (OrderAction = hlACTION_NEW Or OrderAction = hlACTION_MODIFY Or OrderAction = hlACTION_RELEASE Or OrderAction = hlACTION_RENEW Or OrderAction = hlACTION_RENEW_MODIFY) Then
        
            lReturnErr = oHl7Constructor.SendOverrides(OrderType)
            Screen.MousePointer = vbDefault
            Call ShowCHCSErrors(WorkingTablet, OrderType, False, bErr50MoreWarnings, lReturnErr)
            If lReturnErr = hlERROR_PERFORM_ACTION_FAILED Then
                 HandleSecondaryWarnings = True
                 Exit Function
            End If
        Else      'No Order-ID exists at this point ... proceed
            If OrderAction = hlACTION_NEW Then
              lReturnErr = oHl7Constructor.SendOrder(OrderType)
              Screen.MousePointer = vbDefault
              Call ShowCHCSErrors(WorkingTablet, OrderType, False, bErr50MoreWarnings, lReturnErr)
              If lReturnErr = hlERROR_PERFORM_ACTION_FAILED Then
                  HandleSecondaryWarnings = True
                 Exit Function
              End If
            ElseIf OrderAction = hlACTION_PREVERIFY Or OrderAction = hlACTION_PRESET Then
              ''No Action required ... Supplying a reason means we can proceed smoothly ...
            Else
'                'SCR 32369
                If OrderType = hlRX_ORDERS_CODE Then
                    lReturnErr = oHl7Constructor.SendOverrides(OrderType)
                Else
                    lReturnErr = oHl7Constructor.PerformAction(WorkingTablet.orderid, OrderAction, OrderType)
                End If
                Screen.MousePointer = vbDefault
            End If
            Call ShowCHCSErrors(WorkingTablet, OrderType, False, bErr50MoreWarnings, lReturnErr)
            
        End If
        If lReturnErr <> hlERROR_NONE Then
            If lReturnErr = hlERROR_TIMED_OUT Then
                If ReconnectToCHCS <> hlERROR_NONE Then
            
                    HandleSecondaryWarnings = True
                    Exit Function
                End If
            Else
                MsgBxARMd oHl7Constructor.GetErrorText(lReturnErr), vbCritical
                HandleSecondaryWarnings = True
                Exit Function
            End If
        End If
    Else
'        If Len(cwobjCWCommSession.CmdLineSwitch("PDTS")) > 0 Then
                
            'if it is a drug order than need to send cancel order with outcome code
            
            If OrderType = hlRX_ORDERS_CODE And OrderAction <> hlACTION_PREVERIFY Then
                lReturnErr = oHl7Constructor.SendPDTSCancel(RXTablet.WarningPDTSOutcome, RXTablet.WarningReason, RXTablet.WarningHeader, RXTablet.orderid)
                If WorkingTablet.WarningPDTSOutcome = hlPDTSOutcomes_DiffQty Then
                    lReturnErr = hlERROR_USER_CANCELLED_NEWQTY
                ElseIf WorkingTablet.WarningPDTSOutcome = hlPDTSOutcomes_DiffDirections Then
                    lReturnErr = hlERROR_USER_CANCELLED_NEWSIG
                ElseIf WorkingTablet.WarningPDTSOutcome = hlPDTSOutcomes_DiffDose Then
                    lReturnErr = hlERROR_USER_CANCELLED_NEWSIG
                Else
                    lReturnErr = hlERROR_USER_CANCELLED
                End If
                
                Call ShowCHCSErrors(WorkingTablet, OrderType, False, bErr50MoreWarnings, lReturnErr)
                
                
            Else
                lReturnErr = hlERROR_USER_CANCELLED
 
            End If
                
                
                
                
            HandleSecondaryWarnings = True
            Exit Function
    End If
 
    HandleSecondaryWarnings = False
End Function


Public Function GetDefaultCancelOption(OrderType As hlORDERTYPECODE) As hlPDTSOUTCOMES
    
    GetDefaultCancelOption = hlPDTSOutcomes_PrescriptionCancelled
    
    Dim pWarningCollection As Collection
    
    Select Case OrderType
        Case hlRX_ORDERS_CODE
            Set pWarningCollection = oHl7Constructor.RxWarning
       
        Case hlSiGN_ORDERS_CODE
            Set pWarningCollection = oHl7Constructor.colUnsignedOrder(oHl7Constructor.colUnsignedOrder.WarningIndex).WarningCollection
            OrderType = oHl7Constructor.colUnsignedOrder(oHl7Constructor.colUnsignedOrder.WarningIndex).OrderType
            If OrderType <> hlRX_ORDERS_CODE Then
                Exit Function
            End If
    End Select
    If pWarningCollection.Count > 1 Then
        GetDefaultCancelOption = hlPDTSOutcomes_PrescriptionCancelled
        Exit Function
    End If
    Select Case pWarningCollection(1).WarningTypeCodePDTS
        Case hlPDTSWarning_Drug_Interactions: GetDefaultCancelOption = hlPDTSOutcomes_PrescriptionCancelled
        Case hlPDTSWarning_Therapeutic_Duplication: GetDefaultCancelOption = hlPDTSOutcomes_PrescriptionCancelled
        Case hlPDTSWarning_Excessive_Dose: GetDefaultCancelOption = hlPDTSOutcomes_DiffDose
        Case hlPDTSWarning_Insufficient_Dose: GetDefaultCancelOption = hlPDTSOutcomes_DiffDose
        Case hlPDTSWarning_Over_Utilization: GetDefaultCancelOption = hlPDTSOutcomes_PrescriptionCancelled
        Case hlPDTSWarning_Under_Utilization: GetDefaultCancelOption = hlPDTSOutcomes_DiffDirections
        Case hlPDTSWarning_Duration_Of_Therapy: GetDefaultCancelOption = hlPDTSOutcomes_DiffDirections
        Case hlPDTSWarning_Local:
                    If pWarningCollection(1).WarningTypeCode = "5" Then
                        GetDefaultCancelOption = hlPDTSOutcomes_DiffDose
                    Else
                        GetDefaultCancelOption = hlPDTSOutcomes_PrescriptionCancelled
                    End If
    End Select
    
End Function

Public Function GetOverrideOptionFrames(OrderType As hlORDERTYPECODE) As hlRXWARNINGOVERRIDESTYPE
    
    'GetOverrideOptionFrames = hlRXWarningsOverride_DrugDrug
    
    Dim pWarningCollection As Collection
    Dim iCount As Integer
    Dim lResult As hlRXWARNINGOVERRIDESTYPE
    Dim lPrevResult As hlRXWARNINGOVERRIDESTYPE
    
    
    Select Case OrderType
        Case hlRX_ORDERS_CODE
            Set pWarningCollection = oHl7Constructor.RxWarning
       
        Case hlSiGN_ORDERS_CODE
            Set pWarningCollection = oHl7Constructor.colUnsignedOrder(oHl7Constructor.colUnsignedOrder.WarningIndex).WarningCollection
            OrderType = oHl7Constructor.colUnsignedOrder(oHl7Constructor.colUnsignedOrder.WarningIndex).OrderType
            If OrderType <> hlRX_ORDERS_CODE Then
                Exit Function
            End If
    End Select
    
    iCount = 1
    Do
        Select Case pWarningCollection(iCount).WarningTypeCodePDTS
            Case hlPDTSWarning_Drug_Interactions: lResult = hlRXWarningsOverride_DrugDrug
            Case hlPDTSWarning_Therapeutic_Duplication: lResult = hlRXWarningsOverride_DrugDrug
            Case hlPDTSWarning_Excessive_Dose: lResult = hlRXWarningsOverride_Dosage
            Case hlPDTSWarning_Insufficient_Dose: lResult = hlRXWarningsOverride_Dosage
            Case hlPDTSWarning_Over_Utilization: lResult = hlRXWarningsOverride_Dosage
            Case hlPDTSWarning_Under_Utilization: lResult = hlRXWarningsOverride_Dosage
            Case hlPDTSWarning_Duration_Of_Therapy: lResult = hlRXWarningsOverride_Dosage
            Case hlPDTSWarning_Local:
                If pWarningCollection(iCount).WarningTypeCode = "5" Then
                    lResult = hlRXWarningsOverride_Dosage
                Else
                    lResult = hlRXWarningsOverride_DrugDrug
                End If
        End Select
        
        If iCount = 1 Then
            lPrevResult = lResult
        Else
            If lResult <> lPrevResult Then
                lResult = hlRXWarningsOverride_All
                Exit Do
            End If
        End If
        iCount = iCount + 1
    Loop Until iCount > pWarningCollection.Count
    
    GetOverrideOptionFrames = lResult
    
End Function

Public Sub SubmitAllInQueue()

    Dim RadQ As RadTransaction
    Dim RxQ As RxTransaction
    Dim LabQ As LabTransaction
    Dim i As Integer
    Dim lReturnErr As hlERRORS
    
    'OrderAction = hlACTION_NEW
    For i = 1 To goRadTransactionCol.Count
        If Left$(goRadTransactionCol.Item(i).NewOrderid, 1) = "Q" Then
            Set goOESession.RadTablet = goRadTransactionCol.Item(i).RadTablet.CopyTablet
            
            lReturnErr = oHl7Constructor.SendOrder(hlRAD_ORDERS_CODE)
            TrapWarningAndErrorForSubmitAll hlRAD_ORDERS_CODE, goOESession.RadTablet
        End If
        
        
    Next

    
'    For Each RxQ In goRxTransactionCol
'    Next
'
'    For Each LabQ In goLabTransactionCol
'    Next

End Sub

Private Sub TrapWarningAndErrorForSubmitAll(OrderType As hlORDERTYPECODE, WorkingTablet As Object)
    Dim pszTempA As String
    Dim szErrMsg As String
    Dim bErr50MoreWarnings As Boolean
    Dim piReturn As Integer
    Dim iErrCount As Integer
    Dim lReturnErr As hlERRORS
    Dim bOrderIDProvided As Boolean
    Dim orderid As String
    
    If GetHl7Errors(OrderType, szErrMsg, iErrCount) Then
        bErr50MoreWarnings = bAnyErr50(OrderType)
        If WorkingTablet.ErrorIsMessage = False Then                          'added 9-2-99; AAB
            If Not bErr50MoreWarnings Then                                    'ERROR 50 is always special case; added 9-13-99 AKM
                piReturn = IsError_Recoverable(OrderType, hlACTION_NEW)
                If piReturn = giERRORS_ARE_CHOICE Then       'YES Then
                    pszTempA = vbCrLf & vbCrLf & "Enter <YES> to CORRECT order, or <NO> to CLEAR order"   ''& & vbCrLf
                    piReturn = MsgBxARMd(szErrMsg & pszTempA, vbExclamation + vbYesNo + vbDefaultButton1, "CHCS I Errors")
                    If piReturn = vbYes Then
                        lReturnErr = hlERROR_SOFT_ERROR                      'caller can tell if this allows the user to correct and retry
                    Else
                        lReturnErr = hlERROR_PERFORM_ACTION_FAILED       'added 9-2-99; AAB
                    End If
                                                        
                    AskCoreToRefreshAfterModal OrderType
                                                  
                ElseIf piReturn = giERRORS_ARE_HARD Then
                    MsgBxARMd szErrMsg, vbCritical, "CHCS I Errors"
                    AskCoreToRefreshAfterModal OrderType
                    lReturnErr = hlERROR_PERFORM_ACTION_FAILED
                Else      'giERRORS_ARE_SOFT
                                                    
                    MsgBxARMd szErrMsg, vbExclamation, "Please correct and re-submit"
                    AskCoreToRefreshAfterModal OrderType
                    lReturnErr = hlERROR_SOFT_ERROR
                End If
                                        
                GoTo ExitFunctionJmp:
                                        '================================
            Else
                                        'ERROR-50 Detected:   '6-01-01; AAB #11978
                If GetHl7Warnings(OrderType, bOrderIDProvided) Then
                    If HandleSecondaryWarnings(OrderType, hlACTION_NEW, WorkingTablet, bOrderIDProvided, orderid, lReturnErr) Then
                        GoTo ExitFunctionJmp:
                    Else
                        lReturnErr = hlERROR_NONE
                                            
                    End If
                End If
            End If
        End If
    End If
ExitFunctionJmp:

End Sub

Private Sub RemoveFromOrderList(OrderType As hlORDERTYPECODE, Optional orderid As String)
    Dim i As Integer
    Dim j As Integer
    Dim fActive As Form
    Dim itmX As ListItem
    
    For i = 1 To goOESession.SummaryCollection.Count
        
        If goOESession.SummaryCollection(i).orderid = orderid Then

            Select Case goOESession.SummaryCollection(i).OrderTypeCode
                Case hlRX_ORDERS_CODE: Set fActive = Frm_OE_RX
                Case hlLAB_ORDERS_CODE: Set fActive = Frm_OE_Lab
                Case hlRAD_ORDERS_CODE: Set fActive = Frm_OE_Rad
                Case Else
                    Exit For
            End Select
            For j = 1 To fActive.Lst_OrdersBottom.ListItems.Count
                If fActive.Lst_OrdersBottom.ListItems(j).SubItems(colOrderID(OrderType)) = orderid Then
                    fActive.Lst_OrdersBottom.ListItems.Remove j
                    Exit For
                    
                End If

            Next j
            Exit For
        End If
    Next i

End Sub
