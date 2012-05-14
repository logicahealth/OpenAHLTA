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
Attribute VB_Name = "Hl7Constructor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit


'Codes are both Return-STATUS and COMMANDS
'=========================================

'-- for HL7CommStatus codes; 6-25-96-AAB
Public Enum hlCOMMSTATUS
    hlCOMM_NULL = 0        'nothing
    hlCOMM_OKCOMM = 1      'GOOD MESSAGE (ACK)
    hlCOMM_RELOGGED = 2    'NC was successfully relogged for us
    
    '-- These may allow continuation after returning
    hlCOMM_BLOCKED = 4     'steering flag prohibited us from processing
    hlCOMM_PACKET_ID = 8   'Packet IDs dont match ... (stacked outbound messages)
    
    '-- These require some recovery action
    hlCOMM_QUIT = 32       'hit CANCEL at logon screen
    hlCOMM_NOCOMM = 64     'dropped line; unable to restart
    hlCOMM_SOCKET = 128    'PARSE error or MsgHdlr SOCKET ... failure
    hlCOMM_TIMEOUT = 256   'WatchDog Timer has gone off ... 7-19-96-AAB
    
    '-- This helps conditional error returns at ProcessRcvdMsg()
    hlCOMM_RELOGON_YES = 1024    'SendHL7Msg will Relogon before returning
    hlCOMM_RELOGON_NO = 2048     'SendHL7Msg will NOT Relogon before returning
End Enum

' CLT 7/23/99
'-- Tell SendHL7Msg that if ReLogon rqd, then rexmit failed message Y/N
Public Enum hlCOMMRESEND
    hlCOMM_RESEND_YES = 4096
    hlCOMM_RESEND_NO = 0
End Enum

'for login
Public Enum hlLOGIN
    hlUNENCRYPTED = 1
    hlENCRYPTED = 0
End Enum


Public Enum hlLABSUBSCRIPT
    hlAUTOPSY = 0
    hlBONEMARROW = 1
    hlCYTGYN = 2
    hlCYTNONGYN = 3
    hlSURGPATH = 4
    hlCHEMISTRY = 5
    hlOther = 6
End Enum

Public Enum hlCHCS_Boolean
    hlyes = 0
    hlNo = 1
    hlNull = 2
End Enum

'For OrderTypeCode
Public Enum hlORDERTYPECODE
    hlRX_ORDERS_CODE = 9
    hlLAB_ORDERS_CODE = 4
    hlRAD_ORDERS_CODE = 5
    hlCON_ORDERS_CODE = 30
    hlSiGN_ORDERS_CODE = 99
    hlALL_ORDERS_CODE = 0
End Enum

'Used for ALL priorities
Public Enum hlLABPRIORITY
    hlLAB_PRIORITY_ROUTINE = 1
    hlLAB_PRIORITY_STAT = 2
    hlLAB_PRIORITY_ASAP = 3
    hlLAB_PRIORITY_NOTIFY = 4
    hlLAB_PRIORITY_PREOP = 5
End Enum

Public Enum hlLABPRIORITYSTAT
    hlLAB_PRIORITY_STAT_NO = False
    hlLAB_PRIORITY_STAT_YES = True
End Enum

'For RAD Portable codes
' CLT 7/28/99
Public Enum hlRADPORTABLE
    hlPORTABLE_ALWAYS = 1
    hlPORTABLE_SOMETIMES = 2
    hlPORTABLE_NEVER = 3
End Enum

'Mark if an ERROR is actual 'error' or a 'message'
Public Enum hlERRMSGFLAG
    hlERROR_IS_MESSAGE = True
    hlERROR_IS_ERROR = False
End Enum

'Preverify RX (RXO-16) is a 'valid' flag
Public Enum hlRXVALID
    hlDRUG_IS_VALID = True   'used as true for boolean rxtype.drugisvalid
    hlDRUG_IS_NOT_VALID = False 'used as false for boolean rxtype.drugisvalid
End Enum

'For 'dispense full' drug flag
Public Enum hlDISPENSEFULL
    hlDISPENSE_FULL_YES = 1
    hlDISPENSE_FULL_NO = 2
End Enum

'For Approvals, Appointments, etc.
Public Enum hlAPPROVALREQUIRED
    hlAPPROVAL_REQUIRED = True
    hlAPPROVAL_NOT_REQUIRED = False
End Enum

Public Enum hlAPPOINTMENTREQUIRED
    hlAPPOINTMENT_REQUIRED = True
    hlAPPOINTMENT_NOT_REQUIRED = False
End Enum

Public Enum hlCOUNTERSIGNREQUIRED
    hlCOUNTERSIGN_REQUIRED = True
    hlCOUNTERSIGN_NOT_REQUIRED = False
End Enum

Public Enum hlCONTINUOUSALLOWED
    hlCONTINUOUS_ALLOWED = True
    hlCONTINUOUS_NOT_ALLOWED = False
End Enum

'Allergy Override Required Flag
Public Enum hlALLERGYREQUIRED
    hlallergy_required = True
    hlALLERGY_NOT_REQUIRED = False
End Enum

'For RX Child Cap; corrected final 4-13-98
'These are for Sending to CHCS
Public Enum hlCHILDCAP
    hlCHILD_CAP_YES = 1
    hlCHILD_CAP_NO = 0
End Enum

'Rad WET reading
Public Enum hlWETREADING
    hlWET_READING_YES = True
    hlWET_READING_NO = False
End Enum

'For tailoring GetOrders
Public Enum hlSHOWNEWORDER
    hlSHOW_NEW_ORDERS = True
    hlDONT_SHOW_NEW_ORDERS = False
End Enum

Public Enum hlSHOWSYSTEMORDERS
    hlSHOW_SYSTEM_ORDERS = True
    hlDONT_SHOW_SYSTEM_ORDERS = False
End Enum

'Mixed RAD and LAB values; sent to DLL from application:
Public Enum hlRADLAB
    hlONETIME = 1
    hlCONTINUE = 2
End Enum

'Sort BY:
' CLT 7/28/99
Public Enum hlSORTBY
    hlNO_SORT = 1
    hlTYPE_SORT = 2
    hlDATE_SORT = 3
    hlDESCRIP_SORT = 4
    hlSTATUS_SORT = 5
End Enum

'Sort HOW:
' CLT 7/28/99
Public Enum hlSORTHOW
    hlNO_SORT_ORDER = 1
    hlASCEND_SORT_ORDER = 2
    hlDESCEND_SORT_ORDER = 3
End Enum


Public Enum hlMODIFIABLE
    hlIS_MODIFIABLE = 0
    hlIS_NOT_MODIFIABLE = 1
End Enum

'Rad Mobility Selections (transport mode)
' CLT 7/28/99
Public Enum hlRADXPORT
    hlRAD_XPORT_AMBULATE = 1
    hlRAD_XPORT_WHEELCHAIR = 2
    hlRAD_XPORT_STRETCHER = 3
    hlRAD_XPORT_PORTABLE = 4
End Enum

'Allergies: new 3-17-98
' CLT 7/28/99
Public Enum hlALLERGY
    hlALLERGY_ADD = 1
    hlALLERGY_EDIT = 2
    hlALLERGY_DELETE = 3
End Enum

Public Enum hlLABCOLLECT
    hlLAB_COLLECT_SEND = 1
    hlLAB_COLLECT_CLINIC = 2
End Enum

'Rad Portable Requested
' CLT 7/28/99
Public Enum hlPORTABLEREQUESTED
    hlPORTABLE_REQUESTED_YES = 1
    hlPORTABLE_REQUESTED_NO = 0
End Enum

'Perform Action Codes
' CLT 7/28/99
Public Enum hlACTION

     hlACTION_DETAIL = 1
               'redefined 2-15-01; AAB
               ''hlACTION_DELETE = 2
               ''hlACTION_MODIFY = 4
               ''hlACTION_RENEW = 5
               ''hlACTION_NEW = 7
               ''hlACTION_RENEW_MODIFY = 10
               ''hlACTION_ORDER_SET = 11
     hlACTION_DELETE = 44
     hlACTION_MODIFY = 45
     hlACTION_RENEW = 46
     hlACTION_NEW = 47
     hlACTION_RENEW_MODIFY = 48
     hlACTION_ORDER_SET = 49
     hlACTION_ADD_TO_QUEUE = 58

     hlACTION_HOLD = 3
     hlACTION_RELEASE = 64
     hlACTION_PREVERIFY = 8
     hlACTION_PREMODIFY = 9
     hlACTION_MED_LIST_MODIFY = 12
     hlACTION_MED_LIST_RENEW = 13
     hlACTION_EXPAND_SIG = 14                       'added 05-10-00; AAB
     ''hlACTION_NONE = 0                            'added 09-22-01; AAB
     hlACTION_PRESET = 15                           'added 11-12-01; AKM
     hlACTION_RESET_ORDERING_PROVIDER = 16
     
     
End Enum

Public Enum hlALLERGYACTION
  hlAllergyAction_Add = 0
  hlAllergyAction_Edit = 1
  hlAllergyAction_Delete = 2
End Enum

Public Enum hlPDTSWARNING
  hlPDTSWarning_Drug_Interactions = 0
  hlPDTSWarning_Therapeutic_Duplication = 1
  hlPDTSWarning_Excessive_Dose = 2
  hlPDTSWarning_Insufficient_Dose = 3
  hlPDTSWarning_Over_Utilization = 4
  hlPDTSWarning_Under_Utilization = 5
  hlPDTSWarning_Duration_Of_Therapy = 6
  hlPDTSWarning_Local_Drug_Allergies = 7
  hlPDTSWarning_Other_Drug_Warnings = 8
  hlPDTSWarning_ExcessiveUnitsDaysSupply = 9
  hlPDTSWarning_Local = 10
End Enum

Public Enum hlPDTSOUTCOMES
  hlPDTSOutcomes_False_Positive = 0
  hlPDTSOutcomes_DiffDose = 1
  hlPDTSOutcomes_DiffDirections = 2
  hlPDTSOutcomes_DiffDrugs = 3
  hlPDTSOutcomes_DiffQty = 4
  hlPDTSOutcomes_PrescriberApproved = 5
  hlPDTSOutcomes_PrescriptionCancelled = 6
End Enum

Public Enum hlRXWARNINGOVERRIDESTYPE
  hlRXWarningsOverride_DrugDrug = 0
  hlRXWarningsOverride_Dosage = 1
  hlRXWarningsOverride_All = 2
  hlRXWarningsOverride_Other = 3
  hlRXWarningsOverride_None = -1
End Enum


Public Enum hlSingleSpecimenIsComplete
     hlSingleSpecimen_Complete = 1
     hlSingleSpecimen_NotComplete = 2
End Enum

Public Enum hlSIGNBY
    hlRETURNBYPROVIDER = 0
    hlRETURNBYPATIENT = 1
End Enum

Public Enum hlSIGNNUMBER
    hlRETURNBATCH = 0
    hlRETURNALL = 1
End Enum

Public Enum hlORDERSIGNACTION
    hlSIGNED = 0
    hlHOLD = 1
    hlNEW = 2
    hlCANCEL = 3
    hlModIV = 4
    hlSUSPEND = 18
End Enum

Public Enum hlORDERSIGNCURSOR
    hlForward = 0
    hlBack = 1
    hlStart = 2
End Enum

Public Enum hlSIGNATURENEEDED
    hlNoSigNeeded = 0
    hlSigNeeded = 1
End Enum

Public Enum hlSIGNATURECOMPLETION
    hlCOMPLETED = 0
    hlCANCELLED = 1
    hlAUTOMATICALLY_DISCONTINUED = 2
    hlCANCELLED_NEEDS_SIGNATURE = 3
    hlRESULTED_NEEDS_SIGNATURE = 4
    hlBLANK = -1
End Enum

Public oHL7Server As HL7ComMsg
Public goPatient As PatientType
Public gszErrMsg As String
Public Specimen As Collection
Public Summary As Collection
Public Allergy As Collection
Public colUnsignedOrder As UnsignedOrders
 
Public RXTablet As New RxType
Public LabTablet As New LabType
Public RadTablet As New RadType
Public OtherTablet As New BaseTablet

Public RxError As Collection
Public LabError As Collection
Public RadError As Collection
Public OtherError As Collection
Public SignError As Collection

Public RxWarning As Collection
Public LabWarning As Collection
Public RadWarning As Collection
Public OtherWarning As Collection
Public SignWarning As Collection

'local variable(s) to hold property value(s) =============================

Private mvarszClinicIen As String
Private mvarszMeprsIen As String
Private mvarszMeprsCode As String
Private mvarszMTFCode As String
Private mvarszMTFIEN As String
Private mvarszApptIEN As String 'SCR 31315
Private mszLockedPageID As String
Private mbPreModifying As Boolean
Private mvarszOrderContinuation As String

Private mvarDefOrderPage As String      'added 02-22-00; AAB for CT work
Private mvarSignatureClass As Integer

Private COLL_ZQRs As Collection     'added 1-23-02 for RqdComments Modify; AAB
Private COLL_RESPONSE As RequiredCommentsCollection
  
Public Property Get szDefOrderPage() As String
     'created 02-22-00; AAB
     szDefOrderPage = mvarDefOrderPage
End Property



Public Function LoginAndSelectPatient(szAccessCode As String, szVerifyCode As String, bEncrypted As hlLOGIN, _
                          szCHCSIPAddress As String, szCHCSPort As String, szLocalIP As String, _
                          szClinicIEN As String, szMeprsCode As String, szMeprsIen As String, szMTFCode As String, _
                          szPatientIEN As String, szPatientName As String) As hlERRORS
                          
                       
    Dim szEncAccessCode As String
    Dim szEncVerifyCode As String
    Dim retString As String
    'KELLYMCDim szMsgBoxText As String
    
    mszAccessCode = szAccessCode
    mszVerifyCode = szVerifyCode
    
    If LenB(mszAccessCode) = 0 Or LenB(mszVerifyCode) = 0 Then
      LoginAndSelectPatient = hlERRORS.hlERROR_ACCESSVERIFY_CODE
      Exit Function
    End If
    
    If LenB(szCHCSIPAddress) = 0 Then
      LoginAndSelectPatient = hlERRORS.hlERROR_CHCSIPADDRESS_NOT_SET
      Exit Function
    End If
    oHL7Server.SetCHCSAddress (szCHCSIPAddress)
    
    If LenB(szCHCSPort) = 0 Then
      LoginAndSelectPatient = hlERRORS.hlERROR_CHCSPORT_NOT_SET
      Exit Function
    End If
    oHL7Server.SetCHCSPort (szCHCSPort)
    
    If LenB(szLocalIP) = 0 Then
      LoginAndSelectPatient = hlERRORS.hlERROR_CLIENTIPADDRESS_NOT_SET
      Exit Function
    End If
    gszLocalIP = szLocalIP
    
    If bEncrypted = hlLOGIN.hlUNENCRYPTED Then
      szEncAccessCode = szEncrypt(UCase$(mszAccessCode))
      szEncVerifyCode = szEncrypt(mszVerifyCode)
    Else
      szEncAccessCode = mszAccessCode
      szEncVerifyCode = mszVerifyCode
    End If
    retString = Space$(2048)

    'Screen.MousePointer = vbHourglass
       oHL7Server.SetAccessCode (szEncAccessCode)
       oHL7Server.SetVerifyCode (szEncVerifyCode)
       oHL7Server.Login (retString)
       gszHL7RetCode = szGetSegSeq(retString, 1, "MSA", hlMSASeg.AckCode)
       gszHL7RetMsg = retString
       ShowHL7InboundDebugMessages
       
               ''Screen.MousePointer = vbDefault
    
    Select Case gszHL7RetCode
    Case "AA"   ' You are good to go
      LoginAndSelectPatient = hlERRORS.hlERROR_NONE
      gbHL7LoginStatus = True
      
    Case "AE"   ' user error
      LoginAndSelectPatient = hlERROR_ACCESS_VERFIY_IP_NOT_VALID
      GoTo FAILURE
      
    Case "AR"   ' HL7 message format error
      LoginAndSelectPatient = hlERROR_LOGIN_MSG_FORMAT_ERROR
      GoTo FAILURE
        
    Case "AC"
      LoginAndSelectPatient = hlERROR_LOGIN_COMM_ERROR
      GoTo FAILURE
           
    Case Else   ' Booo         "Your message ACK response is not known."
      LoginAndSelectPatient = hlERROR_LOGIN_UNKNOWN_ERROR
      GoTo FAILURE
    
    End Select
    
Dim szTemp As String

    szTemp = Space$(200)
    oHL7Server.GetUserIEN (szTemp)
    gszUserIEN = szGetColumn(szTemp, 1, "\")
    
    szTemp = Space$(200)
    oHL7Server.GetProviderIEN (szTemp)
    gszProviderIEN = szGetColumn(szTemp, 1, "\")
    If Left$(gszProviderIEN, 1) = vbNullChar Then
        gszProviderIEN = ""
    End If
    
    If gszProviderIEN = "" Then
        LoginAndSelectPatient = hlERROR_LOGIN_NOT_A_PROVIDER
        GoTo FAILURE
    End If
    'PR-286
    gszOrderingProviderIEN = gszProviderIEN ' initially set the entering id to the same as provider id
   
   ''GetClinicInfo
    
     '==================================
    'Added new 6-25-98; AAB
    'reset primary objects after logon
    
    Set goPatient = New PatientType
    

    HL7CommStatus = hlCOMMSTATUS.hlCOMM_OKCOMM
    '==================================
              
'    Dim hlReturn As hlERRORS
'
'    hlReturn = Login(szAccessCode, szVerifyCode, bEncrypted, szCHCSIPAddress, szCHCSPort, szLocalIP)
'    If hlReturn <> hlERROR_NONE Then
'        LoginAndSelectPatient = hlReturn
'        Exit Function
'    End If
'
          mvarszClinicIen = szClinicIEN
          clinic.IEN = szClinicIEN
          mvarszMeprsCode = szMeprsCode
          
          mvarszMeprsIen = szMeprsIen
          clinic.MEPRS_STRING = "\" & mvarszMeprsIen & "\" & mvarszMeprsCode & "\99MEP"
          mvarszMTFCode = szMTFCode
          clinic.MTF_CODE = szMTFCode
          FormatPatientSelectMsg szPatientIEN, szPatientName
    
     'SEND HL7 command package >>>>>>>>>>>>>>>
          glngTimeoutStatus = HL7MsgHandler()  'changed 6-30-98;AAB new design
    
          If glngTimeoutStatus = hlCOMMSTATUS.hlCOMM_TIMEOUT Then
               LoginAndSelectPatient = hlERRORS.hlERROR_TIMED_OUT
               GoTo PatSelFailed:
                    ''Exit Function
          Else
               'added 4-7-98; AAB
               If gszHL7RetCode = "AA" Then
                    LoginAndSelectPatient = hlERRORS.hlERROR_NONE
               Else
                    LoginAndSelectPatient = hlERRORS.hlERROR_PATIENT_SELECT_FAILED
                    GoTo PatSelFailed:
                         ''Exit Function
               End If
          End If
    
          goPatient.ParsePID gszHL7RetMsg
    
    '4-6-98; AAB this is a "text quality" copy of the name, NOT PID(5)
          Patient.Name = goPatient.Name
    
    'Want HL7-ready copy of name
          Patient.HL7Name = goPatient.HL7Name
          Patient.IEN = goPatient.IEN
          
Debug.Print "get Default Order Page"

    'Get the default orders page from ZPW/Z11 return with ZOP3=3 (may be more than one ZOP returned)
Dim hlTemp As hlERRORS
    
          hlTemp = szGetDefOrderPage(szPatientIEN)
          If hlTemp <> hlERROR_NONE Then LoginAndSelectPatient = hlTemp
          
          Exit Function
                'LoginAndSelectPatient = SendPatientSelectMsg(szPatientIEN, szPatientName)
          Exit Function
    
FAILURE:

    gbHL7LoginStatus = False
    gszHL7RetCode = ""
    HL7CommStatus = hlCOMMSTATUS.hlCOMM_NOCOMM
    
PatSelFailed:       'added 6-30-98; AAB

    gszDefOrderPage = ""
    Patient.Name = ""
    Patient.HL7Name = ""
    Patient.IEN = ""
    
             
End Function

  
Public Function GetAllergies() As hlERRORS
    
Dim pszAllergyCount As Integer

    If gbHL7LoginStatus = False Then
      GetAllergies = hlERRORS.hlERROR_NOT_LOGGED_IN
      Exit Function
    End If
    
    GetAllergies = hlERRORS.hlERROR_NONE

    ' MSH
    ReDim seginfo(1 To hlMSHSeg.NumberOfSegs) As String
      seginfo(hlMSHSeg.MsgType) = "ZPW\Z19"
    gszHL7SndMsg = szAddSeg("MSH", seginfo())

    ' PID
    ReDim seginfo(1 To hlPIDSeg.NumberOfSegs) As String
      seginfo(hlPIDSeg.SetID) = "1"
      seginfo(hlPIDSeg.IEN) = szFormatPIDSegIEN()
      seginfo(hlPIDSeg.Name) = Patient.HL7Name
    gszHL7SndMsg = gszHL7SndMsg & szAddSeg("PID", seginfo())
 
     'SEND HL7 command package >>>>>>>>>>>>>>>
     glngTimeoutStatus = HL7MsgHandler()  'changed 6-30-98;AAB new design
    
     If glngTimeoutStatus = hlCOMMSTATUS.hlCOMM_TIMEOUT Then
          GetAllergies = hlERRORS.hlERROR_TIMED_OUT
          Exit Function
     Else
          If gszHL7RetCode = "AA" Then
            GetAllergies = hlERRORS.hlERROR_NONE
          Else
            GetAllergies = hlERRORS.hlERROR_GET_ALLERGIES_FAILED
            Exit Function
          End If
     End If
  
'===================================================================
 
    'returns ZOAs
    '
    'zoa1 = allergy ien\name [len 60]
    'zoa2 = comment [len 80]
    'zoa3 = update message code [len 1]; (A=add E=edit D=delete)
    
    pszAllergyCount = iGetSegCount(gszHL7RetMsg, "ZOA")
            
    'ZOA^11084010\STRESS\99DAL^^^
    'ZOA^14047001\BACITRACIN\99DAL^KELLY^^
    'ZOA^36287\VITAMIN\99DAL^^^
            
    Set Summary = Nothing

    If pszAllergyCount Then
      Set Allergy = New Collection
      Dim AllergyObject As Allergies
      
      For iTempA = 1 To pszAllergyCount
        Set AllergyObject = New Allergies

        szTempa = szGetSegSeq(gszHL7RetMsg, iTempA, "ZOA", hlZOASeg.AllergyIen)
        AllergyObject.Name = szGetColumn(szTempa, 2, "\")
        AllergyObject.IEN = szGetColumn(szTempa, 1, "\")
        AllergyObject.Comment = szGetSegSeq(gszHL7RetMsg, iTempA, "ZOA", hlZOASeg.Comment)
        
        Allergy.Add AllergyObject
        
      Next
    End If
      
End Function
Public Function PerformAllergyAction(ActionCode As hlALLERGYACTION, _
                                     AllergyIen As String, _
                                     Optional AllergyName As String, _
                                     Optional Comment As String) As hlERRORS
                                     
'Actions are: ADD new, DELETE and EDIT

    If gbHL7LoginStatus = False Then
      PerformAllergyAction = hlERRORS.hlERROR_NOT_LOGGED_IN
      Exit Function
    End If
    
    ' MSH
    ReDim seginfo(1 To hlMSHSeg.NumberOfSegs) As String
      seginfo(hlMSHSeg.MsgType) = "ZPW\Z20"
    gszHL7SndMsg = szAddSeg("MSH", seginfo())
    
    ' PID
    ReDim seginfo(1 To 27) As String
      seginfo(hlPIDSeg.SetID) = "1"
      seginfo(hlPIDSeg.IEN) = szFormatPIDSegIEN()
      seginfo(hlPIDSeg.Name) = Patient.HL7Name
    gszHL7SndMsg = gszHL7SndMsg & szAddSeg("PID", seginfo())
    
    ' ZOA
    ReDim seginfo(1 To hlZOASeg.NumberOfSegs) As String
      seginfo(hlZOASeg.AllergyIen) = AllergyIen & "\" & AllergyName
      seginfo(hlZOASeg.Comment) = Comment
      seginfo(hlZOASeg.UpdateAllergy) = GetTextfromAllergyAction(ActionCode)    '"A"dd-new D-elete E-dit
    gszHL7SndMsg = gszHL7SndMsg & szAddSeg("ZOA", seginfo())

    '=============================================================

     'SEND HL7 command package >>>>>>>>>>>>>>>
         glngTimeoutStatus = HL7MsgHandler()  'changed 6-30-98;AAB new design
         
         If glngTimeoutStatus = hlCOMMSTATUS.hlCOMM_TIMEOUT Then
           PerformAllergyAction = hlERRORS.hlERROR_TIMED_OUT
         Else
           If gszHL7RetCode = "AA" Then
             PerformAllergyAction = hlERRORS.hlERROR_NONE
           Else
             PerformAllergyAction = hlERRORS.hlERROR_PERFORM_ALLERGY_ACTION_FAILED
           End If
         End If

End Function

Public Function SendPatientReleaseMessage() As hlERRORS

    SendPatientReleaseMessage = hlERRORS.hlERROR_NONE
    If gbHL7LoginStatus = False Then
      SendPatientReleaseMessage = hlERRORS.hlERROR_NOT_LOGGED_IN
      Exit Function
    End If
    
    'If we locked a page for this patient with a ZPW/Z12 then send an unlock ZPW/Z14
    '   This may not be necessary, as we send a ZPW/Z17 'Unlock All Pages' message,
    '   but just in case. AMW 8/31/99
          If LenB(mszLockedPageID) > 0 Then UnlockPage mszLockedPageID
          mszLockedPageID = ""
    
    'This sends a ZPW/Z17
          Build_PatientRelease_Message
               ''Set goPatient = New PatientType
               
     'SEND HL7 command package >>>>>>>>>>>>>>>
          glngTimeoutStatus = HL7MsgHandler()  'changed 6-30-98;AAB new design
          
          If glngTimeoutStatus = hlCOMMSTATUS.hlCOMM_TIMEOUT Then
               SendPatientReleaseMessage = hlERRORS.hlERROR_TIMED_OUT
               Exit Function
          Else

               'added 4-7-98; AAB
               If gszHL7RetCode = "AA" Then
                 SendPatientReleaseMessage = hlERRORS.hlERROR_NONE
               Else
                 SendPatientReleaseMessage = hlERRORS.hlERROR_PATIENT_RELEASE_FAILED
                 Exit Function
               End If
          End If
    
         Set goPatient = New PatientType 'moved 6-30-98;AAB
     
         gszDefOrderPage = ""
         gszBatchId = ""
         gszBatchTime = ""
         Patient.HL7Name = ""
         Patient.IEN = ""

End Function


Public Function Logout() As hlERRORS

    Logout = hlERRORS.hlERROR_NONE
    If gbHL7LoginStatus = False Then
      Logout = hlERRORS.hlERROR_NOT_LOGGED_IN
      Exit Function
    End If
    
    Call Build_Logout_Message
    
     'SEND HL7 command package >>>>>>>>>>>>>>>
    glngTimeoutStatus = HL7MsgHandler()  'changed 6-30-98;AAB new design
    
    If glngTimeoutStatus = hlCOMMSTATUS.hlCOMM_TIMEOUT Then
      Logout = hlERRORS.hlERROR_TIMED_OUT
    Else
        'added 4-7-98; AAB
      If gszHL7RetCode = "AA" Then
        Logout = hlERRORS.hlERROR_NONE
      Else
        Logout = hlERRORS.hlERROR_LOGOUT_FAILED
        ''Exit Function
      End If
    End If
    
    'Reset even on failure
    gszUserIEN = ""
    gszProviderIEN = ""
    gszOrderingProviderIEN = ""
    gbHL7LoginStatus = False
    
End Function

Public Function GetErrorText(lngErrorCode As hlERRORS) As String

    GetErrorText = goErrorTable.Item(lngErrorCode).Text

End Function

Public Function CreateBatch() As hlERRORS

  If gbHL7LoginStatus = False Then
    CreateBatch = hlERRORS.hlERROR_NOT_LOGGED_IN
    Exit Function
  End If
  
  CreateBatch = hlERRORS.hlERROR_NONE
      
    ' MSH
    ReDim seginfo(1 To hlMSHSeg.NumberOfSegs) As String
      seginfo(hlMSHSeg.SendingApp) = "PWS\ORE"   'who is sending
      seginfo(hlMSHSeg.RecvApp) = "CHCS\ORE"     'sent to
      seginfo(hlMSHSeg.MsgType) = "ZPW\Z13"      'message/event
    gszHL7SndMsg = szAddSeg("MSH", seginfo())

     'SEND HL7 command package >>>>>>>>>>>>>>>
    glngTimeoutStatus = HL7MsgHandler()  'changed 6-30-98;AAB new design
    If glngTimeoutStatus = hlCOMMSTATUS.hlCOMM_TIMEOUT Then
      CreateBatch = hlERRORS.hlERROR_TIMED_OUT
      gszBatchId = ""
      gszBatchTime = ""
  
    Else
  
      '  If there are no errors then we have an Order Batch
      If gszHL7RetCode = "AA" Then
        gszBatchId = szGetSegSeq(gszHL7RetMsg, 1, "ZOB", hlZOBSeg.BatchID)
        gszBatchTime = szGetSegSeq(gszHL7RetMsg, 1, "ZOB", hlZOBSeg.BatchTime)
      
      Else
        CreateBatch = hlERRORS.hlERROR_CREATE_BATCH_FAILED
        gszBatchId = ""
        gszBatchTime = ""
      End If
    End If

End Function

Private Function FetchErrors(TabletObject As BaseTablet, _
                            ErrorCollection As Collection) As hlERRMSGFLAG

Dim piTemp As Long
Dim piErrCount As Integer
Dim pszErrorString As String
Dim pszErrorCode As String
Dim pszErrorText As String
Dim piNumErrs As Integer
Dim ErrorObject As ErrorType
'KELLYMCDim iNumMSH As Integer
'KELLYMCDim iNumMSA As Integer
'KELLYMCDim iNumSegs As Integer

     'Clear out ErrorCollection
     Set ErrorCollection = New Collection

    'set defaults
    FetchErrors = hlERROR_IS_MESSAGE
    TabletObject.ErrorIsMessage = hlERROR_IS_ERROR
    'TabletObject.ErrorIsMessage = hlERROR_IS_MESSAGE ### want to do this ... why are the two lines above 'opposites' ?

    piNumErrs = iGetSegCount(gszHL7RetMsg, "ERR")
    'KELLYMCpszNumErrors = Str$(piNumErrs)

    '==============================================================
    'get ERR segment messages
        
 If piNumErrs > 0 Then      'fix 4-2-98;AAB
 
    For piErrCount = 1 To piNumErrs
      pszErrorString = szGetSegSeq(gszHL7RetMsg, piErrCount, "ERR", hlERRSeg.ErrorCodeAndLocation)    'whole \\\n&text&trailer
        piTemp = InStr(pszErrorString, "\") + 3             'strip off leading \\\
      pszErrorString = Mid$(pszErrorString, piTemp)         'without header junk
      pszErrorCode = szGetColumn(pszErrorString, 1, "&")    'CODE
      pszErrorText = szGetColumn(pszErrorString, 2, "&")    'TEXT
      Set ErrorObject = New ErrorType
      ErrorObject.TypeCode = pszErrorCode
      ErrorObject.Text = pszErrorText
      ErrorCollection.Add ErrorObject
    Next piErrCount

    If gszHL7RetCode = "AA" Then
          TabletObject.ErrorIsMessage = hlERRMSGFLAG.hlERROR_IS_MESSAGE
          FetchErrors = hlERRMSGFLAG.hlERROR_IS_MESSAGE      'not passed back to caller
    Else
          TabletObject.ErrorIsMessage = hlERRMSGFLAG.hlERROR_IS_ERROR
          FetchErrors = hlERRMSGFLAG.hlERROR_IS_ERROR
    End If
     
  ElseIf szGetSegSeq(gszHL7RetMsg, 1, "MSA", hlMSASeg.AckCode) = "AR" Then
    'If no ERR segments, but AckCode is AR, treat MSA6 as Error   AMW 10/19/99
    'This error received if order a lab with processpriority=STAT with STAT allowed (ZOL(1)=S from ZP)
    '10/19/1999 11:56:57  Inbound msg: MSH^\|~&^CHCS/ORE^^PWS/ORE^^19991019085400^^ORR\ACK^F1756-3801355^P^2.2^^^^AL^
    'MSA^AR^PWS0821204593^^^^STATPRI+9,ORPWSL2:1, %DSM-E-LINER, Undefined line label, -DSM-I-ECODE, MUMPS error code: M13^

    Set ErrorObject = New ErrorType
    ErrorObject.Text = szGetSegSeq(gszHL7RetMsg, 1, "MSA", hlMSASeg.ErrorCondition)
 '   LabError.Add ErrorObject
    ErrorCollection.Add ErrorObject
 
    FetchErrors = hlERRMSGFLAG.hlERROR_IS_ERROR

  End If
      

    '3-15-98; MSA3 and MSA6 have more error condition info,
    'especially if code = "AR"
    'Typically MSA6 will be first of the two to have text
'    order.ErrorCondition3 = szGetSegSeq(gszHL7RetMsg, 1, "MSA", hlMSASeg.ErrorText)
'    order.ErrorCondition6 = szGetSegSeq(gszHL7RetMsg, 1, "MSA", hlMSASeg.ErrorCondition)
  
End Function
Private Function FetchSignErrors(ErrorCollection As Collection) As hlERRMSGFLAG

Dim piTemp As Long
Dim piErrCount As Integer
Dim pszErrorString As String
Dim pszErrorCode As String
Dim pszErrorText As String
Dim piNumErrs As Integer
Dim piNumZERs As Integer

Dim ErrorObject As ErrorType
'KELLYMCDim iNumMSH As Integer
'KELLYMCDim iNumMSA As Integer
'KELLYMCDim iNumSegs As Integer

     'Clear out ErrorCollection
     Set ErrorCollection = New Collection

    'set defaults
    FetchSignErrors = hlERROR_IS_MESSAGE
    'TabletObject.ErrorIsMessage = hlERROR_IS_ERROR
    'TabletObject.ErrorIsMessage = hlERROR_IS_MESSAGE ### want to do this ... why are the two lines above 'opposites' ?

    piNumErrs = iGetSegCount(gszHL7RetMsg, "ERR")
    piNumZERs = iGetSegCount(gszHL7RetMsg, "ZER")
    'KELLYMCpszNumErrors = Str$(piNumErrs)

    '==============================================================
    'get ERR segment messages
        
 If piNumErrs > 0 Then      'fix 4-2-98;AAB
 
    For piErrCount = 1 To piNumErrs
      pszErrorString = szGetSegSeq(gszHL7RetMsg, piErrCount, "ERR", hlERRSeg.ErrorCodeAndLocation)    'whole \\\n&text&trailer
        piTemp = InStr(pszErrorString, "\") + 3             'strip off leading \\\
      pszErrorString = Mid$(pszErrorString, piTemp)         'without header junk
      pszErrorCode = szGetColumn(pszErrorString, 1, "&")    'CODE
      pszErrorText = szGetColumn(pszErrorString, 2, "&")    'TEXT
      Set ErrorObject = New ErrorType
      ErrorObject.TypeCode = pszErrorCode
      ErrorObject.Text = pszErrorText
      ErrorCollection.Add ErrorObject
    Next piErrCount

    'If gszHL7RetCode = "AA" Then
    '      TabletObject.ErrorIsMessage = hlERRMSGFLAG.hlERROR_IS_MESSAGE
    '      FetchSignErrors = hlERRMSGFLAG.hlERROR_IS_MESSAGE      'not passed back to caller
    'Else
    '      TabletObject.ErrorIsMessage = hlERRMSGFLAG.hlERROR_IS_ERROR
    '      FetchSignErrors = hlERRMSGFLAG.hlERROR_IS_ERROR
    'End If
     
  ElseIf szGetSegSeq(gszHL7RetMsg, 1, "MSA", hlMSASeg.AckCode) = "AR" Then
    'If no ERR segments, but AckCode is AR, treat MSA6 as Error   AMW 10/19/99
    'This error received if order a lab with processpriority=STAT with STAT allowed (ZOL(1)=S from ZP)
    '10/19/1999 11:56:57  Inbound msg: MSH^\|~&^CHCS/ORE^^PWS/ORE^^19991019085400^^ORR\ACK^F1756-3801355^P^2.2^^^^AL^
    'MSA^AR^PWS0821204593^^^^STATPRI+9,ORPWSL2:1, %DSM-E-LINER, Undefined line label, -DSM-I-ECODE, MUMPS error code: M13^

    Set ErrorObject = New ErrorType
    ErrorObject.Text = szGetSegSeq(gszHL7RetMsg, 1, "MSA", hlMSASeg.ErrorCondition)
 '   LabError.Add ErrorObject
    ErrorCollection.Add ErrorObject
 
    FetchSignErrors = hlERRMSGFLAG.hlERROR_IS_ERROR
  ElseIf piNumZERs > 0 Then
    For piErrCount = 1 To piNumZERs
        Set ErrorObject = New ErrorType
        ErrorObject.Text = "Error On Order ID: " & szGetSegSeq(gszHL7RetMsg, piErrCount, "ZOD", hlZODSeg.OrderId) & " " & szGetSegSeq(gszHL7RetMsg, piErrCount, "ZOD", hlZODSeg.OrderDisplayText) & vbCrLf & _
            "Message: " & szGetSegSeq(gszHL7RetMsg, piErrCount, "ZER", 1) & vbCrLf
        ErrorCollection.Add ErrorObject
    Next piErrCount

  End If
      

    '3-15-98; MSA3 and MSA6 have more error condition info,
    'especially if code = "AR"
    'Typically MSA6 will be first of the two to have text
'    order.ErrorCondition3 = szGetSegSeq(gszHL7RetMsg, 1, "MSA", hlMSASeg.ErrorText)
'    order.ErrorCondition6 = szGetSegSeq(gszHL7RetMsg, 1, "MSA", hlMSASeg.ErrorCondition)
  
End Function
Public Property Let szMTFCode(ByVal szData As String)
    mvarszMTFCode = szData
    clinic.MTF_CODE = mvarszMTFCode      'added 11-13-97
End Property


Public Property Get szMTFCode() As String
    szMTFCode = mvarszMTFCode
End Property

Public Property Get szMTFIEN() As String
  szMTFIEN = mvarszMTFIEN
End Property

Public Property Let szMTFIEN(ByVal szData As String)
  mvarszMTFIEN = szData
  clinic.MTF_IEN = szData
End Property
Public Property Get szApptIEN() As String 'SCR 31315  Order Entry - Appt IEN needs to be sent to CHCS

  szApptIEN = mvarszApptIEN
End Property

Public Property Let szApptIEN(ByVal szData As String) 'SCR 31315  Order Entry - Appt IEN needs to be sent to CHCS
  mvarszApptIEN = szData
End Property

Public Property Let szMeprsCode(ByVal vData As String)
    mvarszMeprsCode = vData
    clinic.MEPRS_STRING = "\" & mvarszMeprsIen & "\" & mvarszMeprsCode & "\99MEP"

End Property


Public Property Get szMeprsCode() As String
    szMeprsCode = mvarszMeprsCode
End Property



Public Property Let szMeprsIen(ByVal vData As String)
    mvarszMeprsIen = vData
    clinic.MEPRS_STRING = "\" & mvarszMeprsIen & "\" & mvarszMeprsCode & "\99MEP"

End Property


Public Property Get szMeprsIen() As String
    szMeprsIen = mvarszMeprsIen
End Property


Public Property Get szClinicIEN() As String
    szClinicIEN = mvarszClinicIen
End Property


Public Property Let szClinicIEN(ByVal vData As String)
    mvarszClinicIen = vData
    clinic.IEN = vData
End Property

Public Function GetOrderDetail(OrderId As String, _
                               OrderTypeCode As hlORDERTYPECODE) As hlERRORS
                               
Dim DetailObject As BaseTablet
    
        If gbHL7LoginStatus = False Then
                GetOrderDetail = hlERRORS.hlERROR_NOT_LOGGED_IN
                Exit Function
        End If
        
        GetOrderDetail = hlERRORS.hlERROR_NONE
    
        Select Case OrderTypeCode
        Case hlORDERTYPECODE.hlRX_ORDERS_CODE:
               Set RXTablet = New RxType
               Set DetailObject = RXTablet
                 
        Case hlORDERTYPECODE.hlRAD_ORDERS_CODE:
               Set RadTablet = New RadType
               Set DetailObject = RadTablet
         
        Case hlORDERTYPECODE.hlLAB_ORDERS_CODE:
               Set LabTablet = New LabType
               Set DetailObject = LabTablet
               
        End Select
    
        Call Build_Order_Detail_Message(OrderId, OrderTypeCode)
    
       'SEND HL7 command package >>>>>>>>>>>>>>>
        glngTimeoutStatus = HL7MsgHandler()  'changed 6-30-98;AAB new design
        
        If glngTimeoutStatus = hlCOMMSTATUS.hlCOMM_TIMEOUT Then
                GetOrderDetail = hlERRORS.hlERROR_TIMED_OUT
        Else
                If gszHL7RetCode = "AA" Then
                        GetRequestedInformation DetailObject
                    'defaulted above; GetOrderDetail = ERROR_NONE 'comment added 6-30-98;AAB
                Else
                        GetOrderDetail = hlERRORS.hlERROR_ORDER_DETAIL_FAILED
                End If
        End If

End Function

Public Function PerformAction(OrderId As String, _
                              ActionCode As hlACTION, _
                              OrderTypeCode As hlORDERTYPECODE) As hlERRORS
                              
Dim TabletObject As BaseTablet
Dim ErrorCollection As Collection
Dim WarningCollection As Collection
Dim ReactivationDate As String
    'KELLYMCDim FetchErrorsReturn As hlERRMSGFLAG
Dim bClearTablet As Boolean
     
     
        If gbHL7LoginStatus = False Then
            PerformAction = hlERRORS.hlERROR_NOT_LOGGED_IN
            Exit Function
        End If
    
        PerformAction = hlERRORS.hlERROR_PERFORM_ACTION_FAILED  'default

    Set ErrorCollection = Nothing
    Set ErrorCollection = New Collection
  '  bClearTablet = (ActionCode = hlACTION_DELETE Or ActionCode = hlACTION_HOLD Or ActionCode = hlACTION_DETAIL)
    'SCR 21424
    bClearTablet = (ActionCode = hlACTION_HOLD Or ActionCode = hlACTION_DETAIL)
    
    Select Case OrderTypeCode
    'AMW 2/15/00 Clear out tablet on delete so don't use bogus info from last transaction
    'AMW 2/29/00 Put orderid in cleared out tablet to have that as reference
         Case hlORDERTYPECODE.hlRX_ORDERS_CODE:
            ReactivationDate = RXTablet.ReactivationDate
            If bClearTablet Then
                Set RXTablet = New RxType
                Set WarningCollection = New Collection
                RXTablet.OrderId = OrderId      'AMW 2/29/00 Put orderid in tablet
                RXTablet.ReactivationDate = ReactivationDate
            End If
            Set TabletObject = RXTablet
            Set WarningCollection = RxWarning
         Case hlORDERTYPECODE.hlRAD_ORDERS_CODE:
            ReactivationDate = RadTablet.ReactivationDate
            If bClearTablet Then
                Set RadTablet = New RadType
                Set WarningCollection = New Collection
                RadTablet.OrderId = OrderId
                RadTablet.ReactivationDate = ReactivationDate
            End If
            Set TabletObject = RadTablet
            Set WarningCollection = RadWarning
         Case hlORDERTYPECODE.hlLAB_ORDERS_CODE:
            ReactivationDate = LabTablet.ReactivationDate
            If bClearTablet Then
                Set LabTablet = New LabType
                Set WarningCollection = New Collection
                LabTablet.OrderId = OrderId
                LabTablet.ReactivationDate = ReactivationDate
            End If
            Set TabletObject = LabTablet
            Set WarningCollection = LabWarning
         Case Else 'if cancelling another order type
            Set OtherTablet = New BaseTablet
            Set WarningCollection = LabWarning
            OtherTablet.OrderTypeCode = OrderTypeCode
            Set TabletObject = LabTablet
         
    End Select
    
    If LenB(gszProviderIEN) = 0 Then
        PerformAction = hlERRORS.hlERROR_NO_PROVIDER_IEN
        Exit Function
    End If
    
    Select Case ActionCode
    Case hlACTION_DELETE, hlACTION_RENEW, hlACTION_RELEASE, hlACTION_HOLD, hlACTION_MODIFY
   
      lngTempA = Build_Order_Management_Message(ActionCode, OrderId, TabletObject, WarningCollection)
      If lngTempA <> hlERRORS.hlERROR_NONE Then
        PerformAction = lngTempA
        Exit Function
      End If
      
    Case hlACTION_DETAIL
      'NOTE: Detail calls populate the nnDETAIL object, NOT the nnTABLET object
      Call Build_Order_Detail_Message(OrderId, OrderTypeCode)
    
    Case Else
      PerformAction = hlERRORS.hlERROR_BAD_ACTION_CODE
      Exit Function
    End Select
        
     'SEND HL7 command package >>>>>>>>>>>>>>>
    glngTimeoutStatus = HL7MsgHandler()  'changed 6-30-98;AAB new design
    
    If glngTimeoutStatus = hlCOMMSTATUS.hlCOMM_TIMEOUT Then
      PerformAction = hlERRORS.hlERROR_TIMED_OUT
      Exit Function
    ElseIf gszHL7RetCode = "AA" Then    'If not AA, then performaction =hlERROR_PERFORM_ACTION_FAILED (default)
        GetRequestedInformation TabletObject
        PerformAction = hlERRORS.hlERROR_NONE
    End If
    
     'KELLYMCFetchErrorsReturn = FetchErrors(TabletObject, ErrorCollection)
     FetchErrors TabletObject, ErrorCollection
     HL7GetWarnings WarningCollection, TabletObject
    
     Select Case OrderTypeCode
         Case hlORDERTYPECODE.hlRX_ORDERS_CODE:
             Set RxError = ErrorCollection
             Set RxWarning = WarningCollection
         Case hlORDERTYPECODE.hlRAD_ORDERS_CODE:
             Set RadError = ErrorCollection
             Set RadWarning = WarningCollection
         Case hlORDERTYPECODE.hlLAB_ORDERS_CODE:
             Set LabError = ErrorCollection
             Set LabWarning = WarningCollection
         Case Else
             Set OtherError = ErrorCollection
             'Set LabWarning = WarningCollection
         
      End Select
      
End Function
Public Function SendPDTSCancel(hlOutcome As hlPDTSOUTCOMES, szReason As String, szHeader As String, szOrderID As String) As hlERRORS
                              
    Dim TabletObject As BaseTablet
    Dim ErrorCollection As Collection
    Dim WarningCollection As Collection
   
    If gbHL7LoginStatus = False Then
        SendPDTSCancel = hlERRORS.hlERROR_NOT_LOGGED_IN
        Exit Function
    End If

    SendPDTSCancel = hlERRORS.hlERROR_PERFORM_ACTION_FAILED  'default

    Set ErrorCollection = Nothing
    Set ErrorCollection = New Collection
    
   
    Set WarningCollection = RxWarning
    
    
    If LenB(gszProviderIEN) = 0 Then
        SendPDTSCancel = hlERRORS.hlERROR_NO_PROVIDER_IEN
        Exit Function
    End If
    
   
      lngTempA = Build_PDTS_Cancel_Message(hlOutcome, szReason, szHeader, szOrderID)
      If lngTempA <> hlERRORS.hlERROR_NONE Then
        SendPDTSCancel = lngTempA
        Exit Function
      End If
      
   
        
     'SEND HL7 command package >>>>>>>>>>>>>>>
    glngTimeoutStatus = HL7MsgHandler()  'changed 6-30-98;AAB new design
    
    If glngTimeoutStatus = hlCOMMSTATUS.hlCOMM_TIMEOUT Then
      SendPDTSCancel = hlERRORS.hlERROR_TIMED_OUT
      Exit Function
    ElseIf gszHL7RetCode = "AA" Then    'If not AA, then SendPDTSCancel =hlERROR_PERFORM_ACTION_FAILED (default)
        GetRequestedInformation RXTablet
        SendPDTSCancel = hlERRORS.hlERROR_NONE
    End If
    
      FetchErrors RXTablet, ErrorCollection
     HL7GetWarnings WarningCollection, RXTablet
    

    Set RxError = ErrorCollection
    Set RxWarning = WarningCollection
  
      
End Function

Public Function SubmitBatch() As hlERRORS

    If gbHL7LoginStatus = False Then
      SubmitBatch = hlERRORS.hlERROR_NOT_LOGGED_IN
      Exit Function
    End If
    
'    If gbHL7MessageInMotion = YES Then
'        SubmitBatch = ERROR_HL7_MESSAGE_IN_MOTION
'        Exit Function
'    End If
'
    SubmitBatch = hlERRORS.hlERROR_NONE

'===================================================================
'===================================================================
    
''Sub Btn_Submit_Click()    'taken out of control array 11-13-95-AAB

  
    '===============================================================
    'If a batch does not exist (CreateBatch) then error and abort
      If (LenB(gszBatchId) = 0) Or (LenB(gszBatchTime) = 0) Then
        SubmitBatch = hlERRORS.hlERROR_NO_BATCH
        Exit Function
      End If

    '=============================================================================
    ' MSH
      ReDim seginfo(1 To hlMSHSeg.NumberOfSegs) As String
        seginfo(hlMSHSeg.SendingApp) = "PWS\ORE"
        seginfo(hlMSHSeg.RecvApp) = "CHCS\ORE"
        seginfo(hlMSHSeg.MsgType) = "ZPW\Z16"
      gszHL7SndMsg = szAddSeg("MSH", seginfo())
      
    ' PID
      ''gszHL7SndMsg = gszHL7SndMsg & szAddSeg("PID", gszPatSelArray())   'see PatSelect
      
      ReDim seginfo(1 To hlPIDSeg.NumberOfSegs) As String
        seginfo(hlPIDSeg.SetID) = "1"
        seginfo(hlPIDSeg.IEN) = Patient.IEN  '& "\\\"      '& 'gszMTFIEN
        seginfo(hlPIDSeg.Name) = Patient.HL7Name
      gszHL7SndMsg = gszHL7SndMsg & szAddSeg("PID", seginfo())
        
    ' ZOB
      ReDim seginfo(1 To hlZOBSeg.NumberOfSegs) As String
        seginfo(hlZOBSeg.BatchID) = gszBatchId
        seginfo(hlZOBSeg.BatchTime) = gszBatchTime
      gszHL7SndMsg = gszHL7SndMsg & szAddSeg("ZOB", seginfo())
    
    '=============================================================================

     'SEND HL7 command package >>>>>>>>>>>>>>>
      glngTimeoutStatus = HL7MsgHandler()  'changed 6-30-98;AAB new design
      
      If glngTimeoutStatus = hlCOMMSTATUS.hlCOMM_TIMEOUT Then
        SubmitBatch = hlERRORS.hlERROR_TIMED_OUT
      Else
        If gszHL7RetCode = "AA" Then
          gszBatchId = ""
          gszBatchTime = ""
          SubmitBatch = hlERRORS.hlERROR_NONE
        Else
          SubmitBatch = hlERRORS.hlERROR_SUBMIT_BATCH_FAILED
        End If
      End If

End Function
Public Function PreVerify(OrderTypeCode As hlORDERTYPECODE) As hlERRORS

Dim pszTemp As String
Dim pszSegment As String
Dim FetchErrorsReturn As hlERRMSGFLAG
Dim szOrderName As String
Dim szOrderIen As String
Dim szPharmacyIEN As String
Dim szPharmacyName As String
Dim szComma As String
Dim iLabZocCount As Integer
Dim iLabZRQCount As Integer
Dim szZoc1 As String
Dim szZoc2 As String
Dim szZoc5 As String
Dim pszSegmentData As String    'so we can match HL7 segment-seq-numbers; trim its head off
Dim ZPDsegment As String
Dim pszPharmacyHL7rdy As String
Dim pszAnswers As String
Dim piPosition As Long
Dim piStart As Integer
Dim szOBR4 As String
Dim szOrc7 As String
Dim iTempA As Long
Dim szRadDeftDuratn As String
Dim szZom4 As String
Dim piTempA As Integer
Dim pszTempA As String

Dim SpecimenObject As TypeSpecimen
Dim AnswerObject As Answers

    '====================================================================================
    'NOTE: 7-1-98;AAB: There is a design error in HL7; error-7 should return code "AA"
    '   as syntax is OK, but the HL7 code returns "AR"
    'ALSO, if we send a blank IEN (bad syntax), we get status "AA" with no defaults !!!
    'ALSO, if we send NO patient info, we get "AA" with error segment!!
    
    'MSA^AR^PWS08207024057^^^^Test/Drug/Procedure not currently available.^
    'ERR^\\\7&Test/Drug/Procedure not currently available.&99ERR^
    
    'sent without selecting a patient:
    'MSA^AR^PWS08207024062^^^^Missing required parameter.^
    'ERR^\\\1&Missing required parameter.&99ERR^
    
    '7-1-98; HL7 error-7 drug/test not available returns status "AR" even with good syntax
    'BUT!! Send Preverify with missing IEN returns "AA" with blank info!!
    '====================================================================================
                                                                     
            If gbHL7LoginStatus = False Then
                 PreVerify = hlERRORS.hlERROR_NOT_LOGGED_IN
                 Exit Function
            End If
            
            If LenB(gszProviderIEN) = 0 Then
                PreVerify = hlERRORS.hlERROR_NO_PROVIDER_IEN
                Exit Function
            End If
         
            PreVerify = hlERRORS.hlERROR_NONE  'default to start
        
        'Keep the required order info before ditching the tablet
            Call GrabOrderInfo(szOrderName, szOrderIen, OrderTypeCode)
                       
        '===================================================================================================================
            Select Case OrderTypeCode
            Case hlORDERTYPECODE.hlRX_ORDERS_CODE
                 
                'Preverify means "start over"; get fresh objects
                'If there are real ERRORS, then quick exit leaves all other properties BLANK
                    
                'Need pharmacy info for preverify
                    szPharmacyIEN = RXTablet.PharmacyIenSent
                    szPharmacyName = RXTablet.PharmacyNameSent
                        
                'If this is part of a premodify, we don't want to clear the whole tablet,
                'just the few scragglers
                    If Not mbPreModifying Then
                         Set RXTablet = New RxType
                    End If
                    
                    Set RxError = New Collection
                    Set RxWarning = New Collection
                     
                'Set the order info for new tablet
                    RXTablet.OrderName = szOrderName
                    RXTablet.OrderIen = szOrderIen
                     
                'Set the pharmacy stuff back
                    RXTablet.PharmacyIenSent = szPharmacyIEN
                    RXTablet.PharmacyNameSent = szPharmacyName
                    
                    If LenB(RXTablet.OrderIen) = 0 Then
                         PreVerify = hlERRORS.hlERROR_PREVERIFY_RX_FAILED
                         Exit Function
                    End If
                    
                    Call Build_PreVerifyRx_Message(RXTablet)
                    
                    RXTablet.LastPreverifyNameWas = szOrderName
                    RXTablet.LastPreverifyIenWas = szOrderIen
                          
                'SEND HL7 command package >>>>>>>>>>>>>>>
                    glngTimeoutStatus = HL7MsgHandler()
                    If glngTimeoutStatus = hlCOMMSTATUS.hlCOMM_TIMEOUT Then
                         PreVerify = hlERRORS.hlERROR_TIMED_OUT
                         Exit Function
                    End If
                  
                'FIRST look at any errors.
                'If any, and they are not 'soft' then exit
                    FetchErrorsReturn = FetchErrors(RXTablet, RxError)
                    
                If FetchErrorsReturn = hlERROR_IS_ERROR Then Exit Function
                     
                'get RX defaults
                    
                'note 8-17-95; for preverify RX only,
                '   RXO(16) is "1" if Medication requested is valid
                '  +++(ien not rqd for pre-verify rx)
                         
                    pszTemp = szGetSegSeq(gszHL7RetMsg, 1, "RXO", hlRXOSeg.DrugIsValid)
                    If pszTemp = "1" Then
                         RXTablet.DrugIsValid = True
                    Else
                         RXTablet.DrugIsValid = False
                    End If
                       
                'put return text into status area if AA
                    If (gszHL7RetCode = "AA") Then
                            szTempa = szGetSegSeq(gszHL7RetMsg, 1, "ZOS", 2)
                            szTempb = szGetSegSeq(gszHL7RetMsg, 1, "ZOS", 3)
                        If LenB(szTempa) > 0 Then szComma = ","
                            RXTablet.Status = szTempa & szComma & szTempb
                    End If
                    
                    pszSegment = szGetSegSeq(gszHL7RetMsg, 1, "RXO", hlRXOSeg.DrugName)
                    RXTablet.OrderIen = szGetColumn(pszSegment, 1, "\")  'IEN
                    RXTablet.OrderName = szGetColumn(pszSegment, 2, "\")   'TEXT
                    
                    If RXTablet.LastPreverifyNameWas = RXTablet.OrderName Then
                        RXTablet.LastPreverifyNamesMatch = True
                    Else
                        RXTablet.LastPreverifyNamesMatch = False
                    End If
                        
                    ZPDsegment = GetFullSeg(gszHL7RetMsg, "ZPD")
                     
                ''Offset all by one, as (ZPDsegment, 1, "^") = "ZPD^" 2-5-96
                    pszSegmentData = Mid$(ZPDsegment, 5)    'count past the "ZPD^" == 4 chars
                    pszPharmacyHL7rdy = szGetColumn(pszSegmentData, 1, "^")      'in HL7 format (2\MAIN OUTPATIENT PHARMACY\99PSI)
                    
                    RXTablet.DefaultPharmacyIen = szGetColumn(pszPharmacyHL7rdy, 1, "\")      '2-22-96
                    RXTablet.DefaultPharmacyName = szGetColumn(pszPharmacyHL7rdy, 2, "\")
                    RXTablet.MaxRefills = szGetColumn(pszSegmentData, hlZPDSeg.MaxRefills, "^")
                    RXTablet.DefaultSig = szGetColumn(pszSegmentData, hlZPDSeg.DefaultSig, "^")
                    RXTablet.DefaultUnit = szGetColumn(pszSegmentData, hlZPDSeg.DefaultUnit, "^")
                    RXTablet.MaxQuantity = szGetColumn(pszSegmentData, hlZPDSeg.MaxQuantity, "^")
                    RXTablet.MaxDaysSupply = szGetColumn(pszSegmentData, hlZPDSeg.MaxDaysSupply, "^")
                    RXTablet.NoteToProvider = szGetColumn(pszSegmentData, hlZPDSeg.DispensingComment, "^")
                    RXTablet.Cost = szGetColumn(pszSegmentData, hlZPDSeg.LocalCost, "^")
                    pszTemp = szGetColumn(pszSegmentData, hlZPDSeg.DispenseCompleteContainer, "^")
                     
                    If pszTemp <> "1" And pszTemp <> "2" Then pszTemp = "1"         'Default to Yes if unidentified
                        
                    RXTablet.DispenseFull = Val(pszTemp)
                    RXTablet.DefaultQuantity = szGetColumn(pszSegmentData, hlZPDSeg.DefaultQuantity, "^")    'DEFAULT #
                    RXTablet.DefaultExpire = szGetColumn(pszSegmentData, hlZPDSeg.DefaultExpiration, "^")      'Default Expiration Date
                    RXTablet.DefaultDaysSupply = szGetColumn(pszSegmentData, hlZPDSeg.DefaultDaysSupply, "^")  'DefaultDaysSply
                     
                '======================================================================================
                'NOTES: 5-9-96:
                'DESIGN consideration 5-9-96; ZPD says: "Dispensing Pharmacy is the default pharmacy based
                'on requesting location OR DISPENSING PHARMACY OBTAINED FROM THE SIG CODE."
                '======================================================================================
                      
                    If Not mbPreModifying Then HL7GetWarnings RxWarning, RXTablet   '.WarningHeader, RxTablet.WarningCount

        '===================================================================================================================
            Case hlORDERTYPECODE.hlLAB_ORDERS_CODE
                         
                'Preverify means "start over"; get fresh objects
                'If there are real ERRORS, then quick exit leaves all other properties BLANK
                    If Not mbPreModifying Then
                        Set LabTablet = New LabType
                    End If
                    
                    Set Specimen = New Collection
                    Set LabError = New Collection
                    Set LabWarning = New Collection
                
                'Set the order info for new tablet
                    LabTablet.OrderName = szOrderName
                    LabTablet.OrderIen = szOrderIen
                    
                    If LenB(LabTablet.OrderIen) = 0 Then
                       PreVerify = hlERRORS.hlERROR_PREVERIFY_LAB_FAILED
                       Exit Function
                    End If
                    
                    Call Build_PreVerifyLab_Message(LabTablet)
                    
                    LabTablet.LastPreverifyNameWas = LabTablet.OrderName
                    LabTablet.LastPreverifyIenWas = LabTablet.OrderIen
                          
                'SEND HL7 command package >>>>>>>>>>>>>>>
                    glngTimeoutStatus = HL7MsgHandler()
                    
                    If glngTimeoutStatus = hlCOMMSTATUS.hlCOMM_TIMEOUT Then
                        PreVerify = hlERRORS.hlERROR_TIMED_OUT
                        Exit Function
                    End If
                         
                'NOTE: ZOC(6) now has cap color <<<<
                
                'FIRST look at any errors
                    iReturn = FetchErrors(LabTablet, LabError)
                    
                If iReturn = False Then Exit Function
                         
                        ''Order.ErrorCondition3 = szGetSegSeq(gszHL7RetMsg, 1, "MSA", hlMSASeg.ErrorText)
                        ''Order.ErrorCondition6 = szGetSegSeq(gszHL7RetMsg, 1, "MSA", hlMSASeg.ErrorCondition)
                    pszSegment = szGetSegSeq(gszHL7RetMsg, 1, "OBR", hlOBRSeg.TestName)
                    LabTablet.OrderIen = szGetColumn(pszSegment, 1, "\") 'IEN
                    LabTablet.OrderName = szGetColumn(pszSegment, 2, "\")  'TEXT
                    
                    If LabTablet.LastPreverifyNameWas = LabTablet.OrderName Then
                        LabTablet.LastPreverifyNamesMatch = True
                    Else
                        LabTablet.LastPreverifyNamesMatch = False
                    End If
                    
                'get defaults
                    pszTemp = szGetSegSeq(gszHL7RetMsg, 1, "ZOL", hlZOLSeg.AncillaryApprovalNeeded)
                    If pszTemp = "1" Then
                      LabTablet.ApprovalRqd = hlAPPROVAL_REQUIRED
                    Else
                      LabTablet.ApprovalRqd = hlAPPROVAL_NOT_REQUIRED
                    End If
                    
                    pszSegment = szGetSegSeq(gszHL7RetMsg, 1, "ZOL", hlZOLSeg.LabCollect)    'ien\specimen\99LRT
                    LabTablet.LabWillCollect_SpecimenIen = szGetColumn(pszSegment, 1, "\")
                    LabTablet.LabWillCollect_Specimen = szGetColumn(pszSegment, 2, "\")
                    LabTablet.LabWillCollect_Times = szGetSegSeq(gszHL7RetMsg, 1, "ZOL", hlZOLSeg.CollectionTimes)
                'times are eg: 0700\0900\1100\1300\1500, and need parsing by user
                    
                'ZOL(1) id code LETTER which indicates HIGHEST priority allowed ...
                    pszSegment = Trim$(szGetSegSeq(gszHL7RetMsg, 1, "ZOL", hlZOLSeg.HighestPriority))
                    pszSegment = UCase$(pszSegment)
                    If LenB(Trim$(pszSegment)) = 0 Then
                        LabTablet.MaxPriorityAllowed = hlLABPRIORITY.hlLAB_PRIORITY_ROUTINE
                    Else
                        LabTablet.MaxPriorityAllowed = goPriorityTable.SearchHL7Code(pszSegment)
                    End If
                    
                'Build Specimen collection
                    iLabZocCount = iGetSegCount(gszHL7RetMsg, "ZOC")
                    iLabZRQCount = iGetSegCount(gszHL7RetMsg, "ZRQ")
                    LabTablet.SingleSpecimenIsComplete = hlSingleSpecimen_NotComplete          'default to start
                    
                    LabTablet.Subscript = GetLabSubscriptEnumFromText(Trim$(szGetSegSeq(gszHL7RetMsg, 1, "ZOL", hlZOLSeg.Subscript)))
                    '"ZOC
                    '^4\BLOOD
                    '^72\SERUM
                    '^ORDER ENTRY REMARK FOR GLUCOSE UP TO 75 CHARACTERS
                    '^DEBUG TITLE FOR ZOC5
                    '^A&ATEXT\B&BTEXT\C&CTEXT\D&DTEXT\E&END TEXT^"
                    '
                    'ZOC^137\BODY FLUID\99LRS^8819\BODY FLUID\99LRP^^^^RED TOP TUBE^
                    'ZOC^250\PERITONEAL DIALYSATE\99LRS^76\PERITONEAL FLUID\99LRP^^^^RED TOP TUBE^
                                                                                                                 
                If iLabZocCount > 0 Then        'ONE-OR-MORE returned ZOC -----------------------
                    #If coiRequiredComments = False Then
                    For piTempA = 1 To iLabZocCount
                        Set SpecimenObject = New TypeSpecimen
                        Set SpecimenObject.Answer = New Collection
                        pszTempA = szGetSegSeq(gszHL7RetMsg, piTempA, "ZOC", hlZOCSeg.CollectionSample)
                        SpecimenObject.IEN = szGetColumn(pszTempA, 1, "\")
                        SpecimenObject.Name = szGetColumn(pszTempA, 2, "\")
                    
                         'NOTE: if ZOC1 is missing, then we have a problem ... used to get table, but not for CBC
                    
                        SpecimenObject.CapColorOrContainer = szGetSegSeq(gszHL7RetMsg, piTempA, "ZOC", hlZOCSeg.Container)  'API bug is fixed now
                        pszTempA = szGetSegSeq(gszHL7RetMsg, piTempA, "ZOC", hlZOCSeg.Specimen)   'ZOC(2) has TYPE info
                        pszTempA = Trim$(pszTempA)
                    
                        If LenB(pszTempA) = 0 Then
                            SpecimenObject.TopographyIen = ""
                            SpecimenObject.Topography = ""
                        Else
                            SpecimenObject.TopographyIen = szGetColumn(pszTempA, 1, "\")
                            SpecimenObject.Topography = szGetColumn(pszTempA, 2, "\")
                        End If
                        
                        SpecimenObject.Comments = Trim$(szGetSegSeq(gszHL7RetMsg, piTempA, "ZOC", hlZOCSeg.Remark))  'ZOC(3) has COMMENT info
                        SpecimenObject.TitleForQuestions = Trim$(szGetSegSeq(gszHL7RetMsg, piTempA, "ZOC", hlZOCSeg.RequiredCommentText))  'ZOC(4) has "Title for Questions" info
    
                        ' New 11-5-97;AAB We must break the answers up into separate pieces
                        pszAnswers = Trim$(szGetSegSeq(gszHL7RetMsg, piTempA, "ZOC", hlZOCSeg.RequiredCommentAnswer)) 'ZOC(5) has CODE&TEXT
                        If LenB(pszAnswers) > 0 Then
                            'we must parse the answers out of the string
                            'example string:  "^A&ATEXT\B&BTEXT\C&CTEXT\D&DTEXT\E&END TEXT^"
                            piPosition = 0   'seed
                            piStart = 1  'init
                            
                            Do
                                Set AnswerObject = New Answers
                                piPosition = InStr(piStart, pszAnswers, "\")
                                
                                ' If return = 0 then it is either the LAST or the ONLY answer
                                If piPosition <> 0 Then
                                   pszTemp = Left$(pszAnswers, piPosition - 1)
                                   pszAnswers = Mid$(pszAnswers, piPosition + 1)
                                Else
                                   pszTemp = pszAnswers
                                End If
                    
         'Note this should take all the chars up to the next &, not just the next one char AMW 12/28/99
         'Can set up new answers in CHCS using ^LTE procedure
         'Max length for answercode is 5 char
         'Max length for answertext is 60 char in CHCS, but only 39 char will get sent via HL7 (don't know if this will cause a problem)
                     
                                 piStart = InStr(1, pszTemp, "&")
                                 szTempa = Left$(pszTemp, piStart - 1)
                                 szTempb = Mid$(pszTemp, piStart + 1)
                                 
                                 AnswerObject.AnswerCode = szTempa
                                 AnswerObject.AnswerText = szTempb
                                 SpecimenObject.Answer.Add AnswerObject
                            Loop Until piPosition = 0
    
                        End If
                        Specimen.Add SpecimenObject
                
                    Next piTempA
                    #Else       '  <<<<<<<<<<<<<<<  BEGIN REQUIRED COMMENTS SECTION  >>>>>>>>>>>>>>>>
                                
                'Redesigned 1-08 to 1-11-2002 to manage multiple ZOCs; AAB
                          
                            ''  For piTempA = 1 To iLabZocCount
                                             
Dim piTempB As Integer
Dim piStartPosition As Integer
Dim piZRQ_Pointer As Integer
Dim piZRRcount As Integer

                        piZRRcount = iGetSegCount(gszHL7RetMsg, "ZRR")
                        
                        ReDim iZOC_Start(iLabZocCount) As Integer
                        ReDim iZRQ_Start(iLabZRQCount) As Integer
                        ReDim iZRR_Start(piZRRcount) As Integer
                                
                        ''If iLabZocCount <> 0 Then
        
                        For piTempB = 1 To iLabZocCount
                            pszSegment = szGetSeg(gszHL7RetMsg, piTempB, "ZOC", piStartPosition)
                            iZOC_Start(piTempB) = piStartPosition
                            Debug.Print "ZOC  (" & Str$(piTempB) & ")  " & Str$(piStartPosition) & " = " & pszSegment
                        Next piTempB
                         
                        If iLabZRQCount <> 0 Then
                            For piTempB = 1 To iLabZRQCount
                                pszSegment = szGetSeg(gszHL7RetMsg, piTempB, "ZRQ", piStartPosition)
                                iZRQ_Start(piTempB) = piStartPosition
                                Debug.Print "ZRQ  (" & Str$(piTempB) & ")  " & Str$(piStartPosition) & " = " & pszSegment
                            Next piTempB
                        End If
                                     
                        If piZRRcount <> 0 Then
                            For piTempB = 1 To piZRRcount
                                pszSegment = szGetSeg(gszHL7RetMsg, piTempB, "ZRR", piStartPosition)
                                iZRR_Start(piTempB) = piStartPosition
                                Debug.Print "ZRR  (" & Str$(piTempB) & ")  " & Str$(piStartPosition) & " = " & pszSegment
                            Next piTempB
                        End If
          
                        '' end If
                           
                           
   '============================================================================
   
Dim piThisZOC As Integer
Dim piStartValue As Integer
Dim piEndValue As Integer
Dim piNextZocLocus As Integer
Dim piPointer As Integer
Dim piThisZRQ As Integer
                           
                        piZRQ_Pointer = 1
            
                        '%% what if NO ZOCs
                        For piThisZOC = 1 To iLabZocCount
                         
                    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                            Set SpecimenObject = New TypeSpecimen
                            Set SpecimenObject.colRequiredComments = New RequiredCommentsCollection
                                ''Dim colAnswers As Collection
                                ''Set colAnswers = New Collection
                                ''Dim itemAnswer As Answers
                            
                            If iLabZRQCount <> 0 Then
                                If piThisZOC = iLabZocCount Then        'special case
                                    If piZRQ_Pointer > iLabZRQCount Then   'no ZRQs left
                                        piStartValue = 0
                                        piEndValue = 0
                                    Else        'take the rest
                                        piStartValue = piZRQ_Pointer
                                        piEndValue = iLabZRQCount
                                    End If
                                Else
                                    'Any ZRQs within range call AddVia Hl7
                                    piNextZocLocus = iZOC_Start(piThisZOC + 1)
                                    piStartValue = 0
                                    piEndValue = 0
                                    For piTempA = piZRQ_Pointer To iLabZRQCount
                                        If iZRQ_Start(piTempA) < piNextZocLocus Then
                                            If piStartValue = 0 Then
                                                piStartValue = piTempA
                                                piEndValue = piTempA
                                            Else
                                                piEndValue = piTempA
                                            End If
                                            piZRQ_Pointer = piZRQ_Pointer + 1
                                        Else
                                            piZRQ_Pointer = piTempA
                                            Exit For
                                        End If
                                    Next piTempA
                                
                                End If
                                
                            End If
                        
                            pszTempA = szGetSegSeq(gszHL7RetMsg, piThisZOC, "ZOC", hlZOCSeg.CollectionSample)
                            SpecimenObject.IEN = szGetColumn(pszTempA, 1, "\")
                            SpecimenObject.Name = szGetColumn(pszTempA, 2, "\")
                            
                            'NOTE: if ZOC1 is missing, then we have a problem ... used to get table, but not for CBC
                            SpecimenObject.Comments = Trim$(szGetSegSeq(gszHL7RetMsg, piThisZOC, "ZOC", hlZOCSeg.Remark))  'ZOC(3) has COMMENT info
                            
                            SpecimenObject.CapColorOrContainer = szGetSegSeq(gszHL7RetMsg, piThisZOC, "ZOC", hlZOCSeg.Container)  'API bug is fixed now
                            pszTempA = szGetSegSeq(gszHL7RetMsg, piThisZOC, "ZOC", hlZOCSeg.Specimen)   'ZOC(2) has TYPE info
                            pszTempA = Trim$(pszTempA)
                            
                            If LenB(pszTempA) = 0 Then
                                  SpecimenObject.TopographyIen = ""
                                  SpecimenObject.Topography = ""
                            Else
                                  SpecimenObject.TopographyIen = szGetColumn(pszTempA, 1, "\")
                                  SpecimenObject.Topography = szGetColumn(pszTempA, 2, "\")
                            End If
                            
                'SpecimenObject.Comments = Trim$(szGetSegSeq(gszHL7RetMsg, piThisZOC, "ZOC", hlZOCSeg.Remark))
                'ZOC(3) has COMMENT info
                'prompt = Trim$(szGetSegSeq(gszHL7RetMsg, piThisZOC, "ZOC", hlZOCSeg.RequiredCommentText))
                'ZOC(4) has "Title for Questions" info
                                           
Dim szMessage As String
Dim pilength As Integer
                                
                            'This redesign sends message 'chunks' to the RequiredComments builder routine
                            If piThisZOC = iLabZocCount Then
                                szMessage = Mid$(gszHL7RetMsg, iZOC_Start(piThisZOC))
                            Else
                                szMessage = Mid$(gszHL7RetMsg, iZOC_Start(piThisZOC), _
                                    (iZOC_Start(piThisZOC + 1)) - (iZOC_Start(piThisZOC) - 1))
                            End If
                        
                            'PROBLEM: Since we are sending a partial return-message, then the first ZRQ is ALWAYS = = 1, etc
                            If (piStartValue <> 0) And (piStartValue <> 1) Then
                                pilength = piStartValue - 1
                                piStartValue = piStartValue - pilength
                                piEndValue = piEndValue - pilength
                            End If
                                            
                            If piStartValue <> 0 Then       'If ZERO, then no ZRQs for this ZOC
                                'Any ZOC may have multiple ZRQs
                                 For piThisZRQ = piStartValue To piEndValue   'must call one-by-one
                                     SpecimenObject.colRequiredComments.Add_RequiredComments piThisZRQ, szMessage
                                 Next piThisZRQ
                            End If
                         
                            'This ZOC is ready
                            Specimen.Add SpecimenObject   'added 1-10-02 AAB
                            Set SpecimenObject = Nothing
                                     
                        Next piThisZOC
                         
'''                         '24-JAN-2002; AAB: If we are doing this preverify as part of a MODIFY, then we have ZQRs which
'''                         'hold the original test RqdComments entered .... THESE now become our PRESENT DEFAULTS
'''
'''    '========================================================================================
'''
'''Dim piRqdCount As Integer
'''Dim iZQR_Count As Integer
'''Dim piSpecimenIEN As Integer
'''Dim piThisSpecimenIen As Integer
'''Dim piSpecimenIndex As Integer
'''Dim piZocCount As Integer
'''Dim piLabZocCount As Integer
'''
'''                        piSpecimenIndex = -1
'''                        piSpecimenIEN = Val(LabTablet.Specimen.IEN)
'''                        piLabZocCount = Specimen.Count
'''
'''                        For piZocCount = 1 To piLabZocCount
'''                             piThisSpecimenIen = Val(Specimen(piZocCount).IEN)
'''                            If piThisSpecimenIen = piSpecimenIEN Then
'''                                 piSpecimenIndex = piZocCount
'''                                 Exit For
'''                            End If
'''                        Next piZocCount
'''
'''
'''                        If piSpecimenIndex <> -1 Then
'''                               If Not COLL_ZQRs Is Nothing Then
'''                                       piRqdCount = Specimen.Item(piSpecimenIndex).colRequiredComments.Count
'''                                       For piTempA = 1 To piRqdCount
'''                                               For iZQR_Count = 1 To COLL_ZQRs.Count
'''                                                   With Specimen.Item(piSpecimenIndex).colRequiredComments(piTempA)
'''                                                        If .RequiredCommentID = COLL_ZQRs(iZQR_Count).AnswerText Then
'''                                                             .Default_Value = COLL_ZQRs(iZQR_Count).AnswerCode
'''                                                        End If
'''                                                   End With
'''                                               Next iZQR_Count
'''                                       Next piTempA
'''                                       Set COLL_ZQRs = Nothing
'''                               End If
'''                        End If
    
                    #End If
               
            '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                                    
                Else    'No ZOCs returned
                          Set SpecimenObject = New TypeSpecimen
                End If
          
            'ORIGINAL CODE:
                                    
            '=====================================================================
            'If only one ZOC, then see if it is complete ================================
                    If iLabZocCount = 1 Then
                        'Only ONE Zoc = SPECIAL; we may need to only get PARTIAL info back
                        'these locals are: "IEN&TEXT"
                        'sent format "IEN\TEXT" becomes output format "IEN&TEXT"
                        szZoc1 = szGetSegSeq(gszHL7RetMsg, 1, "ZOC", hlZOCSeg.CollectionSample)  'ZOC(1) has SPECIMEN and CONTAINER info
                        If LenB(szZoc1) > 0 Then
                            'KELLYMCpszZoc1IEN = szGetColumn(szZoc1, 1, "\")
                             szZoc1 = szMohel(szZoc1, 2, "\", "&")
                        Else
                            'KELLYMCpszZoc1IEN = ""
                        End If
                        
                        szZoc2 = szGetSegSeq(gszHL7RetMsg, 1, "ZOC", hlZOCSeg.Specimen)  'ZOC(2) has TOPO  info
                        If LenB(szZoc2) > 0 Then
                            szZoc2 = szMohel(szZoc2, 2, "\", "&")
                        End If
                            
                        'KELLYMCszZoc3 = Trim$(szGetSegSeq(gszHL7RetMsg, 1, "ZOC", hlZOCSeg.Remark))  'ZOC(3) has comments
                        'KELLYMCszZoc4 = Trim$(szGetSegSeq(gszHL7RetMsg, 1, "ZOC", hlZOCSeg.RequiredCommentText)) 'ZOC(4) has TITLE
                            
                        szZoc5 = Trim$(szGetSegSeq(gszHL7RetMsg, 1, "ZOC", _
                        hlZOCSeg.RequiredCommentAnswer)) 'ZOC(5) has SOURCE QUESTIONS
                                      
                        'chosen ANSWER goes to CHCS via OBR18
                        'First see if this single ZOC is complete; if yes then put into status area
                        If (LenB(szZoc1) > 0) And (LenB(szZoc2) > 0) And (LenB(szZoc5) = 0) Then
                             LabTablet.SingleSpecimenIsComplete = hlSingleSpecimen_Complete ' YES it is complete:
                        Else  'INCOMPLETE ... need pop-up; (needs QUESTION answered)
                             LabTablet.SingleSpecimenIsComplete = hlSingleSpecimen_NotComplete  'redundant
                        End If
                    End If
                      
                    If Not mbPreModifying Then HL7GetWarnings LabWarning, LabTablet
                         
        '===================================================================================================================
            Case hlORDERTYPECODE.hlRAD_ORDERS_CODE
                
                    If LenB(Trim$(RadTablet.OrderIen)) = 0 Then
                        PreVerify = hlERRORS.hlERROR_PREVERIFY_RAD_FAILED
                        Exit Function
                    End If
                    
                'Preverify means "start over"; get fresh objects
                
                'If there are real ERRORS, then quick exit leaves all other properties BLANK
                    If Not mbPreModifying Then
                        Set RadTablet = New RadType
                    End If
                    Set RadError = New Collection
                    Set RadWarning = New Collection
                    
                'Set the order info for new tablet
                    RadTablet.OrderName = szOrderName
                    RadTablet.OrderIen = szOrderIen
              
                    If LenB(RadTablet.OrderIen) = 0 Then
                        PreVerify = hlERRORS.hlERROR_PREVERIFY_RAD_FAILED
                        Exit Function
                    End If
                    Call Build_PreVerifyRad_Message(RadTablet)
                    
                    RadTablet.LastPreverifyNameWas = RadTablet.OrderName
                    RadTablet.LastPreverifyIenWas = RadTablet.OrderIen
                    
                'SEND HL7 command package >>>>>>>>>>>>>>>
                    glngTimeoutStatus = HL7MsgHandler()      'changed 6-30-98;AAB new design
                    If glngTimeoutStatus = hlCOMMSTATUS.hlCOMM_TIMEOUT Then 'section added 6-30-98; AAB
                         PreVerify = hlERRORS.hlERROR_TIMED_OUT
                         Exit Function
                    End If
                
                'FIRST look at any errors. If any, and they are not 'soft' then exit
                    iReturn = FetchErrors(RadTablet, RadError)
                    If iReturn = False Then Exit Function
                
                'get defaults
                    szOBR4 = szGetSegSeq(gszHL7RetMsg, 1, "OBR", hlOBRSeg.TestName)
                    RadTablet.OrderIen = szGetColumn(szOBR4, 1, "\")
                    RadTablet.OrderName = szGetColumn(szOBR4, 2, "\")
                    If RadTablet.LastPreverifyNameWas = RadTablet.OrderName Then
                         RadTablet.LastPreverifyNamesMatch = True
                    Else
                         RadTablet.LastPreverifyNamesMatch = False
                    End If
                
                'ZOS(5) zero node text summary
                'handled by HL7GetWarnings()
                ' RadTablet.WarningHeader = szGetSegSeq(gszHL7RetMsg, 1, "ZOS", 5)
                    
                    szOrc7 = szGetSegSeq(gszHL7RetMsg, 1, "ORC", hlORCSeg.Timing)
                  
                'Default CONTINUOUS [Admin-Time and Duration] might be in the returned ORC(7)
                    szTempb = szGetColumn(szOrc7, 2, "\") '  want first packet, but string begins with [\]
                'results are [unexpd & Admin time(s) & Freq]; eg "\&0800\D1"
                    szTempa = szGetColumn(szTempb, 2, "&")   'want Admin time(s) if any
                     
                'if more-than-one, they are comma-separated
                    If LenB(szTempa) > 0 Then
                        iTempA = InStr(szTempa, ",")
                        If iTempA <> 0 Then   'get only the FIRST time sent
                           RadTablet.DefaultAdminTime = szGetColumn(szTempa, 1, ",")
                        Else
                           RadTablet.DefaultAdminTime = szTempa
                        End If
                    End If
              
                    szRadDeftDuratn = szGetColumn(szOrc7, 3, "\")   'want Duration if any
                    If LenB(szRadDeftDuratn) = 0 Then
                         RadTablet.DefaultDuration = 0 '"D1"  **** expand this *** +++
                    Else
                         RadTablet.DefaultDuration = Right$(szRadDeftDuratn, Len(szRadDeftDuratn) - 1)
                    End If
                          
                'ZOM(1) countersign needed       CHCS=1 for Yes
                    pszTemp = szGetSegSeq(gszHL7RetMsg, 1, "ZOM", hlZOMSeg.CounterSignatureNeeded)
                    If pszTemp = "1" Then
                      RadTablet.CounterSignatureRqd = hlCOUNTERSIGN_REQUIRED
                    Else
                      RadTablet.CounterSignatureRqd = hlCOUNTERSIGN_NOT_REQUIRED
                    End If
                        
                'ZOM(2) ancillary approval needed  CHCS=1 for Yes
                    pszTemp = szGetSegSeq(gszHL7RetMsg, 1, "ZOM", hlZOMSeg.AncillaryApprovalRequired)
                    If pszTemp = "1" Then
                      RadTablet.ApprovalRqd = hlAPPROVAL_REQUIRED
                    Else
                      RadTablet.ApprovalRqd = hlAPPROVAL_NOT_REQUIRED
                    End If
            
                'ZOM(3) appointment needed     CHCS=1 for Yes
                    pszTemp = szGetSegSeq(gszHL7RetMsg, 1, "ZOM", hlZOMSeg.AppointmentRequired)
                    If pszTemp = "1" Then
                      RadTablet.AppointmentRqd = hlAPPOINTMENT_REQUIRED
                    Else
                      RadTablet.AppointmentRqd = hlAPPOINTMENT_NOT_REQUIRED
                    End If
                
                'ZOM(4) Location of Procedure
                    szZom4 = szGetSegSeq(gszHL7RetMsg, 1, "ZOM", hlZOMSeg.RadiologyLocation)
                    RadTablet.ProcedureLocationIen = szGetColumn(szZom4, 1, "\")
                    RadTablet.ProcedureLocation = szGetColumn(szZom4, 2, "\")
                
                'ZOM(5) A,S,N returned
                'PORTABILITY indicator A=always S=Sometimes N=Never; If indicator = N then dont allow
                    pszTemp = szGetSegSeq(gszHL7RetMsg, 1, "ZOM", hlZOMSeg.Portable)
                    Select Case pszTemp
                       Case "A": RadTablet.PortableAllowed = hlPORTABLE_ALWAYS
                       Case "S": RadTablet.PortableAllowed = hlPORTABLE_SOMETIMES
                       Case "N": RadTablet.PortableAllowed = hlPORTABLE_NEVER
                       Case Else: RadTablet.PortableAllowed = hlPORTABLE_SOMETIMES
                    End Select
                            
                'ZOM(6) "1" = yes          CHCS=1 for Yes
                    pszTemp = szGetSegSeq(gszHL7RetMsg, 1, "ZOM", hlZOMSeg.Continuous)
                    If pszTemp = "1" Then
                         RadTablet.ContinuousAllowed = hlCONTINUOUS_ALLOWED
                    Else
                         RadTablet.ContinuousAllowed = hlCONTINUOUS_NOT_ALLOWED
                    End If
                
                'Y/N ZOM(7) 1 = required; Allergy Override Required Flag  CHCS=1 for Yes
                    pszTemp = szGetSegSeq(gszHL7RetMsg, 1, "ZOM", hlZOMSeg.AllergyOverrideRequired)
                    If pszTemp = "1" Then
                         RadTablet.AllergyRqd = hlallergy_required
                    Else
                         RadTablet.AllergyRqd = hlALLERGY_NOT_REQUIRED
                    End If
                    
                    If Not mbPreModifying Then HL7GetWarnings RadWarning, RadTablet
                
        '===================================================================================================================
            Case Else   'added 7-1-98;AAB need to micro-manage this call due to bad HL7 status codes and behavior.
                    PreVerify = hlERROR_BAD_ORDERTYPE_CODE
                    Exit Function
            End Select
               
End Function
Public Function ExpandSig(RawSigText As String) As hlERRORS

      If gbHL7LoginStatus = False Then
        ExpandSig = hlERRORS.hlERROR_NOT_LOGGED_IN
        Exit Function
      End If
    
      ExpandSig = hlERRORS.hlERROR_NONE  'changed 4-6-98; AAB
      If LenB(Trim(RawSigText)) = 0 Then
               ExpandSig = hlERRORS.hlERROR_NO_SIG_TEXT
               Exit Function
      End If

      'SIG  ... go expand it ---------------------------------------
      
      Call Build_PreVerifyRx_Message(RXTablet, RawSigText)

          '>>>>   sent sig              in RXO(6)
          '<<<<   get expanded-sig      in RXO(7)
          
     'SEND HL7 command package >>>>>>>>>>>>>>>
          glngTimeoutStatus = HL7MsgHandler()  'changed 6-30-98;AAB new design
      
      If glngTimeoutStatus = hlCOMMSTATUS.hlCOMM_TIMEOUT Then
               ExpandSig = hlERRORS.hlERROR_TIMED_OUT
               RXTablet.ExpandedSig = ""
      Else
          If gszHL7RetCode = "AA" Then
               szTempa = szGetSegSeq(gszHL7RetMsg, 1, "RXO", hlRXOSeg.ExpSigCode) 'expanded sig text
               'ExpandSig = szTempA  'removed 4-6-98; AAB keep as status return only
               'Use property for results
               RXTablet.ExpandedSig = szTempa  'added 11-12-97
               szTempa = szGetSegSeq(gszHL7RetMsg, 1, "RXO", hlRXOSeg.Quantity) 'expanded sig text
               RXTablet.Quantity = szTempa  'added 11-12-97
          Else
               ExpandSig = hlERRORS.hlERROR_EXPAND_SIG_FAILED
               RXTablet.ExpandedSig = ""
          End If
      End If

End Function

Public Function SendOverrides(OrderTypeCode As hlORDERTYPECODE) As hlERRORS

'KELLYMCDim piNumErrs As Integer
'KELLYMCDim FetchErrorsReturn As hlERRMSGFLAG
    If gbHL7LoginStatus = False Then
      SendOverrides = hlERRORS.hlERROR_NOT_LOGGED_IN
      Exit Function
    End If
    If LenB(gszProviderIEN) = 0 Then
        SendOverrides = hlERRORS.hlERROR_NO_PROVIDER_IEN
        Exit Function
    End If
    
    SendOverrides = hlERRORS.hlERROR_NONE

    '===================================================================
    '===================================================================
     'Build_SendOverRide_Message
            Select Case OrderTypeCode
            Case hlORDERTYPECODE.hlRX_ORDERS_CODE
              Build_SendOverRide_Message RXTablet, RxWarning
            Case hlORDERTYPECODE.hlLAB_ORDERS_CODE
              Build_SendOverRide_Message LabTablet, LabWarning
            Case hlORDERTYPECODE.hlRAD_ORDERS_CODE
              Build_SendOverRide_Message RadTablet, RadWarning
            Case hlORDERTYPECODE.hlCON_ORDERS_CODE
              Build_SendOverRide_Message OtherTablet, OtherWarning
            End Select

     'SEND HL7 command package >>>>>>>>>>>>>>>
          glngTimeoutStatus = HL7MsgHandler()      'changed 6-30-98;AAB new design
      
          If glngTimeoutStatus = hlCOMMSTATUS.hlCOMM_TIMEOUT Then 'section added 6-30-98; AAB
            SendOverrides = hlERRORS.hlERROR_TIMED_OUT
            Exit Function
          End If
      
    '=======================    ERRORS AND WARNINGS ====================
    'ERRORS  ===========================================================
      
      Select Case OrderTypeCode
            
       '=====================================================================================
       Case hlORDERTYPECODE.hlRX_ORDERS_CODE
         Set RxError = New Collection
            
         'KELLYMCFetchErrorsReturn = FetchErrors(RXTablet, RxError)
        FetchErrors RXTablet, RxError

'            order.ErrorCondition3 = szGetSegSeq(gszHL7RetMsg, 1, "MSA", hlMSASeg.ErrorText)
'            order.ErrorCondition6 = szGetSegSeq(gszHL7RetMsg, 1, "MSA", hlMSASeg.ErrorCondition)
         
      '=====================================================================================
      Case hlORDERTYPECODE.hlLAB_ORDERS_CODE
        Set LabError = New Collection
          
        'KELLYMCFetchErrorsReturn = FetchErrors(LabTablet, LabError)
        FetchErrors LabTablet, LabError

'          order.ErrorCondition3 = szGetSegSeq(gszHL7RetMsg, 1, "MSA", hlMSASeg.ErrorText)
'          order.ErrorCondition6 = szGetSegSeq(gszHL7RetMsg, 1, "MSA", hlMSASeg.ErrorCondition)
        
      '=====================================================================================
      Case hlORDERTYPECODE.hlRAD_ORDERS_CODE
        Set RadError = New Collection
        'KELLYMCFetchErrorsReturn = FetchErrors(RadTablet, RadError)
        FetchErrors RadTablet, RadError

'          order.ErrorCondition3 = szGetSegSeq(gszHL7RetMsg, 1, "MSA", hlMSASeg.ErrorText)
'          order.ErrorCondition6 = szGetSegSeq(gszHL7RetMsg, 1, "MSA", hlMSASeg.ErrorCondition)
        
      End Select

    'WARNINGS ==========================================================

      Select Case OrderTypeCode
      
      Case hlORDERTYPECODE.hlRX_ORDERS_CODE
        Set RxWarning = New Collection
        If gszHL7RetCode <> "AA" Then
          HL7GetWarnings RxWarning, RXTablet
        End If
          
      Case hlORDERTYPECODE.hlLAB_ORDERS_CODE
        Set LabWarning = New Collection
        If gszHL7RetCode <> "AA" Then
          HL7GetWarnings LabWarning, LabTablet
        End If
          
      Case hlORDERTYPECODE.hlRAD_ORDERS_CODE
        Set RadWarning = New Collection
        If gszHL7RetCode <> "AA" Then
          HL7GetWarnings RadWarning, RadTablet
        End If
          
      End Select

End Function

Public Function SendOrder(OrderTypeCode As hlORDERTYPECODE) As hlERRORS
    

Dim iDuration As Integer
Dim pszClinical As String

               If gbHL7LoginStatus = False Then
                    SendOrder = hlERRORS.hlERROR_NOT_LOGGED_IN
                    GoTo SendOrderIncomplete:
               End If
               If LenB(gszProviderIEN) = 0 Then
                    SendOrder = hlERRORS.hlERROR_NO_PROVIDER_IEN
                    GoTo SendOrderIncomplete:
               End If
    
               SendOrder = hlERRORS.hlERROR_NONE
    
          '==========================================================================
          'HL7 bug; comments must be 3-78 chars long; 0-2 lengths will return an error
          'initially thought only for RX, 11-29-95
          '   learned that ALL test areas are like this; fixed 11-30-95
          
          'ALSO CHCS API bug cannot allow times of 00:00; fixed to comply 11-30-95
          '+++ but should also test UNEXPANDED Time for, eg "T +1@ 0000"
          '==========================================================================
 
Select Case OrderTypeCode
'*********************************************************************************
Case hlORDERTYPECODE.hlRX_ORDERS_CODE
    
     Dim pszRxQuantity As String
     Dim pszRxMaxQuant As String
    
    'Clear out possible old orderid in tablet
    RXTablet.OrderId = ""
    
          'RX CHECKS & MEASURES __________________________________________________________
                  
          '10-9-95 if we HAVE warnings but NO override endorsement, then go away
               If (RxWarning.Count <> 0) And (LenB(RXTablet.WarningReason) = 0) Then
                      SendOrder = hlERRORS.hlERROR_RX_REASON
                      GoTo SendOrderIncomplete:
               End If
               
               If (LenB(RXTablet.OrderName) = 0) Then
                    SendOrder = hlERRORS.hlERROR_RX_NAME
                    GoTo SendOrderIncomplete:
               End If
               
''taken out '03-30-00; AAB HARD
''               If (RXTablet.OrderIen = "") Then
''                    SendOrder = hlERRORS.HLERROR_RX_ORDERIEN
''                    GoTo SendOrderIncomplete:
''               End If
               
               '    '6-19-95-AAB; entered # and RF will be placed on the SIG.
               '    ' -- If the SIG had Embedded # or RF, then these will be replaced;
               '    ' -- therefore ALL SIGS will have #/RF embedded for submittal
    
               pszRxQuantity = RXTablet.Quantity
         
          '4-10-97; if QUANT = "0" then treat it as if BLANK; P2451; CCB-266
               If pszRxQuantity = "0" Then pszRxQuantity = ""
          'bug fix ... incompletion; 3-15-98
               pszRxMaxQuant = Trim$(RXTablet.MaxQuantity)
     
          '4-23-97; P1202 work-around
          'If QUANT > MaxQuant, then a comment is required
               If LenB(pszRxMaxQuant) > 0 Then
                    If Val(pszRxQuantity) > Val(pszRxMaxQuant) And (LenB(Trim$(RXTablet.Comment)) = 0) Then
                         SendOrder = hlERRORS.hlERROR_RX_COMMENT_RQD
                         GoTo SendOrderIncomplete:
                    End If
               End If
      
               If LenB(Trim$(RXTablet.SigCode)) = 0 Then            '4-3-96 trim$
                 SendOrder = hlERRORS.hlERROR_RX_NOSIG  '"SIG Code Missing"  "A SIG code is required"
                 GoTo SendOrderIncomplete:
               End If
                            
               If LenB(RXTablet.OrderDate) = 0 Then
                 SendOrder = hlERRORS.hlERROR_RX_DATE   '"DATE Missing" "A valid DATE is required"
                 GoTo SendOrderIncomplete:
               End If
    
               iTempA = Len(Trim$(RXTablet.Comment))
               If iTempA <> 0 Then
                    If iTempA < 3 Or iTempA > 78 Then
                         SendOrder = hlERRORS.hlERROR_RX_INVALID_COMMENT  'pszCmmtLenErrorText; "Invalid Comment"
                         GoTo SendOrderIncomplete:
                    End If
               End If
        
          '============================================================================
          'update SIG to reflect hard-coding of # RF fields ... REMOVED this functionality
          '============================================================================
     
          'SIG has NO MORE embedded # or RF; add ACTUAL values now
          ' -- append New # and RF to stripped SIG:
        
          '===============================================================================
          'RX Post To SUMMARY_____________________________________________________________
          
          'RxTablet.OrderName = ... obtained at Pre-Verify
          'RxTablet.OrderIen  = ... obtained at Pre-Verify
          
          'added 4-7-98 to test direct-sends, such as SETS records
               RXTablet.LastSentNameWas = RXTablet.OrderName
               RXTablet.LastSentIenWas = RXTablet.OrderIen
               
               lngTempA = Build_SendNewRx_Message(RxWarning, RXTablet)
               
               If lngTempA <> hlERRORS.hlERROR_NONE Then
                         SendOrder = lngTempA
                         GoTo SendOrderIncomplete:
               End If
    
          'SEND HL7 command package >>>>>>>>>>>>>>>
               glngTimeoutStatus = HL7MsgHandler()      'changed 6-30-98;AAB new design
               
               If glngTimeoutStatus = hlCOMMSTATUS.hlCOMM_TIMEOUT Then 'section added 6-30-98; AAB
                         SendOrder = hlERRORS.hlERROR_TIMED_OUT
                         Exit Function
               End If
          
               '      The Orders DLL should fill the
               '      RxTablet.Quantity property after return from SendNeworder.
               '      This is needed because a successful order can be placed without a specific quantity
               '      requested; the quantity was calculated by CHCS from the SIG-code.
    
               Call GetRequestedInformation(RXTablet)
                         
               If RXTablet.OrderName = RXTablet.LastSentNameWas Then
                    RXTablet.LastSentNamesMatch = True
               Else
                    RXTablet.LastSentNamesMatch = False
               End If
          
          'NOTES: 5-9-96: '+++
          'DESIGN consideration 5-9-96; ZPD says: "Dispensing Pharmacy is the default pharmacy based
          'on requesting location OR DISPENSING PHARMACY OBTAINED FROM THE SIG CODE."
          '--------------------------------------------------------------------------
          
    '      '4-9-96; if the QUANTITY was interpreted from the SIG,
    '      'then act as if it were hard-coded
    '
    '    If (Trim$(pszRxQuantity) = "") Then
    '      'use return data from RXO(11); QUANTITY ... calculated if necessary
    '      'also,  RXO (6) is RAW SIG
    '      '       RXO (7) is EXPANDED SIG
    '
    '      returned quantity = szGetSegSeq(gszHL7RetMsg, 1, "RXO", 11)
    '      encoded sig = szGetSegSeq(gszHL7RetMsg, 1, "RXO", 6)
    '      expanded sig = szGetSegSeq(gszHL7RetMsg, 1, "RXO", 7)
    '
    '      sz'MsgBoxText = "The QUANTITY for this order," & gchCrlf
    '      sz'MsgBoxText = sz'MsgBoxText & "was interpreted from the SIG Code:" & gchCrlf
            
    
        'ERRORS AND WARNINGS ==========================
        'ERRORS  =========================================
            Dim RxOldErrors As Collection
            Set RxOldErrors = RxError
            Set RxError = New Collection
            'KELLYMCFetchErrorsReturn = FetchErrors(RXTablet, RxError)
             FetchErrors RXTablet, RxError

            'find dup errors and remove them
            Dim iOld, iNew As Integer
            
            For iOld = RxOldErrors.Count To 1 Step -1
                For iNew = RxError.Count To 1 Step -1
                    If RxOldErrors.Item(iOld).TypeCode = RxError.Item(iNew).TypeCode Then
                        RxError.Remove iNew
                    End If
                Next
            Next
        'WARNINGS ======================================
            Set RxWarning = New Collection
            
            'If gszHL7RetCode <> "AA" Then   'BUG: 6-25-98;
            'secondary warnings come with "AA"
            'Fixed in three palces ... now same as preverify
            '===============================================
            
               HL7GetWarnings RxWarning, RXTablet
            
'**************************************************************************************************
Case hlORDERTYPECODE.hlLAB_ORDERS_CODE
        'Clear out possible old orderid in tablet
        LabTablet.OrderId = ""
        
          'LAB CHECKS & MEASURES ___________________________________________

          '10-9-95 if we HAVE warnings but NO override endorsement, then go away
               If (LabWarning.Count <> 0) And (LenB(LabTablet.WarningReason) = 0) Then
                    SendOrder = hlERRORS.hlERROR_LAB_REASON
                    GoTo SendOrderIncomplete:
               End If
    
               If (LenB(LabTablet.OrderName) = 0) Then
                    SendOrder = hlERRORS.hlERROR_LAB_NAME
                    GoTo SendOrderIncomplete:
               End If
              
          'LAB COLLECT-priority is in ORC(7); PROCESS-priority is in OBR(5)
            
          'for both 1-time and continuous:     '-------------------------------------------
               If (LenB(LabTablet.Timing.StartDate) = 0) Then
                 '5-29-96 editing sets has blank date
                    SendOrder = hlERRORS.hlERROR_LAB_DATE
                    GoTo SendOrderIncomplete:
               End If    '------------------------------------------------------------------
                      
               If LabTablet.Timing.OnexOrContinued = hlONETIME Then
                         '4-11-96; ONETIME LABS can send Admin-Times ...
                         'Only start DATE needed for one-time
                         
                         'GetOrc7 call moved to Build_SendNewLab_Message   AMW 10/20/99
               Else
                      'LAB CONTINUOUS ... needs:  Start DATE [[not time], and no end time] ,
                      '  Admin time(s) or Unexpanded Time, Frequency, and Duration
                      
                      'we need only ONE of the following: Time(LAB) Time(0) or Unexp Time
                         
                       'remmed out for testing 9/21/99 amw
                         If LenB(Trim$(LabTablet.Timing.AdminTime)) = 0 Then
                              SendOrder = hlERRORS.hlERROR_LAB_ADMIN
                              GoTo SendOrderIncomplete:
                         End If
                 
                         If LenB(LabTablet.Timing.Frequency) = 0 Then
                              SendOrder = hlERRORS.hlERROR_LAB_FREQUENCY '"CONTINUOUS needs Frequency"
                              GoTo SendOrderIncomplete:
                         End If
           
                         iDuration = LabTablet.Timing.Duration
                         If iDuration = 0 Then
                              SendOrder = hlERRORS.hlERROR_LAB_DURATION_RQD  '"CONTINUOUS needs Duration"
                              GoTo SendOrderIncomplete:
                         End If
                 
                         If iDuration < 2 Then   'must be GT 1
                              SendOrder = hlERRORS.hlERROR_LAB_BAD_DURATION_VALUE  '"Duration must be TWO or more Days "
                              GoTo SendOrderIncomplete:
                         End If
           
                      ''GetOrc7 call moved to Build_SendNewLab_Message   AMW 10/20/99
                      
               End If
          
               iTempA = Len(Trim$(LabTablet.Comment))
               If iTempA <> 0 Then
                    If iTempA < 3 Or iTempA > 78 Then
                         SendOrder = hlERRORS.hlERROR_LAB_INVALID_COMMENT
                         GoTo SendOrderIncomplete:
                    End If
               End If
    
          '===============================================================================
          'LabTablet.OrderName = loaded from preverify
          'LabTablet.OrderIen  = loaded from preverify
          
          'LAB COST NOT AVAILABLE with ver 5.0
                  
          'added 4-7-98 to test direct-sends, such as SETS records
               LabTablet.LastSentNameWas = LabTablet.OrderName
               LabTablet.LastSentIenWas = LabTablet.OrderIen
               
               lngTempA = Build_SendNewLab_Message(LabWarning, LabTablet)
               If lngTempA <> hlERRORS.hlERROR_NONE Then
                    SendOrder = lngTempA
                    GoTo SendOrderIncomplete:
               End If
          
          'SEND HL7 command package >>>>>>>>>>>>>>>
               glngTimeoutStatus = HL7MsgHandler()                'changed 6-30-98;AAB new design
               
               If glngTimeoutStatus = hlCOMMSTATUS.hlCOMM_TIMEOUT Then   'section added 6-30-98; AAB
                    SendOrder = hlERRORS.hlERROR_TIMED_OUT
                    Exit Function
               End If
    
               Call GetRequestedInformation(LabTablet)
    
          'ERRORS AND WARNINGS ============================
          'ERRORS  ===========================================
               Set LabError = New Collection
               'KELLYMCFetchErrorsReturn = FetchErrors(LabTablet, LabError)
               FetchErrors LabTablet, LabError

          'WARNINGS ========================================
               Set LabWarning = New Collection
               HL7GetWarnings LabWarning, LabTablet
        
'*********************************************************************************
Case hlORDERTYPECODE.hlRAD_ORDERS_CODE
    'Clear out possible old orderid in tablet
    RadTablet.OrderId = ""
          'RAD CHECKS & MEASURES ======================================================================

          '10-9-95 if we HAVE warnings but NO override endorsement, then go away
               If (RadWarning.Count <> 0) And (LenB(RadTablet.WarningReason) = 0) Then
                    SendOrder = hlERRORS.hlERROR_RAD_REASON
                    GoTo SendOrderIncomplete:
               End If
        
               If (LenB(RadTablet.OrderName) = 0) Then
                    SendOrder = hlERRORS.hlERROR_RAD_NAME
                    GoTo SendOrderIncomplete:
               End If
    
               pszClinical = Trim$(RadTablet.ClinicalImpressionText)
               If (LenB(pszClinical) = 0) Then   '"A Clinical Impression MUST be entered"
                    SendOrder = hlERRORS.hlERROR_RAD_CI
                    GoTo SendOrderIncomplete:
               Else
                    'make array strings and count for the NTE segments
                    Call GetCIarray(pszClinical)
               End If
    
        'for both continuous or 1X:
               If (LenB(RadTablet.Timing.StartDate) = 0) Then
                    SendOrder = hlERRORS.hlERROR_RAD_DATE
                    GoTo SendOrderIncomplete:
               End If
          
          '4-11-96 CANNOT specify Admin Time with 1X ...
          'Only start DATE needed for one-time
            If RadTablet.Timing.OnexOrContinued = hlRADLAB.hlCONTINUE Then
                    'RAD CONTINUOUS ... needs:  Start DATE [[not time], and no end time] ,
                    ' -- Admin time(s) or Unexpanded Time, Frequency, and Duration
                
                     iDuration = RadTablet.Timing.Duration
                     If iDuration = 0 Then
                          '"CONTINUOUS needs Duration"
                              SendOrder = hlERRORS.hlERROR_RAD_DURATION_RQD
                              GoTo SendOrderIncomplete:
                     End If
                    
                     If iDuration < 2 Then  'must be GT 1
                          '"Duration must be greater than ONE"
                              SendOrder = hlERRORS.hlERROR_RAD_BAD_DURATION_VALUE
                              GoTo SendOrderIncomplete:
                     End If
                
                'We need only ONE of the following: Time(RAD)  or Unexp Time
                '
                'If Trim$(RadTablet.Timing.AdminTime) = "" Then
                '10-20-97 AAB; we need to see if unexpanded is also present ...
                   If LenB(Trim$(RadTablet.Timing.AdminTime)) = 0 _
                        And (LenB(RadTablet.Timing.UnexpandedTime) = 0) _
                        Then
                            SendOrder = hlERRORS.hlERROR_RAD_ADMIN
                            GoTo SendOrderIncomplete:
                   End If
                
                   If (LenB(RadTablet.Timing.Frequency) = 0) Then
                        SendOrder = hlERRORS.hlERROR_RAD_FREQUENCY '    "CONTINUOUS needs Frequency"
                        GoTo SendOrderIncomplete:
                   End If
              
                  '========================================================
                  'Build the TQ timing element
            
            End If  '1x or continuous
        
            'GetOrc7 call moved to Build_SendNewLab_Message   AMW 10/20/99
    
            '==========================================================
            'comments
               iTempA = Len(Trim$(RadTablet.Comment))
               If iTempA <> 0 Then
                    If iTempA < 3 Or iTempA > 78 Then
                         SendOrder = hlERRORS.hlERROR_RAD_INVALID_COMMENT   'pszCmmtLenErrorText
                         GoTo SendOrderIncomplete:
                    End If
               End If
    
            'if ONE-TIME then all we need is the start date
        
            'RAD Post To SUMMARY______________________________________________________
           
            ''RadTablet.OrderName = loaded from preverify
            ''RadTablet.OrderIen  = loaded from preverify
                 
            'INDEX_RadMobile points to TAG needed for OBR(30)
             
            'NO COLLECT PRIORITY FOR RAD:
            '
            '
            'COST NOT AVAILABLE with rlse 5.0
    
            'added 4-7-98 to test direct-sends, such as SETS records
               RadTablet.LastSentNameWas = RadTablet.OrderName
               RadTablet.LastSentIenWas = RadTablet.OrderIen
                         
               lngTempA = Build_SendNewRad_Message(RadWarning, RadTablet)
               If lngTempA <> hlERRORS.hlERROR_NONE Then
                    SendOrder = lngTempA
                    GoTo SendOrderIncomplete:
               End If
          
          'SEND HL7 command package >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
               glngTimeoutStatus = HL7MsgHandler()      'changed 6-30-98;AAB new design
               
               If glngTimeoutStatus = hlCOMMSTATUS.hlCOMM_TIMEOUT Then 'section added 6-30-98; AAB
                    SendOrder = hlERRORS.hlERROR_TIMED_OUT
                    Exit Function
               End If
              
               Call GetRequestedInformation(RadTablet)
               
          'ERRORS AND WARNINGS ============================
          'ERRORS  ===========================================
               Set RadError = New Collection
                    'KELLYMCFetchErrorsReturn = FetchErrors(RadTablet, RadError)
                FetchErrors RadTablet, RadError
 
          'WARNINGS ========================================
               Set RadWarning = New Collection
               HL7GetWarnings RadWarning, RadTablet
               
End Select

    '*********************************************************
    Exit Function
    '*********************************************************

    'Error code was set before branching here:
    
SendOrderIncomplete:                   'added 3-25-96 for modify-update bug mgmt

End Function

Public Function GetOrders(hlOrderType As hlORDERTYPECODE) As hlERRORS

Dim szOrderCode As String
Dim szOrderStatus As String
Dim szOrderNum As String
Dim szOrderDesc As String
Dim szModifiable As String
Dim szExpirationDate As String
Dim szLastDispenseDate As String
Dim szLastFillDate As String
Dim szRemainingRF As String
Dim szOrderCompleteDesc As String
Dim szTempDescription As String
Dim szNewStr As String
          'Dim SampleZOD As String     'for design/debug only ***
Dim i As Long
Dim iLoopCount As Integer

Dim bMatch As Integer
Dim StrPos As Integer

               'Dim szOrderTypeDesc As String

     'FOR PR296, we ALWAYS clear the Summary so old information won't persist with an error
     '11-01-00 AAB
    'Erase Summary Collection and begin new:
      Set Summary = New Collection
      Dim SummaryRecord As SummaryType

     '=============================================

    If gbHL7LoginStatus = False Then
      GetOrders = hlERRORS.hlERROR_NOT_LOGGED_IN
      Exit Function
    End If
   
    GetOrders = hlERRORS.hlERROR_NONE

    '===================================================================
    '===================================================================
    
'    SampleZOD = "MSH^\|~&^CHCS\ORE^^PWS\ORE^^19970911085800^^ZPW\ACK^F1756-171292^P^2." _
'& "2 ^^^^^ MSA^AA^PWS08206513958^^^^^" _
'& "PID^1^^13959\\\F1756^20/111-11-1119^TEST\PATIENT\\\\^^19751212000000^M" _
'& "^^C^341 WEST TEST STREET\\SCOTT A F" _
'& "B\ILLINOIS\62225\\\^^555-5555^555-5555^^S^UNK^^111-11-1119^^^^^^^^4222" _
'& "\USAF DECEASED SPONSOR\99PAT\\\^ ZOP^1025^^3^^" _
'& "ZOD^951006-00752^4^1^ARTERIAL BLOOD GASES~SEND PATIENT TO LAB~BLOOD~SYRINGE " _
'& "on 06 Oct 1995^ACTIVE^19961005101800^" _
'& "ZOD^960222-01719^4^1^*STAT*STAT* ARTERIAL BLOOD GASES~SEND PATIENT TO" _
'& "LAB~BLOOD~SYRINGE   on 22 Feb 1996@1415^ACTIVE^19970221141400^" _
'& "ZOD^970911-00002^9^^ TERBUTALINE (BRETHINE)--PO 5MG TAB~TAKE 1 TABLET" _
'& "4 TIMES A DAY RF5 #45 DS30  on 11 Sep 1997@0848^ACTIVE^^" _
'& "ZOD^970911-00003^9^^ POTASSIUM ACETATE--INJ 2MEQ/ML SOLN~SHR T1 BID" _
'& "RF0 #60 DS30  on 11 Sep 1997@0849^ACTIVE^^" _
'& "ZOD^970911-00004^4^^*STAT*STAT* ALBUMIN~SEND PATIENT TO LAB~BODY FLUID" _
'& " 0800  QD for 5 days  Starting on 12 Sep 1997@0800^ACTIVE^19970917075900^ ZOD^970911-00005^4^^*STAT*ASAP*" _
'& "MONOSPOT~SEND PATIENT TO LAB~BLOOD~RED TOP  0800  QD for 5 days" _
'& "Starting on 12 Sep 1997@0800^ACTIVE^19970917075900^ZOD^970911-00006^4^^*STAT* GLUCOSE TOLERANCE TEST 2HR~SEND PATIENT TO" _
'& "LAB~BLOOD~NA FLOURIDE  0800  QD for 5 days  Starting on 12 Sep" _
'& "1997@0800^ACTIVE^19970917075900^ ZOD^970911-00007^5^^*STAT*" _
'& "ANGIOGRAM,AORTIC ARCH~ANGIOGRAPHY~AMB  on 11 Sep 1997@0851^ACTIVE^^" _
'& "ZOD^970911-00008^5^^ US,ABDOMINAL~ULTRASOUND~AMB  on 11 Sep" _
'& "1997@0851^ACTIVE^^ ZOD^970911-00009^5^^*PREOP* mri,angiogram~MRI~AMB" _
'& "on 11 Sep 1997@0851^PENDING^^ ZOD^970911-00010^9^^ HEPARIN 25000 UNITS" _
'& "IN 1/2 NS 500ML--IV ~T2 BID RF0 #1 DS30  on 11 Sep 1997@0855^ACTIVE^^"
    
    
'===================================================================
    
    'CALL HANDLER HERE and wait for return
    
       
    FmtOrdersOnPage
    
     'SEND HL7 command package >>>>>>>>>>>>>>>
          glngTimeoutStatus = HL7MsgHandler()  'changed 6-30-98;AAB new design
     
          If glngTimeoutStatus = hlCOMMSTATUS.hlCOMM_TIMEOUT Then
               GetOrders = hlERRORS.hlERROR_TIMED_OUT
               Exit Function
          Else
               'added 4-7-98; AAB
                    If gszHL7RetCode = "AA" Then
                         GetOrders = hlERRORS.hlERROR_NONE
                    Else
                         GetOrders = hlERRORS.hlERROR_GET_ORDERS_FAILED
                         'PR296: Quick exit should erase previous patient information
                         FetchErrors OtherTablet, OtherError
                         Exit Function
                    End If
          End If
        
    'loaded at select-patientgszDefOrderPage = szGetSegSeq(gszHL7RetMsg, 1, "ZOP", hlZOPSeg.PageIEN)
    
        ''*** DESIGN / DEBUG ***
        ''gszHL7RetMsg = SampleZOD
'gszHL7RetMsg = "MSH^\|~&^CHCS\ORE^^PWS\ORE^^20000606124400^^ZPW\ACK^A0101-63619739^P^2.2^^^^^" _
'& "MSA^AA^PWS0041008633^^^^^" _
'& Chr$(13) & "PID^1^^255772\\A0101^20/001-90-0228^TESTZZ\BETSY\\\\^^19960102000000^F^^N^226 PA0 ST\\KANEOHE\HAWAII\96744\^^257-8876^257-2778^^M^^^001-90-0228^^^1^^^^^4199\VETERANS ADMIN BENEFICIARY\99PAT\K61\\99PTC^" _
'& Chr$(13) & "ZOP^105^^3^^" _
'& Chr$(13) & "ZOD^991222-04286^30^1^ PEDIATRIC ASTHMA EDUCATION at PEDIATRIC SPECIALTY CLINIC  on 22 Dec 1999@111023^PENDING^^^" _
'& Chr$(13) & "ZOD^000410-00033^4^^ CERULOPLASMIN~SEND PATIENT TO LAB~BLOOD~RED  on 10 Apr 2000@NOW^ACTIVE^^^" _
'& Chr$(13) & "ZOD^000418-00004^4^^ HIV~SEND PATIENT TO LAB~BLOOD~MARB/RED  on 20 Mar 2000@1443^ACTIVE^^^" _
'& Chr$(13) & "ZOD^000418-00007^5^^ CT, HEAD W/CONTRAST~CAT SCAN~AMB  on 20 Mar 2000^PENDING^^^" _
'& Chr$(13) & "ZOD^000424-00007^4^^ HIV~SEND PATIENT TO LAB~BLOOD~MARB/RED  on 20 Mar 2000@1443^ON HOLD^^^" _
'& Chr$(13) & "ZOD^000424-00012^4^^ HIV~SEND PATIENT TO LAB~BLOOD~MARB/RED  on 20 Mar 2000@1443^ACTIVE^^^" _
'& "ZOD^000429-00022^9^^ ACETAMINOPHEN--PO 325MG TAB~T UD Q4-6H PRN (MAXIMUM 12 TB DAILY) RF2 #50 DS30  on 29 Apr 2000^ACTIVE^^(Not Dispensed)^" _
'& "ZOD^000429-00023^9^^ ACETAMINOPHEN--PO 325MG TAB~T UD Q4-6H PRN (MAXIMUM 12 TB DAILY) RF2 #50 DS30  on 29 Apr 2000^ON HOLD^^^" _
'& Chr$(13) & "ZOD^000429-00025^4^^ HI-KODAK DT FORS HEAD~SEND PATIENT TO LAB~OTHER MISC  on 27 Apr 2000@1503^ACTIVE^^^" _
'& Chr$(13) & "ZOD^000502-00030^9^^ BUPIVACAINE 0.75%/DEXTR 8.25%--INJ AMP~AS DIRECTED RF2 #1 DS30  on 02 May 2000^ACTIVE^^(Order has changed: Expand to View)^" _
'& Chr$(13) & "ZOD^000502-00035^4^^*PREOP*NOTIFY* ANAEROBIC GM ST/CULTURE PANEL~SEND PATIENT TO LAB~ABSCESS  on 03 May 2000@NOW^ACTIVE^^^" _
'& Chr$(13) & "ZOD^000502-00038^5^^*STAT* CT, HEAD W/CONTRAST~CAT SCAN~STRETCH  on 04 May 2000   [CANCELLED] ^ON HOLD^^^" _
'& Chr$(13) & "ZOD^000502-00041^5^^*PREOP* CT, HEAD W/CONTRAST~CAT SCAN~STRETCH  on 08 May 2000^PENDING^^^" _
'& Chr$(13) & "ZOD^000502-00044^9^^ ACETAMINOPHEN--PO 325MG TAB~T UD Q4-6H PRN (MAXIMUM 12 TB DAILY) RF2 #50 DS30  on 02 May 2000^ACTIVE^^(Order has changed: Expand to View)^" _
'& Chr$(13) & "ZOD^000503-00037^4^^ GAMMA GLUTAMYL TRANSFERASE~SEND PATIENT TO LAB~BLOOD~RED  on 03 May 2000@NOW^ACTIVE^^^" _
'& Chr$(13) & "ZOD^000504-00011^9^^ ESTROG/M-TEST (ESTRATEST)--PO TAB~T1 TAB PO QD RF5 #100 DS100  on 04 May 2000^ACTIVE^^(Order has changed: Expand to View)^" _
'& Chr$(13) & "ZOD^000504-00015^9^^ ACETAMINOPHEN--PO 325MG TAB~T UD Q4-6H PRN (MAXIMUM 12 TB DAILY) RF2 #50 DS30  on 04 May 2000^ACTIVE^^(Not Dispensed)^" _
'& Chr$(13) & "ZOD^000504-00016^4^^ 24HR URINE ELECTROLYTES~SEND PATIENT TO LAB~24HR URINE  on 04 May 2000@NOW^ACTIVE^^^" _
'& Chr$(13) & "ZOD^000504-00017^5^^ ANKLE LT (WEIGHT BEARING)~DIAGNOSTIC RAD ORTHO/XRAY~AMB  on 04 May 2000^ACTIVE^^^" _
'& Chr$(13) & "ZOD^000508-00005^4^^ CHEM 4 PANEL~SEND PATIENT TO LAB~BLOOD~RED  on 08 May 2000@NOW^ACTIVE^^^"
'"ZOD^000508-00006^4^^ 24HR URINE CREATININE~SEND PATIENT TO LAB~24HR URINE  on 08 May 2000@NOW^ACTIVE^^^
'"ZOD^000509-00014^30^^ DIABETES EDUCATION INDIVIDUAL at COMMUNITY HEALTH NURSE  on 09 May 2000@1431^PENDING^^^
'"ZOD^000510-00006^4^^ CBC WITH DIFF~SEND PATIENT TO LAB~BLOOD~LAV  on 10 May 2000@NOW^ACTIVE^^^
'"ZOD^000510-00007^4^^ THROAT CULTURE R/O GP A STREP~SEND PATIENT TO LAB~THROAT SWAB  on 10 May 2000@NOW^ACTIVE^^^
'"ZOD^000510-00009^9^^ AZITHROMYCIN--PO 250MG TAB~T 4 TBS NOW RF0 #4 DS30  on 10 May 2000^ON HOLD^^^
'"ZOD^000512-00026^9^^ DILUENT FOR REGULAR INSULIN--SQ INJ~AS DIRECTED RF1 #1 DS30  on 12 May 2000^ACTIVE^^(Not Dispensed)^
'"ZOD^000512-00027^4^^ CHEM 7 PANEL~SEND PATIENT TO LAB~BLOOD~RED  on 12 May 2000@NOW^ACTIVE^^^
'"ZOD^000515-00024^5^^*ASAP* MRI, ANGIO CAROTID W/O GAD (GROUP)~MAGNETIC RESONANCE IMAGING~AMB  on 15 May 2000^PENDING^^^
'"ZOD^000517-00034^5^^*STAT* ANKLE LT~DIAGNOSTIC RADIOLOGY~AMB  on 17 May 2000^ACTIVE^^^
'"ZOD^000518-00004^9^^ BUTALBITAL/ASA/CAFFEINE(FIORINAL)--PO TA~TAKE TWICE DAILY RF1 #300 DS30  on 18 May 2000^ACTIVE^^(Order has changed: Expand to View)^
'"ZOD^000518-00005^9^^ ACETAMINOPHEN--PO 325MG TAB~T UD Q4-6H PRN (MAXIMUM 12 TB DAILY) RF2 #50 DS30  on 01 Jan 2001^ACTIVE^^^
'"ZOD^000518-00006^9^^ CLARITHROMYCIN (BIAXIN)--PO 500MG TAB~T1 TB BID F7 RF0 #14 DS7  on 01 Jan 2001^ACTIVE^^^
'"ZOD^000518-00008^9^^ ITRACONAZOLE--PO 100MG CAP~T 2 CAP BID F7 PER MONTH FOR 3 MONTHS(FINGERS) OR 4 MONTHS(TOES) RF2 #100 DS7  on 18 May 2000^ACTIVE^^(Order has changed: Expand to View)^
'"ZOD^000518-00017^9^^ ALBUTEROL *BCF*--INH 90MCG/DOSE AERO~INH 2 PUFFS QID PRN RF0 #1 DS30  on 18 May 2000^ACTIVE^^(Order has changed: Expand to View)^
'"ZOD^000518-00034^9^^ PEN G BENZ (PERMAPEN)--IM 1.2M UNITS~PO BID RF1 #60 DS30  on 18 May 2000^ACTIVE^^(Not Dispensed)^
'"ZOD^000518-00037^9^^ FLUOCINONIDE--TOP 0.05% GEL~AAA BID RF5 #60 DS30  on 02 Jan 2001^ACTIVE^^^
'"ZOD^000518-00038^9^^ SARRACENIACEAE (SARAPIN)--INJ SOLN~T1 PO QID RF2 #50 DS30  on 01 Jan 2001^ACTIVE^^^
'"ZOD^000518-00047^9^^ PEN G BENZ (PERMAPEN)--IM 1.2M UNITS~PO BID RF0 #30 DS30  on 18 May 2000^ACTIVE^^(Not Dispensed)^
'"ZOD^000518-00055^5^^*NOTIFY* ELBOWS BILAT~DIAGNOSTIC RADIOLOGY~WHEELCHAIR  on 18 May 2000^ACTIVE^^^
'"ZOD^000523-00001^9^^ TAMOXIFEN CITRATE--PO 10MG TAB~T1 TAB PO BID RF5 #120 DS60  on 23 May 2000^ACTIVE^^(Not Dispensed)^
'"ZOD^000523-00040^9^^ ASPIRIN--RECT 120MG SUPP~1 RF1 #32 DS30  on 23 May 2000^ACTIVE^^(Not Dispensed)^
'"ZOD^000525-00015^9^^ NORETHI/ETH-ESTR. (MODICON EQ)--PO TAB~T1 TAB QD RF3 #84 DS84  on 25 May 2000^ACTIVE^^(Not Dispensed)^
'"ZOD^000526-00050^5^^ CT, HEAD W/CONTRAST~CAT SCAN~AMB  on 27 Apr 2000^PENDING^^^
'"ZOD^000526-00051^5^^ FOOT & TOES LT - GROUP~DIAGNOSTIC RADIOLOGY~AMB  on 26 May 2000^ACTIVE^^^
'"ZOD^000530-00020^4^^*STAT* ALCOHOL,ETHYL BLOOD~SEND PATIENT TO LAB~BLOOD~RED  on 31 May 2000@NOW^ACTIVE^^^
'"ZOD^000530-00022^5^^ ABDOMEN SERIES ACUTE (W/PA CHEST)~DIAGNOSTIC RADIOLOGY~AMB  on 30 May 2000^ACTIVE^^^
'"ZOD^000530-00043^5^^*NOTIFY* HAND LT~DIAGNOSTIC RADIOLOGY~AMB  on 30 May 2000^ACTIVE^^^
'"ZOD^000530-00045^4^^*ASAP*ASAP* LEPTOSPIRA ANTIBODY~WARD/CLINIC COLLECT~BLOOD~RED  on 30 May 2000@NOW^ACTIVE^^^
'"ZOD^000530-00051^4^^*STAT*NOTIFY* GLUCOSE~SEND PATIENT TO LAB~URINE  on 30 May 2000@NOW^ACTIVE^^^
'"ZOD^000531-00004^5^^*ASAP* FOOT & ANKLE RT - GROUP~DIAGNOSTIC RADIOLOGY~AMB  on 01 Jun 2000^ACTIVE^^^
'"ZOD^000531-00005^9^^ ACETAMINOPHEN--PO 325MG TAB~T UD Q4-6H PRN (MAXIMUM 12 TB DAILY) RF2 #50 DS30  on 10 Jun 2000^ACTIVE^^^
'"ZOD^000531-00006^4^^*NOTIFY* SPECIAL STAINS~SEND PATIENT TO LAB~BLOOD~LAV  on 31 May 2000@NOW^ACTIVE^^^
'"ZOD^000531-00008^5^^*STAT* FEMUR LT~DIAGNOSTIC RADIOLOGY~AMB  on 31 May 2000   [CANCELLED] ^ON HOLD^^^
'"ZOD^000531-00009^9^^ PSYLLIUM--PO POWD~MIX 1 TABLESPOONFUL IN 8 OZ WATER OR JUICE AND TAKE QD RF0 #1 DS30  on 31 May 2000^ACTIVE^^(Order has changed: Expand to View)^
'"ZOD^000531-00010^4^^*STAT*STAT* GLUCOSE~SEND PATIENT TO LAB~BLOOD~RED  on 31 May 2000@0920^ON HOLD^^^
'"ZOD^000531-00015^5^^*STAT* ABDOMEN SERIES ACUTE (W/PA CHEST)~DIAGNOSTIC RADIOLOGY~AMB  on 31 May 2000^ACTIVE^^^
'"ZOD^000531-00039^9^^ SARGRAMOSTIM--SQ 500MCG VIAL~T1 PO TID RF0 #1 DS30  on 31 May 2000^ACTIVE^^(Not Dispensed)^
'"ZOD^000531-00040^4^^*ASAP*ASAP* HI-DRY BATH CLEAN~WARD/CLINIC COLLECT~OTHER MISC  on 31 May 2000@NOW^ACTIVE^^^
'"ZOD^000531-00041^5^^*STAT* TIBIA/FIBULA, LT~DIAGNOSTIC RADIOLOGY~WHEELCHAIR  on 31 May 2000^ACTIVE^^^
'"ZOD^000531-00059^4^^*STAT*STAT* GLUCOSE~SEND PATIENT TO LAB~BLOOD~RED  on 31 May 2000@NOW^ACTIVE^^^
'"ZOD^000601-00079^4^^ XXAFB IDENTIFICATION~SEND PATIENT TO LAB~OTHER MISC  on 01 Jun 2000@NOW^ACTIVE^^^
'"ZOD^000601-00082^4^^*NOTIFY* HCG,URINE~SEND PATIENT TO LAB~URINE  on 01 Jun 2000@NOW^ACTIVE^^^
'"ZOD^000602-00016^4^^ AUTOPSY~WARD/CLINIC COLLECT~BODY  on 02 Jun 2000@1057^ON HOLD^^^
'"ZOD^000605-00009^4^^ CBC WITH DIFF~SEND PATIENT TO LAB~BLOOD~LAV  on 05 Jun 2000@NOW^ACTIVE^^^
'"ZOD^000605-00010^4^^ UREA NITROGEN~SEND PATIENT TO LAB~BLOOD~RED  on 05 Jun 2000@NOW^ACTIVE^^^
'"ZOD^000606-00029^9^^ ACETAMINOPHEN--PO 100MG/ML SOLN~GV 1 DROPPERFUL (0.8ML) Q4-6H PRN RF0 #15 DS30  on 06 Jun 2000^ACTIVE^^(Not Dispensed)^
'"ZOD^000606-00030^5^^*ASAP* MRI, CARDIAC~MAGNETIC RESONANCE IMAGING~STRETCH  on 06 Jun 2000^PENDING^^^
'"ZOD^000606-00031^4^^*ASAP* PYRUVATE~WARD/CLINIC COLLECT~BLOOD~ARUP TUBE  on 06 Jun 2000@NOW^ACTIVE^^^"
' ZOD^000605-00037^9^^ TRIP BECLOMETHASONE--NAS 4       TRIL QD F2 RF0 #1 DS2  on 05 Jun 2000^ACTIVE^^^0^^

    i = InStr(gszHL7RetMsg, "ZOD")
    If i Then
      szNewStr = Mid$(gszHL7RetMsg, i, Len(gszHL7RetMsg))
      i = InStr(szNewStr, "ZOD")
    Else
      '** no ZOD info ...*** ERROR ***
    End If
    
     'Moved up for PR296; 11-01-00 AAB
''    'Erase Summary Collection and begin new:
''      Set Summary = New Collection
''      Dim SummaryRecord As SummaryType

    iLoopCount = 1    'collections are 1's based
    
    'If at least one ZOD then do this; iParseZOD is also the throttle

    While i
      Set SummaryRecord = New SummaryType
      
'      i = iParseZOD(i, szNewStr, szOrderCode, szOrderStatus, szOrderNum, szOrderDesc, szModifiable, _
'                    szExpirationDate, szLastDispenseDate, szRemainingRF, szLastFillDate)
      i = iParseZODWithSigning(i, szNewStr, SummaryRecord)
     
'      SummaryRecord.OrderTypeCode = Val(szOrderCode)
'      SummaryRecord.Status = szOrderStatus
'      SummaryRecord.OrderId = szOrderNum
'      'Debug.Print SummaryRecord.OrderId, szOrderNum
'
'      SummaryRecord.FullText = szOrderDesc
'      If Trim$(szModifiable) = "1" Then
'        SummaryRecord.Modifiable = hlIS_NOT_MODIFIABLE
'      Else
'        SummaryRecord.Modifiable = hlIS_MODIFIABLE
'      End If
'
'      SummaryRecord.ExpirationDate = szExpirationDate
'      SummaryRecord.LastDispenseDate = szLastDispenseDate
'
'      If SummaryRecord.OrderTypeCode = hlRX_ORDERS_CODE Then
'        SummaryRecord.LastFillDate = szLastFillDate
'        SummaryRecord.RemainingRF = szRemainingRF
'      End If
'
'        szGetOrderTypeDesc SummaryRecord.OrderTypeCode, szOrderCode
'        SummaryRecord.OrderTypeText = szOrderCode

        'are we looking only for LAB RAD RX, or ALL=0
        If hlOrderType = hlALL_ORDERS_CODE Then    'this is what was asked for
          bMatch = True
        ElseIf hlOrderType = SummaryRecord.OrderTypeCode Then  'this is what we found
          bMatch = True
        Else
          bMatch = False
        End If
       
        '==========================================================
        
        If bMatch Then    'else drop through for next record

        szOrderCompleteDesc = SummaryRecord.FullText
        ParseDisplayText szOrderCompleteDesc, SummaryRecord.BriefText
        '****************************************************************************
           
    '        If SummaryRecord.Status = "PRE-ACTIVE" Then  '... force
    '          szModifiable = "1"  'else keep as-is
    '        End If
    '        If szModifiable = "1" Then
    '          szModifiable = "YES"
    '        Else
    '          szModifiable = "NO"
    '        End If
                    
            'Dim SummaryRecord As New SummaryArrayType
          
            'SummaryRecord.OrderId = szOrderNum
            'SummaryRecord.OrderTypeCode = szOrderCode
            'SummaryRecord.OrderTypeText = szOrderTypeDesc
            'SummaryRecord.Status = szOrderStatus
            SummaryRecord.FullText = szOrderCompleteDesc
            'SummaryRecord.BriefText = szTemp
            
          Summary.Add SummaryRecord
           
          iLoopCount = iLoopCount + 1         'starts with "0"
       
        Else
          Set SummaryRecord = Nothing   'destroy this unused object instance
        End If        ' search matches  'order type'
    
    Wend

End Function
Public Function GetOrdersToSign(bReturnOrdersByProvider As hlSIGNBY, bReturnAll As hlSIGNNUMBER, szPatientIEN As String, szNumberToReturn As String, lCursor As hlORDERSIGNCURSOR, iTotalNumberOfOrders As Integer, bMore As Boolean) As hlERRORS
    
    Dim szNewStr As String
    Dim i As Integer
    Dim iLoopCount As Integer
    Dim szTemp As String

     '=============================================

    If gbHL7LoginStatus = False Then
      GetOrdersToSign = hlERRORS.hlERROR_NOT_LOGGED_IN
      Exit Function
    End If
   
    GetOrdersToSign = hlERRORS.hlERROR_NONE

    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    If lCursor = hlStart Then
        Set colUnsignedOrder = New UnsignedOrders
    End If
    
    FmtGetOrdersToSignRequest bReturnOrdersByProvider, bReturnAll, szPatientIEN, szNumberToReturn, lCursor
    
    glngTimeoutStatus = HL7MsgHandler()  'changed 6-30-98;AAB new design
     
    iTotalNumberOfOrders = 0
    If glngTimeoutStatus = hlCOMMSTATUS.hlCOMM_TIMEOUT Then
        GetOrdersToSign = hlERRORS.hlERROR_TIMED_OUT
        Exit Function
    Else
        If gszHL7RetCode = "AA" Then
             GetOrdersToSign = hlERRORS.hlERROR_NONE
             colUnsignedOrder.AddUsingHL7Segments
             mvarszOrderContinuation = szGetSegSeq(gszHL7RetMsg, 1, "ZOX", hlZOXSeg.OrderContinuation)
             If LenB(mvarszOrderContinuation) = 0 Then
                bMore = False
             Else
                bMore = True
             End If
             If lCursor = hlStart Then
                szTemp = szGetSegSeq(gszHL7RetMsg, 1, "ZOX", hlZOXSeg.TotalNumber)
                If LenB(szTemp) > 0 Then
                    iTotalNumberOfOrders = Val(szGetSegSeq(gszHL7RetMsg, 1, "ZOX", hlZOXSeg.TotalNumber))
                End If
                'If iTotalNumberOfOrders > 0 Then
                '    iTotalNumberOfOrders = iTotalNumberOfOrders - 1
                'End If
             End If
        Else
             FetchSignErrors SignError

             GetOrdersToSign = hlERRORS.hlERROR_GET_ORDERS_TO_SIGN_FAILED
             Exit Function
        End If
    End If
        
    
    
End Function
Public Function SendSignOrders() As hlERRORS
    
    Dim szNewStr As String
    Dim i As Integer
    Dim iLoopCount As Integer
    Dim ErrorCollection As Collection
   ' Dim WarningCollection As Collection


     '=============================================

    If gbHL7LoginStatus = False Then
      SendSignOrders = hlERRORS.hlERROR_NOT_LOGGED_IN
      Exit Function
    End If
   
    SendSignOrders = hlERRORS.hlERROR_NONE

    Set SignError = New Collection
    Set ErrorCollection = Nothing
    Set ErrorCollection = New Collection
    
    FmtOrdersToSignRequest
    
    glngTimeoutStatus = HL7MsgHandler()  'changed 6-30-98;AAB new design
     
    If glngTimeoutStatus = hlCOMMSTATUS.hlCOMM_TIMEOUT Then
        SendSignOrders = hlERRORS.hlERROR_TIMED_OUT
        Exit Function
    Else
        If gszHL7RetCode = "AA" Then
             SendSignOrders = hlERRORS.hlERROR_NONE
             
        Else
             SendSignOrders = hlERRORS.hlERROR_SEND_ORDERS_TO_SIGN_FAILED
           '  Exit Function
        End If
    End If
    
    FetchSignErrors ErrorCollection
    HL7GetSignWarnings
    
    Set SignError = ErrorCollection
    'Set SignWarning = WarningCollection
    
    
End Function

Public Function SendPatientSelectMsg(szPatientIEN As String, _
                                     szPatientName As String) As hlERRORS
    Dim hlTemp As hlERRORS
                                  
    SendPatientSelectMsg = hlERRORS.hlERROR_NONE
    If gbHL7LoginStatus = False Then
      SendPatientSelectMsg = hlERRORS.hlERROR_NOT_LOGGED_IN
      Exit Function
    End If
    
    FormatPatientSelectMsg szPatientIEN, szPatientName
    
     'SEND HL7 command package >>>>>>>>>>>>>>>
          glngTimeoutStatus = HL7MsgHandler()  'changed 6-30-98;AAB new design
          
          If glngTimeoutStatus = hlCOMMSTATUS.hlCOMM_TIMEOUT Then
               SendPatientSelectMsg = hlERRORS.hlERROR_TIMED_OUT
               GoTo PatSelFailed:
          Else
               'added 4-7-98; AAB
               If gszHL7RetCode = "AA" Then
                    SendPatientSelectMsg = hlERRORS.hlERROR_NONE
               Else
                    SendPatientSelectMsg = hlERRORS.hlERROR_PATIENT_SELECT_FAILED
                    GoTo PatSelFailed:
               End If
          End If
    
    goPatient.ParsePID gszHL7RetMsg
    
    '4-6-98; AAB this is a "text quality" copy of the name, NOT PID(5)
    Patient.Name = goPatient.Name
    'Want HL7-ready copy of name
    Patient.HL7Name = goPatient.HL7Name
    Patient.IEN = goPatient.IEN
Debug.Print "get Default Order Page"
    'Get the default orders page from ZPW/Z11 return with ZOP3=3 (may be more than one ZOP returned)
    hlTemp = szGetDefOrderPage(szPatientIEN)
    If hlTemp <> hlERROR_NONE Then SendPatientSelectMsg = hlTemp
    
    Exit Function
    
PatSelFailed:   'added 6-30-98; AAB
    gszDefOrderPage = ""
    Patient.Name = ""
    Patient.HL7Name = ""
    Patient.IEN = ""
    
End Function
Friend Sub FmtOrdersOnPage()

    ' MSH
    ReDim seginfo(1 To hlMSHSeg.NumberOfSegs) As String
      seginfo(hlMSHSeg.MsgType) = "ZPW\Z15"
    gszHL7SndMsg = szAddSeg("MSH", seginfo())

    ' PID
    ReDim seginfo(1 To hlPIDSeg.NumberOfSegs) As String
      seginfo(hlPIDSeg.SetID) = "1"
      seginfo(hlPIDSeg.IEN) = szFormatPIDSegIEN()
      seginfo(hlPIDSeg.Name) = Patient.HL7Name
    gszHL7SndMsg = gszHL7SndMsg & szAddSeg("PID", seginfo())

    ' ZOP
    ReDim seginfo(1 To hlZOPSeg.NumberOfSegs) As String
      seginfo(hlZOPSeg.PageIEN) = gszDefOrderPage
    gszHL7SndMsg = gszHL7SndMsg & szAddSeg("ZOP", seginfo())

End Sub
Friend Sub FmtGetOrdersToSignRequest(bReturnOrdersByProvider As hlSIGNBY, bReturnAll As hlSIGNNUMBER, szPatientIEN As String, szNumberToReturn As String, lCursor As hlORDERSIGNCURSOR)

    ' MSH
    ReDim seginfo(1 To hlMSHSeg.NumberOfSegs) As String
      If bReturnOrdersByProvider = hlRETURNBYPATIENT Then
        seginfo(hlMSHSeg.MsgType) = "ZSV\Z23"
      Else
        seginfo(hlMSHSeg.MsgType) = "ZSV\Z22"
      End If
    gszHL7SndMsg = szAddSeg("MSH", seginfo())

    ' ZOX
    ReDim seginfo(1 To hlZOXSeg.NumberOfSegs) As String
      seginfo(hlZOXSeg.ProviderIEN) = gszProviderIEN
      seginfo(hlZOXSeg.ProviderIEN_Name) = gszProviderIEN & "\\99SID"
      seginfo(hlZOXSeg.ReturnOrdersByProvider) = Trim$(Str$(bReturnOrdersByProvider))
      
      If bReturnOrdersByProvider = hlRETURNBYPATIENT Then
        seginfo(hlZOXSeg.PatientIEN) = szPatientIEN
      End If
      
      seginfo(hlZOXSeg.ReturnAllOrders) = Trim$(Str$(bReturnAll))
      If bReturnAll = hlRETURNBATCH Then
        seginfo(hlZOXSeg.NumberOfOrders) = szNumberToReturn
        
      End If
      Select Case lCursor
        Case hlORDERSIGNCURSOR.hlForward:
        'Case hlORDERSIGNCURSOR.hlBack
        Case hlORDERSIGNCURSOR.hlStart:
                mvarszOrderContinuation = ""
      End Select
      
      seginfo(hlZOXSeg.OrderContinuation) = mvarszOrderContinuation
    gszHL7SndMsg = gszHL7SndMsg & szAddSeg("ZOX", seginfo())


End Sub
Friend Sub FmtOrdersToSignRequest()
    Dim i As Integer
    Dim iH As Integer
    Dim bAddToMsg As Boolean
    
    ' MSH
    ReDim seginfo(1 To hlMSHSeg.NumberOfSegs) As String
      seginfo(hlMSHSeg.MsgType) = "ZSV\Z24"
    gszHL7SndMsg = szAddSeg("MSH", seginfo())

    ' ZOD
    For i = 1 To colUnsignedOrder.Count
        bAddToMsg = False
        'If colUnsignedOrder(i).ProviderSignatureNeeded = hlSigNeeded And LenB(colUnsignedOrder(i).ProviderSignatureInfo.IEN) > 0 Then
        If colUnsignedOrder.bSendingOverrides Then
            '12-4-02 If colUnsignedOrder(i).WarningOverriden = True Then
            If colUnsignedOrder(i).UserHasSigned = True And colUnsignedOrder(i).WarningCollection.Count > 0 Then
                bAddToMsg = True
            End If
        Else
            If colUnsignedOrder(i).UserHasSigned = True Then
               bAddToMsg = True
            End If
        End If
        
        If bAddToMsg Then
           
           Debug.Print colUnsignedOrder(i).OrderNumber & " " & colUnsignedOrder(i).ProviderSignatureInfo.IEN
            ReDim seginfo(1 To hlZODSeg.NumberOfSegs) As String
                seginfo(hlZODSeg.OrderId) = colUnsignedOrder(i).OrderNumber
                seginfo(hlZODSeg.OrderType) = GetStrfromEnum(colUnsignedOrder(i).OrderType)
                seginfo(hlZODSeg.OrderStatus) = colUnsignedOrder(i).szStatus
                seginfo(hlZODSeg.EnteredBy) = colUnsignedOrder(i).szEnteredByIEN

                
                seginfo(hlZODSeg.OrderSigningAction) = Trim$(Str$(colUnsignedOrder(i).SigningAction))
                seginfo(hlZODSeg.OrderEventDate) = colUnsignedOrder(i).szOrderEventDateHL7Format
                If mvarSignatureClass > 1 Then
                    seginfo(hlZODSeg.ProviderSigDate) = colUnsignedOrder(i).ProviderSignatureInfo.DateTimeHL7Format
                    seginfo(hlZODSeg.ProviderSignature) = colUnsignedOrder(i).ProviderSignatureInfo.IEN
                    
                    'as per Tho Nguygen 5/1/02
                    seginfo(hlZODSeg.NurseSigDate) = ""
                    seginfo(hlZODSeg.NurseSignature) = ""
                Else
                    'as per Tho Nguygen 5/1/02
                    seginfo(hlZODSeg.ProviderSigDate) = ""
                    seginfo(hlZODSeg.ProviderSignature) = ""
                    
                    seginfo(hlZODSeg.NurseSigDate) = colUnsignedOrder(i).NurseSignatureInfo.DateTimeHL7Format
                    seginfo(hlZODSeg.NurseSignature) = colUnsignedOrder(i).NurseSignatureInfo.IEN
                End If

                gszHL7SndMsg = gszHL7SndMsg & szAddSeg("ZOD", seginfo())
                With colUnsignedOrder(i).UnsignedAddActions
                    For iH = 1 To .Count
                        If .Item(iH).UserHasSigned Then
                            ReDim seginfo(1 To hlZOHSeg.NumberOfSegs) As String
                            seginfo(hlZOHSeg.OrderId) = colUnsignedOrder(i).OrderNumber
                            If mvarSignatureClass > 1 Then
                                seginfo(hlZOHSeg.ProviderSigDate) = .Item(iH).ProviderSignatureInfo.DateTimeHL7Format
                                seginfo(hlZOHSeg.ProviderSignature) = .Item(iH).ProviderSignatureInfo.IEN
                                seginfo(hlZOHSeg.NurseSigDate) = ""
                                seginfo(hlZOHSeg.NurseSignature) = ""
                            Else
                                seginfo(hlZOHSeg.ProviderSigDate) = ""
                                seginfo(hlZOHSeg.ProviderSignature) = ""
                                seginfo(hlZOHSeg.NurseSigDate) = .Item(iH).NurseSignatureInfo.DateTimeHL7Format
                                seginfo(hlZOHSeg.NurseSignature) = .Item(iH).NurseSignatureInfo.IEN
                            End If
                            seginfo(hlZOHSeg.OrderSigningAction) = Trim$(Str$(.Item(iH).SigningAction))
                            seginfo(hlZOHSeg.OrderEventDate) = .Item(iH).szOrderEventDateHL7Format
            
                            gszHL7SndMsg = gszHL7SndMsg & szAddSeg("ZOH", seginfo())
                        End If
                    Next iH
                End With
            '    If colUnsignedOrder(i).WarningOverriden = True Then
                    AddHl7WarningOverrides colUnsignedOrder(i).WarningCollection, colUnsignedOrder(i).WarningHeader, colUnsignedOrder(i).WarningReason
            '    End If

        End If
        
    Next i

End Sub


Friend Sub FormatPatientSelectMsg(szPatientIEN As String, _
                                  szPatientName As String)
    
    ' MSH
    ReDim seginfo(1 To hlMSHSeg.NumberOfSegs) As String
      seginfo(hlMSHSeg.SendingApp) = "PWS\ORE"
      seginfo(hlMSHSeg.RecvApp) = "CHCS\ORE"
      seginfo(hlMSHSeg.MsgType) = "ZPW\Z11"
    gszHL7SndMsg = szAddSeg("MSH", seginfo())
    
    ' PID
    ReDim seginfo(1 To hlPIDSeg.NumberOfSegs) As String
      seginfo(hlPIDSeg.SetID) = "1"
      seginfo(hlPIDSeg.IEN) = szFormatPIDSegIEN(szPatientIEN)
      seginfo(hlPIDSeg.Name) = szPatientName
    gszHL7SndMsg = gszHL7SndMsg & szAddSeg("PID", seginfo())

End Sub

Public Function SearchPatientCHCS(szSearchCrit As String, _
                                  bMore As Boolean) As hlERRORS
    
    If gbHL7LoginStatus = False Then
      SearchPatientCHCS = hlERRORS.hlERROR_NOT_LOGGED_IN
      Exit Function
    End If
    
    mbDSC = bMore
    FormatPatientQueryMsg (szSearchCrit)
    
     'SEND HL7 command package >>>>>>>>>>>>>>>
          glngTimeoutStatus = HL7MsgHandler()  'changed 6-30-98;AAB new design
          
          If glngTimeoutStatus = hlCOMMSTATUS.hlCOMM_TIMEOUT Then
               SearchPatientCHCS = hlERRORS.hlERROR_TIMED_OUT
          Else
                   bMore = mbDSC
               'added 4-7-98; AAB
                   If gszHL7RetCode = "AA" Then
                        SearchPatientCHCS = hlERRORS.hlERROR_NONE
                   Else
                        SearchPatientCHCS = hlERRORS.hlERROR_SEARCH_PATIENT_FAILED
                   End If
          End If
    
End Function
Friend Sub FormatPatientQueryMsg(szPatQry As String)

   ' MSH
    ReDim seginfo(1 To hlMSHSeg.NumberOfSegs) As String
      seginfo(hlMSHSeg.SendingApp) = "PWS\ORE"
      seginfo(hlMSHSeg.RecvApp) = "CHCS\ORE"
      seginfo(hlMSHSeg.MsgType) = "QRY\A19"
    gszHL7SndMsg = szAddSeg("MSH", seginfo())
    
    ' QRD
    ReDim seginfo(1 To hlQRDSeg.NumberOfSegs) As String
      seginfo(hlQRDSeg.DateTime) = Format$(Date$, "yyyymmdd") & Format$(Time$, "hhmmss")
      seginfo(hlQRDSeg.FormatCode) = "R"
      seginfo(hlQRDSeg.Priority) = "I"
      seginfo(hlQRDSeg.ID) = szGetSeq(gszHL7SndMsg, 9)   ' Using MSH Control Id # "Query ID"
      seginfo(hlQRDSeg.QuantityLimitedRequest) = "30\RD"
      seginfo(hlQRDSeg.WhoSubjectFilter) = szPatQry
      seginfo(hlQRDSeg.WhatSubjectFilter) = "APN"
      seginfo(hlQRDSeg.WhatDataCodeValueQual) = ""
    gszHL7SndMsg = gszHL7SndMsg & szAddSeg("QRD", seginfo())

    If mbDSC = True Then
      ReDim seginfo(1 To hlDSCSeg.NumberOfSegs) As String
        seginfo(hlDSCSeg.ContPointer) = mszDSC
      gszHL7SndMsg = gszHL7SndMsg & szAddSeg("DSC", seginfo())
    End If
    
End Sub


Public Function Login(szAccessCode As String, szVerifyCode As String, bEncrypted As hlLOGIN, _
                        szCHCSIPAddress As String, szCHCSPort As String, szLocalIP As String, _
                        szClinicIEN As String, szMeprsCode As String, szMeprsIen As String, szMTFCode As String) As hlERRORS
 
Dim szEncAccessCode As String
Dim szEncVerifyCode As String
Dim retString As String
'KELLYMCDim szMsgBoxText As String
    
    mszAccessCode = szAccessCode
    mszVerifyCode = szVerifyCode
    
    If LenB(mszAccessCode) = 0 Or LenB(mszVerifyCode) = 0 Then
      Login = hlERRORS.hlERROR_ACCESSVERIFY_CODE
      Exit Function
    End If
    
    If LenB(szCHCSIPAddress) = 0 Then
      Login = hlERRORS.hlERROR_CHCSIPADDRESS_NOT_SET
      Exit Function
    End If
    oHL7Server.SetCHCSAddress (szCHCSIPAddress)
    
    If LenB(szCHCSPort) = 0 Then
      Login = hlERRORS.hlERROR_CHCSPORT_NOT_SET
      Exit Function
    End If
    oHL7Server.SetCHCSPort (szCHCSPort)
    
    If LenB(szLocalIP) = 0 Then
      Login = hlERRORS.hlERROR_CLIENTIPADDRESS_NOT_SET
      Exit Function
    End If
    gszLocalIP = szLocalIP
    
    If bEncrypted = hlLOGIN.hlUNENCRYPTED Then
      szEncAccessCode = szEncrypt(UCase$(mszAccessCode))
      szEncVerifyCode = szEncrypt(mszVerifyCode)
    Else
      szEncAccessCode = mszAccessCode
      szEncVerifyCode = mszVerifyCode
    End If
    retString = Space(2048)

    Screen.MousePointer = vbHourglass
       oHL7Server.SetAccessCode (szEncAccessCode)
       oHL7Server.SetVerifyCode (szEncVerifyCode)
       oHL7Server.Login (retString)
       gszHL7RetCode = szGetSegSeq(retString, 1, "MSA", hlMSASeg.AckCode)
       gszHL7RetMsg = retString
       ShowHL7InboundDebugMessages
    Screen.MousePointer = vbDefault
    
    
    mvarszClinicIen = szClinicIEN
    clinic.IEN = szClinicIEN
    mvarszMeprsCode = szMeprsCode
    
    mvarszMeprsIen = szMeprsIen
    clinic.MEPRS_STRING = "\" & mvarszMeprsIen & "\" & mvarszMeprsCode & "\99MEP"
    mvarszMTFCode = szMTFCode
    clinic.MTF_CODE = szMTFCode

    Select Case gszHL7RetCode

    Case "AA"   ' You are good to go
      Login = hlERRORS.hlERROR_NONE
      gbHL7LoginStatus = True
    Case "AE"   ' user error
      Login = hlERROR_ACCESS_VERFIY_IP_NOT_VALID
      GoTo FAILURE
    Case "AR"   ' HL7 message format error
      Login = hlERROR_LOGIN_MSG_FORMAT_ERROR
      GoTo FAILURE
        
    Case "AC"
      Login = hlERROR_LOGIN_COMM_ERROR
      GoTo FAILURE
           
    Case Else   ' Booo         "Your message ACK response is not known."
      Login = hlERROR_LOGIN_UNKNOWN_ERROR
      GoTo FAILURE
    
    End Select
    
Dim szTemp As String

    szTemp = Space$(200)
    oHL7Server.GetUserIEN (szTemp)
    gszUserIEN = szGetColumn(szTemp, 1, "\")
    
    szTemp = Space$(200)
    oHL7Server.GetProviderIEN (szTemp)
    gszProviderIEN = szGetColumn(szTemp, 1, "\")
    If Left$(gszProviderIEN, 1) = vbNullChar Then
        gszProviderIEN = ""
    End If
    gszOrderingProviderIEN = gszProviderIEN
    If gszProviderIEN = "" Then
        Login = hlERROR_LOGIN_NOT_A_PROVIDER
        GoTo FAILURE
    End If
    ''GetClinicInfo
    
     '==================================
    'Added new 6-25-98; AAB
    'reset primary objects after logon
    Set goPatient = New PatientType
    

    HL7CommStatus = hlCOMMSTATUS.hlCOMM_OKCOMM
    '==================================
    
    
    DoEvents
    Exit Function
    
FAILURE:
    gbHL7LoginStatus = False
    Screen.MousePointer = 0
    gszHL7RetCode = ""
    HL7CommStatus = hlCOMMSTATUS.hlCOMM_NOCOMM
End Function

Private Sub Class_Initialize()

    Set oHL7Server = New HL7ComMsg
      gcHl7End = Chr$(28) & vbCr & vbNullChar
      'gszCHCSVersion = "4.5"
      gbHL7LoginStatus = False
        
    Set goPatient = New PatientType
    Set Specimen = New Collection
    Set Summary = New Collection
    
    Set RxWarning = New Collection
    Set LabWarning = New Collection
    Set RadWarning = New Collection
    
    Set RxError = New Collection
    Set LabError = New Collection
    Set RadError = New Collection
    Set OtherError = New Collection
    
  
    Set RXTablet = New RxType
    Set LabTablet = New LabType
    Set RadTablet = New RadType
    Set OtherTablet = New BaseTablet
        
    '9-2-98; Lab AP tests will use the test date if this flag is TRUE (default)
    'Me.LabAP_UseSystemDate = True
    
    Set goErrorTable = New ErrorCollection
    Set goTransportModeTable = New TransportModeCollection
    Set goPriorityTable = New PriorityCollection

    Set colUnsignedOrder = New UnsignedOrders
    
    mbPreModifying = False      'TRUE if using PreModify message
    
'     '**************  Test new TransportModeCollection AMW
'               Dim lLong As hlRADXPORT
'               Dim oTransportMode As TransportModeType
'
'               'Demo how to get TransportModetext from the TransportModenumber
'                lLong = hlRAD_XPORT_STRETCHER
'                Debug.Print "TransportMode Number " & lLong; ":"
'                If Not goTransportModeTable.Item(lLong) Is Nothing Then
'                    With goTransportModeTable.Item(lLong)
'                         Debug.Print .TypeCode, .Text, .HL7Code & vbCrLf
'                    End With
'                End If
'
'               'Demo "For Each" functionality of TransportModecollection
'               Debug.Print "The list of TransportModes is:"
'               Debug.Print "ENUM", "DESCR.", "HL7CODE"
'               For Each oTransportMode In goTransportModeTable
'                Debug.Print oTransportMode.TypeCode, oTransportMode.Text, oTransportMode.HL7Code
'               Next
'               Debug.Print ""
'               'Demo .searchdescription method of TransportModecollection
'               Dim szSearchString As String
'               szSearchString = "bassinet"
'               Set oTransportMode = goTransportModeTable.SearchDescription(szSearchString)
'               Debug.Print "The following TransportMode description contains the text '" & szSearchString & "':"
'               If Not oTransportMode Is Nothing Then
'                With oTransportMode
'                 Debug.Print .TypeCode, .Text, .HL7Code
'                End With
'               End If
'
'               'Demo .searchdescription method of TransportModecollection
'               szSearchString = "whl"
'               Set oTransportMode = goTransportModeTable.SearchHL7Code(szSearchString)
'               Debug.Print "The following TransportMode contains the hl7code text '" & szSearchString & "':"
'               If Not oTransportMode Is Nothing Then
'                With oTransportMode
'                 Debug.Print .TypeCode, .Text, .HL7Code
'                End With
'               End If
'
'    '************** End test new TransportModecollection AMW
'
'    '**************  Test new priorityCollection AMW
'               Dim oPriority As PriorityType
'               Dim lPriority As hlLABPRIORITY
'               Set oPriority = New PriorityType
'
'               'Demo how to get prioritytext from the prioritynumber
'               lPriority = hlLAB_PRIORITY_NOTIFY
'                Debug.Print "Priority Number " & lPriority; ":"
'                If Not goPriorityTable.Item(lPriority) Is Nothing Then
'                    With goPriorityTable.Item(lPriority)
'                         Debug.Print .TypeCode, .Text, .HL7Code & vbCrLf
'                    End With
'                End If
'               'Demo "For Each" functionality and of Prioritycollection
'               Debug.Print "The list of Prioritys is:"
'               Debug.Print "ENUM", "DESCR.", , "HL7CODE"
'               For Each oPriority In goPriorityTable
'                Debug.Print oPriority.TypeCode, oPriority.Text, , oPriority.HL7Code
'               Next
'               Debug.Print ""
'               'Demo .searchdescription method of Prioritycollection
'               szSearchString = "sta"
'               Set oPriority = goPriorityTable.SearchDescription(szSearchString)
'               Debug.Print "The following Priority matches the text '" & szSearchString & "':"
'               If Not oPriority Is Nothing Then
'                With oPriority
'                 Debug.Print .TypeCode, .Text, .HL7Code
'                End With
'               End If
'
'               'Demo .searchdescription method of Prioritycollection
'               szSearchString = "t"
'               Set oPriority = goPriorityTable.SearchHL7Code(szSearchString)
'               Debug.Print "The following Priority matches the hl7code '" & szSearchString & "':"
'               If Not oPriority Is Nothing Then
'                With oPriority
'                 Debug.Print .TypeCode, .Text, .HL7Code
'                End With
'               End If
'
'    '************** End test new Prioritycollection AMW
'
'        '**************  Test new errorCollection AMW
'               Dim oerror As ErrorType
'               Set oerror = New ErrorType
'
'               'Demo how to get errortext from the errornumber
'                Debug.Print "error Number " & hlERROR_TIMED_OUT & ":"
'                With goErrorTable.Item(hlERROR_TIMED_OUT)
'                     Debug.Print .TypeCode, .Text & vbCrLf
'                End With
'
'               'Demo "For Each" functionality and of errorcollection
'               Debug.Print "The list of errors is:"
'               For Each oerror In goErrorTable
'                Debug.Print oerror.TypeCode, oerror.Text
'               Next
'               Debug.Print ""
'               'Demo .searchdescription method of errorcollection
'               Dim colReturn As Collection
'               'Dim szSearchString As String
'               Set colReturn = New Collection
'               szSearchString = "log"
'               Call goErrorTable.SearchDescription(szSearchString, colReturn)
'               Debug.Print "The following error descriptions contain the text '" & szSearchString & "':"
'               If colReturn.Count > 0 Then
'                For Each oerror In colReturn
'                 Debug.Print oerror.TypeCode, oerror.Text
'                Next
'               End If
'    '************** End test new errorcollection AMW

End Sub

Private Sub Class_Terminate()

'    Set oHL7Server = Nothing
'    Set goPatient = Nothing
'    Set Summary = Nothing
    
    '============================
    'added 4-7-98; AAB
        
    Set oHL7Server = Nothing
        
    Set goPatient = Nothing
    Set Specimen = Nothing
    Set Summary = Nothing
    
    Set RxWarning = Nothing
    Set LabWarning = Nothing
    Set RadWarning = Nothing
    
    Set RxError = Nothing
    Set LabError = Nothing
    Set RadError = Nothing
    Set OtherError = Nothing
    
  
    Set LabTablet.Timing = Nothing
    Set RadTablet.Timing = Nothing
    Set RXTablet.Timing = Nothing 'used by GetOrderDetail
    Set OtherTablet.Timing = Nothing
   
    Set LabTablet.Specimen = Nothing
    
    Set RXTablet = Nothing
    Set LabTablet = Nothing
    Set RadTablet = Nothing
    Set OtherTablet = Nothing
    
    
    Set goErrorTable = Nothing
    Set goTransportModeTable = Nothing
    Set goPriorityTable = Nothing
    
    Set colUnsignedOrder = Nothing

End Sub

Public Function HL7MsgHandler(Optional szSndMessage As String) As hlCOMMSTATUS


    Dim lLength As Long
    
    If Len(szSndMessage) > 0 Then
       gszHL7SndMsg = szSndMessage
    End If
    
    ShowHL7OutboundDebugMessages
    
    
    If Right$(gszHL7SndMsg, 3) <> gcHl7End Then
        gszHL7SndMsg = gcHl7Start + gszHL7SndMsg + gcHl7End
    End If
    '---------------------------------------------------
    Debug.Print "Sending Message with SendHL7Message"
    oHL7Server.SendHL7Message (gszHL7SndMsg)
    
    
    
    lLength = oHL7Server.MessageSize
    gszHL7RetMsg = Space(lLength + 1)
    
    'gszHL7RetMsg = Space(32800)
    Debug.Print "Returned from SendHL7Message"
    '---------------------------------------------------
    oHL7Server.GetHL7Message (gszHL7RetMsg)
    gszHL7RetMsg = Trim$(gszHL7RetMsg)
          'RetMsg = gszHL7RetMsg
    Debug.Print "Trimmed Length: " & Len(gszHL7RetMsg) & " PreTrimed Length: " & oHL7Server.MessageSize
    ShowHL7InboundDebugMessages
    
    gszHL7RetCode = szGetSegSeq(gszHL7RetMsg, 1, "MSA", hlMSASeg.AckCode)  'added 10-29-97; AAB
    gszErrMsg = szGetSegSeq(gszHL7RetMsg, 1, "MSA", hlMSASeg.ErrorCondition)
        
    'MSA^AT^^TIMED OUT ^^^^ is sent on time-out and will NOT parse with szGetSegSeq()
    '4-7-98
    If LenB(gszHL7RetCode) = 0 Or gszHL7RetCode = "AT" Then
        HL7CommStatus = hlCOMMSTATUS.hlCOMM_TIMEOUT
        HL7MsgHandler = hlCOMMSTATUS.hlCOMM_TIMEOUT
        gbHL7LoginStatus = False
    Else
        HL7CommStatus = hlCOMMSTATUS.hlCOMM_OKCOMM
        HL7MsgHandler = hlCOMMSTATUS.hlCOMM_OKCOMM
    End If
    
'    DoEvents

End Function

Private Function LockPage(szPageNumber As String, Optional szPatientIEN As String) As hlERRORS
    On Error GoTo ErrorHandler:

    LockPage = hlERRORS.hlERROR_NONE
    If gbHL7LoginStatus = False Then
      LockPage = hlERRORS.hlERROR_NOT_LOGGED_IN
      Exit Function
    End If
    
    FormatLockPageMsg szPageNumber, szPatientIEN
    mszLockedPageID = szPageNumber  'used to keep track of whether a page unlock is reqd
    
     'SEND HL7 command package >>>>>>>>>>>>>>>
          glngTimeoutStatus = HL7MsgHandler()  'changed 6-30-98;AAB new design
          If glngTimeoutStatus = hlCOMMSTATUS.hlCOMM_TIMEOUT Then
               LockPage = hlERRORS.hlERROR_TIMED_OUT
                    ''Exit Function
          Else
                 'added 4-7-98; AAB
                 If gszHL7RetCode = "AA" Then
                   LockPage = hlERRORS.hlERROR_NONE
                 Else
                   LockPage = hlERRORS.hlERROR_LOCK_PAGE_FAILED
                 End If
          End If
          Exit Function
    
ErrorHandler:
  LockPage = hlERRORS.hlERROR_LOCK_PAGE_FAILED
    
End Function

Friend Sub FormatLockPageMsg(szPageNumber As String, Optional szPatientIEN As String)
    ' MSH
    ReDim seginfo(1 To hlMSHSeg.NumberOfSegs) As String
      seginfo(hlMSHSeg.SendingApp) = "PWS\ORE"
      seginfo(hlMSHSeg.RecvApp) = "CHCS\ORE"
      seginfo(hlMSHSeg.MsgType) = "ZPW\Z12"
    gszHL7SndMsg = szAddSeg("MSH", seginfo())
    
    ' PID
    ReDim seginfo(1 To hlPIDSeg.NumberOfSegs) As String
      seginfo(hlPIDSeg.SetID) = "1"
      seginfo(hlPIDSeg.IEN) = szFormatPIDSegIEN(szPatientIEN)
      seginfo(hlPIDSeg.Name) = Patient.HL7Name
    gszHL7SndMsg = gszHL7SndMsg & szAddSeg("PID", seginfo())
    
   ' ZOP
    ReDim seginfo(1 To hlZOPSeg.NumberOfSegs) As String
    seginfo(hlZOPSeg.PageIEN) = szPageNumber
    gszHL7SndMsg = gszHL7SndMsg & szAddSeg("ZOP", seginfo())

End Sub
Private Function UnlockPage(szPageNumber As String) As hlERRORS

          UnlockPage = hlERRORS.hlERROR_NONE
    
          If gbHL7LoginStatus = False Then
                UnlockPage = hlERRORS.hlERROR_NOT_LOGGED_IN
                Exit Function
          End If
    
          FormatUnlockPageMsg szPageNumber
    
     'SEND HL7 command package >>>>>>>>>>>>>>>
          glngTimeoutStatus = HL7MsgHandler()  'changed 6-30-98;AAB new design
          If glngTimeoutStatus = hlCOMMSTATUS.hlCOMM_TIMEOUT Then
               UnlockPage = hlERRORS.hlERROR_TIMED_OUT
                    ''Exit Function
          Else
               'added 4-7-98; AAB
               If gszHL7RetCode = "AA" Then
                    UnlockPage = hlERRORS.hlERROR_NONE
               Else
                    UnlockPage = hlERRORS.hlERROR_UNLOCK_PAGE_FAILED
               End If
          End If
    
End Function

Friend Sub FormatUnlockPageMsg(szPageNumber As String)
    ' MSH
    ReDim seginfo(1 To hlMSHSeg.NumberOfSegs) As String
      seginfo(hlMSHSeg.SendingApp) = "PWS\ORE"
      seginfo(hlMSHSeg.RecvApp) = "CHCS\ORE"
      seginfo(hlMSHSeg.MsgType) = "ZPW\Z14"
    gszHL7SndMsg = szAddSeg("MSH", seginfo())
    
    ' PID
    ReDim seginfo(1 To hlPIDSeg.NumberOfSegs) As String
      seginfo(hlPIDSeg.SetID) = "1"
      seginfo(hlPIDSeg.IEN) = szFormatPIDSegIEN()
      seginfo(hlPIDSeg.Name) = Patient.HL7Name
    gszHL7SndMsg = gszHL7SndMsg & szAddSeg("PID", seginfo())
    
   ' ZOP
    ReDim seginfo(1 To hlZOPSeg.NumberOfSegs) As String
    seginfo(hlZOPSeg.PageIEN) = szPageNumber
    gszHL7SndMsg = gszHL7SndMsg & szAddSeg("ZOP", seginfo())

End Sub


Public Property Get RetMsg() As String
  RetMsg = gszHL7RetMsg
End Property

Public Property Get SndMsg() As String
  SndMsg = gszHL7SndMsg
End Property

Public Property Get szUserIEN() As String
  szUserIEN = gszUserIEN
End Property
Public Property Get szProviderIEN() As String
  szProviderIEN = gszProviderIEN
End Property
Public Property Get szOrderingProviderIEN() As String
  szOrderingProviderIEN = gszOrderingProviderIEN
End Property
Public Property Let szOrderingProviderIEN(szTemp As String)
    'in most cases the ordering provider and entering provider (providerien) are the same
    'but if a tech is ordering for a doc then there could be two differnent iens
    If LenB(szTemp) = 0 Then
        gszOrderingProviderIEN = gszProviderIEN
    Else
        gszOrderingProviderIEN = szTemp
    End If
    
End Property
Sub HL7GetWarnings(WarningCollection As Collection, _
                   TabletObject As BaseTablet)
'Pulls the warnings out of the hl7return message and places into warningcollection
Dim pszZOW3 As String
Dim pszZOW7 As String
'KELLYMCDim pszwarning As String
Dim piZowLoop As Integer
Dim piZOW2 As Integer
Dim pszZOW2 As String
Dim piZowCount As Integer
Dim pszZowCount As String

     'Clear out warning collection
     Set WarningCollection = New Collection

    
  TabletObject.WarningHeader = szGetSegSeq(gszHL7RetMsg, 1, "ZOS", hlZOSseg.DisplayText)
  
  'there is a bug in 4.6 where dupilcate warnings are being sent AKM 3/16/99
  pszZowCount = Str(iGetSegCount(gszHL7RetMsg, "ZOW"))
  'pszZowCount = szGetSegSeq(gszHL7RetMsg, 1, "ZOS", 1)
  
  'bug fix 4-1-98;AAB
  If LenB(Trim$(pszZowCount)) = 0 Then pszZowCount = "0"
  piZowCount = Val(pszZowCount)  'look for warnings; ZOS(1) has count
  
  If piZowCount <> 0 Then

    Select Case UCase$(TypeName(TabletObject))
    
    Case RXTYPENAME
      Dim pszZow3Left As String
      Dim pszZow3Center As String
      Dim pszZow3Right As String
      Dim pszZow7Left As String
      Dim pszZow7Right As String
      Dim RxWarningObject As RxWarningType
    
      For piZowLoop = 1 To piZowCount
        Set RxWarningObject = New RxWarningType

        'pszZOW3 RX: "ordered///reason///conflicts with"; 4-10-97; CHCS 4.5 also sends "ordered//reason//conflicts with"
        pszZOW3 = szGetSegSeq(gszHL7RetMsg, piZowLoop, "ZOW", hlZOWSeg.WarningDesc)  'warning description
        
        'filled by reference;  LEFT=drug  CENTER=problem  RIGHT=why
        HL7ParseZow3 pszZOW3, pszZow3Left, pszZow3Center, pszZow3Right
          pszZOW7 = szGetSegSeq(gszHL7RetMsg, piZowLoop, "ZOW", hlZOWSeg.OrderDisplayText)  'RWS5type text
          pszZow7Left = szGetColumn(pszZOW7, 1, "~")
          pszZow7Right = szGetColumn(pszZOW7, 2, "~")
        '============================================
        RxWarningObject.SequenceNumber = piZowLoop  'same as if ZOW(1)
        RxWarningObject.WarningTypeCode = szGetSegSeq(gszHL7RetMsg, piZowLoop, "ZOW", hlZOWSeg.WarningID)                                          '2
        RxWarningObject.WarningTypeCodePDTS = GetCodefromPDTSWarning(RxWarningObject.WarningTypeCode)
        RxWarningObject.WarningDescription = pszZOW3  '3-FULL; warning description  '3
        RxWarningObject.ReportId = szGetSegSeq(gszHL7RetMsg, piZowLoop, "ZOW", hlZOWSeg.ReportId)  'report ID          '4
        RxWarningObject.OrderFileId = szGetSegSeq(gszHL7RetMsg, piZowLoop, "ZOW", hlZOWSeg.OrderFileId)  '5 order file ID
        RxWarningObject.AllergyId = szGetSegSeq(gszHL7RetMsg, piZowLoop, "ZOW", hlZOWSeg.AllergyId)  '6 allergy ID
        RxWarningObject.ConflictData = pszZOW7  'RWS5type text
        RxWarningObject.ConflictingOrderText = pszZow3Left
        If RxWarningObject.WarningTypeCodePDTS = hlPDTSWarning_Local Then
            RxWarningObject.ReasonOrProblem = pszZow3Center
        Else
            RxWarningObject.ReasonOrProblem = GetTextfromPDTSWarning(RxWarningObject.WarningTypeCodePDTS) & " " & vbTab & pszZow3Center
        End If
        If RxWarningObject.WarningTypeCodePDTS = hlPDTSWarning_Local Then
            RxWarningObject.ConflictData = pszZOW7  'RWS5type text
        Else
            RxWarningObject.PDTSExpandedWarning = GetExpandedTextfromPDTSWarning(pszZOW7, RxWarningObject.WarningTypeCodePDTS)
        End If
        RxWarningObject.ConflictsWithOrWhy = pszZow3Right
        RxWarningObject.ConflictingOrderName = pszZow7Left
        RxWarningObject.ConflictSigAndDate = pszZow7Right
    
        WarningCollection.Add RxWarningObject

      Next piZowLoop

    '================================================================================
    Case LABTYPENAME  'hlORDERTYPECODE.hlLAB_ORDERS_CODE

      'ZOW(7) is proc+coll method+container/date;
      'eg: ^CHEM 7~SEND PATIENT TO LAB~BLOOD~MARB/RED   on 05 Oct 1995^
      'ZOW(3) always says "POSSIBLE DUPLICATE ORDER"
      'LAB Warnings are for DUPLICATES Only
          
      Dim LabWarningObject As LabWarningType
      Dim Zow71 As String
      Dim Zow72 As String
      Dim Zow73 As String
      Dim Zow74 As String
        
       For piZowLoop = 1 To piZowCount
        Set LabWarningObject = New LabWarningType
        pszZOW7 = szGetSegSeq(gszHL7RetMsg, piZowLoop, "ZOW", hlZOWSeg.OrderDisplayText)
        
        'ZOW^1^2^POSSIBLE DUPLICATE ORDER^^1079474^^
        '
        'ZOW(7) ==
            '*ASAP* STOOL CULTURE~
          'SEND PATIENT TO LAB~
          'BLOOD~
          'LT BLUE  Q24H  Q21D for 8 days Starting on 18 Apr 1996@0900^

        'match CHCS screenmode display
            Zow71 = szGetColumn(pszZOW7, 1, "~")    'test
        'use pszTemp
            Zow72 = szGetColumn(pszZOW7, 2, "~")    'send
        '3 is specimen
            Zow73 = szGetColumn(pszZOW7, 3, "~")  'specimen
        '4 is cap and details
            Zow74 = szGetColumn(pszZOW7, 4, "~")   'container/date

        LabWarningObject.SequenceNumber = piZowLoop                         'same as if ZOW(1)
        LabWarningObject.WarningTypeCode = szGetSegSeq(gszHL7RetMsg, piZowLoop, "ZOW", hlZOWSeg.WarningID)                                           '2
        LabWarningObject.WarningDescription = szGetSegSeq(gszHL7RetMsg, piZowLoop, "ZOW", hlZOWSeg.WarningDesc)  'warning description  '3
        LabWarningObject.ReportId = szGetSegSeq(gszHL7RetMsg, piZowLoop, "ZOW", hlZOWSeg.ReportId)   'report ID          '4
        LabWarningObject.OrderFileId = szGetSegSeq(gszHL7RetMsg, piZowLoop, "ZOW", hlZOWSeg.OrderFileId)  'order file ID
        LabWarningObject.AllergyId = szGetSegSeq(gszHL7RetMsg, piZowLoop, "ZOW", hlZOWSeg.AllergyId)  'allergy ID
        LabWarningObject.ConflictData = pszZOW7     'RWS5type text
        LabWarningObject.ConflictingOrderName = Zow71
        LabWarningObject.LabCollectMethod = Zow72
        LabWarningObject.LabSampleType = Zow73
        LabWarningObject.CapColorAndTiming = Zow74
           
          WarningCollection.Add LabWarningObject

      Next piZowLoop

    '================================================================================
    Case RADTYPENAME 'hlORDERTYPECODE.hlRAD_ORDERS_CODE
    
      Dim RadWarningObject As RadWarningType

      'NOTE:9695; format differs slightly from design:
      'ZOW(3) is SHORT or VERBOSE text
      'ZOW(7) is proc+date+status

      'pszFrag = "Rad Test"
      'pszHeader = gchTab & "PROC" & gchTab & "Date/Time" & gchTab & "Status" & gchCRLF
      'pszHeader = ""
        
      For piZowLoop = 1 To piZowCount
        Set RadWarningObject = New RadWarningType
        
          'HL7GetZow2Text gets title text from type code; ZOW(2) ** TEXT **
        pszZOW7 = szGetSegSeq(gszHL7RetMsg, piZowLoop, "ZOW", hlZOWSeg.OrderDisplayText)  'RWS5type text 'Previous test date/time details
        pszZOW2 = szGetSegSeq(gszHL7RetMsg, piZowLoop, "ZOW", hlZOWSeg.WarningID)  'warning type code
        piZOW2 = Val(pszZOW2)
        
        RadWarningObject.SequenceNumber = piZowLoop                                           'same as if ZOW(1)
        RadWarningObject.WarningTypeCode = pszZOW2                                            '2
        RadWarningObject.WarningDescription = szGetSegSeq(gszHL7RetMsg, piZowLoop, "ZOW", hlZOWSeg.WarningDesc)  'warning description  '3
        RadWarningObject.ReportId = szGetSegSeq(gszHL7RetMsg, piZowLoop, "ZOW", hlZOWSeg.ReportId)            'report ID          '4
        RadWarningObject.OrderFileId = szGetSegSeq(gszHL7RetMsg, piZowLoop, "ZOW", hlZOWSeg.OrderFileId)         'order file ID
        RadWarningObject.AllergyId = szGetSegSeq(gszHL7RetMsg, piZowLoop, "ZOW", hlZOWSeg.AllergyId)           'allergy ID
        RadWarningObject.ConflictingOrderName = pszZOW7                                       'RWS5type text
        RadWarningObject.ConflictData = pszZOW7                                               'RWS5type text
        RadWarningObject.SectionTitle = HL7GetZow2Text(piZOW2)

        WarningCollection.Add RadWarningObject

      Next piZowLoop
    Case Else 'Consult
        Dim ConWarningObject As LabWarningType
        OtherTablet.OrderName = szGetColumn(szGetSegSeq(gszHL7RetMsg, 1, "OBR", hlOBRSeg.TestName), 2, "\")
        OtherTablet.OrderId = szGetColumn(szGetSegSeq(gszHL7RetMsg, 1, "ORC", hlORCSeg.OrderId), 1, "\")
        For piZowLoop = 1 To piZowCount
            Set ConWarningObject = New LabWarningType
            pszZOW7 = szGetSegSeq(gszHL7RetMsg, piZowLoop, "ZOW", hlZOWSeg.OrderDisplayText)
            Zow71 = szGetColumn(pszZOW7, 1, "~")    'test

            ConWarningObject.SequenceNumber = piZowLoop                         'same as if ZOW(1)
            ConWarningObject.WarningTypeCode = szGetSegSeq(gszHL7RetMsg, piZowLoop, "ZOW", hlZOWSeg.WarningID)                                           '2
            ConWarningObject.WarningDescription = szGetSegSeq(gszHL7RetMsg, piZowLoop, "ZOW", hlZOWSeg.WarningDesc)  'warning description  '3
            ConWarningObject.ReportId = szGetSegSeq(gszHL7RetMsg, piZowLoop, "ZOW", hlZOWSeg.ReportId)   'report ID          '4
            ConWarningObject.OrderFileId = szGetSegSeq(gszHL7RetMsg, piZowLoop, "ZOW", hlZOWSeg.OrderFileId)  'order file ID
            ConWarningObject.AllergyId = szGetSegSeq(gszHL7RetMsg, piZowLoop, "ZOW", hlZOWSeg.AllergyId)  'allergy ID
            ConWarningObject.ConflictData = pszZOW7     'RWS5type text
            ConWarningObject.ConflictingOrderName = Zow71
               
            WarningCollection.Add ConWarningObject

      Next piZowLoop
    Set OtherWarning = WarningCollection
    End Select

    '================================================================================
  
  End If

End Sub
Sub HL7GetSignWarnings()
'Pulls the warnings out of the hl7return message and places into warningcollection
Dim pszZOW3 As String
Dim pszZOW7 As String
'KELLYMCDim pszwarning As String
Dim piZowLoop As Integer
Dim piZOW2 As Integer
Dim pszZOW2 As String
Dim piZowCount As Integer
Dim pszZowCount As String
Dim piZODCount As Integer
Dim szMsgPart As String
Dim iStart As Long
Dim iStop As Long
Dim iCount As Long
Dim szOrderID As String
Dim pSign As UnsignedOrder

    'Clear out warning collection
   ' Set WarningCollection = New Collection

    piZODCount = iGetSegCount(gszHL7RetMsg, "ZOD")
    If piZODCount = 0 Then
      Exit Sub
    End If
    iStart = 1
    iStop = 1

    For iCount = 1 To piZODCount
        iStart = InStr(iStop, gszHL7RetMsg, "ZOD")
        iStop = InStr(iStart + 3, gszHL7RetMsg, "ZOD")
        If iStop = 0 Then
            iStop = Len(gszHL7RetMsg)
        End If
        
        szMsgPart = Mid$(gszHL7RetMsg, iStart - 1, iStop - iStart)
        szOrderID = szGetSegSeq(szMsgPart, 1, "ZOD", hlZODSeg.OrderId)
        For Each pSign In colUnsignedOrder
            If pSign.OrderNumber = szOrderID Then
                Exit For
            End If
        Next
        pSign.WarningHeader = szGetSegSeq(szMsgPart, 1, "ZOS", hlZOSseg.DisplayText)
        Set pSign.WarningCollection = New Collection
        'there is a bug in 4.6 where dupilcate warnings are being sent AKM 3/16/99
        pszZowCount = Str(iGetSegCount(szMsgPart, "ZOW"))
        
        If LenB(Trim$(pszZowCount)) = 0 Then pszZowCount = "0"
        piZowCount = Val(pszZowCount)  'look for warnings; ZOS(1) has count
          
          If piZowCount <> 0 Then
        
            Select Case pSign.OrderType
            
            Case hlRX_ORDERS_CODE
              Dim pszZow3Left As String
              Dim pszZow3Center As String
              Dim pszZow3Right As String
              Dim pszZow7Left As String
              Dim pszZow7Right As String
              Dim RxWarningObject As RxWarningType
            
              For piZowLoop = 1 To piZowCount
                Set RxWarningObject = New RxWarningType
        
                'pszZOW3 RX: "ordered///reason///conflicts with"; 4-10-97; CHCS 4.5 also sends "ordered//reason//conflicts with"
                pszZOW3 = szGetSegSeq(szMsgPart, piZowLoop, "ZOW", hlZOWSeg.WarningDesc)  'warning description
                
                'filled by reference;  LEFT=drug  CENTER=problem  RIGHT=why
                HL7ParseZow3 pszZOW3, pszZow3Left, pszZow3Center, pszZow3Right
                  pszZOW7 = szGetSegSeq(szMsgPart, piZowLoop, "ZOW", hlZOWSeg.OrderDisplayText)  'RWS5type text
                  pszZow7Left = szGetColumn(pszZOW7, 1, "~")
                  pszZow7Right = szGetColumn(pszZOW7, 2, "~")
                '============================================
                RxWarningObject.SequenceNumber = piZowLoop  'same as if ZOW(1)
                RxWarningObject.WarningTypeCode = szGetSegSeq(szMsgPart, piZowLoop, "ZOW", hlZOWSeg.WarningID)                                          '2
                RxWarningObject.WarningDescription = pszZOW3  '3-FULL; warning description  '3
                RxWarningObject.ReportId = szGetSegSeq(szMsgPart, piZowLoop, "ZOW", hlZOWSeg.ReportId)  'report ID          '4
                RxWarningObject.OrderFileId = szGetSegSeq(szMsgPart, piZowLoop, "ZOW", hlZOWSeg.OrderFileId)  '5 order file ID
                RxWarningObject.AllergyId = szGetSegSeq(szMsgPart, piZowLoop, "ZOW", hlZOWSeg.AllergyId)  '6 allergy ID
                RxWarningObject.WarningTypeCodePDTS = GetCodefromPDTSWarning(RxWarningObject.WarningTypeCode)
                If RxWarningObject.WarningTypeCodePDTS = hlPDTSWarning_Local Then
                    RxWarningObject.ConflictData = pszZOW7  'RWS5type text
                Else
                    RxWarningObject.PDTSExpandedWarning = GetExpandedTextfromPDTSWarning(pszZOW7, RxWarningObject.WarningTypeCodePDTS)
                End If
                RxWarningObject.ConflictingOrderText = pszZow3Left
                If RxWarningObject.WarningTypeCodePDTS = hlPDTSWarning_Local Then
                    RxWarningObject.ReasonOrProblem = pszZow3Center
                Else
                    RxWarningObject.ReasonOrProblem = GetTextfromPDTSWarning(RxWarningObject.WarningTypeCodePDTS) & " " & vbTab & pszZow3Center
                End If
                RxWarningObject.ConflictsWithOrWhy = pszZow3Right
                RxWarningObject.ConflictingOrderName = pszZow7Left
                RxWarningObject.ConflictSigAndDate = pszZow7Right
            
                pSign.WarningCollection.Add RxWarningObject
        
              Next piZowLoop
        
            '================================================================================
            Case hlLAB_ORDERS_CODE  'hlORDERTYPECODE.hlLAB_ORDERS_CODE
        
              'ZOW(7) is proc+coll method+container/date;
              'eg: ^CHEM 7~SEND PATIENT TO LAB~BLOOD~MARB/RED   on 05 Oct 1995^
              'ZOW(3) always says "POSSIBLE DUPLICATE ORDER"
              'LAB Warnings are for DUPLICATES Only
        
              Dim LabWarningObject As LabWarningType
              Dim Zow71 As String
              Dim Zow72 As String
              Dim Zow73 As String
              Dim Zow74 As String
        
              For piZowLoop = 1 To piZowCount
                Set LabWarningObject = New LabWarningType
                pszZOW7 = szGetSegSeq(szMsgPart, piZowLoop, "ZOW", hlZOWSeg.OrderDisplayText)
        
                'ZOW^1^2^POSSIBLE DUPLICATE ORDER^^1079474^^
                '
                'ZOW(7) ==
                    '*ASAP* STOOL CULTURE~
                  'SEND PATIENT TO LAB~
                  'BLOOD~
                  'LT BLUE  Q24H  Q21D for 8 days Starting on 18 Apr 1996@0900^
        
                'match CHCS screenmode display
                    Zow71 = szGetColumn(pszZOW7, 1, "~")    'test
                'use pszTemp
                    Zow72 = szGetColumn(pszZOW7, 2, "~")    'send
                '3 is specimen
                    Zow73 = szGetColumn(pszZOW7, 3, "~")  'specimen
                '4 is cap and details
                    Zow74 = szGetColumn(pszZOW7, 4, "~")   'container/date
        
                LabWarningObject.SequenceNumber = piZowLoop                         'same as if ZOW(1)
                LabWarningObject.WarningTypeCode = szGetSegSeq(szMsgPart, piZowLoop, "ZOW", hlZOWSeg.WarningID)                                           '2
                LabWarningObject.WarningDescription = szGetSegSeq(szMsgPart, piZowLoop, "ZOW", hlZOWSeg.WarningDesc)  'warning description  '3
                LabWarningObject.ReportId = szGetSegSeq(szMsgPart, piZowLoop, "ZOW", hlZOWSeg.ReportId)   'report ID          '4
                LabWarningObject.OrderFileId = szGetSegSeq(szMsgPart, piZowLoop, "ZOW", hlZOWSeg.OrderFileId)  'order file ID
                LabWarningObject.AllergyId = szGetSegSeq(szMsgPart, piZowLoop, "ZOW", hlZOWSeg.AllergyId)  'allergy ID
                LabWarningObject.ConflictData = pszZOW7     'RWS5type text
                LabWarningObject.ConflictingOrderName = Zow71
                LabWarningObject.LabCollectMethod = Zow72
                LabWarningObject.LabSampleType = Zow73
                LabWarningObject.CapColorAndTiming = Zow74
        
                pSign.WarningCollection.Add LabWarningObject
        
              Next piZowLoop
        
            '================================================================================
            Case hlRAD_ORDERS_CODE 'hlORDERTYPECODE.hlRAD_ORDERS_CODE
        
              Dim RadWarningObject As RadWarningType
        
              'NOTE:9695; format differs slightly from design:
              'ZOW(3) is SHORT or VERBOSE text
              'ZOW(7) is proc+date+status
        
              'pszFrag = "Rad Test"
              'pszHeader = gchTab & "PROC" & gchTab & "Date/Time" & gchTab & "Status" & gchCRLF
              'pszHeader = ""
        
              For piZowLoop = 1 To piZowCount
                Set RadWarningObject = New RadWarningType
        
                  'HL7GetZow2Text gets title text from type code; ZOW(2) ** TEXT **
                pszZOW7 = szGetSegSeq(szMsgPart, piZowLoop, "ZOW", hlZOWSeg.OrderDisplayText)  'RWS5type text 'Previous test date/time details
                pszZOW2 = szGetSegSeq(szMsgPart, piZowLoop, "ZOW", hlZOWSeg.WarningID)  'warning type code
                piZOW2 = Val(pszZOW2)
        
                RadWarningObject.SequenceNumber = piZowLoop                                           'same as if ZOW(1)
                RadWarningObject.WarningTypeCode = pszZOW2                                            '2
                RadWarningObject.WarningDescription = szGetSegSeq(szMsgPart, piZowLoop, "ZOW", hlZOWSeg.WarningDesc)  'warning description  '3
                RadWarningObject.ReportId = szGetSegSeq(szMsgPart, piZowLoop, "ZOW", hlZOWSeg.ReportId)            'report ID          '4
                RadWarningObject.OrderFileId = szGetSegSeq(szMsgPart, piZowLoop, "ZOW", hlZOWSeg.OrderFileId)         'order file ID
                RadWarningObject.AllergyId = szGetSegSeq(szMsgPart, piZowLoop, "ZOW", hlZOWSeg.AllergyId)           'allergy ID
                RadWarningObject.ConflictingOrderName = pszZOW7                                       'RWS5type text
                RadWarningObject.ConflictData = pszZOW7                                               'RWS5type text
                RadWarningObject.SectionTitle = HL7GetZow2Text(piZOW2)
        
                pSign.WarningCollection.Add RadWarningObject
        
              Next piZowLoop
        
            End Select
        
            '================================================================================
      
        End If
    
      Next iCount

End Sub

Public Property Get CommStatus() As hlCOMMSTATUS
  CommStatus = HL7CommStatus
End Property

Public Property Get VersionCode() As String
  'e.g. 1.1.3
  VersionCode = Trim$(Str$(App.Major)) & "." & Trim$(Str$(App.Minor)) & "." & Trim$(Str$(App.Revision)) & "; " & Trim$(App.Comments)
End Property

Private Function szGetDefOrderPage(Optional szPatientIEN As String) As hlERRORS
'This function should be called after a patient select (ZPW/Z11) has been sent and a return received
'If only one ZOP in the return, then that will be the default page, and no other work required
'If more than one ZOP in the return, then the ZOP with a ZOP3=3 (outpatient) will be selected
'and if its ZOP4<>1 (default flag) then that page will be locked (ZPW/Z12)
'     AMW 8/31/99

Dim iZOPCount As Integer
Dim iCount As Integer
Dim szPage() As String
'KELLYMCDim iOutpatientPage As Integer

On Error GoTo ErrorHandler:

iZOPCount = iGetSegCount(gszHL7RetMsg, "ZOP")
  ReDim szPage(1 To iZOPCount, 1 To 4) 'i.e. szpage(1,3) is the first ZOP, segment 3
  For iCount = 1 To iZOPCount
    szPage(iCount, 1) = szGetSegSeq(gszHL7RetMsg, iCount, "ZOP", hlZOPSeg.PageIEN)
    'szpage(icount,2) = szGetSegSeq(gszHL7RetMsg, iCount, "ZOP", hlZOPSeg.PageIEN) 'not needed
    szPage(iCount, 3) = szGetSegSeq(gszHL7RetMsg, iCount, "ZOP", hlZOPSeg.PageStatus)
    szPage(iCount, 4) = szGetSegSeq(gszHL7RetMsg, iCount, "ZOP", hlZOPSeg.DefaultPage)
    
    If szPage(iCount, 3) = "3" Then
      gszDefOrderPage = szPage(iCount, 1)
      If szPage(iCount, 4) <> "1" Then
Debug.Print "--------------"
Debug.Print "Lock Page Start"
        szGetDefOrderPage = LockPage(szPage(iCount, 1), szPatientIEN)
Debug.Print "Lock Page Done"
      Else
        szGetDefOrderPage = hlERROR_NONE
      End If
      Exit For
    End If
  Next
  
  mvarDefOrderPage = gszDefOrderPage    'added 02-22-00; AAB
  
Exit Function
  
ErrorHandler:
szGetDefOrderPage = hlERROR_GET_DEF_ORDER_PAGE_FAILED

End Function

Private Sub GrabOrderInfo(ByRef szOrderName As String, ByRef szOrderIen As String, ByVal lOrderTypeCode As hlORDERTYPECODE)
'Used to keep the ordername and orderien of a tablet before ditching the tablet
  Select Case lOrderTypeCode
    Case hlORDERTYPECODE.hlRX_ORDERS_CODE
      szOrderName = Trim$(RXTablet.OrderName)
      szOrderIen = Trim$(RXTablet.OrderIen)
    Case hlORDERTYPECODE.hlLAB_ORDERS_CODE
      szOrderName = Trim$(LabTablet.OrderName)
      szOrderIen = Trim$(LabTablet.OrderIen)
    Case hlORDERTYPECODE.hlRAD_ORDERS_CODE
      szOrderName = Trim$(RadTablet.OrderName)
      szOrderIen = Trim$(RadTablet.OrderIen)
    Case Else
  End Select
  
End Sub

'Public Property Get LabAP_UseSystemDate() As Boolean
'  LabAP_UseSystemDate = gbLabAP_UseSystemDate
'End Property
'
'Public Property Let LabAP_UseSystemDate(ByVal bData As Boolean)
'  gbLabAP_UseSystemDate = bData
'End Property

Public Property Get BatchTime() As String
  BatchTime = gszBatchTime
End Property

Friend Property Let BatchTime(ByVal szData As String)
  gszBatchTime = szData
End Property

Public Property Get BatchID() As String
  BatchID = gszBatchId
End Property

Friend Property Let BatchID(ByVal szData As String)
  gszBatchId = szData
End Property


Public Function PreModify(OrderId As String, OrderTypeCode As hlORDERTYPECODE) As hlERRORS
     
'This function combines a GetOrderDetail and a preverify messages.
'It is used prior to modifying an existing order.
     
Dim lReturn As hlERRORS
Dim szSigCode As String
Dim szPharmacyIenSent As String
Dim szPharmacyNameSent As String
Dim piCount As Integer
Dim piZQRcount As Integer
     
            On Error GoTo Exit_Premodify
     
        'Used at module level flag to inhibit some normal actions,
        'such as: check for warnings on ZP,
            mbPreModifying = True
          
        'GetOrderDetail first to get orderien and ordername in nnTablet
            lReturn = GetOrderDetail(OrderId, OrderTypeCode)
                      
If lReturn <> hlERROR_NONE Then GoTo Exit_Premodify:
'======================================================

     'Need to clear some of the tablet items before Preverify (or they go out in ZP and change the ACK)
     'These items are put back in at end of Premodify
     
          Select Case OrderTypeCode
        '===================================================================
          Case hlRX_ORDERS_CODE
          
                    With RXTablet
                         szSigCode = .SigCode
                         szPharmacyIenSent = .PharmacyIenSent
                         szPharmacyNameSent = .PharmacyNameSent
                         .SigCode = ""
                         .PharmacyIenSent = ""
                         .PharmacyNameSent = ""
                    End With
               
        '===================================================================
            Case hlLAB_ORDERS_CODE        'save any ZQRs for Modify with RqdComments; 1-23-02 AAB
                'added 1-23-02 for RqdComments Modify; AAB
                    piZQRcount = iGetSegCount(gszHL7RetMsg, "ZQR")
                    If piZQRcount > 0 Then
                        Set COLL_ZQRs = New Collection
                        
                        ''End If
                        Dim STRUCT_ZQRs As Answers
                        For piCount = 1 To piZQRcount
                            Set STRUCT_ZQRs = New Answers
                            STRUCT_ZQRs.AnswerText = szGetSegSeq(gszHL7RetMsg, piCount, "ZQR", 2)
                            STRUCT_ZQRs.AnswerCode = szGetSegSeq(gszHL7RetMsg, piCount, "ZQR", 3)
                            COLL_ZQRs.Add STRUCT_ZQRs
                            Set STRUCT_ZQRs = Nothing
                        Next piCount
                    End If
  
        '===================================================================
            End Select
          
    '==============================================================================================
    '==============================================================================================
     
     'Preverify next to fill in the defaults
     'Note - Preverify wil overwrite the orderien and ordername
     '    and, if radtype, then procedurelocation/ien is also overwritten.
     
        lReturn = PreVerify(OrderTypeCode)
        If lReturn <> hlERROR_NONE Then GoTo Exit_Premodify:
        
        '======================================================================
     
            'Put back in the values that we were keeping from the getorderdetail
                Select Case OrderTypeCode
                Case hlRX_ORDERS_CODE
                    With RXTablet
                        .SigCode = szSigCode
                        .PharmacyIenSent = szPharmacyIenSent
                        .PharmacyNameSent = szPharmacyNameSent
                    End With
            
                ''End Select modified 1-24-02
                '===================================================================
                Case hlLAB_ORDERS_CODE  'AAB; code was doing this even if RX or RAD; added CASE steering
                    
                    'The getorderdetail does not get the CapColorOrContainer info,
                    'only the specimen and topography info.  So we need to search the
                    'specimen collection filled by the preverify for the matching
                    'specimen and topography...when we find a match, voila, that is our cap_color-or-container.
        
                    Dim oTypeSpecimen As TypeSpecimen
                    For Each oTypeSpecimen In Specimen
                        With oTypeSpecimen
                            If LabTablet.Specimen.IEN = .IEN And LabTablet.Specimen.TopographyIen = .TopographyIen Then
                                 LabTablet.Specimen.CapColorOrContainer = .CapColorOrContainer
                                Exit For
                            End If
                        End With
                    Next
                    
                    
                    
    '========================================================================================
            'moved here from PreVerify: 02-01-2002
                '24-JAN-2002; AAB: If we are doing this preverify as part of a MODIFY, then we have ZQRs which
                'hold the original test RqdComments entered .... THESE now become our PRESENT DEFAULTS

    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
            
Dim piRqdCount As Integer
Dim iZQR_Count As Integer
Dim piSpecimenIEN As Integer
Dim piThisSpecimenIen As Integer
Dim piSpecimenIndex As Integer
Dim piZocCount As Integer
Dim piLabZocCount As Integer
Dim piTempA As Integer

                    piSpecimenIndex = 0     '1
                    piSpecimenIEN = Val(LabTablet.Specimen.IEN)
                    piLabZocCount = Specimen.Count
                    
                    For piZocCount = 1 To piLabZocCount
                        piThisSpecimenIen = Val(Specimen(piZocCount).IEN)
                        If piThisSpecimenIen = piSpecimenIEN Then
                            piSpecimenIndex = piZocCount
                            Exit For
                        End If
                    Next piZocCount
                    
                    LabTablet.OrderSetHasReqComments = False
                    If Specimen(piZocCount).colRequiredComments.Count > 0 Then
                        LabTablet.OrderSetHasReqComments = True
                    'Begin SCR 27271
                    ElseIf LabTablet.Subscript = hlAUTOPSY Or LabTablet.Subscript = hlBONEMARROW Or LabTablet.Subscript = hlCYTGYN Or LabTablet.Subscript = hlCYTNONGYN Or LabTablet.Subscript = hlSURGPATH Then
                        LabTablet.OrderSetHasReqComments = True
                    End If
                    'end SCR 27271
                    
                    'utilize new tablet property [ModifySpecimenZocIndex]; 02-27-2002; AAB
                    LabTablet.ModifySpecimenZocIndex = piSpecimenIndex
                
                    If piSpecimenIndex <> 0 Then        'was -1
                        'match ZQR answers with ZRQ questions ...
                        If Not COLL_ZQRs Is Nothing Then
                        
                            Set COLL_RESPONSE = New RequiredCommentsCollection  '03-20-2002; AAB

                            piRqdCount = Specimen.Item(piSpecimenIndex).colRequiredComments.Count
                            For piTempA = 1 To piRqdCount
                                For iZQR_Count = 1 To COLL_ZQRs.Count
                                    With Specimen.Item(piSpecimenIndex).colRequiredComments(piTempA)
                                        If .RequiredCommentID = COLL_ZQRs(iZQR_Count).AnswerText Then
                                             .Default_Value = COLL_ZQRs(iZQR_Count).AnswerCode
                                            COLL_RESPONSE.AddUserResponse .RequiredCommentID, COLL_ZQRs(iZQR_Count).AnswerCode
                                             
                                        End If
                                    End With
                                Next iZQR_Count
                            Next piTempA
                            Set COLL_ZQRs = Nothing
                            
                            'added 03-20-2002; AAB: If modify and user never shows the RC screen
                            Set LabTablet.Specimen.colRequiredComments = COLL_RESPONSE
                            Set COLL_RESPONSE = Nothing
                            
                        End If
                    End If
                    
                End Select
                
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    
                mbPreModifying = False
                Exit Function
     
    '========================================================================================
    
Exit_Premodify:

                PreModify = lReturn
                mbPreModifying = False

End Function
Public Function PreSet(OrderTypeCode As hlORDERTYPECODE) As hlERRORS
     
'this function combines an existing Tablet and a preverify message.  It is used prior to
'modifying an item in a order set.
     
Dim lReturn As hlERRORS
     
    On Error GoTo Exit_PreSet
     
    'Used at module level flag to inhibit some normal actions,
    'such as: check for warnings on ZP,
    mbPreModifying = True
          
    

    'Need to clear some of the tablet items before Preverify (or they go out in ZP and change the ACK)
    'These items are put back in at end of PreSet
    Select Case OrderTypeCode
        Case hlRX_ORDERS_CODE
            Dim szSigCode As String
            Dim szPharmacyIenSent As String
            Dim szPharmacyNameSent As String
            With RXTablet
                szSigCode = .SigCode
                szPharmacyIenSent = .PharmacyIenSent
                szPharmacyNameSent = .PharmacyNameSent
                .SigCode = ""
                .PharmacyIenSent = ""
                .PharmacyNameSent = ""
            End With
          
    End Select
     
     'Preverify next to fill in the defaults
     'Note - Preverify wil overwrite the orderien and ordername
     '    and, if radtype, then procedurelocation/ien is also overwritten.
    lReturn = PreVerify(OrderTypeCode)
          
     
     'Put back in the values that we were keeping from the getorderdetail
    Select Case OrderTypeCode
        Case hlRX_ORDERS_CODE
            With RXTablet
                    .SigCode = szSigCode
                    .PharmacyIenSent = szPharmacyIenSent
                    .PharmacyNameSent = szPharmacyNameSent
            End With
        Case hlLAB_ORDERS_CODE
            'Begin SCR 24599
            Dim piSpecimenIEN As Integer
            Dim piZocCount As Integer
            
            piSpecimenIEN = Val(LabTablet.Specimen.IEN)
            LabTablet.OrderSetHasReqComments = False
            LabTablet.ModifySpecimenZocIndex = 0
            For piZocCount = 1 To Specimen.Count
                If Val(Specimen(piZocCount).IEN) = piSpecimenIEN Then
                    If Specimen(piZocCount).colRequiredComments.Count > 0 Then
                        LabTablet.OrderSetHasReqComments = True
                        LabTablet.ModifySpecimenZocIndex = piZocCount
                        Exit For
                    End If
                End If
            Next piZocCount
            'end SCR 24599
            'Begin 27271
            If LabTablet.Subscript = hlAUTOPSY Or LabTablet.Subscript = hlBONEMARROW Or LabTablet.Subscript = hlCYTGYN Or LabTablet.Subscript = hlCYTNONGYN Or LabTablet.Subscript = hlSURGPATH Then
                LabTablet.OrderSetHasReqComments = True
            End If
            'end SCR 27271
            
    End Select
          
     
Exit_PreSet:

    PreSet = lReturn
    mbPreModifying = False

End Function
Public Function Sav2Q(OrderTypeCode As hlORDERTYPECODE) As hlERRORS
     
    On Error GoTo Exit_Sav2Q
     
    Select Case OrderTypeCode
        Case hlRX_ORDERS_CODE
            RXTablet.OrderingProviderIen = gszOrderingProviderIEN
        Case hlLAB_ORDERS_CODE
            LabTablet.OrderingProviderIen = gszOrderingProviderIEN
        Case hlRAD_ORDERS_CODE
            RadTablet.OrderingProviderIen = gszOrderingProviderIEN
    End Select
          
     
Exit_Sav2Q:

    Sav2Q = hlERROR_NONE
    
End Function


Public Property Get ErrorTable() As ErrorCollection
     Set ErrorTable = goErrorTable
End Property

Public Property Get PriorityTable() As PriorityCollection
     Set PriorityTable = goPriorityTable
End Property

Public Property Get TransportModeTable() As TransportModeCollection
     Set TransportModeTable = goTransportModeTable
End Property

Public Function GetControlID() As String
    GetControlID = ControlId
End Function

Public Function GetOrderTypeTextfromEnum(lngEnumVar As hlORDERTYPECODE) As String

    Dim szOrderDesc As String
    szGetOrderTypeDesc lngEnumVar, szOrderDesc
    
'    Select Case lngEnumVar
'    Case hlRAD_ORDERS_CODE: GetOrderTypeTextfromEnum = "RAD"
'    Case hlRX_ORDERS_CODE: GetOrderTypeTextfromEnum = "RX"
'    Case hlLAB_ORDERS_CODE: GetOrderTypeTextfromEnum = "LAB"
'    Case 2: GetOrderTypeTextfromEnum = "APR"
'    Case 7: GetOrderTypeTextfromEnum = "IVF"
'    Case 8: GetOrderTypeTextfromEnum = "ANC"
'    Case 30: GetOrderTypeTextfromEnum = "CON"
'
'    Case Else
        GetOrderTypeTextfromEnum = szOrderDesc
'    End Select

End Function
'
Public Property Let SignatureClass(ByVal iData As Integer)
  mvarSignatureClass = iData
End Property

Public Property Get SignatureClass() As Integer
   SignatureClass = mvarSignatureClass
End Property

