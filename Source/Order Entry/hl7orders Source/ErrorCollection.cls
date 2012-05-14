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
Attribute VB_Name = "ErrorCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'---------------------------------------------------------------------------------------------
'This class is used to create a 'look-up table' of the errors returned from hl7orders.
'To add errors: Add an error code to the hlERRORS enum, and add the code to
'the class initialization procedure to add the error (enum and description) to the collection.
'---------------------------------------------------------------------------------------------


'ERROR CODES and related constants
'=============================================
Public Enum hlERRORS

    hlERROR_NONE = 0
    hlERROR_UNKNOWN = 1
    
    'Category 100 for server and connectivity
    'NOTE: code n01 to n06 similar for all three ORDER types

    hlERROR_ACCESSVERIFY_CODE = 101
    hlERROR_CHCSIPADDRESS_NOT_SET = 102
    hlERROR_CHCSPORT_NOT_SET = 103
    hlERROR_CLIENTIPADDRESS_NOT_SET = 104
    hlERROR_LOGIN_FAILED = 105
    hlERROR_LOGOUT_FAILED = 106
    hlERROR_ACCESS_VERFIY_IP_NOT_VALID = 107
    hlERROR_LOGIN_MSG_FORMAT_ERROR = 108
    hlERROR_LOGIN_COMM_ERROR = 109
    hlERROR_LOGIN_UNKNOWN_ERROR = 110
    hlERROR_LOGIN_NOT_A_PROVIDER = 111

'Category 200 for PATIENT
'added 4-7-98; AAB
    hlERROR_SEARCH_PATIENT_FAILED = 201
    hlERROR_PATIENT_RELEASE_FAILED = 202
    hlERROR_PATIENT_SELECT_FAILED = 203
    hlERROR_LOCK_PAGE_FAILED = 204
    hlERROR_UNLOCK_PAGE_FAILED = 205
    hlERROR_GET_DEF_ORDER_PAGE_FAILED = 206

'Category 300 for GENERAL errors
    hlERROR_TIMED_OUT = 301                'added 6-30-98; AAB
    hlERROR_REQUEST_FAILED = 302           'added 7-1-98; AAB
    hlERROR_PREVERIFY_FAILED = 303         'HL7 Preverify failed; added 7-1-98; AAB
    hlERROR_NO_PROVIDER_IEN = 304
    hlERROR_PREVERIFY_ORDERNAME_DONT_MATCH = 305 'The order name in HL7 Preverify return message does not match the one sent
          'NOTE: error 305 means user aborted test by NOT accepting new name
          
'Category 400 for LAB orders
    hlERROR_PREVERIFY_LAB_FAILED = 401     'HL7 Preverify failed
    hlERROR_LAB_REASON = 402               'No reason for overrides sent
    hlERROR_LAB_NAME = 403                 'No Order name sent
    hlERROR_LAB_INVALID_COMMENT = 404      'Comment is over 78 chars long
    hlERROR_LAB_DATE = 405                 'No date sent
    hlERROR_SEND_LAB_FAILED = 406          'HL7 SendNew Failed
    hlERROR_LAB_ADMIN = 407                'Continuous needs ADMIN time
    hlERROR_LAB_FREQUENCY = 408            'Continuous needs FREQUENCY data
    hlERROR_LAB_DURATION_RQD = 409         'Continuous needs DURATION data
    hlERROR_LAB_BAD_DURATION_FORMAT = 410  'Duration is 'number or 'D-and number' ... data
    hlERROR_LAB_BAD_DURATION_VALUE = 411   'Duration must be 2 or more days

'Category 500 for RAD orders
    hlERROR_PREVERIFY_RAD_FAILED = 501     'HL7 Preverify failed
    hlERROR_RAD_REASON = 502               'No reason for overrides sent
    hlERROR_RAD_NAME = 503                 'No Order name sent
    hlERROR_RAD_INVALID_COMMENT = 504      'Comment is over 78 chars long
    hlERROR_RAD_DATE = 505                 'No date sent
    hlERROR_SEND_RAD_FAILED = 506          'HL7 SendNew Failed
    hlERROR_RAD_ADMIN = 507                'Continuous needs ADMIN time
    hlERROR_RAD_FREQUENCY = 508            'Continuous needs FREQUENCY data
    hlERROR_RAD_DURATION_RQD = 509         'Continuous needs DURATION data
    hlERROR_RAD_BAD_DURATION_FORMAT = 510  'Duration is 'number or 'D-and number' ... data
    hlERROR_RAD_BAD_DURATION_VALUE = 511   'Duration must be 2 or more days
    hlERROR_RAD_CI = 512                   'No clinical impression sent

'Category 600 for Other ORDER errors
    hlERROR_NO_SIG_TEXT = 601              'Called ExpandSig() with blank sig text
    hlERROR_EXPAND_SIG_FAILED = 602        'Error Getting Expanded Sig
    hlERROR_REMOVE_ORDER_FAILED = 603      'Error Removing Order
    hlERROR_CREATE_BATCH_FAILED = 604      'Cannot Create Batch
    hlERROR_NO_BATCH = 605                 'Cannot Submit Missing Batch
    hlERROR_SUBMIT_BATCH_FAILED = 606      'Error Submitting Batch
    hlERROR_SEND_OVERRIDES_FAILED = 607    'Error Sending Overrides
    hlERROR_ORDER_DETAIL_FAILED = 608
    hlERROR_PERFORM_ACTION_FAILED = 609
    hlERROR_BAD_ACTION_CODE = 610            'Bad action code provided to performaction
    hlERROR_MISSING_ACTION_REASON = 611
    hlERROR_GET_ALLERGIES_FAILED = 612
    hlERROR_PERFORM_ALLERGY_ACTION_FAILED = 613
    hlERROR_GET_ORDERS_FAILED = 614
    hlERROR_WRONG_DATE_FORMAT = 615
    hlERROR_WRONG_TIME_FORMAT = 616
    hlERROR_SUMMARYCOLLECTION_IS_EMPTY = 617
    hlERROR_ORDERID_NOT_FOUND = 618
    hlERROR_BAD_ORDERTYPE_CODE = 619         'Bad ordertypecode provided to preverify
    hlERROR_GET_ORDERS_TO_SIGN_FAILED = 620
    hlERROR_SEND_ORDERS_TO_SIGN_FAILED = 621
    hlERROR_ANP_ENCOUNTER_NOT_LOCKED = 622
 
'Category 700 for common Order Entry CONTROL ERRORS
    hlERROR_NOT_LOGGED_IN = 701

'Category 800 is for OE-SPECIFIC errors;          [was for SQL (Stored Procedure/DataBase) Errors]
'Now for OE-specific errors; 04-2000
     hlERROR_SOFT_ERROR = 801           '04-04-00; AAB: allow user to edit and resend
               ''hlERROR_PREVERIFY_ORDERNAME_DONT_MATCH = 305; 'NOTE: error 305 means user aborted test by NOT accepting new name
     hlERROR_ORDER_NAME_SUBSTITUTED = 802
     hlERROR_USER_CANCELLED = 803
     hlERROR_USER_CANCELLED_NEWSIG = 804
     hlERROR_USER_CANCELLED_NEWQTY = 805
     
'Category 900 for RX orders
    
    hlERROR_PREVERIFY_RX_FAILED = 901      'HL7 Preverify failed
    hlERROR_RX_REASON = 902                'No reason for overrides sent
    hlERROR_RX_NAME = 903                  'No Order name sent
    hlERROR_RX_INVALID_COMMENT = 904       'Comment is over 78 chars long
    hlERROR_RX_DATE = 905                  'No date sent
    hlERROR_SEND_RX_FAILED = 906           'HL7 SendNew Failed
    hlERROR_RX_NOSIG = 907                 'No sig code sent
    hlERROR_RX_COMMENT_RQD = 908           'A comment is required if QUANT sent exceeds MAX-Quantity
    hlERROR_RX_REFILLS = 909               'No refills sent
    HLERROR_RX_ORDERIEN = 910
End Enum

'local variable to hold collection
Private mcolErrors As Collection

Public Property Get Count() As Long
     Count = mcolErrors.Count
End Property

Public Property Get Item(IndexKey As hlERRORS) As ErrorType
     
     Dim szIndexKey As String
     
     On Error GoTo ErrorHandler:
     'If item(key) not in collection then error will occur
     If IndexKey > 0 Then
          'Since collection not exposed, assume client will use key(string) not item # to identify item
          szIndexKey = CStr(IndexKey)
          Set Item = mcolErrors.Item(szIndexKey)
     ElseIf mcolErrors.Count > 0 Then  'If indexkey not a collection item, then default to the first item, UNKNOWN ERROR
          Set Item = mcolErrors.Item(1)
     End If     'if collection has no items, return Nothing
     
     Exit Property
     
ErrorHandler:
     'Return item(1) as default, UNKNOWN ERROR
     If mcolErrors.Count > 0 Then
          Set Item = mcolErrors.Item(1)
     End If    'if collection has no items, return Nothing

     
'++++++++++++++++++++++++++++++++++++++++++++++
'     Dim szIndexKey As String
'
'     On Error GoTo ErrorHandler:
'
'     'Since collection not exposed, assume client will use key to identify item
'     szIndexKey = CStr(IndexKey)
'     Set Item = mcolErrors.Item(szIndexKey)
'     Exit Property
'
'ErrorHandler:
'     Set Item = mcolErrors.Item(CStr(hlERROR_UNKNOWN))
     
End Property

'Enable For...Each support
'ProcedureID must be set to -4
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
     Set NewEnum = mcolErrors.[_NewEnum]
End Property


Private Sub Class_Initialize()

Dim oerror As ErrorType

     ''Set oerror = New ErrorType
     
     Set mcolErrors = New Collection
     
          'Load up the collection
          
               Set oerror = New ErrorType
               oerror.TypeCode = hlERROR_UNKNOWN
               oerror.Text = "Unknown Error"
               mcolErrors.Add oerror, CStr(oerror.TypeCode)
     
               Set oerror = New ErrorType
               oerror.TypeCode = hlERROR_ACCESSVERIFY_CODE
               oerror.Text = "Access or Verify Code is blank"
               mcolErrors.Add oerror, CStr(oerror.TypeCode)
               
               Set oerror = New ErrorType
               oerror.TypeCode = hlERROR_CHCSIPADDRESS_NOT_SET
               oerror.Text = "CHCS I IP ADDRESS has not been specified"
               mcolErrors.Add oerror, CStr(oerror.TypeCode)
               
               Set oerror = New ErrorType
               oerror.TypeCode = hlERROR_CHCSPORT_NOT_SET
               oerror.Text = "CHCS I PORT has not been specified"
               mcolErrors.Add oerror, CStr(oerror.TypeCode)
               
               Set oerror = New ErrorType
               oerror.TypeCode = hlERROR_CLIENTIPADDRESS_NOT_SET
               oerror.Text = "IP Address of Client has not been specified"
               mcolErrors.Add oerror, CStr(oerror.TypeCode)
               
               Set oerror = New ErrorType
               oerror.TypeCode = hlERROR_LOGIN_FAILED
               oerror.Text = "CHCS I LOGIN failed"
               mcolErrors.Add oerror, CStr(oerror.TypeCode)
               
               Set oerror = New ErrorType
               oerror.TypeCode = hlERROR_LOGOUT_FAILED
               oerror.Text = "CHCS I LOGOUT failed"
               mcolErrors.Add oerror, CStr(oerror.TypeCode)
     
               Set oerror = New ErrorType
               oerror.TypeCode = hlERROR_ACCESS_VERFIY_IP_NOT_VALID
               oerror.Text = "Access/Verify code not valid.  Or Client IP address is not registered on CHCS I. Or user has no default divison in CHCS I."
               mcolErrors.Add oerror, CStr(oerror.TypeCode)
               
               Set oerror = New ErrorType
               oerror.TypeCode = hlERROR_LOGIN_MSG_FORMAT_ERROR
               oerror.Text = "The login messeage FORMAT is not valid"
               mcolErrors.Add oerror, CStr(oerror.TypeCode)
     
               Set oerror = New ErrorType
               oerror.TypeCode = hlERROR_LOGIN_COMM_ERROR
               oerror.Text = "Could not establish connection to CHCS I"
               mcolErrors.Add oerror, CStr(oerror.TypeCode)
          
               Set oerror = New ErrorType
               oerror.TypeCode = hlERROR_LOGIN_UNKNOWN_ERROR
               oerror.Text = "Unknown LOGIN error"
               mcolErrors.Add oerror, CStr(oerror.TypeCode)
               
               Set oerror = New ErrorType
               oerror.TypeCode = hlERROR_LOGIN_NOT_A_PROVIDER
               oerror.Text = "User is not a CHCS I Provider"
               mcolErrors.Add oerror, CStr(oerror.TypeCode)
               
               '===============================================================================
               
               Set oerror = New ErrorType
               oerror.TypeCode = hlERROR_TIMED_OUT
               oerror.Text = "Message Response Timed Out"
               mcolErrors.Add oerror, CStr(oerror.TypeCode)
               
               Set oerror = New ErrorType
               oerror.TypeCode = hlERROR_PREVERIFY_FAILED
               ''oerror.Text = "No IEN for lab/rx/rad was specified"
               oerror.Text = " Order internal-entry-number (IEN) not specified"    '03-31-00; AAB
               mcolErrors.Add oerror, CStr(oerror.TypeCode)
          
               Set oerror = New ErrorType
               oerror.TypeCode = hlERROR_NO_PROVIDER_IEN
               oerror.Text = "This user is NOT a provider"
               mcolErrors.Add oerror, CStr(oerror.TypeCode)
               
               Set oerror = New ErrorType
               oerror.TypeCode = hlERROR_PREVERIFY_ORDERNAME_DONT_MATCH
               oerror.Text = "The order NAME returned from CHCS I preverify does NOT match the order name sent"
               mcolErrors.Add oerror, CStr(oerror.TypeCode)
                              
               '===============================================================================

               Set oerror = New ErrorType
               oerror.TypeCode = hlERROR_SEARCH_PATIENT_FAILED
               oerror.Text = "Patient SEARCH failed"
               mcolErrors.Add oerror, CStr(oerror.TypeCode)
          
               Set oerror = New ErrorType
               oerror.TypeCode = hlERROR_PATIENT_RELEASE_FAILED
               oerror.Text = "Patient RELEASE failed"
               mcolErrors.Add oerror, CStr(oerror.TypeCode)
               
               Set oerror = New ErrorType
               oerror.TypeCode = hlERROR_PATIENT_SELECT_FAILED
               oerror.Text = "Patient SELECT failed"
               mcolErrors.Add oerror, CStr(oerror.TypeCode)
                              
               '===============================================================================
          
               'this was missing; 03-31-00; AAB
               Set oerror = New ErrorType
               oerror.TypeCode = hlERROR_PREVERIFY_LAB_FAILED
               oerror.Text = "Preverify LAB order failed"
               mcolErrors.Add oerror, CStr(oerror.TypeCode)
          
               Set oerror = New ErrorType
               oerror.TypeCode = hlERROR_LAB_REASON
               ''oerror.Text = "No reason for overrides sent"
               oerror.Text = "No override reason for Lab warning sent"     '03-31-00; AAB
               mcolErrors.Add oerror, CStr(oerror.TypeCode)
               
               Set oerror = New ErrorType
               oerror.TypeCode = hlERROR_LAB_NAME
               oerror.Text = "No order NAME sent with LAB order"
               mcolErrors.Add oerror, CStr(oerror.TypeCode)
               
               Set oerror = New ErrorType
               oerror.TypeCode = hlERROR_LAB_INVALID_COMMENT
               oerror.Text = "Comment is over 80 characters long"
               mcolErrors.Add oerror, CStr(oerror.TypeCode)
               
               Set oerror = New ErrorType
               oerror.TypeCode = hlERROR_LAB_DATE
               oerror.Text = "No DATE sent with LAB order"
               mcolErrors.Add oerror, CStr(oerror.TypeCode)
          
               Set oerror = New ErrorType
               oerror.TypeCode = hlERROR_SEND_LAB_FAILED
               ''oerror.Text = "HL SendNew Failed"
               oerror.Text = "Submit LAB order failed"       '03-31-00; AAB
               mcolErrors.Add oerror, CStr(oerror.TypeCode)
               
               Set oerror = New ErrorType
               oerror.TypeCode = hlERROR_LAB_ADMIN
               oerror.Text = "Continuous Lab orders require an ADMIN time"
               mcolErrors.Add oerror, CStr(oerror.TypeCode)
               
               Set oerror = New ErrorType
               oerror.TypeCode = hlERROR_LAB_FREQUENCY
               oerror.Text = "Continuous Lab orders require FREQUENCY data"
               mcolErrors.Add oerror, CStr(oerror.TypeCode)
          
               Set oerror = New ErrorType
               oerror.TypeCode = hlERROR_LAB_DURATION_RQD
               oerror.Text = "Continuous Lab orders require DURATION data"
               mcolErrors.Add oerror, CStr(oerror.TypeCode)
               
               Set oerror = New ErrorType
               oerror.TypeCode = hlERROR_LAB_BAD_DURATION_FORMAT
               oerror.Text = "Format for Lab Duration is either a number or 'D' and a number"
               mcolErrors.Add oerror, CStr(oerror.TypeCode)
               
               Set oerror = New ErrorType
               oerror.TypeCode = hlERROR_LAB_BAD_DURATION_VALUE
               oerror.Text = "Lab Duration must be 2 or more days"
               mcolErrors.Add oerror, CStr(oerror.TypeCode)
                                        
               '===============================================================================

               Set oerror = New ErrorType
               oerror.TypeCode = hlERROR_PREVERIFY_RAD_FAILED
               ''oerror.Text = "HL Preverify failed"
               oerror.Text = "Preverify RAD order failed"    '03-31-00; AAB
               mcolErrors.Add oerror, CStr(oerror.TypeCode)
               
               Set oerror = New ErrorType
               oerror.TypeCode = hlERROR_RAD_REASON
               ''oerror.Text = "No reason for overrides sent"
               oerror.Text = "No override reason for Rad warning sent"     '03-31-00; AAB
               mcolErrors.Add oerror, CStr(oerror.TypeCode)
               
               Set oerror = New ErrorType
               oerror.TypeCode = hlERROR_RAD_NAME
               oerror.Text = "No order NAME sent with rad order"
               mcolErrors.Add oerror, CStr(oerror.TypeCode)
               
               Set oerror = New ErrorType
               oerror.TypeCode = hlERROR_RAD_INVALID_COMMENT
               oerror.Text = "Rad Comment is over 80 chars long"
               mcolErrors.Add oerror, CStr(oerror.TypeCode)
               
               Set oerror = New ErrorType
               oerror.TypeCode = hlERROR_RAD_DATE
               oerror.Text = "No date sent with Rad order"
               mcolErrors.Add oerror, CStr(oerror.TypeCode)
               
               Set oerror = New ErrorType
               oerror.TypeCode = hlERROR_SEND_RAD_FAILED
               ''oerror.Text = "HL SendNew Failed"
               oerror.Text = "Submit RAD order failed"       '03-31-00; AAB
              mcolErrors.Add oerror, CStr(oerror.TypeCode)
          
               Set oerror = New ErrorType
               oerror.TypeCode = hlERROR_RAD_ADMIN
               oerror.Text = "Continuous Rad orders require ADMIN time"
               mcolErrors.Add oerror, CStr(oerror.TypeCode)
               
               Set oerror = New ErrorType
               oerror.TypeCode = hlERROR_RAD_FREQUENCY
               oerror.Text = "Continuous Rad orders require FREQUENCY data"
               mcolErrors.Add oerror, CStr(oerror.TypeCode)
               
               Set oerror = New ErrorType
               oerror.TypeCode = hlERROR_RAD_DURATION_RQD
               oerror.Text = "Continuous Rad orders require DURATION data"
               mcolErrors.Add oerror, CStr(oerror.TypeCode)
          
               Set oerror = New ErrorType
               oerror.TypeCode = hlERROR_RAD_BAD_DURATION_FORMAT
               oerror.Text = "Format for Rad Duration is 'number or 'D-and number' ... data"
               mcolErrors.Add oerror, CStr(oerror.TypeCode)
               
               Set oerror = New ErrorType
               oerror.TypeCode = hlERROR_RAD_BAD_DURATION_VALUE
               oerror.Text = "Rad Duration must be 1 or more days"
               mcolErrors.Add oerror, CStr(oerror.TypeCode)
               
               Set oerror = New ErrorType
               oerror.TypeCode = hlERROR_RAD_CI
               oerror.Text = "A clinical impression is required"
               mcolErrors.Add oerror, CStr(oerror.TypeCode)
                                        
               '===============================================================================
            
               Set oerror = New ErrorType
               oerror.TypeCode = hlERROR_NO_SIG_TEXT
               ''oerror.Text = "Called ExpandSig() with blank sig text"
               oerror.Text = "Cannot expand blank SIG text"
              mcolErrors.Add oerror, CStr(oerror.TypeCode)
               
               Set oerror = New ErrorType
               oerror.TypeCode = hlERROR_EXPAND_SIG_FAILED
               oerror.Text = "Error getting Expanded Sig text"
               mcolErrors.Add oerror, CStr(oerror.TypeCode)
               
               Set oerror = New ErrorType
               oerror.TypeCode = hlERROR_REMOVE_ORDER_FAILED
               oerror.Text = "Error REMOVING order"
               mcolErrors.Add oerror, CStr(oerror.TypeCode)
               
               Set oerror = New ErrorType
               oerror.TypeCode = hlERROR_CREATE_BATCH_FAILED
               oerror.Text = "Cannot CREATE Batch"
               mcolErrors.Add oerror, CStr(oerror.TypeCode)
          
               Set oerror = New ErrorType
               oerror.TypeCode = hlERROR_NO_BATCH
               oerror.Text = "Cannot Submit a missing Batch"
               mcolErrors.Add oerror, CStr(oerror.TypeCode)
               
               Set oerror = New ErrorType
               oerror.TypeCode = hlERROR_SUBMIT_BATCH_FAILED
               oerror.Text = "Error Submitting Batch"
               mcolErrors.Add oerror, CStr(oerror.TypeCode)
               
               Set oerror = New ErrorType
               oerror.TypeCode = hlERROR_SEND_OVERRIDES_FAILED
               oerror.Text = "Error sending Warning Overrides"
               mcolErrors.Add oerror, CStr(oerror.TypeCode)
               
               Set oerror = New ErrorType
               oerror.TypeCode = hlERROR_ORDER_DETAIL_FAILED
               oerror.Text = "Error sending order-detail message"
               mcolErrors.Add oerror, CStr(oerror.TypeCode)
               
               Set oerror = New ErrorType
               oerror.TypeCode = hlERROR_PERFORM_ACTION_FAILED
               oerror.Text = "Error Sending PerformAction message"
               mcolErrors.Add oerror, CStr(oerror.TypeCode)
               
               Set oerror = New ErrorType
               oerror.TypeCode = hlERROR_BAD_ACTION_CODE
               oerror.Text = "Bad Code sent for PerformAction"
               mcolErrors.Add oerror, CStr(oerror.TypeCode)
               
               Set oerror = New ErrorType
               oerror.TypeCode = hlERROR_MISSING_ACTION_REASON
               oerror.Text = "A Reason for this Action is REQUIRED"
               mcolErrors.Add oerror, CStr(oerror.TypeCode)
               
               Set oerror = New ErrorType
               oerror.TypeCode = hlERROR_GET_ALLERGIES_FAILED
               oerror.Text = "Get Allergies failed"
               mcolErrors.Add oerror, CStr(oerror.TypeCode)
               
               Set oerror = New ErrorType
               oerror.TypeCode = hlERROR_PERFORM_ALLERGY_ACTION_FAILED
               oerror.Text = "Perform Allergy Action failed"
               mcolErrors.Add oerror, CStr(oerror.TypeCode)
               
               Set oerror = New ErrorType
               oerror.TypeCode = hlERROR_GET_ORDERS_FAILED
               oerror.Text = "Get-orders-on-page failed"
               mcolErrors.Add oerror, CStr(oerror.TypeCode)
               
               Set oerror = New ErrorType
               oerror.TypeCode = hlERROR_GET_ORDERS_TO_SIGN_FAILED
               oerror.Text = "Get-orders-to-sign failed"
               mcolErrors.Add oerror, CStr(oerror.TypeCode)
               
               Set oerror = New ErrorType
               oerror.TypeCode = hlERROR_SEND_ORDERS_TO_SIGN_FAILED
               oerror.Text = "Send-orders-to-sign failed"
               mcolErrors.Add oerror, CStr(oerror.TypeCode)
               
               Set oerror = New ErrorType
               oerror.TypeCode = hlERROR_ANP_ENCOUNTER_NOT_LOCKED
               oerror.Text = "The CHCS II Encounter is not locked"
               mcolErrors.Add oerror, CStr(oerror.TypeCode)
               
              
               
               Set oerror = New ErrorType
               oerror.TypeCode = hlERROR_WRONG_DATE_FORMAT
               oerror.Text = "Date format must be YYYYMMDD"
               mcolErrors.Add oerror, CStr(oerror.TypeCode)
               
               Set oerror = New ErrorType
               oerror.TypeCode = hlERROR_WRONG_TIME_FORMAT
               oerror.Text = "Time format must be HHMMSS"
               mcolErrors.Add oerror, CStr(oerror.TypeCode)
               
               Set oerror = New ErrorType
               oerror.TypeCode = hlERROR_SUMMARYCOLLECTION_IS_EMPTY
               ''oerror.Text = "The hl7constructor's Summary collection is empty."
               oerror.Text = "The HL7Orders summary collection is EMPTY."
               mcolErrors.Add oerror, CStr(oerror.TypeCode)
                    
               Set oerror = New ErrorType
               oerror.TypeCode = hlERROR_ORDERID_NOT_FOUND
               oerror.Text = "The order-ID specified does not correspond to any members of the summary collection."
               mcolErrors.Add oerror, CStr(oerror.TypeCode)
               
               Set oerror = New ErrorType
               oerror.TypeCode = hlERROR_BAD_ORDERTYPE_CODE
               oerror.Text = "The orderType-Code supplied is not valid."
               mcolErrors.Add oerror, CStr(oerror.TypeCode)
                                                       
               '===============================================================================

               Set oerror = New ErrorType
               oerror.TypeCode = hlERROR_NOT_LOGGED_IN
               oerror.Text = "User is NOT logged in"
               mcolErrors.Add oerror, CStr(oerror.TypeCode)
                                                       
               '===============================================================================

               Set oerror = New ErrorType
               oerror.TypeCode = hlERROR_PREVERIFY_RX_FAILED
               oerror.Text = "Preverify RX order failed"
               mcolErrors.Add oerror, CStr(oerror.TypeCode)

               Set oerror = New ErrorType
               oerror.TypeCode = hlERROR_RX_REASON
               ''oerror.Text = "No reason for overrides sent"
               oerror.Text = "No override reason for RX warning sent"     '03-31-00; AAB
               mcolErrors.Add oerror, CStr(oerror.TypeCode)
               
               Set oerror = New ErrorType
               oerror.TypeCode = hlERROR_RX_NAME
               oerror.Text = "No order NAME sent with RX order"
               mcolErrors.Add oerror, CStr(oerror.TypeCode)
               
               Set oerror = New ErrorType
               oerror.TypeCode = hlERROR_RX_INVALID_COMMENT
               oerror.Text = "RX Comment is over 80 chars long"
               mcolErrors.Add oerror, CStr(oerror.TypeCode)
               
               Set oerror = New ErrorType
               oerror.TypeCode = hlERROR_RX_DATE
               oerror.Text = "No date sent with RX order"
               mcolErrors.Add oerror, CStr(oerror.TypeCode)
               
               Set oerror = New ErrorType
               oerror.TypeCode = hlERROR_SEND_RX_FAILED
               ''oerror.Text = "HL SendNew Failed for RX"
               oerror.Text = "Submit RX order failed"       '03-31-00; AAB
               mcolErrors.Add oerror, CStr(oerror.TypeCode)
          
               Set oerror = New ErrorType
               oerror.TypeCode = hlERROR_RX_NOSIG
               oerror.Text = "No RX SIG code sent"
               mcolErrors.Add oerror, CStr(oerror.TypeCode)
               
               Set oerror = New ErrorType
               oerror.TypeCode = hlERROR_RX_COMMENT_RQD
               ''oerror.Text = "A comment is required if QUANT sent exceeds MAX-Quantity"
               '03-31-00;AAB
               oerror.Text = "A Comment is required if the ordered quantity exceeds the MAX-quantity value"
               mcolErrors.Add oerror, CStr(oerror.TypeCode)
               
               Set oerror = New ErrorType
               oerror.TypeCode = hlERROR_RX_REFILLS
               oerror.Text = "No Refills sent with RX order"
               mcolErrors.Add oerror, CStr(oerror.TypeCode)
               
               Set oerror = New ErrorType
               oerror.TypeCode = HLERROR_RX_ORDERIEN
               oerror.Text = "No order internal-entry-number (IEN) sent with RX order"
               mcolErrors.Add oerror, CStr(oerror.TypeCode)
          
               '===============================================================================

               Set oerror = New ErrorType
               oerror.TypeCode = hlERROR_SOFT_ERROR         'NOTE: more a message than an error
               oerror.Text = "A Correctable MISSING (or incorrect) parameter was encountered"
               mcolErrors.Add oerror, CStr(oerror.TypeCode)
          
          
               Set oerror = New ErrorType
               oerror.TypeCode = hlERROR_ORDER_NAME_SUBSTITUTED
               oerror.Text = "Order name was substituted and accepted"          'NOTE: more a message than an error
               mcolErrors.Add oerror, CStr(oerror.TypeCode)


               Set oerror = New ErrorType
               oerror.TypeCode = hlERROR_USER_CANCELLED               '803 added 05-30-00; AAB
               oerror.Text = "User Cancelled Activity"          'NOTE: more a message than an error
               mcolErrors.Add oerror, CStr(oerror.TypeCode)


               '===============================================================================
          
          Debug.Print "ErrorCollection Initialize"

End Sub

Public Function SearchDescription(szSearchString As String, ByRef colReturn As Collection)
Attribute SearchDescription.VB_Description = "Returns a Collection of Errortype objects with descriptions that contain the szSearchString supplied. "

Dim oerror As ErrorType

     Set oerror = New ErrorType

For Each oerror In mcolErrors
     If InStr(1, oerror.Text, szSearchString, vbTextCompare) > 0 Then
          colReturn.Add oerror, CStr(oerror.TypeCode)
     End If
Next

End Function

Private Sub Class_Terminate()
     Set mcolErrors = Nothing
     Debug.Print "ErrorCollection Terminate"
End Sub

