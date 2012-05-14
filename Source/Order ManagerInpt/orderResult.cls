VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "orderResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mOrderId As String
Private mComment As String
Private mdOrderDate As Date
Private mdCompletedDate As Date
Private mTestName As String
Private mOrderType As orderTypes
Private mOrderSatus As String
Private mEncounterNumber As String
Private mPatientUnitNumber As String
Private mPatientName As String
Private mPatientSSN As String
Private mPatientFMP As String
Private mSponsorSSN As String
Private mProviderName As String
Private mSEX As String
Private miNormal As enumItemCritical
Private miPriority As enumPriority
Private mResults As String
Private mPatientStatus As String
Private mFrequency  As String
Private mSign       As Boolean
Private mNurseAck   As Boolean
Private mOtherType  As String
Private mStartDate  As Date
Private mDataId     As String
Private mDetails    As String
Private mLastFilledDate As Date

Public Property Get OrderID() As String
    OrderID = mOrderId
End Property

Public Property Get DataId() As String
    DataId = mDataId
End Property

Public Property Get Priority() As enumPriority
    Priority = miPriority
End Property

Public Property Get Normal() As enumItemCritical
    Normal = miNormal
End Property

Public Property Get OrderDate() As Date
    OrderDate = mdOrderDate
End Property


Public Property Get CompletedDate() As Date
    CompletedDate = mdCompletedDate
End Property


Public Property Get TestName() As String
    TestName = mTestName
End Property

Public Property Let TestName(ByVal vsTestName As String)
    mTestName = vsTestName
End Property


Public Property Get ResultType() As orderTypes
    ResultType = mOrderType
End Property

Public Property Get OrderSatus() As String
    OrderSatus = mOrderSatus
End Property

Public Property Get encounterNumber() As String
    encounterNumber = mEncounterNumber
End Property

Public Property Get PatientUnitNumber() As String
    PatientUnitNumber = mPatientUnitNumber
End Property

Public Property Get PatientName() As String
    PatientName = mPatientName
End Property

Public Property Get PatientSSN() As String
    PatientSSN = mPatientSSN
End Property

Public Property Get results() As String
    results = mResults
End Property

Public Property Get ProviderName() As String
    ProviderName = mProviderName
End Property

Public Property Get SEX() As String
    SEX = mSEX
End Property
Public Property Get PatientStatus() As patientType
    '** 6/12/05 - Frank; Fix type mismatched error
'    PatientStatus = mPatientStatus
    PatientStatus = Val(mPatientStatus)
End Property
Public Property Get Frequency() As String
    Frequency = mFrequency
End Property
Public Property Get Sign() As Boolean
    Sign = mSign
End Property
Public Property Get OtherType() As String
    OtherType = mOtherType
End Property
Public Property Get Comment() As String
    Comment = mComment
End Property

Public Property Get NurseAck() As Boolean
    NurseAck = mNurseAck
End Property

Public Property Get StartDate() As Date
    StartDate = mStartDate
End Property

Public Property Get Details() As String
    Details = mDetails
End Property
Public Property Get LastFilledDate() As Date
    LastFilledDate = mLastFilledDate
End Property

Friend Function Deserialize(ByVal rs As Recordset, ByVal order_type As orderTypes) As Boolean
    On Error GoTo ErrHandler

    If rs Is Nothing Then
        Exit Function
    End If

    mOrderId = rs!order_id
    mDataId = rs!data_id
    mdOrderDate = rs!TIME_ORDERED
    mTestName = rs!test_name
    mOrderType = order_type
    mOrderSatus = rs!status
    mEncounterNumber = rs!ENC_NUM
    mPatientUnitNumber = rs!Unit_Number
    mPatientName = rs!Patient_Name
    mProviderName = rs!ordered_by & vbNullString
    miNormal = Val(rs!NORMAL_STATUS & vbNullString)
    If IsDate(rs!COMPLETED_DATE) Then
        mdCompletedDate = (rs!COMPLETED_DATE & vbNullString)
    End If
    If IsDate(rs!TIMING_START_DATE) Then
        mStartDate = rs!TIMING_START_DATE
    End If
    mSEX = rs!SEX
    mComment = rs!prv_comment & vbNullString
    mResults = rs!results & vbNullString
    miPriority = rs!Priority
    mPatientFMP = rs!patient_fmp
    mSponsorSSN = rs!sponsor_ssn
    mPatientSSN = mPatientFMP & "/" & mSponsorSSN
    mFrequency = rs!TIMING_FREQUENCY & vbNullString
    
    If LenB(mFrequency) > 0 Then
        mDetails = "Frequency: " & mFrequency & ", Duration:" & rs!TIMING_DURATION
    End If
        
    If Val(rs!PATIENT_TYPE & vbNullString) = 1 Then
        mPatientStatus = Inpatient
    Else
        mPatientStatus = Outpatient
    End If

    '82728 -- 12/22/05 - Frank
    mSign = Abs(Val(rs!SIGN_FLAG & vbNullString)) = 1
 
    '83102 -- 12/22/05 - Frank
    mNurseAck = Abs(Val(rs!NURSE_ACK_FALG & vbNullString)) = 1

    Deserialize = True
    
    Exit Function
    
ErrHandler:
  goBB.SharedObj.ShowVBError Err.Number, Err.Description, "OrderMgt.Deserialize", "OrderMgt", vbExclamation
End Function

Friend Function DeserializeMed(ByVal rs As Recordset) As Boolean
    On Error GoTo ErrHandler

    If rs Is Nothing Then
        Exit Function
    End If

    '** 6/12/05 - Frank; Fix invalid use of null error
    '** appending & vbNullString

    mOrderId = rs!order_id & vbNullString
    mdOrderDate = rs!TIME_ORDERED & vbNullString

    mTestName = rs!test_name & vbNullString
    mOrderType = hlRX_ORDERS_CODE

    If rs!STATUS_NCID & vbNullString = NCID_ORDERSTATUS_DISCONTINUED Then
        mOrderSatus = isDiscontinued
    ElseIf rs!STATUS_NCID & vbNullString = NCID_ORDERSTATUS_ACTIVE Then
        mOrderSatus = isActive
    ElseIf rs!STATUS_NCID & vbNullString = NCID_ONHOLD Then
        mOrderSatus = isHold
    ElseIf rs!STATUS_NCID & vbNullString = isCompleted_Order Then
        mOrderSatus = isCompleted_Order
        If IsDate(rs!COMPLETED_DATE & vbNullString) Then
            mdCompletedDate = rs!COMPLETED_DATE
        End If
    Else
        mOrderSatus = isPending
    End If
    
    'SF SCR 77233, 77234, 77235, 77236
    If IsDate(rs!ORDER_DATE) Then
        mStartDate = rs!ORDER_DATE
    End If
    
    mEncounterNumber = rs!ENC_NUM & vbNullString
    mPatientUnitNumber = rs!Unit_Number & vbNullString
    mPatientName = rs!Patient_Name & vbNullString
    mProviderName = rs!ordered_by & vbNullString
    mSEX = rs!SEX & vbNullString
    mComment = (rs!prv_comment & vbNullString)
    mPatientFMP = rs!patient_fmp & vbNullString
    mSponsorSSN = rs!sponsor_ssn & vbNullString
    mPatientSSN = mPatientFMP & "/" & mSponsorSSN
    If Val(rs!PATIENT_TYPE & vbNullString) = 1 Then
        mPatientStatus = Inpatient
    Else
        mPatientStatus = Outpatient
    End If
    
    '82728 -- 12/22/05 - Frank
    mSign = Abs(Val(rs!SIGN_FLAG & vbNullString)) = 1
    
    '83102 -- 12/22/05 - Frank
    mNurseAck = Abs(Val(rs!NURSE_ACK_FALG & vbNullString)) = 1
    
    'SCR 105366 - LCI - 9/7/2006
    '"||0" is an Expanded SIG flag stored in the DB
    'so it needs to be removed when being displayed
    'SCR 105874 - LCI - 09/13/2006: Added if/else code to first find the Expanded SIG flag
    'then remove it if found. Otherwise last three characters will be removed no matter what.
    'SCR 106948 - LCI - 09/26/2006
    If rs!SIG_CODE <> "" Then
        mDetails = Right$(rs!SIG_CODE, 3)
        If mDetails = "||0" Then
            mDetails = Left$(rs!SIG_CODE, Len(rs!SIG_CODE) - 3)
        Else
            mDetails = rs!SIG_CODE
        End If
        mDetails = "SIG: " & mDetails
    End If
    
    If IsDate(rs!COMPLETED_DATE & vbNullString) Then
        mLastFilledDate = rs!COMPLETED_DATE & vbNullString
    End If
    
    miPriority = Val(rs!PROCESS_PRIORITY & vbNullString)
    
    DeserializeMed = True
    
    Exit Function
    
ErrHandler:
  goBB.SharedObj.ShowVBError Err.Number, Err.Description, "OrderMgt.DeserializeMed", "OrderMgt", vbExclamation

  Exit Function
  Resume

End Function


Friend Function Deserialize_Other(ByVal rs As Recordset) As Boolean

    If rs Is Nothing Then
        Exit Function
    End If
    On Error GoTo ErrHandler
    
    mDataId = rs!OE_OTR_DATA_ID
    mdOrderDate = rs!ORDER_DATE
    'SF SCR 77233, 77234, 77235, 77236
    If IsDate(rs!ORDER_DATE) Then
        mStartDate = rs!ORDER_DATE
    End If
    mTestName = rs!test_name
    mOtherType = rs!order_type
    mOrderSatus = rs!ORDER_STATUS
    mEncounterNumber = rs!ENC_NUM
    mPatientUnitNumber = rs!Unit_Number
    mPatientName = rs!Patient_Name
    mProviderName = rs!ordered_by
    mSEX = rs!SEX
    mComment = rs!prv_comment & vbNullString
    mPatientFMP = rs!patient_fmp
    mSponsorSSN = rs!sponsor_ssn
    mPatientSSN = mPatientFMP & "/" & mSponsorSSN
    mPatientStatus = Inpatient

    '82728 -- 12/22/05 - Frank
    mSign = Abs(Val(rs!SIGN_FLAG & vbNullString)) = 1
        
    '83102 -- 12/22/05 - Frank
    mNurseAck = Abs(Val(rs!NURSE_ACK_FLAG & vbNullString)) = 1

    Deserialize_Other = True
    
    Exit Function
    
ErrHandler:
  goBB.SharedObj.ShowVBError Err.Number, Err.Description, "OrderMgt.Deserialize", "OrderMgt", vbExclamation
End Function


