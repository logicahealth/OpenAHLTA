VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AnPDiagnosis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Single diagnosis with state.  Problems (or WAMData) saves diagnoses but does not allow
'the data to be cached and saved at later time.

Option Explicit

Implements iGridRow

Public Enum enumDiagnosisState
  dsUnknown
  dsNew
  dsModified
  dsDelete
  dsRemove
  dsSaved
End Enum

Private mChronicity As String         'Acute/ Chronic
Private mClinicNCID As String
Private mComment As String
Private mDataID As String             'Unique ID
Private mEncID As String              'Encounter ID
Private mExtenderCD As String
Private mExtenderNCID As String
Private mExtenderNM As String
Private mFacalityNCID As String
Private mICD9Code As String           'Billing code
Private mKey As String                'Unique Key, during load it is same as mDataID
Private mLongNM As String
Private mSnoID As String
Private mName As String
Private mNCID As String
Private mOnSetDt As Date
Private mPrefix As String             'F for Family History, H for History
Private mPriority As Long             '1..n, unique
Private mShortNM As String
Private mType As String               'Follow up/ New
Private mFromProbList As Boolean      'SCR 49055 - used to determine if ICD9 conversion required in PGUI2

Private mCurrentState As enumDiagnosisState
Private mPrevState As enumDiagnosisState      'Saved state

Private mAssociations As OrderAssociations    'All orders that have been associated or
                                              'unassociated to this diagnosis

Private Sub Class_Initialize()
  Set mAssociations = New OrderAssociations
  
  mComment = Space$(1)    'Cheat to detect if a comment has been changed
End Sub

Private Function iGridRow_IsDiagnosis() As Boolean
  iGridRow_IsDiagnosis = True
End Function

Private Function iGridRow_IsOrderAssociation() As Boolean
  iGridRow_IsOrderAssociation = False
End Function

Public Property Get Chronicity() As String
  Chronicity = mChronicity
End Property

Public Property Let Chronicity(ByVal NewValue As String)
  If Len(mChronicity) > 0 Then State = dsModified
  
  mChronicity = NewValue
End Property

Public Property Get ClinicNCID() As String
  ClinicNCID = mClinicNCID
End Property

Public Property Let ClinicNCID(ByVal NewValue As String)
  mClinicNCID = NewValue
End Property

Public Property Get Comment() As String
  Comment = Trim$(mComment)
End Property

Public Property Let Comment(ByVal NewValue As String)
  If mComment <> Space$(1) Then State = dsModified
  
  mComment = Trim$(NewValue)
End Property

Public Property Get DataID() As String
  DataID = mDataID
End Property

Public Property Let DataID(ByVal NewValue As String)
  If Len(mDataID) > 0 Then Exit Property
  
  mDataID = NewValue
End Property

Public Property Get DiagnosisType() As String
  DiagnosisType = mType
End Property

Public Property Let DiagnosisType(ByVal NewValue As String)
  If Len(mType) > 0 Then State = dsModified
  
  mType = NewValue
End Property

Public Property Get EncounterID() As String
  EncounterID = mEncID
End Property

Public Property Let EncounterID(ByVal NewValue As String)
  mEncID = NewValue
End Property

Public Property Get ExtenderCode() As String
  ExtenderCode = mExtenderCD
End Property

Public Property Let ExtenderCode(ByVal NewValue As String)
  mExtenderCD = NewValue
End Property

Public Property Get ExtenderNCID() As String
  ExtenderNCID = mExtenderNCID
End Property

Public Property Let ExtenderNCID(ByVal NewValue As String)
  mExtenderNCID = NewValue
End Property

Public Property Get ExtenderName() As String
  ExtenderName = mExtenderNM
End Property

Public Property Let ExtenderName(ByVal NewValue As String)
  mExtenderNM = NewValue
End Property

Public Property Get FacilityNCID() As String
  FacilityNCID = mFacalityNCID
End Property

Public Property Let FacilityNCID(ByVal NewValue As String)
  mFacalityNCID = NewValue
End Property

'scr 29451  Phil Crowder 11/20/02
Public Property Get FullICDCode() As String
  If Len(mExtenderCD) > 2 Then
    FullICDCode = mExtenderCD
  Else
    FullICDCode = mICD9Code
  End If
End Property

'scr 29451  Phil Crowder 11/20/02
Public Property Get FullName() As String
  If Len(mExtenderNM) > 0 Then
    FullName = mName & "(" & mExtenderNM & ")"
  Else
    FullName = mName
  End If
End Property

Public Property Get ICD9Code() As String
  ICD9Code = mICD9Code
End Property

Public Property Let ICD9Code(ByVal NewValue As String)
  mICD9Code = NewValue
End Property

Public Property Get Key() As String
  Key = mKey
End Property

Public Property Let Key(ByVal NewValue As String)
  If Len(mKey) > 0 Then Exit Property
  
  mKey = NewValue
End Property

Public Property Get LongName() As String
  LongName = mLongNM
End Property

Public Property Let LongName(ByVal NewValue As String)
  mLongNM = NewValue
End Property

Public Property Get SnoID() As String
  SnoID = mSnoID
End Property

Public Property Let SnoID(ByVal NewValue As String)
  mSnoID = NewValue
End Property

Public Property Get Name() As String
  Name = mName
End Property

Public Property Let Name(ByVal NewValue As String)
  mName = NewValue
End Property

'JT BEGIN SCR 40955 - Added new property to indicate if diagnosis selected from problem list
Public Property Get FromProblemList() As Boolean
   FromProblemList = mFromProbList
End Property
Public Property Let FromProblemList(ByVal NewValue As Boolean)
   mFromProbList = NewValue
End Property
'JT END SCR 40955

Public Property Get NCID() As String
  NCID = mNCID
End Property

Public Property Let NCID(ByVal NewValue As String)
  mNCID = NewValue
End Property

Public Property Get OnsetDate() As Date
  OnsetDate = mOnSetDt
End Property

Public Property Let OnsetDate(ByVal NewValue As Date)
  mOnSetDt = NewValue
End Property

Public Property Get Prefix() As String
  Prefix = mPrefix
End Property

Public Property Let Prefix(ByVal NewValue As String)
  mPrefix = NewValue
End Property

Public Property Get Priority() As Long
  Priority = mPriority
End Property

Public Property Let Priority(ByVal NewValue As Long)
  If NewValue <= 0 Then Exit Property
  If mPriority <> 0 Then State = dsModified
  
  mPriority = NewValue
End Property

Public Property Get ShortName() As String
  ShortName = mShortNM
End Property

Public Property Let ShortName(ByVal NewValue As String)
  mShortNM = NewValue
End Property

Public Property Get State() As enumDiagnosisState
  State = mCurrentState
End Property

Public Property Let State(ByVal NewValue As enumDiagnosisState)

Dim ValidNewState As Boolean
Dim Association As OrderAssociation

  If NewValue = dsDelete And Len(mDataID) = 0 Then NewValue = dsRemove
  
  ValidNewState = IsValidState(NewValue)
  
  If ValidNewState Then
    If mCurrentState = dsUnknown Then
      mPrevState = NewValue
    Else
      mPrevState = mCurrentState
    End If
  
    mCurrentState = NewValue
    
    'SCR 49803 - Unassocation any orders from this diagnosis
    If mCurrentState = dsDelete Or mCurrentState = dsRemove Then
      For Each Association In mAssociations.Filter(typAll, fltActive)
         UnassociateOrder Association.OrderType, Association.OrderKey
      Next
    End If
    
  End If

End Property

Public Function AssociateOrder(ByRef Assoc As OrderAssociation) As Boolean
  AssociateOrder = mAssociations.AssociateOrder(Assoc)
End Function

Public Function Filter(Optional ByVal OrderType As enumOrderType = 0, _
                        Optional ByVal FilterType As enumFilterType = 0) As Collection
  Set Filter = mAssociations.Filter(OrderType, FilterType)
End Function

Public Sub Init(ByVal DataID As String, ByVal Name As String)
  mDataID = DataID
  mKey = DataID
  mName = Name
  mCurrentState = dsSaved
  mPrevState = dsSaved
End Sub

'Checks to see if order is associated to this diagnosis.
Public Function IsOrderAssociated(ByVal OrderType As enumOrderType, ByVal OrderKey As String, Optional ByVal SavedOnly As Boolean = False) As Boolean
Dim Associations As Collection
Dim Association As OrderAssociation

  On Error GoTo ErrHandler
  
  Set Associations = mAssociations.Filter(OrderType, fltActive)
  
  For Each Association In Associations
    If Association.OrderKey = OrderKey Then
      If SavedOnly Then
        IsOrderAssociated = (Association.State = asAssociated Or Association.State = asUnassociate)
      Else
        IsOrderAssociated = (Association.State = asAssociate Or _
                            Association.State = asAssociated Or _
                            Association.State = asQueued)
      End If
    End If
  Next
  
  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "AnPDiagnosis.IsOrderAssociated", "A/P", vbExclamation
  
End Function

Private Function IsValidState(ByVal NewValue As enumDiagnosisState) As Boolean

  Select Case mCurrentState
    Case dsUnknown
      IsValidState = True
      
    Case dsNew
      IsValidState = (NewValue = dsDelete Or NewValue = dsRemove)
    
    Case dsModified
      IsValidState = (NewValue = dsDelete Or NewValue = dsRemove)
      
    Case dsDelete, dsRemove
      IsValidState = False
      
    Case dsSaved
      IsValidState = (NewValue = dsModified Or NewValue = dsDelete Or NewValue = dsRemove)
  End Select

End Function

Private Sub RollbackState()
  mCurrentState = mPrevState
End Sub

'SCR 28663  3/11/02  Phil Crowder  Changed RTF format
'Returns diagnosis RTF and its associated orders RTFs as one string.
Public Function RTF(Optional ByVal SavedOnly As Boolean = False) As String
Dim DiagnosisRTF As String
Dim AssociatedRTF As String
Dim sCodeRTF As String

  On Error GoTo ErrHandler
    
  DiagnosisRTF = DIAGNOSIS_HEADER & mPriority & DIAGNOSIS_PERIOD & mName
  
  'JT BEGIN SCR 38958 - Build the appropriate code string based upon encounter option
  If gBB.EncounterInterface.IsIncludeRTFCodesSelected Then
    sCodeRTF = RTF_COMMENT_BEGIN_TAG & " (" & IIf(Len(mExtenderCD) > 0, mExtenderCD, mICD9Code) & ")" & RTF_COMMENT_END_TAG
  End If
  
  If Len(gShared.CmdLineSwitches("GREEN")) = 0 Then
    'SCR 17941  Phil Crowder 10/22/01     Added check for null value.
    If Len(mExtenderNM) > 0 Then
      DiagnosisRTF = DiagnosisRTF & DIAGNOSIS_FOOTER & sCodeRTF & DIAGNOSIS_EXT_NAME_HEADER & _
                     mExtenderNM & DIAGNOSIS_EXT_NAME_FOOTER_OLD
    Else
      DiagnosisRTF = DiagnosisRTF & DIAGNOSIS_FOOTER & sCodeRTF & END_LINE
    End If
  Else
    'SCR 17941  Phil Crowder 10/22/01     Added check for null value.
    If Len(mExtenderNM) > 0 Then
      DiagnosisRTF = DiagnosisRTF & DIAGNOSIS_FOOTER & sCodeRTF & DIAGNOSIS_EXT_NAME_HEADER & mExtenderNM & _
                      DIAGNOSIS_EXT_NAME_FOOTER
    Else
      DiagnosisRTF = DiagnosisRTF & DIAGNOSIS_FOOTER & sCodeRTF
    End If
  End If
  
  If Len(gShared.CmdLineSwitches("GREEN")) = 0 Then
    DiagnosisRTF = DiagnosisRTF & RTF_COMMENT_BEGIN_TAG & DIAGNOSIS_COMMENT_HEADER_OLD & _
                  RTFCodeFilter(mComment) & RTF_COMMENT_END_TAG & END_LINE
  Else
    If Len(Trim$(mComment)) > 0 Then
      DiagnosisRTF = DiagnosisRTF & RTF_COMMENT_BEGIN_TAG & DIAGNOSIS_COMMENT_HEADER & _
                  RTFCodeFilter(mComment) & RTF_COMMENT_END_TAG
    End If
  End If
  'JT END SCR 38958 - Updated RTF strings and fixed problem with multi-line comments giving sync warning
  
  AssociatedRTF = mAssociations.RTF(Me, SavedOnly)
  
  
  If Len(gShared.CmdLineSwitches("GREEN")) = 0 Then
    RTF = DiagnosisRTF & AssociatedRTF
  Else
    RTF = DiagnosisRTF & END_LINE & AssociatedRTF
  End If
  
  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "Diagnosis.RTF", "A/P", vbExclamation
  
End Function

Private Sub SaveCompleteState()
  
  Select Case mCurrentState
    Case dsNew, dsModified
      mCurrentState = dsSaved
      mPrevState = dsSaved
      
    Case dsDelete
      mCurrentState = dsRemove
      mPrevState = dsRemove
  End Select

End Sub

Public Sub AssociateQueuedOrders(ByVal OrderType As enumOrderType, ByVal OrderKey As String)
  On Error Resume Next
  
  mAssociations.Item(OrderType, OrderKey).State = asAssociate
End Sub

Public Sub UnassociateOrder(ByVal OrderType As enumOrderType, ByVal OrderKey As String)
  On Error Resume Next
  
   If mAssociations.Item(OrderType, OrderKey).State = asAssociate Then
      'SCR Item has not yet been associated so remove it
      mAssociations.RemoveAssocation OrderType, OrderKey
   Else
      'Item previously associated - instruct association to update problems
      mAssociations.Item(OrderType, OrderKey).State = asUnassociate
   End If

End Sub

'JT BEGIN SCR 44681 - used to remove a corrupt association that references a missing order
'Can't use UnassociateOrder the order because the save would then delete it from the DB which would
'only make the order warning message available once.  We want it to show up each time the encounter
'is opened so that we don't lose the reference to the missing order
Public Sub RemoveOrderAssocation(ByVal OrderType As enumOrderType, ByVal OrderKey As String)
   On Error Resume Next
   mAssociations.RemoveAssocation OrderType, OrderKey
End Sub
'JT END SCR 44681

Public Function UserHasPriv(ByVal PrivType As CHCSII_PrivilegeType) As Boolean
  
  On Error GoTo ErrHandler
  
  UserHasPriv = gBB.AnPChild(modDiagnosis).UserHasPriv(PrivType)
  
  Exit Function
    
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "Diagnosis.UserHasPriv", "A/P", vbExclamation
  
End Function

Public Function Save() As Boolean
Dim Required As Boolean
Dim Proxy As ProxyProblems
Dim Success As Boolean

  On Error GoTo ErrHandler
  
  Required = mCurrentState = dsDelete Or mCurrentState = dsModified Or _
              mCurrentState = dsSaved Or mCurrentState = dsNew
  
  If Not Required Then
    Save = True
    Exit Function
  End If
  
  Set Proxy = gBB.ProblemInterface
  
  Select Case mCurrentState
    Case dsNew
      Success = Proxy.AddNewDiagnosis(Me)
      If Success Then Success = SaveAssociations
      
    Case dsModified
      Success = Proxy.UpdateDiagnosis(Me)
      If Success Then Success = SaveAssociations
    
    Case dsDelete
      'Need to save related items for deleted diagnosis because they are all unassociated
      Success = SaveAssociations
      
      If Success Then
        Success = Proxy.DeleteDiagnosis(mDataID)
      End If
      
    Case dsSaved
      Success = SaveAssociations
  End Select
  
  If Success Then SaveCompleteState Else RollbackState
  
  Save = Success
    
  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "AnPDiagnosis.Save", "A/P", vbExclamation
  
End Function

Private Function SaveAssociations() As Boolean
Dim Success As Boolean

  On Error GoTo ErrHandler
  
  Success = mAssociations.Save(mDataID)
  
  If Success And Not gShared.IsAppMode(modeCHCSI_GUI) Then Success = gBB.ProblemInterface.SaveCachedWrites
  
  SaveAssociations = Success
  
  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "AnPDiagnosis.SaveAssociations", "A/P", vbExclamation
  
End Function

'Order data id is not set until the order is saved.  After the order is saved, then all
'order associations are updated with the order data id.  The order data id is used when
'saving the order association.
Public Sub UpdateOrderDataID(ByVal OrderType As enumOrderType, ByVal OrderKey As String, ByVal OrderDataID As String)
  On Error Resume Next
  
  mAssociations.Item(OrderType, OrderKey).OrderDataID = OrderDataID

End Sub
