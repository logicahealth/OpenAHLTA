VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TransactionSearchParams"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'--- begin DTO Builder; DTO=AllergyReactionSearchParam
'--- 1/22/2004 11:03:01 AM
'---
'--- NOTE: Any time you modify this class, increment the version number
'
'

Implements iParameters
Implements iObjectPersist

Private Const OBJECT_VERSION As String = "1.0"                            'object version
Private Const UNIQUE_IDENTIFIER As String = "AllergyReactionSearchParam"

Private mobjParams As iParameters

Private mUnitNumber As String           'unit number
Private mTransType As String            'transaction type
Private mSynchronize As Boolean         'synchronzie
Private mIncludeMMIData As Boolean      'should MMI data be returned?

Private Const cUNIT_NUMBER As String = "UnitNumber"
Private Const cTRANS_TYPE As String = "TransType"
Private Const cSYNCHRONIZE As String = "Synchronize"
Private Const cINCLUDE_MMI_DATA As String = "IncludeMMIData"

Private Sub Class_Initialize()
    Set mobjParams = New ParametersGeneric
End Sub
'--- end init / terminate procedures

'--- begin property lets/gets
Public Property Get UnitNumber() As String
    UnitNumber = GetStringValue(cUNIT_NUMBER)
End Property

Public Property Let UnitNumber(ByVal Value As String)
    SetStringValue cUNIT_NUMBER, Value
End Property

Public Property Get IncludeMMIData() As Boolean
    IncludeMMIData = GetBoolValue(cINCLUDE_MMI_DATA)
End Property

Public Property Let IncludeMMIData(ByVal Value As Boolean)
    SetBoolValue cINCLUDE_MMI_DATA, Value
End Property

Public Property Get Synchronize() As Boolean
    Synchronize = GetBoolValue(cSYNCHRONIZE)
End Property

Public Property Let Synchronize(ByVal Value As Boolean)
    SetBoolValue cSYNCHRONIZE, Value
End Property

Public Property Get TransType() As String
    TransType = GetStringValue(cTRANS_TYPE)
End Property

Public Property Let TransType(ByVal Value As String)
    SetStringValue cTRANS_TYPE, Value
End Property
'--- end property lets/gets

Private Function iObjectPersist_Copy() As Object
    Dim persist As iObjectPersist

    Set persist = mobjParams
    Set iObjectPersist_Copy = persist.Copy
End Function

Private Sub iObjectPersist_Persist(ByRef Writer As ObjectPersist.iWriter)
    Dim persist As iObjectPersist

    Set persist = mobjParams
    persist.persist Writer
End Sub

Private Function iObjectPersist_Restore(ByRef Reader As ObjectPersist.iReader) As Boolean
    Dim persist As iObjectPersist

    Set persist = mobjParams
    iObjectPersist_Restore = persist.Restore(Reader)
End Function


Private Property Get iObjectPersist_UniqueIdentifier() As String
    Dim persist As iObjectPersist

    Set persist = mobjParams
    iObjectPersist_UniqueIdentifier = persist.UniqueIdentifier
End Property

Private Property Get iObjectPersist_Version() As String
    Dim persist As iObjectPersist

    Set persist = mobjParams
    iObjectPersist_Version = persist.Version
End Property
'--- end DTO Builder

Private Function iParameters_Add(ByVal Name As String, ByVal Value As Variant) As DL_Support.Parameter
    mobjParams.Add Name, Value
End Function

Private Function iParameters_Count() As Long
    iParameters_Count = mobjParams.count
End Function

Private Function iParameters_Exists(ByVal Name As String) As Boolean
    iParameters_Exists = mobjParams.Exists(Name)
End Function

Private Function iParameters_Item(ByVal Key As Variant) As DL_Support.Parameter
    Set iParameters_Item = mobjParams.Item(Key)
End Function

Private Sub iParameters_Remove(ByVal Key As Variant)
    mobjParams.Remove (Key)
End Sub


Private Function GetStringValue(ByVal Name As String) As String
  If mobjParams.Exists(Name) Then
    GetStringValue = mobjParams.Item(Name).Value
  Else
    GetStringValue = vbNullString
  End If
End Function

Private Sub SetStringValue(ByVal Name As String, ByVal NewValue As String)
  If mobjParams.Exists(Name) Then
    mobjParams.Item(Name).Value = NewValue
  Else
    mobjParams.Add Name, NewValue
  End If
End Sub

Private Function GetBoolValue(ByVal Name As String) As Boolean
  If mobjParams.Exists(Name) Then
    GetBoolValue = mobjParams.Item(Name).Value
  Else
    GetBoolValue = False
  End If
End Function

Private Sub SetBoolValue(ByVal Name As String, ByVal NewValue As Boolean)
  If mobjParams.Exists(Name) Then
    mobjParams.Item(Name).Value = NewValue
  Else
    mobjParams.Add Name, NewValue
  End If
End Sub


