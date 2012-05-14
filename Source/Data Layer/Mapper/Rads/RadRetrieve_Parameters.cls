VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Retrieve_Parameters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements iParameters
Implements iObjectPersist

Private Const cUNIT_NUMBER As String = "Unit_Number"
Private Const cIEN_NUMBER As String = "IEN_Number"
Private Const cSUMMARYDATAPULL As String = "SummaryDataPull"
Private Const cTIMESEARCHTYPE As String = "TimeSearchType"
Private Const cTIMEOBJECT As String = "TimeObject"
Private Const cFILTERCOLLECTION As String = "FilterCollection"
Private Const cFORCEREFRESH As String = "ForceRefresh"
Private Const cEVENTID As String = "EventID"
Private Const cLOGONFACILITYNCID As String = "LogonFacilityNCID"

Private mParams As iParameters                      'ParametersGeneric class used to store and persist the NCIDs

Private Sub Class_Initialize()
  Set mParams = New ParametersGeneric
End Sub

Private Function iObjectPersist_Copy() As Object
Dim Persist As iObjectPersist

  Set Persist = mParams
  Set iObjectPersist_Copy = Persist.Copy
  
End Function

Private Sub iObjectPersist_Persist(ByRef Writer As ObjectPersist.iWriter)
Dim Persist As iObjectPersist

  Set Persist = mParams
  Persist.Persist Writer
End Sub

Private Function iObjectPersist_Restore(ByRef Reader As ObjectPersist.iReader) As Boolean
Dim Persist As iObjectPersist

  Set Persist = mParams
  iObjectPersist_Restore = Persist.Restore(Reader)
  
End Function

Private Property Get iObjectPersist_UniqueIdentifier() As String
Dim Persist As iObjectPersist

  Set Persist = mParams
  iObjectPersist_UniqueIdentifier = Persist.UniqueIdentifier
  
End Property

Private Property Get iObjectPersist_Version() As String
Dim Persist As iObjectPersist

  Set Persist = mParams
  iObjectPersist_Version = Persist.Version
  
End Property

Private Function iParameters_Add(ByVal Name As String, ByVal Value As Variant) As DL_Support.Parameter
  Set iParameters_Add = mParams.Add(Name, Value)
End Function

Private Function iParameters_Count() As Long
  iParameters_Count = mParams.Count
End Function

Private Function iParameters_Exists(ByVal Name As String) As Boolean
  iParameters_Exists = mParams.Exists(Name)
End Function

Private Function iParameters_Item(ByVal Key As Variant) As DL_Support.Parameter
  Set iParameters_Item = mParams.Item(Key)
End Function

Private Sub iParameters_Remove(ByVal Key As Variant)
  mParams.Remove Key
End Sub

Private Function GetValue(ByVal Name As String) As Variant
  If mParams.Exists(Name) Then
    If IsObject(mParams.Item(Name).Value) Then
      Set GetValue = mParams.Item(Name).Value
    Else
      GetValue = mParams.Item(Name).Value
    End If
  End If
End Function

Private Sub SetValue(ByVal Name As String, ByVal NewValue As Variant)
  If mParams.Exists(Name) Then
    If IsObject(NewValue) Then
      Set mParams.Item(Name).Value = NewValue
    Else
      mParams.Item(Name).Value = NewValue
    End If
  Else
    mParams.Add Name, NewValue
  End If
End Sub

'Unit_Number
Public Property Get Unit_Number() As Double
  Unit_Number = GetValue(cUNIT_NUMBER)
End Property

Public Property Let Unit_Number(ByRef NewValue As Double)
    SetValue cUNIT_NUMBER, NewValue
End Property

'IEN_Number
Public Property Get IEN_Number() As Double
  IEN_Number = GetValue(cIEN_NUMBER)
End Property

Public Property Let IEN_Number(ByRef NewValue As Double)
    SetValue cIEN_NUMBER, NewValue
End Property

'LogonFacilityNCID
Public Property Get LogonFacilityNCID() As Double
  LogonFacilityNCID = GetValue(cLOGONFACILITYNCID)
End Property

Public Property Let LogonFacilityNCID(ByRef NewValue As Double)
    SetValue cLOGONFACILITYNCID, NewValue
End Property

'SummaryDataPull
Public Property Get SummaryDataPull() As Boolean
  SummaryDataPull = GetValue(cSUMMARYDATAPULL)
End Property

Public Property Let SummaryDataPull(ByRef NewValue As Boolean)
    SetValue cSUMMARYDATAPULL, NewValue
End Property

'TimeObject
Public Property Get TimeObject() As CWTime
  Set TimeObject = GetValue(cTIMEOBJECT)
End Property

Public Property Let TimeObject(ByRef NewValue As CWTime)
  SetValue cTIMEOBJECT, NewValue
End Property

'TimeSearchType
Public Property Get TimeSearchType() As Integer
  TimeSearchType = GetValue(cTIMESEARCHTYPE)
End Property

Public Property Let TimeSearchType(ByRef NewValue As Integer)
    SetValue cTIMESEARCHTYPE, NewValue
End Property

'FilterCollection
Public Property Get FilterCollection() As Collection
  Set FilterCollection = GetValue(cFILTERCOLLECTION)
End Property

Public Property Let FilterCollection(ByRef NewValue As Collection)
  SetValue cFILTERCOLLECTION, NewValue
End Property

'EventID
Public Property Get EventID() As String
  EventID = GetValue(cEVENTID)
End Property

Public Property Let EventID(ByRef NewValue As String)
    SetValue cEVENTID, NewValue
End Property

'ForceRefresh
Public Property Get ForceRefresh() As Boolean
  ForceRefresh = GetValue(cFORCEREFRESH)
End Property

Public Property Let ForceRefresh(ByRef NewValue As Boolean)
    SetValue cFORCEREFRESH, NewValue
End Property
