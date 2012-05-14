VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ParamsEncLock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Implements iParameters
Implements iObjectPersist

Public Enum EncSectionEnum
    eAutoCite = 1
    eRFV = 2
    eVitals = 3
    eSO = 4
    eLabs = 5
    eRads = 6
    eQuestionnaire = 7
    eap = 8
    eDisp = 9
    eNotes = 10
    eDental = 11
    eEducation = 12
    eSOOpen = 96
    eHistory = 97
    eHeader = 98
    eEntireEncounter = 99
End Enum

Private Const SECNUM As String = "SectionNumber"
Private Const EVTID As String = "EventID"
Private Const REASN As String = "Reason"
Private Const ITMNME As String = "ItemName"
Private Const ENCNUM As String = "EncounterNumber"
Private Const WRKSTN As String = "WorkStation"
Private Const USRNCID As String = "UserNCID"
Private Const FACNCID As String = "FacilityNCID"

Private mParams As iParameters
Public Property Get FacilityNCID() As String
  FacilityNCID = GetValue(FACNCID)
End Property

Public Property Let FacilityNCID(ByVal NewValue As String)
  SetValue FACNCID, NewValue
End Property
Public Property Get UserNCID() As String
  UserNCID = GetValue(USRNCID)
End Property

Public Property Let UserNCID(ByVal NewValue As String)
  SetValue USRNCID, NewValue
End Property
Public Property Get WorkStation() As String
  WorkStation = GetValue(WRKSTN)
End Property

Public Property Let WorkStation(ByVal NewValue As String)
  SetValue WRKSTN, NewValue
End Property
Public Property Get EncounterNumber() As String
  EncounterNumber = GetValue(ENCNUM)
End Property

Public Property Let EncounterNumber(ByVal NewValue As String)
  SetValue ENCNUM, NewValue
End Property
Public Property Get SectionNumber() As EncSectionEnum
  SectionNumber = GetValue(SECNUM)
End Property

Public Property Let SectionNumber(ByVal NewValue As EncSectionEnum)
  SetValue SECNUM, NewValue
End Property
Public Property Get EventID() As String
  EventID = GetValue(EVTID)
End Property

Public Property Let EventID(ByVal NewValue As String)
  SetValue EVTID, NewValue
End Property
Public Property Get Reason() As String
  Reason = GetValue(REASN)
End Property

Public Property Let Reason(ByVal NewValue As String)
  SetValue REASN, NewValue
End Property
Public Property Get ItemName() As String
  ItemName = GetValue(ITMNME)
End Property

Public Property Let ItemName(ByVal NewValue As String)
  SetValue ITMNME, NewValue
End Property

Private Sub Class_Initialize()
  Set mParams = New ParametersGeneric
End Sub

Private Function iObjectPersist_Copy() As Object
Dim Persist As iObjectPersist

  Set Persist = mParams
  Set iObjectPersist_Copy = Persist.Copy
  
End Function

Private Sub iObjectPersist_Persist(Writer As ObjectPersist.iWriter)
Dim Persist As iObjectPersist

  Set Persist = mParams
  Persist.Persist Writer
End Sub

Private Function iObjectPersist_Restore(Reader As ObjectPersist.iReader) As Boolean
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
    GetValue = mParams.Item(Name).Value
  End If
End Function

Private Sub SetValue(ByVal Name As String, ByVal NewValue As Variant)
  If mParams.Exists(Name) Then
    mParams.Item(Name).Value = NewValue
  Else
    mParams.Add Name, NewValue
  End If
End Sub


