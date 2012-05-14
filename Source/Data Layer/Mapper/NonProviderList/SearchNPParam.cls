VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SearchNPParams"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements iParameters
Implements iObjectPersist

Public Enum ListMode
    Provider = 0
    AdditionalProvider = 1
End Enum

Private mobjParams As iParameters

Private mstrNursingDivisionNCID As String
Private mstrFacilityNCID As String
Private mstrUserNCID As String
Private mstrEncounterProviderNCID As String
Private mstrClinicianNCID As String
Private mstrProviderIEN As String
Private mtypListMode As ListMode


Private Sub Class_Initialize()
    Set mobjParams = New ParametersGeneric
    
    'Default to normal provider
    mtypListMode = Provider
End Sub

Private Function iObjectPersist_Copy() As Object
Dim Persist As iObjectPersist

  Set Persist = mobjParams
  Set iObjectPersist_Copy = Persist.Copy

End Function

Private Sub iObjectPersist_Persist(Writer As ObjectPersist.iWriter)
Dim Persist As iObjectPersist

  Set Persist = mobjParams
  Persist.Persist Writer

End Sub

Private Function iObjectPersist_Restore(Reader As ObjectPersist.iReader) As Boolean
Dim Persist As iObjectPersist

  Set Persist = mobjParams
  iObjectPersist_Restore = Persist.Restore(Reader)
End Function


Private Property Get iObjectPersist_UniqueIdentifier() As String
Dim Persist As iObjectPersist

  Set Persist = mobjParams
  iObjectPersist_UniqueIdentifier = Persist.UniqueIdentifier
End Property

Private Property Get iObjectPersist_Version() As String
Dim Persist As iObjectPersist

  Set Persist = mobjParams
  iObjectPersist_Version = Persist.Version
End Property

Private Function iParameters_Add(ByVal Name As String, ByVal Value As Variant) As DL_Support.Parameter
    mobjParams.Add Name, Value
End Function

Private Function iParameters_Count() As Long
    iParameters_Count = mobjParams.Count
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

Public Property Get NursingDivisionNCID() As String
    NursingDivisionNCID = mstrNursingDivisionNCID
End Property
Public Property Let NursingDivisionNCID(ByVal NewValue As String)
    mstrNursingDivisionNCID = NewValue
End Property

Public Property Get FacilityNCID() As String
    FacilityNCID = mstrFacilityNCID
End Property
Public Property Let FacilityNCID(ByVal NewValue As String)
    mstrFacilityNCID = NewValue
End Property

Public Property Get UserNCID() As String
    UserNCID = mstrUserNCID
End Property
Public Property Let UserNCID(ByVal NewValue As String)
    mstrUserNCID = NewValue
End Property

Public Property Get EncounterProviderNCID() As String
    EncounterProviderNCID = mstrEncounterProviderNCID
End Property
Public Property Let EncounterProviderNCID(ByVal NewValue As String)
    mstrEncounterProviderNCID = NewValue
End Property

Public Property Get ClinicianNCID() As String
    ClinicianNCID = mstrClinicianNCID
End Property
Public Property Let ClinicianNCID(ByVal NewValue As String)
    mstrClinicianNCID = NewValue
End Property

Public Property Get ProviderIEN() As String
    ProviderIEN = mstrProviderIEN
End Property
Public Property Let ProviderIEN(ByVal NewValue As String)
    mstrProviderIEN = NewValue
End Property

Public Property Get ListMode() As ListMode
    ListMode = mtypListMode
End Property
Public Property Let ListMode(ByVal NewValue As ListMode)
    mtypListMode = NewValue
End Property
