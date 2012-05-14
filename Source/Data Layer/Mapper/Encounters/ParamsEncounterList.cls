VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ParamsEncList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Implements iParameters
Implements iObjectPersist

Public Enum EnctoReturn
    eAll = 0
    eOpen = 1
    ePrev = 2
    eCosign = 4
    eInpatientNotes = 8
End Enum

Public Enum ENCTYPE     '-- this is really enc category, for the first 3, and category is appt 'type'
  Inpatient = 0
  Outpatient = 1
  Ambulatory = 2
  TeleConsult = 3
  Dental = 4
  InpatientNote = 5
End Enum

Private Const PROVNCID As String = "ProviderNCID"
Private Const UNITNUM As String = "PatientUnitNumber"
Private Const GROUP As String = "Grouping"
Private Const ENCTYPE As String = "EncType"
Private Const INCSECT As String = "IncludeSections"
Private Const INCDOCS As String = "IncludeDocuments"

Private mParams As iParameters
Public Property Get IncludeDocuments() As Boolean
  IncludeDocuments = GetValue(INCDOCS)
End Property

Public Property Let IncludeDocuments(ByVal NewValue As Boolean)
  SetValue INCDOCS, NewValue
End Property
Public Property Get IncludeSections() As Boolean
  IncludeSections = GetValue(INCSECT)
End Property

Public Property Let IncludeSections(ByVal NewValue As Boolean)
  SetValue INCSECT, NewValue
End Property
Public Property Get Grouping() As EnctoReturn
  Grouping = GetValue(GROUP)
End Property

Public Property Let Grouping(ByVal NewValue As EnctoReturn)
  SetValue GROUP, NewValue
End Property
Public Property Get PatientUnitNumber() As String
  PatientUnitNumber = GetValue(UNITNUM)
End Property

Public Property Let PatientUnitNumber(ByVal NewValue As String)
  SetValue UNITNUM, NewValue
End Property
Public Property Get ProviderNCID() As String
  ProviderNCID = GetValue(PROVNCID)
End Property

Public Property Let ProviderNCID(ByVal NewValue As String)
  SetValue PROVNCID, NewValue
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
