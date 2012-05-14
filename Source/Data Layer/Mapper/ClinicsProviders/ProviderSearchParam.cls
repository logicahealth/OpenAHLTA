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
Attribute VB_Name = "ProviderSearchParam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements iParameters
Implements iObjectPersist

Private Const PROVNCID As String = "ProviderNCID"
Private Const PROVIEN As String = "ProviderIEN"
Private Const PROVNAME As String = "EnterpriseNCID"
Private Const FACNCID As String = "FacilityNCID"
Private Const CLINNCID As String = "ClinicNCID"
Private Const PROVROLE As String = "ProviderRole"
Private Const INCROLES As String = "IncludeRoles"
Private Const INCWORK As String = "IncludeWorkload"
Private Const CLINUSER As String = "ClinicUsersOnly"
Private Const SIGNABLE As String = "SignableOnly"
Private Const INCEXT As String = "IncludeExtendedData"
Private Const UseCPTABLE As String = "UseClinicProviderTable"

Private mobjParams As iParameters

Private Sub Class_Initialize()
    On Error GoTo ErrHandler
    
    Set mobjParams = New ParametersGeneric
    
    'Initialize ProviderSearchParams
    SetValue PROVNCID, ""
    SetValue PROVIEN, ""
    SetValue PROVNAME, ""
    SetValue FACNCID, ""
    SetValue CLINNCID, ""
    SetValue PROVROLE, ""
    SetValue INCROLES, False
    SetValue INCWORK, False
    SetValue CLINUSER, False
    SetValue SIGNABLE, False
    SetValue INCEXT, False
    SetValue UseCPTABLE, True
    
    Exit Sub
    
ErrHandler:
    Err.Raise Err.Number, "ProviderSearchParam.Initialize", Err.Description, Err.HelpFile, Err.HelpContext
    
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

Public Property Get ProviderNCID() As String
    ProviderNCID = GetValue(PROVNCID)
End Property
Public Property Let ProviderNCID(ByVal NewValue As String)
    SetValue PROVNCID, NewValue
End Property

Public Property Get ProviderIEN() As String
    ProviderIEN = GetValue(PROVIEN)
End Property
Public Property Let ProviderIEN(ByVal NewValue As String)
    SetValue PROVIEN, NewValue
End Property

Public Property Get ProviderLastName() As String
    ProviderLastName = GetValue(PROVNAME)
End Property
Public Property Let ProviderLastName(ByVal NewValue As String)
    SetValue PROVNAME, NewValue
End Property

Public Property Get FacilityNCID() As String
    FacilityNCID = GetValue(FACNCID)
End Property
Public Property Let FacilityNCID(ByVal NewValue As String)
    SetValue FACNCID, NewValue
End Property

Public Property Get ClinicNCID() As String
    ClinicNCID = GetValue(CLINNCID)
End Property
Public Property Let ClinicNCID(ByVal NewValue As String)
    SetValue CLINNCID, NewValue
End Property

Public Property Get ProviderRole() As String
    ProviderRole = GetValue(PROVROLE)
End Property
Public Property Let ProviderRole(ByVal NewValue As String)
    SetValue PROVROLE, NewValue
End Property

Public Property Get IncludeRoles() As Boolean
    IncludeRoles = GetValue(INCROLES)
End Property
Public Property Let IncludeRoles(ByVal NewValue As Boolean)
    SetValue INCROLES, NewValue
End Property

Public Property Get IncludeWorkload() As Boolean
    IncludeWorkload = GetValue(INCWORK)
End Property
Public Property Let IncludeWorkload(ByVal NewValue As Boolean)
    SetValue INCWORK, NewValue
End Property

Public Property Get ClinicUsersOnly() As Boolean
    ClinicUsersOnly = GetValue(CLINUSER)
End Property
Public Property Let ClinicUsersOnly(ByVal NewValue As Boolean)
    SetValue CLINUSER, NewValue
End Property

Public Property Get SignableOnly() As Boolean
    SignableOnly = GetValue(SIGNABLE)
End Property
Public Property Let SignableOnly(ByVal NewValue As Boolean)
    SetValue SIGNABLE, NewValue
End Property

Public Property Get IncludeExtendedData() As Boolean
    IncludeExtendedData = GetValue(INCEXT)
End Property
Public Property Let IncludeExtendedData(ByVal NewValue As Boolean)
    SetValue INCEXT, NewValue
End Property

Public Property Get UseClinicProviderTable() As Boolean
    UseClinicProviderTable = GetValue(UseCPTABLE)
End Property
Public Property Let UseClinicProviderTable(ByVal NewValue As Boolean)
    SetValue UseCPTABLE, NewValue
End Property

Private Function GetValue(ByVal Name As String) As String
  If mobjParams.Exists(Name) Then
    GetValue = mobjParams.Item(Name).Value
  End If
End Function

Private Sub SetValue(ByVal Name As String, ByVal NewValue As String)
  If mobjParams.Exists(Name) Then
    mobjParams.Item(Name).Value = NewValue
  Else
    mobjParams.Add Name, NewValue
  End If
End Sub


