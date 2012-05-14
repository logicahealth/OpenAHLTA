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
Attribute VB_Name = "ClinicSearchParam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements iParameters
Implements iObjectPersist

Private Const CLINNCID As String = "ClinicNCID"
Private Const CLINIEN As String = "ClinicIEN"
Private Const FACNCID As String = "FacilityNCID"
Private Const USRNCID As String = "UserNCID"
Private Const INCNONMTF As String = "IncludeNonMTF"
Private Const INCWORK As String = "IncludeWorkload"
Private Const APPTCOD As String = "ApptCode"
Private Const INCUSRONLY As String = "IncludeUserClinicsOnly"
Private Const UseCPTABLE As String = "UseClinicProviderTable"
Private Const INCIBWACLIN As String = "IncludeIBWAClinics"

Private mobjParams As iParameters

Private Sub Class_Initialize()

    On Error GoTo ErrHandler
    
    Set mobjParams = New ParametersGeneric
    
    'Initialize ClinicSearchParams
    SetValue CLINNCID, ""
    SetValue CLINIEN, ""
    SetValue FACNCID, ""
    SetValue USRNCID, ""
    SetValue INCNONMTF, True
    SetValue INCWORK, False
    SetValue APPTCOD, ""
    SetValue INCUSRONLY, False
    SetValue UseCPTABLE, True
    SetValue INCIBWACLIN, False
    
    Exit Sub
    
ErrHandler:
    Err.Raise Err.Number, "ClinicSearchParam.Initialize", Err.Description, Err.HelpFile, Err.HelpContext
    
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

Public Property Get ClinicNCID() As String
    ClinicNCID = GetValue(CLINNCID)
End Property
Public Property Let ClinicNCID(ByVal NewValue As String)
    SetValue CLINNCID, NewValue
End Property

Public Property Get ClinicIEN() As String
    ClinicIEN = GetValue(CLINIEN)
End Property
Public Property Let ClinicIEN(ByVal NewValue As String)
    SetValue CLINIEN, NewValue
End Property

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

Public Property Get IncludeNonMTF() As Boolean
    IncludeNonMTF = GetValue(INCNONMTF)
End Property
Public Property Let IncludeNonMTF(ByVal NewValue As Boolean)
    SetValue INCNONMTF, NewValue
End Property

Public Property Get IncludeWorkload() As Boolean
    IncludeWorkload = GetValue(INCWORK)
End Property
Public Property Let IncludeWorkload(ByVal NewValue As Boolean)
    SetValue INCWORK, NewValue
End Property

Public Property Get ApptCode() As String
    ApptCode = GetValue(APPTCOD)
End Property
Public Property Let ApptCode(ByVal NewValue As String)
    SetValue APPTCOD, NewValue
End Property

Public Property Get IncludeUserClinicsOnly() As Boolean
    IncludeUserClinicsOnly = GetValue(INCUSRONLY)
End Property
Public Property Let IncludeUserClinicsOnly(ByVal NewValue As Boolean)
    SetValue INCUSRONLY, NewValue
End Property

Public Property Get UseClinicProviderTable() As Boolean
    UseClinicProviderTable = GetValue(UseCPTABLE)
End Property
Public Property Let UseClinicProviderTable(ByVal NewValue As Boolean)
    SetValue UseCPTABLE, NewValue
End Property

Public Property Get IncludeIBWAClinics() As Boolean
  IncludeIBWAClinics = GetValue(INCIBWACLIN)
End Property
Public Property Let IncludeIBWAClinics(ByVal NewValue As Boolean)
  SetValue INCIBWACLIN, NewValue
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

