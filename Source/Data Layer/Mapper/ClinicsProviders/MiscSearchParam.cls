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
Attribute VB_Name = "MiscSearchParam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements iParameters

Private Const PHARMID As String = "PharmacyID"
Private Const WORKTYPE As String = "WorkloadType"
Private Const LOCATIONIEN As String = "LocationIEN"

Private mobjParams As iParameters

Private Sub Class_Initialize()
    Set mobjParams = New ParametersGeneric
    
    'Initialize MiscSearchParams
    SetValue PHARMID, True
    SetValue WORKTYPE, True
    SetValue LOCATIONIEN, False
End Sub

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

Public Property Get PharmacyID() As String
    PharmacyID = GetValue(PHARMID)
End Property
Public Property Let PharmacyID(ByVal NewValue As String)
    SetValue PHARMID, NewValue
End Property

Public Property Get WorkloadType() As String
    WorkloadType = GetValue(WORKTYPE)
End Property
Public Property Let WorkloadType(ByVal NewValue As String)
    SetValue WORKTYPE, NewValue
End Property

Public Property Get Location_IEN() As String
    Location_IEN = GetValue(WORKTYPE)
End Property
Public Property Let Location_IEN(ByVal NewValue As String)
    SetValue LOCATIONIEN, NewValue
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


