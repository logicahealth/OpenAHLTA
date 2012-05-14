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
Attribute VB_Name = "RetrieveTypes_Parameters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements iParameters
Implements iObjectPersist

Private Const cDOMAIN As String = "Domain"
Private Const cRELATIONSHIP As String = "Relationship"
Private Const cCONTEXT As String = "Context"
Private Const cPREFONLYFLAG As String = "PrefOnlyFlag"

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

'Domain
Public Property Get Domain() As String
  Domain = GetValue(cDOMAIN)
End Property

Public Property Let Domain(ByRef NewValue As String)
    SetValue cDOMAIN, NewValue
End Property

'Relationship
Public Property Get Relationship() As String
  Relationship = GetValue(cRELATIONSHIP)
End Property

Public Property Let Relationship(ByRef NewValue As String)
    SetValue cRELATIONSHIP, NewValue
End Property

'Context
Public Property Get Context() As String
  Context = GetValue(cCONTEXT)
End Property

Public Property Let Context(ByRef NewValue As String)
    SetValue cCONTEXT, NewValue
End Property

'PrefOnlyFlag
Public Property Get PrefOnlyFlag() As String
  PrefOnlyFlag = GetValue(cPREFONLYFLAG)
End Property

Public Property Let PrefOnlyFlag(ByRef NewValue As String)
    SetValue cPREFONLYFLAG, NewValue
End Property

