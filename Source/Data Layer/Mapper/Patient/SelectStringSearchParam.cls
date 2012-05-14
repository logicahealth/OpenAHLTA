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
Attribute VB_Name = "SelectStringSearchParam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True


Option Explicit

Implements iParameters
Implements iObjectPersist

Private Const cSELECTSTR As String = "SelectStringType"
Private Const cNCID      As String = "NCID"

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

Private Function GetValue(ByVal Name As String) As String
  If mParams.Exists(Name) Then
    GetValue = mParams.Item(Name).Value
  End If
End Function

Private Sub SetValue(ByVal Name As String, ByVal NewValue As String)
  If mParams.Exists(Name) Then
    mParams.Item(Name).Value = NewValue
  Else
    mParams.Add Name, NewValue
  End If
End Sub



Public Property Get SelectStringCategory() As DTOs.enumSelectStringCategory
    If GetValue(cSELECTSTR) = vbNullString Then
        SelectStringCategory = osUnknownCategory
    Else
        SelectStringCategory = GetValue(cSELECTSTR)
    End If
End Property
Public Property Let SelectStringCategory(ByRef NewValue As DTOs.enumSelectStringCategory)
    SetValue cSELECTSTR, NewValue
End Property



Public Property Get NCID() As String
    If GetValue(cNCID) = vbNullString Then
        NCID = vbNullString
    Else
        NCID = GetValue(cNCID)
    End If
End Property
Public Property Let NCID(ByRef NewValue As String)
    SetValue cNCID, NewValue
End Property


