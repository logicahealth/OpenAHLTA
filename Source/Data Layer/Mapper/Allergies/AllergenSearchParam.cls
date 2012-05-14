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
Attribute VB_Name = "AllergenSearchParam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'--- begin DTO Builder; DTO=AllergenSearchParam
'--- 1/22/2004 11:03:01 AM
'---
'--- NOTE: Any time you modify this class, increment the version number
'
'

Implements iParameters
Implements iObjectPersist

Private Const OBJECT_VERSION As String = "1.0"
Private Const UNIQUE_IDENTIFIER As String = "AllergenSearchParam"
Private Const cPREFIX_KEYWORD As String = "PrefixKeyword"

Private mobjParams As iParameters

Private Sub Class_Initialize()
    Set mobjParams = New ParametersGeneric
End Sub
'--- end init / terminate procedures


'--- begin property lets/gets
Public Property Get PrefixKeyword() As String
    PrefixKeyword = GetStringValue(cPREFIX_KEYWORD)
End Property

Public Property Let PrefixKeyword(ByVal Value As String)
    SetStringValue cPREFIX_KEYWORD, Value
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

