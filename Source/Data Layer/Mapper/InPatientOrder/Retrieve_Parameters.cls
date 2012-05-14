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
Attribute VB_Name = "Retrieve_Parameters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

Implements iParameters
Implements iObjectPersist

Private Const cSTATUS As String = "Status"
Private Const cUSERNCID As String = "UserNCID"
Private Const cORDERTYPE As String = "OrderType"
Private Const cOTHERORDERTYPE As String = "OtherOrderType"
Private Const cUNITNUMBER As String = "UnitNumber"
Private Const cORDER_ID As String = "Order_Id"
Private Const cPATIENTSTATUS As String = "PatientStatus"
Private Const cINPATIENTID As String = "InpatientID"
Private Const cENCOUNTERNUMBER As String = "EncounterNumber"

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

On Error GoTo ErrHandler

  If mParams.Exists(Name) Then
    If IsObject(NewValue) Then
      Set mParams.Item(Name).Value = NewValue
    Else
      mParams.Item(Name).Value = NewValue
    End If
  Else
    mParams.Add Name, NewValue
  End If
  
Exit Sub

ErrHandler:

  Err.Raise Err.Number, CACHE_NAME & ".Retrieve_Parameters.SetValue", Err.Description, Err.HelpFile, Err.HelpContext

End Sub
'Patient Status
Public Property Get PatientStatus() As Inpatient_Utilities.enumPatientStatus
  PatientStatus = GetValue(cPATIENTSTATUS)
End Property

Public Property Let PatientStatus(ByRef NewValue As Inpatient_Utilities.enumPatientStatus)
    SetValue cPATIENTSTATUS, NewValue
End Property


'Status
Public Property Get Status() As String
  Status = GetValue(cSTATUS)
End Property

Public Property Let Status(ByRef NewValue As String)
    SetValue cSTATUS, NewValue
End Property

'UserNCID
Public Property Get UserNCID() As String
  UserNCID = GetValue(cUSERNCID)
End Property

Public Property Let UserNCID(ByRef NewValue As String)
    SetValue cUSERNCID, NewValue
End Property

'Order type
Public Property Get OrderType() As String
    OrderType = GetValue(cORDERTYPE)
End Property


Public Property Let OrderType(ByRef NewValue As String)
    SetValue cORDERTYPE, NewValue
End Property

'Unit Number
Public Property Get UnitNumber() As String
    UnitNumber = GetValue(cUNITNUMBER)
End Property


Public Property Let UnitNumber(ByRef NewValue As String)
    SetValue cUNITNUMBER, NewValue
End Property


'Order type for Other Orders
Public Property Get OtherOrderType() As String
    OtherOrderType = GetValue(cOTHERORDERTYPE)
End Property


Public Property Let OtherOrderType(ByRef NewValue As String)
    SetValue cOTHERORDERTYPE, NewValue
End Property

'Order id for Other Orders
Public Property Get OrderId() As String
    OrderId = GetValue(cORDER_ID)
End Property


Public Property Let OrderId(ByRef NewValue As String)
    SetValue cORDER_ID, NewValue
End Property


'Inpatient id for addmission
Public Property Get InpatientId() As String
    InpatientId = GetValue(cINPATIENTID)
End Property


Public Property Let InpatientId(ByRef NewValue As String)
    SetValue cINPATIENTID, NewValue
End Property



'Encounter Number for Inpatient Notes
Public Property Get EncounterNumber() As String
    EncounterNumber = GetValue(cENCOUNTERNUMBER)
End Property


Public Property Let EncounterNumber(ByRef NewValue As String)
    SetValue cENCOUNTERNUMBER, NewValue
End Property


