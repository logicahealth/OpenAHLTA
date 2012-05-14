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
Private Const cUNITNUMBER As String = "UnitNumber"
Private Const cORDER_ID As String = "Order_Id"
Private Const cOrderIEN As String = "Order_IEN"


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


'Unit Number
Public Property Get UnitNumber() As String
    UnitNumber = GetValue(cUNITNUMBER)
End Property


Public Property Let UnitNumber(ByRef NewValue As String)
    SetValue cUNITNUMBER, NewValue
End Property


'Order id for Other Orders
Public Property Get OrderId() As String
    OrderId = GetValue(cORDER_ID)
End Property


Public Property Let OrderId(ByRef NewValue As String)
    SetValue cORDER_ID, NewValue
End Property



'Order IEN
Public Property Get OrderIEN() As String
    OrderIEN = GetValue(cOrderIEN)
End Property


Public Property Let OrderIEN(ByRef NewValue As String)
    SetValue OrderIEN, NewValue
End Property


