VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CacheEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Container for a value to be cached in memory or the file system.

Option Explicit

Implements iCacheEntry

Private mExpirationDate As Date         'Date/Time value is no longer valid
Private mKey As String                  'Unique Identifier for cached item
Private mValid As Boolean               'True if the value is still valid and False if the value is no longer valid (based on expiration date)
Private mValue As Variant               'Cached value

Private Sub Class_Initialize()
  mExpirationDate = #12/12/9999#        'Default used to mean the value does not expire (at least in my lifetime)
End Sub

Private Property Get iCacheEntry_Copy() As ObjectPersist.iObjectPersist
Dim Obj As Object
  
  Set Obj = mValue
  Set iCacheEntry_Copy = CopyObject(Obj)
  
End Property

Private Property Let iCacheEntry_ExpirationDate(ByVal RHS As Date)
  If mExpirationDate = #12/12/9999# Then mExpirationDate = RHS
End Property

Private Property Get iCacheEntry_ExpirationDate() As Date
  iCacheEntry_ExpirationDate = mExpirationDate
End Property

Private Property Let iCacheEntry_Key(ByVal RHS As String)
  If Len(mKey) = 0 Then mKey = RHS
End Property

Private Property Get iCacheEntry_Key() As String
  iCacheEntry_Key = mKey
End Property

Private Property Let iCacheEntry_Valid(ByVal RHS As Boolean)
  mValid = RHS
End Property

Private Property Get iCacheEntry_Valid() As Boolean
  iCacheEntry_Valid = mValid
End Property

Private Property Set iCacheEntry_Value(RHS As Variant)
  
  If Not IsEmpty(mValue) Then Exit Property
  
  If VarType(RHS) = vbObject Then
    Set mValue = RHS
  End If
  
End Property

Private Property Let iCacheEntry_Value(RHS As Variant)

  If Not IsEmpty(mValue) Then Exit Property

  If VarType(mValue) = vbObject Then
    Set mValue = RHS
  Else
    mValue = RHS
  End If

End Property

Private Property Get iCacheEntry_Value() As Variant
  If VarType(mValue) = vbObject Then
    Set iCacheEntry_Value = mValue
  Else
    iCacheEntry_Value = mValue
  End If
End Property
