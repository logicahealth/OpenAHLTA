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
Attribute VB_Name = "CacheNullEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'A Cache Entry without a value that has expired and is not valid.

Option Explicit

Implements iCacheEntry

Private mKey As String        'Unique Identifier for cached item

Private Property Get iCacheEntry_Copy() As ObjectPersist.iObjectPersist
  Set iCacheEntry_Copy = Nothing
End Property

Private Property Let iCacheEntry_ExpirationDate(ByVal RHS As Date)
'
End Property

Private Property Get iCacheEntry_ExpirationDate() As Date
  iCacheEntry_ExpirationDate = #1/1/1970#
End Property

Private Property Let iCacheEntry_Key(ByVal RHS As String)
  If Len(mKey) = 0 Then mKey = RHS
End Property

Private Property Get iCacheEntry_Key() As String
  iCacheEntry_Key = mKey
End Property

Private Property Let iCacheEntry_Valid(ByVal RHS As Boolean)
'
End Property

Private Property Get iCacheEntry_Valid() As Boolean
  iCacheEntry_Valid = False
End Property

Private Property Set iCacheEntry_Value(RHS As Variant)
'
End Property

Private Property Let iCacheEntry_Value(RHS As Variant)
'
End Property

Private Property Get iCacheEntry_Value() As Variant
'
End Property
