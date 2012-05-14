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
Attribute VB_Name = "CacheMapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Cache service interface for data mapper objects.  It reduces the number of parameters needed to save and retrieve
'cached items.

Option Explicit

Private mMapperName As String       'Name of Mapper that is going to access the Cache (unique per mapper interface)
Private mCache As Cache             'Reference to Cache object.

Public Property Get MapperName() As String
  MapperName = mMapperName
End Property

Public Property Let MapperName(ByVal NewValue As String)
  If Len(mMapperName) = 0 Then mMapperName = NewValue
End Property

Public Sub Init(ByRef Cache As Cache)
  Set mCache = Cache
End Sub

Public Sub Remove(ByVal Method As String, Optional ByRef Parameters As iParameters = Nothing)
  mCache.Remove mMapperName, Method, Parameters
End Sub

Public Function Retrieve(ByVal Method As String, Optional ByRef Parameters As iParameters = Nothing) As iCacheEntry
  Set Retrieve = mCache.Retrieve(mMapperName, Method, Parameters)
End Function

Public Function RetrieveLocal(ByVal Method As String, Optional ByRef Parameters As iParameters = Nothing) As iCacheEntry
  Set RetrieveLocal = mCache.Retrieve(mMapperName, Method, Parameters)
End Function

Public Function Save(ByVal Method As String, ByRef Value As iObjectPersist, Optional ByRef Parameters As iParameters = Nothing, Optional ByVal ExpirationDate As Date = #12/12/9999#) As Boolean
  Save = mCache.Save(mMapperName, Method, Value, Parameters, ExpirationDate)
End Function

Public Function SaveLocal(ByVal Method As String, ByVal SerializeData As String, Optional ByRef Parameters As iParameters = Nothing, Optional ByVal ExpirationDate As Date = #12/12/9999#) As Boolean
  SaveLocal = mCache.SaveLocal(mMapperName, Method, SerializeData, Parameters, ExpirationDate)
End Function

