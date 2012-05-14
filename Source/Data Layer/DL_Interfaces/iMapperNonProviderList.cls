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
Attribute VB_Name = "iMapperNonProviderList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Selected is called to add a Provider to the top of the list.  If the Provider
'does not already exist in the list, it will retrieve the information
'for the requested Provider.  Calling this with no Provider info will retrieve the
'current list
Public Function Selected(NPParams As iParameters) As DTOs.Providers

End Function

'This should only be called from the DataAccess library
Public Sub Init(ByRef Support As iParameters, ByRef Cache As CacheMapper)

End Sub

'Save the current list to the OSH (or it's corresponding data store) so that it can be
'retrieved at a later time.
Public Sub SaveList(NPParams As iParameters)

End Sub

