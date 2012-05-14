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
Attribute VB_Name = "ICHCSII_Resource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"378B86E70190"
Option Base 0
Option Explicit

'##ModelId=378B99330091
Public Property Get Name() As String
End Property

'##ModelId=378B993301B3
Public Property Get UserId() As String
End Property

'##ModelId=378B993302CC
Public Property Get Password() As String
End Property

'##ModelId=378B99340006
Public Property Get ResourceData() As String
End Property

Public Function GetAttributeNames() As Collection
End Function

Public Function GetAttribute(AttributeName As String) As String
End Function

