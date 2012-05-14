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
Attribute VB_Name = "DODInvalidCodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private msCODE          As String
Private msDescription   As String
Private msDOMAIN_NCID   As String
Private msSubCode       As String

Public Property Let CODE(value As String)
    msCODE = value
End Property
Public Property Let Description(value As String)
    msDescription = value
End Property
Public Property Let DOMAIN_NCID(value As String)
    msDOMAIN_NCID = value
End Property
Public Property Get CODE() As String
    CODE = msCODE
End Property
Public Property Get Description() As String
    Description = msDescription
End Property
Public Property Get DOMAIN_NCID() As String
    DOMAIN_NCID = msDOMAIN_NCID
End Property

Public Property Let SubCode(ByVal NewValue As String)
  msSubCode = NewValue
End Property

Public Property Get SubCode() As String
  SubCode = msSubCode
End Property
