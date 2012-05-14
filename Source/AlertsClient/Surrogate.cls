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
Attribute VB_Name = "Surrogate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mProviderNcid As String
Private mSurrogateNcid As String
Private mPrecedence As Integer
Private mActive As Boolean


Public Property Get ProviderNcid() As String
    Let ProviderNcid = mProviderNcid
End Property

Public Property Let ProviderNcid(ByVal Value As String)
    Let mProviderNcid = Value
End Property

Public Property Get SurrogateNcid() As String
    Let SurrogateNcid = mSurrogateNcid
End Property

Public Property Let SurrogateNcid(ByVal Value As String)
    Let mSurrogateNcid = Value
End Property

Public Property Get Precedence() As Integer
    Let Precedence = mPrecedence
End Property

Public Property Let Precedence(ByVal Value As Integer)
    Let mPrecedence = Value
End Property

Public Property Get Active() As Boolean
    Let Active = mActive
End Property

Public Property Let Active(ByVal Value As Boolean)
    Let mActive = Value
End Property

