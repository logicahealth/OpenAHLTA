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
Attribute VB_Name = "Imm_People"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Public Key As String

Private mvarSPON_SSAN_FMP As String
Private mvarPT_SSAN As String
Private mvarPT_LAST_NAME As String
Private mvarPT_FIRST_NAME As String
Private mvarPT_MIDDLE_NAME As String
Private mvarPT_SUFFIX As String
Private mvarPT_RANK As String
Private mvarPT_DOB As Date
Private mvarPT_GENDER As String
Private mvarPT_STATUS As String
Private mvarPT_SERVICE As String
Private mvarPT_UNIT As String
Private mvarPt_OFFICE_SYMBOL As String
Private mvarDataSource As String
Private mvarLastEdited As Date
Private mvarCurrent As Boolean
Private mvarDEERS_DupID As String
Private mvarDEERS_Date As Date
Private mvarExemptionExpirationDate As Date
Private mvarExemptionCode As Integer
Private PersonSelected As Boolean
Private mvarComments As String
Private mvarIEN As Long

Public Property Let lIEN(ByVal vData As Long)
  mvarIEN = vData
End Property


Public Property Get lIEN() As Long
  lIEN = mvarIEN
End Property
Public Property Let Comments(ByVal vData As String)
  mvarComments = vData
End Property


Public Property Get Comments() As String
  Comments = mvarComments
End Property

Public Property Let ExemptionCode(ByVal vData As Integer)
  mvarExemptionCode = vData
End Property


Public Property Get ExemptionCode() As Integer
  ExemptionCode = mvarExemptionCode
End Property



Public Property Let ExemptionExpirationDate(ByVal vData As Date)
  mvarExemptionExpirationDate = vData
End Property


Public Property Get ExemptionExpirationDate() As Date
  ExemptionExpirationDate = mvarExemptionExpirationDate
End Property



Public Property Let DEERS_Date(ByVal vData As Date)
  mvarDEERS_Date = vData
End Property


Public Property Get DEERS_Date() As Date
  DEERS_Date = mvarDEERS_Date
End Property



Public Property Let DEERS_DupID(ByVal vData As String)
  mvarDEERS_DupID = vData
End Property


Public Property Get DEERS_DupID() As String
  DEERS_DupID = mvarDEERS_DupID
End Property



Public Property Let Current(ByVal vData As Boolean)
  mvarCurrent = vData
End Property


Public Property Get Current() As Boolean
  Current = mvarCurrent
End Property



Public Property Let LastEdited(ByVal vData As Date)
  mvarLastEdited = vData
End Property


Public Property Get LastEdited() As Date
  LastEdited = mvarLastEdited
End Property



Public Property Let DataSource(ByVal vData As String)
  mvarDataSource = vData
End Property


Public Property Get DataSource() As String
  DataSource = mvarDataSource
End Property



Public Property Let Pt_OFFICE_SYMBOL(ByVal vData As String)
  mvarPt_OFFICE_SYMBOL = vData
End Property


Public Property Get Pt_OFFICE_SYMBOL() As String
  Pt_OFFICE_SYMBOL = mvarPt_OFFICE_SYMBOL
End Property



Public Property Let PT_UNIT(ByVal vData As String)
  mvarPT_UNIT = vData
End Property


Public Property Get PT_UNIT() As String
  PT_UNIT = mvarPT_UNIT
End Property



Public Property Let PT_SERVICE(ByVal vData As String)
  mvarPT_SERVICE = vData
End Property


Public Property Get PT_SERVICE() As String
  PT_SERVICE = mvarPT_SERVICE
End Property



Public Property Let PT_STATUS(ByVal vData As String)
  mvarPT_STATUS = vData
End Property


Public Property Get PT_STATUS() As String
  PT_STATUS = mvarPT_STATUS
End Property



Public Property Let PT_GENDER(ByVal vData As String)
  mvarPT_GENDER = vData
End Property


Public Property Get PT_GENDER() As String
  PT_GENDER = mvarPT_GENDER
End Property



Public Property Let PT_DOB(ByVal vData As Date)
  mvarPT_DOB = vData
End Property


Public Property Get PT_DOB() As Date
  PT_DOB = mvarPT_DOB
End Property



Public Property Let PT_RANK(ByVal vData As String)
  mvarPT_RANK = vData
End Property


Public Property Get PT_RANK() As String
  PT_RANK = mvarPT_RANK
End Property



Public Property Let PT_SUFFIX(ByVal vData As String)
  mvarPT_SUFFIX = vData
End Property


Public Property Get PT_SUFFIX() As String
  PT_SUFFIX = mvarPT_SUFFIX
End Property



Public Property Let PT_MIDDLE_NAME(ByVal vData As String)
  mvarPT_MIDDLE_NAME = vData
End Property


Public Property Get PT_MIDDLE_NAME() As String
  PT_MIDDLE_NAME = mvarPT_MIDDLE_NAME
End Property



Public Property Let PT_FIRST_NAME(ByVal vData As String)
  mvarPT_FIRST_NAME = vData
End Property


Public Property Get PT_FIRST_NAME() As String
  PT_FIRST_NAME = mvarPT_FIRST_NAME
End Property



Public Property Let PT_LAST_NAME(ByVal vData As String)
  mvarPT_LAST_NAME = vData
End Property


Public Property Get PT_LAST_NAME() As String
  PT_LAST_NAME = mvarPT_LAST_NAME
End Property



Public Property Let PT_SSAN(ByVal vData As String)
  mvarPT_SSAN = vData
End Property


Public Property Get PT_SSAN() As String
  PT_SSAN = mvarPT_SSAN
End Property



Public Property Let SPON_SSAN_FMP(ByVal vData As String)
  mvarSPON_SSAN_FMP = vData
End Property


Public Property Get SPON_SSAN_FMP() As String
  SPON_SSAN_FMP = mvarSPON_SSAN_FMP
End Property



