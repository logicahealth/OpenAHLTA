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
Attribute VB_Name = "labOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Private mTextName As String 'local copy
Private mSynonyms As String 'local copy
Private mActive As Boolean 'local copy
Private mValueRange As String 'local copy
Private mIen As Long 'local copy
Private mPanelFlag As Integer
Private mTestTypeCode As Integer
Private mSubScript As String
Private mQCOnly As Integer

Public Property Let Active(ByVal vData As Boolean)
    mActive = vData
End Property

Public Property Get Active() As Boolean
    Active = mActive
End Property

Public Property Let Synonyms(ByVal vData As String)
    mSynonyms = vData
End Property


Public Property Get Synonyms() As String
    Synonyms = mSynonyms
End Property

Public Property Let TextName(ByVal vData As String)
    mTextName = vData
End Property

Public Property Get TextName() As String
    TextName = mTextName
End Property

Public Property Let ValueRange(ByVal vData As String)
    mValueRange = vData
End Property

Public Property Get ValueRange() As String
    ValueRange = mValueRange
End Property

Public Property Let Ien(ByVal vData As Long)
    mIen = vData
End Property

Public Property Get Ien() As Long
    Ien = mIen
End Property

Public Property Let PanelFlag(ByVal vData As Integer)
    mPanelFlag = vData
End Property

Public Property Get PanelFlag() As Integer
    PanelFlag = mPanelFlag
End Property

Public Property Let TestTypeCode(ByVal vData As Integer)
    mTestTypeCode = vData
End Property

Public Property Get TestTypeCode() As Integer
    TestTypeCode = mTestTypeCode
End Property

Public Property Let SubScript(ByVal vData As String)
    mSubScript = vData
End Property

Public Property Get SubScript() As String
    SubScript = mSubScript
End Property

Public Property Let QCOnly(ByVal vData As Integer)
    mQCOnly = vData
End Property

Public Property Get QCOnly() As Integer
    QCOnly = mQCOnly
End Property

