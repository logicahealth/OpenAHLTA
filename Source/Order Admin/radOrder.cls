VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "radOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mTextName As String 'local copy
Private mSynonyms As String 'local copy
Private mActive As Boolean 'local copy
Private mIen As Long 'local copy
Private mTestType As String 'local copy

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

Public Property Let Ien(ByVal vData As Long)
    mIen = vData
End Property

Public Property Get Ien() As Long
    Ien = mIen
End Property

Public Property Let TestType(ByVal vData As String)
    mTestType = vData
End Property

Public Property Get TestType() As String
    TestType = mTestType
End Property

