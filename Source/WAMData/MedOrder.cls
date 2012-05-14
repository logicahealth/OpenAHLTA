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
Attribute VB_Name = "MedOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit
Implements IPersist

Private mComment As String
Private mICD As String
Private mOrderedBy As String
Private mOrderID As String
Private mOrderName As String
Private mSigCd As String
Private mQuantity As Long
Private mRefills As Long
Private mStartDate As String
Private mProvider As String
Private mTimeOrdered As String

Private mvarValidator As Validation

Public Property Let Comment(ByVal NewValue As String)
  mComment = NewValue
End Property

Public Property Get Comment() As String
  Comment = mComment
End Property

Public Property Get ICD() As String
  ICD = mICD
End Property

Public Property Let ICD(ByVal NewValue As String)
  mICD = NewValue
End Property

Public Property Let OrderedBy(ByVal NewValue As String)
  mOrderedBy = NewValue
End Property

Public Property Get OrderedBy() As String
  OrderedBy = mOrderedBy
End Property

Public Property Let OrderID(ByVal NewValue As String)
  mOrderID = NewValue
End Property

Public Property Get OrderID() As String
  OrderID = mOrderID
End Property

Public Property Let OrderName(ByVal NewValue As String)
  mOrderName = NewValue
End Property

Public Property Get OrderName() As String
  OrderName = mOrderName
End Property

Public Property Let Provider(ByVal NewValue As String)
  mProvider = NewValue
End Property

Public Property Get Provider() As String
  Provider = mProvider
End Property

Public Property Let Quantity(ByVal NewValue As Long)
  mQuantity = NewValue
End Property

Public Property Get Quantity() As Long
  Quantity = mQuantity
End Property

Public Property Let Refills(ByVal NewValue As Long)
  mRefills = NewValue
End Property

Public Property Get Refills() As Long
  Refills = mRefills
End Property

Public Property Let SigCode(ByVal NewValue As String)
  mSigCd = NewValue
End Property

Public Property Get SigCode() As String
  SigCode = mSigCd
End Property

Public Property Let StartDate(ByVal NewValue As String)
  mStartDate = NewValue
End Property

Public Property Get StartDate() As String
  StartDate = mStartDate
End Property

Public Property Let TimeOrdered(ByVal NewValue As String)
  mTimeOrdered = NewValue
End Property

Public Property Get TimeOrdered() As String
  TimeOrdered = mTimeOrdered
End Property

Friend Sub Init(Validatator As Validation)
    Set mvarValidator = Validatator
End Sub

''''''''''''''''''''''''''''''''''''''''''''''
' Persistence Block
''''''''''''''''''''''''''''''''''''''''''''''
Private Function IPersist_ChildFactory() As Object

End Function

Private Function IPersist_Persist() As String
    IPersist_Persist = Persist(Me)
End Function

Private Function IPersist_Restore(ByVal Persist As String) As Boolean
    IPersist_Restore = Restore(Me, Persist)
End Function


