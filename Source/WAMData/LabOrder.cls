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
Attribute VB_Name = "LabOrder"
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
Private mPriority As String
Private mProvider As String
Private mTimingStartDt As String
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

Public Property Let Priority(ByVal NewValue As String)
  mPriority = NewValue
End Property

Public Property Get Priority() As String
  Priority = mPriority
End Property

Public Property Let Provider(ByVal NewValue As String)
  mProvider = NewValue
End Property

Public Property Get Provider() As String
  Provider = mProvider
End Property

Public Property Let TimeOrdered(ByVal NewValue As String)
  mTimeOrdered = NewValue
End Property

Public Property Get TimeOrdered() As String
  TimeOrdered = mTimeOrdered
End Property

Public Property Let TimingStartDate(ByVal NewValue As String)
  mTimingStartDt = NewValue
End Property

Public Property Get TimingStartDate() As String
  TimingStartDate = mTimingStartDt
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


