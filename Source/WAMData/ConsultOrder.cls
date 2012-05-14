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
Attribute VB_Name = "ConsultOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit
Implements IPersist

Private mName As String
Private mClinic As String
Private mICD As String
Private mOrderBy As String
Private mOrderDate As String
Private mOrderID As String
Private mPriority As String
Private mProvider As String
Private mReferTo As String
Private mProvisionalDiagnosis As String

Private mvarValidator As Validation

Public Property Let Clinic(ByVal NewValue As String)
  mClinic = NewValue
End Property

Public Property Get Clinic() As String
  Clinic = mClinic
End Property

Public Property Get ICD() As String
  ICD = mICD
End Property

Public Property Let ICD(ByVal NewValue As String)
  mICD = NewValue
End Property

Public Property Let LongName(ByVal NewValue As String)
  mName = NewValue
End Property

Public Property Get LongName() As String
  LongName = mName
End Property

Public Property Let OrderBy(ByVal NewValue As String)
  mOrderBy = NewValue
End Property

Public Property Get OrderBy() As String
  OrderBy = mOrderBy
End Property

Public Property Get OrderID() As String
  OrderID = mOrderID
End Property

Public Property Let OrderID(ByVal NewValue As String)
  mOrderID = NewValue
End Property

Public Property Let OrderDate(ByVal NewValue As String)
  mOrderDate = NewValue
End Property

Public Property Get OrderDate() As String
  OrderDate = mOrderDate
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

Public Property Get ProvisionalDiagnosis() As String
  ProvisionalDiagnosis = mProvisionalDiagnosis
End Property

Public Property Let ProvisionalDiagnosis(ByVal NewValue As String)
  mProvisionalDiagnosis = NewValue
End Property

Public Property Let ReferTo(ByVal NewValue As String)
  mReferTo = NewValue
End Property

Public Property Get ReferTo() As String
  ReferTo = mReferTo
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


