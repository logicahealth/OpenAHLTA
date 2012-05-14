VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NonFDBAllergy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mDescription As String
Private mID As String

Public Property Get Description() As String
  Description = mDescription
End Property

Public Property Let Description(sValue As String)
  mDescription = sValue
End Property

Public Property Get ID() As String
  ID = mID
End Property
Public Property Let ID(sValue As String)
  mID = sValue
End Property
