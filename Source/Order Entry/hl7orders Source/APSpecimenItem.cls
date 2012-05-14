VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "APSpecimenItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mvarNumber As Integer 'local copy
Private mvarContainer As String 'local copy
Private mvarDescription As String 'local copy
Private mvarFrozen As hlCHCS_Boolean 'local copy
Public Property Let Frozen(ByVal vData As hlCHCS_Boolean)
  mvarFrozen = vData
End Property


Public Property Get Frozen() As hlCHCS_Boolean
    Frozen = mvarFrozen
End Property



Public Property Let Description(ByVal vData As String)
    mvarDescription = vData
End Property


Public Property Get Description() As String
    Description = mvarDescription
End Property



Public Property Let Container(ByVal vData As String)
    mvarContainer = vData
End Property


Public Property Get Container() As String
    Container = mvarContainer
End Property



Public Property Let Number(ByVal vData As Integer)
    mvarNumber = vData
End Property


Public Property Get Number() As Integer
    Number = mvarNumber
End Property



