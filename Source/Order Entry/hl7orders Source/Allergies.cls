VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Allergies"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarName As String 'local copy
Private mvarIen As String 'local copy
Private mvarComment As String 'local copy
Private mvarUpdateCode As String 'local copy
Friend Property Let UpdateCode(ByVal vData As String)
    mvarUpdateCode = vData
End Property

Public Property Get UpdateCode() As String
    UpdateCode = mvarUpdateCode
End Property

Friend Property Let Comment(ByVal vData As String)
    mvarComment = vData
End Property

Public Property Get Comment() As String
    Comment = mvarComment
End Property

Friend Property Let Ien(ByVal vData As String)
    mvarIen = vData
End Property

Public Property Get Ien() As String
    Ien = mvarIen
End Property

Friend Property Let Name(ByVal vData As String)
    mvarName = vData
End Property

Public Property Get Name() As String
    Name = mvarName
End Property
