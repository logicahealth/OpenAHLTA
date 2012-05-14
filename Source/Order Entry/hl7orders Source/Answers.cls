VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Answers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

Private mvarAnswerCode As String 'local copy
Private mvarAnswerText As String 'local copy

Friend Property Let AnswerText(ByVal vData As String)
    mvarAnswerText = vData
End Property

Public Property Get AnswerText() As String
    AnswerText = mvarAnswerText
End Property

Friend Property Let AnswerCode(ByVal vData As String)
    mvarAnswerCode = vData
End Property

Public Property Get AnswerCode() As String
    AnswerCode = mvarAnswerCode
End Property
