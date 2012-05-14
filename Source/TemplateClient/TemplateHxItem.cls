VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TemplateHxItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mobjTemplate As TemplateLite
Private mObjHxSOItem As TemplateHxSOItem
Private mObjGeneric As Object
Private msKey As String


Public Property Get TemplateItem() As TemplateLite
    Set TemplateItem = mobjTemplate
End Property
Public Property Set TemplateItem(val As TemplateLite)
    If Not val Is Nothing Then Set mobjTemplate = val
End Property
Public Property Get HxSOItem() As TemplateHxSOItem
    Set HxSOItem = mObjHxSOItem
End Property
Public Property Set HxSOItem(val As TemplateHxSOItem)
     If Not val Is Nothing Then Set mObjHxSOItem = val
End Property
Public Property Get Generic() As Object
    Generic = mObjGeneric
End Property
Public Property Set Generic(val As Object)
    If Not val Is Nothing Then Set mObjGeneric = val
End Property
Public Property Get Key() As String
    Key = msKey
End Property
Public Property Let Key(val As String)
    msKey = val
End Property
