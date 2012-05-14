VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SignatureInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Private mvarIen As String
Private mvarszName As String
Private mvarDateTime As String
Private mvarDateTimeHL7Format As String
Public Property Let DateTime(ByVal vData As String)
    mvarDateTime = vData
End Property
Public Property Get DateTime() As String
    DateTime = mvarDateTime
End Property
Public Property Let DateTimeHL7Format(ByVal vData As String)
    mvarDateTimeHL7Format = vData
End Property
Public Property Get DateTimeHL7Format() As String
    DateTimeHL7Format = mvarDateTimeHL7Format
End Property
Public Property Let szName(ByVal vData As String)
    mvarszName = vData
End Property
Public Property Get szName() As String
    szName = mvarszName
End Property
Public Property Let IEN(ByVal vData As String)
    mvarIen = vData
End Property
Public Property Get IEN() As String
    IEN = mvarIen
End Property



