VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DODSpecificCodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private msNCID          As String
Private msCODE          As String
Private msBASE          As String
Private msEXTENDER      As String
Private msSHORT_TEXT    As String
Private msLONG_TEXT     As String
Private msDOMAIN_NCID   As String

Public Property Let NCID(value As String)
    msNCID = value
End Property
Public Property Let CODE(value As String)
    msCODE = value
End Property
Public Property Let BASE(value As String)
    msBASE = value
End Property
Public Property Let EXTENDER(value As String)
    msEXTENDER = value
End Property
Public Property Let SHORT_TEXT(value As String)
    msSHORT_TEXT = value
End Property
Public Property Let LONG_TEXT(value As String)
    msLONG_TEXT = value
End Property
Public Property Let DOMAIN_NCID(value As String)
    msDOMAIN_NCID = value
End Property

Public Property Get NCID() As String
    NCID = msNCID
End Property
Public Property Get CODE() As String
    CODE = msCODE
End Property
Public Property Get BASE() As String
    BASE = msBASE
End Property
Public Property Get EXTENDER() As String
    EXTENDER = msEXTENDER
End Property
Public Property Get SHORT_TEXT() As String
    SHORT_TEXT = msSHORT_TEXT
End Property
Public Property Get LONG_TEXT() As String
    LONG_TEXT = msLONG_TEXT
End Property
Public Property Get DOMAIN_NCID() As String
    DOMAIN_NCID = msDOMAIN_NCID
End Property
