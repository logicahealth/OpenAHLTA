VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GEMS_Phone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarType            As String
Private mvarCountryCode     As String
Private mvarAreaCode        As String
Private mvarPhone           As String
Private mvarExtension       As String
Private mvarUpdateFlag      As String

Public Property Get UpdateFlag() As String:           UpdateFlag = mvarUpdateFlag:  End Property
Public Property Let UpdateFlag(ByVal RHV As String):  mvarUpdateFlag = RHV:         End Property


Public Property Get PhoneType() As String:            PhoneType = mvarType:         End Property
Attribute PhoneType.VB_MemberFlags = "200"
Public Property Let PhoneType(ByVal RHV As String):   mvarType = RHV:               End Property

Public Property Get CountryCode() As String:          CountryCode = mvarCountryCode: End Property
Public Property Let CountryCode(ByVal RHV As String): mvarCountryCode = RHV:         End Property

Public Property Get AreaCode() As String:            AreaCode = mvarAreaCode:       End Property
Public Property Let AreaCode(ByVal RHV As String):   mvarAreaCode = RHV:            End Property

Public Property Get Phone() As String:               Phone = mvarPhone:             End Property
Public Property Let Phone(ByVal RHV As String):      mvarPhone = RHV:               End Property

Public Property Get Extension() As String:           Extension = mvarExtension:     End Property
Public Property Let Extension(ByVal RHV As String):  mvarExtension = RHV:           End Property



