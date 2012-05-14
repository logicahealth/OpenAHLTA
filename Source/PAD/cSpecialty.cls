VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSpecialty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarSpecName        As String
Private mvarSpecNCID        As String
Private mvarBedCount        As Long

Public Property Let BedCount(ByVal vData As Long):         mvarBedCount = vData:                  End Property
Public Property Get BedCount() As Long:                    BedCount = mvarBedCount:               End Property
Public Property Let SpecName(ByVal vData As String):    mvarSpecName = vData:                   End Property
Public Property Get SpecName() As String:               SpecName = mvarSpecName:                End Property
Public Property Let SpecNCID(ByVal vData As String):    mvarSpecNCID = vData:                   End Property
Public Property Get SpecNCID() As String:               SpecNCID = mvarSpecNCID:                End Property


