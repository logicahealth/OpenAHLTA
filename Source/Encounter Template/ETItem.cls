VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ETItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Enum eDataType
    dt_null = 0
    dt_3m = 1
    dt_blob = 2
End Enum

Private meType As eDataType
Private mlIsRequired As Long
Private mvDataBlob As Variant



Public Property Get eType() As Integer:              eType = meType:             End Property
Public Property Let eType(ByVal RHV As Integer):   meType = RHV:              End Property

Public Property Get IsRequired() As Long:           IsRequired = mlIsRequired: End Property
Public Property Let IsRequired(ByVal RHV As Long):  mlIsRequired = RHV:        End Property

Public Property Get DataBlob() As Variant:          DataBlob = mvDataBlob:     End Property
Public Property Let DataBlob(ByVal RHV As Variant): mvDataBlob = RHV:          End Property


Public Function IsEmpty() As Long
    IsEmpty = (meType = 0)
End Function
Public Function ClearAll()
    meType = 0
    mvDataBlob = ""
End Function
Public Function ClassHandle() As Long

End Function
