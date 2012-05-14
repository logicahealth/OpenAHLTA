VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ETXref"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private msType      As String   ' ncid
Private mnValue     As Long

Public Property Get eType() As String:          eType = msType:   End Property
Public Property Let eType(ByVal RHV As String): msType = RHV:    End Property

Public Property Get Value() As Long:           Value = mnValue: End Property
Public Property Let Value(ByVal RHV As Long):  mnValue = RHV:   End Property

Public Function IsEmpty() As Long
    IsEmpty = (msType = "")
End Function
Public Function ClearAll()
    msType = 0
    mnValue = 0
End Function
Public Function ClassHandle() As Long

End Function

