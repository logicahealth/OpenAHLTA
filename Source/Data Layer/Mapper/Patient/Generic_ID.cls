VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Generic_ID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarId_Type_NCID        As Long
Private mvarId_Value            As String
Private mvarAssigning_Authority As Long

Public Property Get id_type_ncid() As Long
    id_type_ncid = mvarId_Type_NCID
End Property

Public Property Let id_type_ncid(ByVal RHV As Long)
    mvarId_Type_NCID = RHV
End Property

Public Property Get id_value() As String
    id_value = mvarId_Value
End Property
Public Property Let id_value(ByVal RHV As String)
    mvarId_Value = RHV
End Property

Public Property Get assigning_authority() As Long
    assigning_authority = mvarAssigning_Authority
End Property
Public Property Let assigning_authority(ByVal RHV As Long)
    mvarAssigning_Authority = RHV
End Property



